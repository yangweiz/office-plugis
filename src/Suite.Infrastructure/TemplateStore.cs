using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Suite.Core.Batching;

namespace Suite.Infrastructure
{
    public sealed class TemplateDescriptor
    {
        public TemplateDescriptor(string name, string host, string path) { Name = name; Host = host; Path = path; }
        public string Name { get; private set; }
        public string Host { get; private set; }
        public string Path { get; private set; }
    }

    public sealed class TemplateStore
    {
        private static readonly Dictionary<string, string> Extensions = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { ".dotx", "Word" }, { ".xltx", "Excel" }, { ".potx", "PowerPoint" }
        };
        private readonly string _directory;

        public TemplateStore(string directory)
        {
            if (string.IsNullOrWhiteSpace(directory)) throw new ArgumentException("模板目录不能为空。", "directory");
            _directory = NormalizeDirectory(directory);
        }

        public IReadOnlyList<TemplateDescriptor> List()
        {
            if (!Directory.Exists(_directory)) return new List<TemplateDescriptor>();
            return Directory.EnumerateFiles(_directory)
                .Where(path => Extensions.ContainsKey(Path.GetExtension(path)))
                .Select(path => new TemplateDescriptor(Path.GetFileNameWithoutExtension(path), Extensions[Path.GetExtension(path)], path))
                .OrderBy(item => item.Name, StringComparer.CurrentCulture)
                .ToList();
        }

        public TemplateDescriptor Import(string sourcePath)
        {
            if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath)) throw new FileNotFoundException("模板源文件不存在。", sourcePath);
            string extension = Path.GetExtension(sourcePath);
            string host;
            if (!Extensions.TryGetValue(extension, out host)) throw new InvalidOperationException("仅支持无宏的 DOTX、XLTX、POTX 原创模板；宏模板不会自动导入或执行。");
            Directory.CreateDirectory(_directory);
            string destination = OutputPathPolicy.GetAvailablePath(Path.Combine(_directory, Path.GetFileName(sourcePath)));
            File.Copy(Path.GetFullPath(sourcePath), destination, false);
            return new TemplateDescriptor(Path.GetFileNameWithoutExtension(destination), host, destination);
        }

        public string Export(TemplateDescriptor template, string destinationDirectory)
        {
            if (template == null) throw new ArgumentNullException("template");
            if (!IsOwnedTemplate(template.Path) || !File.Exists(template.Path)) throw new FileNotFoundException("模板不存在或不属于当前模板库。", template.Path);
            Directory.CreateDirectory(destinationDirectory);
            string destination = OutputPathPolicy.GetAvailablePath(Path.Combine(destinationDirectory, Path.GetFileName(template.Path)));
            File.Copy(template.Path, destination, false);
            return destination;
        }

        public void Delete(TemplateDescriptor template)
        {
            if (template == null) throw new ArgumentNullException("template");
            if (!IsOwnedTemplate(template.Path)) throw new InvalidOperationException("只能删除当前模板库中的文件。");
            if (File.Exists(template.Path)) File.Delete(template.Path);
        }

        private static string NormalizeDirectory(string directory)
        {
            string full = Path.GetFullPath(directory);
            string root = Path.GetPathRoot(full);
            if (string.Equals(full, root, StringComparison.OrdinalIgnoreCase)) return full;
            return full.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        }

        private bool IsOwnedTemplate(string path)
        {
            string full = Path.GetFullPath(path);
            string parent = Path.GetDirectoryName(full);
            return string.Equals(parent, _directory, StringComparison.OrdinalIgnoreCase) && Extensions.ContainsKey(Path.GetExtension(full));
        }
    }
}
