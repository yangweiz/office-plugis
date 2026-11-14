using System;
using System.IO;

namespace Suite.Core.Batching
{
    public static class OutputPathPolicy
    {
        public static string EnsureDistinct(string inputPath, string outputPath)
        {
            if (string.IsNullOrWhiteSpace(inputPath)) throw new ArgumentException("输入路径不能为空。", "inputPath");
            if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("输出路径不能为空。", "outputPath");
            string input = Path.GetFullPath(inputPath);
            string output = Path.GetFullPath(outputPath);
            if (string.Equals(input, output, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("输出文件不得覆盖输入文件。");
            return output;
        }

        public static string GetAvailablePath(string requestedPath)
        {
            string fullPath = Path.GetFullPath(requestedPath);
            if (!File.Exists(fullPath) && !Directory.Exists(fullPath)) return fullPath;
            string directory = Path.GetDirectoryName(fullPath);
            string name = Path.GetFileNameWithoutExtension(fullPath);
            string extension = Path.GetExtension(fullPath);
            for (int index = 1; index < 10000; index++)
            {
                string candidate = Path.Combine(directory, string.Format("{0} ({1}){2}", name, index, extension));
                if (!File.Exists(candidate) && !Directory.Exists(candidate)) return candidate;
            }
            throw new IOException("无法生成不冲突的输出文件名。");
        }
    }
}
