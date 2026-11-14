using System;
using System.Diagnostics;
using System.IO;
using Suite.Core.Batching;

namespace Suite.Infrastructure
{
    public sealed class ExternalConverter
    {
        private readonly string _executablePath;
        private readonly string _argumentTemplate;

        public ExternalConverter(string executablePath, string argumentTemplate)
        {
            if (string.IsNullOrWhiteSpace(executablePath) || !File.Exists(executablePath)) throw new FileNotFoundException("转换器不存在，请先配置用户合法授权的本地转换器。", executablePath);
            if (!string.Equals(Path.GetExtension(executablePath), ".exe", StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("转换器必须是明确的本地 EXE 文件。");
            if (string.IsNullOrWhiteSpace(argumentTemplate) || !argumentTemplate.Contains("{Input}") || !argumentTemplate.Contains("{Output}"))
                throw new ArgumentException("转换参数必须同时包含 {Input} 和 {Output} 占位符。", "argumentTemplate");
            _executablePath = Path.GetFullPath(executablePath);
            _argumentTemplate = argumentTemplate;
        }

        public string Convert(string inputPath, string requestedOutputPath, TimeSpan timeout)
        {
            if (!File.Exists(inputPath)) throw new FileNotFoundException("待转换文件不存在。", inputPath);
            if (timeout <= TimeSpan.Zero || timeout > TimeSpan.FromMinutes(30)) throw new ArgumentOutOfRangeException("timeout");
            string input = Path.GetFullPath(inputPath);
            string output = OutputPathPolicy.EnsureDistinct(input, OutputPathPolicy.GetAvailablePath(requestedOutputPath));
            Directory.CreateDirectory(Path.GetDirectoryName(output));
            string arguments = _argumentTemplate.Replace("{Input}", Quote(input)).Replace("{Output}", Quote(output));
            var start = new ProcessStartInfo
            {
                FileName = _executablePath,
                Arguments = arguments,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                WorkingDirectory = Path.GetDirectoryName(_executablePath)
            };
            using (var process = new Process { StartInfo = start })
            {
                if (!process.Start()) throw new InvalidOperationException("转换器未能启动。");
                var standardOutput = process.StandardOutput.ReadToEndAsync();
                var standardError = process.StandardError.ReadToEndAsync();
                if (!process.WaitForExit((int)timeout.TotalMilliseconds))
                {
                    try { process.Kill(); } catch (InvalidOperationException) { }
                    process.WaitForExit();
                    throw new TimeoutException("转换器执行超时，已终止；未将任何输出声明为成功。");
                }
                string error = standardError.GetAwaiter().GetResult();
                string standard = standardOutput.GetAwaiter().GetResult();
                if (process.ExitCode != 0) throw new InvalidOperationException("转换器返回退出码 " + process.ExitCode + "：" + Limit(string.IsNullOrWhiteSpace(error) ? standard : error));
            }
            if (!File.Exists(output) || new FileInfo(output).Length == 0) throw new InvalidOperationException("转换器未生成有效的非空输出文件。");
            return output;
        }

        private static string Quote(string path) { return "\"" + path + "\""; }
        private static string Limit(string value) { return (value ?? string.Empty).Trim().Substring(0, Math.Min(1000, (value ?? string.Empty).Trim().Length)); }
    }
}
