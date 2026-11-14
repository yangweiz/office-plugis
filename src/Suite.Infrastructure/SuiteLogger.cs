using System;
using System.IO;
using System.Text.RegularExpressions;

namespace Suite.Infrastructure
{
    public sealed class SuiteLogger
    {
        private static readonly Regex Sensitive = new Regex(@"(?i)(password|passwd|pwd|api[_-]?key|access[_-]?token|secret|authorization)(\s*[:=]\s*)(\""[^\""]*\""|[^\s,;]+)", RegexOptions.Compiled);
        private static readonly Regex Bearer = new Regex(@"(?i)Bearer\s+[A-Za-z0-9._~+/=-]+", RegexOptions.Compiled);
        private readonly object _sync = new object();
        private readonly string _path;

        public SuiteLogger(string path) { _path = path ?? throw new ArgumentNullException("path"); }
        public string Redact(string message) { return Bearer.Replace(Sensitive.Replace(message ?? string.Empty, "$1$2[已脱敏]"), "Bearer [已脱敏]"); }

        public void Info(string message) { Write("INFO", message, null); }
        public void Error(string message, Exception exception) { Write("ERROR", message, exception); }

        private void Write(string level, string message, Exception exception)
        {
            try
            {
                lock (_sync)
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(_path));
                    string detail = exception == null ? string.Empty : " | " + exception.GetType().Name + ": " + exception.Message;
                    File.AppendAllText(_path, DateTime.UtcNow.ToString("o") + " [" + level + "] " + Redact(message + detail) + Environment.NewLine);
                }
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }
}
