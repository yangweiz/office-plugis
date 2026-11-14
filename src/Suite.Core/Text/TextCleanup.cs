using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Suite.Core.Text
{
    public static class TextCleanup
    {
        public static string RemoveInvisibleCharacters(string text)
        {
            if (text == null) return null;
            var builder = new StringBuilder(text.Length);
            foreach (char c in text)
            {
                bool invisible = c == '\u200B' || c == '\u200C' || c == '\u200D' || c == '\u2060' || c == '\uFEFF' || (char.IsControl(c) && c != '\r' && c != '\n' && c != '\t');
                if (!invisible) builder.Append(c);
            }
            return builder.ToString();
        }

        public static string CollapseBlankLines(string text)
        {
            if (text == null) return null;
            string normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
            return Regex.Replace(normalized, @"\n[ \t]*(?:\n[ \t]*)+", "\n").Replace("\n", Environment.NewLine);
        }

        public static string RemoveLeadingSpaces(string text)
        {
            if (text == null) return null;
            string normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
            return string.Join(Environment.NewLine, normalized.Split('\n').Select(x => x.TrimStart(' ', '\t', '\u3000')));
        }

        public static string ReflowBrokenLines(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            string normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
            string[] lines = normalized.Split('\n');
            var output = new List<string>();
            var current = new StringBuilder();
            foreach (string raw in lines)
            {
                string line = raw.Trim();
                if (line.Length == 0)
                {
                    FlushCurrent(output, current);
                    if (output.Count > 0 && output[output.Count - 1].Length != 0) output.Add(string.Empty);
                    continue;
                }
                if (OfficialDocumentText.IsHeading(line))
                {
                    FlushCurrent(output, current);
                    output.Add(line);
                    continue;
                }
                if (current.Length == 0) current.Append(line);
                else current.Append(NeedsSpace(current[current.Length - 1], line[0]) ? " " + line : line);
                if (EndsSentence(line)) FlushCurrent(output, current);
            }
            FlushCurrent(output, current);
            return string.Join(Environment.NewLine, output);
        }

        private static void FlushCurrent(List<string> output, StringBuilder current)
        {
            if (current.Length == 0) return;
            output.Add(current.ToString());
            current.Clear();
        }

        private static bool EndsSentence(string line)
        {
            char last = line[line.Length - 1];
            return "。！？；：.!?;:".IndexOf(last) >= 0;
        }

        private static bool NeedsSpace(char left, char right)
        {
            return (char.IsLetterOrDigit(left) && left <= 127) && (char.IsLetterOrDigit(right) && right <= 127);
        }
    }
}
