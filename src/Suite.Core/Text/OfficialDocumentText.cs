using System;
using System.Text.RegularExpressions;

namespace Suite.Core.Text
{
    public static class OfficialDocumentText
    {
        private static readonly Regex HeadingPattern = new Regex(
            @"^(第[一二三四五六七八九十百千万0-9]+[章节条](?:\s|$)|[一二三四五六七八九十]+[、.．]|[(（][一二三四五六七八九十]+[)）]|\d+[、.．]\s*|\d+(?:\.\d+){1,5}\s+)",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        public static bool IsHeading(string text)
        {
            return !string.IsNullOrWhiteSpace(text) && HeadingPattern.IsMatch(text.Trim());
        }

        public static int PromoteNumberedHeading(string text, out string result)
        {
            result = text;
            if (string.IsNullOrWhiteSpace(text)) return 0;
            string trimmed = text.Trim();
            if (Regex.IsMatch(trimmed, @"^[(（][一二三四五六七八九十]+[)）]"))
            {
                result = Regex.Replace(trimmed, @"^[(（]([一二三四五六七八九十]+)[)）]", "$1、");
                return 1;
            }
            if (Regex.IsMatch(trimmed, @"^\d+[、.．]"))
            {
                result = Regex.Replace(trimmed, @"^(\d+)[、.．]", "（$1）");
                return 1;
            }
            return 0;
        }
    }
}
