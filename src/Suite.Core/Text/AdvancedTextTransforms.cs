using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace Suite.Core.Text
{
    public sealed class RegexMatchPreview
    {
        internal RegexMatchPreview(int index, int length, string value, string replacement)
        {
            Index = index;
            Length = length;
            Value = value;
            Replacement = replacement;
        }

        public int Index { get; private set; }
        public int Length { get; private set; }
        public string Value { get; private set; }
        public string Replacement { get; private set; }
    }

    public sealed class RegexReplacementResult
    {
        internal RegexReplacementResult(string output, int totalMatches, IList<RegexMatchPreview> previews)
        {
            Output = output;
            TotalMatches = totalMatches;
            Previews = previews.AsReadOnly();
        }

        public string Output { get; private set; }
        public int TotalMatches { get; private set; }
        public IReadOnlyList<RegexMatchPreview> Previews { get; private set; }
    }

    public static class AdvancedTextTransforms
    {
        public const int MaximumInputLength = 100000;
        public const int MaximumPatternLength = 1000;
        public const int MaximumPreviewCount = 100;
        public const decimal MaximumCurrencyAmount = 99999999999999.99m;

        private static readonly Regex ProtectedTextPattern = new Regex(
            @"https?://[^\s<>""'，。；！？（）【】]+|(?<![A-Z0-9._%+\-])[A-Z0-9._%+\-]+@[A-Z0-9](?:[A-Z0-9\-]{0,61}[A-Z0-9])?(?:\.[A-Z0-9](?:[A-Z0-9\-]{0,61}[A-Z0-9])?)+",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant,
            TimeSpan.FromSeconds(1));

        private static readonly char[] CurrencyDigits = "零壹贰叁肆伍陆柒捌玖".ToCharArray();
        private static readonly string[] GroupUnits = { string.Empty, "万", "亿", "兆" };

        public static string ToChinesePunctuation(string text)
        {
            bool openingDoubleQuote = true;
            return TransformUnprotected(text, delegate(char value)
            {
                switch (value)
                {
                    case ',': return '，';
                    case '.': return '。';
                    case ':': return '：';
                    case ';': return '；';
                    case '?': return '？';
                    case '!': return '！';
                    case '(': return '（';
                    case ')': return '）';
                    case '"':
                        char quote = openingDoubleQuote ? '“' : '”';
                        openingDoubleQuote = !openingDoubleQuote;
                        return quote;
                    default: return value;
                }
            });
        }

        public static string ToEnglishPunctuation(string text)
        {
            return TransformUnprotected(text, delegate(char value)
            {
                switch (value)
                {
                    case '，': return ',';
                    case '。':
                    case '．': return '.';
                    case '、': return ',';
                    case '：': return ':';
                    case '；': return ';';
                    case '？': return '?';
                    case '！': return '!';
                    case '（': return '(';
                    case '）': return ')';
                    case '【':
                    case '［': return '[';
                    case '】':
                    case '］': return ']';
                    case '“':
                    case '”':
                    case '＂': return '"';
                    case '‘':
                    case '’':
                    case '＇': return '\'';
                    case '～': return '~';
                    default: return value;
                }
            });
        }

        public static string ToFullWidth(string text)
        {
            return TransformUnprotected(text, delegate(char value)
            {
                if (value == ' ') return '\u3000';
                if (value >= '!' && value <= '~') return (char)(value + 0xFEE0);
                return value;
            });
        }

        public static string ToHalfWidth(string text)
        {
            return TransformUnprotected(text, delegate(char value)
            {
                if (value == '\u3000') return ' ';
                if (value >= '\uFF01' && value <= '\uFF5E') return (char)(value - 0xFEE0);
                return value;
            });
        }

        public static string ToChineseCurrency(decimal amount)
        {
            if (amount > MaximumCurrencyAmount || amount < -MaximumCurrencyAmount)
                throw new ArgumentOutOfRangeException("amount", "金额超出可转换范围。");
            if (amount != decimal.Round(amount, 2))
                throw new ArgumentException("金额最多支持两位小数。", "amount");

            bool negative = amount < 0;
            decimal absolute = Math.Abs(amount);
            long integerPart = (long)decimal.Truncate(absolute);
            int cents = (int)((absolute - integerPart) * 100m);
            if (integerPart == 0 && cents == 0) return "零元整";

            var result = new StringBuilder();
            if (negative) result.Append('负');
            if (integerPart > 0)
            {
                result.Append(ConvertInteger(integerPart));
                result.Append('元');
            }

            int jiao = cents / 10;
            int fen = cents % 10;
            if (jiao == 0 && fen == 0)
            {
                result.Append('整');
                return result.ToString();
            }
            if (jiao > 0)
            {
                result.Append(CurrencyDigits[jiao]);
                result.Append('角');
            }
            else if (integerPart > 0 && fen > 0)
            {
                result.Append('零');
            }
            if (fen > 0)
            {
                result.Append(CurrencyDigits[fen]);
                result.Append('分');
            }
            return result.ToString();
        }

        public static RegexReplacementResult ReplaceRegex(string input, string pattern, string replacement, int previewLimit = 10, int timeoutMilliseconds = 1000)
        {
            if (input == null) throw new ArgumentNullException("input");
            if (pattern == null) throw new ArgumentNullException("pattern");
            if (replacement == null) throw new ArgumentNullException("replacement");
            if (input.Length > MaximumInputLength) throw new ArgumentOutOfRangeException("input", "正则输入长度超出限制。");
            if (pattern.Length == 0) throw new ArgumentException("正则模式不能为空。", "pattern");
            if (pattern.Length > MaximumPatternLength) throw new ArgumentOutOfRangeException("pattern", "正则模式长度超出限制。");
            if (previewLimit < 0 || previewLimit > MaximumPreviewCount) throw new ArgumentOutOfRangeException("previewLimit");
            if (timeoutMilliseconds <= 0) throw new ArgumentOutOfRangeException("timeoutMilliseconds");

            var previews = new List<RegexMatchPreview>();
            int totalMatches = 0;
            var regex = new Regex(pattern, RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(timeoutMilliseconds));
            string output = regex.Replace(input, delegate(Match match)
            {
                string substituted = match.Result(replacement);
                totalMatches++;
                if (previews.Count < previewLimit)
                    previews.Add(new RegexMatchPreview(match.Index, match.Length, match.Value, substituted));
                return substituted;
            });
            return new RegexReplacementResult(output, totalMatches, previews);
        }

        private static string TransformUnprotected(string text, Func<char, char> transform)
        {
            if (string.IsNullOrEmpty(text)) return text;
            var result = new StringBuilder(text.Length);
            int position = 0;
            foreach (Match match in ProtectedTextPattern.Matches(text))
            {
                int length = match.Length;
                if (match.Value.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                    match.Value.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                {
                    while (length > 0 && IsTrailingUrlPunctuation(match.Value[length - 1])) length--;
                }
                if (length == 0) continue;

                AppendTransformed(result, text, position, match.Index - position, transform);
                result.Append(text, match.Index, length);
                position = match.Index + length;
            }
            AppendTransformed(result, text, position, text.Length - position, transform);
            return result.ToString();
        }

        private static bool IsTrailingUrlPunctuation(char value)
        {
            return value == '.' || value == ',' || value == ';' || value == ':' || value == '!' || value == '?' ||
                   value == ')' || value == ']' || value == '}';
        }

        private static void AppendTransformed(StringBuilder result, string text, int start, int count, Func<char, char> transform)
        {
            for (int index = start; index < start + count; index++) result.Append(transform(text[index]));
        }

        private static string ConvertInteger(long value)
        {
            var groups = new int[4];
            int groupCount = 0;
            while (value > 0)
            {
                groups[groupCount++] = (int)(value % 10000);
                value /= 10000;
            }

            var result = new StringBuilder();
            bool pendingZero = false;
            for (int index = groupCount - 1; index >= 0; index--)
            {
                int group = groups[index];
                if (group == 0)
                {
                    if (result.Length > 0) pendingZero = true;
                    continue;
                }
                if (result.Length > 0 && (pendingZero || group < 1000)) result.Append('零');
                result.Append(ConvertGroup(group));
                result.Append(GroupUnits[index]);
                pendingZero = false;
            }
            return result.ToString();
        }

        private static string ConvertGroup(int value)
        {
            var divisors = new[] { 1000, 100, 10, 1 };
            var units = new[] { "仟", "佰", "拾", string.Empty };
            var result = new StringBuilder();
            bool pendingZero = false;
            for (int index = 0; index < divisors.Length; index++)
            {
                int digit = value / divisors[index] % 10;
                if (digit == 0)
                {
                    if (result.Length > 0) pendingZero = true;
                    continue;
                }
                if (pendingZero) result.Append('零');
                result.Append(CurrencyDigits[digit]);
                result.Append(units[index]);
                pendingZero = false;
            }
            return result.ToString();
        }
    }
}
