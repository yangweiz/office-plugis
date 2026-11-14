using System;
using System.Linq;

namespace Suite.Core.Tests
{
    internal static class AdvancedTextTransformsTests
    {
        public static void Run()
        {
            TestRunner.Run("AdvancedTextTransforms - 中文标点", ChinesePunctuation);
            TestRunner.Run("AdvancedTextTransforms - 英文标点及 URL/邮箱保护", EnglishPunctuation);
            TestRunner.Run("AdvancedTextTransforms - 全角及保护区间的往返", FullWidth);
            TestRunner.Run("AdvancedTextTransforms - 半角及空输入", HalfWidthAndEmpty);
            TestRunner.Run("AdvancedTextTransforms - 人民币常用金额", CurrencyCommon);
            TestRunner.Run("AdvancedTextTransforms - 人民币跨位零规则", CurrencyZeros);
            TestRunner.Run("AdvancedTextTransforms - 人民币精度和范围异常", CurrencyValidation);
            TestRunner.Run("AdvancedTextTransforms - 正则替换与预览上限", RegexReplacement);
            TestRunner.Run("AdvancedTextTransforms - 正则空结果和零预览", RegexEmptyAndZeroPreview);
            TestRunner.Run("AdvancedTextTransforms - 正则参数限制", RegexValidation);
            TestRunner.Run("AdvancedTextTransforms - 灾难性正则超时", RegexTimeout);
        }

        private static void ChinesePunctuation()
        {
            TestAssert.Equal("“甲”，（OK）！", Suite.Core.Text.AdvancedTextTransforms.ToChinesePunctuation("\"甲\",(OK)!"));
            const string input = "链接 http://example.com/a,b. 邮箱 test+tag@example.com!";
            const string expected = "链接 http://example.com/a,b。 邮箱 test+tag@example.com！";
            TestAssert.Equal(expected, Suite.Core.Text.AdvancedTextTransforms.ToChinesePunctuation(input));
        }

        private static void EnglishPunctuation()
        {
            const string input = "“中文”，测试。http://a.test/x,y。 a@b.com！";
            TestAssert.Equal("\"中文\",测试.http://a.test/x,y. a@b.com!", Suite.Core.Text.AdvancedTextTransforms.ToEnglishPunctuation(input));
            TestAssert.Equal("[甲],'乙'~", Suite.Core.Text.AdvancedTextTransforms.ToEnglishPunctuation("【甲】、‘乙’～"));
        }

        private static void FullWidth()
        {
            const string input = "A1! http://example.com/a test@example.com";
            const string expected = "Ａ１！\u3000http://example.com/a\u3000test@example.com";
            TestAssert.Equal(expected, Suite.Core.Text.AdvancedTextTransforms.ToFullWidth(input));
        }

        private static void HalfWidthAndEmpty()
        {
            const string input = "Ａ１！\u3000http://example.com/a\u3000test@example.com";
            TestAssert.Equal("A1! http://example.com/a test@example.com", Suite.Core.Text.AdvancedTextTransforms.ToHalfWidth(input));
            TestAssert.Equal<string>(null, Suite.Core.Text.AdvancedTextTransforms.ToFullWidth(null));
            TestAssert.Equal(string.Empty, Suite.Core.Text.AdvancedTextTransforms.ToChinesePunctuation(string.Empty));
        }

        private static void CurrencyCommon()
        {
            TestAssert.Equal("零元整", Suite.Core.Text.AdvancedTextTransforms.ToChineseCurrency(0m));
            TestAssert.Equal("壹仟贰佰叁拾肆元伍角陆分", Suite.Core.Text.AdvancedTextTransforms.ToChineseCurrency(1234.56m));
            TestAssert.Equal("壹仟零壹元零伍分", Suite.Core.Text.AdvancedTextTransforms.ToChineseCurrency(1001.05m));
            TestAssert.Equal("负壹拾元壹角", Suite.Core.Text.AdvancedTextTransforms.ToChineseCurrency(-10.10m));
            TestAssert.Equal("伍分", Suite.Core.Text.AdvancedTextTransforms.ToChineseCurrency(0.05m));
            TestAssert.Equal("壹元贰角叁分", Suite.Core.Text.AdvancedTextTransforms.ToChineseCurrency(1.2300m));
        }

        private static void CurrencyZeros()
        {
            TestAssert.Equal("壹亿零壹元整", Suite.Core.Text.AdvancedTextTransforms.ToChineseCurrency(100000001m));
            TestAssert.Equal("壹亿零壹万零壹元整", Suite.Core.Text.AdvancedTextTransforms.ToChineseCurrency(100010001m));
            TestAssert.Equal("壹亿元整", Suite.Core.Text.AdvancedTextTransforms.ToChineseCurrency(100000000m));
            string maximum = Suite.Core.Text.AdvancedTextTransforms.ToChineseCurrency(Suite.Core.Text.AdvancedTextTransforms.MaximumCurrencyAmount);
            TestAssert.True(maximum.Contains("兆"));
            TestAssert.True(maximum.EndsWith("元玖角玖分", StringComparison.Ordinal));
        }

        private static void CurrencyValidation()
        {
            TestAssert.Throws<ArgumentException>(() => Suite.Core.Text.AdvancedTextTransforms.ToChineseCurrency(1.001m));
            TestAssert.Throws<ArgumentOutOfRangeException>(() => Suite.Core.Text.AdvancedTextTransforms.ToChineseCurrency(Suite.Core.Text.AdvancedTextTransforms.MaximumCurrencyAmount + 0.01m));
            TestAssert.Throws<ArgumentOutOfRangeException>(() => Suite.Core.Text.AdvancedTextTransforms.ToChineseCurrency(-Suite.Core.Text.AdvancedTextTransforms.MaximumCurrencyAmount - 0.01m));
        }

        private static void RegexReplacement()
        {
            Suite.Core.Text.RegexReplacementResult result = Suite.Core.Text.AdvancedTextTransforms.ReplaceRegex("a1 a2 a3", @"a(\d)", "[$1]", 2);
            TestAssert.Equal("[1] [2] [3]", result.Output);
            TestAssert.Equal(3, result.TotalMatches);
            TestAssert.Equal(2, result.Previews.Count);
            TestAssert.Equal(0, result.Previews[0].Index);
            TestAssert.Equal(2, result.Previews[0].Length);
            TestAssert.Equal("a1", result.Previews[0].Value);
            TestAssert.Equal("[2]", result.Previews[1].Replacement);
        }

        private static void RegexEmptyAndZeroPreview()
        {
            Suite.Core.Text.RegexReplacementResult empty = Suite.Core.Text.AdvancedTextTransforms.ReplaceRegex(string.Empty, "a", "b");
            TestAssert.Equal(string.Empty, empty.Output);
            TestAssert.Equal(0, empty.TotalMatches);
            Suite.Core.Text.RegexReplacementResult noPreview = Suite.Core.Text.AdvancedTextTransforms.ReplaceRegex("aaa", "a", "b", 0);
            TestAssert.Equal("bbb", noPreview.Output);
            TestAssert.Equal(3, noPreview.TotalMatches);
            TestAssert.Equal(0, noPreview.Previews.Count);
        }

        private static void RegexValidation()
        {
            TestAssert.Throws<ArgumentNullException>(() => Suite.Core.Text.AdvancedTextTransforms.ReplaceRegex(null, "a", "b"));
            TestAssert.Throws<ArgumentException>(() => Suite.Core.Text.AdvancedTextTransforms.ReplaceRegex("a", string.Empty, "b"));
            TestAssert.Throws<ArgumentException>(() => Suite.Core.Text.AdvancedTextTransforms.ReplaceRegex("a", "(a+", "b"));
            TestAssert.Throws<ArgumentOutOfRangeException>(() => Suite.Core.Text.AdvancedTextTransforms.ReplaceRegex(new string('a', Suite.Core.Text.AdvancedTextTransforms.MaximumInputLength + 1), "a", "b"));
            TestAssert.Throws<ArgumentOutOfRangeException>(() => Suite.Core.Text.AdvancedTextTransforms.ReplaceRegex("a", new string('a', Suite.Core.Text.AdvancedTextTransforms.MaximumPatternLength + 1), "b"));
            TestAssert.Throws<ArgumentOutOfRangeException>(() => Suite.Core.Text.AdvancedTextTransforms.ReplaceRegex("a", "a", "b", Suite.Core.Text.AdvancedTextTransforms.MaximumPreviewCount + 1));
            TestAssert.Throws<ArgumentOutOfRangeException>(() => Suite.Core.Text.AdvancedTextTransforms.ReplaceRegex("a", "a", "b", 1, 0));
        }

        private static void RegexTimeout()
        {
            string input = new string('a', 40) + "!";
            TestAssert.Throws<TimeoutException>(() => Suite.Core.Text.AdvancedTextTransforms.ReplaceRegex(input, "^(a+)+$", "b", 1, 50));
        }
    }
}
