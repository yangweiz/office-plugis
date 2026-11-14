using System;
using Suite.Core.Text;
using Word = Microsoft.Office.Interop.Word;

namespace Suite.Word.Services
{
    internal sealed class WordTextEditingService
    {
        public string ReflowSelection(Word.Document document)
        {
            using (var scope = new ComScope())
            {
                Word.Range range = RequireTextSelection(scope, document);
                if (range.Text.IndexOf('\f') >= 0) throw new InvalidOperationException("选区包含分节符，请先缩小范围以避免破坏节属性。");
                string selectedText = range.Text;
                if (selectedText.IndexOf('\a') >= 0) throw new InvalidOperationException("选区包含表格结束标记，请将重排范围限制在普通正文段落。");
                int contentEnd = selectedText.Length;
                while (contentEnd > 0 && selectedText[contentEnd - 1] == '\r') contentEnd--;
                string suffix = selectedText.Substring(contentEnd);
                string original = selectedText.Substring(0, contentEnd);
                string updated = TextCleanup.ReflowBrokenLines(original).Replace("\r\n", "\r").Replace('\n', '\r') + suffix;
                object style = scope.Track(range.Style);
                range.Text = updated;
                range.Style = style;
                return string.Format("断行重排完成：{0} 个字符调整为 {1} 个字符。", original.Length, updated.Length - suffix.Length);
            }
        }

        public string CleanWhitespaceAndInvisible(Word.Document document)
        {
            int removed = 0;
            using (var scope = new ComScope())
            {
                Word.Range selection = RequireTextSelection(scope, document);
                char[] invisible = { '\u200B', '\u200C', '\u200D', '\u2060', '\uFEFF' };
                foreach (char value in invisible)
                {
                    Word.Range range = scope.Track(selection.Duplicate);
                    Word.Find find = scope.Track(range.Find);
                    find.ClearFormatting();
                    Word.Replacement replacement = scope.Track(find.Replacement);
                    replacement.ClearFormatting();
                    find.Text = value.ToString();
                    replacement.Text = string.Empty;
                    find.Forward = true;
                    find.Wrap = Word.WdFindWrap.wdFindStop;
                    find.Format = false;
                    find.MatchWildcards = false;
                    int occurrences = CountCharacter(selection.Text, value);
                    find.Execute(Replace: Word.WdReplace.wdReplaceAll);
                    removed += Math.Max(0, occurrences - CountCharacter(selection.Text, value));
                }

                Word.Paragraphs paragraphs = scope.Track(selection.Paragraphs);
                for (int index = paragraphs.Count; index >= 1; index--)
                {
                    Word.Paragraph paragraph = scope.Track(paragraphs[index]);
                    Word.Range range = scope.Track(paragraph.Range);
                    string text = range.Text;
                    if (text.IndexOf('\f') >= 0 || range.XML.IndexOf("w:sectPr", StringComparison.Ordinal) >= 0) continue;
                    string content = text.TrimEnd('\r', '\a');
                    if (content.Length == 0)
                    {
                        if (paragraphs.Count > 1) { range.Delete(); removed++; }
                        continue;
                    }
                    int leading = 0;
                    while (leading < content.Length && (content[leading] == ' ' || content[leading] == '\t' || content[leading] == '\u3000')) leading++;
                    if (leading > 0)
                    {
                        scope.Track(document.Range(range.Start, range.Start + leading)).Delete();
                        removed += leading;
                    }
                }
            }
            return "空白和不可见字符清理完成，共处理 " + removed + " 处。";
        }

        public string ConvertSelectionPunctuation(Word.Document document, bool toChinese)
        {
            return TransformSelectionCharacters(document,
                toChinese ? "中英文标点转换" : "英文标点转换",
                toChinese ? (Func<string, string>)AdvancedTextTransforms.ToChinesePunctuation : AdvancedTextTransforms.ToEnglishPunctuation);
        }

        public string ConvertSelectionWidth(Word.Document document, bool toFullWidth)
        {
            return TransformSelectionCharacters(document,
                toFullWidth ? "全角转换" : "半角转换",
                toFullWidth ? (Func<string, string>)AdvancedTextTransforms.ToFullWidth : AdvancedTextTransforms.ToHalfWidth);
        }

        public int UnlinkAndClearDirectFormatting(Word.Document document)
        {
            int unlinked = 0;
            using (var scope = new ComScope())
            {
                Word.Range range = RequireTextSelection(scope, document);
                Word.Fields initialFields = scope.Track(range.Fields);
                for (int index = initialFields.Count; index >= 1; index--)
                {
                    Word.Fields fields = scope.Track(range.Fields);
                    if (index > fields.Count) continue;
                    Word.Field field = scope.Track(fields[index]);
                    if (field.Type != Word.WdFieldType.wdFieldHyperlink) continue;
                    field.Unlink();
                    unlinked++;
                }
                Word.Font font = scope.Track(range.Font);
                Word.ParagraphFormat paragraphFormat = scope.Track(range.ParagraphFormat);
                font.Reset();
                paragraphFormat.Reset();
            }
            return unlinked;
        }

        private static string TransformSelectionCharacters(Word.Document document, string operationName, Func<string, string> transform)
        {
            using (var scope = new ComScope())
            {
                Word.Range selection = RequireTextSelection(scope, document);
                string original = selection.Text ?? string.Empty;
                if (original.Length > AdvancedTextTransforms.MaximumInputLength)
                    throw new InvalidOperationException(operationName + "单次最多处理 100000 个字符，请缩小选区。");
                if (original.IndexOf('\f') >= 0 || original.IndexOf('\a') >= 0)
                    throw new InvalidOperationException(operationName + "不支持跨分节符或表格结束标记的选区，请缩小到普通正文文本。");
                Word.Fields fields = scope.Track(selection.Fields);
                if (fields.Count > 0)
                    throw new InvalidOperationException(operationName + "不会改写域或超链接结果，请选择不含域的纯文本范围。");

                string updated = transform(original) ?? string.Empty;
                if (updated.Length != original.Length)
                    throw new InvalidOperationException(operationName + "的内部转换结果长度不一致，已停止修改。");

                int changed = 0;
                for (int index = 0; index < original.Length; index++)
                {
                    if (original[index] == updated[index]) continue;
                    Word.Range character = null;
                    try
                    {
                        character = document.Range(selection.Start + index, selection.Start + index + 1);
                        character.Text = updated[index].ToString();
                        changed++;
                    }
                    finally
                    {
                        if (character != null && System.Runtime.InteropServices.Marshal.IsComObject(character))
                            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(character);
                    }
                }
                return changed == 0
                    ? operationName + "完成，没有需要修改的字符；URL 和邮箱保持不变。"
                    : string.Format("{0}完成，逐字符修改 {1} 处并保留原字符格式；URL 和邮箱保持不变。", operationName, changed);
            }
        }

        private static int CountCharacter(string text, char value)
        {
            int count = 0;
            foreach (char current in text ?? string.Empty) if (current == value) count++;
            return count;
        }

        private static Word.Range RequireTextSelection(ComScope scope, Word.Document document)
        {
            Word.Application application = scope.Track(document.Application);
            Word.Selection selection = scope.Track(application.Selection);
            if (selection == null || selection.Type == Word.WdSelectionType.wdSelectionIP)
                throw new InvalidOperationException("请先选择要处理的文本；该功能不会默认修改整篇文档。");
            return scope.Track(selection.Range);
        }
    }
}
