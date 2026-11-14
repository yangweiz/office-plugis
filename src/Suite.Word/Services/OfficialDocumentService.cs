using System;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Suite.Core.Text;
using Word = Microsoft.Office.Interop.Word;

namespace Suite.Word.Services
{
    internal sealed class OfficialDocumentService
    {
        internal const string BodyStyle = "公文正文";
        internal const string Heading1Style = "公文一级标题";
        internal const string Heading2Style = "公文二级标题";
        internal const string Heading3Style = "公文三级标题";
        internal const string Heading4Style = "公文四级标题";

        public void ApplyPageSetup(Word.Document document, bool applyGrid)
        {
            using (var scope = new ComScope())
            {
                Word.Sections sections = scope.Track(document.Sections);
                for (int index = 1; index <= sections.Count; index++)
                {
                    Word.Section section = scope.Track(sections[index]);
                    Word.PageSetup setup = scope.Track(section.PageSetup);
                    setup.PaperSize = Word.WdPaperSize.wdPaperA4;
                    setup.Orientation = Word.WdOrientation.wdOrientPortrait;
                    setup.TopMargin = Millimeters(document, 37);
                    setup.BottomMargin = Millimeters(document, 35);
                    setup.LeftMargin = Millimeters(document, 28);
                    setup.RightMargin = Millimeters(document, 26);
                    setup.HeaderDistance = Millimeters(document, 15);
                    setup.FooterDistance = Millimeters(document, 17);
                    if (applyGrid) ApplyGrid(setup);
                }
            }
        }

        private static void ApplyGrid(Word.PageSetup setup)
        {
            setup.LayoutMode = Word.WdLayoutMode.wdLayoutModeGrid;
            setup.CharsLine = 28;
            setup.LinesPage = 22;
        }

        public void EnsureStyles(Word.Document document)
        {
            using (var scope = new ComScope())
            {
                EnsureStyle(scope, document, BodyStyle, "仿宋", 16, false, 0, 28, 32);
                EnsureStyle(scope, document, Heading1Style, "黑体", 16, false, 0, 28, 0);
                EnsureStyle(scope, document, Heading2Style, "楷体", 16, false, 0, 28, 0);
                EnsureStyle(scope, document, Heading3Style, "仿宋", 16, true, 0, 28, 0);
                EnsureStyle(scope, document, Heading4Style, "仿宋", 16, false, 0, 28, 0);
            }
        }

        public void ApplyBodyToSelection(Word.Document document)
        {
            EnsureStyles(document);
            using (var scope = new ComScope())
            {
                Word.Selection selection = RequireSelection(scope, document);
                Word.Range range = scope.Track(selection.Paragraphs.Range);
                range.Style = GetStyle(scope, document, BodyStyle);
            }
        }

        public string AutoFormat(Word.Document document)
        {
            EnsureStyles(document);
            int headings = 0;
            int body = 0;
            int skipped = 0;
            using (var scope = new ComScope())
            {
                Word.Paragraphs paragraphs = scope.Track(document.Paragraphs);
                for (int index = 1; index <= paragraphs.Count; index++)
                {
                    Word.Paragraph paragraph = scope.Track(paragraphs[index]);
                    Word.Range range = scope.Track(paragraph.Range);
                    string text = range.Text.Trim('\r', '\a', ' ', '\t', '\u3000');
                    if (text.Length == 0) { skipped++; continue; }
                    if ((bool)range.Information[Word.WdInformation.wdWithInTable]) { skipped++; continue; }
                    int level = HeadingLevel(text);
                    if (level > 0)
                    {
                        range.Style = GetStyle(scope, document, StyleForLevel(level));
                        headings++;
                    }
                    else
                    {
                        range.Style = GetStyle(scope, document, BodyStyle);
                        body++;
                    }
                }
            }
            return string.Format("已处理 {0} 个正文段落、{1} 个标题；跳过 {2} 个空段或表格段落。", body, headings, skipped);
        }

        public int PromoteSelectionHeadings(Word.Document document)
        {
            EnsureStyles(document);
            int changed = 0;
            using (var scope = new ComScope())
            {
                Word.Selection selection = RequireSelection(scope, document);
                Word.Paragraphs paragraphs = scope.Track(selection.Paragraphs);
                for (int index = paragraphs.Count; index >= 1; index--)
                {
                    Word.Paragraph paragraph = scope.Track(paragraphs[index]);
                    Word.Range original = scope.Track(paragraph.Range);
                    string current = original.Text.TrimEnd('\r', '\a');
                    string promoted;
                    if (OfficialDocumentText.PromoteNumberedHeading(current, out promoted) == 0) continue;
                    Word.Range content = scope.Track(document.Range(original.Start, Math.Max(original.Start, original.End - 1)));
                    content.Text = promoted;
                    int level = HeadingLevel(promoted);
                    if (level > 0) content.Style = GetStyle(scope, document, StyleForLevel(level));
                    changed++;
                }
            }
            return changed;
        }

        public void MakeCurrentPageLandscape(Word.Document document)
        {
            using (var scope = new ComScope())
            {
                Word.Selection selection = RequireSelection(scope, document);
                Word.Range selected = scope.Track(selection.Range);
                Word.Bookmark pageBookmark = scope.Track(selected.Bookmarks[@"\Page"]);
                Word.Range page = scope.Track(pageBookmark.Range);
                int start = page.Start;
                int end = page.End;
                Word.Range after = scope.Track(document.Range(end, end));
                after.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);
                Word.Range before = scope.Track(document.Range(start, start));
                before.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);
                Word.Sections selectedSections = scope.Track(selection.Range.Sections);
                Word.Section current = scope.Track(selectedSections[1]);
                Word.PageSetup setup = scope.Track(current.PageSetup);
                float width = setup.PageWidth;
                float height = setup.PageHeight;
                setup.Orientation = Word.WdOrientation.wdOrientLandscape;
                if (width < height)
                {
                    setup.PageWidth = height;
                    setup.PageHeight = width;
                }
            }
        }

        public void InsertPageNumbers(Word.Document document)
        {
            using (var scope = new ComScope())
            {
                Word.Sections sections = scope.Track(document.Sections);
                for (int index = 1; index <= sections.Count; index++)
                {
                    Word.HeaderFooter footer = PrimaryFooter(scope, sections, index);
                    if (!HasPageField(scope, footer) && FooterText(footer).Length > 0)
                        throw new InvalidOperationException("第 " + index + " 节页脚已有内容。为避免覆盖，已停止；请先清理或手动处理该页脚。");
                }
                for (int index = 1; index <= sections.Count; index++)
                {
                    Word.HeaderFooter footer = PrimaryFooter(scope, sections, index);
                    if (HasPageField(scope, footer)) continue;
                    footer.LinkToPrevious = false;
                    Word.Range range = scope.Track(footer.Range);
                    range.Text = "—  —";
                    range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    Word.Range fieldRange = scope.Track(document.Range(range.Start + 2, range.Start + 2));
                    Word.Field field = scope.Track(range.Fields.Add(fieldRange, Word.WdFieldType.wdFieldPage));
                    scope.Track(field.Result).Font.Size = 14;
                    range.Font.NameFarEast = "宋体";
                    range.Font.Size = 14;
                }
            }
        }

        public void ClearHeadersAndFooters(Word.Document document)
        {
            using (var scope = new ComScope())
            {
                Word.Sections sections = scope.Track(document.Sections);
                for (int index = 1; index <= sections.Count; index++)
                {
                    Word.Section section = scope.Track(sections[index]);
                    foreach (Word.WdHeaderFooterIndex kind in new[] { Word.WdHeaderFooterIndex.wdHeaderFooterPrimary, Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage, Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages })
                    {
                        ClearHeaderFooter(scope, scope.Track(section.Headers[kind]));
                        ClearHeaderFooter(scope, scope.Track(section.Footers[kind]));
                    }
                }
            }
        }

        private static void ClearHeaderFooter(ComScope scope, Word.HeaderFooter item)
        {
            if (!item.Exists) return;
            item.LinkToPrevious = false;
            scope.Track(item.Range).Text = string.Empty;
        }

        private static Word.HeaderFooter PrimaryFooter(ComScope scope, Word.Sections sections, int index)
        {
            Word.Section section = scope.Track(sections[index]);
            Word.HeaderFooters footers = scope.Track(section.Footers);
            return scope.Track(footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary]);
        }

        private static bool HasPageField(ComScope scope, Word.HeaderFooter footer)
        {
            Word.Fields fields = scope.Track(scope.Track(footer.Range).Fields);
            for (int index = 1; index <= fields.Count; index++)
                if (scope.Track(fields[index]).Type == Word.WdFieldType.wdFieldPage) return true;
            return false;
        }

        private static string FooterText(Word.HeaderFooter footer)
        {
            return footer.Range.Text.Trim('\r', '\a', ' ', '\t');
        }

        private static Word.Selection RequireSelection(ComScope scope, Word.Document document)
        {
            Word.Selection selection = scope.Track(document.Application.Selection);
            if (selection == null) throw new InvalidOperationException("当前没有可用选区。");
            return selection;
        }

        private static void EnsureStyle(ComScope scope, Word.Document document, string name, string eastAsianFont, float size, bool bold, float spaceBefore, float lineSpacing, float firstLineIndent)
        {
            Word.Style style = GetStyle(scope, document, name);
            Word.Font font = scope.Track(style.Font);
            font.Name = "Times New Roman";
            font.NameFarEast = eastAsianFont;
            font.Size = size;
            font.Bold = bold ? 1 : 0;
            Word.ParagraphFormat format = scope.Track(style.ParagraphFormat);
            format.SpaceBefore = spaceBefore;
            format.SpaceAfter = 0;
            format.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
            format.LineSpacing = lineSpacing;
            format.FirstLineIndent = firstLineIndent;
            format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
            format.WidowControl = -1;
        }

        private static Word.Style GetStyle(ComScope scope, Word.Document document, string name)
        {
            Word.Styles styles = scope.Track(document.Styles);
            try { return scope.Track(styles[name]); }
            catch (COMException) { return scope.Track(styles.Add(name, Word.WdStyleType.wdStyleTypeParagraph)); }
        }

        internal static int HeadingLevel(string text)
        {
            string value = (text ?? string.Empty).Trim();
            if (Regex.IsMatch(value, @"^[一二三四五六七八九十]+[、.．]")) return 1;
            if (Regex.IsMatch(value, @"^[(（][一二三四五六七八九十]+[)）]")) return 2;
            if (Regex.IsMatch(value, @"^\d+[、.．]")) return 3;
            if (Regex.IsMatch(value, @"^[(（]\d+[)）]")) return 4;
            if (Regex.IsMatch(value, @"^第[一二三四五六七八九十百千万0-9]+[章节条](?:\s|$)")) return 1;
            return 0;
        }

        private static string StyleForLevel(int level)
        {
            switch (level)
            {
                case 1: return Heading1Style;
                case 2: return Heading2Style;
                case 3: return Heading3Style;
                default: return Heading4Style;
            }
        }

        private static float Millimeters(Word.Document document, float value)
        {
            return document.Application.MillimetersToPoints(value);
        }
    }
}
