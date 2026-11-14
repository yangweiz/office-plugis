using System;
using Word = Microsoft.Office.Interop.Word;

namespace Suite.Word.Services
{
    internal sealed class OfficialDocumentPartService
    {
        public void InsertTextRedHead(Word.Document document, string agencyName)
        {
            string text = (agencyName ?? string.Empty).Trim();
            if (text.Length == 0 || text.Length > 120) throw new ArgumentException("红头文字应为 1 至 120 个字符。", "agencyName");
            using (var scope = new ComScope())
            {
                Word.Selection selection = scope.Track(document.Application.Selection);
                if (selection == null) throw new InvalidOperationException("请先选择红头插入位置。");
                Word.Range insertion = scope.Track(selection.Range);
                insertion.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                int start = insertion.Start;
                insertion.InsertAfter(text + "\r\r");

                Word.Range title = scope.Track(document.Range(start, start + text.Length));
                title.Font.Name = "Times New Roman";
                title.Font.NameFarEast = "宋体";
                title.Font.Size = 36;
                title.Font.Bold = 1;
                title.Font.Color = Word.WdColor.wdColorRed;
                title.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                title.ParagraphFormat.SpaceAfter = 8;

                int lineStart = start + text.Length + 1;
                Word.Range line = scope.Track(document.Range(lineStart, lineStart + 1));
                line.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                Word.Border border = scope.Track(line.Borders[Word.WdBorderType.wdBorderBottom]);
                border.LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                border.LineWidth = Word.WdLineWidth.wdLineWidth225pt;
                border.Color = Word.WdColor.wdColorRed;
                line.ParagraphFormat.SpaceAfter = 16;
            }
        }

        public void InsertOfficialDate(Word.Document document)
        {
            Word.Selection selection = document.Application.Selection;
            if (selection == null) throw new InvalidOperationException("请先选择落款日期插入位置。");
            Word.Range range = selection.Range;
            try
            {
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.InsertAfter(DateTime.Now.ToString("yyyy年M月d日"));
                range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            }
            finally
            {
                if (System.Runtime.InteropServices.Marshal.IsComObject(range)) System.Runtime.InteropServices.Marshal.FinalReleaseComObject(range);
                if (System.Runtime.InteropServices.Marshal.IsComObject(selection)) System.Runtime.InteropServices.Marshal.FinalReleaseComObject(selection);
            }
        }
    }
}
