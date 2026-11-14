using System;
using System.Collections.Generic;
using System.IO;
using Suite.Core.Batching;
using Word = Microsoft.Office.Interop.Word;

namespace Suite.Word.Services
{
    internal sealed class TableAndLabelService
    {
        public void AutoFitSelectionTableToWindow(Word.Document document)
        {
            using (var scope = new ComScope())
            {
                Word.Table table = RequireSingleTable(scope, document);
                table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);
            }
        }

        public void DistributeSelectionRows(Word.Document document)
        {
            using (var scope = new ComScope())
            {
                Word.Table table = RequireSingleTable(scope, document);
                Word.Rows rows = scope.Track(table.Rows);
                rows.Distribute();
            }
        }

        public void DistributeSelectionColumns(Word.Document document)
        {
            using (var scope = new ComScope())
            {
                Word.Table table = RequireSingleTable(scope, document);
                Word.Columns columns = scope.Track(table.Columns);
                columns.Distribute();
            }
        }

        public string CreateSeatLabels(Word.Application application, IEnumerable<string> names, int columns, string outputPath)
        {
            if (application == null) throw new ArgumentNullException("application");
            if (columns < 1 || columns > 6) throw new ArgumentOutOfRangeException("columns", "桌签列数必须为 1 至 6 列。");
            if (names == null) throw new ArgumentNullException("names");

            var cleanedNames = new List<string>();
            foreach (string name in names)
            {
                if (string.IsNullOrWhiteSpace(name)) continue;
                cleanedNames.Add(name.Trim());
            }
            if (cleanedNames.Count == 0) throw new ArgumentException("至少需要一个非空姓名。", "names");
            if (cleanedNames.Count > 500) throw new ArgumentException("一次最多生成 500 个桌签，请分批处理。", "names");

            string destination = PrepareDocxPath(outputPath);
            destination = OutputPathPolicy.GetAvailablePath(destination);
            string directory = Path.GetDirectoryName(destination);
            if (string.IsNullOrWhiteSpace(directory)) throw new ArgumentException("无法确定桌签输出目录。", "outputPath");
            Directory.CreateDirectory(directory);

            using (var scope = new ComScope())
            {
                Word.Document document = null;
                bool closed = false;
                try
                {
                    Word.Documents documents = scope.Track(application.Documents);
                    document = scope.Track(documents.Add());
                    Word.Range insertion = scope.Track(document.Range(0, 0));
                    Word.Tables tables = scope.Track(document.Tables);
                    int rows = (cleanedNames.Count + columns - 1) / columns;
                    Word.Table table = scope.Track(tables.Add(insertion, rows, columns));
                    table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);
                    Word.Columns tableColumns = scope.Track(table.Columns);
                    tableColumns.Distribute();
                    Word.Borders borders = scope.Track(table.Borders);
                    borders.Enable = 1;

                    for (int index = 0; index < cleanedNames.Count; index++)
                    {
                        Word.Cell cell = scope.Track(table.Cell(index / columns + 1, index % columns + 1));
                        cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellVerticalAlignmentCenter;
                        Word.Range range = scope.Track(cell.Range);
                        range.Text = cleanedNames[index];
                        Word.ParagraphFormat paragraph = scope.Track(range.ParagraphFormat);
                        paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        Word.Font font = scope.Track(range.Font);
                        font.Name = "Times New Roman";
                        font.NameFarEast = "黑体";
                        font.Size = SuitableFontSize(cleanedNames[index], columns);
                    }

                    document.SaveAs2(FileName: destination, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, AddToRecentFiles: false);
                    document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                    closed = true;
                    if (!File.Exists(destination) || new FileInfo(destination).Length == 0)
                        throw new IOException("桌签输出文件未生成或内容为空：" + destination);
                    return destination;
                }
                catch
                {
                    if (!closed) CloseWithoutSaving(document);
                    throw;
                }
            }
        }

        private static Word.Table RequireSingleTable(ComScope scope, Word.Document document)
        {
            if (document == null) throw new ArgumentNullException("document");
            Word.Application application = scope.Track(document.Application);
            Word.Selection selection = scope.Track(application.Selection);
            if (selection == null) throw new InvalidOperationException("当前没有可用选区，请先选择或进入一个 Word 表格。");
            Word.Tables tables = scope.Track(selection.Tables);
            if (tables.Count != 1) throw new InvalidOperationException("请只选择一个 Word 表格，或将插入点放在单个表格内；当前命令不会处理多个表格或表外选区。");
            return scope.Track(tables[1]);
        }

        private static float SuitableFontSize(string name, int columns)
        {
            float size = columns == 1 ? 32 : columns == 2 ? 28 : columns <= 4 ? 22 : 18;
            if (name.Length > 4) size = Math.Max(10f, size * 4f / name.Length);
            return (float)Math.Round(size * 2f, MidpointRounding.AwayFromZero) / 2f;
        }

        private static string PrepareDocxPath(string outputPath)
        {
            if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("桌签输出路径不能为空。", "outputPath");
            string path = outputPath.Trim();
            string extension = Path.GetExtension(path);
            if (string.IsNullOrEmpty(extension)) path += ".docx";
            else if (!string.Equals(extension, ".docx", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("桌签默认保存为 DOCX，请使用 .docx 扩展名。", "outputPath");
            return Path.GetFullPath(path);
        }

        private static void CloseWithoutSaving(Word.Document document)
        {
            if (document == null) return;
            try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); }
            catch { }
        }
    }
}
