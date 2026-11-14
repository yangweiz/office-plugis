using System;
using System.Collections.Generic;
using System.IO;
using Suite.Core.Batching;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace Suite.Word.Services
{
    internal sealed class DocumentSplitResult
    {
        internal DocumentSplitResult()
        {
            OutputPaths = new List<string>();
            SkippedItems = new List<string>();
            FailedItems = new List<string>();
        }

        public IList<string> OutputPaths { get; private set; }
        public IList<string> SkippedItems { get; private set; }
        public IList<string> FailedItems { get; private set; }
    }

    internal sealed class DocumentBatchService
    {
        public string MergeDocuments(Word.Application application, IEnumerable<string> inputPaths, string outputPath)
        {
            if (application == null) throw new ArgumentNullException("application");
            List<string> inputs = ValidateInputs(inputPaths);
            string destination = PrepareDocxPath(outputPath, Path.Combine(Path.GetDirectoryName(inputs[0]), "合并文档.docx"), "outputPath");
            foreach (string input in inputs) destination = OutputPathPolicy.EnsureDistinct(input, destination);
            destination = OutputPathPolicy.GetAvailablePath(destination);
            EnsureParentDirectory(destination);

            using (var scope = new ComScope())
            {
                Word.Document merged = null;
                bool closed = false;
                try
                {
                    Word.Documents documents = scope.Track(application.Documents);
                    merged = scope.Track(documents.Add());
                    for (int index = 0; index < inputs.Count; index++)
                    {
                        Word.Range content = scope.Track(merged.Content);
                        int end = Math.Max(0, content.End - 1);
                        Word.Range insertion = scope.Track(merged.Range(end, end));
                        if (index > 0)
                        {
                            insertion.InsertBreak(Word.WdBreakType.wdPageBreak);
                            insertion.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        }
                        insertion.InsertFile(inputs[index]);
                    }

                    merged.SaveAs2(FileName: destination, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, AddToRecentFiles: false);
                    merged.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                    closed = true;
                    EnsureNonEmptyFile(destination);
                    return destination;
                }
                catch
                {
                    if (!closed) CloseWithoutSaving(merged);
                    throw;
                }
            }
        }

        public DocumentSplitResult SplitByLevelOneHeadings(Word.Document source, string outputDirectory)
        {
            if (source == null) throw new ArgumentNullException("source");
            if (string.IsNullOrWhiteSpace(outputDirectory)) throw new ArgumentException("拆分输出目录不能为空。", "outputDirectory");

            string directory = Path.GetFullPath(outputDirectory.Trim());
            var result = new DocumentSplitResult();
            var headings = new List<HeadingBlock>();
            int contentEnd;

            using (var scope = new ComScope())
            {
                Word.Paragraphs paragraphs = scope.Track(source.Paragraphs);
                for (int index = 1; index <= paragraphs.Count; index++)
                {
                    Word.Paragraph paragraph = scope.Track(paragraphs[index]);
                    Word.Range range = scope.Track(paragraph.Range);
                    string title = CleanText(range.Text);
                    if (OfficialDocumentService.HeadingLevel(title) != 1) continue;
                    if ((bool)range.Information[Word.WdInformation.wdWithInTable])
                    {
                        result.SkippedItems.Add("表格中的候选一级标题“" + Abbreviate(title, 60) + "”未作为拆分边界。");
                        continue;
                    }
                    headings.Add(new HeadingBlock(range.Start, title));
                }

                if (headings.Count < 2)
                    throw new InvalidOperationException("拆分需要正文层至少两个公文一级标题；当前识别到 " + headings.Count + " 个。表格内标题不作为拆分边界，源文档未修改。");

                Word.Range content = scope.Track(source.Content);
                contentEnd = Math.Max(0, content.End - 1);
                if (headings[0].Start > 0)
                {
                    Word.Range leading = scope.Track(source.Range(0, headings[0].Start));
                    if (HasContent(scope, leading)) result.SkippedItems.Add("首个一级标题前的非空内容未纳入任何拆分文档。");
                }
            }

            Directory.CreateDirectory(directory);
            string savedSourcePath = string.IsNullOrWhiteSpace(source.Path) ? null : source.FullName;
            string temporaryDirectory = Path.Combine(Path.GetTempPath(), "OfficeProductivitySuite-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(temporaryDirectory);
            try
            {
                string extension = savedSourcePath == null ? ".docx" : Path.GetExtension(savedSourcePath);
                if (string.IsNullOrWhiteSpace(extension)) extension = ".docx";
                string snapshotPath = Path.Combine(temporaryDirectory, "source" + extension);
                source.SaveCopyAs(snapshotPath);
                EnsureNonEmptyFile(snapshotPath);

                using (var scope = new ComScope())
                {
                    Word.Application application = scope.Track(source.Application);
                    for (int index = 0; index < headings.Count; index++)
                    {
                        int end = index + 1 < headings.Count ? headings[index + 1].Start : contentEnd;
                        string title = Abbreviate(headings[index].Title, 60);
                        string fileName = string.Format("{0:000}_{1}.docx", index + 1, SafeFileName(headings[index].Title));
                        string requestedPath = Path.Combine(directory, fileName);
                        string destination = null;
                        try
                        {
                            destination = OutputPathPolicy.GetAvailablePath(requestedPath);
                            if (savedSourcePath != null) OutputPathPolicy.EnsureDistinct(savedSourcePath, destination);
                            CopyBlock(application, snapshotPath, headings[index].Start, end, destination);
                            result.OutputPaths.Add(destination);
                        }
                        catch (Exception ex)
                        {
                            string residual = destination != null && File.Exists(destination)
                                ? "；请检查可能的残留文件：" + destination
                                : string.Empty;
                            result.FailedItems.Add("第 " + (index + 1) + " 项“" + title + "”生成失败：" + ex.Message + residual);
                        }
                    }
                }
                return result;
            }
            finally
            {
                TryDeleteDirectory(temporaryDirectory);
            }
        }

        public string ExportPdf(Word.Document source, string requestedOutputPath)
        {
            if (source == null) throw new ArgumentNullException("source");
            using (var scope = new ComScope())
            {
                Word.Range content = scope.Track(source.Content);
                if (!HasContent(scope, content)) throw new InvalidOperationException("源文档没有可导出的正文、图片或形状内容。");
            }

            string defaultPath = null;
            if (!string.IsNullOrWhiteSpace(source.Path))
                defaultPath = Path.Combine(source.Path, Path.GetFileNameWithoutExtension(source.Name) + ".pdf");
            string destination = PreparePathWithExtension(requestedOutputPath, defaultPath, ".pdf", "requestedOutputPath");
            if (!string.IsNullOrWhiteSpace(source.Path)) OutputPathPolicy.EnsureDistinct(source.FullName, destination);
            destination = OutputPathPolicy.GetAvailablePath(destination);
            EnsureParentDirectory(destination);

            source.ExportAsFixedFormat(
                OutputFileName: destination,
                ExportFormat: Word.WdExportFormat.wdExportFormatPDF,
                OpenAfterExport: false,
                OptimizeFor: Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                Range: Word.WdExportRange.wdExportAllDocument,
                Item: Word.WdExportItem.wdExportDocumentContent);
            EnsureNonEmptyFile(destination);
            return destination;
        }

        private static List<string> ValidateInputs(IEnumerable<string> inputPaths)
        {
            if (inputPaths == null) throw new ArgumentNullException("inputPaths");
            var inputs = new List<string>();
            foreach (string path in inputPaths)
            {
                if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("合并输入路径不能包含空项。", "inputPaths");
                string fullPath = Path.GetFullPath(path.Trim());
                string extension = Path.GetExtension(fullPath);
                if (!string.Equals(extension, ".docx", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(extension, ".docm", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(extension, ".dotx", StringComparison.OrdinalIgnoreCase))
                    throw new ArgumentException("仅支持合并 DOCX、DOCM、DOTX 文件：" + fullPath, "inputPaths");
                if (!File.Exists(fullPath)) throw new FileNotFoundException("合并输入文件不存在。", fullPath);
                inputs.Add(fullPath);
            }
            if (inputs.Count < 2) throw new ArgumentException("合并文档至少需要两个现有输入文件。", "inputPaths");
            return inputs;
        }

        private static void CopyBlock(Word.Application application, string snapshotPath, int start, int end, string destination)
        {
            using (var scope = new ComScope())
            {
                Word.Document target = null;
                bool closed = false;
                bool saveStarted = false;
                try
                {
                    Word.Documents documents = scope.Track(application.Documents);
                    Office.MsoAutomationSecurity previousSecurity = application.AutomationSecurity;
                    try
                    {
                        application.AutomationSecurity = Office.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
                        target = scope.Track(documents.Open(
                            FileName: snapshotPath,
                            ConfirmConversions: false,
                            ReadOnly: true,
                            AddToRecentFiles: false,
                            Visible: false));
                    }
                    finally
                    {
                        application.AutomationSecurity = previousSecurity;
                    }

                    Word.Range content = scope.Track(target.Content);
                    int targetContentEnd = Math.Max(0, content.End - 1);
                    if (end < targetContentEnd)
                        scope.Track(target.Range(end, targetContentEnd)).Delete();
                    if (start > 0)
                        scope.Track(target.Range(0, start)).Delete();

                    Word.WdAlertLevel previousAlerts = application.DisplayAlerts;
                    try
                    {
                        application.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                        saveStarted = true;
                        target.SaveAs2(FileName: destination, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument, AddToRecentFiles: false);
                    }
                    finally
                    {
                        application.DisplayAlerts = previousAlerts;
                    }
                    EnsureNonEmptyFile(destination);
                    CloseWithoutSaving(target);
                    closed = true;
                }
                catch
                {
                    if (!closed) CloseWithoutSaving(target);
                    if (saveStarted) TryDeleteFile(destination);
                    throw;
                }
            }
        }

        private static bool HasContent(ComScope scope, Word.Range range)
        {
            string text = range.Text ?? string.Empty;
            for (int index = 0; index < text.Length; index++)
            {
                char value = text[index];
                if (value != '\a' && value != '\f' && !char.IsWhiteSpace(value)) return true;
            }
            Word.InlineShapes inlineShapes = scope.Track(range.InlineShapes);
            if (inlineShapes.Count > 0) return true;
            Word.Shapes shapes = scope.Track(range.Shapes);
            return shapes.Count > 0;
        }

        private static string CleanText(string text)
        {
            return (text ?? string.Empty).Trim('\r', '\a', '\f', ' ', '\t', '\u3000');
        }

        private static string SafeFileName(string title)
        {
            string value = CleanText(title);
            char[] invalid = Path.GetInvalidFileNameChars();
            var characters = new char[value.Length];
            for (int index = 0; index < value.Length; index++)
                characters[index] = Array.IndexOf(invalid, value[index]) >= 0 || char.IsControl(value[index]) ? '_' : value[index];
            string safe = new string(characters).Trim().TrimEnd('.', ' ');
            if (safe.Length == 0) safe = "未命名一级标题";
            if (IsReservedDeviceName(safe)) safe = "_" + safe;
            return Abbreviate(safe, 60);
        }

        private static bool IsReservedDeviceName(string value)
        {
            string name = value;
            int dot = name.IndexOf('.');
            if (dot >= 0) name = name.Substring(0, dot);
            if (string.Equals(name, "CON", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "PRN", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "AUX", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "NUL", StringComparison.OrdinalIgnoreCase)) return true;
            if (name.Length != 4) return false;
            bool device = name.StartsWith("COM", StringComparison.OrdinalIgnoreCase) || name.StartsWith("LPT", StringComparison.OrdinalIgnoreCase);
            return device && name[3] >= '1' && name[3] <= '9';
        }

        private static string Abbreviate(string value, int maximumLength)
        {
            return value.Length <= maximumLength ? value : value.Substring(0, maximumLength);
        }

        private static string PrepareDocxPath(string requestedPath, string defaultPath, string parameterName)
        {
            return PreparePathWithExtension(requestedPath, defaultPath, ".docx", parameterName);
        }

        private static string PreparePathWithExtension(string requestedPath, string defaultPath, string extension, string parameterName)
        {
            string path = string.IsNullOrWhiteSpace(requestedPath) ? defaultPath : requestedPath.Trim();
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("输出路径不能为空；未保存的文档必须显式指定输出路径。", parameterName);
            string currentExtension = Path.GetExtension(path);
            if (string.IsNullOrEmpty(currentExtension)) path += extension;
            else if (!string.Equals(currentExtension, extension, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("输出文件必须使用 " + extension.ToUpperInvariant() + " 扩展名。", parameterName);
            return Path.GetFullPath(path);
        }

        private static void EnsureParentDirectory(string path)
        {
            string directory = Path.GetDirectoryName(path);
            if (string.IsNullOrWhiteSpace(directory)) throw new ArgumentException("无法确定输出目录：" + path);
            Directory.CreateDirectory(directory);
        }

        private static void EnsureNonEmptyFile(string path)
        {
            if (!File.Exists(path) || new FileInfo(path).Length == 0) throw new IOException("输出文件未生成或内容为空：" + path);
        }

        private static void TryDeleteFile(string path)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(path) && File.Exists(path)) File.Delete(path);
            }
            catch { }
        }

        private static void TryDeleteDirectory(string path)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(path) && Directory.Exists(path)) Directory.Delete(path, true);
            }
            catch { }
        }

        private static void CloseWithoutSaving(Word.Document document)
        {
            if (document == null) return;
            try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); }
            catch { }
        }

        private sealed class HeadingBlock
        {
            internal HeadingBlock(int start, string title)
            {
                Start = start;
                Title = title;
            }

            internal int Start { get; private set; }
            internal string Title { get; private set; }
        }
    }
}
