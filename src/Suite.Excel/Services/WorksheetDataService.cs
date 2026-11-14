using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Suite.Core.Batching;
using Suite.Core.Data;
using Suite.Core.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace Suite.Excel.Services
{
    public sealed class WorksheetDataService
    {
        private const int MaxCells = 200000;
        private const int MaxGroups = 200;
        private const string OutputMarkerName = "OfficeProductivitySuite.Output";
        private const int NoCellsFound = unchecked((int)0x800A03EC);
        private readonly Excel.Application _application;

        public WorksheetDataService(Excel.Application application)
        {
            if (application == null) throw new ArgumentNullException("application");
            _application = application;
        }

        public bool HasActiveWorkbook()
        {
            Excel.Workbook workbook = null;
            try
            {
                workbook = _application.ActiveWorkbook;
                return workbook != null;
            }
            finally
            {
                ReleaseComObject(workbook);
            }
        }

        public Excel.Range GetSelectionRange()
        {
            object selected = _application.Selection;
            Excel.Range range = selected as Excel.Range;
            if (range == null)
            {
                ReleaseComObject(selected);
                throw new InvalidOperationException("当前选择不是单元格区域，请先选择连续的表格区域。");
            }
            try
            {
                EnsureManageable(range);
                return range;
            }
            catch
            {
                ReleaseComObject(range);
                throw;
            }
        }

        public string CleanupText()
        {
            Excel.Range selection = GetSelectionRange();
            Excel.Range textCells = null;
            Excel.Areas areas = null;
            int changed = 0;
            try
            {
                try
                {
                    textCells = selection.SpecialCells(Excel.XlCellType.xlCellTypeConstants, Excel.XlSpecialCellsValue.xlTextValues);
                }
                catch (COMException ex) when (ex.ErrorCode == NoCellsFound)
                {
                    return "选区中没有可清理的文本常量；数字、日期和公式均未改动。";
                }

                areas = textCells.Areas;
                for (int index = 1; index <= areas.Count; index++)
                {
                    Excel.Range area = null;
                    try
                    {
                        area = (Excel.Range)areas[index];
                        object[,] values = ReadMatrix(area);
                        bool areaChanged = false;
                        for (int row = 1; row <= values.GetLength(0); row++)
                        {
                            for (int column = 1; column <= values.GetLength(1); column++)
                            {
                                string original = values[row, column] as string;
                                if (original == null) continue;
                                string cleaned = TextCleanup.RemoveInvisibleCharacters(original);
                                cleaned = TextCleanup.RemoveLeadingSpaces(cleaned);
                                cleaned = TextCleanup.CollapseBlankLines(cleaned);
                                cleaned = TextCleanup.ReflowBrokenLines(cleaned);
                                if (string.Equals(original, cleaned, StringComparison.Ordinal)) continue;
                                values[row, column] = cleaned;
                                changed++;
                                areaChanged = true;
                            }
                        }
                        if (areaChanged)
                        {
                            area.NumberFormat = "@";
                            area.Value2 = values;
                        }
                    }
                    finally
                    {
                        ReleaseComObject(area);
                    }
                }
                return changed == 0
                    ? "文本检查完成，无需修改；数字、日期和公式均未改动。"
                    : string.Format("已清理 {0} 个文本单元格；数字、日期和公式均未改动。", changed);
            }
            finally
            {
                ReleaseComObject(areas);
                ReleaseComObject(textCells);
                ReleaseComObject(selection);
            }
        }

        public string RemoveBlankRows()
        {
            object activeSheet = _application.ActiveSheet;
            Excel.Worksheet worksheet = activeSheet as Excel.Worksheet;
            if (worksheet == null)
            {
                ReleaseComObject(activeSheet);
                throw new InvalidOperationException("当前没有活动工作表。");
            }
            Excel.Range usedRange = null;
            bool oldScreenUpdating = _application.ScreenUpdating;
            bool oldEnableEvents = _application.EnableEvents;
            try
            {
                usedRange = worksheet.UsedRange;
                EnsureManageable(usedRange);
                EnsureNoFormulas(usedRange, "删除空行");
                List<IList<string>> rows = ReadRows(usedRange);
                List<IList<string>> retained = TabularData.RemoveBlankRows(rows);
                if (retained.Count == 0) return "已用区域没有数据，未删除任何工作表行。";
                int blankCount = rows.Count - retained.Count;
                if (blankCount == 0) return "已用区域中没有空行，工作表未修改。";

                var blankRowIndexes = new List<int>();
                for (int index = 0; index < rows.Count; index++)
                {
                    if (TabularData.IsBlankRow(rows[index])) blankRowIndexes.Add(usedRange.Row + index);
                }

                _application.ScreenUpdating = false;
                _application.EnableEvents = false;
                for (int index = blankRowIndexes.Count - 1; index >= 0; index--)
                {
                    DeleteEntireRow(worksheet, blankRowIndexes[index]);
                }
                return string.Format("已从已用区域删除 {0} 个空行；含公式的区域会被阻止处理。", blankCount);
            }
            finally
            {
                _application.ScreenUpdating = oldScreenUpdating;
                _application.EnableEvents = oldEnableEvents;
                ReleaseComObject(usedRange);
                ReleaseComObject(worksheet);
            }
        }

        public string RemoveDuplicateRows()
        {
            Excel.Range selection = GetSelectionRange();
            Excel.Worksheet worksheet = null;
            bool oldScreenUpdating = _application.ScreenUpdating;
            bool oldEnableEvents = _application.EnableEvents;
            try
            {
                EnsureNoFormulas(selection, "选区去重");
                List<IList<string>> rows = ReadRows(selection);
                if (rows.Count < 2) return "选区不足两行，无需去重。";
                List<IList<string>> uniqueRows = TabularData.DistinctRows(rows, null, false);
                if (uniqueRows.Count == rows.Count) return "选区没有重复行，工作表未修改。";

                var duplicates = new List<int>();
                int uniqueIndex = 0;
                for (int index = 0; index < rows.Count; index++)
                {
                    if (uniqueIndex < uniqueRows.Count && RowsEqual(rows[index], uniqueRows[uniqueIndex]))
                    {
                        uniqueIndex++;
                    }
                    else
                    {
                        duplicates.Add(index);
                    }
                }

                worksheet = selection.Worksheet;
                int firstRow = selection.Row;
                int firstColumn = selection.Column;
                int columnCount = ColumnCount(selection);
                _application.ScreenUpdating = false;
                _application.EnableEvents = false;
                for (int index = duplicates.Count - 1; index >= 0; index--)
                {
                    DeleteSelectionRow(worksheet, firstRow + duplicates[index], firstColumn, columnCount);
                }
                return string.Format("已按整行内容忽略大小写去重，删除 {0} 行并保留首次出现的数据。", duplicates.Count);
            }
            finally
            {
                _application.ScreenUpdating = oldScreenUpdating;
                _application.EnableEvents = oldEnableEvents;
                ReleaseComObject(worksheet);
                ReleaseComObject(selection);
            }
        }

        public string SortSelection(bool hasHeader)
        {
            Excel.Range selection = GetSelectionRange();
            Excel.Range key = null;
            Excel.Range sortColumns = null;
            bool oldScreenUpdating = _application.ScreenUpdating;
            try
            {
                if (RowCount(selection) < 2) return "选区不足两行，无需排序。";
                if (hasHeader && RowCount(selection) < 3) return "含表头的选区不足两行数据，无需排序。";
                sortColumns = selection.Columns;
                key = (Excel.Range)sortColumns[1];
                _application.ScreenUpdating = false;
                selection.Sort(
                    Key1: key,
                    Order1: Excel.XlSortOrder.xlAscending,
                    Header: hasHeader ? Excel.XlYesNoGuess.xlYes : Excel.XlYesNoGuess.xlNo,
                    Orientation: Excel.XlSortOrientation.xlSortRows);
                return string.Format("已按选区首列升序排序（{0}）。", hasHeader ? "首行为表头" : "无表头");
            }
            finally
            {
                _application.ScreenUpdating = oldScreenUpdating;
                ReleaseComObject(key);
                ReleaseComObject(sortColumns);
                ReleaseComObject(selection);
            }
        }

        public string CompareByKey(Excel.Range left, Excel.Range right, int keyColumn, bool hasHeader)
        {
            if (left == null) throw new ArgumentNullException("left");
            if (right == null) throw new ArgumentNullException("right");
            EnsureManageable(left);
            EnsureManageable(right);
            int leftColumns = ColumnCount(left);
            int rightColumns = ColumnCount(right);
            if (keyColumn < 1 || keyColumn > leftColumns || keyColumn > rightColumns)
                throw new ArgumentOutOfRangeException("keyColumn", "比较键列号必须同时存在于左右区域，且从 1 开始。");

            List<IList<string>> leftRows = ReadRows(left);
            List<IList<string>> rightRows = ReadRows(right);
            if (hasHeader && (leftRows.Count < 2 || rightRows.Count < 2))
                throw new InvalidOperationException("含表头的左右区域都需要至少一行数据。");

            List<IList<string>> leftData = leftRows.Skip(hasHeader ? 1 : 0).ToList();
            List<IList<string>> rightData = rightRows.Skip(hasHeader ? 1 : 0).ToList();
            List<DifferenceRow> differences = TabularData.Compare(leftData, rightData, keyColumn - 1, false);
            var output = new List<IList<string>>();
            output.Add(BuildCompareHeader(leftRows, rightRows, leftColumns, rightColumns, hasHeader));
            foreach (DifferenceRow difference in differences)
            {
                var row = new List<string> { DifferenceText(difference.Kind), difference.Key };
                AddPadded(row, difference.Left, leftColumns);
                AddPadded(row, difference.Right, rightColumns);
                output.Add(row);
            }

            Excel.Workbook workbook = GetWorkbook(left);
            Excel.Worksheet resultSheet = null;
            try
            {
                resultSheet = AddOutputSheet(workbook, "按键对比结果");
                WriteRows(resultSheet, output);
                return string.Format("已在工作表“{0}”以文本快照输出 {1} 项差异；原区域未修改，结果尚未保存。", resultSheet.Name, differences.Count);
            }
            finally
            {
                ReleaseComObject(resultSheet);
                ReleaseComObject(workbook);
            }
        }

        public string SummarizeWorksheets(bool hasHeader)
        {
            Excel.Workbook workbook = _application.ActiveWorkbook;
            if (workbook == null) throw new InvalidOperationException("当前没有活动工作簿。");
            Excel.Sheets sheets = null;
            Excel.Worksheet resultSheet = null;
            var collected = new List<CollectedSheet>();
            IList<string> header = null;
            int maximumColumns = 0;
            long totalRows = 0;
            try
            {
                sheets = workbook.Worksheets;
                for (int index = 1; index <= sheets.Count; index++)
                {
                    Excel.Worksheet worksheet = null;
                    Excel.Range usedRange = null;
                    try
                    {
                        object sheetItem = sheets[index];
                        worksheet = sheetItem as Excel.Worksheet;
                        if (worksheet == null)
                        {
                            ReleaseComObject(sheetItem);
                            continue;
                        }
                        if (IsGeneratedOutputSheet(worksheet)) continue;
                        usedRange = worksheet.UsedRange;
                        EnsureManageable(usedRange);
                        List<IList<string>> rows = ReadRows(usedRange);
                        List<IList<object>> values = ReadValueRows(usedRange);
                        var retainedIndexes = new List<int>();
                        for (int row = 0; row < rows.Count; row++)
                            if (!TabularData.IsBlankRow(rows[row])) retainedIndexes.Add(row);
                        if (retainedIndexes.Count == 0) continue;

                        foreach (int row in retainedIndexes) maximumColumns = Math.Max(maximumColumns, rows[row].Count);
                        int dataStart = 0;
                        if (hasHeader)
                        {
                            int headerIndex = retainedIndexes[0];
                            if (header == null) header = new List<string>(rows[headerIndex]);
                            else if (!HeadersEqual(header, rows[headerIndex]))
                                throw new InvalidOperationException("工作表“" + worksheet.Name + "”的表头与首个工作表不一致，已停止汇总。");
                            dataStart = 1;
                        }
                        List<IList<object>> data = retainedIndexes.Skip(dataStart).Select(row => values[row]).ToList();
                        if (data.Count == 0) continue;
                        totalRows += data.Count;
                        if (totalRows * (maximumColumns + 1L) > MaxCells)
                            throw new InvalidOperationException("汇总结果超过 200000 个单元格，请分批处理。");
                        collected.Add(new CollectedSheet(worksheet.Name, data));
                    }
                    finally
                    {
                        ReleaseComObject(usedRange);
                        ReleaseComObject(worksheet);
                    }
                }

                if (collected.Count == 0) return "当前工作簿没有可汇总的数据行，未创建结果表。";
                var output = new List<IList<object>>();
                var outputHeader = new List<object> { "来源工作表" };
                if (hasHeader) foreach (string value in header) outputHeader.Add(value);
                else for (int column = 1; column <= maximumColumns; column++) outputHeader.Add("列" + column);
                output.Add(outputHeader);
                foreach (CollectedSheet sheet in collected)
                {
                    foreach (IList<object> sourceRow in sheet.Rows)
                    {
                        var row = new List<object> { sheet.Name };
                        AddPaddedValues(row, sourceRow, maximumColumns);
                        output.Add(row);
                    }
                }

                resultSheet = AddOutputSheet(workbook, "多表汇总结果", true);
                WriteValues(resultSheet, output);
                return string.Format("已汇总 {0} 个工作表、{1} 行数据到“{2}”；已按原始值写入（不复制原格式），原工作表未修改，结果尚未保存。", collected.Count, totalRows, resultSheet.Name);
            }
            finally
            {
                ReleaseComObject(resultSheet);
                ReleaseComObject(sheets);
                ReleaseComObject(workbook);
            }
        }

        public string SplitSelectionToNewWorkbook(int keyColumn, out string outputPath)
        {
            Excel.Range selection = GetSelectionRange();
            Excel.Workbook outputWorkbook = null;
            Excel.Workbooks workbooks = null;
            bool oldDisplayAlerts = _application.DisplayAlerts;
            outputPath = null;
            try
            {
                List<IList<string>> rows = ReadRows(selection);
                List<IList<object>> values = ReadValueRows(selection);
                if (rows.Count < 2) throw new InvalidOperationException("拆分需要一行表头和至少一行数据。");
                int columns = ColumnCount(selection);
                if (keyColumn < 1 || keyColumn > columns)
                    throw new ArgumentOutOfRangeException("keyColumn", "拆分列号超出选区范围，列号从 1 开始。");
                List<RowGroup> groups = GroupRows(rows, values, keyColumn - 1);

                Excel.Workbook sourceWorkbook = GetWorkbook(selection);
                string sourceFullName;
                string sourceDirectory;
                string sourceName;
                bool sourceIsSaved;
                try
                {
                    sourceFullName = sourceWorkbook.FullName;
                    sourceDirectory = sourceWorkbook.Path;
                    sourceIsSaved = !string.IsNullOrWhiteSpace(sourceDirectory);
                    sourceName = Path.GetFileNameWithoutExtension(sourceWorkbook.Name);
                }
                finally
                {
                    ReleaseComObject(sourceWorkbook);
                }

                if (string.IsNullOrWhiteSpace(sourceDirectory))
                    sourceDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                if (string.IsNullOrWhiteSpace(sourceDirectory))
                    throw new InvalidOperationException("源工作簿尚未保存，且无法确定“我的文档”目录；请先保存源工作簿。");
                if (string.IsNullOrWhiteSpace(sourceName)) sourceName = "拆分输出";
                string groupName = SanitizeFileName(rows[0][keyColumn - 1]);
                string requestedPath = Path.Combine(sourceDirectory, sourceName + "_按" + groupName + "拆分.xlsx");
                outputPath = sourceIsSaved
                    ? OutputPathPolicy.EnsureDistinct(sourceFullName, requestedPath)
                    : requestedPath;
                outputPath = OutputPathPolicy.GetAvailablePath(outputPath);

                workbooks = _application.Workbooks;
                outputWorkbook = workbooks.Add();
                Excel.Sheets initialSheets = null;
                Excel.Worksheet firstSheet = null;
                int initialCount;
                try
                {
                    initialSheets = outputWorkbook.Worksheets;
                    initialCount = initialSheets.Count;
                    firstSheet = (Excel.Worksheet)initialSheets[1];
                }
                finally
                {
                    ReleaseComObject(firstSheet);
                    ReleaseComObject(initialSheets);
                }

                Excel.Sheets sheets = null;
                Excel.Worksheet target = null;
                try
                {
                    sheets = outputWorkbook.Worksheets;
                    target = (Excel.Worksheet)sheets[1];
                    target.Name = GetUniqueSheetName(outputWorkbook, SanitizeSheetName(groups[0].Key));
                    WriteValues(target, WithHeaderValues(values[0], groups[0].Values));
                }
                finally
                {
                    ReleaseComObject(target);
                    ReleaseComObject(sheets);
                }

                for (int index = 1; index < groups.Count; index++)
                {
                    Excel.Worksheet groupSheet = null;
                    try
                    {
                        groupSheet = AddOutputSheet(outputWorkbook, SanitizeSheetName(groups[index].Key));
                        WriteValues(groupSheet, WithHeaderValues(values[0], groups[index].Values));
                    }
                    finally
                    {
                        ReleaseComObject(groupSheet);
                    }
                }

                if (initialCount > 1)
                {
                    _application.DisplayAlerts = false;
                    for (int index = initialCount; index >= 2; index--)
                    {
                        Excel.Sheets remaining = null;
                        Excel.Worksheet extra = null;
                        try
                        {
                            remaining = outputWorkbook.Worksheets;
                            extra = remaining[index] as Excel.Worksheet;
                            if (extra != null) extra.Delete();
                        }
                        finally
                        {
                            ReleaseComObject(extra);
                            ReleaseComObject(remaining);
                        }
                    }
                    _application.DisplayAlerts = oldDisplayAlerts;
                }

                outputWorkbook.SaveAs(
                    Filename: outputPath,
                    FileFormat: Excel.XlFileFormat.xlOpenXMLWorkbook,
                    CreateBackup: false);
                return string.Format("已按“{0}”拆分为 {1} 个工作表并保存到：{2}。已按原始值写入（不复制原格式），源工作簿未被覆盖。", rows[0][keyColumn - 1], groups.Count, outputPath);
            }
            catch
            {
                if (outputWorkbook != null) outputWorkbook.Close(false, Type.Missing, Type.Missing);
                throw;
            }
            finally
            {
                _application.DisplayAlerts = oldDisplayAlerts;
                ReleaseComObject(workbooks);
                ReleaseComObject(outputWorkbook);
                ReleaseComObject(selection);
            }
        }

        public string CreatePayroll(int blankRows)
        {
            Excel.Range selection = GetSelectionRange();
            Excel.Workbook workbook = null;
            Excel.Worksheet resultSheet = null;
            try
            {
                List<IList<string>> rows = ReadRows(selection);
                List<IList<object>> values = ReadValueRows(selection);
                if (rows.Count < 2) throw new InvalidOperationException("工资条需要一行表头和至少一行数据。");
                if (blankRows < 0 || blankRows > 20) throw new ArgumentOutOfRangeException("blankRows");
                var payroll = new List<IList<object>>();
                for (int index = 1; index < rows.Count; index++)
                {
                    if (TabularData.IsBlankRow(rows[index])) continue;
                    payroll.Add(new List<object>(values[0]));
                    payroll.Add(new List<object>(values[index]));
                    for (int blank = 0; blank < blankRows; blank++) payroll.Add(new object[values[0].Count]);
                }
                if (payroll.Count == 0) return "选区没有可生成工资条的数据行。";
                workbook = GetWorkbook(selection);
                resultSheet = AddOutputSheet(workbook, "工资条");
                WriteValues(resultSheet, payroll);
                return string.Format("已在“{0}”生成 {1} 行工资条，数据行间保留 {2} 个空行；已按原始值写入（不复制原格式），原选区未修改，结果尚未保存。", resultSheet.Name, payroll.Count, blankRows);
            }
            finally
            {
                ReleaseComObject(resultSheet);
                ReleaseComObject(workbook);
                ReleaseComObject(selection);
            }
        }

        public string SummarizeSelectionByColumn(int keyColumn, int sumColumn)
        {
            Excel.Range selection = GetSelectionRange();
            Excel.Workbook workbook = null;
            Excel.Worksheet resultSheet = null;
            try
            {
                List<IList<string>> rows = ReadRows(selection);
                List<IList<object>> values = ReadValueRows(selection);
                if (rows.Count < 2) throw new InvalidOperationException("分组求和需要一行表头和至少一行数据。");
                int columns = ColumnCount(selection);
                if (keyColumn < 1 || keyColumn > columns || sumColumn < 1 || sumColumn > columns)
                    throw new ArgumentOutOfRangeException("keyColumn", "分组键列和求和列必须位于选区内，列号从 1 开始。");

                var dataRows = new List<IList<string>>();
                for (int index = 1; index < rows.Count; index++)
                {
                    if (TabularData.IsBlankRow(rows[index])) continue;
                    if (string.IsNullOrWhiteSpace(rows[index][keyColumn - 1]))
                        throw new InvalidOperationException("第 " + (index + 1) + " 行的分组键为空，请先补全数据。");
                    object raw = values[index][sumColumn - 1];
                    if (raw == null || (raw is string && string.IsNullOrWhiteSpace((string)raw)))
                        throw new InvalidOperationException("第 " + (index + 1) + " 行的求和列为空，请先补全或剔除该行。");
                    if (raw is DateTime || raw is bool)
                        throw new InvalidOperationException("第 " + (index + 1) + " 行的求和列不是数值。");

                    decimal numericValue;
                    string text = raw as string;
                    if (text != null)
                    {
                        const NumberStyles styles = NumberStyles.Float | NumberStyles.AllowThousands;
                        if (!decimal.TryParse(text, styles, CultureInfo.CurrentCulture, out numericValue)
                            && !decimal.TryParse(text, styles, CultureInfo.InvariantCulture, out numericValue))
                            throw new InvalidOperationException("第 " + (index + 1) + " 行的求和列包含非数值文本：" + text);
                    }
                    else
                    {
                        try { numericValue = Convert.ToDecimal(raw, CultureInfo.CurrentCulture); }
                        catch (Exception ex) when (ex is FormatException || ex is InvalidCastException || ex is OverflowException)
                        {
                            throw new InvalidOperationException("第 " + (index + 1) + " 行的求和列不是有效数值。", ex);
                        }
                    }

                    var normalized = new List<string>(rows[index]);
                    normalized[sumColumn - 1] = numericValue.ToString(CultureInfo.InvariantCulture);
                    dataRows.Add(normalized);
                }
                if (dataRows.Count == 0) throw new InvalidOperationException("选区没有可汇总的数据行。");

                IReadOnlyList<WorksheetGroup> groups = WorksheetGrouping.GroupByColumns(dataRows, new[] { keyColumn - 1 });
                if (groups.Count > MaxGroups) throw new InvalidOperationException("分组结果超过 200 组，请分批处理。");
                string keyHeader = string.IsNullOrWhiteSpace(rows[0][keyColumn - 1]) ? "分组键" : rows[0][keyColumn - 1].Trim();
                string sumHeader = string.IsNullOrWhiteSpace(rows[0][sumColumn - 1]) ? "求和列" : rows[0][sumColumn - 1].Trim();
                var output = new List<IList<string>> { new List<string> { keyHeader, sumHeader + "合计" } };
                foreach (WorksheetGroup group in groups)
                {
                    decimal total = WorksheetGrouping.SumNumericColumns(group.Rows, new[] { sumColumn - 1 })[0];
                    output.Add(new List<string> { group.KeyValues[0], total.ToString(CultureInfo.CurrentCulture) });
                }

                workbook = GetWorkbook(selection);
                resultSheet = AddOutputSheet(workbook, "分组求和结果");
                WriteRows(resultSheet, output);
                return string.Format("已按“{0}”将 {1} 行数据分为 {2} 组，并对“{3}”求和；结果以文本快照写入“{4}”，原选区未修改，结果尚未保存。",
                    keyHeader, dataRows.Count, groups.Count, sumHeader, resultSheet.Name);
            }
            finally
            {
                ReleaseComObject(resultSheet);
                ReleaseComObject(workbook);
                ReleaseComObject(selection);
            }
        }

        public string CreateGradeStatistics(double maximumScore, double passScore, double excellentScore)
        {
            Excel.Range selection = GetSelectionRange();
            Excel.Workbook workbook = null;
            Excel.Worksheet resultSheet = null;
            try
            {
                object[,] values = ReadMatrix(selection);
                var scores = new List<double>();
                int blankCount = 0;
                for (int row = 1; row <= values.GetLength(0); row++)
                {
                    for (int column = 1; column <= values.GetLength(1); column++)
                    {
                        object raw = values[row, column];
                        if (raw == null || (raw is string && string.IsNullOrWhiteSpace((string)raw)))
                        {
                            blankCount++;
                            continue;
                        }
                        string address = CellAddress(selection.Row + row - 1, selection.Column + column - 1);
                        if (raw is DateTime || raw is bool)
                            throw new InvalidOperationException("单元格 " + address + " 不是分数；请仅选择分数列，空白单元格会自动跳过。");

                        double score;
                        string text = raw as string;
                        if (text != null)
                        {
                            const NumberStyles styles = NumberStyles.Float | NumberStyles.AllowThousands;
                            if (!double.TryParse(text, styles, CultureInfo.CurrentCulture, out score)
                                && !double.TryParse(text, styles, CultureInfo.InvariantCulture, out score))
                                throw new InvalidOperationException("单元格 " + address + " 包含非数值文本：" + text);
                        }
                        else
                        {
                            try { score = Convert.ToDouble(raw, CultureInfo.CurrentCulture); }
                            catch (Exception ex) when (ex is FormatException || ex is InvalidCastException || ex is OverflowException)
                            {
                                throw new InvalidOperationException("单元格 " + address + " 不是有效分数。", ex);
                            }
                        }
                        if (double.IsNaN(score) || double.IsInfinity(score) || score < 0 || score > maximumScore)
                            throw new InvalidOperationException("单元格 " + address + " 的分数超出 0 至满分范围或不是有限数值。");
                        scores.Add(score);
                    }
                }

                GradeSummary summary = GradeStatistics.Analyze(scores, 0, maximumScore, passScore, excellentScore);
                var bands = new List<ScoreBand>();
                for (int index = 0; index < 5; index++)
                {
                    double lower = maximumScore * index / 5.0;
                    double upper = maximumScore * (index + 1) / 5.0;
                    bands.Add(new ScoreBand(string.Format(CultureInfo.CurrentCulture, "{0:p0}–{1:p0}", index / 5.0, (index + 1) / 5.0), lower, upper));
                }
                IReadOnlyList<ScoreBandResult> bandResults = GradeStatistics.CountByBands(scores, 0, maximumScore, bands);

                var output = new List<IList<string>>
                {
                    new List<string> { "指标", "数值", "统计口径" },
                    new List<string> { "有效分数个数", summary.Count.ToString(CultureInfo.CurrentCulture), "空白跳过；非空文本、错误值和越界分数会阻止统计" },
                    new List<string> { "最高分", summary.Maximum.ToString("0.##", CultureInfo.CurrentCulture), "满分 " + maximumScore.ToString("0.##", CultureInfo.CurrentCulture) },
                    new List<string> { "最低分", summary.Minimum.ToString("0.##", CultureInfo.CurrentCulture), "最低允许值 0" },
                    new List<string> { "平均分", summary.Average.ToString("0.##", CultureInfo.CurrentCulture), "算术平均值" },
                    new List<string> { "中位数", summary.Median.ToString("0.##", CultureInfo.CurrentCulture), "排序后中间值；偶数个取中间两数平均" },
                    new List<string> { "及格人数", summary.PassCount.ToString(CultureInfo.CurrentCulture), "分数 ≥ " + passScore.ToString("0.##", CultureInfo.CurrentCulture) },
                    new List<string> { "及格率", summary.PassRate.ToString("0.00", CultureInfo.CurrentCulture) + "%", "及格人数 / 有效分数个数" },
                    new List<string> { "优秀人数", summary.ExcellentCount.ToString(CultureInfo.CurrentCulture), "分数 ≥ " + excellentScore.ToString("0.##", CultureInfo.CurrentCulture) },
                    new List<string> { "优秀率", summary.ExcellentRate.ToString("0.00", CultureInfo.CurrentCulture) + "%", "优秀人数 / 有效分数个数" },
                    new List<string> { string.Empty, string.Empty, string.Empty },
                    new List<string> { "等距分数段", "人数", "区间规则" }
                };
                for (int index = 0; index < bandResults.Count; index++)
                {
                    ScoreBandResult band = bandResults[index];
                    string rule = string.Format(CultureInfo.CurrentCulture, "[{0:0.##}, {1:0.##}{2}", band.LowerBound, band.UpperBound, index == bandResults.Count - 1 ? "]" : ")");
                    output.Add(new List<string> { band.Name, band.Count.ToString(CultureInfo.CurrentCulture), rule });
                }

                workbook = GetWorkbook(selection);
                resultSheet = AddOutputSheet(workbook, "成绩统计结果");
                WriteRows(resultSheet, output);
                return string.Format("已统计 {0} 个有效分数并跳过 {1} 个空白单元格；结果以文本快照写入“{2}”，原选区未修改，结果尚未保存。",
                    summary.Count, blankCount, resultSheet.Name);
            }
            finally
            {
                ReleaseComObject(resultSheet);
                ReleaseComObject(workbook);
                ReleaseComObject(selection);
            }
        }

        public string ValidateIdentityCards()
        {
            Excel.Range selection = GetSelectionRange();
            Excel.Range textCells = null;
            Excel.Areas areas = null;
            int valid = 0;
            int invalid = 0;
            var invalidAddresses = new List<string>();
            try
            {
                try
                {
                    textCells = selection.SpecialCells(Excel.XlCellType.xlCellTypeConstants, Excel.XlSpecialCellsValue.xlTextValues);
                }
                catch (COMException ex) when (ex.ErrorCode == NoCellsFound)
                {
                    return "选区中没有可校验的身份证文本常量；数字和公式未检查、未转换。";
                }

                areas = textCells.Areas;
                for (int index = 1; index <= areas.Count; index++)
                {
                    Excel.Range area = null;
                    try
                    {
                        area = (Excel.Range)areas[index];
                        object[,] values = ReadMatrix(area);
                        for (int row = 1; row <= values.GetLength(0); row++)
                        {
                            for (int column = 1; column <= values.GetLength(1); column++)
                            {
                                string value = values[row, column] as string;
                                if (string.IsNullOrWhiteSpace(value)) continue;
                                if (ChineseIdCard.IsValid(value)) valid++;
                                else
                                {
                                    invalid++;
                                    if (invalidAddresses.Count < 12)
                                        invalidAddresses.Add(CellAddress(area.Row + row - 1, area.Column + column - 1));
                                }
                            }
                        }
                    }
                    finally
                    {
                        ReleaseComObject(area);
                    }
                }
                string detail = invalidAddresses.Count == 0 ? string.Empty : "；部分无效单元格：" + string.Join("、", invalidAddresses);
                if (invalid > invalidAddresses.Count) detail += " 等";
                return string.Format("校验完成：有效 {0} 个，无效 {1} 个{2}。仅检查文本常量，数字和公式未转换。", valid, invalid, detail);
            }
            finally
            {
                ReleaseComObject(areas);
                ReleaseComObject(textCells);
                ReleaseComObject(selection);
            }
        }

        public string MaskIdentityCards()
        {
            Excel.Range selection = GetSelectionRange();
            Excel.Range textCells = null;
            Excel.Areas areas = null;
            int masked = 0;
            int invalid = 0;
            try
            {
                try
                {
                    textCells = selection.SpecialCells(Excel.XlCellType.xlCellTypeConstants, Excel.XlSpecialCellsValue.xlTextValues);
                }
                catch (COMException ex) when (ex.ErrorCode == NoCellsFound)
                {
                    return "选区中没有可脱敏的身份证文本常量；数字和公式未修改、未转换。";
                }

                areas = textCells.Areas;
                for (int index = 1; index <= areas.Count; index++)
                {
                    Excel.Range area = null;
                    try
                    {
                        area = (Excel.Range)areas[index];
                        object[,] values = ReadMatrix(area);
                        bool areaChanged = false;
                        for (int row = 1; row <= values.GetLength(0); row++)
                        {
                            for (int column = 1; column <= values.GetLength(1); column++)
                            {
                                string value = values[row, column] as string;
                                if (string.IsNullOrWhiteSpace(value)) continue;
                                if (!ChineseIdCard.IsValid(value))
                                {
                                    invalid++;
                                    continue;
                                }
                                values[row, column] = ChineseIdCard.Mask(value.Trim());
                                masked++;
                                areaChanged = true;
                            }
                        }
                        if (areaChanged)
                        {
                            area.NumberFormat = "@";
                            area.Value2 = values;
                        }
                    }
                    finally
                    {
                        ReleaseComObject(area);
                    }
                }
                return string.Format("已脱敏 {0} 个有效身份证文本；{1} 个非空无效文本保持不变，数字和公式未转换。", masked, invalid);
            }
            finally
            {
                ReleaseComObject(areas);
                ReleaseComObject(textCells);
                ReleaseComObject(selection);
            }
        }

        private static List<RowGroup> GroupRows(IList<IList<string>> rows, IList<IList<object>> values, int keyColumn)
        {
            if (rows.Count != values.Count) throw new ArgumentException("文本与原始值行数不一致。", "values");
            var groups = new List<RowGroup>();
            var indexByKey = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            int dataCount = 0;
            for (int index = 1; index < rows.Count; index++)
            {
                IList<string> row = rows[index];
                if (TabularData.IsBlankRow(row)) continue;
                dataCount++;
                string key = keyColumn < row.Count ? (row[keyColumn] ?? string.Empty).Trim() : string.Empty;
                if (key.Length == 0) throw new InvalidOperationException("拆分键列存在空值，请先补全或清理数据。");
                int groupIndex;
                if (!indexByKey.TryGetValue(key, out groupIndex))
                {
                    if (groups.Count >= MaxGroups) throw new InvalidOperationException("拆分结果超过 200 组，请分批处理。");
                    groupIndex = groups.Count;
                    indexByKey.Add(key, groupIndex);
                    groups.Add(new RowGroup(key));
                }
                groups[groupIndex].Values.Add(values[index]);
            }
            if (dataCount == 0) throw new InvalidOperationException("选区没有可拆分的非空数据行。");
            return groups;
        }

        private static List<IList<object>> WithHeaderValues(IList<object> header, IEnumerable<IList<object>> rows)
        {
            var output = new List<IList<object>> { new List<object>(header) };
            output.AddRange(rows);
            return output;
        }

        private static List<string> BuildCompareHeader(IList<IList<string>> left, IList<IList<string>> right, int leftColumns, int rightColumns, bool hasHeader)
        {
            var header = new List<string> { "差异类型", "比较键" };
            for (int column = 0; column < leftColumns; column++) header.Add(ColumnHeader(left, hasHeader, column, "左："));
            for (int column = 0; column < rightColumns; column++) header.Add(ColumnHeader(right, hasHeader, column, "右："));
            return header;
        }

        private static string ColumnHeader(IList<IList<string>> rows, bool hasHeader, int column, string prefix)
        {
            if (hasHeader && rows.Count > 0 && column < rows[0].Count && !string.IsNullOrWhiteSpace(rows[0][column]))
                return prefix + rows[0][column];
            return prefix + "列" + (column + 1);
        }

        private static string DifferenceText(DifferenceKind kind)
        {
            if (kind == DifferenceKind.Added) return "仅右侧";
            if (kind == DifferenceKind.Removed) return "仅左侧";
            return "内容变化";
        }

        private static void AddPadded(List<string> target, IList<string> values, int count)
        {
            for (int column = 0; column < count; column++)
                target.Add(values != null && column < values.Count ? values[column] ?? string.Empty : string.Empty);
        }

        private static void AddPaddedValues(List<object> target, IList<object> values, int count)
        {
            for (int column = 0; column < count; column++)
                target.Add(values != null && column < values.Count ? values[column] : null);
        }

        private static bool RowsEqual(IList<string> left, IList<string> right)
        {
            if (left.Count != right.Count) return false;
            for (int index = 0; index < left.Count; index++)
            {
                string leftCell = (left[index] ?? string.Empty).Trim();
                string rightCell = (right[index] ?? string.Empty).Trim();
                if (!string.Equals(leftCell, rightCell, StringComparison.OrdinalIgnoreCase)) return false;
            }
            return true;
        }

        private static bool HeadersEqual(IList<string> left, IList<string> right)
        {
            if (left.Count != right.Count) return false;
            for (int index = 0; index < left.Count; index++)
                if (!string.Equals((left[index] ?? string.Empty).Trim(), (right[index] ?? string.Empty).Trim(), StringComparison.OrdinalIgnoreCase)) return false;
            return true;
        }

        private static List<IList<string>> ReadRows(Excel.Range range)
        {
            object[,] values = ReadMatrix(range);
            var rows = new List<IList<string>>(values.GetLength(0));
            for (int row = 1; row <= values.GetLength(0); row++)
            {
                var data = new List<string>(values.GetLength(1));
                for (int column = 1; column <= values.GetLength(1); column++)
                {
                    object value = values[row, column];
                    data.Add(value == null ? string.Empty : Convert.ToString(value, CultureInfo.CurrentCulture) ?? string.Empty);
                }
                rows.Add(data);
            }
            return rows;
        }

        private static List<IList<object>> ReadValueRows(Excel.Range range)
        {
            object[,] values = ReadMatrix(range);
            var rows = new List<IList<object>>(values.GetLength(0));
            for (int row = 1; row <= values.GetLength(0); row++)
            {
                var data = new List<object>(values.GetLength(1));
                for (int column = 1; column <= values.GetLength(1); column++) data.Add(values[row, column]);
                rows.Add(data);
            }
            return rows;
        }

        private static object[,] ReadMatrix(Excel.Range range)
        {
            object raw = range.Value;
            object[,] values = raw as object[,];
            if (values != null) return values;
            object[,] single = CreateMatrix(1, 1);
            single[1, 1] = raw;
            return single;
        }

        private static object[,] CreateMatrix(int rows, int columns)
        {
            return (object[,])Array.CreateInstance(typeof(object), new[] { rows, columns }, new[] { 1, 1 });
        }

        private static void WriteRows(Excel.Worksheet worksheet, IList<IList<string>> rows)
        {
            if (worksheet == null) throw new ArgumentNullException("worksheet");
            if (rows == null || rows.Count == 0) throw new ArgumentException("没有可写入的数据。", "rows");
            int columns = rows.Max(row => row.Count);
            if (columns == 0 || (long)rows.Count * columns > MaxCells)
                throw new InvalidOperationException("输出为空或超过 200000 个单元格，已停止写入。");
            object[,] values = CreateMatrix(rows.Count, columns);
            for (int row = 0; row < rows.Count; row++)
            {
                for (int column = 0; column < columns; column++)
                    values[row + 1, column + 1] = column < rows[row].Count ? rows[row][column] ?? string.Empty : string.Empty;
            }

            Excel.Range firstCell = null;
            Excel.Range target = null;
            try
            {
                firstCell = (Excel.Range)worksheet.Cells[1, 1];
                target = firstCell.Resize[rows.Count, columns];
                target.NumberFormat = "@";
                target.Value2 = values;
            }
            finally
            {
                ReleaseComObject(target);
                ReleaseComObject(firstCell);
            }
        }

        private static void WriteValues(Excel.Worksheet worksheet, IList<IList<object>> rows)
        {
            if (worksheet == null) throw new ArgumentNullException("worksheet");
            if (rows == null || rows.Count == 0) throw new ArgumentException("没有可写入的数据。", "rows");
            int columns = rows.Max(row => row.Count);
            if (columns == 0 || (long)rows.Count * columns > MaxCells)
                throw new InvalidOperationException("输出为空或超过 200000 个单元格，已停止写入。");
            object[,] values = CreateMatrix(rows.Count, columns);
            for (int row = 0; row < rows.Count; row++)
                for (int column = 0; column < columns; column++)
                    values[row + 1, column + 1] = column < rows[row].Count ? rows[row][column] : null;

            Excel.Range firstCell = null;
            Excel.Range target = null;
            try
            {
                firstCell = (Excel.Range)worksheet.Cells[1, 1];
                target = firstCell.Resize[rows.Count, columns];
                target.NumberFormat = "@";
                target.Value = values;
                target.NumberFormat = "General";
                for (int row = 1; row <= values.GetLength(0); row++)
                {
                    for (int column = 1; column <= values.GetLength(1); column++)
                    {
                        if (!(values[row, column] is DateTime)) continue;
                        DateTime date = (DateTime)values[row, column];
                        Excel.Range cell = null;
                        try
                        {
                            cell = (Excel.Range)target.Cells[row, column];
                            cell.NumberFormat = date.TimeOfDay == TimeSpan.Zero ? "yyyy-mm-dd" : "yyyy-mm-dd hh:mm:ss";
                        }
                        finally
                        {
                            ReleaseComObject(cell);
                        }
                    }
                }
            }
            finally
            {
                ReleaseComObject(target);
                ReleaseComObject(firstCell);
            }
        }

        private static void EnsureManageable(Excel.Range range)
        {
            if (range == null) throw new ArgumentNullException("range");
            Excel.Areas areas = null;
            try
            {
                areas = range.Areas;
                if (areas.Count != 1) throw new InvalidOperationException("请选择一个连续区域，暂不支持多重选区。");
            }
            finally
            {
                ReleaseComObject(areas);
            }
            long cells = (long)RowCount(range) * ColumnCount(range);
            if (cells > MaxCells) throw new InvalidOperationException("选区或已用区域超过 200000 个单元格，请缩小范围后重试。");
            object merged = range.MergeCells;
            if (merged == null || (merged is bool && (bool)merged))
                throw new InvalidOperationException("区域含合并单元格或合并状态不一致，为避免数据错位已停止操作。");
        }

        private static void EnsureNoFormulas(Excel.Range range, string operation)
        {
            object hasFormula = range.HasFormula;
            if (hasFormula == null || (hasFormula is bool && (bool)hasFormula))
                throw new InvalidOperationException(operation + "暂不支持含公式的区域，以免误删或破坏计算关系。");
        }

        private static int RowCount(Excel.Range range)
        {
            Excel.Range rows = null;
            try
            {
                rows = range.Rows;
                return rows.Count;
            }
            finally
            {
                ReleaseComObject(rows);
            }
        }

        private static int ColumnCount(Excel.Range range)
        {
            Excel.Range columns = null;
            try
            {
                columns = range.Columns;
                return columns.Count;
            }
            finally
            {
                ReleaseComObject(columns);
            }
        }

        private static void DeleteEntireRow(Excel.Worksheet worksheet, int rowIndex)
        {
            Excel.Range row = null;
            try
            {
                row = (Excel.Range)worksheet.Rows[rowIndex];
                row.Delete();
            }
            finally
            {
                ReleaseComObject(row);
            }
        }

        private static void DeleteSelectionRow(Excel.Worksheet worksheet, int rowIndex, int firstColumn, int columnCount)
        {
            Excel.Range first = null;
            Excel.Range last = null;
            Excel.Range row = null;
            try
            {
                first = (Excel.Range)worksheet.Cells[rowIndex, firstColumn];
                last = (Excel.Range)worksheet.Cells[rowIndex, firstColumn + columnCount - 1];
                row = worksheet.Range[first, last];
                row.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            }
            finally
            {
                ReleaseComObject(row);
                ReleaseComObject(last);
                ReleaseComObject(first);
            }
        }

        private static Excel.Workbook GetWorkbook(Excel.Range range)
        {
            Excel.Worksheet worksheet = null;
            try
            {
                worksheet = range.Worksheet;
                Excel.Workbook workbook = worksheet.Parent as Excel.Workbook;
                if (workbook == null) throw new InvalidOperationException("无法确定选区所属工作簿。");
                return workbook;
            }
            finally
            {
                ReleaseComObject(worksheet);
            }
        }

        private static void MarkOutputSheet(Excel.Worksheet worksheet)
        {
            Excel.CustomProperties properties = null;
            Excel.CustomProperty marker = null;
            try
            {
                properties = worksheet.CustomProperties;
                marker = properties.Add(OutputMarkerName, "1");
            }
            finally
            {
                ReleaseComObject(marker);
                ReleaseComObject(properties);
            }
        }

        private static bool IsGeneratedOutputSheet(Excel.Worksheet worksheet)
        {
            Excel.CustomProperties properties = null;
            try
            {
                properties = worksheet.CustomProperties;
                for (int index = 1; index <= properties.Count; index++)
                {
                    Excel.CustomProperty marker = null;
                    try
                    {
                        marker = properties[index];
                        if (string.Equals(marker.Name, OutputMarkerName, StringComparison.Ordinal)) return true;
                    }
                    finally
                    {
                        ReleaseComObject(marker);
                    }
                }
                return false;
            }
            finally
            {
                ReleaseComObject(properties);
            }
        }

        private static Excel.Worksheet AddOutputSheet(Excel.Workbook workbook, string requestedName, bool summaryResult = false)
        {
            if (workbook == null) throw new ArgumentNullException("workbook");
            Excel.Sheets sheets = null;
            Excel.Worksheet last = null;
            Excel.Worksheet created = null;
            try
            {
                sheets = workbook.Worksheets;
                last = sheets[sheets.Count] as Excel.Worksheet;
                created = sheets.Add(After: last) as Excel.Worksheet;
                if (created == null) throw new InvalidOperationException("无法创建输出工作表。");
                created.Name = GetUniqueSheetName(workbook, requestedName);
                if (summaryResult) MarkOutputSheet(created);
                return created;
            }
            catch
            {
                ReleaseComObject(created);
                throw;
            }
            finally
            {
                ReleaseComObject(last);
                ReleaseComObject(sheets);
            }
        }

        private static string GetUniqueSheetName(Excel.Workbook workbook, string requestedName)
        {
            string baseName = SanitizeSheetName(requestedName);
            if (!SheetExists(workbook, baseName)) return baseName;
            for (int index = 2; index < 10000; index++)
            {
                string suffix = "_" + index;
                string candidate = baseName.Substring(0, Math.Min(baseName.Length, 31 - suffix.Length)) + suffix;
                if (!SheetExists(workbook, candidate)) return candidate;
            }
            throw new InvalidOperationException("无法生成不冲突的工作表名称。");
        }

        private static bool SheetExists(Excel.Workbook workbook, string name)
        {
            Excel.Sheets sheets = null;
            try
            {
                sheets = workbook.Worksheets;
                for (int index = 1; index <= sheets.Count; index++)
                {
                    Excel.Worksheet sheet = null;
                    try
                    {
                        sheet = sheets[index] as Excel.Worksheet;
                        if (sheet != null && string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase)) return true;
                    }
                    finally
                    {
                        ReleaseComObject(sheet);
                    }
                }
                return false;
            }
            finally
            {
                ReleaseComObject(sheets);
            }
        }

        private static string SanitizeSheetName(string value)
        {
            char[] invalid = { ':', '\\', '/', '?', '*', '[', ']' };
            string name = new string((value ?? string.Empty).Select(c => invalid.Contains(c) || char.IsControl(c) ? '_' : c).ToArray()).Trim();
            if (name.Length == 0) name = "分组";
            return name.Length <= 31 ? name : name.Substring(0, 31);
        }

        private static string SanitizeFileName(string value)
        {
            char[] invalid = Path.GetInvalidFileNameChars();
            string name = new string((value ?? string.Empty).Select(c => invalid.Contains(c) || char.IsControl(c) ? '_' : c).ToArray()).Trim();
            if (name.Length == 0) name = "分组列";
            return name.Length <= 60 ? name : name.Substring(0, 60);
        }

        private static string CellAddress(int row, int column)
        {
            var letters = new Stack<char>();
            while (column > 0)
            {
                column--;
                letters.Push((char)('A' + column % 26));
                column /= 26;
            }
            return "$" + new string(letters.ToArray()) + "$" + row;
        }

        public static void ReleaseComObject(object value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
        }

        private sealed class RowGroup
        {
            public RowGroup(string key)
            {
                Key = key;
                Values = new List<IList<object>>();
            }

            public string Key { get; private set; }
            public List<IList<object>> Values { get; private set; }
        }

        private sealed class CollectedSheet
        {
            public CollectedSheet(string name, List<IList<object>> rows)
            {
                Name = name;
                Rows = rows;
            }

            public string Name { get; private set; }
            public List<IList<object>> Rows { get; private set; }
        }
    }
}
