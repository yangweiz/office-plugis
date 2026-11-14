using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;

namespace Suite.Core.Data
{
    public sealed class WorksheetGroup
    {
        private readonly List<IList<string>> _rows = new List<IList<string>>();
        private readonly ReadOnlyCollection<IList<string>> _readOnlyRows;

        internal WorksheetGroup(string[] keyValues)
        {
            KeyValues = Array.AsReadOnly(keyValues);
            _readOnlyRows = _rows.AsReadOnly();
        }

        public IReadOnlyList<string> KeyValues { get; private set; }
        public IReadOnlyList<IList<string>> Rows { get { return _readOnlyRows; } }

        internal void Add(IList<string> row)
        {
            _rows.Add(row);
        }
    }

    public static class WorksheetGrouping
    {
        public static IReadOnlyList<WorksheetGroup> GroupByColumns(IEnumerable<IList<string>> rows, IEnumerable<int> keyColumns, bool caseSensitive = false)
        {
            int[] columns = ValidateColumnSelection(keyColumns, "keyColumns");
            List<IList<string>> copies = CopyRows(rows);
            var comparer = caseSensitive ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;
            var lookup = new Dictionary<string, WorksheetGroup>(comparer);
            var groups = new List<WorksheetGroup>();

            for (int rowIndex = 0; rowIndex < copies.Count; rowIndex++)
            {
                IList<string> row = copies[rowIndex];
                ValidateRowColumns(row, columns, rowIndex);
                var keyValues = new string[columns.Length];
                for (int index = 0; index < columns.Length; index++) keyValues[index] = Cell(row, columns[index]);
                string encodedKey = EncodeKey(keyValues);

                WorksheetGroup group;
                if (!lookup.TryGetValue(encodedKey, out group))
                {
                    group = new WorksheetGroup(keyValues);
                    lookup.Add(encodedKey, group);
                    groups.Add(group);
                }
                group.Add(row);
            }
            return groups.AsReadOnly();
        }

        public static IReadOnlyList<WorksheetGroup> SplitByKey(IEnumerable<IList<string>> rows, int keyColumn, bool caseSensitive = false)
        {
            return GroupByColumns(rows, new[] { keyColumn }, caseSensitive);
        }

        public static IReadOnlyList<IList<string>> StableSort(IEnumerable<IList<string>> rows, int sortColumn, bool descending = false, bool caseSensitive = false)
        {
            int[] columns = ValidateColumnSelection(new[] { sortColumn }, "sortColumn");
            List<IList<string>> copies = CopyRows(rows);
            for (int index = 0; index < copies.Count; index++) ValidateRowColumns(copies[index], columns, index);
            var comparer = caseSensitive ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;

            IOrderedEnumerable<IList<string>> sorted = descending
                ? copies.OrderByDescending(row => Cell(row, sortColumn), comparer)
                : copies.OrderBy(row => Cell(row, sortColumn), comparer);
            return sorted.ToList().AsReadOnly();
        }

        public static IReadOnlyList<decimal> SumNumericColumns(IEnumerable<IList<string>> rows, IEnumerable<int> columns)
        {
            int[] selectedColumns = ValidateColumnSelection(columns, "columns");
            List<IList<string>> copies = CopyRows(rows);
            var totals = new decimal[selectedColumns.Length];
            const NumberStyles styles = NumberStyles.Float | NumberStyles.AllowThousands;

            for (int rowIndex = 0; rowIndex < copies.Count; rowIndex++)
            {
                IList<string> row = copies[rowIndex];
                ValidateRowColumns(row, selectedColumns, rowIndex);
                for (int index = 0; index < selectedColumns.Length; index++)
                {
                    string cell = Cell(row, selectedColumns[index]);
                    decimal value;
                    if (!decimal.TryParse(cell, styles, CultureInfo.InvariantCulture, out value))
                        throw new FormatException(string.Format("第 {0} 行第 {1} 列不是有效数值：{2}", rowIndex + 1, selectedColumns[index] + 1, cell));
                    totals[index] += value;
                }
            }
            return Array.AsReadOnly(totals);
        }

        public static IReadOnlyDictionary<string, IList<string>> IndexUniqueRows(IEnumerable<IList<string>> rows, int keyColumn, bool caseSensitive = false)
        {
            int[] columns = ValidateColumnSelection(new[] { keyColumn }, "keyColumn");
            List<IList<string>> copies = CopyRows(rows);
            var comparer = caseSensitive ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;
            var result = new Dictionary<string, IList<string>>(comparer);
            for (int rowIndex = 0; rowIndex < copies.Count; rowIndex++)
            {
                IList<string> row = copies[rowIndex];
                ValidateRowColumns(row, columns, rowIndex);
                string key = Cell(row, keyColumn);
                if (key.Length == 0) throw new InvalidOperationException("唯一键不能为空。");
                if (result.ContainsKey(key)) throw new InvalidOperationException("存在重复唯一键：" + key);
                result.Add(key, row);
            }
            return new ReadOnlyDictionary<string, IList<string>>(result);
        }

        private static int[] ValidateColumnSelection(IEnumerable<int> columns, string parameterName)
        {
            if (columns == null) throw new ArgumentNullException(parameterName);
            int[] selected = columns.ToArray();
            if (selected.Length == 0) throw new ArgumentException("至少选择一列。", parameterName);
            var seen = new HashSet<int>();
            foreach (int column in selected)
            {
                if (column < 0) throw new ArgumentOutOfRangeException(parameterName, "列索引不能为负数。");
                if (!seen.Add(column)) throw new ArgumentException("不能重复选择同一列。", parameterName);
            }
            return selected;
        }

        private static List<IList<string>> CopyRows(IEnumerable<IList<string>> rows)
        {
            if (rows == null) throw new ArgumentNullException("rows");
            var copies = new List<IList<string>>();
            foreach (IList<string> row in rows)
            {
                if (row == null) throw new ArgumentException("数据行不能为空。", "rows");
                var copy = new List<string>(row.Count);
                for (int index = 0; index < row.Count; index++) copy.Add(row[index] ?? string.Empty);
                copies.Add(copy);
            }
            return copies;
        }

        private static void ValidateRowColumns(IList<string> row, IEnumerable<int> columns, int rowIndex)
        {
            foreach (int column in columns)
            {
                if (column >= row.Count)
                    throw new ArgumentException(string.Format("第 {0} 行缺少第 {1} 列。", rowIndex + 1, column + 1), "rows");
            }
        }

        private static string Cell(IList<string> row, int column)
        {
            return (row[column] ?? string.Empty).Trim();
        }

        private static string EncodeKey(IEnumerable<string> values)
        {
            var result = new StringBuilder();
            foreach (string value in values)
                result.Append(value.Length).Append(':').Append(value).Append(';');
            return result.ToString();
        }
    }
}
