using System;
using System.Collections.Generic;
using System.Linq;

namespace Suite.Core.Data
{
    public enum DifferenceKind { Added, Removed, Changed }

    public sealed class DifferenceRow
    {
        public DifferenceRow(DifferenceKind kind, string key, IList<string> left, IList<string> right)
        { Kind = kind; Key = key; Left = left; Right = right; }
        public DifferenceKind Kind { get; private set; }
        public string Key { get; private set; }
        public IList<string> Left { get; private set; }
        public IList<string> Right { get; private set; }
    }

    public static class TabularData
    {
        public static bool IsBlankRow(IEnumerable<string> row)
        {
            return row == null || row.All(string.IsNullOrWhiteSpace);
        }

        public static List<IList<string>> RemoveBlankRows(IEnumerable<IList<string>> rows)
        {
            if (rows == null) throw new ArgumentNullException("rows");
            return rows.Where(x => !IsBlankRow(x)).Select(Copy).ToList();
        }

        public static List<IList<string>> DistinctRows(IEnumerable<IList<string>> rows, IEnumerable<int> keyColumns, bool caseSensitive)
        {
            if (rows == null) throw new ArgumentNullException("rows");
            int[] keys = (keyColumns ?? Enumerable.Empty<int>()).ToArray();
            var comparison = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            var seen = new HashSet<string>(StringComparer.Ordinal);
            var output = new List<IList<string>>();
            foreach (var row in rows)
            {
                IEnumerable<int> columns = keys.Length == 0 ? Enumerable.Range(0, row.Count) : keys;
                string key = string.Join("\u001f", columns.Select(column => Cell(row, column)).Select(x => caseSensitive ? x : x.ToUpperInvariant()));
                if (seen.Add(key)) output.Add(Copy(row));
            }
            return output;
        }

        public static List<IList<string>> CreatePayroll(IList<IList<string>> rows, bool hasHeader, int blankRows)
        {
            if (rows == null) throw new ArgumentNullException("rows");
            if (!hasHeader || rows.Count < 2) throw new ArgumentException("工资条需要一行表头和至少一行数据。", "rows");
            if (blankRows < 0 || blankRows > 20) throw new ArgumentOutOfRangeException("blankRows");
            var output = new List<IList<string>>();
            for (int index = 1; index < rows.Count; index++)
            {
                if (IsBlankRow(rows[index])) continue;
                output.Add(Copy(rows[0]));
                output.Add(Copy(rows[index]));
                for (int blank = 0; blank < blankRows; blank++) output.Add(new string[rows[0].Count]);
            }
            return output;
        }

        public static List<DifferenceRow> Compare(IList<IList<string>> left, IList<IList<string>> right, int keyColumn, bool caseSensitive)
        {
            if (left == null || right == null) throw new ArgumentNullException(left == null ? "left" : "right");
            if (keyColumn < 0) throw new ArgumentOutOfRangeException("keyColumn");
            var comparer = caseSensitive ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;
            var leftMap = UniqueByKey(left, keyColumn, comparer, "左侧");
            var rightMap = UniqueByKey(right, keyColumn, comparer, "右侧");
            var output = new List<DifferenceRow>();
            foreach (var pair in leftMap)
            {
                IList<string> rightRow;
                if (!rightMap.TryGetValue(pair.Key, out rightRow)) output.Add(new DifferenceRow(DifferenceKind.Removed, pair.Key, Copy(pair.Value), null));
                else if (!SameRow(pair.Value, rightRow, caseSensitive)) output.Add(new DifferenceRow(DifferenceKind.Changed, pair.Key, Copy(pair.Value), Copy(rightRow)));
            }
            foreach (var pair in rightMap) if (!leftMap.ContainsKey(pair.Key)) output.Add(new DifferenceRow(DifferenceKind.Added, pair.Key, null, Copy(pair.Value)));
            return output;
        }

        private static Dictionary<string, IList<string>> UniqueByKey(IEnumerable<IList<string>> rows, int keyColumn, StringComparer comparer, string side)
        {
            var map = new Dictionary<string, IList<string>>(comparer);
            foreach (var row in rows)
            {
                string key = Cell(row, keyColumn);
                if (key.Length == 0) throw new InvalidOperationException(side + "存在空比较键。");
                if (map.ContainsKey(key)) throw new InvalidOperationException(side + "存在重复比较键：" + key);
                map.Add(key, row);
            }
            return map;
        }

        private static bool SameRow(IList<string> left, IList<string> right, bool caseSensitive)
        {
            int count = Math.Max(left.Count, right.Count);
            var comparison = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            for (int index = 0; index < count; index++) if (!string.Equals(Cell(left, index), Cell(right, index), comparison)) return false;
            return true;
        }

        private static string Cell(IList<string> row, int column)
        {
            if (column < 0 || column >= row.Count || row[column] == null) return string.Empty;
            return row[column].Trim();
        }

        private static IList<string> Copy(IEnumerable<string> row) { return row.Select(x => x ?? string.Empty).ToList(); }
    }
}
