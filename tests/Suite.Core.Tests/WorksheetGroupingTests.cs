using System;
using System.Collections.Generic;
using System.Linq;
using Suite.Core.Data;

namespace Suite.Core.Tests
{
    internal static class WorksheetGroupingTests
    {
        public static void Run()
        {
            TestRunner.Run("WorksheetGrouping - 多列分组及组内稳定顺序", Grouping);
            TestRunner.Run("WorksheetGrouping - 大小写敏感与输入副本", GroupCaseAndCopies);
            TestRunner.Run("WorksheetGrouping - 按键拆分", Splitting);
            TestRunner.Run("WorksheetGrouping - 稳定升序和降序", Sorting);
            TestRunner.Run("WorksheetGrouping - 严格数值汇总", NumericSums);
            TestRunner.Run("WorksheetGrouping - 非数值和空值拒绝", NumericErrors);
            TestRunner.Run("WorksheetGrouping - 唯一键索引", UniqueIndex);
            TestRunner.Run("WorksheetGrouping - 重复键、空键和列参数异常", Validation);
        }

        private static void Grouping()
        {
            var rows = Rows(
                Row(" A ", "x", "1"),
                Row("b", "y", "2"),
                Row("a", "z", "3"),
                Row("a", "X", "4"));
            IReadOnlyList<WorksheetGroup> groups = WorksheetGrouping.GroupByColumns(rows, new[] { 0, 1 });
            TestAssert.Equal(3, groups.Count);
            TestAssert.SequenceEqual(new[] { "A", "x" }, groups[0].KeyValues);
            TestAssert.SequenceEqual(new[] { "1", "4" }, groups[0].Rows.Select(x => x[2]));
            TestAssert.SequenceEqual(new[] { "2" }, groups[1].Rows.Select(x => x[2]));
            TestAssert.SequenceEqual(new[] { "3" }, groups[2].Rows.Select(x => x[2]));
        }

        private static void GroupCaseAndCopies()
        {
            var rows = Rows(Row("A", "1"), Row("a", "2"));
            TestAssert.Equal(2, WorksheetGrouping.GroupByColumns(rows, new[] { 0 }, true).Count);
            IReadOnlyList<WorksheetGroup> groups = WorksheetGrouping.GroupByColumns(rows, new[] { 0 });
            TestAssert.Equal(1, groups.Count);
            rows[0][0] = "改变";
            TestAssert.Equal("A", groups[0].KeyValues[0]);
            TestAssert.Equal("A", groups[0].Rows[0][0]);
            TestAssert.Equal(0, WorksheetGrouping.GroupByColumns(Rows(), new[] { 0 }).Count);
        }

        private static void Splitting()
        {
            var groups = WorksheetGrouping.SplitByKey(Rows(Row("乙", "1"), Row("甲", "2"), Row("乙", "3")), 0);
            TestAssert.Equal(2, groups.Count);
            TestAssert.Equal("乙", groups[0].KeyValues[0]);
            TestAssert.SequenceEqual(new[] { "1", "3" }, groups[0].Rows.Select(x => x[1]));
        }

        private static void Sorting()
        {
            var rows = Rows(Row("b", "0"), Row("a", "first"), Row("A", "second"), Row("a", "third"));
            var ascending = WorksheetGrouping.StableSort(rows, 0);
            TestAssert.SequenceEqual(new[] { "first", "second", "third", "0" }, ascending.Select(x => x[1]));
            var descending = WorksheetGrouping.StableSort(rows, 0, true);
            TestAssert.SequenceEqual(new[] { "0", "first", "second", "third" }, descending.Select(x => x[1]));
            TestAssert.SequenceEqual(new[] { "b", "0" }, rows[0]);
        }

        private static void NumericSums()
        {
            var totals = WorksheetGrouping.SumNumericColumns(
                Rows(Row("1.25", "2"), Row("2.75", "3"), Row("1e2", "1,000")), new[] { 0, 1 });
            TestAssert.SequenceEqual(new[] { 104m, 1005m }, totals);
            var empty = WorksheetGrouping.SumNumericColumns(Rows(), new[] { 0 });
            TestAssert.SequenceEqual(new[] { 0m }, empty);
        }

        private static void NumericErrors()
        {
            TestAssert.Throws<FormatException>(() => WorksheetGrouping.SumNumericColumns(Rows(Row("abc")), new[] { 0 }));
            TestAssert.Throws<FormatException>(() => WorksheetGrouping.SumNumericColumns(Rows(Row(" ")), new[] { 0 }));
            TestAssert.Throws<FormatException>(() => WorksheetGrouping.SumNumericColumns(Rows(Row("NaN")), new[] { 0 }));
            TestAssert.Throws<ArgumentException>(() => WorksheetGrouping.SumNumericColumns(Rows(Row("1")), new[] { 1 }));
        }

        private static void UniqueIndex()
        {
            var result = WorksheetGrouping.IndexUniqueRows(Rows(Row(" A ", "1"), Row("b", "2")), 0);
            TestAssert.Equal(2, result.Count);
            TestAssert.SequenceEqual(new[] { " A ", "1" }, result["a"]);
        }

        private static void Validation()
        {
            TestAssert.Throws<InvalidOperationException>(() => WorksheetGrouping.IndexUniqueRows(Rows(Row("A"), Row("a")), 0));
            TestAssert.Throws<InvalidOperationException>(() => WorksheetGrouping.IndexUniqueRows(Rows(Row(" ")), 0));
            TestAssert.Throws<ArgumentException>(() => WorksheetGrouping.GroupByColumns(Rows(Row("A")), new[] { 0, 0 }));
            TestAssert.Throws<ArgumentException>(() => WorksheetGrouping.GroupByColumns(Rows(Row("A")), new int[0]));
            TestAssert.Throws<ArgumentOutOfRangeException>(() => WorksheetGrouping.StableSort(Rows(Row("A")), -1));
            TestAssert.Throws<ArgumentException>(() => WorksheetGrouping.SplitByKey(Rows(Row("A")), 1));
            TestAssert.Throws<ArgumentException>(() => WorksheetGrouping.GroupByColumns(new List<IList<string>> { null }, new[] { 0 }));
            TestAssert.Throws<ArgumentNullException>(() => WorksheetGrouping.GroupByColumns(null, new[] { 0 }));
        }

        private static IList<string> Row(params string[] cells)
        {
            return new List<string>(cells);
        }

        private static List<IList<string>> Rows(params IList<string>[] rows)
        {
            return new List<IList<string>>(rows);
        }
    }
}
