using System;
using System.Linq;
using Suite.Core.Data;

namespace Suite.Core.Tests
{
    internal static class GradeStatisticsTests
    {
        public static void Run()
        {
            TestRunner.Run("GradeStatistics - 偶数人数完整指标", EvenSummary);
            TestRunner.Run("GradeStatistics - 奇数人数中位数", OddSummary);
            TestRunner.Run("GradeStatistics - 空集合明确失败", EmptySummary);
            TestRunner.Run("GradeStatistics - 非法分数和分数线", InvalidScores);
            TestRunner.Run("GradeStatistics - 分段边界及优秀线参数", Bands);
            TestRunner.Run("GradeStatistics - 空分段数据和非法分段", BandValidation);
        }

        private static void EvenSummary()
        {
            GradeSummary result = GradeStatistics.Analyze(new[] { 50.0, 60.0, 70.0, 80.0, 90.0, 100.0 }, 0, 100, 60, 90);
            TestAssert.Equal(6, result.Count);
            TestAssert.NearlyEqual(100, result.Maximum);
            TestAssert.NearlyEqual(50, result.Minimum);
            TestAssert.NearlyEqual(75, result.Average);
            TestAssert.NearlyEqual(75, result.Median);
            TestAssert.Equal(5, result.PassCount);
            TestAssert.Equal(2, result.ExcellentCount);
            TestAssert.NearlyEqual(500.0 / 6.0, result.PassRate);
            TestAssert.NearlyEqual(200.0 / 6.0, result.ExcellentRate);
        }

        private static void OddSummary()
        {
            GradeSummary result = GradeStatistics.Analyze(new[] { 9.0, 1.0, 2.0 }, 0, 10, 2, 9);
            TestAssert.NearlyEqual(4, result.Average);
            TestAssert.NearlyEqual(2, result.Median);
            TestAssert.NearlyEqual(200.0 / 3.0, result.PassRate);
            TestAssert.NearlyEqual(100.0 / 3.0, result.ExcellentRate);
        }

        private static void EmptySummary()
        {
            TestAssert.Throws<InvalidOperationException>(() => GradeStatistics.Analyze(Enumerable.Empty<double>(), 0, 100, 60, 90));
            TestAssert.Throws<ArgumentNullException>(() => GradeStatistics.Analyze(null, 0, 100, 60, 90));
        }

        private static void InvalidScores()
        {
            TestAssert.Throws<ArgumentOutOfRangeException>(() => GradeStatistics.Analyze(new[] { -0.01 }, 0, 100, 60, 90));
            TestAssert.Throws<ArgumentOutOfRangeException>(() => GradeStatistics.Analyze(new[] { 100.01 }, 0, 100, 60, 90));
            TestAssert.Throws<ArgumentOutOfRangeException>(() => GradeStatistics.Analyze(new[] { double.NaN }, 0, 100, 60, 90));
            TestAssert.Throws<ArgumentOutOfRangeException>(() => GradeStatistics.Analyze(new[] { double.PositiveInfinity }, 0, 100, 60, 90));
            TestAssert.Throws<ArgumentException>(() => GradeStatistics.Analyze(new[] { 0.0 }, 0, 0, 0, 0));
            TestAssert.Throws<ArgumentOutOfRangeException>(() => GradeStatistics.Analyze(new[] { 60.0 }, 0, 100, 101, 90));
            TestAssert.Throws<ArgumentException>(() => GradeStatistics.Analyze(new[] { 60.0 }, 0, 100, 60, 59));
        }

        private static void Bands()
        {
            var result = GradeStatistics.CountByBands(new[] { 0.0, 59.9, 60.0, 89.9, 90.0, 100.0 }, 0, 100, ValidBands());
            TestAssert.Equal(3, result.Count);
            TestAssert.SequenceEqual(new[] { 2, 2, 2 }, result.Select(x => x.Count));
            TestAssert.SequenceEqual(new[] { "不及格", "合格", "优秀" }, result.Select(x => x.Name));
            TestAssert.NearlyEqual(100, result[2].UpperBound);
        }

        private static void BandValidation()
        {
            var empty = GradeStatistics.CountByBands(Enumerable.Empty<double>(), 0, 100, ValidBands());
            TestAssert.SequenceEqual(new[] { 0, 0, 0 }, empty.Select(x => x.Count));
            TestAssert.Throws<ArgumentException>(() => GradeStatistics.CountByBands(new[] { 60.0 }, 0, 100,
                new[] { new ScoreBand("低", 0, 50), new ScoreBand("高", 60, 100) }));
            TestAssert.Throws<ArgumentException>(() => GradeStatistics.CountByBands(new[] { 60.0 }, 0, 100,
                new[] { new ScoreBand("仅一段", 0, 90) }));
            TestAssert.Throws<ArgumentException>(() => GradeStatistics.CountByBands(new[] { 101.0 }, 0, 100, ValidBands()));
            TestAssert.Throws<ArgumentException>(() => new ScoreBand(" ", 0, 10));
        }

        private static ScoreBand[] ValidBands()
        {
            return new[]
            {
                new ScoreBand("不及格", 0, 60),
                new ScoreBand("合格", 60, 90),
                new ScoreBand("优秀", 90, 100)
            };
        }
    }
}
