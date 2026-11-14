using System;
using System.Collections.Generic;
using System.Linq;

namespace Suite.Core.Data
{
    public sealed class GradeSummary
    {
        internal GradeSummary(int count, double maximum, double minimum, double average, double median, int passCount, int excellentCount)
        {
            Count = count;
            Maximum = maximum;
            Minimum = minimum;
            Average = average;
            Median = median;
            PassCount = passCount;
            ExcellentCount = excellentCount;
            PassRate = 100.0 * passCount / count;
            ExcellentRate = 100.0 * excellentCount / count;
        }

        public int Count { get; private set; }
        public double Maximum { get; private set; }
        public double Minimum { get; private set; }
        public double Average { get; private set; }
        public double Median { get; private set; }
        public int PassCount { get; private set; }
        public int ExcellentCount { get; private set; }
        public double PassRate { get; private set; }
        public double ExcellentRate { get; private set; }
    }

    public sealed class ScoreBand
    {
        public ScoreBand(string name, double lowerBound, double upperBound)
        {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("分段名称不能为空。", "name");
            Name = name;
            LowerBound = lowerBound;
            UpperBound = upperBound;
        }

        public string Name { get; private set; }
        public double LowerBound { get; private set; }
        public double UpperBound { get; private set; }
    }

    public sealed class ScoreBandResult
    {
        internal ScoreBandResult(ScoreBand band, int count)
        {
            Name = band.Name;
            LowerBound = band.LowerBound;
            UpperBound = band.UpperBound;
            Count = count;
        }

        public string Name { get; private set; }
        public double LowerBound { get; private set; }
        public double UpperBound { get; private set; }
        public int Count { get; private set; }
    }

    public static class GradeStatistics
    {
        public static GradeSummary Analyze(IEnumerable<double> scores, double minimumScore, double maximumScore, double passScore, double excellentScore)
        {
            ValidateRange(minimumScore, maximumScore);
            ValidateThreshold(passScore, minimumScore, maximumScore, "passScore");
            ValidateThreshold(excellentScore, minimumScore, maximumScore, "excellentScore");
            if (excellentScore < passScore) throw new ArgumentException("优秀线不得低于及格线。", "excellentScore");

            List<double> values = ReadAndValidate(scores, minimumScore, maximumScore);
            if (values.Count == 0) throw new InvalidOperationException("没有分数，无法计算统计指标。");
            values.Sort();

            int passCount = values.Count(x => x >= passScore);
            int excellentCount = values.Count(x => x >= excellentScore);
            double average = values.Sum() / values.Count;
            int middle = values.Count / 2;
            double median = values.Count % 2 == 1
                ? values[middle]
                : values[middle - 1] / 2.0 + values[middle] / 2.0;
            return new GradeSummary(values.Count, values[values.Count - 1], values[0], average, median, passCount, excellentCount);
        }

        public static IReadOnlyList<ScoreBandResult> CountByBands(IEnumerable<double> scores, double minimumScore, double maximumScore, IEnumerable<ScoreBand> bands)
        {
            ValidateRange(minimumScore, maximumScore);
            if (bands == null) throw new ArgumentNullException("bands");
            List<ScoreBand> orderedBands = bands.ToList();
            if (orderedBands.Count == 0) throw new ArgumentException("至少需要一个分数段。", "bands");

            double expectedLower = minimumScore;
            for (int index = 0; index < orderedBands.Count; index++)
            {
                ScoreBand band = orderedBands[index];
                if (band == null) throw new ArgumentException("分数段不能为空。", "bands");
                if (!IsFinite(band.LowerBound) || !IsFinite(band.UpperBound) || band.LowerBound >= band.UpperBound)
                    throw new ArgumentException("分数段边界无效。", "bands");
                if (band.LowerBound != expectedLower || band.UpperBound > maximumScore)
                    throw new ArgumentException("分数段必须连续、不重叠并覆盖整个分数范围。", "bands");
                expectedLower = band.UpperBound;
            }
            if (expectedLower != maximumScore)
                throw new ArgumentException("分数段必须连续、不重叠并覆盖整个分数范围。", "bands");

            List<double> values = ReadAndValidate(scores, minimumScore, maximumScore);
            var counts = new int[orderedBands.Count];
            foreach (double score in values)
            {
                for (int index = 0; index < orderedBands.Count; index++)
                {
                    ScoreBand band = orderedBands[index];
                    bool isLast = index == orderedBands.Count - 1;
                    if (score >= band.LowerBound && (score < band.UpperBound || (isLast && score == band.UpperBound)))
                    {
                        counts[index]++;
                        break;
                    }
                }
            }

            return orderedBands.Select((band, index) => new ScoreBandResult(band, counts[index])).ToList().AsReadOnly();
        }

        private static List<double> ReadAndValidate(IEnumerable<double> scores, double minimumScore, double maximumScore)
        {
            if (scores == null) throw new ArgumentNullException("scores");
            var values = new List<double>();
            foreach (double score in scores)
            {
                if (!IsFinite(score) || score < minimumScore || score > maximumScore)
                    throw new ArgumentOutOfRangeException("scores", "存在非有限或超出合法范围的分数。");
                values.Add(score);
            }
            return values;
        }

        private static void ValidateRange(double minimumScore, double maximumScore)
        {
            if (!IsFinite(minimumScore) || !IsFinite(maximumScore) || minimumScore >= maximumScore)
                throw new ArgumentException("分数范围必须为有限值，且最低分小于最高分。", "minimumScore");
        }

        private static void ValidateThreshold(double threshold, double minimumScore, double maximumScore, string parameterName)
        {
            if (!IsFinite(threshold) || threshold < minimumScore || threshold > maximumScore)
                throw new ArgumentOutOfRangeException(parameterName, "分数线必须位于合法分数范围内。");
        }

        private static bool IsFinite(double value)
        {
            return !double.IsNaN(value) && !double.IsInfinity(value);
        }
    }
}
