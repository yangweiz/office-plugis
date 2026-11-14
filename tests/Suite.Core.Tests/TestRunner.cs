using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Suite.Core.Tests
{
    internal static class TestAssert
    {
        public static void True(bool condition, string message = null)
        {
            if (!condition) throw new InvalidOperationException(message ?? "断言失败：条件应为 true。");
        }

        public static void False(bool condition, string message = null)
        {
            True(!condition, message ?? "断言失败：条件应为 false。");
        }

        public static void Equal<T>(T expected, T actual)
        {
            if (!EqualityComparer<T>.Default.Equals(expected, actual))
                throw new InvalidOperationException(string.Format("断言失败：期望 <{0}>，实际 <{1}>。", Display(expected), Display(actual)));
        }

        public static void NearlyEqual(double expected, double actual, double tolerance = 0.000001)
        {
            if (double.IsNaN(actual) || Math.Abs(expected - actual) > tolerance)
                throw new InvalidOperationException(string.Format("断言失败：期望 {0}，实际 {1}。", expected, actual));
        }

        public static void SequenceEqual<T>(IEnumerable<T> expected, IEnumerable<T> actual)
        {
            T[] expectedArray = expected.ToArray();
            T[] actualArray = actual.ToArray();
            if (!expectedArray.SequenceEqual(actualArray))
                throw new InvalidOperationException(string.Format("断言失败：期望序列 [{0}]，实际 [{1}]。",
                    string.Join(", ", expectedArray.Select(x => Display(x))), string.Join(", ", actualArray.Select(x => Display(x)))));
        }

        public static TException Throws<TException>(Action action) where TException : Exception
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                TException expected = ex as TException;
                if (expected != null) return expected;
                throw new InvalidOperationException(string.Format("断言失败：期望异常 {0}，实际 {1}。", typeof(TException).Name, ex.GetType().Name), ex);
            }
            throw new InvalidOperationException("断言失败：未抛出期望异常 " + typeof(TException).Name + "。");
        }

        private static string Display<T>(T value)
        {
            return value == null ? "null" : Convert.ToString(value, CultureInfo.InvariantCulture);
        }
    }

    internal static class TestRunner
    {
        private static int _passed;
        private static int _failed;

        public static void Run(string name, Action test)
        {
            try
            {
                test();
                _passed++;
                Console.WriteLine("[PASS] " + name);
            }
            catch (Exception ex)
            {
                _failed++;
                Console.WriteLine("[FAIL] " + name);
                Console.WriteLine(ex);
            }
        }

        public static int Complete()
        {
            Console.WriteLine(string.Format("测试结束：通过 {0}，失败 {1}。", _passed, _failed));
            return _failed == 0 ? 0 : 1;
        }
    }
}
