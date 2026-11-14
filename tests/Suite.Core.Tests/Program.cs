namespace Suite.Core.Tests
{
    internal static class Program
    {
        public static int Main()
        {
            ExistingAlgorithmTests.Run();
            AdvancedTextTransformsTests.Run();
            GradeStatisticsTests.Run();
            WorksheetGroupingTests.Run();
            return TestRunner.Complete();
        }
    }
}
