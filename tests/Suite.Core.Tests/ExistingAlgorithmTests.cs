using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Suite.Core.Batching;
using Suite.Core.Commands;
using Suite.Core.Data;
using Suite.Core.Text;
using Suite.Infrastructure;

namespace Suite.Core.Tests
{
    internal static class ExistingAlgorithmTests
    {
        public static void Run()
        {
            TestRunner.Run("TextCleanup - 不可见字符与 null", TextCleanupInvisible);
            TestRunner.Run("TextCleanup - 空行和行首空白", TextCleanupWhitespace);
            TestRunner.Run("TextCleanup - 断行重排与空输入", TextCleanupReflow);
            TestRunner.Run("OfficialDocumentText - 标题识别", OfficialHeadings);
            TestRunner.Run("OfficialDocumentText - 标题升级及不命中", OfficialPromotion);
            TestRunner.Run("ChineseIdCard - 校验、出生日期、性别和脱敏", IdCardNormal);
            TestRunner.Run("ChineseIdCard - 年龄边界", IdCardAge);
            TestRunner.Run("ChineseIdCard - 非法输入", IdCardInvalid);
            TestRunner.Run("TabularData - 空行清理", TabularBlankRows);
            TestRunner.Run("TabularData - 按键去重及大小写", TabularDistinct);
            TestRunner.Run("TabularData - 工资条及参数边界", TabularPayroll);
            TestRunner.Run("TabularData - 差异比较", TabularCompare);
            TestRunner.Run("TabularData - 重复键和空键异常", TabularCompareErrors);
            TestRunner.Run("BatchExecutor - 预检与失败隔离", BatchPrecheckAndFailure);
            TestRunner.Run("BatchExecutor - 执行前取消", BatchCancelledBeforeStart);
            TestRunner.Run("BatchExecutor - 执行中取消", BatchCancelledDuringExecution);
            TestRunner.Run("BatchExecutor - 空批次和空回调", BatchBoundaries);
            TestRunner.Run("OutputPathPolicy - 仅临时目录中的路径行为", OutputPaths);
            TestRunner.Run("TemplateStore - 尾部分隔符与归属边界", TemplateStoreOwnership);
            TestRunner.Run("Commands - 注册、查找和搜索", CommandRegistryBehavior);
            TestRunner.Run("Commands - 执行结果、阻塞和异常", CommandExecutorBehavior);
            TestRunner.Run("Commands - 取消和构造参数", CommandCancellationAndValidation);
        }

        private static void TextCleanupInvisible()
        {
            TestAssert.Equal<string>(null, TextCleanup.RemoveInvisibleCharacters(null));
            TestAssert.Equal("甲\r\n\t乙", TextCleanup.RemoveInvisibleCharacters("甲\u200B\0\r\n\t\uFEFF乙"));
            TestAssert.Equal(string.Empty, TextCleanup.RemoveInvisibleCharacters("\u200C\u2060"));
        }

        private static void TextCleanupWhitespace()
        {
            string newLine = Environment.NewLine;
            TestAssert.Equal("甲" + newLine + "乙", TextCleanup.CollapseBlankLines("甲\r\n \t\r\n\r\n乙"));
            TestAssert.Equal("甲" + newLine + "乙", TextCleanup.RemoveLeadingSpaces(" \t甲\r\n\u3000乙"));
            TestAssert.Equal<string>(null, TextCleanup.CollapseBlankLines(null));
        }

        private static void TextCleanupReflow()
        {
            string actual = TextCleanup.ReflowBrokenLines("一、标题\n这是第一\n行。\n\nHello\nworld!");
            TestAssert.Equal("一、标题" + Environment.NewLine + "这是第一行。" + Environment.NewLine + Environment.NewLine + "Hello world!", actual);
            TestAssert.Equal(string.Empty, TextCleanup.ReflowBrokenLines(string.Empty));
        }

        private static void OfficialHeadings()
        {
            TestAssert.True(OfficialDocumentText.IsHeading("第一章 总则"));
            TestAssert.True(OfficialDocumentText.IsHeading("1.1 标题"));
            TestAssert.False(OfficialDocumentText.IsHeading("普通正文"));
            TestAssert.False(OfficialDocumentText.IsHeading("  "));
        }

        private static void OfficialPromotion()
        {
            string result;
            TestAssert.Equal(1, OfficialDocumentText.PromoteNumberedHeading("（一）总则", out result));
            TestAssert.Equal("一、总则", result);
            TestAssert.Equal(1, OfficialDocumentText.PromoteNumberedHeading("1.总则", out result));
            TestAssert.Equal("（1）总则", result);
            TestAssert.Equal(0, OfficialDocumentText.PromoteNumberedHeading("第一章 总则", out result));
            TestAssert.Equal("第一章 总则", result);
            TestAssert.Equal(0, OfficialDocumentText.PromoteNumberedHeading(" ", out result));
        }

        private static void IdCardNormal()
        {
            const string female = "11010519491231002X";
            DateTime birth;
            TestAssert.True(ChineseIdCard.IsValid(female));
            TestAssert.True(ChineseIdCard.IsValid(female.ToLowerInvariant()));
            TestAssert.True(ChineseIdCard.TryGetBirthDate(female, out birth));
            TestAssert.Equal(new DateTime(1949, 12, 31), birth);
            TestAssert.Equal("女", ChineseIdCard.GetGender(female));
            TestAssert.Equal("男", ChineseIdCard.GetGender("110105194912310011"));
            TestAssert.Equal("110105********002X", ChineseIdCard.Mask(female));
        }

        private static void IdCardAge()
        {
            const string id = "11010519491231002X";
            TestAssert.Equal(74, ChineseIdCard.GetAge(id, new DateTime(2024, 12, 30)));
            TestAssert.Equal(75, ChineseIdCard.GetAge(id, new DateTime(2024, 12, 31)));
            TestAssert.Throws<ArgumentException>(() => ChineseIdCard.GetAge(id, new DateTime(1949, 12, 30)));
        }

        private static void IdCardInvalid()
        {
            DateTime ignored;
            TestAssert.False(ChineseIdCard.IsValid("110105194912310020"));
            TestAssert.False(ChineseIdCard.TryGetBirthDate("11010519491331002X", out ignored));
            TestAssert.False(ChineseIdCard.TryGetBirthDate("123", out ignored));
            TestAssert.Throws<ArgumentException>(() => ChineseIdCard.GetGender("invalid"));
            TestAssert.Throws<ArgumentException>(() => ChineseIdCard.Mask("invalid"));
        }

        private static void TabularBlankRows()
        {
            TestAssert.True(TabularData.IsBlankRow(null));
            var rows = Rows(Row(" ", ""), Row("甲", null));
            List<IList<string>> result = TabularData.RemoveBlankRows(rows);
            TestAssert.Equal(1, result.Count);
            TestAssert.SequenceEqual(new[] { "甲", "" }, result[0]);
            TestAssert.Throws<ArgumentNullException>(() => TabularData.RemoveBlankRows(null));
        }

        private static void TabularDistinct()
        {
            var rows = Rows(Row("A", "1"), Row("a", "2"), Row("B", "3"));
            TestAssert.Equal(2, TabularData.DistinctRows(rows, new[] { 0 }, false).Count);
            TestAssert.Equal(3, TabularData.DistinctRows(rows, new[] { 0 }, true).Count);
            TestAssert.SequenceEqual(new[] { "A", "1" }, TabularData.DistinctRows(rows, new[] { 0 }, false)[0]);
        }

        private static void TabularPayroll()
        {
            var rows = Rows(Row("姓名", "工资"), Row("甲", "1"), Row("", ""), Row("乙", "2"));
            List<IList<string>> result = TabularData.CreatePayroll(rows, true, 1);
            TestAssert.Equal(6, result.Count);
            TestAssert.SequenceEqual(new[] { "姓名", "工资" }, result[3]);
            TestAssert.True(TabularData.IsBlankRow(result[2]));
            TestAssert.Throws<ArgumentException>(() => TabularData.CreatePayroll(Rows(Row("姓名")), true, 0));
            TestAssert.Throws<ArgumentOutOfRangeException>(() => TabularData.CreatePayroll(rows, true, 21));
        }

        private static void TabularCompare()
        {
            var left = Rows(Row("A", "1"), Row("Gone", "x"), Row("Same", "s"));
            var right = Rows(Row("a", "2"), Row("New", "n"), Row("same", "s"));
            List<DifferenceRow> result = TabularData.Compare(left, right, 0, false);
            TestAssert.Equal(3, result.Count);
            DifferenceRow changed = result.Single(x => x.Kind == DifferenceKind.Changed);
            TestAssert.Equal("A", changed.Key);
            TestAssert.SequenceEqual(new[] { "a", "2" }, changed.Right);
            TestAssert.Equal("Gone", result.Single(x => x.Kind == DifferenceKind.Removed).Key);
            TestAssert.Equal("New", result.Single(x => x.Kind == DifferenceKind.Added).Key);
        }

        private static void TabularCompareErrors()
        {
            TestAssert.Throws<InvalidOperationException>(() => TabularData.Compare(Rows(Row("A"), Row("a")), Rows(Row("B")), 0, false));
            TestAssert.Throws<InvalidOperationException>(() => TabularData.Compare(Rows(Row(" ")), Rows(Row("B")), 0, false));
            TestAssert.Throws<ArgumentOutOfRangeException>(() => TabularData.Compare(Rows(), Rows(), -1, false));
        }

        private static void BatchPrecheckAndFailure()
        {
            var executor = new BatchExecutor();
            var executed = new List<int>();
            var progress = new ListProgress();
            BatchExecutionResult<int> result = executor.Execute(new[] { 1, 2, 3, 4, 5 },
                delegate(int item)
                {
                    if (item == 2) return "缺少输入";
                    if (item == 3) throw new InvalidOperationException("预检异常");
                    return null;
                },
                delegate(int item, CancellationToken token)
                {
                    executed.Add(item);
                    if (item == 4) throw new InvalidOperationException("单项失败");
                    return "输出" + item;
                }, progress, CancellationToken.None);

            TestAssert.Equal(2, result.Succeeded);
            TestAssert.Equal(2, result.Blocked);
            TestAssert.Equal(1, result.Failed);
            TestAssert.False(result.IsCompleteSuccess);
            TestAssert.SequenceEqual(new[] { 1, 4, 5 }, executed);
            TestAssert.Equal("输出5", result.Items[4].Output);
            TestAssert.Equal("单项失败", result.Items[3].Message);
            TestAssert.True(result.Items[2].Message.Contains("预检异常"));
            TestAssert.Equal(3, progress.Messages.Count);
        }

        private static void BatchCancelledBeforeStart()
        {
            using (var source = new CancellationTokenSource())
            {
                source.Cancel();
                int executions = 0;
                BatchExecutionResult<int> result = new BatchExecutor().Execute(new[] { 1, 2 }, null,
                    delegate(int item, CancellationToken token) { executions++; return string.Empty; }, null, source.Token);
                TestAssert.Equal(2, result.Cancelled);
                TestAssert.Equal(0, executions);
            }
        }

        private static void BatchCancelledDuringExecution()
        {
            using (var source = new CancellationTokenSource())
            {
                BatchExecutionResult<int> result = new BatchExecutor().Execute(new[] { 1, 2, 3 }, null,
                    delegate(int item, CancellationToken token)
                    {
                        if (item == 1) source.Cancel();
                        return item.ToString();
                    }, null, source.Token);
                TestAssert.Equal(BatchItemStatus.Succeeded, result.Items[0].Status);
                TestAssert.Equal(2, result.Cancelled);
            }
        }

        private static void BatchBoundaries()
        {
            BatchExecutionResult<int> empty = new BatchExecutor().Execute(new int[0], null,
                (item, token) => string.Empty, null, CancellationToken.None);
            TestAssert.Equal(0, empty.Items.Count);
            TestAssert.True(empty.IsCompleteSuccess);
            TestAssert.Throws<ArgumentNullException>(() => new BatchExecutor().Execute(new[] { 1 }, null, null, null, CancellationToken.None));
        }

        private static void OutputPaths()
        {
            string root = Path.Combine(Path.GetTempPath(), "Suite.Core.Tests-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try
            {
                string input = Path.Combine(root, "input.txt");
                string output = Path.Combine(root, "output.txt");
                File.WriteAllText(input, "input");
                TestAssert.Equal(Path.GetFullPath(output), OutputPathPolicy.EnsureDistinct(input, output));
                TestAssert.Throws<InvalidOperationException>(() => OutputPathPolicy.EnsureDistinct(input, input));
                TestAssert.Throws<ArgumentException>(() => OutputPathPolicy.EnsureDistinct(input, " "));

                string requested = Path.Combine(root, "report.txt");
                TestAssert.Equal(Path.GetFullPath(requested), OutputPathPolicy.GetAvailablePath(requested));
                File.WriteAllText(requested, "0");
                string first = Path.Combine(root, "report (1).txt");
                TestAssert.Equal(first, OutputPathPolicy.GetAvailablePath(requested));
                File.WriteAllText(first, "1");
                TestAssert.Equal(Path.Combine(root, "report (2).txt"), OutputPathPolicy.GetAvailablePath(requested));
            }
            finally
            {
                Directory.Delete(root, true);
            }
        }

        private static void TemplateStoreOwnership()
        {
            string root = Path.Combine(Path.GetTempPath(), "Suite.Core.Tests-" + Guid.NewGuid().ToString("N"));
            string templateDirectory = Path.Combine(root, "templates");
            string exportDirectory = Path.Combine(root, "export");
            Directory.CreateDirectory(templateDirectory);
            try
            {
                string sourcePath = Path.Combine(root, "source.dotx");
                File.WriteAllText(sourcePath, "original template");
                var store = new TemplateStore(templateDirectory + Path.DirectorySeparatorChar);
                TemplateDescriptor imported = store.Import(sourcePath);
                TestAssert.Equal(1, store.List().Count);
                string exported = store.Export(imported, exportDirectory);
                TestAssert.True(File.Exists(exported));

                string outsidePath = Path.Combine(root, "outside.dotx");
                File.WriteAllText(outsidePath, "outside");
                var outside = new TemplateDescriptor("outside", "Word", outsidePath);
                TestAssert.Throws<InvalidOperationException>(() => store.Delete(outside));
                TestAssert.True(File.Exists(outsidePath));

                store.Delete(imported);
                TestAssert.False(File.Exists(imported.Path));
            }
            finally
            {
                Directory.Delete(root, true);
            }
        }

        private static void CommandRegistryBehavior()
        {
            var registry = new CommandRegistry();
            DelegateSuiteCommand command = Command("Text.Clean", "清理文本", "文本 清理", SuiteHost.Word, true, null);
            registry.Register(command);
            TestAssert.True(object.ReferenceEquals(command, registry.Find("text.clean")));
            TestAssert.Equal<string>(null, registry.Find(null));
            TestAssert.Equal(1, registry.Search(SuiteHost.Word, "文本").Count);
            TestAssert.Equal(0, registry.Search(SuiteHost.Excel, string.Empty).Count);
            TestAssert.Throws<InvalidOperationException>(() => registry.Register(Command("text.clean", "重复", string.Empty, SuiteHost.Word, true, null)));
            TestAssert.Throws<ArgumentNullException>(() => registry.Register(null));
        }

        private static void CommandExecutorBehavior()
        {
            var executor = new CommandExecutor();
            var context = new CommandContext(SuiteHost.Word, null, null, null, null);
            CommandResult success = executor.Execute(Command("ok", "正常", string.Empty, SuiteHost.Word, true, null), context, CancellationToken.None);
            TestAssert.Equal(CommandStatus.Succeeded, success.Status);
            CommandResult blocked = executor.Execute(Command("blocked", "阻塞", string.Empty, SuiteHost.Word, false, null), context, CancellationToken.None);
            TestAssert.Equal(CommandStatus.Blocked, blocked.Status);
            CommandResult failure = executor.Execute(Command("bad", "失败功能", string.Empty, SuiteHost.Word, true,
                (contextValue, token) => { throw new InvalidOperationException("错误"); }), context, CancellationToken.None);
            TestAssert.Equal(CommandStatus.Failed, failure.Status);
            TestAssert.True(failure.Message.Contains("失败功能"));
            TestAssert.Equal(CommandStatus.Failed, executor.Execute(null, context, CancellationToken.None).Status);
            TestAssert.Equal(CommandStatus.Failed, executor.Execute(Command("ok2", "正常", string.Empty, SuiteHost.Word, true, null), null, CancellationToken.None).Status);
        }

        private static void CommandCancellationAndValidation()
        {
            using (var source = new CancellationTokenSource())
            {
                source.Cancel();
                bool executed = false;
                DelegateSuiteCommand command = Command("cancel", "取消", string.Empty, SuiteHost.Word, true,
                    delegate(CommandContext context, CancellationToken token) { executed = true; return CommandResult.Success("不应执行"); });
                var context = new CommandContext(SuiteHost.Word, null, null, null, null);
                TestAssert.Equal(CommandStatus.Cancelled, new CommandExecutor().Execute(command, context, source.Token).Status);
                TestAssert.False(executed);
            }
            TestAssert.Throws<ArgumentException>(() => new DelegateSuiteCommand(" ", "名称", string.Empty, SuiteHost.Word, CommandRisk.ReadOnly, false, null,
                (context, token) => CommandResult.Success("ok")));
            TestAssert.Throws<ArgumentNullException>(() => new DelegateSuiteCommand("id", "名称", string.Empty, SuiteHost.Word, CommandRisk.ReadOnly, false, null, null));
        }

        private static DelegateSuiteCommand Command(string id, string displayName, string keywords, SuiteHost hosts, bool canExecute,
            Func<CommandContext, CancellationToken, CommandResult> execute)
        {
            Func<CommandContext, bool> predicate = canExecute ? (Func<CommandContext, bool>)null : context => false;
            if (execute == null) execute = (context, token) => CommandResult.Success("ok");
            return new DelegateSuiteCommand(id, displayName, keywords, hosts, CommandRisk.ReadOnly, false, predicate, execute);
        }

        private static IList<string> Row(params string[] cells)
        {
            return new List<string>(cells);
        }

        private static List<IList<string>> Rows(params IList<string>[] rows)
        {
            return new List<IList<string>>(rows);
        }

        private sealed class ListProgress : IProgress<string>
        {
            public List<string> Messages { get; private set; }

            public ListProgress()
            {
                Messages = new List<string>();
            }

            public void Report(string value)
            {
                Messages.Add(value);
            }
        }
    }
}
