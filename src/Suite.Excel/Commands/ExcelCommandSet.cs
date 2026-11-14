using System;
using System.Globalization;
using System.Threading;
using System.Windows.Forms;
using Suite.Core.Commands;
using Suite.Excel.Services;
using Excel = Microsoft.Office.Interop.Excel;

namespace Suite.Excel.Commands
{
    internal sealed class ExcelCommandSet
    {
        private readonly Excel.Application _application;
        private readonly WorksheetDataService _dataService;

        public ExcelCommandSet(Excel.Application application, WorksheetDataService dataService)
        {
            if (application == null) throw new ArgumentNullException("application");
            if (dataService == null) throw new ArgumentNullException("dataService");
            _application = application;
            _dataService = dataService;
        }

        public void RegisterAll(CommandRegistry registry)
        {
            if (registry == null) throw new ArgumentNullException("registry");
            Register(registry, "E-DEL-001-TEXT", "清理选区文本", "文本 清理 隐形字符 断行 空格", CommandRisk.ModifyDocument,
                (context, token) => Simple(_dataService.CleanupText, token));
            Register(registry, "E-DEL-001-BLANK", "删除空行", "空行 已用区域 清理", CommandRisk.ModifyDocument,
                (context, token) => Simple(_dataService.RemoveBlankRows, token));
            Register(registry, "E-DEL-002-ROWS", "选区去重", "重复 去重 唯一 整行", CommandRisk.ModifyDocument,
                (context, token) => Simple(_dataService.RemoveDuplicateRows, token));
            Register(registry, "E-SORT-001-FIRST", "选区排序", "排序 升序 首列", CommandRisk.ModifyDocument, Sort);
            Register(registry, "E-CMP-001-KEY", "按键列对比", "比较 对比 差异 新增 删除 变化", CommandRisk.ModifyDocument, Compare);
            Register(registry, "E-MERGE-001-SHEETS", "多表汇总", "汇总 合并 工作表", CommandRisk.ModifyDocument, Summarize);
            Register(registry, "E-SPLIT-002-COLUMN", "按列拆分", "拆分 分组 新工作簿", CommandRisk.ModifyFiles, Split);
            Register(registry, "E-MERGE-001-GROUPSUM", "分组求和", "分组 汇总 求和 多列", CommandRisk.ModifyDocument, GroupSum);
            Register(registry, "E-EDU-001-SUMMARY", "成绩统计", "成绩 分数 平均分 中位数 及格率 优秀率 分数段", CommandRisk.ModifyDocument, GradeSummary);
            Register(registry, "E-PAY-001", "生成工资条", "工资 工资条 表头", CommandRisk.ModifyDocument,
                (context, token) => Simple(() => _dataService.CreatePayroll(1), token));
            Register(registry, "E-ID-001-VALIDATE", "身份证校验", "身份证 校验 有效性 文本", CommandRisk.ReadOnly,
                (context, token) => Simple(_dataService.ValidateIdentityCards, token));
            Register(registry, "E-ID-002-MASK", "身份证脱敏", "身份证 脱敏 隐私 掩码", CommandRisk.ModifyDocument,
                (context, token) => Simple(_dataService.MaskIdentityCards, token));
        }

        private void Register(CommandRegistry registry, string id, string name, string keywords, CommandRisk risk,
            Func<CommandContext, CancellationToken, CommandResult> execute)
        {
            registry.Register(new DelegateSuiteCommand(
                id,
                name,
                keywords,
                SuiteHost.Excel,
                risk,
                false,
                context => _dataService.HasActiveWorkbook(),
                execute));
        }

        private static CommandResult Simple(Func<string> action, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            return CommandResult.Success(action());
        }

        private CommandResult Sort(CommandContext context, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            bool? hasHeader = ConfirmHeader("选区第一行是否为表头？");
            if (!hasHeader.HasValue) return CommandResult.Cancelled("已取消排序，选区未修改。");
            return CommandResult.Success(_dataService.SortSelection(hasHeader.Value));
        }

        private CommandResult Compare(CommandContext context, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            Excel.Range left = _dataService.GetSelectionRange();
            Excel.Range right = null;
            try
            {
                bool? hasHeader = ConfirmHeader("左右两个区域是否都以第一行为表头？");
                if (!hasHeader.HasValue) return CommandResult.Cancelled("已取消对比，原区域未修改。");
                int? keyColumn = RequestColumn("请输入左右区域中作为比较键的列号（从 1 开始）：");
                if (!keyColumn.HasValue) return CommandResult.Cancelled("已取消对比，原区域未修改。");
                object selected = _application.InputBox(
                    "当前选区作为左侧。请选择右侧对比区域：",
                    "按键列对比",
                    Type: 8);
                if (selected is bool && !(bool)selected) return CommandResult.Cancelled("已取消对比，原区域未修改。");
                right = selected as Excel.Range;
                if (right == null) return CommandResult.Failed("右侧选择不是有效单元格区域。");
                return CommandResult.Success(_dataService.CompareByKey(left, right, keyColumn.Value, hasHeader.Value));
            }
            finally
            {
                WorksheetDataService.ReleaseComObject(right);
                WorksheetDataService.ReleaseComObject(left);
            }
        }

        private CommandResult Summarize(CommandContext context, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            bool? hasHeader = ConfirmHeader("各工作表是否具有一致的第一行表头？选择“否”将汇总所有行并生成通用列名。");
            if (!hasHeader.HasValue) return CommandResult.Cancelled("已取消汇总，工作簿未修改。");
            return CommandResult.Success(_dataService.SummarizeWorksheets(hasHeader.Value));
        }

        private CommandResult Split(CommandContext context, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            int? keyColumn = RequestColumn("当前选区首行必须是表头。请输入拆分依据列号（从 1 开始）：");
            if (!keyColumn.HasValue) return CommandResult.Cancelled("已取消拆分，源工作簿未修改。");
            string outputPath;
            string message = _dataService.SplitSelectionToNewWorkbook(keyColumn.Value, out outputPath);
            return CommandResult.Success(message, outputPath);
        }

        private CommandResult GroupSum(CommandContext context, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            int? keyColumn = RequestColumn("当前选区首行必须是表头。请输入分组键列号（从 1 开始）：");
            if (!keyColumn.HasValue) return CommandResult.Cancelled("已取消分组求和，源区域未修改。");
            int? sumColumn = RequestColumn("请输入要求和的数值列号（从 1 开始）：");
            if (!sumColumn.HasValue) return CommandResult.Cancelled("已取消分组求和，源区域未修改。");
            token.ThrowIfCancellationRequested();
            return CommandResult.Success(_dataService.SummarizeSelectionByColumn(keyColumn.Value, sumColumn.Value));
        }

        private CommandResult GradeSummary(CommandContext context, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            double? maximum = RequestNumber("请输入满分（例如 100）：");
            if (!maximum.HasValue) return CommandResult.Cancelled("已取消成绩统计，原区域未修改。");
            double? pass = RequestNumber("请输入及格线（例如 60）：");
            if (!pass.HasValue) return CommandResult.Cancelled("已取消成绩统计，原区域未修改。");
            double? excellent = RequestNumber("请输入优秀线（例如 90）：");
            if (!excellent.HasValue) return CommandResult.Cancelled("已取消成绩统计，原区域未修改。");
            token.ThrowIfCancellationRequested();
            return CommandResult.Success(_dataService.CreateGradeStatistics(maximum.Value, pass.Value, excellent.Value));
        }

        private double? RequestNumber(string prompt)
        {
            object result = _application.InputBox(prompt, "办公增效套件", Type: 1);
            if (result is bool && !(bool)result) return null;
            double value;
            try
            {
                value = Convert.ToDouble(result, CultureInfo.CurrentCulture);
            }
            catch (Exception ex) when (ex is FormatException || ex is InvalidCastException || ex is OverflowException)
            {
                throw new InvalidOperationException("请输入有效的有限数值。", ex);
            }
            if (double.IsNaN(value) || double.IsInfinity(value)) throw new InvalidOperationException("请输入有效的有限数值。");
            return value;
        }

        private int? RequestColumn(string prompt)
        {
            object result = _application.InputBox(prompt, "办公增效套件", Type: 1);
            if (result is bool && !(bool)result) return null;
            double value;
            try
            {
                value = Convert.ToDouble(result, CultureInfo.CurrentCulture);
            }
            catch (Exception ex) when (ex is FormatException || ex is InvalidCastException || ex is OverflowException)
            {
                throw new InvalidOperationException("列号必须是大于等于 1 的整数。", ex);
            }
            if (value < 1 || value > int.MaxValue || Math.Truncate(value) != value)
                throw new InvalidOperationException("列号必须是大于等于 1 的整数。");
            return (int)value;
        }

        private static bool? ConfirmHeader(string message)
        {
            DialogResult result = MessageBox.Show(message, "办公增效套件", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (result == DialogResult.Yes) return true;
            if (result == DialogResult.No) return false;
            return null;
        }
    }
}
