using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using Suite.Core.Commands;
using Suite.Word.Services;
using Word = Microsoft.Office.Interop.Word;

namespace Suite.Word
{
    internal sealed class WordCommands
    {
        private readonly CommandRegistry _registry = new CommandRegistry();
        private readonly OfficialDocumentService _official = new OfficialDocumentService();
        private readonly WordTextEditingService _text = new WordTextEditingService();
        private readonly OfficialDocumentPartService _parts = new OfficialDocumentPartService();
        private readonly DocumentBatchService _batch = new DocumentBatchService();
        private readonly TableAndLabelService _tablesAndLabels = new TableAndLabelService();

        public CommandRegistry Registry { get { return _registry; } }

        public WordCommands()
        {
            Register("W-PAGE-001", "公文页面设置", "A4 页边距 版心 公文", true, (document, token) =>
            {
                _official.ApplyPageSetup(document, false);
                return CommandResult.Success("已对全部节应用 A4 公文页面参数。请结合样张核验打印效果。");
            });
            Register("W-PAGE-002", "设置 22 行 28 字版面", "文档网格 行 字", true, (document, token) =>
            {
                _official.ApplyPageSetup(document, true);
                return CommandResult.Success("已设置文档网格、页面和页边距；字体与打印机可能影响最终分页。");
            });
            Register("W-PAGE-004", "当前页转为横向独立节", "横页 横向 分节", true, (document, token) =>
            {
                _official.MakeCurrentPageLandscape(document);
                return CommandResult.Success("已在当前页前后插入分节符并设置横向，请检查页码和空白页。");
            });
            Register("W-PAGE-007", "插入居中公文页码", "页码 PAGE 域 一字线", true, (document, token) =>
            {
                _official.InsertPageNumbers(document);
                return CommandResult.Success("已插入页码；已有 PAGE 域的节未重复插入。");
            });
            Register("W-PAGE-008", "清理全部页眉页脚", "页眉 页脚 横线 清理", true, (document, token) =>
            {
                _official.ClearHeadersAndFooters(document);
                return CommandResult.Success("已清理全部节的首页、奇偶页页眉页脚及其中的横线。");
            });
            Register("W-STYLE-001", "应用公文正文样式", "正文 仿宋 三号 样式", true, (document, token) =>
            {
                _official.ApplyBodyToSelection(document);
                return CommandResult.Success("已将选区段落应用为公文正文样式。");
            });
            Register("W-STYLE-002", "自动识别并排版正文标题", "一键排版 标题 正文", true, (document, token) =>
            {
                return CommandResult.Success(_official.AutoFormat(document));
            });
            Register("W-STYLE-003", "选区标题提级", "标题 提级 序号", true, (document, token) =>
            {
                int changed = _official.PromoteSelectionHeadings(document);
                return changed == 0 ? CommandResult.Blocked("选区中没有可按规则提级的编号标题。") : CommandResult.Success("已提级 " + changed + " 个标题。");
            });
            Register("W-TEXT-001", "选区断行重排", "断行 换行 段落 重排", true, (document, token) =>
            {
                return CommandResult.Success(_text.ReflowSelection(document));
            });
            Register("W-TEXT-002", "清理选区空白和不可见字符", "空格 空段 不可见字符 清理", true, (document, token) =>
            {
                return CommandResult.Success(_text.CleanWhitespaceAndInvisible(document));
            });
            Register("W-TEXT-003", "清理选区超链接和直接格式", "超链接 格式 清理", true, (document, token) =>
            {
                int count = _text.UnlinkAndClearDirectFormatting(document);
                return CommandResult.Success("已解除 " + count + " 个超链接域，并将选区恢复为段落样式格式。");
            });
            Register("W-TEXT-005-CN", "英文标点转中文标点", "标点 中英文 转换 URL 邮箱 保护", true, (document, token) =>
            {
                return CommandResult.Success(_text.ConvertSelectionPunctuation(document, true));
            });
            Register("W-TEXT-005-EN", "中文标点转英文标点", "标点 中英文 转换 URL 邮箱 保护", true, (document, token) =>
            {
                return CommandResult.Success(_text.ConvertSelectionPunctuation(document, false));
            });
            Register("W-TEXT-005-FULL", "选区半角转全角", "全角 半角 字符 宽度 URL 邮箱 保护", true, (document, token) =>
            {
                return CommandResult.Success(_text.ConvertSelectionWidth(document, true));
            });
            Register("W-TEXT-005-HALF", "选区全角转半角", "全角 半角 字符 宽度 URL 邮箱 保护", true, (document, token) =>
            {
                return CommandResult.Success(_text.ConvertSelectionWidth(document, false));
            });
            Register("W-PART-001", "插入文字红头与红线", "红头 红线 发文机关", true, (document, token) =>
            {
                string agency = Interaction.InputBox("输入发文机关名称：", "文字红头与红线");
                if (string.IsNullOrWhiteSpace(agency)) return CommandResult.Cancelled("未输入红头文字，操作已取消。");
                _parts.InsertTextRedHead(document, agency);
                return CommandResult.Success("已插入原创文字红头和红线，可在撤销后调整位置或字号。");
            });
            Register("W-PART-004", "插入当前公文日期", "日期 落款 年月日", true, (document, token) =>
            {
                _parts.InsertOfficialDate(document);
                return CommandResult.Success("已插入当前日期。");
            });
            RegisterExtendedCommands();
        }

        private void RegisterExtendedCommands()
        {
            RegisterContext("W-DOC-001-MERGE", "合并 Word 文档", "文档 合并 追加 批量 DOCX", CommandRisk.ModifyFiles, false, false, (context, token) =>
            {
                string[] inputs;
                using (var dialog = new OpenFileDialog
                {
                    Title = "选择至少两个要合并的 Word 文档",
                    Filter = "Word 文档 (*.docx;*.docm;*.dotx)|*.docx;*.docm;*.dotx",
                    Multiselect = true,
                    CheckFileExists = true,
                    ValidateNames = true
                })
                {
                    if (dialog.ShowDialog() != DialogResult.OK) return CommandResult.Cancelled("未选择合并输入文件，操作已取消。");
                    inputs = dialog.FileNames;
                }
                if (inputs.Length < 2) return CommandResult.Blocked("合并文档至少需要选择两个文件。");

                string requestedOutput;
                using (var dialog = new SaveFileDialog
                {
                    Title = "选择合并文档输出位置（不会覆盖已有文件）",
                    Filter = "Word 文档 (*.docx)|*.docx",
                    DefaultExt = "docx",
                    AddExtension = true,
                    FileName = "合并文档.docx",
                    OverwritePrompt = false,
                    ValidateNames = true
                })
                {
                    if (dialog.ShowDialog() != DialogResult.OK) return CommandResult.Cancelled("未选择合并输出位置，操作已取消。");
                    requestedOutput = dialog.FileName;
                }

                token.ThrowIfCancellationRequested();
                string output = _batch.MergeDocuments((Word.Application)context.Application, inputs, requestedOutput);
                return CommandResult.Success("已按选择顺序合并 " + inputs.Length + " 个文档；宏内容不会写入无宏 DOCX 输出。", output);
            });

            RegisterContext("W-DOC-001-SPLIT-H1", "按一级标题拆分文档", "文档 拆分 一级标题 批量 DOCX", CommandRisk.ModifyFiles, false, true, (context, token) =>
            {
                var document = (Word.Document)context.ActiveDocument;
                string outputDirectory;
                using (var dialog = new FolderBrowserDialog
                {
                    Description = "选择一级标题拆分文档的输出目录；源文档不会修改，已有文件不会覆盖。",
                    SelectedPath = document.Path,
                    ShowNewFolderButton = true
                })
                {
                    if (dialog.ShowDialog() != DialogResult.OK) return CommandResult.Cancelled("未选择拆分输出目录，操作已取消。");
                    outputDirectory = dialog.SelectedPath;
                }

                token.ThrowIfCancellationRequested();
                DocumentSplitResult split = _batch.SplitByLevelOneHeadings(document, outputDirectory);
                string message = "拆分结果：成功 " + split.OutputPaths.Count + " 项，失败 " + split.FailedItems.Count + " 项；源文档未修改。";
                if (split.OutputPaths.Count > 0)
                    message += "\r\n已生成文件：\r\n" + string.Join("\r\n", split.OutputPaths);
                if (split.FailedItems.Count > 0)
                    message += "\r\n失败项：\r\n" + string.Join("\r\n", split.FailedItems);
                if (split.SkippedItems.Count > 0)
                    message += "\r\n跳过与注意事项：\r\n" + string.Join("\r\n", split.SkippedItems);
                if (split.OutputPaths.Count == 0) return CommandResult.Failed(message);
                if (split.FailedItems.Count > 0 || split.SkippedItems.Count > 0) return CommandResult.Partial(message);
                return CommandResult.Success(message);
            });

            RegisterContext("W-DOC-003-PDF", "导出当前文档为 PDF", "PDF 导出 固定格式", CommandRisk.ModifyFiles, false, true, (context, token) =>
            {
                var document = (Word.Document)context.ActiveDocument;
                using (var dialog = new SaveFileDialog
                {
                    Title = "选择 PDF 输出位置（不会覆盖已有文件）",
                    Filter = "PDF 文档 (*.pdf)|*.pdf",
                    DefaultExt = "pdf",
                    AddExtension = true,
                    FileName = Path.GetFileNameWithoutExtension(document.Name) + ".pdf",
                    InitialDirectory = document.Path,
                    OverwritePrompt = false,
                    ValidateNames = true
                })
                {
                    if (dialog.ShowDialog() != DialogResult.OK) return CommandResult.Cancelled("未选择 PDF 输出位置，操作已取消。");
                    token.ThrowIfCancellationRequested();
                    string output = _batch.ExportPdf(document, dialog.FileName);
                    return CommandResult.Success("已导出当前文档的全部页面为 PDF。", output);
                }
            });

            Register("W-TBL-001-FIT", "选区表格适应窗口", "表格 自动调整 适应窗口 列宽", true, (document, token) =>
            {
                _tablesAndLabels.AutoFitSelectionTableToWindow(document);
                return CommandResult.Success("已将当前选区涉及的单个表格调整为适应窗口。");
            });
            Register("W-TBL-001-ROWS", "均分选区表格行", "表格 均分行 行高", true, (document, token) =>
            {
                _tablesAndLabels.DistributeSelectionRows(document);
                return CommandResult.Success("已均分当前选区涉及的单个表格的行高。");
            });
            Register("W-TBL-001-COLS", "均分选区表格列", "表格 均分列 列宽", true, (document, token) =>
            {
                _tablesAndLabels.DistributeSelectionColumns(document);
                return CommandResult.Success("已均分当前选区涉及的单个表格的列宽。");
            });

            RegisterContext("W-LABEL-001-SEAT", "批量生成桌签", "桌签 座牌 姓名 名单 标签", CommandRisk.ModifyFiles, false, false, (context, token) =>
            {
                string input = Interaction.InputBox("输入姓名，使用逗号、顿号、分号或竖线分隔：\r\n例如：张三、李四、王五", "批量生成桌签");
                if (string.IsNullOrWhiteSpace(input)) return CommandResult.Cancelled("未输入姓名，操作已取消。");
                string[] names = input.Split(new[] { ',', '，', '、', ';', '；', '|', '｜', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                names = Array.FindAll(names, name => !string.IsNullOrWhiteSpace(name));
                if (names.Length == 0) return CommandResult.Blocked("至少需要一个非空姓名。");
                if (names.Length > 500) return CommandResult.Blocked("一次最多生成 500 个桌签，请分批处理。");

                string columnText = Interaction.InputBox("输入每行桌签列数（1 至 6）：", "批量生成桌签", "2");
                if (string.IsNullOrWhiteSpace(columnText)) return CommandResult.Cancelled("未输入桌签列数，操作已取消。");
                int columns;
                if (!int.TryParse(columnText, out columns) || columns < 1 || columns > 6)
                    return CommandResult.Blocked("桌签列数必须是 1 至 6 的整数，未生成任何文件。");

                using (var dialog = new SaveFileDialog
                {
                    Title = "选择桌签 DOCX 输出位置（不会覆盖已有文件）",
                    Filter = "Word 文档 (*.docx)|*.docx",
                    DefaultExt = "docx",
                    AddExtension = true,
                    FileName = "桌签.docx",
                    OverwritePrompt = false,
                    ValidateNames = true
                })
                {
                    if (dialog.ShowDialog() != DialogResult.OK) return CommandResult.Cancelled("未选择桌签输出位置，操作已取消。");
                    token.ThrowIfCancellationRequested();
                    string output = _tablesAndLabels.CreateSeatLabels((Word.Application)context.Application, names, columns, dialog.FileName);
                    return CommandResult.Success("已生成 " + names.Length + " 个桌签，每行 " + columns + " 列；请打印前核验纸张、字号和分页。", output);
                }
            });
        }

        public void ExecuteWithUndo(Word.Application application, string name, Action action)
        {
            Word.UndoRecord undo = application.UndoRecord;
            bool started = false;
            try
            {
                undo.StartCustomRecord(name);
                started = true;
                action();
            }
            finally
            {
                if (started) undo.EndCustomRecord();
                if (System.Runtime.InteropServices.Marshal.IsComObject(undo)) System.Runtime.InteropServices.Marshal.FinalReleaseComObject(undo);
            }
        }

        private void Register(string id, string name, string keywords, bool undo, Func<Word.Document, CancellationToken, CommandResult> execute)
        {
            RegisterContext(id, name, keywords, CommandRisk.ModifyDocument, undo, true,
                (context, token) => execute((Word.Document)context.ActiveDocument, token));
        }

        private void RegisterContext(string id, string name, string keywords, CommandRisk risk, bool undo, bool requiresDocument,
            Func<CommandContext, CancellationToken, CommandResult> execute)
        {
            _registry.Register(new DelegateSuiteCommand(
                id, name, keywords, SuiteHost.Word, risk, undo,
                context => !requiresDocument || context.ActiveDocument is Word.Document,
                (context, token) =>
                {
                    token.ThrowIfCancellationRequested();
                    return execute(context, token);
                }));
        }
    }
}
