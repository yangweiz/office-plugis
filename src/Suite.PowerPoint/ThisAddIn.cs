using System;
using System.IO;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Suite.Core.Batching;
using Suite.Core.Commands;
using Suite.Infrastructure;
using Suite.PowerPoint.Services;
using Suite.Windowing;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Suite.PowerPoint
{
    public sealed partial class ThisAddIn : IRibbonExtensibility
    {
        private CommandRegistry _registry;
        private CommandExecutor _executor;
        private PresentationCleanupService _cleanupService;
        private ShapeService _shapeService;
        private PresentationExportService _exportService;
        private SuiteLogger _logger;
        private SuiteSettings _settings;
        private JsonSettingsStore<SuiteSettings> _settingsStore;
        private SynchronizationContext _uiContext;
        private int _uiThreadId;
        private string _searchText = string.Empty;
        private Microsoft.Office.Tools.CustomTaskPane _documentPane;
        private DocumentSessionService _documentSessions;
        private DocumentSwitcherControl _documentSwitcher;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                _uiThreadId = Thread.CurrentThread.ManagedThreadId;
                _uiContext = SynchronizationContext.Current as WindowsFormsSynchronizationContext ?? new WindowsFormsSynchronizationContext();
                string root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeProductivitySuite");
                _logger = new SuiteLogger(Path.Combine(root, "logs", "powerpoint.log"));
                _settingsStore = new JsonSettingsStore<SuiteSettings>(Path.Combine(root, "settings.json"));
                _settings = _settingsStore.Load();
                if (_settings.DocumentColors == null) _settings.DocumentColors = new System.Collections.Generic.Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                _registry = new CommandRegistry();
                _executor = new CommandExecutor();
                _cleanupService = new PresentationCleanupService();
                _shapeService = new ShapeService();
                _exportService = new PresentationExportService();
                RegisterCommands();
                _documentSessions = new DocumentSessionService(new PowerPointDocumentSessionAdapter(Application), _settings.DocumentColors, SaveSettings);
                _documentSwitcher = new DocumentSwitcherControl(_documentSessions);
                _documentPane = CustomTaskPanes.Add(_documentSwitcher, "文档管理");
                _documentPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                _documentPane.Width = 300;
                Application.PresentationOpen += Application_PresentationOpen;
                Application.NewPresentation += Application_NewPresentation;
                Application.PresentationBeforeClose += Application_PresentationBeforeClose;
                Application.PresentationSave += Application_PresentationSave;
                Application.WindowActivate += Application_WindowActivate;
                _documentSessions.Refresh();
            }
            catch (Exception ex)
            {
                if (_logger != null) _logger.Error("PowerPoint 加载项启动失败。", ex);
                MessageBox.Show("办公增效 PowerPoint 启动失败：" + ex.Message, "办公增效", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Application.PresentationOpen -= Application_PresentationOpen;
            Application.NewPresentation -= Application_NewPresentation;
            Application.PresentationBeforeClose -= Application_PresentationBeforeClose;
            Application.PresentationSave -= Application_PresentationSave;
            Application.WindowActivate -= Application_WindowActivate;
            if (_documentPane != null) _documentPane.Dispose();
            if (_logger != null) _logger.Info("PowerPoint 加载项正在卸载。");
            _registry = null;
            _executor = null;
            _cleanupService = null;
            _shapeService = null;
            _exportService = null;
            _uiContext = null;
        }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject() { return this; }

        public string GetCustomUI(string ribbonId)
        {
            if (!string.Equals(ribbonId, "Microsoft.PowerPoint.Presentation", StringComparison.Ordinal)) return null;
            using (Stream stream = typeof(ThisAddIn).Assembly.GetManifestResourceStream("Suite.PowerPoint.Ribbon.xml"))
            {
                if (stream == null) throw new InvalidOperationException("未找到嵌入式 Ribbon.xml 资源。");
                using (var reader = new StreamReader(stream)) return reader.ReadToEnd();
            }
        }

        public void OnRibbonAction(IRibbonControl control)
        {
            try { RunOnUiThread(delegate { ExecuteCommand(control == null ? null : control.Tag); }); }
            catch (Exception ex)
            {
                if (_logger != null) _logger.Error("Ribbon 命令调度失败。", ex);
                MessageBox.Show("命令调度失败：" + ex.Message, "办公增效", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnSearchChanged(IRibbonControl control, string text) { _searchText = text; }
        public void OnSearch(IRibbonControl control) { ShowCommandSearch(); }
        public void OnDocuments(IRibbonControl control) { if (_documentPane != null) _documentPane.Visible = !_documentPane.Visible; }
        public void OnNativeTabs(IRibbonControl control) { MessageBox.Show(WindowIntegration.GetNativeTabStatus().Message, "原生标签状态", MessageBoxButtons.OK, MessageBoxIcon.Information); }
        public void OnSettings(IRibbonControl control) { ShowSettings(); }
        public void OnAbout(IRibbonControl control)
        {
            MessageBox.Show("办公增效套件 " + Assembly.GetExecutingAssembly().GetName().Version + "\n\nPowerPoint 修改命令执行前创建独立备份，不宣称支持单次撤销。\n日志位置：%LOCALAPPDATA%\\OfficeProductivitySuite\\logs", "帮助与版本", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void RunOnUiThread(Action action)
        {
            if (Thread.CurrentThread.ManagedThreadId == _uiThreadId) { action(); return; }
            if (_uiContext == null) throw new InvalidOperationException("PowerPoint UI 线程尚未就绪。");
            Exception error = null;
            _uiContext.Send(delegate { try { action(); } catch (Exception ex) { error = ex; } }, null);
            if (error != null) ExceptionDispatchInfo.Capture(error).Throw();
        }

        private void ExecuteCommand(string commandId)
        {
            PowerPoint.Presentation presentation = null;
            try
            {
                if (_registry == null || _executor == null) { MessageBox.Show("加载项尚未完成启动。", "办公增效", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
                presentation = GetActivePresentation();
                var context = new CommandContext(SuiteHost.PowerPoint, Application, presentation, null, null);
                ISuiteCommand command = _registry.Find(commandId);
                string backupPath = null;
                if (command != null && command.Risk == CommandRisk.ModifyDocument && presentation != null && command.CanExecute(context))
                {
                    string error;
                    if (!TryCreateBackup(presentation, out backupPath, out error)) { ShowResult(CommandResult.Blocked(error)); return; }
                }
                CommandResult result = _executor.Execute(command, context, CancellationToken.None);
                ShowResult(result, backupPath);
            }
            catch (Exception ex)
            {
                if (_logger != null) _logger.Error("执行 Ribbon 命令时发生未处理异常。", ex);
                MessageBox.Show("命令执行失败：" + ex.Message, "办公增效", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Release(presentation);
            }
        }

        private void ShowResult(CommandResult result, string backupPath = null)
        {
            if (result == null) return;
            MessageBoxIcon icon = result.Status == CommandStatus.Failed ? MessageBoxIcon.Error : result.Status == CommandStatus.Blocked || result.Status == CommandStatus.Partial ? MessageBoxIcon.Warning : MessageBoxIcon.Information;
            string message = result.Message;
            if (!string.IsNullOrWhiteSpace(result.OutputPath)) message += Environment.NewLine + "输出位置：" + result.OutputPath;
            if (!string.IsNullOrWhiteSpace(backupPath)) message += Environment.NewLine + "修改前备份：" + backupPath;
            MessageBox.Show(message, "办公增效", MessageBoxButtons.OK, icon);
        }

        private bool TryCreateBackup(PowerPoint.Presentation presentation, out string backupPath, out string error)
        {
            backupPath = null;
            error = null;
            if (string.IsNullOrWhiteSpace(presentation.Path))
            {
                error = "请先将演示文稿保存到本地路径，以便在修改前创建独立备份。";
                return false;
            }
            try
            {
                string directory = Path.Combine(presentation.Path, "办公增效备份");
                Directory.CreateDirectory(directory);
                string name = Path.GetFileNameWithoutExtension(presentation.Name) + "_备份_" + DateTime.Now.ToString("yyyyMMdd_HHmmss_fff") + Path.GetExtension(presentation.Name);
                backupPath = OutputPathPolicy.GetAvailablePath(Path.Combine(directory, name));
                presentation.SaveCopyAs(backupPath);
                return true;
            }
            catch (Exception ex)
            {
                error = "修改前备份失败，命令未执行：" + ex.Message;
                return false;
            }
        }

        private void RegisterCommands()
        {
            Register("P-CLEAN-001-BLANK", "清理空白页", "空白页 删除", CommandRisk.ModifyDocument, false, delegate(CommandContext context, CancellationToken token) { return _cleanupService.RemoveBlankSlides(GetPresentation(context)); });
            Register("P-CLEAN-001-HIDDEN", "清理隐藏页", "隐藏页 删除", CommandRisk.ModifyDocument, false, ConfirmAndRemoveHiddenSlides);
            Register("P-CLEAN-002-ANIMATION", "清理动画", "动画 清除", CommandRisk.ModifyDocument, false, delegate(CommandContext context, CancellationToken token) { return _cleanupService.RemoveAnimations(GetPresentation(context)); });
            Register("P-CLEAN-002-TRANSITION", "清理切换", "切换 自动换片", CommandRisk.ModifyDocument, false, delegate(CommandContext context, CancellationToken token) { return _cleanupService.RemoveTransitions(GetPresentation(context)); });
            Register("P-SHAPE-001-LEFT", "左对齐", "形状 对齐", CommandRisk.ModifyDocument, false, delegate(CommandContext context, CancellationToken token) { return ExecuteShapeCommand(_shapeService.AlignLeft); });
            Register("P-SHAPE-001-RIGHT", "右对齐", "形状 对齐", CommandRisk.ModifyDocument, false, delegate(CommandContext context, CancellationToken token) { return ExecuteShapeCommand(_shapeService.AlignRight); });
            Register("P-SHAPE-001-TOP", "上对齐", "形状 对齐", CommandRisk.ModifyDocument, false, delegate(CommandContext context, CancellationToken token) { return ExecuteShapeCommand(_shapeService.AlignTop); });
            Register("P-SHAPE-001-BOTTOM", "下对齐", "形状 对齐", CommandRisk.ModifyDocument, false, delegate(CommandContext context, CancellationToken token) { return ExecuteShapeCommand(_shapeService.AlignBottom); });
            Register("P-SHAPE-001-FLIPH", "水平翻转", "形状 翻转", CommandRisk.ModifyDocument, false, delegate(CommandContext context, CancellationToken token) { return ExecuteShapeCommand(_shapeService.FlipHorizontal); });
            Register("P-SHAPE-001-FLIPV", "垂直翻转", "形状 翻转", CommandRisk.ModifyDocument, false, delegate(CommandContext context, CancellationToken token) { return ExecuteShapeCommand(_shapeService.FlipVertical); });
            Register("P-SHAPE-001-WIDTH", "统一宽度", "形状 等宽 首个", CommandRisk.ModifyDocument, false, delegate(CommandContext context, CancellationToken token) { return ExecuteShapeCommand(_shapeService.UniformWidth); });
            Register("P-SHAPE-001-HEIGHT", "统一高度", "形状 等高 首个", CommandRisk.ModifyDocument, false, delegate(CommandContext context, CancellationToken token) { return ExecuteShapeCommand(_shapeService.UniformHeight); });
            Register("P-DESIGN-001-ACCENT1", "主题强调色 1", "形状 配色 主题", CommandRisk.ModifyDocument, false, delegate(CommandContext context, CancellationToken token) { return ExecuteShapeCommand(delegate(PowerPoint.Selection selection) { return _shapeService.ApplyThemeColor(selection, MsoThemeColorIndex.msoThemeColorAccent1, "主题强调色 1"); }); });
            Register("P-DESIGN-001-ACCENT2", "主题强调色 2", "形状 配色 主题", CommandRisk.ModifyDocument, false, delegate(CommandContext context, CancellationToken token) { return ExecuteShapeCommand(delegate(PowerPoint.Selection selection) { return _shapeService.ApplyThemeColor(selection, MsoThemeColorIndex.msoThemeColorAccent2, "主题强调色 2"); }); });
            Register("P-DESIGN-001-ACCENT3", "主题强调色 3", "形状 配色 主题", CommandRisk.ModifyDocument, false, delegate(CommandContext context, CancellationToken token) { return ExecuteShapeCommand(delegate(PowerPoint.Selection selection) { return _shapeService.ApplyThemeColor(selection, MsoThemeColorIndex.msoThemeColorAccent3, "主题强调色 3"); }); });
            Register("P-DOC-003-PNG", "逐页导出 PNG", "图片 导出 幻灯片", CommandRisk.ModifyFiles, false, ExportPresentationAsPng);
        }

        private void Register(string id, string displayName, string keywords, CommandRisk risk, bool canUndo, Func<CommandContext, CancellationToken, CommandResult> execute)
        {
            _registry.Register(new DelegateSuiteCommand(id, displayName, keywords, SuiteHost.PowerPoint, risk, canUndo, null, execute));
        }

        private static PowerPoint.Presentation GetPresentation(CommandContext context) { return context.ActiveDocument as PowerPoint.Presentation; }

        private CommandResult ExecuteShapeCommand(Func<PowerPoint.Selection, CommandResult> action)
        {
            PowerPoint.DocumentWindows windows = Application.Windows;
            PowerPoint.DocumentWindow window = null;
            PowerPoint.Selection selection = null;
            try
            {
                if (windows.Count == 0) return CommandResult.Blocked("请先打开演示文稿窗口并选择形状。");
                window = Application.ActiveWindow;
                selection = window.Selection;
                return action(selection);
            }
            finally { Release(selection); Release(window); Release(windows); }
        }

        private CommandResult ConfirmAndRemoveHiddenSlides(CommandContext context, CancellationToken token)
        {
            PowerPoint.Presentation presentation = GetPresentation(context);
            if (presentation == null) return CommandResult.Blocked("请先打开要清理的演示文稿。");
            DialogResult answer = MessageBox.Show("此命令会删除当前演示文稿中的全部隐藏页，且已创建修改前备份。是否继续？", "确认清理隐藏页", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            return answer == DialogResult.Yes ? _cleanupService.RemoveHiddenSlides(presentation) : CommandResult.Cancelled("已取消隐藏页清理，演示文稿未更改。");
        }

        private CommandResult ExportPresentationAsPng(CommandContext context, CancellationToken token)
        {
            PowerPoint.Presentation presentation = GetPresentation(context);
            if (presentation == null) return CommandResult.Blocked("请先打开要导出的演示文稿。");
            using (var dialog = new FolderBrowserDialog { Description = "选择逐页 PNG 图片的导出目录", ShowNewFolderButton = true })
            {
                if (dialog.ShowDialog() != DialogResult.OK) return CommandResult.Cancelled("未选择导出目录，操作已取消。");
                return _exportService.ExportSlidesAsPng(presentation, dialog.SelectedPath);
            }
        }

        private PowerPoint.Presentation GetActivePresentation()
        {
            PowerPoint.Presentations presentations = Application.Presentations;
            try { return presentations.Count == 0 ? null : Application.ActivePresentation; }
            finally { Release(presentations); }
        }

        private void ShowCommandSearch()
        {
            using (var form = new CommandSearchForm(_registry, SuiteHost.PowerPoint, _searchText))
            {
                if (form.ShowDialog() == DialogResult.OK && form.SelectedCommand != null) ExecuteCommand(form.SelectedCommand.Id);
            }
        }

        private void ShowSettings()
        {
            try
            {
                using (var dialog = new FolderBrowserDialog { Description = "选择默认输出目录", SelectedPath = _settings.OutputDirectory, ShowNewFolderButton = true })
                {
                    if (dialog.ShowDialog() != DialogResult.OK) return;
                    _settings.OutputDirectory = dialog.SelectedPath;
                    SaveSettings();
                }
            }
            catch (Exception ex) { MessageBox.Show("设置保存失败：" + ex.Message, "办公增效", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }

        private void SaveSettings() { if (_settingsStore != null && _settings != null) _settingsStore.Save(_settings); }

        private void Application_PresentationOpen(PowerPoint.Presentation presentation) { QueueDocumentRefresh(); }
        private void Application_NewPresentation(PowerPoint.Presentation presentation) { QueueDocumentRefresh(); }
        private void Application_PresentationBeforeClose(PowerPoint.Presentation presentation, ref bool cancel) { QueueDocumentRefresh(); }
        private void Application_PresentationSave(PowerPoint.Presentation presentation) { QueueDocumentRefresh(); }
        private void Application_WindowActivate(PowerPoint.DocumentWindow window, PowerPoint.Presentation presentation) { QueueDocumentRefresh(); }

        private void QueueDocumentRefresh()
        {
            if (_documentSwitcher == null || _documentSwitcher.IsDisposed || _documentSwitcher.Disposing) return;
            try
            {
                if (!_documentSwitcher.IsHandleCreated) _documentSwitcher.CreateControl();
                _documentSwitcher.BeginInvoke(new Action(RefreshDocumentSessions));
            }
            catch (InvalidOperationException) { }
        }

        private void RefreshDocumentSessions()
        {
            try { if (_documentSessions != null) _documentSessions.Refresh(); }
            catch (Exception ex) { if (_logger != null) _logger.Error("文档列表刷新失败。", ex); }
        }

        private static void Release(object value) { if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value); }
    }
}
