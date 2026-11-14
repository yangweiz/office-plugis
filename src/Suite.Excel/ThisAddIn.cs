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
using Suite.Excel.Commands;
using Suite.Excel.Services;
using Suite.Infrastructure;
using Suite.Windowing;
using Excel = Microsoft.Office.Interop.Excel;

namespace Suite.Excel
{
    public sealed partial class ThisAddIn : IRibbonExtensibility
    {
        private CommandRegistry _registry;
        private CommandExecutor _executor;
        private WorksheetDataService _dataService;
        private ExcelCommandSet _commandSet;
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
                _logger = new SuiteLogger(Path.Combine(root, "logs", "excel.log"));
                _settingsStore = new JsonSettingsStore<SuiteSettings>(Path.Combine(root, "settings.json"));
                _settings = _settingsStore.Load();
                if (_settings.DocumentColors == null) _settings.DocumentColors = new System.Collections.Generic.Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                _registry = new CommandRegistry();
                _executor = new CommandExecutor();
                _dataService = new WorksheetDataService(Application);
                _commandSet = new ExcelCommandSet(Application, _dataService);
                _commandSet.RegisterAll(_registry);
                _documentSessions = new DocumentSessionService(new ExcelDocumentSessionAdapter(Application), _settings.DocumentColors, SaveSettings);
                _documentSwitcher = new DocumentSwitcherControl(_documentSessions);
                _documentPane = CustomTaskPanes.Add(_documentSwitcher, "文档管理");
                _documentPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                _documentPane.Width = 300;
                Application.WorkbookOpen += Application_WorkbookOpen;
                Application.NewWorkbook += Application_NewWorkbook;
                Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;
                Application.WorkbookAfterSave += Application_WorkbookAfterSave;
                Application.WindowActivate += Application_WindowActivate;
                _documentSessions.Refresh();
            }
            catch (Exception ex)
            {
                if (_logger != null) _logger.Error("Excel 加载项启动失败。", ex);
                MessageBox.Show("办公增效 Excel 启动失败：" + ex.Message, "办公增效", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Application.WorkbookOpen -= Application_WorkbookOpen;
            Application.NewWorkbook -= Application_NewWorkbook;
            Application.WorkbookBeforeClose -= Application_WorkbookBeforeClose;
            Application.WorkbookAfterSave -= Application_WorkbookAfterSave;
            Application.WindowActivate -= Application_WindowActivate;
            if (_documentPane != null) _documentPane.Dispose();
            if (_logger != null) _logger.Info("Excel 加载项正在卸载。");
            _registry = null;
            _executor = null;
            _commandSet = null;
            _dataService = null;
            _uiContext = null;
        }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject() { return this; }

        public string GetCustomUI(string ribbonId)
        {
            if (!string.Equals(ribbonId, "Microsoft.Excel.Workbook", StringComparison.Ordinal)) return null;
            using (Stream stream = typeof(ThisAddIn).Assembly.GetManifestResourceStream("Suite.Excel.Ribbon.SuiteRibbon.xml"))
            {
                if (stream == null) throw new InvalidOperationException("未找到嵌入式 SuiteRibbon.xml 资源。");
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

        public void OnSearchChanged(IRibbonControl control, string text) { _searchText = text ?? string.Empty; }
        public void OnSearch(IRibbonControl control) { ShowCommandSearch(); }
        public void OnDocuments(IRibbonControl control) { if (_documentPane != null) _documentPane.Visible = !_documentPane.Visible; }
        public void OnNativeTabs(IRibbonControl control) { MessageBox.Show(WindowIntegration.GetNativeTabStatus().Message, "原生标签状态", MessageBoxButtons.OK, MessageBoxIcon.Information); }
        public void OnSettings(IRibbonControl control) { ShowSettings(); }
        public void OnAbout(IRibbonControl control)
        {
            MessageBox.Show("办公增效套件 " + Assembly.GetExecutingAssembly().GetName().Version + "\n\nExcel 修改命令执行前创建独立备份，不宣称支持单次撤销。\n日志位置：%LOCALAPPDATA%\\OfficeProductivitySuite\\logs", "帮助与版本", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void RunOnUiThread(Action action)
        {
            if (Thread.CurrentThread.ManagedThreadId == _uiThreadId) { action(); return; }
            if (_uiContext == null) throw new InvalidOperationException("Excel UI 线程尚未就绪。");
            Exception error = null;
            _uiContext.Send(delegate { try { action(); } catch (Exception ex) { error = ex; } }, null);
            if (error != null) ExceptionDispatchInfo.Capture(error).Throw();
        }

        private void ExecuteCommand(string commandId)
        {
            Excel.Workbook workbook = null;
            try
            {
                if (_registry == null || _executor == null) { MessageBox.Show("加载项尚未完成启动。", "办公增效", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
                workbook = Application.ActiveWorkbook;
                var context = new CommandContext(SuiteHost.Excel, Application, workbook, null, null);
                ISuiteCommand command = _registry.Find(commandId);
                string backupPath = null;
                if (command != null && command.Risk == CommandRisk.ModifyDocument && workbook != null && command.CanExecute(context))
                {
                    string error;
                    if (!TryCreateBackup(workbook, out backupPath, out error)) { ShowResult(CommandResult.Blocked(error)); return; }
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
                Release(workbook);
            }
        }

        private void ShowResult(CommandResult result, string backupPath = null)
        {
            if (result == null) return;
            if (result.Status == CommandStatus.Failed && _logger != null) _logger.Error(result.Message, null);
            MessageBoxIcon icon = result.Status == CommandStatus.Failed ? MessageBoxIcon.Error : result.Status == CommandStatus.Blocked || result.Status == CommandStatus.Partial ? MessageBoxIcon.Warning : MessageBoxIcon.Information;
            string message = result.Message;
            if (!string.IsNullOrWhiteSpace(result.OutputPath)) message += Environment.NewLine + "输出位置：" + result.OutputPath;
            if (!string.IsNullOrWhiteSpace(backupPath)) message += Environment.NewLine + "修改前备份：" + backupPath;
            MessageBox.Show(message, "办公增效", MessageBoxButtons.OK, icon);
        }

        private static bool TryCreateBackup(Excel.Workbook workbook, out string backupPath, out string error)
        {
            backupPath = null;
            error = null;
            if (string.IsNullOrWhiteSpace(workbook.Path))
            {
                error = "请先将工作簿保存到本地路径，以便在修改前创建独立备份。";
                return false;
            }
            try
            {
                string directory = Path.Combine(workbook.Path, "办公增效备份");
                Directory.CreateDirectory(directory);
                string name = Path.GetFileNameWithoutExtension(workbook.Name) + "_备份_" + DateTime.Now.ToString("yyyyMMdd_HHmmss_fff") + Path.GetExtension(workbook.Name);
                backupPath = OutputPathPolicy.GetAvailablePath(Path.Combine(directory, name));
                workbook.SaveCopyAs(backupPath);
                return true;
            }
            catch (Exception ex)
            {
                error = "修改前备份失败，命令未执行：" + ex.Message;
                return false;
            }
        }

        private void ShowCommandSearch()
        {
            using (var form = new CommandSearchForm(_registry, SuiteHost.Excel, _searchText))
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

        private void Application_WorkbookOpen(Excel.Workbook workbook) { QueueDocumentRefresh(); }
        private void Application_NewWorkbook(Excel.Workbook workbook) { QueueDocumentRefresh(); }
        private void Application_WorkbookBeforeClose(Excel.Workbook workbook, ref bool cancel) { QueueDocumentRefresh(); }
        private void Application_WorkbookAfterSave(Excel.Workbook workbook, bool success) { QueueDocumentRefresh(); }
        private void Application_WindowActivate(Excel.Workbook workbook, Excel.Window window) { QueueDocumentRefresh(); }

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
