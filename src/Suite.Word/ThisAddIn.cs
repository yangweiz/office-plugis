using System;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using Suite.Core.Commands;
using Suite.Infrastructure;
using Suite.Windowing;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace Suite.Word
{
    public sealed partial class ThisAddIn
    {
        private readonly WordCommands _commands = new WordCommands();
        private readonly CommandExecutor _executor = new CommandExecutor();
        private SuiteSettings _settings;
        private JsonSettingsStore<SuiteSettings> _settingsStore;
        private SuiteLogger _logger;
        private Microsoft.Office.Tools.CustomTaskPane _documentPane;
        private DocumentSessionService _documentSessions;
        private DocumentSwitcherControl _documentSwitcher;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            string root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeProductivitySuite");
            _settingsStore = new JsonSettingsStore<SuiteSettings>(Path.Combine(root, "settings.json"));
            _settings = _settingsStore.Load();
            _logger = new SuiteLogger(Path.Combine(root, "logs", "word.log"));
            _logger.Info("Word 加载项已启动。版本 " + Assembly.GetExecutingAssembly().GetName().Version);
            if (_settingsStore.LastError != null) _logger.Error("设置已回退到默认值。", new InvalidDataException(_settingsStore.LastError));
            if (_settings.DocumentColors == null) _settings.DocumentColors = new System.Collections.Generic.Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            _documentSessions = new DocumentSessionService(new WordDocumentSessionAdapter(Application), _settings.DocumentColors, SaveSettings);
            _documentSwitcher = new DocumentSwitcherControl(_documentSessions);
            _documentPane = CustomTaskPanes.Add(_documentSwitcher, "文档管理");
            _documentPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            _documentPane.Width = 300;
            Application.DocumentOpen += Application_DocumentOpen;
            Application.NewDocument += Application_NewDocument;
            Application.DocumentBeforeClose += Application_DocumentBeforeClose;
            Application.DocumentBeforeSave += Application_DocumentBeforeSave;
            Application.WindowActivate += Application_WindowActivate;
            _documentSessions.Refresh();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Application.DocumentOpen -= Application_DocumentOpen;
            Application.NewDocument -= Application_NewDocument;
            Application.DocumentBeforeClose -= Application_DocumentBeforeClose;
            Application.DocumentBeforeSave -= Application_DocumentBeforeSave;
            Application.WindowActivate -= Application_WindowActivate;
            if (_documentPane != null) _documentPane.Dispose();
            if (_logger != null) _logger.Info("Word 加载项正在卸载。");
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new WordRibbon(this);
        }

        internal void ExecuteCommand(string id)
        {
            ISuiteCommand command = _commands.Registry.Find(id);
            if (command == null)
            {
                MessageBox.Show("未找到功能：" + id, "办公增效", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            Word.Document document = null;
            try
            {
                document = GetActiveDocument();
                var context = new CommandContext(SuiteHost.Word, Application, document, null, null);
                CommandResult result = null;
                Action run = delegate { result = _executor.Execute(command, context, CancellationToken.None); };
                if (command.CanUndo && document != null) _commands.ExecuteWithUndo(Application, command.DisplayName, run);
                else run();
                if (result != null && result.Status == CommandStatus.Failed && command.CanUndo && document != null)
                    result = CommandResult.Failed(result.Message + "\n命令可能已部分修改文档，请检查后立即使用 Word 撤销。");
                ShowResult(result);
            }
            catch (Exception ex)
            {
                if (_logger != null) _logger.Error("命令执行异常：" + id, ex);
                MessageBox.Show("操作未完成：" + ex.Message, "办公增效", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (document != null && System.Runtime.InteropServices.Marshal.IsComObject(document))
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(document);
            }
        }

        internal void ShowCommandSearch(string query)
        {
            using (var form = new CommandSearchForm(_commands.Registry, SuiteHost.Word, query))
            {
                if (form.ShowDialog() != DialogResult.OK || form.SelectedCommand == null) return;
                ExecuteCommand(form.SelectedCommand.Id);
            }
        }

        internal void ToggleDocumentPane()
        {
            if (_documentPane != null) _documentPane.Visible = !_documentPane.Visible;
        }

        internal void ShowNativeTabStatus()
        {
            MessageBox.Show(WindowIntegration.GetNativeTabStatus().Message, "原生标签状态", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        internal void ShowSettings()
        {
            EnsureInitialized();
            using (var dialog = new FolderBrowserDialog { Description = "选择批处理和导出文件的默认输出目录", SelectedPath = _settings.OutputDirectory, ShowNewFolderButton = true })
            {
                if (dialog.ShowDialog() != DialogResult.OK) return;
                _settings.OutputDirectory = dialog.SelectedPath;
                SaveSettings();
            }
            MessageBox.Show("设置已保存。服务连接凭据仅通过 DPAPI 凭据存储使用，不写入此设置文件。", "办公增效", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        internal void ShowAbout()
        {
            MessageBox.Show(
                "办公增效套件 " + Assembly.GetExecutingAssembly().GetName().Version + "\n\n" +
                "原创 clean-room 实现。页面与样式参数需结合单位模板和打印样张核验。\n" +
                "修改类命令合并为一次 Word 撤销；批处理默认不覆盖原文件。\n" +
                "日志位置：%LOCALAPPDATA%\\OfficeProductivitySuite\\logs",
                "帮助与版本", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void Application_DocumentOpen(Word.Document document) { QueueDocumentRefresh(); }
        private void Application_NewDocument(Word.Document document) { QueueDocumentRefresh(); }
        private void Application_DocumentBeforeClose(Word.Document document, ref bool cancel) { QueueDocumentRefresh(); }
        private void Application_DocumentBeforeSave(Word.Document document, ref bool saveAsUi, ref bool cancel) { QueueDocumentRefresh(); }
        private void Application_WindowActivate(Word.Document document, Word.Window window) { QueueDocumentRefresh(); }

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

        private void SaveSettings()
        {
            if (_settingsStore != null && _settings != null) _settingsStore.Save(_settings);
        }

        private Word.Document GetActiveDocument()
        {
            Word.Documents documents = Application.Documents;
            try { return documents.Count == 0 ? null : Application.ActiveDocument; }
            finally { if (System.Runtime.InteropServices.Marshal.IsComObject(documents)) System.Runtime.InteropServices.Marshal.FinalReleaseComObject(documents); }
        }

        private void EnsureInitialized()
        {
            if (_settings != null) return;
            string root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeProductivitySuite");
            _settingsStore = new JsonSettingsStore<SuiteSettings>(Path.Combine(root, "settings.json"));
            _settings = _settingsStore.Load();
            _logger = new SuiteLogger(Path.Combine(root, "logs", "word.log"));
        }

        private static void ShowResult(CommandResult result)
        {
            if (result == null) return;
            MessageBoxIcon icon = result.Status == CommandStatus.Succeeded ? MessageBoxIcon.Information :
                result.Status == CommandStatus.Cancelled ? MessageBoxIcon.Information : MessageBoxIcon.Warning;
            string message = result.Message;
            if (!string.IsNullOrWhiteSpace(result.OutputPath)) message += "\r\n输出文件：" + result.OutputPath;
            MessageBox.Show(message, "办公增效", MessageBoxButtons.OK, icon);
        }
    }
}
