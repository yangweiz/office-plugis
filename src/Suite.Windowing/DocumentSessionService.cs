using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;

namespace Suite.Windowing
{
    public sealed class DocumentSessionService
    {
        private readonly IDocumentSessionAdapter _adapter;
        private readonly IDictionary<string, string> _persistentColors;
        private readonly Action _saveSettings;
        private readonly Dictionary<string, Color> _sessionColors = new Dictionary<string, Color>(StringComparer.OrdinalIgnoreCase);
        private readonly List<DocumentSession> _sessions = new List<DocumentSession>();

        public DocumentSessionService(IDocumentSessionAdapter adapter, IDictionary<string, string> persistentColors, Action saveSettings)
        {
            _adapter = adapter ?? throw new ArgumentNullException("adapter");
            _persistentColors = persistentColors ?? throw new ArgumentNullException("persistentColors");
            _saveSettings = saveSettings ?? throw new ArgumentNullException("saveSettings");
        }

        public event EventHandler Changed;
        public IReadOnlyList<DocumentSession> Sessions { get { return _sessions.AsReadOnly(); } }

        public void Refresh()
        {
            foreach (DocumentSession old in _sessions.Where(x => x.Color != Color.Empty && !string.IsNullOrEmpty(x.Key)))
                _sessionColors[old.Key] = old.Color;
            _sessions.Clear();
            foreach (DocumentSession session in _adapter.GetSessions() ?? new List<DocumentSession>())
            {
                string path = NormalizePath(session.FullPath);
                string stored;
                if (path.Length > 0 && _persistentColors.TryGetValue(path, out stored))
                {
                    try { session.Color = ColorTranslator.FromHtml(stored); }
                    catch (ArgumentException) { _persistentColors.Remove(path); }
                }
                Color transient;
                if (_sessionColors.TryGetValue(session.Key, out transient)) session.Color = transient;
                _sessions.Add(session);
            }
            Changed?.Invoke(this, EventArgs.Empty);
        }

        public void Activate(DocumentSession session) { Require(session); _adapter.Activate(session); }
        public void Save(DocumentSession session) { Require(session); _adapter.Save(session); Refresh(); }

        public int SaveAll()
        {
            int saved = 0;
            foreach (DocumentSession session in _sessions.Where(x => x.IsDirty).ToList())
            {
                _adapter.Save(session);
                saved++;
            }
            Refresh();
            return saved;
        }

        public bool Close(DocumentSession session, Func<DocumentSession, CloseDecision> confirm)
        {
            Require(session);
            bool closed = CloseCore(session, confirm);
            Refresh();
            return closed;
        }

        public int CloseAll(Func<DocumentSession, CloseDecision> confirm)
        {
            int closed = 0;
            foreach (DocumentSession session in _sessions.ToList())
            {
                if (!CloseCore(session, confirm)) break;
                closed++;
            }
            Refresh();
            return closed;
        }

        private bool CloseCore(DocumentSession session, Func<DocumentSession, CloseDecision> confirm)
        {
            bool save = false;
            if (session.IsDirty)
            {
                if (confirm == null) throw new InvalidOperationException("未保存文档必须提供保存、放弃或取消选择。");
                CloseDecision decision = confirm(session);
                if (decision == CloseDecision.Cancel) return false;
                save = decision == CloseDecision.Save;
            }
            _adapter.Close(session, save);
            return true;
        }

        public void SetColor(DocumentSession session, Color color)
        {
            Require(session);
            session.Color = color;
            _sessionColors[session.Key] = color;
            string path = NormalizePath(session.FullPath);
            if (path.Length > 0)
            {
                if (color == Color.Empty) _persistentColors.Remove(path);
                else _persistentColors[path] = "#" + color.R.ToString("X2", CultureInfo.InvariantCulture) + color.G.ToString("X2", CultureInfo.InvariantCulture) + color.B.ToString("X2", CultureInfo.InvariantCulture);
                _saveSettings();
            }
            Changed?.Invoke(this, EventArgs.Empty);
        }

        private static string NormalizePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return string.Empty;
            try { return System.IO.Path.GetFullPath(path); }
            catch (Exception ex) when (ex is ArgumentException || ex is NotSupportedException || ex is System.IO.PathTooLongException) { return string.Empty; }
        }

        private void Require(DocumentSession session)
        {
            if (session == null || !_sessions.Any(x => ReferenceEquals(x, session))) throw new InvalidOperationException("文档列表已更新，请重新选择。");
        }
    }
}
