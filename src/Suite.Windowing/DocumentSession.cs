using System.Drawing;
using Suite.Core.Commands;

namespace Suite.Windowing
{
    public sealed class DocumentSession
    {
        public DocumentSession(string key, string displayName, string fullPath, SuiteHost host, bool isDirty, object nativeDocument)
        {
            Key = key ?? string.Empty;
            DisplayName = displayName ?? string.Empty;
            FullPath = fullPath ?? string.Empty;
            Host = host;
            IsDirty = isDirty;
            NativeDocument = nativeDocument;
            Color = Color.Empty;
        }

        public string Key { get; private set; }
        public string DisplayName { get; private set; }
        public string FullPath { get; private set; }
        public SuiteHost Host { get; private set; }
        public bool IsDirty { get; internal set; }
        public object NativeDocument { get; private set; }
        public Color Color { get; internal set; }
        public override string ToString() { return DisplayName + (IsDirty ? " *" : string.Empty); }
    }

    public interface IDocumentSessionAdapter
    {
        System.Collections.Generic.IReadOnlyList<DocumentSession> GetSessions();
        void Activate(DocumentSession session);
        void Save(DocumentSession session);
        void Close(DocumentSession session, bool saveChanges);
    }

    public enum CloseDecision { Save, Discard, Cancel }
}
