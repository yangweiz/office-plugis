using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Suite.Core.Commands;
using Suite.Windowing;
using Word = Microsoft.Office.Interop.Word;

namespace Suite.Word
{
    internal sealed class WordDocumentSessionAdapter : IDocumentSessionAdapter
    {
        private readonly Word.Application _application;
        public WordDocumentSessionAdapter(Word.Application application) { _application = application ?? throw new ArgumentNullException("application"); }

        public IReadOnlyList<DocumentSession> GetSessions()
        {
            var sessions = new List<DocumentSession>();
            Word.Documents documents = _application.Documents;
            try
            {
                for (int index = 1; index <= documents.Count; index++)
                {
                    Word.Document document = documents[index];
                    try
                    {
                        string path = document.Path.Length == 0 ? string.Empty : document.FullName;
                        string key = path.Length > 0 ? path : "word:unsaved:" + GetWindowId(document) + ":" + document.Name;
                        sessions.Add(new DocumentSession(key, document.Name, path, SuiteHost.Word, !document.Saved, document.Name));
                    }
                    finally { Release(document); }
                }
            }
            finally { Release(documents); }
            return sessions;
        }

        public void Activate(DocumentSession session)
        {
            Word.Document document = Find(session);
            try { document.Activate(); }
            finally { Release(document); }
        }

        public void Save(DocumentSession session)
        {
            Word.Document document = Find(session);
            try { document.Save(); }
            finally { Release(document); }
        }

        public void Close(DocumentSession session, bool saveChanges)
        {
            Word.Document document = Find(session);
            try { document.Close(saveChanges ? Word.WdSaveOptions.wdSaveChanges : Word.WdSaveOptions.wdDoNotSaveChanges); }
            finally { Release(document); }
        }

        private Word.Document Find(DocumentSession session)
        {
            Word.Documents documents = _application.Documents;
            try
            {
                for (int index = 1; index <= documents.Count; index++)
                {
                    Word.Document document = documents[index];
                    bool matches = session.FullPath.Length > 0
                        ? string.Equals(document.FullName, session.FullPath, StringComparison.OrdinalIgnoreCase)
                        : string.Equals(document.Name, session.NativeDocument as string, StringComparison.Ordinal);
                    if (matches) return document;
                    Release(document);
                }
            }
            finally { Release(documents); }
            throw new InvalidOperationException("文档已关闭或已更名，请刷新文档列表。");
        }

        private static int GetWindowId(Word.Document document)
        {
            Word.Windows windows = document.Windows;
            try
            {
                Word.Window window = windows[1];
                try { return window.Hwnd; }
                finally { Release(window); }
            }
            finally { Release(windows); }
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
        }
    }
}
