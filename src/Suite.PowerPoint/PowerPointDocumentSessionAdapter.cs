using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Suite.Core.Commands;
using Suite.Windowing;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Suite.PowerPoint
{
    internal sealed class PowerPointDocumentSessionAdapter : IDocumentSessionAdapter
    {
        private readonly PowerPoint.Application _application;
        public PowerPointDocumentSessionAdapter(PowerPoint.Application application) { _application = application ?? throw new ArgumentNullException("application"); }

        public IReadOnlyList<DocumentSession> GetSessions()
        {
            var sessions = new List<DocumentSession>();
            PowerPoint.Presentations presentations = _application.Presentations;
            try
            {
                for (int index = 1; index <= presentations.Count; index++)
                {
                    PowerPoint.Presentation presentation = presentations[index];
                    try
                    {
                        string path = presentation.Path.Length == 0 ? string.Empty : presentation.FullName;
                        string key = path.Length > 0 ? path : "powerpoint:unsaved:" + GetWindowId(presentation) + ":" + presentation.Name;
                        sessions.Add(new DocumentSession(key, presentation.Name, path, SuiteHost.PowerPoint, presentation.Saved != MsoTriState.msoTrue, presentation.Name));
                    }
                    finally { Release(presentation); }
                }
            }
            finally { Release(presentations); }
            return sessions;
        }

        public void Activate(DocumentSession session)
        {
            PowerPoint.Presentation presentation = Find(session);
            PowerPoint.DocumentWindows windows = null;
            PowerPoint.DocumentWindow window = null;
            try
            {
                windows = presentation.Windows;
                window = windows[1];
                window.Activate();
            }
            finally { Release(window); Release(windows); Release(presentation); }
        }

        public void Save(DocumentSession session)
        {
            PowerPoint.Presentation presentation = Find(session);
            try { presentation.Save(); }
            finally { Release(presentation); }
        }

        public void Close(DocumentSession session, bool saveChanges)
        {
            PowerPoint.Presentation presentation = Find(session);
            try
            {
                if (saveChanges) presentation.Save();
                presentation.Close();
            }
            finally { Release(presentation); }
        }

        private PowerPoint.Presentation Find(DocumentSession session)
        {
            PowerPoint.Presentations presentations = _application.Presentations;
            try
            {
                for (int index = 1; index <= presentations.Count; index++)
                {
                    PowerPoint.Presentation presentation = presentations[index];
                    bool matches = session.FullPath.Length > 0
                        ? string.Equals(presentation.FullName, session.FullPath, StringComparison.OrdinalIgnoreCase)
                        : string.Equals(presentation.Name, session.NativeDocument as string, StringComparison.Ordinal);
                    if (matches) return presentation;
                    Release(presentation);
                }
            }
            finally { Release(presentations); }
            throw new InvalidOperationException("演示文稿已关闭或已更名，请刷新文档列表。");
        }

        private static int GetWindowId(PowerPoint.Presentation presentation)
        {
            PowerPoint.DocumentWindows windows = presentation.Windows;
            try
            {
                PowerPoint.DocumentWindow window = windows[1];
                try { return window.HWND; }
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
