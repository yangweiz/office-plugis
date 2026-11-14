using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Suite.Core.Commands;
using Suite.Windowing;
using Excel = Microsoft.Office.Interop.Excel;

namespace Suite.Excel
{
    internal sealed class ExcelDocumentSessionAdapter : IDocumentSessionAdapter
    {
        private readonly Excel.Application _application;
        public ExcelDocumentSessionAdapter(Excel.Application application) { _application = application ?? throw new ArgumentNullException("application"); }

        public IReadOnlyList<DocumentSession> GetSessions()
        {
            var sessions = new List<DocumentSession>();
            Excel.Workbooks workbooks = _application.Workbooks;
            try
            {
                for (int index = 1; index <= workbooks.Count; index++)
                {
                    Excel.Workbook workbook = workbooks[index];
                    try
                    {
                        string path = workbook.Path.Length == 0 ? string.Empty : workbook.FullName;
                        string key = path.Length > 0 ? path : "excel:unsaved:" + GetWindowId(workbook) + ":" + workbook.Name;
                        sessions.Add(new DocumentSession(key, workbook.Name, path, SuiteHost.Excel, !workbook.Saved, workbook.Name));
                    }
                    finally { Release(workbook); }
                }
            }
            finally { Release(workbooks); }
            return sessions;
        }

        public void Activate(DocumentSession session)
        {
            Excel.Workbook workbook = Find(session);
            try { workbook.Activate(); }
            finally { Release(workbook); }
        }

        public void Save(DocumentSession session)
        {
            Excel.Workbook workbook = Find(session);
            try { workbook.Save(); }
            finally { Release(workbook); }
        }

        public void Close(DocumentSession session, bool saveChanges)
        {
            Excel.Workbook workbook = Find(session);
            try { workbook.Close(saveChanges); }
            finally { Release(workbook); }
        }

        private Excel.Workbook Find(DocumentSession session)
        {
            Excel.Workbooks workbooks = _application.Workbooks;
            try
            {
                for (int index = 1; index <= workbooks.Count; index++)
                {
                    Excel.Workbook workbook = workbooks[index];
                    bool matches = session.FullPath.Length > 0
                        ? string.Equals(workbook.FullName, session.FullPath, StringComparison.OrdinalIgnoreCase)
                        : string.Equals(workbook.Name, session.NativeDocument as string, StringComparison.Ordinal);
                    if (matches) return workbook;
                    Release(workbook);
                }
            }
            finally { Release(workbooks); }
            throw new InvalidOperationException("工作簿已关闭或已更名，请刷新文档列表。");
        }

        private static int GetWindowId(Excel.Workbook workbook)
        {
            Excel.Windows windows = workbook.Windows;
            try
            {
                if (windows.Count == 0) return 0;
                Excel.Window window = windows[1];
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
