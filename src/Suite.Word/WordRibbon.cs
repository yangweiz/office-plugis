using System.IO;
using System.Reflection;
using System.Text;
using Office = Microsoft.Office.Core;

namespace Suite.Word
{
    internal sealed class WordRibbon : Office.IRibbonExtensibility
    {
        private readonly ThisAddIn _addIn;
        private string _searchText = string.Empty;

        public WordRibbon(ThisAddIn addIn) { _addIn = addIn; }

        public string GetCustomUI(string ribbonId)
        {
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Suite.Word.WordRibbon.xml"))
            {
                if (stream == null) throw new FileNotFoundException("找不到内嵌的 Word Ribbon XML。");
                using (var reader = new StreamReader(stream, Encoding.UTF8)) return reader.ReadToEnd();
            }
        }

        public void OnCommand(Office.IRibbonControl control) { _addIn.ExecuteCommand(control.Tag); }
        public void OnSearchChanged(Office.IRibbonControl control, string text) { _searchText = text; }
        public void OnSearch(Office.IRibbonControl control) { _addIn.ShowCommandSearch(_searchText); }
        public void OnDocuments(Office.IRibbonControl control) { _addIn.ToggleDocumentPane(); }
        public void OnNativeTabs(Office.IRibbonControl control) { _addIn.ShowNativeTabStatus(); }
        public void OnSettings(Office.IRibbonControl control) { _addIn.ShowSettings(); }
        public void OnAbout(Office.IRibbonControl control) { _addIn.ShowAbout(); }
    }
}
