using System;
using System.Runtime.InteropServices;
using OfficeToMarkdown.AddIns;
using OfficeToMarkdown.Core.Word;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace WordToMarkdown
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private readonly Func<Word.Application> _application;

        public Ribbon(Func<Word.Application> application)
        {
            _application = application;
        }

        public string GetCustomUI(string ribbonID) => ExportCommand.RibbonXml;

        public void OnExport(Office.IRibbonControl control)
        {
            var app = _application();
            if (app.Documents.Count == 0)
            {
                ExportCommand.ShowNoDocument(IntPtr.Zero);
                return;
            }

            var document = app.ActiveDocument;
            var owner = new IntPtr(app.ActiveWindow.Hwnd);
            // WordOpenXML 은 문서 전체를 Flat OPC XML 로 돌려주므로 COM 왕복 한 번으로 내용을 모두 읽는다.
            ExportCommand.Run(owner, document.Path, document.Name,
                              assetFolder => WordOpenXmlConverter.Convert(document.WordOpenXML, assetFolder));
        }
    }
}
