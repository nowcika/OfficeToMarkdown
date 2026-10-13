using System;
using System.Runtime.InteropServices;
using OfficeToMarkdown.AddIns;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelToMarkdown
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private readonly Func<Excel.Application> _application;

        public Ribbon(Func<Excel.Application> application)
        {
            _application = application;
        }

        public string GetCustomUI(string ribbonID) => ExportCommand.RibbonXml;

        public void OnExport(Office.IRibbonControl control)
        {
            var app = _application();
            var owner = new IntPtr(app.Hwnd);
            var workbook = app.ActiveWorkbook;
            if (workbook == null)
            {
                ExportCommand.ShowNoDocument(owner);
                return;
            }

            ExportCommand.Run(owner, workbook.Path, workbook.Name,
                              assetFolder => WorkbookExtractor.Extract(workbook, assetFolder));
        }
    }
}
