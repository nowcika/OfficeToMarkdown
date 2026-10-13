using System;
using System.Runtime.InteropServices;
using OfficeToMarkdown.AddIns;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointToMarkdown
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private readonly Func<PowerPoint.Application> _application;

        public Ribbon(Func<PowerPoint.Application> application)
        {
            _application = application;
        }

        public string GetCustomUI(string ribbonID) => ExportCommand.RibbonXml;

        public void OnExport(Office.IRibbonControl control)
        {
            var app = _application();
            var owner = new IntPtr(app.HWND);
            if (app.Presentations.Count == 0)
            {
                ExportCommand.ShowNoDocument(owner);
                return;
            }

            var presentation = app.ActivePresentation;
            ExportCommand.Run(owner, presentation.Name,
                              assetFolder => PresentationExtractor.Extract(presentation, assetFolder));
        }
    }
}
