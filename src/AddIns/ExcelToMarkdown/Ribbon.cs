using System;
using System.Runtime.InteropServices;
using OfficeToMarkdown.AddIns;
using OfficeToMarkdown.Core.Excel;
using OfficeToMarkdown.Core.Model;
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

            ExportCommand.Run(owner, workbook.Path, workbook.Name, _ => Extract(workbook));
        }

        /// <summary>보이는 워크시트마다 시트 이름 제목 + 사용 범위 표를 만든다. (차트 시트·숨긴 시트 제외)</summary>
        private static DocumentContent Extract(Excel.Workbook workbook)
        {
            var content = new DocumentContent();
            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                    continue;

                // Value2 와 달리 Value 는 날짜 셀을 DateTime 으로 돌려준다. 한 번에 배열로 읽어 COM 호출을 줄인다.
                object value = sheet.UsedRange.Value[Excel.XlRangeValueDataType.xlRangeValueDefault];
                var values = value as object[,];
                if (values == null && value != null)
                    values = new object[,] { { value } }; // 사용 범위가 셀 하나면 배열이 아닌 값이 온다.

                ExcelSheetConverter.AppendSheet(content, sheet.Name, values);
            }
            return content;
        }
    }
}
