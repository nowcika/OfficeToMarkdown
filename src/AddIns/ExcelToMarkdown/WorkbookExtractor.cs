using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using OfficeToMarkdown.Core;
using OfficeToMarkdown.Core.Excel;
using OfficeToMarkdown.Core.Model;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToMarkdown
{
    /// <summary>
    /// 보이는 시트마다 시트 이름 제목 + 사용 범위 표 + 차트 이미지를 만든다. 숨긴 시트는 제외한다.
    /// 차트 시트는 차트 이미지만, 워크시트에 삽입된 차트는 표 아래에 위→아래 순서로 넣는다.
    /// </summary>
    internal sealed class WorkbookExtractor
    {
        private readonly string _assetFolder;
        private readonly DocumentContent _content = new DocumentContent();

        private WorkbookExtractor(string assetFolder)
        {
            _assetFolder = assetFolder.TrimEnd('/');
        }

        /// <param name="assetFolder">차트 이미지를 저장할 폴더의 Markdown 파일 기준 상대 경로 (예: "매출_files").</param>
        public static DocumentContent Extract(Excel.Workbook workbook, string assetFolder)
        {
            var extractor = new WorkbookExtractor(assetFolder);
            int sheetIndex = 0;
            foreach (object sheet in workbook.Sheets)
            {
                sheetIndex++;
                var worksheet = sheet as Excel.Worksheet;
                if (worksheet != null)
                {
                    extractor.AppendWorksheet(worksheet, sheetIndex);
                    continue;
                }

                var chartSheet = sheet as Excel.Chart;
                if (chartSheet != null && chartSheet.Visible == Excel.XlSheetVisibility.xlSheetVisible)
                {
                    var image = extractor.ExportChart(chartSheet, chartSheet.Name, sheetIndex, 1);
                    ExcelSheetConverter.AppendSheet(extractor._content, chartSheet.Name, null,
                                                    image == null ? null : new[] { image });
                }
            }
            return extractor._content;
        }

        private void AppendWorksheet(Excel.Worksheet sheet, int sheetIndex)
        {
            if (sheet.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                return;

            // Value2 와 달리 Value 는 날짜 셀을 DateTime 으로 돌려준다. 한 번에 배열로 읽어 COM 호출을 줄인다.
            object value = sheet.UsedRange.Value[Excel.XlRangeValueDataType.xlRangeValueDefault];
            var values = value as object[,];
            if (values == null && value != null)
                values = new object[,] { { value } }; // 사용 범위가 셀 하나면 배열이 아닌 값이 온다.

            var chartObjects = ((Excel.ChartObjects)sheet.ChartObjects()).Cast<Excel.ChartObject>()
                .Where(c => c.Visible)
                .OrderBy(c => c.Top)
                .ThenBy(c => c.Left);

            var images = new List<Inline>();
            int chartNumber = 0;
            foreach (var chartObject in chartObjects)
            {
                var image = ExportChart(chartObject.Chart, chartObject.Name, sheetIndex, ++chartNumber);
                if (image != null)
                    images.Add(image);
            }

            ExcelSheetConverter.AppendSheet(_content, sheet.Name, values, images);
        }

        /// <summary>Chart.Export 로 PNG 를 만들어 리소스에 등록한다. 실패하면 null.</summary>
        private Inline ExportChart(Excel.Chart chart, string fallbackName, int sheetIndex, int chartNumber)
        {
            byte[] data;
            try
            {
                data = TempFileExport.Capture(file => chart.Export(file, "PNG", false), ".png");
            }
            catch (COMException)
            {
                return null;
            }
            if (data == null)
                return null;

            var path = _assetFolder + "/sheet" + sheetIndex + "_chart" + chartNumber + ".png";
            _content.Resources[path] = data;
            return Inline.Image(path, ChartTitle(chart) ?? fallbackName);
        }

        private static string ChartTitle(Excel.Chart chart)
        {
            try
            {
                if (!chart.HasTitle)
                    return null;
                var title = chart.ChartTitle.Text.Replace("\r", " ").Replace("\n", " ").Trim();
                return title.Length > 0 ? title : null;
            }
            catch (COMException)
            {
                return null;
            }
        }
    }
}
