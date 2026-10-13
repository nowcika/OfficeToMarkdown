using System;
using System.Collections.Generic;
using System.Globalization;
using OfficeToMarkdown.Core.Model;

namespace OfficeToMarkdown.Core.Excel
{
    /// <summary>Range.Value 로 읽은 2차원 배열을 Markdown 표로 변환한다.</summary>
    public static class ExcelSheetConverter
    {
        // Range.Value 는 오류 셀을 CVErr 정수(Int32)로 돌려준다.
        private static readonly Dictionary<int, string> ErrorValues = new Dictionary<int, string>
        {
            { -2146826281, "#DIV/0!" },
            { -2146826246, "#N/A" },
            { -2146826259, "#NAME?" },
            { -2146826288, "#NULL!" },
            { -2146826252, "#NUM!" },
            { -2146826265, "#REF!" },
            { -2146826273, "#VALUE!" },
        };

        /// <summary>
        /// 시트 이름을 제목으로, 값을 표로 추가한다. 내용이 없는 시트는 건너뛴다.
        /// </summary>
        /// <param name="values">Range.Value 결과. COM 에서 오는 배열은 1부터 시작하지만 하한에 관계없이 처리한다.</param>
        public static void AppendSheet(DocumentContent content, string sheetName, object[,] values)
        {
            var table = ToTable(values);
            if (table == null)
                return;
            content.Blocks.Add(new HeadingBlock(2, sheetName));
            content.Blocks.Add(table);
        }

        /// <summary>뒤쪽의 빈 행·열을 잘라낸 표를 만든다. 모든 셀이 비어 있으면 null.</summary>
        public static TableBlock ToTable(object[,] values)
        {
            if (values == null)
                return null;

            int r0 = values.GetLowerBound(0), c0 = values.GetLowerBound(1);
            int rowCount = values.GetLength(0), colCount = values.GetLength(1);

            var text = new string[rowCount, colCount];
            int lastRow = -1, lastCol = -1;
            for (int r = 0; r < rowCount; r++)
            {
                for (int c = 0; c < colCount; c++)
                {
                    text[r, c] = FormatValue(values[r0 + r, c0 + c]);
                    if (text[r, c].Length > 0)
                    {
                        lastRow = Math.Max(lastRow, r);
                        lastCol = Math.Max(lastCol, c);
                    }
                }
            }
            if (lastRow < 0)
                return null;

            var rows = new List<List<List<Inline>>>();
            for (int r = 0; r <= lastRow; r++)
            {
                var row = new List<List<Inline>>();
                for (int c = 0; c <= lastCol; c++)
                    row.Add(Inline.FromPlainText(text[r, c]));
                rows.Add(row);
            }
            return new TableBlock(rows);
        }

        public static string FormatValue(object value)
        {
            if (value == null)
                return "";
            if (value is string)
                return (string)value;
            if (value is bool)
                return (bool)value ? "TRUE" : "FALSE";
            if (value is DateTime)
            {
                var date = (DateTime)value;
                return date.TimeOfDay == TimeSpan.Zero
                    ? date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                    : date.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            }
            if (value is double)
                return ((double)value).ToString("G15", CultureInfo.InvariantCulture);
            if (value is int)
            {
                string error;
                return ErrorValues.TryGetValue((int)value, out error) ? error : ((int)value).ToString(CultureInfo.InvariantCulture);
            }
            var formattable = value as IFormattable;
            return formattable != null ? formattable.ToString(null, CultureInfo.InvariantCulture) : value.ToString();
        }
    }
}
