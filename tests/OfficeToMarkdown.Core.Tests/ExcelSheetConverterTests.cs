using System;
using OfficeToMarkdown.Core.Excel;
using OfficeToMarkdown.Core.Model;
using Xunit;

namespace OfficeToMarkdown.Core.Tests
{
    public class ExcelSheetConverterTests
    {
        /// <summary>COM 의 Range.Value 처럼 하한이 1인 배열을 만든다.</summary>
        private static object[,] OneBased(object[,] source)
        {
            int rows = source.GetLength(0), cols = source.GetLength(1);
            var result = (object[,])Array.CreateInstance(typeof(object), new[] { rows, cols }, new[] { 1, 1 });
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols; c++)
                    result[r + 1, c + 1] = source[r, c];
            return result;
        }

        [Fact]
        public void Sheet_becomes_heading_and_table_with_trailing_blanks_trimmed()
        {
            var values = OneBased(new object[,]
            {
                { "품목", "수량", "날짜", null },
                { "사과", 3.0, new DateTime(2026, 10, 13), null },
                { "배\n(국산)", 0.1 + 0.2, true, null },
                { null, null, null, null },
            });

            var content = new DocumentContent();
            ExcelSheetConverter.AppendSheet(content, "Sheet1", values);

            Assert.Equal(
                "## Sheet1\n\n" +
                "| 품목 | 수량 | 날짜 |\n" +
                "| --- | --- | --- |\n" +
                "| 사과 | 3 | 2026-10-13 |\n" +
                "| 배<br>(국산) | 0.3 | TRUE |\n",
                MarkdownRenderer.Render(content.Blocks));
        }

        [Fact]
        public void Empty_sheet_is_skipped()
        {
            var content = new DocumentContent();
            ExcelSheetConverter.AppendSheet(content, "빈 시트", OneBased(new object[,] { { null, "" } }));

            Assert.Empty(content.Blocks);
        }

        [Theory]
        [InlineData(-2146826281, "#DIV/0!")]
        [InlineData(-2146826246, "#N/A")]
        public void Error_values_are_shown_as_excel_error_text(int code, string expected)
        {
            Assert.Equal(expected, ExcelSheetConverter.FormatValue(code));
        }

        [Fact]
        public void Date_with_time_keeps_time_part()
        {
            Assert.Equal("2026-10-13 09:30:00", ExcelSheetConverter.FormatValue(new DateTime(2026, 10, 13, 9, 30, 0)));
        }
    }
}
