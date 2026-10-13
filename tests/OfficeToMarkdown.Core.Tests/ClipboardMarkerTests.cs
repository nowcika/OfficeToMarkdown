using System.Text;
using Xunit;

namespace OfficeToMarkdown.Core.Tests
{
    public class ClipboardMarkerTests
    {
        [Fact]
        public void Round_trips_id_and_source_name()
        {
            var marker = ClipboardMarker.Create("한글 보고서.docx");
            var parsed = ClipboardMarker.Parse(marker.ToBytes());

            Assert.Equal(marker.Id, parsed.Id);
            Assert.Equal("한글 보고서.docx", parsed.SourceName);
        }

        [Fact]
        public void Each_marker_has_new_id()
        {
            Assert.NotEqual(ClipboardMarker.Create("a").Id, ClipboardMarker.Create("a").Id);
        }

        [Fact]
        public void Ignores_trailing_nulls_and_crlf()
        {
            var parsed = ClipboardMarker.Parse(Encoding.UTF8.GetBytes("id=abc\r\nsource=a.xlsx\r\n\0\0"));

            Assert.Equal("abc", parsed.Id);
            Assert.Equal("a.xlsx", parsed.SourceName);
        }

        [Fact]
        public void Returns_null_without_id()
        {
            Assert.Null(ClipboardMarker.Parse(Encoding.UTF8.GetBytes("source=a.xlsx\n")));
            Assert.Null(ClipboardMarker.Parse(null));
        }
    }
}
