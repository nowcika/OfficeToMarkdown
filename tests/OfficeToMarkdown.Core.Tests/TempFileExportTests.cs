using System.IO;
using Xunit;

namespace OfficeToMarkdown.Core.Tests
{
    public class TempFileExportTests
    {
        [Fact]
        public void Returns_written_bytes_and_deletes_temp_file()
        {
            string written = null;
            var data = TempFileExport.Capture(path =>
            {
                written = path;
                File.WriteAllBytes(path, new byte[] { 1, 2, 3 });
            }, ".png");

            Assert.Equal(new byte[] { 1, 2, 3 }, data);
            Assert.EndsWith(".png", written);
            Assert.False(File.Exists(written));
        }

        [Fact]
        public void Returns_null_when_nothing_was_written()
        {
            Assert.Null(TempFileExport.Capture(_ => { }, ".png"));
        }
    }
}
