using System;
using System.IO;

namespace OfficeToMarkdown.Core
{
    /// <summary>
    /// 파일 경로로만 내보낼 수 있는 Office API(Shape.Export, Chart.Export 등)의 결과를
    /// 임시 파일을 거쳐 바이트 배열로 받는다.
    /// </summary>
    public static class TempFileExport
    {
        /// <param name="export">임시 파일 경로를 받아 그 경로에 파일을 쓰는 동작.</param>
        /// <param name="extension">임시 파일 확장자 (예: ".png"). Office 는 확장자로 형식을 판단하기도 한다.</param>
        /// <returns>파일 내용. 파일이 만들어지지 않았거나 비어 있으면 null.</returns>
        public static byte[] Capture(Action<string> export, string extension)
        {
            var path = Path.Combine(Path.GetTempPath(), "OfficeToMarkdown_" + Guid.NewGuid().ToString("N") + extension);
            try
            {
                export(path);
                if (!File.Exists(path))
                    return null;
                var data = File.ReadAllBytes(path);
                return data.Length == 0 ? null : data;
            }
            finally
            {
                try
                {
                    File.Delete(path);
                }
                catch (IOException)
                {
                }
                catch (UnauthorizedAccessException)
                {
                }
            }
        }
    }
}
