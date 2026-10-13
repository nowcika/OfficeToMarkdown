using System.Text;

namespace OfficeToMarkdown.Core
{
    /// <summary>
    /// Windows 클립보드 "HTML Format"(CF_HTML) 문자열을 만든다.
    /// 머리글의 오프셋은 UTF-8 바이트 기준이며, .NET Framework 4.5 이상의 DataObject 는 이 문자열을 UTF-8 로 기록한다.
    /// https://learn.microsoft.com/windows/win32/dataxchg/html-clipboard-format
    /// </summary>
    public static class ClipboardHtml
    {
        private const string HeaderFormat =
            "Version:0.9\r\nStartHTML:{0:D10}\r\nEndHTML:{1:D10}\r\nStartFragment:{2:D10}\r\nEndFragment:{3:D10}\r\n";

        private const string Prefix = "<html><head><meta charset=\"utf-8\"></head><body><!--StartFragment-->";
        private const string Suffix = "<!--EndFragment--></body></html>";

        public static string Format(string fragment)
        {
            var utf8 = Encoding.UTF8;
            // 오프셋 자리는 10자리 고정이므로 0 으로 채운 머리글 길이가 실제 길이와 같다.
            int headerLength = utf8.GetByteCount(string.Format(HeaderFormat, 0, 0, 0, 0));
            int startFragment = headerLength + utf8.GetByteCount(Prefix);
            int endFragment = startFragment + utf8.GetByteCount(fragment);
            int endHtml = endFragment + utf8.GetByteCount(Suffix);

            return string.Format(HeaderFormat, headerLength, endHtml, startFragment, endFragment)
                   + Prefix + fragment + Suffix;
        }

        /// <summary>CF_HTML 문자열에서 StartFragment~EndFragment 사이의 HTML 조각을 꺼낸다.</summary>
        /// <returns>머리글이 없거나 오프셋이 맞지 않으면 null.</returns>
        public static string ExtractFragment(string cfHtml)
        {
            if (string.IsNullOrEmpty(cfHtml))
                return null;
            int start = ReadOffset(cfHtml, "StartFragment:");
            int end = ReadOffset(cfHtml, "EndFragment:");
            var bytes = Encoding.UTF8.GetBytes(cfHtml);
            if (start < 0 || end < start || end > bytes.Length)
                return null;
            return Encoding.UTF8.GetString(bytes, start, end - start);
        }

        private static int ReadOffset(string cfHtml, string key)
        {
            int i = cfHtml.IndexOf(key, System.StringComparison.Ordinal);
            if (i < 0)
                return -1;
            i += key.Length;
            int j = i;
            while (j < cfHtml.Length && char.IsDigit(cfHtml[j]))
                j++;
            int value;
            return int.TryParse(cfHtml.Substring(i, j - i), out value) ? value : -1;
        }
    }
}
