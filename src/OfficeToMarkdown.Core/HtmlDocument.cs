using System.Net;

namespace OfficeToMarkdown.Core
{
    /// <summary><see cref="HtmlRenderer"/> 가 만든 HTML 조각을 브라우저에서 보기 좋은 완전한 문서로 감싼다.</summary>
    public static class HtmlDocument
    {
        private const string Style = @"
body { font-family: 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif; max-width: 1100px; margin: 2em auto; padding: 0 1em; line-height: 1.6; color: #222; }
img { max-width: 100%; height: auto; }
blockquote { margin: 1em 0; padding-left: 1em; border-left: 4px solid #ddd; color: #555; }
/* 열이 많은 표는 셀을 누르지 않고 가로 스크롤한다. */
.table-wrap { overflow-x: auto; margin: 1em 0; }
table { border-collapse: collapse; font-size: 14px; line-height: 1.4; }
th, td { border: 1px solid #d0d0d0; padding: 4px 10px; vertical-align: top; text-align: left;
         word-break: keep-all; overflow-wrap: break-word; min-width: 3em; max-width: 28em; }
th { background: #eef1f5; font-weight: 600; white-space: nowrap; }
tr:nth-child(even) td { background: #fafbfc; }
";

        public static string Build(string title, string fragment)
        {
            // 텍스트는 HtmlRenderer 가 인코딩하므로 "<table" 은 실제 표 태그에만 나온다.
            var body = fragment.Replace("<table", "<div class=\"table-wrap\"><table")
                               .Replace("</table>", "</table></div>");
            return "<!DOCTYPE html>\n<html lang=\"ko\">\n<head>\n<meta charset=\"utf-8\">\n"
                   + "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n"
                   + "<title>" + WebUtility.HtmlEncode(title ?? "") + "</title>\n"
                   + "<style>" + Style + "</style>\n</head>\n<body>\n"
                   + body
                   + "\n</body>\n</html>\n";
        }
    }
}
