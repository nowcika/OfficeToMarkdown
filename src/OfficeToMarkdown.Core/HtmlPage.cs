using System.Net;

namespace OfficeToMarkdown.Core
{
    /// <summary><see cref="HtmlRenderer"/> 가 만든 HTML 조각을 브라우저에서 보기 좋은 완전한 문서로 감싼다.</summary>
    public static class HtmlPage
    {
        private const string Style = @"
/* 본문 글은 읽기 좋은 폭으로 제한하고, 표는 폭 제한 없이 펼쳐 넓으면 페이지 전체가 가로로 스크롤되게 한다. */
body { font-family: 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif; margin: 2em; line-height: 1.6; color: #222; }
p, h1, h2, h3, h4, h5, h6, ul, ol, blockquote, hr { max-width: 1100px; }
img { max-width: 100%; height: auto; }
blockquote { margin: 1em 0; padding-left: 1em; border-left: 4px solid #ddd; color: #555; }
table { border-collapse: collapse; font-size: 14px; line-height: 1.4; margin: 1em 0; }
th, td { border: 1px solid #d0d0d0; padding: 4px 10px; vertical-align: top; text-align: left;
         word-break: keep-all; overflow-wrap: break-word; min-width: 3em; max-width: 28em; }
th { background: #eef1f5; font-weight: 600; white-space: nowrap; }
tr:nth-child(even) td { background: #fafbfc; }
";

        public static string Build(string title, string fragment)
        {
            return "<!DOCTYPE html>\n<html lang=\"ko\">\n<head>\n<meta charset=\"utf-8\">\n"
                   + "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n"
                   + "<title>" + WebUtility.HtmlEncode(title ?? "") + "</title>\n"
                   + "<style>" + Style + "</style>\n</head>\n<body>\n"
                   + fragment
                   + "\n</body>\n</html>\n";
        }
    }
}
