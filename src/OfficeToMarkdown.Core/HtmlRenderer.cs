using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using OfficeToMarkdown.Core.Model;

namespace OfficeToMarkdown.Core
{
    /// <summary>
    /// 블록 모델을 HTML 조각으로 렌더링한다. 클립보드 HTML 형식으로 Word·Notion·메일 등에
    /// 서식째 붙여넣기 위한 것이므로 문서 골격(html/head/body) 없이 본문만 만든다.
    /// </summary>
    public static class HtmlRenderer
    {
        /// <param name="imageSource">Markdown 기준 상대 경로를 img src 로 바꾸는 함수. null 을 돌려주면 그 이미지는 빠진다.</param>
        public static string Render(IEnumerable<Block> blocks, Func<string, string> imageSource)
        {
            var sb = new StringBuilder();
            // 열려 있는 목록들의 번호 여부. 깊이 = Count.
            var openLists = new List<bool>();

            foreach (var block in blocks)
            {
                var item = block as ListItemBlock;
                if (item != null)
                {
                    var text = RenderInlines(item.Inlines, imageSource);
                    if (text.Length == 0)
                        continue;
                    // Markdown 렌더러와 같이 한 번에 한 단계씩만 깊어진다.
                    int level = Math.Max(0, Math.Min(item.Level, openLists.Count));
                    while (openLists.Count > level + 1)
                        CloseList(sb, openLists);
                    if (openLists.Count == level + 1)
                    {
                        if (openLists[level] == item.Ordered)
                            sb.Append("</li>");
                        else
                            CloseList(sb, openLists);
                    }
                    if (openLists.Count == level)
                    {
                        sb.Append(item.Ordered ? "<ol>" : "<ul>");
                        openLists.Add(item.Ordered);
                    }
                    sb.Append("<li>").Append(text);
                    continue;
                }

                var html = RenderBlock(block, imageSource);
                if (html == null)
                    continue;
                while (openLists.Count > 0)
                    CloseList(sb, openLists);
                sb.Append(html);
            }

            while (openLists.Count > 0)
                CloseList(sb, openLists);
            return sb.ToString();
        }

        /// <summary>리소스 바이트를 img src 에 넣을 data URI 로 만든다.</summary>
        public static string DataUri(string path, byte[] data)
        {
            return "data:" + MimeTypeFor(path) + ";base64," + Convert.ToBase64String(data);
        }

        private static string MimeTypeFor(string path)
        {
            var dot = path.LastIndexOf('.');
            var ext = dot < 0 ? "" : path.Substring(dot + 1).ToLowerInvariant();
            switch (ext)
            {
                case "jpg":
                case "jpeg":
                    return "image/jpeg";
                case "gif":
                    return "image/gif";
                case "bmp":
                    return "image/bmp";
                case "svg":
                    return "image/svg+xml";
                case "emf":
                    return "image/x-emf";
                case "wmf":
                    return "image/x-wmf";
                case "tif":
                case "tiff":
                    return "image/tiff";
                default:
                    return "image/png";
            }
        }

        private static void CloseList(StringBuilder sb, List<bool> openLists)
        {
            sb.Append(openLists[openLists.Count - 1] ? "</li></ol>" : "</li></ul>");
            openLists.RemoveAt(openLists.Count - 1);
        }

        private static string RenderBlock(Block block, Func<string, string> imageSource)
        {
            var heading = block as HeadingBlock;
            if (heading != null)
            {
                var text = RenderInlines(heading.Inlines, imageSource);
                if (text.Length == 0)
                    return null;
                var tag = "h" + Math.Max(1, Math.Min(6, heading.Level));
                return "<" + tag + ">" + text + "</" + tag + ">";
            }

            var paragraph = block as ParagraphBlock;
            if (paragraph != null)
            {
                var text = RenderInlines(paragraph.Inlines, imageSource);
                return text.Length == 0 ? null : "<p>" + text + "</p>";
            }

            var quote = block as QuoteBlock;
            if (quote != null)
            {
                var text = RenderInlines(quote.Inlines, imageSource);
                return text.Length == 0 ? null : "<blockquote><p>" + text + "</p></blockquote>";
            }

            var table = block as TableBlock;
            if (table != null)
                return RenderTable(table, imageSource);

            if (block is RuleBlock)
                return "<hr>";

            throw new NotSupportedException("알 수 없는 블록 형식: " + block.GetType().Name);
        }

        private static string RenderTable(TableBlock table, Func<string, string> imageSource)
        {
            int columns = table.Rows.Count == 0 ? 0 : table.Rows.Max(r => r.Count);
            if (columns == 0)
                return null;

            var sb = new StringBuilder("<table border=\"1\" style=\"border-collapse:collapse\">");
            for (int i = 0; i < table.Rows.Count; i++)
            {
                var tag = i == 0 ? "th" : "td";
                sb.Append("<tr>");
                for (int c = 0; c < columns; c++)
                {
                    var inlines = c < table.Rows[i].Count ? table.Rows[i][c] : null;
                    var cell = inlines != null ? RenderInlines(inlines, imageSource) : "";
                    sb.Append('<').Append(tag);
                    if (i > 0 && inlines != null)
                        AppendCellStyle(sb, inlines);
                    sb.Append('>').Append(cell).Append("</").Append(tag).Append('>');
                }
                sb.Append("</tr>");
            }
            return sb.Append("</table>").ToString();
        }

        /// <summary>짧은 셀은 이 글자 수 이하면 줄바꿈하지 않는다.</summary>
        private const int ShortCellLength = 20;

        /// <summary>
        /// Excel 표가 읽기 좋도록 셀 모양을 정한다. 붙여넣는 프로그램(Word 등)은 CSS 를 받지 않으므로 인라인 스타일로 넣는다.
        /// - 숫자: 오른쪽 정렬, 줄바꿈 없음
        /// - 짧은 값(날짜, 이름 등): 줄바꿈 없음 — "2026-" / "01-03" 처럼 잘리지 않게
        /// - 긴 문장: 최소 폭을 줘서 좁고 긴 셀이 되지 않게
        /// </summary>
        private static void AppendCellStyle(StringBuilder sb, List<Inline> inlines)
        {
            if (inlines.Count == 0 || inlines.Any(x => x.Kind != InlineKind.Text))
                return;
            var text = string.Concat(inlines.Select(x => x.Text)).Trim();
            if (text.Length == 0)
                return;
            if (IsNumber(text))
                sb.Append(" style=\"text-align:right;white-space:nowrap\"");
            else if (text.Length <= ShortCellLength)
                sb.Append(" style=\"white-space:nowrap\"");
            else
                sb.Append(" style=\"min-width:14em\"");
        }

        // Excel 표시 형식으로 나오는 숫자: -1,234.5 / ₩1,000 / $ 12.00 / 35% / (1,000) / 1.23E+11 / 1,000원
        private static readonly Regex NumberPattern = new Regex(
            @"^\(?[-+]?[$₩€¥£]?\s?[-+]?(\d{1,3}(,\d{3})+|\d+)?(\.\d+)?([eE][-+]?\d+)?\s?[%원]?\)?$",
            RegexOptions.Compiled);

        private static bool IsNumber(string text) => text.Any(char.IsDigit) && NumberPattern.IsMatch(text);

        private static string RenderInlines(List<Inline> inlines, Func<string, string> imageSource)
        {
            var sb = new StringBuilder();
            foreach (var inline in inlines)
            {
                switch (inline.Kind)
                {
                    case InlineKind.LineBreak:
                        sb.Append("<br>");
                        break;
                    case InlineKind.Image:
                        var src = imageSource(inline.ImagePath);
                        if (src != null)
                            sb.Append("<img src=\"").Append(WebUtility.HtmlEncode(src))
                              .Append("\" alt=\"").Append(WebUtility.HtmlEncode(inline.Text)).Append("\">");
                        break;
                    default:
                        sb.Append(FormatRun(inline));
                        break;
                }
            }

            // 앞뒤의 줄바꿈·공백만 있는 내용은 빈 블록으로 본다.
            var html = sb.ToString();
            while (html.StartsWith("<br>", StringComparison.Ordinal))
                html = html.Substring(4);
            while (html.EndsWith("<br>", StringComparison.Ordinal))
                html = html.Substring(0, html.Length - 4);
            return html.Trim();
        }

        private static string FormatRun(Inline inline)
        {
            var html = WebUtility.HtmlEncode(inline.Text.Replace('\t', ' '));
            if (html.Length == 0)
                return html;
            if (inline.Bold)
                html = "<b>" + html + "</b>";
            if (inline.Italic)
                html = "<i>" + html + "</i>";
            if (inline.Strike)
                html = "<s>" + html + "</s>";
            if (inline.Link != null)
                html = "<a href=\"" + WebUtility.HtmlEncode(inline.Link) + "\">" + html + "</a>";
            return html;
        }
    }
}
