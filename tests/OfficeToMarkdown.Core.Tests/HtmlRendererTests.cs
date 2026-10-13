using System.Collections.Generic;
using System.Text;
using OfficeToMarkdown.Core.Model;
using Xunit;

namespace OfficeToMarkdown.Core.Tests
{
    public class HtmlRendererTests
    {
        private static List<Inline> Text(string text) => Inline.FromPlainText(text);

        private static string Render(params Block[] blocks) => HtmlRenderer.Render(blocks, p => "src:" + p);

        [Fact]
        public void Nested_lists_are_nested_inside_parent_item()
        {
            var html = Render(
                new ListItemBlock(0, true, Text("가")),
                new ListItemBlock(1, false, Text("가-1")),
                new ListItemBlock(0, true, Text("나")),
                new ParagraphBlock(Text("끝")));

            Assert.Equal("<ol><li>가<ul><li>가-1</li></ul></li><li>나</li></ol><p>끝</p>", html);
        }

        [Fact]
        public void List_levels_cannot_jump_more_than_one_step()
        {
            var html = Render(
                new ListItemBlock(2, false, Text("너무 깊음")),
                new ListItemBlock(3, false, Text("하위")));

            Assert.Equal("<ul><li>너무 깊음<ul><li>하위</li></ul></li></ul>", html);
        }

        [Fact]
        public void List_type_change_at_same_level_starts_new_list()
        {
            var html = Render(
                new ListItemBlock(0, false, Text("a")),
                new ListItemBlock(0, true, Text("b")));

            Assert.Equal("<ul><li>a</li></ul><ol><li>b</li></ol>", html);
        }

        [Fact]
        public void Text_is_escaped_and_formatted()
        {
            var html = Render(new ParagraphBlock(new List<Inline>
            {
                Inline.Plain("a < b & "),
                new Inline { Text = "굵게", Bold = true, Link = "http://x/?a=1&b=2" },
                Inline.Break(),
                Inline.Plain("둘째 줄"),
            }));

            Assert.Equal("<p>a &lt; b &amp; <a href=\"http://x/?a=1&amp;b=2\"><b>굵게</b></a><br>둘째 줄</p>", html);
        }

        [Fact]
        public void Images_use_source_function_and_missing_ones_are_dropped()
        {
            var html = HtmlRenderer.Render(new Block[]
            {
                new ParagraphBlock(new List<Inline> { Inline.Image("a_files/1.png", "그림") }),
                new ParagraphBlock(new List<Inline> { Inline.Image("a_files/없음.png", "x") }),
            }, p => p == "a_files/1.png" ? HtmlRenderer.DataUri(p, new byte[] { 1, 2, 3 }) : null);

            Assert.Equal("<p><img src=\"data:image/png;base64,AQID\" alt=\"그림\"></p>", html);
        }

        [Fact]
        public void Table_pads_short_rows_and_uses_header_cells()
        {
            var html = Render(new TableBlock(new List<List<List<Inline>>>
            {
                new List<List<Inline>> { Text("A"), Text("B") },
                new List<List<Inline>> { Text("1") },
            }));

            Assert.Equal("<table border=\"1\" style=\"border-collapse:collapse\">"
                         + "<tr><th>A</th><th>B</th></tr><tr><td>1</td><td></td></tr></table>", html);
        }

        [Fact]
        public void Clipboard_offsets_are_utf8_byte_positions()
        {
            var fragment = "<p>한글</p>";
            var cf = ClipboardHtml.Format(fragment);
            var bytes = Encoding.UTF8.GetBytes(cf);

            int Offset(string key)
            {
                var i = cf.IndexOf(key + ":") + key.Length + 1;
                return int.Parse(cf.Substring(i, 10));
            }

            Assert.Equal(bytes.Length, Offset("EndHTML"));
            Assert.StartsWith("<html>", Encoding.UTF8.GetString(bytes, Offset("StartHTML"), 6));
            var start = Offset("StartFragment");
            Assert.Equal(fragment, Encoding.UTF8.GetString(bytes, start, Offset("EndFragment") - start));
        }

        [Fact]
        public void Clipboard_fragment_round_trips()
        {
            var fragment = "<h1>제목</h1><p>본문 😀</p>";

            Assert.Equal(fragment, ClipboardHtml.ExtractFragment(ClipboardHtml.Format(fragment)));
        }

        [Fact]
        public void Extract_fragment_returns_null_for_invalid_input()
        {
            Assert.Null(ClipboardHtml.ExtractFragment(null));
            Assert.Null(ClipboardHtml.ExtractFragment("<p>머리글 없음</p>"));
            Assert.Null(ClipboardHtml.ExtractFragment("StartFragment:0000000010\r\nEndFragment:0000009999\r\n"));
        }
    }
}
