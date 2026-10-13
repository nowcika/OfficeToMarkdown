using System.Collections.Generic;
using OfficeToMarkdown.Core.Model;
using Xunit;

namespace OfficeToMarkdown.Core.Tests
{
    public class MarkdownRendererTests
    {
        private static List<Inline> Text(string text) => Inline.FromPlainText(text);

        [Fact]
        public void List_levels_cannot_jump_more_than_one_step()
        {
            var md = MarkdownRenderer.Render(new Block[]
            {
                new ListItemBlock(2, false, Text("너무 깊음")),
                new ListItemBlock(3, true, Text("하위")),
            });

            Assert.Equal("- 너무 깊음\n    1. 하위\n", md);
        }

        [Fact]
        public void Ordered_numbering_restarts_for_nested_level()
        {
            var md = MarkdownRenderer.Render(new Block[]
            {
                new ListItemBlock(0, true, Text("가")),
                new ListItemBlock(1, true, Text("가-1")),
                new ListItemBlock(1, true, Text("가-2")),
                new ListItemBlock(0, true, Text("나")),
                new ListItemBlock(1, true, Text("나-1")),
            });

            Assert.Equal("1. 가\n    1. 가-1\n    2. 가-2\n2. 나\n    1. 나-1\n", md);
        }

        [Fact]
        public void Quote_prefixes_every_line()
        {
            var md = MarkdownRenderer.Render(new Block[] { new QuoteBlock(Text("발표자 노트\n- 둘째 줄")) });

            Assert.Equal("> 발표자 노트  \n> \\- 둘째 줄\n", md);
        }

        [Fact]
        public void Bold_italic_link_combination()
        {
            var md = MarkdownRenderer.Render(new Block[]
            {
                new ParagraphBlock(new List<Inline>
                {
                    new Inline { Text = "링크", Bold = true, Italic = true, Link = "https://example.com" },
                }),
                new RuleBlock(),
            });

            Assert.Equal("[***링크***](https://example.com)\n\n---\n", md);
        }
    }
}
