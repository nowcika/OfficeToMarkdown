using System.Collections.Generic;

namespace OfficeToMarkdown.Core.Model
{
    /// <summary>Markdown 으로 렌더링되는 블록 단위 요소.</summary>
    public abstract class Block
    {
    }

    public sealed class HeadingBlock : Block
    {
        public HeadingBlock(int level, List<Inline> inlines)
        {
            Level = level;
            Inlines = inlines;
        }

        public HeadingBlock(int level, string text)
            : this(level, Inline.FromPlainText(text))
        {
        }

        public int Level { get; }
        public List<Inline> Inlines { get; }
    }

    public sealed class ParagraphBlock : Block
    {
        public ParagraphBlock(List<Inline> inlines)
        {
            Inlines = inlines;
        }

        public List<Inline> Inlines { get; }
    }

    public sealed class ListItemBlock : Block
    {
        public ListItemBlock(int level, bool ordered, List<Inline> inlines)
        {
            Level = level;
            Ordered = ordered;
            Inlines = inlines;
        }

        /// <summary>0부터 시작하는 들여쓰기 수준.</summary>
        public int Level { get; }
        public bool Ordered { get; }
        public List<Inline> Inlines { get; }
    }

    public sealed class TableBlock : Block
    {
        public TableBlock(List<List<List<Inline>>> rows)
        {
            Rows = rows;
        }

        /// <summary>행 → 셀 → 셀 내용. 첫 행은 머리글로 렌더링된다.</summary>
        public List<List<List<Inline>>> Rows { get; }
    }

    public sealed class QuoteBlock : Block
    {
        public QuoteBlock(List<Inline> inlines)
        {
            Inlines = inlines;
        }

        public List<Inline> Inlines { get; }
    }

    public sealed class RuleBlock : Block
    {
    }
}
