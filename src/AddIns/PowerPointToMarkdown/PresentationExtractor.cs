using System.Collections.Generic;
using System.Linq;
using OfficeToMarkdown.Core.Model;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointToMarkdown
{
    /// <summary>
    /// 슬라이드마다 "## 번호. 제목" 과 도형 내용(텍스트·표·SmartArt), 발표자 노트를 블록으로 만든다.
    /// 도형은 위→아래, 왼쪽→오른쪽 순서로 읽는다.
    /// </summary>
    internal static class PresentationExtractor
    {
        public static DocumentContent Extract(PowerPoint.Presentation presentation)
        {
            var content = new DocumentContent();
            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                if (content.Blocks.Count > 0)
                    content.Blocks.Add(new RuleBlock());

                var titleShape = slide.Shapes.HasTitle == Office.MsoTriState.msoTrue ? slide.Shapes.Title : null;
                var title = titleShape != null ? PlainText(titleShape) : "";
                content.Blocks.Add(new HeadingBlock(2, title.Length > 0
                    ? slide.SlideIndex + ". " + title
                    : "슬라이드 " + slide.SlideIndex));

                var shapes = slide.Shapes.Cast<PowerPoint.Shape>()
                    .Where(s => titleShape == null || s.Id != titleShape.Id);
                foreach (var shape in InReadingOrder(shapes))
                    AppendShape(shape, content.Blocks);

                AppendNotes(slide, content.Blocks);
            }
            return content;
        }

        private static IEnumerable<PowerPoint.Shape> InReadingOrder(IEnumerable<PowerPoint.Shape> shapes) =>
            shapes.OrderBy(s => s.Top).ThenBy(s => s.Left);

        private static void AppendShape(PowerPoint.Shape shape, List<Block> blocks)
        {
            if (shape.Type == Office.MsoShapeType.msoGroup)
            {
                foreach (var item in InReadingOrder(shape.GroupItems.Cast<PowerPoint.Shape>()))
                    AppendShape(item, blocks);
            }
            else if (shape.HasTable == Office.MsoTriState.msoTrue)
            {
                blocks.Add(ConvertTable(shape.Table));
            }
            else if (shape.HasSmartArt == Office.MsoTriState.msoTrue)
            {
                foreach (Office.SmartArtNode node in shape.SmartArt.AllNodes)
                {
                    var text = node.TextFrame2.TextRange.Text.Trim();
                    if (text.Length > 0)
                        blocks.Add(new ListItemBlock(node.Level - 1, false, Inline.FromPlainText(text)));
                }
            }
            else if (shape.HasTextFrame == Office.MsoTriState.msoTrue
                     && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
            {
                AppendTextRange(shape.TextFrame.TextRange, blocks);
            }
            // 그림·차트·미디어는 현재 건너뛴다.
        }

        /// <summary>문단 단위로 읽어 글머리 기호가 있으면 목록 항목, 없으면 일반 문단으로 만든다.</summary>
        private static void AppendTextRange(PowerPoint.TextRange range, List<Block> blocks)
        {
            int paragraphCount = range.Paragraphs().Count;
            for (int i = 1; i <= paragraphCount; i++)
            {
                var paragraph = range.Paragraphs(i, 1);
                var inlines = ConvertRuns(paragraph);
                if (inlines.All(x => string.IsNullOrWhiteSpace(x.Text)))
                    continue;

                var bullet = paragraph.ParagraphFormat.Bullet;
                if (bullet.Visible == Office.MsoTriState.msoTrue)
                {
                    bool ordered = bullet.Type == PowerPoint.PpBulletType.ppBulletNumbered;
                    blocks.Add(new ListItemBlock(paragraph.IndentLevel - 1, ordered, inlines));
                }
                else
                {
                    blocks.Add(new ParagraphBlock(inlines));
                }
            }
        }

        private static List<Inline> ConvertRuns(PowerPoint.TextRange paragraph)
        {
            var result = new List<Inline>();
            int runCount = paragraph.Runs().Count;
            for (int i = 1; i <= runCount; i++)
            {
                var run = paragraph.Runs(i, 1);
                var text = run.Text.TrimEnd('\r');
                if (text.Length == 0)
                    continue;

                bool bold = run.Font.Bold == Office.MsoTriState.msoTrue;
                bool italic = run.Font.Italic == Office.MsoTriState.msoTrue;
                string link = HyperlinkOf(run);
                foreach (var inline in Inline.FromPlainText(text)) // \v(Shift+Enter) → 줄바꿈
                {
                    inline.Bold = bold;
                    inline.Italic = italic;
                    inline.Link = link;
                    result.Add(inline);
                }
            }
            return result;
        }

        private static string HyperlinkOf(PowerPoint.TextRange run)
        {
            var action = run.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick];
            if (action.Action != PowerPoint.PpActionType.ppActionHyperlink)
                return null;
            var address = action.Hyperlink.Address;
            return string.IsNullOrEmpty(address) ? null : address;
        }

        private static TableBlock ConvertTable(PowerPoint.Table table)
        {
            var rows = new List<List<List<Inline>>>();
            for (int r = 1; r <= table.Rows.Count; r++)
            {
                var row = new List<List<Inline>>();
                for (int c = 1; c <= table.Columns.Count; c++)
                    row.Add(Inline.FromPlainText(PlainText(table.Cell(r, c).Shape)));
                rows.Add(row);
            }
            return new TableBlock(rows);
        }

        private static void AppendNotes(PowerPoint.Slide slide, List<Block> blocks)
        {
            if (slide.HasNotesPage != Office.MsoTriState.msoTrue)
                return;

            foreach (PowerPoint.Shape placeholder in slide.NotesPage.Shapes.Placeholders)
            {
                if (placeholder.PlaceholderFormat.Type != PowerPoint.PpPlaceholderType.ppPlaceholderBody)
                    continue;
                var text = PlainText(placeholder);
                if (text.Length == 0)
                    continue;

                var inlines = new List<Inline> { new Inline { Text = "발표자 노트", Bold = true }, Inline.Break() };
                inlines.AddRange(Inline.FromPlainText(text));
                blocks.Add(new QuoteBlock(inlines));
            }
        }

        private static string PlainText(PowerPoint.Shape shape)
        {
            if (shape.HasTextFrame != Office.MsoTriState.msoTrue || shape.TextFrame.HasText != Office.MsoTriState.msoTrue)
                return "";
            return shape.TextFrame.TextRange.Text.Trim('\r', '\v', ' ');
        }
    }
}
