using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using OfficeToMarkdown.Core;
using OfficeToMarkdown.Core.Model;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointToMarkdown
{
    /// <summary>
    /// 슬라이드마다 "## 번호. 제목" 과 도형 내용(텍스트·표·SmartArt·그림·차트), 발표자 노트를 블록으로 만든다.
    /// 도형은 위→아래, 왼쪽→오른쪽 순서로 읽는다.
    /// </summary>
    internal sealed class PresentationExtractor
    {
        // Office 2016 이상의 SVG 아이콘 등(msoGraphic). Office 2013 interop 의 MsoShapeType 에는 없는 값이다.
        private const int MsoGraphic = 28;

        private readonly string _assetFolder;
        private readonly DocumentContent _content = new DocumentContent();
        private int _slideNumber;
        private int _imageNumber;

        private PresentationExtractor(string assetFolder)
        {
            _assetFolder = assetFolder.TrimEnd('/');
        }

        /// <param name="assetFolder">그림을 저장할 폴더의 Markdown 파일 기준 상대 경로 (예: "발표_files").</param>
        public static DocumentContent Extract(PowerPoint.Presentation presentation, string assetFolder)
        {
            var extractor = new PresentationExtractor(assetFolder);
            foreach (PowerPoint.Slide slide in presentation.Slides)
                extractor.AppendSlide(slide);
            return extractor._content;
        }

        private void AppendSlide(PowerPoint.Slide slide)
        {
            var blocks = _content.Blocks;
            _slideNumber = slide.SlideIndex;
            _imageNumber = 0;

            if (blocks.Count > 0)
                blocks.Add(new RuleBlock());

            var titleShape = slide.Shapes.HasTitle == Office.MsoTriState.msoTrue ? slide.Shapes.Title : null;
            var title = titleShape != null ? PlainText(titleShape) : "";
            blocks.Add(new HeadingBlock(2, title.Length > 0
                ? _slideNumber + ". " + title
                : "슬라이드 " + _slideNumber));

            var shapes = slide.Shapes.Cast<PowerPoint.Shape>()
                .Where(s => titleShape == null || s.Id != titleShape.Id);
            foreach (var shape in InReadingOrder(shapes))
                AppendShape(shape, blocks);

            AppendNotes(slide, blocks);
        }

        private static IEnumerable<PowerPoint.Shape> InReadingOrder(IEnumerable<PowerPoint.Shape> shapes) =>
            shapes.OrderBy(s => s.Top).ThenBy(s => s.Left);

        private void AppendShape(PowerPoint.Shape shape, List<Block> blocks)
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
            else if (IsImageLike(shape))
            {
                var path = ExportImage(shape);
                if (path != null)
                    blocks.Add(new ParagraphBlock(new List<Inline> { Inline.Image(path, AltText(shape)) }));
            }
            else if (shape.HasTextFrame == Office.MsoTriState.msoTrue
                     && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
            {
                AppendTextRange(shape.TextFrame.TextRange, blocks);
            }
            // 미디어(동영상·오디오)와 그 밖의 도형은 건너뛴다.
        }

        /// <summary>그림, 연결된 그림, 그림이 들어간 개체 틀, SVG 아이콘, 차트를 이미지로 내보낸다.</summary>
        private static bool IsImageLike(PowerPoint.Shape shape)
        {
            var type = shape.Type;
            if (type == Office.MsoShapeType.msoPicture || type == Office.MsoShapeType.msoLinkedPicture
                || (int)type == MsoGraphic || shape.HasChart == Office.MsoTriState.msoTrue)
                return true;
            return type == Office.MsoShapeType.msoPlaceholder
                   && shape.PlaceholderFormat.ContainedType == Office.MsoShapeType.msoPicture;
        }

        /// <summary>
        /// 문서화되지 않은 Shape.Export 로 도형을 PNG 로 렌더링해 리소스에 등록한다.
        /// 자르기·크기 조절이 반영된 화면 모습 그대로 저장되며, 실패한 도형은 건너뛴다.
        /// </summary>
        private string ExportImage(PowerPoint.Shape shape)
        {
            byte[] data;
            try
            {
                data = TempFileExport.Capture(file => shape.Export(file, PowerPoint.PpShapeFormat.ppShapeFormatPNG), ".png");
            }
            catch (COMException)
            {
                return null;
            }
            if (data == null)
                return null;

            _imageNumber++;
            var path = _assetFolder + "/slide" + _slideNumber + "_" + _imageNumber + ".png";
            _content.Resources[path] = data;
            return path;
        }

        private static string AltText(PowerPoint.Shape shape)
        {
            var alt = (shape.AlternativeText ?? "").Replace("\r", " ").Replace("\n", " ").Trim();
            if (alt.Length == 0)
                alt = shape.HasChart == Office.MsoTriState.msoTrue ? "chart" : "image";
            return alt;
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
