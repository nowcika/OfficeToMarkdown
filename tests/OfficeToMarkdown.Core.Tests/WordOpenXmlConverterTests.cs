using System;
using System.Linq;
using System.Text;
using OfficeToMarkdown.Core;
using OfficeToMarkdown.Core.Word;
using Xunit;

namespace OfficeToMarkdown.Core.Tests
{
    public class WordOpenXmlConverterTests
    {
        private const string Ns =
            "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
            "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
            "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\"";

        private const string Styles = @"
<w:styles " + Ns + @">
  <w:style w:type=""paragraph"" w:styleId=""1""><w:name w:val=""heading 1""/></w:style>
  <w:style w:type=""paragraph"" w:styleId=""2""><w:name w:val=""heading 2""/><w:basedOn w:val=""1""/></w:style>
  <w:style w:type=""paragraph"" w:styleId=""MyHeading""><w:name w:val=""Custom""/><w:basedOn w:val=""2""/></w:style>
  <w:style w:type=""character"" w:styleId=""Strong""><w:name w:val=""Strong""/><w:rPr><w:b/></w:rPr></w:style>
</w:styles>";

        private const string Numbering = @"
<w:numbering " + Ns + @">
  <w:abstractNum w:abstractNumId=""10"">
    <w:lvl w:ilvl=""0""><w:numFmt w:val=""bullet""/></w:lvl>
    <w:lvl w:ilvl=""1""><w:numFmt w:val=""bullet""/></w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId=""20"">
    <w:lvl w:ilvl=""0""><w:numFmt w:val=""decimal""/></w:lvl>
  </w:abstractNum>
  <w:num w:numId=""1""><w:abstractNumId w:val=""10""/></w:num>
  <w:num w:numId=""2""><w:abstractNumId w:val=""20""/></w:num>
</w:numbering>";

        private const string Rels = @"
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rIdLink"" Type=""hyperlink"" Target=""https://example.com/a b"" TargetMode=""External""/>
  <Relationship Id=""rIdImg"" Type=""image"" Target=""media/image1.png""/>
</Relationships>";

        private static string Package(string bodyXml, params (string Name, byte[] Data)[] binaries)
        {
            var sb = new StringBuilder();
            sb.Append("<pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">");
            void XmlPart(string name, string xml) =>
                sb.Append($"<pkg:part pkg:name=\"{name}\"><pkg:xmlData>{xml}</pkg:xmlData></pkg:part>");
            XmlPart("/word/document.xml", $"<w:document {Ns}><w:body>{bodyXml}</w:body></w:document>");
            XmlPart("/word/styles.xml", Styles);
            XmlPart("/word/numbering.xml", Numbering);
            XmlPart("/word/_rels/document.xml.rels", Rels);
            foreach (var (name, data) in binaries)
                sb.Append($"<pkg:part pkg:name=\"{name}\"><pkg:binaryData>{Convert.ToBase64String(data)}</pkg:binaryData></pkg:part>");
            sb.Append("</pkg:package>");
            return sb.ToString();
        }

        private static string ToMarkdown(string bodyXml, params (string, byte[])[] binaries) =>
            MarkdownRenderer.Render(WordOpenXmlConverter.Convert(Package(bodyXml, binaries), "doc_files").Blocks);

        [Fact]
        public void Headings_are_resolved_from_style_names_and_basedOn_chain()
        {
            var md = ToMarkdown(@"
<w:p><w:pPr><w:pStyle w:val=""1""/></w:pPr><w:r><w:t>제목</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val=""MyHeading""/></w:pPr><w:r><w:t>소제목</w:t></w:r></w:p>
<w:p><w:r><w:t>본문</w:t></w:r></w:p>");

            Assert.Equal("# 제목\n\n## 소제목\n\n본문\n", md);
        }

        [Fact]
        public void Runs_with_same_format_are_merged_and_whitespace_moved_outside_markers()
        {
            var md = ToMarkdown(@"
<w:p>
  <w:r><w:t xml:space=""preserve"">일반 </w:t></w:r>
  <w:r><w:rPr><w:b/></w:rPr><w:t xml:space=""preserve"">굵게 </w:t></w:r>
  <w:r><w:rPr><w:b/></w:rPr><w:t>이어짐</w:t></w:r>
  <w:r><w:rPr><w:b w:val=""0""/><w:i/></w:rPr><w:t xml:space=""preserve""> 기울임</w:t></w:r>
  <w:r><w:rPr><w:rStyle w:val=""Strong""/></w:rPr><w:t>스타일</w:t></w:r>
  <w:r><w:rPr><w:strike/></w:rPr><w:t>취소</w:t></w:r>
</w:p>");

            Assert.Equal("일반 **굵게 이어짐** *기울임***스타일**~~취소~~\n", md);
        }

        [Fact]
        public void Lists_use_numbering_definitions_for_bullet_vs_ordered()
        {
            var md = ToMarkdown(@"
<w:p><w:pPr><w:numPr><w:ilvl w:val=""0""/><w:numId w:val=""1""/></w:numPr></w:pPr><w:r><w:t>사과</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val=""1""/><w:numId w:val=""1""/></w:numPr></w:pPr><w:r><w:t>홍옥</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val=""0""/><w:numId w:val=""1""/></w:numPr></w:pPr><w:r><w:t>배</w:t></w:r></w:p>
<w:p><w:r><w:t>중간 문단</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val=""0""/><w:numId w:val=""2""/></w:numPr></w:pPr><w:r><w:t>첫째</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val=""0""/><w:numId w:val=""2""/></w:numPr></w:pPr><w:r><w:t>둘째</w:t></w:r></w:p>");

            Assert.Equal("- 사과\n    - 홍옥\n- 배\n\n중간 문단\n\n1. 첫째\n2. 둘째\n", md);
        }

        [Fact]
        public void External_hyperlinks_become_markdown_links()
        {
            var md = ToMarkdown(@"
<w:p><w:r><w:t xml:space=""preserve"">참고: </w:t></w:r>
  <w:hyperlink r:id=""rIdLink""><w:r><w:t>예제</w:t></w:r></w:hyperlink>
  <w:hyperlink w:anchor=""_Toc1""><w:r><w:t xml:space=""preserve""> 목차</w:t></w:r></w:hyperlink>
</w:p>");

            Assert.Equal("참고: [예제](<https://example.com/a b>) 목차\n", md);
        }

        [Fact]
        public void Tables_handle_header_merges_and_pipes()
        {
            var md = ToMarkdown(@"
<w:tbl>
  <w:tr>
    <w:tc><w:p><w:r><w:t>이름</w:t></w:r></w:p></w:tc>
    <w:tc><w:tcPr><w:gridSpan w:val=""2""/></w:tcPr><w:p><w:r><w:t>값</w:t></w:r></w:p></w:tc>
  </w:tr>
  <w:tr>
    <w:tc><w:tcPr><w:vMerge w:val=""restart""/></w:tcPr><w:p><w:r><w:t>a|b</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>1</w:t></w:r></w:p><w:p><w:r><w:t>2</w:t></w:r></w:p></w:tc>
    <w:tc><w:p/></w:tc>
  </w:tr>
  <w:tr>
    <w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:r><w:t>무시</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>3</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>4</w:t></w:r></w:p></w:tc>
  </w:tr>
</w:tbl>");

            Assert.Equal(
                "| 이름 | 값 |  |\n" +
                "| --- | --- | --- |\n" +
                "| a\\|b | 1<br>2 |  |\n" +
                "|  | 3 | 4 |\n", md);
        }

        [Fact]
        public void Images_are_exported_as_resources_and_referenced_relatively()
        {
            var png = new byte[] { 0x89, 0x50, 0x4E, 0x47 };
            var package = Package(@"
<w:p><w:r><w:drawing><wp:inline><wp:docPr id=""1"" name=""그림 1"" descr=""로고""/>
  <a:graphic><a:graphicData><a:blip r:embed=""rIdImg""/></a:graphicData></a:graphic>
</wp:inline></w:drawing></w:r></w:p>", ("/word/media/image1.png", png));

            var content = WordOpenXmlConverter.Convert(package, "my doc_files");

            Assert.Equal(png, content.Resources["my doc_files/image1.png"]);
            Assert.Equal("![로고](<my doc_files/image1.png>)\n", MarkdownRenderer.Render(content.Blocks));
        }

        [Fact]
        public void Content_controls_tracked_deletions_and_empty_paragraphs_are_handled()
        {
            var md = ToMarkdown(@"
<w:sdt><w:sdtContent><w:p><w:r><w:t>컨트롤 안</w:t></w:r></w:p></w:sdtContent></w:sdt>
<w:p/>
<w:p><w:r><w:br w:type=""page""/></w:r></w:p>
<w:p><w:r><w:t>남김</w:t></w:r><w:del><w:r><w:delText>삭제됨</w:delText></w:r></w:del><w:ins><w:r><w:t>추가됨</w:t></w:r></w:ins></w:p>");

            Assert.Equal("컨트롤 안\n\n남김추가됨\n", md);
        }

        [Fact]
        public void Markdown_special_characters_are_escaped()
        {
            var md = ToMarkdown(@"
<w:p><w:r><w:t># 해시로 시작</w:t></w:r></w:p>
<w:p><w:r><w:t>1. 숫자로 시작</w:t></w:r></w:p>
<w:p><w:r><w:t>a*b_c [x] &lt;tag&gt;</w:t></w:r></w:p>");

            Assert.Equal("\\# 해시로 시작\n\n1\\. 숫자로 시작\n\na\\*b\\_c \\[x\\] \\<tag\\>\n", md);
        }

        [Fact]
        public void Line_breaks_in_paragraph_become_hard_breaks()
        {
            var md = ToMarkdown(@"<w:p><w:r><w:t>첫 줄</w:t><w:br/><w:t>둘째 줄</w:t></w:r></w:p>");

            Assert.Equal("첫 줄  \n둘째 줄\n", md);
        }
    }
}
