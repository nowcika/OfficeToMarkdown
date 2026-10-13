using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using OfficeToMarkdown.Core.Model;

namespace OfficeToMarkdown.Core.Word
{
    /// <summary>
    /// Word 의 Document.WordOpenXML (Flat OPC XML) 문자열을 블록 모델로 변환한다.
    /// COM 으로 문단·글자를 하나씩 읽는 방식보다 훨씬 빠르고, Office 없이 테스트할 수 있다.
    /// </summary>
    public sealed class WordOpenXmlConverter
    {
        private static readonly XNamespace Pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
        private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private static readonly XNamespace Wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        private static readonly XNamespace V = "urn:schemas-microsoft-com:vml";
        private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        private static readonly XNamespace Rels = "http://schemas.openxmlformats.org/package/2006/relationships";

        private const string MainDocumentPart = "/word/document.xml";

        private readonly Dictionary<string, XElement> _parts;
        private readonly Dictionary<string, Relationship> _relationships;
        private readonly Dictionary<string, StyleInfo> _styles;
        private readonly Dictionary<string, string> _numIdToAbstract;
        private readonly Dictionary<string, XElement> _abstractNums;
        private readonly string _assetFolder;
        private readonly DocumentContent _content = new DocumentContent();
        private readonly Dictionary<string, string> _exportedImages = new Dictionary<string, string>();

        private WordOpenXmlConverter(string flatOpcXml, string assetFolder)
        {
            _assetFolder = assetFolder.TrimEnd('/');
            _parts = XDocument.Parse(flatOpcXml).Root
                .Elements(Pkg + "part")
                .ToDictionary(p => (string)p.Attribute(Pkg + "name"), StringComparer.OrdinalIgnoreCase);

            _relationships = LoadRelationships();
            _styles = LoadStyles();

            var numbering = PartXml("/word/numbering.xml");
            _numIdToAbstract = numbering == null
                ? new Dictionary<string, string>()
                : numbering.Elements(W + "num").ToDictionary(
                    n => (string)n.Attribute(W + "numId"),
                    n => Val(n.Element(W + "abstractNumId")));
            _abstractNums = numbering == null
                ? new Dictionary<string, XElement>()
                : numbering.Elements(W + "abstractNum").ToDictionary(n => (string)n.Attribute(W + "abstractNumId"));
        }

        /// <param name="flatOpcXml">Document.WordOpenXML 값.</param>
        /// <param name="assetFolder">이미지를 저장할 폴더의 Markdown 파일 기준 상대 경로 (예: "보고서_files").</param>
        public static DocumentContent Convert(string flatOpcXml, string assetFolder)
        {
            if (flatOpcXml == null)
                throw new ArgumentNullException(nameof(flatOpcXml));
            if (string.IsNullOrEmpty(assetFolder))
                throw new ArgumentException("이미지 폴더 이름이 필요합니다.", nameof(assetFolder));

            var converter = new WordOpenXmlConverter(flatOpcXml, assetFolder);
            var body = converter.PartXml(MainDocumentPart)?.Element(W + "body")
                ?? throw new FormatException("WordOpenXML 에 본문(/word/document.xml)이 없습니다.");
            converter.ConvertContainer(body);
            return converter._content;
        }

        // ── 본문 구조 ─────────────────────────────────────────────

        private void ConvertContainer(XElement container)
        {
            foreach (var element in Unwrap(container))
            {
                if (element.Name == W + "p")
                {
                    var block = ConvertParagraph(element);
                    if (block != null)
                        _content.Blocks.Add(block);
                }
                else if (element.Name == W + "tbl")
                {
                    _content.Blocks.Add(ConvertTable(element));
                }
            }
        }

        /// <summary>콘텐츠 컨트롤(sdt), customXml 같은 래퍼를 벗겨 실제 자식 요소를 열거한다.</summary>
        private static IEnumerable<XElement> Unwrap(XElement container)
        {
            foreach (var child in container.Elements())
            {
                if (child.Name == W + "sdt")
                {
                    var inner = child.Element(W + "sdtContent");
                    if (inner != null)
                        foreach (var e in Unwrap(inner))
                            yield return e;
                }
                else if (child.Name == W + "customXml")
                {
                    foreach (var e in Unwrap(child))
                        yield return e;
                }
                else
                {
                    yield return child;
                }
            }
        }

        private Block ConvertParagraph(XElement p)
        {
            var inlines = ConvertInlines(p, new RunFormat());
            if (inlines.All(i => i.Kind != InlineKind.Image && string.IsNullOrWhiteSpace(i.Text)))
                return null;

            var pPr = p.Element(W + "pPr");
            var style = StyleOf(Val(pPr?.Element(W + "pStyle")));

            int? heading = OutlineLevel(pPr) ?? style?.HeadingLevel;
            if (heading.HasValue)
                return new HeadingBlock(heading.Value, inlines);

            var numPr = pPr?.Element(W + "numPr");
            string numId = Val(numPr?.Element(W + "numId")) ?? style?.NumId;
            if (numId != null && numId != "0")
            {
                string ilvl = Val(numPr?.Element(W + "ilvl")) ?? style?.NumLevel ?? "0";
                int level;
                int.TryParse(ilvl, out level);
                return new ListItemBlock(level, IsOrdered(numId, ilvl), inlines);
            }

            return new ParagraphBlock(inlines);
        }

        private TableBlock ConvertTable(XElement tbl)
        {
            var rows = new List<List<List<Inline>>>();
            foreach (var tr in Unwrap(tbl).Where(e => e.Name == W + "tr"))
            {
                var row = new List<List<Inline>>();
                foreach (var tc in Unwrap(tr).Where(e => e.Name == W + "tc"))
                {
                    var tcPr = tc.Element(W + "tcPr");
                    var vMerge = tcPr?.Element(W + "vMerge");
                    bool mergedFromAbove = vMerge != null && Val(vMerge) != "restart";

                    row.Add(mergedFromAbove ? new List<Inline>() : ConvertCell(tc));

                    int span;
                    if (int.TryParse(Val(tcPr?.Element(W + "gridSpan")), out span))
                        for (int i = 1; i < span; i++)
                            row.Add(new List<Inline>());
                }
                rows.Add(row);
            }
            return new TableBlock(rows);
        }

        /// <summary>셀 안의 문단(중첩 표 포함)을 줄바꿈으로 이어 붙인다.</summary>
        private List<Inline> ConvertCell(XElement tc)
        {
            var result = new List<Inline>();
            foreach (var p in tc.Descendants(W + "p"))
            {
                var inlines = ConvertInlines(p, new RunFormat());
                if (inlines.All(i => i.Kind != InlineKind.Image && string.IsNullOrWhiteSpace(i.Text)))
                    continue;
                if (result.Count > 0)
                    result.Add(Inline.Break());
                result.AddRange(inlines);
            }
            return result;
        }

        // ── 문단 안의 런 ─────────────────────────────────────────

        private List<Inline> ConvertInlines(XElement container, RunFormat format)
        {
            var result = new List<Inline>();
            AppendInlines(container, format, result);
            return result;
        }

        private void AppendInlines(XElement container, RunFormat format, List<Inline> result)
        {
            foreach (var child in container.Elements())
            {
                var name = child.Name;
                if (name == W + "r")
                {
                    AppendRun(child, format, result);
                }
                else if (name == W + "hyperlink")
                {
                    var rel = Relationship.Find(_relationships, (string)child.Attribute(R + "id"));
                    var link = rel != null && rel.External ? rel.Target : format.Link;
                    AppendInlines(child, format.WithLink(link), result);
                }
                else if (name == W + "ins" || name == W + "smartTag" || name == W + "fldSimple"
                         || name == W + "customXml" || name == W + "dir" || name == W + "bdo")
                {
                    AppendInlines(child, format, result);
                }
                else if (name == W + "sdt")
                {
                    var inner = child.Element(W + "sdtContent");
                    if (inner != null)
                        AppendInlines(inner, format, result);
                }
                else if (name == M + "oMath" || name == M + "oMathPara")
                {
                    var text = string.Concat(child.Descendants(M + "t").Select(t => t.Value));
                    if (text.Length > 0)
                        result.Add(format.Apply(Inline.Plain(text)));
                }
                // w:del, w:moveFrom(삭제된 변경 내용), w:pPr, 책갈피 등은 무시한다.
            }
        }

        private void AppendRun(XElement r, RunFormat inherited, List<Inline> result)
        {
            var format = inherited.WithRunProperties(r.Element(W + "rPr"), _styles);
            foreach (var child in r.Elements())
            {
                var name = child.Name;
                if (name == W + "t")
                    result.Add(format.Apply(Inline.Plain(child.Value)));
                else if (name == W + "tab" || name == W + "ptab")
                    result.Add(format.Apply(Inline.Plain(" ")));
                else if (name == W + "noBreakHyphen")
                    result.Add(format.Apply(Inline.Plain("-")));
                else if (name == W + "br")
                {
                    var type = (string)child.Attribute(W + "type");
                    if (type == null || type == "textWrapping")
                        result.Add(Inline.Break());
                }
                else if (name == W + "cr")
                    result.Add(Inline.Break());
                else if (name == W + "drawing" || name == W + "pict" || name == W + "object")
                    AppendImages(child, result);
            }
        }

        private void AppendImages(XElement element, List<Inline> result)
        {
            var docPr = element.Descendants(Wp + "docPr").FirstOrDefault();
            var alt = (string)docPr?.Attribute("descr");
            if (string.IsNullOrWhiteSpace(alt))
                alt = (string)docPr?.Attribute("title");
            if (string.IsNullOrWhiteSpace(alt))
                alt = "image";

            var ids = element.Descendants(A + "blip").Select(b => (string)b.Attribute(R + "embed"))
                .Concat(element.Descendants(V + "imagedata").Select(d => (string)d.Attribute(R + "id")))
                .Where(id => id != null);

            foreach (var id in ids)
            {
                var path = ExportImage(id);
                if (path != null)
                    result.Add(Inline.Image(path, alt));
            }
        }

        /// <summary>관계 ID 가 가리키는 이미지 파트를 리소스로 등록하고 Markdown 상대 경로를 돌려준다.</summary>
        private string ExportImage(string relationshipId)
        {
            var rel = Relationship.Find(_relationships, relationshipId);
            if (rel == null || rel.External)
                return null;

            string existing;
            if (_exportedImages.TryGetValue(rel.Target, out existing))
                return existing;

            XElement part;
            if (!_parts.TryGetValue(rel.Target, out part))
                return null;
            var data = part.Element(Pkg + "binaryData")?.Value;
            if (data == null)
                return null;

            var fileName = rel.Target.Substring(rel.Target.LastIndexOf('/') + 1);
            var path = _assetFolder + "/" + fileName;
            for (int n = 2; _content.Resources.ContainsKey(path); n++)
                path = _assetFolder + "/" + n + "_" + fileName;

            _content.Resources[path] = System.Convert.FromBase64String(data);
            _exportedImages[rel.Target] = path;
            return path;
        }

        // ── 스타일 / 번호 매기기 ─────────────────────────────────

        private static int? OutlineLevel(XElement pPr)
        {
            int level;
            if (int.TryParse(Val(pPr?.Element(W + "outlineLvl")), out level) && level < 9)
                return level + 1;
            return null;
        }

        private StyleInfo StyleOf(string styleId)
        {
            StyleInfo style;
            return styleId != null && _styles.TryGetValue(styleId, out style) ? style : null;
        }

        private bool IsOrdered(string numId, string ilvl)
        {
            string abstractId;
            XElement abstractNum;
            if (!_numIdToAbstract.TryGetValue(numId, out abstractId)
                || abstractId == null
                || !_abstractNums.TryGetValue(abstractId, out abstractNum))
                return false;

            var lvl = abstractNum.Elements(W + "lvl").FirstOrDefault(l => (string)l.Attribute(W + "ilvl") == ilvl);
            var numFmt = Val(lvl?.Element(W + "numFmt"));
            return numFmt != null && numFmt != "bullet" && numFmt != "none";
        }

        private Dictionary<string, StyleInfo> LoadStyles()
        {
            var raw = PartXml("/word/styles.xml")?.Elements(W + "style")
                .Where(s => s.Attribute(W + "styleId") != null)
                .GroupBy(s => (string)s.Attribute(W + "styleId"))
                .ToDictionary(g => g.Key, g => g.First())
                ?? new Dictionary<string, XElement>();

            var result = new Dictionary<string, StyleInfo>();
            foreach (var id in raw.Keys)
                result[id] = StyleInfo.Resolve(id, raw);
            return result;
        }

        private Dictionary<string, Relationship> LoadRelationships()
        {
            var rels = PartXml("/word/_rels/document.xml.rels");
            var result = new Dictionary<string, Relationship>();
            if (rels == null)
                return result;

            foreach (var rel in rels.Elements(Rels + "Relationship"))
            {
                var id = (string)rel.Attribute("Id");
                var target = (string)rel.Attribute("Target");
                if (id == null || target == null)
                    continue;
                bool external = (string)rel.Attribute("TargetMode") == "External";
                result[id] = new Relationship(external ? target : ResolvePartName("/word/", target), external);
            }
            return result;
        }

        private static string ResolvePartName(string baseFolder, string target)
        {
            var path = target.StartsWith("/") ? target : baseFolder + target;
            var segments = new List<string>();
            foreach (var segment in path.Split('/'))
            {
                if (segment.Length == 0 || segment == ".")
                    continue;
                if (segment == "..")
                {
                    if (segments.Count > 0)
                        segments.RemoveAt(segments.Count - 1);
                }
                else
                {
                    segments.Add(segment);
                }
            }
            return "/" + string.Join("/", segments);
        }

        private XElement PartXml(string name)
        {
            XElement part;
            return _parts.TryGetValue(name, out part) ? part.Element(Pkg + "xmlData")?.Elements().FirstOrDefault() : null;
        }

        private static string Val(XElement element) => (string)element?.Attribute(W + "val");

        private static bool? Toggle(XElement element)
        {
            if (element == null)
                return null;
            var val = Val(element);
            return val == null || !(val == "0" || val == "false" || val == "off");
        }

        // ── 보조 타입 ────────────────────────────────────────────

        private sealed class Relationship
        {
            public Relationship(string target, bool external)
            {
                Target = target;
                External = external;
            }

            public string Target { get; }
            public bool External { get; }

            public static Relationship Find(Dictionary<string, Relationship> map, string id)
            {
                Relationship rel;
                return id != null && map.TryGetValue(id, out rel) ? rel : null;
            }
        }

        private sealed class StyleInfo
        {
            public int? HeadingLevel { get; private set; }
            public string NumId { get; private set; }
            public string NumLevel { get; private set; }
            public bool? Bold { get; private set; }
            public bool? Italic { get; private set; }

            /// <summary>basedOn 체인을 따라가며 제목 수준·번호 매기기·굵게/기울임을 결정한다.</summary>
            public static StyleInfo Resolve(string id, Dictionary<string, XElement> styles)
            {
                var info = new StyleInfo();
                var visited = new HashSet<string>();
                XElement style;
                for (var current = id; current != null && visited.Add(current) && styles.TryGetValue(current, out style);
                     current = Val(style.Element(W + "basedOn")))
                {
                    var name = (Val(style.Element(W + "name")) ?? "").ToLowerInvariant();
                    var pPr = style.Element(W + "pPr");
                    var rPr = style.Element(W + "rPr");

                    if (info.HeadingLevel == null)
                    {
                        int level;
                        if (name == "title")
                            info.HeadingLevel = 1;
                        else if (name.StartsWith("heading ") && int.TryParse(name.Substring(8), out level))
                            info.HeadingLevel = level;
                        else
                            info.HeadingLevel = OutlineLevel(pPr);
                    }

                    var numPr = pPr?.Element(W + "numPr");
                    if (info.NumId == null)
                        info.NumId = Val(numPr?.Element(W + "numId"));
                    if (info.NumLevel == null)
                        info.NumLevel = Val(numPr?.Element(W + "ilvl"));
                    if (info.Bold == null)
                        info.Bold = Toggle(rPr?.Element(W + "b"));
                    if (info.Italic == null)
                        info.Italic = Toggle(rPr?.Element(W + "i"));
                }
                return info;
            }
        }

        private sealed class RunFormat
        {
            public bool Bold { get; private set; }
            public bool Italic { get; private set; }
            public bool Strike { get; private set; }
            public string Link { get; private set; }

            public RunFormat WithLink(string link) =>
                new RunFormat { Bold = Bold, Italic = Italic, Strike = Strike, Link = link };

            /// <summary>문자 스타일(rStyle) → 직접 서식 순서로 적용한다.</summary>
            public RunFormat WithRunProperties(XElement rPr, Dictionary<string, StyleInfo> styles)
            {
                var result = new RunFormat { Link = Link };
                StyleInfo style;
                var styleId = Val(rPr?.Element(W + "rStyle"));
                if (styleId == null || !styles.TryGetValue(styleId, out style))
                    style = null;

                result.Bold = Toggle(rPr?.Element(W + "b")) ?? style?.Bold ?? false;
                result.Italic = Toggle(rPr?.Element(W + "i")) ?? style?.Italic ?? false;
                result.Strike = (Toggle(rPr?.Element(W + "strike")) ?? false) || (Toggle(rPr?.Element(W + "dstrike")) ?? false);
                return result;
            }

            public Inline Apply(Inline inline)
            {
                inline.Bold = Bold;
                inline.Italic = Italic;
                inline.Strike = Strike;
                inline.Link = Link;
                return inline;
            }
        }
    }
}
