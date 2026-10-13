using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeToMarkdown.Core.Model;

namespace OfficeToMarkdown.Core
{
    /// <summary>블록 모델을 GitHub Flavored Markdown 문자열로 렌더링한다.</summary>
    public static class MarkdownRenderer
    {
        private enum InlineMode
        {
            Block,   // 일반 문단: 줄바꿈 → "  \n"
            Heading, // 제목: 줄바꿈 → 공백
            Cell,    // 표 셀/목록 항목: 줄바꿈 → <br>, 표에서는 '|' 이스케이프
        }

        private const int ListIndentWidth = 4;

        private static readonly Regex OrderedMarker = new Regex(@"^(\d{1,9})([.)])(\s|$)", RegexOptions.Compiled);

        public static string Render(IEnumerable<Block> blocks)
        {
            var sb = new StringBuilder();
            var listCounters = new List<int>();
            int prevListLevel = -1;
            Block prev = null;

            foreach (var block in blocks)
            {
                string text;
                var item = block as ListItemBlock;
                if (item != null)
                {
                    // 수준이 한 번에 두 단계 이상 깊어지면 코드 블록으로 해석되므로 한 단계씩만 허용한다.
                    int level = Math.Max(0, Math.Min(item.Level, prevListLevel + 1));
                    text = RenderListItem(item, level, listCounters);
                    if (text == null)
                        continue;
                    prevListLevel = level;
                }
                else
                {
                    text = RenderBlock(block);
                    if (text == null)
                        continue;
                    listCounters.Clear();
                    prevListLevel = -1;
                }

                if (prev != null)
                    sb.Append(prev is ListItemBlock && item != null ? "\n" : "\n\n");
                sb.Append(text);
                prev = block;
            }

            if (sb.Length > 0)
                sb.Append('\n');
            return sb.ToString();
        }

        private static string RenderBlock(Block block)
        {
            var heading = block as HeadingBlock;
            if (heading != null)
            {
                var text = RenderInlines(heading.Inlines, InlineMode.Heading);
                if (text.Length == 0)
                    return null;
                return new string('#', Math.Max(1, Math.Min(6, heading.Level))) + " " + text;
            }

            var paragraph = block as ParagraphBlock;
            if (paragraph != null)
            {
                var text = RenderInlines(paragraph.Inlines, InlineMode.Block);
                return text.Length == 0 ? null : EscapeLineStarts(text);
            }

            var quote = block as QuoteBlock;
            if (quote != null)
            {
                var text = RenderInlines(quote.Inlines, InlineMode.Block);
                if (text.Length == 0)
                    return null;
                return string.Join("\n", EscapeLineStarts(text).Split('\n').Select(l => "> " + l));
            }

            var table = block as TableBlock;
            if (table != null)
                return RenderTable(table);

            if (block is RuleBlock)
                return "---";

            throw new NotSupportedException("알 수 없는 블록 형식: " + block.GetType().Name);
        }

        private static string RenderListItem(ListItemBlock item, int level, List<int> counters)
        {
            var text = RenderInlines(item.Inlines, InlineMode.Cell);
            if (text.Length == 0)
                return null;

            // 더 얕은 수준이 나오면 그보다 깊은 수준의 번호는 다시 시작한다.
            if (counters.Count > level + 1)
                counters.RemoveRange(level + 1, counters.Count - level - 1);
            while (counters.Count <= level)
                counters.Add(0);
            counters[level]++;

            var marker = item.Ordered ? counters[level] + "." : "-";
            return new string(' ', level * ListIndentWidth) + marker + " " + text;
        }

        private static string RenderTable(TableBlock table)
        {
            var rows = table.Rows
                .Select(r => r.Select(c => RenderInlines(c, InlineMode.Cell)).ToList())
                .ToList();
            int columns = rows.Count == 0 ? 0 : rows.Max(r => r.Count);
            if (columns == 0)
                return null;

            var sb = new StringBuilder();
            for (int i = 0; i < rows.Count; i++)
            {
                var cells = rows[i];
                while (cells.Count < columns)
                    cells.Add("");
                if (i > 0)
                    sb.Append('\n');
                sb.Append("| ").Append(string.Join(" | ", cells)).Append(" |");
                if (i == 0)
                    sb.Append('\n').Append('|').Append(string.Concat(Enumerable.Repeat(" --- |", columns)));
            }
            return sb.ToString();
        }

        private static string RenderInlines(List<Inline> inlines, InlineMode mode)
        {
            var sb = new StringBuilder();
            int i = 0;
            while (i < inlines.Count)
            {
                var current = inlines[i];
                if (current.Kind == InlineKind.LineBreak)
                {
                    sb.Append(mode == InlineMode.Block ? "  \n" : mode == InlineMode.Heading ? " " : "<br>");
                    i++;
                    continue;
                }
                if (current.Kind == InlineKind.Image)
                {
                    sb.Append("![").Append(Escape(current.Text, mode == InlineMode.Cell)).Append("](")
                      .Append(FormatDestination(current.ImagePath)).Append(')');
                    i++;
                    continue;
                }

                // 서식이 같은 연속 조각을 합쳐서 "**a****b**" 같은 출력을 피한다.
                var text = new StringBuilder();
                int j = i;
                while (j < inlines.Count && inlines[j].Kind == InlineKind.Text && SameFormat(inlines[j], current))
                {
                    text.Append(inlines[j].Text);
                    j++;
                }
                sb.Append(FormatRun(text.ToString(), current, mode));
                i = j;
            }

            return TrimLines(sb.ToString());
        }

        private static bool SameFormat(Inline a, Inline b) =>
            a.Bold == b.Bold && a.Italic == b.Italic && a.Strike == b.Strike && a.Link == b.Link;

        private static string FormatRun(string raw, Inline format, InlineMode mode)
        {
            var escaped = Escape(raw, mode == InlineMode.Cell);
            if (!format.Bold && !format.Italic && !format.Strike && format.Link == null)
                return escaped;

            // 강조 기호 안쪽에 공백이 붙으면 강조로 인식되지 않으므로 공백은 바깥으로 뺀다.
            var core = escaped.Trim();
            if (core.Length == 0)
                return escaped;
            int lead = escaped.Length - escaped.TrimStart().Length;
            int trail = escaped.Length - escaped.TrimEnd().Length;

            var marker = format.Bold && format.Italic ? "***" : format.Bold ? "**" : format.Italic ? "*" : "";
            core = marker + core + marker;
            if (format.Strike)
                core = "~~" + core + "~~";
            if (format.Link != null)
                core = "[" + core + "](" + FormatDestination(format.Link) + ")";

            return escaped.Substring(0, lead) + core + escaped.Substring(escaped.Length - trail);
        }

        private static string Escape(string text, bool inTable)
        {
            var sb = new StringBuilder(text.Length);
            foreach (var c in text)
            {
                switch (c)
                {
                    case '\\':
                    case '`':
                    case '*':
                    case '_':
                    case '[':
                    case ']':
                    case '<':
                    case '>':
                    case '~':
                        sb.Append('\\').Append(c);
                        break;
                    case '|':
                        sb.Append(inTable ? "\\|" : "|");
                        break;
                    case '\t':
                        sb.Append(' ');
                        break;
                    default:
                        sb.Append(c);
                        break;
                }
            }
            return sb.ToString();
        }

        private static string FormatDestination(string url)
        {
            if (url.IndexOfAny(new[] { ' ', '(', ')', '<', '>' }) < 0)
                return url;
            return "<" + url.Replace("<", "%3C").Replace(">", "%3E") + ">";
        }

        /// <summary>줄 맨 앞에서 제목·목록·인용으로 오인될 수 있는 문자를 이스케이프한다.</summary>
        private static string EscapeLineStarts(string text)
        {
            var lines = text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                var line = lines[i];
                if (line.Length == 0)
                    continue;
                char c = line[0];
                if (c == '#' || c == '-' || c == '+' || c == '=')
                {
                    lines[i] = "\\" + line;
                    continue;
                }
                var m = OrderedMarker.Match(line);
                if (m.Success)
                    lines[i] = m.Groups[1].Value + "\\" + line.Substring(m.Groups[1].Length);
            }
            return string.Join("\n", lines);
        }

        private static string TrimLines(string text)
        {
            // 줄 앞 공백 4칸 이상은 코드 블록이 되므로 각 줄의 앞 공백을 제거한다. (뒤쪽 "  " 줄바꿈 표시는 유지)
            var lines = text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
                lines[i] = lines[i].TrimStart(' ');
            return string.Join("\n", lines).Trim(' ', '\n');
        }
    }
}
