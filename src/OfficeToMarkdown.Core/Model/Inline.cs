using System.Collections.Generic;

namespace OfficeToMarkdown.Core.Model
{
    public enum InlineKind
    {
        Text,
        LineBreak,
        Image,
    }

    /// <summary>블록 안의 텍스트 조각. 서식이 같은 인접 조각은 렌더링 시 하나로 합쳐진다.</summary>
    public sealed class Inline
    {
        public InlineKind Kind { get; set; } = InlineKind.Text;
        public string Text { get; set; } = "";
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Strike { get; set; }

        /// <summary>하이퍼링크 주소. 없으면 null.</summary>
        public string Link { get; set; }

        /// <summary>Kind 가 Image 일 때 Markdown 파일 기준 상대 경로.</summary>
        public string ImagePath { get; set; }

        public static Inline Plain(string text) => new Inline { Text = text };

        public static Inline Break() => new Inline { Kind = InlineKind.LineBreak };

        public static Inline Image(string path, string alt) =>
            new Inline { Kind = InlineKind.Image, ImagePath = path, Text = alt ?? "" };

        /// <summary>줄바꿈 문자(\r\n, \n, \r, \v)를 LineBreak 로 바꾼 서식 없는 조각 목록을 만든다.</summary>
        public static List<Inline> FromPlainText(string text)
        {
            var result = new List<Inline>();
            if (string.IsNullOrEmpty(text))
                return result;

            var lines = text.Replace("\r\n", "\n").Split('\n', '\r', '\v');
            for (int i = 0; i < lines.Length; i++)
            {
                if (i > 0)
                    result.Add(Break());
                if (lines[i].Length > 0)
                    result.Add(Plain(lines[i]));
            }
            return result;
        }
    }
}
