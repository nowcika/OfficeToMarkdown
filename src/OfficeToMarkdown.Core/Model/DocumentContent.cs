using System.Collections.Generic;

namespace OfficeToMarkdown.Core.Model
{
    /// <summary>변환 결과: 블록 목록과, Markdown 파일 옆에 함께 저장할 리소스(이미지 등).</summary>
    public sealed class DocumentContent
    {
        public List<Block> Blocks { get; } = new List<Block>();

        /// <summary>Markdown 파일 기준 상대 경로('/' 구분) → 파일 내용.</summary>
        public Dictionary<string, byte[]> Resources { get; } = new Dictionary<string, byte[]>();
    }
}
