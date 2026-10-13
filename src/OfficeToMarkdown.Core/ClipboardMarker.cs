using System;
using System.Text;

namespace OfficeToMarkdown.Core
{
    /// <summary>
    /// 추가 기능이 복사한 클립보드 내용임을 알리는 표식. 클립보드를 감시하는 프로그램은 이 형식이 있을 때만 반응하면 된다.
    /// 내용은 "키=값" 줄로 된 UTF-8 바이트이며, .NET 런타임 간 직렬화 차이를 피하려고 Stream 으로 주고받는다.
    /// </summary>
    public sealed class ClipboardMarker
    {
        /// <summary>클립보드 형식 이름 (RegisterClipboardFormat 에 쓰이는 이름).</summary>
        public const string FormatName = "OfficeToMarkdown";

        public ClipboardMarker(string id, string sourceName)
        {
            Id = id;
            SourceName = sourceName;
        }

        /// <summary>복사할 때마다 새로 만드는 값. 같은 복사로 알림이 여러 번 와도 한 번만 처리하는 데 쓴다.</summary>
        public string Id { get; }

        /// <summary>원본 파일 이름 (예: 보고서.docx).</summary>
        public string SourceName { get; }

        public static ClipboardMarker Create(string sourceName) =>
            new ClipboardMarker(Guid.NewGuid().ToString("N"), sourceName ?? "");

        public byte[] ToBytes()
        {
            // 파일 이름에는 줄바꿈이 들어갈 수 없지만 혹시 모를 값에 대비해 공백으로 바꾼다.
            var source = SourceName.Replace('\r', ' ').Replace('\n', ' ');
            return Encoding.UTF8.GetBytes("id=" + Id + "\nsource=" + source + "\n");
        }

        /// <returns>형식이 맞지 않으면 null.</returns>
        public static ClipboardMarker Parse(byte[] data)
        {
            if (data == null)
                return null;
            string id = null, source = "";
            foreach (var line in Encoding.UTF8.GetString(data).TrimEnd('\0').Split('\n'))
            {
                var eq = line.IndexOf('=');
                if (eq < 0)
                    continue;
                var key = line.Substring(0, eq);
                var value = line.Substring(eq + 1).TrimEnd('\r');
                if (key == "id")
                    id = value;
                else if (key == "source")
                    source = value;
            }
            return string.IsNullOrEmpty(id) ? null : new ClipboardMarker(id, source);
        }
    }
}
