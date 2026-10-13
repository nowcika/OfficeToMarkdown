using System.IO;
using System.Text;
using OfficeToMarkdown.Core.Model;

namespace OfficeToMarkdown.Core
{
    public static class MarkdownExporter
    {
        private static readonly Encoding Utf8NoBom = new UTF8Encoding(false);

        /// <summary>
        /// 이미지 등 리소스를 저장할 폴더 이름(Markdown 파일 기준 상대 경로).
        /// 예: C:\docs\보고서.md → "보고서_files"
        /// </summary>
        public static string AssetFolderFor(string markdownPath) =>
            Path.GetFileNameWithoutExtension(markdownPath) + "_files";

        /// <summary>Markdown 파일과 리소스 파일을 저장한다.</summary>
        public static void Save(DocumentContent content, string markdownPath)
        {
            var directory = Path.GetDirectoryName(Path.GetFullPath(markdownPath));
            foreach (var resource in content.Resources)
            {
                var path = Path.Combine(directory, resource.Key.Replace('/', Path.DirectorySeparatorChar));
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                File.WriteAllBytes(path, resource.Value);
            }
            File.WriteAllText(markdownPath, MarkdownRenderer.Render(content.Blocks), Utf8NoBom);
        }
    }
}
