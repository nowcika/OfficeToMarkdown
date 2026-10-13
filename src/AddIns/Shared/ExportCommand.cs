using System;
using System.IO;
using System.Windows.Forms;
using OfficeToMarkdown.Core;
using OfficeToMarkdown.Core.Model;

namespace OfficeToMarkdown.AddIns
{
    /// <summary>
    /// 세 추가 기능이 공유하는 "Markdown으로 저장" 흐름: 저장 위치 선택 → 변환 → 저장 → 결과 안내.
    /// 각 VSTO 프로젝트에 "링크로 추가" 해서 사용한다.
    /// </summary>
    internal static class ExportCommand
    {
        private const string Caption = "Markdown으로 저장";

        /// <summary>홈 탭 끝에 "Markdown" 그룹과 버튼을 추가하는 리본 XML.</summary>
        public const string RibbonXml = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon>
    <tabs>
      <tab idMso=""TabHome"">
        <group id=""MarkdownGroup"" label=""Markdown"">
          <button id=""ExportMarkdownButton""
                  label=""MD로 저장""
                  screentip=""Markdown으로 저장""
                  supertip=""현재 열려 있는 파일을 Markdown(.md) 파일로 변환해 저장합니다.""
                  size=""large""
                  imageMso=""FileSaveAs""
                  onAction=""OnExport"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

        /// <param name="ownerHandle">대화상자의 부모가 될 Office 창 핸들.</param>
        /// <param name="sourceDirectory">원본 파일 폴더 (저장 전 문서·클라우드 문서는 빈 값이나 URL일 수 있음).</param>
        /// <param name="sourceName">원본 파일 이름 (예: 보고서.docx).</param>
        /// <param name="convert">이미지 폴더 상대 경로를 받아 변환 결과를 돌려주는 함수.</param>
        public static void Run(IntPtr ownerHandle, string sourceDirectory, string sourceName,
                               Func<string, DocumentContent> convert)
        {
            var owner = new OwnerWindow(ownerHandle);
            string path;
            using (var dialog = new SaveFileDialog())
            {
                dialog.Title = Caption;
                dialog.Filter = "Markdown (*.md)|*.md|모든 파일 (*.*)|*.*";
                dialog.DefaultExt = "md";
                dialog.AddExtension = true;
                dialog.OverwritePrompt = true;
                dialog.FileName = Path.GetFileNameWithoutExtension(sourceName) + ".md";
                dialog.InitialDirectory = IsLocalDirectory(sourceDirectory)
                    ? sourceDirectory
                    : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                if (dialog.ShowDialog(owner) != DialogResult.OK)
                    return;
                path = dialog.FileName;
            }

            try
            {
                Cursor.Current = Cursors.WaitCursor;
                var assetFolder = MarkdownExporter.AssetFolderFor(path);
                var content = convert(assetFolder);
                MarkdownExporter.Save(content, path);
                Cursor.Current = Cursors.Default;

                var message = "저장했습니다.\n\n" + path;
                if (content.Resources.Count > 0)
                    message += "\n\n이미지 " + content.Resources.Count + "개: " + assetFolder + "\\";
                MessageBox.Show(owner, message, Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Cursor.Current = Cursors.Default;
                MessageBox.Show(owner, "변환 중 오류가 발생했습니다.\n\n" + ex.Message, Caption,
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void ShowNoDocument(IntPtr ownerHandle)
        {
            MessageBox.Show(new OwnerWindow(ownerHandle), "열려 있는 파일이 없습니다.", Caption,
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private static bool IsLocalDirectory(string directory)
        {
            try
            {
                return !string.IsNullOrEmpty(directory) && Directory.Exists(directory);
            }
            catch (Exception)
            {
                return false; // OneDrive/SharePoint URL 등
            }
        }

        private sealed class OwnerWindow : IWin32Window
        {
            public OwnerWindow(IntPtr handle)
            {
                Handle = handle;
            }

            public IntPtr Handle { get; }
        }
    }
}
