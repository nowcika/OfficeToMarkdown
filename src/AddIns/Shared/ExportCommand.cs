using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using OfficeToMarkdown.Core;
using OfficeToMarkdown.Core.Model;

namespace OfficeToMarkdown.AddIns
{
    /// <summary>
    /// 세 추가 기능이 공유하는 "Markdown으로 복사" 흐름: 변환 → 클립보드에 복사 → 결과 안내.
    /// 각 VSTO 프로젝트에 "링크로 추가" 해서 사용한다.
    /// </summary>
    internal static class ExportCommand
    {
        private const string Caption = "Markdown으로 복사";

        /// <summary>홈 탭 끝에 "Markdown" 그룹과 버튼을 추가하는 리본 XML.</summary>
        public const string RibbonXml = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon>
    <tabs>
      <tab idMso=""TabHome"">
        <group id=""MarkdownGroup"" label=""Markdown"">
          <button id=""ExportMarkdownButton""
                  label=""MD로 복사""
                  screentip=""Markdown으로 복사""
                  supertip=""현재 열려 있는 파일을 Markdown 으로 변환해 클립보드에 복사합니다. 메모장 등에는 Markdown 텍스트, Word·Notion·메일에는 이미지를 포함한 서식으로 붙여넣어집니다.""
                  size=""large""
                  imageMso=""Copy""
                  onAction=""OnExport"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

        /// <param name="ownerHandle">메시지 상자의 부모가 될 Office 창 핸들.</param>
        /// <param name="sourceName">원본 파일 이름 (예: 보고서.docx). 이미지 경로 이름에 쓰인다.</param>
        /// <param name="convert">이미지 폴더 상대 경로를 받아 변환 결과를 돌려주는 함수.</param>
        public static void Run(IntPtr ownerHandle, string sourceName, Func<string, DocumentContent> convert)
        {
            var owner = new OwnerWindow(ownerHandle);
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                // 파일로 저장하지 않으므로 이 폴더는 실제로 만들어지지 않는다. Markdown 텍스트의 이미지 참조 이름으로만 쓰인다.
                var assetFolder = Path.GetFileNameWithoutExtension(sourceName) + "_files";
                var content = convert(assetFolder);
                var markdown = MarkdownRenderer.Render(content.Blocks);
                if (markdown.Length == 0)
                {
                    Cursor.Current = Cursors.Default;
                    MessageBox.Show(owner, "복사할 내용이 없습니다.", Caption, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                var copied = TryCopyToClipboard(markdown, content, sourceName);
                Cursor.Current = Cursors.Default;

                if (!copied)
                {
                    MessageBox.Show(owner, "클립보드에 복사하지 못했습니다. 다른 프로그램이 클립보드를 사용 중일 수 있습니다.",
                                    Caption, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                var message = "클립보드에 복사했습니다.";
                if (content.Resources.Count > 0)
                    message += "\n\n이미지 " + content.Resources.Count + "개는 Word·Notion·메일 등 서식을 지원하는 곳에 붙여넣을 때 함께 들어갑니다.";
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

        /// <summary>
        /// Markdown 텍스트와, 이미지를 data URI 로 넣은 HTML 을 함께 클립보드에 넣는다.
        /// 붙여넣는 프로그램이 둘 중 지원하는 형식을 고른다. 클립보드 감시 프로그램용 표식(<see cref="ClipboardMarker"/>)도 함께 넣는다.
        /// 클립보드는 다른 프로그램이 잠깐 잡고 있을 수 있어 몇 번 재시도한다.
        /// </summary>
        private static bool TryCopyToClipboard(string markdown, DocumentContent content, string sourceName)
        {
            try
            {
                var html = HtmlRenderer.Render(content.Blocks, path =>
                {
                    byte[] data;
                    return content.Resources.TryGetValue(path, out data) ? HtmlRenderer.DataUri(path, data) : null;
                });
                var dataObject = new DataObject();
                dataObject.SetData(DataFormats.UnicodeText, markdown);
                dataObject.SetData(DataFormats.Html, new MemoryStream(ClipboardHtml.FormatBytes(html)));
                dataObject.SetData(ClipboardMarker.FormatName, new MemoryStream(ClipboardMarker.Create(sourceName).ToBytes()));
                Clipboard.SetDataObject(dataObject, true, 5, 100);
                return true;
            }
            catch (ExternalException)
            {
                return false;
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
