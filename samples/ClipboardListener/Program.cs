using System;
using System.Threading;
using System.Windows.Forms;

namespace ClipboardListener
{
    /// <summary>
    /// 알림 영역(트레이)에서 클립보드를 감시하다가 OfficeToMarkdown 추가 기능이 복사한 내용이 들어오면
    /// 곧바로 HTML 파일로 저장한다. 종료는 트레이 아이콘 오른쪽 클릭 → 종료.
    /// </summary>
    internal static class Program
    {
        [STAThread] // 클립보드 API 는 STA 스레드에서만 동작한다.
        private static void Main()
        {
            // 두 번 실행하면 같은 복사를 두 번 저장하게 되므로 하나만 띄운다.
            using (var mutex = new Mutex(true, @"Local\OfficeToMarkdown.ClipboardListener", out bool first))
            {
                if (!first)
                {
                    MessageBox.Show("이미 실행 중입니다. 알림 영역(트레이)의 아이콘을 확인하세요.", TrayApp.Title,
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                ApplicationConfiguration.Initialize();
                Application.Run(new TrayApp());
            }
        }
    }
}
