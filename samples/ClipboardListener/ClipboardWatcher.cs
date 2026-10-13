using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ClipboardListener
{
    /// <summary>
    /// 화면에 보이지 않는 창으로 WM_CLIPBOARDUPDATE 알림을 받는다.
    /// 설치 관리자(MSI)가 업그레이드·제거 전에 보내는 WM_CLOSE 를 받을 수 있도록 메시지 전용 창이 아닌 숨긴 최상위 창으로 만든다.
    /// </summary>
    internal sealed class ClipboardWatcher : NativeWindow, IDisposable
    {
        private const int WM_CLOSE = 0x0010;
        private const int WM_CLIPBOARDUPDATE = 0x031D;

        private readonly Action _onUpdate;
        private readonly Action _onClose;

        public ClipboardWatcher(Action onUpdate, Action onClose)
        {
            _onUpdate = onUpdate;
            _onClose = onClose;
            CreateHandle(new CreateParams { Caption = "OfficeToMarkdown.ClipboardListener" });
            if (!AddClipboardFormatListener(Handle))
                throw new InvalidOperationException("클립보드 감시를 시작하지 못했습니다. (오류 " + Marshal.GetLastWin32Error() + ")");
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_CLIPBOARDUPDATE)
                _onUpdate();
            if (m.Msg == WM_CLOSE)
            {
                _onClose(); // 창만 닫히고 프로세스가 남지 않도록 앱 전체를 끝낸다.
                return;
            }
            base.WndProc(ref m);
        }

        public void Dispose()
        {
            if (Handle != IntPtr.Zero)
            {
                RemoveClipboardFormatListener(Handle);
                DestroyHandle();
            }
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool AddClipboardFormatListener(IntPtr hwnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RemoveClipboardFormatListener(IntPtr hwnd);
    }
}
