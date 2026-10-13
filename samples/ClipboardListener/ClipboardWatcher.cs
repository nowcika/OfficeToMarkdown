using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ClipboardListener
{
    /// <summary>
    /// 화면에 보이지 않는 메시지 전용 창으로 WM_CLIPBOARDUPDATE 알림을 받는다.
    /// </summary>
    internal sealed class ClipboardWatcher : NativeWindow, IDisposable
    {
        private const int WM_CLIPBOARDUPDATE = 0x031D;
        private static readonly IntPtr HWND_MESSAGE = new IntPtr(-3);

        private readonly Action _onUpdate;

        public ClipboardWatcher(Action onUpdate)
        {
            _onUpdate = onUpdate;
            CreateHandle(new CreateParams { Parent = HWND_MESSAGE });
            if (!AddClipboardFormatListener(Handle))
                throw new InvalidOperationException("클립보드 감시를 시작하지 못했습니다. (오류 " + Marshal.GetLastWin32Error() + ")");
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_CLIPBOARDUPDATE)
                _onUpdate();
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
