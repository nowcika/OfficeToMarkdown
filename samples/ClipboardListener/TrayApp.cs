using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using OfficeToMarkdown.Core;

namespace ClipboardListener
{
    /// <summary>트레이 아이콘과 클립보드 감시를 묶은 앱. 받은 내용은 저장 폴더(<see cref="Settings"/>)에 HTML 로 저장한다.</summary>
    internal sealed class TrayApp : ApplicationContext
    {
        public const string Title = "OfficeToMarkdown 수신기";

        private static readonly Encoding Utf8NoBom = new UTF8Encoding(false);

        private readonly NotifyIcon _tray;
        private readonly ClipboardWatcher _watcher;
        private string _lastId;
        private string _lastSavedPath;
        private string _outputFolder = Settings.LoadOutputFolder();

        public TrayApp()
        {
            var menu = new ContextMenuStrip();
            menu.Items.Add("저장 폴더 열기", null, (s, e) => OpenFolder());
            menu.Items.Add("저장 폴더 변경...", null, (s, e) => ChangeFolder());
            menu.Items.Add("마지막 파일 열기", null, (s, e) => OpenLastFile());
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("종료", null, (s, e) => ExitThread());

            _tray = new NotifyIcon
            {
                Icon = SystemIcons.Application,
                Text = Title,
                ContextMenuStrip = menu,
                Visible = true,
            };
            _tray.DoubleClick += (s, e) => OpenFolder();
            _tray.BalloonTipClicked += (s, e) => OpenLastFile();

            _watcher = new ClipboardWatcher(OnClipboardUpdate, ExitThread);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _watcher.Dispose();
                _tray.Visible = false; // 종료 뒤 아이콘이 알림 영역에 남지 않게 한다.
                _tray.Dispose();
            }
            base.Dispose(disposing);
        }

        private void OnClipboardUpdate()
        {
            IDataObject data;
            try
            {
                data = Clipboard.GetDataObject();
            }
            catch (ExternalException)
            {
                return; // 다른 프로그램이 아직 클립보드를 잡고 있음
            }
            if (data == null || !data.GetDataPresent(ClipboardMarker.FormatName))
                return; // 추가 기능이 복사한 것이 아님

            var marker = ClipboardMarker.Parse(ReadBytes(data.GetData(ClipboardMarker.FormatName)));
            // 같은 복사로 알림이 여러 번 올 수 있으므로 Id 로 한 번만 처리한다.
            if (marker == null || marker.Id == _lastId)
                return;
            _lastId = marker.Id;

            try
            {
                var fragment = ClipboardHtml.ExtractFragment(data.GetData(DataFormats.Html) as string);
                if (fragment == null)
                {
                    _tray.ShowBalloonTip(5000, Title, "HTML 내용이 없어 저장하지 못했습니다.", ToolTipIcon.Warning);
                    return;
                }
                _lastSavedPath = SaveHtml(_outputFolder, marker.SourceName, fragment);
                _tray.ShowBalloonTip(3000, Title, "저장했습니다: " + Path.GetFileName(_lastSavedPath) + "\n클릭하면 엽니다.",
                                     ToolTipIcon.Info);
            }
            catch (Exception ex)
            {
                _tray.ShowBalloonTip(5000, Title, "저장 중 오류: " + ex.Message, ToolTipIcon.Error);
            }
        }

        /// <returns>저장한 파일 경로. 예: 문서\OfficeToMarkdown\보고서_20261013-153012.html</returns>
        private static string SaveHtml(string outputFolder, string sourceName, string fragment)
        {
            Directory.CreateDirectory(outputFolder);
            var baseName = SafeFileName(Path.GetFileNameWithoutExtension(sourceName));
            var stem = Path.Combine(outputFolder, baseName + "_" + DateTime.Now.ToString("yyyyMMdd-HHmmss"));
            var path = stem + ".html";
            for (int i = 2; File.Exists(path); i++)
                path = stem + " (" + i + ").html";

            var title = string.IsNullOrEmpty(sourceName) ? baseName : sourceName;
            var html = HtmlPage.Build(title, fragment);
            File.WriteAllText(path, html, Utf8NoBom);
            return path;
        }

        private static string SafeFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "문서";
            foreach (var c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return name;
        }

        private static byte[] ReadBytes(object value)
        {
            switch (value)
            {
                case MemoryStream memory:
                    return memory.ToArray();
                case Stream stream:
                    using (var copy = new MemoryStream())
                    {
                        stream.CopyTo(copy);
                        return copy.ToArray();
                    }
                case byte[] bytes:
                    return bytes;
                default:
                    return null;
            }
        }

        private void OpenFolder()
        {
            Directory.CreateDirectory(_outputFolder);
            Shell(_outputFolder);
        }

        private void ChangeFolder()
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "HTML 파일을 저장할 폴더를 선택하세요.";
                dialog.UseDescriptionForTitle = true;
                dialog.ShowNewFolderButton = true;
                if (Directory.Exists(_outputFolder))
                    dialog.InitialDirectory = _outputFolder;
                if (dialog.ShowDialog() != DialogResult.OK)
                    return;
                _outputFolder = dialog.SelectedPath;
            }

            try
            {
                Settings.SaveOutputFolder(_outputFolder);
                _tray.ShowBalloonTip(3000, Title, "저장 폴더: " + _outputFolder, ToolTipIcon.Info);
            }
            catch (Exception ex)
            {
                // 설정 파일을 못 써도 이번 실행 동안은 바뀐 폴더를 쓴다.
                _tray.ShowBalloonTip(5000, Title, "저장 폴더를 바꿨지만 설정을 기록하지 못했습니다: " + ex.Message,
                                     ToolTipIcon.Warning);
            }
        }

        private void OpenLastFile()
        {
            if (_lastSavedPath != null && File.Exists(_lastSavedPath))
                Shell(_lastSavedPath);
        }

        private static void Shell(string path)
        {
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }
    }
}
