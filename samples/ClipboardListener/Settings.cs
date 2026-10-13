using System;
using System.IO;
using System.Text;

namespace ClipboardListener
{
    /// <summary>
    /// 사용자 설정. %AppData%\OfficeToMarkdown\ClipboardListener.txt 에 저장 폴더 경로 한 줄만 기록한다.
    /// </summary>
    internal static class Settings
    {
        private static readonly string SettingsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OfficeToMarkdown", "ClipboardListener.txt");

        public static readonly string DefaultOutputFolder =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "OfficeToMarkdown");

        /// <returns>저장된 폴더. 설정이 없거나 읽을 수 없으면 기본 폴더(문서\OfficeToMarkdown).</returns>
        public static string LoadOutputFolder()
        {
            try
            {
                if (File.Exists(SettingsPath))
                {
                    var folder = File.ReadAllText(SettingsPath, Encoding.UTF8).Trim();
                    if (folder.Length > 0 && Path.IsPathRooted(folder))
                        return folder;
                }
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
            return DefaultOutputFolder;
        }

        public static void SaveOutputFolder(string folder)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath));
            File.WriteAllText(SettingsPath, folder, new UTF8Encoding(false));
        }
    }
}
