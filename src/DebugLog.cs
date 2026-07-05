using System;
using System.IO;
using System.Text;

namespace SharePointShortcutMaker
{
    internal static partial class Program
    {
        private static string appDebugLogPath;

        private static string GetAppDebugLogPath()
        {
            if (!TeamsDebugLogEnabled)
            {
                return string.Empty;
            }

            if (appDebugLogPath != null)
            {
                return appDebugLogPath;
            }

            try
            {
                string baseDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "M365LinkShortcut",
                    "Logs");
                Directory.CreateDirectory(baseDir);
                appDebugLogPath = Path.Combine(baseDir, "app-debug-" + DateTime.Now.ToString("yyyyMMdd-HHmmss-fff") + ".log");
            }
            catch
            {
                appDebugLogPath = string.Empty;
            }

            return appDebugLogPath;
        }

        internal static void WriteAppDebugLog(string category, string message)
        {
            string path = GetAppDebugLogPath();
            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }

            try
            {
                File.AppendAllText(
                    path,
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " [" + category + "] " + message + Environment.NewLine,
                    Encoding.UTF8);
            }
            catch
            {
            }
        }

        internal static void WriteAppDebugLog(string category, Exception ex)
        {
            if (ex == null)
            {
                return;
            }

            WriteAppDebugLog(category, ex.GetType().FullName + ": " + ex.Message);
        }
    }
}
