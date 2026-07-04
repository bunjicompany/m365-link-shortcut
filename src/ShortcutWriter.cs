using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Script.Serialization;
using System.Windows.Automation;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.VisualBasic;
using Microsoft.Win32;

namespace SharePointShortcutMaker
{
    internal static partial class Program
    {
        private static ShortcutIcon GetShortcutIcon(string url, string baseName)
        {
            if (IsTeamsChatLink(url))
            {
                string chatIconPath = GetEmbeddedIconPath("M365TeamsChat.ico");
                if (!string.IsNullOrWhiteSpace(chatIconPath))
                {
                    return new ShortcutIcon(chatIconPath, 0);
                }

                return GetSystemIcon("imageres.dll", 102);
            }

            if (IsTeamsMeetingLink(url))
            {
                string meetingIconPath = GetEmbeddedIconPath("M365TeamsMeeting.ico");
                if (!string.IsNullOrWhiteSpace(meetingIconPath))
                {
                    return new ShortcutIcon(meetingIconPath, 0);
                }

                return GetSystemIcon("imageres.dll", 102);
            }

            if (IsTeamsLink(url))
            {
                ShortcutIcon teamsIcon = GetProtocolIcon("msteams");
                if (teamsIcon != null)
                {
                    return teamsIcon;
                }

                return GetSystemIcon("shell32.dll", 14);
            }

            if (IsFolderSharePointUrl(url))
            {
                return GetSystemIcon("shell32.dll", 3);
            }

            string extension = GetExtensionFromNameOrUrl(baseName, ExtractSharePointUrl(url));
            if (string.IsNullOrWhiteSpace(extension))
            {
                extension = GetExtensionFromOfficeOrSharePointKind(url);
            }

            if (!string.IsNullOrWhiteSpace(extension))
            {
                ShortcutIcon preferredIcon = GetPreferredIconForExtension(extension);
                if (preferredIcon != null)
                {
                    return preferredIcon;
                }

                ShortcutIcon associatedIcon = GetAssociatedIcon(extension);
                if (associatedIcon != null)
                {
                    return associatedIcon;
                }
            }

            return GetSystemIcon("shell32.dll", 13);
        }

        private static ShortcutIcon GetSystemIcon(string fileName, int index)
        {
            string windowsDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            string iconPath = Path.Combine(windowsDirectory, "System32", fileName);
            return new ShortcutIcon(iconPath, index);
        }

        private static string GetEmbeddedIconPath(string iconFileName)
        {
            if (string.IsNullOrWhiteSpace(embeddedFolderPath) || string.IsNullOrWhiteSpace(iconFileName))
            {
                return string.Empty;
            }

            string iconPath = Path.Combine(embeddedFolderPath, iconFileName);
            return File.Exists(iconPath) ? iconPath : string.Empty;
        }

        private static ShortcutIcon GetPreferredIconForExtension(string extension)
        {
            if (string.IsNullOrWhiteSpace(extension))
            {
                return null;
            }

            if (!extension.StartsWith(".", StringComparison.Ordinal))
            {
                extension = "." + extension;
            }

            if (IsOneOfExtensions(extension, ".txt", ".log", ".md", ".ini", ".cfg", ".json", ".xml", ".yaml", ".yml", ".bat", ".cmd", ".ps1", ".psm1", ".psd1", ".vbs", ".js"))
            {
                ShortcutIcon textIcon = GetIconFromRegistryKey(@"SystemFileAssociations\text\DefaultIcon");
                return textIcon ?? GetSystemIcon("notepad.exe", 0);
            }

            if (IsOneOfExtensions(extension, ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff", ".webp", ".heic"))
            {
                ShortcutIcon imageIcon = GetIconFromRegistryKey(@"SystemFileAssociations\image\DefaultIcon");
                return imageIcon ?? GetSystemIcon("imageres.dll", 67);
            }

            return null;
        }

        internal static string GetExtensionFromNameOrUrl(string baseName, string sharePointUrl)
        {
            string extension = Path.GetExtension(baseName);
            if (!string.IsNullOrWhiteSpace(extension))
            {
                return extension;
            }

            try
            {
                Uri uri = new Uri(sharePointUrl);
                extension = Path.GetExtension(GetLastPathSegment(uri.AbsolutePath));
                return string.IsNullOrWhiteSpace(extension) ? string.Empty : extension;
            }
            catch
            {
                return string.Empty;
            }
        }

        internal static string GetExtensionFromOfficeOrSharePointKind(string url)
        {
            if (url.StartsWith("ms-excel:", StringComparison.OrdinalIgnoreCase))
            {
                return ".xlsx";
            }

            if (url.StartsWith("ms-word:", StringComparison.OrdinalIgnoreCase))
            {
                return ".docx";
            }

            if (url.StartsWith("ms-powerpoint:", StringComparison.OrdinalIgnoreCase))
            {
                return ".pptx";
            }

            if (url.StartsWith("ms-visio:", StringComparison.OrdinalIgnoreCase))
            {
                return ".vsdx";
            }

            string sharePointUrl = ExtractSharePointUrl(url);
            if (sharePointUrl.IndexOf("/:i:/", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return ".png";
            }

            if (sharePointUrl.IndexOf("/:t:/", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return ".txt";
            }

            if (sharePointUrl.IndexOf("/:x:/", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return ".xlsx";
            }

            if (sharePointUrl.IndexOf("/:w:/", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return ".docx";
            }

            if (sharePointUrl.IndexOf("/:p:/", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return ".pptx";
            }

            return string.Empty;
        }

        private static bool IsExcelExtension(string extension)
        {
            return IsOneOfExtensions(extension, ".xlsx", ".xlsm", ".xls", ".xlsb", ".xltx", ".xltm", ".csv");
        }

        private static bool IsWordExtension(string extension)
        {
            return IsOneOfExtensions(extension, ".docx", ".docm", ".doc", ".dotx", ".dotm", ".rtf");
        }

        private static bool IsPowerPointExtension(string extension)
        {
            return IsOneOfExtensions(extension, ".pptx", ".pptm", ".ppt", ".potx", ".potm", ".ppsx", ".ppsm");
        }

        private static bool IsVisioExtension(string extension)
        {
            return IsOneOfExtensions(extension, ".vsdx", ".vsdm", ".vsd", ".vssx", ".vstx");
        }

        private static bool IsOneOfExtensions(string extension, params string[] expected)
        {
            if (string.IsNullOrWhiteSpace(extension))
            {
                return false;
            }

            foreach (string item in expected)
            {
                if (string.Equals(extension, item, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        internal static bool IsFolderSharePointUrl(string url)
        {
            string sharePointUrl = ExtractSharePointUrl(url);
            if (sharePointUrl.IndexOf("/:f:/", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            if (sharePointUrl.IndexOf("/folder/", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            try
            {
                Uri uri = new Uri(sharePointUrl);
                foreach (string key in new[] { "RootFolder", "rootFolder", "folder", "id" })
                {
                    string value = GetQueryValue(uri, key);
                    if (string.IsNullOrWhiteSpace(value))
                    {
                        continue;
                    }

                    string leaf = GetLastPathSegment(value);
                    if (!string.IsNullOrWhiteSpace(leaf) && string.IsNullOrWhiteSpace(Path.GetExtension(leaf)))
                    {
                        return true;
                    }
                }
            }
            catch
            {
            }

            return false;
        }

        private static ShortcutIcon GetAssociatedIcon(string extension)
        {
            try
            {
                if (!extension.StartsWith(".", StringComparison.Ordinal))
                {
                    extension = "." + extension;
                }

                string className = string.Empty;
                using (RegistryKey extensionKey = Registry.ClassesRoot.OpenSubKey(extension))
                {
                    if (extensionKey != null)
                    {
                        className = Convert.ToString(extensionKey.GetValue(string.Empty));
                    }
                }

                if (string.IsNullOrWhiteSpace(className))
                {
                    return null;
                }

                using (RegistryKey iconKey = Registry.ClassesRoot.OpenSubKey(className + @"\DefaultIcon"))
                {
                    if (iconKey == null)
                    {
                        return null;
                    }

                    string rawIcon = Convert.ToString(iconKey.GetValue(string.Empty));
                    return ParseIconLocation(rawIcon);
                }
            }
            catch
            {
                return null;
            }
        }

        private static ShortcutIcon GetProtocolIcon(string scheme)
        {
            if (string.IsNullOrWhiteSpace(scheme))
            {
                return null;
            }

            try
            {
                using (RegistryKey iconKey = Registry.ClassesRoot.OpenSubKey(scheme + @"\DefaultIcon"))
                {
                    if (iconKey == null)
                    {
                        return null;
                    }

                    string rawIcon = Convert.ToString(iconKey.GetValue(string.Empty));
                    return ParseIconLocation(rawIcon);
                }
            }
            catch
            {
                return null;
            }
        }

        private static ShortcutIcon GetIconFromRegistryKey(string keyPath)
        {
            try
            {
                using (RegistryKey iconKey = Registry.ClassesRoot.OpenSubKey(keyPath))
                {
                    if (iconKey == null)
                    {
                        return null;
                    }

                    string rawIcon = Convert.ToString(iconKey.GetValue(string.Empty));
                    return ParseIconLocation(rawIcon);
                }
            }
            catch
            {
                return null;
            }
        }

        private static ShortcutIcon ParseIconLocation(string rawIcon)
        {
            if (string.IsNullOrWhiteSpace(rawIcon))
            {
                return null;
            }

            string value = Environment.ExpandEnvironmentVariables(rawIcon.Trim());
            string file = value;
            int index = 0;

            if (value.StartsWith("\"", StringComparison.Ordinal))
            {
                int closingQuote = value.IndexOf('"', 1);
                if (closingQuote > 1)
                {
                    file = value.Substring(1, closingQuote - 1);
                    string remainder = value.Substring(closingQuote + 1).Trim();
                    if (remainder.StartsWith(",", StringComparison.Ordinal))
                    {
                        int.TryParse(remainder.Substring(1).Trim(), out index);
                    }
                }
            }
            else
            {
                int comma = value.LastIndexOf(',');
                if (comma > 0)
                {
                    file = value.Substring(0, comma).Trim();
                    int.TryParse(value.Substring(comma + 1).Trim(), out index);
                }
            }

            if (string.IsNullOrWhiteSpace(file))
            {
                return null;
            }

            return new ShortcutIcon(file, index);
        }

        internal static string GetSafeFileName(string name)
        {
            string safeName = name;
            foreach (char invalid in Path.GetInvalidFileNameChars())
            {
                safeName = safeName.Replace(invalid, '_');
            }

            safeName = safeName.Trim();
            return string.IsNullOrWhiteSpace(safeName)
                ? "M365\u30ea\u30f3\u30af"
                : safeName;
        }

        private static void CreateWindowsShortcut(string shortcutPath, string url, ShortcutIcon icon)
        {
            Type shellType = Type.GetTypeFromProgID("WScript.Shell");
            if (shellType == null)
            {
                throw new InvalidOperationException("Windows shortcut service was not found.");
            }

            object shell = null;
            object shortcut = null;
            try
            {
                shell = Activator.CreateInstance(shellType);
                shortcut = shellType.InvokeMember(
                    "CreateShortcut",
                    BindingFlags.InvokeMethod,
                    null,
                    shell,
                    new object[] { shortcutPath });

                Type shortcutType = shortcut.GetType();
                shortcutType.InvokeMember("TargetPath", BindingFlags.SetProperty, null, shortcut, new object[] { Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "explorer.exe") });
                shortcutType.InvokeMember("Arguments", BindingFlags.SetProperty, null, shortcut, new object[] { "\"" + url.Replace("\"", "%22") + "\"" });
                shortcutType.InvokeMember("IconLocation", BindingFlags.SetProperty, null, shortcut, new object[] { icon.File + "," + icon.Index.ToString() });
                shortcutType.InvokeMember("Save", BindingFlags.InvokeMethod, null, shortcut, null);
            }
            finally
            {
                if (shortcut != null && Marshal.IsComObject(shortcut))
                {
                    Marshal.ReleaseComObject(shortcut);
                }

                if (shell != null && Marshal.IsComObject(shell))
                {
                    Marshal.ReleaseComObject(shell);
                }
            }
        }

        private static string GetUniquePath(string directory, string baseName, string extension)
        {
            string candidate = Path.Combine(directory, baseName + extension);
            if (!File.Exists(candidate))
            {
                return candidate;
            }

            for (int i = 2; i < 1000; i++)
            {
                candidate = Path.Combine(directory, string.Format("{0} ({1}){2}", baseName, i, extension));
                if (!File.Exists(candidate))
                {
                    return candidate;
                }
            }

            return Path.Combine(directory, baseName + "-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + extension);
        }
    }
}
