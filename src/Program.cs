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

[assembly: AssemblyTitle("M365リンクをアイコン化")]
[assembly: AssemblyDescription("Microsoft 365 links shortcut creator")]
[assembly: AssemblyCompany("ぶんじカンパニー")]
[assembly: AssemblyProduct("M365リンクをアイコン化")]
[assembly: AssemblyCopyright("Copyright (c) 2026 ぶんじカンパニー")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

namespace SharePointShortcutMaker
{
    internal static partial class Program
    {
        private const string AppTitle = "M365\u30ea\u30f3\u30af\u3092\u30a2\u30a4\u30b3\u30f3\u5316";
        private const string MenuName = "M365\u30ea\u30f3\u30af\u3092\u30a2\u30a4\u30b3\u30f3\u5316";
        private const string BackgroundShellKey = @"Software\Classes\Directory\Background\shell\M365LinkShortcut";
        private const string DirectoryShellKey = @"Software\Classes\Directory\shell\M365LinkShortcut";
        private const int ShcneAssocChanged = 0x08000000;
        private const int ShcnfIdList = 0x0000;
        private const int SwShowMinNoActive = 7;
        private const int SwRestore = 9;
        private static readonly bool TeamsDebugLogEnabled = false;
        private static bool embeddedDependenciesInitialized;
        private static string embeddedFolderPath = string.Empty;
        private static readonly Dictionary<string, string> EmbeddedAssemblyResources = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "Microsoft.Web.WebView2.Core", "Embedded.Microsoft.Web.WebView2.Core.dll" },
            { "Microsoft.Web.WebView2.WinForms", "Embedded.Microsoft.Web.WebView2.WinForms.dll" }
        };

        [DllImport("shell32.dll")]
        private static extern void SHChangeNotify(int eventId, int flags, IntPtr item1, IntPtr item2);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool SetDllDirectory(string lpPathName);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr SetActiveWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr processId);

        [DllImport("user32.dll")]
        private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool attach);

        [DllImport("user32.dll")]
        private static extern bool BringWindowToTop(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr SetFocus(IntPtr hWnd);

        [STAThread]
        private static int Main(string[] args)
        {
            try
            {
                InitializeEmbeddedDependencies();

                bool quiet = false;
                bool hasTargetPath = false;
                bool toggle = false;
                bool install = false;
                bool uninstall = false;
                string targetPath = Environment.CurrentDirectory;
                foreach (string arg in args)
                {
                    if (string.Equals(arg, "--quiet", StringComparison.OrdinalIgnoreCase))
                    {
                        quiet = true;
                    }
                    else if (string.Equals(arg, "--install", StringComparison.OrdinalIgnoreCase))
                    {
                        install = true;
                    }
                    else if (string.Equals(arg, "--uninstall", StringComparison.OrdinalIgnoreCase))
                    {
                        uninstall = true;
                    }
                    else
                    {
                        hasTargetPath = true;
                        targetPath = arg;
                    }
                }

                toggle = !hasTargetPath && !install && !uninstall;

                if (toggle)
                {
                    if (IsContextMenuRegisteredForThisExe())
                    {
                        UninstallContextMenu();
                        if (!quiet)
                        {
                            MessageBox.Show(
                                "\u53f3\u30af\u30ea\u30c3\u30af\u30e1\u30cb\u30e5\u30fc\u3092\u524a\u9664\u3057\u307e\u3057\u305f\u3002",
                                AppTitle,
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                        }
                    }
                    else
                    {
                        InstallContextMenu();
                        if (!quiet)
                        {
                            MessageBox.Show(
                                "\u53f3\u30af\u30ea\u30c3\u30af\u30e1\u30cb\u30e5\u30fc\u306b\u767b\u9332\u3057\u307e\u3057\u305f\u3002\r\n\r\nM365\u30ea\u30f3\u30af\u3092\u30b3\u30d4\u30fc\u3057\u3066\u304b\u3089\u3001\u30d5\u30a9\u30eb\u30c0\u306e\u4f59\u767d\u3092\u53f3\u30af\u30ea\u30c3\u30af\u3057\u3066\u4f7f\u3063\u3066\u304f\u3060\u3055\u3044\u3002",
                                AppTitle,
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                        }
                    }

                    return 0;
                }

                if (uninstall)
                {
                    UninstallContextMenu();
                    if (!quiet)
                    {
                        MessageBox.Show(
                            "\u53f3\u30af\u30ea\u30c3\u30af\u30e1\u30cb\u30e5\u30fc\u3092\u524a\u9664\u3057\u307e\u3057\u305f\u3002",
                            AppTitle,
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }

                    return 0;
                }

                if (install)
                {
                    InstallContextMenu();
                    if (!quiet)
                    {
                        MessageBox.Show(
                            "\u53f3\u30af\u30ea\u30c3\u30af\u30e1\u30cb\u30e5\u30fc\u306b\u767b\u9332\u3057\u307e\u3057\u305f\u3002\r\n\r\nM365\u30ea\u30f3\u30af\u3092\u30b3\u30d4\u30fc\u3057\u3066\u304b\u3089\u3001\u30d5\u30a9\u30eb\u30c0\u306e\u4f59\u767d\u3092\u53f3\u30af\u30ea\u30c3\u30af\u3057\u3066\u4f7f\u3063\u3066\u304f\u3060\u3055\u3044\u3002",
                            AppTitle,
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }

                    return 0;
                }

                string directory = GetTargetDirectory(targetPath);
                ClipboardLink clipboardLink = GetClipboardLink();
                string url = GetLaunchUrl(clipboardLink.Url);
                bool isTeamsLink = IsTeamsLink(clipboardLink.Url);
                string nameSourceUrl = isTeamsLink ? clipboardLink.Url : ExtractSharePointUrl(url);
                string baseName = GetSafeFileName(GetShortcutName(nameSourceUrl, quiet, clipboardLink.TitleCandidate));
                bool isFolderLink = !isTeamsLink && IsFolderSharePointUrl(url);
                string shortcutPath = GetUniquePath(directory, baseName, isFolderLink ? ".lnk" : ".url");
                ShortcutIcon icon = GetShortcutIcon(url, baseName);

                if (isFolderLink)
                {
                    CreateWindowsShortcut(shortcutPath, ExtractSharePointUrl(url), icon);
                }
                else
                {
                    File.WriteAllLines(shortcutPath, new[]
                    {
                        "[InternetShortcut]",
                        "URL=" + url,
                        "IconFile=" + icon.File,
                        "IconIndex=" + icon.Index.ToString()
                    }, Encoding.GetEncoding(932));
                }

                if (!quiet)
                {
                    IntPtr foregroundWindowBeforeCompletionDialog = GetForegroundWindow();
                    ShowTopMostMessage(
                        "\u30b7\u30e7\u30fc\u30c8\u30ab\u30c3\u30c8\u3092\u4f5c\u6210\u3057\u307e\u3057\u305f\u3002\r\n\r\n" + shortcutPath,
                        AppTitle,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    RestoreForegroundWindow(foregroundWindowBeforeCompletionDialog);
                }

                return 0;
            }
            catch (OperationCanceledException)
            {
                return 0;
            }
            catch (Exception ex)
            {
                if (!HasQuietArg(args))
                {
                    MessageBox.Show(ex.Message, AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                return 1;
            }
        }

        private static void InitializeEmbeddedDependencies()
        {
            if (embeddedDependenciesInitialized)
            {
                return;
            }

            AppDomain.CurrentDomain.AssemblyResolve += ResolveEmbeddedAssembly;
            string nativeFolder = ExtractEmbeddedNativeDependencies();
            if (!string.IsNullOrWhiteSpace(nativeFolder))
            {
                embeddedFolderPath = nativeFolder;
                SetDllDirectory(nativeFolder);
                string path = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
                if (path.IndexOf(nativeFolder, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    Environment.SetEnvironmentVariable("PATH", nativeFolder + Path.PathSeparator + path);
                }
            }

            embeddedDependenciesInitialized = true;
        }

        private static Assembly ResolveEmbeddedAssembly(object sender, ResolveEventArgs args)
        {
            string assemblyName = new AssemblyName(args.Name).Name;
            string resourceName;
            if (!EmbeddedAssemblyResources.TryGetValue(assemblyName, out resourceName))
            {
                return null;
            }

            Assembly currentAssembly = Assembly.GetExecutingAssembly();
            using (Stream stream = currentAssembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    return null;
                }

                byte[] bytes = new byte[stream.Length];
                int offset = 0;
                while (offset < bytes.Length)
                {
                    int read = stream.Read(bytes, offset, bytes.Length - offset);
                    if (read <= 0)
                    {
                        break;
                    }

                    offset += read;
                }

                return Assembly.Load(bytes);
            }
        }

        private static string ExtractEmbeddedNativeDependencies()
        {
            string baseFolder = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            if (string.IsNullOrWhiteSpace(baseFolder))
            {
                baseFolder = Path.GetTempPath();
            }

            string folder = Path.Combine(baseFolder, "M365LinkShortcut", "Embedded");
            Directory.CreateDirectory(folder);
            ExtractEmbeddedResourceToFile("Embedded.WebView2Loader.dll", Path.Combine(folder, "WebView2Loader.dll"));
            ExtractEmbeddedResourceToFile("Embedded.ChatShortcut.ico", Path.Combine(folder, "M365TeamsChat.ico"), true);
            ExtractEmbeddedResourceToFile("Embedded.MeetingShortcut.ico", Path.Combine(folder, "M365TeamsMeeting.ico"), true);
            return folder;
        }

        private static void ExtractEmbeddedResourceToFile(string resourceName, string destinationPath, bool forceOverwrite = false)
        {
            Assembly currentAssembly = Assembly.GetExecutingAssembly();
            using (Stream stream = currentAssembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    return;
                }

                FileInfo destination = new FileInfo(destinationPath);
                if (!forceOverwrite && destination.Exists && destination.Length == stream.Length)
                {
                    return;
                }

                string tempPath = destinationPath + ".tmp";
                using (FileStream output = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    stream.CopyTo(output);
                }

                if (File.Exists(destinationPath))
                {
                    File.Delete(destinationPath);
                }

                File.Move(tempPath, destinationPath);
            }
        }

        private sealed class ClipboardLink
        {
            public ClipboardLink(string url)
                : this(url, string.Empty)
            {
            }

            public ClipboardLink(string url, string titleCandidate)
            {
                Url = url;
                TitleCandidate = titleCandidate ?? string.Empty;
            }

            public string Url { get; private set; }
            public string TitleCandidate { get; private set; }
        }

        private sealed class ShortcutIcon
        {
            public ShortcutIcon(string file, int index)
            {
                File = file;
                Index = index;
            }

            public string File { get; private set; }
            public int Index { get; private set; }
        }

        private sealed class BrowserNameResult
        {
            public BrowserNameResult(string name, string failureReason)
            {
                Name = name ?? string.Empty;
                FailureReason = failureReason ?? string.Empty;
            }

            public string Name { get; private set; }
            public string FailureReason { get; private set; }
        }

        private sealed class TeamsAppNameResult
        {
            public TeamsAppNameResult(string name, string suggestedFallbackName)
            {
                Name = name ?? string.Empty;
                SuggestedFallbackName = suggestedFallbackName ?? string.Empty;
            }

            public string Name { get; private set; }
            public string SuggestedFallbackName { get; private set; }
        }

        private sealed class TeamsTitleScanResult
        {
            public TeamsTitleScanResult(string title, string suggestedFallbackName)
            {
                Title = title ?? string.Empty;
                SuggestedFallbackName = suggestedFallbackName ?? string.Empty;
            }

            public string Title { get; private set; }
            public string SuggestedFallbackName { get; private set; }
        }

        private sealed class TeamsChatInfo
        {
            public TeamsChatInfo(string targetMessageId, bool isChannelLink, string groupId, string teamName, string chatName)
            {
                TargetMessageId = targetMessageId ?? string.Empty;
                IsChannelLink = isChannelLink;
                GroupId = groupId ?? string.Empty;
                TeamName = teamName ?? string.Empty;
                ChatName = chatName ?? string.Empty;
            }

            public string TargetMessageId { get; private set; }
            public bool IsChannelLink { get; private set; }
            public string GroupId { get; private set; }
            public string TeamName { get; set; }
            public string ChatName { get; set; }
            public string LastTreeTeamCandidate { get; set; }
            public string Speaker { get; set; }
            public string LastSpeakerCandidate { get; set; }
            public string TargetSpeakerCandidate { get; set; }
            public string Timestamp { get; set; }
            public bool HasTargetTimestamp { get; set; }
            public bool HasTargetBody { get; set; }
            public int ScanElementIndex { get; set; }
            public int LastSpeakerCandidateIndex { get; set; }
            public int TargetBodyElementIndex { get; set; }

            public bool HasShortcutName
            {
                get
                {
                    return !string.IsNullOrWhiteSpace(ChatName)
                        && !string.IsNullOrWhiteSpace(Speaker)
                        && !string.IsNullOrWhiteSpace(Timestamp)
                        && (!IsChannelLink || !string.IsNullOrWhiteSpace(TeamName));
                }
            }

            public bool HasWholeChatShortcutName
            {
                get
                {
                    return string.IsNullOrWhiteSpace(TargetMessageId)
                        && !string.IsNullOrWhiteSpace(ChatName)
                        && (!IsChannelLink || !string.IsNullOrWhiteSpace(TeamName));
                }
            }

            public bool HasExactShortcutName
            {
                get
                {
                    return HasShortcutName
                        && (string.IsNullOrWhiteSpace(TargetMessageId) || HasTargetTimestamp);
                }
            }
        }

        private static bool HasQuietArg(string[] args)
        {
            foreach (string arg in args)
            {
                if (string.Equals(arg, "--quiet", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static void InstallContextMenu()
        {
            string exePath = Application.ExecutablePath;
            DeleteKeyIfExists(BackgroundShellKey);
            DeleteKeyIfExists(DirectoryShellKey);
            RegisterMenuEntry(BackgroundShellKey, "\"" + exePath + "\" \"%V\"", exePath);
            RegisterMenuEntry(DirectoryShellKey, "\"" + exePath + "\" \"%1\"", exePath);
            RefreshExplorerShell();
        }

        private static bool IsContextMenuRegisteredForThisExe()
        {
            string exePath = Application.ExecutablePath;
            return IsMenuEntryRegisteredForThisExe(BackgroundShellKey, exePath)
                || IsMenuEntryRegisteredForThisExe(DirectoryShellKey, exePath);
        }

        private static bool IsMenuEntryRegisteredForThisExe(string shellKeyPath, string exePath)
        {
            using (RegistryKey commandKey = Registry.CurrentUser.OpenSubKey(shellKeyPath + @"\command"))
            {
                if (commandKey == null)
                {
                    return false;
                }

                string command = Convert.ToString(commandKey.GetValue(string.Empty));
                return !string.IsNullOrEmpty(command)
                    && command.IndexOf(exePath, StringComparison.OrdinalIgnoreCase) >= 0;
            }
        }

        private static void RegisterMenuEntry(string shellKeyPath, string command, string iconPath)
        {
            using (RegistryKey shellKey = Registry.CurrentUser.CreateSubKey(shellKeyPath))
            {
                if (shellKey == null)
                {
                    throw new InvalidOperationException("\u53f3\u30af\u30ea\u30c3\u30af\u30e1\u30cb\u30e5\u30fc\u3092\u767b\u9332\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002");
                }

                shellKey.SetValue("MUIVerb", MenuName, RegistryValueKind.String);
                shellKey.SetValue("Icon", iconPath, RegistryValueKind.String);
            }

            using (RegistryKey commandKey = Registry.CurrentUser.CreateSubKey(shellKeyPath + @"\command"))
            {
                if (commandKey == null)
                {
                    throw new InvalidOperationException("\u53f3\u30af\u30ea\u30c3\u30af\u30e1\u30cb\u30e5\u30fc\u3092\u767b\u9332\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002");
                }

                commandKey.SetValue(string.Empty, command, RegistryValueKind.String);
            }
        }

        private static void UninstallContextMenu()
        {
            DeleteKeyIfExists(BackgroundShellKey);
            DeleteKeyIfExists(DirectoryShellKey);
            RefreshExplorerShell();
        }

        private static void RefreshExplorerShell()
        {
            try
            {
                SHChangeNotify(ShcneAssocChanged, ShcnfIdList, IntPtr.Zero, IntPtr.Zero);
            }
            catch
            {
            }
        }

        private static void DeleteKeyIfExists(string keyPath)
        {
            try
            {
                Registry.CurrentUser.DeleteSubKeyTree(keyPath);
            }
            catch (ArgumentException)
            {
            }
        }

        private static string GetTargetDirectory(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return Environment.CurrentDirectory;
            }

            string expanded = Environment.ExpandEnvironmentVariables(path.Trim('"'));
            if (!File.Exists(expanded) && !Directory.Exists(expanded))
            {
                throw new InvalidOperationException("\u30b7\u30e7\u30fc\u30c8\u30ab\u30c3\u30c8\u4f5c\u6210\u5148\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093: " + expanded);
            }

            if (Directory.Exists(expanded))
            {
                return Path.GetFullPath(expanded);
            }

            string parent = Path.GetDirectoryName(Path.GetFullPath(expanded));
            if (string.IsNullOrEmpty(parent))
            {
                throw new InvalidOperationException("\u30b7\u30e7\u30fc\u30c8\u30ab\u30c3\u30c8\u4f5c\u6210\u5148\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093: " + expanded);
            }

            return parent;
        }

    }
}
