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
        private static ClipboardLink GetClipboardLink()
        {
            string text = Clipboard.ContainsText() ? Clipboard.GetText(TextDataFormat.UnicodeText).Trim() : string.Empty;
            string html = Clipboard.ContainsText(TextDataFormat.Html) ? Clipboard.GetText(TextDataFormat.Html) : string.Empty;
            if (string.IsNullOrWhiteSpace(text))
            {
                throw new InvalidOperationException("\u30af\u30ea\u30c3\u30d7\u30dc\u30fc\u30c9\u306bM365\u30ea\u30f3\u30af\u304c\u3042\u308a\u307e\u305b\u3093\u3002SharePoint\u3001Teams\u3001OneDrive\u3067\u30ea\u30f3\u30af\u3092\u30b3\u30d4\u30fc\u3057\u3066\u304b\u3089\u5b9f\u884c\u3057\u3066\u304f\u3060\u3055\u3044\u3002");
            }

            Match officeMatch = Regex.Match(text, "(?:ms-excel|ms-word|ms-powerpoint|ms-visio|ms-access|onenote):[^\\s\"'<>()]+", RegexOptions.IgnoreCase);
            if (officeMatch.Success)
            {
                string officeUrl = CleanUrlCandidate(officeMatch.Value);
                if (IsSharePointUrl(ExtractSharePointUrl(officeUrl)))
                {
                    return new ClipboardLink(officeUrl);
                }
            }

            Match teamsProtocolMatch = Regex.Match(text, "(?:msteams|ms-teams):[^\\s\"'<>()]+", RegexOptions.IgnoreCase);
            if (teamsProtocolMatch.Success)
            {
                string teamsProtocolUrl = CleanUrlCandidate(teamsProtocolMatch.Value);
                if (IsTeamsLink(teamsProtocolUrl))
                {
                    return new ClipboardLink(teamsProtocolUrl, GetLinkTitleCandidate(text, html, teamsProtocolUrl));
                }
            }

            Match match = Regex.Match(text, "https?://[^\\s\"'<>()]+", RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                throw new InvalidOperationException("\u30af\u30ea\u30c3\u30d7\u30dc\u30fc\u30c9\u304b\u3089URL\u3092\u898b\u3064\u3051\u3089\u308c\u307e\u305b\u3093\u3067\u3057\u305f\u3002M365\u30ea\u30f3\u30af\u3092\u30b3\u30d4\u30fc\u3057\u3066\u304b\u3089\u5b9f\u884c\u3057\u3066\u304f\u3060\u3055\u3044\u3002");
            }

            string url = CleanUrlCandidate(match.Value);
            if (!IsSharePointUrl(url) && !IsTeamsLink(url))
            {
                throw new InvalidOperationException("M365\u30ea\u30f3\u30af\u3067\u306f\u306a\u3044\u3088\u3046\u3067\u3059\u3002\r\n\r\n\u691c\u51fa\u3057\u305fURL:\r\n" + url);
            }

            return new ClipboardLink(url, GetLinkTitleCandidate(text, html, url));
        }

        private static bool LooksLikeOpaqueSharePointId(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            string trimmed = value.Trim();
            return trimmed.IndexOf('.') < 0
                && Regex.IsMatch(trimmed, "^[A-Za-z0-9_-]{18,}$");
        }

        private static bool IsSharePointUrl(string url)
        {
            return Regex.IsMatch(url, "^https://[^/]+\\.sharepoint\\.com/", RegexOptions.IgnoreCase);
        }

        internal static bool IsTeamsLink(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                return false;
            }

            string cleaned = CleanUrlCandidate(url);
            if (Regex.IsMatch(cleaned, "^(?:msteams|ms-teams):", RegexOptions.IgnoreCase))
            {
                return true;
            }

            return Regex.IsMatch(cleaned, "^https://teams\\.(?:microsoft|live)\\.com/", RegexOptions.IgnoreCase)
                || Regex.IsMatch(cleaned, "^https://[^/]+\\.teams\\.microsoft\\.com/", RegexOptions.IgnoreCase);
        }

        internal static bool IsTeamsChatLink(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                return false;
            }

            string cleaned = CleanUrlCandidate(url);
            return Regex.IsMatch(cleaned, "(?:/l/message/|/l/chat/|/chat/|/l/channel/|/channel/|/l/team/|[?&]chatId=)", RegexOptions.IgnoreCase);
        }

        internal static bool IsTeamsTeamLink(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                return false;
            }

            return Regex.IsMatch(CleanUrlCandidate(url), "/l/team/", RegexOptions.IgnoreCase);
        }

        internal static bool IsTeamsMeetingLink(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                return false;
            }

            string cleaned = CleanUrlCandidate(url);
            return Regex.IsMatch(cleaned, "(?:/meetup-join/|/meet/|[?&]meetingId=)", RegexOptions.IgnoreCase);
        }

        private static bool IsTeamsChannelChatLink(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                return false;
            }

            string cleaned = CleanUrlCandidate(url);
            return Regex.IsMatch(cleaned, "(?:/l/channel/|/channel/|@thread\\.tacv2|[?&]groupId=)", RegexOptions.IgnoreCase);
        }

        private static string GetQueryParameter(string url, string name)
        {
            if (string.IsNullOrWhiteSpace(url) || string.IsNullOrWhiteSpace(name))
            {
                return string.Empty;
            }

            Match match = Regex.Match(url, "[?&]" + Regex.Escape(name) + "=(?<value>[^&#]*)", RegexOptions.IgnoreCase);
            return match.Success ? DecodeUrlComponent(match.Groups["value"].Value) : string.Empty;
        }

        private static string GetTeamsChannelTeamNameFromUrlOrCache(string url)
        {
            string groupId = GetQueryParameter(url, "groupId");
            string teamName = GetQueryParameter(url, "teamName");
            if (!string.IsNullOrWhiteSpace(groupId) && !string.IsNullOrWhiteSpace(teamName))
            {
                SaveTeamsChannelTeamNameCache(groupId, teamName);
                return teamName.Trim();
            }

            return ReadTeamsChannelTeamNameCache(groupId);
        }

        internal static string GetTeamsChannelNameFromUrl(string url)
        {
            string channelName = GetQueryParameter(url, "channelName");
            if (!string.IsNullOrWhiteSpace(channelName))
            {
                return channelName.Trim();
            }

            Match match = Regex.Match(CleanUrlCandidate(url), "/l/channel/[^/]+/(?<name>[^?/#]+)", RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                match = Regex.Match(CleanUrlCandidate(url), "/channel/[^/]+/(?<name>[^?/#]+)", RegexOptions.IgnoreCase);
            }

            return match.Success ? DecodeUrlComponent(match.Groups["name"].Value).Trim() : string.Empty;
        }

        private static string GetTeamsChannelTeamCachePath()
        {
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            if (string.IsNullOrWhiteSpace(localAppData))
            {
                localAppData = Path.GetTempPath();
            }

            return Path.Combine(localAppData, "M365LinkShortcut", "TeamsChannelTeamCache.tsv");
        }

        private static string ReadTeamsChannelTeamNameCache(string groupId)
        {
            if (string.IsNullOrWhiteSpace(groupId))
            {
                return string.Empty;
            }

            try
            {
                string path = GetTeamsChannelTeamCachePath();
                if (!File.Exists(path))
                {
                    return string.Empty;
                }

                foreach (string line in File.ReadAllLines(path, Encoding.UTF8))
                {
                    string[] parts = line.Split('\t');
                    if (parts.Length >= 2 && string.Equals(parts[0], groupId, StringComparison.OrdinalIgnoreCase))
                    {
                        return parts[1].Trim();
                    }
                }
            }
            catch (Exception ex)
            {
                WriteAppDebugLog("ReadTeamsChannelTeamNameCache", ex);
            }

            return string.Empty;
        }

        private static void SaveTeamsChannelTeamNameCache(string groupId, string teamName)
        {
            if (string.IsNullOrWhiteSpace(groupId) || string.IsNullOrWhiteSpace(teamName))
            {
                return;
            }

            try
            {
                string path = GetTeamsChannelTeamCachePath();
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                Dictionary<string, string> cache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                if (File.Exists(path))
                {
                    foreach (string line in File.ReadAllLines(path, Encoding.UTF8))
                    {
                        string[] parts = line.Split('\t');
                        if (parts.Length >= 2 && !string.IsNullOrWhiteSpace(parts[0]) && !string.IsNullOrWhiteSpace(parts[1]))
                        {
                            cache[parts[0]] = parts[1];
                        }
                    }
                }

                cache[groupId.Trim()] = teamName.Trim();
                List<string> lines = new List<string>();
                foreach (KeyValuePair<string, string> entry in cache)
                {
                    lines.Add(entry.Key + "\t" + entry.Value);
                }

                File.WriteAllLines(path, lines.ToArray(), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                WriteAppDebugLog("SaveTeamsChannelTeamNameCache", ex);
            }
        }

        private static string GetLinkTitleCandidate(string text, string html, string url)
        {
            string title = GetMarkdownLinkTitleCandidate(text, url);
            if (IsUsableTitleCandidate(title))
            {
                return title;
            }

            title = GetHtmlLinkTitleCandidate(html, url);
            if (IsUsableTitleCandidate(title))
            {
                return title;
            }

            title = GetPlainTextTitleCandidate(text, url);
            return IsUsableTitleCandidate(title) ? title : string.Empty;
        }

        private static string GetMarkdownLinkTitleCandidate(string text, string url)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            foreach (Match match in Regex.Matches(text, "\\[(?<title>[^\\]]{1,200})\\]\\((?<url>https?://[^\\)]+)\\)", RegexOptions.IgnoreCase))
            {
                if (IsSameLink(match.Groups["url"].Value, url))
                {
                    return CleanTitleCandidate(match.Groups["title"].Value);
                }
            }

            return string.Empty;
        }

        private static string GetHtmlLinkTitleCandidate(string html, string url)
        {
            if (string.IsNullOrWhiteSpace(html))
            {
                return string.Empty;
            }

            foreach (Match match in Regex.Matches(html, "<a\\b[^>]*href\\s*=\\s*[\"'](?<url>[^\"']+)[\"'][^>]*>(?<title>.*?)</a>", RegexOptions.IgnoreCase | RegexOptions.Singleline))
            {
                if (IsSameLink(match.Groups["url"].Value, url))
                {
                    return CleanTitleCandidate(match.Groups["title"].Value);
                }
            }

            return string.Empty;
        }

        private static string GetPlainTextTitleCandidate(string text, string url)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            string[] lines = text.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string rawLine in lines)
            {
                string line = CleanTitleCandidate(rawLine);
                if (line.IndexOf("http://", StringComparison.OrdinalIgnoreCase) >= 0
                    || line.IndexOf("https://", StringComparison.OrdinalIgnoreCase) >= 0
                    || IsSameLink(line, url))
                {
                    continue;
                }

                if (IsUsableTitleCandidate(line))
                {
                    return line;
                }
            }

            return string.Empty;
        }

        private static bool IsSameLink(string candidateUrl, string url)
        {
            string left = CleanUrlCandidate(candidateUrl);
            string right = CleanUrlCandidate(url);
            if (string.Equals(left, right, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return IsTeamsLink(left) && IsTeamsLink(right);
        }

        private static string CleanTitleCandidate(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            string decoded = HttpUtility.HtmlDecode(value);
            decoded = Regex.Replace(decoded, "<[^>]+>", " ");
            decoded = Regex.Replace(decoded, "\\s+", " ").Trim();
            return RepairUtf8AsShiftJisMojibake(decoded);
        }

        internal static bool IsUsableTitleCandidate(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            string candidate = value.Trim();
            if (candidate.Length < 2 || candidate.Length > 120)
            {
                return false;
            }

            if (candidate.IndexOf('�') >= 0 || LooksLikeUtf8AsShiftJisMojibake(candidate))
            {
                return false;
            }

            if (candidate.IndexOf("teams.microsoft.com", StringComparison.OrdinalIgnoreCase) >= 0
                || candidate.IndexOf("teams.live.com", StringComparison.OrdinalIgnoreCase) >= 0
                || candidate.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
                || candidate.StartsWith("https://", StringComparison.OrdinalIgnoreCase)
                || candidate.StartsWith("msteams:", StringComparison.OrdinalIgnoreCase)
                || candidate.StartsWith("ms-teams:", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return !Regex.IsMatch(candidate, "^(?:Microsoft Teams|Teams|Join|Open|会議に参加|参加)$", RegexOptions.IgnoreCase);
        }

        private static string ExtractSharePointUrl(string url)
        {
            Match match = Regex.Match(url, "https?://[^\\s\"'<>()]+", RegexOptions.IgnoreCase);
            return match.Success ? CleanUrlCandidate(match.Value) : CleanUrlCandidate(url);
        }

        internal static string CleanUrlCandidate(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                return string.Empty;
            }

            string cleaned = HttpUtility.HtmlDecode(url).Trim();
            cleaned = cleaned.Trim('<', '>', '"', '\'');

            while (cleaned.Length > 0 && "].,;".IndexOf(cleaned[cleaned.Length - 1]) >= 0)
            {
                cleaned = cleaned.Substring(0, cleaned.Length - 1).TrimEnd();
            }

            return cleaned;
        }

        private static string GetLaunchUrl(string url)
        {
            if (IsTeamsLink(url))
            {
                return GetTeamsLaunchUrl(url);
            }

            if (IsOfficeLaunchUrl(url))
            {
                return url;
            }

            string sharePointUrl = ExtractSharePointUrl(url);
            string scheme = GetOfficeSchemeFromSharePointUrl(sharePointUrl);
            if (string.IsNullOrWhiteSpace(scheme))
            {
                return url;
            }

            if (!IsProtocolRegistered(scheme))
            {
                return sharePointUrl;
            }

            return scheme + ":ofe|u|" + sharePointUrl;
        }

        private static string GetTeamsLaunchUrl(string url)
        {
            string cleaned = CleanUrlCandidate(url);
            if (Regex.IsMatch(cleaned, "^(?:msteams|ms-teams):", RegexOptions.IgnoreCase))
            {
                return cleaned;
            }

            string teamsUrl = ExtractSharePointUrl(cleaned);
            if (IsProtocolRegistered("msteams"))
            {
                try
                {
                    Uri uri = new Uri(teamsUrl);
                    return "msteams://" + uri.Host + uri.PathAndQuery + uri.Fragment;
                }
                catch (Exception ex)
                {
                    WriteAppDebugLog("GetTeamsLaunchUrl", ex);
                    return teamsUrl;
                }
            }

            return teamsUrl;
        }

        private static bool IsProtocolRegistered(string scheme)
        {
            if (string.IsNullOrWhiteSpace(scheme))
            {
                return false;
            }

            try
            {
                using (RegistryKey key = Registry.ClassesRoot.OpenSubKey(scheme))
                {
                    return key != null && key.GetValue("URL Protocol") != null;
                }
            }
            catch (Exception ex)
            {
                WriteAppDebugLog("IsProtocolRegistered", ex);
                return false;
            }
        }

        private static bool IsOfficeLaunchUrl(string url)
        {
            return Regex.IsMatch(url, "^(?:ms-excel|ms-word|ms-powerpoint|ms-visio|ms-access|onenote):", RegexOptions.IgnoreCase);
        }

        private static string GetOfficeSchemeFromSharePointUrl(string sharePointUrl)
        {
            string extension = GetExtensionFromNameOrUrl(string.Empty, sharePointUrl);
            if (IsExcelExtension(extension) || sharePointUrl.IndexOf("/:x:/", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "ms-excel";
            }

            if (IsWordExtension(extension) || sharePointUrl.IndexOf("/:w:/", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "ms-word";
            }

            if (IsPowerPointExtension(extension) || sharePointUrl.IndexOf("/:p:/", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "ms-powerpoint";
            }

            if (IsVisioExtension(extension))
            {
                return "ms-visio";
            }

            return string.Empty;
        }

        private static string GetShortcutName(string url, bool quiet, string titleCandidate)
        {
            if (IsTeamsLink(url))
            {
                string teamsPrefix = GetTeamsShortcutPrefix(url);
                string teamsFallbackName = GetTeamsShortcutName(url, titleCandidate);
                if (!quiet)
                {
                    TeamsAppNameResult appName = GetTeamsNameFromApp(url);
                    if (!string.IsNullOrWhiteSpace(appName.Name))
                    {
                        return appName.Name;
                    }

                    if (IsTeamsMeetingPrefix(teamsPrefix))
                    {
                        string teamsManualFallbackName = string.IsNullOrWhiteSpace(appName.SuggestedFallbackName)
                            ? teamsPrefix
                            : appName.SuggestedFallbackName;
                        return GetShortcutNameFromUserInputOrFallback(
                            quiet,
                            teamsManualFallbackName,
                            "\u4f1a\u8b70\u540d\u3092Teams\u306e\u5de6\u4e0a\u30bf\u30a4\u30c8\u30eb\u304b\u3089\u53d6\u5f97\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002");
                    }

                    if (IsTeamsTeamLink(url) && !IsBuiltTeamsChatShortcutTitle(teamsFallbackName))
                    {
                        return GetShortcutNameFromUserInputOrFallback(
                            quiet,
                            teamsFallbackName,
                            "\u30c1\u30fc\u30e0\u306e\u30ea\u30f3\u30af\u304b\u3089\u30c1\u30fc\u30e0\u540d\u3068\u30c1\u30e3\u30cd\u30eb\u540d\u3092\u53d6\u5f97\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002");
                    }
                }

                return teamsFallbackName;
            }

            string fallbackName = "M365\u30ea\u30f3\u30af";
            string fallbackExtension = GetExtensionFromOfficeOrSharePointKind(url);
            if (!IsFolderSharePointUrl(url) && !string.IsNullOrWhiteSpace(fallbackExtension))
            {
                fallbackName += fallbackExtension;
            }

            if (!quiet)
            {
                BrowserNameResult browserResult = GetNameFromBrowser(url);
                if (!string.IsNullOrWhiteSpace(browserResult.Name))
                {
                    return browserResult.Name;
                }

                return GetShortcutNameFromUserInputOrFallback(quiet, fallbackName, browserResult.FailureReason);
            }

            return fallbackName;
        }

        private static string GetTeamsShortcutName(string url, string titleCandidate)
        {
            string cleaned = CleanUrlCandidate(url);
            if (Regex.IsMatch(cleaned, "(?:/meetup-join/|/meet/|[?&]meetingId=)", RegexOptions.IgnoreCase))
            {
                return AppendTeamsTitle("Teams \u4f1a\u8b70", titleCandidate, url);
            }

            if (IsTeamsChatLink(cleaned))
            {
                return AppendTeamsTitle("Teams \u30c1\u30e3\u30c3\u30c8", titleCandidate, url);
            }

            return AppendTeamsTitle("Teams \u30ea\u30f3\u30af", titleCandidate, url);
        }

        internal static bool IsTeamsMeetingChatLink(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                return false;
            }

            return Regex.IsMatch(CleanUrlCandidate(url), "19(?::|%3A)meeting_", RegexOptions.IgnoreCase);
        }

        internal static string GetTeamsChatDisplayPrefix(string url)
        {
            if (IsTeamsChannelChatLink(url))
            {
                return "Teams\u6295\u7a3f";
            }

            if (IsTeamsMeetingChatLink(url))
            {
                return "Teams\u4f1a\u8b70\u30c1\u30e3\u30c3\u30c8";
            }

            return "Teams\u30c1\u30e3\u30c3\u30c8";
        }

        private static string GetTeamsShortcutPrefix(string url)
        {
            string cleaned = CleanUrlCandidate(url);
            if (Regex.IsMatch(cleaned, "(?:/meetup-join/|/meet/|[?&]meetingId=)", RegexOptions.IgnoreCase))
            {
                return "Teams \u4f1a\u8b70";
            }

            if (IsTeamsChatLink(cleaned))
            {
                return "Teams \u30c1\u30e3\u30c3\u30c8";
            }

            return "Teams \u30ea\u30f3\u30af";
        }

        private static string AppendTeamsTitle(string prefix, string titleCandidate, string url)
        {
            string title = CleanTeamsBrowserTitleCandidate(titleCandidate);
            if (!IsUsableTitleCandidate(title))
            {
                return IsTeamsChatPrefix(prefix) ? GetTeamsChatDisplayPrefix(url) : prefix;
            }

            if (IsTeamsMeetingPrefix(prefix))
            {
                title = StripTeamsMeetingPrefix(title);
                return string.IsNullOrWhiteSpace(title)
                    ? "Teams\u4f1a\u8b70"
                    : "Teams\u4f1a\u8b70 \u3010" + title + "\u3011";
            }

            if (IsTeamsChatPrefix(prefix))
            {
                title = StripTeamsChatPrefix(title);
                string displayPrefix = GetTeamsChatDisplayPrefix(url);
                return string.IsNullOrWhiteSpace(title)
                    ? displayPrefix
                    : BuildTeamsWholeChatShortcutTitle(displayPrefix, title);
            }

            if (title.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                return title;
            }

            return prefix + " - " + title;
        }

        private static string GetTeamsNameFromBrowserCandidates(string candidates, string url)
        {
            string prefix = GetTeamsShortcutPrefix(url);
            foreach (string rawLine in candidates.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries))
            {
                string title = CleanTeamsBrowserTitleCandidate(rawLine);
                if (IsUsableTeamsBrowserTitle(title))
                {
                    return AppendTeamsTitle(prefix, title, url);
                }
            }

            return string.Empty;
        }

    }
}
