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
        private static TeamsAppNameResult GetTeamsNameFromApp(string url)
        {
            IntPtr foregroundWindowToRestore = GetForegroundWindow();
            try
            {
            string prefix = GetTeamsShortcutPrefix(url);
            string launchUrl = GetTeamsLaunchUrl(url);
            string targetMessageId = GetTeamsMessageId(url);
            string debugLogPath = CreateTeamsDebugLogPath();
            WriteTeamsDebugLog(debugLogPath, "Teams title debug started");
            WriteTeamsDebugLog(debugLogPath, "url=" + url);
            WriteTeamsDebugLog(debugLogPath, "launchUrl=" + launchUrl);
            WriteTeamsDebugLog(debugLogPath, "prefix=" + prefix);
            WriteTeamsDebugLog(debugLogPath, "targetMessageId=" + ToLogValue(targetMessageId));
            Dictionary<int, TeamsWindowInfo> beforeWindows = GetTeamsWindows();
            WriteTeamsDebugLog(debugLogPath, "beforeWindows=" + beforeWindows.Count);
            foreach (TeamsWindowInfo beforeWindow in beforeWindows.Values)
            {
                if (TeamsDebugLogEnabled)
                {
                    WriteTeamsDebugLog(debugLogPath, FormatTeamsWindowSummary("before", beforeWindow));
                }
            }

            TeamsWindowInfo openedWindow = null;
            bool openedWindowMinimized = false;
            bool chatPrefix = IsTeamsChatPrefix(prefix);
            string lastAcceptedTitle = string.Empty;
            string suggestedFallbackName = string.Empty;
            int stableTitleCount = 0;

            try
            {
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = launchUrl;
                startInfo.UseShellExecute = true;
                startInfo.WindowStyle = ProcessWindowStyle.Minimized;
                using (Process startedProcess = Process.Start(startInfo))
                {
                }
                WriteTeamsDebugLog(debugLogPath, "Process.Start succeeded with WindowStyle=Minimized");
            }
            catch (Exception ex)
            {
                WriteTeamsDebugLog(debugLogPath, "Process.Start failed: " + ex.GetType().FullName + ": " + ex.Message);
                return new TeamsAppNameResult(string.Empty, string.Empty);
            }

            DateTime quickMinimizeDeadline = DateTime.Now.AddSeconds(3);
            while (!chatPrefix && DateTime.Now < quickMinimizeDeadline && openedWindow == null)
            {
                System.Threading.Thread.Sleep(50);
                openedWindow = FindNewTeamsWindow(beforeWindows);
                if (openedWindow != null)
                {
                    MinimizeTeamsWindow(openedWindow);
                    openedWindowMinimized = true;
                    WriteTeamsDebugLog(debugLogPath, "quick minimized opened Teams window: handle=" + openedWindow.Handle);
                    break;
                }
            }

            DateTime deadline = DateTime.Now.AddSeconds(5);
            int round = 0;
            while (DateTime.Now < deadline)
            {
                round++;
                System.Threading.Thread.Sleep(250);
                Dictionary<int, TeamsWindowInfo> windows = GetTeamsWindows();
                WriteTeamsDebugLog(debugLogPath, "round=" + round + ", windows=" + windows.Count);
                foreach (TeamsWindowInfo window in windows.Values)
                {
                    bool existingWindow = beforeWindows.ContainsKey(window.Handle);
                    if (existingWindow && !chatPrefix)
                    {
                        continue;
                    }

                    if (!existingWindow)
                    {
                        openedWindow = window;
                        if (!openedWindowMinimized)
                        {
                            MinimizeTeamsWindow(window);
                            openedWindowMinimized = true;
                            WriteTeamsDebugLog(debugLogPath, "minimized opened Teams window: handle=" + window.Handle);
                        }
                    }

                    if (TeamsDebugLogEnabled)
                    {
                        WriteTeamsDebugLog(debugLogPath, FormatTeamsWindowSummary(existingWindow ? "existingAfterLaunch" : "new", window));
                    }
                    TeamsTitleScanResult scanResult = GetTeamsVisibleTitle(window, prefix, debugLogPath, targetMessageId, IsTeamsChannelChatLink(url), url);
                    string title = scanResult.Title;
                    if (string.IsNullOrWhiteSpace(suggestedFallbackName) && !string.IsNullOrWhiteSpace(scanResult.SuggestedFallbackName))
                    {
                        suggestedFallbackName = scanResult.SuggestedFallbackName;
                        WriteTeamsDebugLog(debugLogPath, "suggestedFallbackName=" + ToLogValue(suggestedFallbackName));
                    }

                    if (string.IsNullOrWhiteSpace(title) && !IsTeamsMeetingPrefix(prefix) && !IsTeamsChatPrefix(prefix))
                    {
                        title = CleanTeamsBrowserTitleCandidate(window.Title);
                        WriteTeamsDebugLog(debugLogPath, "fallbackToAutomationWindowTitle=" + ToLogValue(window.Title) + ", cleaned=" + ToLogValue(title));
                    }
                    else if (string.IsNullOrWhiteSpace(title) && IsTeamsChatPrefix(prefix))
                    {
                        WriteTeamsDebugLog(debugLogPath, "chat title not ready. skipped fallback automation title.");
                    }
                    else if (string.IsNullOrWhiteSpace(title))
                    {
                        WriteTeamsDebugLog(debugLogPath, "meeting title not ready. skipped fallback automation title.");
                    }

                    bool builtChatTitle = chatPrefix && IsBuiltTeamsChatShortcutTitle(title);
                    WriteTeamsDebugLog(debugLogPath, "acceptedCandidateBeforeStableCheck=" + ToLogValue(title) + ", usable=" + IsUsableTeamsBrowserTitle(title) + ", builtChatTitle=" + builtChatTitle);
                    if (builtChatTitle)
                    {
                        WriteTeamsDebugLog(debugLogPath, "selectedChatTitle=" + ToLogValue(title));
                        return new TeamsAppNameResult(title, suggestedFallbackName);
                    }

                    if (IsUsableTeamsBrowserTitle(title))
                    {
                        if (string.Equals(title, lastAcceptedTitle, StringComparison.OrdinalIgnoreCase))
                        {
                            stableTitleCount++;
                        }
                        else
                        {
                            lastAcceptedTitle = title;
                            stableTitleCount = 1;
                        }

                        if (stableTitleCount >= 2)
                        {
                            WriteTeamsDebugLog(debugLogPath, "selectedTitle=" + ToLogValue(title));
                            CloseTeamsWindow(window);
                            return new TeamsAppNameResult(AppendTeamsTitle(prefix, title), suggestedFallbackName);
                        }
                    }
                }
            }

            if (openedWindow != null && !chatPrefix)
            {
                WriteTeamsDebugLog(debugLogPath, "timeout. closing opened window.");
                CloseTeamsWindow(openedWindow);
            }

            WriteTeamsDebugLog(debugLogPath, "Teams title debug finished without title.");
            return new TeamsAppNameResult(string.Empty, suggestedFallbackName);
            }
            finally
            {
                RestoreForegroundWindow(foregroundWindowToRestore);
            }
        }

        private static TeamsWindowInfo FindNewTeamsWindow(Dictionary<int, TeamsWindowInfo> beforeWindows)
        {
            Dictionary<int, TeamsWindowInfo> windows = GetTeamsWindows();
            foreach (TeamsWindowInfo window in windows.Values)
            {
                if (!beforeWindows.ContainsKey(window.Handle))
                {
                    return window;
                }
            }

            return null;
        }

        private static TeamsTitleScanResult GetTeamsVisibleTitle(TeamsWindowInfo window, string prefix, string debugLogPath, string targetMessageId, bool isChannelChatLink, string sourceUrl)
        {
            if (window == null || window.Element == null)
            {
                return new TeamsTitleScanResult(string.Empty, string.Empty);
            }

            string rawWindowCaption = GetNativeWindowTitle(window.Handle);
            string windowCaption = CleanTeamsBrowserTitleCandidate(rawWindowCaption);
            WriteTeamsDebugLog(debugLogPath, "nativeWindowTitle raw=" + ToLogValue(rawWindowCaption) + ", cleaned=" + ToLogValue(windowCaption) + ", visibleUsable=" + IsLikelyTeamsVisibleTitle(windowCaption, prefix) + ", nativeUsable=" + IsLikelyTeamsNativeMeetingTitle(windowCaption, prefix) + ", score=" + ScoreTeamsVisibleTitleCandidate(windowCaption, prefix));
            if (IsLikelyTeamsNativeMeetingTitle(windowCaption, prefix))
            {
                return new TeamsTitleScanResult(windowCaption, string.Empty);
            }

            bool meetingPrefix = IsTeamsMeetingPrefix(prefix);
            bool chatPrefix = IsTeamsChatPrefix(prefix);
            string channelGroupId = isChannelChatLink ? GetQueryParameter(sourceUrl, "groupId") : string.Empty;
            string channelTeamName = isChannelChatLink ? GetTeamsChannelTeamNameFromUrlOrCache(sourceUrl) : string.Empty;
            string channelName = isChannelChatLink ? GetTeamsChannelNameFromUrl(sourceUrl) : string.Empty;
            if (isChannelChatLink)
            {
                WriteTeamsDebugLog(debugLogPath, "channelUrlInfo groupId=" + ToLogValue(channelGroupId) + ", teamName=" + ToLogValue(channelTeamName) + ", channelName=" + ToLogValue(channelName));
            }

            TeamsChatInfo chatInfo = chatPrefix ? new TeamsChatInfo(targetMessageId, isChannelChatLink, channelGroupId, channelTeamName, channelName) : null;
            if (meetingPrefix)
            {
                WriteTeamsDebugLog(debugLogPath, "meeting prefix detected. UI element scan will be logged for investigation, but element titles will not be auto-selected.");
            }
            else if (chatPrefix)
            {
                WriteTeamsDebugLog(debugLogPath, "chat prefix detected. UI element scan will be logged for speaker/date/message investigation.");
                UpdateTeamsChatNameFromRaw(chatInfo, rawWindowCaption);
                UpdateTeamsChatNameFromRaw(chatInfo, window.Title);
                WriteTeamsDebugLog(debugLogPath, "chat initial chatName=" + ToLogValue(chatInfo.ChatName));
            }

            string bestTitle = string.Empty;
            string bestFallbackName = string.Empty;
            int bestScore = -1;

            try
            {
                bool detailScan = meetingPrefix || chatPrefix;
                Condition scanCondition = detailScan
                    ? new OrCondition(
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Hyperlink),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Header),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Document),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Group),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.RadioButton),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ToolBar),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.DataItem),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Tree),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.List),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Table),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.HeaderItem),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Image),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Custom))
                    : new OrCondition(
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Hyperlink),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Header),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Document),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane));
                AutomationElementCollection elements = window.Element.FindAll(
                    TreeScope.Descendants,
                    scanCondition);

                int inspected = 0;
                int maxInspect = isChannelChatLink ? 5000 : (meetingPrefix ? 1500 : (chatPrefix && !string.IsNullOrWhiteSpace(targetMessageId) ? 2500 : (chatPrefix ? 900 : 500)));
                WriteTeamsDebugLog(debugLogPath, "element scan started. meetingPrefix=" + meetingPrefix + ", chatPrefix=" + chatPrefix + ", elementCount=" + elements.Count + ", maxInspect=" + maxInspect);
                foreach (AutomationElement element in elements)
                {
                    if (++inspected > maxInspect)
                    {
                        WriteTeamsDebugLog(debugLogPath, "element scan stopped at " + maxInspect);
                        break;
                    }

                    string rawName = string.Empty;
                    string candidate = string.Empty;
                    string controlType = string.Empty;
                    string localizedControlType = string.Empty;
                    string automationId = string.Empty;
                    string className = string.Empty;
                    string helpText = string.Empty;
                    int score = -999;
                    bool usable = false;
                    try
                    {
                        rawName = element.Current.Name ?? string.Empty;
                        try { controlType = element.Current.ControlType.ProgrammaticName; } catch { }
                        try { localizedControlType = element.Current.LocalizedControlType; } catch { }
                        try { automationId = element.Current.AutomationId; } catch { }
                        try { className = element.Current.ClassName; } catch { }
                        try { helpText = element.Current.HelpText; } catch { }
                        candidate = CleanTeamsBrowserTitleCandidate(rawName);
                        usable = IsLikelyTeamsVisibleTitle(candidate, prefix);
                        score = ScoreTeamsVisibleTitleCandidate(candidate, prefix);
                        if (TeamsDebugLogEnabled)
                        {
                            WriteTeamsDebugLog(debugLogPath, FormatTeamsElementDebug(inspected, element, rawName, candidate, usable, score));
                            if (isChannelChatLink && IsPotentialTeamsChannelTeamElement(rawName, candidate, automationId, localizedControlType))
                            {
                                WriteTeamsDebugLog(
                                    debugLogPath,
                                    "channelTeamProbe element[" + inspected + "] rawName=" + ToLogValue(rawName)
                                    + ", cleaned=" + ToLogValue(candidate)
                                    + ", controlType=" + ToLogValue(controlType)
                                    + ", localizedControlType=" + ToLogValue(localizedControlType)
                                    + ", automationId=" + ToLogValue(automationId)
                                    + ", helpText=" + ToLogValue(helpText));
                            }
                        }
                        if (chatPrefix)
                        {
                            UpdateTeamsChatInfo(chatInfo, rawName, helpText, controlType, localizedControlType, automationId, className, debugLogPath);
                            if (chatInfo.HasWholeChatShortcutName && !isChannelChatLink)
                            {
                                string chatTitle = BuildTeamsWholeChatShortcutTitle(GetTeamsChatDisplayName(chatInfo));
                                WriteTeamsDebugLog(debugLogPath, "wholeChatShortcutTitle=" + ToLogValue(chatTitle));
                                return new TeamsTitleScanResult(chatTitle, string.Empty);
                            }

                            if (chatInfo.HasExactShortcutName)
                            {
                                string chatTitle = BuildTeamsChatShortcutTitle(chatInfo);
                                WriteTeamsDebugLog(debugLogPath, "chatShortcutTitle exact=" + ToLogValue(chatTitle));
                                return new TeamsTitleScanResult(chatTitle, string.Empty);
                            }
                        }

                        if (meetingPrefix && string.IsNullOrWhiteSpace(bestFallbackName))
                        {
                            bestFallbackName = GetTeamsMeetingDateFallbackName(candidate);
                            if (!string.IsNullOrWhiteSpace(bestFallbackName))
                            {
                                WriteTeamsDebugLog(debugLogPath, "dateFallbackCandidate=" + ToLogValue(candidate) + ", fallbackName=" + ToLogValue(bestFallbackName));
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        WriteTeamsDebugLog(debugLogPath, "element[" + inspected + "] read failed: " + ex.GetType().FullName + ": " + ex.Message);
                        continue;
                    }

                    if (!usable)
                    {
                        continue;
                    }

                    if (score > bestScore)
                    {
                        bestScore = score;
                        bestTitle = candidate;
                        WriteTeamsDebugLog(debugLogPath, "bestTitleUpdated score=" + bestScore + ", title=" + ToLogValue(bestTitle));
                    }
                }
            }
            catch (Exception ex)
            {
                WriteTeamsDebugLog(debugLogPath, "FindAll failed: " + ex.GetType().FullName + ": " + ex.Message);
            }

            WriteTeamsDebugLog(debugLogPath, "bestTitle=" + ToLogValue(bestTitle) + ", bestScore=" + bestScore);
            if (meetingPrefix)
            {
                WriteTeamsDebugLog(debugLogPath, "meeting prefix detected. ignoring UI element bestTitle and falling back to manual input if native title was not available.");
                return new TeamsTitleScanResult(string.Empty, bestFallbackName);
            }

            if (chatPrefix && chatInfo != null)
            {
                FillTeamsChatTimestampFromTargetMessageId(chatInfo, debugLogPath);
            }

            if (chatPrefix && chatInfo != null && chatInfo.HasShortcutName)
            {
                string chatTitle = BuildTeamsChatShortcutTitle(chatInfo);
                WriteTeamsDebugLog(debugLogPath, "chatShortcutTitle fallback=" + ToLogValue(chatTitle) + ", targetMatched=" + chatInfo.HasTargetTimestamp);
                return new TeamsTitleScanResult(chatTitle, string.Empty);
            }

            if (chatPrefix && chatInfo != null && chatInfo.HasWholeChatShortcutName)
            {
                string chatTitle = BuildTeamsWholeChatShortcutTitle(GetTeamsChatDisplayName(chatInfo));
                WriteTeamsDebugLog(debugLogPath, "wholeChatShortcutTitle fallback=" + ToLogValue(chatTitle));
                return new TeamsTitleScanResult(chatTitle, string.Empty);
            }

            if (chatPrefix)
            {
                WriteTeamsDebugLog(debugLogPath, "chat prefix detected. ignoring generic bestTitle because chat speaker/timestamp was not completed.");
                return new TeamsTitleScanResult(string.Empty, string.Empty);
            }

            return new TeamsTitleScanResult(bestTitle, string.Empty);
        }

        private static string GetNativeWindowTitle(int handle)
        {
            if (handle == 0)
            {
                return string.Empty;
            }

            try
            {
                IntPtr hWnd = new IntPtr(handle);
                int length = GetWindowTextLength(hWnd);
                if (length <= 0)
                {
                    return string.Empty;
                }

                StringBuilder buffer = new StringBuilder(length + 1);
                if (GetWindowText(hWnd, buffer, buffer.Capacity) <= 0)
                {
                    return string.Empty;
                }

                return buffer.ToString();
            }
            catch
            {
                return string.Empty;
            }
        }

        private static void RestoreForegroundWindow(IntPtr foregroundWindow)
        {
            if (foregroundWindow == IntPtr.Zero)
            {
                return;
            }

            try
            {
                if (!IsWindow(foregroundWindow) || GetForegroundWindow() == foregroundWindow)
                {
                    return;
                }

                SetForegroundWindow(foregroundWindow);
            }
            catch
            {
            }
        }

        private static string CreateTeamsDebugLogPath()
        {
            if (!TeamsDebugLogEnabled)
            {
                return string.Empty;
            }

            try
            {
                string baseDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "M365LinkShortcut",
                    "Logs");
                Directory.CreateDirectory(baseDir);
                return Path.Combine(baseDir, "teams-title-debug-" + DateTime.Now.ToString("yyyyMMdd-HHmmss-fff") + ".log");
            }
            catch
            {
                return string.Empty;
            }
        }

        private static void WriteTeamsDebugLog(string path, string message)
        {
            if (!TeamsDebugLogEnabled || string.IsNullOrWhiteSpace(path))
            {
                return;
            }

            try
            {
                File.AppendAllText(
                    path,
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " " + message + Environment.NewLine,
                    Encoding.UTF8);
            }
            catch
            {
            }
        }

        private static string FormatTeamsWindowSummary(string label, TeamsWindowInfo window)
        {
            if (window == null)
            {
                return label + " window=null";
            }

            string nativeTitle = GetNativeWindowTitle(window.Handle);
            return label
                + " handle=" + window.Handle
                + ", processName=" + ToLogValue(window.ProcessName)
                + ", className=" + ToLogValue(window.ClassName)
                + ", nativeTitle=" + ToLogValue(nativeTitle)
                + ", automationName=" + ToLogValue(window.Title)
                + ", nativeCleaned=" + ToLogValue(CleanTeamsBrowserTitleCandidate(nativeTitle))
                + ", automationCleaned=" + ToLogValue(CleanTeamsBrowserTitleCandidate(window.Title));
        }

        private static string FormatTeamsElementDebug(int index, AutomationElement element, string rawName, string candidate, bool usable, int score)
        {
            string controlType = string.Empty;
            string localizedControlType = string.Empty;
            string className = string.Empty;
            string automationId = string.Empty;
            string helpText = string.Empty;
            string itemStatus = string.Empty;
            string itemType = string.Empty;
            string frameworkId = string.Empty;
            string accessKey = string.Empty;
            string acceleratorKey = string.Empty;
            string valuePatternValue = string.Empty;
            string boundingRectangle = string.Empty;
            bool isOffscreen = false;

            try { controlType = element.Current.ControlType.ProgrammaticName; } catch { }
            try { localizedControlType = element.Current.LocalizedControlType; } catch { }
            try { className = element.Current.ClassName; } catch { }
            try { automationId = element.Current.AutomationId; } catch { }
            try { helpText = element.Current.HelpText; } catch { }
            try { itemStatus = element.Current.ItemStatus; } catch { }
            try { itemType = element.Current.ItemType; } catch { }
            try { frameworkId = element.Current.FrameworkId; } catch { }
            try { accessKey = element.Current.AccessKey; } catch { }
            try { acceleratorKey = element.Current.AcceleratorKey; } catch { }
            try
            {
                object valuePatternObject;
                if (element.TryGetCurrentPattern(ValuePattern.Pattern, out valuePatternObject))
                {
                    valuePatternValue = ((ValuePattern)valuePatternObject).Current.Value;
                }
            }
            catch { }
            try { isOffscreen = element.Current.IsOffscreen; } catch { }
            try
            {
                System.Windows.Rect rect = element.Current.BoundingRectangle;
                boundingRectangle = string.Format("{0:0},{1:0},{2:0},{3:0}", rect.Left, rect.Top, rect.Width, rect.Height);
            }
            catch
            {
            }

            return "element[" + index + "]"
                + " usable=" + usable
                + ", score=" + score
                + ", controlType=" + ToLogValue(controlType)
                + ", localizedControlType=" + ToLogValue(localizedControlType)
                + ", className=" + ToLogValue(className)
                + ", automationId=" + ToLogValue(automationId)
                + ", isOffscreen=" + isOffscreen
                + ", rect=" + ToLogValue(boundingRectangle)
                + ", rawName=" + ToLogValue(rawName)
                + ", cleaned=" + ToLogValue(candidate)
                + ", helpText=" + ToLogValue(helpText)
                + ", itemStatus=" + ToLogValue(itemStatus)
                + ", itemType=" + ToLogValue(itemType)
                + ", frameworkId=" + ToLogValue(frameworkId)
                + ", accessKey=" + ToLogValue(accessKey)
                + ", acceleratorKey=" + ToLogValue(acceleratorKey)
                + ", value=" + ToLogValue(valuePatternValue);
        }

        private static string ToLogValue(string value)
        {
            if (value == null)
            {
                return "<null>";
            }

            string normalized = value.Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t");
            if (normalized.Length > 500)
            {
                normalized = normalized.Substring(0, 500) + "...";
            }

            return "\"" + normalized + "\"";
        }

        private static bool IsLikelyTeamsVisibleTitle(string candidate, string prefix)
        {
            if (!IsUsableTeamsBrowserTitle(candidate))
            {
                return false;
            }

            if (candidate.Length < 4 || candidate.Length > 120)
            {
                return false;
            }

            if (candidate.IndexOf("@", StringComparison.OrdinalIgnoreCase) >= 0
                || candidate.IndexOf("Microsoft Teams", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return false;
            }

            if (IsTeamsMeetingPrefix(prefix) && IsTeamsMeetingOptionText(candidate))
            {
                return false;
            }

            string compact = Regex.Replace(candidate, "\\s+", string.Empty);
            if (Regex.IsMatch(compact, "^(?:OK|Cancel|Share|Open|Close|Copy|More|Join|Chat|Calendar|\\u5171\\u6709|\\u958b\\u304f|\\u9589\\u3058\\u308b|\\u30b3\\u30d4\\u30fc|\\u305d\\u306e\\u4ed6|\\u53c2\\u52a0|\\u30c1\\u30e3\\u30c3\\u30c8|\\u4f1a\\u8b70|\\u4e88\\u5b9a\\u8868|Teams\\u4f1a\\u8b70|Teams\\u30c1\\u30e3\\u30c3\\u30c8|Teams\\u30ea\\u30f3\\u30af)$", RegexOptions.IgnoreCase))
            {
                return false;
            }

            if (Regex.IsMatch(compact, "\\u30d3\\u30c7\\u30aa.*\\u30aa\\u30fc\\u30c7\\u30a3\\u30aa.*\\u30aa\\u30d7\\u30b7\\u30e7\\u30f3.*\\u9078\\u629e", RegexOptions.IgnoreCase)
                || Regex.IsMatch(compact, "video.*audio.*option", RegexOptions.IgnoreCase))
            {
                return false;
            }

            return true;
        }

        private static bool IsTeamsMeetingPrefix(string prefix)
        {
            return string.Equals(prefix, "Teams \u4f1a\u8b70", StringComparison.OrdinalIgnoreCase);
        }

        private static string StripTeamsMeetingPrefix(string title)
        {
            if (string.IsNullOrWhiteSpace(title))
            {
                return string.Empty;
            }

            string value = title.Trim();
            Match wrappedMatch = Regex.Match(value, "^Teams\\s*\u4f1a\u8b70\\s*\u3010(?<title>.+)\u3011$", RegexOptions.IgnoreCase);
            if (wrappedMatch.Success)
            {
                return wrappedMatch.Groups["title"].Value.Trim();
            }

            return Regex.Replace(value, "^Teams\\s*\u4f1a\u8b70\\s*[-_\u2013\u2014]\\s*", string.Empty, RegexOptions.IgnoreCase).Trim();
        }

        private static bool IsTeamsChatPrefix(string prefix)
        {
            return string.Equals(prefix, "Teams \u30c1\u30e3\u30c3\u30c8", StringComparison.OrdinalIgnoreCase);
        }

        private static string StripTeamsChatPrefix(string title)
        {
            if (string.IsNullOrWhiteSpace(title))
            {
                return string.Empty;
            }

            string value = title.Trim();
            Match wrappedMatch = Regex.Match(value, "^Teams\\s*\u30c1\u30e3\u30c3\u30c8\\s*\u3010(?<title>.+)\u3011$", RegexOptions.IgnoreCase);
            if (wrappedMatch.Success)
            {
                return wrappedMatch.Groups["title"].Value.Trim();
            }

            return Regex.Replace(value, "^Teams\\s*\u30c1\u30e3\u30c3\u30c8\\s*[-_\u2013\u2014]\\s*", string.Empty, RegexOptions.IgnoreCase).Trim();
        }

        private static bool IsBuiltTeamsChatShortcutTitle(string title)
        {
            return !string.IsNullOrWhiteSpace(title)
                && title.StartsWith("Teams\u30c1\u30e3\u30c3\u30c8 \u3010", StringComparison.OrdinalIgnoreCase);
        }

        private static string GetTeamsMessageId(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                return string.Empty;
            }

            Match match = Regex.Match(url, "/message/[^/]+/(?<id>\\d+)", RegexOptions.IgnoreCase);
            return match.Success ? match.Groups["id"].Value : string.Empty;
        }

        private static void FillTeamsChatTimestampFromTargetMessageId(TeamsChatInfo chatInfo, string debugLogPath)
        {
            if (chatInfo == null
                || string.IsNullOrWhiteSpace(chatInfo.TargetMessageId)
                || !string.IsNullOrWhiteSpace(chatInfo.Timestamp))
            {
                return;
            }

            DateTime timestamp;
            if (TryGetTeamsChatTimestampFromMessageId(chatInfo.TargetMessageId, out timestamp))
            {
                chatInfo.Timestamp = FormatDateTimeForJapaneseTitle(timestamp);
                chatInfo.HasTargetTimestamp = true;
                WriteTeamsDebugLog(debugLogPath, "chat timestampFromMessageId=" + ToLogValue(chatInfo.Timestamp));
            }
        }

        internal static bool TryGetTeamsChatTimestampFromMessageId(string messageId, out DateTime value)
        {
            value = DateTime.MinValue;
            long milliseconds;
            if (string.IsNullOrWhiteSpace(messageId)
                || !long.TryParse(messageId, out milliseconds)
                || milliseconds < 946684800000L)
            {
                return false;
            }

            try
            {
                value = DateTimeOffset.FromUnixTimeMilliseconds(milliseconds).LocalDateTime;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void UpdateTeamsChatInfo(
            TeamsChatInfo chatInfo,
            string rawName,
            string helpText,
            string controlType,
            string localizedControlType,
            string automationId,
            string className,
            string debugLogPath)
        {
            if (chatInfo == null)
            {
                return;
            }

            chatInfo.ScanElementIndex++;
            UpdateTeamsChannelTeamFromTreeItem(chatInfo, rawName, controlType, localizedControlType, className, debugLogPath);
            UpdateTeamsChatNameFromRaw(chatInfo, rawName);

            string speaker = ExtractTeamsChatSpeaker(rawName);
            if (!string.IsNullOrWhiteSpace(speaker)
                && string.Equals(controlType, "ControlType.Text", StringComparison.OrdinalIgnoreCase)
                && string.Equals(localizedControlType, "\u898b\u51fa\u3057", StringComparison.OrdinalIgnoreCase))
            {
                chatInfo.LastSpeakerCandidate = speaker;
                chatInfo.LastSpeakerCandidateIndex = chatInfo.ScanElementIndex;
                if (!string.IsNullOrWhiteSpace(chatInfo.TargetMessageId))
                {
                    WriteTeamsDebugLog(debugLogPath, "chat speakerCandidateForTarget=" + ToLogValue(speaker));
                }
                else if (string.IsNullOrWhiteSpace(chatInfo.Speaker))
                {
                    chatInfo.Speaker = speaker;
                    WriteTeamsDebugLog(debugLogPath, "chat speakerCandidate=" + ToLogValue(speaker));
                }
                else if (!string.Equals(chatInfo.Speaker, speaker, StringComparison.OrdinalIgnoreCase)
                    && !chatInfo.HasTargetTimestamp)
                {
                    chatInfo.Speaker = speaker;
                    WriteTeamsDebugLog(debugLogPath, "chat speakerCandidateUpdated=" + ToLogValue(speaker));
                }
            }

            Match messageBodyMatch = Regex.Match(automationId ?? string.Empty, "^message-body-(?<id>\\d+)$", RegexOptions.IgnoreCase);
            if (messageBodyMatch.Success)
            {
                string id = messageBodyMatch.Groups["id"].Value;
                bool targetBodyMatch = !string.IsNullOrWhiteSpace(chatInfo.TargetMessageId)
                    && string.Equals(id, chatInfo.TargetMessageId, StringComparison.OrdinalIgnoreCase);
                if (targetBodyMatch)
                {
                    chatInfo.HasTargetBody = true;
                    chatInfo.TargetBodyElementIndex = chatInfo.ScanElementIndex;
                    string bodySpeaker = ExtractTeamsChatSpeakerFromMessageBody(rawName);
                    if (!string.IsNullOrWhiteSpace(bodySpeaker))
                    {
                        chatInfo.Speaker = bodySpeaker;
                        chatInfo.LastSpeakerCandidate = bodySpeaker;
                        chatInfo.TargetSpeakerCandidate = bodySpeaker;
                        WriteTeamsDebugLog(debugLogPath, "chat targetBodySpeaker=" + ToLogValue(bodySpeaker));
                    }
                    else if (IsRecentTeamsChatSpeakerCandidate(chatInfo))
                    {
                        chatInfo.TargetSpeakerCandidate = chatInfo.LastSpeakerCandidate;
                        WriteTeamsDebugLog(debugLogPath, "chat targetBodyNearbySpeaker=" + ToLogValue(chatInfo.TargetSpeakerCandidate));
                    }
                }
            }

            Match timestampMatch = Regex.Match(automationId ?? string.Empty, "^timestamp-(?<id>\\d+)$", RegexOptions.IgnoreCase);
            if (timestampMatch.Success
                && string.Equals(controlType, "ControlType.Text", StringComparison.OrdinalIgnoreCase)
                && string.Equals(localizedControlType, "\u65e5\u6642", StringComparison.OrdinalIgnoreCase))
            {
                string id = timestampMatch.Groups["id"].Value;
                string timestamp = CleanTeamsChatTimestamp(string.IsNullOrWhiteSpace(helpText) ? rawName : helpText);
                bool targetMatch = !string.IsNullOrWhiteSpace(chatInfo.TargetMessageId)
                    && string.Equals(id, chatInfo.TargetMessageId, StringComparison.OrdinalIgnoreCase);

                if (!string.IsNullOrWhiteSpace(chatInfo.TargetMessageId) && !targetMatch)
                {
                    return;
                }

                if (targetMatch || string.IsNullOrWhiteSpace(chatInfo.Timestamp))
                {
                    chatInfo.Timestamp = timestamp;
                    chatInfo.HasTargetTimestamp = targetMatch;
                    if (!string.IsNullOrWhiteSpace(chatInfo.TargetSpeakerCandidate))
                    {
                        chatInfo.Speaker = chatInfo.TargetSpeakerCandidate;
                    }
                    else if (targetMatch && IsRecentTeamsChatSpeakerCandidate(chatInfo))
                    {
                        chatInfo.Speaker = chatInfo.LastSpeakerCandidate;
                    }
                    else if (!targetMatch && !string.IsNullOrWhiteSpace(chatInfo.LastSpeakerCandidate))
                    {
                        chatInfo.Speaker = chatInfo.LastSpeakerCandidate;
                    }

                    WriteTeamsDebugLog(
                        debugLogPath,
                        "chat timestampCandidate id=" + ToLogValue(id)
                        + ", targetMatch=" + targetMatch
                        + ", timestamp=" + ToLogValue(timestamp)
                        + ", speaker=" + ToLogValue(chatInfo.Speaker));
                }
            }
        }

        private static bool IsRecentTeamsChatSpeakerCandidate(TeamsChatInfo chatInfo)
        {
            if (chatInfo == null
                || string.IsNullOrWhiteSpace(chatInfo.LastSpeakerCandidate)
                || chatInfo.LastSpeakerCandidateIndex <= 0)
            {
                return false;
            }

            int referenceIndex = chatInfo.HasTargetBody && chatInfo.TargetBodyElementIndex > 0
                ? chatInfo.TargetBodyElementIndex
                : chatInfo.ScanElementIndex;
            int distance = Math.Abs(referenceIndex - chatInfo.LastSpeakerCandidateIndex);
            return distance <= 25;
        }

        private static void UpdateTeamsChatNameFromRaw(TeamsChatInfo chatInfo, string rawValue)
        {
            if (chatInfo == null)
            {
                return;
            }

            if (chatInfo.IsChannelLink)
            {
                string channelDisplayName = ExtractTeamsChatName(rawValue);
                if (!string.IsNullOrWhiteSpace(channelDisplayName))
                {
                    if (string.IsNullOrWhiteSpace(chatInfo.ChatName)
                        || ShouldPreferTeamsChannelDisplayName(chatInfo.ChatName, channelDisplayName))
                    {
                        chatInfo.ChatName = channelDisplayName;
                    }
                }
            }

            if (!string.IsNullOrWhiteSpace(chatInfo.ChatName))
            {
                return;
            }

            string chatName = ExtractTeamsChatName(rawValue);
            if (!string.IsNullOrWhiteSpace(chatName))
            {
                chatInfo.ChatName = chatName;
            }
        }

        private static void UpdateTeamsChannelTeamFromTreeItem(TeamsChatInfo chatInfo, string rawValue, string controlType, string localizedControlType, string className, string debugLogPath)
        {
            if (chatInfo == null || !chatInfo.IsChannelLink || string.IsNullOrWhiteSpace(rawValue))
            {
                return;
            }

            bool isTreeItem = string.Equals(controlType, "ControlType.TreeItem", StringComparison.OrdinalIgnoreCase)
                || string.Equals(localizedControlType, "\u30c4\u30ea\u30fc\u9805\u76ee", StringComparison.OrdinalIgnoreCase);
            if (!isTreeItem)
            {
                return;
            }

            string value = CleanTitleCandidate(rawValue);
            if (string.IsNullOrWhiteSpace(value))
            {
                return;
            }

            bool isSelectedChannel = IsSelectedTeamsTreeItemClass(className) && IsTeamsChannelTreeItemForChannel(value, chatInfo.ChatName);
            string displayChannelName = ExtractTeamsChannelDisplayNameFromTreeItem(value, chatInfo.ChatName);
            if (!string.IsNullOrWhiteSpace(displayChannelName)
                && (string.IsNullOrWhiteSpace(chatInfo.ChatName)
                    || ShouldPreferTeamsChannelDisplayName(chatInfo.ChatName, displayChannelName)))
            {
                chatInfo.ChatName = displayChannelName;
                WriteTeamsDebugLog(debugLogPath, "channel displayNameFromTree=" + ToLogValue(displayChannelName) + ", raw=" + ToLogValue(value));
            }

            string teamCandidate = ExtractTeamsTeamNameFromTreeItem(value);
            if (!string.IsNullOrWhiteSpace(teamCandidate))
            {
                chatInfo.LastTreeTeamCandidate = teamCandidate;
                WriteTeamsDebugLog(debugLogPath, "channel tree teamCandidate=" + ToLogValue(teamCandidate) + ", raw=" + ToLogValue(value));
            }

            string directTeam = ExtractTeamsTeamNameFromChannelTreeItem(value, chatInfo.ChatName);
            if (!string.IsNullOrWhiteSpace(directTeam)
                && (isSelectedChannel || string.IsNullOrWhiteSpace(chatInfo.TeamName)))
            {
                chatInfo.TeamName = directTeam;
                SaveTeamsChannelTeamNameCacheFromInfo(chatInfo);
                WriteTeamsDebugLog(debugLogPath, "channel directTeamFromTree=" + ToLogValue(directTeam) + ", selected=" + isSelectedChannel + ", raw=" + ToLogValue(value));
                return;
            }

            if (!string.IsNullOrWhiteSpace(chatInfo.LastTreeTeamCandidate)
                && IsTeamsChannelTreeItemForChannel(value, chatInfo.ChatName)
                && (isSelectedChannel || string.IsNullOrWhiteSpace(chatInfo.TeamName)))
            {
                chatInfo.TeamName = chatInfo.LastTreeTeamCandidate;
                SaveTeamsChannelTeamNameCacheFromInfo(chatInfo);
                WriteTeamsDebugLog(debugLogPath, "channel teamFromTreeParent=" + ToLogValue(chatInfo.TeamName) + ", selected=" + isSelectedChannel + ", channelRaw=" + ToLogValue(value));
            }
        }

        private static string ExtractTeamsTeamNameFromTreeItem(string value)
        {
            Match match = Regex.Match(value ?? string.Empty, "^(?:\u672a\u8aad\\s+)?\u30c1\u30fc\u30e0\\s+(?<team>.+?)(?:\\s+\u30b3\u30f3\u30c6\u30ad\u30b9\u30c8\\s+\u30e1\u30cb\u30e5\u30fc\u3042\u308a)?$", RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                return string.Empty;
            }

            string team = match.Groups["team"].Value.Trim();
            return IsUsableTitleCandidate(team) ? team : string.Empty;
        }

        private static string ExtractTeamsTeamNameFromChannelTreeItem(string value, string channelName)
        {
            if (string.IsNullOrWhiteSpace(value) || string.IsNullOrWhiteSpace(channelName))
            {
                return string.Empty;
            }

            string matchedChannel;
            string tail;
            if (!TryMatchTeamsChannelTreeItem(value, channelName, out matchedChannel, out tail))
            {
                return string.Empty;
            }

            string team = tail.Trim();
            team = Regex.Replace(team, "\\s+(?:\\d{1,2}/\\d{1,2}|\\d{4}/\\d{1,2}/\\d{1,2})$", string.Empty).Trim();
            team = Regex.Replace(team, "\\s+\u30e1\u30f3\u30b7\u30e7\u30f3\u3055\u308c\u305f(?:\u30c1\u30fc\u30e0|\u30c1\u30e3\u30cd\u30eb)$", string.Empty).Trim();
            if (IsLikelyTeamsTreeDateOrStatus(team))
            {
                return string.Empty;
            }

            return IsUsableTitleCandidate(team) ? team : string.Empty;
        }

        private static bool IsTeamsChannelTreeItemForChannel(string value, string channelName)
        {
            if (string.IsNullOrWhiteSpace(value) || string.IsNullOrWhiteSpace(channelName))
            {
                return false;
            }

            string matchedChannel;
            string tail;
            return TryMatchTeamsChannelTreeItem(value, channelName, out matchedChannel, out tail);
        }

        private static string ExtractTeamsChannelDisplayNameFromTreeItem(string value, string channelName)
        {
            string matchedChannel;
            string tail;
            if (!TryMatchTeamsChannelTreeItem(value, channelName, out matchedChannel, out tail))
            {
                return string.Empty;
            }

            return IsUsableTitleCandidate(matchedChannel) ? matchedChannel : string.Empty;
        }

        private static bool TryMatchTeamsChannelTreeItem(string value, string channelName, out string matchedChannel, out string tail)
        {
            matchedChannel = string.Empty;
            tail = string.Empty;
            if (string.IsNullOrWhiteSpace(value) || string.IsNullOrWhiteSpace(channelName))
            {
                return false;
            }

            List<string> names = new List<string>();
            names.Add(channelName.Trim());
            if (string.Equals(channelName.Trim(), "General", StringComparison.OrdinalIgnoreCase))
            {
                names.Add("\u4e00\u822c");
            }
            else if (string.Equals(channelName.Trim(), "\u4e00\u822c", StringComparison.OrdinalIgnoreCase))
            {
                names.Add("General");
            }

            foreach (string name in names)
            {
                Match match = Regex.Match(
                    value,
                    "^(?:\u672a\u8aad\\s+)?(?:\u5171\u6709\\s+)?(?:\u30d7\u30e9\u30a4\u30d9\u30fc\u30c8\\s+)?\u30c1\u30e3\u30cd\u30eb\\s+"
                    + Regex.Escape(name)
                    + "(?:\\s+(?<tail>.*))?$",
                    RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    matchedChannel = name;
                    tail = match.Groups["tail"].Value.Trim();
                    return true;
                }
            }

            return false;
        }

        private static bool IsSelectedTeamsTreeItemClass(string className)
        {
            return Regex.IsMatch(className ?? string.Empty, "(?:^|\\s)ferormf(?:\\s|$)", RegexOptions.IgnoreCase);
        }

        private static bool ShouldPreferTeamsChannelDisplayName(string currentName, string displayName)
        {
            if (string.IsNullOrWhiteSpace(displayName))
            {
                return false;
            }

            return IsSameTeamsChannelName(currentName, displayName)
                && !string.Equals(currentName.Trim(), displayName.Trim(), StringComparison.Ordinal);
        }

        private static bool IsSameTeamsChannelName(string left, string right)
        {
            string a = (left ?? string.Empty).Trim();
            string b = (right ?? string.Empty).Trim();
            if (string.Equals(a, b, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return (string.Equals(a, "General", StringComparison.OrdinalIgnoreCase) && string.Equals(b, "\u4e00\u822c", StringComparison.OrdinalIgnoreCase))
                || (string.Equals(a, "\u4e00\u822c", StringComparison.OrdinalIgnoreCase) && string.Equals(b, "General", StringComparison.OrdinalIgnoreCase));
        }

        private static bool IsLikelyTeamsTreeDateOrStatus(string value)
        {
            string text = CleanTitleCandidate(value);
            if (string.IsNullOrWhiteSpace(text))
            {
                return true;
            }

            return Regex.IsMatch(text, "^(?:\\d{1,2}[/_]\\d{1,2}|\\d{4}[/_]\\d{1,2}[/_]\\d{1,2}|\\d{1,2}\u6708\\d{1,2}\u65e5)(?:\\s+\u30e1\u30f3\u30b7\u30e7\u30f3\u3055\u308c\u305f(?:\u30c1\u30fc\u30e0|\u30c1\u30e3\u30cd\u30eb))?$", RegexOptions.IgnoreCase);
        }

        private static void SaveTeamsChannelTeamNameCacheFromInfo(TeamsChatInfo chatInfo)
        {
            if (chatInfo == null || string.IsNullOrWhiteSpace(chatInfo.TeamName))
            {
                return;
            }

            SaveTeamsChannelTeamNameCache(chatInfo.GroupId, chatInfo.TeamName);
        }

        private static bool TryExtractTeamsChannelNames(string rawValue, out string teamName, out string channelName)
        {
            teamName = string.Empty;
            channelName = string.Empty;
            if (string.IsNullOrWhiteSpace(rawValue))
            {
                return false;
            }

            string value = CleanTitleCandidate(rawValue);
            value = RepairUtf8AsShiftJisMojibake(value);
            string[] parts = value.Split(new[] { '|', '\uFF5C' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < parts.Length; i++)
            {
                parts[i] = CleanTitleCandidate(parts[i]).Trim();
            }

            for (int i = 0; i < parts.Length - 2; i++)
            {
                if (IsTeamsChannelLabel(parts[i]))
                {
                    string first = parts[i + 1];
                    string second = parts[i + 2];
                    if (IsUsableTitleCandidate(first) && IsUsableTitleCandidate(second))
                    {
                        teamName = first;
                        channelName = second;
                        return true;
                    }
                }
            }

            Match match = Regex.Match(value, "(?:^|[\\s\\r\\n])(?:\u30c1\u30fc\u30e0\u3068\u30c1\u30e3\u30cd\u30eb|\u30c1\u30e3\u30cd\u30eb|Teams\\s+and\\s+channels|Channel)\\s*[\\|\\uFF5C]\\s*(?<team>[^\\|\\uFF5C\\r\\n]+)\\s*[\\|\\uFF5C]\\s*(?<channel>[^\\|\\uFF5C\\r\\n]+)", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                string team = match.Groups["team"].Value.Trim();
                string channel = match.Groups["channel"].Value.Trim();
                if (IsUsableTitleCandidate(team) && IsUsableTitleCandidate(channel))
                {
                    teamName = team;
                    channelName = channel;
                    return true;
                }
            }

            return false;
        }

        private static bool IsTeamsChannelLabel(string value)
        {
            string compact = Regex.Replace(value ?? string.Empty, "\\s+", string.Empty);
            return string.Equals(compact, "\u30c1\u30fc\u30e0\u3068\u30c1\u30e3\u30cd\u30eb", StringComparison.OrdinalIgnoreCase)
                || string.Equals(compact, "\u30c1\u30e3\u30cd\u30eb", StringComparison.OrdinalIgnoreCase)
                || string.Equals(compact, "Teamsandchannels", StringComparison.OrdinalIgnoreCase)
                || string.Equals(compact, "Channel", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsPotentialTeamsChannelTeamElement(string rawName, string candidate, string automationId, string localizedControlType)
        {
            string text = ((rawName ?? string.Empty) + " " + (candidate ?? string.Empty) + " " + (automationId ?? string.Empty) + " " + (localizedControlType ?? string.Empty)).Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            if (Regex.IsMatch(text, "(?:\u30c1\u30fc\u30e0|\u30c1\u30e3\u30cd\u30eb|team|channel|thread|group)", RegexOptions.IgnoreCase))
            {
                return true;
            }

            string clean = CleanTitleCandidate(candidate);
            return IsUsableTitleCandidate(clean)
                && clean.Length <= 80
                && !Regex.IsMatch(clean, "(?:Ctrl\\+|Microsoft Teams|https?://|@|\\d{4}\u5e74|\\d{1,2}:\\d{2})", RegexOptions.IgnoreCase);
        }

        private static string ExtractTeamsChatName(string rawValue)
        {
            if (string.IsNullOrWhiteSpace(rawValue))
            {
                return string.Empty;
            }

            string value = CleanTitleCandidate(rawValue);
            value = RepairUtf8AsShiftJisMojibake(value);
            Match match = Regex.Match(value, "(?:^|[\\s\\r\\n])(?:\u30c1\u30e3\u30c3\u30c8|\u30c1\u30fc\u30e0\u3068\u30c1\u30e3\u30cd\u30eb|\u30c1\u30e3\u30cd\u30eb)\\s*[\\|\\uFF5C]\\s*(?<name>[^\\|\\uFF5C\\r\\n]+)", RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                return string.Empty;
            }

            string name = match.Groups["name"].Value.Trim();
            return IsUsableTitleCandidate(name) ? name : string.Empty;
        }

        private static string ExtractTeamsChatSpeaker(string rawValue)
        {
            if (string.IsNullOrWhiteSpace(rawValue))
            {
                return string.Empty;
            }

            string value = CleanTitleCandidate(rawValue);
            value = RepairUtf8AsShiftJisMojibake(value);
            MatchCollection matches = Regex.Matches(value, "\u3001(?<speaker>[^\\u3001\\r\\n]{1,120})\\s+\u304c\u4f5c\u6210");
            if (matches.Count == 0)
            {
                return string.Empty;
            }

            string speaker = matches[matches.Count - 1].Groups["speaker"].Value.Trim();
            return IsUsableTitleCandidate(speaker) ? speaker : string.Empty;
        }

        private static string ExtractTeamsChatSpeakerFromMessageBody(string rawValue)
        {
            if (string.IsNullOrWhiteSpace(rawValue))
            {
                return string.Empty;
            }

            string value = CleanTitleCandidate(rawValue);
            value = RepairUtf8AsShiftJisMojibake(value);
            Match unknownUser = Regex.Match(value, "^(?<speaker>\u4e0d\u660e\u306a\u30e6\u30fc\u30b6\u30fc|Unknown\\s+user)(?:\\s|$)", RegexOptions.IgnoreCase);
            if (unknownUser.Success)
            {
                string speaker = unknownUser.Groups["speaker"].Value.Trim();
                return IsUsableTitleCandidate(speaker) ? speaker : string.Empty;
            }

            Match japaneseEnglishName = Regex.Match(value, "^(?<speaker>(?:[\\p{IsCJKUnifiedIdeographs}\\u3040-\\u30ff\\u3005\\u3006\\u30fc]+\\s+){1,3}[A-Z][A-Za-z]+\\s+[A-Z][A-Za-z]+)(?:\\s|$)");
            if (japaneseEnglishName.Success)
            {
                string speaker = japaneseEnglishName.Groups["speaker"].Value.Trim();
                return IsUsableTitleCandidate(speaker) ? speaker : string.Empty;
            }

            Match englishJapaneseName = Regex.Match(value, "^(?<speaker>[A-Z][A-Za-z]+\\s+[A-Z][A-Za-z]+\\([^\\)\\r\\n]{1,40}\\))(?:\\s|$)");
            if (englishJapaneseName.Success)
            {
                string speaker = englishJapaneseName.Groups["speaker"].Value.Trim();
                return IsUsableTitleCandidate(speaker) ? speaker : string.Empty;
            }

            return string.Empty;
        }

        private static string CleanTeamsChatTimestamp(string rawValue)
        {
            string value = CleanTitleCandidate(rawValue);
            value = value.Trim().TrimEnd('.');
            return value;
        }

        private static string BuildTeamsChatShortcutTitle(TeamsChatInfo chatInfo)
        {
            if (chatInfo == null || !chatInfo.HasShortcutName)
            {
                return string.Empty;
            }

            string timestamp = FormatTeamsChatTimestampForTitle(chatInfo.Timestamp);
            if (string.IsNullOrWhiteSpace(timestamp))
            {
                timestamp = chatInfo.Timestamp.Trim();
            }

            return "Teams\u30c1\u30e3\u30c3\u30c8 \u3010"
                + GetTeamsChatDisplayName(chatInfo)
                + "\u3011 "
                + timestamp
                + " - "
                + chatInfo.Speaker.Trim();
        }

        private static string BuildTeamsWholeChatShortcutTitle(string chatName)
        {
            if (string.IsNullOrWhiteSpace(chatName))
            {
                return string.Empty;
            }

            return "Teams\u30c1\u30e3\u30c3\u30c8 \u3010" + chatName.Trim() + "\u3011";
        }

        private static string GetTeamsChatDisplayName(TeamsChatInfo chatInfo)
        {
            if (chatInfo == null)
            {
                return string.Empty;
            }

            string chatName = (chatInfo.ChatName ?? string.Empty).Trim();
            string teamName = (chatInfo.TeamName ?? string.Empty).Trim();
            if (chatInfo.IsChannelLink
                && !string.IsNullOrWhiteSpace(teamName)
                && !string.IsNullOrWhiteSpace(chatName)
                && !string.Equals(teamName, chatName, StringComparison.OrdinalIgnoreCase))
            {
                return teamName + " - " + chatName;
            }

            return chatName;
        }

        private static string FormatTeamsChatTimestampForTitle(string rawTimestamp)
        {
            DateTime value;
            if (!TryParseTeamsChatTimestamp(rawTimestamp, out value))
            {
                return string.Empty;
            }

            return FormatDateTimeForJapaneseTitle(value);
        }

        internal static string FormatDateTimeForJapaneseTitle(DateTime value)
        {
            string[] weekdays = new[] { "\u65e5", "\u6708", "\u706b", "\u6c34", "\u6728", "\u91d1", "\u571f" };
            return value.Year
                + "\u5e74"
                + value.Month
                + "\u6708"
                + value.Day
                + "\u65e5("
                + weekdays[(int)value.DayOfWeek]
                + ")"
                + value.Hour
                + "\u6642"
                + value.Minute
                + "\u5206";
        }

        internal static bool TryParseTeamsChatTimestamp(string rawTimestamp, out DateTime value)
        {
            value = DateTime.MinValue;
            if (string.IsNullOrWhiteSpace(rawTimestamp))
            {
                return false;
            }

            string normalized = CleanTitleCandidate(rawTimestamp).Trim().TrimEnd('.');
            Match absoluteMatch = Regex.Match(normalized, "^(?<year>\\d{4})\u5e74(?<month>\\d{1,2})\u6708(?<day>\\d{1,2})\u65e5\\s*(?<hour>\\d{1,2}):(?<minute>\\d{2})$");
            if (absoluteMatch.Success)
            {
                return TryCreateDateTime(
                    absoluteMatch.Groups["year"].Value,
                    absoluteMatch.Groups["month"].Value,
                    absoluteMatch.Groups["day"].Value,
                    absoluteMatch.Groups["hour"].Value,
                    absoluteMatch.Groups["minute"].Value,
                    out value);
            }

            Match relativeMatch = Regex.Match(normalized, "^(?<relative>\u4eca\u65e5|\u6628\u65e5)\u306e\\s*(?<hour>\\d{1,2}):(?<minute>\\d{2})$");
            if (relativeMatch.Success)
            {
                DateTime baseDate = DateTime.Today;
                if (string.Equals(relativeMatch.Groups["relative"].Value, "\u6628\u65e5", StringComparison.Ordinal))
                {
                    baseDate = baseDate.AddDays(-1);
                }

                return TryCreateDateTime(
                    baseDate.Year.ToString(),
                    baseDate.Month.ToString(),
                    baseDate.Day.ToString(),
                    relativeMatch.Groups["hour"].Value,
                    relativeMatch.Groups["minute"].Value,
                    out value);
            }

            return false;
        }

        private static bool TryCreateDateTime(string year, string month, string day, string hour, string minute, out DateTime value)
        {
            value = DateTime.MinValue;
            int y;
            int m;
            int d;
            int h;
            int min;
            if (!int.TryParse(year, out y)
                || !int.TryParse(month, out m)
                || !int.TryParse(day, out d)
                || !int.TryParse(hour, out h)
                || !int.TryParse(minute, out min))
            {
                return false;
            }

            try
            {
                value = new DateTime(y, m, d, h, min, 0);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsTeamsMeetingOptionText(string candidate)
        {
            if (string.IsNullOrWhiteSpace(candidate))
            {
                return false;
            }

            return Regex.IsMatch(candidate, "(?:\u9078\u629e\u3057\u305f(?:\u30de\u30a4\u30af|\u30b9\u30d4\u30fc\u30ab\u30fc)|\u30de\u30a4\u30af|\u30b9\u30d4\u30fc\u30ab\u30fc|\u30ab\u30e1\u30e9|\u97f3\u58f0|\u30aa\u30fc\u30c7\u30a3\u30aa|\u30aa\u30d7\u30b7\u30e7\u30f3|\u30df\u30e5\u30fc\u30c8|Ctrl\\+Shift|Shokz|Loop)", RegexOptions.IgnoreCase);
        }

        internal static string GetTeamsMeetingDateFallbackName(string candidate)
        {
            if (string.IsNullOrWhiteSpace(candidate))
            {
                return string.Empty;
            }

            string normalized = Regex.Replace(candidate, "\\s+", " ").Trim();
            Match japaneseDateRangeMatch = Regex.Match(
                normalized,
                "(?<month>\\d{1,2})\\u6708(?<day>\\d{1,2})\\u65e5(?<weekday>\\u6708\\u66dc\\u65e5|\\u706b\\u66dc\\u65e5|\\u6c34\\u66dc\\u65e5|\\u6728\\u66dc\\u65e5|\\u91d1\\u66dc\\u65e5|\\u571f\\u66dc\\u65e5|\\u65e5\\u66dc\\u65e5|[\\u6708\\u706b\\u6c34\\u6728\\u91d1\\u571f\\u65e5])?\\s*(?<startHour>\\d{1,2}):(?<startMinute>\\d{2})\\s*[\\u2013\\u2014\\-]\\s*(?<endHour>\\d{1,2}):(?<endMinute>\\d{2})(?:\\s*JST)?",
                RegexOptions.IgnoreCase);
            if (japaneseDateRangeMatch.Success)
            {
                string weekday = japaneseDateRangeMatch.Groups["weekday"].Value;
                if (weekday.EndsWith("\u66dc\u65e5", StringComparison.Ordinal))
                {
                    weekday = weekday.Substring(0, 1);
                }

                string weekdayPart = string.IsNullOrWhiteSpace(weekday) ? string.Empty : "(" + weekday + ")";
                string display = japaneseDateRangeMatch.Groups["month"].Value
                    + "\u6708"
                    + japaneseDateRangeMatch.Groups["day"].Value
                    + "\u65e5"
                    + weekdayPart
                    + TrimLeadingZero(japaneseDateRangeMatch.Groups["startHour"].Value)
                    + "\u6642"
                    + japaneseDateRangeMatch.Groups["startMinute"].Value
                    + "\u5206"
                    + "-"
                    + TrimLeadingZero(japaneseDateRangeMatch.Groups["endHour"].Value)
                    + "\u6642"
                    + japaneseDateRangeMatch.Groups["endMinute"].Value
                    + "\u5206";
                return "Teams\u4f1a\u8b70 \u3010" + display + "\u3011";
            }

            Match japaneseDateMatch = Regex.Match(
                normalized,
                "(?<month>\\d{1,2})\\u6708(?<day>\\d{1,2})\\u65e5.*?(?<hour>\\d{1,2}):(?<minute>\\d{2})",
                RegexOptions.IgnoreCase);
            if (japaneseDateMatch.Success)
            {
                return "Teams\u4f1a\u8b70 \u3010" + japaneseDateMatch.Value.Trim() + "\u3011";
            }

            Match numericDateMatch = Regex.Match(
                normalized,
                "(?<year>20\\d{2})[/-](?<month>\\d{1,2})[/-](?<day>\\d{1,2}).*?(?<hour>\\d{1,2}):(?<minute>\\d{2})",
                RegexOptions.IgnoreCase);
            if (numericDateMatch.Success)
            {
                return FormatTeamsMeetingDateFallbackName(
                    numericDateMatch.Groups["year"].Value,
                    numericDateMatch.Groups["month"].Value,
                    numericDateMatch.Groups["day"].Value,
                    numericDateMatch.Groups["hour"].Value,
                    numericDateMatch.Groups["minute"].Value);
            }

            return string.Empty;
        }

        private static string TrimLeadingZero(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            int number;
            return int.TryParse(value, out number) ? number.ToString() : value.TrimStart('0');
        }

        private static string FormatTeamsMeetingDateFallbackName(object yearValue, string monthValue, string dayValue, string hourValue, string minuteValue)
        {
            int year;
            int month;
            int day;
            int hour;
            int minute;
            if (!int.TryParse(Convert.ToString(yearValue), out year)
                || !int.TryParse(monthValue, out month)
                || !int.TryParse(dayValue, out day)
                || !int.TryParse(hourValue, out hour)
                || !int.TryParse(minuteValue, out minute))
            {
                return string.Empty;
            }

            try
            {
                DateTime start = new DateTime(year, month, day, hour, minute, 0);
                return "Teams\u4f1a\u8b70 \u3010" + FormatDateTimeForJapaneseTitle(start) + "\u3011";
            }
            catch
            {
                return string.Empty;
            }
        }

        private static bool IsLikelyTeamsNativeMeetingTitle(string candidate, string prefix)
        {
            if (!IsLikelyTeamsVisibleTitle(candidate, prefix))
            {
                return false;
            }

            if (candidate.StartsWith("Microsoft Teams", StringComparison.OrdinalIgnoreCase)
                || candidate.StartsWith("Teams", StringComparison.OrdinalIgnoreCase)
                || candidate.IndexOf("@", StringComparison.OrdinalIgnoreCase) >= 0
                || candidate.IndexOf("|", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return false;
            }

            return true;
        }

        private static int ScoreTeamsVisibleTitleCandidate(string candidate, string prefix)
        {
            int score = 0;

            if (candidate.Length >= 8)
            {
                score += 10;
            }

            if (candidate.Length >= 16)
            {
                score += 5;
            }

            foreach (char c in candidate)
            {
                if (c >= 0x3000)
                {
                    score += 8;
                    break;
                }
            }

            if (candidate.IndexOf(" ", StringComparison.Ordinal) >= 0)
            {
                score += 2;
            }

            if (candidate.EndsWith(".com", StringComparison.OrdinalIgnoreCase)
                || candidate.IndexOf("http://", StringComparison.OrdinalIgnoreCase) >= 0
                || candidate.IndexOf("https://", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                score -= 50;
            }

            if (candidate.IndexOf(prefix, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                score += candidate.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) ? 30 : -10;
            }

            return score;
        }

        private sealed class TeamsWindowInfo
        {
            public TeamsWindowInfo(int handle, string title, string className, string processName, AutomationElement element)
            {
                Handle = handle;
                Title = title ?? string.Empty;
                ClassName = className ?? string.Empty;
                ProcessName = processName ?? string.Empty;
                Element = element;
            }

            public int Handle { get; private set; }
            public string Title { get; private set; }
            public string ClassName { get; private set; }
            public string ProcessName { get; private set; }
            public AutomationElement Element { get; private set; }
        }

        private static Dictionary<int, TeamsWindowInfo> GetTeamsWindows()
        {
            Dictionary<int, TeamsWindowInfo> windowsByHandle = new Dictionary<int, TeamsWindowInfo>();

            try
            {
                AutomationElementCollection windows = AutomationElement.RootElement.FindAll(
                    TreeScope.Children,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window));

                foreach (AutomationElement window in windows)
                {
                    string className = string.Empty;
                    string processName = string.Empty;
                    if (!IsTeamsAutomationElement(window, out className, out processName))
                    {
                        continue;
                    }

                    int handle = window.Current.NativeWindowHandle;
                    if (handle == 0 || windowsByHandle.ContainsKey(handle))
                    {
                        continue;
                    }

                    windowsByHandle.Add(handle, new TeamsWindowInfo(handle, window.Current.Name, className, processName, window));
                }
            }
            catch
            {
            }

            return windowsByHandle;
        }

        private static bool IsTeamsAutomationElement(AutomationElement element, out string className, out string processName)
        {
            className = string.Empty;
            processName = string.Empty;
            try
            {
                className = element.Current.ClassName ?? string.Empty;
                int processId = element.Current.ProcessId;
                try
                {
                    processName = Process.GetProcessById(processId).ProcessName;
                }
                catch
                {
                }

                return processName.IndexOf("Teams", StringComparison.OrdinalIgnoreCase) >= 0
                    || processName.IndexOf("ms-teams", StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch
            {
                return false;
            }
        }

        private static DialogResult ShowTopMostMessage(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            using (Form form = new Form())
            {
                form.Text = caption;
                form.StartPosition = FormStartPosition.CenterScreen;
                form.Width = 460;
                form.Height = 180;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MaximizeBox = false;
                form.MinimizeBox = false;
                form.ShowInTaskbar = true;
                form.TopMost = true;

                Label label = new Label();
                label.Left = 18;
                label.Top = 18;
                label.Width = 410;
                label.Height = 82;
                label.Text = text;
                label.AutoEllipsis = true;

                Button okButton = new Button();
                okButton.Text = "OK";
                okButton.Width = 90;
                okButton.Height = 28;
                okButton.Left = form.ClientSize.Width - okButton.Width - 18;
                okButton.Top = form.ClientSize.Height - okButton.Height - 18;
                okButton.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;
                okButton.DialogResult = DialogResult.OK;
                okButton.TabIndex = 0;

                form.Controls.Add(label);
                form.Controls.Add(okButton);
                form.AcceptButton = okButton;
                form.ActiveControl = okButton;
                form.Shown += delegate
                {
                    ScheduleDialogActivation(form, okButton);
                };

                return form.ShowDialog();
            }
        }

        private static void ScheduleDialogActivation(Form form, Control focusControl)
        {
            form.BeginInvoke(new Action(delegate
            {
                ActivateDialog(form, focusControl);
            }));

            Timer timer = new Timer();
            int attempts = 0;
            timer.Interval = 120;
            timer.Tick += delegate
            {
                attempts++;
                if (form.IsDisposed || attempts > 6)
                {
                    timer.Stop();
                    timer.Dispose();
                    return;
                }

                ActivateDialog(form, focusControl);
            };
            timer.Start();
        }

        private static void ActivateDialog(Form form, Control focusControl)
        {
            if (form == null || form.IsDisposed)
            {
                return;
            }

            try
            {
                IntPtr foregroundWindow = GetForegroundWindow();
                uint currentThreadId = GetCurrentThreadId();
                uint foregroundThreadId = foregroundWindow == IntPtr.Zero
                    ? 0
                    : GetWindowThreadProcessId(foregroundWindow, IntPtr.Zero);
                bool attached = false;

                if (form.WindowState == FormWindowState.Minimized)
                {
                    form.WindowState = FormWindowState.Normal;
                }

                try
                {
                    if (foregroundThreadId != 0 && foregroundThreadId != currentThreadId)
                    {
                        attached = AttachThreadInput(currentThreadId, foregroundThreadId, true);
                    }

                    form.TopMost = false;
                    form.TopMost = true;
                    form.Show();
                    form.Activate();
                    form.BringToFront();

                    if (form.IsHandleCreated)
                    {
                        ShowWindow(form.Handle, SwRestore);
                        BringWindowToTop(form.Handle);
                        SetActiveWindow(form.Handle);
                        SetForegroundWindow(form.Handle);
                    }

                    if (focusControl != null && focusControl.CanFocus)
                    {
                        form.ActiveControl = focusControl;
                        focusControl.Focus();
                        focusControl.Select();

                        if (focusControl.IsHandleCreated)
                        {
                            SetFocus(focusControl.Handle);
                        }
                    }
                }
                finally
                {
                    if (attached)
                    {
                        AttachThreadInput(currentThreadId, foregroundThreadId, false);
                    }
                }
            }
            catch
            {
            }
        }

        private static void CloseTeamsWindow(TeamsWindowInfo window)
        {
            if (window == null || window.Element == null)
            {
                return;
            }

            try
            {
                object pattern;
                if (window.Element.TryGetCurrentPattern(WindowPattern.Pattern, out pattern))
                {
                    ((WindowPattern)pattern).Close();
                }
            }
            catch
            {
            }
        }

        private static void MinimizeTeamsWindow(TeamsWindowInfo window)
        {
            if (window == null || window.Handle == 0)
            {
                return;
            }

            try
            {
                ShowWindow(new IntPtr(window.Handle), SwShowMinNoActive);
            }
            catch
            {
            }
        }

        private static string CleanTeamsBrowserTitleCandidate(string value)
        {
            string candidate = CleanTitleCandidate(value);
            if (string.IsNullOrWhiteSpace(candidate))
            {
                return string.Empty;
            }

            candidate = RepairUtf8AsShiftJisMojibake(candidate);
            candidate = RemoveTeamsJoinWindowSuffix(candidate);
            candidate = RemoveTeamsMojibakeJoinWindowSuffix(candidate);
            candidate = RemoveTeamsPipeAccountSuffix(candidate);
            candidate = Regex.Replace(candidate, "\\s*[\\-|\\u2013|\\u2014|\\uFF5C|\\|]\\s*Microsoft Teams\\s*$", string.Empty, RegexOptions.IgnoreCase).Trim();
            candidate = Regex.Replace(candidate, "\\s*Microsoft Teams\\s*$", string.Empty, RegexOptions.IgnoreCase).Trim();
            candidate = Regex.Replace(candidate, "^Microsoft Teams\\s*[\\-|\\u2013|\\u2014|\\uFF5C|\\|]\\s*", string.Empty, RegexOptions.IgnoreCase).Trim();
            candidate = Regex.Replace(candidate, "\\s*_\\s*[^_\\r\\n]{1,80}\\s*_\\s*[A-Z0-9._%+\\-]+@[A-Z0-9.\\-]+\\.[A-Z]{2,}\\s*$", string.Empty, RegexOptions.IgnoreCase).Trim();
            candidate = Regex.Replace(candidate, "\\s*_\\s*[A-Z0-9._%+\\-]+@[A-Z0-9.\\-]+\\.[A-Z]{2,}\\s*$", string.Empty, RegexOptions.IgnoreCase).Trim();
            candidate = RemoveTeamsAccountSuffix(candidate);
            candidate = RemoveTeamsJoinWindowSuffix(candidate);
            candidate = RemoveTeamsMojibakeJoinWindowSuffix(candidate);
            candidate = RemoveTeamsPipeAccountSuffix(candidate);
            return candidate;
        }

        private static string RemoveTeamsPipeAccountSuffix(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            string candidate = Regex.Replace(value, "[\\u200E\\u200F\\u202A-\\u202E]", string.Empty).Trim();
            candidate = Regex.Replace(candidate, "\\s*[\\|\\uFF5C]\\s*[^\\|\\uFF5C\\r\\n]{0,80}\\s*[\\|\\uFF5C]\\s*[^\\s\\|\\uFF5C@]+@[^\\s\\|\\uFF5C@]+\\s*(?:[\\|\\uFF5C]\\s*Microsoft Teams)?\\s*$", string.Empty, RegexOptions.IgnoreCase).Trim();
            candidate = Regex.Replace(candidate, "\\s*[\\|\\uFF5C]\\s*[^\\s\\|\\uFF5C@]+@[^\\s\\|\\uFF5C@]+\\s*(?:[\\|\\uFF5C]\\s*Microsoft Teams)?\\s*$", string.Empty, RegexOptions.IgnoreCase).Trim();
            return candidate;
        }

        internal static string RepairUtf8AsShiftJisMojibake(string value)
        {
            if (string.IsNullOrWhiteSpace(value) || !LooksLikeUtf8AsShiftJisMojibake(value))
            {
                return value;
            }

            try
            {
                string repaired = Encoding.UTF8.GetString(Encoding.GetEncoding(932).GetBytes(value));
                return LooksBetterRepairedTitle(value, repaired) ? repaired : value;
            }
            catch
            {
                return value;
            }
        }

        private static bool LooksLikeUtf8AsShiftJisMojibake(string value)
        {
            return Regex.IsMatch(value, "(?:\u7e5d|\u7e67|\u7e5e|\u7e5a|\u7e5f|\u7e5b|\u7e5c|\u83a8|\u8709|\u9a3e|\u8c3a|\u8b10|\u601c|\u58fb|\u87c7|\u9af1|\u8389|\u7fd5|\u9727|\u8b8c|\u870d)");
        }

        private static bool LooksBetterRepairedTitle(string original, string repaired)
        {
            if (string.IsNullOrWhiteSpace(repaired))
            {
                return false;
            }

            int originalMojibakeScore = CountMojibakeMarkers(original);
            int repairedMojibakeScore = CountMojibakeMarkers(repaired);
            return repairedMojibakeScore < originalMojibakeScore;
        }

        private static int CountMojibakeMarkers(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return 0;
            }

            int count = 0;
            foreach (Match match in Regex.Matches(value, "(?:\u7e5d|\u7e67|\u7e5e|\u7e5a|\u7e5f|\u7e5b|\u7e5c|\u83a8|\u8709|\u9a3e|\u8c3a|\u8b10|\u601c|\u58fb|\u87c7|\u9af1|\u8389|\u7fd5|\u9727|\u8b8c|\u870d)"))
            {
                count++;
            }

            return count;
        }

        private static string RemoveTeamsJoinWindowSuffix(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            string candidate = Regex.Replace(value, "[\\u200E\\u200F\\u202A-\\u202E]", string.Empty).Trim();
            candidate = Regex.Replace(candidate, "\\s*[_\\uFF3F]\\s*(?:\u4f1a\u8b70\u306b\u53c2\u52a0|Join meeting)\\s*[_\\uFF3F]\\s*Microsoft Teams\\s*$", string.Empty, RegexOptions.IgnoreCase).Trim();
            candidate = Regex.Replace(candidate, "\\s*[_\\uFF3F]\\s*(?:\u4f1a\u8b70\u306b\u53c2\u52a0|Join meeting)\\s*$", string.Empty, RegexOptions.IgnoreCase).Trim();
            candidate = Regex.Replace(candidate, "\\s*[_\\uFF3F]\\s*Microsoft Teams\\s*$", string.Empty, RegexOptions.IgnoreCase).Trim();
            return candidate;
        }

        private static string RemoveTeamsMojibakeJoinWindowSuffix(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            string candidate = Regex.Replace(value, "[\\u200E\\u200F\\u202A-\\u202E]", string.Empty).Trim();
            candidate = Regex.Replace(candidate, "\\s*[_\\uFF3F]\\s*莨夊ｭｰ縺ｫ蜿ょ刈\\s*[_\\uFF3F]\\s*Microsoft Teams\\s*$", string.Empty, RegexOptions.IgnoreCase).Trim();
            candidate = Regex.Replace(candidate, "\\s*[_\\uFF3F]\\s*莨夊ｭｰ縺ｫ蜿ょ刈\\s*$", string.Empty, RegexOptions.IgnoreCase).Trim();
            candidate = Regex.Replace(candidate, "\\s*[_\\uFF3F]\\s*Microsoft Teams\\s*$", string.Empty, RegexOptions.IgnoreCase).Trim();
            return candidate;
        }

        private static string RemoveTeamsAccountSuffix(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            string candidate = Regex.Replace(value, "[\\u200E\\u200F\\u202A-\\u202E]", string.Empty).Trim();
            candidate = Regex.Replace(candidate, "\\s*[\\x5F\\uFF3F]\\s*[^\\x5F\\uFF3F\\r\\n]{0,80}\\s*[\\x5F\\uFF3F]\\s*[^\\s\\x5F\\uFF3F@]+@[^\\s\\x5F\\uFF3F@]+\\s*$", string.Empty, RegexOptions.IgnoreCase).Trim();
            candidate = Regex.Replace(candidate, "\\s*[\\x5F\\uFF3F]\\s*[^\\s\\x5F\\uFF3F@]+@[^\\s\\x5F\\uFF3F@]+\\s*$", string.Empty, RegexOptions.IgnoreCase).Trim();
            return candidate;
        }

        private static bool IsUsableTeamsBrowserTitle(string value)
        {
            if (!IsUsableTitleCandidate(value))
            {
                return false;
            }

            string candidate = value.Trim();
            if (Regex.IsMatch(candidate, "^(?:Microsoft Teams|Microsoft Teams\\s+\\u4f1a\\u8b70|Teams|Teams\\s+\\u4f1a\\u8b70|M365|Office|Loading|\\u8aad\\u307f\\u8fbc\\u307f\\u4e2d|\\u30b5\\u30a4\\u30f3\\u30a4\\u30f3|\\u30ed\\u30b0\\u30a4\\u30f3|Sign in|Join meeting|\\u4f1a\\u8b70\\u306b\\u53c2\\u52a0|\\u53c2\\u52a0|\\u30c1\\u30e3\\u30c3\\u30c8|\\u4f1a\\u8b70|\\u4e88\\u5b9a\\u8868|Calendar|Chat|Activity)$", RegexOptions.IgnoreCase))
            {
                return false;
            }

            if (Regex.IsMatch(candidate, "^Microsoft Teams\\s+会議[_\\s-].+@.+", RegexOptions.IgnoreCase)
                || Regex.IsMatch(candidate, "^Teams\\s+会議[_\\s-].+@.+", RegexOptions.IgnoreCase)
                || Regex.IsMatch(candidate, "^会議[_\\s-].+@.+", RegexOptions.IgnoreCase)
                || Regex.IsMatch(candidate, "^[^\\s]+@[^\\s]+$"))
            {
                return false;
            }

            if (candidate.IndexOf("http://", StringComparison.OrdinalIgnoreCase) >= 0
                || candidate.IndexOf("https://", StringComparison.OrdinalIgnoreCase) >= 0
                || candidate.IndexOf("msteams:", StringComparison.OrdinalIgnoreCase) >= 0
                || candidate.IndexOf("teams.microsoft.com", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return false;
            }

            return true;
        }

        private static string GetShortcutNameFromUserInputOrFallback(bool quiet, string fallbackName, string failureReason)
        {
            if (quiet)
            {
                return fallbackName;
            }

            string prompt = "\u30ea\u30f3\u30af\u304b\u3089\u30d5\u30a1\u30a4\u30eb\u540d\u3092\u53d6\u5f97\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002";
            if (!string.IsNullOrWhiteSpace(failureReason))
            {
                prompt += "\r\n\r\n\u7406\u7531: " + failureReason;
            }

            prompt += "\r\n\r\n\u30b7\u30e7\u30fc\u30c8\u30ab\u30c3\u30c8\u540d\u3092\u5165\u529b\u3057\u3066\u304f\u3060\u3055\u3044\u3002";

            string input = ShowShortcutNameInputDialog(prompt, fallbackName);
            if (input == null)
            {
                throw new OperationCanceledException();
            }

            if (!string.IsNullOrWhiteSpace(input))
            {
                return input.Trim();
            }

            return fallbackName;
        }

        private static string ShowShortcutNameInputDialog(string prompt, string fallbackName)
        {
            using (Form form = new Form())
            using (Label label = new Label())
            using (TextBox textBox = new TextBox())
            using (Button okButton = new Button())
            using (Button cancelButton = new Button())
            {
                form.Text = AppTitle;
                form.StartPosition = FormStartPosition.CenterScreen;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MinimizeBox = false;
                form.MaximizeBox = false;
                form.ShowInTaskbar = true;
                form.TopMost = true;
                form.ClientSize = new System.Drawing.Size(430, 170);

                label.AutoSize = false;
                label.Left = 12;
                label.Top = 10;
                label.Width = 405;
                label.Height = 88;
                label.Text = prompt;

                textBox.Left = 12;
                textBox.Top = 105;
                textBox.Width = 405;
                textBox.Text = fallbackName;
                textBox.SelectAll();

                okButton.Text = "OK";
                okButton.Left = 242;
                okButton.Top = 136;
                okButton.Width = 82;
                okButton.DialogResult = DialogResult.OK;

                cancelButton.Text = "\u30ad\u30e3\u30f3\u30bb\u30eb";
                cancelButton.Left = 335;
                cancelButton.Top = 136;
                cancelButton.Width = 82;
                cancelButton.DialogResult = DialogResult.Cancel;

                form.Controls.Add(label);
                form.Controls.Add(textBox);
                form.Controls.Add(okButton);
                form.Controls.Add(cancelButton);
                form.AcceptButton = okButton;
                form.CancelButton = cancelButton;
                form.ActiveControl = textBox;
                form.Shown += delegate
                {
                    ScheduleDialogActivation(form, textBox);
                };

                return form.ShowDialog() == DialogResult.OK ? textBox.Text : null;
            }
        }

    }
}
