using System;
using System.Text;

namespace SharePointShortcutMaker
{
    internal static class TestRunner
    {
        private static int failed;

        private static void AssertEqual(string expected, string actual, string name)
        {
            if (!string.Equals(expected, actual, StringComparison.Ordinal))
            {
                Console.WriteLine("NG " + name + ": expected=[" + expected + "] actual=[" + actual + "]");
                failed++;
            }
        }

        private static void AssertTrue(bool condition, string name)
        {
            if (!condition)
            {
                Console.WriteLine("NG " + name + ": expected true");
                failed++;
            }
        }

        private static void AssertFalse(bool condition, string name)
        {
            if (condition)
            {
                Console.WriteLine("NG " + name + ": expected false");
                failed++;
            }
        }

        private static void TestCleanUrlCandidate()
        {
            AssertEqual("https://x.sharepoint.com/a?b=1", Program.CleanUrlCandidate("<https://x.sharepoint.com/a?b=1>"), "CleanUrlCandidate trims brackets");
            AssertEqual("https://x.sharepoint.com/a?b=1", Program.CleanUrlCandidate("https://x.sharepoint.com/a?b=1."), "CleanUrlCandidate trims trailing punctuation");
        }

        private static void TestTeamsLinkKinds()
        {
            AssertTrue(Program.IsTeamsLink("https://teams.microsoft.com/l/message/19:abc@thread.v2/123"), "IsTeamsLink message");
            AssertTrue(Program.IsTeamsChatLink("https://teams.microsoft.com/l/message/19:abc@thread.v2/123"), "IsTeamsChatLink message");
            AssertTrue(Program.IsTeamsMeetingLink("https://teams.microsoft.com/l/meetup-join/19:meeting_abc@thread.v2/0"), "IsTeamsMeetingLink meetup");
            AssertFalse(Program.IsTeamsChatLink("https://contoso.sharepoint.com/:x:/s/site/id"), "IsTeamsChatLink sharepoint false");
        }

        private static void TestTeamsChannelName()
        {
            AssertEqual("一般", Program.GetTeamsChannelNameFromUrl("https://teams.microsoft.com/l/channel/19%3Aabc%40thread.tacv2/%E4%B8%80%E8%88%AC?groupId=xxx&tenantId=yyy"), "GetTeamsChannelNameFromUrl Japanese");
            AssertEqual("General", Program.GetTeamsChannelNameFromUrl("https://teams.microsoft.com/channel/19:abc@thread.tacv2/General?groupId=xxx"), "GetTeamsChannelNameFromUrl plain");
        }

        private static void TestTeamsTimestamp()
        {
            DateTime parsed;
            AssertTrue(Program.TryParseTeamsChatTimestamp("2026年7月3日 17:41.", out parsed), "TryParseTeamsChatTimestamp absolute");
            AssertEqual("2026年7月3日(金)17時41分", Program.FormatDateTimeForJapaneseTitle(parsed), "FormatDateTimeForJapaneseTitle");
        }

        private static void TestTeamsMessageIdTimestamp()
        {
            DateTime parsed;
            AssertTrue(Program.TryGetTeamsChatTimestampFromMessageId("1783058499225", out parsed), "TryGetTeamsChatTimestampFromMessageId");
            AssertEqual("2026年7月3日(金)15時1分", Program.FormatDateTimeForJapaneseTitle(parsed), "FormatDateTimeForJapaneseTitle from Teams message id");
        }

        private static void TestTeamsMeetingDateFallbackName()
        {
            AssertEqual("Teams会議 【7月3日(金)16時00分-17時00分】", Program.GetTeamsMeetingDateFallbackName("7月3日金曜日 16:00–17:00JST"), "GetTeamsMeetingDateFallbackName long weekday");
            AssertEqual("Teams会議 【7月3日(金)16時00分-17時00分】", Program.GetTeamsMeetingDateFallbackName("7月3日金 16:00-17:00"), "GetTeamsMeetingDateFallbackName short weekday");
        }

        private static void TestMojibakeRepair()
        {
            string original = "会議に参加";
            string mojibake = "莨夊ｭｰ縺ｫ蜿ょ刈";
            AssertEqual(original, Program.RepairUtf8AsShiftJisMojibake(mojibake), "RepairUtf8AsShiftJisMojibake");
            AssertEqual(original, Program.RepairUtf8AsShiftJisMojibake(original), "RepairUtf8AsShiftJisMojibake keeps normal text");
        }

        private static void TestSafeFileName()
        {
            AssertEqual("abc_def", Program.GetSafeFileName("abc:def"), "GetSafeFileName invalid char");
            AssertEqual("M365リンク", Program.GetSafeFileName("   "), "GetSafeFileName blank fallback");
            AssertEqual("abc", Program.GetSafeFileName("abc.."), "GetSafeFileName trailing dots");
        }

        private static void TestReservedWindowsFileName()
        {
            AssertTrue(Program.IsReservedWindowsFileName("CON"), "IsReservedWindowsFileName CON");
            AssertTrue(Program.IsReservedWindowsFileName("con.xlsx"), "IsReservedWindowsFileName con.xlsx");
            AssertTrue(Program.IsReservedWindowsFileName("COM1"), "IsReservedWindowsFileName COM1");
            AssertFalse(Program.IsReservedWindowsFileName("CONFIG"), "IsReservedWindowsFileName CONFIG is not reserved");
            AssertFalse(Program.IsReservedWindowsFileName("COM10"), "IsReservedWindowsFileName COM10 is not reserved");
            AssertEqual("CON_", Program.GetSafeFileName("CON"), "GetSafeFileName reserved name");
            AssertEqual("NUL_.xlsx", Program.GetSafeFileName("NUL.xlsx"), "GetSafeFileName reserved name with extension");
            AssertEqual("週次定例.docx", Program.GetSafeFileName("週次定例.docx"), "GetSafeFileName normal name unchanged");
        }

        private static void TestBaseNameTruncation()
        {
            AssertEqual("abcde", Program.TruncateBaseName("abcde", 10), "TruncateBaseName short name unchanged");
            AssertEqual("aaaaaaaaaa", Program.TruncateBaseName(new string('a', 30), 10), "TruncateBaseName plain cut");

            string longStemWithExtension = new string('a', 200) + ".xlsx";
            string truncatedWithExtension = Program.TruncateBaseName(longStemWithExtension, 50);
            AssertTrue(truncatedWithExtension.Length == 50, "TruncateBaseName keeps max length");
            AssertTrue(truncatedWithExtension.EndsWith(".xlsx", StringComparison.Ordinal), "TruncateBaseName keeps extension");

            string surrogateName = new string('a', 9) + "𠮷";
            AssertEqual(new string('a', 9), Program.TruncateBaseName(surrogateName, 10), "TruncateBaseName does not split surrogate pair");

            string longName = new string('あ', 300);
            string budgeted = Program.TruncateBaseNameForPath("C:\\", longName, ".url");
            AssertTrue(budgeted.Length <= 120, "TruncateBaseNameForPath caps at 120");
        }

        private static void TestTitleCandidateMojibakeGuard()
        {
            AssertFalse(Program.IsUsableTitleCandidate("abc�def"), "IsUsableTitleCandidate rejects replacement char");
            AssertFalse(Program.IsUsableTitleCandidate("莨夊ｭｰ縺ｫ蜿ょ刈"), "IsUsableTitleCandidate rejects mojibake");
            AssertTrue(Program.IsUsableTitleCandidate("週次定例メモ"), "IsUsableTitleCandidate accepts normal text");
        }

        private static void TestTeamsChatDisplayPrefix()
        {
            AssertEqual("Teams投稿", Program.GetTeamsChatDisplayPrefix("https://teams.microsoft.com/l/message/19:abc@thread.tacv2/123?groupId=00000000-0000-0000-0000-000000000007&teamName=SampleTeam&channelName=General"), "GetTeamsChatDisplayPrefix channel post");
            AssertEqual("Teams会議チャット", Program.GetTeamsChatDisplayPrefix("https://teams.microsoft.com/l/message/19:meeting_QWJjZGVm@thread.v2/1783058499225?tenantId=00000000-0000-0000-0000-000000000008"), "GetTeamsChatDisplayPrefix meeting chat");
            AssertEqual("Teamsチャット", Program.GetTeamsChatDisplayPrefix("https://teams.microsoft.com/l/message/19:abc@thread.v2/1783058499225"), "GetTeamsChatDisplayPrefix plain chat");
            AssertTrue(Program.IsTeamsMeetingChatLink("https://teams.microsoft.com/l/message/19%3Ameeting_QWJjZGVm%40thread.v2/1783058499225"), "IsTeamsMeetingChatLink url encoded");
        }

        private static void TestTeamsTeamLink()
        {
            string teamUrl = "https://teams.microsoft.com/l/team/19%3A00000000000000000000000000000000%40thread.tacv2/conversations?groupId=00000000-0000-0000-0000-000000000009&tenantId=00000000-0000-0000-0000-00000000000a";
            AssertTrue(Program.IsTeamsTeamLink(teamUrl), "IsTeamsTeamLink team url");
            AssertTrue(Program.IsTeamsChatLink(teamUrl), "IsTeamsChatLink accepts team url");
            AssertEqual("Teams投稿", Program.GetTeamsChatDisplayPrefix(teamUrl), "GetTeamsChatDisplayPrefix team url");
            AssertFalse(Program.IsTeamsTeamLink("https://teams.microsoft.com/l/channel/19:abc@thread.tacv2/General?groupId=x"), "IsTeamsTeamLink channel url false");
        }

        private static void TestPreLaunchChatTitleMatch()
        {
            System.Collections.Generic.List<string> beforeTitles = new System.Collections.Generic.List<string>();
            beforeTitles.Add("サンプル 太郎 Taro Sample");
            AssertTrue(Program.MatchesPreLaunchChatTitle(beforeTitles, "サンプル 太郎 Taro Sample"), "MatchesPreLaunchChatTitle same chat");
            AssertFalse(Program.MatchesPreLaunchChatTitle(beforeTitles, "別のチャット名"), "MatchesPreLaunchChatTitle different chat");
        }

        private static void TestTeamsTargetBodyPrefixSpeaker()
        {
            Program.TeamsChatInfo chatInfo = new Program.TeamsChatInfo("1782374400529", true, "00000000-0000-0000-0000-000000000005", "サンプルチーム", "サンプルチャネル");
            Program.UpdateTeamsChatInfo(chatInfo, "発言者 一郎 サンプル件名 本文のテキストが続きます。", "", "ControlType.Group", "グループ", "message-body-1782374400529", "", "");
            AssertTrue(chatInfo.HasTargetBody, "target body detected");

            Program.UpdateTeamsChatInfo(chatInfo, "発言者 一郎", "", "ControlType.Text", "テキスト", "", "", "");
            AssertEqual("発言者 一郎", chatInfo.Speaker, "target body prefix speaker adopted");

            Program.TeamsChatInfo otherInfo = new Program.TeamsChatInfo("1782374400529", true, "00000000-0000-0000-0000-000000000006", "サンプルチーム", "サンプルチャネル");
            Program.UpdateTeamsChatInfo(otherInfo, "発言者 一郎 サンプル件名 本文のテキストが続きます。", "", "ControlType.Group", "グループ", "message-body-1782374400529", "", "");
            Program.UpdateTeamsChatInfo(otherInfo, "無関係なテキスト要素", "", "ControlType.Text", "テキスト", "", "", "");
            AssertTrue(string.IsNullOrEmpty(otherInfo.Speaker), "non prefix text is not adopted as speaker");
        }

        private static void TestTeamsBrowserTitleUsability()
        {
            AssertFalse(Program.IsUsableTeamsBrowserTitle("チームとチャネル"), "IsUsableTeamsBrowserTitle rejects nav view title");
            AssertFalse(Program.IsUsableTeamsBrowserTitle("チームとチャネル | サンプルチャネル"), "IsUsableTeamsBrowserTitle rejects nav view prefix");
            AssertFalse(Program.IsUsableTeamsBrowserTitle("アクティビティ"), "IsUsableTeamsBrowserTitle rejects activity view");
            AssertTrue(Program.IsUsableTeamsBrowserTitle("一般連絡"), "IsUsableTeamsBrowserTitle accepts normal channel name");
        }

        private static void TestTeamsChannelUrlFallbackTitle()
        {
            string channelMessageUrl = "https://teams.microsoft.com/l/message/19:00000000000000000000000000000000@thread.skype/1782439160187?tenantId=00000000-0000-0000-0000-000000000001&groupId=00000000-0000-0000-0000-000000000002&parentMessageId=1782115880684&teamName=%E3%82%B5%E3%83%B3%E3%83%97%E3%83%AB%E3%83%81%E3%83%BC%E3%83%A0&channelName=%E9%80%A3%E7%B5%A1%E7%94%A8&createdTime=1782439160187";
            AssertEqual(
                "Teams投稿 【サンプルチーム - 連絡用】 2026年6月26日(金)10時59分",
                Program.BuildTeamsChannelUrlFallbackTitle(channelMessageUrl),
                "BuildTeamsChannelUrlFallbackTitle channel message with timestamp");

            string channelMessageUrl2 = "https://teams.microsoft.com/l/message/19:11111111111111111111111111111111@thread.tacv2/1782374400529?tenantId=00000000-0000-0000-0000-000000000003&groupId=00000000-0000-0000-0000-000000000004&parentMessageId=1782374400529&teamName=Sample%20AI%20Team&channelName=Weekly%20Report&createdTime=1782374400529";
            string fallbackTitle2 = Program.BuildTeamsChannelUrlFallbackTitle(channelMessageUrl2);
            AssertTrue(fallbackTitle2.StartsWith("Teams投稿 【Sample AI Team - Weekly Report】", StringComparison.Ordinal), "BuildTeamsChannelUrlFallbackTitle tacv2 channel message");

            AssertEqual(string.Empty, Program.BuildTeamsChannelUrlFallbackTitle("https://contoso.sharepoint.com/:x:/s/site/id"), "BuildTeamsChannelUrlFallbackTitle non teams url");
        }

        private static void TestDebugLogArguments()
        {
            AssertTrue(Program.IsTeamsDebugLogArg("--debug"), "IsTeamsDebugLogArg debug");
            AssertTrue(Program.IsTeamsDebugLogArg("--debug-log"), "IsTeamsDebugLogArg debug-log");
            AssertFalse(Program.IsTeamsDebugLogArg("--install"), "IsTeamsDebugLogArg install false");
            AssertEqual("\"C:\\App\\M365LinkShortcut.exe\" \"%V\"", Program.BuildContextMenuCommand("C:\\App\\M365LinkShortcut.exe", "%V", false), "BuildContextMenuCommand default");
            AssertEqual("\"C:\\App\\M365LinkShortcut.exe\" --debug \"%V\"", Program.BuildContextMenuCommand("C:\\App\\M365LinkShortcut.exe", "%V", true), "BuildContextMenuCommand debug");
        }

        private static void TestSharePointKindExtensions()
        {
            AssertEqual(".xlsx", Program.GetExtensionFromOfficeOrSharePointKind("https://contoso.sharepoint.com/:x:/s/site/id?e=abc"), "SharePoint xlsx kind");
            AssertEqual(".docx", Program.GetExtensionFromOfficeOrSharePointKind("https://contoso.sharepoint.com/:w:/s/site/id?e=abc"), "SharePoint docx kind");
            AssertEqual(".pptx", Program.GetExtensionFromOfficeOrSharePointKind("https://contoso.sharepoint.com/:p:/s/site/id?e=abc"), "SharePoint pptx kind");
            AssertEqual(".txt", Program.GetExtensionFromOfficeOrSharePointKind("https://contoso.sharepoint.com/:t:/s/site/id?e=abc"), "SharePoint txt kind");
            AssertEqual(".png", Program.GetExtensionFromOfficeOrSharePointKind("https://contoso.sharepoint.com/:i:/s/site/id?e=abc"), "SharePoint image kind");
            AssertEqual(".xlsx", Program.GetExtensionFromOfficeOrSharePointKind("ms-excel:ofe|u|https://contoso.sharepoint.com/:x:/s/site/id"), "Office scheme xlsx kind");
        }

        private static void TestSharePointFolderAndNameExtensions()
        {
            AssertTrue(Program.IsFolderSharePointUrl("https://contoso.sharepoint.com/:f:/s/site/id?e=abc"), "SharePoint folder kind");
            AssertFalse(Program.IsFolderSharePointUrl("https://contoso.sharepoint.com/:x:/s/site/id?e=abc"), "SharePoint file kind is not folder");
            AssertEqual(".xlsx", Program.GetExtensionFromNameOrUrl("テスト結果.xlsx", "https://contoso.sharepoint.com/:x:/s/site/id"), "Extension from base name xlsx");
            AssertEqual(".jpg", Program.GetExtensionFromNameOrUrl("写真.jpg", "https://contoso.sharepoint.com/:i:/s/site/id"), "Extension from base name jpg");
            AssertEqual(".bat", Program.GetExtensionFromNameOrUrl("サンプルバックアップ.bat", "https://contoso.sharepoint.com/:u:/s/site/id"), "Extension from base name bat");
            AssertEqual(".pdf", Program.GetExtensionFromNameOrUrl("資料", "https://contoso.sharepoint.com/sites/site/Shared%20Documents/sample_guide.pdf"), "Extension from URL path");
        }

        internal static int Main()
        {
            TestCleanUrlCandidate();
            TestTeamsLinkKinds();
            TestTeamsChannelName();
            TestTeamsTimestamp();
            TestTeamsMessageIdTimestamp();
            TestTeamsMeetingDateFallbackName();
            TestMojibakeRepair();
            TestSafeFileName();
            TestReservedWindowsFileName();
            TestBaseNameTruncation();
            TestTitleCandidateMojibakeGuard();
            TestTeamsChatDisplayPrefix();
            TestTeamsTeamLink();
            TestPreLaunchChatTitleMatch();
            TestTeamsTargetBodyPrefixSpeaker();
            TestTeamsBrowserTitleUsability();
            TestTeamsChannelUrlFallbackTitle();
            TestDebugLogArguments();
            TestSharePointKindExtensions();
            TestSharePointFolderAndNameExtensions();

            Console.WriteLine(failed == 0 ? "ALL OK" : failed + " failed");
            return failed == 0 ? 0 : 1;
        }
    }
}
