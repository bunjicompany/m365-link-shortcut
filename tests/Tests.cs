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
            AssertEqual(".pdf", Program.GetExtensionFromNameOrUrl("資料", "https://contoso.sharepoint.com/sites/site/Shared%20Documents/totori_family_trip_guide.pdf"), "Extension from URL path");
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
            TestDebugLogArguments();
            TestSharePointKindExtensions();
            TestSharePointFolderAndNameExtensions();

            Console.WriteLine(failed == 0 ? "ALL OK" : failed + " failed");
            return failed == 0 ? 0 : 1;
        }
    }
}
