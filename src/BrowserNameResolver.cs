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
        private static BrowserNameResult GetNameFromBrowser(string url)
        {
            try
            {
                using (BrowserNameResolver resolver = new BrowserNameResolver(url))
                {
                    BrowserNameResult result = resolver.ResolveName();
                    if (!resolver.NeedsInteractiveLogin)
                    {
                        return result;
                    }
                }

                using (BrowserNameResolver resolver = new BrowserNameResolver(url, true))
                {
                    return resolver.ResolveName();
                }
            }
            catch (Exception ex)
            {
                WriteAppDebugLog("GetNameFromBrowser", ex.ToString());
                return new BrowserNameResult(string.Empty, "\u4e88\u671f\u3057\u306a\u3044\u30a8\u30e9\u30fc: " + ex.Message);
            }
        }

        private sealed class BrowserNameResolver : Form
        {
            private readonly string url;
            private readonly WebView2 webView;
            private readonly Label message;
            private readonly Timer timeoutTimer;
            private readonly Timer pollTimer;
            private DateTime deadline;
            private bool resolving;
            private readonly bool interactiveLogin;
            private bool needsInteractiveLogin;
            private string resolvedName = string.Empty;
            private string failureReason = string.Empty;
            private string lastCandidates = string.Empty;

            public BrowserNameResolver(string url)
                : this(url, false)
            {
            }

            public BrowserNameResolver(string url, bool interactiveLogin)
            {
                this.url = url;
                this.interactiveLogin = interactiveLogin;
                this.deadline = interactiveLogin ? DateTime.Now.AddMinutes(10) : DateTime.Now.AddSeconds(8);
                this.Text = AppTitle;
                this.Width = 1280;
                this.Height = 900;
                this.StartPosition = FormStartPosition.Manual;
                this.MinimizeBox = false;
                this.MaximizeBox = false;

                if (interactiveLogin)
                {
                    ConfigureInteractiveLoginWindow();
                }
                else
                {
                    this.Location = new System.Drawing.Point(-32000, -32000);
                    this.FormBorderStyle = FormBorderStyle.None;
                    this.ShowInTaskbar = false;
                    this.Opacity = 0;
                }

                this.message = new Label();
                this.message.Dock = DockStyle.Top;
                this.message.Height = 44;
                this.message.Padding = new Padding(10, 8, 10, 4);
                this.message.Text = interactiveLogin
                    ? "Microsoft 365\u306b\u30ed\u30b0\u30a4\u30f3\u3057\u3066\u304f\u3060\u3055\u3044\u3002\u30ed\u30b0\u30a4\u30f3\u5f8c\u3001\u540d\u524d\u3092\u53d6\u5f97\u3067\u304d\u305f\u3089\u81ea\u52d5\u3067\u9589\u3058\u307e\u3059\u3002"
                    : "\u30ea\u30f3\u30af\u5148\u304b\u3089\u30d5\u30a1\u30a4\u30eb\u540d\u3092\u53d6\u5f97\u3057\u3066\u3044\u307e\u3059\u3002";

                this.webView = new WebView2();
                this.webView.Dock = DockStyle.Fill;

                this.timeoutTimer = new Timer();
                this.timeoutTimer.Interval = 1000;
                this.timeoutTimer.Tick += TimeoutTimerTick;

                this.pollTimer = new Timer();
                this.pollTimer.Interval = 1200;
                this.pollTimer.Tick += PollTimerTick;

                this.Controls.Add(this.webView);
                this.Controls.Add(this.message);
                this.Shown += BrowserNameResolverShown;
            }

            public bool NeedsInteractiveLogin
            {
                get { return this.needsInteractiveLogin; }
            }

            public BrowserNameResult ResolveName()
            {
                this.ShowDialog();
                if (!string.IsNullOrWhiteSpace(this.resolvedName))
                {
                    return new BrowserNameResult(this.resolvedName, string.Empty);
                }

                if (string.IsNullOrWhiteSpace(this.failureReason))
                {
                    this.failureReason = "\u540d\u524d\u5019\u88dc\u3092\u53d6\u5f97\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002";
                }

                return new BrowserNameResult(string.Empty, this.failureReason);
            }

            protected override bool ShowWithoutActivation
            {
                get { return !this.interactiveLogin; }
            }

            private async void BrowserNameResolverShown(object sender, EventArgs e)
            {
                try
                {
                    string userDataFolder = GetWebView2UserDataFolder();
                    CoreWebView2EnvironmentOptions options = CreateWebView2EnvironmentOptions();
                    CoreWebView2Environment environment = await CoreWebView2Environment.CreateAsync(null, userDataFolder, options);
                    await this.webView.EnsureCoreWebView2Async(environment);
                    this.webView.CoreWebView2.NavigationCompleted += WebViewNavigationCompleted;
                    this.timeoutTimer.Start();
                    this.pollTimer.Start();
                    this.webView.CoreWebView2.Navigate(this.url);
                    if (this.interactiveLogin)
                    {
                        this.BringToFront();
                        this.Activate();
                        this.webView.Focus();
                    }
                }
                catch (Exception ex)
                {
                    WriteAppDebugLog("BrowserNameResolverShown", ex.ToString());
                    this.failureReason = "WebView2\u306e\u521d\u671f\u5316\u306b\u5931\u6557\u3057\u307e\u3057\u305f: " + ex.Message;
                    this.Close();
                }
            }

            private static CoreWebView2EnvironmentOptions CreateWebView2EnvironmentOptions()
            {
                CoreWebView2EnvironmentOptions options = new CoreWebView2EnvironmentOptions();

                try
                {
                    PropertyInfo ssoProperty = typeof(CoreWebView2EnvironmentOptions).GetProperty("AllowSingleSignOnUsingOSPrimaryAccount");
                    if (ssoProperty != null && ssoProperty.CanWrite)
                    {
                        ssoProperty.SetValue(options, true, null);
                    }
                }
                catch (Exception ex)
                {
                    WriteAppDebugLog("CreateWebView2EnvironmentOptions", ex);
                }

                return options;
            }

            private static string GetWebView2UserDataFolder()
            {
                string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                if (string.IsNullOrWhiteSpace(localAppData))
                {
                    localAppData = Path.GetTempPath();
                }

                string folder = Path.Combine(localAppData, "M365LinkShortcut", "WebView2");
                Directory.CreateDirectory(folder);
                return folder;
            }

            private void WebViewNavigationCompleted(object sender, CoreWebView2NavigationCompletedEventArgs e)
            {
                TryResolveFromPage();
            }

            private void TimeoutTimerTick(object sender, EventArgs e)
            {
                if (DateTime.Now >= this.deadline)
                {
                    this.failureReason = BuildTimeoutReason();
                    this.Close();
                    return;
                }

                TryResolveFromPage();
            }

            private void PollTimerTick(object sender, EventArgs e)
            {
                TryResolveFromPage();
            }

            private async void TryResolveFromPage()
            {
                if (this.resolving || this.webView.CoreWebView2 == null)
                {
                    return;
                }

                this.resolving = true;
                try
                {
                    string script =
                        "(function(){" +
                        "function v(x){return x==null?'':String(x).trim();}" +
                        "var c=[];" +
                        "function add(x){x=v(x);if(x&&c.indexOf(x)<0)c.push(x);}" +
                        "function addQuery(name){try{add(new URL(location.href).searchParams.get(name));}catch(e){}}" +
                        "var login=false;" +
                        "try{" +
                        "var h=location.hostname.toLowerCase();" +
                        "var ttl=v(document.title).toLowerCase();" +
                        "login=h.indexOf('login.microsoftonline.com')>=0||h.indexOf('login.live.com')>=0||h.indexOf('login.windows.net')>=0||" +
                        "ttl.indexOf('sign in')>=0||ttl.indexOf('サインイン')>=0||ttl.indexOf('ログイン')>=0||" +
                        "document.querySelector('input[type=email],input[name=loginfmt],input[name=passwd],input[type=password]')!=null;" +
                        "}catch(e){}" +
                        "if(login)return 'LOGIN_REQUIRED\\n'+document.title+'\\n'+location.href;" +
                        "add(document.title);" +
                        "add(location.href);" +
                        "addQuery('id');addQuery('file');addQuery('filename');addQuery('fileName');addQuery('name');" +
                        "var metas=document.querySelectorAll('meta[property=\"og:title\"],meta[name=\"twitter:title\"],meta[name=\"title\"]');" +
                        "for(var i=0;i<metas.length;i++)add(metas[i].getAttribute('content'));" +
                        "var attrs=document.querySelectorAll('[title],[aria-label],[data-title],[data-name],[data-item-name],[data-automationid]');" +
                        "for(var j=0;j<attrs.length&&j<300;j++){" +
                        "add(attrs[j].getAttribute('title'));" +
                        "add(attrs[j].getAttribute('aria-label'));" +
                        "add(attrs[j].getAttribute('data-title'));" +
                        "add(attrs[j].getAttribute('data-name'));" +
                        "add(attrs[j].getAttribute('data-item-name'));" +
                        "}" +
                        "var textNodes=document.querySelectorAll('h1,h2,h3,[role=\"heading\"],button,span,div,a');" +
                        "var fileLike=/[^\\\\/:*?\"<>|\\r\\n]{1,180}\\.(?:xlsx|xlsm|xlsb|xls|docx|docm|doc|pptx|pptm|ppt|txt|csv|pdf|png|jpg|jpeg|gif|bmp|bat|cmd|ps1|vbs|js|json|xml|zip|vsdx|vsd)\\b/i;" +
                        "for(var k=0;k<textNodes.length&&k<600;k++){" +
                        "var t=v(textNodes[k].textContent);" +
                        "if(t&&t.length<=160&&(textNodes[k].matches('h1,h2,h3,[role=\"heading\"]')))add(t);" +
                        "if(t&&t.length<=220&&fileLike.test(t))add(t);" +
                        "}" +
                        "return c.join('\\n');" +
                        "})();";

                    string raw = await this.webView.CoreWebView2.ExecuteScriptAsync(script);
                    string candidates = DecodeJavaScriptString(raw);
                    this.lastCandidates = candidates;
                    if (IsLoginRequiredBrowserResult(candidates))
                    {
                        this.failureReason = "Microsoft\u306e\u30ed\u30b0\u30a4\u30f3\u753b\u9762\u307e\u305f\u306f\u8a8d\u8a3c\u5165\u529b\u6b04\u304c\u691c\u51fa\u3055\u308c\u307e\u3057\u305f\u3002SSO\u304c\u52b9\u3044\u3066\u3044\u306a\u3044\u53ef\u80fd\u6027\u304c\u3042\u308a\u307e\u3059\u3002";
                        if (!this.interactiveLogin)
                        {
                            this.needsInteractiveLogin = true;
                            this.Close();
                            return;
                        }

                        this.deadline = DateTime.Now.AddMinutes(10);
                        return;
                    }

                    string name = GetNameFromBrowserCandidates(candidates, this.url);
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        this.resolvedName = name;
                        this.Close();
                    }
                    else if (!string.IsNullOrWhiteSpace(candidates))
                    {
                        this.failureReason = "\u30da\u30fc\u30b8\u5019\u88dc\u306f\u53d6\u5f97\u3067\u304d\u307e\u3057\u305f\u304c\u3001\u30d5\u30a1\u30a4\u30eb\u540d/\u30d5\u30a9\u30eb\u30c0\u540d\u3068\u3057\u3066\u63a1\u7528\u3067\u304d\u308b\u5019\u88dc\u304c\u3042\u308a\u307e\u305b\u3093\u3067\u3057\u305f\u3002\u5019\u88dc: " + TruncateForDisplay(candidates);
                    }
                }
                catch (Exception ex)
                {
                    WriteAppDebugLog("TryResolveFromPage", ex.ToString());
                    this.failureReason = "\u30da\u30fc\u30b8\u5019\u88dc\u306e\u53d6\u5f97\u4e2d\u306b\u30a8\u30e9\u30fc\u304c\u767a\u751f\u3057\u307e\u3057\u305f: " + ex.Message;
                }
                finally
                {
                    this.resolving = false;
                }
            }

            private string BuildTimeoutReason()
            {
                if (this.interactiveLogin)
                {
                    if (string.IsNullOrWhiteSpace(this.lastCandidates)
                        || IsLoginRequiredBrowserResult(this.lastCandidates))
                    {
                        return "Microsoft\u306e\u30ed\u30b0\u30a4\u30f3\u753b\u9762\u3092\u8868\u793a\u3057\u307e\u3057\u305f\u304c\u3001\u5236\u9650\u6642\u9593\u5185\u306b\u8a8d\u8a3c\u304c\u5b8c\u4e86\u305b\u305a\u3001\u30d5\u30a1\u30a4\u30eb\u540d/\u30d5\u30a9\u30eb\u30c0\u540d\u3092\u53d6\u5f97\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002";
                    }

                    return "\u30ed\u30b0\u30a4\u30f3\u5f8c\u306b\u30da\u30fc\u30b8\u5019\u88dc\u306f\u53d6\u5f97\u3067\u304d\u307e\u3057\u305f\u304c\u3001\u63a1\u7528\u3067\u304d\u308b\u540d\u524d\u3092\u78ba\u5b9a\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002\u5019\u88dc: " + TruncateForDisplay(this.lastCandidates);
                }

                if (string.IsNullOrWhiteSpace(this.lastCandidates))
                {
                    return "WebView2\u3067\u30da\u30fc\u30b8\u3092\u958b\u304d\u307e\u3057\u305f\u304c\u30018\u79d2\u4ee5\u5185\u306b\u540d\u524d\u5019\u88dc\u3092\u53d6\u5f97\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002\u8a8d\u8a3c\u4e2d\u3001\u8aad\u307f\u8fbc\u307f\u9045\u5ef6\u3001\u307e\u305f\u306f\u30da\u30fc\u30b8\u69cb\u9020\u5909\u66f4\u306e\u53ef\u80fd\u6027\u304c\u3042\u308a\u307e\u3059\u3002";
                }

                return "\u540d\u524d\u5019\u88dc\u306f\u53d6\u5f97\u3067\u304d\u307e\u3057\u305f\u304c\u30018\u79d2\u4ee5\u5185\u306b\u63a1\u7528\u3067\u304d\u308b\u540d\u524d\u3092\u78ba\u5b9a\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002\u5019\u88dc: " + TruncateForDisplay(this.lastCandidates);
            }

            private void ConfigureInteractiveLoginWindow()
            {
                this.FormBorderStyle = FormBorderStyle.Sizable;
                this.ShowInTaskbar = true;
                this.Opacity = 1;
                this.WindowState = FormWindowState.Normal;
                this.StartPosition = FormStartPosition.Manual;

                System.Drawing.Rectangle area = Screen.PrimaryScreen.WorkingArea;
                this.Width = Math.Min(1100, Math.Max(640, area.Width - 80));
                this.Height = Math.Min(800, Math.Max(480, area.Height - 80));
                this.Left = area.Left + ((area.Width - this.Width) / 2);
                this.Top = area.Top + ((area.Height - this.Height) / 2);

                if (this.WindowState == FormWindowState.Minimized)
                {
                    this.WindowState = FormWindowState.Normal;
                }
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    this.timeoutTimer.Dispose();
                    this.pollTimer.Dispose();
                    this.webView.Dispose();
                    this.message.Dispose();
                }

                base.Dispose(disposing);
            }
        }

        private static string DecodeJavaScriptString(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                return string.Empty;
            }

            try
            {
                return new JavaScriptSerializer().Deserialize<string>(raw);
            }
            catch
            {
                return raw.Trim('"');
            }
        }

        private static string TruncateForDisplay(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            string normalized = value.Replace("\r", " ").Replace("\n", " / ").Trim();
            normalized = Regex.Replace(normalized, "\\s+", " ");
            return normalized.Length <= 300 ? normalized : normalized.Substring(0, 300) + "...";
        }

        private static string GetNameFromBrowserCandidates(string candidates, string url)
        {
            if (string.IsNullOrWhiteSpace(candidates))
            {
                return string.Empty;
            }

            bool isFolderLink = IsFolderSharePointUrl(url);
            string inferredExtension = GetExtensionFromOfficeOrSharePointKind(url);

            if (IsTeamsLink(url))
            {
                return GetTeamsNameFromBrowserCandidates(candidates, url);
            }

            if (isFolderLink)
            {
                string folderNameFromUrl = GetFolderNameFromBrowserUrlCandidates(candidates);
                if (!string.IsNullOrWhiteSpace(folderNameFromUrl))
                {
                    return folderNameFromUrl;
                }

                string[] rawLines = candidates.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                if (rawLines.Length > 0)
                {
                    string folderName = CleanResolvedFolderTitle(rawLines[0]);
                    if (IsUsableBrowserResolvedName(folderName, true, string.Empty))
                    {
                        return folderName;
                    }
                }
            }

            foreach (string rawLine in candidates.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries))
            {
                string line = CleanResolvedPageTitle(rawLine);
                if (IsUsableBrowserResolvedName(line, isFolderLink, inferredExtension))
                {
                    if (string.IsNullOrWhiteSpace(Path.GetExtension(line)) && !string.IsNullOrWhiteSpace(inferredExtension))
                    {
                        return line + inferredExtension;
                    }

                    return line;
                }

                string relaxedLine = NormalizeBrowserFileNameCandidate(rawLine, isFolderLink, inferredExtension);
                if (!string.IsNullOrWhiteSpace(relaxedLine))
                {
                    return relaxedLine;
                }
            }

            return string.Empty;
        }

        private static string GetFolderNameFromBrowserUrlCandidates(string candidates)
        {
            foreach (string rawLine in candidates.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries))
            {
                string value = HttpUtility.HtmlDecode(rawLine).Trim().Trim('"', '\'');
                if (value.IndexOf("http://", StringComparison.OrdinalIgnoreCase) < 0
                    && value.IndexOf("https://", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    continue;
                }

                try
                {
                    Uri uri = new Uri(value);
                    foreach (string key in new[] { "id", "RootFolder", "rootFolder", "folder" })
                    {
                        string queryValue = GetQueryValue(uri, key);
                        string candidate = CleanResolvedFolderTitle(GetLastPathSegment(queryValue));
                        if (IsUsableBrowserResolvedName(candidate, true, string.Empty))
                        {
                            return candidate;
                        }
                    }
                }
                catch
                {
                }
            }

            return string.Empty;
        }

        private static bool IsLoginRequiredBrowserResult(string candidates)
        {
            if (string.IsNullOrWhiteSpace(candidates))
            {
                return false;
            }

            string value = candidates.Trim();
            if (value.StartsWith("LOGIN_REQUIRED", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return value.IndexOf("login.microsoftonline.com", StringComparison.OrdinalIgnoreCase) >= 0
                || value.IndexOf("login.live.com", StringComparison.OrdinalIgnoreCase) >= 0
                || value.IndexOf("login.windows.net", StringComparison.OrdinalIgnoreCase) >= 0
                || value.IndexOf("someone@example.com", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string NormalizeBrowserFileNameCandidate(string value, bool isFolderLink, string inferredExtension)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            string candidate = HttpUtility.HtmlDecode(value).Trim().Trim('"', '\'');
            if (candidate.IndexOf('/') >= 0 || candidate.IndexOf('\\') >= 0)
            {
                string leaf = GetLastPathSegment(candidate);
                if (!string.IsNullOrWhiteSpace(leaf))
                {
                    candidate = leaf;
                }
            }

            candidate = candidate.Trim().Trim('"', '\'');
            if (candidate.Length == 0 || candidate.Length > 220)
            {
                return string.Empty;
            }

            if (IsGenericBrowserTitle(candidate) || IsNonFileNameBrowserCandidate(candidate) || LooksLikeOpaqueSharePointId(candidate))
            {
                return string.Empty;
            }

            if (isFolderLink)
            {
                return candidate;
            }

            string extension = Path.GetExtension(candidate);
            if (string.IsNullOrWhiteSpace(extension))
            {
                return string.Empty;
            }

            if (!string.IsNullOrWhiteSpace(inferredExtension)
                && !string.Equals(extension, inferredExtension, StringComparison.OrdinalIgnoreCase)
                && !IsCompatibleOfficeExtension(extension, inferredExtension))
            {
                return string.Empty;
            }

            return candidate;
        }

        private static bool IsUsableBrowserResolvedName(string name, bool isFolderLink, string inferredExtension)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return false;
            }

            if (IsGenericBrowserTitle(name) || IsNonFileNameBrowserCandidate(name))
            {
                return false;
            }

            if (isFolderLink)
            {
                return true;
            }

            string extension = Path.GetExtension(name);
            if (string.IsNullOrWhiteSpace(extension))
            {
                return false;
            }

            return string.IsNullOrWhiteSpace(inferredExtension)
                || string.Equals(extension, inferredExtension, StringComparison.OrdinalIgnoreCase)
                || IsCompatibleOfficeExtension(extension, inferredExtension);
        }

        private static bool IsCompatibleOfficeExtension(string extension, string inferredExtension)
        {
            if (string.IsNullOrWhiteSpace(extension) || string.IsNullOrWhiteSpace(inferredExtension))
            {
                return false;
            }

            if (IsExcelExtension(inferredExtension))
            {
                return IsExcelExtension(extension);
            }

            if (IsWordExtension(inferredExtension))
            {
                return IsWordExtension(extension);
            }

            if (IsPowerPointExtension(inferredExtension))
            {
                return IsPowerPointExtension(extension);
            }

            if (IsVisioExtension(inferredExtension))
            {
                return IsVisioExtension(extension);
            }

            return false;
        }

        private static bool IsNonFileNameBrowserCandidate(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return true;
            }

            string candidate = value.Trim().Trim('$').Trim();
            string extension = Path.GetExtension(candidate);
            if (string.Equals(extension, ".aspx", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".html", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".htm", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (Regex.IsMatch(candidate, "^[^\\s@<>]+@[^\\s@<>]+\\.[^\\s@<>]+$", RegexOptions.IgnoreCase))
            {
                return true;
            }

            foreach (string generic in new[]
            {
                "someone@example.com",
                "user@example.com",
                "name@example.com",
                "example@example.com",
                "AllItems.aspx",
                "Forms.aspx",
                "DispForm.aspx",
                "EditForm.aspx",
                "NewForm.aspx"
            })
            {
                if (string.Equals(candidate, generic, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static string CleanResolvedPageTitle(string title)
        {
            if (string.IsNullOrWhiteSpace(title))
            {
                return string.Empty;
            }

            string candidate = HttpUtility.HtmlDecode(title).Trim().Trim('"', '\'');
            if (candidate.IndexOf('/') >= 0 || candidate.IndexOf('\\') >= 0)
            {
                string leaf = GetLastPathSegment(candidate);
                if (!string.IsNullOrWhiteSpace(leaf))
                {
                    candidate = leaf;
                }
            }

            foreach (string suffix in new[]
            {
                " - Excel",
                " - Word",
                " - PowerPoint",
                " - Microsoft 365",
                " - SharePoint",
                " | Microsoft 365",
                " | SharePoint"
            })
            {
                if (candidate.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                {
                    candidate = candidate.Substring(0, candidate.Length - suffix.Length).Trim();
                }
            }

            return IsGenericBrowserTitle(candidate) ? string.Empty : candidate;
        }

        private static string CleanResolvedFolderTitle(string title)
        {
            string candidate = CleanResolvedPageTitle(title);
            if (string.IsNullOrWhiteSpace(candidate))
            {
                return string.Empty;
            }

            foreach (string suffix in new[]
            {
                " - Microsoft OneDrive",
                " | Microsoft OneDrive",
                " - OneDrive",
                " | OneDrive",
                " - Microsoft SharePoint",
                " | Microsoft SharePoint",
                " - SharePoint",
                " | SharePoint",
                " - All Documents",
                " | All Documents",
                " - Shared Documents",
                " | Shared Documents",
                " - Documents",
                " | Documents",
                " - \u3059\u3079\u3066\u306e\u30c9\u30ad\u30e5\u30e1\u30f3\u30c8",
                " | \u3059\u3079\u3066\u306e\u30c9\u30ad\u30e5\u30e1\u30f3\u30c8",
                " - \u5171\u6709\u30c9\u30ad\u30e5\u30e1\u30f3\u30c8",
                " | \u5171\u6709\u30c9\u30ad\u30e5\u30e1\u30f3\u30c8",
                " - \u30c9\u30ad\u30e5\u30e1\u30f3\u30c8",
                " | \u30c9\u30ad\u30e5\u30e1\u30f3\u30c8"
            })
            {
                if (candidate.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                {
                    candidate = candidate.Substring(0, candidate.Length - suffix.Length).Trim();
                }
            }

            if (candidate.IndexOf("http://", StringComparison.OrdinalIgnoreCase) >= 0
                || candidate.IndexOf("https://", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return string.Empty;
            }

            string extension = Path.GetExtension(candidate);
            if (string.Equals(extension, ".aspx", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".html", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".htm", StringComparison.OrdinalIgnoreCase))
            {
                return string.Empty;
            }

            return IsGenericBrowserTitle(candidate) || IsNonFileNameBrowserCandidate(candidate)
                ? string.Empty
                : candidate;
        }

        private static bool IsGenericBrowserTitle(string title)
        {
            if (string.IsNullOrWhiteSpace(title))
            {
                return true;
            }

            string candidate = title.Trim();
            candidate = Regex.Replace(candidate, "\\s*\\(\\d+\\)$", string.Empty).Trim();
            foreach (string generic in new[]
            {
                "Microsoft 365",
                "SharePoint",
                "OneDrive",
                "Excel",
                "Word",
                "PowerPoint",
                "Sign in to your account",
                "Sign in",
                "Login",
                "Account sign in",
                "Organization background image",
                "Background image",
                "\u30a2\u30ab\u30a6\u30f3\u30c8\u306b\u30b5\u30a4\u30f3\u30a4\u30f3",
                "\u30b5\u30a4\u30f3\u30a4\u30f3\u3057\u3066\u304f\u3060\u3055\u3044",
                "\u30b5\u30a4\u30f3\u30a4\u30f3",
                "\u30ed\u30b0\u30a4\u30f3",
                "\u7d44\u7e54\u306e\u80cc\u666f\u753b\u50cf",
                "\u80cc\u666f\u753b\u50cf",
                "\u8aad\u307f\u8fbc\u307f\u4e2d"
            })
            {
                if (string.Equals(candidate, generic, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            if (candidate.IndexOf("\u7d44\u7e54\u306e\u80cc\u666f\u753b\u50cf", StringComparison.OrdinalIgnoreCase) >= 0
                || candidate.IndexOf("\u80cc\u666f\u753b\u50cf", StringComparison.OrdinalIgnoreCase) >= 0
                || candidate.IndexOf("background image", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            return false;
        }

        private static string GetQueryValue(Uri uri, string key)
        {
            string query = uri.Query;
            if (string.IsNullOrWhiteSpace(query))
            {
                return string.Empty;
            }

            if (query.StartsWith("?", StringComparison.Ordinal))
            {
                query = query.Substring(1);
            }

            foreach (string pair in query.Split('&'))
            {
                int equals = pair.IndexOf('=');
                string rawKey = equals >= 0 ? pair.Substring(0, equals) : pair;
                if (!string.Equals(DecodeUrlComponent(rawKey), key, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string rawValue = equals >= 0 ? pair.Substring(equals + 1) : string.Empty;
                return DecodeUrlComponent(rawValue);
            }

            return string.Empty;
        }

        private static string DecodeUrlComponent(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            string plusAsSpace = value.Replace("+", "%20");
            return HttpUtility.UrlDecode(plusAsSpace, Encoding.UTF8);
        }

        private static string GetLastPathSegment(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            string normalized = DecodeUrlComponent(value).Trim().TrimEnd('/', '\\');
            int queryIndex = normalized.IndexOf('?');
            if (queryIndex >= 0)
            {
                normalized = normalized.Substring(0, queryIndex);
            }

            int fragmentIndex = normalized.IndexOf('#');
            if (fragmentIndex >= 0)
            {
                normalized = normalized.Substring(0, fragmentIndex);
            }

            string[] parts = normalized.Split(new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries);
            return parts.Length == 0 ? string.Empty : parts[parts.Length - 1].Trim();
        }

    }
}
