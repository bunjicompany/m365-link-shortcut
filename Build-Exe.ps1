Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$sourceDir = Join-Path $PSScriptRoot "src"
$testsDir = Join-Path $PSScriptRoot "tests"
$readmePath = Join-Path $PSScriptRoot "README.md"
$distDir = Join-Path $PSScriptRoot "dist"
$objDir = Join-Path $PSScriptRoot "obj"
$exePath = Join-Path $distDir "M365LinkShortcut.exe"
$testExePath = Join-Path $objDir "M365LinkShortcut.Tests.exe"
$distIndexPath = Join-Path $distDir "index.html"
$iconPath = Join-Path $PSScriptRoot "M365LinkShortcut.ico"
$chatIconPath = Join-Path $PSScriptRoot "chat.ico"
$meetingIconPath = Join-Path $PSScriptRoot "meeting.ico"
$compilerPath = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$webView2Root = Join-Path $PSScriptRoot "vendor\WebView2"
$appVersion = "1.0.0"
$releaseDate = "2026-07-05"

if (-not (Test-Path -LiteralPath $compilerPath)) {
    $compilerPath = Join-Path $env:WINDIR "Microsoft.NET\Framework\v4.0.30319\csc.exe"
}

if (-not (Test-Path -LiteralPath $compilerPath)) {
    throw "C# compiler was not found."
}

New-Item -ItemType Directory -Path $distDir -Force | Out-Null
New-Item -ItemType Directory -Path $objDir -Force | Out-Null

$sourceFiles = @(Get-ChildItem -LiteralPath $sourceDir -Filter "*.cs" | Sort-Object Name | ForEach-Object { $_.FullName })
if ($sourceFiles.Count -eq 0) {
    throw "Source files were not found under src."
}

$testSourcePath = Join-Path $testsDir "Tests.cs"

function New-AppIcon {
    param([string]$path)

    Add-Type -AssemblyName System.Drawing

    if (-not ("NativeIconMethods" -as [type])) {
        Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class NativeIconMethods {
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool DestroyIcon(IntPtr hIcon);
}
"@
    }

    $bitmap = New-Object System.Drawing.Bitmap 64, 64
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $brushBlue = $null
    $brushGreen = $null
    $brushYellow = $null
    $brushRed = $null
    $brushText = $null
    $pen = $null
    $font = $null
    $format = $null
    $stream = $null
    $icon = $null
    $handle = [IntPtr]::Zero

    try {
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $graphics.Clear([System.Drawing.Color]::Transparent)

        $brushBlue = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(0, 120, 212))
        $brushGreen = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(16, 124, 16))
        $brushYellow = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 185, 0))
        $brushRed = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(216, 59, 1))
        $brushText = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::White)
        $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::White), 3
        $font = New-Object System.Drawing.Font "Segoe UI", 13, ([System.Drawing.FontStyle]::Bold), ([System.Drawing.GraphicsUnit]::Pixel)
        $format = New-Object System.Drawing.StringFormat
        $format.Alignment = [System.Drawing.StringAlignment]::Center
        $format.LineAlignment = [System.Drawing.StringAlignment]::Center

        $graphics.FillRectangle($brushBlue, 8, 8, 22, 22)
        $graphics.FillRectangle($brushGreen, 34, 8, 22, 22)
        $graphics.FillRectangle($brushYellow, 8, 34, 22, 22)
        $graphics.FillRectangle($brushRed, 34, 34, 22, 22)
        $graphics.DrawEllipse($pen, 18, 23, 28, 18)
        $graphics.DrawString("365", $font, $brushText, ([System.Drawing.RectangleF]::new(0, 23, 64, 18)), $format)

        $handle = $bitmap.GetHicon()
        $icon = [System.Drawing.Icon]::FromHandle($handle)
        $stream = [System.IO.File]::Create($path)
        $icon.Save($stream)
    }
    finally {
        if ($stream -ne $null) { $stream.Dispose() }
        if ($icon -ne $null) { $icon.Dispose() }
        if ($handle -ne [IntPtr]::Zero) { [void][NativeIconMethods]::DestroyIcon($handle) }
        if ($format -ne $null) { $format.Dispose() }
        if ($font -ne $null) { $font.Dispose() }
        if ($pen -ne $null) { $pen.Dispose() }
        if ($brushText -ne $null) { $brushText.Dispose() }
        if ($brushRed -ne $null) { $brushRed.Dispose() }
        if ($brushYellow -ne $null) { $brushYellow.Dispose() }
        if ($brushGreen -ne $null) { $brushGreen.Dispose() }
        if ($brushBlue -ne $null) { $brushBlue.Dispose() }
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

if (-not (Test-Path -LiteralPath $iconPath)) {
    New-AppIcon $iconPath
}

foreach ($shortcutIcon in @($chatIconPath, $meetingIconPath)) {
    if (-not (Test-Path -LiteralPath $shortcutIcon)) {
        throw "Shortcut icon file was not found: $shortcutIcon"
    }
}

$webViewCore = $null
$webViewWinForms = $null
$webViewLoader = $null

function Set-WebView2ReferencesFromFolder {
    param([string]$folder)

    if ([string]::IsNullOrWhiteSpace($folder) -or -not (Test-Path -LiteralPath $folder)) {
        return $false
    }

    $core = Join-Path $folder "Microsoft.Web.WebView2.Core.dll"
    $winForms = Join-Path $folder "Microsoft.Web.WebView2.WinForms.dll"
    $loader = Join-Path $folder "runtimes\win-x64\native\WebView2Loader.dll"
    if ((Test-Path -LiteralPath $core) -and
        (Test-Path -LiteralPath $winForms) -and
        (Test-Path -LiteralPath $loader)) {
        $script:webViewCore = $core
        $script:webViewWinForms = $winForms
        $script:webViewLoader = $loader
        return $true
    }

    return $false
}

function Set-Utf8Text {
    param(
        [string]$Path,
        [string]$Value
    )

    $encoding = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Value, $encoding)
}

function Update-ReleaseMetadata {
    param([string]$Hash)

    if (Test-Path -LiteralPath $readmePath) {
        $readme = [System.IO.File]::ReadAllText($readmePath, [System.Text.Encoding]::UTF8)
        $readme = [regex]::Replace($readme, '- バージョン: .+', "- バージョン: $appVersion")
        $readme = [regex]::Replace($readme, '- 更新日: .+', "- 更新日: $releaseDate")
        $readme = [regex]::Replace($readme, '- SHA-256: `[A-Fa-f0-9]+`', "- SHA-256: ``$Hash``")
        Set-Utf8Text -Path $readmePath -Value $readme
    }

    if (Test-Path -LiteralPath $distIndexPath) {
        $index = [System.IO.File]::ReadAllText($distIndexPath, [System.Text.Encoding]::UTF8)
        $index = [regex]::Replace($index, '"softwareVersion": "[^"]+"', ('"softwareVersion": "' + $appVersion + '"'))
        $index = [regex]::Replace($index, 'バージョン: [^/]+ / 更新日: [^<]+', ("バージョン: $appVersion / 更新日: $releaseDate"))
        $index = [regex]::Replace($index, 'SHA-256: <code>[A-Fa-f0-9]+</code>', ("SHA-256: <code>$Hash</code>"))
        Set-Utf8Text -Path $distIndexPath -Value $index
    }
}

[void](Set-WebView2ReferencesFromFolder $webView2Root)

foreach ($requiredFile in @($webViewCore, $webViewWinForms, $webViewLoader)) {
    if (-not (Test-Path -LiteralPath $requiredFile)) {
        throw "Required WebView2 file was not found. Keep Microsoft.Web.WebView2.Core.dll, Microsoft.Web.WebView2.WinForms.dll, and runtimes\win-x64\native\WebView2Loader.dll under vendor\WebView2."
    }
}

$frameworkDir = Split-Path -Parent $compilerPath
$uiAutomationClient = Join-Path $frameworkDir "WPF\UIAutomationClient.dll"
$uiAutomationTypes = Join-Path $frameworkDir "WPF\UIAutomationTypes.dll"
$windowsBase = Join-Path $frameworkDir "WPF\WindowsBase.dll"
foreach ($requiredFile in @($uiAutomationClient, $uiAutomationTypes, $windowsBase)) {
    if (-not (Test-Path -LiteralPath $requiredFile)) {
        throw "Required UI Automation reference was not found: $requiredFile"
    }
}

if (Test-Path -LiteralPath $testSourcePath) {
    & $compilerPath `
        /nologo `
        /target:exe `
        /platform:x64 `
        /out:$testExePath `
        /main:SharePointShortcutMaker.TestRunner `
        /reference:System.dll `
        /reference:System.Core.dll `
        /reference:System.Web.dll `
        /reference:System.Web.Extensions.dll `
        /reference:System.Windows.Forms.dll `
        /reference:$uiAutomationClient `
        /reference:$uiAutomationTypes `
        /reference:$windowsBase `
        /reference:Microsoft.VisualBasic.dll `
        /reference:$webViewCore `
        /reference:$webViewWinForms `
        $sourceFiles `
        $testSourcePath

    if ($LASTEXITCODE -ne 0) {
        throw "Test compile failed. Exit code: $LASTEXITCODE"
    }

    & $testExePath
    if ($LASTEXITCODE -ne 0) {
        throw "Tests failed. Distribution exe was not updated."
    }

    Remove-Item -LiteralPath $testExePath -Force
}

& $compilerPath `
    /nologo `
    /target:winexe `
    /platform:x64 `
    /out:$exePath `
    /win32icon:$iconPath `
    "/resource:$webViewCore,Embedded.Microsoft.Web.WebView2.Core.dll" `
    "/resource:$webViewWinForms,Embedded.Microsoft.Web.WebView2.WinForms.dll" `
    "/resource:$webViewLoader,Embedded.WebView2Loader.dll" `
    "/resource:$chatIconPath,Embedded.ChatShortcut.ico" `
    "/resource:$meetingIconPath,Embedded.MeetingShortcut.ico" `
    /reference:System.dll `
    /reference:System.Core.dll `
    /reference:System.Web.dll `
    /reference:System.Web.Extensions.dll `
    /reference:System.Windows.Forms.dll `
    /reference:$uiAutomationClient `
    /reference:$uiAutomationTypes `
    /reference:$windowsBase `
    /reference:Microsoft.VisualBasic.dll `
    /reference:$webViewCore `
    /reference:$webViewWinForms `
    $sourceFiles

if ($LASTEXITCODE -ne 0) {
    throw "C# compile failed. Exit code: $LASTEXITCODE"
}

foreach ($oldDll in @(
    (Join-Path $distDir "Microsoft.Web.WebView2.Core.dll"),
    (Join-Path $distDir "Microsoft.Web.WebView2.WinForms.dll"),
    (Join-Path $distDir "WebView2Loader.dll")
)) {
    if (Test-Path -LiteralPath $oldDll) {
        Remove-Item -LiteralPath $oldDll -Force
    }
}

$licensePath = Join-Path $PSScriptRoot "LICENSE.txt"
if (Test-Path -LiteralPath $licensePath) {
    $licenseText = [System.IO.File]::ReadAllText($licensePath, [System.Text.Encoding]::UTF8)
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText((Join-Path $distDir "LICENSE.txt"), $licenseText, $utf8Bom)
}

foreach ($batchName in @("Install-DebugMode.bat", "Install-NormalMode.bat")) {
    $batchPath = Join-Path $PSScriptRoot $batchName
    if (Test-Path -LiteralPath $batchPath) {
        Copy-Item -LiteralPath $batchPath -Destination (Join-Path $distDir $batchName) -Force
    }
}

$hash = Get-FileHash -Algorithm SHA256 -LiteralPath $exePath
Set-Content -LiteralPath (Join-Path $distDir "SHA256SUMS.txt") -Value ($hash.Hash + "  M365LinkShortcut.exe") -Encoding ASCII
Update-ReleaseMetadata -Hash $hash.Hash

Write-Host "exeを作成しました: $exePath"
Write-Host "SHA-256: $($hash.Hash)"
