Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$exePath = Join-Path $PSScriptRoot "dist\M365LinkShortcut.exe"

if (-not (Test-Path -LiteralPath $exePath)) {
    throw "M365LinkShortcut.exe が見つかりません。先に Build-Exe.ps1 を実行してください: $exePath"
}

& $exePath --install --quiet
Write-Host "右クリックメニューを登録しました: M365リンクをアイコン化"


