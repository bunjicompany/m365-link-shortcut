Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$exePath = Join-Path $PSScriptRoot "dist\M365LinkShortcut.exe"

if (Test-Path -LiteralPath $exePath) {
    & $exePath --uninstall --quiet
}
else {
    $keys = @(
        "HKCU:\Software\Classes\Directory\Background\shell\M365LinkShortcut",
        "HKCU:\Software\Classes\Directory\shell\M365LinkShortcut"
    )

    foreach ($key in $keys) {
        if (Test-Path -LiteralPath $key) {
            Remove-Item -LiteralPath $key -Recurse -Force
        }
    }
}

Write-Host "右クリックメニューを削除しました。"



