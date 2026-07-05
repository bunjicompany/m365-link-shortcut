@echo off
chcp 932 >nul
setlocal

set "EXE_PATH=%~dp0M365LinkShortcut.exe"
if not exist "%EXE_PATH%" (
    echo M365LinkShortcut.exe が同じフォルダに見つかりません。
    pause
    exit /b 1
)

"%EXE_PATH%" --install --quiet

if exist "%~dp0M365LinkShortcut_Logs.lnk" del "%~dp0M365LinkShortcut_Logs.lnk"

echo 通常モード（デバッグログなし）で右クリックメニューに登録しました。
pause
