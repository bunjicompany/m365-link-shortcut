@echo off
chcp 932 >nul
setlocal

set "EXE_PATH=%~dp0M365LinkShortcut.exe"
if not exist "%EXE_PATH%" (
    echo M365LinkShortcut.exe が同じフォルダに見つかりません。
    pause
    exit /b 1
)

"%EXE_PATH%" --install --debug --quiet

set "LOGS_DIR=%LOCALAPPDATA%\M365LinkShortcut\Logs"
if not exist "%LOGS_DIR%" mkdir "%LOGS_DIR%" >nul 2>&1

powershell -NoProfile -ExecutionPolicy Bypass -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%~dp0M365LinkShortcut_Logs.lnk'); $s.TargetPath = '%LOGS_DIR%'; $s.Description = 'M365LinkShortcut debug logs'; $s.Save()"

echo デバッグモードで右クリックメニューに登録しました。
echo ログ出力先: %LOGS_DIR%
echo Logsフォルダへのショートカットを作成しました: %~dp0M365LinkShortcut_Logs.lnk
echo 通常モードへ戻すには Install-NormalMode.bat を実行してください。
pause
