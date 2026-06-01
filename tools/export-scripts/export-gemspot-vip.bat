@echo off
chcp 65001 >nul
echo ========================================
echo   GemSpot VIP 中文聊天室 匯出工具
echo ========================================
echo.

if "%TOKEN%"=="" (
    set /p TOKEN=請輸入你的 Discord Token:
)

if "%CHANNEL_ID%"=="" (
    set /p CHANNEL_ID=請輸入 VIP 中文聊天室的 Channel ID:
)

if "%CHANNEL_ID%"=="" (
    echo 錯誤：未輸入 Channel ID
    pause
    exit /b 1
)

set OUTPUT_DIR=%~dp0exports
if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"

echo.
echo 開始匯出頻道 %CHANNEL_ID% ...
echo 輸出資料夾：%OUTPUT_DIR%
echo.

"%~dp0DiscordChatExporter.Cli.exe" export ^
    -t "%TOKEN%" ^
    -c "%CHANNEL_ID%" ^
    -f HtmlDark ^
    -o "%OUTPUT_DIR%" ^
    --media

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ✓ 匯出成功！
    echo 檔案位於：%OUTPUT_DIR%
    explorer "%OUTPUT_DIR%"
) else (
    echo.
    echo ✗ 匯出失敗，請確認 Channel ID 是否正確。
)

pause
