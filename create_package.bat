@echo off
chcp 65001 >nul
echo ========================================
echo Medicom自動化ツール パッケージ作成
echo ========================================
echo.

REM タイムスタンプ生成
for /f "tokens=1-3 delims=/ " %%a in ('date /t') do set YYYYMMDD=%%a%%b%%c
for /f "tokens=1-2 delims=: " %%a in ('time /t') do set HHMM=%%a%%b
set TIMESTAMP=%YYYYMMDD%-%HHMM%

set DEST=%USERPROFILE%\Desktop\medicom-auto-browser-%TIMESTAMP%.zip

echo 以下のファイルをパッケージに含めます:
echo.
echo [Pythonファイル]
echo   - main.py
echo   - auth.py
echo   - operations.py
echo   - utils.py
echo   - debug_page_structure.py
echo.
echo [起動スクリプト]
echo   - run.bat
echo   - run.ps1
echo   - run.vbs
echo.
echo [ドキュメント]
echo   - README.md
echo   - SETUP_GUIDE.md
echo   - PRINT_SETUP.md
echo.
echo [ユーティリティ]
echo   - create_shortcut.ps1
echo   - convert_to_word.bat
echo   - create_word_template.bat
echo.
echo 圧縮中...
echo.

REM PowerShellで圧縮
powershell -Command "Compress-Archive -Path main.py,auth.py,operations.py,utils.py,debug_page_structure.py,run.bat,run.ps1,run.vbs,README.md,SETUP_GUIDE.md,PRINT_SETUP.md,create_shortcut.ps1,convert_to_word.bat,create_word_template.bat,'Word書式設定ガイド.md','WORD変換手順.txt' -DestinationPath '%DEST%' -Force"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo ✓ パッケージ作成完了
    echo ========================================
    echo.
    echo 保存先: %DEST%
    echo.
    echo デスクトップを開きますか? [Y/N]
    choice /C YN /N /M "選択: "
    if errorlevel 2 goto :end
    if errorlevel 1 explorer.exe "%USERPROFILE%\Desktop"
) else (
    echo.
    echo ========================================
    echo ✗ エラーが発生しました
    echo ========================================
)

:end
echo.
pause
