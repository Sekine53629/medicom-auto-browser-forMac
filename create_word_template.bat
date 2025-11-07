@echo off
chcp 65001 >nul
echo ========================================
echo Word参照テンプレート作成ツール
echo ========================================
echo.

REM Pandocがインストールされているか確認
where pandoc >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo [エラー] Pandocがインストールされていません。
    echo Pandocをインストールしてから再実行してください。
    echo.
    pause
    exit /b 1
)

echo 参照テンプレート（reference.docx）を作成します...
echo.

REM デフォルトの参照テンプレートを出力
pandoc --print-default-data-file reference.docx > reference.docx

if %ERRORLEVEL% EQU 0 (
    echo ✓ reference.docx を作成しました
    echo.
    echo [次のステップ]
    echo 1. reference.docx を開く
    echo 2. 以下のスタイルをカスタマイズ:
    echo    - 見出し 1, 2, 3 のフォント・サイズ・色
    echo    - 本文のフォント（推奨: 游ゴシック, Meiryo）
    echo    - 表のスタイル
    echo    - コードブロックのフォント（推奨: Consolas, 源ノ角ゴシック）
    echo 3. 上書き保存
    echo 4. convert_to_word.bat を実行
    echo.
    echo reference.docx を開きますか？ [Y/N]
    choice /C YN /N /M "選択: "
    if errorlevel 2 goto :end
    if errorlevel 1 start reference.docx
) else (
    echo ✗ 作成に失敗しました
)

:end
echo.
pause
