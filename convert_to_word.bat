@echo off
chcp 65001 >nul
echo ========================================
echo Markdown → Word 変換ツール
echo ========================================
echo.

REM Pandocがインストールされているか確認
where pandoc >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo [エラー] Pandocがインストールされていません。
    echo.
    echo Pandocをインストールしてください:
    echo   1. https://pandoc.org/installing.html にアクセス
    echo   2. 「Download the latest installer」をクリック
    echo   3. Windows用インストーラー（.msi）をダウンロード
    echo   4. インストール後、このバッチファイルを再実行
    echo.
    pause
    exit /b 1
)

echo Pandocが見つかりました。変換を開始します...
echo.

REM SETUP_GUIDE.mdをWord形式に変換（高品質設定）
pandoc SETUP_GUIDE.md -o SETUP_GUIDE.docx ^
    --toc ^
    --toc-depth=3 ^
    --number-sections ^
    --highlight-style=tango ^
    --reference-doc=reference.docx ^
    --metadata title="Medicom自動化ツール セットアップガイド" ^
    --metadata author="開発チーム" ^
    --metadata date="%date%" ^
    2>nul

if not exist reference.docx (
    echo [情報] 初回実行: 標準スタイルでWord文書を作成します...
    pandoc SETUP_GUIDE.md -o SETUP_GUIDE.docx ^
        --toc ^
        --toc-depth=3 ^
        --number-sections ^
        --highlight-style=tango ^
        --metadata title="Medicom自動化ツール セットアップガイド" ^
        --metadata author="開発チーム" ^
        --metadata date="%date%"
)

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo ✓ 変換完了: SETUP_GUIDE.docx
    echo ========================================
    echo.
    echo [注意] より良い書式にするには:
    echo   1. SETUP_GUIDE.docx を開く
    echo   2. スタイルを調整（見出し、フォントなど）
    echo   3. reference.docx として保存
    echo   4. このバッチファイルを再実行
    echo.
    echo ファイルを開きますか？ [Y/N]
    choice /C YN /N /M "選択: "
    if errorlevel 2 goto :end
    if errorlevel 1 start SETUP_GUIDE.docx
) else (
    echo.
    echo ========================================
    echo ✗ 変換に失敗しました
    echo ========================================
    echo.
    echo エラーが発生しました。以下を確認してください:
    echo   - SETUP_GUIDE.md が存在するか
    echo   - ファイルが他のプログラムで開かれていないか
    echo   - 書き込み権限があるか
)

:end
echo.
pause
