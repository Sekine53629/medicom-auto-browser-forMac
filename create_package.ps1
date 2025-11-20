# パッケージ作成スクリプト
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$destPath = "$env:USERPROFILE\Desktop\medicom-auto-browser-$timestamp.zip"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Medicom自動化ツール パッケージ作成" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 含めるファイルのリスト
$files = @(
    # Pythonファイル
    "main.py",
    "auth.py",
    "operations.py",
    "utils.py",
    "debug_page_structure.py",

    # 起動スクリプト
    "run.bat",
    "run.ps1",
    "run.vbs",

    # ドキュメント
    "README.md",
    "SETUP_GUIDE.md",
    "PRINT_SETUP.md",
    "Word書式設定ガイド.md",
    "WORD変換手順.txt",

    # ユーティリティ
    "create_shortcut.ps1",
    "convert_to_word.bat",
    "create_word_template.bat",

    # 設定ファイル（サンプル）
    "requirements.txt"
)

Write-Host "以下のファイルをパッケージに含めます:" -ForegroundColor Yellow
$includedFiles = @()
foreach ($file in $files) {
    if (Test-Path $file) {
        Write-Host "  ✓ $file" -ForegroundColor Green
        $includedFiles += $file
    } else {
        Write-Host "  - $file (存在しません)" -ForegroundColor Gray
    }
}

Write-Host ""
Write-Host "圧縮中..." -ForegroundColor Yellow

try {
    Compress-Archive -Path $includedFiles -DestinationPath $destPath -Force
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "✓ パッケージ作成完了" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "保存先: $destPath" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "デスクトップを開きますか? [Y/N]"
    $response = Read-Host
    if ($response -eq "Y" -or $response -eq "y") {
        explorer.exe "$env:USERPROFILE\Desktop"
    }
} catch {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "✗ エラーが発生しました" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}

Write-Host ""
Write-Host "Press Enter to exit..."
$null = Read-Host
