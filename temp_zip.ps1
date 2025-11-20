$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Medicom自動化ツール パッケージ作成" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$files = @(
    "main.py",
    "auth.py",
    "operations.py",
    "utils.py",
    "README.md",
    "SETUP_GUIDE.md",
    "PRINT_SETUP.md",
    "requirements.txt",
    "run.bat",
    "run.ps1",
    "run.vbs"
)

Write-Host "含めるファイル:" -ForegroundColor Yellow
foreach ($file in $files) {
    if (Test-Path $file) {
        Write-Host "  ✓ $file" -ForegroundColor Green
    } else {
        Write-Host "  ✗ $file (存在しません)" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "圧縮中..." -ForegroundColor Yellow

$dest = Join-Path $env:USERPROFILE "Desktop\medicom-auto-browser-windows-$timestamp.zip"

try {
    Compress-Archive -Path $files -DestinationPath $dest -Force -ErrorAction Stop
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "✓ ZIP作成完了" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "保存先: $dest" -ForegroundColor Cyan
    Write-Host ""
} catch {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "✗ エラーが発生しました" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}
