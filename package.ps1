$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"

Write-Host "========================================"
Write-Host "Medicom Automation Tool - Windows Edition"
Write-Host "========================================"
Write-Host ""

$files = @(
    'main.py',
    'auth.py',
    'operations.py',
    'utils.py',
    'README.md',
    'SETUP_GUIDE.md',
    'PRINT_SETUP.md',
    'requirements.txt',
    'run.bat',
    'run.ps1',
    'run.vbs'
)

Write-Host "Checking files..."
foreach ($file in $files) {
    if (Test-Path $file) {
        Write-Host "  OK: $file"
    } else {
        Write-Host "  MISSING: $file"
    }
}

Write-Host ""
Write-Host "Creating ZIP package..."

$dest = Join-Path $env:USERPROFILE "Desktop\medicom-auto-browser-windows-$timestamp.zip"

try {
    Compress-Archive -Path $files -DestinationPath $dest -Force -ErrorAction Stop
    Write-Host ""
    Write-Host "SUCCESS!"
    Write-Host "Package created: $dest"
    Write-Host ""
} catch {
    Write-Host ""
    Write-Host "ERROR: $($_.Exception.Message)"
}
