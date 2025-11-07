# UTF-8エンコーディングを設定
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8
chcp 65001 | Out-Null

# スクリプトのディレクトリに移動
$scriptDir = "C:\Users\0053629\Documents\GitHub\medicom-auto-browser-forMac"
Set-Location $scriptDir

# Pythonスクリプトを実行
python main.py @args

# 終了時にウィンドウを閉じないようにする
if ($LASTEXITCODE -ne $null) {
    Write-Host ""
    Write-Host "Press Enter to close window..."
    $null = Read-Host
}
