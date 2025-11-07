# PDF自動印刷セットアップガイド

## Windows環境での自動印刷

このスクリプトは以下のPDFビューワーに対応しています（優先順位順）：

1. **Adobe Acrobat Reader DC** （推奨・通常プリインストール済み）
2. **SumatraPDF** （軽量・高速）
3. デフォルトのPDFビューワー（印刷ダイアログが表示される可能性あり）

### Adobe Acrobat Reader DC

通常、Windowsには標準でインストールされています。

#### インストール確認
```powershell
# PowerShellで以下のいずれかを実行
Test-Path "C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
Test-Path "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
# True が返ればOK
```

#### インストールされていない場合
1. 公式サイトからダウンロード
   - URL: https://get.adobe.com/reader/
2. インストーラーを実行してデフォルト設定でインストール

### SumatraPDF（代替・軽量版）

Adobe Acrobat Readerより高速で軽量です。

1. 公式サイトからダウンロード
   - URL: https://www.sumatrapdfreader.org/download-free-pdf-viewer
   - **Installer版**をダウンロード（推奨）

2. インストール
   - ダウンロードした `.exe` ファイルを実行
   - デフォルト設定でインストール
   - インストール先: `C:\Program Files\SumatraPDF\`

3. インストール確認
   ```powershell
   # PowerShellで以下を実行
   Test-Path "C:\Program Files\SumatraPDF\SumatraPDF.exe"
   # True が返ればOK
   ```

### 印刷機能の動作確認

テストスクリプトを実行して印刷機能を確認できます：

```bash
python test_print.py
```

## 印刷の仕組み

### Adobe Acrobat Reader DCがインストールされている場合（推奨）
- **完全自動印刷**が可能
- `/t` オプションで印刷後に自動的に閉じる
- デフォルトプリンタに自動送信
- 通常はプリインストールされているため追加作業不要

### SumatraPDFがインストールされている場合
- **完全自動印刷**が可能（Adobe Acrobatより高速）
- 印刷ダイアログが表示されず、バックグラウンドで印刷
- `-print-to-default` オプションでデフォルトプリンタに自動送信
- `-silent` オプションでウィンドウを表示しない

### どちらもない場合
- デフォルトのPDFビューワーで印刷
- 印刷ダイアログが表示される可能性あり
- **手動での確認が必要になる場合がある**

## トラブルシューティング

### プリンタが見つからない
```bash
# デフォルトプリンタを確認
wmic printer where default=TRUE get name
```

### Adobe Acrobat Reader または SumatraPDF が検出されない
- インストール先が標準と異なる場合は、`utils.py` の該当パスリストを編集してください
  ```python
  # Adobe Acrobat Readerのパス
  acrobat_paths = [
      r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
      r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
      r"カスタムインストールパス\AcroRd32.exe",  # ← 追加
  ]

  # SumatraPDFのパス
  sumatra_paths = [
      r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
      r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
      r"カスタムインストールパス\SumatraPDF.exe",  # ← 追加
  ]
  ```

### 印刷ジョブが実行されない
1. プリンタがオンラインか確認
2. プリンタドライバが最新か確認
3. テスト印刷を実行:
   ```bash
   python test_print.py
   ```

## Mac環境

Macでは `lpr` コマンドを使用するため、追加のインストールは不要です。

```bash
# デフォルトプリンタ確認
lpstat -d
```

## 設定

印刷のON/OFFは、メインメニューの「設定」から変更できます：
- `設定` → `PDF自動印刷` → ON/OFF切り替え
