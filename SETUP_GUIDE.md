# Medicom自動化ツール セットアップガイド (Windows版)

## 目次
1. [Python のインストール](#python-のインストール)
2. [必要なライブラリのインストール](#必要なライブラリのインストール)
3. [Chrome WebDriver のセットアップ](#chrome-webdriver-のセットアップ)
4. [ツールの起動方法](#ツールの起動方法)
5. [トラブルシューティング](#トラブルシューティング)

---

## Python のインストール

### 1. Python 3.8以降のダウンロード

1. **公式サイトにアクセス**
   - URL: https://www.python.org/downloads/

2. **Windows用インストーラーをダウンロード**
   - 「Download Python 3.x.x」ボタンをクリック
   - 推奨バージョン: **Python 3.10以降**

### 2. Python のインストール

1. **インストーラーを実行**
   - ダウンロードした `.exe` ファイルをダブルクリック

2. **重要: 「Add Python to PATH」にチェックを入れる**
   - インストーラー起動時、下部の「Add Python to PATH」に**必ずチェック**を入れてください
   - これを忘れると、コマンドラインでPythonが認識されません

3. **「Install Now」をクリック**
   - デフォルト設定でインストールが開始されます
   - 管理者権限が必要な場合は「はい」をクリック

4. **インストール完了を確認**
   - 「Setup was successful」と表示されたら完了
   - 「Close」をクリック

### 3. インストール確認

PowerShellまたはコマンドプロンプトを開いて以下を実行：

```powershell
python --version
```

**期待される出力:**
```
Python 3.10.x
```

バージョン番号が表示されればインストール成功です。

---

## 必要なライブラリのインストール

### 1. プロジェクトフォルダに移動

PowerShellを開いて、プロジェクトフォルダに移動します：

```powershell
cd C:\Users\YourName\Documents\medicom-auto-browser
```

### 2. 必要なライブラリ一覧

このツールでは以下のPythonライブラリが必要です：

| ライブラリ名 | 用途 | バージョン |
|------------|------|----------|
| selenium | Webブラウザ自動操作 | 4.x以降 |
| pywin32 | Windows API操作（印刷機能） | 最新版 |

### 3. ライブラリのインストール

プロジェクトフォルダで以下のコマンドを実行してください：

#### すべてのライブラリを一括インストール
```powershell
pip install -r requirements.txt
```

**期待される出力:**
```
Successfully installed selenium-4.x.x pywin32-xxx
```

または、個別にインストールする場合：

#### Selenium のインストール
```powershell
pip install selenium
```

#### pywin32 のインストール（Windows専用印刷機能）
```powershell
pip install pywin32
```

### 4. インストール確認

すべてのライブラリが正しくインストールされたか確認：

```powershell
pip list
```

以下のライブラリが表示されればOK：
- selenium
- pywin32

---

## Chrome WebDriver のセットアップ

### 自動管理（推奨）

Selenium 4.6以降では、Chrome WebDriverが**自動的に管理**されます。

- 追加のダウンロードやインストールは**不要**です
- Chromeブラウザさえインストールされていれば、初回実行時に自動的にWebDriverがダウンロードされます

### 前提条件: Google Chrome のインストール

1. **Google Chrome がインストールされているか確認**
   - ブラウザで「chrome://version」にアクセス
   - バージョン情報が表示されればOK

2. **インストールされていない場合**
   - URL: https://www.google.com/chrome/
   - インストーラーをダウンロードして実行

---

## ツールの起動方法

### 方法1: バッチファイルから起動（推奨）

1. **エクスプローラーでプロジェクトフォルダを開く**

2. **`run.bat` をダブルクリック**
   - コマンドプロンプトウィンドウが開き、ツールが起動します
   - UTF-8エンコーディングが自動設定されるため、日本語入力が正常に動作します

### 方法2: PowerShellから直接起動

PowerShellを開いて以下を実行：

```powershell
cd [プロジェクトフォルダのパス]
python main.py
```

### 方法3: デスクトップショートカット作成

1. **PowerShellを管理者として実行**

2. **プロジェクトフォルダに移動して以下のコマンドを実行**
   ```powershell
   powershell -ExecutionPolicy Bypass -File create_shortcut.ps1
   ```

3. **デスクトップに「Medicom自動化ツール」ショートカットが作成されます**

---

## トラブルシューティング

### エラー: 'python' は、内部コマンドまたは外部コマンド...として認識されていません

**原因:** PythonがPATHに追加されていません

**解決方法:**

1. **Pythonを再インストール**
   - インストーラーを再度実行
   - 「Add Python to PATH」に**必ずチェック**を入れる

または

2. **手動でPATHに追加**
   - システム環境変数を開く（「Windows + R」→「sysdm.cpl」→「詳細設定」タブ→「環境変数」）
   - 「Path」を選択して「編集」
   - 以下を追加:
     ```
     C:\Users\<ユーザー名>\AppData\Local\Programs\Python\Python310
     C:\Users\<ユーザー名>\AppData\Local\Programs\Python\Python310\Scripts
     ```

### エラー: No module named 'selenium'

**原因:** Seleniumライブラリがインストールされていません

**解決方法:**
```powershell
pip install selenium
```

### エラー: No module named 'win32print' または 'win32api'

**原因:** pywin32ライブラリがインストールされていません

**解決方法:**
```powershell
pip install pywin32
```

### エラー: WebDriver の起動に失敗

**原因1:** Google Chrome がインストールされていない

**解決方法:**
- Google Chrome をインストール: https://www.google.com/chrome/

**原因2:** Chrome のバージョンが古すぎる

**解決方法:**
- Chrome を最新バージョンに更新

### 日本語が文字化けする

**原因:** エンコーディング設定が正しくない

**解決方法:**
- `run.bat` から起動してください（UTF-8が自動設定されます）
- または、PowerShellで以下を実行してから起動:
  ```powershell
  chcp 65001
  ```

### PDF自動印刷が動作しない

**詳細は別途「PDF自動印刷セットアップガイド」を参照:**
- `PRINT_SETUP.md` ファイルを開いてください

**推奨設定:**
- Adobe Acrobat Reader DC（通常プリインストール済み）
- または SumatraPDF（軽量・高速）

---

## セットアップ完了チェックリスト

インストールが完了したら、以下を確認してください：

- [ ] Python がインストールされている（`python --version` で確認）
- [ ] pip が使える（`pip --version` で確認）
- [ ] Selenium がインストールされている（`pip list` で確認）
- [ ] pywin32 がインストールされている（`pip list` で確認）
- [ ] Google Chrome がインストールされている
- [ ] `run.bat` からツールが起動できる

すべてチェックできたら、セットアップ完了です！

---

## サポート情報

### 主要ファイル
- `main.py` - メインプログラム
- `auth.py` - 認証処理
- `operations.py` - 業務処理
- `utils.py` - ユーティリティ関数（Windows専用）
- `run.bat` - 起動用バッチファイル（推奨）
- `run.ps1` - PowerShell起動スクリプト
- `run.vbs` - バックグラウンド起動スクリプト
- `accounts.json` - アカウント情報（自動生成）

### ログファイル
実行ログは `logs/` フォルダに保存されます：
```
logs/operation_YYYYMMDD_HHMMSS.log
```

### 印刷設定

Windows環境で自動印刷を行うには、以下のいずれかをインストールしてください：

**推奨:**
- Adobe Acrobat Reader DC（多くのPCにプリインストール済み）
- SumatraPDF（軽量で高速）

詳細は `PRINT_SETUP.md` を参照してください。

---

**セットアップガイド 作成日: 2025年11月**

**バージョン: 2.0 (Windows版)**
