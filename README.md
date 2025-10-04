# medicom-auto-browser-forMac

Medicom自動ブラウザシステムのMac版です。

## 概要

このシステムは、MedicomのWebサイトに自動ログインし、毎日在庫や自動発注などの業務を自動化するためのツールです。Windows版からMac用に移植されました。

## 主な機能

- アカウント管理（複数店舗対応）
- 自動ログイン
- パスワード有効期限管理（30日）
- 毎日在庫処理
- 自動発注処理
- PDFダウンロード・印刷（Mac対応）

## Mac版での変更点

- Windows専用の`pywin32`ライブラリを削除
- 印刷機能をMacの`lpr`コマンドに変更
- プリンタ管理をMacの`lpstat`コマンドに変更

## 必要な環境

- Python 3.7以上
- Chrome ブラウザ
- ChromeDriver（Selenium用）

## インストール

1. リポジトリをクローン
```bash
git clone <repository-url>
cd medicom-auto-browser-forMac
```

2. 依存関係をインストール
```bash
pip install -r requirements.txt
```

3. ChromeDriverをインストール
```bash
# Homebrewを使用する場合
brew install chromedriver
```

## 使用方法

1. プログラムを実行
```bash
python main.py
```

2. 初回実行時は新しいアカウントを追加してください
3. 既存アカウントでログインして業務を実行できます

## ファイル構成

- `main.py` - メインエントリーポイント
- `auth.py` - 認証関連の機能
- `operations.py` - 業務処理関連の機能
- `utils.py` - ユーティリティ関数（Mac版）
- `requirements.txt` - 依存関係
- `accounts.json` - アカウント情報（自動生成）
- `downloads/` - PDFダウンロードフォルダ

## 注意事項

- プリンタが正しく設定されていることを確認してください
- ChromeDriverのバージョンがChromeブラウザと互換性があることを確認してください
- アカウント情報は安全に管理してください

## トラブルシューティング

### 印刷ができない場合
```bash
# デフォルトプリンタを確認
lpstat -d

# 利用可能なプリンタを確認
lpstat -p
```

### ChromeDriverのエラー
ChromeDriverのバージョンがChromeブラウザと合わない場合は、適切なバージョンをインストールしてください。
