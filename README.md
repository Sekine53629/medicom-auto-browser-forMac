# medicom-auto-browser

Medicom自動ブラウザシステムです。

## 概要

このシステムは、MedicomのWebサイトに自動ログインし、毎日在庫や自動発注などの業務を自動化するためのツールです。

## 主な機能

- アカウント管理（複数店舗対応）
- 自動ログイン
- パスワード有効期限管理（30日）
- 毎日在庫処理
- 自動発注処理
- PDFダウンロード・印刷

## 必要な環境

- Python 3.7以上
- Chrome ブラウザ
- ChromeDriver（Selenium用）

## インストール

1. リポジトリをクローン
```bash
git clone <repository-url>
cd medicom-auto-browser
```

2. 依存関係をインストール
```bash
pip install -r requirements.txt
```

3. ChromeDriverをインストール
Chromeブラウザと互換性のあるChromeDriverをインストールしてください。

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
- `utils.py` - ユーティリティ関数
- `requirements.txt` - 依存関係
- `accounts.json` - アカウント情報（自動生成）
- `downloads/` - PDFダウンロードフォルダ

## 注意事項

- プリンタが正しく設定されていることを確認してください
- ChromeDriverのバージョンがChromeブラウザと互換性があることを確認してください
- アカウント情報は安全に管理してください

## トラブルシューティング

### 印刷ができない場合
プリンタの設定を確認してください。

### ChromeDriverのエラー
ChromeDriverのバージョンがChromeブラウザと合わない場合は、適切なバージョンをインストールしてください。
