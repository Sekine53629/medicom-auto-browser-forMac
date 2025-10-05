# デバッグツール

このフォルダには、HTML要素を調査するためのスクリプトが含まれています。

## 📋 スクリプト一覧

### 1. `terminal_capture.py` ⭐⭐⭐ 最推奨
**ターミナルコマンド操作モード** - ターミナルでコマンド入力して保存

- ✅ ターミナルで 's' または Enter を押すだけで保存
- ✅ ブラウザとターミナルを横並びで使用
- ✅ JavaScript生成ページにも対応
- ✅ ウィンドウ切り替え機能
- ✅ 'q' で終了
- ✅ 最もシンプルで確実

### 2. `capture_on_demand.py`
**撮影ボタンモード** - 画面に表示される撮影ボタンをクリックして保存

- ✅ 画面右上に「📸 撮影」ボタンを表示
- ⚠️ Mac環境でボタンが表示されない場合がある

### 3. `auto_capture.py`
**自動キャプチャモード** - ページの変化を自動検出して保存

- ✅ ページ遷移を自動検出して保存
- ⚠️ JavaScript生成ページは検出できない

### 4. `manual_inspector.py`
**手動保存モード** - メニュー操作型

- ⚠️ メニュー操作が複雑

### 5. `element_inspector.py`
**自動操作モード** - スクリプトが自動的にボタンをクリック

- ⚠️ サイトの構造変化に弱い（エラーが出やすい）

## 使用方法

### 最推奨: ターミナルコマンド操作モード（`terminal_capture.py`）⭐⭐⭐

#### 1. スクリプト実行（日中のサービス時間帯に実行）

```bash
cd /Users/yoshipc/Documents/GitHub/GitHub_Sekine53629/medicom-auto-browser-forMac
python debug_tools/terminal_capture.py
```

#### 2. 推奨レイアウト

```
┌─────────────────┬─────────────────┐
│                 │                 │
│   ブラウザ      │  ターミナル     │
│   (左半分)      │  (右半分)       │
│                 │                 │
│  [月次処理]     │  コマンド:      │
│  [発注]         │  s = 保存       │
│                 │  w = ウィンドウ  │
│                 │  q = 終了       │
└─────────────────┴─────────────────┘
```

#### 3. 使い方

1. アカウントを選択 → ログイン自動
2. ブラウザが開く（画面左側に配置される）
3. ターミナルを画面右側に配置
4. **メインページを保存したい** → ターミナルで `s` または Enter
5. **月次処理ボタンをクリック**（ブラウザ上で手動）
6. **ページが変わったら** → ターミナルで `s` または Enter
7. **繰り返し...**
8. **終了** → ターミナルで `q` + Enter

**コマンド一覧:**
- `s` または Enter = 現在のページを保存
- `w` = ウィンドウ一覧を表示
- `[数字]` = 指定したウィンドウに切り替え
- `h` = ヘルプ表示
- `q` = 終了

**特徴:**
- ✅ ターミナルを見ながら操作できる
- ✅ JavaScript生成ページにも対応
- ✅ ウィンドウ切り替えが簡単
- ✅ Mac環境で確実に動作

---

## 生成されるファイル

### ターミナルコマンド操作モードの場合（推奨）

`debug_tools/inspection_results/terminal_session_YYYYMMDD_HHMMSS/` に以下のファイルが生成されます：

```
00_terminal_summary.json            # 全体のサマリー情報

# 撮影ボタンをクリックするたびに保存（番号順）
01_page_YYYYMMDD_HHMMSS.html        # 1番目に撮影したページHTML
01_page_YYYYMMDD_HHMMSS.png         # 1番目に撮影したページスクリーンショット
01_page_elements.json               # 1番目に撮影したページの要素情報

02_page_YYYYMMDD_HHMMSS.html        # 2番目に撮影したページHTML
02_page_YYYYMMDD_HHMMSS.png
02_page_elements.json

03_page_YYYYMMDD_HHMMSS.html        # 3番目に撮影したページHTML
03_page_YYYYMMDD_HHMMSS.png
03_page_elements.json
...
```

### その他のモード

- `capture_on_demand.py` → `capture_session_YYYYMMDD_HHMMSS/`
- `auto_capture.py` → `auto_capture_YYYYMMDD_HHMMSS/`
- `manual_inspector.py` → `manual_session_YYYYMMDD_HHMMSS/`
- `element_inspector.py` → `session_YYYYMMDD_HHMMSS/`

---

## ファイルをClaude Codeに共有

生成されたフォルダ内のファイルをClaude Codeと共有してください。

**推奨共有ファイル（優先順）:**
1. `00_terminal_summary.json` - 全体の構造把握
2. 各ページの`.html`ファイル - 要素の詳細分析（最重要）
3. 各ページの`.png`ファイル - 視覚的確認
4. 各ページの`_elements.json`ファイル - 要素のクイックリファレンス

Claude CodeがHTMLを解析し、スクリーンショットと照らし合わせて、必要な要素を正確に特定して実装を完成させます。

**共有方法:**
VSCode上で、ファイルを選択して Claude Code とチャットで送信してください。

## トラブルシューティング

### エラーが発生した場合

- サービス時間帯（日中）に実行していることを確認
- アカウント情報が正しいことを確認
- Chromeブラウザがインストールされていることを確認

### ブラウザを表示したくない場合

`element_inspector.py` の19行目のコメントを解除：
```python
chrome_options.add_argument('--headless')  # この行のコメントを解除
```

## 注意事項

- このスクリプトは調査目的のみで使用します
- 生成されるHTMLファイルには機密情報が含まれる可能性があるため、取り扱いに注意してください
- `inspection_results/` フォルダは .gitignore に追加されています
