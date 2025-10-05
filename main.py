"""メインエントリーポイント"""
import os
import json
from auth import (
    add_account,
    select_account,
    update_password_menu,
    check_password_expiration,
    login,
    load_accounts
)
from operations import daily_inventory, auto_order, check_messages
from utils import setup_driver


def load_config():
    """設定ファイルを読み込む"""
    config_file = "config.json"
    default_config = {
        "download_path": os.path.join(os.getcwd(), "downloads"),
        "should_print_pdf": True
    }

    if os.path.exists(config_file):
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
                # デフォルト値とマージ
                return {**default_config, **config}
        except Exception as e:
            print(f"設定ファイルの読み込みエラー: {e}")
            return default_config
    else:
        # 設定ファイルが存在しない場合は作成
        try:
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(default_config, f, indent=2, ensure_ascii=False)
            print(f"設定ファイル {config_file} を作成しました。")
        except Exception as e:
            print(f"設定ファイルの作成エラー: {e}")
        return default_config


def save_config(config):
    """設定ファイルを保存する"""
    config_file = "config.json"
    try:
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"設定ファイルの保存エラー: {e}")
        return False


def main():
    """メイン処理"""
    print("=== PDF自動ダウンロード・印刷システム（Mac版） ===\n")

    # 設定を読み込む
    config = load_config()

    # アカウント管理メニュー
    while True:
        print("\n1. 既存アカウントでログイン")
        print("2. 新しいアカウントを追加")
        print("3. パスワードを更新")
        print("4. 設定")
        print("5. 終了")

        choice = input("\n選択してください: ")

        if choice == "1":
            account = select_account()
            if not account:
                continue

            # パスワード有効期限チェック
            check_password_expiration(account)

            # パスワード更新後、最新のアカウント情報を再読み込み
            accounts = load_accounts()
            for acc in accounts:
                if acc['user_id'] == account['user_id']:
                    account = acc
                    break

            break
        elif choice == "2":
            add_account()
        elif choice == "3":
            update_password_menu()
        elif choice == "4":
            # 設定メニュー
            print("\n=== 設定 ===")
            print(f"1. PDF保存先: {config['download_path']}")
            print(f"2. PDF自動印刷: {'ON' if config['should_print_pdf'] else 'OFF'}")
            print("3. 戻る")

            setting_choice = input("\n変更する項目を選択してください: ")

            if setting_choice == "1":
                new_path = input(f"新しいPDF保存先を入力してください（現在: {config['download_path']}）: ")
                if new_path:
                    config['download_path'] = new_path
                    save_config(config)
                    print(f"✓ PDF保存先を変更しました: {new_path}")
            elif setting_choice == "2":
                config['should_print_pdf'] = not config['should_print_pdf']
                save_config(config)
                print(f"✓ PDF自動印刷を{'ON' if config['should_print_pdf'] else 'OFF'}に変更しました")
            continue
        elif choice == "5":
            return
        else:
            print("無効な選択です。")

    # ダウンロードフォルダを設定
    download_path = config['download_path']
    os.makedirs(download_path, exist_ok=True)

    driver = None
    try:
        # Seleniumドライバーをセットアップ
        driver = setup_driver(download_path)

        # ログイン
        if not login(driver, account):
            print("ログインに失敗しました。")
            return

        # ログイン後のメニュー
        while True:
            print("\n=== 作業メニュー ===")
            print("1. 毎日在庫")
            print("2. 自動発注")
            print("3. 連絡板確認（テスト：1件のみ）")
            print("4. 終了")

            work_choice = input("\n作業を選択してください: ")

            if work_choice == "1":
                if daily_inventory(driver, download_path, config['should_print_pdf']):
                    input("\n処理が完了しました。Enterキーを押して続行...")
            elif work_choice == "2":
                if auto_order(driver, download_path, config['should_print_pdf']):
                    input("\n処理が完了しました。Enterキーを押して続行...")
            elif work_choice == "3":
                if check_messages(driver):
                    input("\n処理が完了しました。Enterキーを押して続行...")
            elif work_choice == "4":
                break
            else:
                print("無効な選択です。")

        print("\n処理を終了します。")

    except Exception as e:
        print(f"エラーが発生しました: {e}")

    finally:
        if driver:
            driver.quit()


if __name__ == "__main__":
    main()
