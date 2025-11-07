"""メインエントリーポイント"""
import os
import json
from auth import (
    add_account,
    select_account,
    update_password_menu,
    check_password_expiration,
    login,
    logout,
    load_accounts
)
from operations import daily_inventory, auto_order, check_messages
from utils import setup_driver


def normalize_input(text):
    """全角数字を半角数字に変換する

    Args:
        text: 入力文字列

    Returns:
        str: 半角数字に変換された文字列
    """
    # 全角数字を半角数字に変換
    translation_table = str.maketrans('０１２３４５６７８９', '0123456789')
    return text.translate(translation_table)


def load_config():
    """設定ファイルを読み込む"""
    config_file = "config.json"
    default_config = {
        "download_path": os.path.join(os.getcwd(), "downloads"),
        "should_print_pdf": True,
        "message_processing": {
            "購入伺い": True,
            "マッチング：使用期限": True,
            "不動在庫転送": True,
            "返信": True
        },
        "max_message_count": 10  # 連絡板の最大処理件数
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
    print("=== PDF自動ダウンロード・印刷システム ===\n")

    # 設定を読み込む
    config = load_config()

    # アカウント管理メニュー
    while True:
        print("\n1. 既存アカウントでログイン")
        print("2. 新しいアカウントを追加")
        print("3. パスワードを更新")
        print("4. 設定")
        print("5. 終了")

        choice = normalize_input(input("\n選択してください: "))

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
            print(f"3. スキップするメールの種類")
            print(f"4. 連絡板の最大処理件数: {config.get('max_message_count', 10)}件")
            print("5. 戻る")

            setting_choice = normalize_input(input("\n変更する項目を選択してください: "))

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
            elif setting_choice == "3":
                # スキップするメールの種類設定
                if 'message_processing' not in config:
                    config['message_processing'] = {
                        "購入伺い": True,
                        "マッチング：使用期限": True,
                        "不動在庫転送": True,
                        "返信": True
                    }

                msg_proc = config['message_processing']

                # 処理可能なメールタイトルのリスト
                mail_types = [
                    "購入伺い",
                    "マッチング：使用期限",
                    "不動在庫転送",
                    "返信"
                ]

                print("\n=== スキップするメールの種類 ===")
                print("処理をスキップするメールの種類を選択してください")
                print("\n現在の設定:")
                for i, mail_type in enumerate(mail_types, 1):
                    status = "処理する" if msg_proc.get(mail_type, True) else "スキップ"
                    print(f"  {i}. {mail_type:<20} [{status}]")

                print("\n番号を入力してON/OFFを切り替えます (0で戻る)")

                while True:
                    try:
                        toggle_choice = int(normalize_input(input("\n切り替える番号を入力: ")))
                        if toggle_choice == 0:
                            break
                        if 1 <= toggle_choice <= len(mail_types):
                            mail_type = mail_types[toggle_choice - 1]
                            # ON/OFF切り替え
                            current = msg_proc.get(mail_type, True)
                            msg_proc[mail_type] = not current
                            config['message_processing'] = msg_proc
                            save_config(config)

                            new_status = "処理する" if msg_proc[mail_type] else "スキップ"
                            print(f"✓ '{mail_type}' を [{new_status}] に変更しました")

                            # 更新後の一覧を表示
                            print("\n現在の設定:")
                            for i, mt in enumerate(mail_types, 1):
                                status = "処理する" if msg_proc.get(mt, True) else "スキップ"
                                print(f"  {i}. {mt:<20} [{status}]")
                        else:
                            print("無効な番号です")
                    except ValueError:
                        print("数字を入力してください")
            elif setting_choice == "4":
                # 連絡板の最大処理件数設定
                try:
                    current_count = config.get('max_message_count', 10)
                    print(f"\n現在の設定: {current_count}件")
                    new_count = input("新しい処理件数を入力してください (1-50): ").strip()
                    new_count = int(normalize_input(new_count))

                    if 1 <= new_count <= 50:
                        config['max_message_count'] = new_count
                        save_config(config)
                        print(f"✓ 連絡板の最大処理件数を {new_count}件 に変更しました")
                    else:
                        print("⚠️ 1から50の間で指定してください")
                except ValueError:
                    print("⚠️ 数字を入力してください")
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

        # 店舗IDを抽出して保持（セッション中使用）
        from operations import extract_store_id
        current_store_id = extract_store_id(account['user_id'])
        print(f"\n現在の店舗ID: {current_store_id}")

        # ログイン後のメニュー
        while True:
            max_count = config.get('max_message_count', 10)
            print("\n=== 作業メニュー ===")
            print("1. 毎日在庫")
            print("2. 自動発注")
            print(f"3. 連絡板確認（最大{max_count}件）")
            print("0. 終了")

            work_choice = normalize_input(input("\n作業を選択してください: "))

            if work_choice == "1":
                if daily_inventory(driver, download_path, config['should_print_pdf']):
                    input("\n処理が完了しました。Enterキーを押して続行...")
            elif work_choice == "2":
                if auto_order(driver, download_path, config['should_print_pdf']):
                    input("\n処理が完了しました。Enterキーを押して続行...")
            elif work_choice == "3":
                if check_messages(driver, account['user_id'], config):
                    input("\n処理が完了しました。Enterキーを押して続行...")
            elif work_choice == "0":
                # ログアウト処理
                logout(driver)
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
