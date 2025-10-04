"""メインエントリーポイント"""
import os
from auth import (
    add_account,
    select_account,
    update_password_menu,
    check_password_expiration,
    login,
    load_accounts
)
from operations import daily_inventory, auto_order
from utils import setup_driver


def main():
    """メイン処理"""
    print("=== PDF自動ダウンロード・印刷システム（Mac版） ===\n")

    # アカウント管理メニュー
    while True:
        print("\n1. 既存アカウントでログイン")
        print("2. 新しいアカウントを追加")
        print("3. パスワードを更新")
        print("4. 終了")

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
            return
        else:
            print("無効な選択です。")

    # ダウンロードフォルダを設定
    download_path = os.path.join(os.getcwd(), "downloads")
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
            print("3. 終了")

            work_choice = input("\n作業を選択してください: ")

            if work_choice == "1":
                if daily_inventory(driver, download_path):
                    # TODO: PDFダウンロード・印刷処理
                    input("\n処理が完了しました。Enterキーを押して続行...")
            elif work_choice == "2":
                if auto_order(driver, download_path):
                    # TODO: PDFダウンロード・印刷処理
                    input("\n処理が完了しました。Enterキーを押して続行...")
            elif work_choice == "3":
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
