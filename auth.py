"""認証関連の機能"""
import json
import os
import time
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def load_accounts():
    """アカウント情報をJSONファイルから読み込む"""
    if not os.path.exists('accounts.json'):
        print("accounts.jsonが見つかりません。新規作成します。")
        return []

    with open('accounts.json', 'r', encoding='utf-8') as f:
        return json.load(f)


def save_accounts(accounts):
    """アカウント情報をJSONファイルに保存"""
    with open('accounts.json', 'w', encoding='utf-8') as f:
        json.dump(accounts, indent=2, ensure_ascii=False, fp=f)


def update_last_login(account):
    """最終ログイン日時を更新"""
    accounts = load_accounts()

    # 該当アカウントを探して更新
    for acc in accounts:
        if acc['user_id'] == account['user_id']:
            acc['last_login'] = datetime.now().isoformat()
            break

    save_accounts(accounts)


def add_account():
    """新しいアカウントを追加"""
    store_name = input("店舗名を入力してください (0でキャンセル): ")
    if store_name == "0":
        print("キャンセルしました。")
        return False

    user_id = input("ユーザーIDを入力してください (0でキャンセル): ")
    if user_id == "0":
        print("キャンセルしました。")
        return False

    password = input("パスワードを入力してください (0でキャンセル): ")
    if password == "0":
        print("キャンセルしました。")
        return False

    accounts = load_accounts()
    accounts.append({
        "store_name": store_name,
        "user_id": user_id,
        "password": password,
        "password_updated": datetime.now().isoformat()  # パスワード更新日時
    })
    save_accounts(accounts)
    print(f"アカウント '{store_name}' を追加しました。")
    return True


def update_password(account):
    """パスワードを更新"""
    accounts = load_accounts()

    # 該当アカウントを探して更新
    for acc in accounts:
        if acc['user_id'] == account['user_id']:
            new_password = input(f"{acc.get('store_name', acc['user_id'])} の新しいパスワードを入力してください: ")
            acc['password'] = new_password
            acc['password_updated'] = datetime.now().isoformat()
            save_accounts(accounts)
            print("パスワードを更新しました。")
            return True

    return False


def update_password_menu():
    """メニューからパスワードを更新"""
    accounts = load_accounts()

    if not accounts:
        print("保存されたアカウントがありません。")
        return

    print("\n保存されたアカウント:")
    for i, account in enumerate(accounts, 1):
        store_name = account.get('store_name', account['user_id'])
        print(f"{i}. {store_name}")

    while True:
        try:
            choice = int(input("\nパスワードを更新する店舗番号を選択してください (0でキャンセル): "))
            if choice == 0:
                return
            if 1 <= choice <= len(accounts):
                selected_account = accounts[choice - 1]
                new_password = input(f"\n{selected_account.get('store_name', selected_account['user_id'])} の新しいパスワードを入力してください: ")

                # パスワードを更新（更新日は登録日として現在日時を設定）
                selected_account['password'] = new_password
                selected_account['password_updated'] = datetime.now().isoformat()
                save_accounts(accounts)

                print(f"✓ {selected_account.get('store_name', selected_account['user_id'])} のパスワードを更新しました。")
                return
            else:
                print("無効な番号です。")
        except ValueError:
            print("数字を入力してください。")


def check_password_expiration(account):
    """パスワード有効期限をチェック（30日）"""
    password_updated = account.get('password_updated')

    if not password_updated:
        print("\n⚠️  警告: パスワード更新日が記録されていません。")
        update_choice = input("パスワードを更新しますか？ (y/n): ")
        if update_choice.lower() == 'y':
            update_password(account)
        return

    # パスワード更新日からの経過日数を計算
    updated_date = datetime.fromisoformat(password_updated)
    days_passed = (datetime.now() - updated_date).days
    days_remaining = 30 - days_passed

    if days_remaining <= 0:
        print(f"\n❌ パスワードの有効期限が切れています！（{abs(days_remaining)}日超過）")
        print("パスワードを更新してください。")
        update_password(account)
    elif days_remaining <= 5:
        print(f"\n⚠️  警告: パスワードの有効期限まで残り {days_remaining} 日です。")
        update_choice = input("今すぐパスワードを更新しますか？ (y/n): ")
        if update_choice.lower() == 'y':
            update_password(account)
    else:
        print(f"パスワード有効期限: 残り {days_remaining} 日")


def select_account():
    """保存されたアカウントから選択"""
    accounts = load_accounts()

    if not accounts:
        print("保存されたアカウントがありません。")
        return None

    # 最終ログイン日時順にソート（降順：新しい順）
    sorted_accounts = sorted(
        accounts,
        key=lambda x: x.get('last_login', ''),
        reverse=True
    )

    print("\n保存されたアカウント:")
    for i, account in enumerate(sorted_accounts, 1):
        store_name = account.get('store_name', account['user_id'])  # 旧データ対応
        last_login = account.get('last_login', '未使用')
        if last_login != '未使用':
            # 日時を読みやすい形式に変換
            try:
                dt = datetime.fromisoformat(last_login)
                last_login_str = dt.strftime('%Y/%m/%d %H:%M')
            except:
                last_login_str = last_login
            print(f"{i}. {store_name} (最終: {last_login_str})")
        else:
            print(f"{i}. {store_name} ({last_login})")

    while True:
        try:
            choice = int(input("\n使用する店舗番号を選択してください (0で戻る): "))
            if choice == 0:
                return None
            if 1 <= choice <= len(sorted_accounts):
                return sorted_accounts[choice - 1]
            else:
                print("無効な番号です。")
        except ValueError:
            print("数字を入力してください。")


def login(driver, account):
    """ログイン処理"""
    driver.get("https://www.ph-netmaster.jp/medicom/LoginTop.aspx")

    wait = WebDriverWait(driver, 10)

    # ユーザーID入力
    user_field = wait.until(EC.presence_of_element_located((By.ID, "txtUser")))
    user_field.clear()
    user_field.send_keys(account['user_id'])

    # パスワード入力
    pass_field = driver.find_element(By.ID, "txtPass")
    pass_field.clear()
    pass_field.send_keys(account['password'])

    # ログインボタンをクリック
    login_button = driver.find_element(By.ID, "btnLogin")
    login_button.click()

    time.sleep(3)  # ログイン完了を待つ

    # エラーチェック
    try:
        error_label = driver.find_element(By.ID, "lblErr")
        error_text = error_label.text
        if error_text:
            print(f"ログインエラー: {error_text}")

            # 同時ログインエラーの検出
            if "すでにログインされています" in error_text or "ログイン" in error_text:
                print("\n⚠️  他の端末で既にログインされている可能性があります。")
                print("他のブラウザやタブを閉じてから再度お試しください。")

            return False
    except:
        pass

    # ページ内容から同時ログインエラーを検出（念のため）
    try:
        page_text = driver.page_source
        if "すでにログインされています" in page_text:
            print("\n⚠️  同時ログインエラー: 既に他の場所でログインされています。")
            print("他のブラウザやタブを閉じてから再度お試しください。")
            return False
    except:
        pass

    # 不要なウィンドウ（統計画面など）を閉じる
    time.sleep(2)  # ウィンドウが開くのを待つ
    all_windows = driver.window_handles

    print(f"開いているウィンドウ数: {len(all_windows)}")

    # 不要なURLのウィンドウを閉じる
    unwanted_urls = [
        "I_Tenpo_Sakutaihi_Login.aspx",
        "SubFrameSanyo.aspx",
        "about:blank"
    ]

    main_window = None
    for window in all_windows:
        driver.switch_to.window(window)
        current_url = driver.current_url
        print(f"ウィンドウURL: {current_url}")

        # 不要なURLか確認
        should_close = any(unwanted in current_url for unwanted in unwanted_urls)

        if should_close:
            print(f"不要なウィンドウを閉じます: {current_url}")
            driver.close()
        else:
            main_window = window

    # メインウィンドウに戻る
    if main_window:
        driver.switch_to.window(main_window)

    print("ログイン成功")

    # 最終ログイン日時を更新
    update_last_login(account)

    return True


def logout(driver):
    """ログアウト処理

    Args:
        driver: Seleniumドライバー

    Returns:
        bool: ログアウト成功時True、失敗時False
    """
    try:
        print("\nログアウト処理を開始します...")

        # デフォルトコンテンツに戻る
        driver.switch_to.default_content()

        wait = WebDriverWait(driver, 10)

        # ログアウトボタンを探す（リトライ処理付き）
        logout_button = None
        for attempt in range(3):
            try:
                # まずデフォルトコンテンツで探す
                driver.switch_to.default_content()
                logout_button = driver.find_element(By.ID, "btnLogout")
                print(f"✓ ログアウトボタンが見つかりました（メインコンテンツ、試行 {attempt + 1}）")
                break
            except:
                # フレーム内を探す
                print(f"フレーム内を探しています...（試行 {attempt + 1}/3）")
                driver.switch_to.default_content()
                frames = driver.find_elements(By.TAG_NAME, "frame") + driver.find_elements(By.TAG_NAME, "iframe")

                for frame in frames:
                    try:
                        driver.switch_to.default_content()
                        driver.switch_to.frame(frame)
                        logout_button = driver.find_element(By.ID, "btnLogout")
                        print(f"✓ ログアウトボタンが見つかりました（フレーム内、試行 {attempt + 1}）")
                        break
                    except:
                        continue

                if logout_button:
                    break

                if attempt < 2:
                    print(f"⚠️ ログアウトボタンが見つかりません。再試行します...（{attempt + 1}/3）")
                    time.sleep(2)

        if not logout_button:
            print("⚠️ ログアウトボタンが見つかりませんでした")
            return False

        # ログアウトボタンをクリック
        print("✓ ログアウトボタンをクリックします...")
        logout_button.click()

        time.sleep(2)

        # アラートダイアログが表示されるか確認（リトライ処理付き）
        for alert_attempt in range(3):
            try:
                alert = wait.until(EC.alert_is_present())
                alert_text = alert.text
                print(f"確認ダイアログ: {alert_text}")
                alert.accept()  # OKをクリック
                print("✓ ログアウトしました")
                time.sleep(2)
                return True
            except Exception as e:
                if alert_attempt < 2:
                    print(f"⚠️ ダイアログを探しています...（{alert_attempt + 1}/3）")
                    time.sleep(2)
                else:
                    # アラートが表示されない場合もログアウト成功とみなす
                    print("✓ ログアウトしました（ダイアログなし）")
                    return True

        return True

    except Exception as e:
        print(f"ログアウトエラー: {e}")
        return False
