"""HTML要素調査スクリプト

このスクリプトは日中のサービス時間帯に実行し、
各ページのHTML要素とスクリーンショットを収集します。
"""
import os
import sys
import time
import json
from datetime import datetime

# 親ディレクトリのモジュールをインポート
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from auth import load_accounts, select_account


def setup_debug_driver(output_path):
    """デバッグ用のChromeドライバーをセットアップ"""
    chrome_options = Options()

    # ウィンドウを表示（デバッグ用）
    # chrome_options.add_argument('--headless')  # 必要に応じてコメント解除

    # PDFダウンロード設定
    prefs = {
        "download.default_directory": output_path,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(options=chrome_options)
    driver.maximize_window()  # ウィンドウを最大化
    return driver


def save_page_info(driver, page_name, output_dir, window_handle=None):
    """ページ情報を保存（HTML + スクリーンショット）"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # HTMLソースを保存
    html_filename = f"{page_name}_{timestamp}.html"
    html_path = os.path.join(output_dir, html_filename)
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(driver.page_source)
    print(f"✓ HTML保存: {html_filename}")

    # スクリーンショットを保存
    screenshot_filename = f"{page_name}_{timestamp}.png"
    screenshot_path = os.path.join(output_dir, screenshot_filename)
    driver.save_screenshot(screenshot_path)
    print(f"✓ スクリーンショット保存: {screenshot_filename}")

    # URLを保存
    url_info = {
        "page_name": page_name,
        "url": driver.current_url,
        "timestamp": timestamp,
        "html_file": html_filename,
        "screenshot_file": screenshot_filename,
        "window_handle": window_handle or driver.current_window_handle
    }

    return url_info


def wait_for_page_load(driver, timeout=10):
    """JavaScriptの実行完了を待つ"""
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        time.sleep(1)  # 追加の待機時間
        return True
    except Exception as e:
        print(f"⚠️ ページ読み込み待機タイムアウト: {e}")
        return False


def check_new_windows(driver, previous_handles, output_dir, action_description):
    """新しいウィンドウを検出して保存"""
    current_handles = driver.window_handles
    new_windows = [h for h in current_handles if h not in previous_handles]

    window_info = []

    if new_windows:
        print(f"\n🔔 新しいウィンドウが{len(new_windows)}個開きました")

        original_window = driver.current_window_handle

        for i, window_handle in enumerate(new_windows):
            driver.switch_to.window(window_handle)
            wait_for_page_load(driver)

            window_name = f"window_{len(previous_handles) + i + 1}"
            print(f"\n--- ウィンドウ {window_name} を保存 ---")
            print(f"URL: {driver.current_url}")

            # ページ情報を保存
            info = save_page_info(driver, window_name, output_dir, window_handle)
            info["opened_by"] = action_description
            window_info.append(info)

            # 要素情報も収集
            collect_page_elements(driver, window_name, output_dir)

        # 元のウィンドウに戻る（まだ存在する場合）
        if original_window in driver.window_handles:
            driver.switch_to.window(original_window)
        else:
            # 元のウィンドウが閉じられた場合は最後のウィンドウに切り替え
            driver.switch_to.window(driver.window_handles[-1])

    return window_info


def collect_page_elements(driver, page_name, output_dir):
    """ページ内の全要素を収集"""
    print(f"\n--- {page_name} の要素を収集 ---")

    # すべてのボタンを検索
    buttons = driver.find_elements(By.TAG_NAME, "input")
    button_info = []
    for i, btn in enumerate(buttons):
        btn_type = btn.get_attribute("type")
        btn_id = btn.get_attribute("id")
        btn_name = btn.get_attribute("name")
        btn_value = btn.get_attribute("value")
        btn_src = btn.get_attribute("src")
        btn_onclick = btn.get_attribute("onclick")

        if btn_type in ["button", "submit", "image"]:
            button_info.append({
                "index": i,
                "type": btn_type,
                "id": btn_id,
                "name": btn_name,
                "value": btn_value,
                "src": btn_src,
                "onclick": btn_onclick
            })
            print(f"  ボタン{i}: type={btn_type}, id={btn_id}, name={btn_name}, value={btn_value}")

    # すべてのリンクを検索（フィルタリングなし）
    links = driver.find_elements(By.TAG_NAME, "a")
    link_info = []
    for i, link in enumerate(links):
        link_href = link.get_attribute("href")
        link_text = link.text.strip()
        link_id = link.get_attribute("id")
        link_onclick = link.get_attribute("onclick")

        if link_href or link_text or link_onclick:
            link_info.append({
                "index": i,
                "href": link_href,
                "text": link_text,
                "id": link_id,
                "onclick": link_onclick
            })
            if link_text or "pdf" in str(link_href).lower():
                print(f"  リンク{i}: href={link_href}, text={link_text}, id={link_id}")

    # 要素情報をJSONで保存
    elements_data = {
        "page": page_name,
        "url": driver.current_url,
        "buttons": button_info,
        "links": link_info,
        "total_buttons": len(button_info),
        "total_links": len(link_info)
    }

    elements_file = os.path.join(output_dir, f"{page_name}_elements.json")
    with open(elements_file, 'w', encoding='utf-8') as f:
        json.dump(elements_data, f, indent=2, ensure_ascii=False)
    print(f"✓ 要素情報保存: {page_name}_elements.json (ボタン: {len(button_info)}, リンク: {len(link_info)})")


def login_and_inspect(driver, account, output_dir):
    """ログインして各ページを調査"""
    print("\n=== ログイン処理 ===")
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

    # ログイン前のウィンドウハンドルを記録
    handles_before_login = driver.window_handles

    # ログインボタンをクリック
    login_button = driver.find_element(By.ID, "btnLogin")
    login_button.click()

    time.sleep(3)
    wait_for_page_load(driver)

    # ログイン後に開いた全ウィンドウを確認
    print(f"\nログイン後のウィンドウ数: {len(driver.window_handles)}")

    # 不要なウィンドウを閉じる
    unwanted_urls = [
        "I_Tenpo_Sakutaihi_Login.aspx",
        "SubFrameSanyo.aspx",
        "about:blank"
    ]

    all_windows = driver.window_handles
    main_window = None
    closed_windows = []

    for window in all_windows:
        driver.switch_to.window(window)
        current_url = driver.current_url
        should_close = any(unwanted in current_url for unwanted in unwanted_urls)

        if should_close:
            print(f"不要なウィンドウを閉じます: {current_url}")
            closed_windows.append(current_url)
            driver.close()
        else:
            main_window = window

    if main_window:
        driver.switch_to.window(main_window)

    print("✓ ログイン成功")
    time.sleep(2)

    # メインページの情報を保存
    all_info = []
    print("\n=== メインページを保存 ===")
    info = save_page_info(driver, "01_main_page", output_dir)
    all_info.append(info)

    # メインページの全要素を収集
    collect_page_elements(driver, "01_main_page", output_dir)

    return all_info


def inspect_daily_inventory(driver, output_dir):
    """毎日在庫（月次処理）ページを調査"""
    print("\n=== 毎日在庫（月次処理）ページ調査 ===")
    all_info = []

    try:
        wait = WebDriverWait(driver, 10)

        # クリック前のウィンドウハンドルを記録
        handles_before = driver.window_handles.copy()

        # 月次処理ボタンをクリック
        monthly_button = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//input[@type='image' and contains(@src, '00_getuji.gif')]"))
        )
        print("月次処理ボタンをクリック...")
        monthly_button.click()

        time.sleep(3)
        wait_for_page_load(driver)

        # 新しいウィンドウが開いたかチェック
        new_windows = check_new_windows(driver, handles_before, output_dir, "月次処理ボタンクリック")
        all_info.extend(new_windows)

        # 現在のウィンドウ（メインまたは新ウィンドウ）の情報を保存
        if not new_windows:
            # ページ遷移の場合
            print("\n新しいウィンドウは開かず、ページ遷移しました")
            info = save_page_info(driver, "02_daily_inventory", output_dir)
            all_info.append(info)
            collect_page_elements(driver, "02_daily_inventory", output_dir)

    except Exception as e:
        print(f"エラー: {e}")
        import traceback
        traceback.print_exc()

    return all_info


def inspect_auto_order(driver, output_dir):
    """自動発注ページを調査"""
    print("\n=== 自動発注ページ調査 ===")
    all_info = []

    try:
        # メインページに戻る
        print("メインページに戻ります...")
        driver.back()
        time.sleep(2)
        wait_for_page_load(driver)

        wait = WebDriverWait(driver, 10)

        # クリック前のウィンドウハンドルを記録
        handles_before = driver.window_handles.copy()

        # 発注ボタンをクリック
        order_button = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//input[@type='image' and contains(@src, '00_hattyu.gif')]"))
        )
        print("発注ボタンをクリック...")
        order_button.click()

        time.sleep(3)
        wait_for_page_load(driver)

        # 新しいウィンドウが開いたかチェック
        new_windows = check_new_windows(driver, handles_before, output_dir, "発注ボタンクリック")
        all_info.extend(new_windows)

        # 現在のウィンドウ（メインまたは新ウィンドウ）の情報を保存
        if not new_windows:
            # ページ遷移の場合
            print("\n新しいウィンドウは開かず、ページ遷移しました")
            info = save_page_info(driver, "03_auto_order", output_dir)
            all_info.append(info)
            collect_page_elements(driver, "03_auto_order", output_dir)

    except Exception as e:
        print(f"エラー: {e}")
        import traceback
        traceback.print_exc()

    return all_info


def main():
    """メイン処理"""
    print("=== HTML要素調査スクリプト ===\n")

    # アカウント選択
    account = select_account()
    if not account:
        print("アカウントが選択されませんでした。")
        return

    # 出力ディレクトリを作成
    output_dir = os.path.join(os.path.dirname(__file__), "inspection_results")
    os.makedirs(output_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    session_dir = os.path.join(output_dir, f"session_{timestamp}")
    os.makedirs(session_dir, exist_ok=True)

    print(f"出力先: {session_dir}\n")

    driver = None
    all_page_info = []
    operation_flow = []

    try:
        # ドライバーをセットアップ
        driver = setup_debug_driver(session_dir)

        # ログインしてメインページを保存
        print("\n" + "="*60)
        print("ステップ1: ログインとメインページ保存")
        print("="*60)
        info = login_and_inspect(driver, account, session_dir)
        all_page_info.extend(info)
        operation_flow.append({
            "step": 1,
            "action": "ログイン",
            "pages_saved": len(info),
            "details": "メインページを保存しました"
        })

        input("\nメインページの確認が完了しました。Enterキーを押して毎日在庫ページへ進みます...")

        # 毎日在庫ページを調査
        print("\n" + "="*60)
        print("ステップ2: 毎日在庫（月次処理）ページ調査")
        print("="*60)
        info = inspect_daily_inventory(driver, session_dir)
        all_page_info.extend(info)
        operation_flow.append({
            "step": 2,
            "action": "月次処理ボタンクリック",
            "pages_saved": len(info),
            "details": f"{len(info)}個のページ/ウィンドウを保存しました"
        })

        input("\n毎日在庫ページの確認が完了しました。Enterキーを押して自動発注ページへ進みます...")

        # 自動発注ページを調査
        print("\n" + "="*60)
        print("ステップ3: 自動発注ページ調査")
        print("="*60)
        info = inspect_auto_order(driver, session_dir)
        all_page_info.extend(info)
        operation_flow.append({
            "step": 3,
            "action": "発注ボタンクリック",
            "pages_saved": len(info),
            "details": f"{len(info)}個のページ/ウィンドウを保存しました"
        })

        # サマリー情報を保存
        summary = {
            "inspection_date": timestamp,
            "account": account.get('store_name', account['user_id']),
            "total_pages": len(all_page_info),
            "operation_flow": operation_flow,
            "pages": all_page_info
        }

        summary_file = os.path.join(session_dir, "00_inspection_summary.json")
        with open(summary_file, 'w', encoding='utf-8') as f:
            json.dump(summary, f, indent=2, ensure_ascii=False)

        # 操作フローを別ファイルでも保存
        flow_file = os.path.join(session_dir, "00_operation_flow.json")
        with open(flow_file, 'w', encoding='utf-8') as f:
            json.dump(operation_flow, f, indent=2, ensure_ascii=False)

        print("\n" + "="*60)
        print("✅ 調査完了！")
        print("="*60)
        print(f"\n📁 保存先: {session_dir}")
        print(f"📊 保存ページ数: {len(all_page_info)}")
        print("\n📄 以下のファイルが作成されました:")
        for file in sorted(os.listdir(session_dir)):
            print(f"  - {file}")

        print("\n" + "="*60)
        print("次のステップ:")
        print("="*60)
        print("1. 上記フォルダ内のファイルを確認")
        print("2. HTMLファイル、PNGファイル、JSONファイルをClaude Codeに共有")
        print("3. Claude Codeが要素を解析して実装を完成させます")
        print("\n推奨: 以下のファイルを優先的に共有してください")
        print("  - 00_inspection_summary.json (全体のサマリー)")
        print("  - 00_operation_flow.json (操作フロー)")
        print("  - 各ページのHTMLファイル")
        print("  - 各ページのスクリーンショット")

    except Exception as e:
        print(f"\n❌ エラーが発生しました: {e}")
        import traceback
        traceback.print_exc()

    finally:
        if driver:
            input("\nEnterキーを押してブラウザを閉じます...")
            driver.quit()


if __name__ == "__main__":
    main()
