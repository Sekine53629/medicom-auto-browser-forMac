"""手動操作型HTML要素調査スクリプト

ユーザーが手動でボタンをクリックし、
Enterキーを押すたびに現在のページを保存します。
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
from auth import select_account


def setup_debug_driver(output_path):
    """デバッグ用のChromeドライバーをセットアップ"""
    chrome_options = Options()

    # ウィンドウを表示（手動操作用）
    # chrome_options.add_argument('--headless')  # ヘッドレスモードは使用しない

    # PDFダウンロード設定
    prefs = {
        "download.default_directory": output_path,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(options=chrome_options)
    driver.maximize_window()
    return driver


def save_current_state(driver, session_dir, page_counter, description=""):
    """現在のページ状態を保存"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    page_num = f"{page_counter:02d}"

    print(f"\n{'='*60}")
    print(f"📸 ページ {page_num} を保存中...")
    print(f"{'='*60}")

    # 現在のウィンドウ情報
    current_url = driver.current_url
    print(f"URL: {current_url}")

    # HTMLソースを保存
    html_filename = f"{page_num}_page_{timestamp}.html"
    html_path = os.path.join(session_dir, html_filename)
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(driver.page_source)
    print(f"✓ HTML保存: {html_filename}")

    # スクリーンショットを保存
    screenshot_filename = f"{page_num}_page_{timestamp}.png"
    screenshot_path = os.path.join(session_dir, screenshot_filename)
    driver.save_screenshot(screenshot_path)
    print(f"✓ スクリーンショット保存: {screenshot_filename}")

    # 要素情報を収集
    print(f"\n要素を収集中...")

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

    # すべてのリンクを検索
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

    # 要素情報をJSONで保存
    elements_data = {
        "page_number": page_num,
        "description": description,
        "url": current_url,
        "timestamp": timestamp,
        "buttons": button_info,
        "links": link_info,
        "total_buttons": len(button_info),
        "total_links": len(link_info)
    }

    elements_filename = f"{page_num}_page_elements.json"
    elements_path = os.path.join(session_dir, elements_filename)
    with open(elements_path, 'w', encoding='utf-8') as f:
        json.dump(elements_data, f, indent=2, ensure_ascii=False)
    print(f"✓ 要素情報保存: {elements_filename} (ボタン: {len(button_info)}, リンク: {len(link_info)})")

    return {
        "page_number": page_num,
        "description": description,
        "url": current_url,
        "timestamp": timestamp,
        "html_file": html_filename,
        "screenshot_file": screenshot_filename,
        "elements_file": elements_filename,
        "button_count": len(button_info),
        "link_count": len(link_info)
    }


def show_window_info(driver):
    """現在のウィンドウ情報を表示"""
    handles = driver.window_handles
    current = driver.current_window_handle

    print(f"\n{'='*60}")
    print(f"📋 現在のウィンドウ情報")
    print(f"{'='*60}")
    print(f"開いているウィンドウ数: {len(handles)}")

    for i, handle in enumerate(handles, 1):
        driver.switch_to.window(handle)
        marker = "👉 " if handle == current else "   "
        print(f"{marker}ウィンドウ{i}: {driver.current_url[:80]}")

    # 元のウィンドウに戻る
    driver.switch_to.window(current)
    print(f"{'='*60}")


def switch_window_menu(driver):
    """ウィンドウ切り替えメニュー"""
    handles = driver.window_handles

    if len(handles) == 1:
        print("\n他のウィンドウはありません。")
        return

    print(f"\n{'='*60}")
    print("ウィンドウを選択してください:")
    print(f"{'='*60}")

    for i, handle in enumerate(handles, 1):
        driver.switch_to.window(handle)
        print(f"{i}. {driver.current_url[:80]}")

    while True:
        try:
            choice = input(f"\nウィンドウ番号を入力 (1-{len(handles)}, 0=キャンセル): ")
            choice = int(choice)

            if choice == 0:
                return
            if 1 <= choice <= len(handles):
                driver.switch_to.window(handles[choice - 1])
                print(f"✓ ウィンドウ{choice}に切り替えました")
                print(f"URL: {driver.current_url}")
                return
            else:
                print("無効な番号です。")
        except ValueError:
            print("数字を入力してください。")


def main():
    """メイン処理"""
    print("="*60)
    print("手動操作型HTML要素調査スクリプト")
    print("="*60)
    print("\n使い方:")
    print("1. ブラウザで自由に操作してください")
    print("2. ページを保存したいタイミングでEnterキーを押す")
    print("3. 'q' を入力すると終了します")
    print("="*60)

    # アカウント選択
    account = select_account()
    if not account:
        print("アカウントが選択されませんでした。")
        return

    # 出力ディレクトリを作成
    output_dir = os.path.join(os.path.dirname(__file__), "inspection_results")
    os.makedirs(output_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    session_dir = os.path.join(output_dir, f"manual_session_{timestamp}")
    os.makedirs(session_dir, exist_ok=True)

    print(f"\n📁 保存先: {session_dir}\n")

    driver = None
    all_pages = []
    page_counter = 0

    try:
        # ドライバーをセットアップ
        driver = setup_debug_driver(session_dir)

        # ログイン
        print("\nログインページを開きます...")
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

        time.sleep(3)

        # 不要なウィンドウを閉じる
        unwanted_urls = [
            "I_Tenpo_Sakutaihi_Login.aspx",
            "SubFrameSanyo.aspx",
            "about:blank"
        ]

        all_windows = driver.window_handles
        main_window = None

        for window in all_windows:
            driver.switch_to.window(window)
            current_url = driver.current_url
            should_close = any(unwanted in current_url for unwanted in unwanted_urls)

            if should_close:
                print(f"不要なウィンドウを閉じます: {current_url}")
                driver.close()
            else:
                main_window = window

        if main_window:
            driver.switch_to.window(main_window)

        print("\n✓ ログイン成功")
        time.sleep(2)

        # メインループ
        while True:
            show_window_info(driver)

            print(f"\n{'='*60}")
            print("コマンド:")
            print("  Enter = 現在のページを保存")
            print("  w = ウィンドウを切り替え")
            print("  q = 終了")
            print(f"{'='*60}")

            command = input("\nコマンドを入力 (または説明を入力してEnter): ").strip()

            if command.lower() == 'q':
                print("\n調査を終了します...")
                break
            elif command.lower() == 'w':
                switch_window_menu(driver)
            else:
                # ページを保存
                page_counter += 1
                description = command if command else f"ページ{page_counter}"

                page_info = save_current_state(driver, session_dir, page_counter, description)
                all_pages.append(page_info)

                print(f"\n✅ 保存完了 (合計: {page_counter}ページ)")
                print("\n次の操作:")
                print("  - ブラウザで次の操作を行ってください")
                print("  - 保存したいタイミングでEnterキーを押してください")

        # サマリー情報を保存
        summary = {
            "inspection_date": timestamp,
            "account": account.get('store_name', account['user_id']),
            "total_pages": len(all_pages),
            "pages": all_pages
        }

        summary_file = os.path.join(session_dir, "00_manual_inspection_summary.json")
        with open(summary_file, 'w', encoding='utf-8') as f:
            json.dump(summary, f, indent=2, ensure_ascii=False)

        print("\n" + "="*60)
        print("✅ 調査完了！")
        print("="*60)
        print(f"\n📁 保存先: {session_dir}")
        print(f"📊 保存ページ数: {len(all_pages)}")
        print("\n📄 以下のファイルが作成されました:")
        for file in sorted(os.listdir(session_dir)):
            print(f"  - {file}")

        print("\n" + "="*60)
        print("次のステップ:")
        print("="*60)
        print("1. 上記フォルダ内のファイルを確認")
        print("2. HTMLファイル、PNGファイル、JSONファイルをClaude Codeに共有")
        print("3. Claude Codeが要素を解析して実装を完成させます")

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
