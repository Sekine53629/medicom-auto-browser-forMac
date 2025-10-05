"""自動キャプチャスクリプト

ページの変化を自動検出して、HTML/スクリーンショットを保存します。
ユーザーは自由にブラウザを操作するだけでOK。
"""
import os
import sys
import time
import json
from datetime import datetime
from threading import Thread, Event

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

    try:
        # 現在のウィンドウ情報
        current_url = driver.current_url

        print(f"\n{'='*60}")
        print(f"📸 ページ {page_num} を自動保存中...")
        print(f"URL: {current_url}")
        print(f"{'='*60}")

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
        print(f"✓ 要素情報保存 (ボタン: {len(button_info)}, リンク: {len(link_info)})")

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
    except Exception as e:
        print(f"⚠️ 保存中にエラー: {e}")
        return None


def monitor_page_changes(driver, session_dir, all_pages, stop_event):
    """ページの変化を監視して自動保存"""
    page_counter = 0
    previous_url = None
    previous_handles = set()

    print("\n" + "="*60)
    print("🔍 ページ変化の監視を開始しました")
    print("自由にブラウザを操作してください。ページが変わると自動保存します。")
    print("="*60)

    while not stop_event.is_set():
        try:
            time.sleep(1)  # 1秒ごとにチェック

            current_handles = set(driver.window_handles)

            # 新しいウィンドウが開いたかチェック
            new_windows = current_handles - previous_handles
            if new_windows:
                print(f"\n🔔 新しいウィンドウが {len(new_windows)} 個開きました")
                for handle in new_windows:
                    driver.switch_to.window(handle)
                    time.sleep(1)  # ページ読み込み待機
                    page_counter += 1
                    page_info = save_current_state(driver, session_dir, page_counter, "新しいウィンドウ")
                    if page_info:
                        all_pages.append(page_info)
                previous_handles = current_handles

            # 閉じられたウィンドウがあるかチェック
            closed_windows = previous_handles - current_handles
            if closed_windows:
                print(f"\n🔔 ウィンドウが {len(closed_windows)} 個閉じられました")
                previous_handles = current_handles
                if current_handles:
                    driver.switch_to.window(list(current_handles)[0])

            # 現在のURLをチェック
            try:
                current_url = driver.current_url
            except:
                continue

            # URLが変わったら保存
            if current_url != previous_url and previous_url is not None:
                time.sleep(1.5)  # ページの完全な読み込みを待つ
                page_counter += 1
                page_info = save_current_state(driver, session_dir, page_counter, "ページ遷移")
                if page_info:
                    all_pages.append(page_info)
                previous_url = current_url
            elif previous_url is None:
                # 初回
                previous_url = current_url
                previous_handles = current_handles

        except Exception as e:
            # ブラウザが閉じられた場合など
            if "invalid session id" in str(e).lower() or "no such window" in str(e).lower():
                break
            time.sleep(0.5)


def main():
    """メイン処理"""
    print("="*60)
    print("自動キャプチャスクリプト")
    print("="*60)
    print("\n🎯 このスクリプトの動作:")
    print("- ログイン後、ブラウザを自由に操作してください")
    print("- ページが変わると自動的にHTML/スクリーンショットを保存します")
    print("- 終了するには、ターミナルで Ctrl+C を押してください")
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
    session_dir = os.path.join(output_dir, f"auto_capture_{timestamp}")
    os.makedirs(session_dir, exist_ok=True)

    print(f"\n📁 保存先: {session_dir}\n")

    driver = None
    all_pages = []
    stop_event = Event()

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

        # 初期ページを保存
        page_info = save_current_state(driver, session_dir, 1, "ログイン後のメインページ")
        if page_info:
            all_pages.append(page_info)

        # 監視スレッドを開始
        monitor_thread = Thread(target=monitor_page_changes, args=(driver, session_dir, all_pages, stop_event))
        monitor_thread.daemon = True
        monitor_thread.start()

        print("\n" + "="*60)
        print("✅ 準備完了！ブラウザを操作してください")
        print("終了するには Ctrl+C を押してください")
        print("="*60 + "\n")

        # メインスレッドは待機
        while True:
            time.sleep(1)

    except KeyboardInterrupt:
        print("\n\n" + "="*60)
        print("🛑 ユーザーによる中断")
        print("="*60)
        stop_event.set()
        time.sleep(1)

    except Exception as e:
        print(f"\n❌ エラーが発生しました: {e}")
        import traceback
        traceback.print_exc()

    finally:
        stop_event.set()

        # サマリー情報を保存
        if all_pages:
            summary = {
                "inspection_date": timestamp,
                "account": account.get('store_name', account['user_id']),
                "total_pages": len(all_pages),
                "pages": all_pages
            }

            summary_file = os.path.join(session_dir, "00_auto_capture_summary.json")
            with open(summary_file, 'w', encoding='utf-8') as f:
                json.dump(summary, f, indent=2, ensure_ascii=False)

        print("\n" + "="*60)
        print("✅ 調査完了！")
        print("="*60)
        print(f"\n📁 保存先: {session_dir}")
        print(f"📊 保存ページ数: {len(all_pages)}")

        if all_pages:
            print("\n📄 以下のファイルが作成されました:")
            for file in sorted(os.listdir(session_dir)):
                print(f"  - {file}")

            print("\n" + "="*60)
            print("次のステップ:")
            print("="*60)
            print("1. 上記フォルダ内のファイルを確認")
            print("2. HTMLファイル、PNGファイル、JSONファイルをClaude Codeに共有")
            print("3. Claude Codeが要素を解析して実装を完成させます")

        if driver:
            print("\nブラウザを閉じています...")
            try:
                driver.quit()
            except:
                pass


if __name__ == "__main__":
    main()
