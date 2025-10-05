"""ターミナルコマンド操作型キャプチャスクリプト

ターミナルでコマンドを入力して操作します。
ブラウザとターミナルを横並びで表示して使用してください。
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

    # PDFダウンロード設定
    prefs = {
        "download.default_directory": output_path,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(options=chrome_options)

    # ウィンドウサイズを調整（画面の半分程度）
    driver.set_window_size(1200, 1000)
    driver.set_window_position(0, 0)

    return driver


def save_current_state(driver, session_dir, page_counter, description=""):
    """現在のページ状態を保存"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    page_num = f"{page_counter:02d}"

    try:
        # 現在のウィンドウ情報
        current_url = driver.current_url

        print(f"\n{'='*70}")
        print(f"📸 ページ {page_num} を保存中...")
        print(f"URL: {current_url[:60]}...")
        print(f"{'='*70}")

        # HTMLソースを保存
        html_filename = f"{page_num}_page_{timestamp}.html"
        html_path = os.path.join(session_dir, html_filename)
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(driver.page_source)
        print(f"✓ HTML: {html_filename}")

        # スクリーンショットを保存
        screenshot_filename = f"{page_num}_page_{timestamp}.png"
        screenshot_path = os.path.join(session_dir, screenshot_filename)
        driver.save_screenshot(screenshot_path)
        print(f"✓ PNG: {screenshot_filename}")

        # 要素情報を収集
        buttons = driver.find_elements(By.TAG_NAME, "input")
        button_info = []
        for i, btn in enumerate(buttons):
            try:
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
            except:
                continue

        # すべてのリンクを検索
        links = driver.find_elements(By.TAG_NAME, "a")
        link_info = []
        for i, link in enumerate(links):
            try:
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
            except:
                continue

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
        print(f"✓ JSON: {elements_filename}")
        print(f"  (ボタン: {len(button_info)}, リンク: {len(link_info)})")
        print(f"{'='*70}")

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


def show_help():
    """ヘルプを表示"""
    print("\n" + "="*70)
    print("📋 コマンド一覧")
    print("="*70)
    print("  s または Enter  = 現在のページを保存（スクリーンショット）")
    print("  w               = 開いているウィンドウ一覧を表示")
    print("  [数字]          = 指定したウィンドウに切り替え")
    print("  h               = このヘルプを表示")
    print("  q               = 終了")
    print("="*70 + "\n")


def show_windows(driver):
    """ウィンドウ一覧を表示"""
    handles = driver.window_handles
    current = driver.current_window_handle

    print("\n" + "="*70)
    print(f"📋 開いているウィンドウ ({len(handles)}個)")
    print("="*70)

    for i, handle in enumerate(handles, 1):
        driver.switch_to.window(handle)
        marker = "👉 " if handle == current else "   "
        url = driver.current_url
        print(f"{marker}{i}. {url[:65]}")

    # 元のウィンドウに戻る
    driver.switch_to.window(current)
    print("="*70 + "\n")


def main():
    """メイン処理"""
    print("="*70)
    print("ターミナルコマンド操作型キャプチャスクリプト")
    print("="*70)
    print("\n🎯 推奨レイアウト:")
    print("  左半分: ブラウザ")
    print("  右半分: このターミナル")
    print("\n📸 ページを保存したいタイミングで:")
    print("  → 's' または Enter キーを押してください")
    print("="*70)

    # アカウント選択
    account = select_account()
    if not account:
        print("アカウントが選択されませんでした。")
        return

    # 出力ディレクトリを作成
    output_dir = os.path.join(os.path.dirname(__file__), "inspection_results")
    os.makedirs(output_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    session_dir = os.path.join(output_dir, f"terminal_session_{timestamp}")
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

        print("\n" + "="*70)
        print("✅ 準備完了！")
        print("="*70)
        print("ブラウザを操作して、保存したいタイミングで 's' または Enter を押してください")
        print("コマンド一覧は 'h' で表示できます")
        print("="*70 + "\n")

        # メインループ
        while True:
            try:
                # コマンド入力を待つ
                command = input("コマンド ('h'=ヘルプ, 's'=保存, 'q'=終了): ").strip().lower()

                if command == 'q':
                    print("\n終了します...")
                    break

                elif command == 'h':
                    show_help()

                elif command == 'w':
                    show_windows(driver)

                elif command.isdigit():
                    # ウィンドウ切り替え
                    window_num = int(command)
                    handles = driver.window_handles
                    if 1 <= window_num <= len(handles):
                        driver.switch_to.window(handles[window_num - 1])
                        print(f"✓ ウィンドウ {window_num} に切り替えました")
                        print(f"  URL: {driver.current_url}")
                    else:
                        print(f"⚠️ ウィンドウ {window_num} は存在しません（1-{len(handles)}）")

                elif command == 's' or command == '':
                    # ページを保存
                    page_counter += 1
                    page_info = save_current_state(driver, session_dir, page_counter, f"ページ{page_counter}")
                    if page_info:
                        all_pages.append(page_info)
                        print(f"\n✅ 保存完了！ (合計: {page_counter}ページ)\n")

                else:
                    print(f"⚠️ 不明なコマンド: {command}")
                    print("'h' でヘルプを表示できます")

            except KeyboardInterrupt:
                print("\n\n終了します...")
                break

            except Exception as e:
                print(f"⚠️ エラー: {e}")

    except Exception as e:
        print(f"\n❌ エラーが発生しました: {e}")
        import traceback
        traceback.print_exc()

    finally:
        # サマリー情報を保存
        if all_pages:
            summary = {
                "inspection_date": timestamp,
                "account": account.get('store_name', account['user_id']),
                "total_pages": len(all_pages),
                "pages": all_pages
            }

            summary_file = os.path.join(session_dir, "00_terminal_summary.json")
            with open(summary_file, 'w', encoding='utf-8') as f:
                json.dump(summary, f, indent=2, ensure_ascii=False)

        print("\n" + "="*70)
        print("✅ 調査完了！")
        print("="*70)
        print(f"\n📁 保存先: {session_dir}")
        print(f"📊 保存ページ数: {len(all_pages)}")

        if all_pages:
            print("\n📄 生成されたファイル:")
            files = sorted(os.listdir(session_dir))
            for file in files[:10]:  # 最初の10件を表示
                print(f"  - {file}")
            if len(files) > 10:
                print(f"  ... 他 {len(files) - 10} 件")

            print("\n" + "="*70)
            print("次のステップ:")
            print("="*70)
            print("1. 上記フォルダ内のファイルを確認")
            print("2. HTMLファイル、PNGファイルをClaude Codeに共有")
            print("3. Claude Codeが要素を解析して実装を完成させます")

        if driver:
            print("\nブラウザを閉じています...")
            try:
                driver.quit()
            except:
                pass


if __name__ == "__main__":
    main()
