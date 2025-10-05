"""撮影ボタン付きキャプチャスクリプト

画面上に「📸 撮影」ボタンを表示し、
クリックするとHTML/スクリーンショットを保存します。
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
    driver.maximize_window()
    return driver


def inject_capture_button(driver):
    """ページに撮影ボタンを注入"""
    script = """
    (function() {
        // 既存のボタンを削除
        var existingButton = document.getElementById('claude-capture-button');
        if (existingButton) {
            existingButton.remove();
        }

        // bodyが存在することを確認
        if (!document.body) {
            console.error('document.body が存在しません');
            return false;
        }

        // 撮影ボタンを作成
        var button = document.createElement('div');
        button.id = 'claude-capture-button';
        button.innerHTML = '📸 撮影';
        button.style.position = 'fixed';
        button.style.top = '10px';
        button.style.right = '10px';
        button.style.zIndex = '2147483647';  // 最大値
        button.style.padding = '15px 30px';
        button.style.backgroundColor = '#4CAF50';
        button.style.color = 'white';
        button.style.border = '3px solid #fff';
        button.style.borderRadius = '8px';
        button.style.cursor = 'pointer';
        button.style.fontSize = '20px';
        button.style.fontWeight = 'bold';
        button.style.boxShadow = '0 4px 6px rgba(0,0,0,0.5)';
        button.style.transition = 'all 0.3s';
        button.style.userSelect = 'none';

        // ホバー効果
        button.onmouseover = function() {
            this.style.backgroundColor = '#45a049';
            this.style.transform = 'scale(1.1)';
        };
        button.onmouseout = function() {
            this.style.backgroundColor = '#4CAF50';
            this.style.transform = 'scale(1)';
        };

        // クリック時の視覚効果
        button.onclick = function(e) {
            e.preventDefault();
            e.stopPropagation();
            this.style.backgroundColor = '#FFA500';
            this.innerHTML = '✅ 保存中...';
            this.setAttribute('data-clicked', 'true');
            setTimeout(() => {
                this.style.backgroundColor = '#4CAF50';
                this.innerHTML = '📸 撮影';
                this.setAttribute('data-clicked', 'false');
            }, 1000);
        };

        button.setAttribute('data-clicked', 'false');

        document.body.appendChild(button);
        console.log('撮影ボタンを追加しました');
        return true;
    })();
    """

    try:
        result = driver.execute_script(script)
        return result
    except Exception as e:
        print(f"⚠️ ボタン注入エラー: {e}")
        return False


def check_button_clicked(driver):
    """撮影ボタンがクリックされたかチェック"""
    try:
        script = """
        var button = document.getElementById('claude-capture-button');
        if (button && button.getAttribute('data-clicked') === 'true') {
            button.setAttribute('data-clicked', 'false');
            return true;
        }
        return false;
        """
        return driver.execute_script(script)
    except:
        return False


def save_current_state(driver, session_dir, page_counter, description=""):
    """現在のページ状態を保存"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    page_num = f"{page_counter:02d}"

    try:
        # 現在のウィンドウ情報
        current_url = driver.current_url

        print(f"\n{'='*60}")
        print(f"📸 ページ {page_num} を保存中...")
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
        print(f"✓ 要素情報保存 (ボタン: {len(button_info)}, リンク: {len(link_info)})")
        print(f"{'='*60}")

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


def main():
    """メイン処理"""
    print("="*60)
    print("撮影ボタン付きキャプチャスクリプト")
    print("="*60)
    print("\n🎯 使い方:")
    print("1. ログイン後、画面右上に「📸 撮影」ボタンが表示されます")
    print("2. ボタンをクリックすると、その時点のページを保存します")
    print("3. 自由にブラウザを操作して、保存したいタイミングでボタンをクリック")
    print("4. 終了するには、このターミナルで 'q' + Enter を押してください")
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
    session_dir = os.path.join(output_dir, f"capture_session_{timestamp}")
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
        time.sleep(3)

        # ページが完全に読み込まれるまで待機
        try:
            wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
        except:
            pass

        time.sleep(2)

        # 撮影ボタンを注入（複数回試行）
        print("\n撮影ボタンを配置中...")
        for attempt in range(3):
            try:
                inject_capture_button(driver)
                # ボタンが実際に存在するか確認
                button_exists = driver.execute_script("return document.getElementById('claude-capture-button') !== null;")
                if button_exists:
                    print("✓ 撮影ボタンの配置に成功しました")
                    break
                else:
                    print(f"⚠️ ボタン配置失敗（試行 {attempt + 1}/3）再試行中...")
                    time.sleep(1)
            except Exception as e:
                print(f"⚠️ ボタン配置エラー（試行 {attempt + 1}/3）: {e}")
                time.sleep(1)

        print("\n" + "="*60)
        print("✅ 準備完了！")
        print("="*60)
        print("画面右上の「📸 撮影」ボタンをクリックすると保存します")
        print("※ ボタンが見えない場合は、ページを少しスクロールしてみてください")
        print("終了するには 'q' + Enter を入力してください")
        print("="*60 + "\n")

        # 監視ループ
        previous_handles = set(driver.window_handles)

        while True:
            # ユーザー入力をチェック（非ブロッキング）
            import select
            if select.select([sys.stdin], [], [], 0.1)[0]:
                user_input = sys.stdin.readline().strip()
                if user_input.lower() == 'q':
                    print("\n終了します...")
                    break

            try:
                # 撮影ボタンがクリックされたかチェック
                if check_button_clicked(driver):
                    page_counter += 1
                    page_info = save_current_state(driver, session_dir, page_counter, f"ページ{page_counter}")
                    if page_info:
                        all_pages.append(page_info)
                        print(f"\n✅ 保存完了！ (合計: {page_counter}ページ)\n")

                    # ボタンを再注入（ページが変わった可能性があるため）
                    time.sleep(0.5)
                    inject_capture_button(driver)

                # 新しいウィンドウが開いたかチェック
                current_handles = set(driver.window_handles)
                new_windows = current_handles - previous_handles

                if new_windows:
                    print(f"\n🔔 新しいウィンドウが {len(new_windows)} 個開きました")
                    for handle in new_windows:
                        driver.switch_to.window(handle)
                        time.sleep(1)
                        inject_capture_button(driver)
                        print("新しいウィンドウに撮影ボタンを追加しました")

                    previous_handles = current_handles

                time.sleep(0.3)  # CPU負荷軽減

            except Exception as e:
                # ブラウザが閉じられた場合など
                if "invalid session id" in str(e).lower() or "no such window" in str(e).lower():
                    print("\nブラウザが閉じられました")
                    break
                time.sleep(0.5)

    except KeyboardInterrupt:
        print("\n\n終了します...")

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

            summary_file = os.path.join(session_dir, "00_capture_summary.json")
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
