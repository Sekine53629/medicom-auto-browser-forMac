"""ページ構造調査スクリプト

出庫処理ボタンをクリックした後のページ構造を調査する
"""
import time
from selenium.webdriver.common.by import By


def debug_page_structure(driver, operation_logger=None):
    """現在のページ構造を詳しく調査する

    Args:
        driver: Seleniumドライバー
        operation_logger: ロガー（オプション）
    """
    def log(msg):
        if operation_logger:
            operation_logger.info(msg)
        print(msg)

    log("\n========== ページ構造調査開始 ==========")

    # 現在のURL
    try:
        current_url = driver.current_url
        log(f"現在のURL: {current_url}")
    except Exception as e:
        log(f"URL取得エラー: {e}")

    # ページタイトル
    try:
        title = driver.title
        log(f"ページタイトル: {title}")
    except Exception as e:
        log(f"タイトル取得エラー: {e}")

    # ウィンドウ数
    try:
        windows = driver.window_handles
        log(f"ウィンドウ数: {len(windows)}")
        log(f"現在のウィンドウハンドル: {driver.current_window_handle}")
    except Exception as e:
        log(f"ウィンドウ情報取得エラー: {e}")

    # フレーム構造を確認
    log("\n--- フレーム構造 ---")
    try:
        # デフォルトコンテンツに戻る
        driver.switch_to.default_content()

        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        log(f"iframeの数: {len(iframes)}")

        for i, iframe in enumerate(iframes):
            log(f"\niframe {i}:")
            try:
                log(f"  - id: {iframe.get_attribute('id')}")
                log(f"  - name: {iframe.get_attribute('name')}")
                log(f"  - src: {iframe.get_attribute('src')}")
            except Exception as e:
                log(f"  - 属性取得エラー: {e}")

        # 各フレーム内を調査
        for i, iframe in enumerate(iframes):
            try:
                driver.switch_to.frame(iframe)
                log(f"\n  iframe {i} 内部:")

                # btnRecalcを探す
                try:
                    btn_recalc = driver.find_element(By.ID, "btnRecalc")
                    log(f"    ✓ btnRecalc が見つかりました！")
                    log(f"      - type: {btn_recalc.get_attribute('type')}")
                    log(f"      - value: {btn_recalc.get_attribute('value')}")
                    log(f"      - displayed: {btn_recalc.is_displayed()}")
                except:
                    log(f"    ✗ btnRecalc は見つかりません")

                # btnSyukoを探す
                try:
                    btn_syuko = driver.find_element(By.ID, "btnSyuko")
                    log(f"    ✓ btnSyuko が見つかりました！")
                    log(f"      - type: {btn_syuko.get_attribute('type')}")
                    log(f"      - value: {btn_syuko.get_attribute('value')}")
                    log(f"      - displayed: {btn_syuko.is_displayed()}")
                except:
                    log(f"    ✗ btnSyuko は見つかりません")

                # フレーム内のボタンを全て列挙
                buttons = driver.find_elements(By.TAG_NAME, "input")
                submit_buttons = [b for b in buttons if b.get_attribute('type') in ['submit', 'button']]
                if submit_buttons:
                    log(f"    フレーム内のボタン数: {len(submit_buttons)}")
                    for j, btn in enumerate(submit_buttons[:10]):  # 最初の10個
                        try:
                            log(f"      ボタン {j}: id={btn.get_attribute('id')}, "
                                f"name={btn.get_attribute('name')}, "
                                f"value={btn.get_attribute('value')}")
                        except:
                            pass

                driver.switch_to.default_content()
            except Exception as e:
                log(f"  iframe {i} 調査エラー: {e}")
                driver.switch_to.default_content()

    except Exception as e:
        log(f"フレーム構造調査エラー: {e}")

    # メインコンテンツ（フレーム外）を調査
    log("\n--- メインコンテンツ（フレーム外） ---")
    try:
        driver.switch_to.default_content()

        # btnRecalcを探す
        try:
            btn_recalc = driver.find_element(By.ID, "btnRecalc")
            log(f"✓ btnRecalc が見つかりました（フレーム外）！")
            log(f"  - type: {btn_recalc.get_attribute('type')}")
            log(f"  - value: {btn_recalc.get_attribute('value')}")
            log(f"  - displayed: {btn_recalc.is_displayed()}")
        except:
            log(f"✗ btnRecalc は見つかりません（フレーム外）")

        # btnSyukoを探す
        try:
            btn_syuko = driver.find_element(By.ID, "btnSyuko")
            log(f"✓ btnSyuko が見つかりました（フレーム外）！")
            log(f"  - type: {btn_syuko.get_attribute('type')}")
            log(f"  - value: {btn_syuko.get_attribute('value')}")
            log(f"  - displayed: {btn_syuko.is_displayed()}")
        except:
            log(f"✗ btnSyuko は見つかりません（フレーム外）")

        # 全てのボタンを列挙
        buttons = driver.find_elements(By.TAG_NAME, "input")
        submit_buttons = [b for b in buttons if b.get_attribute('type') in ['submit', 'button']]
        log(f"\nメインコンテンツのボタン数: {len(submit_buttons)}")
        for i, btn in enumerate(submit_buttons[:20]):  # 最初の20個
            try:
                btn_id = btn.get_attribute('id')
                btn_name = btn.get_attribute('name')
                btn_value = btn.get_attribute('value')
                log(f"  ボタン {i}: id={btn_id}, name={btn_name}, value={btn_value}")
            except:
                pass

    except Exception as e:
        log(f"メインコンテンツ調査エラー: {e}")

    log("\n========== ページ構造調査終了 ==========\n")


# operations.pyで使用するためのエクスポート
__all__ = ['debug_page_structure']
