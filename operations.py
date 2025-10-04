"""業務処理関連の機能"""
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def daily_inventory(driver, download_path):
    """毎日在庫処理（月次処理ボタン）"""
    try:
        wait = WebDriverWait(driver, 10)

        # 月次処理ボタンをクリック（画像のsrc属性で特定）
        monthly_button = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//input[@type='image' and contains(@src, '00_getuji.gif')]"))
        )
        print("月次処理ボタンをクリックします...")
        monthly_button.click()

        time.sleep(3)  # ページ遷移を待つ

        # TODO: ここから先の処理（PDFダウンロードなど）を追加
        print("月次処理ページに移動しました。")

        return True
    except Exception as e:
        print(f"毎日在庫処理エラー: {e}")
        return False


def auto_order(driver, download_path):
    """自動発注処理（発注ボタン）"""
    try:
        wait = WebDriverWait(driver, 10)

        # 発注ボタンをクリック（画像のsrc属性で特定）
        order_button = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//input[@type='image' and contains(@src, '00_hattyu.gif')]"))
        )
        print("発注ボタンをクリックします...")
        order_button.click()

        time.sleep(3)  # ページ遷移を待つ

        # TODO: ここから先の処理（PDFダウンロードなど）を追加
        print("発注ページに移動しました。")

        return True
    except Exception as e:
        print(f"自動発注処理エラー: {e}")
        return False
