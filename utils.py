"""ユーティリティ関数"""
import os
import time
import subprocess
from selenium import webdriver
from selenium.webdriver.chrome.options import Options


def setup_driver(download_path):
    """Chromeドライバーをセットアップ"""
    chrome_options = Options()

    # PDFダウンロード設定
    prefs = {
        "download.default_directory": download_path,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True  # PDFを自動ダウンロード
    }
    chrome_options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(options=chrome_options)
    return driver


def download_pdf(driver, download_path):
    """PDFファイルをダウンロード"""
    time.sleep(5)  # ダウンロード完了を待つ

    # ダウンロードされたPDFファイルを取得
    files = [f for f in os.listdir(download_path) if f.endswith('.pdf')]
    if files:
        # 最新のPDFファイルを取得
        latest_file = max([os.path.join(download_path, f) for f in files],
                         key=os.path.getctime)
        return latest_file

    return None


def print_pdf(pdf_path):
    """PDFファイルを印刷"""
    if not pdf_path or not os.path.exists(pdf_path):
        print("印刷するPDFファイルが見つかりません。")
        return False

    try:
        # Macのlprコマンドを使用してPDFを印刷
        # デフォルトプリンタに印刷
        result = subprocess.run(
            ['lpr', pdf_path],
            capture_output=True,
            text=True,
            check=True
        )
        print(f"印刷ジョブを送信しました: {pdf_path}")
        return True
    except subprocess.CalledProcessError as e:
        print(f"印刷エラー: {e}")
        print("プリンタが設定されているか確認してください。")
        return False
    except Exception as e:
        print(f"印刷エラー: {e}")
        return False


def get_default_printer():
    """デフォルトプリンタを取得"""
    try:
        result = subprocess.run(
            ['lpstat', '-d'],
            capture_output=True,
            text=True,
            check=True
        )
        # 出力からプリンタ名を抽出
        output = result.stdout.strip()
        if "system default destination:" in output:
            printer_name = output.split("system default destination:")[-1].strip()
            return printer_name
        return None
    except Exception as e:
        print(f"プリンタ情報取得エラー: {e}")
        return None


def print_pdf_with_printer(pdf_path, printer_name=None):
    """指定されたプリンタでPDFファイルを印刷"""
    if not pdf_path or not os.path.exists(pdf_path):
        print("印刷するPDFファイルが見つかりません。")
        return False

    try:
        if printer_name:
            # 指定されたプリンタで印刷
            result = subprocess.run(
                ['lpr', '-P', printer_name, pdf_path],
                capture_output=True,
                text=True,
                check=True
            )
            print(f"プリンタ '{printer_name}' で印刷ジョブを送信しました: {pdf_path}")
        else:
            # デフォルトプリンタで印刷
            return print_pdf(pdf_path)
        
        return True
    except subprocess.CalledProcessError as e:
        print(f"印刷エラー: {e}")
        print("プリンタが設定されているか確認してください。")
        return False
    except Exception as e:
        print(f"印刷エラー: {e}")
        return False
