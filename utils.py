"""ユーティリティ関数"""
import os
import time
import subprocess
import platform
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
    """PDFファイルを印刷（WindowsとMacの両方に対応）"""
    if not pdf_path or not os.path.exists(pdf_path):
        print("印刷するPDFファイルが見つかりません。")
        return False

    try:
        system = platform.system()

        if system == "Windows":
            # Windowsの場合: Adobe Acrobat Reader、SumatraPDF、またはデフォルトのPDFビューワーで印刷
            try:
                # 方法1: Adobe Acrobat Readerを使用
                acrobat_paths = [
                    r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
                    r"C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
                    r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
                    r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
                ]

                acrobat_exe = None
                for path in acrobat_paths:
                    if os.path.exists(path):
                        acrobat_exe = path
                        break

                if acrobat_exe:
                    # Adobe Acrobat Readerで自動印刷
                    # /t オプション: 印刷後に自動的に閉じる
                    # /h オプション: 最小化して実行
                    result = subprocess.run(
                        [acrobat_exe, '/t', pdf_path],
                        capture_output=True,
                        text=True,
                        timeout=60
                    )
                    print(f"✓ 印刷ジョブを送信しました (Adobe Acrobat): {os.path.basename(pdf_path)}")
                    time.sleep(2)  # Adobe Acrobatが印刷ジョブを送信するまで待機
                    return True

                # 方法2: SumatraPDFを使用（高速で確実）
                sumatra_paths = [
                    r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
                    r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
                ]

                sumatra_exe = None
                for path in sumatra_paths:
                    if os.path.exists(path):
                        sumatra_exe = path
                        break

                if sumatra_exe:
                    # SumatraPDFで自動印刷（-print-to-default オプション）
                    result = subprocess.run(
                        [sumatra_exe, '-print-to-default', '-silent', pdf_path],
                        capture_output=True,
                        text=True,
                        timeout=30
                    )
                    print(f"✓ 印刷ジョブを送信しました (SumatraPDF): {os.path.basename(pdf_path)}")
                    return True

                # 方法3: デフォルトのPDFビューワーで印刷
                # 注意: この方法は印刷ダイアログが表示される可能性がある
                os.startfile(pdf_path, "print")
                print(f"✓ 印刷ジョブを送信しました（デフォルトビューワー）: {os.path.basename(pdf_path)}")
                print("⚠️ 自動印刷には Adobe Acrobat Reader または SumatraPDF のインストールを推奨します")
                return True

            except subprocess.TimeoutExpired:
                print(f"⚠️ 印刷処理がタイムアウトしました: {pdf_path}")
                return False
            except Exception as e:
                print(f"Windows印刷エラー: {e}")
                return False

        elif system == "Darwin":  # Mac
            # Macのlprコマンドを使用してPDFを印刷
            result = subprocess.run(
                ['lpr', pdf_path],
                capture_output=True,
                text=True,
                check=True
            )
            print(f"✓ 印刷ジョブを送信しました: {os.path.basename(pdf_path)}")
            return True

        else:
            print(f"⚠️ 未対応のOS: {system}")
            return False

    except subprocess.CalledProcessError as e:
        print(f"印刷エラー: {e}")
        print("プリンタが設定されているか確認してください。")
        return False
    except Exception as e:
        print(f"印刷エラー: {e}")
        return False


def get_default_printer():
    """デフォルトプリンタを取得（WindowsとMac対応）"""
    try:
        system = platform.system()

        if system == "Windows":
            # Windowsの場合: wmic コマンドを使用
            result = subprocess.run(
                ['wmic', 'printer', 'where', 'default=TRUE', 'get', 'name'],
                capture_output=True,
                text=True,
                check=True
            )
            output = result.stdout.strip().split('\n')
            if len(output) > 1:
                printer_name = output[1].strip()
                return printer_name
            return None

        elif system == "Darwin":  # Mac
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
        else:
            return None

    except Exception as e:
        print(f"プリンタ情報取得エラー: {e}")
        return None


def print_pdf_with_printer(pdf_path, printer_name=None):
    """指定されたプリンタでPDFファイルを印刷（WindowsとMac対応）"""
    if not pdf_path or not os.path.exists(pdf_path):
        print("印刷するPDFファイルが見つかりません。")
        return False

    try:
        system = platform.system()

        if printer_name:
            if system == "Windows":
                # Windows: Adobe Acrobat Reader または SumatraPDFで指定プリンタに印刷

                # 方法1: Adobe Acrobat Readerを使用
                acrobat_paths = [
                    r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
                    r"C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
                    r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
                    r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
                ]

                acrobat_exe = None
                for path in acrobat_paths:
                    if os.path.exists(path):
                        acrobat_exe = path
                        break

                if acrobat_exe:
                    # Adobe Acrobatで指定プリンタに印刷
                    # /t オプションでプリンタ指定: /t <file> <printer>
                    result = subprocess.run(
                        [acrobat_exe, '/t', pdf_path, printer_name],
                        capture_output=True,
                        text=True,
                        timeout=60
                    )
                    print(f"✓ プリンタ '{printer_name}' で印刷ジョブを送信しました (Adobe Acrobat): {os.path.basename(pdf_path)}")
                    time.sleep(2)
                    return True

                # 方法2: SumatraPDFを使用
                sumatra_paths = [
                    r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
                    r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
                ]

                sumatra_exe = None
                for path in sumatra_paths:
                    if os.path.exists(path):
                        sumatra_exe = path
                        break

                if sumatra_exe:
                    result = subprocess.run(
                        [sumatra_exe, '-print-to', printer_name, '-silent', pdf_path],
                        capture_output=True,
                        text=True,
                        timeout=30
                    )
                    print(f"✓ プリンタ '{printer_name}' で印刷ジョブを送信しました (SumatraPDF): {os.path.basename(pdf_path)}")
                    return True

                print("⚠️ Adobe Acrobat Reader または SumatraPDF が見つかりません")
                return False

            elif system == "Darwin":  # Mac
                # 指定されたプリンタで印刷
                result = subprocess.run(
                    ['lpr', '-P', printer_name, pdf_path],
                    capture_output=True,
                    text=True,
                    check=True
                )
                print(f"✓ プリンタ '{printer_name}' で印刷ジョブを送信しました: {os.path.basename(pdf_path)}")
                return True
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
