"""業務処理関連の機能"""
import time
import os
import json
import logging
import re
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from utils import download_pdf, print_pdf


# ログ設定
def setup_logger():
    """操作ログを設定"""
    log_dir = "logs"
    os.makedirs(log_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(log_dir, f"operation_{timestamp}.log")

    # ロガーの設定
    logger = logging.getLogger("operations")
    logger.setLevel(logging.DEBUG)

    # ファイルハンドラ
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)

    # コンソールハンドラ
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)

    # フォーマット
    formatter = logging.Formatter('[%(asctime)s] %(levelname)s: %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger, log_file


# グローバルロガー
operation_logger = None
current_log_file = None


def safe_click(driver, element, description, wait_time=2, logger=None):
    """要素を安全にクリックし、標準待機時間（デフォルト2秒）で待機する

    メッセージ閲覧、伝票作成など通常の操作用。

    Args:
        driver: Seleniumドライバー
        element: クリック対象の要素
        description: 操作の説明（ログ出力用）
        wait_time: クリック後の待機時間（秒）デフォルト2秒
        logger: ロガーオブジェクト

    Returns:
        bool: 成功時True、失敗時False
    """
    try:
        if logger:
            logger.info(f"{description}をクリックします...")
        print(f"{description}をクリックします...")

        element.click()
        time.sleep(wait_time)

        return True
    except Exception as e:
        if logger:
            logger.error(f"{description}のクリックに失敗: {e}")
        print(f"⚠️ {description}のクリックに失敗: {e}")
        return False


def safe_click_heavy(driver, element, description, wait_time=10, logger=None):
    """要素を安全にクリックし、長めの待機時間（デフォルト10秒）で待機する

    発注処理など、100件近い医薬品集計を伴う重い処理用。

    Args:
        driver: Seleniumドライバー
        element: クリック対象の要素
        description: 操作の説明（ログ出力用）
        wait_time: クリック後の待機時間（秒）デフォルト10秒
        logger: ロガーオブジェクト

    Returns:
        bool: 成功時True、失敗時False
    """
    try:
        if logger:
            logger.info(f"{description}をクリックします（重い処理：約{wait_time}秒待機）...")
        print(f"{description}をクリックします（重い処理：約{wait_time}秒待機）...")

        element.click()
        time.sleep(wait_time)

        return True
    except Exception as e:
        if logger:
            logger.error(f"{description}のクリックに失敗: {e}")
        print(f"⚠️ {description}のクリックに失敗: {e}")
        return False


def safe_find_and_click(driver, by, value, description, wait_time=2, logger=None):
    """要素を検索してクリックし、標準待機時間で待機する

    メッセージ閲覧、伝票作成など通常の操作用。

    Args:
        driver: Seleniumドライバー
        by: 検索方法（By.ID, By.XPATH等）
        value: 検索値
        description: 操作の説明（ログ出力用）
        wait_time: クリック後の待機時間（秒）デフォルト2秒
        logger: ロガーオブジェクト

    Returns:
        bool: 成功時True、失敗時False
    """
    try:
        # 明示的な待機を使用
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((by, value)))

        if logger:
            logger.info(f"{description}が見つかりました")

        return safe_click(driver, element, description, wait_time, logger)
    except Exception as e:
        if logger:
            logger.error(f"{description}の要素が見つかりません: {e}")
        print(f"⚠️ {description}の要素が見つかりません: {e}")
        return False


def safe_find_and_click_heavy(driver, by, value, description, wait_time=10, logger=None):
    """要素を検索してクリックし、長めの待機時間で待機する

    発注処理など、100件近い医薬品集計を伴う重い処理用。

    Args:
        driver: Seleniumドライバー
        by: 検索方法（By.ID, By.XPATH等）
        value: 検索値
        description: 操作の説明（ログ出力用）
        wait_time: クリック後の待機時間（秒）デフォルト10秒
        logger: ロガーオブジェクト

    Returns:
        bool: 成功時True、失敗時False
    """
    try:
        element = driver.find_element(by, value)
        return safe_click_heavy(driver, element, description, wait_time, logger)
    except Exception as e:
        if logger:
            logger.error(f"{description}の要素が見つかりません: {e}")
        print(f"⚠️ {description}の要素が見つかりません: {e}")
        return False


def safe_wait(seconds, description=None, logger=None):
    """指定時間待機する（通常用）

    1〜3秒程度の通常待機に使用。

    Args:
        seconds: 待機時間（秒）
        description: 待機の説明（ログ出力用）
        logger: ロガーオブジェクト
    """
    if description and logger:
        logger.debug(f"{description}: {seconds}秒待機")
    time.sleep(seconds)


def safe_wait_heavy(seconds, description=None, logger=None):
    """指定時間待機する（重い処理用）

    10秒以上の長時間待機に使用。発注処理など。

    Args:
        seconds: 待機時間（秒）
        description: 待機の説明（ログ出力用）
        logger: ロガーオブジェクト
    """
    if description and logger:
        logger.info(f"{description}: {seconds}秒待機（重い処理）")
    else:
        print(f"{description}: {seconds}秒待機" if description else f"{seconds}秒待機（重い処理）")
    time.sleep(seconds)


def wait_for_page_load(driver, wait, max_retries=3, retry_delay=5):
    """ページが完全に読み込まれるまで待機（リトライ機能付き）

    Args:
        driver: Seleniumドライバー
        wait: WebDriverWaitオブジェクト
        max_retries: 最大リトライ回数（デフォルト: 3回）
        retry_delay: リトライ間隔（秒、デフォルト: 5秒）

    Returns:
        bool: 成功した場合True、失敗した場合False
    """
    for attempt in range(max_retries):
        try:
            # document.readyStateが'complete'になるまで待機
            wait.until(lambda d: d.execute_script('return document.readyState') == 'complete')

            # さらに少し待機（JavaScriptの実行完了を待つ）
            time.sleep(2)

            print(f"✓ ページ読み込み完了（試行 {attempt + 1}/{max_retries}）")
            return True
        except Exception as e:
            if attempt < max_retries - 1:
                print(f"⚠️ ページ読み込み待機中... 再試行します（{attempt + 1}/{max_retries}）")
                time.sleep(retry_delay)
            else:
                print(f"⚠️ ページ読み込みタイムアウト（{max_retries}回試行後）")
                return False

    return False


def wait_for_shipping_screen(driver, logger=None, timeout=15):
    """出庫画面への遷移を確認する

    出庫処理ボタンをクリックした後、メインウィンドウが出庫画面に切り替わるまで待機する。
    画面の完全な読み込みを確認してからTrueを返す。

    Args:
        driver: Seleniumドライバー
        logger: ロガーオブジェクト
        timeout: タイムアウト時間（秒）

    Returns:
        bool: 出庫画面への遷移が確認できた場合True、タイムアウトの場合False
    """
    try:
        if logger:
            logger.info("出庫画面への遷移を確認中...")

        # 現在のURLを記録
        initial_url = driver.current_url
        if logger:
            logger.debug(f"初期URL: {initial_url}")

        # URLが変わるまで待機（HomeMain.aspxから別のページへ）
        wait = WebDriverWait(driver, timeout)
        try:
            wait.until(lambda d: d.current_url != initial_url)
            if logger:
                logger.info(f"✓ URLが変更されました: {driver.current_url}")
        except:
            if logger:
                logger.warning(f"⚠️ URLが変更されませんでした（タイムアウト{timeout}秒）")
                logger.warning(f"現在のURL: {driver.current_url}")
            return False

        # document.readyStateがcompleteになるまで待機
        wait.until(lambda d: d.execute_script('return document.readyState') == 'complete')

        if logger:
            logger.debug("ページの基本読み込みが完了しました")

        # 追加で3秒待機（JavaScriptやフレームの読み込み完了を待つ）
        time.sleep(3)

        # 出庫画面の特徴的な要素が存在するか確認
        # （btnRecalcやbtnSyukoが存在するか、またはフレーム内に存在するか）
        try:
            # デフォルトコンテンツで確認
            driver.switch_to.default_content()

            # ページ内に btnRecalc または btnSyuko が存在するか確認
            # （フレーム内にある可能性もあるので、存在確認のみ）
            wait_short = WebDriverWait(driver, 3)

            # まずフレーム外を確認
            try:
                wait_short.until(
                    lambda d: len(d.find_elements(By.ID, "btnRecalc")) > 0 or
                             len(d.find_elements(By.ID, "btnSyuko")) > 0
                )
                if logger:
                    logger.info("✓ 出庫画面への遷移を確認しました（フレーム外にボタンを検出）")
                return True
            except:
                # フレーム外に見つからない場合、フレーム内を確認
                iframes = driver.find_elements(By.TAG_NAME, "iframe")
                if logger:
                    logger.debug(f"フレーム数: {len(iframes)}")

                for i, iframe in enumerate(iframes):
                    try:
                        driver.switch_to.frame(iframe)
                        if len(driver.find_elements(By.ID, "btnRecalc")) > 0 or \
                           len(driver.find_elements(By.ID, "btnSyuko")) > 0:
                            if logger:
                                logger.info(f"✓ 出庫画面への遷移を確認しました（フレーム{i}内にボタンを検出）")
                            driver.switch_to.default_content()
                            return True
                        driver.switch_to.default_content()
                    except:
                        driver.switch_to.default_content()
                        continue

                # どのフレームにも見つからなかった場合は失敗
                if logger:
                    logger.error("⚠️ 出庫画面のボタンが見つかりません")
                return False

        except Exception as e:
            if logger:
                logger.error(f"出庫画面の要素確認エラー: {e}")
            return False

    except Exception as e:
        if logger:
            logger.error(f"出庫画面への遷移確認エラー: {e}")
        print(f"⚠️ 出庫画面への遷移確認エラー: {e}")
        return False


def go_back_to_main(driver, wait, max_attempts=3):
    """戻るボタンをクリックしてメインメニューに戻る

    Args:
        driver: Seleniumドライバー
        wait: WebDriverWaitオブジェクト
        max_attempts: 最大試行回数

    Returns:
        bool: 成功した場合True、失敗した場合False
    """
    for attempt in range(max_attempts):
        try:
            print(f"戻るボタンを探しています... (試行 {attempt + 1}/{max_attempts})")

            # 戻るボタンを探す
            if switch_to_frame_with_element(driver, "//input[@value='戻る'] | //input[@type='button' and contains(@value, '戻る')] | //a[contains(text(), '戻る')]"):
                back_button = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//input[@value='戻る'] | //input[@type='button' and contains(@value, '戻る')] | //a[contains(text(), '戻る')]"))
                )
                print("✓ 戻るボタンをクリックします...")
                back_button.click()

                # ページ読み込みを待機
                if wait_for_page_load(driver, wait):
                    print("✓ メインメニューに戻りました")
                    return True
            else:
                print(f"⚠️ 戻るボタンが見つかりません (試行 {attempt + 1}/{max_attempts})")
                time.sleep(2)

        except Exception as e:
            print(f"戻るボタンクリックエラー (試行 {attempt + 1}/{max_attempts}): {e}")
            if attempt < max_attempts - 1:
                time.sleep(2)

    return False


def switch_to_frame_with_element(driver, xpath, timeout=10):
    """指定された要素を含むフレームに切り替える"""
    # まずメインフレームに戻る
    driver.switch_to.default_content()

    # 全フレームを検索
    frames = driver.find_elements(By.TAG_NAME, "frame") + driver.find_elements(By.TAG_NAME, "iframe")

    for frame in frames:
        try:
            driver.switch_to.default_content()
            driver.switch_to.frame(frame)

            # 要素が存在するかチェック
            elements = driver.find_elements(By.XPATH, xpath)
            if elements and len(elements) > 0:
                print(f"要素が見つかりました: フレーム内")
                return True

            # ネストされたフレームもチェック
            nested_frames = driver.find_elements(By.TAG_NAME, "frame") + driver.find_elements(By.TAG_NAME, "iframe")
            for nested_frame in nested_frames:
                try:
                    driver.switch_to.frame(nested_frame)
                    elements = driver.find_elements(By.XPATH, xpath)
                    if elements and len(elements) > 0:
                        print(f"要素が見つかりました: ネストされたフレーム内")
                        return True
                    driver.switch_to.parent_frame()
                except:
                    driver.switch_to.parent_frame()
        except:
            pass

    # フレーム内で見つからなかった場合、メインコンテンツをチェック
    driver.switch_to.default_content()
    elements = driver.find_elements(By.XPATH, xpath)
    if elements and len(elements) > 0:
        print(f"要素が見つかりました: メインコンテンツ")
        return True

    return False


def daily_inventory(driver, download_path, should_print=True):
    """毎日在庫処理（月次処理→棚卸タブ→印刷）

    Args:
        driver: Seleniumドライバー
        download_path: PDFダウンロードパス
        should_print: PDFを印刷するかどうか（デフォルト: True）
    """
    # ログ設定
    operation_logger, log_file_path = setup_logger()
    operation_logger.info(f"ログファイル: {log_file_path}")
    operation_logger.info("============================================================")
    operation_logger.info("毎日在庫処理を開始します")
    operation_logger.info("============================================================")

    try:
        wait = WebDriverWait(driver, 10)

        # ステップ1: 月次処理ボタンをクリック
        print("月次処理ボタンを探しています...")
        if not switch_to_frame_with_element(driver, "//input[@type='image' and contains(@src, '00_getuji.gif')]"):
            print("月次処理ボタンが見つかりません")
            return False

        monthly_button = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//input[@type='image' and contains(@src, '00_getuji.gif')]"))
        )
        print("✓ 月次処理ボタンをクリックします...")
        monthly_button.click()

        # ページが完全に読み込まれるまで待機
        if not wait_for_page_load(driver, wait):
            print("⚠️ ページ読み込みに時間がかかっています。続行しますか？")
            input("Enterキーを押して続行...")

        # ステップ2: 棚卸タブをクリック
        operation_logger.info("棚卸タブを探しています...")
        print("棚卸タブを探しています...")

        # 棚卸タブのXPathパターン
        tab_xpaths = [
            "//div[@id='GYOUMU_ID_8']",  # IDで直接指定
            "//div[contains(@class, 'GyoumuButton')]//a[contains(text(), '棚卸')]/ancestor::div[contains(@class, 'GyoumuButton')]",
            "//div[contains(@class, 'GyoumuButton') and contains(., '棚卸')]",
            "//a[contains(text(), '棚卸')]/ancestor::div[contains(@class, 'GyoumuButton')]",
            "//div[contains(@class, 'GyoumuButton')]/div[contains(text(), '棚卸')]/parent::div"
        ]

        tanaoroshi_tab = None
        found_xpath = None

        for i, xpath in enumerate(tab_xpaths, 1):
            try:
                operation_logger.debug(f"XPathパターン {i}/{len(tab_xpaths)}: {xpath}")
                tanaoroshi_tab = driver.find_element(By.XPATH, xpath)
                if tanaoroshi_tab:
                    operation_logger.info(f"✓ 棚卸タブが見つかりました (パターン {i}: {xpath})")
                    found_xpath = xpath
                    break
            except:
                continue

        if not tanaoroshi_tab:
            operation_logger.error("⚠️ 棚卸タブが見つかりません")
            print("⚠️ 棚卸タブが見つかりません")
            return False

        # 要素の詳細情報をログに記録
        operation_logger.debug("========================================")
        operation_logger.debug("棚卸タブ要素情報:")
        operation_logger.debug(f"  - タグ名: {tanaoroshi_tab.tag_name}")
        operation_logger.debug(f"  - テキスト: {tanaoroshi_tab.text}")
        try:
            operation_logger.debug(f"  - onclick属性: {tanaoroshi_tab.get_attribute('onclick')}")
        except:
            operation_logger.debug("  - onclick属性: None")
        try:
            operation_logger.debug(f"  - href属性: {tanaoroshi_tab.get_attribute('href')}")
        except:
            operation_logger.debug("  - href属性: None")
        try:
            operation_logger.debug(f"  - class属性: {tanaoroshi_tab.get_attribute('class')}")
        except:
            operation_logger.debug("  - class属性: None")
        operation_logger.debug(f"  - 表示状態: {tanaoroshi_tab.is_displayed()}")
        operation_logger.debug("========================================")

        # JavaScriptクリックで実行
        operation_logger.info("棚卸タブをJavaScriptクリックで実行します...")
        driver.execute_script("arguments[0].click();", tanaoroshi_tab)
        operation_logger.info("✓ 棚卸タブのクリックに成功しました")

        # ページ読み込み待機（画面遷移完了を待つ）
        operation_logger.info("画面遷移の完了を待機しています...")
        print("画面遷移の完了を待機しています...")
        if not wait_for_page_load(driver, wait):
            operation_logger.warning("⚠️ ページ読み込みに時間がかかっています")
            print("⚠️ ページ読み込みに時間がかかっています")

        # ステップ3: 備考に"毎日"を入力（リトライ処理付き）
        operation_logger.info("備考フィールドに「毎日」を入力します...")
        print("備考フィールドに「毎日」を入力します...")

        remark_input_success = False
        for attempt in range(3):  # 最大3回リトライ
            try:
                operation_logger.debug(f"備考フィールド入力試行 {attempt + 1}/3")

                # フレーム内の要素を探す
                if not switch_to_frame_with_element(driver, "//input[@id='txtReMark']"):
                    operation_logger.debug(f"フレーム切り替え失敗（試行 {attempt + 1}）")
                    time.sleep(3)
                    continue

                # 備考フィールド（txtReMark）を探す
                remark_field = wait.until(
                    EC.presence_of_element_located((By.ID, "txtReMark"))
                )
                remark_field.clear()
                remark_field.send_keys("毎日")
                operation_logger.info("✓ 備考フィールドに「毎日」を入力しました")
                print("✓ 備考フィールドに「毎日」を入力しました")
                remark_input_success = True
                break
            except Exception as e:
                operation_logger.warning(f"備考フィールド入力エラー（試行 {attempt + 1}/3）: {e}")
                if attempt < 2:
                    print(f"⚠️ 備考フィールドが見つかりません。再試行します...（{attempt + 1}/3）")
                    time.sleep(5)
                else:
                    print(f"⚠️ 備考フィールドが見つからないか、入力できません")

        if not remark_input_success:
            operation_logger.warning("備考フィールドへの入力に失敗しました")

        # ステップ4: 在庫なし薬品を表示チェックボックスをオフにする（リトライ処理付き）
        operation_logger.info("在庫なし薬品を表示チェックボックスを探しています...")
        print("在庫なし薬品を表示チェックボックスを探しています...")

        checkbox_success = False
        for attempt in range(3):  # 最大3回リトライ
            try:
                operation_logger.debug(f"チェックボックス操作試行 {attempt + 1}/3")

                # フレーム内の要素を探す
                if not switch_to_frame_with_element(driver, "//input[@id='chkDISP_ZERO' and @type='checkbox']"):
                    operation_logger.debug(f"フレーム切り替え失敗（試行 {attempt + 1}）")
                    time.sleep(3)
                    continue

                # チェックボックスのXPath
                checkbox_xpath = "//input[@id='chkDISP_ZERO' and @type='checkbox']"
                checkbox = wait.until(
                    EC.presence_of_element_located((By.XPATH, checkbox_xpath))
                )

                # チェック状態を確認
                is_checked = checkbox.is_selected()
                operation_logger.debug(f"チェックボックス状態: checked={is_checked}")

                if is_checked:
                    # チェックを外す
                    checkbox.click()
                    operation_logger.info("✓ 在庫なし薬品を表示チェックボックスをオフにしました")
                    print("✓ 在庫なし薬品を表示チェックボックスをオフにしました")
                else:
                    operation_logger.info("在庫なし薬品を表示チェックボックスは既にオフです")
                    print("在庫なし薬品を表示チェックボックスは既にオフです")

                checkbox_success = True
                break
            except Exception as e:
                operation_logger.warning(f"チェックボックス処理エラー（試行 {attempt + 1}/3）: {e}")
                if attempt < 2:
                    print(f"⚠️ チェックボックスが見つかりません。再試行します...（{attempt + 1}/3）")
                    time.sleep(5)
                else:
                    print(f"⚠️ チェックボックスが見つからないか、操作できません")

        if not checkbox_success:
            operation_logger.warning("チェックボックス操作に失敗しました")

        # ステップ5: 印刷ボタンをクリック（集計完了を待つため、3段階で待機してリトライ）
        operation_logger.info("棚卸集計の完了を待機しています...")
        print("棚卸集計の完了を待機しています...")

        print_button_found = False

        # 第1回試行: 5秒待機
        operation_logger.info("[第1回] 5秒待機してから印刷ボタンを探します...")
        print("[第1回] 5秒待機してから印刷ボタンを探します...")
        safe_wait(5, "第1回印刷ボタン待機", operation_logger)

        if switch_to_frame_with_element(driver, "//input[@value='印刷'] | //input[@type='button' and contains(@onclick, '印刷')] | //a[contains(text(), '印刷')]"):
            print_button_found = True
            operation_logger.info("✓ 印刷ボタンが見つかりました")
            print("✓ 印刷ボタンが見つかりました")
        else:
            # 第2回試行: 30秒待機（集計処理中）
            operation_logger.warning("⚠️ 印刷ボタンが見つかりませんでした。")
            operation_logger.info("[第2回] 30秒待機してから再度探します...")
            print("⚠️ 印刷ボタンが見つかりませんでした。")
            print("[第2回] 30秒待機してから再度探します...")
            safe_wait_heavy(30, "第2回印刷ボタン待機（集計処理中）", operation_logger)

            if switch_to_frame_with_element(driver, "//input[@value='印刷'] | //input[@type='button' and contains(@onclick, '印刷')] | //a[contains(text(), '印刷')]"):
                print_button_found = True
                operation_logger.info("✓ 印刷ボタンが見つかりました")
                print("✓ 印刷ボタンが見つかりました")
            else:
                # 第3回試行: 300秒(5分)待機（重い集計処理）
                operation_logger.warning("⚠️ 印刷ボタンが見つかりませんでした。集計に時間がかかっている可能性があります。")
                operation_logger.info("[第3回] 300秒(5分)待機してから再度探します...")
                print("⚠️ 集計に時間がかかっている可能性があります。")
                print("[第3回] 300秒(5分)待機してから再度探します...")
                safe_wait_heavy(300, "第3回印刷ボタン待機（重い集計処理）", operation_logger)

                if switch_to_frame_with_element(driver, "//input[@value='印刷'] | //input[@type='button' and contains(@onclick, '印刷')] | //a[contains(text(), '印刷')]"):
                    print_button_found = True
                    operation_logger.info("✓ 印刷ボタンが見つかりました")
                    print("✓ 印刷ボタンが見つかりました")
                else:
                    operation_logger.error("⚠️ 5分待機後も印刷ボタンが見つかりませんでした")
                    print("⚠️ 5分待機後も印刷ボタンが見つかりませんでした")
                    return False

        if print_button_found:
            try:
                print_button = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//input[@value='印刷'] | //input[@type='button' and contains(@onclick, '印刷')] | //a[contains(text(), '印刷')]"))
                )
                operation_logger.info("✓ 印刷ボタンをクリックします...")
                print("✓ 印刷ボタンをクリックします...")
                print_button.click()

                # 確認ダイアログ（JavaScript alert）が表示されるまで待機
                operation_logger.info("確認ダイアログの表示を待機しています...")
                print("確認ダイアログの表示を待機しています...")

                alert_handled = False
                for alert_attempt in range(3):  # 最大3回リトライ
                    try:
                        operation_logger.debug(f"確認ダイアログ検索試行 {alert_attempt + 1}/3")
                        time.sleep(2)  # ダイアログ表示待機

                        # JavaScript alertを処理
                        alert = wait.until(EC.alert_is_present())
                        alert_text = alert.text
                        operation_logger.info(f"確認ダイアログ: {alert_text}")
                        print(f"確認ダイアログ: {alert_text}")
                        alert.accept()  # OKをクリック
                        operation_logger.info("✓ 確認ダイアログでOKをクリックしました")
                        print("✓ 確認ダイアログでOKをクリックしました")
                        alert_handled = True
                        break
                    except Exception as e:
                        operation_logger.debug(f"確認ダイアログ検索エラー（試行 {alert_attempt + 1}/3）: {e}")
                        if alert_attempt < 2:
                            print(f"⚠️ 確認ダイアログを探しています...（{alert_attempt + 1}/3）")
                        else:
                            operation_logger.warning("確認ダイアログが見つかりませんでした")
                            print("⚠️ 確認ダイアログが見つかりません。")

                if not alert_handled:
                    operation_logger.warning("確認ダイアログの処理に失敗しました")

                # PDFダウンロード・印刷処理
                time.sleep(3)
                if download_pdf(driver, download_path):
                    if should_print:
                        print_pdf(download_path)
                    print("✓ PDFダウンロード・印刷が完了しました")
                else:
                    print("⚠️ PDFのダウンロードに失敗しました")

            except Exception as e:
                print(f"印刷ボタンクリックエラー: {e}")

        # ウィンドウ整理：about:blank や印刷ページを閉じる
        try:
            operation_logger.info("ウィンドウ状態を確認しています...")
            print("ウィンドウ状態を確認しています...")
            time.sleep(3)  # ウィンドウが開くまで待機

            initial_windows = driver.window_handles
            operation_logger.info(f"初期ウィンドウ数: {len(initial_windows)}")
            print(f"現在のウィンドウ数: {len(initial_windows)}")

            # 全ウィンドウの詳細情報を収集
            window_info = []
            for i, window in enumerate(initial_windows, 1):
                try:
                    driver.switch_to.window(window)
                    url = driver.current_url
                    title = driver.title
                    info = {
                        'handle': window,
                        'index': i,
                        'url': url,
                        'title': title,
                        'is_medicom': "medicom" in url.lower() or "Medicom" in title
                    }
                    window_info.append(info)
                    operation_logger.info(f"ウィンドウ {i}: URL={url[:80]}, タイトル={title[:50]}, Medicom={info['is_medicom']}")
                    print(f"  ウィンドウ {i}: {'[Medicom]' if info['is_medicom'] else '[その他]'} {url[:50]}")
                except Exception as e:
                    operation_logger.debug(f"ウィンドウ {i} アクセスエラー: {e}")
                    continue

            # メインウィンドウ（Medicom）を特定
            main_windows = [w for w in window_info if w['is_medicom']]
            other_windows = [w for w in window_info if not w['is_medicom']]

            operation_logger.info(f"Medicomウィンドウ: {len(main_windows)}個")
            operation_logger.info(f"その他のウィンドウ: {len(other_windows)}個")

            if main_windows:
                main_window = main_windows[0]['handle']
                operation_logger.info(f"メインウィンドウ: {main_windows[0]['url']}")
            else:
                main_window = None
                operation_logger.warning("⚠️ Medicomウィンドウが見つかりません")

            # Medicom以外の全ウィンドウを閉じる（about:blank、印刷ページなど）
            closed_count = 0
            for window in other_windows:
                try:
                    driver.switch_to.window(window['handle'])
                    operation_logger.info(f"余分なウィンドウを閉じます: {window['url'][:80]}")
                    driver.close()
                    closed_count += 1
                    print(f"✓ ウィンドウを閉じました: {window['url'][:50]}")
                except Exception as e:
                    operation_logger.debug(f"ウィンドウクローズエラー（既に閉じられている）: {e}")

            # メインウィンドウに戻る
            if main_window:
                driver.switch_to.window(main_window)
                operation_logger.info("✓ メインウィンドウに戻りました")

            final_count = len(driver.window_handles)
            operation_logger.info(f"✓ ウィンドウ整理完了。{closed_count}個のウィンドウを閉じました。残り: {final_count}個")
            print(f"✓ ウィンドウ整理完了（{closed_count}個のウィンドウを閉じました）")
        except Exception as e:
            operation_logger.warning(f"ウィンドウ整理エラー: {e}")
            print(f"⚠️ ウィンドウ整理でエラーが発生しましたが、処理を継続します")
            # エラーが発生してもメインウィンドウには戻る
            try:
                remaining = driver.window_handles
                if remaining:
                    driver.switch_to.window(remaining[0])
            except:
                pass

        # 戻るボタンをクリックしてメインメニューに戻る
        print("メインメニューに戻ります...")
        if not go_back_to_main(driver, wait):
            print("⚠️ 自動で戻れませんでした。手動で戻ってください。")
            input("メインメニューに戻ったらEnterキーを押してください...")

        operation_logger.info("毎日在庫処理が正常に完了しました")
        operation_logger.info(f"ログファイル: {log_file_path}")
        return True

    except Exception as e:
        print(f"毎日在庫処理エラー: {e}")
        import traceback
        traceback.print_exc()
        return False
    finally:
        # フレームをリセット
        try:
            driver.switch_to.default_content()
        except:
            pass


def auto_order(driver, download_path, should_print=True):
    """自動発注処理（発注ボタン→発注表示ボタン→印刷→発注実行）

    Args:
        driver: Seleniumドライバー
        download_path: PDFダウンロードパス
        should_print: PDFを印刷するかどうか（デフォルト: True）
    """
    # ログ設定
    operation_logger, log_file_path = setup_logger()
    operation_logger.info(f"ログファイル: {log_file_path}")
    operation_logger.info("============================================================")
    operation_logger.info("自動発注処理を開始します")
    operation_logger.info("============================================================")

    try:
        wait = WebDriverWait(driver, 10)

        # ステップ1: 発注ボタンをクリック
        operation_logger.info("発注ボタンを探しています...")
        print("発注ボタンを探しています...")
        if not switch_to_frame_with_element(driver, "//input[@type='image' and contains(@src, '00_hattyu.gif')]"):
            operation_logger.error("発注ボタンが見つかりません")
            print("発注ボタンが見つかりません")
            return False

        order_button = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//input[@type='image' and contains(@src, '00_hattyu.gif')]"))
        )
        operation_logger.info("✓ 発注ボタンをクリックします...")
        print("✓ 発注ボタンをクリックします...")
        order_button.click()

        # ページが完全に読み込まれるまで待機
        if not wait_for_page_load(driver, wait):
            operation_logger.warning("⚠️ ページ読み込みに時間がかかっています")
            print("⚠️ ページ読み込みに時間がかかっています。続行しますか？")
            input("Enterキーを押して続行...")

        # ステップ2: 発注表示ボタンをクリック
        operation_logger.info("発注表示ボタンを探しています...")
        print("発注表示ボタンを探しています...")
        if not switch_to_frame_with_element(driver, "//input[@value='発注表示'] | //input[@type='button' and contains(@value, '発注表示')] | //*[contains(text(), '発注表示')]"):
            operation_logger.warning("⚠️ 発注表示ボタンが見つかりません")
            print("⚠️ 発注表示ボタンが見つかりません。手動で発注表示ボタンをクリックしてください。")
            print("準備ができたらEnterキーを押してください...")
            input()

            # 再度発注表示ボタンを探す
            if not switch_to_frame_with_element(driver, "//input[@value='発注表示'] | //input[@type='button' and contains(@value, '発注表示')] | //*[contains(text(), '発注表示')]"):
                operation_logger.warning("⚠️ 発注表示ボタンが見つかりません。次のステップに進みます。")
                print("⚠️ 発注表示ボタンが見つかりません。次のステップに進みます。")
            else:
                try:
                    display_button = driver.find_element(By.XPATH, "//input[@value='発注表示'] | //input[@type='button' and contains(@value, '発注表示')] | //*[contains(text(), '発注表示')]")
                    operation_logger.info("✓ 発注表示ボタンをクリックします...")
                    print("✓ 発注表示ボタンをクリックします...")
                    display_button.click()
                    time.sleep(3)
                except Exception as e:
                    operation_logger.error(f"発注表示ボタンクリックエラー: {e}")
                    print(f"発注表示ボタンクリックエラー: {e}")
        else:
            try:
                display_button = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//input[@value='発注表示'] | //input[@type='button' and contains(@value, '発注表示')] | //*[contains(text(), '発注表示')]"))
                )
                operation_logger.info("✓ 発注表示ボタンをクリックします...")
                print("✓ 発注表示ボタンをクリックします...")
                display_button.click()

                # ページが完全に読み込まれるまで待機
                if not wait_for_page_load(driver, wait):
                    operation_logger.warning("⚠️ ページ読み込みに時間がかかっています")
                    print("⚠️ ページ読み込みに時間がかかっています。続行しますか？")
                    input("Enterキーを押して続行...")
            except Exception as e:
                operation_logger.error(f"発注表示ボタンクリックエラー: {e}")
                print(f"発注表示ボタンクリックエラー: {e}")
                print("手動で発注表示ボタンをクリックしてください。")
                input("準備ができたらEnterキーを押してください...")

        # ステップ3: 印刷ボタンをクリック（発注前にリストを印刷）
        # 集計完了を待つため、3段階で待機してリトライ
        operation_logger.info("発注集計の完了を待機しています...")
        print("発注集計の完了を待機しています...")

        print_button_found = False

        # 第1回試行: 5秒待機
        operation_logger.info("[第1回] 5秒待機してから印刷ボタンを探します...")
        print("[第1回] 5秒待機してから印刷ボタンを探します...")
        safe_wait(5, "第1回印刷ボタン待機", operation_logger)

        if switch_to_frame_with_element(driver, "//input[@value='印刷'] | //input[@type='button' and contains(@onclick, '印刷')] | //a[contains(text(), '印刷')]"):
            print_button_found = True
            operation_logger.info("✓ 印刷ボタンが見つかりました")
            print("✓ 印刷ボタンが見つかりました")
        else:
            # 第2回試行: 30秒待機（集計処理中）
            operation_logger.warning("⚠️ 印刷ボタンが見つかりませんでした。")
            operation_logger.info("[第2回] 30秒待機してから再度探します...")
            print("⚠️ 印刷ボタンが見つかりませんでした。")
            print("[第2回] 30秒待機してから再度探します...")
            safe_wait_heavy(30, "第2回印刷ボタン待機（集計処理中）", operation_logger)

            if switch_to_frame_with_element(driver, "//input[@value='印刷'] | //input[@type='button' and contains(@onclick, '印刷')] | //a[contains(text(), '印刷')]"):
                print_button_found = True
                operation_logger.info("✓ 印刷ボタンが見つかりました")
                print("✓ 印刷ボタンが見つかりました")
            else:
                # 第3回試行: 300秒(5分)待機（重い集計処理）
                operation_logger.warning("⚠️ 印刷ボタンが見つかりませんでした。集計に時間がかかっている可能性があります。")
                operation_logger.info("[第3回] 300秒(5分)待機してから再度探します...")
                print("⚠️ 集計に時間がかかっている可能性があります。")
                print("[第3回] 300秒(5分)待機してから再度探します...")
                safe_wait_heavy(300, "第3回印刷ボタン待機（重い集計処理）", operation_logger)

                if switch_to_frame_with_element(driver, "//input[@value='印刷'] | //input[@type='button' and contains(@onclick, '印刷')] | //a[contains(text(), '印刷')]"):
                    print_button_found = True
                    operation_logger.info("✓ 印刷ボタンが見つかりました")
                    print("✓ 印刷ボタンが見つかりました")
                else:
                    operation_logger.error("⚠️ 5分待機後も印刷ボタンが見つかりませんでした")
                    print("⚠️ 5分待機後も印刷ボタンが見つかりませんでした")
                    return False

        if print_button_found:
            try:
                print_button = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//input[@value='印刷'] | //input[@type='button' and contains(@onclick, '印刷')] | //a[contains(text(), '印刷')]"))
                )
                operation_logger.info("✓ 印刷ボタンをクリックします...")
                print("✓ 印刷ボタンをクリックします...")
                print_button.click()

                # 確認ダイアログ（JavaScript alert）が表示されるまで待機
                operation_logger.info("確認ダイアログの表示を待機しています...")
                print("確認ダイアログの表示を待機しています...")

                alert_handled = False
                for alert_attempt in range(3):  # 最大3回リトライ
                    try:
                        operation_logger.debug(f"確認ダイアログ検索試行 {alert_attempt + 1}/3")
                        time.sleep(2)  # ダイアログ表示待機

                        # JavaScript alertを処理
                        alert = wait.until(EC.alert_is_present())
                        alert_text = alert.text
                        operation_logger.info(f"確認ダイアログ: {alert_text}")
                        print(f"確認ダイアログ: {alert_text}")
                        alert.accept()  # OKをクリック
                        operation_logger.info("✓ 確認ダイアログでOKをクリックしました")
                        print("✓ 確認ダイアログでOKをクリックしました")
                        alert_handled = True
                        break
                    except Exception as e:
                        operation_logger.debug(f"確認ダイアログ検索エラー（試行 {alert_attempt + 1}/3）: {e}")
                        if alert_attempt < 2:
                            print(f"⚠️ 確認ダイアログを探しています...（{alert_attempt + 1}/3）")
                        else:
                            operation_logger.warning("確認ダイアログが見つかりませんでした")
                            print("⚠️ 確認ダイアログが見つかりません。")

                if not alert_handled:
                    operation_logger.warning("確認ダイアログの処理に失敗しました")

                # PDFダウンロード・印刷処理
                time.sleep(3)
                if download_pdf(driver, download_path):
                    if should_print:
                        print_pdf(download_path)
                    print("✓ PDFダウンロード・印刷が完了しました")
                else:
                    print("⚠️ PDFのダウンロードに失敗しました")

            except Exception as e:
                print(f"印刷ボタンクリックエラー: {e}")

        # ウィンドウ整理：about:blank や印刷ページを閉じる
        try:
            operation_logger.info("ウィンドウ状態を確認しています...")
            print("ウィンドウ状態を確認しています...")
            time.sleep(3)  # ウィンドウが開くまで待機

            initial_windows = driver.window_handles
            operation_logger.info(f"初期ウィンドウ数: {len(initial_windows)}")
            print(f"現在のウィンドウ数: {len(initial_windows)}")

            # 全ウィンドウの詳細情報を収集
            window_info = []
            for i, window in enumerate(initial_windows, 1):
                try:
                    driver.switch_to.window(window)
                    url = driver.current_url
                    title = driver.title
                    info = {
                        'handle': window,
                        'index': i,
                        'url': url,
                        'title': title,
                        'is_medicom': "medicom" in url.lower() or "Medicom" in title
                    }
                    window_info.append(info)
                    operation_logger.info(f"ウィンドウ {i}: URL={url[:80]}, タイトル={title[:50]}, Medicom={info['is_medicom']}")
                    print(f"  ウィンドウ {i}: {'[Medicom]' if info['is_medicom'] else '[その他]'} {url[:50]}")
                except Exception as e:
                    operation_logger.debug(f"ウィンドウ {i} アクセスエラー: {e}")
                    continue

            # メインウィンドウ（Medicom）を特定
            main_windows = [w for w in window_info if w['is_medicom']]
            other_windows = [w for w in window_info if not w['is_medicom']]

            operation_logger.info(f"Medicomウィンドウ: {len(main_windows)}個")
            operation_logger.info(f"その他のウィンドウ: {len(other_windows)}個")

            if main_windows:
                main_window = main_windows[0]['handle']
                operation_logger.info(f"メインウィンドウ: {main_windows[0]['url']}")
            else:
                main_window = None
                operation_logger.warning("⚠️ Medicomウィンドウが見つかりません")

            # Medicom以外の全ウィンドウを閉じる（about:blank、印刷ページなど）
            closed_count = 0
            for window in other_windows:
                try:
                    driver.switch_to.window(window['handle'])
                    operation_logger.info(f"余分なウィンドウを閉じます: {window['url'][:80]}")
                    driver.close()
                    closed_count += 1
                    print(f"✓ ウィンドウを閉じました: {window['url'][:50]}")
                except Exception as e:
                    operation_logger.debug(f"ウィンドウクローズエラー（既に閉じられている）: {e}")

            # メインウィンドウに戻る
            if main_window:
                driver.switch_to.window(main_window)
                operation_logger.info("✓ メインウィンドウに戻りました")

            final_count = len(driver.window_handles)
            operation_logger.info(f"✓ ウィンドウ整理完了。{closed_count}個のウィンドウを閉じました。残り: {final_count}個")
            print(f"✓ ウィンドウ整理完了（{closed_count}個のウィンドウを閉じました）")
        except Exception as e:
            operation_logger.warning(f"ウィンドウ整理エラー: {e}")
            print(f"⚠️ ウィンドウ整理でエラーが発生しましたが、処理を継続します")
            # エラーが発生してもメインウィンドウには戻る
            try:
                remaining = driver.window_handles
                if remaining:
                    driver.switch_to.window(remaining[0])
            except:
                pass

        else:
            operation_logger.warning("⚠️ PDFのダウンロードに失敗しました")
            print("⚠️ PDFのダウンロードに失敗しました")

        # ステップ4: 薬品リストの一番下にある発注ボタンをクリック（印刷後に発注実行）
        operation_logger.info("発注ボタン（薬品リスト下部）を探しています...")
        print("発注ボタン（薬品リスト下部）を探しています...")

        try:
            # 印刷処理後、ページが安定するまで待機
            time.sleep(2)

            # メインウィンドウに戻る（念のため）
            driver.switch_to.default_content()
            operation_logger.info("メインフレームに切り替えました")

            # 現在のURLを確認
            current_url = driver.current_url
            operation_logger.info(f"現在のURL: {current_url}")
            print(f"現在のURL: {current_url[:80]}")

            # 発注ボタンを含むフレームを探す
            operation_logger.info("発注ボタンを含むフレームを探しています...")
            frames = driver.find_elements(By.TAG_NAME, "frame") + driver.find_elements(By.TAG_NAME, "iframe")
            operation_logger.info(f"フレーム数: {len(frames)}")

            # 発注ボタンのXPathパターン（id="btnHatyu" または value="発注する"）
            hatyu_button_xpaths = [
                "//input[@id='btnHatyu']",
                "//input[@name='btnHatyu']",
                "//input[@type='submit' and @value='発注する']",
                "//input[@type='submit' and contains(@value, '発注')]"
            ]

            hatyu_button = None
            found_frame = None

            # まずメインコンテンツで検索
            for xpath in hatyu_button_xpaths:
                try:
                    hatyu_button = driver.find_element(By.XPATH, xpath)
                    if hatyu_button:
                        operation_logger.info(f"✓ 発注ボタンが見つかりました（メインコンテンツ）: {xpath}")
                        print(f"✓ 発注ボタンが見つかりました（メインコンテンツ）")
                        found_frame = "main"
                        break
                except:
                    continue

            # メインで見つからなければフレーム内を検索
            if not hatyu_button:
                for i, frame in enumerate(frames):
                    try:
                        driver.switch_to.default_content()
                        driver.switch_to.frame(frame)
                        operation_logger.debug(f"フレーム {i+1} を確認中...")

                        for xpath in hatyu_button_xpaths:
                            try:
                                hatyu_button = driver.find_element(By.XPATH, xpath)
                                if hatyu_button:
                                    operation_logger.info(f"✓ 発注ボタンが見つかりました（フレーム {i+1}）: {xpath}")
                                    print(f"✓ 発注ボタンが見つかりました（フレーム {i+1}）")
                                    found_frame = f"frame_{i}"
                                    break
                            except:
                                continue

                        if hatyu_button:
                            break

                    except Exception as e:
                        operation_logger.debug(f"フレーム {i+1} アクセスエラー: {e}")
                        driver.switch_to.default_content()
                        continue

            if not hatyu_button:
                operation_logger.warning("⚠️ 発注ボタンが見つかりません")
                print("⚠️ 発注ボタン（薬品リスト下部）が見つかりません。手動で発注ボタンをクリックしてください。")
                print("準備ができたらEnterキーを押してください...")
                input()

                # 再度探す（ユーザーがページを操作した可能性）
                driver.switch_to.default_content()
                for xpath in hatyu_button_xpaths:
                    try:
                        hatyu_button = driver.find_element(By.XPATH, xpath)
                        if hatyu_button:
                            operation_logger.info(f"✓ 発注ボタンが見つかりました（再試行）: {xpath}")
                            break
                    except:
                        continue

                if not hatyu_button:
                    operation_logger.error("⚠️ 発注ボタンが見つかりません")
                    print("⚠️ 発注ボタンが見つかりません")
                    return False

            # ボタンまでスクロール（ページ下部にあるため）
            operation_logger.info("発注ボタンまでスクロールしています...")
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", hatyu_button)
            time.sleep(1)

            # クリック
            operation_logger.info("発注ボタンをクリックします...")
            print("✓ 発注ボタン（薬品リスト下部）をクリックします...")
            hatyu_button.click()

            # 確認ダイアログ（"発注しても宜しいですか？"）を処理
            time.sleep(1)
            try:
                alert = driver.switch_to.alert
                alert_text = alert.text
                operation_logger.info(f"確認ダイアログ: {alert_text}")
                print(f"確認ダイアログ: {alert_text}")
                alert.accept()  # OKをクリック
                operation_logger.info("✓ 確認ダイアログでOKをクリックしました")
                print("✓ 確認ダイアログでOKをクリックしました")
            except:
                operation_logger.debug("確認ダイアログなし、またはすでに処理済み")

            # ページが完全に読み込まれるまで待機
            if not wait_for_page_load(driver, wait):
                print("⚠️ ページ読み込みに時間がかかっています。続行しますか？")
                input("Enterキーを押して続行...")

            operation_logger.info("✓ 発注処理が完了しました")
            print("✓ 発注処理が完了しました")

        except Exception as e:
            operation_logger.error(f"発注ボタンクリックエラー: {e}")
            print(f"発注ボタンクリックエラー: {e}")
            print("手動で発注ボタンをクリックしてください。")
            input("準備ができたらEnterキーを押してください...")

        # 戻るボタンをクリックしてメインメニューに戻る
        operation_logger.info("メインメニューに戻ります...")
        print("メインメニューに戻ります...")
        if not go_back_to_main(driver, wait):
            operation_logger.warning("⚠️ 自動で戻れませんでした")
            print("⚠️ 自動で戻れませんでした。手動で戻ってください。")
            input("メインメニューに戻ったらEnterキーを押してください...")

        operation_logger.info("自動発注処理が正常に完了しました")
        operation_logger.info(f"ログファイル: {log_file_path}")
        return True

    except Exception as e:
        operation_logger.error(f"自動発注処理エラー: {e}")
        print(f"自動発注処理エラー: {e}")
        import traceback
        traceback.print_exc()
        return False
    finally:
        # フレームをリセット
        try:
            driver.switch_to.default_content()
        except:
            pass


# ===========================
# 不動在庫転送機能
# ===========================

def load_immobile_stock(store_id):
    """不動医薬品リストのJSON読み込み（店舗IDごと）

    Args:
        store_id: 店舗ID（4桁）

    Returns:
        dict: 不動医薬品リストデータ
    """
    stock_file = f"data/immobile_stock_{store_id}.json"

    # ディレクトリが存在しない場合は作成
    os.makedirs("data", exist_ok=True)

    if os.path.exists(stock_file):
        with open(stock_file, 'r', encoding='utf-8') as f:
            return json.load(f)
    else:
        return {"medicines": []}


def save_immobile_stock(data, store_id):
    """不動医薬品リストのJSON保存（店舗IDごと）

    Args:
        data: 保存するデータ
        store_id: 店舗ID（4桁）
    """
    stock_file = f"data/immobile_stock_{store_id}.json"

    # ディレクトリが存在しない場合は作成
    os.makedirs("data", exist_ok=True)

    with open(stock_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def add_medicine_to_immobile_stock(store_id, medicine_data):
    """不動医薬品リストに薬品を追加

    Args:
        store_id: 店舗ID（4桁）
        medicine_data: 薬品データ（辞書形式）
            {
                'medicine_name': 薬品名,
                'quantity': 数量,
                'unit': 単位,
                'lot_number': 製造番号（ロット番号）,
                'expiry_date': 使用期限,
                'source_message_id': 元メッセージID,
                'target_stores': [
                    {
                        'store_id': 送り先店舗ID,
                        'store_name': 送り先店舗名,
                        'status': 'pending' | 'accepted' | 'rejected',
                        'sent_at': 送信日時,
                        'responded_at': 返信日時
                    }
                ]
            }

    Returns:
        bool: 追加成功時True
    """
    try:
        # 不動在庫リストを読み込み
        immobile_stock = load_immobile_stock(store_id)

        # 新しい薬品データに必要なフィールドを追加
        new_medicine = {
            'medicine_id': f"{store_id}_{datetime.now().strftime('%Y%m%d%H%M%S')}",
            'medicine_name': medicine_data.get('medicine_name'),
            'quantity': medicine_data.get('quantity'),
            'unit': medicine_data.get('unit'),
            'lot_number': medicine_data.get('lot_number', ''),
            'expiry_date': medicine_data.get('expiry_date'),
            'source_message_id': medicine_data.get('source_message_id', ''),
            'status': 'active',  # active, completed, cancelled
            'created_at': datetime.now().isoformat(),
            'target_stores': []  # 送り先店舗リスト（後で追加）
        }

        # リストに追加
        immobile_stock['medicines'].append(new_medicine)

        # 保存
        save_immobile_stock(immobile_stock, store_id)

        return True

    except Exception as e:
        print(f"不動医薬品リストへの追加エラー: {e}")
        return False


def update_target_store_status(store_id, medicine_id, target_store_id, status, message_id=None):
    """送り先店舗の受け入れ可否ステータスを更新

    Args:
        store_id: 自店舗ID（4桁）
        medicine_id: 薬品ID
        target_store_id: 送り先店舗ID
        status: ステータス（'accepted' | 'rejected'）
        message_id: 返信メッセージID（オプション）

    Returns:
        bool: 更新成功時True
    """
    try:
        # 不動在庫リストを読み込み
        immobile_stock = load_immobile_stock(store_id)

        # 該当する薬品を探す
        for medicine in immobile_stock['medicines']:
            if medicine['medicine_id'] == medicine_id:
                # 該当する送り先店舗を探す
                for target in medicine['target_stores']:
                    if target['store_id'] == target_store_id:
                        # ステータスを更新
                        target['status'] = status
                        target['responded_at'] = datetime.now().isoformat()
                        if message_id:
                            target['response_message_id'] = message_id

                        # 保存
                        save_immobile_stock(immobile_stock, store_id)
                        return True

        return False

    except Exception as e:
        print(f"ステータス更新エラー: {e}")
        return False


# ===========================
# 連絡板関連の機能
# ===========================

def extract_store_id(user_id):
    """user_idから店舗ID（4桁）を抽出

    Args:
        user_id: ユーザーID（例: TRH170501）

    Returns:
        str: 店舗ID 4桁（例: 1705）
    """
    # TRHを除去し、最後の2桁を除外した4桁を取得
    # 例: TRH170501 → 170501 → 1705
    match = re.search(r'TRH(\d{4})\d{2}', user_id)
    if match:
        return match.group(1)
    else:
        # パターンにマッチしない場合はuser_idをそのまま使用
        return user_id


def load_message_stock(store_id):
    """メッセージストックのJSON読み込み（店舗IDごと）

    Args:
        store_id: 店舗ID（4桁）
    """
    stock_file = f"data/message_stock_{store_id}.json"

    # ディレクトリが存在しない場合は作成
    os.makedirs("data", exist_ok=True)

    if os.path.exists(stock_file):
        with open(stock_file, 'r', encoding='utf-8') as f:
            return json.load(f)
    else:
        return {"messages": []}


def save_message_stock(data, store_id):
    """メッセージストックのJSON保存（店舗IDごと）

    Args:
        data: 保存するデータ
        store_id: 店舗ID（4桁）
    """
    stock_file = f"data/message_stock_{store_id}.json"

    # ディレクトリが存在しない場合は作成
    os.makedirs("data", exist_ok=True)

    with open(stock_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def parse_message_content(content):
    """メッセージ本文をパースして必要な情報を抽出

    Args:
        content: メッセージ本文（テキスト）

    Returns:
        list: 抽出された医薬品情報のリスト（複数ロット対応）
              各要素は dict: {送信店舗名、医薬品名、数量、単位、使用期限}
              パース失敗時は空リスト
    """
    try:
        # 送信店舗名の抽出（***で囲まれた部分から正式名称を取得）
        # 例: "*  ウォンツ薬局　日赤病院前店"
        store_match = re.search(r'\*+\s*\n\s*\*\s+(.+?)\s*\n', content)
        if not store_match:
            # 旧パターン: "調剤薬局ツルハドラッグ..."
            store_match = re.search(r'調剤薬局(.+?)[\s\*\n]', content)

        sender_store = store_match.group(1).strip() if store_match else None

        # 薬品情報の抽出（--------で囲まれた部分全体）
        # 全てのダッシュ行で区切られたセクションを取得
        dash_sections = re.findall(r'-{10,}\n(.+?)(?=\n-{10,}|$)', content, re.DOTALL)

        if not dash_sections or len(dash_sections) < 2:
            return []

        # 2番目のセクション（ヘッダーの次）が実際のデータ
        medicine_text = dash_sections[1].strip() if len(dash_sections) > 1 else dash_sections[0].strip()

        # 薬品名と数量の抽出（複数ロット対応）
        lines = medicine_text.split('\n')

        # ヘッダー行をスキップするためのフラグ
        header_keywords = ['薬品名', '数量', '単位', '使用期限', '余剰在庫']

        results = []
        current_medicine = None

        i = 0
        while i < len(lines):
            line = lines[i].strip()
            if not line or any(keyword in line for keyword in header_keywords):
                i += 1
                continue

            # 薬品名と数量が同じ行にある場合
            qty_match = re.search(r'(\d+\.?\d*)\s+(\S+?)\s*$', line)
            if qty_match:
                try:
                    quantity_str = qty_match.group(1)
                    unit_str = qty_match.group(2).strip()

                    # 単位は通常1-10文字程度
                    if len(unit_str) <= 10:
                        medicine_name = line[:qty_match.start()].strip()
                        quantity = float(quantity_str)
                        unit = unit_str

                        # 次の行に使用期限があるかチェック
                        expiry_date = None
                        if i + 1 < len(lines):
                            next_line = lines[i + 1].strip()
                            if '/' in next_line and re.match(r'\s*\d{4}/\d{2}', next_line):
                                expiry_match = re.search(r'(\d{4}/\d{2})', next_line)
                                if expiry_match:
                                    expiry_date = expiry_match.group(1)
                                    i += 1  # 期限行をスキップ

                        # 結果に追加
                        results.append({
                            'sender_store': sender_store,
                            'medicine_name': medicine_name,
                            'quantity': quantity,
                            'unit': unit,
                            'expiry_date': expiry_date
                        })
                        current_medicine = medicine_name
                except Exception:
                    pass

            i += 1

        return results

    except Exception as e:
        print(f"メッセージパースエラー: {e}")
        return []


def check_messages(driver, user_id, config=None):
    """連絡板の未読メッセージを確認（連続処理）

    Args:
        driver: Seleniumドライバー
        user_id: ユーザーID（店舗ID抽出に使用）
        config: 設定情報（タイトルスキップ設定、最大処理件数を含む）
    """
    # config未指定の場合はデフォルト値を使用
    if config is None:
        config = {
            "message_processing": {
                "購入伺い": True,
                "マッチング：使用期限": True,
                "不動在庫転送": True,
                "返信": True
            },
            "max_message_count": 10
        }

    # 最大処理件数を取得
    max_message_count = config.get('max_message_count', 10)

    # ログ設定
    operation_logger, log_file_path = setup_logger()
    operation_logger.info(f"ログファイル: {log_file_path}")
    operation_logger.info("============================================================")
    operation_logger.info(f"連絡板メッセージ確認処理を開始します（最大{max_message_count}件連続処理）")
    operation_logger.info("============================================================")

    # メッセージ処理設定を表示
    msg_processing = config.get('message_processing', {})
    operation_logger.info(f"メッセージ処理設定: {msg_processing}")

    # 処理可能なメールタイトルとその設定
    enabled_count = sum(1 for enabled in msg_processing.values() if enabled)
    disabled_count = len(msg_processing) - enabled_count

    if enabled_count > 0:
        print(f"処理対象メール: {enabled_count}種類 (スキップ: {disabled_count}種類)")
    else:
        print("\n⚠️ すべてのメールタイプがスキップ設定されています")
        print("設定メニューから処理するメールを選択してください")
        return False

    try:
        wait = WebDriverWait(driver, 10)

        # 店舗IDを抽出
        store_id = extract_store_id(user_id)
        operation_logger.info(f"ユーザーID: {user_id}")
        operation_logger.info(f"店舗ID: {store_id}")
        print(f"\n店舗ID: {store_id}")

        # メッセージストックを読み込み（店舗IDごと）
        message_stock = load_message_stock(store_id)
        operation_logger.info(f"現在のストック数: {len(message_stock['messages'])}")

        # 受信一覧フレームに切り替え
        operation_logger.info("受信一覧フレームに切り替えます...")
        print("受信一覧フレームに切り替えます...")

        try:
            # メインフレームに切り替え
            driver.switch_to.default_content()
            driver.switch_to.frame("Itiran")
            operation_logger.info("受信一覧フレームに切り替えました")
            time.sleep(2)
        except Exception as e:
            operation_logger.error(f"フレームの切り替えに失敗しました: {e}")
            print(f"⚠️ フレームの切り替えに失敗しました")
            return False

        # 画面下部の受信一覧から最初の未読メッセージを取得
        operation_logger.info("未読メッセージ一覧を確認しています...")
        print("未読メッセージ一覧を確認しています...")

        # メッセージ一覧のテーブル行を取得（ヘッダー行を除く）
        try:
            # デバッグ: テーブルの存在確認
            tables = driver.find_elements(By.TAG_NAME, "table")
            operation_logger.info(f"ページ内のテーブル数: {len(tables)}件")

            # テーブルのIDを確認
            for i, table in enumerate(tables):
                table_id = table.get_attribute("id")
                operation_logger.info(f"テーブル{i}: id={table_id}")

            # テーブル内のメッセージ行を取得
            message_rows = driver.find_elements(By.XPATH, "//table[@id='grdJushin']//tr[position()>2]")
            operation_logger.info(f"未読メッセージ: {len(message_rows)}件")
            print(f"未読メッセージ: {len(message_rows)}件")

            if len(message_rows) == 0:
                operation_logger.info("未読メッセージはありません")
                print("未読メッセージはありません")
                return True

            # 処理対象メッセージの件数をカウント
            target_count = 0
            for row in message_rows:
                try:
                    title_link = row.find_element(By.XPATH, "./td[3]/a")
                    title = title_link.text.strip()

                    # タイトルが処理対象かチェック
                    should_process = False

                    # 「購入伺い」の完全一致チェック
                    if title == "購入伺い" and msg_processing.get("購入伺い", True):
                        should_process = True
                    # 「マッチング：使用期限」の前方一致チェック
                    elif title.startswith("マッチング：使用期限") and msg_processing.get("マッチング：使用期限", True):
                        should_process = True
                    # その他のタイトル（将来の拡張用）
                    elif title.startswith("不動在庫転送") and msg_processing.get("不動在庫転送", True):
                        should_process = True
                    elif title.startswith("Re:") and msg_processing.get("返信", True):
                        should_process = True

                    if not should_process:
                        operation_logger.info(f"スキップ: {title} (設定で無効)")
                        continue

                    target_count += 1
                    operation_logger.info(f"処理対象メッセージを発見: {title}")
                    if target_count >= max_message_count:  # 設定された最大件数
                        break
                except:
                    continue

            if target_count == 0:
                operation_logger.info("処理対象メッセージが見つかりませんでした")
                print("処理対象メッセージが見つかりませんでした")
                return True

            operation_logger.info(f"処理対象メッセージ: {target_count}件")
            print(f"\n処理対象メッセージ: {target_count}件")

            # 各メッセージを順番に処理（F5リロードで画面更新）
            for idx in range(1, target_count + 1):
                operation_logger.info(f"========== メッセージ {idx}/{target_count} の処理開始 ==========")
                print(f"\n========== メッセージ {idx}/{target_count} の処理開始 ==========")

                # 2件目以降はF5でリロード（DOM要素をリフレッシュ）
                if idx > 1:
                    try:
                        operation_logger.debug(f"メッセージ {idx} 処理前: F5リロード実行")

                        # メインコンテンツに戻る
                        driver.switch_to.default_content()

                        # driver.refresh()を使用（F5と同等）
                        driver.refresh()
                        time.sleep(3)  # リロード完了待機

                        # 再度フレームに切り替え
                        driver.switch_to.frame("Itiran")
                        time.sleep(1)

                        operation_logger.info("✓ ページをリロードしました")
                        print("✓ ページをリロードしました")
                    except Exception as e:
                        operation_logger.error(f"リロードエラー: {e}")
                        print(f"⚠️ リロードに失敗しました")
                        continue

                # メッセージリストを再取得
                try:
                    message_rows = driver.find_elements(By.XPATH, "//table[@id='grdJushin']//tr[position()>2]")

                    # 最初の処理対象メッセージを探す
                    row = None
                    for msg_row in message_rows:
                        try:
                            title_link = msg_row.find_element(By.XPATH, "./td[3]/a")
                            title = title_link.text.strip()

                            # タイトルが処理対象かチェック
                            should_process = False

                            # 「購入伺い」の完全一致チェック
                            if title == "購入伺い" and msg_processing.get("購入伺い", True):
                                should_process = True
                            # 「マッチング：使用期限」の前方一致チェック
                            elif title.startswith("マッチング：使用期限") and msg_processing.get("マッチング：使用期限", True):
                                should_process = True
                            # その他のタイトル
                            elif title.startswith("不動在庫転送") and msg_processing.get("不動在庫転送", True):
                                should_process = True
                            elif title.startswith("Re:") and msg_processing.get("返信", True):
                                should_process = True

                            if should_process:
                                row = msg_row
                                break
                        except:
                            continue

                    if not row:
                        operation_logger.warning(f"メッセージ {idx} が見つかりません（スキップ）")
                        print(f"⚠️ メッセージ {idx} が見つかりません（スキップ）")
                        continue

                    # 受信日時を取得
                    date_cell = row.find_element(By.XPATH, "./td[2]")
                    received_datetime = date_cell.text.strip()

                    # タイトルリンクを取得
                    title_link = row.find_element(By.XPATH, "./td[3]/a")
                    title = title_link.text.strip()

                    # 送信者を取得
                    sender_cell = row.find_element(By.XPATH, "./td[4]")
                    sender = sender_cell.text.strip()
                except Exception as e:
                    operation_logger.error(f"メッセージ {idx} の情報取得エラー: {e}")
                    print(f"⚠️ メッセージ {idx} の情報取得に失敗しました")
                    continue

                # メッセージIDを取得（リンクのonclick属性から抽出）
                onclick_attr = title_link.get_attribute("onclick")
                # 例: __doPostBack('grdJushin$_ctl3$_ctl0','')
                message_id = None
                if onclick_attr:
                    # URLパラメータにIDが含まれる可能性があるため、後でウィンドウから取得
                    pass

                operation_logger.info(f"処理対象メッセージ: {title}")
                operation_logger.info(f"  受信日時: {received_datetime}")
                operation_logger.info(f"  送信者: {sender}")
                print(f"\n処理対象メッセージ:")
                print(f"  タイトル: {title}")
                print(f"  受信日時: {received_datetime}")
                print(f"  送信者: {sender}")

                # 現在のウィンドウハンドルを保存
                main_window = driver.current_window_handle

                # タイトルリンクをクリック（新しいウィンドウで開く）
                safe_click(driver, title_link, "メッセージ詳細", 2, operation_logger)

                # 新しいウィンドウに切り替え
                all_windows = driver.window_handles
                for window in all_windows:
                    if window != main_window:
                        driver.switch_to.window(window)
                        break

                # URLからメッセージIDを取得
                current_url = driver.current_url
                operation_logger.info(f"メッセージURL: {current_url}")

                # URLパラメータからtarget値を抽出
                target_match = re.search(r'target=(\d+)', current_url)
                if target_match:
                    message_id = target_match.group(1)
                    operation_logger.info(f"メッセージID: {message_id}")

                # タイトルに応じて処理を分岐
                if title == "購入伺い":
                    # 購入伺いメッセージ：医薬品情報を抽出してストック保存
                    operation_logger.info("「購入伺い」メッセージを処理します")
                    print("\n「購入伺い」メッセージを処理します")

                    # メッセージ本文を取得
                    safe_wait(1, "メッセージ本文読み込み待機", operation_logger)
                    body_element = driver.find_element(By.TAG_NAME, "body")
                    message_content = body_element.text

                    operation_logger.info("メッセージ本文:")
                    operation_logger.info("-" * 50)
                    operation_logger.info(message_content)
                    operation_logger.info("-" * 50)

                    print("\nメッセージ本文:")
                    print("-" * 50)
                    print(message_content)
                    print("-" * 50)

                    # メッセージ本文をパース（複数ロット対応）
                    parsed_data_list = parse_message_content(message_content)

                    if parsed_data_list:
                        operation_logger.info(f"抽出されたデータ: {len(parsed_data_list)}件のロット")
                        print(f"\n抽出されたデータ: {len(parsed_data_list)}件のロット")

                        # 各ロットをストックに保存
                        saved_count = 0
                        for lot_idx, parsed_data in enumerate(parsed_data_list, 1):
                            operation_logger.info(f"  [{lot_idx}] 送信店舗: {parsed_data['sender_store']}")
                            operation_logger.info(f"      医薬品名: {parsed_data['medicine_name']}")
                            operation_logger.info(f"      数量: {parsed_data['quantity']}")
                            operation_logger.info(f"      単位: {parsed_data['unit']}")
                            operation_logger.info(f"      使用期限: {parsed_data['expiry_date']}")

                            print(f"  [{lot_idx}] 送信店舗: {parsed_data['sender_store']}")
                            print(f"      医薬品名: {parsed_data['medicine_name']}")
                            print(f"      数量: {parsed_data['quantity']}")
                            print(f"      単位: {parsed_data['unit']}")
                            print(f"      使用期限: {parsed_data['expiry_date']}")

                            # メッセージストックに保存（重複チェック）
                            if message_id:
                                # 重複チェック（メッセージID + 医薬品名 + 使用期限で判定）
                                lot_key = f"{message_id}_{parsed_data['medicine_name']}_{parsed_data['expiry_date']}"
                                existing = [m for m in message_stock['messages']
                                           if m.get('message_id') == message_id
                                           and m.get('medicine_name') == parsed_data['medicine_name']
                                           and m.get('expiry_date') == parsed_data['expiry_date']]

                                if not existing:
                                    new_message = {
                                        'message_id': message_id,
                                        'lot_key': lot_key,
                                        'received_datetime': received_datetime,
                                        'title': title,
                                        'sender': sender,
                                        'sender_store': parsed_data['sender_store'],
                                        'medicine_name': parsed_data['medicine_name'],
                                        'quantity': parsed_data['quantity'],
                                        'unit': parsed_data['unit'],
                                        'expiry_date': parsed_data['expiry_date'],
                                        'status': 'unprocessed',
                                        'created_at': datetime.now().isoformat()
                                    }

                                    message_stock['messages'].append(new_message)
                                    saved_count += 1
                                else:
                                    operation_logger.info(f"      → このロットは既にストックに存在します")
                                    print(f"      → このロットは既にストックに存在します")

                        if saved_count > 0:
                            save_message_stock(message_stock, store_id)
                            operation_logger.info(f"✓ メッセージをストックに保存しました（{saved_count}件のロット）")
                            print(f"\n✓ メッセージをストックに保存しました（{saved_count}件のロット）")
                        else:
                            operation_logger.info("全てのロットが既にストックに存在します")
                            print("\n全てのロットが既にストックに存在します")
                    else:
                        operation_logger.warning("メッセージのパースに失敗しました")
                        print("\n⚠️ メッセージのパースに失敗しました")

                elif title.startswith("マッチング：使用期限"):
                    # マッチングメッセージ：出庫処理を自動実行（最大3回リトライ）
                    operation_logger.info("「マッチング：使用期限」メッセージを処理します")
                    print("\n「マッチング：使用期限」メッセージを処理します")

                    max_retries = 3
                    success = False

                    for retry_count in range(max_retries):
                        try:
                            if retry_count > 0:
                                operation_logger.info(f"リトライ {retry_count}/{max_retries - 1} を実行します...")
                                print(f"\nリトライ {retry_count}/{max_retries - 1} を実行します...")
                                safe_wait(3, "リトライ前待機", operation_logger)
                            else:
                                # 初回のみ：出庫処理リンクをクリック（このクリックでメッセージウィンドウが閉じる）
                                safe_wait(1, "出庫処理ボタン検索前", operation_logger)

                                operation_logger.info(f"出庫処理ボタン（lnkSyukko）を探します...")
                                if not safe_find_and_click(driver, By.ID, "lnkSyukko", "出庫処理ボタン", 2, operation_logger):
                                    raise Exception("出庫処理ボタンのクリックに失敗")

                                # 出庫処理ボタンをクリックするとメッセージウィンドウが閉じて
                                # メインウィンドウに出庫画面が表示される
                                operation_logger.info("メインウィンドウに切り替えます...")
                                safe_wait(2, "ウィンドウ切り替え待機", operation_logger)

                                # メインウィンドウに切り替え
                                try:
                                    driver.switch_to.window(main_window)
                                    operation_logger.info("メインウィンドウに切り替わりました")
                                except Exception as e:
                                    operation_logger.warning(f"メインウィンドウ切り替えエラー: {e}")
                                    # フォールバック：最初のウィンドウに切り替え
                                    windows = driver.window_handles
                                    if windows:
                                        driver.switch_to.window(windows[0])
                                        operation_logger.info("最初のウィンドウに切り替えました")

                                # 出庫画面への遷移を確認（ページ読み込み完了＋ボタン存在確認）
                                if not wait_for_shipping_screen(driver, operation_logger, timeout=15):
                                    raise Exception("出庫画面への遷移がタイムアウトしました")

                            # 再計算ボタンをクリック
                            operation_logger.info(f"再計算ボタン（btnRecalc）を探します...")
                            if not safe_find_and_click(driver, By.ID, "btnRecalc", "再計算ボタン", 3, operation_logger):
                                # ボタンが見つからない場合、初回のみページ構造を調査
                                if retry_count == 0:
                                    operation_logger.warning("再計算ボタンが見つかりません。ページ構造を調査します...")
                                    try:
                                        from debug_page_structure import debug_page_structure
                                        debug_page_structure(driver, operation_logger)
                                    except Exception as e:
                                        operation_logger.warning(f"ページ構造調査エラー: {e}")
                                raise Exception("再計算ボタンのクリックに失敗")

                            # 出庫するボタンをクリック
                            operation_logger.info(f"出庫するボタン（btnSyuko）を探します...")
                            if not safe_find_and_click(driver, By.ID, "btnSyuko", "出庫するボタン", 2, operation_logger):
                                raise Exception("出庫するボタンのクリックに失敗")

                            # アラート（確認ダイアログ）の処理
                            try:
                                alert = driver.switch_to.alert
                                alert_text = alert.text
                                operation_logger.info(f"確認ダイアログ: {alert_text}")
                                print(f"確認ダイアログ: {alert_text}")
                                alert.accept()  # OKをクリック
                                time.sleep(2)
                                operation_logger.info("✓ 出庫処理が完了しました")
                                print("\n✓ 出庫処理が完了しました")
                                success = True
                                break
                            except:
                                # アラートがない場合も成功とみなす
                                operation_logger.info("✓ 出庫処理が完了しました（確認ダイアログなし）")
                                print("\n✓ 出庫処理が完了しました")
                                success = True
                                break

                        except Exception as e:
                            operation_logger.error(f"出庫処理エラー (試行 {retry_count + 1}/{max_retries}): {e}")
                            print(f"\n⚠️ 出庫処理に失敗しました (試行 {retry_count + 1}/{max_retries}): {e}")

                            if retry_count < max_retries - 1:
                                operation_logger.info("待機時間を3秒に延長してリトライします...")
                                print("待機時間を3秒に延長してリトライします...")
                            else:
                                operation_logger.error("最大リトライ回数に達しました。このメッセージの処理をスキップします。")
                                print("\n⚠️ 最大リトライ回数に達しました。このメッセージの処理をスキップします。")

                    if not success:
                        operation_logger.warning("出庫処理が完了できませんでした")
                        print("\n⚠️ 出庫処理が完了できませんでした")

                # ウィンドウ処理：マッチング処理ではメッセージウィンドウは既に閉じられている
                # それ以外の場合はメッセージウィンドウを閉じる
                if not title.startswith("マッチング：使用期限"):
                    try:
                        # メッセージウィンドウを閉じる
                        driver.close()
                        driver.switch_to.window(main_window)
                        operation_logger.info("メッセージウィンドウを閉じました")
                    except Exception as e:
                        operation_logger.warning(f"ウィンドウクローズエラー: {e}")
                        # エラーが発生した場合はメインウィンドウに切り替え
                        try:
                            driver.switch_to.window(main_window)
                        except:
                            windows = driver.window_handles
                            if windows:
                                driver.switch_to.window(windows[0])
                else:
                    # マッチング処理後は既にメインウィンドウにいるはず
                    operation_logger.info("マッチング処理完了（既にメインウィンドウ）")

            operation_logger.info(f"連絡板メッセージ確認処理が完了しました（{target_count}件処理）")
            operation_logger.info(f"ログファイル: {log_file_path}")
            print(f"\n✓ 連絡板メッセージ確認処理が完了しました（{target_count}件処理）")

            return True

        except Exception as e:
            operation_logger.error(f"メッセージ取得エラー: {e}")
            print(f"メッセージ取得エラー: {e}")
            import traceback
            traceback.print_exc()
            return False

    except Exception as e:
        operation_logger.error(f"連絡板メッセージ確認エラー: {e}")
        print(f"連絡板メッセージ確認エラー: {e}")
        import traceback
        traceback.print_exc()
        return False
    finally:
        # メインウィンドウに戻る
        try:
            windows = driver.window_handles
            if windows:
                driver.switch_to.window(windows[0])
        except:
            pass
