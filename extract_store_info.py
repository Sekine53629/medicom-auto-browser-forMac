"""
Sample/id.md から店舗情報を抽出してCSV化するスクリプト

使い方:
    python extract_store_info.py
"""

import re
import csv
import json
import os
from datetime import datetime
from bs4 import BeautifulSoup


def load_accounts(accounts_path='accounts.json'):
    """accounts.jsonを読み込み"""
    if not os.path.exists(accounts_path):
        return []

    with open(accounts_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def extract_store_id_from_user_id(user_id):
    """user_idから店舗ID（4桁）を抽出

    Args:
        user_id: ユーザーID（例: TRH170501）

    Returns:
        str: 店舗ID（例: 1705）
    """
    match = re.search(r'TRH(\d{4})\d{2}', user_id)
    if match:
        return match.group(1)
    return None


def parse_id_md(file_path='Sample/id.md'):
    """id.mdファイルから店舗名を抽出

    HTML要素、または単純なテキスト行から店舗名を抽出

    Args:
        file_path: id.mdファイルのパス

    Returns:
        list: 店舗名のリスト
    """
    if not os.path.exists(file_path):
        print(f"❌ ファイルが見つかりません: {file_path}")
        return []

    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    store_names = []

    # 方法1: HTML要素から抽出
    try:
        soup = BeautifulSoup(content, 'html.parser')
        elements = soup.find_all('span', {'id': 'lblShopName'})

        for elem in elements:
            store_name = elem.get_text(strip=True)
            if store_name:
                store_names.append(store_name)
    except:
        pass

    # 方法2: 正規表現でHTML要素から抽出
    html_pattern = r'<span[^>]*id="lblShopName"[^>]*>([^<]+)</span>'
    matches = re.findall(html_pattern, content)
    store_names.extend(matches)

    # 方法3: テキスト行から抽出（「ツ)」で始まる行）
    lines = content.split('\n')
    for line in lines:
        line = line.strip()
        # コードブロックやコメント行をスキップ
        if line.startswith('#') or line.startswith('```') or line.startswith('<!--'):
            continue

        # 「ツ)」で始まる店舗名を抽出
        if line.startswith('ツ)'):
            store_names.append(line)

    # 重複を削除
    return list(set(store_names))


def match_store_with_account(store_name, accounts):
    """店舗名からaccounts.jsonの該当アカウントを特定

    Args:
        store_name: 店舗名（例: "ツ)旭川末広5条店"）
        accounts: アカウントリスト

    Returns:
        dict or None: マッチしたアカウント情報
    """
    # 店舗名の正規化（「ツ)」「調剤薬局ツルハドラッグ」などを除去）
    normalized_name = store_name.replace('ツ)', '').replace('調剤薬局ツルハドラッグ', '').strip()

    for account in accounts:
        acc_store_name = account.get('store_name', '')

        # 部分一致で検索
        if normalized_name in acc_store_name or acc_store_name in normalized_name:
            return account

    return None


def save_store_mapping_csv(store_data, output_path='data/store_mapping.csv'):
    """店舗情報をCSV保存

    Args:
        store_data: [{"store_id": "1830", "store_name": "...", "user_id": "..."}]
        output_path: 出力CSVパス
    """
    # ディレクトリ作成
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    # 既存データを読み込み（重複チェック用）
    existing_data = {}
    if os.path.exists(output_path):
        with open(output_path, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                existing_data[row['store_id']] = row

    # 新規データを追加
    for data in store_data:
        existing_data[data['store_id']] = {
            'store_id': data['store_id'],
            'store_name': data['store_name'],
            'user_id': data.get('user_id', ''),
            'last_updated': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }

    # CSV書き込み
    with open(output_path, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['store_id', 'store_name', 'user_id', 'last_updated'])

        for data in existing_data.values():
            writer.writerow([
                data['store_id'],
                data['store_name'],
                data['user_id'],
                data['last_updated']
            ])

    print(f"✓ CSV保存完了: {output_path}")
    print(f"  保存件数: {len(existing_data)} 件")


def main():
    """メイン処理"""
    print("=" * 60)
    print("店舗情報抽出ツール")
    print("=" * 60)

    # id.mdから店舗名を抽出
    print("\n1. Sample/id.md から店舗名を抽出中...")
    store_names = parse_id_md()

    if not store_names:
        print("❌ 店舗名が見つかりませんでした")
        print("\nSample/id.md に以下の形式でデータを追加してください:")
        print('  <span id="lblShopName" class="cssTenpoName">ツ)旭川末広5条店</span>')
        print("  または")
        print('  ツ)旭川末広5条店')
        return

    print(f"✓ {len(store_names)} 件の店舗名を抽出しました")
    for name in store_names:
        print(f"  - {name}")

    # accounts.jsonを読み込み
    print("\n2. accounts.json からアカウント情報を読み込み中...")
    accounts = load_accounts()

    if not accounts:
        print("⚠️ accounts.json が見つかりません")
        print("店舗IDをCSVに含めることができません")
    else:
        print(f"✓ {len(accounts)} 件のアカウントを読み込みました")

    # 店舗名とアカウントをマッチング
    print("\n3. 店舗名とアカウント情報をマッチング中...")
    store_data = []

    for store_name in store_names:
        account = match_store_with_account(store_name, accounts)

        if account:
            store_id = extract_store_id_from_user_id(account['user_id'])
            user_id = account['user_id']

            print(f"✓ {store_name} → ID: {store_id}, User: {user_id}")

            store_data.append({
                'store_id': store_id,
                'store_name': store_name,
                'user_id': user_id
            })
        else:
            print(f"⚠️ {store_name} → マッチするアカウントが見つかりません")

            # アカウントが見つからない場合でも、店舗名だけCSVに保存
            store_data.append({
                'store_id': 'UNKNOWN',
                'store_name': store_name,
                'user_id': ''
            })

    # CSV保存
    if store_data:
        print("\n4. CSVファイルに保存中...")
        save_store_mapping_csv(store_data)

        print("\n" + "=" * 60)
        print("✓ 処理完了")
        print("=" * 60)
    else:
        print("\n❌ 保存するデータがありません")


if __name__ == "__main__":
    main()
