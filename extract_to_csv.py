#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Sample/id.md から店舗ID・店舗名を抽出してCSV出力

想定される入力形式:
  <option value='1830'>ツ)旭川末広5条店</option>
  <option value='1705'>ツ)旭川7条店</option>

出力: data/store_mapping.csv
"""
import re
import os
import csv
from datetime import datetime


def read_file_with_encoding(file_path):
    """複数のエンコーディングを試してファイルを読み込み"""
    encodings = ['utf-8', 'utf-16', 'utf-16-le', 'shift-jis', 'cp932', 'euc-jp']

    for encoding in encodings:
        try:
            with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
                content = f.read()
                if content and len(content) > 0:
                    return content
        except:
            continue

    return None


def extract_store_data(content):
    """HTML contentから店舗ID・店舗名を抽出"""
    results = []

    # パターン1: <option value='数字'>店舗名
    pattern1 = r"<option\s+value='(\d+)'>([^<]+)</option>"
    matches1 = re.findall(pattern1, content, re.IGNORECASE)
    results.extend(matches1)

    # パターン2: <option value="数字">店舗名
    pattern2 = r'<option\s+value="(\d+)">([^<]+)</option>'
    matches2 = re.findall(pattern2, content, re.IGNORECASE)
    results.extend(matches2)

    # パターン3: value='数字'>店舗名（<option タグなし）
    pattern3 = r"value='(\d+)'>([^\n<]+)"
    matches3 = re.findall(pattern3, content)
    results.extend(matches3)

    # 重複削除
    unique_results = []
    seen = set()

    for store_id, store_name in results:
        store_name = store_name.strip()
        if (store_id, store_name) not in seen:
            seen.add((store_id, store_name))
            unique_results.append((store_id, store_name))

    return unique_results


def save_to_csv(data, output_path='data/store_mapping.csv'):
    """CSVに保存"""
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    with open(output_path, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['store_id', 'store_name', 'last_updated'])

        for store_id, store_name in data:
            writer.writerow([
                store_id,
                store_name,
                datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            ])

    print(f"\u2713 CSV\u4fdd\u5b58\u5b8c\u4e86: {output_path}")
    print(f"  \u4fdd\u5b58\u4ef6\u6570: {len(data)} \u4ef6")


def main():
    """メイン処理"""
    input_file = 'Sample/id.md'

    print("=" * 60)
    print("店舗情報抽出ツール")
    print("=" * 60)
    print(f"\n入力ファイル: {input_file}")

    if not os.path.exists(input_file):
        print(f"\nエラー: {input_file} が見つかりません")
        return

    # ファイル読み込み
    content = read_file_with_encoding(input_file)

    if not content:
        print("\nエラー: ファイルを読み込めませんでした")
        return

    # データ抽出
    results = extract_store_data(content)

    if not results:
        print("\nデータが見つかりませんでした")
        print("\nSample/id.md に以下のような形式でデータを追加してください:")
        print("  <option value='1830'>ツ)旭川末広5条店</option>")
        print("  <option value='1705'>ツ)旭川7条店</option>")
        return

    # 結果表示
    print(f"\n抽出されたデータ ({len(results)} 件):")
    print("-" * 60)
    print(f"{'店舗ID':<15} 店舗名")
    print("-" * 60)

    for store_id, store_name in results:
        print(f"{store_id:<15} {store_name}")

    # CSV保存
    print("\n" + "=" * 60)
    save_to_csv(results)
    print("=" * 60)


if __name__ == "__main__":
    main()
