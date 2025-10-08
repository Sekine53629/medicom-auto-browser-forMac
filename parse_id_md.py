"""
Sample/id.md からvalueの後ろの数字と店舗名を抽出してCSV出力
"""
import re
import sys

# UTF-8で標準出力を設定
sys.stdout.reconfigure(encoding='utf-8')

def parse_id_md(file_path='Sample/id.md'):
    """id.mdファイルを解析"""

    # さまざまなエンコーディングで試す
    encodings = ['utf-16-le', 'utf-16', 'utf-8', 'shift-jis', 'cp932']

    content = None
    used_encoding = None

    for enc in encodings:
        try:
            with open(file_path, 'r', encoding=enc, errors='ignore') as f:
                test_content = f.read()
                # valueが含まれていればOK
                if 'value' in test_content.lower():
                    content = test_content
                    used_encoding = enc
                    break
        except:
            continue

    if not content:
        print("エラー: ファイルを読み込めませんでした", file=sys.stderr)
        return []

    # パターン1: <option value='数字'>店舗名</option>
    pattern1 = r"value='(\d+)'>([^<\n]+)"

    # パターン2: <option value="数字">店舗名</option>
    pattern2 = r'value="(\d+)">([^<\n]+)'

    # パターン3: value='数字' で始まり、>の後に店舗名
    pattern3 = r"value='([^']+)'[^>]*>([^\n<]+)"

    matches = []

    # 各パターンで試す
    for pattern in [pattern1, pattern2, pattern3]:
        found = re.findall(pattern, content)
        if found:
            matches.extend(found)

    # 重複を削除
    unique_matches = []
    seen = set()
    for store_id, store_name in matches:
        key = (store_id, store_name.strip())
        if key not in seen:
            seen.add(key)
            unique_matches.append(key)

    return unique_matches

def main():
    """メイン処理"""
    results = parse_id_md()

    if results:
        # CSV形式で出力
        print("store_id,store_name")
        for store_id, store_name in results:
            print(f"{store_id},{store_name}")
    else:
        print("データが見つかりませんでした", file=sys.stderr)
        print("", file=sys.stderr)
        print("Sample/id.md に以下のような形式でデータを追加してください:", file=sys.stderr)
        print("<option value='1830'>ツ)旭川末広5条店</option>", file=sys.stderr)

if __name__ == "__main__":
    main()
