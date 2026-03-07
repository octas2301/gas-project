# -*- coding: utf-8 -*-
"""
楽天CSVの「バリエーション1選択肢定義」を検証する。
楽天の仕様: 「各選択肢」（|で区切った1つ1つ）が32バイトまで。列全体ではない。
コード.js の truncateToByteLength と同じ数え方: charCode < 128 → 1バイト、それ以外 → 2バイト
"""
import csv
import os
import sys

COL_NAME = 'バリエーション1選択肢定義'
MAX_BYTES = 32

def byte_length(s):
    if s is None or s == '':
        return 0
    s = str(s).strip()
    n = 0
    for c in s:
        n += 1 if ord(c) < 128 else 2
    return n

def check_row(val):
    """|で区切った各選択肢が32バイト以下か。超過した選択肢があれば (選択肢の表示用, バイト数) のリストで返す"""
    if not val or not str(val).strip():
        return []
    over = []
    for part in str(val).split('|'):
        seg = part.strip()
        if not seg:
            continue
        blen = byte_length(seg)
        if blen > MAX_BYTES:
            over.append((seg[:30] + ('…' if len(seg) > 30 else ''), blen))
    return over

def main():
    csv_dir = os.path.join(os.path.dirname(__file__), '楽天CSVファイル')
    if not os.path.isdir(csv_dir):
        print('楽天CSVファイル フォルダが見つかりません')
        sys.exit(1)
    files = [f for f in os.listdir(csv_dir) if f.endswith('.csv')]
    files.sort()
    total_rows = 0
    total_over = 0
    results = []
    for fname in files:
        path = os.path.join(csv_dir, fname)
        try:
            with open(path, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                header = next(reader, None)
                if not header:
                    results.append((fname, '空', 0, []))
                    continue
                try:
                    idx = header.index(COL_NAME)
                except ValueError:
                    results.append((fname, 'ヘッダーに列なし', 0, []))
                    continue
                over = []
                count = 0
                for row in reader:
                    if len(row) <= idx:
                        continue
                    count += 1
                    val = (row[idx] or '').strip()
                    if not val:
                        continue
                    seg_over = check_row(val)
                    if seg_over:
                        item = (row[0] if row else '') or ('行' + str(count + 1))
                        for disp, blen in seg_over:
                            over.append((count + 1, item, disp, blen))
                total_rows += count
                total_over += len(over)
                results.append((fname, count, over, None))
        except UnicodeDecodeError:
            try:
                with open(path, 'r', encoding='cp932') as f:
                    reader = csv.reader(f)
                    header = next(reader, None)
                    if not header:
                        results.append((fname, '空', 0, []))
                        continue
                    try:
                        idx = header.index(COL_NAME)
                    except ValueError:
                        results.append((fname, 'ヘッダーに列なし', 0, []))
                        continue
                    over = []
                    count = 0
                    for row in reader:
                        if len(row) <= idx:
                            continue
                        count += 1
                        val = (row[idx] or '').strip()
                        if not val:
                            continue
                        seg_over = check_row(val)
                        if seg_over:
                            item = (row[0] if row else '') or ('行' + str(count + 1))
                            for disp, blen in seg_over:
                                over.append((count + 1, item, disp, blen))
                    total_rows += count
                    total_over += len(over)
                    results.append((fname, count, over, None))
            except Exception as e:
                results.append((fname, str(e), 0, []))
        except Exception as e:
            results.append((fname, str(e), 0, []))
    # 出力
    print('=== 楽天「バリエーション1選択肢定義」32バイト超過チェック（各選択肢ごと）===')
    print('楽天仕様: |で区切った「各選択肢」が32バイトまで。数え方: ASCII=1バイト、それ以外=2バイト\n')
    for r in results:
        if len(r) == 4 and r[3] is None:
            fname, data_rows, over, _ = r
            print('ファイル:', fname)
            print('  データ行数:', data_rows, '  32バイト超過:', len(over))
            if over:
                for row_num, item, val, blen in over[:15]:
                    print('    行{}  {}  バイト数={}  値="{}"'.format(row_num, item, blen, val))
                if len(over) > 15:
                    print('    ... 他', len(over) - 15, '件')
            print()
        else:
            print('ファイル:', r[0], '  ', r[1])
            print()
    print('--- まとめ ---')
    print('対象データ行合計:', total_rows)
    print('32バイト超過した選択肢の数:', total_over)
    if total_over == 0 and total_rows > 0:
        print('→ すべての選択肢が32バイト以下でした。')
    elif total_over > 0:
        print('→ 上記の選択肢が楽天の32バイト制限を超えています。')

if __name__ == '__main__':
    main()
