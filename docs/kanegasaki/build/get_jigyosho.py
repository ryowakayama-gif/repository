# -*- coding: utf-8 -*-
"""町内介護サービス提供事業所一覧を、介護サービス情報公表システムのオープンデータから組む。

　厚生労働省が同システムの事業所データをＣＳＶで公表している（毎年6月末・12月末時点）。
　全国一括・サービス種類別の35ファイルから、市区町村名が「金ケ崎」「金ヶ崎」の行を抜き出す。
　画面の検索は自動のアクセスを受け付けないが、同じ中身がこちらで取れる。

　  python3 -I get_jigyosho.py [出力先CSV] [時点のスタンプ]
　　　時点のスタンプの既定は 20260709（令和8年6月末時点）。
　　　別の時点にするときは一覧のページでファイル名のスタンプを確かめる。

　要 ネットワーク　www.mhlw.go.jp
"""
import csv
import io
import re
import sys
import urllib.request

PAGE = 'https://www.mhlw.go.jp/stf/kaigo-kouhyou_opendata.html'
BASE = 'https://www.mhlw.go.jp'
COLS = ('都道府県コード又は市町村コード', '市区町村名', '事業所名', 'サービスの種類',
        '住所', '法人の名称', '事業所番号', '定員')

out_path = sys.argv[1] if len(sys.argv) > 1 else 'out/町内介護サービス提供事業所一覧.csv'
stamp = sys.argv[2] if len(sys.argv) > 2 else '20260709'

with urllib.request.urlopen(PAGE, timeout=120) as r:
    html = r.read().decode('utf-8', errors='replace')

links, seen = [], set()
for u, code in re.findall(
        r'href="(/content/12300000/jigyosho_(\d+)_all_%s\d+\.csv)"' % stamp, html):
    if u not in seen:
        seen.add(u)
        links.append((u, code))
if not links:
    sys.exit('時点 %s のファイルが一覧にありません。ページでスタンプを確かめてください。' % stamp)
print('対象ファイル %d件（時点 %s）' % (len(links), stamp))

rows = []
for u, code in links:
    try:
        with urllib.request.urlopen(BASE + u, timeout=300) as r:
            txt = r.read().decode('utf-8-sig', errors='replace')
    except Exception as e:                      # 1ファイルの失敗で全体を止めない
        print('  取得できず %s　%s' % (code, e))
        continue
    rdr = csv.reader(io.StringIO(txt))
    hdr = next(rdr)
    idx = {c: i for i, c in enumerate(hdr)}
    n = 0
    for row in rdr:
        if len(row) < len(hdr):
            continue
        city = row[idx['市区町村名']]
        if '金ケ崎' in city or '金ヶ崎' in city:
            rows.append([code] + [row[idx[k]] for k in COLS])
            n += 1
    if n:
        print('  %s　%d件' % (code, n))

rows.sort(key=lambda r: (r[0], r[3]))
with io.open(out_path, 'w', encoding='utf-8-sig', newline='') as f:
    w = csv.writer(f)
    w.writerow(['サービスコード'] + list(COLS))
    w.writerows(rows)
print('%d事業所　%s' % (len(rows), out_path))
