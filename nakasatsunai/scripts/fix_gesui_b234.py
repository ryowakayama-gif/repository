# -*- coding: utf-8 -*-
"""公表中の経営戦略の成果品（検討報告書・将来人口予測ファイル）で根拠が確定した
B-2（将来人口の推計根拠）・B-3（普及率等の設定年度）・B-4（普及率の値）を本文へ反映する。"""
import zipfile, sys
from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
XS = '{http://www.w3.org/XML/1998/namespace}space'
SRC, DST = sys.argv[1], sys.argv[2]


def replace_in_para(p, old, new):
    ts = list(p.iter(W + 't'))
    texts = [t.text or '' for t in ts]
    full = ''.join(texts)
    idx = full.find(old)
    if idx < 0:
        return False
    end = idx + len(old)
    spans, pos = [], 0
    for t, tx in zip(ts, texts):
        spans.append((pos, pos + len(tx), t, tx)); pos += len(tx)
    first = True
    for s, e, t, tx in spans:
        if e <= idx or s >= end:
            continue
        a, b = max(idx, s) - s, min(end, e) - s
        t.text = tx[:a] + (new if first else '') + tx[b:]
        first = False
        if t.text != t.text.strip():
            t.set(XS, 'preserve')
    return True


EDITS = [
    (837,
     '村内での下水道普及率は６９．５％（令和５年度決算）であり',
     '村内での下水道普及率は７０．９％であり',
     'B-4　検討報告書「将来の普及率=70.9%」および将来人口予測ファイルの採用値に合わせる'),
    (938,
     '有収水量の予測に際し、総人口は国立社会保障・人口問題研究所による推計から算出しています。処理区域内人口・水洗化人口は、普及率・水洗化率の令和４（２０２２）年度から令和６（２０２４）年度の３か年実績の平均を基に算出しています。',
     '有収水量の予測に際し、総人口は「第７期中札内村まちづくり計画」（令和４年３月策定）における村独自推計を採用しています。処理区域内人口・水洗化人口は、普及率・水洗化率が過去５年間でほぼ横ばいで推移していることから、令和５（２０２３）年度の実績値を基準として算出しています。',
     'B-2・B-3　検討報告書「行政人口の将来推計値は…第7期中札内村まちづくり計画における将来推計値を採用する」ほか'),
    (939,
     'また、令和４（２０２２）年度から令和６（２０２４）年度の３か年実績から一人当たり処理水量及び有収率を設定し、年間有収水量の推計を行っています。',
     'また、令和５（２０２３）年度実績の一人当たり有収水量原単位（０．２４９㎥／人／日）を基に、年間有収水量の推計を行っています。',
     'B-3　将来人口予測ファイルの原単位は令和5年度実績の据置'),
    (944,
     '令和3（2021）年度から令和5（2023）年度までの過去3か年実績の平均で算出',
     '令和5（2023）年度実績の値を採用（過去5年間でほぼ横ばいのため）',
     'B-3　前提条件ボックスを実際の設定方法に合わせる'),
]

z = zipfile.ZipFile(SRC)
root = etree.fromstring(z.read('word/document.xml'))
ps = list(root.iter(W + 'p'))
applied, failed = [], []
for i, old, new, why in EDITS:
    (applied if replace_in_para(ps[i], old, new) else failed).append((i, old, new, why))

parts = {n: z.read(n) for n in z.namelist()}
parts['word/document.xml'] = etree.tostring(root, xml_declaration=True,
                                            encoding='UTF-8', standalone=True)
zo = zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED)
for it in z.infolist():
    zo.writestr(it, parts[it.filename])
zo.close(); z.close()

print(f'=== 適用 {len(applied)} 件 / 失敗 {len(failed)} 件 ===')
for i, old, new, why in applied:
    print(f'  [段落{i}] {old[:46]}')
    print(f'        → {new[:46]}')
    print(f'          {why}')
for f in failed:
    print('  !! 未適用:', f[0], f[1][:40])
sys.exit(1 if failed else 0)
