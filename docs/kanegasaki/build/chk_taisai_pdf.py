# -*- coding: utf-8 -*-
"""前回計画書（PDF）の体裁を実測する。フォント・サイズ・色・余白・ページ番号。"""
import sys, collections, pymupdf
d = pymupdf.open(sys.argv[1])
cnt = collections.Counter()
L = R = T = B = None
for i, p in enumerate(d):
    w, h = p.rect.width, p.rect.height
    for blk in p.get_text('dict')['blocks']:
        if blk['type'] != 0:
            continue
        for ln in blk['lines']:
            for sp in ln['spans']:
                t = sp['text'].strip()
                if not t:
                    continue
                cnt[(sp['font'], round(sp['size'], 1), '%06X' % sp['color'])] += len(t)
                x0, y0, x1, y1 = sp['bbox']
                L = x0 if L is None else min(L, x0)
                R = x1 if R is None else max(R, x1)
                T = y0 if T is None else min(T, y0)
                B = y1 if B is None else max(B, y1)
mm = lambda v: round(v * 25.4 / 72, 1)
print('ページ数 %d／用紙 %smm × %smm' % (d.page_count, mm(d[0].rect.width), mm(d[0].rect.height)))
print('文字の占める範囲（全ページの外周）　左%smm 右端%smm 上%smm 下端%smm'
      % (mm(L), mm(d[0].rect.width - R), mm(T), mm(d[0].rect.height - B)))
print()
print('フォント・サイズ・色の分布（文字数の多い順・上位14）')
tot = sum(cnt.values())
for (f, s, c), n in cnt.most_common(14):
    print('  %-28s %5s pt  #%s  %7d字  %5.1f%%' % (f, s, c, n, 100 * n / tot))
print('  （全 %d字・%d種）' % (tot, len(cnt)))
