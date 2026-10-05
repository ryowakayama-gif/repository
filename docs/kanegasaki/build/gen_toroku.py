# -*- coding: utf-8 -*-
"""所得段階別第1号被保険者数の登録用データ（xlsx・csv）を作る。

　見える化システムの推計ワークシート「5_保険料推計」シートへ登録する値である。
　13段階×3年度。記録（所得段階別被保険者数の登録後の確認結果）の値による。
"""
import csv
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

DATA = [
    ('第1段階',   580,   579,   579),
    ('第2段階',   435,   435,   434),
    ('第3段階',   373,   372,   372),
    ('第4段階',   665,   664,   664),
    ('第5段階',  1096,  1094,  1094),
    ('第6段階',   794,   793,   792),
    ('第7段階',   504,   504,   503),
    ('第8段階',   172,   172,   171),
    ('第9段階',    68,    68,    68),
    ('第10段階',   24,    24,    24),
    ('第11段階',   10,    10,    10),
    ('第12段階',    7,     7,     7),
    ('第13段階',   28,    28,    28),
]
YEARS = ['令和9年度', '令和10年度', '令和11年度']
tot = [sum(r[i + 1] for r in DATA) for i in range(3)]
assert tot == [4756, 4750, 4746], tot
assert sum(tot) == 14252

# ── csv ────────────────────────────────────────────
with open('out/所得段階別第1号被保険者数_登録用.csv', 'w', encoding='utf-8-sig', newline='') as f:
    w = csv.writer(f)
    w.writerow(['所得段階'] + YEARS)
    for r in DATA:
        w.writerow(list(r))
    w.writerow(['合計'] + tot)

# ── xlsx ───────────────────────────────────────────
wb = Workbook()
ws = wb.active
ws.title = '所得段階別第1号被保険者数'
HEAD = PatternFill('solid', fgColor='1F3864')
WH = Font(name='游ゴシック', size=10, bold=True, color='FFFFFF')
BD = Font(name='游ゴシック', size=10, bold=True)
NM = Font(name='游ゴシック', size=10)
thin = Side(style='thin', color='BFBFBF')
BR = Border(left=thin, right=thin, top=thin, bottom=thin)

ws.append(['所得段階'] + YEARS)
for r in DATA:
    ws.append(list(r))
ws.append(['合計'] + tot)
ws.append([])
ws.append(['3か年計', sum(tot)])

for row in ws.iter_rows(min_row=1, max_row=15, max_col=4):
    for c in row:
        c.border = BR
        c.alignment = Alignment(horizontal='center' if c.column > 1 else 'left', vertical='center')
        if c.row == 1:
            c.fill = HEAD; c.font = WH
        elif c.row == 15:
            c.font = BD
        else:
            c.font = NM
        if c.column > 1 and c.row > 1:
            c.number_format = '#,##0'
ws['A17'].font = BD
ws['B17'].font = BD
ws['B17'].number_format = '#,##0'
ws.column_dimensions['A'].width = 14
for col in 'BCD':
    ws.column_dimensions[col].width = 13

ws['A19'] = '登録先：見える化システム 推計ワークシート「5_保険料推計」シート'
ws['A20'] = '保険料段階設定数は0とし、①標準段階区分・割合（第2段階0.685）で登録する'
ws['A21'] = '登録後の所得段階別加入割合補正係数は 0.9770／0.9767／0.9770、補正後被保険者数は 13,920.560人'
ws['A22'] = '保険料の分母は 13,920.560 × 12 × 99.77％ ＝ 166,663人・月'
for r in range(19, 23):
    ws.cell(row=r, column=1).font = Font(name='游ゴシック', size=9)

wb.save('out/所得段階別第1号被保険者数_登録用.xlsx')
print('登録用データを作成　合計 %s／%s／%s　3か年計 %s' % (*tot, sum(tot)))
