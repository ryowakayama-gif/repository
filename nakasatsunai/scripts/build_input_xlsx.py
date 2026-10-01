# -*- coding: utf-8 -*-
"""打合せ確認事項の「確認結果入力シート」を生成する。

この環境の LibreOffice は Office 形式を読み込めず recalc.py が使えないため、
数式は Excel で開いた時点で計算されるよう fullCalcOnLoad を立てる。
入力欄は黄色塗り、回答列はプルダウン（入力規則）を設定する。
"""
import json, sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

SRC, DST = sys.argv[1], sys.argv[2]
D = json.load(open(SRC, encoding='utf-8'))
ITEMS = D['items']

FONT = 'Yu Gothic'
INK, TEAL, GREY = '1A2422', '0E5E62', '5E6B67'
RED, AMBER, GREEN = 'A33520', '8A6512', '2F6B45'
HEADBG, ZEBRA, INPUTBG, NOTEBG = 'DCE6E3', 'F5F8F7', 'FFF7D6', 'EFF4F2'
RULE = 'C8D2CE'

thin = Side(style='thin', color=RULE)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
IMPACT_COLOR = {'大': RED, '中': AMBER, '小': GREEN}
STATE_COLOR = {'仮反映済': GREEN, '照会中': AMBER, '方針待ち': AMBER, '未確認': GREY}

def style(c, *, size=10, bold=False, color=INK, bg=None, wrap=True,
          ha='left', va='top', border=True):
    c.font = Font(name=FONT, size=size, bold=bold, color=color)
    if bg:
        c.fill = PatternFill('solid', fgColor=bg)
    c.alignment = Alignment(horizontal=ha, vertical=va, wrap_text=wrap)
    if border:
        c.border = BOX

wb = Workbook()

# ══════════════════ シート1：確認結果入力 ══════════════════
ws = wb.active
ws.title = '確認結果入力'

COLS = [
    ('ID', 8), ('区分', 20), ('影響度', 8), ('状態', 10), ('確認事項', 26),
    ('確認したいこと', 48), ('選択肢', 32),
    ('ご回答', 28), ('補足・回答内容', 36), ('決定者', 12), ('決定日', 12), ('社内メモ', 24),
]
INPUT_COLS = ('H', 'I', 'J', 'K')          # 村にご記入いただく列
ANSWER_COL = 'H'

# タイトル行
ws['A1'] = D['meta']['subtitle'] + '　確認結果入力シート'
style(ws['A1'], size=14, bold=True, color=INK, bg=None, wrap=False, va='center', border=False)
ws.merge_cells('A1:G1')
ws['H1'] = '← 黄色のセルにご記入ください'
style(ws['H1'], size=10, bold=True, color=AMBER, wrap=False, va='center', border=False)
ws.merge_cells('H1:L1')
ws.row_dimensions[1].height = 24

ws['A2'] = f"作成日 {D['meta']['date']}　／　全{len(ITEMS)}件（影響度 大{sum(1 for i in ITEMS if i['impact']=='大')}件・中{sum(1 for i in ITEMS if i['impact']=='中')}件・小{sum(1 for i in ITEMS if i['impact']=='小')}件）　／　IDは別添Wordの確認事項IDと対応"
style(ws['A2'], size=9, color=GREY, wrap=False, va='center', border=False)
ws.merge_cells('A2:L2')
ws.row_dimensions[2].height = 18

HEAD = 4
for j, (name, w) in enumerate(COLS, start=1):
    col = get_column_letter(j)
    ws.column_dimensions[col].width = w
    c = ws.cell(row=HEAD, column=j, value=name)
    style(c, size=10, bold=True, color=INK, bg=HEADBG, ha='center', va='center')
ws.row_dimensions[HEAD].height = 30

GENERIC = ['提供済み', '提供予定', '提供不可', '未定']

# 選択肢は別シートに置いて範囲参照する（桁区切りのカンマが入るため、直書きリストは使えない）
mst = wb.create_sheet('選択肢マスタ')
mst['A1'] = '※　このシートは「確認結果入力」シートのプルダウンの元データです。編集しないでください。'
style(mst['A1'], size=9, color=GREY, wrap=False, border=False)
mst.column_dimensions['A'].width = 10
for col in 'BCDEF':
    mst.column_dimensions[col].width = 34
mst_head = mst.cell(row=2, column=1, value='ID')
style(mst_head, size=9, bold=True, bg=HEADBG, ha='center', va='center')
for j in range(2, 7):
    style(mst.cell(row=2, column=j, value=f'選択肢{j-1}'), size=9, bold=True, bg=HEADBG, ha='center', va='center')

for n, it in enumerate(ITEMS):
    r = HEAD + 1 + n
    mr = 3 + n                                        # 選択肢マスタ上の行
    opts = list(it['options']) if it['options'] else list(GENERIC)
    if '未定' not in opts:
        opts.append('未定')
    assert len(opts) <= 5, f'選択肢が多すぎます: {it["id"]}'
    style(mst.cell(row=mr, column=1, value=it['id']), size=9, bold=True, ha='center', va='center')
    for j, o in enumerate(opts):
        style(mst.cell(row=mr, column=2 + j, value=o), size=9)

    vals = [it['id'], f"{it['cat']}．{D['categories'][it['cat']]}", it['impact'], it.get('state', '未確認'),
            it['title'], it['ask'], '／'.join(it['options']) if it['options'] else '（資料のご提供）',
            None, None, None, None, it.get('note')]
    zebra = ZEBRA if n % 2 else None
    for j, v in enumerate(vals, start=1):
        col = get_column_letter(j)
        c = ws.cell(row=r, column=j, value=v)
        if col in INPUT_COLS:
            style(c, size=10, bg=INPUTBG, ha='left')
        elif col == 'A':
            style(c, size=10, bold=True, bg=zebra, ha='center', va='center')
        elif col == 'C':
            style(c, size=10, bold=True, color=IMPACT_COLOR[it['impact']], bg=zebra, ha='center', va='center')
        elif col == 'D':
            style(c, size=9, bold=True, color=STATE_COLOR.get(it.get('state'), GREY), bg=zebra, ha='center', va='center')
        else:
            style(c, size=9, bg=zebra)
    ws.cell(row=r, column=11).number_format = 'yyyy/mm/dd'

    longest = max(len(it['ask']), len(it['title']) * 2)
    ws.row_dimensions[r].height = max(46, min(132, 14 + (longest // 26) * 13))

    # ご回答列のプルダウン（行ごとに選択肢が異なるため行単位で設定）
    last_col = get_column_letter(1 + len(opts))
    dv = DataValidation(type='list',
                        formula1=f"'選択肢マスタ'!$B${mr}:${last_col}${mr}",
                        allow_blank=True, showDropDown=False)
    dv.promptTitle = it['id']
    dv.prompt = '一覧から選択してください'
    ws.add_data_validation(dv)
    dv.add(ws[f'{ANSWER_COL}{r}'])

LAST = HEAD + len(ITEMS)
ws.auto_filter.ref = f'A{HEAD}:L{LAST}'
ws.freeze_panes = f'E{HEAD+1}'
ws.sheet_view.zoomScale = 90
ws.print_title_rows = f'{HEAD}:{HEAD}'
ws.page_setup.orientation = 'landscape'
ws.page_setup.paperSize = ws.PAPERSIZE_A3
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_options.horizontalCentered = True

# ══════════════════ シート2：凡例・記入要領 ══════════════════
lg = wb.create_sheet('凡例・記入要領')
for col, w in zip('ABCDEF', (4, 16, 20, 46, 20, 20)):
    lg.column_dimensions[col].width = w

def put(r, c, v, **kw):
    cell = lg.cell(row=r, column=c, value=v)
    style(cell, border=kw.pop('border', False), **kw)
    return cell

put(1, 2, '記入要領', size=14, bold=True, wrap=False)
lg.row_dimensions[1].height = 24

put(3, 2, '１．ご記入いただく列', size=11, bold=True, color=TEAL, wrap=False)
rows = [
    ('H　ご回答', '一覧から選択（セルをクリックすると選択肢が出ます）。該当がなければ「その他」を選び、H列に内容をご記入ください。'),
    ('I　補足・回答内容', '自由記述。資料をご提供いただく事項（A-1・A-6）はファイル名や提供時期をご記入ください。'),
    ('J　決定者', 'ご決定された方・部署名。'),
    ('K　決定日', '2026/10/15 の形式。'),
]
r = 4
for a, b in rows:
    put(r, 2, a, size=10, bold=True, bg=INPUTBG, border=True)
    put(r, 3, b, size=9, border=True)
    lg.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
    lg.row_dimensions[r].height = 30
    r += 1
put(r, 2, 'L　社内メモ', size=10, bold=True, border=True)
put(r, 3, '当方の作業メモ欄です。ご記入は不要です。', size=9, color=GREY, border=True)
lg.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
r += 2

put(r, 2, '２．影響度の区分', size=11, bold=True, color=TEAL, wrap=False); r += 1
for k, v in D['impactDef']:
    put(r, 2, k, size=10, bold=True, color=IMPACT_COLOR[k], ha='center', va='center', border=True)
    put(r, 3, v, size=9, border=True)
    lg.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
    lg.row_dimensions[r].height = 28
    r += 1
r += 1

put(r, 2, '３．記入例', size=11, bold=True, color=TEAL, wrap=False); r += 1
ex_head = ['ID', 'ご回答', '補足・回答内容', '決定者', '決定日']
for j, h in enumerate(ex_head):
    put(r, 2 + j, h, size=9, bold=True, bg=HEADBG, ha='center', va='center', border=True)
lg.row_dimensions[r].height = 22
r += 1
ex = ['B-4', '69.48%（経営比較分析表・普及率）', '公表している経営比較分析表の値に揃える。指標名も「普及率」とする。', '施設課', '2026/10/15']
for j, v in enumerate(ex):
    c = put(r, 2 + j, v, size=9, bg=NOTEBG, border=True)
    if j == 4:
        c.number_format = 'yyyy/mm/dd'
lg.row_dimensions[r].height = 32
r += 1
put(r, 2, '※　上の行は記入例です。実際のご記入は「確認結果入力」シートにお願いします。', size=9, color=GREY, wrap=False)
r += 2

put(r, 2, '４．回答状況', size=11, bold=True, color=TEAL, wrap=False); r += 1
put(r, 2, '※　Excelで開くと自動計算されます。', size=9, color=GREY, wrap=False); r += 1
base = r
prog = [
    ('確認事項の総数', f'=COUNTA(確認結果入力!$A${HEAD+1}:$A${LAST})'),
    ('回答済み',       f'=COUNTA(確認結果入力!${ANSWER_COL}${HEAD+1}:${ANSWER_COL}${LAST})'),
    ('未回答',         f'=C{base}-C{base+1}'),
    ('うち影響度「大」の未回答',
     f'=COUNTIFS(確認結果入力!$C${HEAD+1}:$C${LAST},"大",確認結果入力!${ANSWER_COL}${HEAD+1}:${ANSWER_COL}${LAST},"")'),
]
for k, (label, f) in enumerate(prog):
    put(base + k, 2, label, size=10, bold=(k == 3), border=True)
    c = lg.cell(row=base + k, column=3, value=f)
    style(c, size=10, bold=True, color=RED if k == 3 else INK, ha='center', va='center', border=True)
    lg.row_dimensions[base + k].height = 22

lg.sheet_view.showGridLines = False

# ══════════════════ シート3：確認事項の詳細（参照用） ══════════════════
dt = wb.create_sheet('確認事項の詳細')
for col, w in zip('ABCDEFG', (8, 8, 10, 24, 58, 32, 32)):
    dt.column_dimensions[col].width = w
head = ['ID', '影響度', '状態', '確認事項', '背景・現状', '影響範囲', '未決の場合']
for j, h in enumerate(head, start=1):
    c = dt.cell(row=1, column=j, value=h)
    style(c, size=10, bold=True, bg=HEADBG, ha='center', va='center')
dt.row_dimensions[1].height = 26
for n, it in enumerate(ITEMS):
    r = 2 + n
    zebra = ZEBRA if n % 2 else None
    for j, v in enumerate([it['id'], it['impact'], it.get('state', '未確認'), it['title'],
                           it['background'], it['scope'], it['risk']], start=1):
        c = dt.cell(row=r, column=j, value=v)
        if j == 1:
            style(c, size=10, bold=True, bg=zebra, ha='center', va='center')
        elif j == 2:
            style(c, size=10, bold=True, color=IMPACT_COLOR[it['impact']], bg=zebra, ha='center', va='center')
        elif j == 3:
            style(c, size=9, bold=True, color=STATE_COLOR.get(it.get('state'), GREY), bg=zebra, ha='center', va='center')
        else:
            style(c, size=9, bg=zebra)
    dt.row_dimensions[r].height = max(50, min(150, 14 + (len(it['background']) // 60) * 13))
dt.freeze_panes = 'D2'
dt.auto_filter.ref = f'A1:G{1+len(ITEMS)}'
dt.page_setup.orientation = 'landscape'
dt.page_setup.paperSize = dt.PAPERSIZE_A3
dt.page_setup.fitToWidth = 1
dt.page_setup.fitToHeight = 0
dt.sheet_properties.pageSetUpPr.fitToPage = True

# Excelで開いた時点で数式を計算させる（この環境では再計算できないため）
wb.calculation.fullCalcOnLoad = True
# シート順：入力 → 凡例 → 詳細 → 選択肢マスタ
wb.move_sheet('選択肢マスタ', offset=len(wb.sheetnames) - 1 - wb.sheetnames.index('選択肢マスタ'))
wb.active = 0
wb.save(DST)
print(f'wrote {DST} / {len(ITEMS)} items / sheets: {wb.sheetnames}')
