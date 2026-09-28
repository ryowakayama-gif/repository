# -*- coding: utf-8 -*-
"""採用計算表（Excel）の作成。

中間報告書に載せた指標を、この1冊から再現できるようにするためのもの。
`tariff_model.py` の算定結果をそのまま書き出し、07_本文照合表で
本文の数値と突き合わせられるようにしている。
"""
import os, io, sys, contextlib, importlib.util
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, '..', 'onagawa_adopted_calc.xlsx')

spec = importlib.util.spec_from_file_location('tm', os.path.join(HERE, 'tariff_model.py'))
tm = importlib.util.module_from_spec(spec)
with contextlib.redirect_stdout(io.StringIO()):
    spec.loader.exec_module(tm)

M = tm.M_A
K_FULL, K_HALF = tm.K_FULL, tm.K_HALF
GENKA = tm.kyusui_genka
TANKA = tm.kyoukyu_tanka

NAVY, BLUE, RED, ORANGE, GREEN = '1F3864', '2E75B6', 'C00000', 'C55A11', '375623'
GREY, LIGHT, YELLOW = '595959', 'DDEBF7', 'FFF2CC'
F = 'Arial'
thin = Side(style='thin', color='BFBFBF')
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)


def title(ws, row, text, span):
    ws.cell(row=row, column=1, value=text).font = Font(name=F, size=12, bold=True, color='FFFFFF')
    for i in range(1, span + 1):
        ws.cell(row=row, column=i).fill = PatternFill('solid', fgColor=NAVY)
    ws.row_dimensions[row].height = 22


def note(ws, row, text, color=GREY):
    ws.cell(row=row, column=1, value=text).font = Font(name=F, size=9, color=color)


def header(ws, row, headers):
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = Font(name=F, size=9, bold=True, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor=BLUE)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = BOX
    ws.row_dimensions[row].height = 28


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def rows(ws, start, data, fmts=None, center=None, bold_rows=()):
    r = start
    for rec in data:
        for i, v in enumerate(rec, start=1):
            c = ws.cell(row=r, column=i, value=v)
            c.font = Font(name=F, size=9, bold=(r - start) in bold_rows)
            c.border = BOX
            c.alignment = Alignment(vertical='top', wrap_text=(i > 1 and isinstance(v, str)),
                                    horizontal='center' if (center and i in center) else
                                    ('right' if isinstance(v, (int, float)) else 'left'))
            if fmts and i in fmts:
                c.number_format = fmts[i]
        r += 1
    return r


wb = Workbook()

# ============================== 01_採用前提 ==============================
ws = wb.active
ws.title = '01_採用前提'
ws.sheet_view.showGridLines = False
title(ws, 1, '採用した前提条件（中間報告書 第4章・第5章）', 5)
note(ws, 2, '★の行は未確定。町の資料を受領して確定させる。金額は千円、税区分は収益的収支・総括原価とも税抜。')
header(ws, 4, ['項目', '採用値', '単位', '確定／未確定', '出所・備考'])
widths(ws, [30, 16, 10, 12, 62])
DATA = [
 ['料金算定対象期間', '令和8〜12年度', '5年', '確定', '水道法施行規則第12条第2号（長期収支試算を行った場合の3〜5年）'],
 ['有収水量（5年計）', tm.YUSHU_5Y, '㎥', '採用', '経営戦略の水需要予測。R7実績1,024,576㎥からR8推計901千㎥への段差は未説明'],
 ['給水収益（5年計）', tm.REV_5Y, '千円', '採用', '経営戦略の財政シミュレーション（税抜）。令和8年度の基本料金等免除は未反映'],
 ['維持管理費（5年計）', tm.MAINT_5Y, '千円', '採用', '経営戦略。動力費・薬品費の補正＋47,167千円を含む'],
 ['減価償却費（5年計）', tm.DEP_5Y, '千円', '採用',
  '既存資産1,855,337（R6末取得分1,822,881＋R7取得分32,456）＋調査票に基づく新規分147,642'],
 ['支払利息（既往債・5年計）', 50_585, '千円', '採用', '経営戦略'],
 ['支払利息（新規債・5年計）', 106_760, '千円', '★未確定', '借入利率2.1％は仮置き。調査票4事業とも未記入'],
 ['資産維持費（5年計）', 134_445, '千円', '★未確定', '対象資産896,312千円×3.0％×5年。対象資産の出所は平成31年度末の残高'],
 ['控除：長期前受金戻入益', -1_565_456, '千円', '採用', '経営戦略の実額。新規取得資産に対応する戻入は未計上'],
 ['控除：基準内繰入金', -tm.KIJUNNAI_5Y, '千円', '★未確定', '他会計補助金。繰出基準に基づく区分を町に確認するまで条件付き'],
 ['資産維持率', 0.03, '', '採用', '日本水道協会「水道料金算定要領」の標準値'],
 ['新規取得資産の償却年数', 38, '年', '★未確定', '経営戦略の設定。機械・電気設備の区分が未確定'],
 ['端数処理', '1円未満切捨て', '', '★未確定', '現行料金の取扱い。適用範囲を町に確認のうえ条例案と統一する'],
]
r = rows(ws, 5, DATA, fmts={2: '#,##0'}, center={3, 4})

# ============================== 02_総括原価 ==============================
ws = wb.create_sheet('02_総括原価')
ws.sheet_view.showGridLines = False
title(ws, 1, '総括原価（令和8〜12年度 5年計・税抜・千円）', 4)
note(ws, 2, '中間報告書 4-2 と同じ。基準内繰入金を控除する前提の値。')
header(ws, 4, ['費目', '5年計', '構成', '摘要'])
widths(ws, [30, 16, 10, 74])
COST = [
 ['Ⅰ　維持管理費', tm.MAINT_5Y, '＋', '人件費・動力費・薬品費・修繕費・委託料等（減価償却費を除く）'],
 ['Ⅱ　減価償却費', tm.DEP_5Y, '＋', '既存資産1,855,337＋新規分147,642'],
 ['Ⅲ　支払利息（既往債）', 50_585, '＋', ''],
 ['Ⅳ　支払利息（新規債）', 106_760, '＋', '借入利率2.1％（仮置き）'],
 ['Ⅴ　資産維持費', 134_445, '＋', '対象資産896,312千円×3.0％×5年'],
 ['Ⅵ　控除：長期前受金戻入益', -1_565_456, '－', '補助財源で取得した資産の償却見合い'],
 ['Ⅶ　控除：基準内繰入金', -tm.KIJUNNAI_5Y, '－', '他会計補助金。控除条件は未確定'],
 ['総括原価　合計', tm.TOTAL_COST_5Y, '', 'Ⅰ〜Ⅶの合計'],
 ['（参考）繰入金を控除しない場合', tm.TOTAL_COST_MAX, '', 'Ⅶを0とした場合'],
]
r = rows(ws, 5, COST, fmts={2: '#,##0'}, center={3}, bold_rows=(7, 8))
chk = sum(x[1] for x in COST[:7])
r += 1
ws.cell(row=r, column=1, value='検算（Ⅰ〜Ⅶの単純合計）').font = Font(name=F, size=9, bold=True)
c = ws.cell(row=r, column=2, value=chk)
c.number_format = '#,##0'
c.font = Font(name=F, size=9, bold=True, color=(GREEN if chk == tm.TOTAL_COST_5Y else RED))
ws.cell(row=r, column=4, value=('合計と一致' if chk == tm.TOTAL_COST_5Y else '不一致')).font = \
    Font(name=F, size=9, bold=True, color=(GREEN if chk == tm.TOTAL_COST_5Y else RED))

# ============================== 03_指標 ==============================
ws = wb.create_sheet('03_指標')
ws.sheet_view.showGridLines = False
title(ws, 1, '給水原価・供給単価・総括原価回収割合', 4)
note(ws, 2, '供給単価は3種類あり、税区分と対象年度が異なる。同じ指標として比較しないこと。')
header(ws, 4, ['指標', '値', '単位', '算定式・注意'])
widths(ws, [32, 14, 10, 76])
IND = [
 ['給水原価', round(GENKA, 1), '円/㎥', f'総括原価{tm.TOTAL_COST_5Y:,}千円×1,000÷有収水量{tm.YUSHU_5Y:,}㎥'],
 ['供給単価（R8〜R12・税抜）', round(TANKA, 1), '円/㎥',
  f'給水収益{tm.REV_5Y/5:,.0f}千円/年×1,000÷有収水量{tm.YUSHU_5Y/5:,.0f}㎥/年'],
 ['供給単価（R7実績・税込）', tm.TANKA_R7_INC, '円/㎥',
  '調定額131,428,825円÷1,024,576㎥。★将来推計と同じ指標として比較しない'],
 ['供給単価（R7実績・税抜）', tm.TANKA_R7_EX, '円/㎥',
  '調定額から消費税相当額11,948,075円を除いた119,480,750円÷1,024,576㎥'],
 ['総括原価回収割合', round(TANKA / GENKA, 3), '', '供給単価（R8〜R12・税抜）÷給水原価。決算の料金回収率とは分母が異なる'],
 ['料金回収率（決算・R5）', 0.5287, '', '給水収益120,553千円÷（総費用564,429千円−長期前受金戻入336,409千円）'],
 ['年間不足額', round((tm.TOTAL_COST_5Y - tm.REV_5Y) / 5), '千円/年', '（総括原価−給水収益）÷5年'],
 ['必要改定率（試算値）', round(K_FULL, 3), '倍',
  '総括原価÷給水収益。★基準内繰入金を控除できる場合の値。控除しない場合は2.452倍'],
 ['必要改定率（繰入金控除なし）', round(tm.TOTAL_COST_MAX / tm.REV_5Y, 3), '倍', 'Ⅶを0とした場合'],
 ['料金不足額（定義③）', round(tm.cash_short), '千円/年',
  '給水収益＋基準内繰入金−（維持管理費＋支払利息＋資産維持費）。元金償還・建設改良費を含まない'],
 ['赤字半減の改定率', round(K_HALF, 6), '倍', '定義③の半減。1.12倍は丸め表示'],
 ['3-Aの水準係数', round(M, 6), '', '★総括原価と一致するよう逆算した設定値。実調定で検証した値ではない'],
]
rows(ws, 5, IND, fmts={2: '#,##0.000'}, center={3})

# ============================== 04_パターン別 ==============================
ws = wb.create_sheet('04_パターン別')
ws.sheet_view.showGridLines = False
title(ws, 1, 'パターン別の年間料金収入と標準家庭の月額', 7)
note(ws, 2, '収入は水量ゾーン別の代表値による試算。実際の調定データによる検証は未了（中間報告書 5-6）。')
header(ws, 4, ['パターン', '改定率', '年間料金収入\n（百万円）', '現行比増収\n（百万円）',
               '供給単価\n（円/㎥・税抜）', '総括原価\n回収割合', '標準家庭\n20mm・20㎥（円）'])
widths(ws, [30, 11, 14, 14, 14, 12, 16])
cur_rev = tm.revenue(lambda v, d: tm.current(v, d))
PAT = []
for name, k, fn in [
        ('現行', 1.0, lambda v, d: tm.current(v, d)),
        ('①現行体系維持 +30％', 1.30, lambda v, d: tm.p1(v, d, 1.30)),
        ('①現行体系維持 +50％', 1.50, lambda v, d: tm.p1(v, d, 1.50)),
        ('①現行体系維持 +99％', K_FULL, lambda v, d: tm.p1(v, d, K_FULL)),
        ('②赤字半減', K_HALF, lambda v, d: tm.p1(v, d, K_HALF)),
        ('③口径別3-A', None, lambda v, d: tm.p3a(v, d, M))]:
    rev = tm.revenue(fn)
    PAT.append([name, (round(k, 3) if k else '―'), round(rev / 1e6, 1),
                round((rev - cur_rev) / 1e6, 1), round(rev / (tm.YUSHU_5Y / 5), 1),
                round(rev / (tm.YUSHU_5Y / 5) / GENKA, 3), fn(20, 20)])
PAT.append(['③口径別3-B', '―', '算定していない', '―', '―', '―', 4166])
r = rows(ws, 5, PAT, fmts={3: '#,##0.0', 4: '#,##0.0', 5: '#,##0.0', 6: '0.0%', 7: '#,##0'}, center={2})
note(ws, r + 1, '※ 3-Bは対象者の母集団が接続しないため、町全体の収入・減収額を算定していない（中間報告書 5-7）。')
note(ws, r + 2, '※ ①+99％と③3-Aの100.0％は、総括原価と一致するよう料金水準を逆算した結果である。')

# ============================== 05_料金表3A ==============================
ws = wb.create_sheet('05_料金表3A')
ws.sheet_view.showGridLines = False
title(ws, 1, 'パターン3-A　料金表案（税込・円）', 4)
note(ws, 2, '塩竈市の体系比率×水準係数。掲載は丸め後、料金例は丸め前の単価で算定している。')
header(ws, 4, ['口径', '現行メーター使用料', '案3-A 基本料金', '丸め前の値'])
widths(ws, [10, 20, 20, 20])
BASE = [[f'{k}mm', tm.METER[k], round(v * M), round(v * M, 2)] for k, v in tm.BASE_A.items()]
r = rows(ws, 5, BASE, fmts={2: '#,##0', 3: '#,##0', 4: '#,##0.00'}, center={1})
r += 1
header(ws, r, ['水量区分', '案3-A 従量料金', '現行（一般用）', '丸め前の値'])
r += 1
TIERS = [['〜10㎥', 89.1, '基本水量', round(91.3 * M, 4)],
         ['11〜20㎥', 187.9, '121円/㎥', round(192.5 * M, 4)],
         ['21〜50㎥', 252.3, '121円/㎥', round(258.5 * M, 4)],
         ['51〜100㎥', 273.8, '121円/㎥', round(280.5 * M, 4)],
         ['101㎥超', 316.8, '110円/㎥（逓減）', round(324.5 * M, 4)]]
r = rows(ws, r, TIERS, fmts={2: '#,##0.0', 4: '#,##0.0000'}, center={1})

# ============================== 06_利用者影響 ==============================
ws = wb.create_sheet('06_利用者影響')
ws.sheet_view.showGridLines = False
title(ws, 1, '利用者への影響（月額・税込）', 8)
note(ws, 2, '代表値による試算。件数は原票との照合が未了（中間報告書 5-6）。')
header(ws, 4, ['水量ゾーン', '月平均水量', '件数', '適用口径', '現行', '②赤字半減', '①+99％', '③3-A'])
widths(ws, [14, 12, 8, 10, 12, 12, 12, 12])
ZR = [[n, v, c, f'{d}mm', tm.current(v, d), tm.p1(v, d, K_HALF),
       tm.p1(v, d, K_FULL), tm.p3a(v, d, M)] for n, v, c, d in tm.ZONES]
r = rows(ws, 5, ZR, fmts={2: '#,##0.0', 3: '#,##0', 5: '#,##0', 6: '#,##0', 7: '#,##0', 8: '#,##0'},
         center={4})
r += 1
ws.cell(row=r, column=1, value='口径と使用水量の組合せによる3-Aの現行比').font = \
    Font(name=F, size=10, bold=True, color=NAVY)
r += 1
VOLS = [0, 5, 10, 20, 50, 100, 200, 500, 1000]
header(ws, r, ['口径'] + [f'{v}㎥' for v in VOLS])
r += 1
GR = [[f'{d}mm'] + [round(tm.p3a(v, d, M) / tm.current(v, d), 2) for v in VOLS]
      for d in (13, 20, 25, 40, 50)]
r = rows(ws, r, GR, fmts={i: '0.00"倍"' for i in range(2, 11)}, center={1})
r += 1
ws.cell(row=r, column=1, value='特殊用途区分（区分を廃止して口径別を適用した場合）').font = \
    Font(name=F, size=10, bold=True, color=NAVY)
r += 1
header(ws, r, ['区分', '月使用量', '現行', '②赤字半減', '①+99％', '③3-A', '③現行比', '仮定口径'])
r += 1


def yuya(v):
    return 36300 + max(0, v - 500) * 110


def sen(v):
    return 2420 + max(0, v - 10) * 242


SP = []
for nm, f, vols, dia in [('湯屋用', yuya, [500, 800], 40), ('船舶用（直接）', sen, [30, 100], 20)]:
    for v in vols:
        cur = f(v)
        SP.append([nm, v, cur, round(cur * K_HALF), round(cur * K_FULL),
                   tm.p3a(v, dia, M), round(tm.p3a(v, dia, M) / cur, 2), f'{dia}mm'])
r = rows(ws, r, SP, fmts={2: '#,##0', 3: '#,##0', 4: '#,##0', 5: '#,##0', 6: '#,##0', 7: '0.00"倍"'},
         center={8})
note(ws, r + 1, '※ 現行欄は用途別の基本料金と従量料金の合計であり、メーター使用料を含まない。'
                '船舶用（委託）は請求単位が異なるため試算していない。')

# ============================== 07_本文照合表 ==============================
ws = wb.create_sheet('07_本文照合表')
ws.sheet_view.showGridLines = False
title(ws, 1, '中間報告書の数値と本計算表の対応', 5)
note(ws, 2, 'この表の値がすべて一致していれば、本文は本計算表から再現できている。', RED)
header(ws, 4, ['中間報告書の箇所', '指標', '本文の値', '本計算表', '一致'])
widths(ws, [22, 30, 18, 18, 10])
MAP = [
 ['4-2', '総括原価（5年計・千円）', 1_175_810, tm.TOTAL_COST_5Y],
 ['4-2', '減価償却費（5年計・千円）', 2_002_979, tm.DEP_5Y],
 ['4-3', '給水原価（円/㎥）', 272.4, round(GENKA, 1)],
 ['4-3', '供給単価 R8〜R12 税抜（円/㎥）', 136.9, round(TANKA, 1)],
 ['4-3', '総括原価回収割合', 0.503, round(TANKA / GENKA, 3)],
 ['4-3', '年間不足額（千円/年）', 116_993, round((tm.TOTAL_COST_5Y - tm.REV_5Y) / 5)],
 ['4-3', '必要改定率（倍）', 1.99, round(K_FULL, 2)],
 ['4-3', '繰入金控除なしの改定率（倍）', 2.45, round(tm.TOTAL_COST_MAX / tm.REV_5Y, 2)],
 ['5-2', '供給単価 R7実績 税抜（円/㎥）', 116.6, tm.TANKA_R7_EX],
 ['7-1', '料金不足額 定義③（千円/年）', -29_488, round(tm.cash_short)],
 ['7-2', '赤字半減の改定率（倍）', 1.125, round(K_HALF, 3)],
 ['8-2', '3-A 基本料金 20mm（円）', 1_396, round(tm.BASE_A[20] * M)],
 ['9-2', '標準家庭 ①+99％（円）', 4_969, tm.p1(20, 20, K_FULL)],
 ['9-2', '標準家庭 ②赤字半減（円）', 2_809, tm.p1(20, 20, K_HALF)],
 ['9-2', '標準家庭 ③3-A（円）', 4_166, tm.p3a(20, 20, M)],
 ['9-4', '湯屋用500㎥ ③3-A（円）', 156_111, tm.p3a(500, 40, M)],
 ['8-3', '25mm 499㎥ 3-A（円）', 152_680, tm.p3a(499, 25, M)],
 ['8-3', '特例 500㎥（円）', 57_500, 500 * 115],
]
DAT = []
ok_all = True
for loc, name, book, calc in MAP:
    same = abs(float(book) - float(calc)) < 0.0005
    ok_all = ok_all and same
    DAT.append([loc, name, book, calc, '一致' if same else '不一致'])
r = rows(ws, 5, DAT, fmts={3: '#,##0.000', 4: '#,##0.000'}, center={1, 5})
for i in range(5, r):
    v = ws.cell(row=i, column=5).value
    ws.cell(row=i, column=5).font = Font(name=F, size=9, bold=True,
                                         color=(GREEN if v == '一致' else RED))
r += 1
ws.cell(row=r, column=1, value=('全項目一致' if ok_all else '不一致あり')).font = \
    Font(name=F, size=11, bold=True, color=(GREEN if ok_all else RED))
r += 2
note(ws, r, '※ 一致は本計算表と本文の算術的な整合を示すもので、前提の妥当性を示すものではない。')
note(ws, r + 1, '※ 未確定の前提は 01_採用前提 の★印を参照。')

wb.properties.title = '女川町水道料金改定の検討　採用計算表'
wb.properties.subject = '令和8年度上下水道経営指標評価書作成等業務委託'
wb.properties.creator = '若山諒太'
wb.properties.lastModifiedBy = '若山諒太'
wb.properties.category = 'ビズアップ公共コンサルティング株式会社'
wb.save(OUT)
print('作成:', os.path.abspath(OUT))
print('照合結果:', '全項目一致' if ok_all else '不一致あり')
