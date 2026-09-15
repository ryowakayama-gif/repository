# -*- coding: utf-8 -*-
"""⑦保険料算定シートを、所得段階別第1号被保険者数（13段階・令和7年度末実績）で更新する。

出所：介護保険事業状況報告（年報・令和7年度／様式1 所得段階別・列20 年度末現在被保険者数）
"""
import importlib.util
import io
import contextlib
import warnings

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

warnings.filterwarnings("ignore")

BASE = "/home/user/repository/kawasaki_project"
XLSX = f"{BASE}/05_試算・管理シート/川崎町_第10期_計画見込量_R8.9.15.xlsx"

spec = importlib.util.spec_from_file_location(
    "m", f"{BASE}/07_ソーススクリプト/project_mikomi_R8.9.15.py")
M = importlib.util.module_from_spec(spec)
with contextlib.redirect_stdout(io.StringIO()):
    spec.loader.exec_module(M)

YEARS, HIHO3 = M.YEARS, M.hihoken3
SHUNO3 = sum(M.FIN[y]["保険料収納必要額"] for y in YEARS)

RITSU = [0.455, 0.685, 0.690, 0.90, 1.00, 1.20, 1.30, 1.50, 1.70,
         1.90, 2.10, 2.30, 2.40]
R6 = [412, 308, 304, 383, 699, 375, 419, 218, 67, 25, 13, 7, 31]
R7 = [383, 298, 305, 343, 679, 410, 427, 232, 85, 26, 17, 7, 28]
KEIGEN = {0: 0.285, 1: 0.485, 2: 0.685}      # 公費による軽減後の乗率
FUND = 152_500_000


def hosei(w):
    return sum(n * r for n, r in zip(w, RITSU)) / sum(w)


def premium(h, torikuzushi=0.0):
    return SHUNO3 / 0.96 / 12 / (HIHO3 * h) - torikuzushi / (12 * HIHO3 * h)


H_NOW = M.HOSEI                 # 見える化システムに登録されている補正係数
H_R6, H_R7 = hosei(R6), hosei(R7)
BASE_PREM = premium(H_R7)

wb = openpyxl.load_workbook(XLSX)
ws = wb["⑦保険料算定"]

THIN = Side(style="thin", color="B0B0B0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD = PatternFill("solid", fgColor="DCE6F1")
BOLD = Font(bold=True)
R = Alignment(horizontal="right", vertical="center")
L = Alignment(horizontal="left", vertical="center")
C = Alignment(horizontal="center", vertical="center")


def put(row, col, value, *, head=False, num=None, align=None):
    c = ws.cell(row, col, value)
    c.border = BORDER
    if head:
        c.fill, c.font = HEAD, BOLD
    if num:
        c.number_format = num
    c.alignment = align or (R if isinstance(value, (int, float)) else L)
    return c


def clear(row_from, row_to):
    for r in range(row_from, row_to + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            cell.value = None
            cell.border = Border()
            cell.fill = PatternFill()
            cell.font = Font()


# ── 1. 補正係数と保険料基準額（14〜18行）を令和7年度末実績に差し替え ──────────
ws.cell(15, 1).value = "所得段階別加入割合による補正係数（令和7年度末実績・13段階）"
ws.cell(15, 2).value = round(H_R7, 6)
ws.cell(16, 2).value = HIHO3 * H_R7
ws.cell(18, 2).value = BASE_PREM
for r in (15, 16, 18):
    ws.cell(r, 2).number_format = "#,##0.000000" if r == 15 else "#,##0.0"
ws.cell(18, 2).number_format = "#,##0.00"

# ── 2. 21行以降を全面的に書き直す ────────────────────────────────────
clear(20, ws.max_row)
row = 21

put(row, 1, "【内訳】所得段階別第1号被保険者数（令和7年度末・年報 様式1 所得段階別）",
    head=True)
row += 1
put(row, 1, "段階", head=True, align=C)
put(row, 2, "基準額に対する乗率", head=True, align=C)
put(row, 3, "被保険者数（人）", head=True, align=C)
put(row, 4, "構成比", head=True, align=C)
put(row, 5, "加重（人×乗率）", head=True, align=C)
row += 1
tot = sum(R7)
for i, (n, r) in enumerate(zip(R7, RITSU), 1):
    label = f"第{i}段階"
    if i - 1 in KEIGEN:
        label += f"（軽減後 {KEIGEN[i - 1]:.3f}）"
    put(row, 1, label)
    put(row, 2, r, num="0.000")
    put(row, 3, n, num="#,##0")
    put(row, 4, n / tot, num="0.0%")
    put(row, 5, n * r, num="#,##0.0")
    row += 1
put(row, 1, "計", head=True)
put(row, 2, "─", head=True, align=C)
put(row, 3, tot, head=True, num="#,##0")
put(row, 4, 1.0, head=True, num="0.0%")
put(row, 5, sum(n * r for n, r in zip(R7, RITSU)), head=True, num="#,##0.0")
row += 1
put(row, 1, "補正係数（＝加重計÷被保険者数計）", head=True)
put(row, 2, H_R7, head=True, num="0.000000")
row += 3

# ── 3. 分布別の比較 ───────────────────────────────────────────────
put(row, 1, "【比較】所得段階別の分布のとり方による保険料基準額", head=True)
row += 1
put(row, 1, "分布", head=True, align=C)
put(row, 2, "補正係数", head=True, align=C)
put(row, 3, "補正後被保険者数", head=True, align=C)
put(row, 4, "A 取崩なし", head=True, align=C)
put(row, 5, "採用との差", head=True, align=C)
row += 1
for nm, h in [("見える化システムの現在の登録値（第10〜13段階が0人）", H_NOW),
              ("令和6年度末の実績分布（13段階）", H_R6),
              ("令和7年度末の実績分布（13段階）★採用", H_R7)]:
    put(row, 1, nm)
    put(row, 2, h, num="0.000000")
    put(row, 3, HIHO3 * h, num="#,##0.0")
    put(row, 4, premium(h), num="#,##0")
    put(row, 5, premium(h) - BASE_PREM, num="+#,##0;-#,##0;0")
    row += 1
row += 1
put(row, 1, "※ 見える化システムには第10〜13段階が0人として登録されているため、"
            "補正係数が実績より 0.0497 低く出て、保険料が 352円／月 高く算定される。")
row += 3

# ── 4. 準備基金の取崩し ──────────────────────────────────────────
put(row, 1, "【感度分析】介護給付費準備基金の取崩し"
            "（補正係数 1.013748 ＝ 令和7年度末実績を前提）", head=True)
row += 1
put(row, 1, "取崩額（千円）", head=True, align=C)
put(row, 2, "月額換算（円）", head=True, align=C)
put(row, 3, "保険料基準額（月額）", head=True, align=C)
put(row, 4, "第9期6,500円との差", head=True, align=C)
row += 1
for t in (0, 30_000_000, 50_000_000, 76_250_000, 100_000_000, 152_500_000):
    p = premium(H_R7, t)
    put(row, 1, t / 1000, num="#,##0")
    put(row, 2, p - BASE_PREM, num="#,##0.00")
    put(row, 3, p, num="#,##0")
    put(row, 4, p - 6500, num="+#,##0;-#,##0;0")
    row += 1
row += 1
put(row, 1, "※ 基金残高は令和8年度末見込み152,500千円（町財政担当に確定値を照会中）。"
            "76,250千円は50％取崩し、152,500千円は全額取崩し。")
row += 1
put(row, 1, "※ 第9期は78,000千円（月額698.59円相当）を取り崩し、"
            "条例上の基準額を6,508.33円としている。")

ws.column_dimensions["A"].width = 58
for col in "BCDE":
    ws.column_dimensions[col].width = 20

wb.save(XLSX)

print("■ 更新後の⑦保険料算定シート")
print(f"補正係数（令和7年度末実績・13段階）　{H_R7:.6f}")
print(f"補正後被保険者数（3か年計）　　　　　{HIHO3 * H_R7:,.1f} 人")
print(f"保険料基準額（月額・取崩なし）　　　 {BASE_PREM:,.2f} 円")
print(f"　　　　　　 （50％取崩し）　　　　　{premium(H_R7, FUND * .5):,.2f} 円")
print(f"　　　　　　 （全額取崩し）　　　　　{premium(H_R7, FUND):,.2f} 円")
print(f"現在の登録値との差　{BASE_PREM - premium(H_NOW):+,.1f} 円／月")
