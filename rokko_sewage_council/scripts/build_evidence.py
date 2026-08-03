# -*- coding: utf-8 -*-
"""エビデンス集に第2回審議会資料の算定根拠（パターン別シミュレーション結果）を追加する."""
import json
import pathlib
import shutil

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HERE = pathlib.Path(__file__).resolve().parent.parent
SRC = HERE / "source" / "六戸町_使用料改定_成果物エビデンス集.xlsx"
DST = HERE / "output" / "六戸町_使用料改定_成果物エビデンス集.xlsx"
M = json.loads((HERE / "data" / "metrics.json").read_text(encoding="utf-8"))

YEARS = ["令和7年度", "令和8年度", "令和9年度", "令和10年度", "令和11年度",
         "令和12年度", "令和13年度", "令和14年度", "令和15年度", "令和16年度"]
ROWS = [("現行（改定なし）", "現行"), ("パターン①（標準型）2か年", "①2"), ("パターン①（標準型）3か年", "①3"),
        ("パターン②（家庭軽減型）2か年", "②2"), ("パターン②（家庭軽減型）3か年", "②3"),
        ("パターン④（段階累進型）2か年", "④2"), ("パターン④（段階累進型）3か年", "④3")]

RATES = {
    "現行":   [1000, 120, 120, 130, 130, 140, 140, 140, 160],
    "①中間": [1250, 135, 138, 145, 148, 158, 163, 173, 193],
    "①最終": [1500, 150, 155, 160, 165, 175, 185, 205, 225],
    "②中間": [1100, 150, 153, 160, 163, 173, 178, 188, 208],
    "②最終": [1200, 180, 185, 190, 195, 205, 215, 235, 255],
    "④中間": [1200, 140, 143, 150, 153, 163, 168, 178, 198],
    "④最終": [1400, 160, 165, 170, 175, 185, 195, 215, 235],
}
BLOCKS = ["基本(0〜10㎥)", "11〜20㎥", "21〜30㎥", "31〜40㎥", "41〜50㎥",
          "51〜70㎥", "71〜100㎥", "101〜150㎥", "151㎥〜"]
CAPS = [20, 30, 40, 50, 70, 100, 150]


def fee(rate, vol):
    total, prev = rate[0], 10
    for i, cap in enumerate(CAPS, start=1):
        if vol <= prev:
            break
        total += (min(vol, cap) - prev) * rate[i]
        prev = cap
    if vol > 150:
        total += (vol - 150) * rate[8]
    return total


NAVY = PatternFill("solid", fgColor="1E3A8A")
GREY = PatternFill("solid", fgColor="F1F5F9")
BAND = PatternFill("solid", fgColor="F8FAFC")
THIN = Side(style="thin", color="CBD5E1")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD = Font(bold=True, color="FFFFFF", size=10)
BOLD = Font(bold=True, size=10)
BODY = Font(size=10)

DST.parent.mkdir(parents=True, exist_ok=True)
shutil.copy(SRC, DST)
wb = openpyxl.load_workbook(DST)

# ------------------------------------------------ 成果物一覧の更新
ws = wb["成果物一覧"]
for row in ws.iter_rows():
    if row[1].value and "第2回審議会資料" in str(row[1].value):
        row[1].value = "六戸町_第2回審議会資料.pptx"
        row[2].value = "PPTX（31スライド）"
        row[4].value = ("第2回審議会資料。パターン①②④それぞれの収支シミュレーション（01-2〜01-4）・"
                        "段階単価一覧（02-2〜02-4）・経常収支比率見込み（03-1〜03-3）・"
                        "家庭/事業所別影響額（04-1〜04-3）を3パターン並列で収録。"
                        "指標値は各使用料改定.xlsx「パターン別比較」（社人研人口推計）から転記。")
    if row[0].value == "No.":
        hdr_row = row[0].row
last = ws.max_row
ws.cell(last + 1, 1, len(list(ws.iter_rows(min_row=hdr_row + 1, max_row=last))) + 1)
ws.cell(last + 1, 2, "本ブック「パターン別SIM結果」シート")
ws.cell(last + 1, 3, "XLSX（算定根拠）")
ws.cell(last + 1, 4, "✅ 追加")
ws.cell(last + 1, 5, "第2回審議会資料に掲載した経常収支比率・経費回収率・モデルケース別月額の元数値。")
for c in range(1, 6):
    ws.cell(last + 1, c).font = BODY
    ws.cell(last + 1, c).alignment = Alignment(vertical="center", wrap_text=True)

# ------------------------------------------------ 料金データ確認の補正
ws = wb["料金データ確認"]
for r, blk in enumerate(BLOCKS, start=4):
    for c, key in enumerate(["現行", "①中間", "①最終", "②中間", "②最終", "④中間", "④最終"], start=2):
        ws.cell(r, c).value = RATES[key][r - 4]
ws.cell(13, 1).value = ("※ R8中間単価は算定ブックの数式どおり（現行＋最終）÷2の四捨五入。"
                        "第1回審議会資料は切り捨てで提示しており、①71〜100㎥・101〜150㎥・151㎥〜、"
                        "②21〜30㎥・41〜50㎥の5区分で1円の差があります（R8単年の収入影響は約16万円）。")
ws.cell(13, 1).font = Font(size=9, color="B45309")
ws.cell(16, 1).value = ("20㎥時の月額：現行 税抜2,200円／税込2,420円　→　改定後いずれも 税抜3,000円／税込3,300円"
                        "（税抜＋800円・税込＋880円・＋36.4%）")
for r, key in ((18, "①最終"), (19, "②最終"), (20, "④最終")):
    ws.cell(r, 2).value = f"{fee(RATES[key], 20):,}円"
    ws.cell(r, 4).value = f"税込 {int(round(fee(RATES[key], 20) * 1.1)):,}円"
ws.cell(17, 4).value = "税込 2,420円"

# ------------------------------------------------ パターン別SIM結果シート
if "パターン別SIM結果" in wb.sheetnames:
    del wb["パターン別SIM結果"]
sim = wb.create_sheet("パターン別SIM結果")
sim.sheet_view.showGridLines = False
r = 1
sim.cell(r, 1, "第2回審議会資料　掲載数値の算定根拠（パターン別シミュレーション結果）").font = Font(bold=True, size=12)
r += 1
sim.cell(r, 1, "出典：六戸町_公共_使用料改定.xlsx／六戸町_農集_使用料改定.xlsx「パターン別比較」シート"
               "（人口推計＝社人研／使用料収入以外の収支項目は全パターン共通）").font = Font(size=9, color="475569")
r += 2

for biz in ("公共", "農集"):
    label = "公共下水道事業" if biz == "公共" else "農業集落排水事業"
    for metric, goal in (("経常収支比率", "目標：100%以上（単年度黒字）"),
                         ("経費回収率", "目標：47%以上")):
        sim.cell(r, 1, f"■ {label}　{metric}　{goal}").font = BOLD
        r += 1
        sim.cell(r, 1, "パターン／年度").fill = NAVY
        sim.cell(r, 1).font = HEAD
        for j, y in enumerate(YEARS, start=2):
            c = sim.cell(r, j, y)
            c.fill, c.font, c.alignment, c.border = NAVY, HEAD, Alignment(horizontal="center"), BOX
        sim.cell(r, 1).border = BOX
        r += 1
        for i, (name, key) in enumerate(ROWS):
            c = sim.cell(r, 1, name)
            c.font, c.border = (BOLD if "②" in name else BODY), BOX
            if i % 2:
                c.fill = BAND
            for j, v in enumerate(M[biz][metric][key], start=2):
                c = sim.cell(r, j, round(v, 4))
                c.number_format = "0.0%"
                c.alignment, c.border = Alignment(horizontal="center"), BOX
                c.font = BOLD if "②" in name else BODY
                if i % 2:
                    c.fill = BAND
            r += 1
        r += 1

    sim.cell(r, 1, f"■ {label}　使用料収入（千円）").font = BOLD
    r += 1
    sim.cell(r, 1, "パターン／年度").fill = NAVY
    sim.cell(r, 1).font = HEAD
    sim.cell(r, 1).border = BOX
    for j, y in enumerate(YEARS, start=2):
        c = sim.cell(r, j, y)
        c.fill, c.font, c.alignment, c.border = NAVY, HEAD, Alignment(horizontal="center"), BOX
    r += 1
    for i, (name, key) in enumerate(ROWS):
        c = sim.cell(r, 1, name)
        c.font, c.border = (BOLD if "②" in name else BODY), BOX
        if i % 2:
            c.fill = BAND
        for j, v in enumerate(M[biz]["使用料収入"][key], start=2):
            c = sim.cell(r, j, v)
            c.number_format = "#,##0"
            c.alignment, c.border = Alignment(horizontal="right"), BOX
            c.font = BOLD if "②" in name else BODY
            if i % 2:
                c.fill = BAND
        r += 1
    r += 1

# モデルケース別月額
sim.cell(r, 1, "■ モデルケース別　月額使用料（円）　※各区分の単価を段階累進で積み上げて算定").font = BOLD
r += 1
hdr = ["使用水量", "現行(税抜)", "現行(税込)"]
for p in ("①", "②", "④"):
    hdr += [f"{p}最終(税抜)", f"{p}最終(税込)", f"{p}増加額(税込)"]
for j, h in enumerate(hdr, start=1):
    c = sim.cell(r, j, h)
    c.fill, c.font, c.alignment, c.border = NAVY, HEAD, Alignment(horizontal="center", wrap_text=True), BOX
r += 1
for i, vol in enumerate([10, 20, 30, 40, 50, 100, 200, 500]):
    cur = fee(RATES["現行"], vol)
    curt = int(round(cur * 1.1))
    vals = [f"{vol}㎥", cur, curt]
    for p in ("①", "②", "④"):
        v = fee(RATES[p + "最終"], vol)
        vals += [v, int(round(v * 1.1)), int(round(v * 1.1)) - curt]
    for j, v in enumerate(vals, start=1):
        c = sim.cell(r, j, v)
        c.border = BOX
        c.font = BOLD if vol == 20 else BODY
        if j > 1:
            c.number_format = "#,##0"
            c.alignment = Alignment(horizontal="right")
        if vol == 20:
            c.fill = GREY
        elif i % 2:
            c.fill = BAND
    r += 1
r += 1
sim.cell(r, 1, "※ 20㎥（一般家庭の標準使用量）は3パターンとも税抜3,000円／税込3,300円で共通（現行比＋880円・＋36.4%）。").font = Font(size=9, color="475569")

sim.column_dimensions["A"].width = 26
for j in range(2, len(hdr) + 1):
    sim.column_dimensions[get_column_letter(j)].width = 12
sim.freeze_panes = "B1"

wb.save(DST)
print("saved", DST)
