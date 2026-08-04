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

# ------------------------------------------------ 汚水処理費の算定シート
if "汚水処理費の算定" in wb.sheetnames:
    del wb["汚水処理費の算定"]
oc = wb.create_sheet("汚水処理費の算定")
oc.sheet_view.showGridLines = False
r = 1
oc.cell(r, 1, "経費回収率の分母（汚水処理費）の算定　― 維持管理費ベース ―").font = Font(bold=True, size=12)
r += 1
for line in ("算定式：汚水処理費（維持管理費分）＝ 経常費用 − 減価償却費 − 支払利息 − その他営業外費用",
             "　　　　資本費 ＝ 減価償却費 − 長期前受金戻入 ＋ 支払利息"
             "　…　分流式下水道等に要する経費（一般会計繰入金）でカバーされるため、分母から全額控除",
             "出典：六戸町_公共／農集_使用料改定.xlsx「財政計画（ベース）」（社人研人口推計・現行料金ベース）"):
    oc.cell(r, 1, line).font = Font(size=9, color="475569")
    r += 1
r += 1

CALC = [("経常費用（経常支出D）", "経常支出", "#,##0", False),
        ("　− 減価償却費", "減価償却費", "#,##0", False),
        ("　− 支払利息", "支払利息", "#,##0", False),
        ("　− その他営業外費用", "営業外費用その他", "#,##0", False),
        ("▶ 汚水処理費（維持管理費分）", None, "#,##0", True),
        ("　【内訳】職員給与費", "職員給与費", "#,##0", False),
        ("　【内訳】動力費", "動力費", "#,##0", False),
        ("　【内訳】修繕費", "修繕費", "#,##0", False),
        ("　【内訳】材料費", "材料費", "#,##0", False),
        ("　【内訳】その他（主に委託料）", "経費その他", "#,##0", False),
        ("（参考）長期前受金戻入", "長期前受金戻入", "#,##0", False),
        ("（参考）資本費", None, "#,##0", False),
        ("（参考）一般会計繰入金（補助金）", "補助金", "#,##0", False),
        ("（参考）繰入金 ÷ 資本費（倍）", None, "0.00", False),
        ("使用料収入（現行・改定なし）", None, "#,##0", False),
        ("経費回収率（現行・改定なし）", None, "0.0%", True)]

for biz in ("公共", "農集"):
    label = "公共下水道事業" if biz == "公共" else "農業集落排水事業"
    fp = M[biz]["財政計画"]
    cap = [fp["減価償却費"][i] - fp["長期前受金戻入"][i] + fp["支払利息"][i] for i in range(10)]
    mnt = [fp["経常支出"][i] - fp["減価償却費"][i] - fp["支払利息"][i] - fp["営業外費用その他"][i]
           for i in range(10)]
    rev = M[biz]["使用料収入"]["現行"]
    oc.cell(r, 1, f"■ {label}（単位：千円）").font = BOLD
    r += 1
    oc.cell(r, 1, "項目").fill, oc.cell(r, 1).font, oc.cell(r, 1).border = NAVY, HEAD, BOX
    for j, y in enumerate(YEARS, start=2):
        c = oc.cell(r, j, y)
        c.fill, c.font, c.alignment, c.border = NAVY, HEAD, Alignment(horizontal="center"), BOX
    r += 1
    for i, (name, key, fmt, strong) in enumerate(CALC):
        vals = (fp[key] if key else
                mnt if name.startswith("▶") else
                cap if "資本費" in name and "÷" not in name else
                [fp["補助金"][k] / cap[k] for k in range(10)] if "÷" in name else
                rev if "使用料収入" in name else
                [rev[k] / mnt[k] for k in range(10)])
        c = oc.cell(r, 1, name)
        c.font, c.border = (BOLD if strong else BODY), BOX
        if strong:
            c.fill = GREY
        elif i % 2:
            c.fill = BAND
        for j, v in enumerate(vals, start=2):
            c = oc.cell(r, j, round(v, 4) if fmt == "0.0%" else v)
            c.number_format, c.border = fmt, BOX
            c.alignment = Alignment(horizontal="right")
            c.font = BOLD if strong else BODY
            if strong:
                c.fill = GREY
            elif i % 2:
                c.fill = BAND
        r += 1
    r += 1

# 令和6年度決算との接続
R6 = {"公共": {"経常費用": 497395, "汚水処理費": 119503, "使用料収入": 56302, "回収率": 0.4711},
      "農集": {"経常費用": 106700, "汚水処理費": 32870, "使用料収入": 11831, "回収率": 0.3599}}
oc.cell(r, 1, "■ 令和6年度決算（第1回審議会資料）との接続（単位：千円）").font = BOLD
r += 1
for j, h in enumerate(["事業／項目", "R6決算", "R7見込", "差", "備考"], start=1):
    c = oc.cell(r, j, h)
    c.fill, c.font, c.alignment, c.border = NAVY, HEAD, Alignment(horizontal="center"), BOX
r += 1
for biz in ("公共", "農集"):
    fp, a = M[biz]["財政計画"], R6[biz]
    mnt7 = fp["経常支出"][0] - fp["減価償却費"][0] - fp["支払利息"][0] - fp["営業外費用その他"][0]
    resid7 = fp["減価償却費"][0] + fp["支払利息"][0] + fp["営業外費用その他"][0]
    rows = [(f"{biz}　経常費用", a["経常費用"], fp["経常支出"][0], ""),
            (f"{biz}　汚水処理費（維持管理費分）", a["汚水処理費"], mnt7, ""),
            (f"{biz}　差引（減価償却費＋支払利息＋その他）", a["経常費用"] - a["汚水処理費"], resid7,
             "R6決算とR7見込がほぼ一致 → 同一基準で算定されていることの確認")]
    for i, (name, v6, v7, memo) in enumerate(rows):
        oc.cell(r, 1, name).font = BODY
        for j, v in enumerate((v6, v7, v7 - v6), start=2):
            c = oc.cell(r, j, v)
            c.number_format, c.alignment, c.font = "#,##0;△#,##0", Alignment(horizontal="right"), BODY
        oc.cell(r, 5, memo).font = Font(size=9, color="475569")
        for j in range(1, 6):
            oc.cell(r, j).border = BOX
            if i == 2:
                oc.cell(r, j).fill = GREY
        r += 1
r += 1
for line in ("※ 令和6年度決算の汚水処理費（公共119,503千円・農集32,870千円）も、"
             "経常費用から減価償却費・支払利息を控除した維持管理費ベースで算定されており、本シミュレーションと同一基準。",
             "※【内訳】その他は、職員給与費・動力費・修繕費・材料費以外の経費（主に委託料）。"
             "公共下水道ではR7の汚水処理費168,331千円のうち155,280千円（92%）を占める。",
             "※ 公共下水道はR6→R7で維持管理費が119,503千円→168,331千円（＋40.9%）と増加する見込みのため、"
             "経費回収率は47.11%→34.4%に低下する。どの費目の増加によるものかは要確認（確認事項A）。"):
    oc.cell(r, 1, line).font = Font(size=9, color="B45309")
    r += 1

oc.column_dimensions["A"].width = 34
for j in range(2, 12):
    oc.column_dimensions[get_column_letter(j)].width = 12
oc.freeze_panes = "B1"

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
