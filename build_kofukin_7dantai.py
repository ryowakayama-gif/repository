# -*- coding: utf-8 -*-
"""交付金獲得状況（対象7団体）の一覧表を作成する.

出力 output/05_交付金獲得状況_7団体.xlsx
     総括 / 3か年の得点と順位 / 条件を揃えた比較 / 令和6年度の交付額 /
     目標別内訳 / 取り戻す指標
"""

import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import data_kofukin_7dantai as D

OUT_DIR = "/home/user/repository/output"
os.makedirs(OUT_DIR, exist_ok=True)
OUT = os.path.join(OUT_DIR, "05_交付金獲得状況_7団体.xlsx")

COLORS = {
    "header":  "1F3864",
    "subhead": "2E75B6",
    "band":    "DDEBF7",
    "alt":     "F7FAFC",
    "warn":    "FFE2E2",   # 要注意（交付額の未消化・平均を大きく下回る）
    "good":    "E2F0D9",   # 良好
    "white":   "FFFFFF",
}
THIN = Side(border_style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# 令和8年度の得点順
NAMES = ["平取町", "西会津町", "矢祭町", "富良野市", "北広島市", "奥尻町", "音威子府村"]
MOKUHYO = ["推進Ⅰ", "推進Ⅱ", "推進Ⅲ", "推進Ⅳ", "支援Ⅰ", "支援Ⅱ", "支援Ⅲ", "支援Ⅳ"]


def title(ws, text, ncol, note=None):
    ws.cell(1, 1, text).font = Font(bold=True, size=13, color=COLORS["white"])
    ws.cell(1, 1).fill = PatternFill("solid", fgColor=COLORS["header"])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    ws.cell(1, 1).alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height = 24
    if note:
        ws.cell(2, 1, note).font = Font(size=9, color="595959")
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncol)


def head(ws, row, labels, widths=None):
    for j, lab in enumerate(labels, 1):
        c = ws.cell(row, j, lab)
        c.font = Font(bold=True, size=10, color=COLORS["white"])
        c.fill = PatternFill("solid", fgColor=COLORS["subhead"])
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[row].height = 30
    if widths:
        for j, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = ws.cell(row + 1, 1)


def put(ws, row, values, fill=None, bolds=()):
    for j, v in enumerate(values, 1):
        c = ws.cell(row, j, v)
        c.border = BORDER
        c.alignment = Alignment(horizontal="center" if not isinstance(v, str) or len(v) < 12
                                else "left", vertical="center", wrap_text=True)
        c.font = Font(size=10, bold=(j in bolds))
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            c.number_format = "#,##0"
        if fill:
            c.fill = PatternFill("solid", fgColor=fill)


wb = Workbook()

# ── 1 総括 ────────────────────────────────────────────────
ws = wb.active
ws.title = "総括"
title(ws, "交付金の獲得状況　対象7団体の総括", 11,
      "出典 厚生労働省 保険者機能強化推進交付金及び介護保険保険者努力支援交付金"
      "（市町村分）に係る全国集計結果（令和6〜8年度）／母数1,741市町村／800点満点")
head(ws, 4, ["団体", "都道府県", "第1号\n被保険者", "規模\n区分", "過疎",
             "令和8年度\n得点", "全国順位", "同規模・同条件\n順位", "県内順位",
             "令和6年度\n交付額消化率", "令和6年度\n1人当たり(円)"],
     [12, 9, 10, 6, 7, 10, 10, 15, 11, 14, 14])
r = 5
for nm in NAMES:
    e = D.DANTAI[nm]
    k = e["R6交付額"]
    rate = k["合計確定"] / k["合計案"] * 100
    fill = COLORS["warn"] if rate < 99.5 else (COLORS["alt"] if r % 2 else COLORS["white"])
    put(ws, r, [nm, e["都道府県"], e["第1号被保険者"], e["規模区分"], e["過疎"],
                e["R8"]["合計"], e["R8"]["順位"],
                f'{e["同条件"]["順位"]}位/{e["同条件"]["団体数"]}',
                f'{e["県内"]["順位"]}位/{e["県内"]["団体数"]}',
                f"{rate:.1f}％", k["1人当たり"]],
        fill=fill, bolds=(1, 6, 10, 11))
    # 1人当たりが全国平均795円を大きく下回る場合は色を変える
    if k["1人当たり"] < 700:
        ws.cell(r, 11).fill = PatternFill("solid", fgColor=COLORS["warn"])
    r += 1
put(ws, r, ["全国平均", "", "", "", "", D.ZEN["R8"]["合計"], "", "", "", "100.0％", 795],
    fill=COLORS["band"], bolds=(1, 6))
ws.cell(r, 6).number_format = "0.0"
r += 2
for line in [
    "■ 赤の行は令和6年度に交付見込額（案）を満額受け取れていない団体です。",
    "　 全国1,741団体のうち案を下回ったのは24団体（1.4％）で、音威子府村は消化率が最も低く、平取町は3番目です。",
    "■ 同規模・同条件は、規模区分と過疎地域該当がいずれも同じ団体の中での順位です。",
    "　 富良野市は全国866位ですが同条件では115位／285位に上がり、北広島市は全国1,218位から321位／380位に下がります。",
    "■ 1人当たりの赤は全国平均795円を大きく下回るものです。",
    "　 奥尻町は消化率100％ですが1人当たり552円で全国1,664位／1,741、区分1の447団体では412位です。",
    "　 平取町・音威子府村は受け取れなかった額があり、奥尻町は配分そのものが小さいという別の問題です。",
]:
    ws.cell(r, 1, line).font = Font(size=10)
    r += 1

# ── 2 3か年の得点と順位 ──────────────────────────────────
ws = wb.create_sheet("3か年の得点と順位")
title(ws, "推進・支援合計の3か年推移（800点満点・母数1,741市町村）", 8,
      "評価年度と交付年度は1年ずれる。令和8年度交付金は令和7年度に実施した取組を評価したもの。")
head(ws, 4, ["団体", "R6得点", "R6順位", "R7得点", "R7順位", "R8得点", "R8順位", "R6→R8"],
     [12, 9, 9, 9, 9, 9, 9, 9])
r = 5
for nm in NAMES:
    e = D.DANTAI[nm]
    d = e["R8"]["合計"] - e["R6"]["合計"]
    fill = COLORS["good"] if d >= 100 else (COLORS["warn"] if d < 0 else (COLORS["alt"] if r % 2 else COLORS["white"]))
    put(ws, r, [nm, e["R6"]["合計"], e["R6"]["順位"], e["R7"]["合計"], e["R7"]["順位"],
                e["R8"]["合計"], e["R8"]["順位"], d], fill=fill, bolds=(1, 6, 8))
    r += 1
# 全国平均は 422.35 / 434.99 / 455.13 の丸め値。増減は丸め前の差（32.8）を用いる。
put(ws, r, ["全国平均", D.ZEN["R6"]["合計"], "", D.ZEN["R7"]["合計"], "",
            D.ZEN["R8"]["合計"], "", 32.8], fill=COLORS["band"], bolds=(1,))
for j in (2, 4, 6, 8):
    ws.cell(r, j).number_format = "0.0"
r += 2
ws.cell(r, 1, "■ 全国平均は3年で32.8点上がっています。平均が上がる中で順位を保つには同じだけ伸ばす必要があります。").font = Font(size=10)

# ── 3 条件を揃えた比較 ────────────────────────────────────
ws = wb.create_sheet("条件を揃えた比較")
title(ws, "規模区分別・過疎地域該当別の全国分布（令和8年度）", 5,
      "交付金の得点は規模と強く相関する（区分1と区分5で120.2点の差）。過疎地域は非該当より平均36.1点低い。")
head(ws, 4, ["区分", "第1号被保険者数", "団体数", "平均", "中央値"], [10, 20, 10, 10, 10])
BAND = {1: "3千人未満", 2: "3千〜1万人未満", 3: "1万〜5万人未満",
        4: "5万〜10万人未満", 5: "10万人以上"}
r = 5
for k in [1, 2, 3, 4, 5]:
    v = D.KUBUN[k]
    put(ws, r, [f"区分{k}", BAND[k], v["団体数"], v["平均"], v["中央値"]],
        fill=COLORS["alt"] if r % 2 else COLORS["white"])
    r += 1
r += 1
head(ws, r, ["過疎地域該当", "", "団体数", "平均", "中央値"])
r += 1
for lab in ["該当", "非該当"]:
    v = D.KASO[lab]
    put(ws, r, [lab, "一部過疎を含む・R5.4.1" if lab == "該当" else "", v["団体数"], v["平均"], v["中央値"]],
        fill=COLORS["alt"] if r % 2 else COLORS["white"])
    r += 1

# ── 4 令和6年度の交付額 ──────────────────────────────────
ws = wb.create_sheet("令和6年度の交付額")
title(ws, "令和6年度　交付見込額（案）と確定額　単位：千円", 9,
      "確定額は「交付見込額（案）を下回る所要見込額の場合は当該額」となる。"
      "金額が公表されているのは令和6年度分のみ。")
head(ws, 4, ["団体", "得点", "推進\n案", "推進\n確定", "支援\n案", "支援\n確定",
             "合計\n案", "合計\n確定", "受け取れな\nかった額"],
     [12, 8, 10, 10, 10, 10, 11, 11, 12])
r = 5
for nm in NAMES:
    e = D.DANTAI[nm]
    k = e["R6交付額"]
    loss = k["合計案"] - k["合計確定"]
    fill = COLORS["warn"] if loss > 0 else (COLORS["alt"] if r % 2 else COLORS["white"])
    put(ws, r, [nm, e["R6"]["合計"], k["推進案"], k["推進確定"], k["支援案"], k["支援確定"],
                k["合計案"], k["合計確定"], loss], fill=fill, bolds=(1, 8, 9))
    r += 1
r += 2
ws.cell(r, 1, "令和6年度に交付見込額（案）を下回った全国24団体（消化率の低い順）").font = Font(bold=True, size=11)
r += 1
head(ws, r, ["順", "団体", "得点", "第1号被保険者", "合計案", "合計確定", "消化率(％)", "", ""])
r += 1
for m in D.MISHOKA:
    mark = m["団体"] in NAMES
    put(ws, r, [m["順"], m["団体"], m["得点"], m["第1号被保険者"], m["合計案"], m["合計確定"], m["消化率"], "", ""],
        fill=COLORS["warn"] if mark else (COLORS["alt"] if r % 2 else COLORS["white"]),
        bolds=(2,) if mark else ())
    r += 1

# ── 5 目標別内訳 ──────────────────────────────────────────
ws = wb.create_sheet("目標別内訳")
title(ws, "目標別の得点（各100点満点）と全国平均との差", 10,
      "目標Ⅳ（成果指標群）は2つの交付金で同一の指標であり、同じ得点が推進・支援の双方に計上される。")
head(ws, 4, ["団体", "年度"] + MOKUHYO, [12, 7] + [9] * 8)
r = 5
for nm in NAMES:
    e = D.DANTAI[nm]
    for yr in ["R6", "R7", "R8"]:
        vals = [e[yr][k] for k in MOKUHYO]
        put(ws, r, [nm if yr == "R6" else "", yr] + vals,
            fill=COLORS["band"] if yr == "R8" else COLORS["white"], bolds=(1,))
        for j, k in enumerate(MOKUHYO, 3):
            diff = e[yr][k] - D.ZEN[yr][k]
            c = ws.cell(r, j)
            if diff <= -20:
                c.fill = PatternFill("solid", fgColor=COLORS["warn"])
            elif diff >= 20:
                c.fill = PatternFill("solid", fgColor=COLORS["good"])
        r += 1
    r += 0
for yr in ["R6", "R7", "R8"]:
    put(ws, r, ["全国平均", yr] + [D.ZEN[yr][k] for k in MOKUHYO], fill=COLORS["alt"], bolds=(1,))
    r += 1
r += 1
for line in ["■ 緑は全国平均を20点以上上回る目標、赤は20点以上下回る目標です。",
             "■ 成果指標群（目標Ⅳ）は小規模団体ほど年度間の振れが大きくなります"
             "（区分1の標準偏差23.1に対し区分5は13.2）。単年度の上下を取組の成否として読まないでください。"]:
    ws.cell(r, 1, line).font = Font(size=10)
    r += 1

# ── 6 取り戻す指標 ────────────────────────────────────────
ws = wb.create_sheet("取り戻す指標")
title(ws, "同規模・同条件の団体の多くが得点しているのに失点している指標", 9,
      "同規模かつ過疎条件の同じ団体の30％以上が満点を取っている指標に限る。"
      "体制・取組指標群は計画・要綱への記載で得点化できる（得点に現れるのは令和10年度交付金）。")
head(ws, 4, ["団体", "交付金", "目標", "指標群", "指標名", "自団体", "満点", "失点", "同規模\n満点率(％)"],
     [12, 8, 8, 12, 40, 8, 8, 8, 12])
r = 5
for nm in NAMES:
    rows = D.TORIMODOSHI[nm]
    tot = sum(x["失"] for x in rows)
    for i, x in enumerate(rows):
        put(ws, r, [nm if i == 0 else "", x["交付金"], x["目標"], x["指標群"], x["指標名"],
                    x["自"], x["満"], x["失"], x["同規模満点率"]],
            fill=COLORS["alt"] if r % 2 else COLORS["white"], bolds=(1,))
        r += 1
    put(ws, r, ["", "", "", "", f"{nm}　取り戻せる可能性の合計", "", "", tot, ""],
        fill=COLORS["band"], bolds=(5, 8))
    r += 1

wb.save(OUT)
print("作成しました:", OUT)
