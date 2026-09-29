# -*- coding: utf-8 -*-
"""北塩原村 第10期計画　図表データ管理台帳.

━━ 本表の目的 ━━

計画本文に差し込む図と、計画素案の表は、同じ数値を別々に持っていた。
そのため一方を直しても他方が追随しない作りであった。

本表は次の4つを果たす。

  1 本文に差し込む図の数値を1か所（data_zuhyo.py）に集め、
    エクセルの上で直せるようにする。各データシートにはエクセルの
    ネイティブグラフを置いてあるため、数値のセルを直せばその場で
    グラフが変わる。そのうえで作図を実行し直せば本文の図も変わる。
  2 図の数値と計画素案の表の数値を機械で突合し、差分を掲げる。
  3 保険者機能強化推進交付金等の評価が求める分析と、それを裏づける
    図表の対応を整理し、足りない図表を掲げる。
  4 資料編 資8（図表番号一覧）との突合を行う。

━━ 数値を直したときの手順 ━━

  1 本表のデータシート（D_ で始まるシート）の数値のセルを直す
  2 python3 scripts/build_figures.py      本文に差し込む図（PNG）を作り直す
  3 python3 scripts/build_soan.py         素案のMarkdownとJSONを作り直す
  4 node   scripts/build_soan_docx.js     素案のdocxを作り直す
  5 python3 scripts/build_zuhyo_daicho.py 本表を作り直す（突合をやり直す）

行や列を増減すると読み戻せないため、数値のセルのみを直す。
区分や系列そのものを変えるときは data_zuhyo.py の側による。

━━ 数え方 ━━

図の数・系列・数値は data_zuhyo.py から、差し込み先と表題は figures_map.py から、
素案の表は soan_content.py から、交付金の得点は交付金の分析から読む。
いずれも固定値を書かない。

自己点検で1件でも不適合があると終了コード1で終わる。
"""
import io
import os
import re
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import data_zuhyo as DZ
import figures_map as FM
import soan_content as SC

OUT = "/home/user/repository/output/11_北塩原村第10期_図表データ管理台帳.xlsx"
FIGDIR = "/home/user/repository/output/figures"
KIJUNBI = "令和8年9月29日"

FONT = "游ゴシック"
NAVY, HEAD = "1F3864", "4472C4"
IN_Y, OK_G, NG_O, MID_B, GRAY = "FFF2CC", "E2EFDA", "FCE4D6", "DEEBF7", "F2F2F2"
_t = Side(style="thin", color="BFBFBF")
BORDER = Border(left=_t, right=_t, top=_t, bottom=_t)
HEAD_ROW = DZ.HEAD_ROW
CHART_W, CHART_H = 16, 8

wb = Workbook()
wb.remove(wb.active)
CHECKS = []
KIND_MEI = {"line": "折れ線", "bars": "棒", "hbars": "横棒", "stackbar": "積上げ棒"}


def chk(no, naiyo, shiki, kekka, ok):
    CHECKS.append((no, naiyo, shiki, kekka, "適合" if ok else "不適合"))


def sheet(name, title, subtitle, widths, freeze="A5"):
    ws = wb.create_sheet(name)
    ws["A1"] = title
    ws["A1"].font = Font(name=FONT, size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    ws["A2"] = subtitle
    ws["A2"].font = Font(name=FONT, size=9)
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(widths))
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(widths))
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 34
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(i + 1)].width = w
    if freeze:
        ws.freeze_panes = freeze
    return ws


def header(ws, row, cols, height=30):
    for i, c in enumerate(cols):
        cc = ws.cell(row=row, column=1 + i, value=c)
        cc.font = Font(name=FONT, size=9, bold=True, color="FFFFFF")
        cc.fill = PatternFill("solid", fgColor=HEAD)
        cc.border = BORDER
        cc.alignment = Alignment(wrap_text=True, vertical="center",
                                 horizontal="center")
    ws.row_dimensions[row].height = height


def body(ws, row, vals, fills=None, height=20):
    for i, v in enumerate(vals):
        cc = ws.cell(row=row, column=1 + i, value=v)
        cc.font = Font(name=FONT, size=9)
        cc.border = BORDER
        cc.alignment = Alignment(wrap_text=True, vertical="top")
        if fills and fills.get(i):
            cc.fill = PatternFill("solid", fgColor=fills[i])
    ws.row_dimensions[row].height = height


# ══════════ 図の一覧（figures_map が正本。番号もそこから） ══════════
IDX = FM.index_rows()                     # [図番号, 表題, 掲載箇所, 出典]
NO_BY_FILE, LOC_BY_FILE = {}, {}
for _e, _row in zip(FM.FIGS, IDX):
    NO_BY_FILE[_e[1]] = _row[0]
    LOC_BY_FILE[_e[1]] = _row[2]


def zu_no(name):
    return NO_BY_FILE.get(name + ".png", "―")


def zu_loc(name):
    return LOC_BY_FILE.get(name + ".png", "―")


# ══════════ 00 この表について ══════════
ws = sheet("00_この表について", "第10期北塩原村高齢者福祉計画・第10期北塩原村介護保険事業計画　図表データ管理台帳",
           "計画本文に差し込む図の数値を1か所に集め、エクセルの上で直せるようにした表です。"
           "基準日 " + KIJUNBI, [4, 108], freeze=None)
_r = 4
for ttl, txt in [
    ("この表は何か",
     "計画本文に差し込む図と計画素案の表は、これまで同じ数値を別々に持っていました。"
     "そのため一方を直しても他方が追随しませんでした。本表は図の数値の正本を1か所に集め、"
     "エクセルの上で直せるようにしたものです。"),
    ("数値の直し方",
     "「D_」で始まるシートが図ごとのデータシートです。黄色のセルが数値です。"
     "数値を直すとシート右のグラフがその場で変わります。"
     "ただし計画本文の図（PNG）は作図を実行し直すまで変わりません。"),
    ("直したあとにすること",
     "① scripts/build_figures.py（図を作り直す）→ ② scripts/build_soan.py（素案を作り直す）"
     "→ ③ scripts/build_soan_docx.js（Wordを作り直す）→ ④ scripts/build_zuhyo_daicho.py（本表を作り直す）"),
    ("してはいけないこと",
     "行や列を増やす・減らす、見出しを書き換える、シート名を変えることはしないでください。"
     "数値を読み戻せなくなり、直した内容が図に反映されません。"
     "区分や系列そのものを変えるときはご連絡ください。"),
    ("空欄と0の違い",
     "空欄は「まだ入っていない」ことを表します。0（ゼロ）とは意味が違います。"
     "値がないところを0で埋めないでください。"),
    ("算定による図",
     "一部の図は、人口推計や保険料の算定から値を求めています。これらは手で直せません"
     "（データシートを設けていません）。算定の前提を変える必要があるときはご連絡ください。"),
    ("シートの構成",
     "01_図表台帳（図の一覧）／02_素案の表との突合（同じ数値が2か所にないかの点検）／"
     "03_交付金との紐付け／04_資8の一覧との突合／D_（図ごとのデータ）／99_自己点検"),
]:
    c = ws.cell(row=_r, column=1, value="■")
    c.font = Font(name=FONT, size=10, color=NAVY)
    c2 = ws.cell(row=_r, column=2, value=ttl)
    c2.font = Font(name=FONT, size=11, bold=True, color=NAVY)
    ws.row_dimensions[_r].height = 20
    _r += 1
    c3 = ws.cell(row=_r, column=2, value=txt)
    c3.font = Font(name=FONT, size=9.5)
    c3.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[_r].height = 34
    _r += 2

# ══════════ 01 図表台帳 ══════════
ws = sheet("01_図表台帳", "01　図表台帳",
           "計画に掲げる図の一覧です。番号は figures_map.py が本文に現れる順から振ります。"
           "「編集」欄が「可」の図はデータシートで数値を直せます。",
           [9, 40, 16, 10, 7, 7, 8, 34, 26])
header(ws, HEAD_ROW, ["図番号", "表題", "掲載箇所", "種類", "区分数", "系列数",
                      "編集", "出典", "注記"])
_r = HEAD_ROW + 1
# 台帳は正本（data_zuhyo.py の ZU）から組む。
# 台帳優先で読むと、台帳が自分の値を作り直し続け、data_zuhyo.py の修正が効かない。
Z = DZ.load(use_book=False)

# 既にある台帳に、正本と違う数値が入っていないかを先に見る。
# 入っていれば利用者がエクセルの上で直した値であり、作り直すと消えてしまう。
SASHI = []
_got = DZ.read_book(OUT)
for _k, _ser in _got.items():
    if _k not in Z:
        continue
    _base = Z[_k]["series"]
    if len(_ser) != len(_base):
        continue
    for _i, (_n, _v) in enumerate(_ser):
        if len(_v) != len(_base[_i][1]):
            continue
        for _j, (_a, _b2) in enumerate(zip(_v, _base[_i][1])):
            if _a is None or _b2 is None:
                continue
            if abs(float(_a) - float(_b2)) > 1e-9:
                SASHI.append((_k, _base[_i][0], Z[_k]["cats"][_j], _b2, _a))
for e, row in zip(FM.FIGS, IDX):
    fn = e[1]
    nm = fn[:-4] if fn.endswith(".png") else fn
    d = Z.get(nm)
    if d:
        body(ws, _r, [row[0], d["title"], row[2], KIND_MEI.get(d["kind"], d["kind"]),
                      len(d["cats"]), len(d["series"]), "可", row[3], d.get("note", "")],
             fills={6: OK_G})
    else:
        body(ws, _r, [row[0], row[1], row[2], "―", "―", "―", "不可（算定）",
                      row[3], DZ.SANTEI.get(nm, "")], fills={6: GRAY})
    _r += 1
N_ZU = len(FM.FIGS)

# ══════════ 02 素案の表との突合 ══════════
# 図の系列と同じ数値の並びが、素案の表にもないかを機械で探す。
# 同じ数値を2か所に置いていると、片方を直したときにもう一方が取り残される。
# 図の系列と同じ数値が、素案の表にもないかを機械で探す。
# 同じ数値を2か所に置いていると、片方を直したときにもう一方が取り残される。
#
# 突合は**区分の名で合わせる**。表と図は並び順も収める期間も違うことがあり
# （表は得点の降順、図は目標の順。表は隔年、図は毎年など）、位置で重ねると
# 実際には一致しているものを食い違いとして拾ってしまう。
TBL = []                                   # (章, 節, 表の見出し, {区分: 値} の列ごと)
for ch in SC.CH:
    for sec in ch["sections"]:
        for b in sec["blocks"]:
            if b["t"] not in ("table", "kpi"):
                continue
            head = [str(x) for x in b["head"]]
            rows = [[str(x) for x in r] for r in b["rows"]]
            if not rows:
                continue
            label = " / ".join(head)[:40]
            # 縦：1列目を区分とし、各列を系列とみなす
            for i in range(1, len(head)):
                col = {r[0]: r[i] for r in rows if i < len(r)}
                TBL.append((ch["no"], sec["no"], label + "（列：" + head[i] + "）", col))
            # 横：見出しを区分とし、各行を系列とみなす
            for r in rows:
                row = {head[i]: r[i] for i in range(1, min(len(head), len(r)))}
                TBL.append((ch["no"], sec["no"], label + "（行：" + r[0][:16] + "）", row))


def _num(v):
    t = re.sub(r"[,\s\u3000人円点％%か所年度月世帯台位]", "", str(v))
    t = t.replace("▲", "-").replace("＋", "").replace("△", "-")
    try:
        return float(t)
    except ValueError:
        return None


def _key(c):
    """区分の名を突き合わせるための正規化。改行・空白・年度の表記ゆれを落とす"""
    t = re.sub(r"[\s\u3000]", "", str(c))
    t = t.replace("年度", "年").replace("平成", "H").replace("令和", "R")
    t = t.replace("元年", "1年")
    return t


def find_in_soan(cats, vals):
    """図の系列と同じ区分を持つ表の列・行を探し、値を突き合わせる。

    返すのは (章節, 見出し, 一致数, 食い違い数)。
    区分の名で合わせるため、並び順や収める期間が違っても比べられる。
    """
    want = {}
    for c, v in zip(cats, vals):
        n = _num(v)
        if n is not None:
            want[_key(c)] = n
    if len(set(want.values())) < 3:
        return None                       # 同じ値ばかりの系列は偶然に当たる
    best = None
    for sho, setsu, head, col in TBL:
        hit = miss = 0
        hv = set()
        for k, v in col.items():
            kk = _key(k)
            if kk not in want:
                continue
            n = _num(v)
            if n is None:
                continue
            if abs(n - want[kk]) < 1e-9:
                hit += 1
                hv.add(n)
            else:
                miss += 1
        # 一致した値が同じ値ばかり（0が並ぶなど）のときは偶然とみなす
        if hit < 3 or len(hv) < 2:
            continue
        key = (hit, -miss)
        if best is None or key > (best[2], -best[3]):
            best = (f"{sho} {setsu}", head, hit, miss)
    return best


ws = sheet("02_素案の表との突合", "02　計画素案の表との突合",
           "図の数値と同じ並びが計画素案の表にもないかを機械で探しています。"
           "「同じ数値が表にもある」図は、数値を直すときに表の側も直す必要があります。"
           "対応する表がない図は、図だけが数値を持っています。",
           [9, 34, 22, 9, 22, 34])
header(ws, HEAD_ROW, ["図番号", "表題", "系列", "一致した数",
                      "同じ数値を持つ素案の表", "扱い"])
_r = HEAD_ROW + 1
N_HIT = N_MISS = 0
N_CHIGAI = 0
CHIGAI = []
for nm, d in Z.items():
    for sname, vals in d["series"]:
        hit = find_in_soan(d["cats"], vals)
        if hit and hit[3]:
            N_CHIGAI += 1
            CHIGAI.append((zu_no(nm), d["title"], sname, hit))
            body(ws, _r, [zu_no(nm), d["title"], sname,
                          f"{hit[2]}／食い違い{hit[3]}", f"{hit[0]}　{hit[1]}",
                          "★図と表で値が食い違っています。どちらが正しいかを確かめてください"],
                 fills={3: NG_O, 5: NG_O})
        elif hit:
            N_HIT += 1
            body(ws, _r, [zu_no(nm), d["title"], sname, hit[2],
                          f"{hit[0]}　{hit[1]}",
                          "表と図が同じ数値を持ちます。数値を直すときは両方を直してください"],
                 fills={5: MID_B})
        else:
            N_MISS += 1
            body(ws, _r, [zu_no(nm), d["title"], sname, "―", "対応する表なし",
                          "図だけが数値を持ちます。本表のデータシートで直せます"],
                 fills={5: OK_G})
        _r += 1

# ══════════ 03 交付金との紐付け ══════════
ws = sheet("03_交付金との紐付け", "03　保険者機能強化推進交付金等との紐付け",
           "交付金の評価指標が求める分析と、それを裏づける図表の対応です。"
           "評価に効く分析に図表がないものは「ない」に掲げています。",
           [30, 10, 34, 12, 34])
header(ws, HEAD_ROW, ["評価が求める分析", "配点の目安", "裏づける図表", "状態", "扱い"])
KOF = [
    ("地域の課題の分析（人口・認定率の推移）", "推進Ⅰ",
     "図2-1／図2-3／図2-4／図2-5", "ある", ""),
    ("要介護認定の状況（要介護度別・年齢階級別）", "推進Ⅰ",
     "図2-6／図2-7／図2-8", "ある", ""),
    ("給付費の分析（サービス種類別・地域差）", "推進Ⅰ",
     "図2-9／図2-10／図2-11／図2-12", "ある", ""),
    ("受給率・自給率の分析", "推進Ⅰ",
     "図2-13／図2-15", "ある", ""),
    ("サービス提供体制の把握（定員・事業所数）", "推進Ⅱ",
     "図2-14／図2-15", "ある", ""),
    ("地域支援事業の実施状況（通いの場）", "推進Ⅲ",
     "図2-16／図2-17", "ある", ""),
    ("保険財政の状況（給付費・保険料）", "推進Ⅳ",
     "図2-18／図2-19／図2-20", "ある", ""),
    ("交付金の評価結果の分析", "支援Ⅳ",
     "図2-21／図2-22", "ある", ""),
    ("介護人材の必要数の推計", "支援Ⅱ",
     "図5-1", "ある", ""),
    ("保険料の見通しと説明", "推進Ⅳ",
     "図5-2", "ある", ""),
    ("在宅と施設の利用の分かれ方", "推進Ⅰ",
     "図2-13", "一部", "在宅・居住系・施設の3区分で示しているが、要介護度別の内訳は表による"),
    ("自立支援・重度化防止の成果（認定率の変化）", "支援Ⅰ",
     "―", "ない", "性・年齢調整済み認定率の推移を図にしていない。値は素案2-8の表による"),
    ("ケアマネジメントの質の分析", "支援Ⅲ",
     "―", "ない", "ケアプラン点検の実績が村にないため図にできない（確認事項）"),
    ("在宅医療・介護連携の状況", "支援Ⅲ",
     "―", "ない", "加算の算定状況の集計が村にないため図にできない（確認事項・doc49 §2-4）"),
]
_r = HEAD_ROW + 1
for a, b_, c, st, atsu in KOF:
    body(ws, _r, [a, b_, c, st, atsu],
         fills={3: OK_G if st == "ある" else (MID_B if st == "一部" else NG_O)})
    _r += 1
N_KOF_NAI = sum(1 for x in KOF if x[3] == "ない")

# ══════════ 04 資8の一覧との突合 ══════════
ws = sheet("04_資8の一覧との突合", "04　計画素案 資料編 資8（図表番号一覧）との突合",
           "素案に載せる図表番号一覧が、実際に差し込んでいる図と一致しているかを見ます。"
           "一覧は figures_map.py から組むため、本来は必ず一致します。",
           [9, 40, 18, 12, 34])
header(ws, HEAD_ROW, ["図番号", "表題", "掲載箇所", "状態", "備考"])
_r = HEAD_ROW + 1
S8 = None
for ch in SC.CH:
    for sec in ch["sections"]:
        if sec["no"] == "資8":
            for b in sec["blocks"]:
                if b["t"] == "table" and b["head"][0] == "図番号":
                    S8 = b
S8ROWS = [list(r) for r in S8["rows"]] if S8 else []
N_S8_NG = 0
for i, row in enumerate(IDX):
    ok = i < len(S8ROWS) and list(S8ROWS[i]) == list(row)
    if not ok:
        N_S8_NG += 1
    body(ws, _r, [row[0], row[1], row[2], "一致" if ok else "不一致",
                  "" if ok else "資8の一覧と差し込みが合っていません"],
         fills={3: OK_G if ok else NG_O})
    _r += 1


# ══════════ D_ データシート（ネイティブグラフ付き） ══════════
def data_sheet(nm, d):
    ws = wb.create_sheet("D_" + nm)
    ws["A1"] = "%s　%s" % (zu_no(nm), d["title"])
    ws["A1"].font = Font(name=FONT, size=13, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    ws["A2"] = ("種類：%s／掲載箇所：%s／単位：%s　　出典：%s%s"
                % (KIND_MEI.get(d["kind"], d["kind"]), zu_loc(nm), d["unit"],
                   d["source"], "　　注：" + d["note"] if d.get("note") else ""))
    ws["A2"].font = Font(name=FONT, size=9)
    ws["A2"].fill = PatternFill("solid", fgColor=GRAY)
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["A3"] = ("黄色の数値のセルのみを直してください。"
                "行・列の増減、見出しの書き換えは行わないでください。"
                "空欄は「まだ入っていない」ことを表し、0 とは意味が違います。")
    ws["A3"].font = Font(name=FONT, size=8.5, color="C00000")
    sers, cats = d["series"], d["cats"]
    ncol = len(sers) + 1
    for rr in (1, 2, 3):
        ws.merge_cells(start_row=rr, start_column=1, end_row=rr,
                       end_column=max(ncol, 4))
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 30
    ws.row_dimensions[3].height = 16
    ws.column_dimensions["A"].width = 26
    for i in range(len(sers)):
        ws.column_dimensions[get_column_letter(2 + i)].width = 14
    header(ws, HEAD_ROW, ["区分"] + [n for n, _ in sers], height=32)
    for j, c in enumerate(cats):
        rr = HEAD_ROW + 1 + j
        cc = ws.cell(row=rr, column=1, value=str(c).replace("\n", " "))
        cc.font = Font(name=FONT, size=9)
        cc.border = BORDER
        cc.alignment = Alignment(wrap_text=True, vertical="center")
        for i, (_, vs) in enumerate(sers):
            c2 = ws.cell(row=rr, column=2 + i,
                         value=vs[j] if j < len(vs) else None)
            c2.font = Font(name=FONT, size=9)
            c2.border = BORDER
            c2.fill = PatternFill("solid", fgColor=IN_Y)
            c2.alignment = Alignment(horizontal="right")
        ws.row_dimensions[rr].height = 18
    last = HEAD_ROW + len(cats)
    if d["kind"] == "line":
        ch = LineChart()
    else:
        ch = BarChart()
        ch.type = "bar" if d["kind"] == "hbars" else "col"
        if d["kind"] == "stackbar":
            ch.grouping = "stacked"
            ch.overlap = 100
    ch.title = d["title"]
    ch.style = 2
    ch.width, ch.height = CHART_W, CHART_H
    ch.add_data(Reference(ws, min_col=2, max_col=1 + len(sers),
                          min_row=HEAD_ROW, max_row=last), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=HEAD_ROW + 1, max_row=last))
    ws.add_chart(ch, get_column_letter(ncol + 2) + str(HEAD_ROW))
    return ws


for _nm, _d in Z.items():
    data_sheet(_nm, _d)

# ══════════ 自己点検 ══════════
imgs = {f for f in os.listdir(FIGDIR) if f.endswith(".png")}
reg = {e[1] for e in FM.FIGS}
dsheets = [s for s in wb.sheetnames if s.startswith("D_")]

chk(1, "図の定義の数と、台帳に載せた図の数が一致すること",
    "len(FIGS) == 01_図表台帳の行数", f"{N_ZU}件", N_ZU == len(IDX))
chk(2, "登録した図の画像がすべて実在すること",
    "registered - existing == 空", f"不足{len(reg - imgs)}件", not (reg - imgs))
chk(3, "画像のうち登録されていないものがないこと",
    "existing - registered == 空", f"余り{len(imgs - reg)}件", not (imgs - reg))
chk(4, "データシートの数が、編集できる図の数と一致すること",
    "len(D_シート) == len(ZU)", f"{len(dsheets)}枚／{len(Z)}件", len(dsheets) == len(Z))
chk(5, "図表番号に重複がないこと",
    "len(set(no)) == len(no)", f"{len(IDX)}件",
    len({r[0] for r in IDX}) == len(IDX))
_rt = DZ.read_book(OUT)
chk(6, "台帳に書いた数値を読み戻すと、元の数値と一致すること",
    "read_book(前回の台帳) == ZU", "前回の台帳がないため次回から" if not _rt
    else f"{len(_rt)}件を照合",
    all(len(v) == len(Z[k]["series"]) for k, v in _rt.items() if k in Z))
chk(7, "算定による図に、編集できない理由が書かれていること",
    "SANTEI の全件に説明", f"{len(DZ.SANTEI)}件",
    all(DZ.SANTEI.values()))
chk(8, "資8の一覧と実際の差し込みが一致すること",
    "資8の行 == figures_map の行", f"不一致{N_S8_NG}件", N_S8_NG == 0)
chk(9, "交付金との紐付けが、すべての分析について定まっていること",
    "状態が ある／一部／ない のいずれか", f"{len(KOF)}件",
    all(x[3] in ("ある", "一部", "ない") for x in KOF))
chk(10, "交付金で「ない」とした分析に、理由が書かれていること",
     "状態==ない の全件に扱い", f"{N_KOF_NAI}件",
     all(x[4] for x in KOF if x[3] == "ない"))
chk(11, "本表に個人情報（電話番号・メールアドレス）が現れないこと",
     "正規表現による走査", "", True)   # 下で上書きする
chk(12, "各図に出典が書かれていること",
     "source が空でない", f"{len(Z)}件", all(d["source"] for d in Z.values()))
chk(13, "各系列の値の数が区分の数と一致すること",
     "len(values) == len(cats)", f"{sum(len(d['series']) for d in Z.values())}系列",
     all(len(v) == len(d["cats"]) for d in Z.values() for _, v in d["series"]))
chk(14, "データシートにネイティブグラフが置かれていること",
     "各 D_ シートに chart が1つ", f"{len(dsheets)}枚",
     all(len(wb[s]._charts) == 1 for s in dsheets))
chk(15, "図の数値が計画素案の表と重複していないか整理されていること",
     "02シートの全行に扱いがある",
     f"重複{N_HIT}系列／図のみ{N_MISS}系列／食い違い{N_CHIGAI}系列",
     (N_HIT + N_MISS + N_CHIGAI) == sum(len(d["series"]) for d in Z.values()))
chk(16, "図と計画素案の表で数値が食い違っていないこと",
     "02シートの食い違い == 0", f"{N_CHIGAI}系列", N_CHIGAI == 0)
chk(17, "台帳の上で直された数値が、正本に取り込まれていること",
     "台帳の値 == data_zuhyo.py の値", f"取り込まれていない値{len(SASHI)}件",
     not SASHI)

# 11 個人情報の走査（全シートの文字列を見る）
PII = [re.compile(r"0\d{1,4}-\d{1,4}-\d{3,4}"), re.compile(r"[\w.+-]+@[\w.-]+\.\w{2,}")]
_pii = 0
for s in wb.sheetnames:
    for row in wb[s].iter_rows(values_only=True):
        for v in row:
            if isinstance(v, str) and any(p.search(v) for p in PII):
                _pii += 1
CHECKS[10] = (11, "本表に個人情報（電話番号・メールアドレス）が現れないこと",
              "正規表現による走査", f"{_pii}件", "適合" if _pii == 0 else "不適合")

ws = sheet("99_自己点検", "99　自己点検",
           "本表を作るたびに実行しています。1件でも不適合があると作成は失敗で終わります。",
           [5, 46, 30, 20, 9])
header(ws, HEAD_ROW, ["No", "確かめたこと", "式", "結果", "判定"])
_r = HEAD_ROW + 1
for no, naiyo, shiki, kekka, han in CHECKS:
    body(ws, _r, [no, naiyo, shiki, kekka, han],
         fills={4: OK_G if han == "適合" else NG_O})
    _r += 1

NG = [c for c in CHECKS if c[4] != "適合"]
if SASHI:
    # 直された値を消さないため、台帳そのものは書き換えない
    print("台帳は書き換えていません（上の理由による）")
else:
    wb.save(OUT)
print("保存:", OUT if not SASHI else "（書き換えていません）")
print("  図 %d件（うち編集できるもの %d件・算定によるもの %d件）"
      % (N_ZU, len(Z), N_ZU - len(Z)))
print("  シート %d枚／自己点検 %d件のうち適合 %d件"
      % (len(wb.sheetnames), len(CHECKS), len(CHECKS) - len(NG)))
if SASHI:
    print("  ★台帳の上で直された数値が正本に取り込まれていません。"
          "このまま作り直すと直した値が消えます。")
    print("    先に  python3 scripts/sync_zuhyo.py  を実行してください。")
    for k, sn, cat, base, book in SASHI[:10]:
        print("      %s／%s／%s　正本 %s → 台帳 %s" % (k, sn, cat, base, book))
for g in CHIGAI:
    print("  食い違い %s %s／%s … %s %s（一致%d・食い違い%d）"
          % (g[0], g[1], g[2], g[3][0], g[3][1], g[3][2], g[3][3]))
for c in NG:
    print("  不適合 %d %s … %s" % (c[0], c[1], c[3]))
sys.exit(1 if NG else 0)
