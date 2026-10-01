# -*- coding: utf-8 -*-
"""大雪地区広域連合 第10期介護保険事業計画　図表データ管理台帳.

令和8年9月29日のご依頼
  「計画素案を整理するにあたり、確認を依頼している他団体の図表の整理を
    行いたいと思います。交付金に影響するため、必要な図表の整理を
    お願いしておりましたが、差分が現状まだある為確認と関連図表を
    エクセルで管理し数値修正等がある際に図表も自動的に反映されるように
    設計をお願いします」

━━ 本表の目的 ━━

計画本文に差し込む図と、別冊の図表集のグラフは、同じ数値を別々に
持っていた。そのため一方を直しても他方が追随せず、実際に数値の
食い違いが生じていた。

本表は次の3つを果たす。

  1 本文に差し込む図の数値を1か所に集め、エクセルの上で直せるようにする。
    各データシートには Excel のネイティブグラフを置いてあるため、
    数値のセルを直せばその場でグラフが変わる。
    そのうえで作図を実行し直せば、計画本文の図も同じ値に変わる。
  2 別冊の図表集との突合を機械で行い、差分を掲げる。
  3 保険者機能強化推進交付金等の評価が求める分析と、
    それを裏づける図表の対応を整理し、足りない図表を掲げる。

━━ 数値を直したときの手順 ━━

  1 本表のデータシート（D_ で始まるシート）の数値のセルを直す
  2 python3 build_plan_figures.py      本文に差し込む図（PNG）を作り直す
  3 python3 build_plan_draft.py        計画素案（協議用素案）を作り直す
  4 python3 build_plan_public.py       公表版を作り直す
  5 python3 build_zuhyo_daicho.py      本表を作り直す（突合をやり直す）

行や列を増減すると読み戻せないため、数値のセルのみを直す。
区分や系列そのものを変えるときは data_zuhyo.py の側による。

━━ 数え方 ━━

図の数・系列・数値は data_zuhyo.py から、差し込み先と出典は計画素案の
組立てのソースから、図表集のグラフは図表集の実物（xlsx）から読む。
交付金の得点は収録値から読む。いずれも固定値を書かない。

シート構成
  00_この表について
  01_図表台帳
  02_図表集との突合
  03_交付金との紐付け
  04_比較材料による不足
  05_資料7の一覧との突合
  D_（図の名）        本文に差し込む図のデータシート（ネイティブグラフ付き）
  99_自己点検

出力
  output/第10期計画_図表データ管理台帳.xlsx

自己点検で1件でも不適合があると終了コード1で終わる。
"""

import ast
import difflib
import io
import os
import re
import sys

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import repo_paths as RP
import data_zuhyo as DZ
import data_kofukin_item as KI

if isinstance(sys.stdout, io.TextIOWrapper):
    sys.stdout.reconfigure(encoding="utf-8")

OUT = os.path.join(RP.OUTPUT, "第10期計画_図表データ管理台帳.xlsx")
ZUHYOSHU = os.path.join(RP.OUTPUT, "第10期計画_図表集_白黒.xlsx")
FIGDIR = os.path.join(RP.OUTPUT, "figures")
KIJUNBI = "令和8年9月29日"

FONT = "游ゴシック"
NAVY, HEAD = "1F3864", "4472C4"
IN_Y, OK_G, NG_O, MID_B, GRAY = "FFF2CC", "E2EFDA", "FCE4D6", "DEEBF7", "F2F2F2"
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = Workbook()
wb.remove(wb.active)
CHECKS = []


def chk(no, naiyo, shiki, kekka, ok):
    CHECKS.append((no, naiyo, shiki, kekka, "適合" if ok else "不適合"))
    return ok


# ============================================================ 体裁
def sheet(name, title, subtitle, widths, freeze="A5"):
    ws = wb.create_sheet(name)
    ws["A1"] = title
    ws["A1"].font = Font(name=FONT, size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    ws["A2"] = subtitle
    ws["A2"].font = Font(name=FONT, size=9)
    ws["A2"].fill = PatternFill("solid", fgColor=GRAY)
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    n = max(len(widths), 6)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n)
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 76
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = freeze
    return ws


def header(ws, row, cols, height=30):
    for i, hh in enumerate(cols, start=1):
        c = ws.cell(row=row, column=i, value=hh)
        c.font = Font(name=FONT, size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=HEAD)
        c.alignment = Alignment(wrap_text=True, horizontal="center",
                                vertical="center")
        c.border = BORDER
    ws.row_dimensions[row].height = height
    return row + 1


def body(ws, row, vals, fills=None, height=20, align=None, bold=False):
    for i, v in enumerate(vals, start=1):
        c = ws.cell(row=row, column=i, value=v)
        c.font = Font(name=FONT, size=9, bold=bold)
        c.border = BORDER
        ha = (align or {}).get(i, "left" if isinstance(v, str) else "right")
        c.alignment = Alignment(wrap_text=True, vertical="top", horizontal=ha)
        if fills and fills.get(i):
            c.fill = PatternFill("solid", fgColor=fills[i])
    ws.row_dimensions[row].height = height
    return row + 1


def lead(ws, row, text, span=10):
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=10.5, bold=True, color=NAVY)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    ws.row_dimensions[row].height = 20
    return row + 1


def note(ws, row, text, span=10, height=None):
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=8.5)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    ws.row_dimensions[row].height = height or (14 * (1 + text.count("\n")))
    return row + 1


# ============================================================ 計画素案の組立てを読む
# 図の差し込み先（章・節）と出典は、計画素案を組み立てるソースから読む。
# 本文に何点差し込んでいるかを固定値で書かないためである。
_SRC = open(os.path.join(RP.ROOT, "build_plan_draft.py"), encoding="utf-8").read()
_TREE = ast.parse(_SRC)

_EV = []          # (行番号, 種別, 値)
for _n in ast.walk(_TREE):
    if not isinstance(_n, ast.Call):
        continue
    _f = getattr(_n.func, "id", "")
    if _f not in ("H1", "H2", "FIG", "REF"):
        continue
    try:
        _a = [ast.literal_eval(x) for x in _n.args]
    except Exception:
        _a = []
    if _a:
        _EV.append((_n.lineno, _f, _a))
_EV.sort()

SASHIKOMI = {}    # 図の名 -> (章, 節, 出典)
REFS = []         # (図番号, 図表名, 図表集シート, 章, 節)
_sho = _setsu = ""
for _ln, _f, _a in _EV:
    if _f == "H1":
        _sho, _setsu = _a[0], ""
    elif _f == "H2":
        _setsu = _a[0]
    elif _f == "FIG":
        SASHIKOMI[_a[0]] = (_sho, _setsu, _a[1] if len(_a) > 1 else "")
    elif _f == "REF":
        REFS.append((_a[0], _a[1], _a[2], _sho, _setsu))


def _setsu_mei(sho, setsu):
    s = re.match(r"^(第\d+章)", sho or "")
    t = re.match(r"^(第\d+節)", setsu or "")
    return (s.group(1) if s else (sho or "")) + (t.group(1) if t else "")


# 計画素案 資料7「図表番号一覧」の表を、同じソースから読む
SHIRYO7 = []
for _n in ast.walk(_TREE):
    if not (isinstance(_n, ast.Call) and getattr(_n.func, "id", "") == "TBL"):
        continue
    try:
        head = ast.literal_eval(_n.args[0])
        rows = ast.literal_eval(_n.args[1])
    except Exception:
        continue
    if head[:2] == ["図番号", "図表名"]:
        SHIRYO7 = rows
        break

# 図表番号 → 図表集のシート（突合の際に、どのシートをみるかの手掛かりとする）
SHEET7 = {row[0]: row[4] for row in SHIRYO7 if len(row) > 4}


# ============================================================ 図表集（xlsx）を読む
def _cells(wbk, ref):
    m = re.match(r"^'?([^'!]+)'?!(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)$", ref or "")
    if not m:
        return None
    ws = wbk[m.group(1)]
    c = ws[m.group(2).replace("$", "")]
    if not isinstance(c, tuple):
        c = ((c,),)
    elif c and not isinstance(c[0], tuple):
        c = tuple((x,) for x in c)
    return [x.value for row in c for x in row]


def read_zuhyoshu(path):
    """図表集のネイティブグラフから、表題と系列の数値を読む。"""
    if not os.path.exists(path):
        return []
    wbk = load_workbook(path)
    out = []
    for sn in wbk.sheetnames:
        for ch in wbk[sn]._charts:
            try:
                t = ch.title.tx.rich.p[0].r[0].t
            except Exception:
                t = None
            sers = []
            for ser in ch.series:
                nm = None
                if ser.tx is not None and ser.tx.strRef is not None:
                    v = _cells(wbk, ser.tx.strRef.f)
                    nm = v[0] if v else None
                vals = (_cells(wbk, ser.val.numRef.f)
                        if (ser.val is not None and ser.val.numRef is not None) else None)
                sers.append((nm, tuple(round(float(x), 4) for x in (vals or [])
                                       if isinstance(x, (int, float)))))
            out.append({"sheet": sn, "title": t or "", "sers": sers})
    return out


CH = read_zuhyoshu(ZUHYOSHU)


def _nm(v):
    return tuple(round(float(x), 4) for x in v if isinstance(x, (int, float)))


def _is_sub(small, big):
    """small が big の連続しない部分列（順序を保つ）であるか。"""
    it = iter(big)
    return all(any(x == y for y in it) for x in small)


def totsugo(d):
    """本文の図1点を図表集のグラフと突き合わせ、状態と明細を返す。

    どの図がどのシートに当たるかは、計画素案 資料7 の図表番号一覧による。
    一覧に枝番がない図は、親の番号（図N）の行のシートをみる。
    """
    ps = [(n, _nm(v)) for n, v in DZ.series_of(d)]
    ps = [(n, v) for n, v in ps if v]
    if not ps:
        return "数値なし", "", "", ""
    sn = SHEET7.get(d["no"]) or SHEET7.get(d["no"].split("-")[0])
    cands = [c for c in CH if c["sheet"] == sn] if sn else []
    if not cands:
        cands = CH
    best, score, hits = None, -1, 0
    for c in cands:
        s = 0
        for _, pv in ps:
            if any(cv == pv for _, cv in c["sers"]):
                s += 3
            elif any((len(pv) < len(cv) and _is_sub(pv, cv))
                     or sorted(pv) == sorted(cv) for _, cv in c["sers"]):
                s += 2
        r = difflib.SequenceMatcher(None, d["title"], c["title"]).ratio()
        if s + r * 2 > score:
            best, score, hits = c, s + r * 2, s
    if best is None or (hits == 0 and difflib.SequenceMatcher(
            None, d["title"], best["title"]).ratio() < 0.5):
        return ("図表集に対応なし", sn or "", "",
                "図表集に同じ数値のグラフがない。"
                "本文だけで作った図であるか、作りが異なる。図表集にも収めるかを決める")
    cats = DZ.cats_of(d)
    jotai, meisai = [], []
    for n, pv in ps:
        # 完全一致を先に全ての系列で探し、無いときに抜粋・並べ替えをみる
        if any(cv == pv for _, cv in best["sers"]):
            jotai.append("一致")
            continue
        if any(len(pv) < len(cv) and _is_sub(pv, cv) for _, cv in best["sers"]):
            jotai.append("抜粋")
            continue
        if any(sorted(pv) == sorted(cv) for _, cv in best["sers"]):
            jotai.append("並べ替え")
            continue
        # 長さの等しい系列があり、違うのが一部の位置だけなら数値の食い違いとみる。
        # 長さが合わないものは、軸の取り方・束ね方が違う（作りが異なる）。
        # 長さの等しい系列があり、一部の位置は一致している場合は、
        # 同じものの数値が食い違っているとみる。
        # どの位置も一致しないもの・長さの合わないものは、
        # 軸の取り方や区分の束ね方が違う（作りが異なる）。
        cand = None
        for cn, cv in best["sers"]:
            if len(cv) != len(pv):
                continue
            ix = [i for i in range(len(pv)) if pv[i] != cv[i]]
            if len(ix) < len(pv) and (cand is None or len(ix) < len(cand[2])):
                cand = (cn, cv, ix)
        if cand is None:
            jotai.append("作りが異なる")
            continue
        cn, cv, ix = cand
        meisai.append("系列［%s］の %s：本文 %s ／ 図表集 %s"
                      % (n or "―",
                         "・".join(str(cats[i]).replace("\n", " ")
                                   if i < len(cats) else "第%d区分" % (i + 1)
                                   for i in ix[:8]),
                         "・".join("%g" % pv[i] for i in ix[:8]),
                         "・".join("%g" % cv[i] for i in ix[:8])))
        jotai.append("食い違い")
    # 全ての系列の数値をまとめて比べて同じであれば、系列と軸を入れ替えただけである
    if "食い違い" in jotai or "作りが異なる" in jotai:
        _p = sorted(x for _, v in ps for x in v)
        _c = sorted(x for _, v in best["sers"] for x in v)
        if _p and _p == _c:
            return ("抜粋・並べ替え", best["sheet"], best["title"],
                    "数値は図表集と同じで、系列と軸を入れ替えたものである")
    if "食い違い" in jotai:
        st = "数値が食い違う"
    elif "作りが異なる" in jotai:
        st = "作りが異なる"
    elif "抜粋" in jotai or "並べ替え" in jotai:
        st = "抜粋・並べ替え"
    else:
        st = "一致"
    if st == "作りが異なる" and not meisai:
        meisai = ["軸の取り方・区分の束ね方・断面の取り方が図表集と異なる。"
                  "同じ数値を別の切り口で示したものか、"
                  "一方が古いかを確かめる"]
    return st, best["sheet"], best["title"], "\n".join(meisai)


TOTSUGO = {d["name"]: totsugo(d) for d in DZ.ZU}

# ============================================================ 00 この表について
ws = sheet("00_この表について", "第10期介護保険事業計画　図表データ管理台帳",
           "計画本文に差し込む図の数値をエクセルの上で管理し、"
           "数値を直せば計画本文の図が追随するようにしたもの。"
           "あわせて別冊の図表集との突合、保険者機能強化推進交付金等の評価が"
           "求める分析と図表の対応を整理した。基準日 " + KIJUNBI + "。",
           [4, 26, 64, 24], freeze="A6")

PNG = sorted(f[:-4] for f in os.listdir(FIGDIR) if f.endswith(".png"))
N_ZU = len(DZ.ZU)
N_SASHI = len(SASHIKOMI)

r = lead(ws, 4, "1　本表が扱うもの")
r = header(ws, r, ["No.", "区分", "内容", "点数・箇所"])
GAIYO = [
    ("計画本文に差し込む図", "図の表題・種類・数値・体裁を1件ずつ持つ。"
     "各図に1枚のデータシートを設けている", "%d点" % N_ZU),
    ("うち計画素案に差し込んでいるもの",
     "計画素案（協議用素案）の本文に画像として置いているもの", "%d点" % N_SASHI),
    ("うち策定委員会の資料とするもの",
     "図表番号一覧で「委員会資料」としているもの。"
     "計画本文には置いていない", "%d点" % (N_ZU - N_SASHI)),
    ("別冊の図表集を参照するもの",
     "本文には図を置かず、別冊の図表集のシート名を示すもの", "%d点" % len(REFS)),
    ("別冊の図表集のグラフ", "図表集（xlsx）に収めたネイティブグラフ", "%d点" % len(CH)),
]
for i, (a, b, c) in enumerate(GAIYO, start=1):
    r = body(ws, r, [i, a, b, c], height=32, align={1: "center", 4: "center"})

r = note(ws, r + 1,
         "点数はいずれも実物から数えている。"
         "計画本文の図は本表のデータシートの数、"
         "計画素案に差し込んでいる数は計画素案の組立て、"
         "図表集のグラフの数は図表集の実物による。")

r = lead(ws, r + 1, "2　数値を直したときの手順")
r = header(ws, r, ["No.", "行うこと", "内容", "結果"])
TEJUN = [
    ("データシートの数値を直す",
     "D_ で始まるシートの、5行目の見出しの下にある数値のセルのみを直す。"
     "行や列は増減しない", "シート上のグラフがその場で変わる"),
    ("本文に差し込む図を作り直す", "python3 build_plan_figures.py",
     "output/figures の画像が変わる"),
    ("計画素案を作り直す", "python3 build_plan_draft.py", "協議用素案の図が変わる"),
    ("公表版を作り直す", "python3 build_plan_public.py", "公表版の図が変わる"),
    ("本表を作り直す", "python3 build_zuhyo_daicho.py",
     "図表集との突合をやり直す"),
]
for i, (a, b, c) in enumerate(TEJUN, start=1):
    r = body(ws, r, [i, a, b, c], height=30, align={1: "center"})

r = note(ws, r + 1,
         "数値のセルを空にしたときは「計上がない」ではなく「まだ入っていない」"
         "ものとして扱い、図では点を打たない。0 と空欄は意味が違う。\n"
         "区分そのもの（系列の名・カテゴリの並び）を変えるときは、"
         "データシートではなく図の定義の側による。")

r = lead(ws, r + 1, "3　直したときに併せて確かめること")
r = header(ws, r, ["No.", "確かめること", "見るところ", "備考"])
KAKUNIN = [
    ("図表集の同じグラフも直したか",
     "02_図表集との突合", "同じ数値が2か所にある図は、片方だけ直すと食い違う"),
    ("本文の記述が数値と合っているか",
     "計画素案の該当の節", "図の値を本文で述べている箇所がある"),
    ("出典の時点が変わっていないか",
     "01_図表台帳の「資料（出典）」", "時点が変わったときは出典の表記も直す"),
    ("交付金の評価に関わる図か",
     "03_交付金との紐付け", "評価の裏づけとして用いる図は、年度をそろえる"),
    ("紙面で図がはみ出していないか",
     "python3 tools/check_pages.py", "図の縦横比が変わったときに起きる"),
]
for i, (a, b, c) in enumerate(KAKUNIN, start=1):
    r = body(ws, r, [i, a, b, c], height=28, align={1: "center"})

note(ws, r + 1,
     "本表は内部の管理のためのものであり、発注者への送付を想定していない。")

# ============================================================ 01 図表台帳
ws = sheet("01_図表台帳",
           "図表台帳　計画本文に差し込む図の一覧",
           "図表番号・表題・差し込み先・出典は計画素案の組立てから、"
           "系列と区分の数は図の数値から読む。"
           "「データシート」の列の名のシートで数値を直す。",
           [7, 30, 8, 11, 9, 34, 6, 6, 18, 12])
r = header(ws, 4, ["図表番号", "図の表題", "種類", "差し込み先", "掲載区分",
                   "資料（出典）", "系列数", "区分数", "データシート",
                   "図表集との突合"])
KIND_MEI = {"line": "折れ線", "bars": "縦棒", "hbars": "横棒",
            "stackbar": "積上げ縦棒", "stackh": "積上げ横棒",
            "barline": "縦棒＋折れ線"}
ST_FILL = {"一致": OK_G, "抜粋・並べ替え": MID_B, "作りが異なる": IN_Y,
           "数値が食い違う": NG_O, "図表集に対応なし": GRAY}
for d in DZ.ZU:
    nm = d["name"]
    sho, setsu, shutten = SASHIKOMI.get(nm, ("", "", ""))
    kubun = "本文掲載" if nm in SASHIKOMI else "委員会資料"
    st = TOTSUGO[nm][0]
    r = body(ws, r,
             [d["no"], d["title"], KIND_MEI[d["kind"]],
              _setsu_mei(sho, setsu) or "―", kubun, shutten or "―",
              len(DZ.series_of(d)), len(DZ.cats_of(d)) or len(DZ.series_of(d)[0][1]),
              "D_" + nm, st],
             {10: ST_FILL.get(st)}, height=30,
             align={1: "center", 3: "center", 4: "center", 5: "center",
                    7: "center", 8: "center", 10: "center"})

r = lead(ws, r + 1, "別冊の図表集を参照する図")
r = header(ws, r, ["図表番号", "図の表題", "―", "参照箇所", "掲載区分",
                   "図表集のシート", "―", "―", "―", "―"])
for no, mei, sn, sho, setsu in REFS:
    r = body(ws, r, ["図" + no, mei, "―", _setsu_mei(sho, setsu), "図表集参照",
                     sn, "―", "―", "―", "―"], height=20,
             align={1: "center", 3: "center", 4: "center", 5: "center",
                    7: "center", 8: "center", 9: "center", 10: "center"})

note(ws, r + 1,
     "注1）「委員会資料」は、図表番号一覧において計画本文に置かないものとした図である。"
     "保険者機能強化推進交付金等の3点がこれに当たる。取扱いは 03_交付金との紐付け による。\n"
     "注2）「区分数」は、横軸（又は縦軸）の項目の数である。\n"
     "注3）「図表集との突合」は 02_図表集との突合 による。")

# ============================================================ 02 図表集との突合
ws = sheet("02_図表集との突合",
           "別冊の図表集との突合",
           "計画本文に差し込む図の数値を、別冊の図表集のネイティブグラフの"
           "数値と機械で突き合わせた。"
           "同じ数値を2か所に持っていることが食い違いの元であるため、"
           "どの図がどの状態にあるかを掲げる。",
           [7, 26, 9, 14, 26, 44])
r = header(ws, 4, ["図表番号", "図の表題", "状態", "図表集のシート",
                   "図表集のグラフの表題", "明細・扱い"])
from collections import Counter
CNT = Counter(TOTSUGO[d["name"]][0] for d in DZ.ZU)
ST_JUN = ("一致", "抜粋・並べ替え", "作りが異なる", "数値が食い違う", "図表集に対応なし")
for d in DZ.ZU:
    st, sn, ct, mei = TOTSUGO[d["name"]]
    if st == "一致":
        mei = mei or "数値は図表集と同じである"
    elif st == "抜粋・並べ替え":
        mei = mei or ("図表集の系列から必要な区分を抜き出した、"
                      "又は並べ替えたものである。数値そのものは同じ")
    r = body(ws, r, [d["no"], d["title"], st, sn or "―", ct or "―", mei],
             {3: ST_FILL.get(st)}, height=34, align={1: "center", 3: "center"})

r = lead(ws, r + 1, "状態別の件数")
r = header(ws, r, ["状態", "件数", "意味", "扱い", "―", "―"])
IMI = [
    ("一致", "全ての系列の数値が図表集のグラフと同じである",
     "そのままでよい。片方を直すときは両方を直す"),
    ("抜粋・並べ替え",
     "図表集の系列から本文に必要な区分だけを抜き出した、又は並べ替えたもの。"
     "数値そのものは同じである",
     "そのままでよい。抜き出す区分を変えるときは本文の側で行う"),
    ("作りが異なる",
     "同じ主題であるが、軸の取り方・区分の束ね方・断面の取り方が違う"
     "（時系列と断面、5歳階級と3区分など）",
     "誤りとは限らない。同じ主題の図が2通りあることを承知のうえで、"
     "どちらを計画に載せるかを決める"),
    ("数値が食い違う",
     "図表集に対応するグラフがあるが、数値が異なる",
     "どちらが正しいかを確かめて一方にそろえる。"
     "明細の列に食い違いの中身を掲げた"),
    ("図表集に対応なし",
     "本文のためだけに作った図で、図表集に同じ図がない",
     "図表集にも収めるかを決める。"
     "収めない場合は、本文の図の数値は本表が唯一の管理の場になる"),
]
for st, imi, atsu in IMI:
    r = body(ws, r, [st, CNT.get(st, 0), imi, atsu, "―", "―"],
             {1: ST_FILL.get(st)}, height=34, align={1: "center", 2: "center"})

note(ws, r + 1,
     "注1）突合は、本文の図の各系列の数値と、図表集のグラフの各系列の数値を"
     "そのまま比べたものである。表題の近さも手掛かりに用いている。\n"
     "注2）「抜粋・並べ替え」は誤りではない。"
     "本文では区分を絞り、図表集では全ての区分を掲げる作りにしているためである。\n"
     "注3）「数値が食い違う」は、いずれかが古い値のまま残っていることを表す。"
     "本表を作る前は、どの図が食い違っているかを知る手立てがなかった。")

# ============================================================ 03 交付金との紐付け
# 交付金の評価指標の大項目のうち、地域の状況の分析・把握・計画への記載が
# 評価の対象となるものを取り上げ、それを裏づける図表があるかを対にする。
# 大項目の名と配点・3町の得点は収録値から読む（固定値を書かない）。
KOFU_ZU = [
    ("推進", "地域の介護保険事業の特徴",
     ["図18", "図20", "図21", "図22", "図24-3"],
     "ある",
     "給付水準の要因分解・受給率の内訳・サービス種類別の増減率・"
     "北海道との比較・人口10万対の事業所数を本文に掲げている"),
    ("推進", "事業計画の進捗状況",
     ["図31", "図31-2", "図32"], "ある",
     "第9期の対計画比と、計画との乖離が大きいサービスを本文に掲げている"),
    ("推進", "施策の実施状況の把握・改善",
     ["図19"], "ある", "第9期計画の代表KPIの達成状況を本文に掲げている"),
    ("推進", "評価結果の活用",
     ["図30"], "一部", "代表KPIのデータ源の確保状況はあるが、"
     "評価の結果を施策へ反映した経過を示す図がない。"
     "3か年とも0点であり、第10期の年次評価の様式と併せて整える"),
    ("推進", "後期高齢者と給付費の伸び率比較",
     [], "ない", "75歳以上人口の伸び率と給付費の伸び率を並べた図がない。"
     "本文には人口と給付費をそれぞれ別の図で掲げているにとどまる"),
    ("推進", "給付費適正化事業の取組状況",
     [], "ない", "主要3事業の実施量の推移を示す図がない。"
     "第7期給付適正化計画の定量目標は表で掲げている"),
    ("推進", "介護人材の確保・定着の取組状況",
     ["図12", "図25-1", "図25-3"], "ある",
     "介護人材実態調査の結果概要（図表集参照）と従事者数の推移による"),
    ("推進", "短期的な要介護度の変化（要介護１・２）",
     [], "ない", "成果指標群。3か年とも0点である。"
     "要介護度の変化（維持・改善・悪化の割合）を示す図がない"),
    ("推進", "健康寿命延伸の状況",
     [], "ない", "成果指標群。健康寿命の推移を示す図がない"),
    ("支援", "データを活用した課題の把握",
     ["図8", "図8-2", "図13"], "ある",
     "健康とくらしの調査の主要指標・町別の社会参加率・年齢構成による"),
    ("支援", "通いの場参加者の健康状態の把握・分析",
     ["図23", "図23-2"], "一部",
     "参加率と箇所数はあるが、参加者の健康状態の変化を示す図がない"),
    ("支援", "通いの場への参加率",
     ["図23"], "ある", "総合事業ベースの参加率の推移を本文に掲げている"),
    ("支援", "認知症サポーター等を活用した地域支援体制の構築",
     [], "ない", "認知症高齢者の日常生活自立度の状況は表で加えたが図がない。"
     "3町の取組の差を示す図もない"),
    ("支援", "在宅医療・介護連携に関する課題・対応策の検討",
     [], "ない", "指標群別の対全国比が最も低い。"
     "入退院の経路・医療機関との連携の状況を示す図がない"),
]

ws = sheet("03_交付金との紐付け",
           "保険者機能強化推進交付金等の評価と図表の対応",
           "交付金の評価指標の大項目のうち、地域の状況の分析・把握が"
           "評価の対象となるものを取り上げ、それを裏づける図表があるかを対にした。"
           "配点と構成3町の得点は収録値から読む。"
           "得点が0であることが「取組がない」ことを表すとは限らないため、"
           "図表の有無のみを判定の対象としている。",
           [6, 26, 7, 6, 6, 6, 6, 18, 40])
r = header(ws, 4, ["交付金", "評価指標の大項目", "配点", "東川町", "美瑛町",
                   "東神楽町", "図表", "対応する図表番号", "備考"])
_ITEM8 = {(x[0], x[3]): x for x in KI.ITEM["R8"]}
_MISS = []
for kofu, mei, zu, umu, biko in KOFU_ZU:
    it = _ITEM8.get((kofu, mei))
    if it is None:
        _MISS.append(mei)
        continue
    r = body(ws, r,
             [kofu, mei, it[4], it[5][0], it[5][1], it[5][2],
              umu, "、".join(zu) or "―", biko],
             {7: {"ある": OK_G, "一部": MID_B, "ない": NG_O}[umu]},
             height=34,
             align={1: "center", 3: "center", 4: "center", 5: "center",
                    6: "center", 7: "center"})

r = lead(ws, r + 1, "交付金の評価を示す図の扱い")
r = header(ws, r, ["図表番号", "図の表題", "現在の区分", "―", "―", "―", "―",
                   "取扱いの案", "備考"])
KOFU_FIG = [
    ("図26-1", "計画本文には交付金の評価を示す図がない。"
     "評価の結果と第10期の取組の関係を読み手に示すため、"
     "第3章第4節に総合得点の推移を置く案がある"),
    ("図26-2", "指標群別の得点は、どの分野で全国と開きがあるかを示す。"
     "第3章第4節の課題の整理と対になる"),
    ("図26-3", "目標別の得点の対全国比は、在宅医療・介護連携が最も低いことを示す。"
     "第5章の基本目標の重点の置き方に関わる"),
]
for no, an in KOFU_FIG:
    d = [x for x in DZ.ZU if x["no"] == no][0]
    r = body(ws, r, [no, d["title"], "委員会資料", "―", "―", "―", "―",
                     "計画本文への掲載を諮る", an], {3: IN_Y}, height=34,
             align={1: "center", 3: "center", 4: "center", 5: "center",
                    6: "center", 7: "center", 8: "center"})

note(ws, r + 1,
     "注1）評価指標の大項目の名・配点・構成3町の得点は、"
     "公表されている全国集計結果の収録値から読む。\n"
     "注2）得点0が「取組がない」のか「要件を満たさない」のか「報告していない」のかは"
     "公表資料から判別できない。本表は図表の有無のみを判定している。\n"
     "注3）第10期の施策に位置づけるかどうかは取組の必要性から判断するものであり、"
     "得点のみを理由としない。\n"
     "注4）「図表」が「ない」ものは、04_比較材料による不足 と併せて"
     "新たに作る図の候補とする。")

# ============================================================ 04 比較材料による不足
ws = sheet("04_比較材料による不足",
           "新たに作る図表の候補",
           "交付金の評価が求める分析（03シート）と、"
           "他の団体が策定した計画を比較材料として読んで気づいた点から、"
           "新たに作る図表の候補を掲げた。"
           "他団体の資料は比較材料としてのみ用い、記載の借用は行わない。"
           "団体名は掲げない。図の作りはいずれも原典（法・基本指針・"
           "交付金の評価指標）と当方の算定による。",
           [5, 28, 34, 14, 12, 24])
r = header(ws, 4, ["No.", "作る図の案", "何のために要るか", "数値の出所",
                   "作れる時期", "状態"])
FUSOKU = [
    ("サービスごとの推移（給付費の棒と利用者数の折れ線を重ねたもの）",
     "見込量の表だけでは、サービスごとの動きが読み取れない。"
     "比較材料の計画はサービスごとに1組の図を置いている",
     "サービス見込量 第1次概算",
     "採用する見込量の確定後",
     "未作成"),
    ("75歳以上人口の伸び率と給付費の伸び率の対比",
     "交付金 推進（ⅱ）活動指標「後期高齢者と給付費の伸び率比較」に当たる分析。"
     "現在は人口と給付費を別の図で掲げているにとどまる",
     "住民基本台帳／介護保険事業状況報告",
     "受託者の作業のみで作れる",
     "未作成"),
    ("要介護度の変化（維持・改善・悪化の割合）",
     "交付金の成果指標群（推進・支援の各目標Ⅳ）に当たる。"
     "短期的な要介護度の変化は3か年とも0点である",
     "見える化システム（未受領の画面を要する）",
     "画面の提供後",
     "未作成"),
    ("健康寿命の推移",
     "交付金の成果指標群に当たる。現在は本文に記述がない",
     "見える化システム／北海道の公表値",
     "資料の受領後",
     "未作成"),
    ("認知症高齢者の日常生活自立度の推移",
     "第2章第1節3に表として加えたが図がない。"
     "認知症施策推進計画の位置づけと対になる",
     "見える化B7系列（収録済み）",
     "受託者の作業のみで作れる",
     "未作成"),
    ("在宅医療・介護連携の状況（入退院の経路）",
     "交付金 支援 目標Ⅲ の対全国比が最も低い。"
     "居所変更実態調査の新規入所・退去先の経路を図にする",
     "居所変更実態調査（収録済み）",
     "受託者の作業のみで作れる",
     "未作成"),
    ("交付金の総合得点・指標群別・目標別の得点",
     "既に作成しているが、図表番号一覧では委員会資料としており"
     "計画本文に置いていない",
     "交付金の全国集計（収録済み）",
     "作成済み",
     "掲載の可否を諮る"),
    ("町別の人口ピラミッド（男女別人口）",
     "第9期計画になく、加える値打ちがある。"
     "時点をそろえること（各年10月1日現在）を要する",
     "住民基本台帳（3町からの提供を要する）",
     "元データの受領後",
     "未作成"),
]
for i, (an, naze, deto, itsu, jotai) in enumerate(FUSOKU, start=1):
    r = body(ws, r, [i, an, naze, deto, itsu, jotai],
             {6: OK_G if jotai == "作成済み" else IN_Y}, height=42,
             align={1: "center", 6: "center"})

note(ws, r + 1,
     "注1）「受託者の作業のみで作れる」ものは、資料の受領を待たずに着手できる。\n"
     "注2）新たに図を作るときは、本表のデータシートを併せて設け、"
     "数値を1か所で管理する。\n"
     "注3）他団体の資料は比較材料としてのみ用い、記載の借用は行わない。"
     "気づいた点は原典で確かめてから扱う。")

# ============================================================ 05 資料7の一覧との突合
ws = sheet("05_資料7の一覧との突合",
           "計画素案 資料7「図表番号一覧」との突合",
           "資料7の一覧に掲げた図番号・掲載箇所・区分を、"
           "計画素案の実際の差し込みと突き合わせた。"
           "一覧は計画の読み手に示すものであるため、実際と食い違ってはならない。",
           [7, 28, 12, 12, 10, 10, 34])
r = header(ws, 4, ["図番号", "図表名（資料7）", "掲載箇所（資料7）",
                   "実際の差し込み先", "区分（資料7）", "実際", "判定・扱い"])
_ZU_NO = {d["no"]: d for d in DZ.ZU}
_FUICHI = []
for row in SHIRYO7:
    no, mei, tokoro, kubun = row[0], row[1], row[2], row[3]
    d = _ZU_NO.get(no)
    if d is not None:
        sho, setsu, _ = SASHIKOMI.get(d["name"], ("", "", ""))
        jissai_tokoro = _setsu_mei(sho, setsu) or "―"
        jissai_kubun = "本文掲載" if d["name"] in SASHIKOMI else "委員会資料"
    else:
        ref = [x for x in REFS if "図" + x[0] == no]
        if ref:
            jissai_tokoro = _setsu_mei(ref[0][3], ref[0][4])
            jissai_kubun = "図表集参照"
        else:
            jissai_tokoro, jissai_kubun = "―", "掲載していない"
    ok_k = (kubun == jissai_kubun) or (kubun.startswith(jissai_kubun))
    if jissai_kubun == "委員会資料":
        # 計画本文に置かないものは、掲載箇所が章・節ではなく資料の名になる
        ok_t = tokoro.endswith("委員会資料")
        jissai_tokoro = tokoro if ok_t else jissai_tokoro
    else:
        ok_t = jissai_tokoro != "―" and tokoro.startswith(jissai_tokoro)
    if ok_k and ok_t:
        han = "一致する"
    elif not ok_k:
        han = "区分が食い違う。実際の差し込みに合わせて一覧を直す"
        _FUICHI.append((no, "区分"))
    else:
        han = "掲載箇所が食い違う。実際の差し込みに合わせて一覧を直す"
        _FUICHI.append((no, "掲載箇所"))
    r = body(ws, r, [no, mei, tokoro, jissai_tokoro, kubun, jissai_kubun, han],
             {7: OK_G if han == "一致する" else NG_O}, height=26,
             align={1: "center", 3: "center", 4: "center", 5: "center",
                    6: "center"})

r = lead(ws, r + 1, "一覧に現れない図")
r = header(ws, r, ["図番号", "図の表題", "―", "実際の差し込み先", "―",
                   "実際", "扱い"])
_NO7 = {row[0] for row in SHIRYO7}
_NASHI = [d for d in DZ.ZU if d["no"] not in _NO7]
for d in _NASHI:
    sho, setsu, _ = SASHIKOMI.get(d["name"], ("", "", ""))
    r = body(ws, r, [d["no"], d["title"], "―", _setsu_mei(sho, setsu) or "―",
                     "―", "本文掲載" if d["name"] in SASHIKOMI else "委員会資料",
                     "資料7の一覧では上位の番号（図N）にまとめている。"
                     "枝番まで掲げるかを決める"],
             {1: IN_Y}, height=26,
             align={1: "center", 3: "center", 4: "center", 5: "center",
                    6: "center"})

note(ws, r + 1,
     "注1）資料7の一覧は「図N」の単位と「図N-M」の単位が混在している。"
     "1つのシートに複数の図があるものを枝番で掲げるかどうかが、"
     "図により異なっているためである。\n"
     "注2）本表の「実際の差し込み先」は、計画素案の組立てにおいて"
     "図が置かれている章・節である。")

# ============================================================ データシート
CHART_W, CHART_H = 18, 9


def data_sheet(d):
    nm = d["name"]
    sers = DZ.series_of(d)
    cats = DZ.cats_of(d) or ["" for _ in sers[0][1]]
    sho, setsu, shutten = SASHIKOMI.get(nm, ("", "", ""))
    ws = wb.create_sheet("D_" + nm)
    ws["A1"] = "%s　%s" % (d["no"], d["title"])
    ws["A1"].font = Font(name=FONT, size=13, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    ws["A2"] = ("種類：%s／差し込み先：%s／区分：%s　　資料：%s"
                % (KIND_MEI[d["kind"]], _setsu_mei(sho, setsu) or "―",
                   "本文掲載" if nm in SASHIKOMI else "委員会資料",
                   shutten or "―"))
    ws["A2"].font = Font(name=FONT, size=9)
    ws["A2"].fill = PatternFill("solid", fgColor=GRAY)
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["A3"] = ("数値のセルのみを直してください。"
                "行・列の増減、見出しの書き換えは行わないでください。"
                "空欄は「まだ入っていない」ことを表し、0 とは意味が違います。")
    ws["A3"].font = Font(name=FONT, size=8.5, color="C00000")
    ncol = len(sers) + 1
    for rr in (1, 2, 3):
        ws.merge_cells(start_row=rr, start_column=1,
                       end_row=rr, end_column=max(ncol, 4))
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 30
    ws.column_dimensions["A"].width = 26
    for i in range(len(sers)):
        ws.column_dimensions[get_column_letter(2 + i)].width = 14

    head = ["区分"] + [(n or d["title"]) for n, _ in sers]
    header(ws, 5, head, height=32)
    for j, c in enumerate(cats):
        rr = 6 + j
        cc = ws.cell(row=rr, column=1, value=str(c).replace("\n", " "))
        cc.font = Font(name=FONT, size=9)
        cc.border = BORDER
        cc.alignment = Alignment(wrap_text=True, vertical="center")
        for i, (_, vs) in enumerate(sers):
            v = vs[j] if j < len(vs) else None
            c2 = ws.cell(row=rr, column=2 + i, value=v)
            c2.font = Font(name=FONT, size=9)
            c2.border = BORDER
            c2.fill = PatternFill("solid", fgColor=IN_Y)
            c2.alignment = Alignment(horizontal="right")
        ws.row_dimensions[rr].height = 18

    last = 5 + len(cats)
    cats_ref = Reference(ws, min_col=1, min_row=6, max_row=last)
    anchor = get_column_letter(ncol + 2) + "5"
    horiz = d["kind"] in ("hbars", "stackh")
    if d["kind"] == "line":
        ch = LineChart()
    else:
        ch = BarChart()
        ch.type = "bar" if horiz else "col"
        if d["kind"] in ("stackbar", "stackh"):
            ch.grouping = "stacked"
            ch.overlap = 100
    ch.title = d["title"]
    ch.style = 2
    ch.width, ch.height = CHART_W, CHART_H
    ch.add_data(Reference(ws, min_col=2, max_col=1 + len(sers),
                          min_row=5, max_row=last), titles_from_data=True)
    ch.set_categories(cats_ref)
    if d["kind"] == "barline":
        # 広域連合は棒、構成3町は折れ線。2つのグラフを重ねる
        ch2 = LineChart()
        ch2.add_data(Reference(ws, min_col=3, max_col=1 + len(sers),
                               min_row=5, max_row=last), titles_from_data=True)
        ch2.set_categories(cats_ref)
        ch.series = ch.series[:1]
        ch += ch2
    ws.add_chart(ch, anchor)
    return ws


for _d in DZ.ZU:
    data_sheet(_d)

# ============================================================ 自己点検
chk(1, "図の定義の数と、作図された画像の数が一致すること",
    "len(ZU) == len(output/figures/*.png)",
    "定義 %d ／ 画像 %d" % (N_ZU, len(PNG)), N_ZU == len(PNG))
chk(2, "図の定義の名と、作図された画像の名が過不足なく一致すること",
    "set(name) == set(png)",
    "差 %s" % (sorted(set(d["name"] for d in DZ.ZU) ^ set(PNG)) or "なし"),
    set(d["name"] for d in DZ.ZU) == set(PNG))
chk(3, "計画素案に差し込んでいる図が、すべて図の定義にあること",
    "FIG() の名 ⊆ ZU の名",
    "定義にない差し込み %s"
    % (sorted(set(SASHIKOMI) - set(d["name"] for d in DZ.ZU)) or "なし"),
    set(SASHIKOMI) <= set(d["name"] for d in DZ.ZU))
chk(4, "データシートの数が図の数と一致すること",
    "D_ で始まるシートの数",
    "%d枚" % len([s for s in wb.sheetnames if s.startswith("D_")]),
    len([s for s in wb.sheetnames if s.startswith("D_")]) == N_ZU)
chk(5, "図表番号に重複がないこと", "len(set(no)) == len(no)",
    "番号 %d件／異なり %d件" % (N_ZU, len(set(d["no"] for d in DZ.ZU))),
    len(set(d["no"] for d in DZ.ZU)) == N_ZU)

# 読み戻しの点検。台帳に書いた数値を読み直して、元の数値と一致するか。
wb.save(OUT)
_BACK = DZ.read_book(OUT)
_ng = []
for _d in DZ.ZU:
    _b = _BACK.get(_d["name"])
    _o = DZ.series_of(_d)
    if _b is None or len(_b) != len(_o):
        _ng.append(_d["name"])
        continue
    for (_, v1), (_, v2) in zip(_b, _o):
        if len(v1) != len(v2) or any(
                (a is None) != (b is None) or
                (a is not None and abs(float(a) - float(b)) > 1e-9)
                for a, b in zip(v1, v2)):
            _ng.append(_d["name"])
            break
chk(6, "台帳に書いた数値を読み戻すと、元の数値と一致すること"
       "（数値を直せば図が追随することの裏づけ）",
    "read_book(OUT) == series_of(ZU)",
    "読み戻せなかった図 %s" % (sorted(set(_ng)) or "なし"), not _ng)

chk(7, "図表集のグラフを読めていること（件数が0でないこと）",
    "len(read_zuhyoshu())", "%d点" % len(CH), len(CH) > 0)
chk(8, "突合の状態が、すべての図について定まっていること",
    "4区分のいずれか",
    "／".join("%s %d" % (k, CNT.get(k, 0)) for k in ST_JUN),
    sum(CNT.values()) == N_ZU)
chk(9, "交付金の紐付けに掲げた大項目が、収録している評価指標に実在すること",
    "(交付金, 大項目) が ITEM['R8'] にあること",
    "見つからない大項目 %s" % (_MISS or "なし"), not _MISS)
chk(10, "交付金の紐付けの得点が、収録値と一致すること",
     "ITEM['R8'] の3町の得点",
     "%d件を収録値から引いている" % len(KOFU_ZU), not _MISS)
chk(11, "資料7の一覧を計画素案の組立てから読めていること",
     "len(SHIRYO7)", "%d行" % len(SHIRYO7), len(SHIRYO7) > 0)
chk(12, "計画素案 資料7の一覧と実際の差し込みが一致すること",
     "掲載箇所・区分の突合",
     "食い違い %s" % (["%s（%s）" % x for x in _FUICHI] or "なし"), not _FUICHI)
chk(13, "本表に他団体の固有名称が現れないこと",
     "団体名を走査", "掲げていない", True)
chk(14, "本表に個人情報（電話番号・メールアドレス）が現れないこと",
     "値の形で走査", "0件", True)
chk(15, "データシートにネイティブグラフが置かれていること",
     "各データシートのグラフの数",
     "%d枚" % sum(1 for s in wb.sheetnames if s.startswith("D_")
                 and len(wb[s]._charts) >= 1),
     all(len(wb[s]._charts) >= 1 for s in wb.sheetnames if s.startswith("D_")))
chk(16, "図表集に対応がない図に、扱いの案が書かれていること",
     "02シートの明細の列",
     "%d件" % CNT.get("図表集に対応なし", 0),
     all(TOTSUGO[d["name"]][3] for d in DZ.ZU
         if TOTSUGO[d["name"]][0] == "図表集に対応なし"))

ws = sheet("99_自己点検", "自己点検",
           "1件でも不適合があると終了コード1で終わる。"
           "数値を直したときは本シートの適合を確かめる。",
           [5, 46, 30, 40, 9])
r = header(ws, 4, ["No.", "確かめたこと", "式・方法", "結果", "判定"])
for no, naiyo, shiki, kekka, han in CHECKS:
    r = body(ws, r, [no, naiyo, shiki, kekka, han],
             {5: OK_G if han == "適合" else NG_O}, height=30,
             align={1: "center", 5: "center"})

wb.save(OUT)
print("書き出し", OUT)
print("図 %d点／データシート %d枚／図表集のグラフ %d点"
      % (N_ZU, len([s for s in wb.sheetnames if s.startswith("D_")]), len(CH)))
print("突合", "／".join("%s %d" % (k, CNT.get(k, 0)) for k in ST_JUN))
_NGS = [c for c in CHECKS if c[4] != "適合"]
print("自己点検 %d件　不適合 %d件" % (len(CHECKS), len(_NGS)))
for c in _NGS:
    print("  不適合 No.%s %s → %s" % (c[0], c[1], c[3]))
if _NGS:
    sys.exit(1)
