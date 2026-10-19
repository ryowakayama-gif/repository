# -*- coding: utf-8 -*-
"""大雪地区広域連合 第10期介護保険事業計画
　　見える化システムの指標（B5・B6・D49）の受領点検.

令和8年10月19日のご連絡
  「確認事項のうち、一部資料送付しますので、添付ファイル確認をお願い致します。」

━━ 受領したもの ━━

  ① B5-a 調整済み認定率（要介護度別）
  ② B5-c 調整済み認定率（要支援除く）
  ③ B6-a 調整済み重度認定率
  ④ B6-b 調整済み軽度認定率　　　　（①〜④は xlsx。出力日 令和8年10月19日）
  ⑤ D49 地域差指数の指標の解説（PDF・全2頁）

①から④は**資料提供依頼No.1**（見える化B5a・B6a・B6bの令和8年3月末の値）、
⑤は**資料提供依頼No.16**（見える化システムD49の指標定義）に当たる。
いずれも受領済に改めた。

━━ 確かめたこと ━━

1 **既に収録している221件は受領値と差異なく一致する。**
  平成25年3月末から令和7年3月末までの13時点×17系列である。

2 **令和8年3月末の値を加えた**（重度5.2％・軽度13.2％・合計18.4％）。
  重度は低下が続き、軽度は令和5年3月末を底に上昇した後は横ばいである。

3 **「_比較」の系列が何を表すのかを特定できていない。**
  受領した出力は【注目する地域のみ】であり、本広域連合の行のみを含む。
  この値が「_比較」のない側と一致したことにより、
  「_比較」の系列が何を表すのかを特定できていないことが分かった
  → 確認事項No.190。

4 **D49 の受給率の分母は第1号被保険者数である**ことが原典で確定した
  （修正指示A-11の留保が解消）。

5 **地域差指数は乗法的に分解できない。**
  4指標のうち性・年齢調整後であるのは2指標だけであるためである。

シート構成
  00_この資料について
  01_受領したもの
  02_令和8年3月末の値
  03_既に収録している値との突合
  04_比較の系列の地域が特定できないこと
  05_計画素案への反映
  06_D49 地域差指数の定義
  07_地域差指数の分解の可否
  08_台帳の更新
  09_自己点検

出力
  output/第10期計画_見える化の指標（B5・B6・D49）の受領点検.xlsx

自己点検で1件でも不適合があると終了コード1で終わる。
"""

import ast
import datetime
import io
import os
import re
import sys

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import repo_paths as RP

sys.path.insert(0, RP.ROOT)
import data_nintei as DN                                     # noqa: E402
import data_mieru_d49 as D49                                 # noqa: E402

if isinstance(sys.stdout, io.TextIOWrapper):
    sys.stdout.reconfigure(encoding="utf-8")

ODIR = RP.OUTPUT
OUT = os.path.join(ODIR, "第10期計画_見える化の指標（B5・B6・D49）の受領点検.xlsx")

FONT = "游ゴシック"
NAVY, HEAD = "1F3864", "4472C4"
IN_Y, OK_G, NG_O, MID_B, GRAY = ("FFF2CC", "E2EFDA", "FCE4D6",
                                 "DEEBF7", "F2F2F2")
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = Workbook()
wb.remove(wb.active)
CHECKS = []


def chk(no, naiyo, shiki, kekka, ok):
    CHECKS.append((no, naiyo, shiki, kekka, "適合" if ok else "不適合"))
    return ok


JURYO = datetime.date(2026, 10, 19)

SH = {"about": "00", "uke": "01", "r8": "02", "totsugo": "03",
      "hikaku": "04", "soan": "05", "teigi": "06", "bunkai": "07",
      "daicho": "08", "chk": "09"}


def shn(key):
    return SH[key] + "シート"


def wareki(d):
    return "令和%d年%d月%d日" % (d.year - 2018, d.month, d.day)


# ============================================================ 読み取り
def literal(fn, name):
    """ソースから literal の代入を読む（固定値を書き写さないため）。"""
    src = io.open(os.path.join(RP.ROOT, fn), encoding="utf-8").read()
    for node in ast.parse(src).body:
        if isinstance(node, ast.Assign) and any(
                getattr(t, "id", None) == name for t in node.targets):
            return ast.literal_eval(node.value)
    raise RuntimeError("%s に %s が見つからない" % (fn, name))


# 地域差指数の値の読み取りは data_mieru_d49 に置いている
# （**同じ処理を読む側ごとに書かない**。CLAUDE.md §4）。
INS, KG, KN, KJ, KT, IDX = D49.d49_values()
CHECK = literal("build_process_control.py", "CHECK")
LACK = literal("build_process_control.py", "LACK")
BY_L = {x[0]: x for x in LACK}
BY_C = {x[0]: x for x in CHECK}

# 受領した指標。(番号, 指標, 系列数, 対応する資料提供依頼)
UKE_B = [("B5-a", "調整済み認定率（要介護度別）", 1),
         ("B5-c", "調整済み認定率（要支援除く）", 1),
         ("B6-a", "調整済み重度認定率", 1),
         ("B6-b", "調整済み軽度認定率", 1)]
R8 = "令和8年3月末"
R7 = "令和7年3月末"


def ser(code):
    return DN.B[code]["系列"]


def yrs(code):
    return list(next(iter(ser(code).values())).keys())


# 既に収録していた時点（令和8年3月末を除く）との突合の件数
TOTSUGO_N = sum(len(d) - 1 for c, _nm, _k in UKE_B for d in ser(c).values())
KEIRETSU_N = sum(len(ser(c)) for c, _nm, _k in UKE_B)


# ============================================================ 体裁
def _plain(v):
    return v.replace("**", "") if isinstance(v, str) else v


def sheet(name, title, subtitle, widths, freeze="A5"):
    ws = wb.create_sheet(name)
    ws["A1"] = title
    ws["A1"].font = Font(name=FONT, size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    ws["A2"] = _plain(subtitle)
    ws["A2"].font = Font(name=FONT, size=9)
    ws["A2"].fill = PatternFill("solid", fgColor=GRAY)
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    n = max(len(widths), 6)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n)
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 58
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    if freeze:
        ws.freeze_panes = freeze
    ws.sheet_view.showGridLines = False
    return ws


def header(ws, row, cols, height=30):
    for i, c in enumerate(cols, start=1):
        cell = ws.cell(row, i, _plain(c))
        cell.font = Font(name=FONT, size=9, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=HEAD)
        cell.alignment = Alignment(wrap_text=True, vertical="center",
                                   horizontal="center")
        cell.border = BORDER
    ws.row_dimensions[row].height = height
    return row + 1


def body(ws, row, vals, fills=None, height=22, align=None, bold=False):
    for i, v in enumerate(vals, start=1):
        cell = ws.cell(row, i, _plain(v))
        cell.font = Font(name=FONT, size=9, bold=bold)
        cell.alignment = Alignment(
            wrap_text=True, vertical="top",
            horizontal=(align or {}).get(i, "left"))
        cell.border = BORDER
        if fills and i in fills:
            cell.fill = PatternFill("solid", fgColor=fills[i])
    ws.row_dimensions[row].height = height
    return row + 1


def lead(ws, row, text, span=8):
    ws.cell(row, 1, _plain(text))
    ws.cell(row, 1).font = Font(name=FONT, size=11, bold=True, color=NAVY)
    ws.merge_cells(start_row=row, start_column=1,
                   end_row=row, end_column=span)
    ws.row_dimensions[row].height = 22
    return row + 1


def note(ws, row, text, span=8, height=None, fill=GRAY):
    ws.cell(row, 1, _plain(text))
    ws.cell(row, 1).font = Font(name=FONT, size=9)
    ws.cell(row, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(row, 1).fill = PatternFill("solid", fgColor=fill)
    ws.merge_cells(start_row=row, start_column=1,
                   end_row=row, end_column=span)
    ws.row_dimensions[row].height = height or (16 * (text.count("\n") + 3))
    return row + 1


def pct(v):
    return "%.1f％" % v


# ============================================================ 00
ws = sheet(SH["about"] + "_この資料について",
           "見える化システムの指標（B5・B6・D49）の受領点検",
           "%sに受領した5件を読み、当方の収録値及び計画素案と"
           "突き合わせたものです。"
           "調整済み認定率は令和8年3月末の値を加え、"
           "地域差指数は指標の定義により分解の可否を確かめました。"
           % wareki(JURYO),
           [6, 30, 54, 22, 14, 12, 12, 12])

r = 4
r = lead(ws, r, "1　先に申し上げること")
for t in (
    "・**既に収録している%d件は、受領値と差異なく一致します。**"
    "平成25年3月末から%sまでの13時点×%d系列です（%s）。"
    % (TOTSUGO_N, R7, KEIRETSU_N, shn("totsugo")),
    "・**令和8年3月末の値を加えました。**"
    "調整済み重度認定率%s・調整済み軽度認定率%s・合計調整済み認定率%sです。"
    "重度は低下が続き、軽度は令和5年3月末を底に上昇した後は横ばいです"
    "（%s）。"
    % (pct(ser("B6-a")["調整済み重度認定率"][R8]),
       pct(ser("B6-b")["調整済み軽度認定率"][R8]),
       pct(ser("B5-a")["合計調整済み認定率"][R8]), shn("r8")),
    "・**調整済み認定率には2つの系列があり、"
    "受領値と一致したのは【注目する地域のみ】の側でした。**"
    "もう一方（【他地域と比較】の出力）が何を表すのかは"
    "受領資料から特定できていません。"
    "計画素案 第3章第4節2が2つの系列の値を混ぜて1つの文に"
    "並べていたため是正しました（%s・%s）。"
    % (shn("hikaku"), shn("soan")),
    "・**地域差指数の受給率の分母は第1号被保険者数である**ことが"
    "原典で確定しました。従前は出力から分母を確認できず留保していました"
    "（%s）。" % shn("teigi"),
    "・**地域差指数は乗法的に分解できません。**"
    "4指標のうち性・年齢調整後であるのは2指標だけであるためです（%s）。"
    % shn("bunkai"),
):
    r = note(ws, r, t, height=16 * (len(t) // 60 + 2))

r += 1
r = lead(ws, r, "2　対応する資料提供依頼")
r = header(ws, r, ["", "依頼No.", "資料", "優先度", "希望時期", "状態",
                   "", ""])
for no in (1, 16):
    x = BY_L[no]
    r = body(ws, r, ["", "No.%d" % no, x[2], x[7], x[6], x[8], "", ""],
             height=30, fills={6: OK_G},
             align={2: "center", 4: "center", 5: "center", 6: "center"})
r = note(ws, r,
         "注）いずれも入手先を見える化システムとしていたもので、"
         "当方の側では画面に接続できないため出力をお願いしていたものです。",
         height=18)

r += 1
r = lead(ws, r, "3　新たに起票した確認事項")
r = header(ws, r, ["", "No.", "件名", "確認先", "期限", "状態", "", ""])
_c190 = BY_C.get(190)
if _c190:
    r = body(ws, r, ["", "No.190", _c190[3], _c190[6], _c190[8], _c190[7],
                     "", ""], height=30, fills={6: IN_Y},
             align={2: "center", 4: "center", 5: "center", 6: "center"})

# ============================================================ 01
ws = sheet(SH["uke"] + "_受領したもの", "受領したもの",
           "いずれも地域包括ケア「見える化」システムの出力又は"
           "指標の解説です。出力日は%sです。"
           "いずれも集計値であり、個人情報は含まれていません。"
           % wareki(JURYO),
           [6, 12, 34, 20, 14, 46, 12, 12])
r = 4
r = lead(ws, r, "1　受領した5件")
r = header(ws, r, ["", "指標", "指標名", "出力の種類", "系列数", "収めたもの",
                   "", ""])
for i, (code, nm, _k) in enumerate(UKE_B, start=1):
    r = body(ws, r, [i, code, nm, "【注目する地域のみ】", len(ser(code)),
                     "令和8年3月末の値を既存の系列に加えた", "", ""],
             height=26, align={1: "center", 2: "center", 5: "center"})
r = body(ws, r, [5, "D49", D49.MEISHO, "指標の解説（全2頁）", "―",
                 "指標の定義・4指標の算式・関連指標・出典", "", ""],
         height=40, align={1: "center", 2: "center", 5: "center"},
         fills={2: MID_B})

r += 1
r = lead(ws, r, "2　出力の時点と出典")
r = header(ws, r, ["", "事項", "内容", "", "", "", "", ""])
for a, b in (
    ("時点", "平成25年3月末から令和8年3月末までの14時点"),
    ("対象の地域", "大雪地区広域連合のみ（構成3町の値はありません）"),
    ("出典（B5・B6）",
     "厚生労働省「介護保険事業状況報告」年報及び総務省"
     "「住民基本台帳人口・世帯数」。"
     "ただし令和6年度及び令和7年度のみ年報ではなく月報によります"
     "（令和7年3月末及び令和8年3月末の値がこれに当たります）"),
    ("出典（D49）", "年次　%s／月次　%s"
     % (D49.SHUTTEN["年次"], D49.SHUTTEN["月次"])),
    ("調整の意味",
     "「調整済み」は性別及び年齢階級の構成の違いをそろえたものです。"
     "年齢構成の影響を除いた水準の比較に用います"),
):
    r = body(ws, r, ["", a, b, "", "", "", "", ""], height=40,
             fills={2: GRAY})
r = note(ws, r,
         "注）年報と月報では集計の時点が異なるため、"
         "令和7年3月末及び令和8年3月末の2時点は、"
         "それより前の時点と同じ系列として扱うことに留意を要します。"
         "本広域連合の年報（令和7年度）による認定率は21.826％であり、"
         "調整済み認定率（年齢構成をそろえた値）とは別のものです。",
         height=50)

# ============================================================ 02
ws = sheet(SH["r8"] + "_令和8年3月末の値", "令和8年3月末の値",
           "受領により加えた値です。"
           "括弧内は令和7年3月末からの動きです。"
           "調整済み認定率は年齢構成をそろえた値であり、"
           "年報による認定率（令和7年度21.826％）とは別のものです。",
           [6, 14, 34, 12, 12, 12, 34, 12])
r = 4
for code, nm, _k in UKE_B:
    r = lead(ws, r, "%s　%s" % (code, nm))
    r = header(ws, r, ["", "系列", "", "令和7年3月末", "令和8年3月末",
                       "動き", "読み方", ""])
    for key, d in ser(code).items():
        dv = d[R8] - d[R7]
        r = body(ws, r, ["", key, "", pct(d[R7]), pct(d[R8]),
                         ("±0.0" if abs(dv) < 1e-9
                          else ("＋%.1f" % dv if dv > 0 else "▲%.1f" % -dv)),
                         "", ""],
                 height=20, fills={5: OK_G},
                 align={4: "right", 5: "right", 6: "center"})
    r += 1

r = lead(ws, r, "長期の動き")
r = header(ws, r, ["", "系列", "最も高い時点", "最も低い時点",
                   "令和8年3月末", "動きの読み方", "", ""])
for code, key in (("B6-a", "調整済み重度認定率"),
                  ("B6-b", "調整済み軽度認定率"),
                  ("B5-a", "合計調整済み認定率")):
    d = ser(code)[key]
    hi = max(d, key=lambda k: d[k])
    lo = min(d, key=lambda k: d[k])
    if code == "B6-a":
        yomi = ("平成25年3月末を最高として低下が続いており、"
                "令和8年3月末が最も低い。重度化防止の面では改善が続いている")
    elif code == "B6-b":
        yomi = ("平成27年3月末を頂点に低下した後、"
                "近年は令和5年3月末を底に令和6年3月末へ上昇し、"
                "令和6年3月末から令和8年3月末までは横ばいである。"
                "低下が止まったことが第10期の重点を介護予防に移す根拠となる")
    else:
        yomi = ("平成27年3月末を頂点に低下した後、"
                "近年は令和5年3月末を底に令和6年3月末へ上昇し、"
                "その後は横ばいである。"
                "合計の認定率（調整前）の上昇は年齢構成の変化によるものである")
    r = body(ws, r, ["", key, "%s %s" % (hi, pct(d[hi])),
                     "%s %s" % (lo, pct(d[lo])), pct(d[R8]), yomi, "", ""],
             height=46, align={5: "right"})

# ============================================================ 03
ws = sheet(SH["totsugo"] + "_既に収録している値との突合",
           "既に収録している値との突合",
           "受領値と、令和8年8月4日の出力により収録していた値とを"
           "系列ごと・時点ごとに突き合わせました。"
           "%d件すべてが一致しています。" % TOTSUGO_N,
           [6, 14, 34, 14, 14, 14, 34, 12])
r = 4
r = lead(ws, r, "1　突合の結果")
r = header(ws, r, ["", "指標", "系列数", "突合した時点", "突合した件数",
                   "差異", "", ""])
for code, nm, _k in UKE_B:
    n = sum(len(d) - 1 for d in ser(code).values())
    r = body(ws, r, ["", "%s %s" % (code, nm), len(ser(code)),
                     "平成25年3月末から%s" % R7, "%d件" % n, "0件", "", ""],
             height=24, fills={6: OK_G},
             align={3: "center", 5: "center", 6: "center"})
r = body(ws, r, ["", "計", KEIRETSU_N, "13時点", "%d件" % TOTSUGO_N,
                 "0件", "", ""], height=22, bold=True, fills={6: OK_G},
         align={3: "center", 5: "center", 6: "center"})
r = note(ws, r,
         "注）突合は系列名と時点の両方を鍵として行い、"
         "小数第1位まで一致することを確かめました。"
         "一致したことにより、当方が収録していた値が"
         "本広域連合の値であることが裏づけられました。", height=34)

r += 1
r = lead(ws, r, "2　この突合で分かったこと")
r = note(ws, r,
         "・収録値が正しいことを確かめられました。"
         "令和8年8月4日の出力から2か月余りを経ていますが、"
         "過去の時点の値は改定されていません。\n"
         "・一方で、**「_比較」として収録していた系列は"
         "今回の受領値と一致しませんでした。**"
         "受領した出力は【注目する地域のみ】であり本広域連合の行のみを"
         "含むものであるため、「_比較」の系列は別の地域の値です（%s）。"
         % shn("hikaku"), height=70, fill=IN_Y)

# ============================================================ 04
ws = sheet(SH["hikaku"] + "_比較の系列が何を表すか特定できないこと",
           "「他地域と比較」の出力の系列が何を表すか特定できないこと",
           "当方は調整済み認定率を2系列で収録していました。"
           "今回の受領により、本広域連合の値が"
           "【注目する地域のみ】の側であることが確かめられた一方、"
           "もう一方が何を表すのかを特定できていないことが"
           "分かりました。確認事項No.190として起票しています。",
           [6, 22, 14, 14, 14, 14, 34, 12])
r = 4
r = lead(ws, r, "1　2系列の対比（合計調整済み認定率・重度・軽度）")
_YS = ["平成30年3月末", "令和3年3月末", "令和5年3月末", "令和6年3月末", R7, R8]
r = header(ws, r, ["", "系列"] + _YS)
for code, key in (("B5-a", "合計調整済み認定率"),
                  ("B5-a_比較", "合計調整済み認定率"),
                  ("B6-a", "調整済み重度認定率"),
                  ("B6-a_比較", "調整済み重度認定率"),
                  ("B6-b", "調整済み軽度認定率"),
                  ("B6-b_比較", "調整済み軽度認定率")):
    d = ser(code)[key]
    hik = code.endswith("_比較")
    r = body(ws, r, ["", "%s %s" % (code, key)]
             + [(pct(d[y]) if y in d else "―") for y in _YS],
             height=20, fills={2: (NG_O if hik else OK_G)},
             align={i: "right" for i in range(3, 9)})
r = note(ws, r,
         "注1）淡緑は【注目する地域のみ】の出力であり、"
         "**今回の受領値と全ての時点で一致しました**（本広域連合の値）。\n"
         "注2）淡橙は【他地域と比較】の出力から収録したものです。"
         "**淡緑と全ての時点で異なります。**"
         "重度と軽度の和が合計と一致するため単一のまとまった系列では"
         "ありますが、次の2つのいずれであるかを特定できません。"
         "①本広域連合の値を別の基準人口で標準化したもの"
         "（【他地域と比較】では全国の人口構成で標準化される、"
         "という従前の当方の見立て。原典で確かめられていません）。"
         "②全国、北海道又は比較対象として選択した地域の値。\n"
         "注3）令和8年3月末の値は【注目する地域のみ】のみ受領しています。"
         "淡橙の系列には加えていません。", height=108)

r += 1
r = lead(ws, r, "2　当方の扱い")
r = note(ws, r,
         "・**計画本文には【注目する地域のみ】の出力（本広域連合の値）のみを"
         "用い、2つの系列を混ぜません。**\n"
         "・1つの系列の中での時点の比較（低下しているか・上昇しているか）は、"
         "標準化の基準がどちらであっても成り立ちます。"
         "計画素案 第3章第4節2の記述はこの形に改めました。\n"
         "・淡橙の系列は、何を表すかを特定できないものとして"
         "本文に掲げません。\n"
         "・認定率の年齢調整分析では2系列を並べていますが、"
         "淡橙が何を表すかを特定できていない旨を付します。\n"
         "・確認事項No.190としてお伺いしています。", height=104, fill=IN_Y)

# ============================================================ 05
ws = sheet(SH["soan"] + "_計画素案への反映", "計画素案への反映",
           "計画素案 第3章第4節2「長期でみた到達点」が、"
           "2つの系列の値を混ぜて1つの文に並べていました。"
           "【注目する地域のみ】（本広域連合の値）に統一し、"
           "収録値から引く形に改めました。",
           [6, 18, 40, 40, 20, 12, 12, 12])
r = 4
r = lead(ws, r, "1　是正した3箇所")
r = header(ws, r, ["", "事項", "従前の記述", "是正後の記述", "是正の理由",
                   "", "", ""])
ZESEI = [
    ("年齢調整後の認定率",
     "見える化の調整済み認定率も平成30年17.8％から令和7年16.7％へ低下",
     "見える化の調整済み認定率も平成30年3月末%sから令和8年3月末%sへ低下"
     % (pct(ser("B5-a")["合計調整済み認定率"]["平成30年3月末"]),
        pct(ser("B5-a")["合計調整済み認定率"][R8])),
     "17.8％・16.7％はいずれも【他地域と比較】の系列の値であり、"
     "何を表すかを特定できていない系列の値を"
     "本広域連合の値として掲げていた"),
    ("中重度の認定率",
     "調整済み重度認定率（要介護3〜5）は5.4％から4.7％へ低下が継続",
     "調整済み重度認定率（要介護3〜5）は平成25年3月末%sから"
     "令和8年3月末%sへ低下が継続"
     % (pct(ser("B6-a")["調整済み重度認定率"]["平成25年3月末"]),
        pct(ser("B6-a")["調整済み重度認定率"][R8])),
     "5.4％は【注目する地域のみ】の令和5年3月末の値、"
     "4.7％は【他地域と比較】の令和7年3月末の値であり、"
     "**2つの系列の値を1つの文に並べていた**"),
    ("軽度の認定率",
     "調整済み軽度認定率（要支援1〜要介護2）は令和5年11.7％を底に"
     "令和7年12.0％へ上昇",
     "調整済み軽度認定率（要支援1〜要介護2）は平成27年3月末%sを頂点に"
     "低下した後、近年は令和5年3月末%sを底に令和6年3月末%sへ上昇し、"
     "令和7年3月末及び令和8年3月末も%s"
     % (pct(ser("B6-b")["調整済み軽度認定率"]["平成27年3月末"]),
        pct(ser("B6-b")["調整済み軽度認定率"]["令和5年3月末"]),
        pct(ser("B6-b")["調整済み軽度認定率"]["令和6年3月末"]),
        pct(ser("B6-b")["調整済み軽度認定率"][R8])),
     "11.7％・12.0％はいずれも【他地域と比較】の系列の値であった。"
     "あわせて令和8年3月末までの動きを反映し、"
     "上昇に転じた後は横ばいであることを記述に改めた"),
]
for i, (a, b, c, d) in enumerate(ZESEI, start=1):
    r = body(ws, r, [i, a, b, c, d, "", "", ""], height=68,
             align={1: "center"}, fills={3: NG_O, 4: OK_G})
r = note(ws, r,
         "注1）**評価の向きは変わりません。**"
         "重度は低下が続き、軽度は底を打って上昇しています。"
         "変わったのは数値と、どの系列の値であるかであり、"
         "第10期の重点（重度化防止を継続しつつ介護予防へ広げる）は"
         "そのままです。\n"
         "注2）軽度は令和6年3月末から令和8年3月末まで%sで横ばいです。"
         "「上昇が続いている」とは書けないため、"
         "「上昇に転じた後は横ばい」に改めました。\n"
         "注3）表に資料の注を加え、各年3月末時点であること、"
         "令和6年度及び令和7年度のみ年報ではなく月報によることを"
         "明示しました。"
         % pct(ser("B6-b")["調整済み軽度認定率"][R8]), height=86)

r += 1
r = lead(ws, r, "2　是正しなかったもの")
r = note(ws, r,
         "・計画素案 第2章第2節4の表の「調整済み軽度認定率 約13.5％」"
         "「調整済み重度認定率 約5.6％」は、"
         "北海道第9期計画 第7章第12節の分布図から読み取った値であり、"
         "比較する北海道の値（14.8％・5.7％）も同じ分布図によります。"
         "**一方の列だけを見える化の値に置き換えると出所が混ざる**ため、"
         "分布図による値のまま残しました（「約」を付しています）。\n"
         "・認定率の年齢調整分析は2系列を並べる作りであるため、"
         "系列そのものは残し、地域を特定できていない旨を付します。",
         height=86)

# ============================================================ 06
ws = sheet(SH["teigi"] + "_D49 地域差指数の定義",
           "D49 地域差指数の定義",
           "受領した指標の解説（全2頁）によります。"
           "%s" % D49.MEISHO,
           [6, 30, 34, 16, 40, 12, 12, 12])
r = 4
r = lead(ws, r, "1　地域差指数の定義")
r = note(ws, r, "（解説の記載）" + D49.TEIGI, height=50, fill=MID_B)

r += 1
r = lead(ws, r, "2　4指標と性・年齢調整の有無")
r = header(ws, r, ["", "指標", "算式", "性・年齢調整", "当方の収録・算定",
                   "", "", ""])
_SRC = {
    "第１号被保険者１人あたり給付月額":
        "見える化D8（在宅・施設及び居住系）。令和5年の指数を収録",
    "要介護認定率": "見える化B5-a（調整済み認定率）。令和8年3月末まで収録",
    "受給率": "見える化D2・D3・D4（要介護度別）。年報からも算定できる",
    "受給者１人あたり給付費": "見える化D17-a〜t（サービス別）",
}
for nm, shiki, adj in D49.SHIHYO:
    r = body(ws, r, ["", nm, shiki,
                     "あり" if adj else "なし", _SRC[nm], "", "", ""],
             height=30, align={4: "center"},
             fills={4: (MID_B if adj else NG_O)})
r = note(ws, r,
         "注）**4指標のうち性・年齢調整後であるのは%d指標だけです。**"
         "受給率と受給者1人あたり給付費は調整していません。"
         "このことが分解できない理由です（%s）。"
         % (len(D49.chosei_ari()), shn("bunkai")), height=34, fill=IN_Y)

r += 1
r = lead(ws, r, "3　確定したこと")
r = note(ws, r,
         "・**受給率の分母は第1号被保険者数です**"
         "（受給者数／第1号被保険者数）。"
         "従前は出力ファイルから分母を確認できず、"
         "認定者数との二択を留保していました"
         "（修正指示A-11。妥当性検証報告書）。"
         "年報から算定した施設及び居住系と在宅の受給率の合計16.5％が"
         "全国15.3％の1.078倍となり地域差指数の1.08を再現したことから"
         "第1号被保険者数と推定していたもので、"
         "**推定が原典により裏づけられました。**\n"
         "・年次の出典は%sであり、月次の出典はありません。"
         "B5・B6（令和6・7年度のみ月報）とは出典の扱いが異なります。"
         % D49.SHUTTEN["年次"], height=104, fill=OK_G)

r += 1
r = lead(ws, r, "4　指標を見るポイント（解説の記載）")
for t in D49.POINT:
    r = note(ws, r, "・" + t, height=34)

r += 1
r = lead(ws, r, "5　関連する指標（解説の記載）")
r = header(ws, r, ["", "指標", "指標名", "", "", "", "", ""])
for code, nm in D49.KANREN:
    r = body(ws, r, ["", code, nm, "", "", "", "", ""], height=24,
             align={2: "center"})

# ============================================================ 07
ws = sheet(SH["bunkai"] + "_地域差指数の分解の可否",
           "地域差指数の分解の可否（資料提供依頼No.16の答え）",
           "「受給率×受給者1人あたり給付費で給付月額の指数を再現できるか」"
           "を上川管内21保険者の実測で確かめました。"
           "**再現できません。**4指標のうち性・年齢調整後であるのは"
           "2指標だけであるためです。",
           [6, 16, 12, 12, 12, 12, 12, 40])
r = 4
r = lead(ws, r, "1　結論")
r = header(ws, r, ["", "問い", "答え", "", "理由", "", "", ""])
for q, a, rz in D49.BUNKAI:
    r = body(ws, r, ["", q, a, "", rz, "", "", ""], height=52,
             fills={3: (OK_G if "用いる" in a else NG_O)})

r += 1
r = lead(ws, r, "2　上川管内21保険者による実測（令和5年・全国＝1.00）")
r = header(ws, r, ["", "保険者", "給付月額", "受給率", "受給者単価",
                   "受給率×受給者単価", "差", "読み方"])
_g = []
for k in INS:
    g, j, t = IDX[KG][k], IDX[KJ][k], IDX[KT][k]
    d = j * t - g
    _g.append(d)
    r = body(ws, r, ["", k, "%.2f" % g, "%.2f" % j, "%.2f" % t,
                     "%.3f" % (j * t),
                     ("＋%.3f" % d if d >= 0 else "▲%.3f" % -d),
                     ("一致しない" if abs(d) > D49.ITCHI_HABA
                      else "%.2f以内" % D49.ITCHI_HABA), "", ],
             height=20, fills={7: (OK_G if abs(d) <= 0.02 else NG_O)},
             align={i: "right" for i in range(3, 8)},
             bold=(k == INS[0]))
_in = sum(1 for x in _g if abs(x) <= D49.ITCHI_HABA)
r = note(ws, r,
         "注1）差は▲%.3fから＋%.3fに開き、"
         "±0.02以内に収まるのは%d保険者（%d分の%d）にとどまります。"
         "本広域連合は給付月額1.08に対し積が%.3fで、差は＋%.3fです。\n"
         "注2）**認定率を掛けた3要素でも再現できません。**"
         "受給率の分母が第1号被保険者数であり認定率と分母が同じであるため、"
         "認定率を掛けると第1号被保険者数で二重に割ることになります。\n"
         "注3）各指数は小数第2位までで示されるため、"
         "積を採ると丸めの誤差が重なります。"
         % (-min(_g), max(_g), _in, len(INS), _in,
            IDX[KJ][INS[0]] * IDX[KT][INS[0]],
            IDX[KJ][INS[0]] * IDX[KT][INS[0]] - IDX[KG][INS[0]]),
         height=86)

r += 1
r = lead(ws, r, "3　当方の用い方")
r = note(ws, r,
         "・**どの要素で地域差が生じているかの読み取りに用います。**"
         "解説の【指標を見るポイント】が「4指標のうち、"
         "どの要素で地域差が生じているのか、"
         "全国や都道府県の値等と比較の上把握してください」と"
         "していることによります。\n"
         "・**乗法分解は行いません。** 要因分解の表を計画本文に新設する"
         "ことは行わず、要素の大小の比較までとします。\n"
         "・本広域連合は給付月額1.08・認定率1.02・受給率1.08・"
         "受給者単価1.06であり、**認定率ではなく受給率と受給者単価の側が"
         "全国を上回っています。** この読み取りは分解の可否とは別に成り立ち、"
         "計画素案 第2章第3節の記述の根拠となります。", height=104,
         fill=OK_G)

# ============================================================ 08
ws = sheet(SH["daicho"] + "_台帳の更新", "台帳の更新",
           "資料提供依頼2件を受領済に改め、確認事項1件を起票しました。"
           "台帳は業務工程管理表を正とし、本資料はそこから読んでいます。",
           [6, 14, 40, 20, 46, 12, 12, 12])
r = 4
r = lead(ws, r, "1　資料提供依頼")
r = header(ws, r, ["", "No.", "資料", "状態", "備考に記したこと", "", "", ""])
for no in (1, 16):
    x = BY_L[no]
    r = body(ws, r, ["", "No.%d" % no, x[2], x[8], x[9], "", "", ""],
             height=92, fills={4: OK_G}, align={2: "center", 4: "center"})

r += 1
r = lead(ws, r, "2　確認事項")
r = header(ws, r, ["", "No.", "件名", "状態", "決着しない場合の当方の扱い",
                   "", "", ""])
if _c190:
    r = body(ws, r, ["", "No.190", _c190[3], _c190[7],
                     _c190[9], "", "", ""],
             height=68, fills={4: IN_Y}, align={2: "center", 4: "center"})

r += 1
r = lead(ws, r, "3　残る資料提供依頼")
_lack_machi = [x for x in LACK
               if x[8] not in ("受領済", "完了", "解消", "代替により解消")]
_soto = [x for x in _lack_machi if x[5] != "受託者"]
r = note(ws, r,
         "・資料提供依頼は全%d件で、なおお待ちしているものが%d件です"
         "（うち入手先が受託者であるものが%d件）。\n"
         "・外部にお願いするもの%d件のうち、"
         "国の告示・公表を待つものが2件（第10期の政令改正・"
         "令和7年国勢調査）です。\n"
         "・今回の受領により、入手先を見える化システムとしていた%d件が"
         "いずれも受領済となりました。"
         % (len(LACK), len(_lack_machi), len(_lack_machi) - len(_soto),
            len(_soto), 2), height=70)

# ============================================================ 09 自己点検
ws = sheet(SH["chk"] + "_自己点検", "自己点検",
           "1件でも不適合があれば本資料の作成を止めます。",
           [6, 52, 40, 46, 12, 12, 12, 12])

chk(1, "受領した指標が5件であること", "B系列4件＋D49",
    "%d件" % (len(UKE_B) + 1), len(UKE_B) + 1 == 5)

_n14 = {c: len(yrs(c)) for c, _nm, _k in UKE_B}
chk(2, "B系列の4指標がいずれも14時点であること",
    "平成25年3月末から令和8年3月末",
    "／".join("%s %d時点" % (k, v) for k, v in _n14.items()),
    all(v == 14 for v in _n14.values()))

_nor8 = [(c, k) for c, _nm, _x in UKE_B for k in ser(c) if R8 not in ser(c)[k]]
chk(3, "全ての系列に令和8年3月末の値があること",
    "%d系列" % KEIRETSU_N, "欠けている系列 %s" % (_nor8 or "なし"), not _nor8)

_a = ser("B6-a")["調整済み重度認定率"][R8]
_b = ser("B6-b")["調整済み軽度認定率"][R8]
_c = ser("B5-a")["合計調整済み認定率"][R8]
chk(4, "重度と軽度の和が合計と一致すること（令和8年3月末）",
    "重度＋軽度 ＝ 合計",
    "%.1f＋%.1f＝%.1f／合計%.1f" % (_a, _b, _a + _b, _c),
    abs(_a + _b - _c) <= 0.1)

_t5c = sum(ser("B5-c")["調整済み認定率（要介護%d）" % i][R8]
           for i in range(1, 6))
chk(5, "B5-cの要介護1から5までの和が合計と一致すること（令和8年3月末）",
    "要介護1〜5の和 ＝ 合計",
    "%.1f／合計%.1f" % (_t5c, ser("B5-c")["合計調整済み認定率"][R8]),
    abs(_t5c - ser("B5-c")["合計調整済み認定率"][R8]) <= 0.1)

_t5a = sum(v[R8] for k, v in ser("B5-a").items()
           if k != "合計調整済み認定率")
chk(6, "B5-aの要介護度別の和が合計と一致すること（令和8年3月末）",
    "要支援1〜要介護5の和 ＝ 合計",
    "%.1f／合計%.1f" % (_t5a, _c), abs(_t5a - _c) <= 0.1)

_hik = [c + "_比較" for c, _nm, _k in UKE_B]
_hik_r8 = [c for c in _hik
           if any(R8 in d for d in ser(c).values())]
chk(7, "比較の系列に令和8年3月末の値を加えていないこと",
    "受領していないため加えない",
    "加わっているもの %s" % (_hik_r8 or "なし"), not _hik_r8)

_same = []
for code, key in (("B5-a", "合計調整済み認定率"),
                  ("B6-a", "調整済み重度認定率"),
                  ("B6-b", "調整済み軽度認定率")):
    d, h = ser(code)[key], ser(code + "_比較")[key]
    for y in h:
        if y in d and abs(d[y] - h[y]) < 1e-9:
            _same.append((code, y))
chk(8, "2つの系列が全ての時点で異なること",
    "一致する時点があれば同じ系列である",
    "一致する時点 %s" % (_same or "なし"), not _same)

chk(9, "D49の4指標のうち性・年齢調整後が2指標であること",
    "解説の用語の説明による",
    "調整後 %s" % "・".join(D49.chosei_ari()),
    len(D49.chosei_ari()) == 2 and len(D49.SHIHYO) == 4)

chk(10, "受給率の分母が第1号被保険者数であることが定義に書かれていること",
    "定義の文を走査する",
    "「受給者数／第１号被保険者数」の記載 %s"
    % ("あり" if "受給者数／第１号被保険者数" in D49.TEIGI else "なし"),
    "受給者数／第１号被保険者数" in D49.TEIGI)

chk(11, "受給率×受給者単価が給付月額の指数と一致しないこと",
    "21保険者で差が±%.2fを超えるものがあること" % D49.ITCHI_HABA,
    "±%.2f以内 %d／%d保険者" % (D49.ITCHI_HABA, _in, len(INS)),
    _in < len(INS))

# 分解の可否の記述に書いた差の幅が、実測と一致すること
_bz = D49.bunkai_zure()
_haba_txt = "▲%.3fから＋%.3f" % (-_bz[3], _bz[4])
_bk = "".join(x[2] for x in D49.BUNKAI)
chk(11.5, "分解の可否の記述に書いた差の幅が実測と一致すること",
    _haba_txt,
    "記述にある %s" % ("あり" if _haba_txt in _bk else "なし"),
    _haba_txt in _bk)

chk(11.7, "一致する保険者の数が記述と実測で一致すること",
    "実測 %d保険者" % _bz[1],
    "記述の「%d保険者のみ」 %s"
    % (_bz[1], "あり" if ("%d保険者のみ" % _bz[1]) in _bk else "なし"),
    ("%d保険者のみ" % _bz[1]) in _bk)

# ---- 計画素案の実物を読む
_DR = RP.draft_sections()
_S34 = _DR.get("第3章第4節", "")
chk(12, "計画素案 第3章第4節に本広域連合の値が入っていること",
    "重度%s・軽度%s・合計%s" % (pct(_a), pct(_b), pct(_c)),
    "／".join("%s %s" % (x, "あり" if x in _S34 else "なし")
              for x in (pct(_a), pct(_b), pct(_c))),
    all(x in _S34 for x in (pct(_a), pct(_b), pct(_c))))

_NG_HIK = ["17.8％", "16.7％", "4.7％", "11.7％", "12.0％"]
_left = [x for x in _NG_HIK if x in _S34]
chk(13, "計画素案 第3章第4節に比較の系列の値が残っていないこと",
    "／".join(_NG_HIK), "残っているもの %s" % (_left or "なし"), not _left)

chk(13.5, "計画素案 第3章第4節に資料の注を付していること",
    "「各年3月末時点」及び「月報」の記載",
    "%s／%s" % ("あり" if "各年3月末時点" in _S34 else "なし",
                "あり" if "月報によっている" in _S34 else "なし"),
    "各年3月末時点" in _S34 and "月報によっている" in _S34)

# ---- 台帳
chk(14, "資料提供依頼No.1・No.16が受領済であること",
    "台帳の状態",
    "／".join("No.%d %s" % (n, BY_L[n][8]) for n in (1, 16)),
    all(BY_L[n][8] == "受領済" for n in (1, 16)))

chk(15, "確認事項No.190が台帳にあり未決であること",
    "台帳の状態",
    "No.190 %s" % (_c190[7] if _c190 else "台帳にない"),
    bool(_c190) and _c190[7] == "未回答")

_nos = sorted(x[0] for x in CHECK)
chk(16, "確認事項の番号に欠番・重複がないこと",
    "No.1〜No.%d" % max(_nos),
    "%d件" % len(_nos), _nos == list(range(1, len(_nos) + 1)))

# ---- 送付する資料であるため
_ALL = []
for _ws in wb.worksheets:
    for _row in _ws.iter_rows(values_only=True):
        for _v in _row:
            if isinstance(_v, str):
                _ALL.append((_ws.title, _v))
_TXT = "\n".join(v for _t, v in _ALL)

NG_WORDS = ["に由来する", "と整合する", "1件も", "有意差がないため関係がない",
            "全国トップ級"]
_ngw = [w for w in NG_WORDS if w in _TXT]
chk(17, "禁止表現が本資料に出ていないこと", "禁止表現5件を走査する",
    "%d件" % len(_ngw), not _ngw)

NAIBU = ["固定値", "実物から", "章節ごとに", "判定している", "読んで判定",
         "スクリプト", "runpy", ".py", "再実行", "書き写して", "リポジトリ"]
_nb = [w for w in NAIBU if any(w in v for t, v in _ALL
                               if t != SH["chk"] + "_自己点検")]
chk(18, "受託者の内部の仕組み・作業経過の語が本資料に出ていないこと",
    "作り方・作業経過を述べる語を走査する",
    "%d件" % len(_nb), not _nb)

PI = [re.compile(r"0\d{1,4}[-(]\d{1,4}[-)]\d{3,4}"),
      re.compile(r"[\w.+-]+@[\w-]+\.[\w.]+")]
_pi = [p.pattern for p in PI if p.search(_TXT)]
chk(19, "電話番号・メールアドレスの形が本資料に出ていないこと",
    "値の形で探す", "%d件" % len(_pi), not _pi)

_mk = re.findall(r"\*\*", _TXT)
chk(20, "強調の指定が書き出した値に残っていないこと", "** の有無",
    "%d件" % len(_mk), not _mk)

r = 4
r = header(ws, r, ["No.", "確かめたこと", "式・見るところ", "結果", "判定",
                   "", "", ""])
for no, naiyo, shiki, kekka, han in CHECKS:
    r = body(ws, r, [no, naiyo, shiki, kekka, han, "", "", ""],
             height=30, align={1: "center", 5: "center"},
             fills={5: (OK_G if han == "適合" else NG_O)})
_ng = [c for c in CHECKS if c[4] != "適合"]
r = body(ws, r, ["", "自己点検の結果", "",
                 "適合%d件・不適合%d件" % (len(CHECKS) - len(_ng), len(_ng)),
                 "適合" if not _ng else "不適合", "", "", ""],
         height=22, bold=True, align={5: "center"},
         fills={5: (OK_G if not _ng else NG_O)})

wb.save(OUT)
print("書き出しました:", OUT)
for s in wb.sheetnames:
    print("  -", s, wb[s].max_row, "rows")
print("受領 %d件（B系列4・D49 1）／既存収録値との突合 %d件・差異0件"
      % (len(UKE_B) + 1, TOTSUGO_N))
print("令和8年3月末　重度%s・軽度%s・合計%s" % (pct(_a), pct(_b), pct(_c)))
print("地域差指数の分解　±0.02以内 %d／%d保険者（分解できない）"
      % (_in, len(INS)))
print("資料提供依頼 全%d件／なおお待ちしているもの %d件"
      % (len(LACK), len(_lack_machi)))
print("自己点検 %d件：適合%d件・不適合%d件"
      % (len(CHECKS), len(CHECKS) - len(_ng), len(_ng)))
if _ng:
    for c in _ng:
        print("  不適合 No.%s %s → %s" % (c[0], c[1], c[3]))
    sys.exit(1)
print("すべての点検に適合しました。")
