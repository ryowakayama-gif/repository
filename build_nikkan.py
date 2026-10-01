# -*- coding: utf-8 -*-
"""大雪地区広域連合 第10期介護保険事業計画　日次の状況と翌日の作業順位.

令和8年10月1日のご指示
  「毎日23時にWBSの更新をするようにお願いします。
    またペンディングしている内容を整理し、確認必要な項目の洗い出しと
    影響度を整理して優先順位づけするようにお願いします。
    ペンディングしておらず現状で作業継続可能な項目も洗い出しし、
    翌日の作業順位を整理するようにして下さい。」

━━ 本表の目的 ━━

毎日の終わりに、次の4つを1枚にする。

  ・WBS（業務内容別の進捗）の現況と、前に報告した時点からの動き
  ・ペンディング（決定待ち・資料待ち）の棚卸しと、影響度による優先順位
  ・ペンディングしておらず、当方の作業だけで進められるもの
  ・翌日の作業順位

━━ 数え方 ━━

**件数・進捗率・確認事項・据え置きはいずれも実物又はソースから読む。**
確認事項は業務工程管理表（正の台帳）を `ast` で読み、
据え置きは見込量算定を `runpy` で読む。**台帳を二重に持たない。**

**優先順位は機械で付ける。付け方は03シートに全て書く。**
人が並べ替えるときは、その理由を進捗の理由（`data_progress`）に残す。

シート構成
  00_この表について
  01_WBSの状況
  02_ペンディングの一覧と優先順位
  03_優先順位の付け方
  04_作業継続可能なもの
  05_翌日の作業順位
  06_自己点検

出力
  output/第10期計画_日次の状況と翌日の作業順位.xlsx

自己点検で1件でも不適合があると終了コード1で終わる。
内部保管（発注者への送付は想定しない）。
"""

import ast
import datetime
import io
import os
import re
import runpy
import subprocess
import sys

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import repo_paths as RP

sys.path.insert(0, RP.ROOT)
import data_progress as DP                                    # noqa: E402
import data_kitei as DK                                      # noqa: E402

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

ODIR = RP.OUTPUT
OUT = os.path.join(ODIR, "第10期計画_日次の状況と翌日の作業順位.xlsx")

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


# ============================================================ 暦
def _wareki(d):
    """西暦の date を和暦の文字列にする（令和のみ）。"""
    return "令和%d年%d月%d日" % (d.year - 2018, d.month, d.day)


def _yyyymm(d):
    return (d.year - 2018) * 12 + d.month        # 令和の通算月


# 基準日は実行した日とする（引数で上書きできる）。
_arg = [a for a in sys.argv[1:] if re.match(r"^\d{4}-\d{2}-\d{2}$", a)]
TODAY = (datetime.date.fromisoformat(_arg[0]) if _arg
         else datetime.date.today())
KIJUNBI = _wareki(TODAY)
ASU = _wareki(TODAY + datetime.timedelta(days=1))
NOW_M = _yyyymm(TODAY)


# ============================================================ 体裁
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
    ws.row_dimensions[2].height = 64
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = freeze
    return ws


def _plain(v):
    """xlsx は Markdown を解釈しないため、書き出しの時点で ** を落とす。"""
    return v.replace("**", "") if isinstance(v, str) else v


def header(ws, row, cols, height=30):
    for i, hh in enumerate(cols, start=1):
        c = ws.cell(row=row, column=i, value=_plain(hh))
        c.font = Font(name=FONT, size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=HEAD)
        c.alignment = Alignment(wrap_text=True, horizontal="center",
                                vertical="center")
        c.border = BORDER
    ws.row_dimensions[row].height = height
    return row + 1


def body(ws, row, vals, fills=None, height=20, align=None, bold=False,
         numfmt=None):
    for i, v in enumerate(vals, start=1):
        c = ws.cell(row=row, column=i, value=_plain(v))
        c.font = Font(name=FONT, size=9, bold=bold)
        c.border = BORDER
        ha = (align or {}).get(i, "left" if isinstance(v, str) else "right")
        c.alignment = Alignment(wrap_text=True, vertical="top", horizontal=ha)
        if numfmt and numfmt.get(i):
            c.number_format = numfmt[i]
        if fills and fills.get(i):
            c.fill = PatternFill("solid", fgColor=fills[i])
    ws.row_dimensions[row].height = height
    return row + 1


def lead(ws, row, text, span=10):
    c = ws.cell(row=row, column=1, value=_plain(text))
    c.font = Font(name=FONT, size=10.5, bold=True, color=NAVY)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    ws.row_dimensions[row].height = 20
    return row + 1


def note(ws, row, text, span=10, height=None, fill=GRAY):
    c = ws.cell(row=row, column=1, value=_plain(text))
    c.font = Font(name=FONT, size=8.5)
    c.fill = PatternFill("solid", fgColor=fill)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    ws.row_dimensions[row].height = height or (14 * (1 + text.count("\n")))
    return row + 1


# ============================================================ 台帳を読む
def _literal(fn, name):
    """ソースから literal の代入を読む（固定値を書き写さないため）。"""
    src = io.open(os.path.join(RP.ROOT, fn), encoding="utf-8").read()
    for node in ast.parse(src).body:
        if isinstance(node, ast.Assign) and any(
                getattr(t, "id", None) == name for t in node.targets):
            return ast.literal_eval(node.value)
    raise RuntimeError("%s に %s が見つからない" % (fn, name))


def _run(fn):
    """スクリプトを読み込んで名前空間を得る（標準出力は捨てる）。"""
    d = open(os.devnull, "w")
    o, e = sys.stdout, sys.stderr
    sys.stdout = sys.stderr = d
    try:
        return runpy.run_path(os.path.join(RP.ROOT, fn))
    finally:
        sys.stdout, sys.stderr = o, e
        d.close()


# 確認事項（正の台帳）。(0)No (1)業務内容 (2)工程 (3)表題 (4)内容
#                       (5)止めている成果物 (6)確認先 (7)状態 (8)期限 (9)回答
CHECK = _literal("build_process_control.py", "CHECK")
# 資料提供依頼。(0)No (1)領域 (2)資料 (3)内容 (4)確定する主張 (5)入手先
#               (6)希望時期 (7)優先度 (8)状態 (9)備考
LACK = _literal("build_process_control.py", "LACK")
# 決着しない場合の当方の扱い（既定値）。1か所から引く。
KITEI = DK.all_kitei()
# 据え置きの出どころ（A＝受託者で確定できる／B＝発注者・3町／C＝国）
KUBUN = _literal("build_mikomi_juryo_nashi.py", "KUBUN")
# 既定値のとおり計画素案へ反映し終えた確認事項（ご決定はなお待っているもの）。
# 反映し終えたものを翌日の作業として繰り返し挙げないために除く。
HANEI = {x[0] for x in _literal("build_ikenkokankai.py", "SUSUMETA")}

_S = _run("build_mikomiryo_santei.py")
# 据え置き。(0)区分 (1)項目 (2)理由 (3)当方の扱い (4)関係する確認事項 (5)効き
SUEOKI = _S["SUEOKI"]

KANRYO = ("完了", "了承済", "了承済（保管せず廃棄）", "代替により解消", "解決")
MACHI = [x for x in CHECK if x[7] not in KANRYO]
LACK_MACHI = [x for x in LACK if x[8] not in ("受領済", "完了", "解消")]


# ============================================================ 期限を読む
def _kigen_m(s):
    """「R8.10」のような期限を令和の通算月にする。読めなければ None。"""
    m = re.search(r"R(\d+)\s*[.．]\s*(\d+)", str(s or ""))
    return (int(m.group(1)) * 12 + int(m.group(2))) if m else None


# ============================================================ 月額への効き
# 据え置きの「効き」の欄から、月額を何円動かし得るかを読む。
# 確認事項No. は据え置きの「関係する確認事項」の欄から拾う。
_YEN = re.compile(r"([0-9,]+)\s*円")
_NO = re.compile(r"No\.\s*([0-9]+)")
# 「月額」に続く範囲だけを月額とみる。
# 効きの欄には月額でない額も現れる（「差5,854,697円で月額約▲7円」のように、
# 同じ文に総額と月額が並ぶ）。文中の円を無差別に拾うと総額を月額と読む。
_TSUKI_WIN = 24          # 「月額」の後ろ何文字までを月額の記載とみるか
GETSU_MAX = 2000         # 1件の月額の効きとしてあり得る上限（点検に用いる）


def _max_yen(text):
    """「月額」に続いて現れる円のうち最大のものを返す。無ければ None。"""
    t = text or ""
    v = []
    for m in re.finditer("月額", t):
        seg = t[m.end():m.end() + _TSUKI_WIN]
        v += [int(x.replace(",", "")) for x in _YEN.findall(seg)]
    return max(v) if v else None


GETSUGAKU = {}          # 確認事項No. -> (円, 据え置きの項目)
for i, s in enumerate(SUEOKI, start=1):
    y = _max_yen(s[5])
    if y is None:
        continue
    for no in _NO.findall(s[4]):
        no = int(no)
        if y > GETSUGAKU.get(no, (0, ""))[0]:
            GETSUGAKU[no] = (y, s[1])


# ============================================================ 影響度の採点
# 付け方は03シートに全て書く。ここを変えたら03シートも変わる（同じ表から作る）。
#
# 期限は当区域ではほとんどが超過又は当月であり、それだけでは順位が付かない。
# 順位を分けるのは**放置したときに何が動くか**であるため、
# 月額への効きと「止めている成果物」を重く、期限を軽くしている。
RULE = [
    ("期限", "回答期限が基準日の月より前（超過している）", 2),
    ("期限", "回答期限が基準日と同じ月", 1),
    ("期限", "回答期限が翌月", 0),
    ("止めている成果物", "概算（第1次・第2次）を止めている", 3),
    ("止めている成果物", "保険料に関わるものを止めている", 2),
    ("止めている成果物", "計画素案を止めている", 1),
    ("月額への効き", "据え置きに紐づき月額100円以上を動かし得る", 4),
    ("月額への効き", "同 50円以上100円未満", 2),
    ("月額への効き", "同 50円未満（効きの記載がある）", 1),
    ("法定記載事項", "介護保険法第117条の記載事項に直結する", 2),
    ("既定値", "決着しない場合の当方の扱いを置けていない", 3),
    ("当方で動かせるか", "国の告示・公布を待つもの（催促できない）", -2),
]
TOME_MAX = 4            # 「止めている成果物」の加点の上限
BAND = [("至急", 9, 99, "翌日に着手する。相手がある場合は翌日の午前に出す"),
        ("優先", 6, 8, "今週のうちに着手する"),
        ("通常", -99, 5, "期限の月に入ってから着手する")]

# 国の告示・公布を待つ据え置き（区分C）に紐づく確認事項No.
KUNI_NO = set()
for i, s in enumerate(SUEOKI, start=1):
    if KUBUN.get(i) == "C":
        KUNI_NO.update(int(n) for n in _NO.findall(s[4]))

_HOTEI = ("法第117条", "法定記載事項", "第117条")


def score(x):
    """確認事項1件の影響度を採点し、(合計, [内訳]) を返す。"""
    uchi, pt = [], 0
    km = _kigen_m(x[8])
    if km is not None:
        if km < NOW_M:
            uchi.append(("期限", "超過（%s）" % x[8], 2))
        elif km == NOW_M:
            uchi.append(("期限", "当月（%s）" % x[8], 1))
    tome, t = 0, str(x[5])
    if "概算" in t:
        tome += 3
        uchi.append(("止めている成果物", "概算", 3))
    if "保険料" in t:
        tome += 2
        uchi.append(("止めている成果物", "保険料", 2))
    if "計画素案" in t:
        tome += 1
        uchi.append(("止めている成果物", "計画素案", 1))
    if tome > TOME_MAX:          # 上限で頭打ちにする
        uchi = [u for u in uchi if u[0] != "止めている成果物"]
        uchi.append(("止めている成果物", "概算・保険料・計画素案（上限）",
                     TOME_MAX))
        tome = TOME_MAX
    g = GETSUGAKU.get(x[0])
    if g:
        p = 4 if g[0] >= 100 else (2 if g[0] >= 50 else 1)
        uchi.append(("月額への効き", "最大%s円（%s）"
                     % ("{:,}".format(g[0]), g[1]), p))
    if any(w in str(x[4]) for w in _HOTEI):
        uchi.append(("法定記載事項", "法第117条に直結", 2))
    if x[0] not in KITEI and "決着しない場合" not in str(x[9]):
        uchi.append(("既定値", "置けていない", 3))
    if x[0] in KUNI_NO:
        uchi.append(("当方で動かせるか", "国の告示・公布待ち", -2))
    pt = sum(u[2] for u in uchi)
    return pt, uchi


SCORED = []
for x in MACHI:
    pt, uchi = score(x)
    SCORED.append((pt, uchi, x))
SCORED.sort(key=lambda s: (-s[0], _kigen_m(s[2][8]) or 9999, s[2][0]))


# ============================================================ 当日のコミット
def _commits():
    """基準日のコミットを拾う（当日に何をしたかの記録）。"""
    try:
        out = subprocess.run(
            ["git", "-C", RP.ROOT, "log",
             "--since=%s 00:00" % TODAY.isoformat(),
             "--until=%s 23:59" % TODAY.isoformat(),
             "--format=%h\t%s"],
            capture_output=True, text=True, timeout=30)
        if out.returncode != 0:
            return None
        return [l.split("\t", 1) for l in out.stdout.splitlines() if "\t" in l]
    except Exception:
        return None


COMMITS = _commits()


# ============================================================ 00
ws = sheet("00_この表について", "日次の状況と翌日の作業順位",
           "毎日の終わりにWBSの現況・ペンディングの優先順位・"
           "当方の作業だけで進められるもの・翌日の作業順位を1枚にしたものです。"
           "基準日 " + KIJUNBI + "。内部保管。",
           [4, 26, 14, 56])
r = 4
r = lead(ws, r, "1　本表が答えるもの", span=4)
r = header(ws, r, ["", "問い", "シート", "何から求めているか"])
for i, (a, b, c_) in enumerate([
    ("WBSはいまどこまで進んでいるか", "01",
     "進捗率は進捗の記録から読みます。前に報告した時点の値も記録から読むため、"
     "更新しても過去の報告は書き換わりません"),
    ("何が決まらないと進まないか", "02",
     "業務工程管理表 03_確認事項一覧（正の台帳）と"
     "06_資料提供依頼一覧から、完了していないものを読みます"),
    ("どれから片付けるか", "02・03",
     "期限・止めている成果物・月額への効き・法定記載事項・既定値の有無の"
     "5つで点を付けて並べます。付け方は03シートに全て書いています"),
    ("決定を待たずに進められるものはあるか", "04",
     "据え置きのうち当方の側で確定できるもの（区分A）と、"
     "決着しない場合の扱い（既定値）を置けているものです"),
    ("明日は何からやるか", "05", "04を上に、02の上位への催促をその次に並べます"),
], start=1):
    r = body(ws, r, [i, a, b, c_], height=40, align={1: "center",
                                                     3: "center"})
r += 1
r = lead(ws, r, "2　この表の限界", span=4)
r = note(ws, r,
         "・**優先順位は機械で付けたものです。** 期限の近さと影響の大きさだけを"
         "見ており、相手のご都合・会議の日程・作業の段取りは入っていません。\n"
         "・**人が並べ替えてよいものです。** 並べ替えたときは、"
         "その理由を進捗の理由に残してください。\n"
         "・「作業継続可能」は決定を待たずに着手できるという意味であり、"
         "**その作業だけで成果品が確定するという意味ではありません**"
         "（既定値で進めたものは、決定を受けた時点で見直します）。\n"
         "・当日のコミットは作業の記録であり、作業量ではありません。",
         span=4, height=90, fill=IN_Y)

# ============================================================ 01
ws = sheet("01_WBSの状況", "WBS（業務内容別の進捗）の状況",
           "仕様書４（1）〜（12）ごとの進捗率と、前に報告した時点からの動きです。"
           "基準日 " + KIJUNBI + "。",
           [8, 34, 10, 12, 12, 46])
r = 4
_prev = sorted(DP.KIROKU)[-1] if DP.KIROKU else None
r = lead(ws, r, "1　進捗率", span=6)
r = header(ws, r, ["業務", "内容", "進捗", "状態", "前回からの動き",
                   "前に報告した時点（%s）" % (_prev or "―")])
N_WBS = 0
for no, nm, v, st, d in DP.PROGRESS:
    p = DP.KIROKU.get(_prev, {}).get(no) if _prev else None
    r = body(ws, r,
             [no, nm, ("―" if v is None else v), st, d,
              ("―" if p is None else "{:.0%}".format(p))],
             fills={3: (OK_G if (v or 0) >= 0.95 else
                        (NG_O if (v or 0) < 0.5 else IN_Y))},
             align={1: "center", 3: "center", 4: "center", 5: "center",
                    6: "center"},
             numfmt={3: "0%"})
    if v is not None:
        N_WBS += 1
r = body(ws, r, ["", "全体（数値で表せる%d件の平均）" % N_WBS,
                 DP.overall(), "", "",
                 ("―" if not _prev else
                  "{:.0%}".format(
                      sum(DP.KIROKU[_prev].values())
                      / len(DP.KIROKU[_prev])))],
         fills={3: MID_B}, bold=True, numfmt={3: "0%"},
         align={3: "center", 6: "center"})
r = note(ws, r,
         "注1）進捗の記録の基準日は %s です。"
         "**本表の基準日（%s）と違うときは、進捗の更新がまだ行われていません。**\n"
         "注2）進捗率を更新したら、更新前の値を記録へ移します。"
         "移さないと、過去の時点を記す資料（月次の報告・議事録）の数値が"
         "書き換わります。"
         % (DP.KIJUNBI, KIJUNBI), span=6, height=46)

r += 1
r = lead(ws, r, "2　当日の作業の記録（コミット）", span=6)
if COMMITS is None:
    r = note(ws, r, "git を読めませんでした（未実施）。", span=6, height=18,
             fill=NG_O)
elif not COMMITS:
    r = note(ws, r, "基準日のコミットはありません。", span=6, height=18)
else:
    r = header(ws, r, ["", "コミット", "内容", "", "", ""])
    for i, (h, s) in enumerate(COMMITS, start=1):
        r = body(ws, r, [i, h, s, "", "", ""], height=18,
                 align={1: "center", 2: "center"})
r = note(ws, r,
         "注）進捗の理由は、この記録をもとに書きます。"
         "**何をしたか**ではなく**それにより何が確定したか**を書きます。",
         span=6, height=28)

# ============================================================ 02
ws = sheet("02_ペンディングの一覧と優先順位",
           "ペンディング（決定待ち・資料待ち）の一覧と優先順位",
           "完了していない確認事項と資料提供依頼を、影響度の点の高い順に並べたものです。"
           "点の付け方は03シートに全て書いています。基準日 " + KIJUNBI + "。",
           [4, 5, 30, 20, 9, 9, 7, 34, 30])
r = 4
r = lead(ws, r, "1　確認事項（決定を求めるもの）　%d件" % len(MACHI), span=9)
r = header(ws, r, ["順", "No.", "確認事項", "止めている成果物", "確認先",
                   "期限", "点", "点の内訳", "決着しない場合の当方の扱い"])
for i, (pt, uchi, x) in enumerate(SCORED, start=1):
    kt = KITEI.get(x[0])
    if not kt:
        m = re.search(r"決着しない場合[^。]*。?", str(x[9]))
        kt = m.group(0) if m else "**置けていない**"
    r = body(ws, r,
             [i, x[0], x[3], x[5], x[6], x[8], pt,
              "／".join("%s %s（%+d）" % (u[0], u[1], u[2]) for u in uchi)
              or "―", kt],
             fills={7: (NG_O if pt >= BAND[0][1] else
                        (IN_Y if pt >= BAND[1][1] else MID_B)),
                    9: (NG_O if not KITEI.get(x[0])
                        and "決着しない場合" not in str(x[9]) else None)},
             height=48,
             align={1: "center", 2: "center", 5: "center", 6: "center",
                    7: "center"})
r = note(ws, r,
         "注1）点が%d以上を**至急**、%d以上%d以下を**優先**、"
         "%d以下を**通常**とみます（03シート）。\n"
         "注2）「決着しない場合の当方の扱い」が空のものは、"
         % (BAND[0][1], BAND[1][1], BAND[1][2], BAND[2][2]) +
         "**決まらないと当方の作業が止まります。** 先に扱いを決めます。\n"
         "注3）国の告示・公布を待つものは催促できないため点を2下げています。"
         "待つ以外にできることがないという意味であり、重要でないという意味ではありません。",
         span=9, height=56)

r += 1
r = lead(ws, r, "2　資料提供依頼（資料の提供を求めるもの）　%d件"
         % len(LACK_MACHI), span=9)
r = header(ws, r, ["順", "No.", "資料", "確定する主張", "入手先",
                   "希望時期", "優先", "状態", "領域"])
_PRI_N = {"最高": 0, "高": 1, "中": 2, "低": 3}
for i, x in enumerate(sorted(LACK_MACHI,
                             key=lambda y: (_PRI_N.get(y[7], 9),
                                            _kigen_m(y[6]) or 9999, y[0])),
                      start=1):
    r = body(ws, r, [i, x[0], x[2], x[4], x[5], x[6], x[7], x[8], x[1]],
             fills={7: {"最高": NG_O, "高": IN_Y, "中": MID_B}.get(x[7], GRAY)},
             height=40,
             align={1: "center", 2: "center", 6: "center", 7: "center",
                    8: "center", 9: "center"})
r = note(ws, r,
         "注）03シートの確認事項が「決定」を求めるものであるのに対し、"
         "本表は「資料の提供」を求めるものです。"
         "資料が届かない場合の扱いは04シートの区分で示しています。",
         span=9, height=28)

# ============================================================ 03
ws = sheet("03_優先順位の付け方", "優先順位の付け方（採点の規則）",
           "02シートの点は次の規則だけで付けています。"
           "規則を変えれば02シートの並びも変わります。基準日 " + KIJUNBI + "。",
           [4, 20, 50, 8, 40])
r = 4
r = lead(ws, r, "1　加点の規則", span=5)
r = header(ws, r, ["", "観点", "条件", "点", "該当件数"])
_cnt = {}
for pt, uchi, x in SCORED:
    for u in uchi:
        _cnt[(u[0], u[2])] = _cnt.get((u[0], u[2]), 0) + 1
for i, (kan, joken, p) in enumerate(RULE, start=1):
    r = body(ws, r, [i, kan, joken, p, "%d件" % _cnt.get((kan, p), 0)],
             fills={4: (NG_O if p >= 3 else (IN_Y if p > 0 else GRAY))},
             height=24, align={1: "center", 4: "center", 5: "center"})
r = note(ws, r,
         "注1）「止めている成果物」は重ねて当たることがあるため、"
         "合計%d点で頭打ちにしています"
         "（概算・保険料・計画素案の全てを止めていても%d点）。\n"
         "注2）「月額への効き」は、据え置きの一覧の「効き」の欄に現れる円のうち"
         "最も大きいものを採り、紐づく確認事項に割り当てています。"
         "据え置きに紐づかない確認事項には点が付きません"
         "（**効きが無いという意味ではなく、測っていないという意味です**）。\n"
         "注3）法定記載事項は、確認をお願いする内容に"
         "「法第117条」「法定記載事項」が現れるかで見ています。"
         "条文を引かずに書いているものは拾えません。",
         span=5, height=78)

r += 1
r = lead(ws, r, "2　点の分布", span=5)
r = header(ws, r, ["", "区分", "点", "件数", "扱い"])
for i, (nm, lo, hi, atsu) in enumerate(BAND, start=1):
    n = sum(1 for pt, _u, _x in SCORED if lo <= pt <= hi)
    r = body(ws, r, [i, nm,
                     ("%d以上" % lo if hi == 99 else
                      ("%d以下" % hi if lo == -99 else "%d〜%d" % (lo, hi))),
                     "%d件" % n, atsu],
             fills={2: (NG_O if nm == "至急" else
                        (IN_Y if nm == "優先" else MID_B))},
             height=22, align={1: "center", 2: "center", 3: "center",
                               4: "center"})

r += 1
r = lead(ws, r, "3　据え置きの出どころ（区分）", span=5)
r = header(ws, r, ["", "区分", "意味", "件数", "翌日の作業になるか"])
for i, (k, imi, naru) in enumerate([
    ("Z", "受託者の側で確定し、据え置きを解消した",
     "**ならない。** 終わっている"),
    ("A", "受託者の側で確定できる。資料は要らない", "**なる。** 04シートに掲げる"),
    ("B", "発注者・3町の資料待ち。届かなければ据え置きで確定する",
     "催促が作業になる。算定は止まらない"),
    ("C", "国の告示・公布を待つ。「受領できない」では済まない",
     "**ならない。** 待つ以外にできることがない"),
], start=1):
    n = sum(1 for v in KUBUN.values() if v == k)
    r = body(ws, r, [i, ("解消済み" if k == "Z" else k), imi, "%d件" % n, naru],
             fills={2: (OK_G if k in ("A", "Z") else
                        (IN_Y if k == "B" else GRAY))},
             height=30, align={1: "center", 2: "center", 4: "center"})

# ============================================================ 04
ws = sheet("04_作業継続可能なもの", "ペンディングしておらず作業継続可能なもの",
           "決定・資料の受領を待たずに当方の作業だけで進められるものです。"
           "基準日 " + KIJUNBI + "。",
           [4, 32, 52, 12, 34])
r = 4
TSUZUKI = []            # (区分, 作業, 内容, 出どころ)
for i, s in enumerate(SUEOKI, start=1):
    if KUBUN.get(i) != "A":
        continue
    # 据え置きの「項目」の欄は「同上」「―」のことがあるため、
    # 項目の名（s[0]）を主にし、具体の欄があるときだけ添える。
    nm = s[0] if s[1] in ("―", "同上", "") else "%s（%s）" % (s[0], s[1])
    TSUZUKI.append(("据え置き（区分A）", nm, s[3] + "。" + s[2],
                    "見込量算定の据え置き %d" % i))
for pt, uchi, x in SCORED:
    kt = KITEI.get(x[0])
    if kt and x[0] not in HANEI:
        TSUZUKI.append(("既定値で進める", x[3], kt,
                        "確認事項No.%d（点%d）" % (x[0], pt)))

r = lead(ws, r, "1　当方の作業だけで進められるもの　%d件" % len(TSUZUKI),
         span=5)
r = header(ws, r, ["", "区分", "作業", "内容", "出どころ"])
for i, (k, a, b, src) in enumerate(TSUZUKI, start=1):
    r = body(ws, r, [i, k, a, b, src],
             fills={2: (OK_G if k.startswith("据え置き") else MID_B)},
             height=44, align={1: "center"})
r = note(ws, r,
         "注1）「据え置き（区分A）」は、資料の受領を待たずに"
         "当方の側で確定できるものです。\n"
         "注2）「既定値で進める」は、決着しない場合の扱いを置けているため"
         "**決定を待たずに作業を進められる**ものです。"
         "決定を受けた時点で見直します。\n"
         "注3）ここに現れないもの（既定値を置けていない確認事項）は、"
         "**扱いを決めること自体が翌日の作業**になります。02シートを見ます。\n"
         "注4）既定値のとおり計画素案へ反映し終えたもの%d件"
         "（確認事項No.%s）は、作業としては終わっているため"
         "本表から除いています。ご決定はなお待っています。"
         % (len(HANEI), "・No.".join(str(x) for x in sorted(HANEI))),
         span=5, height=76)

r += 1
_NOKITEI = [x for _p, _u, x in SCORED
            if x[0] not in KITEI and "決着しない場合" not in str(x[9])]
r = lead(ws, r, "2　先に扱いを決めるもの（既定値を置けていない）　%d件"
         % len(_NOKITEI), span=5)
if not _NOKITEI:
    r = note(ws, r, "ありません。未決の確認事項の全件に扱いを置けています。",
             span=5, height=18, fill=OK_G)
else:
    r = header(ws, r, ["", "No.", "確認事項", "期限", "止めている成果物"])
    for i, x in enumerate(_NOKITEI, start=1):
        r = body(ws, r, [i, x[0], x[3], x[8], x[5]],
                 fills={2: NG_O}, height=36,
                 align={1: "center", 2: "center", 4: "center"})

# ============================================================ 05
ws = sheet("05_翌日の作業順位", "翌日（%s）の作業順位" % ASU,
           "04シートの作業継続可能なものを上に、"
           "02シートの上位への催促をその次に並べたものです。"
           "**機械で並べたものであり、人が並べ替えてよいものです。**",
           [4, 14, 34, 52, 30])
r = 4
JUNI = []
for k, a, b, src in TSUZUKI[:8]:
    JUNI.append(("進める", a, b, src))
# 催促・照会は優先以上の上位 N_SAISOKU 件まで。
# 既定値を置いたことで至急が0件になることがあるが、
# **相手のある作業は決着しない場合の扱いがあっても出す。**
N_SAISOKU = 5
for pt, uchi, x in SCORED:
    if pt < BAND[1][1]:
        break
    if sum(1 for j in JUNI if j[0] == "催促・照会") >= N_SAISOKU:
        break
    JUNI.append(("催促・照会", "%s（%s）" % (x[3], x[6]),
                 "点%d。%s" % (pt, "／".join(
                     "%s %s" % (u[0], u[1]) for u in uchi)),
                 "確認事項No.%d" % x[0]))
r = header(ws, r, ["順", "種別", "作業", "なぜ翌日か", "出どころ"])
for i, (k, a, b, src) in enumerate(JUNI, start=1):
    r = body(ws, r, [i, k, a, b, src],
             fills={2: (OK_G if k == "進める" else NG_O)},
             height=44, align={1: "center", 2: "center"})
r = note(ws, r,
         "注1）「進める」は決定を待たずに着手できるもの、"
         "「催促・照会」は相手のある作業です。"
         "**相手のある作業は午前のうちに出します**（返事に日数がかかるため）。\n"
         "注2）04シートの作業継続可能なものは上位8件まで、"
         "催促・照会は優先以上の上位5件までを掲げています。"
         "残りは02・04シートを見ます。\n"
         "注3）**この並びは機械が付けたものです。** 会議の日程・相手のご都合・"
         "作業の段取りは入っていません。並べ替えたときは理由を進捗の理由に残します。",
         span=5, height=62)

# ============================================================ 06 自己点検
_no_set = sorted(x[0] for x in CHECK)
_ketsu = [n for n in range(1, max(_no_set) + 1) if n not in set(_no_set)]
chk(1, "確認事項の台帳に欠番・重複がないこと",
    "No.1〜No.%d を数える" % max(_no_set),
    "%d件／欠番 %s／重複 %s"
    % (len(CHECK), _ketsu or "なし",
       "あり" if len(_no_set) != len(set(_no_set)) else "なし"),
    not _ketsu and len(_no_set) == len(set(_no_set)))

chk(2, "ペンディングの件数が台帳の未決の件数と一致すること",
    "状態が完了・了承済でないもの",
    "確認事項%d件／資料提供依頼%d件" % (len(MACHI), len(LACK_MACHI)),
    len(SCORED) == len(MACHI))

_re = all(score(x)[0] == pt for pt, _u, x in SCORED)
chk(3, "02シートの点が採点の規則から再現できること",
    "全件を採点し直して突合",
    "%d件／合わない %s" % (len(SCORED), "なし" if _re else "あり"), _re)

_desc = all(SCORED[i][0] >= SCORED[i + 1][0] for i in range(len(SCORED) - 1))
chk(4, "02シートが点の高い順に並んでいること",
    "隣り合う行の点を比べる", "降順 %s" % ("適合" if _desc else "不適合"),
    _desc)

_rule_kan = set(u[0] for _p, uchi, _x in SCORED for u in uchi)
_rule_def = set(k for k, _j, _p in RULE)
chk(5, "02シートで用いた観点が03シートの規則に全て載っていること",
    "内訳の観点 ⊆ 規則の観点",
    "内訳%d種／規則%d種／載っていない %s"
    % (len(_rule_kan), len(_rule_def), sorted(_rule_kan - _rule_def) or "なし"),
    not (_rule_kan - _rule_def))

# Z＝受託者の側で確定し、据え置きを解消したもの
_na, _nb, _nc, _nz = (sum(1 for v in KUBUN.values() if v == k)
                      for k in "ABCZ")
_yen_ng = ["No.%d %s円（%s）" % (no, "{:,}".format(y), nm)
           for no, (y, nm) in GETSUGAKU.items() if y > GETSU_MAX]
chk(5.5, "月額への効きとして総額を拾っていないこと",
    "「月額」に続く範囲だけを拾い、%s円を超えるものがないこと"
    % "{:,}".format(GETSU_MAX),
    "%d件を拾った（最大%s円）／上限超え %s"
    % (len(GETSUGAKU),
       "{:,}".format(max([v[0] for v in GETSUGAKU.values()] or [0])),
       _yen_ng or "なし"), not _yen_ng)

chk(6, "据え置きの区分の件数が据え置きの件数と合うこと",
    "Z（解消済み）＋A＋B＋C ＝ 据え置き",
    "%d＋%d＋%d＋%d＝%d／据え置き%d件"
    % (_nz, _na, _nb, _nc, _nz + _na + _nb + _nc, len(SUEOKI)),
    _nz + _na + _nb + _nc == len(SUEOKI))

_pv = [v for _n, _nm, v, _s, _d in DP.PROGRESS if v is not None]
chk(7, "進捗率が0から1の間にあること",
    "進捗を数値で表せる%d件" % len(_pv),
    "最小%.2f・最大%.2f・全体%d％"
    % (min(_pv), max(_pv), DP.overall_pct()),
    all(0.0 <= v <= 1.0 for v in _pv))

_same = (DP.KIJUNBI == KIJUNBI)
chk(8, "進捗の記録の基準日が本表の基準日と一致すること",
    "data_progress の基準日 ＝ 本表の基準日",
    "記録 %s／本表 %s%s"
    % (DP.KIJUNBI, KIJUNBI, "" if _same else "（WBSの更新がまだです）"),
    _same)

_kir = DP.KIJUNBI not in DP.KIROKU
chk(9, "現時点の値が記録（過去の時点）に紛れていないこと",
    "記録に現時点の基準日が無いこと",
    "記録%d件（最新 %s）" % (len(DP.KIROKU), sorted(DP.KIROKU)[-1]), _kir)

chk(10, "翌日の作業順位が1件以上あること",
    "進める＋催促・照会",
    "%d件（進める%d・催促%d）"
    % (len(JUNI), sum(1 for j in JUNI if j[0] == "進める"),
       sum(1 for j in JUNI if j[0] == "催促・照会")), len(JUNI) >= 1)

chk(11, "作業継続可能なものを数えられていること",
    "区分A＋既定値のあるもの",
    "%d件（区分A %d・既定値 %d）"
    % (len(TSUZUKI), sum(1 for t in TSUZUKI if t[0].startswith("据え置き")),
       sum(1 for t in TSUZUKI if t[0] == "既定値で進める")),
    len(TSUZUKI) > 0)

NG_WORDS = ["に由来する", "と整合する", "1件も", "有意差がないため関係がない",
            "全国トップ級"]
PI = [re.compile(r"0\d{1,4}[-(]\d{1,4}[-)]\d{3,4}"),
      re.compile(r"[\w.+-]+@[\w-]+\.[\w.]+")]
_ngw, _pi = [], []
for _ws in wb.worksheets:
    for _row in _ws.iter_rows(values_only=True):
        for _v in _row:
            if not isinstance(_v, str):
                continue
            if _ws.title != "06_自己点検":
                _ngw += [w for w in NG_WORDS if w in _v]
            _pi += [p.pattern for p in PI if p.search(_v)]
chk(12, "禁止表現が本表に出ていないこと", "／".join(NG_WORDS),
    "%d件" % len(_ngw), not _ngw)
chk(13, "電話番号・メールアドレスの形が本表に出ていないこと",
    "値の形（正規表現）で走査", "%d件" % len(_pi), not _pi)

_ast_n = sum(1 for _ws in wb.worksheets
             for _row in _ws.iter_rows(values_only=True)
             for _v in _row if isinstance(_v, str) and "**" in _v)
chk(14, "強調の指定（**）がセルに残っていないこと",
    "xlsx は Markdown を解釈しないため書き出しの時点で落とす",
    "残り%d件" % _ast_n, _ast_n == 0)

ws = sheet("06_自己点検", "自己点検",
           "本表の値が台帳・算定と合っていることを機械で確かめたものです。"
           "1件でも不適合があると終了コード1で終わります。",
           [5, 40, 40, 44, 9])
r = header(ws, 4, ["No.", "点検した内容", "式・条件", "結果", "判定"])
for c in CHECKS:
    body(ws, r, list(c), fills={5: (OK_G if c[4] == "適合" else NG_O)},
         height=30)
    r += 1
_NG = sum(1 for c in CHECKS if c[4] != "適合")
body(ws, r, ["", "自己点検の結果", "",
             "適合%d件・不適合%d件" % (len(CHECKS) - _NG, _NG),
             "適合" if _NG == 0 else "不適合"],
     fills={5: (OK_G if _NG == 0 else NG_O)}, height=24, bold=True)

# ============================================================ 出力
os.makedirs(ODIR, exist_ok=True)
wb.save(OUT)

print("書き出しました: %s" % OUT)
for _ws in wb.worksheets:
    print("  - %s %d rows" % (_ws.title, _ws.max_row))
print("基準日 %s（翌日 %s）" % (KIJUNBI, ASU))
print("WBS 全体%d％（進捗の記録の基準日 %s）"
      % (DP.overall_pct(), DP.KIJUNBI))
print("ペンディング 確認事項%d件（%s）／資料提供依頼%d件"
      % (len(MACHI),
         "・".join("%s%d" % (b[0], sum(1 for p, _u, _x in SCORED
                                      if b[1] <= p <= b[2])) for b in BAND),
         len(LACK_MACHI)))
print("作業継続可能 %d件／翌日の作業 %d件" % (len(TSUZUKI), len(JUNI)))
print("自己点検 %d件：適合%d件・不適合%d件"
      % (len(CHECKS), len(CHECKS) - _NG, _NG))
if _NG:
    print("不適合があります。")
    sys.exit(1)
print("すべての点検に適合しました。")
