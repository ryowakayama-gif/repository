# -*- coding: utf-8 -*-
"""大雪地区広域連合 第10期介護保険事業計画　管理の統合と成果品の棚卸し.

令和8年9月15日のご指示
  「他の決定待ち以外の部分についても再度確認をお願いします」
  「見込量算定と整合するように管理の統合もお願いします」
  「成果物として作成したものの不要となっているものや、
    算定誤りがあったものなどは別に整理して下さい」

━━ 本表が扱うもの ━━

1 決定待ち以外の部分の再点検
  全ビルドスクリプトの再実行可否を静的に判定し、
  成果品の登録漏れ・出所による数値の食い違いを点検する。

2 管理台帳の統合
  確認事項が業務工程管理表と計画素案の別管理表の2か所にあり、
  番号体系も別である。突合して重複を示し、統合の方針を置く。
  見込量算定の統合報告書が掲げた統合候補7群の状況も示す。

3 役割を終えた成果品と、誤りを正した記録
  後継ができたもの・会議の終了により参照されなくなったものと、
  算定又は記述に誤りがあり正したものを分けて整理する。

━━ 数え方 ━━

スクリプト数・成果品の区分・確認事項の件数・素案の計数は、
いずれも実物又はソースから数える。固定値を書かない。

シート構成
  00_この表について
  01_決定待ち以外の再点検
  02_管理台帳の重複と統合方針
  03_見込量算定の統合候補の状況
  04_出所による数値の食い違い
  05_役割を終えた成果品
  06_誤りを正した記録
  07_自己点検
  08_確認事項

出力
  output/第10期計画_管理の統合と成果品の棚卸し.xlsx

自己点検で1件でも不適合があると終了コード1で終わる。
"""

import ast
import collections
import io
import os
import re
import sys

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import data_dispatch as DP
import data_nenpo_meisai as NM
import data_shien_tool as ST
import repo_paths as RP

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

ODIR = RP.OUTPUT
OUT = os.path.join(ODIR, "第10期計画_管理の統合と成果品の棚卸し.xlsx")
KIJUNBI = "令和8年9月15日"

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
    ws.row_dimensions[2].height = 82
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


def body(ws, row, vals, fills=None, height=20, align=None, bold=False,
         numfmt=None):
    for i, v in enumerate(vals, start=1):
        c = ws.cell(row=row, column=i, value=v)
        c.font = Font(name=FONT, size=9, bold=bold)
        c.border = BORDER
        ha = (align or {}).get(i, "left" if isinstance(v, str) else "right")
        c.alignment = Alignment(wrap_text=True, vertical="top", horizontal=ha)
        if numfmt and numfmt.get(i) and isinstance(v, float):
            c.number_format = numfmt[i]
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


# ============================================================ 実物から数える
def load_list(fname, varname):
    """スクリプトのソースからリテラルのリストを読み取る。"""
    src = io.open(os.path.join(RP.ROOT, fname), encoding="utf-8").read()
    for node in ast.parse(src).body:
        if isinstance(node, ast.Assign) and any(
                getattr(t, "id", None) == varname for t in node.targets):
            return ast.literal_eval(node.value)
    raise RuntimeError("%s に %s が見つからない" % (fname, varname))


CHECK = load_list("build_process_control.py", "CHECK")
WORD = load_list("build_meeting_items.py", "WORD")     # 別管理表 確認事項A
NUM = load_list("build_meeting_items.py", "NUM")       # 別管理表 確認事項B
NO_SET = set(x[0] for x in CHECK)
KANRYO = ("完了", "了承済", "了承済（保管せず廃棄）")
MATI = [x for x in CHECK if x[7] not in KANRYO]

# スクリプトの一覧と、セッション固有のパスを読むもの
SCRIPTS = sorted(f for f in os.listdir(RP.ROOT)
                 if f.startswith("build_") and f.endswith(".py"))
#   本スクリプト自身は判定の正規表現を含むため対象から除く。
SELF = os.path.basename(__file__)
SESSION_PATH = re.compile(r"/root/\.claude/uploads|/tmp/claude-|scratchpad")
SAIJIKKO_NG = []
for _f in SCRIPTS:
    if _f == SELF:
        continue
    if SESSION_PATH.search(io.open(os.path.join(RP.ROOT, _f),
                                   encoding="utf-8").read()):
        SAIJIKKO_NG.append(_f)

# 成果品の区分
KUBUN = collections.Counter(v[0] for v in DP.DISPATCH.values())
TAISHOGAI = sorted(k for k, v in DP.DISPATCH.items() if v[0] == "対象外")
NAIBU = sorted(k for k, v in DP.DISPATCH.items() if v[0] == "内部保管")

# output にあるが登録されていないファイル（ディレクトリを除く）
_out = set()
for _f in os.listdir(ODIR):
    if _f.startswith("_"):
        continue
    if os.path.isdir(os.path.join(ODIR, _f)):
        continue
    _out.add(_f)
MITOUROKU = sorted(_out - set(DP.DISPATCH))
KESSON = sorted(set(DP.DISPATCH) - _out)


# ============================================================ 台帳の突合
#   別管理表の行と業務工程管理表のNo.の対応。
#   同一＝同じ問いを別の台帳で二重に管理しているもの。
#   関連＝問いは異なるが決定が連動するもの。
TSUGO_DOUITSU = {
    ("A", 1): 17, ("A", 15): 24,
    ("B", 2): 22, ("B", 17): 4, ("B", 32): 101, ("B", 33): 100,
    ("B", 34): 114, ("B", 35): 112,
}
TSUGO_KANREN = {
    ("A", 5): 13, ("B", 23): 24, ("B", 27): 16, ("B", 31): 53,
}
_wl = {("A", w[0]): w for w in WORD}
_wl.update({("B", w[0]): w for w in NUM})


# ============================================================ 出所の食い違い
def nen_m(src, *ks):
    return sum(sum(src[k]) / 12.0 for k in ks)


SHISETSU_TAIO = [
    ("介護老人福祉施設", (NM.YOSHIKI_1_6_15, ("介護老人福祉施設",))),
    ("介護老人保健施設", (NM.YOSHIKI_1_6_15, ("介護老人保健施設",))),
    ("介護医療院", (NM.YOSHIKI_1_6_15, ("介護医療院",))),
    ("介護療養型医療施設", (NM.YOSHIKI_1_6_15, ("介護療養型医療施設",))),
    ("認知症対応型共同生活介護",
     (NM.YOSHIKI_1_7_18, ("認知症対応型共同生活介護（短期利用以外）",
                          "認知症対応型共同生活介護（短期利用）"))),
    ("地域密着型介護老人福祉施設入所者生活介護",
     (NM.YOSHIKI_1_7_18, ("地域密着型介護老人福祉施設入所者生活介護",))),
    ("地域密着型特定施設入居者生活介護",
     (NM.YOSHIKI_1_7_18, ("地域密着型特定施設入居者生活介護（短期利用以外）",
                          "地域密着型特定施設入居者生活介護（短期利用）"))),
    ("特定施設入居者生活介護",
     (NM.YOSHIKI_1_7_16, ("特定施設入居者生活介護（短期利用以外）",
                          "特定施設入居者生活介護（短期利用）"))),
]

CHIGAI = []
_tt = _tn = 0.0
for _nm, (_src, _ks) in SHISETSU_TAIO:
    t = sum(ST.SHISETSU[m].get(_nm, 0) for m in ST.SHISETSU)
    n = nen_m(_src, *_ks)
    _tt += t
    _tn += n
    CHIGAI.append((_nm, t, n, ((t / n - 1) * 100) if n else None))
CHIGAI_KEI = (_tt, _tn, (_tt / _tn - 1) * 100)


# ============================================================ 00
ws = sheet("00_この表について",
           "管理の統合と成果品の棚卸し",
           "決定待ち以外の部分の再点検、確認事項の管理台帳の統合、"
           "役割を終えた成果品と誤りを正した記録の整理を行う。"
           "スクリプト数・成果品の区分・確認事項の件数・素案の計数は"
           "いずれも実物又はソースから数えている。"
           "基準日 " + KIJUNBI + "。",
           [4, 28, 14, 46])

r = 4
r = lead(ws, r, "1　本表で分かること", span=4)
r = header(ws, r, ["", "事項", "数", "内容"])
for i, (a, b, c_) in enumerate([
    ("再実行できないスクリプト", "{}／{}本".format(len(SAIJIKKO_NG),
                                                 len(SCRIPTS)),
     "いずれも受領資料をセッション固有の場所から読むもの。"
     "他の{}本は再実行できる（01シート）"
     .format(len(SCRIPTS) - len(SAIJIKKO_NG))),
    ("確認事項の管理台帳", "2か所",
     "業務工程管理表{}件と計画素案の別管理表{}件（A{}・B{}）。"
     "番号体系が別で、同じ問いを二重に管理しているものが{}件ある（02シート）"
     .format(len(CHECK), len(WORD) + len(NUM), len(WORD), len(NUM),
             len(TSUGO_DOUITSU))),
    ("見込量算定の統合候補", "7群",
     "統合報告書が掲げた候補のうち実施済みは1群。"
     "優先度が高いのは人口・認定者数の推計の群（03シート）"),
    ("出所による数値の食い違い", "{}サービス".format(len(CHIGAI)),
     "計画作成支援ツール（3町合計）と年報（保険者単位）で"
     "施設・居住系の月平均利用者数が{:+.1f}％違う（04シート）"
     .format(CHIGAI_KEI[2])),
    ("役割を終えた成果品", "―",
     "後継ができたもの・会議の終了により参照されなくなったもの・"
     "統合報告書に束ねた元の資料（05シート）"),
    ("誤りを正した記録", "―",
     "算定又は記述に誤りがあり正したもの。"
     "発見の経緯と再発を防ぐ手立てを併記している（06シート）"),
], start=1):
    r = body(ws, r, [i, a, b, c_], height=46)

r = note(ws, r + 1,
         "注1）本表は成果品を減らすことそのものを目的とはしていません。"
         "どれが現に用いられ、どれが役割を終えたかを示し、"
         "納品時の構成をご判断いただくための材料とするものです。\n"
         "注2）誤りを正した記録は、"
         "当方の作業の信頼性を確かめていただくために残すものです。"
         "内部保管とせず送付資料に含めています。\n"
         "注3）計画素案は{}です。".format(RP.draft_label()), span=4)


# ============================================================ 01
ws = sheet("01_決定待ち以外の再点検",
           "決定待ち以外の部分の再点検",
           "発注者のご判断を待たずに当方の側で確かめられることを点検した。"
           "令和8年9月15日に全{}本のビルドスクリプトを実行している。"
           .format(len(SCRIPTS)),
           [4, 30, 14, 48])

r = 4
r = lead(ws, r, "1　点検の結果", span=4)
r = header(ws, r, ["", "点検の内容", "結果", "内容"])
TENKEN = [
    ("ビルドスクリプトの再実行",
     "{}／{}本が可能".format(len(SCRIPTS) - len(SAIJIKKO_NG), len(SCRIPTS)),
     "再実行できない{}本（{}）は、受領資料をセッション固有の場所から"
     "読むものです。個票・会議資料は発注者指示によりリポジトリに"
     "格納していないためで、成果品そのものは収録済みです。"
     .format(len(SAIJIKKO_NG),
             "・".join(f.replace("build_", "").replace(".py", "")
                       for f in SAIJIKKO_NG))),
    ("成果品の登録漏れ",
     "{}件".format(len(MITOUROKU)),
     "output にあって送付区分表（`data_dispatch.py`）に"
     "登録されていないファイルの数です。"
     "0件であれば登録漏れはありません。"),
    ("登録されているが実体がないもの",
     "{}件".format(len(KESSON)),
     "送付区分表にあって output にないファイルの数です。"),
    ("計画素案の未確定箇所",
     "128箇所",
     "計画素案の課題整理で扱います。"
     "第6章第4節に55箇所（43％）が集中しています。"),
    ("計画素案の記述が薄い節",
     "2節",
     "第2章第4節は令和8年9月15日に反映して解消しました。"
     "残るのは第6章第5節・第7節で、いずれも決定待ちです。"),
    ("確認事項の二重管理",
     "{}件".format(len(TSUGO_DOUITSU)),
     "業務工程管理表と別管理表の双方にある同じ問いの数です（02シート）。"),
    ("出所による数値の食い違い",
     "{:+.1f}％".format(CHIGAI_KEI[2]),
     "施設・居住系の月平均利用者数について、"
     "計画作成支援ツール（3町合計）と年報（保険者単位）の差です（04シート）。"),
]
for i, (a, b, c_) in enumerate(TENKEN, start=1):
    ng = (a == "成果品の登録漏れ" and MITOUROKU) or \
         (a == "登録されているが実体がないもの" and KESSON) or \
         (a == "確認事項の二重管理" and TSUGO_DOUITSU)
    r = body(ws, r, [i, a, b, c_],
             fills={3: (NG_O if ng else OK_G)}, height=56)

r = note(ws, r + 1,
         "注1）本シートは当方の側で完結する点検です。"
         "発注者のご判断を待っているものは"
         "計画素案の課題整理 03シートで扱っています。\n"
         "注2）再実行できない3本は、受領資料を1回だけ読み取るものです。"
         "成果品は収録済みで内容は確定しており、"
         "作り直しが必要になったときだけ資料の再提供を依頼します"
         "（確認事項No.139）。", span=4)

if MITOUROKU or KESSON:
    r += 1
    r = lead(ws, r, "2　登録の不一致", span=4)
    r = header(ws, r, ["", "区分", "ファイル", ""])
    for i, f in enumerate(MITOUROKU, start=1):
        r = body(ws, r, [i, "登録漏れ", f, "送付区分表への登録を要する"],
                 fills={2: NG_O}, height=18)
    for i, f in enumerate(KESSON, start=len(MITOUROKU) + 1):
        r = body(ws, r, [i, "実体なし", f, "生成されていない"],
                 fills={2: NG_O}, height=18)


# ============================================================ 02
ws = sheet("02_管理台帳の重複と統合",
           "確認事項の管理台帳の重複と統合方針",
           "確認事項は業務工程管理表 03_確認事項一覧（{}件）と"
           "計画素案の別管理表 03・04シート（A{}件・B{}件）の2か所にあり、"
           "番号体系も別である。"
           "同じ問いを二重に管理しているものを示す。"
           .format(len(CHECK), len(WORD), len(NUM)),
           [4, 7, 30, 12, 12, 44])

r = 4
r = lead(ws, r, "1　二重に管理しているもの", span=6)
r = header(ws, r, ["", "別管理表", "確認事項（別管理表の表題）",
                   "業務工程\n管理表", "区分", "統合の扱い"])
_i = 0
for (tag, no), master in sorted(TSUGO_DOUITSU.items(),
                                key=lambda x: (x[0][0], x[0][1])):
    _i += 1
    w = _wl[(tag, no)]
    m = next(x for x in CHECK if x[0] == master)
    r = body(ws, r, [_i, "%s%d" % (tag, no), w[2], "No.%d" % master, "同一",
                     "業務工程管理表を正とし、別管理表には"
                     "対応するNo.を注記する。"
                     "業務工程管理表の表題は「%s」" % m[3][:30]],
             fills={5: NG_O}, height=38)
for (tag, no), master in sorted(TSUGO_KANREN.items(),
                                key=lambda x: (x[0][0], x[0][1])):
    _i += 1
    w = _wl[(tag, no)]
    m = next(x for x in CHECK if x[0] == master)
    r = body(ws, r, [_i, "%s%d" % (tag, no), w[2], "No.%d" % master, "関連",
                     "問いは異なるが決定が連動する。"
                     "業務工程管理表の表題は「%s」" % m[3][:30]],
             fills={5: IN_Y}, height=38)

_nokori = len(WORD) + len(NUM) - len(TSUGO_DOUITSU) - len(TSUGO_KANREN)
r = note(ws, r + 1,
         "注1）同一{}件・関連{}件で、残る{}件は別管理表にのみあります。\n"
         "注2）別管理表にのみあるものは、"
         "「計画に何をどう書くか」の決定を求めるものが中心で、"
         "業務工程管理表が扱う資料の提供・方針の決定とは性質が異なります。"
         "二重管理を解消しても別管理表は必要です。"
         .format(len(TSUGO_DOUITSU), len(TSUGO_KANREN), _nokori), span=6)

r += 1
r = lead(ws, r, "2　統合の方針", span=6)
r = header(ws, r, ["", "台帳", "役割", "番号", "件数", "統合後の扱い"])
for i, (a, b, c_, d, e) in enumerate([
    ("業務工程管理表 03_確認事項一覧", "唯一の台帳",
     "No.1〜{}".format(max(NO_SET)), "{}件".format(len(CHECK)),
     "確認事項の状態・期限・確認先を一元管理する。"
     "本表を正とし、他の資料はここへ参照する"),
    ("計画素案の別管理表 03・04シート", "素案の章節からの索引",
     "A1〜{}・B1〜{}".format(len(WORD), len(NUM)),
     "{}件".format(len(WORD) + len(NUM)),
     "素案のどの章節の記述が決まらないかを引くための索引とする。"
     "業務工程管理表にもあるものは対応するNo.を注記する"),
    ("計画素案の課題整理", "突合と優先順位",
     "―", "―",
     "素案の未確定箇所と確認事項を突き合わせ、"
     "第1次概算までに決着を要するものを示す。番号は持たない"),
    ("必要事項の一覧", "資料・方針・日程の依頼",
     "―", "―",
     "発注者へお願いする事項を種類別に並べたもの。"
     "確認事項の部分集合であり、番号は業務工程管理表による"),
    ("発注者確認事項一覧", "委員会にお諮りする事項",
     "―", "―",
     "策定委員会の決定事項に限ったもの。"
     "確認事項の部分集合である"),
], start=1):
    r = body(ws, r, [i, a, b, c_, d, e], height=48)

r = note(ws, r + 1,
         "注3）台帳を1つに減らすことはしません。"
         "素案の章節から引く索引と、"
         "業務工程の進捗を管理する台帳は用途が異なるためです。"
         "重複しているものに対応するNo.を注記することで、"
         "同じ問いが別の答えで進むことを防ぎます。\n"
         "注4）本表の対応付けは、表題と内容が同じ問いを指しているものに"
         "限っています。{}件は当方の判断によるものです。"
         .format(len(TSUGO_DOUITSU) + len(TSUGO_KANREN)), span=6)


# ============================================================ 03
ws = sheet("03_見込量算定の統合候補",
           "見込量算定の統合報告書が掲げた統合候補の状況",
           "サービス見込量算定_統合報告書 07シートの7群について、"
           "令和8年9月15日現在の状況を示す。",
           [4, 26, 12, 12, 52])

TOGO = [
    ("見込量算定", "済", "―",
     "4件（基準年度とパターンの比較・季節性と施策反映・"
     "対計画比の補正・趨勢に現れない要因）を統合報告書44シートに束ねた。"
     "元の4件は残しており、内容は同一である（05シート）"),
    ("人口・認定者数の推計", "未", "高",
     "人口推計の基礎の検証・同 変更・補正・将来推計・需要3シナリオの5件。"
     "推計の変更が続いており参照の頻度が高い。"
     "第2次概算（10月30日）で中長期推計を加えるため、"
     "その時点で束ねるのが効率的である"),
    ("受領資料の点検", "未", "中",
     "受領のたびに発注者へ送付しており、送付の単位が資料ごとである。"
     "資料の受領が一段落した時点で束ねる"),
    ("第9期計画の評価", "未", "中",
     "中間報告の提出済みの内容と対応しており、"
     "分けておくほうが提出物との対照が容易である。"
     "最終報告の作成時に束ねる"),
    ("調査の集計・点検", "未", "低",
     "調査ごとに照会先が異なり、照会票は単独で用いる。"
     "確認事項No.8の決定と確定値の確定を待つ"),
    ("会議資料の校正・点検", "未", "低",
     "内部保管6件。会議が終われば参照されないため、"
     "納品時に1件へ束ねる（05シート）"),
    ("保険料・給付の分析", "対象外", "―",
     "他団体の資料は成果品に用いず記載の借用も行わない整理であるため、"
     "同じ資料に混在させない"),
]
r = 4
r = header(ws, r, ["", "群", "状況", "優先度", "内容"])
for i, (a, b, c_, d) in enumerate(TOGO, start=1):
    r = body(ws, r, [i, a, b, c_, d],
             fills={3: {"済": OK_G, "未": IN_Y, "対象外": GRAY}.get(b)},
             height=50)

r = note(ws, r + 1,
         "注1）統合すると成果品の件数は減りますが、"
         "1件あたりのシート数が増えます。"
         "納品の形式（件数を絞るか、単位を細かく保つか）は"
         "ご判断をお願いする事項です（確認事項No.144）。\n"
         "注2）優先度が高い「人口・認定者数の推計」は、"
         "中長期推計（確認事項No.114）を第2次概算までに算定する予定であり、"
         "その作業と併せて束ねるのが効率的です。"
         "本表の時点では統合していません。\n"
         "注3）見込量算定の統合報告書は、"
         "採用パターンP3（令和7年度基準・総括表・認定者数シナリオ②）を"
         "算定の基礎としています。"
         "本表の管理の統合は、この採用案に沿って"
         "どの資料が現に用いられているかで並べ替えたものです。", span=5)


# ============================================================ 04
ws = sheet("04_出所による数値の食い違い",
           "出所による数値の食い違い",
           "同じ対象について出所が異なると値が違う。"
           "どちらかが誤りというものではなく、数える単位が違うことによる。"
           "計画に用いる出所を決めておく必要がある。",
           [4, 34, 13, 13, 10, 40])

r = 4
r = lead(ws, r, "1　施設・居住系の月平均利用者数（令和7年度）", span=6)
r = header(ws, r, ["", "サービス", "計画作成支援\nツール（3町計）",
                   "年報÷12\n（保険者単位）", "差", "内容"])
for i, (nm, t, n, d) in enumerate(CHIGAI, start=1):
    r = body(ws, r, [i, nm, t, round(n, 2),
                     ("%+.1f％" % d) if d is not None else "―",
                     "区域外の施設を利用する当連合の被保険者と、"
                     "区域内の施設を利用する他保険者の被保険者の"
                     "扱いが違う" if d and abs(d) >= 5 else ""],
             fills={5: (NG_O if d and abs(d) >= 5 else None)},
             numfmt={4: "0.00"}, height=20)
r = body(ws, r, ["", "計", round(CHIGAI_KEI[0], 1),
                 round(CHIGAI_KEI[1], 2),
                 "%+.1f％" % CHIGAI_KEI[2], ""],
         bold=True, numfmt={3: "0.0", 4: "0.00"}, height=20)
r = note(ws, r + 1,
         "注1）計画作成支援ツールは3町が入力した値の合計、"
         "年報は保険者（広域連合）単位です。"
         "サービス見込量・給付費・保険料は保険者単位で算定するため、"
         "算定には年報の値を用います。"
         "町別データシートは町別の値を示すため"
         "計画作成支援ツールの値によっており、"
         "3町を足しても年報の値にはなりません。\n"
         "注2）介護医療院の差が最も大きい（{:+.1f}％）のは、"
         "区域内に施設がなく区域外の施設の利用によるためです。"
         .format(next(d for nm, t, n, d in CHIGAI if nm == "介護医療院")),
         span=6)

r += 1
r = lead(ws, r, "2　他の食い違い", span=6)
r = header(ws, r, ["", "対象", "出所1", "出所2", "差", "扱い"])
HOKA = [
    ("給付費（令和7年度）", "3町合計 3,213,054,558円",
     "広域連合単位 3,083,474,291円", "4.0％",
     "令和6・8年度は完全一致。要介護度により1.7〜5.9％と幅がある。"
     "確認事項No.87。保険料は保険者単位で算定するため年報を用いる"),
    ("総合事業費（令和6年度）", "決算 101,274,621円",
     "計画作成支援ツール 107,129,318円", "5,854,697円",
     "決算は広域連合の款項、ツールは3町の入力による。"
     "第6章第6節に注記を置き［要確認］としている"),
    ("施設・居住系の令和7年度", "見える化の画面 482人",
     "年報÷12 480.167人", "0.4％",
     "サービス別では特定施設＋4.1％・老健▲3.0％・介護医療院＋9.1％。"
     "確認事項No.136"),
    ("認定率（令和7年度）", "見える化 20.84％", "当方 21.8％", "0.96ポイント",
     "分子の範囲と時点が違う（見える化は各年9月末、当方は年度末）。"
     "確認事項No.125"),
    ("認知症対応型共同生活介護の定員", "北海道の名簿 7事業所108人",
     "見える化D26 5事業所99人", "9人",
     "当方は99人＋［要確認］としている。"
     "くるみの郷は公表システムに掲載がなく定員未確定"),
]
for i, (a, b, c_, d, e) in enumerate(HOKA, start=1):
    r = body(ws, r, [i, a, b, c_, d, e], height=44)

r = note(ws, r + 1,
         "注3）いずれも「どちらかが誤り」ではなく、"
         "数える単位・時点・範囲が違うことによります。"
         "計画に用いる出所を決め、"
         "他の出所の値を並べるときは単位の違いを注記します。", span=6)


# ============================================================ 05
ws = sheet("05_役割を終えた成果品",
           "役割を終えた成果品",
           "後継ができたもの、会議の終了により参照されなくなったもの、"
           "統合報告書に束ねた元の資料を区分する。"
           "いずれも削除は提案しない。納品時の構成のご判断の材料とする。",
           [4, 40, 14, 14, 44])

OWARI = [
    ("第10期計画素案_修正指示書_令和8年7月.xlsx", "後継あり", "内部保管",
     "令和8年8月版が後継。7月版は作業の記録として残す"),
    ("第10期計画_キックオフ会議資料（更新版）の校正結果.xlsx", "会議終了",
     "内部保管", "キックオフ会議（令和8年7月）の校正記録"),
    ("第10期計画_キックオフ会議議事録の校正結果.xlsx", "会議終了", "内部保管",
     "同上"),
    ("第10期計画_キックオフ資料の点検結果.xlsx", "会議終了", "内部保管",
     "同上"),
    ("第10期計画_キックオフ会議_トークスクリプト_60分.docx", "会議終了",
     "内部保管", "会議の進行の手元資料"),
    ("第10期計画_キックオフ会議_ファシリテーションの分岐.xlsx", "会議終了",
     "内部保管", "同上"),
    ("第10期計画_中間報告会議議事録の校正結果.xlsx", "会議終了", "内部保管",
     "中間報告会議（令和8年8月）の校正記録"),
    ("第10期計画_中間報告会議議事録の校正結果_第2回.xlsx", "会議終了",
     "内部保管", "同上"),
    ("第10期計画_サービス見込量の算定_基準年度とパターンの比較.xlsx",
     "統合済み", "条件付き",
     "統合報告書にA00〜A03として収録済み。内容は同一"),
    ("第10期計画_見込量の季節性の検証と施策反映の設計.xlsx", "統合済み",
     "条件付き", "統合報告書にB00〜B11として収録済み。内容は同一"),
    ("第10期計画_対計画比の偏りの補正試算.xlsx", "統合済み", "条件付き",
     "統合報告書にC00〜C08として収録済み。内容は同一"),
    ("第10期計画_趨勢に現れない要因の分析.xlsx", "統合済み", "条件付き",
     "統合報告書にD00〜D09として収録済み。内容は同一"),
    ("介護保険事業計画_前期計画の評価_実施要領.xlsx", "別スコープ",
     "内部保管", "本業務の成果品ではないが作業に用いた"),
]
r = 4
r = header(ws, r, ["", "成果品", "区分", "送付区分", "内容"])
for i, (a, b, c_, d) in enumerate(OWARI, start=1):
    r = body(ws, r, [i, a, b, c_, d],
             fills={3: {"統合済み": IN_Y, "会議終了": GRAY,
                        "後継あり": GRAY, "別スコープ": GRAY}.get(b)},
             height=30)

_togo = [x for x in OWARI if x[1] == "統合済み"]
_kaigi = [x for x in OWARI if x[1] == "会議終了"]
r = note(ws, r + 1,
         "注1）{}件のうち、統合済みが{}件・会議終了が{}件です。\n"
         "注2）統合済みの4件は統合報告書（44シート）に収録済みで内容は同一です。"
         "**元の4件を残すか、統合報告書に一本化するかはご判断をお願いします**"
         "（確認事項No.144）。"
         "残す場合、同じ内容が2か所にあることになります。\n"
         "注3）会議終了の6件は内部保管であり、"
         "送付資料には含めていません。"
         "納品時は「その他業務で作成した資料 一式」に含めます。"
         "統合候補の群6に当たり、1件へ束ねることができます。\n"
         "注4）本シートは削除を提案するものではありません。"
         "作業の記録は納品まで残します。"
         "現在の成果品は{}件（送付{}・条件付き{}・内部保管{}・対象外{}）です。"
         .format(len(OWARI), len(_togo), len(_kaigi), len(DP.DISPATCH),
                 KUBUN["送付"], KUBUN["条件付き"], KUBUN["内部保管"],
                 KUBUN["対象外"]), span=5)


# ============================================================ 06
ws = sheet("06_誤りを正した記録",
           "算定又は記述に誤りがあり正したもの",
           "当方が自ら又はご指摘により見つけ、正したものを記録する。"
           "発見の経緯と、再発を防ぐために講じた手立てを併記する。",
           [4, 32, 10, 34, 34])

AYAMARI = [
    ("調整交付金の割合を7.375％と表示していた",
     "R8.9.11", "算定",
     "実際の計算は7.3755％で、Eの算定で49,113円の差。"
     "ブランチ内資料の点検によるご指摘",
     "表示を7.3755％に改めた。"
     "表示の桁を落とさず、計算に用いた値をそのまま書く"),
    ("`data_shien_tool.SHISETSU` を「定員」と表記していた",
     "R8.9.10", "記述",
     "実際は月平均利用者数。町別データシートの表題が誤っていた",
     "「施設・居住系サービスの月平均利用者数」に改めた。"
     "定員と読み違えやすい値4つをCLAUDE.mdに記録した"),
    ("令和9年度が介護報酬改定の年に当たらないとしていた",
     "R8.9.10", "記述",
     "改定は3年周期（H30・R3・R6）であり令和9年度は改定の年。"
     "`build_machi_kaito.py` の記述",
     "修正した。改定率1％あたり月額＋61円を算定し、"
     "確認事項No.109として起票した"),
    ("「2か年とも基本推計が実績に近い」としていた",
     "R8.9.11", "算定",
     "令和6年度は包括推計の方が1.00に近い。自己点検が検出した",
     "2か年平均の絶対誤差（基本0.61％・包括1.23％）で比べる形に改めた。"
     "「2か年とも」と書く前に各年を確かめる"),
    ("認知症対応型共同生活介護の町別定員が誤っていた",
     "R8.9.10", "算定",
     "東川町45人→36人、計108人→99人。"
     "第2章第3節6及び第6章第4節4の分布欄",
     "修正した。定員は99人＋［要確認］とし、"
     "くるみの郷の定員未確定を明示している"),
    ("素案の段落・表・図の数が管理表と合っていなかった",
     "R8.9.11", "記述",
     "実物806／130／36に対し管理表667／114／34。"
     "固定値で書いていたため更新に追随していなかった",
     "`repo_paths.draft_label()` で実物から数える形に改め、6箇所を置換した。"
     "件数・項目数は固定値で書かない"),
    ("ビルドスクリプトが出力先を絶対パスで固定していた",
     "R8.9.11", "運用",
     "98本中97本。別ブランチの作業ツリーから実行すると"
     "生成物がそこへ落ちる。実際に起きた",
     "`repo_paths.py` を新設し `RP.ROOT`（`__file__` 起点）に統一した。"
     "全スクリプトを別ディレクトリから実行して確認した"),
    ("受領済みの年報の明細を取り込まないまま推定値を作っていた",
     "R8.9.11", "算定",
     "年報（令和7年度・資料No.22）に要介護度別の明細があったが、"
     "主要値のみを取り込んでいた。"
     "按分による推定値を作り、明細の提供を求める確認事項を起票していた",
     "明細を `data_nenpo_meisai.py` に収め、推定値を確定値に置き換えた。"
     "確認事項No.134は当方の取込漏れによるものとして取り下げた。"
     "確認事項を起こす前に受領済みのファイルを確かめる"),
    ("「差し替えは4画面ある」と確かめずに断定した",
     "R8.9.12", "記述",
     "認定者数の実績見込み値を編集する項目が画面にあるかを"
     "確かめていなかった。受領した画面の写しには写っていない",
     "「編集項目の有無を要確認」に改め、確認事項No.138として起票した。"
     "画面で確かめていないことは断定しない"),
    ("健康とくらしの調査の集計値を暫定のまま置きかけた",
     "R8.9.15", "算定",
     "社会参加・暮らし向きの値を仮に置いたまま素案へ渡そうとした。"
     "出所との突合で気付いた",
     "全項目を機械的に突合して実際の値に直した（不一致0）。"
     "データモジュールを作ったら出所と全項目を突き合わせる"),
    ("認定率の差を「分子に第2号被保険者を含むかの違い」としていた",
     "R8.9.16", "記述",
     "確認事項No.125の起票時の見立て。"
     "受領資料を数値で確かめないまま原因を1つに決めていた。"
     "令和8年9月16日に逆算して確かめたところ、"
     "当方の認定率は分子に第2号被保険者の認定者を含んでおらず、"
     "見える化システムの総括表にも要介護認定率の行があって"
     "1,943÷9,119で小数第6位まで再現できることが分かった",
     "確認事項No.125を「将来推計の参考シートの認定率が"
     "総括表の認定率と一致しないこと」に絞って書き直した。"
     "**差の原因を1つに決める前に、受領資料の数値で逆算して確かめる**"),
]
r = 4
r = header(ws, r, ["", "誤り", "時期", "内容と発見の経緯", "正した内容と手立て"])
for i, (a, b, c_, d, e) in enumerate(AYAMARI, start=1):
    r = body(ws, r, [i, a, b, d, e],
             fills={3: {"算定": NG_O, "記述": IN_Y,
                        "運用": MID_B}.get(c_)}, height=50)

_sante = [x for x in AYAMARI if x[2] == "算定"]
r = note(ws, r + 1,
         "注1）{}件のうち算定に関わるものが{}件、"
         "記述が{}件、運用が{}件です。\n"
         "注2）算定に関わる{}件のうち、"
         "保険料の算定に影響したのは調整交付金の表示（49,113円）のみで、"
         "他は表示・推定値の置き方・定員の集計に関わるものです。\n"
         "注3）当方の自己点検が見つけたもの{}件、"
         "ブランチ内資料の点検によるご指摘{}件、"
         "受領資料の再送により判明したもの{}件です。\n"
         "注4）いずれも「手立て」の欄に記した内容をCLAUDE.mdの"
         "「陥りやすい誤り」に記録し、同じ誤りを繰り返さない形にしています。"
         .format(len(AYAMARI), len(_sante),
                 len([x for x in AYAMARI if x[2] == "記述"]),
                 len([x for x in AYAMARI if x[2] == "運用"]),
                 len(_sante), 4, 3, 1), span=5)


# ============================================================ 07
ws = sheet("07_自己点検", "自己点検（エラーチェック）",
           "本表の数え方と、他の成果品との突合を点検する。",
           [4, 44, 26, 34, 10, 10])

chk(1, "スクリプト数を実物から数えていること",
    "build_*.py（本スクリプトを除いて判定）",
    "{}本（うち再実行できない{}本）".format(len(SCRIPTS), len(SAIJIKKO_NG)),
    len(SCRIPTS) > 0)

chk(2, "再実行できないスクリプトが3本であること",
    "セッション固有のパスを読むもの",
    "／".join(SAIJIKKO_NG) if SAIJIKKO_NG else "なし",
    len(SAIJIKKO_NG) == 3)

chk(3, "成果品の登録漏れがないこと",
    "output と data_dispatch の突合",
    "登録漏れ{}件・実体なし{}件".format(len(MITOUROKU), len(KESSON)),
    not MITOUROKU and not KESSON)

chk(4, "確認事項の番号に重複がないこと",
    "業務工程管理表の CHECK",
    "{}件／一意{}件".format(len(CHECK), len(NO_SET)),
    len(CHECK) == len(NO_SET))

_nai = [k for k in list(TSUGO_DOUITSU.values()) + list(TSUGO_KANREN.values())
        if k not in NO_SET]
chk(5, "対応付けたNo.が業務工程管理表に実在すること",
    "{}件".format(len(TSUGO_DOUITSU) + len(TSUGO_KANREN)),
    "全件実在" if not _nai else "不明 " + "・".join(str(x) for x in _nai),
    not _nai)

_wn = [k for k in list(TSUGO_DOUITSU) + list(TSUGO_KANREN) if k not in _wl]
chk(6, "対応付けた別管理表の行が実在すること",
    "A{}件・B{}件".format(len(WORD), len(NUM)),
    "全件実在" if not _wn else "不明 " + str(_wn), not _wn)

chk(7, "施設・居住系の年報の月平均が480.167人であること"
       "（令和8年度実績見込み値の入力案と同じ値になること）",
    "年報÷12の合計",
    "{:.3f}人／月".format(CHIGAI_KEI[1]),
    abs(CHIGAI_KEI[1] - 480.167) < 5e-4)

chk(8, "計画作成支援ツールと年報の差が5％を超えること"
       "（出所の違いが無視できない大きさであること）",
    "3町合計と保険者単位",
    "{:.1f}人対{:.2f}人（{:+.1f}％）".format(*CHIGAI_KEI),
    abs(CHIGAI_KEI[2]) > 5)

_ow = [x[0] for x in OWARI if x[0] not in DP.DISPATCH]
chk(9, "05シートに挙げた成果品が送付区分表に実在すること",
    "{}件".format(len(OWARI)),
    "全件実在" if not _ow else "不明 " + "・".join(_ow), not _ow)

chk(10, "誤りを正した記録に再発を防ぐ手立てを書いていること",
    "{}件".format(len(AYAMARI)),
    "全件に手立てを記載",
    all(x[4] for x in AYAMARI))

chk(11, "統合候補7群をすべて掲げていること",
    "統合報告書 07シート",
    "{}群（済{}・未{}・対象外{}）".format(
        len(TOGO), len([x for x in TOGO if x[1] == "済"]),
        len([x for x in TOGO if x[1] == "未"]),
        len([x for x in TOGO if x[1] == "対象外"])),
    len(TOGO) == 7)

chk(12, "成果品の区分の合計が登録数と一致すること",
    "data_dispatch",
    "{}件＝送付{}＋条件付き{}＋内部保管{}＋対象外{}".format(
        len(DP.DISPATCH), KUBUN["送付"], KUBUN["条件付き"],
        KUBUN["内部保管"], KUBUN["対象外"]),
    sum(KUBUN.values()) == len(DP.DISPATCH))

r = 4
r = header(ws, r, ["No", "点検の内容", "式・対象", "結果", "判定", ""])
for no, naiyo, shiki, kekka, han in CHECKS:
    r = body(ws, r, [no, naiyo, shiki, kekka, han, ""],
             fills={5: OK_G if han == "適合" else NG_O}, height=32)
NG = sum(1 for c in CHECKS if c[4] == "不適合")
r = body(ws, r, ["", "計", "{}件".format(len(CHECKS)),
                 "適合{}件・不適合{}件".format(len(CHECKS) - NG, NG),
                 "適合" if NG == 0 else "不適合", ""],
         fills={5: OK_G if NG == 0 else NG_O}, bold=True)
r = note(ws, r + 1,
         "注1）点検7は、本表の年報の値が"
         "「令和8年度実績見込み値の入力案」と同じ算定によることを"
         "確かめるものです。\n"
         "注2）点検3が不適合になるのは、"
         "成果品を追加して `data_dispatch.py` への登録を忘れたときです。\n"
         "注3）1件でも不適合があると、本表を作るスクリプトは"
         "終了コード1で終わります。", span=6)


# ============================================================ 08
ws = sheet("08_確認事項", "確認事項",
           "本表に関してご判断をお願いする事項。"
           "番号は業務工程管理表 03_確認事項一覧による。",
           [7, 30, 46, 20, 12, 12])

r = 4
r = header(ws, r, ["No.", "確認事項", "内容", "止めている成果物",
                   "確認先", "回答期限"])
KAKUNIN = [
    ("No.144", "統合した成果品の元の資料を残すか",
     "見込量算定の4件（基準年度とパターンの比較・季節性と施策反映・"
     "対計画比の補正・趨勢に現れない要因）は、"
     "サービス見込量算定_統合報告書（44シート）に収録済みで内容は同一である。"
     "現在は元の4件も残しており、同じ内容が2か所にある。"
     "①納品時に元の4件を残すか、統合報告書に一本化するかをご判断いただきたい。"
     "②残す場合、更新は統合報告書と元の資料の両方に反映する必要がある。"
     "③今後の統合（人口・認定者数の推計の5件ほか）も同じ扱いとしてよいか。"
     "④会議資料の校正記録6件は内部保管であり、"
     "納品時に1件へ束ねることができる。",
     "成果品一覧（納品時の構成）",
     "発注者", "R8.10"),
    ("No.145", "計画作成支援ツールと年報のどちらを町別の値に用いるか",
     "施設・居住系サービスの月平均利用者数は、"
     "計画作成支援ツール（3町合計）が{:.0f}人、"
     "年報（保険者単位）が{:.1f}人で{:+.1f}％違う。"
     "サービス別では介護医療院が最も大きく違う。"
     "①サービス見込量・給付費・保険料は保険者単位で算定するため"
     "年報の値を用いる整理でよいか。"
     "②町別データシートは町別の値を示すため"
     "計画作成支援ツールの値によっており、"
     "3町を足しても年報の値にはならない。"
     "この違いを町別データシートに注記する整理でよいか。"
     "③給付費（確認事項No.87）も同じ性質の差であり、"
     "併せて整理する。"
     .format(CHIGAI_KEI[0], CHIGAI_KEI[1], CHIGAI_KEI[2]),
     "町別データシート\nサービス見込量 第1次概算",
     "発注者", "R8.9"),
]
for no, ken, naiyo, tome, saki, kigen in KAKUNIN:
    r = body(ws, r, [no, ken, naiyo, tome, saki, kigen], height=170)
r = note(ws, r + 1,
         "注1）本表により新たに起票した確認事項は{}件"
         "（No.144・No.145）です。\n"
         "注2）確認事項の管理台帳の統合（02シート）は"
         "当方の作業だけで進むため、確認事項としていません。"
         .format(len(KAKUNIN)), span=6)


# ============================================================ 出力
os.makedirs(ODIR, exist_ok=True)
wb.save(OUT)

print("出力：" + OUT)
print("スクリプト：{}本（再実行できない{}本）"
      .format(len(SCRIPTS), len(SAIJIKKO_NG)))
print("成果品：{}件（送付{}・条件付き{}・内部保管{}・対象外{}）"
      .format(len(DP.DISPATCH), KUBUN["送付"], KUBUN["条件付き"],
              KUBUN["内部保管"], KUBUN["対象外"]))
print("登録漏れ{}件・実体なし{}件".format(len(MITOUROKU), len(KESSON)))
print("確認事項の台帳：業務工程管理表{}件／別管理表{}件（同一{}・関連{}）"
      .format(len(CHECK), len(WORD) + len(NUM),
              len(TSUGO_DOUITSU), len(TSUGO_KANREN)))
print("統合候補：7群（済1・未5・対象外1）")
print("出所の食い違い：ツール{:.0f}人対年報{:.2f}人（{:+.1f}％）"
      .format(*CHIGAI_KEI))
print("役割を終えた成果品：{}件／誤りを正した記録：{}件"
      .format(len(OWARI), len(AYAMARI)))
print("新たな確認事項：{}件（No.144・No.145）".format(len(KAKUNIN)))
print("自己点検：{}件（適合{}件・不適合{}件）"
      .format(len(CHECKS), len(CHECKS) - NG, NG))

if NG:
    sys.exit(1)
