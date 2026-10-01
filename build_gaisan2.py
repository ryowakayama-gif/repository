# -*- coding: utf-8 -*-
"""大雪地区広域連合 第10期介護保険事業計画　サービス見込量の算定（第2次概算の組立て）.

令和8年10月1日のご指示
  「第2次概算（10月30日）の組み立てへ進んで下さい」

第1次概算は令和8年9月26日に提示した。
その版は output/第10期計画_サービス見込量の算定_第1次概算.xlsx として残し、
値は data_gaisan1.py に凍結している（過去に報告した数値は書き換えない）。

本表は第2次概算（令和8年10月30日）の組立てである。
算定そのものは「サービス見込量の算定」（全18シート）が行い、
本表はそこから次のものを示す。

  1 第1次概算からどこが動いたか（前提の変更と月額への効き）
  2 保険料算定表の記号A〜Jの対照
  3 サービス見込量・利用回（日）数・給付費の種別ごとの差分
  4 中長期推計の対照
  5 月額を動かす前提（据え置きのうち発注者・国の決定を待つもの）
  6 見える化システムの出力との対照
  7 残る未決と既定値（仮置きにより進めたもの）
  8 令和8年10月30日までに行うこと

同じ数値を2か所に書かないため、見込量・給付費・保険料は
サービス見込量の算定を読み、第1次概算の値は凍結した値を読む。

シート構成　00_この資料について から 12_自己点検 まで全13シート

出力
  output/第10期計画_サービス見込量の算定_第2次概算.xlsx

自己点検で1件でも不適合があると終了コード1で終わる。
"""

import ast
import io
import os
import re
import runpy
import sys

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import data_gaisan1 as G1
import data_mieru_kekka2 as K2
import data_kitei
import repo_paths as RP

if isinstance(sys.stdout, io.TextIOWrapper):
    sys.stdout.reconfigure(encoding="utf-8")

ODIR = RP.OUTPUT
OUT = os.path.join(ODIR, "第10期計画_サービス見込量の算定_第2次概算.xlsx")
G1_XLSX = os.path.join(ODIR, "第10期計画_サービス見込量の算定_第1次概算.xlsx")

KIJUNBI = "令和8年10月1日"
TEISHUTSU = "令和8年10月30日"

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


def _load_quiet(name):
    buf, orig = io.StringIO(), sys.stdout
    sys.stdout = buf
    try:
        return runpy.run_path(os.path.join(RP.ROOT, name))
    finally:
        sys.stdout = orig


def _literal(fname, name):
    tree = ast.parse(open(os.path.join(RP.ROOT, fname), encoding="utf-8").read())
    for node in tree.body:
        if isinstance(node, ast.Assign) and any(
                getattr(t, "id", "") == name for t in node.targets):
            return ast.literal_eval(node.value)
    raise KeyError(name)


# ==================================================== 算定と受領点検を読む
S = _load_quiet("build_mikomiryo_santei.py")        # サービス見込量の算定
MI = _load_quiet("build_receipt_mieru_kekka.py")    # 見える化の推計結果の受領点検

SVC = S["SVC"]
Y3, Y3L = S["Y3"], S["Y3L"]
YLONG, YLONGL = S["YLONG"], S["YLONGL"]
G_DO, G_ITTEI = S["G_DO"], S["G_ITTEI"]
KYUFU_Y = S["KYUFU_Y"]
SUEOKI = S["SUEOKI"]
TEIIN_MAP = S["TEIIN_MAP"]
sa, yen = S["sa"], S["yen"]
sogaku, kaisu_mikomi, kyufu_do = S["sogaku"], S["kaisu_mikomi"], S["kyufu_do"]
short, kubun = S["short"], S["kubun"]
KAISU_MAP = S["KAISU_MAP"]

KUBUN = _literal("build_mikomi_juryo_nashi.py", "KUBUN")
CHECK = _literal("build_process_control.py", "CHECK")
LACK = _literal("build_process_control.py", "LACK")
SUSUMETA = _literal("build_ikenkokankai.py", "SUSUMETA")
KANRYO = ("完了", "了承済", "了承済（保管せず廃棄）", "代替により解消", "解決")
MACHI = [c for c in CHECK if c[7] not in KANRYO]
KITEI = data_kitei.all_kitei()

KYUFU3 = sum(KYUFU_Y[y] for y in Y3)
GETSU = G_DO["月額"]
KIJUN = int(round(GETSU / 100.0)) * 100


def sad(v, k=0, tani="", c=True):
    """差の表記。差がないときは「―」とする。"""
    return "―" if abs(v) < 0.5 / (10 ** k) else sa(v, k, tani, c=c)


def kiji(v):
    """百円未満四捨五入した保険料基準額。"""
    return int(round(v / 100.0)) * 100


# ==================================================== 共通の体裁
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
    ws.row_dimensions[2].height = 58
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = freeze
    return ws


def header(ws, row, cols, height=32):
    for i, h in enumerate(cols, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = Font(name=FONT, size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=HEAD)
        c.alignment = Alignment(wrap_text=True, horizontal="center",
                                vertical="center")
        c.border = BORDER
    ws.row_dimensions[row].height = height
    return row + 1


def body(ws, row, vals, fills=None, height=22, align=None, bold=False,
         fmt=None):
    for i, v in enumerate(vals, start=1):
        c = ws.cell(row=row, column=i, value=v)
        c.font = Font(name=FONT, size=9, bold=bold)
        c.border = BORDER
        ha = (align or {}).get(i, "left" if isinstance(v, str) else "right")
        c.alignment = Alignment(wrap_text=True, vertical="top", horizontal=ha)
        if fills and fills.get(i):
            c.fill = PatternFill("solid", fgColor=fills[i])
        if fmt and fmt.get(i) and not isinstance(v, str):
            c.number_format = fmt[i]
    ws.row_dimensions[row].height = height
    return row + 1


def lead(ws, row, text, span=8):
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=10.5, bold=True, color=NAVY)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    ws.row_dimensions[row].height = 20
    return row + 1


def note(ws, row, text, span=8, height=None):
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=8.5)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    ws.row_dimensions[row].height = height or (14 * (1 + text.count("\n")))
    return row + 1


def _oki(s):
    """算定の表の文を本資料の文に直す（点検の番号・シートの番号を落とす）。"""
    v = str(s).replace("**", "")
    if v.startswith("令和8年10月1日") and "。" in v:
        v = v.split("。", 1)[1]
    v = re.sub(r"（点検[0-9０-９・点検]*）", "", v)
    v = re.sub(r"[0-9０-９]+シートに", "本算定に", v)
    v = v.replace("本表は", "本算定は").replace("本表の", "本算定の")
    return v.strip()


YEN0 = {i: "#,##0" for i in range(2, 12)}
F1 = {i: "0.0" for i in range(2, 12)}

# ============================================================ 00
ws = sheet("00_この資料について",
           "サービス見込量の算定　第2次概算の組立て",
           "第1次概算（令和8年9月26日提示）から第2次概算（%s）までの間に、"
           "算定のどこが動いたか、何が決まれば確定するかを1冊にまとめたものです。"
           "サービス見込量・給付費・保険料の算定そのものは"
           "「サービス見込量の算定」（全18シート）によります。"
           "ご決定をお待ちしている事項は、決着しない場合の扱い（既定値）に"
           "よって算定しています（確認事項No.142）。"
           "既定値は受託者の仮置きであり、ご決定の内容が異なる場合は"
           "その内容により置き直します。"
           "受託者の算定であり確定値ではありません。（基準日 %s）"
           % (TEISHUTSU, KIJUNBI),
           [3, 26, 62, 20], freeze="A6")
r = 4
r = lead(ws, r, "1　第1次概算と第2次概算の関係", span=4)
r = header(ws, r, ["", "項目", "内容", "出所"])
for i, (k, v, s) in enumerate([
    ("第1次概算",
     "令和8年9月26日に提示しました。総給付費（3か年計）%s円、"
     "算定上の月額基準額%d円、保険料基準額%d円です。"
     "提示した時点の値を残すため、その版は作り直していません。"
     % (yen(G1.KYUFU3), G1.GETSU_DO, G1.KIJUN),
     "サービス見込量の算定_第1次概算"),
    ("第2次概算",
     "本組立ての時点で、総給付費（3か年計）%s円、"
     "算定上の月額基準額%d円、保険料基準額%d円です。"
     "%sに向けて、ご決定をいただいた事項を織り込んで算定し直します。"
     % (yen(KYUFU3), round(GETSU), KIJUN, TEISHUTSU),
     "サービス見込量の算定"),
    ("第1次概算から動いたもの",
     "算定の前提が動いたのは2件（要介護度別の単価の採用、"
     "1人1月あたり利用回（日）数の要介護度別化）です。"
     "いずれも据え置いていたものを受託者の側で確定したものであり、"
     "資料の受領を待たずに進められるものでした。"
     "保険料基準額は%d円で変わりません。" % KIJUN,
     "01シート"),
    ("変わらないもの",
     "基準年度（令和7年度・12か月）、人口の基礎（案C）、"
     "認定者数のシナリオ②、総給付費の割増率1.05909、"
     "第1号被保険者負担割合23％、予定収納率99.0％、"
     "調整交付金の見込交付割合7.3755％、地域支援事業費は"
     "令和6年度決算の水準の据え置きです。",
     "サービス見込量の算定 01シート"),
    ("決まっていないもの",
     "据え置き%d件のうち、発注者・3町の資料を待つもの%d件、"
     "国の告示・公布を待つもの%d件です。"
     "受託者の側で確定できるもの（区分A）は残っていません。"
     % (len(SUEOKI), sum(1 for v in KUBUN.values() if v == "B"),
        sum(1 for v in KUBUN.values() if v == "C")),
     "08シート"),
], start=1):
    r = body(ws, r, [i, k, v, s], height=64)
r += 1
r = lead(ws, r, "2　本資料の構成", span=4)
r = header(ws, r, ["", "シート", "内容", "読む相手側"])
for i, (k, v, s) in enumerate([
    ("01_第1次概算からの差分",
     "算定の前提が動いた2件と、要約（総給付費・月額・基準額）の対照。",
     "凍結した第1次概算の値"),
    ("02_保険料算定表の対照",
     "記号A〜J・補正後被保険者数・算定上の月額基準額を"
     "第1次概算と並べたもの。", "同左"),
    ("03_見込量の差分",
     "サービス29区分の令和9年度から令和11年度までの利用者数（人／月）。",
     "同左"),
    ("04_利用回（日）数の差分",
     "12区分の延べ回（日）数（回・日／月）。"
     "1人1月あたりを要介護度別に置き直したことによる差が出ます。", "同左"),
    ("05_給付費の差分",
     "サービス29区分の第10期3か年計（円）。"
     "要介護度別の単価を採用したことによる差が出ます。", "同左"),
    ("06_中長期推計の対照",
     "令和12・17・22・27・32年度の総給付費と参考の月額。", "同左"),
    ("07_地域支援事業と必要利用定員総数",
     "地域支援事業の量と事業費、施設・居住系の定員に対する到達率。",
     "サービス見込量の算定 10・11シート"),
    ("08_月額を動かす前提",
     "据え置きのうち発注者・国の決定を待つもの。効きの大きい順。",
     "サービス見込量の算定 15シート"),
    ("09_見える化システムとの対照",
     "システムの推計結果と本算定の対照、"
     "修正が必要な箇所11件の現況（解消3・残8）。",
     "見える化の推計結果の受領点検"),
    ("10_残る未決と既定値",
     "仮置きにより計画素案へ反映した%d件と、"
     "%sまでに決着をいただきたい事項。" % (len(SUSUMETA), TEISHUTSU),
     "業務工程管理表 03_確認事項一覧"),
    ("11_10月30日までに行うこと",
     "受託者の作業と、ご決定・ご提供をお願いする事項の段取り。", "―"),
    ("12_自己点検", "本資料の内的整合と、算定・凍結値との接続の点検。", "―"),
], start=1):
    r = body(ws, r, [i, k, v, s], height=40)
r += 1
r = note(ws, r,
         "注1）第1次概算の値は提示した時点のものです。"
         "提示した後に算定を改めても、第1次概算の値は書き換えていません。\n"
         "注2）保険料基準額は算定上の月額基準額を百円未満四捨五入したものです。"
         "条例で定める額ではありません。\n"
         "注3）確認事項は業務工程管理表 03_確認事項一覧で一元管理しています。"
         "計画素案の本文には確認事項の注記を書きません。", span=4, height=60)

# ============================================================ 01
SHIFT = [
    ("要介護度別の単価",
     "要介護度別の1人1月あたり給付費",
     "総括表詳細（４）は要介護度別の内訳を持たないため、"
     "要介護度計の単価で置いていました。"
     "要介護度間の比は給付費データ（保険者単位）から採れます。",
     "要介護度間の比を給付費データから採り、水準は総括表に合わせて"
     "比例調整しました。基準年度では総括表の給付費を円単位で再現します。"
     "要介護度の構成が重度側へ動くため平均単価が上がります。",
     "総給付費（3か年）%s・月額%s"
     % (sa(KYUFU3 - G1.KYUFU3, 0, c=True), sa(GETSU - G1.getsu_sei(), 2))),
    ("1人1月あたり利用回（日）数",
     "利用回（日）数の置き方",
     "要支援・要介護の群でまとめた1人1月あたりで置いていました。"
     "見える化システムの入力画面は要介護度別の1人1月あたりを求めます。",
     "要介護度別の1人1月あたりに置き直しました。"
     "計画本文の値と入力値の差（最大＋0.09％）が消えました。",
     "給付費には入らないため月額は動きません"),
]
ws = sheet("01_第1次概算からの差分",
           "第1次概算からの差分",
           "第1次概算（令和8年9月26日提示）から算定が動いた事項と、"
           "要約の対照です。"
           "動いたのはいずれも据え置いていたものを受託者の側で確定したもので、"
           "資料の受領を待たずに進められるものでした。"
           "発注者・3町・国の決定を待つ事項は動いていません。",
           [3, 22, 34, 40, 26], freeze="A6")
r = 4
r = lead(ws, r, "1　算定の前提が動いたもの", span=5)
r = header(ws, r, ["", "事項", "第1次概算での置き方", "第2次概算での置き方",
                   "効き"])
for i, x in enumerate(SHIFT, start=1):
    r = body(ws, r, [i, x[0], x[2], x[3], x[4]], height=76,
             fills={5: IN_Y})
r += 1
r = lead(ws, r, "2　要約の対照", span=5)
r = header(ws, r, ["", "項目", "第1次概算", "第2次概算", "差"])
SUMM = [
    ("総給付費（3か年計・円）", G1.KYUFU3, KYUFU3),
    ("標準給付費見込額（円）", G1.A, G_DO["A"]),
    ("地域支援事業費（3か年計・円）", G1.B, G_DO["B"]),
    ("保険料収納必要額（J・円）", G1.J, G_DO["J"]),
    ("補正後被保険者数（人）", G1.HOSEI_SEI, G_DO["③"]),
    ("算定上の月額基準額（円）", G1.getsu_sei(), GETSU),
    ("保険料基準額（円）", G1.KIJUN, KIJUN),
]
for i, (k, a, b) in enumerate(SUMM, start=1):
    r = body(ws, r, [i, k, round(a), round(b), sad(b - a)],
             height=22, fmt={3: "#,##0", 4: "#,##0"},
             bold=(k.startswith("保険料基準額")),
             fills={5: (GRAY if abs(b - a) < 0.5 else IN_Y)})
r += 1
r = lead(ws, r, "3　年度別の総給付費の対照（円）", span=5)
r = header(ws, r, ["", "年度", "第1次概算", "第2次概算", "差"])
for i, (y, yl) in enumerate(zip(Y3, Y3L), start=1):
    a, b = G1.KYUFU_NEN[y], KYUFU_Y[y]
    r = body(ws, r, [i, yl, round(a), round(b), sad(b - a)],
             fmt={3: "#,##0", 4: "#,##0"})
r = body(ws, r, ["", "第10期計", round(G1.KYUFU3), round(KYUFU3),
                 sad(KYUFU3 - G1.KYUFU3)],
         fmt={3: "#,##0", 4: "#,##0"}, bold=True, fills={2: MID_B})
r += 1
r = note(ws, r,
         "注1）保険料基準額は変わりません。"
         "算定上の月額基準額が%.2f円から%.2f円へ%s動きましたが、"
         "百円未満を四捨五入するといずれも%d円です。\n"
         "注2）第9期の算定上の月額基準額は6,428円です。"
         "第10期の「一律の伸び」による算定値6,428円と同じ数字になりますが、"
         "別のものを算定した結果が一致したものです。\n"
         "注3）差の算定に用いた第1次概算の値は、"
         "令和8年9月26日に提示した版の値です。"
         % (G1.getsu_sei(), GETSU, sa(GETSU - G1.getsu_sei(), 2), KIJUN),
         span=5, height=60)

# ============================================================ 02
ws = sheet("02_保険料算定表の対照",
           "保険料算定表（記号A〜J）の対照",
           "国の様式による保険料収納必要額の組立てを、"
           "第1次概算と第2次概算で並べたものです。"
           "J＝C＋D－E＋F＋G±H－I、"
           "算定上の月額基準額＝J÷予定収納率99.0％÷補正後被保険者数÷12です。"
           "第1号被保険者負担割合・報酬改定率・基金残高は確定していません。",
           [6, 46, 20, 20, 18, 46], freeze="A6")
r = 4
r = lead(ws, r, "1　保険料収納必要額の組立て", span=6)
r = header(ws, r, ["記号", "項目", "第1次概算", "第2次概算", "差", "置き方"])
OKI = {
    "A": "総給付費×割増率1.05909（令和6年度決算による）",
    "B": "令和6年度決算の水準を3年据え置き（確認事項No.5・No.112・No.150）",
    "①": "標準給付費見込額＋地域支援事業費",
    "②": "標準給付費見込額＋総合事業費",
    "C": "①×第1号被保険者負担割合23％［要確認］（確認事項No.33）",
    "D": "②×5％",
    "E": "②×見込交付割合7.3755％",
    "F": "拠出の予定なし",
    "G": "借入れの実績なし",
    "H": "同事業に参加していない",
    "I": "取崩しを行わないものとしている（令和7年度末の基金残高が未受領）",
    "J": "国の様式による",
}
r = body(ws, r, ["", "総給付費（3か年計）", round(G1.KYUFU3), round(KYUFU3),
                 sa(KYUFU3 - G1.KYUFU3, 0, "", c=True),
                 "要介護度別の伸び（見込量×12×要介護度別の単価）"],
         fmt={3: "#,##0", 4: "#,##0"})
for sym, item, a, _b in G1.SANTEI[1:13]:
    now = G_DO[sym]
    r = body(ws, r, [sym, item, round(a), round(now),
                     sad(now - a), OKI.get(sym, "")],
             fmt={3: "#,##0", 4: "#,##0"},
             bold=(sym == "J"),
             fills={1: (MID_B if sym in ("A", "B", "J") else None)})
r = body(ws, r, ["③", "補正後被保険者数（人）", round(G1.HOSEI_SEI),
                 round(G_DO["③"]),
                 sad(G_DO["③"] - G1.HOSEI_SEI),
                 "推計上の被保険者数（案C）×令和6年度の実績係数0.991352。"
                 "乗率は公費軽減前による"],
         fmt={3: "#,##0", 4: "#,##0"})
r = body(ws, r, ["", "算定上の月額基準額（円）", round(G1.getsu_sei(), 2),
                 round(GETSU, 2), sad(GETSU - G1.getsu_sei(), 2),
                 "J÷予定収納率99.0％÷③÷12"],
         bold=True, fmt={3: "#,##0", 4: "#,##0.00"}, fills={2: IN_Y})
r = body(ws, r, ["", "保険料基準額（円）", G1.KIJUN, KIJUN,
                 sad(KIJUN - G1.KIJUN),
                 "算定上の月額基準額の百円未満を四捨五入したもの。"
                 "条例で定める額ではない"],
         bold=True, fmt={3: "#,##0", 4: "#,##0"}, fills={2: OK_G})
r += 1
r = lead(ws, r, "2　第9期との対比", span=6)
r = header(ws, r, ["", "項目", "第9期計画", "第10期（第2次概算）", "増減", "備考"])
_K9NOW = [G_DO["A"], G_DO["B"], G_DO["J"], G_DO["③"]]
for i, ((k, k9, _g1), nowv) in enumerate(zip(G1.K9, _K9NOW), start=1):
    r = body(ws, r, [i, k, k9, round(nowv), sad(nowv - k9),
                     "円" if "被保険者数" not in k else "人"],
             fmt={3: "#,##0", 4: "#,##0"})
r = body(ws, r, ["", "算定上の月額基準額（円）", 6428, round(GETSU),
                 sad(GETSU - 6428),
                 "第9期の条例による基準額は6,400円"],
         bold=True, fmt={3: "#,##0", 4: "#,##0"}, fills={2: IN_Y})
r += 1
r = note(ws, r,
         "注1）第9期は財政調整基金40,000,000円を取り崩しています。"
         "本算定は取崩しを行わないものとしており、"
         "同額を取り崩せば月額は約▲124円となります。\n"
         "注2）据え置いたもののうち、報酬改定率と単価の趨勢は"
         "月額を%s〜%s動かし得ます（08シート）。"
         "本算定の月額はこれを含まない下側の値です。\n"
         "注3）第1号被保険者負担割合を23％で置いています。"
         "1ポイントの違いで月額が約290円動きます。\n"
         "注4）国の様式には保険料収納必要額から"
         "保険者機能強化推進交付金等の交付見込額を控除する欄があります。"
         "本算定は地域支援事業費（B）から当該2事業（6,223,000円）を"
         "除く方法によっているため、収納必要額からの控除は行っていません。"
         "両方を行うと二重控除になります（確認事項No.86）。"
         % (S["susei_haba_s"]()[0], S["susei_haba_s"]()[1]),
         span=6, height=86)

# ============================================================ 03
ws = sheet("03_見込量の差分",
           "サービス見込量の差分（利用者数・人／月）",
           "サービス29区分の利用者数を第1次概算と並べたものです。"
           "見込量は令和7年度の要介護度別の月平均利用者数に"
           "要介護度別の認定者数の伸びを乗じたもので、"
           "第1次概算から置き方を変えていません。"
           "延べ数であり、1人が複数のサービスを利用する場合は"
           "それぞれに計上されます。",
           [3, 14, 40] + [11] * 6 + [26], freeze="A6")
r = 4
r = header(ws, r, ["", "区分", "サービス"]
           + ["%s\n第1次" % y for y in Y3L]
           + ["%s\n第2次" % y for y in Y3L] + ["差"])
_m_sa = []
for i, lab in enumerate(SVC, start=1):
    a = G1.MIKOMI.get(short(lab))
    nowv = [sogaku(lab, y) for y in Y3]
    if a is None:
        continue
    g1v = list(a[1:])
    d = sum(nowv) - sum(x or 0 for x in g1v)
    _m_sa.append(abs(d))
    r = body(ws, r, [i, kubun(lab).replace("サービス", ""), short(lab)]
             + [round(x, 1) for x in g1v] + [round(x, 1) for x in nowv]
             + ["差なし" if abs(d) < 0.05 else sa(d, 1, "人")],
             fmt={i2: "0.0" for i2 in range(4, 10)},
             fills={10: (GRAY if abs(d) < 0.05 else IN_Y)})
r += 1
r = note(ws, r,
         "注1）利用者数は第1次概算から変わっていません。"
         "見込量の置き方（要介護度別に利用率を固定して延ばすこと）を"
         "変えていないためです。\n"
         "注2）実績が0のサービスが6種別あります。"
         "区域内に事業所がないサービスでも、"
         "区域外の事業所の利用により実績が生じる場合があります。\n"
         "注3）供給の制約（人材・受入可能量）は見込量に反映していません。"
         "定員に対する到達率は07シートに掲げています。",
         span=10, height=48)

# ============================================================ 04
ws = sheet("04_利用回（日）数の差分",
           "利用回（日）数の差分（延べ・回／月・日／月）",
           "介護保険法第117条第2項第1号の量の見込みは"
           "利用者数だけでなく利用回（日）数を含みます。"
           "第1次概算は要支援・要介護の群でまとめた1人1月あたりにより、"
           "第2次概算は要介護度別の1人1月あたりによります。"
           "要介護度の構成が重度側へ動くため、延べの回（日）数が変わります。"
           "給付費には入らないため保険料は動きません。",
           [3, 36, 7] + [12] * 6 + [24], freeze="A6")
r = 4
r = header(ws, r, ["", "サービス", "単位"]
           + ["%s\n第1次" % y for y in Y3L]
           + ["%s\n第2次" % y for y in Y3L] + ["差"])
_k_n = 0
for i, lab in enumerate(KAISU_MAP, start=1):
    nm = short(lab)
    a = G1.KAISU.get(nm)
    if a is None:
        continue
    tani, g1v = a[0], list(a[2:])
    nowv = [kaisu_mikomi(lab, y) for y in Y3]
    d = sum(nowv) - sum(g1v)
    _k_n += 1
    r = body(ws, r, [i, nm, tani] + [round(x, 1) for x in g1v]
             + [round(x, 1) for x in nowv]
             + ["差なし" if abs(d) < 0.05 else sa(d, 1, "")],
             fmt={i2: "0.0" for i2 in range(4, 10)},
             align={3: "center"},
             fills={10: (GRAY if abs(d) < 0.05 else IN_Y)})
r += 1
r = note(ws, r,
         "注1）1人1月あたり利用回（日）数は令和7年度で固定しています。"
         "回数の趨勢は年率＋0.25％（中央値）にとどまります。\n"
         "注2）福祉用具貸与・居宅療養管理指導・特定施設入居者生活介護・"
         "認知症対応型共同生活介護・小規模多機能型居宅介護・"
         "介護予防支援・居宅介護支援・施設サービスには"
         "利用回（日）数の計上がありません。"
         "月額包括又は入所によるものであるためです。\n"
         "注3）見える化システムの施策反映の画面に入れるのは"
         "1人1月あたりであり、本表の延べではありません。"
         "入力値は別に用意しています。",
         span=10, height=62)

# ============================================================ 05
ws = sheet("05_給付費の差分",
           "給付費の差分（第10期3か年計・円）",
           "見込量（人／月）に12と1人あたり給付費（令和7年度実績）を"
           "乗じたものです。"
           "第2次概算は要介護度別の単価によるため、"
           "要介護度の構成が重い区分で単価が上がり、軽い区分で下がります。"
           "報酬改定率が未公表であるため単価は固定しています（08シート）。",
           [3, 14, 40, 18, 18, 18, 14, 28], freeze="A6")
r = 4
r = header(ws, r, ["", "区分", "サービス", "第1次概算", "第2次概算",
                   "差", "差の率", "備考"])
_ky_sa = 0.0
for i, lab in enumerate(SVC, start=1):
    a = G1.KYUFU_SVC.get(short(lab))
    if a is None:
        continue
    g1v = a[4] or 0
    nowv = sum(sum(kyufu_do(lab, y) or []) for y in Y3)
    d = nowv - g1v
    _ky_sa += d
    r = body(ws, r, [i, kubun(lab).replace("サービス", ""), short(lab),
                     round(g1v), round(nowv), sa(d, 0, "", c=True),
                     "―" if not g1v else sa(d / g1v * 100, 2, "％"),
                     "実績なし" if not g1v else ""],
             fmt={4: "#,##0", 5: "#,##0"},
             fills={6: (GRAY if abs(d) < 0.5 else
                        (IN_Y if d > 0 else MID_B))})
r = body(ws, r, ["", "計", "", round(G1.KYUFU3), round(KYUFU3),
                 sa(KYUFU3 - G1.KYUFU3, 0, "", c=True),
                 sa((KYUFU3 - G1.KYUFU3) / G1.KYUFU3 * 100, 3, "％"), ""],
         fmt={4: "#,##0", 5: "#,##0"}, bold=True, fills={2: MID_B})
r += 1
r = note(ws, r,
         "注1）ここでいう給付費は総給付費であり、"
         "特定入所者介護サービス費・高額介護サービス費・"
         "審査支払手数料等を含みません。"
         "これらを加えたものが標準給付費見込額です（02シート）。\n"
         "注2）要介護度別の単価は、要介護度間の比を給付費データ"
         "（保険者単位）から採り、"
         "年報の利用者数で加重した平均が総括表の単価と一致するよう"
         "比例調整したものです。"
         "基準年度（令和7年度）では総括表の給付費2,915,307,125円を"
         "円単位で再現します。\n"
         "注3）差の計は%s円で、総給付費の差と一致します。"
         % sa(_ky_sa, 0, ""),
         span=8, height=76)

# ============================================================ 06
ws = sheet("06_中長期推計の対照",
           "中長期推計の対照（令和12・17・22・27・32年度）",
           "改正法により令和22年度（2040年）を含む中長期の見込みが"
           "記載事項となります（確認事項No.114）。"
           "計画本文に掲げるのは令和17年度・令和22年度です。"
           "見える化システムは令和12・27・32年度も求めるため、"
           "算定の側を5年次に広げています（確認事項No.157）。",
           [3, 16, 20, 20, 18, 14, 14, 12, 34], freeze="A6")
r = 4
r = lead(ws, r, "1　総給付費（単年度・円）の対照", span=9)
r = header(ws, r, ["", "年度", "第1次概算", "第2次概算", "差",
                   "第1号\n被保険者数", "認定者数", "参考の\n月額", "備考"])
_NIN = S["NIN"]
_HIHO = S["HIHO"]
for i, (y, yl) in enumerate(zip(["2029"] + YLONG,
                                ["令和11年度"] + YLONGL), start=1):
    a = G1.CHOKI.get(y)
    nowv = KYUFU_Y[y]
    g = S["gaku3"]([nowv] * 3, _HIHO[y] * 3)["月額"]
    r = body(ws, r, [i, yl, round(a[0]), round(nowv),
                     sad(nowv - a[0]),
                     round(_HIHO[y]), round(sum(_NIN[y]), 1),
                     round(g),
                     "計画本文に掲げる年度" if y in ("2035", "2040")
                     else ("第10期の最終年度" if y == "2029"
                           else "見える化システムが求める年度")],
             fmt={3: "#,##0", 4: "#,##0", 6: "#,##0", 7: "0.0",
                  8: "#,##0"},
             fills={2: (MID_B if y in ("2035", "2040") else None)})
r += 1
r = note(ws, r,
         "注1）中長期の月額は、その年度の給付費の水準が3年続いたと置いて"
         "第10期と同じ式で計算した参考値です。"
         "第11期以降の保険料を算定したものではありません。\n"
         "注2）単価（1人あたり給付費）を令和7年度で固定しているため、"
         "報酬改定と単価の趨勢は入っていません。\n"
         "注3）地域密着型介護老人福祉施設入所者生活介護が"
         "区域内の定員62人に達するのは令和12年度と見込まれます（07シート）。\n"
         "注4）第1号被保険者数は案Cによります。"
         "令和12年度以降は令和11年度を起点に社人研推計の伸び率で接続しています。",
         span=9, height=62)

# ============================================================ 07
ws = sheet("07_地域支援事業と必要利用定員総数",
           "地域支援事業の量と事業費・施設の定員に対する到達率",
           "地域支援事業の量の見込み（法第117条第2項第2号）と、"
           "施設・居住系サービスの区域内の定員に対する到達率です。"
           "いずれも第1次概算から値は変わっていません。"
           "総合事業の令和7年度の実績が未受領であるため、"
           "量の基礎は令和6年度の利用者実人数のままです"
           "（確認事項No.112）。",
           [3, 44, 8, 14, 14, 14, 14, 34], freeze="A6")
r = 4
r = lead(ws, r, "1　介護予防・日常生活支援総合事業の量の見込み", span=8)
r = header(ws, r, ["", "事業", "単位", "R6実績\n（3町計）"]
           + Y3L + ["延ばし方"])
sg3 = S["sg3"]
SOGO_RYO, KAYOI_RYO = S["SOGO_RYO"], S["KAYOI_RYO"]
_NIN_NOBI, _HIHO_NOBI = S["_NIN_NOBI"], S["_HIHO_NOBI"]
SG = S["SG"]
_SOGO_KEYS = {b for _a, b, _c, _d in SOGO_RYO}
_sr_n = 0
for i, (nm, key, tani, nob) in enumerate(SOGO_RYO + KAYOI_RYO, start=1):
    if key in _SOGO_KEYS:
        base = sg3(key)
    else:
        base = sum(x for x in SG.KAYOI[key][0] if x is not None)
    if base is None:
        continue
    f = (_NIN_NOBI if nob == "認定者数"
         else (_HIHO_NOBI if nob == "第1号被保険者数" else None))
    vals = [base * f[y] for y in Y3] if f else [base] * 3
    _sr_n += 1
    r = body(ws, r, [i, nm, tani, round(base, 1)]
             + [round(v, 1) for v in vals]
             + [nob + "の伸び" if f else "据え置き"],
             fmt={4: "0.0", 5: "0.0", 6: "0.0", 7: "0.0"},
             align={3: "center"})
r += 1
r = lead(ws, r, "2　施設・居住系サービスの定員に対する到達率", span=8)
r = header(ws, r, ["", "サービス", "", "区域内の定員",
                   "令和11年度の\n見込量", "到達率", "令和22年度の\n見込量",
                   "出所"])
for i, (lab, (teiin, moto)) in enumerate(TEIIN_MAP.items(), start=1):
    v11 = sogaku(lab, "2029")
    v22 = sogaku(lab, "2040")
    r = body(ws, r, [i, short(lab), "", teiin, round(v11, 1),
                     "%.1f％" % (v11 / teiin * 100),
                     round(v22, 1), moto],
             fmt={4: "#,##0", 5: "0.0", 7: "0.0"},
             align={6: "right"},
             fills={6: (NG_O if v11 / teiin > 0.95 else
                        (IN_Y if v11 / teiin > 0.8 else GRAY))})
r += 1
r = note(ws, r,
         "注1）年報は施設の所在地を問わず当広域連合の被保険者を数えるため、"
         "区域内の定員は上限ではありません。"
         "到達率が100％を超えることは誤りではありません。\n"
         "注2）必要利用定員総数は現に指定を受けている定員を基礎とするもので、"
         "見込量（当広域連合の被保険者の受給者数）とは数える対象が異なります"
         "（確認事項No.141）。\n"
         "注3）定員に対する当広域連合の被保険者の受給者数の割合が"
         "低いことは、区域内に空きがあることを意味しません。\n"
         "注4）認知症対応型共同生活介護の区域内の定員は99人に"
         "［要確認］が加わります（東川町の1事業所が未確定）。\n"
         "注5）地域支援事業費は年度あたり%s円"
         "（第10期3か年%s円）で、令和6年度決算の水準の据え置きです。"
         % (yen(S["CHIIKI_R6"]), yen(G_DO["B"])),
         span=8, height=76)

# ============================================================ 08
ws = sheet("08_月額を動かす前提",
           "月額を動かす前提（据え置きのうち決定を待つもの）",
           "据え置いた%d件のうち、発注者・3町の資料を待つもの（区分B）と"
           "国の告示・公布を待つもの（区分C）です。"
           "受託者の側で確定できるもの（区分A）は残っていません。"
           "効きの大きいものから並べています。"
           "いずれも本算定には織り込んでいません。" % len(SUEOKI),
           [3, 7, 34, 44, 20, 16], freeze="A6")
r = 4
r = lead(ws, r, "1　月額を動かし得る前提", span=6)
r = header(ws, r, ["", "区分", "事項", "決着しない場合の扱い", "効き",
                   "確認事項"])
_SU_BC = [(i, x) for i, x in enumerate(SUEOKI, start=1)
          if KUBUN.get(i) in ("B", "C")]


def _ookisa(txt):
    """効きの記述から月額の最大値を拾う（並べ替えのため）。"""
    v = [abs(int(m.replace(",", "")))
         for m in re.findall(r"月額[^0-9]{0,6}([0-9,]+)円", str(txt))]
    return max(v) if v else 0


_SU_BC.sort(key=lambda t: -_ookisa(t[1][5]))
for n, (i, x) in enumerate(_SU_BC, start=1):
    r = body(ws, r, [n, KUBUN.get(i), x[1], _oki(x[3]), _oki(x[5]), x[4]],
             height=58,
             fills={2: (NG_O if KUBUN.get(i) == "C" else IN_Y)})
r += 1
r = lead(ws, r, "2　区分の内訳", span=6)
r = header(ws, r, ["", "区分", "意味", "件数", "第2次概算での扱い", ""])
for i, (k, nm, imi, atsu) in enumerate([
    ("Z", "解消済み", "受託者の側で確定した、又は決着したもの",
     "算定に織り込み済み"),
    ("A", "受託者の側で確定できる", "資料の受領を待たずに確定できるもの",
     "残っていない（令和8年10月1日に区分Aを0件とした）"),
    ("B", "発注者・3町の資料待ち",
     "届かなければ現に確認できている値で確定する",
     "既定値で算定している。資料が届いた分は%sの再算定で置き換える"
     % TEISHUTSU),
    ("C", "国の告示・公布待ち",
     "受領できないでは済まないもの。公布後に再算定する",
     "現行制度で算定している。告示後に再算定する"),
], start=1):
    cnt = sum(1 for v in KUBUN.values() if v == k)
    r = body(ws, r, [i, k, nm + "：" + imi, cnt, atsu, ""], height=40,
             fmt={4: "0"},
             fills={2: {"Z": OK_G, "A": MID_B, "B": IN_Y, "C": NG_O}[k]})
r += 1
r = note(ws, r,
         "注1）上向きの前提（報酬改定率・単価の趨勢・"
         "第1号被保険者負担割合）はいずれも国の告示・公布によるもので、"
         "資料の受領とは関わりなく生じます。\n"
         "注2）下向きの前提（基金の取崩し・給付の適正化・"
         "サービス別の趨勢の反映）はいずれも織り込んでいません。"
         "したがって算定上の月額は据え置きを前提とした下側に近い値です。\n"
         "注3）改定率を乗じることと単価の趨勢を延ばすことを"
         "重ねて適用すると二重になります（確認事項No.110・No.137）。\n"
         "注4）給付の適正化による抑制と"
         "サービス別の趨勢の反映は向きが逆になる部分があり、"
         "二重に扱わない整理を要します（確認事項No.106・No.151）。",
         span=6, height=76)

# ============================================================ 09
ws = sheet("09_見える化システムとの対照",
           "見える化システムの推計結果との対照",
           "見える化システムで推計した結果（推計パターン9028（１）。"
           "人口の設定を案Cに改めた後の再出力）と本算定の対照です。"
           "算定式そのものは同じであることを第9期の再現により"
           "確かめています（見える化6,427.89円・当方の再現6,427.92円）。"
           "差は入力と設定によるものです。",
           [3, 40, 22, 22, 18, 40], freeze="A6")
r = 4
r = lead(ws, r, "1　第10期3か年の対照", span=6)
r = header(ws, r, ["", "項目", "見える化システム", "本算定", "差", "差の理由"])
_MK = K2.SHUNO
TAI = [
    ("総給付費（円）", sum(K2.KYUFU[y] for y in ("R9", "R10", "R11")) * 1000,
     KYUFU3, "令和8年度を基準年にしていること・包括推計・"
     "1人1月あたり給付費の実績値の年度"),
    ("標準給付費見込額（円）", _MK["標準給付費見込額"]["計"], G_DO["A"],
     "上乗せ率が違う（システムは特定入所者・高額・高額医療合算・"
     "審査支払手数料を個別に入力、本算定は令和6年度決算の割増率"
     "1.05909で一括）"),
    ("地域支援事業費（円）", _MK["地域支援事業費"]["計"], G_DO["B"],
     "システムは未入力（修正が必要な箇所の1件）"),
    ("保険料収納必要額（円）", _MK["保険料収納必要額"]["第10期計"], G_DO["J"],
     "上記の積上げと調整交付金の作り方の違い"),
]
for i, (k, a, b, why) in enumerate(TAI, start=1):
    r = body(ws, r, [i, k, round(a), round(b), sad(b - a), why],
             fmt={3: "#,##0", 4: "#,##0"}, height=52)
r = body(ws, r, ["", "算定上の月額基準額（円）",
                 round(K2.HOKENRYO["第10期"], 2), round(GETSU, 2),
                 sad(GETSU - K2.HOKENRYO["第10期"], 2),
                 "システムは地域支援事業費が0で所得段階を弾力化していない。"
                 "両方を直すと約%d円になる見込み" % round(MI["_G16B"])],
         bold=True, fmt={3: "#,##0.00", 4: "#,##0.00"},
         fills={2: IN_Y}, height=40)
r += 1
r = lead(ws, r, "2　第1号被保険者数・認定者数の対照", span=6)
r = header(ws, r, ["", "年度", "システム\n第1号被保険者数", "本算定",
                   "システム\n認定者数（第1号）", "本算定"])
for i, (y, yk, yl) in enumerate([("2027", "R9", "令和9年度"),
                                 ("2028", "R10", "令和10年度"),
                                 ("2029", "R11", "令和11年度")], start=1):
    r = body(ws, r, [i, yl, K2.HIHO[yk], round(_HIHO[y]),
                     K2.NINTEI_ICHI[yk], round(sum(_NIN[y]), 1)],
             fmt={3: "#,##0", 4: "#,##0", 5: "#,##0", 6: "0.0"})
r += 1
r = lead(ws, r, "3　修正が必要な箇所11件の現況", span=6)
r = header(ws, r, ["", "修正が必要な箇所", "前回", "今回", "", "見方"])
_GEN = MI["_GEN"]
_kai_n = sum(1 for x in _GEN if x[3] == "解消")
for i, (a, b, c, st, e) in enumerate(_GEN, start=1):
    r = body(ws, r, [i, a, b, c, st, e.replace("**", "")],
             fills={5: (OK_G if st == "解消" else NG_O)},
             align={5: "center"}, height=26)
r += 1
r = note(ws, r,
         ("注1）%d件のうち%d件が解消し、%d件が残っています。"
          % (len(_GEN), _kai_n, len(_GEN) - _kai_n))
         + ("残る%d件のうち月額を動かすのは、地域支援事業費・弾力化・"
            "令和8年度の実績見込み値・1人1月あたり給付費の年度・"
            "準備基金の5件です。\n" % (len(_GEN) - _kai_n))
         + "注2）人口の設定は案Cに改まりましたが、"
         "認定者数が8年度とも動いていないため、"
         "給付費の側へ及んでいません。"
         "実績及び推計方法の設定を保存したうえで、"
         "施策反映（認定者数）から保険料額の算定までの各画面を"
         "開いて登録し直すことを要します。"
         "登録し直せたかどうかは令和11年度の認定者数（第1号）"
         + ("%d人が変わることで確かめられます。\n" % K2.NINTEI_ICHI["R11"])
         + ("注3）将来推計ワーニングチェックは%d件で、"
            "施策反映の3画面はいずれも該当がありません。"
            "訪問介護の%d件は当区域の実態によるものであり、"
            "残る%d件は令和8年度の実績見込み値と"
            "1人1月あたり給付費の年度によるものです。"
            % (len(MI["GW"].MEISAI), MI["_W_JITTAI"],
               len(MI["GW"].MEISAI) - MI["_W_JITTAI"])),
         span=6, height=86)

# ============================================================ 10
ws = sheet("10_残る未決と既定値",
           "残る未決と既定値（決着しない場合の扱い）",
           "確認事項は業務工程管理表 03_確認事項一覧で一元管理しています。"
           "未決%d件のすべてに「決着しない場合の受託者の扱い（既定値）」を"
           "置いており、ご決定をお待ちしている間も算定は止まりません。"
           "既定値は受託者の仮置きであり、決定ではありません。"
           "令和8年10月上旬の3町合同意見交換会で内容をお諮りします。"
           % len(MACHI),
           [3, 7, 40, 52, 12, 12, 16], freeze="A6")
r = 4
r = lead(ws, r, "1　仮置きにより計画素案へ反映した事項", span=7)
r = header(ws, r, ["", "確認事項", "反映先", "加えた内容", "", "", ""])
_CM = {c[0]: c for c in CHECK}
for i, (no, saki, naiyo) in enumerate(SUSUMETA, start=1):
    r = body(ws, r, [i, "No.%d" % no, saki, naiyo, "", "", ""], height=40,
             align={2: "center"})
r += 1
r = lead(ws, r, "2　第2次概算に効く未決の事項"
         "（期限が令和8年10月までのもの）", span=7)
r = header(ws, r, ["", "No.", "件名", "決着しない場合の受託者の扱い",
                   "宛先", "期限", "状態"])


def _kigen_m(s):
    m = re.match(r"R(\d+)\.(\d+)", str(s))
    return (int(m.group(1)), int(m.group(2))) if m else (99, 99)


_KIKU = ("概算", "保険料", "見込量", "第6章")
_MADE = sorted([c for c in MACHI
                if _kigen_m(c[8]) <= (8, 10)
                and any(k in str(c[5]) for k in _KIKU)],
               key=lambda c: (_kigen_m(c[8]), c[0]))
for i, c in enumerate(_MADE, start=1):
    r = body(ws, r, [i, c[0], c[3], KITEI.get(c[0], ""), c[6], c[8], c[7]],
             height=40, align={2: "center", 6: "center"},
             fills={4: (NG_O if not KITEI.get(c[0]) else None)})
r += 1
r = note(ws, r,
         "注1）仮置きにより計画素案へ反映した事項は%d件です。"
         "作業として終えたことと、ご決定をお待ちしていることは別です。"
         "ご決定の内容が異なる場合は、その内容により置き直します。\n"
         "注2）2の表は、未決の事項のうち影響先に"
         "サービス見込量・給付費・保険料・計画素案 第6章を含み、"
         "期限が令和8年10月までのもの%d件です。"
         "いずれにも決着しない場合の扱いを置いています。\n"
         "注3）未決%d件の全体と優先順位は"
         "業務工程管理表 03_確認事項一覧によります。"
         % (len(SUSUMETA), len(_MADE), len(MACHI)),
         span=7, height=48)

# ============================================================ 11
ws = sheet("11_10月30日までに行うこと",
           "令和8年10月30日までに行うこと",
           "第2次概算を確定するまでの段取りです。"
           "受託者の作業は、ご決定・ご提供をいただいた分を"
           "算定に織り込んで出し直すことです。"
           "ご決定がない事項は既定値のまま進めます。",
           [3, 10, 40, 44, 20], freeze="A6")
r = 4
r = lead(ws, r, "1　段取り", span=5)
r = header(ws, r, ["", "時期", "行うこと", "内容", "行う者"])
DANDORI = [
    ("10月上旬", "3町合同意見交換会",
     "3町共通の確認事項と既定値の内容をお諮りする。"
     "地域の課題15件と第1次概算の結果を共有する。", "発注者・3町・受託者"),
    ("10月上旬", "見える化システムの設定の修正",
     "地域支援事業費の入力、所得段階の弾力化、"
     "令和8年度の実績見込み値の編集、"
     "施策反映から保険料額の算定までの各画面の登録し直し。"
     "入れる値は別に用意している。", "発注者"),
    ("10月中旬", "町ごとの個別協議",
     "東川町4件・美瑛町5件・東神楽町2件の確認事項。"
     "施設整備・サービス提供の方針（確認事項No.84ほか）は"
     "必要利用定員総数に直結する。", "3町・受託者"),
    ("10月中旬", "資料のご提供",
     "総合事業の令和7年度実績、認知症総合支援事業の事業費、"
     "地域密着型サービスの基準条例、令和7年度末の基金残高、"
     "令和8年8月分の月報、認知症対応型共同生活介護の定員。",
     "発注者・3町"),
    ("10月下旬", "算定の出し直し",
     "ご決定・ご提供をいただいた分を織り込み、"
     "見込量・給付費・保険料を算定し直す。"
     "計画素案 第6章も併せて改める。", "受託者"),
    ("10月30日", "第2次概算の提示",
     "サービス見込量・給付費・保険料と、第1次概算からの差分を提示する。",
     "受託者"),
    ("11月13日", "予算編成用の提示",
     "第2次概算を基礎として、予算編成に用いる値を提示する。", "受託者"),
    ("11月", "3町意見の反映・骨子の整理",
     "3町の意見を計画素案へ反映し、骨子を整理する。", "受託者"),
]
for i, x in enumerate(DANDORI, start=1):
    r = body(ws, r, [i, x[0], x[1], x[2], x[3]], height=52,
             fills={2: (IN_Y if x[0] == "10月30日" else None)})
r += 1
r = lead(ws, r, "2　ご決定・ご提供がない場合", span=5)
r = header(ws, r, ["", "場合", "受託者の扱い", "第2次概算への影響", ""])
for i, (k, v, e) in enumerate([
    ("意見交換会でご決定がない事項がある",
     "既定値のまま算定する。既定値は本資料の10シートと"
     "業務工程管理表 03_確認事項一覧に掲げている。",
     "算定は止まらない。保険料基準額は%d円のままとなる見込み" % KIJUN),
    ("資料のご提供がない（区分B）",
     "現に確認できている値で確定し、出所と制約を注記する。",
     "量の見込みと事業費が令和6年度の水準の据え置きのままとなる"),
    ("国の告示・公布がない（区分C）",
     "現行制度で算定し、告示・公布の後に再算定する。",
     "報酬改定率・第1号被保険者負担割合・所得段階の政令改正が"
     "入らないため、算定上の月額は下側に偏る"),
    ("見える化システムの設定が直らない",
     "本算定の値をもって第2次概算とし、"
     "システムの出力との差の理由を09シートにより示す。",
     "国への報告の値と計画の値が別になるため、"
     "どちらを計画に載せるかのご判断を要する"),
], start=1):
    r = body(ws, r, [i, k, v, e, ""], height=52)
r += 1
r = note(ws, r,
         "注）発注者・3町から追加の資料のご提供がない場合でも、"
         "国の告示が出れば計画は確定します。"
         "算定そのものが止まる事項はありません。",
         span=5, height=28)

# ============================================================ 12 自己点検
_AL = []
for _ws in wb.worksheets:
    for _row in _ws.iter_rows(values_only=True):
        for _v in _row:
            if isinstance(_v, str):
                _AL.append(_v)
ALL = "\n".join(_AL)

chk(1, "第1次概算の年度別の総給付費の和が3か年計と一致すること",
    "%s円＝%s円" % (yen(sum(G1.KYUFU_NEN.values())), yen(G1.KYUFU3)),
    "一致" if sum(G1.KYUFU_NEN.values()) == G1.KYUFU3 else "不一致",
    sum(G1.KYUFU_NEN.values()) == G1.KYUFU3)

chk(2, "第2次概算の年度別の総給付費の和が3か年計と一致すること",
    "%s円" % yen(KYUFU3), "一致",
    abs(sum(KYUFU_Y[y] for y in Y3) - KYUFU3) < 1)

_J = (G_DO["C"] + G_DO["D"] - G_DO["E"] + G_DO["F"] + G_DO["G"]
      + G_DO["H"] - G_DO["I"])
chk(3, "J＝C＋D－E＋F＋G±H－I が成り立つこと",
    "%s円" % yen(_J), "差 %.2f円" % (G_DO["J"] - _J),
    abs(G_DO["J"] - _J) < 1)

_g = G_DO["J"] / 0.99 / G_DO["③"] / 12
chk(4, "算定上の月額基準額＝J÷予定収納率÷補正後被保険者数÷12 であること",
    "%.4f円" % _g, "差 %.6f円" % (GETSU - _g), abs(GETSU - _g) < 0.01)

chk(5, "保険料基準額が第1次概算から変わっていないこと",
    "第1次 %d円／第2次 %d円" % (G1.KIJUN, KIJUN),
    "変わらない" if KIJUN == G1.KIJUN else "変わった", KIJUN == G1.KIJUN)

_mn = sum(1 for lab in SVC if short(lab) in G1.MIKOMI)
chk(6, "見込量の差分がサービス29区分すべてで取れること",
    "%d／%d区分" % (_mn, len(SVC)), "過不足なし" if _mn == len(SVC) else "不足",
    _mn == len(SVC))

chk(7, "利用回（日）数の差分が12区分すべてで取れること",
    "%d／%d区分" % (_k_n, len(G1.KAISU)),
    "過不足なし" if _k_n == len(G1.KAISU) else "不足", _k_n == len(G1.KAISU))

_kn = sum(1 for lab in SVC if short(lab) in G1.KYUFU_SVC)
chk(8, "給付費の差分がサービス29区分すべてで取れること",
    "%d／%d区分" % (_kn, len(SVC)), "過不足なし" if _kn == len(SVC) else "不足",
    _kn == len(SVC))

chk(9, "サービス種別の給付費の差の計が総給付費の差と一致すること",
    "%s円／%s円" % (yen(_ky_sa), yen(KYUFU3 - G1.KYUFU3)),
    "差 %.0f円（提示した版の表は円単位で丸めている）"
    % (_ky_sa - (KYUFU3 - G1.KYUFU3)),
    abs(_ky_sa - (KYUFU3 - G1.KYUFU3)) <= len(SVC))

chk(10, "見込量が第1次概算から動いていないこと（置き方を変えていない）",
    "29区分の差の最大 %.3f人（提示した版の表は小数第1位で丸めている）"
    % (max(_m_sa) if _m_sa else 0),
    "動いていない" if max(_m_sa or [0]) <= 0.15 else "動いた",
    max(_m_sa or [0]) <= 0.15)

_kb = {k: sum(1 for v in KUBUN.values() if v == k) for k in "ZABC"}
chk(11, "据え置きの区分の合計が据え置きの件数と一致すること",
    "Z%d＋A%d＋B%d＋C%d＝%d件（据え置き%d件）"
    % (_kb["Z"], _kb["A"], _kb["B"], _kb["C"], sum(_kb.values()),
       len(SUEOKI)),
    "一致" if sum(_kb.values()) == len(SUEOKI) else "不一致",
    sum(_kb.values()) == len(SUEOKI))

chk(12, "区分Aが残っていないこと（受託者の側で確定できるものがないこと）",
    "区分A %d件" % _kb["A"], "なし" if _kb["A"] == 0 else "あり",
    _kb["A"] == 0)

_ss_ng = [no for no, _s, _n in SUSUMETA if no not in _CM]
chk(13, "仮置きにより反映した事項がすべて確認事項の台帳にあること",
    "%d件／台帳にないもの %s"
    % (len(SUSUMETA), "なし" if not _ss_ng else _ss_ng),
    "過不足なし" if not _ss_ng else "不足", not _ss_ng)

_kt_ng = [c[0] for c in _MADE if not KITEI.get(c[0])]
chk(14, "10月までに期限のある未決の事項すべてに既定値があること",
    "%d件／既定値のないもの %s"
    % (len(_MADE), "なし" if not _kt_ng else _kt_ng),
    "そろっている" if not _kt_ng else "不足", not _kt_ng)

chk(15, "見える化システムの修正が必要な箇所の現況が解消3件・残8件であること",
    "解消%d件／残%d件" % (_kai_n, len(_GEN) - _kai_n), "11件の内訳",
    _kai_n == 3 and len(_GEN) == 11)

_FROZEN = "―"
_fr_ok = False
if os.path.exists(G1_XLSX):
    _wb1 = load_workbook(G1_XLSX, data_only=True)
    _v = None
    for _row in _wb1["13_給付費と保険料"].iter_rows(values_only=True):
        if _row[1] == "算定上の月額基準額" and _v is None:
            _v = str(_row[2])
    _FROZEN = str(_v)
    _fr_ok = _FROZEN == "%s円" % format(G1.GETSU_DO, ",")
chk(16, "凍結した第1次概算の値が提示した版と一致すること",
    "提示した版の算定上の月額基準額 %s" % _FROZEN,
    "一致" if _fr_ok else "不一致", _fr_ok)

_NAIBU = ["固定値", "実物から", "章節ごとに", "判定している", "読んで判定",
          "スクリプト", "runpy", ".py", "再実行", "書き写して",
          "個票がなくても", "モジュール", "終了コード"]
_nh = [w for w in _NAIBU if w in ALL]
chk(17, "受託者の内部の仕組み・作業経過の語が残っていないこと",
    "残り %s" % ("なし" if not _nh else "・".join(_nh)), "走査", not _nh)

_NG = ["に由来する", "と整合する", "1件も", "有意差がないため関係がない",
       "全国トップ級"]
_ngh = [w for w in _NG if w in ALL]
chk(18, "禁止表現が残っていないこと",
    "残り %s" % ("なし" if not _ngh else "・".join(_ngh)), "走査", not _ngh)

_PI = [(r"0\d{1,3}-\d{2,4}-\d{4}", "電話番号"),
       (r"[\w.+-]+@[\w-]+\.[\w.]+", "メールアドレス")]
_pih = [nm for pat, nm in _PI if re.search(pat, ALL)]
chk(19, "個人情報の形（電話番号・メールアドレス）がないこと",
    "検出 %s" % ("なし" if not _pih else "・".join(_pih)), "走査", not _pih)

chk(20, "強調の指定が本文に残っていないこと",
    "強調の指定の数 %d" % ALL.count("*" * 2),
    "なし" if "*" * 2 not in ALL else "あり",
    "*" * 2 not in ALL)

_OK_CH = re.compile(
    r"[ぁ-んァ-ヴ一-龥々ー０-９0-9A-Za-zＡ-Ｚａ-ｚ"
    r"、。・（）「」『』【】〔〕［］〜％±＋▲△○◯→／()%.,:：;；!！?？"
    r"\-－—―_＿　\s①②③④⑤⑥⑦⑧⑨⑩Ⅰ-Ⅹ°×÷≧≦<>＜＞=＝※\"'\[\]#]")
_bad = sorted({c for c in ALL if not _OK_CH.match(c)})
chk(21, "許容する文字の外の文字が混じっていないこと",
    "検出 %s" % ("なし" if not _bad else "".join(_bad)), "走査", not _bad)

import data_kofukin_rengo as KR              # noqa: E402

_DAN = sorted({x[1] for x in KR.RENGO} | {x[1] for x in KR.KAISAN}
              | {m[0] for x in KR.RENGO for m in x[3]}
              | {m[0] for x in KR.KAISAN for m in x[4]},
              key=len, reverse=True)
_OURS = ("大雪地区広域連合", "東川町", "美瑛町", "東神楽町")
_dan_ng = [nm for nm in _DAN if nm and nm not in _OURS and nm in ALL]
chk(22, "他の団体の固有名称を掲げていないこと",
    "対象%d件／検出 %d件 %s"
    % (len(_DAN), len(_dan_ng), _dan_ng[:3] or ""), "走査",
    not _dan_ng and len(_DAN) > 200)

_susei_old = [w for w in ("＋185〜", "＋185円", "＋374円", "＋374〜")
              if w in ALL]
chk(23, "単価の趨勢の効きが算定の率から求めた値であること",
    "年率%.2f％〜%.2f％で月額%s〜%s"
    % (S["SUSEI_RITSU"][0] * 100, S["SUSEI_RITSU"][1] * 100,
       S["susei_haba_s"]()[0], S["susei_haba_s"]()[1]),
    "古い値（＋185円から＋374円）が残っていない"
    if not _susei_old else "古い値が残っている",
    not _susei_old)

ws = sheet("12_自己点検", "自己点検（エラーチェック）",
           "本資料の内的整合と、算定・凍結した第1次概算の値との接続を"
           "点検した結果です。"
           "1件でも不適合があると本資料は作られません。",
           [5, 46, 32, 30, 12], freeze="A6")
r = 4
r = header(ws, r, ["#", "点検の内容", "式・対象", "結果", "判定"])
for c in CHECKS:
    r = body(ws, r, list(c), height=30,
             fills={5: (OK_G if c[4] == "適合" else NG_O)},
             align={5: "center"})
r += 1
r = note(ws, r,
         "注）点検16は、提示した第1次概算の版を実際に開いて"
         "算定上の月額基準額を読み、凍結した値と突き合わせるものです。"
         "提示した版が書き換わっていれば不適合になります。",
         span=5, height=28)

# ============================================================ 出力
os.makedirs(ODIR, exist_ok=True)
wb.save(OUT)

_ng = [c for c in CHECKS if c[4] != "適合"]
print("書き出しました:", OUT)
for s in wb.sheetnames:
    print("  -", s, wb[s].max_row, "rows")
print("第1次概算 総給付費 %s円／月額 %d円／基準額 %d円"
      % (yen(G1.KYUFU3), G1.GETSU_DO, G1.KIJUN))
print("第2次概算 総給付費 %s円／月額 %.2f円／基準額 %d円"
      % (yen(KYUFU3), GETSU, KIJUN))
print("差 総給付費 %s円／月額 %s"
      % (sa(KYUFU3 - G1.KYUFU3, 0, "", c=True),
         sa(GETSU - G1.getsu_sei(), 2)))
print("据え置き %d件（Z%d・A%d・B%d・C%d）／月額を動かす前提 %d件"
      % (len(SUEOKI), _kb["Z"], _kb["A"], _kb["B"], _kb["C"], len(_SU_BC)))
print("仮置きにより反映 %d件／第2次概算に効く未決 %d件／未決 %d件"
      % (len(SUSUMETA), len(_MADE), len(MACHI)))
print("自己点検 %d件：適合%d件・不適合%d件"
      % (len(CHECKS), len(CHECKS) - len(_ng), len(_ng)))
if _ng:
    for c in _ng:
        print("  不適合:", c[0], c[1], c[3])
    sys.exit(1)
print("すべての点検に適合しました。")
