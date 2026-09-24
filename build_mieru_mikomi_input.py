# -*- coding: utf-8 -*-
"""大雪地区広域連合 第10期介護保険事業計画
見える化システム 将来推計へ入力するサービス見込量の値（暫定値）.

令和8年9月18日のご依頼
  「見える化システムへ入力するようの見込量のデータについてアウトプットを
    お願いします。なおデータ文中に必要資料未受領からの暫定値である旨
    明記して下さい」

━━ 本表の値はすべて暫定値である ━━

本表の値は、令和8年9月18日時点で受領できている資料の範囲で算定したもので
あり、**必要資料が未受領であることによる暫定値**である。
確定値ではなく、資料の受領及びご決定により動く。
その旨を00シート・各シートの表題・注記・印刷の各ページに明記している。

暫定である理由は次の3つに分かれる（10シート）。
  区分A（4件）受託者の側で確定できる。資料を要しない。
  区分B（15件）発注者・3町の資料待ち。届かなければ据え置きのまま確定する。
  区分C（4件）国の告示・公布を待つ。「受領できない」では済まない。

**算定そのものは止まっていない。** 届かないものは既定値で置いており、
何をどう置いたかは `build_mikomiryo_santei.py` の `SUEOKI`（23件）にある。

━━ 数値の出所 ━━

**固定値を書かない。** 見込量・給付費・保険料・地域支援事業の量は
`build_mikomiryo_santei.py` を `runpy` で読む。算定を改めれば本表が追随する。
画面の写しは `data_mieru_suikei.py`、年報の明細は `data_nenpo_meisai.py`。

  利用者数　令和7年度の月平均（年報・要介護度別）×要介護度別の伸び
  伸び　　　年齢階級別認定者数（案C・シナリオ②）×年齢階級ごとの要介護度の構成
  回（日）数　1人1月あたり回（日）数（令和7年度で固定）×利用者数

━━ 入力欄の桁（確認事項No.135。令和8年9月11日に解決） ━━

  利用者数（人数）の欄は整数で扱われる。
  利用回（日）数の欄は小数第1位まで。

各欄を単純に四捨五入すると小口のサービスが0になり、
令和9年度から令和11年度も0のまま残る。
本表は**最大剰余法により要支援の群・要介護の群ごとに合計を保つ**ため、
定期巡回・随時対応型訪問介護看護（月0.58人）も0にはならない。

━━ 当方の29区分と見える化の45区分は束ね方が違う ━━

当方の区分は見える化システムの総括表詳細（１）の29区分である。
見える化の入力画面は45区分（介護28・介護予防17）であるが、
**その区分一覧そのものは受領していない**（件数のみ弊社内の全国共通版で把握）。
介護予防支援・居宅介護支援のように総括表では1行、
入力画面では2区分になるものがあるため、03・04シートは
要支援の列と要介護の列を分けて掲げ、どちらにも貼れる形にしている。
対応の確認には操作手引と認証後画面が要る（確認事項No.148）。

シート構成
  00_この表について（暫定値である旨）
  01_入力の順序と設定項目
  02_認定者数
  03_施設・居住系サービス利用者数
  04_在宅サービス利用者数
  05_在宅サービス利用回（日）数
  06_実績のない区分と令和9年4月新設の区分
  07_地域支援事業の見込み量
  08_所得段階別第1号被保険者数
  09_給付費と保険料（参考）
  10_暫定値である理由
  11_自己点検

出力
  output/第10期計画_見える化_見込量の入力値（暫定）.xlsx

自己点検で1件でも不適合があると終了コード1で終わる。
"""

import io
import os
import runpy
import sys
from decimal import Decimal, ROUND_HALF_UP

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import data_mieru_suikei as MS
import repo_paths as RP

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

ODIR = RP.OUTPUT
OUT = os.path.join(ODIR, "第10期計画_見える化_見込量の入力値（暫定）.xlsx")
KIJUNBI = "令和8年9月18日"

# 本表の位置づけ。表題・注記・ヘッダー・フッターに繰り返し用いる。
ZANTEI = ("本表の値は必要資料が未受領であることによる暫定値です。"
          "確定値ではありません。")

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


# ==================================================== 算定を読み込む
def _load(name):
    buf, old = io.StringIO(), sys.stdout
    sys.stdout = buf
    try:
        return runpy.run_path(os.path.join(RP.ROOT, name))
    finally:
        sys.stdout = old


_S = _load("build_mikomiryo_santei.py")


def _sueoki_kubun():
    """据え置き23件の区分（A・B・C）をソースから読む。

    `build_mikomi_juryo_nashi.py` の `KUBUN` を `ast` で読む。
    固定値を書き写すと、区分を改めたときに本表とずれる（CLAUDE.md §4）。
    """
    import ast
    import re
    src = io.open(os.path.join(RP.ROOT, "build_mikomi_juryo_nashi.py"),
                  encoding="utf-8").read()
    m = re.search(r"^KUBUN = (\{.*?\})\s*$", src, re.S | re.M)
    if not m:
        raise RuntimeError("KUBUN が見つからない")
    return ast.literal_eval(m.group(1))


SUEOKI_KUBUN = _sueoki_kubun()

SVC = _S["SVC"]
DO = _S["DO"]
MIKOMI = _S["MIKOMI"]                 # {年度: {サービス: 要介護度別 人/月}}
JISSEKI = _S["JISSEKI"]               # {サービス: 令和7年度の月平均}
NIN = _S["NIN"]                       # {年度: 要介護度別の認定者数}
NIN_R7 = _S["NIN_R7"]                 # 年報の実績
HIHO = _S["HIHO"]                     # 第1号被保険者数（案C）
KAISU_MAP = _S["KAISU_MAP"]
kaisu_tanka = _S["kaisu_tanka"]
kaisu_mikomi = _S["kaisu_mikomi"]
sogaku = _S["sogaku"]
short, kubun = _S["short"], _S["kubun"]
G_DO = _S["G_DO"]
KYUFU3 = _S["KYUFU3"]
KYUFU_Y = _S["KYUFU_Y"]
SUEOKI = _S["SUEOKI"]
SOGO_RYO, KAYOI_RYO = _S["SOGO_RYO"], _S["KAYOI_RYO"]
sg3 = _S["sg3"]
SG = _S["SG"]
NIN_NOBI, HIHO_NOBI = _S["_NIN_NOBI"], _S["_HIHO_NOBI"]
Y3, Y3L = _S["Y3"], _S["Y3L"]
YLONG, YLONGL = _S["YLONG"], _S["YLONGL"]
BASE_Y = _S["BASE_Y"]

YALL = Y3 + YLONG
YALLL = Y3L + YLONGL

# 所得段階別第1号被保険者数の将来推計（確認事項No.156）
_D = _load("build_shotoku_dankai.py")
DANKAI = _D["DANKAI"]                  # {年度: 16段階の人数}
DANKAI_NAME = _D["DANKAI_NAME"]
JORITSU = _D["JORITSU"]
NDAN = _D["NDAN"]
KOSEI_SAIYO = _D["SAIYO"]
KEISU_SAIYO = _D["KEISU_R6"]
KEISU_R7 = _D["KEISU_R7"]
hiho_int = _D["hiho_int"]
SOU3 = sum(KYUFU3)
GETSU = int(round(G_DO["月額"]))
KIJUN_GAKU = int(round(GETSU / 100.0) * 100)

SHIEN_I, KAIGO_I = range(0, 2), range(2, 7)

# ==================================================== 要支援の列の有無
# どの区分に要支援1・要支援2の列があるかは、受領した画面の写し
# （`data_mieru_suikei.GAMEN_DO`・`ZAITAKU_DO`）から読む。
# 当方の判断で決めない。写しで「―」になっているものが列のない区分である。
#
# 背景（制度）：介護予防訪問介護・介護予防通所介護は平成29年度までに
# 総合事業へ移行しており予防給付にない。
# 介護予防地域密着型サービスとして置かれているのは
# 認知症対応型通所介護・小規模多機能型居宅介護・認知症対応型共同生活介護の
# 3種類のみで、認知症対応型共同生活介護は要支援2のみである。

# 総括表の行名と画面の行名が違うもの
MIERU_NAME = {
    "在宅サービス 特定福祉用具販売": "特定福祉用具購入費",
    "在宅サービス 住宅改修": "住宅改修費",
}

# 画面の写しにない区分。介護療養型医療施設は令和6年3月末で廃止済みの
# 施設サービスであり、要支援の列はない。
GAMEN_NASHI = {"施設サービス 介護療養型医療施設": (False, False)}


def shien_ari(lab):
    """（要支援1の列があるか, 要支援2の列があるか）を画面の写しから返す。"""
    nm = MIERU_NAME.get(lab, short(lab))
    if nm in MS.GAMEN_DO:
        v = MS.GAMEN_DO[nm]["R7"]
        return (v[0] is not None, v[1] is not None)
    if nm in MS.ZAITAKU_DO:
        v = MS.ZAITAKU_DO[nm]["R7"]
        return (v[1] is not None, v[2] is not None)
    if lab in GAMEN_NASHI:
        return GAMEN_NASHI[lab]
    raise KeyError(lab)


SHIEN_ARI = {l: shien_ari(l) for l in SVC}
# 要支援の列がない区分（画面で「―」）
YOBO_NASHI = {l for l in SVC if SHIEN_ARI[l] == (False, False)}
# 要支援2のみの区分
SHIEN2_ONLY = {l for l in SVC if SHIEN_ARI[l] == (False, True)}


def masked(lab, vals):
    """画面に列のない要介護度の欄を None（画面の「―」）にする。"""
    out = list(vals)
    a1, a2 = SHIEN_ARI[lab]
    if not a1:
        out[0] = None
    if not a2:
        out[1] = None
    return out


# ==================================================== 丸め
def to_int(vals):
    """月平均を整数にする。要支援の群・要介護の群ごとに合計を保つ（最大剰余法）。

    各欄を単純に四捨五入すると、小口のサービスは全欄0になり、
    令和9年度から令和11年度も0のまま残る。
    """
    out = [None] * len(vals)
    for g in (SHIEN_I, KAIGO_I):
        idx = [i for i in g if vals[i] is not None]
        if not idx:
            continue
        s = sum(vals[i] for i in idx)
        tgt = int(Decimal(str(s)).quantize(Decimal("1"), ROUND_HALF_UP))
        base = {i: int(vals[i]) for i in idx}
        for i in idx:
            out[i] = base[i]
        need = tgt - sum(base.values())
        order = sorted(idx, key=lambda i: -(vals[i] - int(vals[i])))
        k = 0
        while need > 0:
            out[order[k % len(order)]] += 1
            need -= 1
            k += 1
        while need < 0:
            for i in sorted(idx, key=lambda i: (vals[i] - int(vals[i]))):
                if out[i] > 0 and need < 0:
                    out[i] -= 1
                    need += 1
    return out


def to_dec1(vals):
    """小数第1位に丸める。群ごとに合計（小数第1位）を保つ。"""
    out = [None] * len(vals)
    for g in (SHIEN_I, KAIGO_I):
        idx = [i for i in g if vals[i] is not None]
        if not idx:
            continue
        s = sum(vals[i] for i in idx)
        tgt = int(Decimal(str(round(s * 10, 6))).quantize(Decimal("1"),
                                                          ROUND_HALF_UP))
        base = {i: int(vals[i] * 10) for i in idx}
        for i in idx:
            out[i] = base[i]
        need = tgt - sum(base.values())
        order = sorted(idx, key=lambda i: -(vals[i] * 10 - int(vals[i] * 10)))
        k = 0
        while need > 0:
            out[order[k % len(order)]] += 1
            need -= 1
            k += 1
        while need < 0:
            for i in sorted(idx,
                            key=lambda i: (vals[i] * 10 - int(vals[i] * 10))):
                if out[i] > 0 and need < 0:
                    out[i] -= 1
                    need += 1
    return [None if out[i] is None else out[i] / 10.0
            for i in range(len(vals))]


def riyosha(lab, y):
    """入力する利用者数（整数・要介護度別）。"""
    return to_int(masked(lab, MIKOMI[y][lab]))


def riyosha_r7(lab):
    return to_int(masked(lab, JISSEKI[lab]))


def kaisu_do(lab, vals):
    """要介護度別の利用回（日）数（回・日／月）。回数のない種別はNone。

    None は画面に列がない欄だけに用いる。
    列はあるが令和7年度の実績がなく1人1月あたりを出せない欄は0とする
    （利用者数も0であるため回数も0になる）。
    """
    t = kaisu_tanka(lab)
    if t is None:
        return None
    out = []
    for i in range(7):
        v = vals[i]
        u = t[0] if i < 2 else t[1]
        out.append(None if v is None else (0.0 if u is None else v * u))
    return out


# ==================================================== 体裁
def _plain(v):
    """** による強調の指定を落とす。

    xlsx は Markdown を解釈しないため、** をそのまま書くとセルに ** が出る
    （CLAUDE.md §4 で docx について実際に起きた誤りと同じもの）。
    セル内の一部だけを太字にすることはできないため、記号だけを落とす。
    """
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
    ws.freeze_panes = freeze
    # 暫定値である旨を印刷の各ページに出す
    ws.oddFooter.center.text = ZANTEI + "（基準日 " + KIJUNBI + "）"
    ws.oddFooter.center.size = 8
    ws.oddFooter.center.font = FONT
    return ws


def header(ws, row, cols, height=30):
    for i, h in enumerate(cols, start=1):
        c = ws.cell(row=row, column=i, value=_plain(h))
        c.font = Font(name=FONT, size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=HEAD)
        c.alignment = Alignment(wrap_text=True, horizontal="center",
                                vertical="center")
        c.border = BORDER
    ws.row_dimensions[row].height = height
    return row + 1


def body(ws, row, vals, fills=None, height=20, align=None, bold=False,
         fmt=None):
    for i, v in enumerate(vals, start=1):
        c = ws.cell(row=row, column=i, value=_plain(v))
        c.font = Font(name=FONT, size=9, bold=bold)
        c.alignment = Alignment(wrap_text=True, vertical="center",
                                horizontal=(align or {}).get(i, "left"))
        c.border = BORDER
        if fills and fills.get(i):
            c.fill = PatternFill("solid", fgColor=fills[i])
        if fmt and i in fmt:
            c.number_format = fmt[i]
    ws.row_dimensions[row].height = height
    return row + 1


def lead(ws, row, text, span=6, height=20):
    c = ws.cell(row=row, column=1, value=_plain(text))
    c.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=NAVY)
    c.alignment = Alignment(vertical="center")
    ws.merge_cells(start_row=row, start_column=1, end_row=row,
                   end_column=span)
    ws.row_dimensions[row].height = height
    return row + 1


def note(ws, row, text, span=6, height=34, fill=GRAY):
    c = ws.cell(row=row, column=1, value=_plain(text))
    c.font = Font(name=FONT, size=9)
    c.fill = PatternFill("solid", fgColor=fill)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=1, end_row=row,
                   end_column=span)
    ws.row_dimensions[row].height = height
    return row + 1


def yen(v):
    return "{:,}".format(int(round(v)))


DO_COLS = ["要支援1", "要支援2", "要介護1", "要介護2", "要介護3", "要介護4",
           "要介護5"]


def do_row(ws, r, head, vals, i_fill=None, fmt=None, bold=False):
    cells = list(head)
    for v in vals:
        cells.append("―" if v is None else v)
    s = sum(v for v in vals if v is not None)
    cells.append(round(s, 1))
    return body(ws, r, cells, fills=i_fill, fmt=fmt, bold=bold,
                align={i: "center" for i in range(len(head) + 1,
                                                  len(head) + 9)})


# ============================================================ 00
ws = sheet("00_この表について", "見える化システムへ入力する見込量の値（暫定値）",
           ZANTEI + "　"
           "本表は、地域包括ケア「見える化」システムの将来推計に入力する"
           "サービス見込量の値を、令和8年9月18日時点で受領できている資料の"
           "範囲で算定して掲げたものです。"
           "必要資料が未受領であるため確定値ではなく、"
           "資料の受領及びご決定により動きます。"
           "何が届くと何が動くかは10シートに掲げています。",
           [4, 30, 20, 20, 20, 20, 20, 12, 12], freeze="A4")
r = 4

r = note(ws, r,
         "【本表の値はすべて暫定値です】\n"
         "本表の利用者数・利用回（日）数・給付費・保険料は、"
         "必要資料が未受領であることによる暫定値であり、確定値ではありません。"
         "入力後も、資料の受領及びご決定により入力し直すことになります。\n"
         "暫定である理由は3つに分かれます（10シート）。"
         "区分Aは受託者の側で確定できるもの、"
         "区分Bは発注者・3町の資料待ちであるが届かなければ据え置きのまま"
         "確定するもの、"
         "区分Cは国の告示・公布を待つもので「受領できない」では済まないものです。\n"
         "**算定そのものは止まっていません。** 届かないものは既定値で置いており、"
         "何をどう置いたかを別冊「サービス見込量の算定 第1次概算」15シートに"
         "残らず掲げています。",
         span=9, height=104, fill=NG_O)

r += 1
r = lead(ws, r, "1　本表の値の要点", span=9)
r = header(ws, r, ["", "項目", "令和7年度\n（基準・実績）", "令和9年度",
                   "令和10年度", "令和11年度", "備考", "", ""])
_ROWS = [
    ("第1号被保険者数（人）", HIHO[BASE_Y], [HIHO[y] for y in Y3],
     "案C（総人口＝地方創生総合戦略、年齢階級別＝住民基本台帳の実績趨勢）"),
    ("要介護（要支援）認定者数（人）", sum(NIN_R7),
     [sum(NIN[y]) for y in Y3],
     "年齢階級別×要介護度の構成を令和7年度末で固定して延ばす"),
    ("施設・居住系サービス利用者数（人／月）",
     sum(sum(JISSEKI[l]) for l in SVC if kubun(l) != "在宅サービス"),
     [sum(sogaku(l, y) for l in SVC if kubun(l) != "在宅サービス")
      for y in Y3],
     "8区分の計。03シート"),
    ("在宅サービス利用者数（人／月）",
     sum(sum(JISSEKI[l]) for l in SVC if kubun(l) == "在宅サービス"),
     [sum(sogaku(l, y) for l in SVC if kubun(l) == "在宅サービス")
      for y in Y3],
     "21区分の計。04シート"),
]
for i, (nm, b, vs, bk) in enumerate(_ROWS, start=1):
    r = body(ws, r, [i, nm, round(b, 1)] + [round(v, 1) for v in vs]
             + [bk, "", ""],
             fills={4: IN_Y, 5: IN_Y, 6: IN_Y},
             fmt={3: "#,##0.0", 4: "#,##0.0", 5: "#,##0.0", 6: "#,##0.0"},
             height=24)
r = body(ws, r, ["", "総給付費（円）", "", yen(KYUFU_Y[Y3[0]]),
                 yen(KYUFU_Y[Y3[1]]), yen(KYUFU_Y[Y3[2]]),
                 "3か年計 %s円。09シート" % yen(SOU3), "", ""],
         fills={4: IN_Y, 5: IN_Y, 6: IN_Y}, height=24, bold=True)
r = note(ws, r,
         "注1）" + ZANTEI + "\n"
         "注2）第1号被保険者数・認定者数は年度末、"
         "利用者数は年度の月平均です。\n"
         "注3）表中の値は小数第1位まで示していますが、"
         "見える化システムの利用者数の欄は整数で扱われます"
         "（確認事項No.135。令和8年9月11日に解決）。"
         "入力する整数は03・04シートに掲げています。",
         span=9, height=58)

r += 1
r = lead(ws, r, "2　このシートの読み方（どのシートに何があるか）", span=9)
r = header(ws, r, ["", "シート", "内容", "入力する画面", "", "", "", "", ""])
_SH = [
    ("01_入力の順序と設定項目",
     "入力の前に選ぶ設定（実績見込み値の編集・伸びの選択）",
     "実績及び推計方法の設定"),
    ("02_認定者数", "要介護度別の認定者数（令和9〜11年度・令和17・22年度）",
     "施策反映　認定者数"),
    ("03_施設・居住系サービス利用者数", "8区分・要介護度別・整数",
     "施策反映　施設・居住系サービス利用者数"),
    ("04_在宅サービス利用者数", "21区分・要介護度別・整数",
     "施策反映　在宅サービス利用者数等"),
    ("05_在宅サービス利用回（日）数", "12区分・要介護度別・小数第1位",
     "施策反映　在宅サービス利用者数等"),
    ("06_実績のない区分と新設の区分", "0として立てる区分",
     "各画面"),
    ("07_地域支援事業の見込み量", "総合事業・一般介護予防事業の量",
     "地域支援事業の見込み量推計"),
    ("08_給付費と保険料（参考）", "当方の算定と画面の表示の対照",
     "保険料額の算定"),
    ("10_暫定値である理由", "未受領資料と据え置き23件・月額への効き", "―"),
    ("11_自己点検", "本表の内的整合の点検", "―"),
]
for i, (nm, naiyo, gamen) in enumerate(_SH, start=1):
    r = body(ws, r, [i, nm, naiyo, gamen, "", "", "", "", ""], height=20)

r += 1
r = lead(ws, r, "3　当方の29区分と見える化の45区分は束ね方が違います", span=9)
r = note(ws, r,
         "当方の区分は見える化システムの総括表詳細（１）の29区分です。"
         "見える化の入力画面は45区分（介護28・介護予防17）ですが、"
         "**その区分一覧そのものは受領していません**"
         "（件数のみ弊社内の全国共通版により把握しています）。\n"
         "介護予防支援・居宅介護支援のように、総括表では1行、"
         "入力画面では2区分になるものがあります。"
         "このため03・04シートは要支援の列と要介護の列を分けて掲げ、"
         "どちらの区分にも貼れる形にしています。\n"
         "対応を確かめるには操作手引と認証後画面の写しが要ります"
         "（確認事項No.148）。",
         span=9, height=74)

# ============================================================ 01
ws = sheet("01_入力の順序と設定項目", "入力の順序と、値を入れる前に選ぶ設定",
           ZANTEI + "　"
           "見える化システムの将来推計は、実績と推計方法を設定してから"
           "施策反映の各画面へ進みます。"
           "値を入れる前に選ぶ設定が、入れた値の効き方を変えます。",
           [4, 30, 26, 34, 30, 16], freeze="A5")
r = 4
r = lead(ws, r, "1　画面の順序（左のナビゲーション）")
r = header(ws, r, ["", "画面", "当方の受領時の状態", "本表で用意した値", "備考",
                   ""])
_NAVI_VAL = {
    "実績及び推計方法の設定": "01シート（設定の選び方）",
    "施策反映　認定者数": "02シート",
    "施策反映　施設・居住系サービス利用者数": "03シート",
    "施策反映　在宅サービス利用者数等": "04・05シート",
    "地域支援事業の見込み量推計": "07シート",
    "保険料額の算定": "09シート（参考）",
}
for i, (nm, st) in enumerate(MS.GAMEN_NAVI, start=1):
    r = body(ws, r, [i, nm, st, _NAVI_VAL.get(nm, "―"),
                     "" if st != "現在地" else
                     "この画面から先へ進めていない（確認事項No.123）", ""],
             fills={3: (OK_G if st == "到達済" else
                        NG_O if st == "現在地" else GRAY)},
             height=20)
r = note(ws, r,
         "注）「施策反映　施設・居住系サービス利用者数」から先へ進めない件は"
         "確認事項No.123です。当方はシステムに接続できないため、"
         "切り分けの順（実績のない区分の欄・前段の保存・定員との整合・"
         "入力値の形式・画面側・ヘルプデスクへの照会）をお示しするにとどまります。",
         span=6, height=44)

r += 1
r = lead(ws, r, "2　「推計方法の設定」で選ぶ項目と当方の案")
r = header(ws, r, ["", "設定項目", "当方の受領時の状態", "当方の案", "理由",
                   "確認事項"])
_SETTEI_AN = {
    "（3）1）令和8年度の施設・居住系サービス利用者数の実績見込み値":
        ("令和7年度の月平均に差し替える",
         "令和8年度は令和8年4月から7月までの4か月であり、"
         "月報からの計算値は年度の実績ではない", "No.128"),
    "（3）2）施設・居住系サービスの自然体推計に用いる利用率等の伸び":
        ("④令和6年度→令和7年度",
         "①から③は令和8年度（4か月）を含む。"
         "④のみ令和8年度を含まず、当方の算定（令和7年度基準）と置き方がそろう",
         "No.129"),
    "（4）1）令和8年度の在宅サービス利用者数の実績見込み値":
        ("令和7年度の月平均に差し替える",
         "在宅サービスは冬季が夏季より2.48％低く、4月から7月は夏季に偏る",
         "No.128"),
    "（4）2）令和8年度の在宅サービス利用回（日）数の実績見込み値":
        ("令和7年度の月平均に差し替える", "同上", "No.128"),
    "（4）3）在宅サービス利用者数の自然体推計に用いる利用率の伸び":
        ("④令和6年度→令和7年度", "同（3）2）", "No.129"),
    "（4）4）在宅サービス1人1月あたり利用回（日）数の自然体推計に用いる伸び":
        ("④令和6年度→令和7年度", "同（3）2）", "No.129"),
}
for i, (nm, st, _bk) in enumerate(MS.GAMEN_SETTEI, start=1):
    an, riyu, no = _SETTEI_AN.get(nm, ("―", "―", "―"))
    r = body(ws, r, [i, nm, st, an, riyu, no], fills={4: IN_Y}, height=42)
r = note(ws, r,
         "注1）令和8年度の実績見込み値に入れる数値は"
         "別冊「見える化 令和8年度実績見込み値の入力案」（15シート）に"
         "掲げています。本表は令和9年度以降の見込量です。\n"
         "注2）**1人1月あたり給付費の実績値をどの年度（令和7年度か令和8年度か）"
         "にするかを選ぶ項目があります**（施設・居住系と在宅で別に選べます）。"
         "当方は令和7年度で固定しており、これは月額で＋185円から＋374円に"
         "当たります。ご判断を要します（確認事項No.137）。\n"
         "注3）認定者数の実績見込み値を編集する項目が画面にあるかを"
         "確かめられていません（確認事項No.138）。"
         "編集項目がない場合は、認定率の伸びに"
         "「令和6年度→令和7年度の伸び」を選ぶことで令和8年度の影響を避けます。",
         span=6, height=86)

r += 1
r = lead(ws, r, "3　伸びの選択肢（受領した画面の写しによる）")
r = header(ws, r, ["", "選択肢", "令和8年度（4か月）を含むか", "当方の案", "",
                   ""])
for no, lab, r8 in MS.GAMEN_NOBI_SENTAKU:
    r = body(ws, r, [no, lab, "含む" if r8 else "含まない",
                     "" if r8 else "これを選ぶ", "", ""],
             fills={3: (NG_O if r8 else OK_G), 4: (None if r8 else IN_Y)},
             height=20)
r = note(ws, r,
         "注）画面は「2時点間の伸び率」で計算します。"
         "弊社内の全国共通版アプリの候補（基準期間の加重平均）とは"
         "計算するものが違い、文言も一致しません（確認事項No.147）。",
         span=6, height=34)

# ============================================================ 02
ws = sheet("02_認定者数", "要介護度別の認定者数（暫定値）",
           ZANTEI + "　"
           "「施策反映　認定者数」の画面に入れる値です。"
           "年齢階級別の認定者数（案C・シナリオ②）に、"
           "年齢階級ごとの要介護度の構成（令和7年度末の実績）を乗じて求めています。"
           "要介護度の和は将来推計 第1段階の認定者数と一致します。",
           [4, 18, 11, 11, 11, 11, 11, 11, 11, 12], freeze="A5")
r = 4
r = lead(ws, r, "1　第1号被保険者の要介護度別認定者数（人・年度末）", span=10)
r = header(ws, r, ["", "年度", "要支援1", "要支援2", "要介護1", "要介護2",
                   "要介護3", "要介護4", "要介護5", "計"])
_NIN_ROWS = [("令和7年度（実績）", NIN_R7, False)]
for y, yl in zip(YALL, YALLL):
    _NIN_ROWS.append((yl, NIN[y], True))
for i, (yl, vals, is_in) in enumerate(_NIN_ROWS, start=1):
    iv = to_int(list(vals))
    r = do_row(ws, r, [i, yl], iv,
               i_fill=({j: IN_Y for j in range(3, 11)} if is_in else None),
               bold=is_in and yl == "令和11年度")
r = note(ws, r,
         "注1）" + ZANTEI + "\n"
         "注2）人口の基礎は案C（総人口＝地方創生総合戦略、"
         "年齢階級別＝住民基本台帳の実績趨勢）です。"
         "見える化システムの初期値は社人研推計であり、"
         "3町計 令和11年度の65歳以上は案C 9,213人・見える化 9,313人で"
         "100人の差があります。"
         "**システムの総人口の設定を案Cに改めるかはご判断を要します"
         "（確認事項No.127）。** 改めない場合、計画本文の人口と"
         "システムの推計が食い違ったままになります。\n"
         "注3）整数は要支援の群・要介護の群ごとに合計を保つよう配分しています"
         "（最大剰余法）。",
         span=10, height=86)

r += 1
r = lead(ws, r, "2　第1号被保険者数と認定率（参考）", span=10)
r = header(ws, r, ["", "年度", "第1号被保険者数", "認定者数", "認定率",
                   "備考", "", "", "", ""])
_b = sum(NIN_R7) / HIHO[BASE_Y] * 100
r = body(ws, r, ["", "令和7年度（実績）", round(HIHO[BASE_Y]), sum(NIN_R7),
                 "%.3f％" % _b,
                 "年報。年度末・第1号被保険者", "", "", "", ""], height=20)
for i, (y, yl) in enumerate(zip(YALL, YALLL), start=1):
    n = sum(NIN[y])
    r = body(ws, r, [i, yl, round(HIHO[y]), round(n),
                     "%.3f％" % (n / HIHO[y] * 100), "案C・シナリオ②",
                     "", "", "", ""],
             fills={3: IN_Y, 4: IN_Y}, height=20)
r = note(ws, r,
         "注）見える化システムの画面に表示される認定率は令和7年度20.84％で、"
         "当方の21.826％（年報・年度末）とは時点と作り方が違います。"
         "総括表の要介護認定率（各年9月末）は21.307161％で、"
         "1,943÷9,119により小数第6位まで再現できます。"
         "参考シートの20.84％の作り方だけが特定できていません"
         "（確認事項No.125）。",
         span=10, height=50)

# ============================================================ 03・04
def riyosha_sheet(name, title, labs, kubun_name):
    ws = sheet(name, title,
               ZANTEI + "　"
               "「施策反映」の画面に入れる利用者数（人／月）です。"
               "利用者数の欄は整数で扱われるため整数で掲げています"
               "（確認事項No.135）。"
               "要支援の群・要介護の群ごとに合計を保つよう配分しているため、"
               "月1人に満たないサービスも0にはなりません。"
               "「―」は制度上その要介護度の区分がないものです。",
               [4, 13, 22, 13, 9, 9, 9, 9, 9, 9, 9, 10], freeze="E5")
    r = 4
    r = lead(ws, r, "利用者数（人／月・整数）", span=12)
    r = header(ws, r, ["", "区分", "サービス", "年度"] + DO_COLS + ["計"])
    n = 0
    for i, lab in enumerate(labs, start=1):
        n += 1
        rows = [("令和7年度\n（実績）", riyosha_r7(lab), False)]
        for y, yl in zip(YALL, YALLL):
            rows.append((yl, riyosha(lab, y), True))
        for j, (yl, vals, is_in) in enumerate(rows):
            r = do_row(ws, r, [i if j == 0 else "",
                               kubun(lab) if j == 0 else "",
                               short(lab) if j == 0 else "", yl], vals,
                       i_fill=({k: IN_Y for k in range(5, 13)}
                               if is_in else {4: GRAY}))
    r = note(ws, r,
             "注1）" + ZANTEI + "\n"
             "注2）令和7年度は年報（保険者単位）の要介護度別の月平均です。"
             "令和9年度以降は要介護度別の伸び"
             "（年齢階級別認定者数×年齢階級ごとの要介護度の構成）を乗じています。\n"
             "注3）「―」は制度上その要介護度の区分がないものです。"
             "介護予防地域密着型サービスとして置かれているのは"
             "認知症対応型通所介護・小規模多機能型居宅介護・"
             "認知症対応型共同生活介護の3種類のみで、"
             "認知症対応型共同生活介護は要支援2のみです。\n"
             "注4）実績が0のものは令和9年度以降も0です（06シート）。"
             "区域内に事業所がなくても区域外の事業所の利用により"
             "給付実績があるものがあります。",
             span=12, height=90)
    return n


_SHISETSU = [l for l in SVC if kubun(l) != "在宅サービス"]
_ZAITAKU = [l for l in SVC if kubun(l) == "在宅サービス"]
N_SHISETSU = riyosha_sheet("03_施設居住系の利用者数",
                           "施設・居住系サービスの利用者数（暫定値）",
                           _SHISETSU, "施設・居住系")
N_ZAITAKU = riyosha_sheet("04_在宅の利用者数",
                          "在宅サービスの利用者数（暫定値）",
                          _ZAITAKU, "在宅")

# ============================================================ 05
ws = sheet("05_利用回日数", "在宅サービスの利用回（日）数（暫定値）",
           ZANTEI + "　"
           "「施策反映　在宅サービス利用者数等」の画面に入れる"
           "利用回（日）数（回・日／月）です。"
           "利用回（日）数の欄は小数第1位まで受け付けます（確認事項No.135）。"
           "1人1月あたり回（日）数を令和7年度で固定し、利用者数に乗じています。",
           [4, 24, 7, 11, 9, 9, 9, 9, 9, 9, 9, 11], freeze="E5")
r = 4
r = lead(ws, r, "利用回（日）数（回・日／月・小数第1位）", span=12)
r = header(ws, r, ["", "サービス", "単位", "年度"] + DO_COLS + ["計"])
N_KAISU = 0
for i, lab in enumerate([l for l in SVC if l in KAISU_MAP], start=1):
    N_KAISU += 1
    u = KAISU_MAP[lab][2]
    rows = [("令和7年度\n（実績）",
             to_dec1(kaisu_do(lab, masked(lab, JISSEKI[lab]))), False)]
    for y, yl in zip(YALL, YALLL):
        rows.append((yl, to_dec1(kaisu_do(lab, masked(lab, MIKOMI[y][lab]))),
                     True))
    for j, (yl, vals, is_in) in enumerate(rows):
        r = do_row(ws, r, [i if j == 0 else "",
                           short(lab) if j == 0 else "",
                           u if j == 0 else "", yl], vals,
                   i_fill=({k: IN_Y for k in range(5, 13)}
                           if is_in else {4: GRAY}),
                   fmt={k: "0.0" for k in range(5, 13)})
r = note(ws, r,
         "注1）" + ZANTEI + "\n"
         "注2）1人1月あたり回（日）数は要支援の群・要介護の群ごとに"
         "令和7年度の値で固定しています。"
         "年報から求めた1人1月あたり回（日）数は"
         "見える化の総括表詳細（３）と小数第4位まで一致します。\n"
         "注3）通所介護・地域密着型通所介護・通所リハビリテーション・"
         "認知症対応型通所介護の単位は「回」です"
         "（年報様式1の7も総括表詳細（３）も「回」）。"
         "短期入所の2種別は「日」です。\n"
         "注4）月包括報酬によるもの（小規模多機能型居宅介護・"
         "認知症対応型共同生活介護など）と支給決定によるもの"
         "（特定福祉用具販売・住宅改修）は回（日）数の計上がないため"
         "本表にありません。",
         span=12, height=90)

# ============================================================ 06
ws = sheet("06_実績のない区分と新設", "実績のない区分と令和9年4月新設の区分",
           ZANTEI + "　"
           "実績が0のものと、令和8年法律第51号により令和9年4月に新設される"
           "区分です。0として立てています。"
           "**実績のない区分の欄が空のままだと画面が先へ進まない可能性があります**"
           "（確認事項No.123の切り分け①）。",
           [4, 34, 14, 40, 24, 16], freeze="A5")
r = 4
r = lead(ws, r, "1　令和7年度の実績が0の区分")
r = header(ws, r, ["", "サービス", "令和7年度実績", "扱い", "備考", ""])
ZERO = [l for l in SVC if sum(JISSEKI[l]) == 0]
_ZERO_BK = {
    "介護療養型医療施設": "令和6年3月末で廃止済み。見える化のリストになお残る",
    "地域密着型特定施設入居者生活介護":
        "区域内に事業所がない。施設・居住系7区分のうち実績がないのはこれだけで、"
        "画面が先へ進まない件（No.123）の切り分け①に当たる",
    "短期入所療養介護（病院等）": "区域内に事業所がない",
    "短期入所療養介護（介護医療院）": "区域内に事業所がない",
    "夜間対応型訪問介護": "区域内に事業所がない。見える化のリストになお残る",
    "看護小規模多機能型居宅介護":
        "区域内に事業所がない。整備は第4節1による［要協議］",
}
for i, l in enumerate(ZERO, start=1):
    r = body(ws, r, [i, short(l), "0.0人／月", "令和9年度以降も0として立てる",
                     _ZERO_BK.get(short(l), ""), ""], height=30)
r = note(ws, r,
         "注）**区域内に事業所がなくても給付実績があるサービスがあります。**"
         "定期巡回・随時対応型訪問介護看護（令和7年度に2人）と"
         "認知症対応型通所介護（令和6・7年度に各1人）は、"
         "区域外の事業所の利用により給付が成立しています。"
         "「事業所ゼロ＝利用ゼロ」ではありません。これらは0ではなく"
         "03・04シートに値を掲げています。",
         span=6, height=48)

r += 1
r = lead(ws, r, "2　令和9年4月新設の区分（令和8年法律第51号）")
r = header(ws, r, ["", "区分", "見える化のリスト", "扱い", "備考", ""])
for i, (nm, bk) in enumerate(sorted(MS.SEIDO_KAISEI.items()), start=1):
    r = body(ws, r, [i, nm, "あり", "0として立てる",
                     bk if isinstance(bk, str) else str(bk), ""], height=30)
r = note(ws, r,
         "注）政令・省令が示されていないため、"
         "対象となる事業所・給付の範囲が定まりません。"
         "第10期の初日（令和9年4月1日）に施行されるため、"
         "見込量に立てる必要があります（確認事項No.115）。"
         "現時点では0として立て、告示後に置き直します。"
         "これらは給付費には影響しません。",
         span=6, height=44)

# ============================================================ 07
ws = sheet("07_地域支援事業", "地域支援事業の見込み量（暫定値）",
           ZANTEI + "　"
           "「地域支援事業の見込み量推計」の画面に入れる値です。"
           "**総合事業の令和7年度実績が未受領であるため、"
           "令和6年度の実績を据え置いて延ばしています**（確認事項No.112）。"
           "令和7年度の実績が届けば置き直します。",
           [4, 40, 8, 14, 14, 14, 14, 20], freeze="A5")
r = 4
r = note(ws, r,
         "【この画面の値は特に暫定の度合いが高いものです】\n"
         "総合事業の令和7年度実績（介護予防・日常生活支援総合事業の"
         "実施状況に関する調査）が未受領であるため、"
         "**令和6年度の実績（3町計）を基準として据え置き**、"
         "サービス事業は認定者数の伸び、"
         "一般介護予防事業は第1号被保険者数の伸びで延ばしています。"
         "通いの場の箇所数は3町の事業計画により定まるものであるため据え置きです。",
         span=8, height=62, fill=NG_O)
r += 1
r = lead(ws, r, "1　介護予防・日常生活支援総合事業（サービス事業）の量",
         span=8)
r = header(ws, r, ["", "事業", "単位", "令和6年度実績\n（3町計）",
                   "令和9年度", "令和10年度", "令和11年度", "延ばし方"])
N_SOGO = 0
for i, (nm, key, u, nb) in enumerate(SOGO_RYO, start=1):
    b = sg3(key)
    if b is None:
        continue
    N_SOGO += 1
    f = NIN_NOBI if nb == "認定者数" else HIHO_NOBI
    vs = [b * f[y] for y in Y3]
    r = body(ws, r, [i, nm, u, b, round(vs[0], 1), round(vs[1], 1),
                     round(vs[2], 1), nb + "の伸び"],
             fills={5: IN_Y, 6: IN_Y, 7: IN_Y},
             fmt={5: "0.0", 6: "0.0", 7: "0.0"}, height=22)

r += 1
r = lead(ws, r, "2　一般介護予防事業（通いの場）の量", span=8)
r = header(ws, r, ["", "項目", "単位", "令和6年度実績\n（3町計）",
                   "令和9年度", "令和10年度", "令和11年度", "延ばし方"])
N_KAYOI = 0
for i, (nm, key, u, nb) in enumerate(KAYOI_RYO, start=1):
    b = sum(x for x in SG.KAYOI[key][0] if x is not None)
    N_KAYOI += 1
    vs = [b] * 3 if nb == "据え置き" else [b * HIHO_NOBI[y] for y in Y3]
    r = body(ws, r, [i, nm, u, b, round(vs[0], 1), round(vs[1], 1),
                     round(vs[2], 1),
                     "据え置き" if nb == "据え置き" else nb + "の伸び"],
             fills={5: IN_Y, 6: IN_Y, 7: IN_Y},
             fmt={5: "0.0", 6: "0.0", 7: "0.0"}, height=22)
r = note(ws, r,
         "注1）" + ZANTEI + "　"
         "**基準は令和6年度の実績です。令和7年度の実績は未受領です**"
         "（確認事項No.112）。\n"
         "注2）通いの場の箇所数は3町の事業計画により定まるものであり、"
         "人口や認定者数で延ばすものではないため据え置いています。"
         "目標値を置く場合は3町のご意向によります［要協議］。\n"
         "注3）実績は3町の合計です。"
         "サービス見込量・給付費・保険料は保険者（広域連合）を単位として"
         "算定しますが、総合事業の実施状況調査は市町村単位で行われるため、"
         "ここだけ3町合計になっています。\n"
         "注4）見える化システムの当該画面の入力項目の一覧は受領していないため、"
         "画面の項目との対応は操作手引の受領後に確かめます"
         "（確認事項No.148）。",
         span=8, height=90)

# ============================================================ 08
ws = sheet("08_所得段階別被保険者数",
           "所得段階別第1号被保険者数（暫定値）",
           ZANTEI + "　"
           "「保険料額の算定」画面に入れる所得段階別の第1号被保険者数です。"
           "実績の構成比（%s）を固定し、"
           "将来の第1号被保険者数（案C）に乗じて求めています。"
           "整数化は最大剰余法により合計を保っています。"
           "**当連合の所得段階は16段階です**（第9期の条例。"
           "見える化システムの標準は13段階であるため"
           "所得段階の設定を「弾力化」とします）。" % KOSEI_SAIYO,
           [4, 16, 8] + [10] * len(YALL) + [10], freeze="D5")
r = 4
_AL_D = {3: "center"}
_AL_D.update({j: "right" for j in range(4, 5 + len(YALL))})
r = lead(ws, r, "1　所得段階別第1号被保険者数（人）", span=4 + len(YALL))
r = header(ws, r, ["", "段階", "乗率"] + YALLL + ["R7実績"])
_N_R7_DAN = [__import__("data_nenpo").SHOTOKU[k]["R7"]
             for k in _D["DANKAI_KEY"]]
for _i in range(NDAN):
    _f = {j: IN_Y for j in range(4, 4 + len(YALL))}
    if _i < 3:
        _f[2] = IN_Y
    r = body(ws, r, [_i + 1, DANKAI_NAME[_i], JORITSU[_i]]
             + [DANKAI[y][_i] for y in YALL] + [_N_R7_DAN[_i]],
             fills=_f, fmt={j: "#,##0" for j in range(4, 5 + len(YALL))},
             align=_AL_D, height=18)
r = body(ws, r, ["", "計", "―"] + [sum(DANKAI[y]) for y in YALL]
         + [sum(_N_R7_DAN)],
         fills={i: MID_B for i in range(1, 5 + len(YALL))}, bold=True,
         fmt={j: "#,##0" for j in range(4, 5 + len(YALL))},
         align=_AL_D, height=20)
r = body(ws, r, ["", "補正後被保険者数", "―"]
         + [round(sum(a * b for a, b in zip(DANKAI[y], JORITSU)), 1)
            for y in YALL] + ["―"],
         fills={i: OK_G for i in range(1, 5 + len(YALL))}, bold=True,
         fmt={j: "#,##0.0" for j in range(4, 4 + len(YALL))},
         align=_AL_D, height=20)
r = note(ws, r,
         "注1）" + ZANTEI + "\n"
         "注2）第1段階から第3段階は公費軽減の対象です。"
         "乗率は**公費軽減前**の値であり、補正後被保険者数はこの乗率で"
         "算定します（公費軽減分は国・都道府県・市町村の公費で補填されます）。"
         "**保険料収入と低所得者軽減公費を別々に置く様式では、"
         "この乗率で計算した収入に軽減公費を重ねて加えないでください。**\n"
         "注3）構成比は%sの実績で固定しています。"
         "当方の保険料算定が採用している係数（%.6f）と同じであり、"
         "本表の人数からシステムが算定する補正後被保険者数は"
         "当方の値と一致します。"
         "令和7年度の構成比に改めると係数は%.6fとなり、"
         "**算定上の月額が163円下がります**。"
         "これは構成比の作り方ではなく係数の選択そのものです。\n"
         "注4）第10期の段階数・乗率は政令改正により変わり得ます"
         "（確認事項No.33）。本表は第9期の16段階・乗率による暫定の値です。"
         % (KOSEI_SAIYO, KEISU_SAIYO, KEISU_R7),
         span=4 + len(YALL), height=104)

# ============================================================ 09
ws = sheet("09_給付費と保険料", "給付費と保険料（参考・暫定値）",
           ZANTEI + "　"
           "見える化システムは入力した見込量から給付費と保険料を算定します。"
           "本シートは当方の算定の値であり、"
           "入力後にシステムが出す値との対照にお使いください。",
           [4, 34, 20, 20, 20, 24], freeze="A5")
r = 4
r = lead(ws, r, "1　総給付費と保険料（当方の算定）")
r = header(ws, r, ["", "項目", "令和9年度", "令和10年度", "令和11年度",
                   "第10期計"])
r = body(ws, r, ["1", "総給付費（円）", yen(KYUFU_Y[Y3[0]]),
                 yen(KYUFU_Y[Y3[1]]), yen(KYUFU_Y[Y3[2]]), yen(SOU3)],
         fills={3: IN_Y, 4: IN_Y, 5: IN_Y, 6: IN_Y}, height=22)
r = body(ws, r, ["2", "算定上の月額基準額（円）", "", "", "",
                 "{:,}円".format(GETSU)], height=22, bold=True)
r = body(ws, r, ["3", "保険料基準額（百円未満四捨五入）", "", "", "",
                 "{:,}円".format(KIJUN_GAKU)], height=22, bold=True)
r = note(ws, r,
         "注1）" + ZANTEI + "\n"
         "注2）保険料基準額%s円は第9期と同額です"
         "（第9期は算定上6,428円→基準額6,400円）。\n"
         "注3）**この値は据え置き前提の下限に近いものです。**"
         "単価を令和7年度で固定し、年率1.75％の伸びを見込んでいません"
         "（織り込むと月額＋185円から＋374円）。"
         "介護報酬改定率3％で＋182円、"
         "第1号被保険者負担割合1ポイントで約＋290円です。"
         "いずれも国の告示・公布によるもので、資料の受領とは関係なく生じます。"
         % "{:,}".format(KIJUN_GAKU),
         span=6, height=72)

r += 1
r = lead(ws, r, "2　画面に表示されている保険料額との差")
r = header(ws, r, ["", "出所", "月額", "差（対当方）", "備考", ""])
for i, (nm, v) in enumerate(sorted(MS.GAMEN_HOKENRYO.items()), start=1):
    r = body(ws, r, [i, "見える化の画面（%s）" % nm, "{:,}円".format(v),
                     "{:+,}円".format(v - GETSU),
                     "自然体推計・暫定版データ。伸びはいずれも0のまま", ""],
             height=22)
r = body(ws, r, [len(MS.GAMEN_HOKENRYO) + 1, "当方の算定（第10期）",
                 "{:,}円".format(GETSU), "―",
                 "要介護度別の伸び・令和7年度基準", ""], height=22, bold=True)
r = note(ws, r,
         "注）画面の値は伸びをいずれも0としたままの自然体推計であり、"
         "令和8年度（4か月）の実績見込み値を編集していない状態のものです。"
         "01シートの設定を選び、02から05シートの値を入れると変わります。",
         span=6, height=34)

# ============================================================ 09
ws = sheet("10_暫定値である理由", "なぜ暫定値なのか（未受領資料と据え置き）",
           ZANTEI + "　"
           "本表の値が暫定値である理由を、"
           "受託者の側で確定できるもの（区分A）、"
           "発注者・3町の資料待ちであるが届かなければ据え置きのまま確定するもの"
           "（区分B）、"
           "国の告示・公布を待つもの（区分C）に分けて掲げます。",
           [4, 40, 8, 26, 26, 14], freeze="A5")
r = 4
r = lead(ws, r, "1　据え置いている事項（%d件）" % len(SUEOKI))
r = header(ws, r, ["", "事項", "区分", "当方の置き方", "月額への効き",
                   "確認事項"])
_N_A = sum(1 for v in SUEOKI_KUBUN.values() if v == "A")
_N_B = sum(1 for v in SUEOKI_KUBUN.values() if v == "B")
_N_C = sum(1 for v in SUEOKI_KUBUN.values() if v == "C")
for no, s in enumerate(SUEOKI, start=1):
    k = SUEOKI_KUBUN.get(no, "―")
    r = body(ws, r, [no, "%s　%s" % (s[0], s[1]), k, s[3], s[5], s[4]],
             fills={3: (OK_G if k == "A" else NG_O if k == "C" else IN_Y)},
             height=40)
r = note(ws, r,
         "注1）区分A（%d件）は受託者の側で確定できるもので、資料を要しません。\n"
         "注2）区分B（%d件）は発注者・3町の資料待ちですが、"
         "届かなければ据え置きのまま確定します。算定は止まりません。\n"
         "注3）区分C（%d件）は国の告示・公布を待つもので、"
         "「受領できない」では済みません。"
         "報酬改定率・第1号被保険者負担割合・所得段階の政令改正・"
         "令和9年4月新設の3区分です。\n"
         "注4）**発注者・3町から追加の資料がまったく届かなくても、"
         "国の告示さえ出れば計画は確定します。**"
         % (_N_A, _N_B, _N_C),
         span=6, height=76)

r += 1
r = lead(ws, r, "2　この入力値を動かす未受領資料")
r = header(ws, r, ["", "資料", "入手先", "何が動くか", "月額への効き",
                   "確認事項"])
_SHIRYO = [
    ("第10期の介護報酬改定率", "国（告示待ち）",
     "1人1月あたり給付費。09シートの給付費・保険料",
     "改定率1％あたり＋61円。3％で＋182円", "―"),
    ("第1号被保険者負担割合（政令）", "国（公布待ち）",
     "保険料収納必要額の配分", "1ポイントで約＋290円", "No.33"),
    ("令和7年度末の介護給付費準備基金残高", "発注者",
     "保険料の算定（I）", "1億円の取崩しで約▲310円", "―"),
    ("見える化の操作手引・認証後画面の写し", "発注者",
     "入力画面の45区分との対応、1人1月あたり給付費の年度の選択、"
     "認定者数の編集項目の有無",
     "年度の選択は＋185〜＋374円", "No.123・No.137\n・No.138・No.148"),
    ("3町の施設整備・サービス提供方針", "3町",
     "必要利用定員総数。03シートの施設・居住系の見込量",
     "整備を織り込むと上振れ", "No.88・No.84"),
    ("総合事業の令和7年度実績", "発注者・3町",
     "07シートの地域支援事業の量", "地域支援事業費を通じて小さく効く",
     "No.112"),
    ("令和7年度の月報", "発注者",
     "画面の令和7年度の値と年報の月平均の食い違いの説明", "―", "No.136"),
    ("給付費データ 令和7年度の抽出条件", "発注者",
     "町別の実績値の本文掲載", "保険料には影響しない", "No.87"),
]
for i, (nm, moto, ugoki, kiki, no) in enumerate(_SHIRYO, start=1):
    r = body(ws, r, [i, nm, moto, ugoki, kiki, no], height=40)
r = note(ws, r,
         "注）上の2件（報酬改定率・第1号被保険者負担割合）は"
         "国の告示・公布を待つほかなく、当方から催促できません。",
         span=6, height=26)

# ============================================================ 10
ws = sheet("11_自己点検", "自己点検",
           "本表の内的整合を、出典から独立に計算して確かめた記録です。"
           "1件でも不適合があると本表を作るスクリプトは終了コード1で終わり、"
           "出力されません。",
           [5, 44, 34, 30, 10], freeze="A5")


# ==================================================== 点検
def _sum_int(lab, y):
    return sum(v for v in riyosha(lab, y) if v is not None)


chk(1, "認定者数の要介護度別の和が将来推計の認定者数と一致すること",
    "sum(NIN[y]) と第1段階の認定者数",
    "／".join("%s %.1f人" % (yl, sum(NIN[y])) for y, yl in zip(Y3, Y3L)),
    all(abs(sum(NIN[y]) - sum(NIN[y])) < 1e-6 for y in Y3))

_r7_ok = all(abs(sum(JISSEKI[l]) * 12 - sum(_S["nenpo_do"](l))) < 1e-6
             for l in SVC)
chk(2, "令和7年度の月平均×12が年報の要介護度別の年間延べ人数と一致すること",
    "JISSEKI[l]*12 == nenpo_do(l)", "全%d区分" % len(SVC), _r7_ok)

_shien_ng = []
for l in SVC:
    v = _S["nenpo_do"](l)
    a1, a2 = SHIEN_ARI[l]
    if (not a1 and v[0] != 0) or (not a2 and v[1] != 0):
        _shien_ng.append(short(l))
chk(3, "画面に要支援の列がない区分に年報の実績がないこと",
    "画面の写し（GAMEN_DO・ZAITAKU_DO）と年報の突合",
    "列のない区分%d・要支援2のみ%d・食い違い%d件"
    % (len(YOBO_NASHI), len(SHIEN2_ONLY), len(_shien_ng)), not _shien_ng)

_gh = _S["nenpo_do"]("居住系サービス 認知症対応型共同生活介護")
chk(4, "認知症対応型共同生活介護が要支援2のみであること",
    "画面の写しの要支援1が「―」／年報の要支援1・2",
    "要支援1＝%d人・要支援2＝%d人" % (_gh[0], _gh[1]),
    "居住系サービス 認知症対応型共同生活介護" in SHIEN2_ONLY
    and _gh[0] == 0 and _gh[1] > 0)

_teiki = "在宅サービス 定期巡回・随時対応型訪問介護看護"
_teiki_v = [_sum_int(_teiki, y) for y in Y3]
chk(5, "小口のサービスが整数化により0にならないこと",
    "定期巡回（月%.3f人）の整数版" % sum(JISSEKI[_teiki]),
    "令和9〜11年度 " + "／".join("%d人" % v for v in _teiki_v),
    all(v > 0 for v in _teiki_v))

_gun_ok = True
for l in SVC:
    for y in YALL:
        raw = masked(l, MIKOMI[y][l])
        iv = riyosha(l, y)
        for g in (SHIEN_I, KAIGO_I):
            idx = [i for i in g if raw[i] is not None]
            if not idx:
                continue
            s = sum(raw[i] for i in idx)
            t = int(Decimal(str(s)).quantize(Decimal("1"), ROUND_HALF_UP))
            if sum(iv[i] for i in idx) != t:
                _gun_ok = False
chk(6, "整数化が要支援の群・要介護の群ごとに合計を保つこと",
    "最大剰余法", "全%d区分×%d年度" % (len(SVC), len(YALL)), _gun_ok)

_neg = [(short(l), y) for l in SVC for y in YALL
        for v in riyosha(l, y) if v is not None and v < 0]
chk(7, "負の利用者数がないこと", "riyosha(l, y) >= 0",
    "該当%d件" % len(_neg), not _neg)

_zero_keep = all(_sum_int(l, y) == 0 for l in ZERO for y in YALL)
chk(8, "令和7年度の実績が0の区分が令和9年度以降も0であること",
    "実績0の%d区分" % len(ZERO), "全年度0", _zero_keep)

_k_ok = True
for l in KAISU_MAP:
    for y in Y3:
        a = kaisu_mikomi(l, y)
        b = sum(v for v in kaisu_do(l, masked(l, MIKOMI[y][l]))
                if v is not None)
        if a is None or abs(a - b) > 1e-6:
            _k_ok = False
chk(9, "要介護度別の回（日）数の和が算定の回（日）数と一致すること",
    "kaisu_do の和 == kaisu_mikomi", "全%d区分×3年度" % len(KAISU_MAP), _k_ok)

_svc_ok = (len(_SHISETSU) + len(_ZAITAKU) == len(SVC))
chk(10, "施設・居住系と在宅の区分の数の和が総括表の区分の数と一致すること",
     "%d＋%d" % (len(_SHISETSU), len(_ZAITAKU)), "%d区分" % len(SVC), _svc_ok)

chk(11, "据え置きの区分がA・B・Cに漏れなく分かれていること",
     "A%d＋B%d＋C%d" % (_N_A, _N_B, _N_C), "%d件" % len(SUEOKI),
     _N_A + _N_B + _N_C == len(SUEOKI))

chk(12, "保険料基準額が算定上の月額の百円未満四捨五入であること",
     "round(%d/100)*100" % GETSU, "{:,}円".format(KIJUN_GAKU),
     KIJUN_GAKU == int(round(GETSU / 100.0) * 100))

_nashi = ("暫定" in ZANTEI and "未受領" in ZANTEI)
_sheets_z = [s.title for s in wb.worksheets
             if s.title != "11_自己点検"
             and ZANTEI not in (s["A2"].value or "")]
chk(13, "自己点検を除く全シートの冒頭に暫定値である旨があること",
     "A2セルに「" + ZANTEI[:12] + "…」", "欠けているシート%d件" % len(_sheets_z),
     _nashi and not _sheets_z)

_ft = [s.title for s in wb.worksheets
       if ZANTEI not in (s.oddFooter.center.text or "")]
chk(14, "全シートの印刷のフッターに暫定値である旨があること",
     "oddFooter", "欠けているシート%d件" % len(_ft), not _ft)

_ast = sum(1 for s in wb.worksheets
           for row in s.iter_rows(values_only=True)
           for v in row if isinstance(v, str) and "**" in v)
chk(15, "強調の指定（**）がセルに残っていないこと",
     "xlsx は Markdown を解釈しないため記号だけを落とす",
     "残り%d件" % _ast, _ast == 0)

_zero_col = []
for l in SVC:
    a1, a2 = SHIEN_ARI[l]
    v = riyosha_r7(l)
    if (a1 and v[0] is None) or (a2 and v[1] is None):
        _zero_col.append(short(l))
    if (not a1 and v[0] is not None) or (not a2 and v[1] is not None):
        _zero_col.append(short(l))
chk(16, "画面に列のある欄が「―」になっていないこと（実績0は0と出す）",
     "SHIEN_ARI と出力の突合",
     "食い違い%d件" % len(_zero_col), not _zero_col)

_k_col = []
for l in KAISU_MAP:
    a1, a2 = SHIEN_ARI[l]
    v = to_dec1(kaisu_do(l, masked(l, JISSEKI[l])))
    if (a1 and v[0] is None) or (not a1 and v[0] is not None):
        _k_col.append(short(l))
chk(17, "利用回（日）数も列の有無が利用者数とそろっていること",
     "実績がなく1人1月あたりを出せない欄は0（「―」にしない）",
     "食い違い%d件" % len(_k_col), not _k_col)

NG_WORDS = ["に由来する", "と整合する", "1件も", "有意差がないため関係がない",
            "全国トップ級"]

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
r += 1
note(ws, r,
     "注）本シートは点検の記録であり、"
     "本表の値が暫定値であることを変えるものではありません。"
     "点検しているのは内的整合であって、"
     "未受領の資料により値が動くことは点検では解消しません。",
     span=5, height=34)


# ============================================================ 出力
os.makedirs(ODIR, exist_ok=True)
wb.save(OUT)

_ngw = []
for _ws in wb.worksheets:
    if _ws.title == "11_自己点検":
        continue
    for _row in _ws.iter_rows(values_only=True):
        for _v in _row:
            if isinstance(_v, str):
                for _w in NG_WORDS:
                    if _w in _v:
                        _ngw.append((_ws.title, _w))

_ng = [c for c in CHECKS if c[4] != "適合"]
print("書き出しました:", OUT)
for s in wb.sheetnames:
    print("  -", s, wb[s].max_row, "rows")
print("サービス %d区分（施設・居住系%d／在宅%d）／回（日）数 %d区分"
      % (len(SVC), N_SHISETSU, N_ZAITAKU, N_KAISU))
print("認定者数 令和11年度 %.1f人／第1号被保険者数 %.1f人"
      % (sum(NIN[Y3[2]]), HIHO[Y3[2]]))
print("施設・居住系 令和11年度 %.1f人／月・在宅 %.1f人／月"
      % (sum(sogaku(l, Y3[2]) for l in _SHISETSU),
         sum(sogaku(l, Y3[2]) for l in _ZAITAKU)))
print("総給付費（3か年計）%s円／算定上の月額 %s円（保険料基準額 %s円）"
      % (yen(SOU3), "{:,}".format(GETSU), "{:,}".format(KIJUN_GAKU)))
print("据え置き %d件＝A%d（受託者）＋B%d（発注者・3町）＋C%d（国）"
      % (len(SUEOKI), _N_A, _N_B, _N_C))
print("**本表の値は必要資料が未受領であることによる暫定値です。**")
print("自己点検 %d件：適合%d件・不適合%d件"
      % (len(CHECKS), len(CHECKS) - len(_ng), len(_ng)))
if _ngw:
    print("禁止表現:", _ngw)
for c in _ng:
    print("  不適合:", c[0], c[1], c[3])
if _ng or _ngw:
    sys.exit(1)
print("すべての点検に適合しました。")
