# -*- coding: utf-8 -*-
"""大雪地区広域連合 第10期介護保険事業計画
見える化システムに入力する箇所と入力する値（暫定値）.

令和8年9月30日のご依頼
  「データ破損でファイルが消えてしまったので、推計を新しく入力しますので、
    先ほどから整理頂いている内容について再度取りまとめの上、
    エクセルファイルのアウトプットをお願いします。
    エクセル内容については見える化システム上、値を変更する該当箇所と
    入力する値のみご教示ください」

━━ 本表の値はすべて暫定値である ━━

必要資料が未受領であることによる暫定値であり、確定値ではない。
その旨を00シート・各シートの表題・注記・印刷のフッターに置いている。

━━ 画面ごとに入力の様式が違う ━━

令和8年9月30日のご指示
  「施設居住系、在宅、地域支援事業と入力フォーマットが異なるので
    修正をお願いします」

**3つの画面は入力の様式が違う。同じ形で並べると入力できない。**
様式（区分の別・並び順・単位・年度・量の欄の有無）は
`data_mieru_yoshiki.py` に収め、本表はそこから組み立てる。

  手順4 施設・居住系　画面の区分は（１）居宅サービス・（２）地域密着型サービス・
      （３）施設サービスの3つで、総括表の「施設サービス／居住系サービス」とは
      束ね方が違う。入れるのは利用者数（人・整数）。
      **介護療養型医療施設は画面に行がない。**
  手順5 在宅　画面の区分は（１）居宅サービス・（２）地域密着型サービス・
      （４）居宅介護支援の3つ。**施策反映は利用率によって行う。**
      利用率の分母は認定者数−施設・居住系サービス利用者数であり、
      要介護度別の列の前に「全体」の列がある。
      利用回（日）数は別画面で、入れるのは**1人1月あたり**の回（日）数であり、
      1月当たりの延べ回（日）数ではない（合計の列もない）。
  手順6・7 地域支援事業　**量（人）と事業費（円）が同じ1つの表に混在する。**
      **量の欄があるのは4項目だけ**で、ほかの事業は事業費の欄しかない。
      年度は11年度。４．の計はシステムが１．から３．の和として計算する。

━━ 本表に載せるもの・載せないもの ━━

**載せるのは、画面の値を変える箇所と、そこに入れる値だけ**である。
初期値のままでよい箇所は載せていない。
ただし、初期値のままでよいことを確かめる必要がある箇所は
03シートに「確認」として区分を分けて掲げている。

━━ 数値の出所 ━━

**固定値を書かない。** 見込量・給付費・保険料は
`build_mikomiryo_santei.py`、所得段階別は `build_shotoku_dankai.py`、
令和8年度の実績見込み値は `build_mieru_r8_input.py`、
人口の設定は `mieru_anc.py`、画面の様式は `data_mieru_yoshiki.py` を読む。
算定を改めれば本表が追随する。

シート構成（見える化システムの画面の順）
  00_この表について
  01_変更する箇所の一覧
  02_人口の設定（総人口と被保険者数の設定）
  03_設定の選択
  04_令和8年度の実績見込み値
  05_認定者数
  06_施設居住系の利用者数（画面の3区分の並び。利用者数）
  07_在宅の利用率と利用者数（画面の3区分の並び。利用率で施策反映する）
  08_在宅の利用回日数（画面の区分の並び）
  09_地域支援事業の量と事業費（様式どおりの1表。量の欄は4項目）
  10_地域支援事業費の計と入力欄のない量
  11_保険料額の算定
  12_入力後の確かめ方
  13_自己点検

出力
  output/第10期計画_見える化_入力箇所と入力値.xlsx

自己点検で1件でも不適合があると終了コード1で終わる。
"""

import io
import os
import runpy
import sys

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import data_mieru_danryoku as GD
import data_mieru_jinko as J
import data_mieru_kekka2 as K2
import data_mieru_suikei as MS
import data_mieru_yoshiki as YS
import mieru_anc as AN
import repo_paths as RP

if isinstance(sys.stdout, io.TextIOWrapper):
    sys.stdout.reconfigure(encoding="utf-8")

OUT = os.path.join(RP.OUTPUT, "第10期計画_見える化_入力箇所と入力値.xlsx")
KIJUNBI = "令和8年9月30日"

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
_NULLS = []


def _load(name):
    """算定のスクリプトを読み込む。画面への出力は捨てる。

    読み込む側は `sys.stdout.buffer` を包み直すため、
    読み込みごとに別の書き出し先を開く（使い回すと閉じられる）。
    """
    old = sys.stdout
    f = open(os.devnull, "w")
    _NULLS.append(f)
    sys.stdout = f
    try:
        return runpy.run_path(os.path.join(RP.ROOT, name))
    finally:
        sys.stdout = old


_M = _load("build_mieru_mikomi_input.py")     # 令和9年度以降の見込量
_R8 = _load("build_mieru_r8_input.py")        # 令和8年度の実績見込み値
_S = _M["_S"]                                 # サービス見込量の算定

SVC, DO_COLS = _M["SVC"], _M["DO_COLS"]
YALL, YALLL = _M["YALL"], _M["YALLL"]
Y3, Y3L = _M["Y3"], _M["Y3L"]
NIN, NIN_R7 = _M["NIN"], _M["NIN_R7"]
MIKOMI, JISSEKI = _M["MIKOMI"], _M["JISSEKI"]
KAISU_MAP = _M["KAISU_MAP"]
riyosha, riyosha_r7 = _M["riyosha"], _M["riyosha_r7"]
to_int, to_dec1 = _M["to_int"], _M["to_dec1"]
kaisu_do, masked = _M["kaisu_do"], _M["masked"]
kaisu_tanka_do = _M["kaisu_tanka_do"]
short, kubun = _M["short"], _M["kubun"]
SOGO_RYO, KAYOI_RYO = _M["SOGO_RYO"], _M["KAYOI_RYO"]
sg3, SG = _M["sg3"], _M["SG"]
NIN_NOBI, HIHO_NOBI = _M["NIN_NOBI"], _M["HIHO_NOBI"]
DANKAI, JORITSU = _M["DANKAI"], _M["JORITSU"]
DANKAI_NAME, NDAN = _M["DANKAI_NAME"], _M["NDAN"]
KIJUN_SHOTOKU = _M["KIJUN_SHOTOKU"]
G_DO = _M["G_DO"]
GETSU, KIJUN_GAKU = _M["GETSU"], _M["KIJUN_GAKU"]
SHIEN_ARI = _M["SHIEN_ARI"]

CHIIKI_R6, SOGO_R6 = _S["CHIIKI_R6"], _S["SOGO_R6"]
HOKATSU_R6 = CHIIKI_R6 - SOGO_R6

ANC, JY = AN.ANC, J.YEARS
ichigo, kouki = AN.ichigo, AN.kouki


# ============================================================ 体裁
def sheet(name, title, subtitle, widths, freeze="A5"):
    ws = wb.create_sheet(name)
    ws["A1"] = title
    ws["A1"].font = Font(name=FONT, size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    ws["A2"] = _plain("【暫定値】　" + subtitle)
    ws["A2"].font = Font(name=FONT, size=9)
    ws["A2"].fill = PatternFill("solid", fgColor=GRAY)
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    n = max(len(widths), 6)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n)
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 46
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    if freeze:
        ws.freeze_panes = freeze
    return ws


def _plain(v):
    """xlsx は Markdown を解釈しないため、書き出しの時点で ** を落とす。"""
    return v.replace("**", "") if isinstance(v, str) else v


def header(ws, row, cols, height=30):
    for j, c in enumerate(cols, start=1):
        cell = ws.cell(row=row, column=j, value=_plain(c))
        cell.font = Font(name=FONT, size=9, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=HEAD)
        cell.alignment = Alignment(wrap_text=True, vertical="center",
                                   horizontal="center")
        cell.border = BORDER
    ws.row_dimensions[row].height = height
    return row + 1


def body(ws, row, vals, fills=None, height=20, align=None, bold=False,
         fmt=None):
    for j, v in enumerate(vals, start=1):
        cell = ws.cell(row=row, column=j, value=_plain(v))
        cell.font = Font(name=FONT, size=9, bold=bold)
        cell.alignment = Alignment(
            wrap_text=True, vertical="center",
            horizontal=(align or {}).get(j, "left"))
        cell.border = BORDER
        if fills and fills.get(j):
            cell.fill = PatternFill("solid", fgColor=fills[j])
        if fmt and fmt.get(j):
            cell.number_format = fmt[j]
    ws.row_dimensions[row].height = height
    return row + 1


def lead(ws, row, text, span=6, height=20):
    cell = ws.cell(row=row, column=1, value=_plain(text))
    cell.font = Font(name=FONT, size=10, bold=True, color=NAVY)
    cell.alignment = Alignment(vertical="center")
    ws.merge_cells(start_row=row, start_column=1,
                   end_row=row, end_column=span)
    ws.row_dimensions[row].height = height
    return row + 1


def note(ws, row, text, span=6, height=34, fill=GRAY):
    cell = ws.cell(row=row, column=1, value=_plain(text))
    cell.font = Font(name=FONT, size=8)
    cell.fill = PatternFill("solid", fgColor=fill)
    cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=1,
                   end_row=row, end_column=span)
    ws.row_dimensions[row].height = height
    return row + 1


def do_row(ws, r, head, vals, i_fill=None, fmt=None, bold=False):
    """要介護度別の1行（見出し＋7区分＋計）。"""
    cells = list(head) + [("―" if v is None else v) for v in vals]
    num = [v for v in vals if v is not None]
    cells.append(round(sum(num), 1) if num else "―")
    n = len(head)
    al = {j: "right" for j in range(n + 1, n + 9)}
    return body(ws, r, cells, fills=i_fill, align=al, height=18,
                fmt=fmt or {}, bold=bold)


# ============================================================ 00
ws = sheet("00_この表について",
           "見える化システムに入力する箇所と入力する値",
           ZANTEI + "　"
           "推計を新しく作り直す場合に、"
           "画面の値を変える箇所と、そこに入れる値だけをまとめたものです。"
           "初期値のままでよい箇所は載せていません。"
           "画面の順（手順1から手順8）に並べています。",
           [4, 30, 78], freeze="A5")
r = 4
r = lead(ws, r, "1　この表の使い方", span=3)
r = header(ws, r, ["#", "項目", "内容"])
for i, (a, b) in enumerate([
        ("載せているもの",
         "見える化システムの画面のうち、初期値から値を変える箇所と、"
         "そこに入れる値です。"),
        ("載せていないもの",
         "初期値のままでよい箇所。ただし初期値のままでよいことを"
         "確かめる必要がある箇所は03シートに「確認」として掲げています。"),
        ("並べ方",
         "見える化システムの画面の順（手順1 推計の開始と保存 → 手順2 実績及び"
         "推計方法の設定 → 手順3から手順5 施策反映 → 手順6 地域支援事業の"
         "見込み量 → 手順7 地域支援事業費 → 手順8 保険料額の算定）です。"),
        ("入力欄の桁",
         "利用者数（人数）の欄は整数で扱われます。"
         "利用回（日）数の欄は小数第1位までです。"
         "本表はその桁に合わせて掲げています。"),
        ("画面ごとに様式が違うこと",
         "施設・居住系・在宅・利用回（日）数・地域支援事業は"
         "入力の様式が違います。"
         "施設・居住系は利用者数（人）、在宅は利用率（％）、"
         "利用回（日）数は1人1月あたりの回（日）数、"
         "地域支援事業は量（人）と事業費（円）です。"
         "本表は画面ごとの様式に合わせて分けています。"),
        ("同じ「回（日）数」でも画面により意味が違うこと",
         "手順5の施策反映の画面に入れるのは**1人1月あたり**の回（日）数です"
         "（08シート）。"
         "手順2の令和8年度の実績見込み値の欄に入れるのは"
         "**1月当たりの延べ**の回（日）数です（04シート）。"
         "延べの値を施策反映の画面に入れると桁が大きく違います。"),
        ("「―」の意味",
         "制度上その要介護度の区分がないもの（画面に列がないもの）です。"
         "0とは意味が違います。0は「区分はあるが計上がない」ことを表します。"),
        ("入れる値の色",
         "黄色の欄が入力する値です。灰色は実績・参考であり入力しません。"),
        ("暫定値であること",
         ZANTEI + "　必要資料が届いたとき、"
         "及び国の告示・政令が示されたときに置き直します。"),
        ("入力後の確かめ方",
         "12シートに、入力が正しく効いたかどうかを確かめる値を掲げています。"),
], start=1):
    r = body(ws, r, [i, a, b], height=36)

r += 1
r = lead(ws, r, "2　シートの構成", span=3)
r = header(ws, r, ["#", "シート", "画面（手順）"])
_SH = [
    ("01_変更する箇所の一覧", "全体"),
    ("02_人口の設定", "手順2　実績及び推計方法の設定／総人口と被保険者数の設定"),
    ("03_設定の選択", "手順2から手順5　推計方法の設定で選ぶ項目"),
    ("04_令和8年度の実績見込み値",
     "手順2　実績及び推計方法の設定／令和8年度の実績見込み値"),
    ("05_認定者数", "手順3　施策反映　認定者数"),
    ("06_施設居住系の利用者数", "手順4　施策反映　施設・居住系サービス利用者数"),
    ("07_在宅の利用率と利用者数", "手順5　施策反映　在宅サービス利用者数（利用率）"),
    ("08_在宅の利用回日数", "手順5　施策反映　在宅サービス利用回（日）数"),
    ("09_地域支援事業の量", "手順6　地域支援事業の見込み量推計"),
    ("10_地域支援事業費", "手順7　地域支援事業費"),
    ("11_保険料額の算定", "手順8　保険料額の算定"),
    ("12_入力後の確かめ方", "推計結果概要の確認"),
    ("13_自己点検", "―"),
]
for i, (a, b) in enumerate(_SH, start=1):
    r = body(ws, r, [i, a, b], height=20)
r = note(ws, r,
         "注）本表は入力する値を示すものであり、"
         "画面の操作の手引ではありません。"
         "操作の手引と認証後の画面の一覧は受領していないため、"
         "画面の項目名は受領した画面の写しによっています"
         "（確認事項No.148）。",
         span=3, height=40)

# ============================================================ 01
_HYO = {
    "02": "02_人口の設定", "03": "03_設定の選択",
    "04": "04_令和8年度の実績見込み値", "05": "05_認定者数",
    "06": "06_施設居住系の利用者数", "07": "07_在宅の利用率と利用者数",
    "08": "08_在宅の利用回日数", "09": "09_地域支援事業の量",
    "10": "10_地域支援事業費",
    "11": "11_保険料額の算定",
}

ICHIRAN = [
    ("手順1", "推計の開始と保存", "推計パターン名",
     "任意の名で保存する", "―", "―", ""),
    ("手順2", "総人口と被保険者数の設定", "人口の基礎",
     "「独自データを登録する」を選び、令和8年度以降を案Cに改める",
     "02", "入力欄165。令和6年度・令和7年度は画面の値のまま",
     "給付費と調整交付金の双方に及ぶ"),
    ("手順2", "介護保険事業状況報告の設定", "用いる報告の年度",
     "初期値のままでよいことを確かめる", "03", "―", ""),
    ("手順2", "令和8年度の実績見込み値", "認定者数",
     "令和7年度の月平均に差し替える", "04",
     "画面に編集の項目があるかを確かめる", "見込量の起点が変わる"),
    ("手順2", "令和8年度の実績見込み値", "施設・居住系サービス利用者数",
     "令和7年度の月平均に差し替える", "04", "―", "同上"),
    ("手順2", "令和8年度の実績見込み値", "在宅サービス利用者数",
     "令和7年度の月平均に差し替える", "04", "―", "同上"),
    ("手順2", "令和8年度の実績見込み値", "在宅サービス利用回（日）数",
     "令和7年度の月平均に差し替える", "04", "―",
     "ワーニング36件のうち33件が解消する見込み"),
    ("手順3", "施策反映　認定者数", "認定者数の自然体推計手法",
     "初期値（性別／年齢5歳階級別／要介護度別）のままでよい", "03", "―", ""),
    ("手順3", "施策反映　認定者数", "認定率の伸び",
     "「令和6年度→令和7年度の伸び」を選ぶ", "03",
     "令和8年度（4か月）を含まない唯一の選択肢", ""),
    ("手順3", "施策反映　認定者数", "認定者数（令和9年度以降）",
     "当方の算定値を入れる", "05", "要介護度別・全8年度", ""),
    ("手順4", "施策反映　施設・居住系", "利用率の伸び",
     "「令和6年度→令和7年度の伸び」を選ぶ", "03", "―", ""),
    ("手順4", "施策反映　施設・居住系", "1人1月あたり給付費の実績値の年度",
     "「令和7年度」に改める", "03",
     "現況は令和8年度。特定施設 要介護5のワーニング3件はこれによる",
     "月額を下げる向き（確認事項No.137）"),
    ("手順4", "施策反映　施設・居住系", "利用者数（令和9年度以降）",
     "当方の算定値を入れる", "06", "要介護度別・全8年度", ""),
    ("手順5", "施策反映　在宅", "介護サービス利用者数の自然体推計手法",
     "「要介護度別に推計する」に改める", "03",
     "現況は包括推計。当方の検証では要介護度別の方が誤差が小さい",
     "確認事項No.126"),
    ("手順5", "施策反映　在宅", "利用率の伸び・利用回（日）数の伸び",
     "いずれも「令和6年度→令和7年度の伸び」を選ぶ", "03", "―", ""),
    ("手順5", "施策反映　在宅", "1人1月あたり給付費の実績値の年度",
     "「令和7年度」に改める", "03", "施設・居住系と別に選べる",
     "月額を下げる向き（確認事項No.137）"),
    ("手順5", "施策反映　在宅", "利用率（令和9年度以降）",
     "当方の算定値を入れる", "07",
     "この画面は利用率で施策反映する。分母は認定者数−施設・居住系利用者数",
     ""),
    ("手順5", "施策反映　在宅", "利用回（日）数（令和9年度以降）",
     "当方の算定値を入れる", "08",
     "別画面。入れるのは1人1月あたり（延べではない）。小数第1位まで", ""),
    ("手順6", "地域支援事業の見込み量推計", "登録の方法",
     "「サービスごとの総数で登録する」を選ぶ", "09",
     "年齢階級別の実人数を持たないため", ""),
    ("手順6", "地域支援事業の見込み量推計", "量（利用者数）",
     "4区分のうち3区分に当方の算定値を入れる", "09",
     "画面の年度は令和6年度から令和11年度。令和7年度実績が未受領のため据え置く",
     ""),
    ("手順7", "地域支援事業費", "事業費",
     "総合事業費と包括的支援事業・任意事業費の計を入れる", "10",
     "現況は全年度0。事業別の内訳は未受領のため計のみ",
     "月額 約＋332円"),
    ("手順8", "保険料額の算定", "所得段階の設定",
     "「弾力化」を選び16段階で入れる", "11",
     "当連合の条例は16段階。標準13段階では9段階で割合が違う",
     "月額 約▲49円"),
    ("手順8", "保険料額の算定", "所得段階別第1号被保険者数",
     "当方の算定値を入れる", "11", "16段階×全8年度", ""),
    ("手順8", "保険料額の算定", "基準所得金額",
     "第7段階から第16段階の10件を入れる", "11",
     "第1段階から第6段階は入力欄がない", ""),
    ("手順8", "保険料額の算定", "基準額に対する割合",
     "16段階の割合（公費軽減前）を入れる", "11",
     "公費軽減後の値（第1段階0.285など）は入れない", ""),
    ("手順8", "保険料額の算定", "準備基金取崩額",
     "令和7年度末の基金残高の提供を受けてから入れる", "11",
     "現時点では0のまま", "1億円の取崩しで月額 約▲310円"),
    ("手順8", "保険料額の算定",
     "保険者機能強化推進交付金等の交付見込額",
     "0のままとする", "11",
     "地域支援事業費から2事業を既に除いており、併用すると二重控除になる",
     "確認事項No.86"),
    ("手順8", "保険料額の算定", "予定保険料収納率",
     "0.99であることを確かめる", "03", "―", ""),
    ("手順8", "施策反映の内容", "実績見込み値を編集した理由",
     "記入する", "03", "出力の「施策反映の内容」に載る", ""),
]

ws = sheet("01_変更する箇所の一覧",
           "画面の値を変える箇所と操作の一覧",
           ZANTEI + "　"
           "見える化システムの画面のうち、値を変える箇所を画面の順に並べたものです。"
           "入れる値は「本表のシート」の欄のシートにあります。",
           [7, 26, 30, 42, 16, 44, 30], freeze="A5")
r = 4
r = lead(ws, r, "変更する箇所（%d件）" % len(ICHIRAN), span=7)
r = header(ws, r, ["手順", "画面", "項目", "操作", "本表のシート",
                   "備考", "月額への効き・確認事項"])
for te, ga, ko, so, sh, bi, ki in ICHIRAN:
    r = body(ws, r, [te, ga, ko, so,
                     _HYO.get(sh, "―"), bi, ki],
             fills={4: IN_Y}, height=30)
r = note(ws, r,
         "注1）" + ZANTEI + "\n"
         "注2）**「初期値のままでよい」と書いた箇所も、"
         "画面を開いて確かめてください。**"
         "推計を作り直すと初期値に戻るものがあります。\n"
         "注3）月額への効きは、当方が現況の推計結果から試算したものです。"
         "実際の値はシステムが算定します。\n"
         "注4）**手順2の人口の設定を改めたら、手順3から手順8までを"
         "順に開いて登録し直してください。**"
         "人口だけを改めても、認定者数・サービス見込量・給付費は"
         "計算し直されないことを確かめています。",
         span=7, height=76)

# ============================================================ 02
ws = sheet("02_人口の設定",
           "総人口と被保険者数の設定（案C）",
           ZANTEI + "　"
           "「独自データを登録する」を選んだ場合の入力値です。"
           "**令和6年度・令和7年度は実績であるため画面の値をそのまま入れ、"
           "令和8年度以降を案Cに改めます。**"
           "白い欄（6階級と第2号被保険者）が入力欄で、"
           "第1号被保険者と総数は自動計算されます。",
           [4, 20] + [10] * len(JY), freeze="C5")
r = 4
_AL2 = {j: "right" for j in range(3, 3 + len(JY))}
_FMT2 = {j: "#,##0" for j in range(3, 3 + len(JY))}
_HD2 = ["", "区分"] + JY
_SPAN2 = 2 + len(JY)


def _in2(y):
    return {} if y in AN.JITSU else {"fill": IN_Y}


_INCOL = {j: (None if JY[j - 3] in AN.JITSU else IN_Y)
          for j in range(3, 3 + len(JY))}

r = lead(ws, r, "1　総人口（人）", span=_SPAN2)
r = header(ws, r, _HD2)
r = body(ws, r, ["", "案C（入力する値）"] + [ANC[y]["総人口"] for y in JY],
         fills=_INCOL, fmt=_FMT2, align=_AL2, height=20, bold=True)
r = body(ws, r, ["", "画面の現在値"] + [J.SOJINKO[y] for y in JY],
         fills={j: GRAY for j in range(1, 3 + len(JY))},
         fmt=_FMT2, align=_AL2, height=20)

for _sex, _fn in (("（1）男", J.dan), ("（2）女", J.jo)):
    r += 1
    r = lead(ws, r, "%s　被保険者数%s（人）"
             % ("2" if _sex.startswith("（1）") else "3", _sex),
             span=_SPAN2)
    r = header(ws, r, _HD2)
    _key = "男" if _sex.startswith("（1）") else "女"
    for a in J.AGE:
        r = body(ws, r, ["", J.AGEL[a]] + [ANC[y][_key][a] for y in JY],
                 fills=_INCOL, fmt=_FMT2, align=_AL2, height=18)
    r = body(ws, r, ["", "第2号被保険者"] + [_fn("第2号", y) for y in JY],
             fills=_INCOL, fmt=_FMT2, align=_AL2, height=18)
    r = body(ws, r, ["", "第1号被保険者（自動計算）"]
             + [sum(ANC[y][_key].values()) for y in JY],
             fills={j: GRAY for j in range(1, 3 + len(JY))},
             fmt=_FMT2, align=_AL2, height=18)

r += 1
r = lead(ws, r, "4　入力後に画面が表示する第1号被保険者数（自動計算）",
         span=_SPAN2)
r = header(ws, r, _HD2)
r = body(ws, r, ["", "案C（この値になる）"] + [ichigo(y) for y in JY],
         fills={j: OK_G for j in range(3, 3 + len(JY))},
         fmt=_FMT2, align=_AL2, height=20, bold=True)
r = body(ws, r, ["", "画面の現在値"] + [J.ichigo(y) for y in JY],
         fills={j: GRAY for j in range(1, 3 + len(JY))},
         fmt=_FMT2, align=_AL2, height=20)
r = body(ws, r, ["", "差（案C−画面）"]
         + [ichigo(y) - J.ichigo(y) for y in JY],
         fills={j: GRAY for j in range(1, 3 + len(JY))},
         fmt={j: "+#,##0;-#,##0;0" for j in range(3, 3 + len(JY))},
         align=_AL2, height=20)
r = note(ws, r,
         "注1）" + ZANTEI + "\n"
         "注2）**令和6年度・令和7年度は実績であるため画面の値をそのまま入れます**"
         "（灰色ではなく色を付けていない列）。令和8年度以降が案Cによる値です。\n"
         "注3）令和8年度の第1号被保険者数は案Cでも%s人で画面と同じであり、"
         "動くのは年齢階級別の内訳だけです。\n"
         "注4）当方の案Cは65歳以上を65〜74歳・75〜84歳・85歳以上の3区分で"
         "推計しています。画面は6階級×男女を求めるため、"
         "3区分の合計を保ったまま、"
         "画面が現に表示している補正データの同じ年の構成比で分けています。\n"
         "注5）第2号被保険者（40歳から64歳）は画面の値のままとしています。"
         "案Cは40歳から64歳の内訳を持たず、"
         "第1号被保険者の保険料の算定にも用いられません。\n"
         "注6）**後期高齢者の割合が画面より高くなります**"
         "（令和11年度 %.4f対%.4f）。調整交付金が増え保険料を下げる向きに、"
         "認定者数が増え給付費を上げる向きに働きます。"
         % ("{:,}".format(ichigo("R8")),
            kouki("R11") / ichigo("R11"),
            sum(J.kei(a, "R11") for a in J.AGE[2:]) / J.ichigo("R11")),
         span=_SPAN2, height=118)

# ============================================================ 03
SETTEI = [
    ("変更", "手順2", "介護保険事業状況報告の設定",
     "「令和6年度年報及び令和7年度・令和8年度の最新の月報」（初期値）",
     "初期値のまま", "変更しません。初期値であることを確かめてください"),
    ("変更", "手順3", "認定者数の自然体推計手法",
     "性別／年齢5歳階級別／要介護度別に推計する（初期値）",
     "初期値のまま",
     "当方の算定（年齢階級別×要介護度の構成）と置き方がそろいます"),
    ("変更", "手順3", "認定率の伸び",
     "4つの選択肢から選ぶ",
     "令和6年度→令和7年度の伸び",
     "令和8年度（4か月）を含まない唯一の選択肢です"),
    ("変更", "手順4", "施設・居住系サービスの利用率の伸び",
     "4つの選択肢から選ぶ", "令和6年度→令和7年度の伸び", "同上"),
    ("要変更", "手順4",
     "施設・居住系サービスの1人1月あたり給付費の実績値",
     "令和7年度／令和8年度から選ぶ", "令和7年度",
     "現況は令和8年度です。令和8年度は4か月であり年度の実績ではありません"
     "（確認事項No.137）"),
    ("要変更", "手順5", "介護サービス利用者数の自然体推計手法",
     "要介護度別／要介護度を包括して の2つから選ぶ",
     "要介護度別に推計する",
     "現況は包括推計です。当方の検証では2か年平均の絶対誤差が"
     "基本0.61％・包括1.23％で要介護度別の方が小さく振れも小さい"
     "（確認事項No.126）"),
    ("変更", "手順5", "在宅サービスの利用率の伸び",
     "4つの選択肢から選ぶ", "令和6年度→令和7年度の伸び", "同上"),
    ("変更", "手順5", "在宅サービスの1人1月あたり利用回（日）数の伸び",
     "4つの選択肢から選ぶ", "令和6年度→令和7年度の伸び", "同上"),
    ("要変更", "手順5", "在宅サービスの1人1月あたり給付費の実績値",
     "令和7年度／令和8年度から選ぶ", "令和7年度",
     "施設・居住系と別に選べます（確認事項No.137）"),
    ("要変更", "手順8", "所得段階の設定",
     "標準13段階／弾力化 から選ぶ", "弾力化",
     "当連合の条例は16段階です。11シートの値を入れます"),
    ("確認", "手順8", "第1号被保険者負担割合", "―", "23％",
     "画面の値を確かめます。政令改正により変わり得ます"),
    ("確認", "手順8", "調整交付金相当額", "―", "5％", "同上"),
    ("確認", "手順8", "予定保険料収納率", "―", "0.99", "同上"),
    ("記入", "手順5", "実績見込み値を編集した理由",
     "自由記入", "下の文例のとおり",
     "出力の「施策反映の内容」に載ります"),
]

ws = sheet("03_設定の選択",
           "推計方法の設定で選ぶ項目と入れる値",
           ZANTEI + "　"
           "画面で選ぶ項目です。"
           "「要変更」は現況から改めるもの、「変更」は選び直すもの、"
           "「確認」は画面の値を確かめるだけのものです。",
           [8, 7, 34, 34, 26, 52], freeze="A5")
r = 4
r = lead(ws, r, "1　設定の選択（%d件）" % len(SETTEI), span=6)
r = header(ws, r, ["区分", "手順", "項目", "画面の選択肢",
                   "選ぶ・入れる値", "理由・備考"])
for ku, te, ko, se, an, ri in SETTEI:
    r = body(ws, r, [ku, te, ko, se, an, ri],
             fills={1: (NG_O if ku == "要変更" else
                        (OK_G if ku == "確認" else MID_B)), 5: IN_Y},
             align={1: "center", 2: "center"}, height=34)

r += 1
r = lead(ws, r, "2　伸びの選択肢（受領した画面の写しによる）", span=6)
r = header(ws, r, ["", "選択肢", "令和8年度（4か月）を含むか",
                   "選ぶ", "", ""])
for no, lab, r8 in MS.GAMEN_NOBI_SENTAKU:
    r = body(ws, r, [no, lab, "含む" if r8 else "含まない",
                     "" if r8 else "これを選ぶ", "", ""],
             fills={3: (NG_O if r8 else OK_G), 4: (None if r8 else IN_Y)},
             height=20)

r += 1
r = lead(ws, r, "3　実績見込み値を編集した理由の記入例", span=6)
r = note(ws, r,
         "令和8年度の値は令和8年4月から7月までの4か月の月報から"
         "計算されるものであり、年度の実績ではない。"
         "在宅サービスは冬季が夏季より低く、4月から7月は夏季に偏る。"
         "このため令和8年度の実績見込み値は"
         "令和7年度（年報・確定値）の月平均に差し替えた。"
         "人口は地方創生総合戦略の総人口と住民基本台帳の年齢階級別の"
         "実績趨勢により推計した値を用いた。",
         span=6, height=56, fill=IN_Y)
r = note(ws, r,
         "注）**「令和8年度の実績見込み値を編集したかどうか」は、"
         "出力の令和8年度の人数の欄が整数かどうかで分かります。**"
         "システムが月報から計算した値は整数、"
         "年報の年間延べを12で除した月平均は小数になります。",
         span=6, height=34)

# ============================================================ 04
ws = sheet("04_令和8年度の実績見込み値",
           "令和8年度の実績見込み値（令和7年度の月平均に差し替える）",
           ZANTEI + "　"
           "令和8年度の実績見込み値は、編集しない限り"
           "令和8年4月から7月までの4か月の月報から計算した値が使われます。"
           "年度の実績ではないため、"
           "**令和7年度（年報・確定値）の月平均に差し替えます。**",
           [4, 30, 8] + [10] * 7 + [11], freeze="D5")
r = 4
_AL4 = {j: "right" for j in range(4, 12)}
r = lead(ws, r, "1　認定者数（第1号被保険者・人）", span=11)
r = header(ws, r, ["", "区分", "単位"] + DO_COLS + ["計"])
_N8 = [_R8["NINTEI"][d] for d in _R8["DO"]]
r = do_row(ws, r, ["1", "認定者数（令和7年度の実績）", "人"], _N8,
           i_fill={j: IN_Y for j in range(4, 11)},
           fmt={j: "#,##0" for j in range(4, 12)})
r = note(ws, r,
         "注）**認定者数の実績見込み値を編集する項目が画面にあるかを"
         "確かめられていません**（確認事項No.138）。"
         "編集する項目がない場合は、認定率の伸びに"
         "「令和6年度→令和7年度の伸び」を選ぶことで"
         "令和8年度の影響を避けます（03シート）。",
         span=11, height=40)

for _no, _ttl, _src, _unit, _dec in (
        ("2", "施設・居住系サービス利用者数（人／月）", "SHISETSU_DO",
         "人／月", False),
        ("3", "在宅サービス利用者数（人／月）", "ZAITAKU_DO", "人／月", False)):
    r += 1
    r = lead(ws, r, "%s　%s" % (_no, _ttl), span=11)
    r = header(ws, r, ["", "サービス", "単位"] + DO_COLS + ["計"])
    _d = _R8[_src]
    for i, lab in enumerate(_d, start=1):
        r = do_row(ws, r, [i, lab, _unit], _R8["to_int"](_d[lab]),
                   i_fill={j: IN_Y for j in range(4, 11)},
                   fmt={j: "#,##0" for j in range(4, 12)})

r += 1
r = lead(ws, r, "4　在宅サービス利用回（日）数（**1月当たりの延べ**・回・日）",
         span=11)
r = header(ws, r, ["", "サービス", "単位"] + DO_COLS + ["計"])
for i, row in enumerate(_R8["KAISU"], start=1):
    _lab, _u, _vals = row[0], row[1], _R8["to_dec1"](row[2])
    r = do_row(ws, r, [i, _lab, _u], _vals,
               i_fill={j: IN_Y for j in range(4, 11)},
               fmt={j: "0.0" for j in range(4, 12)})
r = note(ws, r,
         "注1）" + ZANTEI + "\n"
         "注2）出所は年報 令和7年度（確定値・要介護度別）です。\n"
         "注3）人数の欄は整数で扱われるため整数で掲げています。"
         "要支援の群・要介護の群ごとに合計を保つよう配分しているため、"
         "月1人に満たないサービスも0にはなりません。\n"
         "注4）「―」は制度上その要介護度の区分がないものです。\n"
         "注5）**認知症対応型共同生活介護の要支援2（月0.42人）のように、"
         "月1人に満たない区分は整数の欄では0になります。**"
         "システムのワーニングチェックで挙がりますが、"
         "整数の欄である以上ほかに置きようがないため、理由の記入で足ります。\n"
         "注6）**4の回（日）数は1月当たりの延べです。**"
         "手順5の施策反映の画面（08シート）に入れるのは"
         "1人1月あたりの回（日）数であり、同じ「回（日）数」でも"
         "画面により意味が違います。"
         "本表の値は年報の年間延べを12で除したもので、"
         "3の利用者数で除した1人1月あたりが"
         "総括表詳細（３）と3％以内で一致することを確かめています。",
         span=11, height=96)

# ============================================================ 05
ws = sheet("05_認定者数",
           "要介護度別の認定者数（令和9年度以降）",
           ZANTEI + "　"
           "「施策反映　認定者数」の画面に入れる値です。"
           "年齢階級別の認定者数（案C）に、"
           "年齢階級ごとの要介護度の構成（令和7年度末の実績）を乗じて求めています。",
           [4, 18] + [10] * 7 + [11], freeze="C5")
r = 4
r = lead(ws, r, "第1号被保険者の要介護度別認定者数（人・年度末）", span=10)
r = header(ws, r, ["", "年度"] + DO_COLS + ["計"])
_NROWS = [("令和7年度（実績）", NIN_R7, False)]
for y, yl in zip(YALL, YALLL):
    _NROWS.append((yl, NIN[y], True))
for i, (yl, vals, is_in) in enumerate(_NROWS, start=1):
    r = do_row(ws, r, [i, yl], to_int(list(vals)),
               i_fill=({j: IN_Y for j in range(3, 10)} if is_in
                       else {j: GRAY for j in range(1, 11)}),
               fmt={j: "#,##0" for j in range(3, 11)})
r = note(ws, r,
         "注1）" + ZANTEI + "\n"
         "注2）人口の基礎は案C（総人口＝地方創生総合戦略、"
         "年齢階級別＝住民基本台帳の実績趨勢）です。"
         "02シートの人口の設定と対になっています。\n"
         "注3）整数は要支援の群・要介護の群ごとに合計を保つよう"
         "配分しています。",
         span=10, height=50)


# ============================================================ 06・07・08
# 画面の様式（区分の別・並び順・年度・列）は `data_mieru_yoshiki` による。
# **4つの画面は様式が違う。**
#   手順4 施設・居住系　区分のプルダウン。年度×サービスが縦、合計が左端。11年度。
#   手順5 在宅　サービスごとに1画面。上段が利用率（合計の列なし）、
#         下段が利用者数（在宅サービス対象者数×利用率。自動計算）。11年度。
#   手順6 地域支援事業の量　4区分。R6〜R11の6年度（令和12年度以降は別画面）。
#   手順7 地域支援事業費　5群33項目。R6〜R11の6年度（令和12年度以降は別画面）。
_BY_NAME = {short(l): l for l in SVC}

_SHI_GROUPS = [(no, [_BY_NAME[n] for n in names])
               for no, names in YS.SHISETSU_KUBUN]
_ZAI_GROUPS = [(no, [_BY_NAME[n] for n in names])
               for no, names in YS.ZAITAKU_KUBUN]

_SHISETSU = [l for g in _SHI_GROUPS for l in g[1]]
_ZAITAKU = [l for g in _ZAI_GROUPS for l in g[1]]
_GAMEN_NASHI = [l for l in SVC if short(l) in YS.GAMEN_NASHI]

# 画面の年度（令和7年度は実績として灰色で示す。令和6・8年度は当方が持たない）
_GY = list(YALLL)                      # 令和9年度以降（入力の対象）


def _zai_taisho(vals_nin, riyo):
    """要介護度別の在宅サービス対象者数
    ＝認定者数−施設・居住系サービス利用者数（画面の分母）。"""
    return [vals_nin[i] - sum((riyo[l][i] or 0) for l in _SHISETSU
                              + _GAMEN_NASHI) for i in range(7)]


_ZN_R7 = _zai_taisho(list(NIN_R7), JISSEKI)
_ZN = {y: _zai_taisho(list(NIN[y]), MIKOMI[y]) for y in YALL}


def _ritsu(lab, vals, zn):
    """要介護度別の在宅サービス利用率（％）。画面に列がない欄は None。
    桁は画面に合わせて小数第1位とする。"""
    v = masked(lab, vals)
    return [None if v[i] is None or zn[i] <= 0
            else round(v[i] / zn[i] * 100, YS.RITSU_KETA) for i in range(7)]


def _kei_row(vals):
    """合計（左端）。None を除いて足す。"""
    num = [x for x in vals if x is not None]
    return sum(num) if num else None


def do_row_kei(ws, r, head, vals, i_fill=None, fmt=None, bold=False):
    """合計を左端に置く要介護度別の1行（画面の並び）。"""
    num = [v for v in vals if v is not None]
    cells = list(head) + [(round(sum(num), 1) if num else "―")] \
        + [("―" if v is None else v) for v in vals]
    n = len(head)
    al = {j: "right" for j in range(n + 1, n + 9)}
    return body(ws, r, cells, fills=i_fill, align=al, height=18,
                fmt=fmt or {}, bold=bold)


# ---------------------------------------------------------------- 06
ws = sheet("06_施設居住系の利用者数",
           "施設・居住系サービス利用者数の施策反映（手順4）",
           ZANTEI + "　"
           "**画面は区分のプルダウン（居宅サービス・地域密着型サービス・"
           "施設サービス）で1つずつ開きます。**"
           "表は行が年度×サービス、列は**合計が左端**で、"
           "その右に要支援1・要支援2・要介護1から要介護5が並びます。"
           "合計は自動計算です。利用者数の欄は整数で扱われます。"
           "「―」は画面にその要介護度の列がないものです。",
           [4, 26, 13, 10] + [9] * 7, freeze="D5")
r = 4
_i = 0
for no, labs in _SHI_GROUPS:
    r = lead(ws, r, "プルダウンで「%s」を選ぶ　　利用者数（人）" % no, span=11)
    r = header(ws, r, ["", "サービス", "年度", "合計"] + DO_COLS)
    for lab in labs:
        _i += 1
        rows = [("令和7年度\n（実績）", riyosha_r7(lab), False)]
        for y, yl in zip(YALL, YALLL):
            rows.append((yl, riyosha(lab, y), True))
        for j, (yl, vals, is_in) in enumerate(rows):
            r = do_row_kei(ws, r, [_i if j == 0 else "",
                                   short(lab) if j == 0 else "", yl], vals,
                           i_fill=({k: IN_Y for k in range(5, 12)}
                                   if is_in else {3: GRAY, 4: GRAY}))
N_SHISETSU = _i
r = note(ws, r,
         "注1）" + ZANTEI + "\n"
         "注2）区分の別・並び順・列は画面の写しによります。"
         "総括表の「施設サービス／居住系サービス」とは束ね方が違います。\n"
         "注3）**画面は令和6年度から令和32年度までの11年度を1つの表に"
         "縦に並べます。**令和6年度から令和8年度は実績として灰色で表示され、"
         "令和9年度以降が入力の対象です。本表は令和7年度（当広域連合の実績）と"
         "令和9年度以降を掲げています。\n"
         "注4）**介護療養型医療施設は画面に行がありません**"
         "（令和6年3月末で廃止された施設サービスであり、"
         "当広域連合の実績もありません）。\n"
         "注5）**認知症対応型共同生活介護の要支援2は月0.42人であり、"
         "整数の欄では0になります。**"
         "ワーニングチェックで「下限を下回る」として挙がりますが、"
         "整数の欄である以上ほかに置きようがないため理由の記入で足ります。",
         span=11, height=100)

# ---------------------------------------------------------------- 07
ws = sheet("07_在宅の利用率と利用者数",
           "在宅サービスの利用者数等の施策反映（手順5）",
           ZANTEI + "　"
           "**画面はサービスのプルダウンで1サービスにつき1つ開きます。**"
           "上段が【入力】在宅サービス利用率（％）、"
           "下段が【自動計算】在宅サービス利用者数"
           "（%s×利用率）です。"
           "**利用率の表には合計の列がありません。**"
           "1に利用率、2に対応する利用者数を掲げています。"
           % YS.ZAITAKU_BUNBO,
           [4, 26, 13] + [9] * 8, freeze="D5")
r = 4
r = note(ws, r,
         "【利用率の分母】\n"
         "画面の下段にあるとおり、利用者数は**%s×利用率**で自動計算されます。"
         "%sは**認定者数−施設・居住系サービス利用者数**（要介護度別）であり、"
         "第1号被保険者数ではありません。"
         "令和7年度は %s人（要介護度別 %s）。\n"
         "画面が表示している利用率と当広域連合の実績から逆算した率が"
         "全21区分で近い値になることを確かめています（13シート 点検18）。"
         % (YS.ZAITAKU_BUNBO, YS.ZAITAKU_BUNBO,
            "{:,.1f}".format(sum(_ZN_R7)),
            "／".join("%.1f" % x for x in _ZN_R7)),
         span=11, height=62, fill=MID_B)
r += 1
_i = 0
for no, labs in _ZAI_GROUPS:
    r = lead(ws, r, "1　" + no + "　【入力】在宅サービス利用率（％・小数第1位）",
             span=11)
    r = header(ws, r, ["", "サービス（プルダウンで選ぶ）", "年度"] + DO_COLS)
    for lab in labs:
        _i += 1
        rows = [("令和7年度\n（実績）", JISSEKI[lab], _ZN_R7, False)]
        for y, yl in zip(YALL, YALLL):
            rows.append((yl, MIKOMI[y][lab], _ZN[y], True))
        for j, (yl, vals, zn, is_in) in enumerate(rows):
            rr = _ritsu(lab, vals, zn)
            r = body(ws, r,
                     [_i if j == 0 else "", short(lab) if j == 0 else "", yl]
                     + ["―" if x is None else x for x in rr],
                     fills=({k: IN_Y for k in range(4, 11)} if is_in
                            else {3: GRAY}),
                     align={k: "right" for k in range(4, 11)},
                     fmt={k: "0.0" for k in range(4, 11)}, height=18)
N_ZAITAKU = _i
r = note(ws, r,
         "注1）" + ZANTEI + "\n"
         "注2）率は利用者数÷%s（要介護度別）です。"
         "**画面と同じく合計の列は置いていません**"
         "（率は各列の和になりません）。\n"
         "注3）「―」は画面にその要介護度の列がないものです。"
         "**訪問介護・通所介護・地域密着型通所介護・定期巡回・夜間対応型・"
         "看護小規模多機能型居宅介護は要支援の列がありません**"
         "（介護予防訪問介護・介護予防通所介護は総合事業へ移行済みです）。\n"
         "注4）画面には「利用者数で施策反映する」のボタンもあり、"
         "利用者数で入れることもできます。その場合は2の値を用います。"
         % YS.ZAITAKU_BUNBO,
         span=11, height=72)

r += 1
_i = 0
for no, labs in _ZAI_GROUPS:
    r = lead(ws, r, "2　" + no
             + "　【自動計算】在宅サービス利用者数（人／月・整数）", span=11)
    r = header(ws, r, ["", "サービス", "年度", "合計"] + DO_COLS)
    for lab in labs:
        _i += 1
        rows = [("令和7年度\n（実績）", riyosha_r7(lab), False)]
        for y, yl in zip(YALL, YALLL):
            rows.append((yl, riyosha(lab, y), True))
        for j, (yl, vals, is_in) in enumerate(rows):
            r = do_row_kei(ws, r, [_i if j == 0 else "",
                                   short(lab) if j == 0 else "", yl], vals,
                           i_fill=({k: IN_Y for k in range(5, 12)}
                                   if is_in else {3: GRAY, 4: GRAY}))
r = note(ws, r,
         "注）1の利用率に分母を乗じた値です。"
         "画面では自動計算されるため、通常は入力しません。"
         "「利用者数で施策反映する」を選んだとき、及び手順2の"
         "令和8年度の実績見込み値を編集するときにこの値を用います。"
         "整数は要支援の群・要介護の群ごとに合計を保つよう配分しています。",
         span=11, height=44)

# ============================================================ 08
ws = sheet("08_在宅の利用回日数",
           "在宅サービスの利用回（日）数の施策反映（手順5の別画面）",
           ZANTEI + "　"
           "**画面はサービスのプルダウンで1サービスにつき1つ開きます。**"
           "入れるのは**%s**であり、"
           "1月当たりの延べ回（日）数ではありません。"
           "**画面に合計の列はありません**（列は要支援1・要支援2・"
           "要介護1〜5の7つだけです）。欄は小数第1位まで受け付けます。"
           % YS.KAISU_SEISHITSU,
           [4, 26, 7, 13] + [9] * 7, freeze="E5")
r = 4
r = note(ws, r,
         "【画面の説明文（逐語）】\n" + YS.KAISU_SETSUMEI + "\n"
         "プルダウンに並ぶのは回（日）数の計上がある%d区分だけです"
         "（居宅療養管理指導・福祉用具貸与・月包括報酬によるものは並びません）。"
         % sum(len(x[1]) for x in YS.KAISU_KUBUN),
         span=11, height=52, fill=MID_B)
r += 1
N_KAISU = 0
_i = 0
for no, names in YS.KAISU_KUBUN:
    r = lead(ws, r, no + "　【入力】1人1月あたりの利用回（日）数"
             "（回・日／月・小数第1位）", span=11)
    r = header(ws, r,
               ["", "サービス（プルダウンで選ぶ）", "単位", "年度"] + DO_COLS)
    for nm in names:
        lab = _BY_NAME[nm]
        _i += 1
        N_KAISU += 1
        u = KAISU_MAP[lab][2]
        vals = [None if v is None else round(v, 1)
                for v in kaisu_tanka_do(lab)]
        rows = [("令和7年度\n（実績）", False)]
        rows += [(yl, True) for yl in YALLL]
        for j, (yl, is_in) in enumerate(rows):
            r = body(ws, r,
                     [_i if j == 0 else "", nm if j == 0 else "",
                      u if j == 0 else "", yl]
                     + ["―" if x is None else x for x in vals],
                     fills=({k: IN_Y for k in range(5, 12)} if is_in
                            else {4: GRAY}),
                     align={k: "right" for k in range(5, 12)},
                     fmt={k: "0.0" for k in range(5, 12)}, height=18)
r = note(ws, r,
         "注1）" + ZANTEI + "\n"
         "注2）**1人1月あたりを令和7年度で固定しているため、"
         "令和9年度以降は令和7年度と同じ値になります。**"
         "利用者数の側が動くことにより延べの回（日）数は年度ごとに変わります。\n"
         "注3）年報（令和7年度）の要介護度別の利用回（日）数を"
         "同じ区分の受給者数で除して求めた値です。"
         "**画面が令和7年度に表示している値と一致します**"
         "（14シート 点検23）。\n"
         "注4）**通所介護・地域密着型通所介護・通所リハビリテーション・"
         "認知症対応型通所介護の単位は「回」です**（年報・総括表による）。"
         "短期入所2種別は「日」です。\n"
         "注5）「―」は画面にその要介護度の列がないものです。"
         "画面と同じく合計の列は置いていません"
         "（1人1月あたりは各列の和になりません）。",
         span=11, height=90)

# ============================================================ 09
# 地域支援事業は手順6（量）と手順7（事業費）の2つの画面に分かれ、
# それぞれ「実績値及び計画値の入力」（R6〜R11）と
# 「自然体推計値の確認と施策反映値の入力」（R12・R17・R22・R27・R32）がある。
_CY6 = list(YS.NENDO_CHIIKI)           # 令和6〜11年度（第10期の画面）
_CY_L = list(YS.NENDO_CHOKI)           # 令和12年度以降の画面
_CY_IN = set(YS.NENDO_10KI)

_RYO_SRC = {nm: (key, u, nb) for nm, key, u, nb in SOGO_RYO}
_RYO_IN = {}
for _nm, _src in YS.CHIIKI_RYO_MAP.items():
    _key, _u, _nb = _RYO_SRC[_src]
    _b = sg3(_key)
    _f = NIN_NOBI if _nb == "認定者数" else HIHO_NOBI
    _RYO_IN[_nm] = (_b, {yl: (None if _b is None else _b * _f[y])
                         for y, yl in zip(Y3, Y3L)})

ws = sheet("09_地域支援事業の量",
           "訪問型・通所型サービス利用者数の入力（手順6）",
           ZANTEI + "　"
           "「地域支援事業の見込み量推計」の画面に入れる値です。"
           "**区分は4つだけ**で、単位は人／月です。"
           "画面は「年齢階級別に登録する」と"
           "「サービスごとの総数で登録する」を選べます。"
           "**当広域連合は年齢階級別の実績を持たないため"
           "「サービスごとの総数で登録する」を選びます。**"
           "**年度は令和6年度から令和11年度の6年度**で、"
           "令和12年度以降は別画面です。",
           [4, 40, 12] + [13] * (len(_CY6) + len(_CY_L)), freeze="D5")
r = 4
r = note(ws, r,
         "【この画面の値は特に暫定の度合いが高いものです】\n"
         "**総合事業の令和7年度実績が未受領である**ため、"
         "令和6年度の利用者実人数（3町計）を基準として据え置き、"
         "認定者数の伸びで延ばしています（確認事項No.112）。\n"
         "**訪問型サービスAは総合事業の実施状況に関する調査に区分がなく、"
         "値を置けません。**",
         span=3 + len(_CY6) + len(_CY_L), height=48, fill=NG_O)
r += 1
r = lead(ws, r, "1　サービスごとの総数で登録する（推奨）",
         span=3 + len(_CY6) + len(_CY_L))
r = header(ws, r, ["", "サービス種別・項目", "単位"] + _CY6 + _CY_L)
N_SOGO = 0
for _i2, _nm in enumerate(YS.CHIIKI_RYO_KUBUN, start=1):
    _in = _RYO_IN.get(_nm)
    if _in is None:
        _vals = ["" for _ in _CY6 + _CY_L]
        _fl = {j: NG_O for j, y in enumerate(_CY6 + _CY_L, start=4)
               if y in _CY_IN or y in _CY_L}
    else:
        _b, _fy = _in
        N_SOGO += 1
        _vals = ([_b if y == "令和6年度" else
                  (round(_fy[y], 1) if y in _CY_IN else "") for y in _CY6]
                 + ["" for _ in _CY_L])
        _fl = {j: (IN_Y if y in _CY_IN else GRAY)
               for j, y in enumerate(_CY6, start=4)}
        _fl.update({j: NG_O for j, y in
                    enumerate(_CY_L, start=4 + len(_CY6))})
    r = body(ws, r, [_i2, _nm, "人／月"] + _vals, fills=_fl,
             align={j: "right" for j in range(4, 4 + len(_CY6) + len(_CY_L))},
             fmt={j: "0.0" for j in range(4, 4 + len(_CY6) + len(_CY_L))},
             height=20)
r = note(ws, r,
         "注1）" + ZANTEI + "\n"
         "注2）黄色は当広域連合で値を用意できる欄、"
         "赤は資料の受領又はご判断を要する欄、灰色は画面の実績値です。\n"
         "注3）**令和12年度以降は「訪問型・通所型サービス利用者数／事業費の"
         "自然体推計値の確認と施策反映値の入力」という別の画面**です。"
         "当広域連合は令和6年度の水準を据え置いており、"
         "中長期の量は自然体推計のままとします。\n"
         "注4）「年齢階級別に登録する」を選ぶと、"
         "各区分の下に%sの%d階級の欄が開きます。"
         "総合事業の実施状況に関する調査は年齢階級別の実人数を持たないため、"
         "**「サービスごとの総数で登録する」を選びます。**"
         % ("・".join(YS.CHIIKI_NENREI), len(YS.CHIIKI_NENREI)),
         span=3 + len(_CY6) + len(_CY_L), height=76)

r += 1
r = lead(ws, r, "2　画面に量の欄がない事業（入力しません）",
         span=3 + len(_CY6) + len(_CY_L))
r = header(ws, r, ["", "項目", "単位", "令和6年度実績\n（3町計）",
                   "令和9年度", "令和11年度", "扱い"]
           + [""] * (len(_CY6) + len(_CY_L) - 4))
_NO_IN = [(nm, key, u, nb) for nm, key, u, nb in SOGO_RYO
          if nm not in YS.CHIIKI_RYO_MAP.values()]
N_KAYOI = N_NOIN = 0
for nm, key, u, nb in _NO_IN:
    b = sg3(key)
    if b is None:
        continue
    N_NOIN += 1
    f = NIN_NOBI if nb == "認定者数" else HIHO_NOBI
    r = body(ws, r, [N_NOIN, nm, u, b, round(b * f[Y3[0]], 1),
                     round(b * f[Y3[2]], 1), YS.CHIIKI_RYO_NASHI_RIYU]
             + [""] * (len(_CY6) + len(_CY_L) - 4),
             fmt={4: "#,##0", 5: "0.0", 6: "0.0"},
             align={j: "right" for j in range(4, 7)}, height=28)
for nm, key, u, nb in KAYOI_RYO:
    b = sum(x for x in SG.KAYOI[key][0] if x is not None)
    N_KAYOI += 1
    N_NOIN += 1
    if nb == "据え置き":
        v9 = v11 = b
    else:
        f = NIN_NOBI if nb == "認定者数" else HIHO_NOBI
        v9, v11 = b * f[Y3[0]], b * f[Y3[2]]
    r = body(ws, r, [N_NOIN, nm, u, b, round(v9, 1), round(v11, 1),
                     YS.CHIIKI_RYO_NASHI_RIYU]
             + [""] * (len(_CY6) + len(_CY_L) - 4),
             fmt={4: "#,##0", 5: "0.0", 6: "0.0"},
             align={j: "right" for j in range(4, 7)}, height=28)
r = note(ws, r,
         "注1）**画面には量（利用者数）の欄が4区分しかありません。**"
         "これらは入力せず、計画本文（第6章第3節2）に掲げます。\n"
         "注2）3町の地域包括支援センター事業実施報告書（令和7年度）により"
         "事業ごとの実施状況は確認できましたが、"
         "**量の単位が延べ人数・食数・回数・団体数で混在**しており、"
         "国の実施状況調査の利用者実人数とは数えているものが違います。"
         "足さない・比べないこととし、"
         "見込みの基礎は令和6年度の利用者実人数のまま据え置いています。",
         span=3 + len(_CY6) + len(_CY_L), height=56)

# ============================================================ 10
ws = sheet("10_地域支援事業費",
           "地域支援事業費の入力（手順7）",
           ZANTEI + "　"
           "「地域支援事業費」の画面に入れる値です。単位は円（年間累計）。"
           "**5つの群に分かれ、計はシステムが計算します。**"
           "**訪問介護相当サービス・訪問型サービスA・通所介護相当サービス・"
           "通所型サービスAの4項目は手順6の量から計算されるため入力しません。**"
           "**年度は令和6年度から令和11年度の6年度**で、"
           "令和12年度以降は別画面です。",
           [4, 48, 14] + [14] * len(_CY6), freeze="D5")
r = 4
r = note(ws, r,
         "【この画面の値は特に暫定の度合いが高いものです】\n"
         "**事業別の内訳（総合事業23項目・包括的支援事業8項目）は未受領です**"
         "（確認事項No.5・No.112・No.150）。"
         "当広域連合が用意できるのは総合事業費の計と"
         "包括的支援事業・任意事業費の計の2つだけであり、"
         "事業別に割り付けられません。\n"
         "**未受領の欄は空欄のままとし、0を入れません**"
         "（0を入れると「計上がない」ことを表してしまいます）。",
         span=3 + len(_CY6), height=56, fill=NG_O)
r += 1
_CHI_ROWS = 0
for _gno, (_gname, _items) in enumerate(YS.CHIIKI_HI_KUBUN, start=1):
    r = lead(ws, r, _gname, span=3 + len(_CY6))
    r = header(ws, r, ["", "サービス種別・項目", "区分"] + _CY6)
    for _nm, _auto in _items:
        _CHI_ROWS += 1
        _shin = _nm in YS.CHIIKI_HI_SHINSETSU
        _vals = [("―" if (_shin and y in YS.NENDO_JISSEKI) else "")
                 for y in _CY6]
        if _auto:
            _bk = "手順6から自動計算"
            _fl = {j: GRAY for j in range(1, 4 + len(_CY6))}
        else:
            _bk = "事業費（円）"
            _fl = {j: NG_O for j, y in enumerate(_CY6, start=4)
                   if y in _CY_IN}
        r = body(ws, r, ["", _nm, _bk] + _vals, fills=_fl,
                 align={j: "right" for j in range(4, 4 + len(_CY6))},
                 fmt={j: "#,##0" for j in range(4, 4 + len(_CY6))},
                 height=18)

r = lead(ws, r, "地域支援事業費計（システムが計算します）", span=3 + len(_CY6))
r = header(ws, r, ["", "サービス種別・項目", "区分"] + _CY6)
for _nm in YS.CHIIKI_HI_KEI:
    _shin = _nm in YS.CHIIKI_HI_SHINSETSU
    _vals = [("―" if (_shin and y in YS.NENDO_JISSEKI) else "")
             for y in _CY6]
    r = body(ws, r, ["", _nm, "自動計算"] + _vals,
             fills={j: GRAY for j in range(1, 4 + len(_CY6))},
             align={j: "right" for j in range(4, 4 + len(_CY6))},
             height=18, bold=(_nm == "地域支援事業費"))

r += 1
_B_Y = CHIIKI_R6
r = lead(ws, r, "参考　当広域連合が用意できる計", span=3 + len(_CY6))
r = header(ws, r, ["", "事業区分", "区分", "令和9年度", "令和10年度",
                   "令和11年度"] + [""] * (len(_CY6) - 3))
_CHI = [
    ("介護予防・日常生活支援総合事業費", SOGO_R6,
     "令和6年度決算（3町の計画作成支援ツールによる）"),
    ("包括的支援事業（センター運営）及び任意事業費"
     "／包括的支援事業（社会保障充実分）", HOKATSU_R6,
     "6事業別・任意事業別の内訳が未受領のため区分できません（確認事項No.5）"),
    ("地域支援事業費（計。保険料算定上のB）", _B_Y, ""),
]
for _i2, (_nm2, _v2, _bk2) in enumerate(_CHI, start=1):
    r = body(ws, r, [_i2, _nm2, "事業費（円）", _v2, _v2, _v2]
             + [""] * (len(_CY6) - 3),
             fills={4: IN_Y, 5: IN_Y, 6: IN_Y},
             fmt={j: "#,##0" for j in range(4, 7)},
             align={j: "right" for j in range(4, 7)},
             height=26, bold=(_i2 == 3))
r = note(ws, r,
         "注1）" + ZANTEI + "\n"
         "注2）令和6年度決算の水準を3年据え置いたものです。"
         "この額を入れると算定上の月額が約＋332円上がります。\n"
         "注3）**保険者機能強化推進事業費及び保険者努力支援事業費は"
         "この額から除いています**（令和6年度決算の款4の全体183,384,259円から"
         "2事業を除いた額が%s円）。"
         "このため、手順8の「保険者機能強化推進交付金等の交付見込額」の欄は"
         "**0のままとします。**両方を行うと二重控除になります（確認事項No.86）。\n"
         "注4）**特定地域居宅サービス等事業・介護情報利活用事業は"
         "令和9年4月の新設**であり、令和6年度から令和8年度は「―」です。"
         "額が定まらないため空欄とします。\n"
         "注5）**令和12年度以降は「地域支援事業費の自然体推計値の確認と"
         "施策反映値の入力」という別の画面**です。"
         % "{:,}".format(CHIIKI_R6),
         span=3 + len(_CY6), height=92)

# ============================================================ 11
ws = sheet("11_保険料額の算定",
           "保険料額の算定（手順8）",
           ZANTEI + "　"
           "「所得段階別第1号被保険者数・基準額に対する割合の登録」の画面に"
           "入れる値です。"
           "**所得段階の設定は「弾力化」を選びます**"
           "（当連合の条例は16段階で、標準13段階とは9段階で割合が違い、"
           "第14段階から第16段階は標準にありません）。",
           [4, 40, 9, 14] + [10] * len(YALL), freeze="E5")
r = 4
_AL11 = {j: "right" for j in range(3, 5 + len(YALL))}
r = lead(ws, r, "1　所得段階別第1号被保険者数・割合・基準所得金額",
         span=4 + len(YALL))
r = header(ws, r, ["段階", "対象者要件", "割合", "基準所得金額（円）"] + YALLL)
for _i in range(NDAN):
    _kin = KIJUN_SHOTOKU[_i]
    r = body(ws, r, [_i + 1, DANKAI_NAME[_i], JORITSU[_i],
                     _kin if _kin else "―"]
             + [DANKAI[y][_i] for y in YALL],
             fills={3: IN_Y, 4: (IN_Y if _kin else GRAY),
                    **{j: IN_Y for j in range(5, 5 + len(YALL))}},
             fmt={3: "0.000", 4: "#,##0",
                  **{j: "#,##0" for j in range(5, 5 + len(YALL))}},
             align=_AL11, height=26)
r = body(ws, r, ["", "計", "―", "―"]
         + [sum(DANKAI[y]) for y in YALL],
         fills={j: MID_B for j in range(1, 5 + len(YALL))}, bold=True,
         fmt={j: "#,##0" for j in range(5, 5 + len(YALL))},
         align=_AL11, height=22)
r = body(ws, r, ["", "補正後被保険者数（自動計算）", "―", "―"]
         + [round(sum(a * b for a, b in zip(DANKAI[y], JORITSU)), 1)
            for y in YALL],
         fills={j: GRAY for j in range(1, 5 + len(YALL))},
         fmt={j: "#,##0.0" for j in range(5, 5 + len(YALL))},
         align=_AL11, height=22)
r = note(ws, r,
         "注1）" + ZANTEI + "\n"
         "注2）**割合は公費軽減前の値です。**"
         "第1段階から第3段階は公費軽減の対象ですが、"
         "補正後被保険者数はこの割合で算定します。"
         "公費軽減後の値（第1段階0.285など）は入れません。\n"
         "注3）**基準所得金額は第7段階から第16段階の10件です。**"
         "第1段階から第6段階は本人の合計所得金額の下限による区分ではないため、"
         "画面も当該欄を受け付けません。\n"
         "注4）**計は「実績及び推計方法の設定」の第1号被保険者数と"
         "一致させます**（02シートの値）。一致していないと登録できません。\n"
         "注5）第17段階から第30段階は空欄のままとします。\n"
         "注6）第10期の段階数・割合・基準所得金額は政令改正により"
         "変わり得ます（確認事項No.33）。"
         "保険料段階の基準額を80.9万円から82.65万円に改める政令改正が"
         "済んでおり、条例改正の手続を要します。",
         span=4 + len(YALL), height=112)

r += 1
r = lead(ws, r, "2　そのほかの欄", span=4 + len(YALL))
r = header(ws, r, ["#", "項目", "入れる値", "備考"]
           + [""] * len(YALL))
_HOKA = [
    ("準備基金取崩額", "0（当面）",
     "令和7年度末の基金残高の提供を受けてから入れます。"
     "1億円の取崩しで算定上の月額が約310円下がります"),
    ("保険者機能強化推進交付金等の交付見込額", "0",
     "地域支援事業費から2事業を既に除いているため、"
     "ここで控除すると二重控除になります（確認事項No.86）"),
    ("財政安定化基金拠出金見込額・償還金", "0", "いずれも0です"),
    ("市町村特別給付費等", "0", "当連合では行っていません"),
]
for i, (a, b, c) in enumerate(_HOKA, start=1):
    r = body(ws, r, [i, a, b, c] + [""] * len(YALL),
             fills={3: IN_Y}, height=30)

# ============================================================ 12
_HOSEI3 = sum(sum(a * b for a, b in zip(DANKAI[y], JORITSU)) for y in Y3)
ws = sheet("12_入力後の確かめ方",
           "入力が正しく効いたかどうかを確かめる値",
           ZANTEI + "　"
           "入力を終えたあと、出力（総括表）の値が次のようになっているかを"
           "確かめてください。"
           "**一致するのは第1号被保険者数と認定者数までです。**"
           "給付費から先は上乗せ率と調整交付金の作り方が当方と違うため、"
           "差が残ります。",
           [4, 34, 24, 24, 44], freeze="A5")
r = 4
# 利用回（日）数の延べ（令和11年度）。画面は1人1月あたりを要介護度別に
# 入れるため、システムが求める延べは要介護度の構成の動きを取り込む。
# 計画本文は要支援・要介護の群でまとめた1人1月あたりによっている。
_kaisu_mikomi = _S["kaisu_mikomi"]
_KAI_G = _KAI_D = 0.0
_KAI_MAX, _KAI_MAXN = 0.0, "―"
for _l in [l for l in SVC if l in KAISU_MAP]:
    _a = _kaisu_mikomi(_l, Y3[2])
    _td = kaisu_tanka_do(_l)
    _b = sum((MIKOMI[Y3[2]][_l][i] or 0.0) * (_td[i] or 0.0)
             for i in range(7))
    _KAI_G += _a
    _KAI_D += _b
    if _a and abs(_b / _a - 1) > abs(_KAI_MAX):
        _KAI_MAX, _KAI_MAXN = _b / _a - 1, short(_l)
_KAI_MAXS = "{:+.2f}％".format(_KAI_MAX * 100)

r = lead(ws, r, "1　一致するはずの値", span=5)
r = header(ws, r, ["#", "確かめる値", "当方の算定", "どこで見るか", "備考"])
_KAKU = [
    ("第1号被保険者数（令和9年度）", "{:,}人".format(ichigo("R9")),
     "総括表 1．被保険者数", "02シートの入力が効いたかどうか"),
    ("第1号被保険者数（令和11年度）", "{:,}人".format(ichigo("R11")),
     "総括表 1．被保険者数", "同上"),
    ("第1号被保険者数（第10期3か年）",
     "{:,}人".format(sum(ichigo(y) for y in ("R9", "R10", "R11"))),
     "総括表 5．保険料推計 6．第1号被保険者数関係", "11シートの計と一致します"),
    ("要介護（支援）認定者数（令和11年度・第1号）",
     "{:,}人".format(sum(to_int(list(NIN[Y3[2]])))),
     "総括表 2．要介護（支援）認定者数",
     "**現況は2,031人です。この値が変わらなければ、"
     "人口の設定が認定者数へ及んでいません**"),
    ("補正後被保険者数（第10期3か年）", "{:,.1f}人".format(_HOSEI3),
     "総括表 5．保険料推計", "11シートの弾力化が効いたかどうか"),
    ("所得段階別加入割合の合計", "全年度 1.0000",
     "総括表 5．保険料推計", "1.0を超えていたら計がそろっていません"),
    ("地域支援事業費（第10期3か年）", "{:,}円".format(G_DO["B"]),
     "総括表 3．地域支援事業費・5．保険料推計",
     "**現況は0円です**"),
]
for i, (a, b, c, d) in enumerate(_KAKU, start=1):
    r = body(ws, r, [i, a, b, c, d], fills={3: OK_G},
             align={3: "right"}, height=30)

r += 1
r = lead(ws, r, "2　差が残る値（一致しません）", span=5)
r = header(ws, r, ["#", "項目", "当方の算定", "見える化（現況）", "差の理由"])
_SA = [
    ("総給付費（第10期3か年）", "{:,}円".format(sum(_S["KYUFU3"])),
     "{:,}円".format(int(sum(K2.KYUFU[y] for y in ("R9", "R10", "R11"))
                         * 1000)),
     "基準年度・推計手法・人口が違います。"
     "03シートの設定を改めると近づきます"),
    ("標準給付費見込額（3か年）", "{:,}円".format(G_DO["A"]),
     "{:,}円".format(int(K2.SHUNO["標準給付費見込額"]["計"])),
     "上乗せ率が違います（当方は令和6年度決算の割増率1.05909で一括、"
     "見える化は特定入所者・高額・高額医療合算・審査支払手数料を個別に入力）"),
    ("算定上の月額", "{:,.2f}円".format(G_DO["月額"]),
     "{:,.2f}円".format(K2.HOKENRYO["第10期"]),
     "上記の積み重ねと調整交付金の作り方の違いによります"),
    ("保険料基準額（月額）", "{:,}円".format(KIJUN_GAKU), "―",
     "算定上の月額を百円未満四捨五入した値です"),
    ("利用回（日）数（延べ・令和11年度）",
     "計画本文は{:,.1f}回・日／月".format(_KAI_G),
     "本表の入力から {:,.1f}回・日／月".format(_KAI_D),
     "画面は1人1月あたりを要介護度別に入れるため、"
     "システムが求める延べは要介護度の構成の動きを取り込みます。"
     "計画本文は要支援・要介護の群でまとめた1人1月あたりによっており、"
     "差は最大 %s（%s）です。保険料には入りません" % (_KAI_MAXS, _KAI_MAXN)),
]
for i, (a, b, c, d) in enumerate(_SA, start=1):
    r = body(ws, r, [i, a, b, c, d], fills={3: MID_B},
             align={3: "right", 4: "right"}, height=34)
r = note(ws, r,
         "注1）" + ZANTEI + "\n"
         "注2）**算定式そのものは当方と見える化で同じです。**"
         "第9期の保険料基準額は見える化%.2f円・当方の再現6,427.92円で"
         "一致しており、差は入力と設定によるものです。\n"
         "注3）当方の算定は、単価を令和7年度で固定し"
         "介護報酬改定率・単価の趨勢を織り込んでいません。"
         "これらは保険料を上げる向きに働くため、"
         "**当方の値は下限に近いもの**です。"
         % K2.HOKENRYO_K9,
         span=5, height=62)

# ============================================================ 13 自己点検
ws = sheet("13_自己点検", "自己点検",
           "本表の内的整合を機械で確かめた記録です。"
           "1件でも不適合があると出力そのものを止めます。",
           [6, 44, 34, 34, 10])

_h_ng = [y for y in YALL if sum(DANKAI[y]) != ichigo(
    {"2027": "R9", "2028": "R10", "2029": "R11", "2030": "R12",
     "2035": "R17", "2040": "R22", "2045": "R27", "2050": "R32"}[y])]
chk(1, "所得段階別の計が人口の設定の第1号被保険者数と一致すること",
    "11シートの計 ＝ 02シートの第1号被保険者数（自動計算）",
    "全%d年度／合わない %s" % (len(YALL), _h_ng or "なし"), not _h_ng)

_PER = 1 + 6 * 2 + 2          # 総人口1＋男女6階級12＋第2号2
_inp = sum(1 for y in JY if y not in AN.JITSU) * _PER   # 改める欄
_all = len(JY) * _PER                                   # 画面の入力欄の総数
chk(2, "人口の設定の欄の数",
    "（総人口1＋男女6階級12＋第2号2）×年度数",
    "画面の入力欄 %d欄／改める欄 %d欄（令和6年度・令和7年度は画面の値のまま）"
    % (_all, _inp), _all == 165 and _inp == 135)

_jis_ok = all(ANC[y]["出所"] == "画面（実績）" for y in AN.JITSU)
chk(3, "令和6年度・令和7年度が画面の値のままであること",
    "ANC の出所が「画面（実績）」であること",
    "令和6年度・令和7年度とも画面の値", _jis_ok)

def _grp_keep(vals):
    """整数化が要支援の群・要介護の群ごとに合計を保っているか。"""
    iv = to_int(list(vals))
    for g in ((0, 1), (2, 3, 4, 5, 6)):
        idx = [i for i in g if vals[i] is not None]
        if not idx:
            continue
        if sum(iv[i] for i in idx) != AN.r0(sum(vals[i] for i in idx)):
            return False
    return True


_n_ng = [yl for y, yl in zip(YALL, YALLL) if not _grp_keep(list(NIN[y]))]
chk(4, "認定者数の整数化が要支援の群・要介護の群ごとに合計を保つこと",
    "群ごとに Σ整数 ＝ Σ実数の四捨五入",
    "全%d年度／合わない %s" % (len(YALL), _n_ng or "なし"), not _n_ng)

_shien_ng = []
for lab in SVC:
    a1, a2 = SHIEN_ARI[lab]
    for y in YALL:
        v = riyosha(lab, y)
        if (v[0] is None) != (not a1) or (v[1] is None) != (not a2):
            _shien_ng.append(short(lab))
            break
chk(5, "要支援の列の有無が画面の写しと一致すること",
    "MS.GAMEN_DO・MS.ZAITAKU_DO の「―」と突き合わせる",
    "全%d区分／合わない %s" % (len(SVC), _shien_ng or "なし"), not _shien_ng)

_ks_none = [i + 1 for i in range(NDAN) if KIJUN_SHOTOKU[i] is None]
_ks_val = [x for x in KIJUN_SHOTOKU if x is not None]
chk(6, "基準所得金額が第7段階から第16段階の10件で昇順であること",
    "第1〜6段階は欄なし／第7段階以降は金額かつ昇順",
    "欄なし 第%s段階／金額%d件"
    % ("・".join(str(x) for x in _ks_none), len(_ks_val)),
    _ks_none == [1, 2, 3, 4, 5, 6] and len(_ks_val) == 10
    and all(a < b for a, b in zip(_ks_val, _ks_val[1:])))

chk(7, "割合が公費軽減前の値であること",
    "第1段階＝0.455（軽減後は0.285）",
    "第1段階 %.3f／第2段階 %.3f／第3段階 %.3f"
    % (JORITSU[0], JORITSU[1], JORITSU[2]),
    abs(JORITSU[0] - 0.455) < 1e-9)

chk(8, "地域支援事業費の年度別の3倍が保険料算定上のBと一致すること",
    "CHIIKI_R6×3 ＝ G_DO['B']",
    "%s円×3＝%s円／B %s円"
    % ("{:,}".format(CHIIKI_R6), "{:,}".format(CHIIKI_R6 * 3),
       "{:,}".format(G_DO["B"])), CHIIKI_R6 * 3 == G_DO["B"])

chk(9, "地域支援事業費の内訳の和が計と一致すること",
    "総合事業費＋包括的支援事業・任意事業費 ＝ B（年度別）",
    "%s＋%s＝%s" % ("{:,}".format(SOGO_R6), "{:,}".format(HOKATSU_R6),
                    "{:,}".format(SOGO_R6 + HOKATSU_R6)),
    SOGO_R6 + HOKATSU_R6 == CHIIKI_R6)

_ich_ng = [sh for _, _, _, _, sh, _, _ in ICHIRAN
           if sh not in _HYO and sh != "―"]
chk(10, "一覧の「本表のシート」が実在するシートを指すこと",
    "_HYO の値が wb.sheetnames に含まれること",
    "参照先の不明 %s" % (_ich_ng or "なし"),
    not _ich_ng and all(v in wb.sheetnames for v in _HYO.values()))

_so_ng = [x[3] for x in ICHIRAN if not x[3]]
chk(11, "一覧の全ての行に操作が書かれていること",
    "操作の欄が空でないこと",
    "%d件中 空 %d件" % (len(ICHIRAN), len(_so_ng)), not _so_ng)

# 08_在宅の利用回日数 は画面に合計の列がないため対象外
# （1人1月あたりは各列の和にならない。07シートの利用率と同じ）。
_KEI_SHEETS = ("04_令和8年度の実績見込み値", "05_認定者数",
               "06_施設居住系の利用者数", "07_在宅の利用率と利用者数")
_kei_ng, _kei_n = [], 0
_kei_by = {s: 0 for s in _KEI_SHEETS}
for _w in wb.worksheets:
    if _w.title not in _KEI_SHEETS:
        continue
    _hd = _kei = None
    for _row in _w.iter_rows():
        _v = [c.value for c in _row]
        if "要支援1" in _v:
            _hd = _v.index("要支援1")
            # 「計」は右端（従前の並び）、「合計」は左端（画面の並び）
            _kei = (_hd - 1 if (_hd > 0 and _v[_hd - 1] == "合計")
                    else _hd + 7)
            continue
        if _hd is None or len(_v) <= max(_hd + 6, _kei):
            continue
        _seg = _v[_hd:_hd + 7]
        _tot = _v[_kei]
        _num = [x for x in _seg if isinstance(x, (int, float))]
        if not _num or not isinstance(_tot, (int, float)):
            continue
        _kei_n += 1
        _kei_by[_w.title] += 1
        if abs(sum(_num) - _tot) > 0.051:
            _kei_ng.append("%s!%d" % (_w.title, _row[0].row))
# 件数が0のときは、数えられているかを先に疑う（CLAUDE.md §4）。
_kei_zero = [s for s, n in _kei_by.items() if n == 0]
chk(12, "書き出した表の内訳の和と「合計」の欄が一致すること",
    "04・05・06・07シートの全行を読み直して検算"
    "（合計の欄は画面と同じく左端。従前の右端の「計」も拾う）",
    "%d行を検算（%s）／合わない %s／数えられていないシート %s"
    % (_kei_n, "・".join("%s %d" % (s[:2], _kei_by[s]) for s in _KEI_SHEETS),
       _kei_ng or "なし", _kei_zero or "なし"),
    not _kei_ng and not _kei_zero)

chk(12.5, "地域支援事業の量の件数が0でないこと",
    "画面に欄がある量・欄がない量の行数",
    "画面に欄がある量%d件／欄がない量%d件（うち通いの場%d件）"
    % (N_SOGO, N_NOIN, N_KAYOI),
    N_SOGO > 0 and N_NOIN > 0 and N_KAYOI > 0)

_ast = []
for _w in wb.worksheets:
    for _row in _w.iter_rows():
        for _c in _row:
            if isinstance(_c.value, str) and "**" in _c.value:
                _ast.append("%s!%s" % (_w.title, _c.coordinate))
chk(13, "強調の指定（**）がセルに残っていないこと",
    "全シートの全セルを走査（xlsx は Markdown を解釈しない）",
    "%d件" % len(_ast), not _ast)

import re as _re
_PI = [_re.compile(r"\d{2,4}-\d{2,4}-\d{3,4}"),
       _re.compile(r"[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}")]
_pi = []
for _w in wb.worksheets:
    for _row in _w.iter_rows():
        for _c in _row:
            if isinstance(_c.value, str):
                for _p in _PI:
                    if _p.search(_c.value):
                        _pi.append("%s!%s" % (_w.title, _c.coordinate))
chk(14, "個人を特定する値を収めていないこと",
    "電話番号・メールアドレスの形を全セルで走査",
    "%d件" % len(_pi), not _pi)

_zan = [_w.title for _w in wb.worksheets
        if _w.title != "13_自己点検"
        and "暫定値" not in str(_w["A2"].value or "")]
chk(15, "自己点検を除く全シートの冒頭に暫定値である旨があること",
    "各シートのA2に「暫定値」の語があること",
    "無いシート %s" % (_zan or "なし"), not _zan)

# ---------------------------------------------- 様式（画面ごとに違うこと）
_sk = [no for no, _ns in YS.SHISETSU_KUBUN]
_sk_gm = set(MS.GAMEN_KUBUN.keys())
chk(16, "施設・居住系の区分と並び順が画面の写しと一致すること",
    "画面の区分のプルダウン（居宅・地域密着型・施設）とサービスの並び",
    "%d区分・%dサービス（%s）／画面にない区分 %s"
    % (len(_SHI_GROUPS), N_SHISETSU, "・".join(_sk),
       "・".join(short(l) for l in _GAMEN_NASHI) or "なし"),
    len(_SHI_GROUPS) == 3 and N_SHISETSU == 7
    and set(_sk) == _sk_gm
    and [short(l) for l in _GAMEN_NASHI] == YS.GAMEN_NASHI)

_zset = {short(l) for l in SVC if kubun(l) == "在宅サービス"}
_yset = {n for _no, ns in YS.ZAITAKU_KUBUN for n in ns}
chk(17, "在宅の様式の区分が当方の在宅サービスと過不足なく一致すること",
    "YS.ZAITAKU_KUBUN の名称の集合 ＝ 総括表の在宅サービスの集合",
    "様式%d区分／当方%d区分／差 %s"
    % (len(_yset), len(_zset), (_yset ^ _zset) or "なし"),
    _yset == _zset and N_ZAITAKU == len(_zset))

# 在宅サービス利用率の分母（認定者数−施設・居住系利用者数）の裏づけ。
# 画面が表示している令和7年度の利用率と、当方の実績から逆算した率を比べる。
_MN = {"特定福祉用具販売": "特定福祉用具購入費", "住宅改修": "住宅改修費"}
_rd, _rn = 0.0, 0
for _lab in _ZAITAKU:
    _nm = _MN.get(short(_lab), short(_lab))
    if _nm not in MS.ZAITAKU_DO:
        continue
    _g = MS.ZAITAKU_DO[_nm]["R7"][0] * 100
    _v = masked(_lab, JISSEKI[_lab])
    _num = [x for x in _v if x is not None]
    _m = (sum(_num) / sum(_ZN_R7) * 100) if _num else 0.0
    _rn += 1
    _rd = max(_rd, abs(_g - _m))
chk(18, "在宅サービス利用率の分母が画面の率から裏づけられること",
    "分母＝認定者数−施設・居住系利用者数として逆算した率と画面の率を比べる",
    "%d区分／最大の差 %.2fポイント（時点の違いによるもの）" % (_rn, _rd),
    _rn == len(_ZAITAKU) and _rd < 1.0)

_ryo_n = len(YS.CHIIKI_RYO_KUBUN)
_auto_n = sum(1 for _g, _it in YS.CHIIKI_HI_KUBUN for _nm, _a in _it if _a)
chk(19, "地域支援事業の量の欄が4区分であること",
    "画面の量の区分の数／当方が値を置ける数／事業費の自動計算の欄の数",
    "量の区分 %d／当方が置ける %d（訪問型サービスAは調査に区分なし）"
    "／事業費の自動計算 %d欄"
    % (_ryo_n, N_SOGO, _auto_n),
    _ryo_n == 4 and N_SOGO == 3 and _auto_n == 4)

_ci = sum(len(_it) for _g, _it in YS.CHIIKI_HI_KUBUN)
chk(20, "地域支援事業費の画面の行を残らず掲げていること",
    "書き出した項目の数 ＝ 画面の5群の項目の数",
    "画面%d項目（%d群）／書き出し%d項目／計の行%d"
    % (_ci, len(YS.CHIIKI_HI_KUBUN), _CHI_ROWS, len(YS.CHIIKI_HI_KEI)),
    _ci == _CHI_ROWS and len(YS.CHIIKI_HI_KUBUN) == 5
    and len(YS.CHIIKI_HI_KEI) == 6)

_ny = [(w.title, y) for w in wb.worksheets
       for y in ("令和12年度",)
       if w.title in ("09_地域支援事業の量", "10_地域支援事業費")]
chk(21, "地域支援事業の画面の年度が令和6年度から令和11年度であること",
    "手順6・手順7の第10期の画面は6年度（令和12年度以降は別画面）",
    "第10期の画面 %d年度（%s）／令和12年度以降の画面 %d年次"
    % (len(YS.NENDO_CHIIKI), "・".join(YS.NENDO_CHIIKI),
       len(YS.NENDO_CHOKI)),
    len(YS.NENDO_CHIIKI) == 6 and len(YS.NENDO_CHOKI) == 5)

chk(22, "施設・居住系と在宅の画面が11年度を1つの表に並べること",
    "実績3年度＋第10期3年度＋中長期5年次",
    "%d年度" % len(YS.NENDO_SHISAKU), len(YS.NENDO_SHISAKU) == 11)

# 利用回（日）数の画面（手順5の別画面。令和8年9月30日受領の写しによる）
_kk = [n for _, ns in YS.KAISU_KUBUN for n in ns]
_km = [short(l) for l in SVC if l in KAISU_MAP]
_kk_ng = sorted(set(_kk) ^ set(_km))
chk(22.5, "利用回（日）数の画面の区分が当方の回（日）数の区分と一致すること",
    "画面のプルダウン ＝ 回（日）数の計上がある区分",
    "画面%d区分・当方%d区分／差 %s"
    % (len(_kk), len(_km), _kk_ng or "なし"), not _kk_ng and len(_kk) == 12)

# 見出し行は「要支援1」を含む行として探す（行番号を書かない）。
# 見出しだけを見ると、行を書く側が計を足していても気づかないため、
# 値の入っている行が見出しより右に及んでいないことも見る。
_w08 = wb["08_在宅の利用回日数"]
_h08 = [[c.value for c in _row] for _row in _w08.iter_rows()
        if "要支援1" in [c.value for c in _row]]
# iter_rows は max_column まで空セルを付けて返すため、
# 見出しの列数は「中身のある最も右の列」で数える
# （len() で数えると、右に値を足した行があっても見出しも同じ長さになり気づけない）。
_n08 = (1 + max(k for k, x in enumerate(_h08[0]) if x not in (None, ""))
        if _h08 else 0)
_x08 = []
for _row in _w08.iter_rows():
    _v = [c.value for c in _row]
    if len(_v) < 4 or not str(_v[3] or "").startswith("令和"):
        continue
    _last = max((k for k, x in enumerate(_v) if x not in (None, "")),
                default=-1)
    if _last + 1 > _n08:
        _x08.append("%d行目" % _row[0].row)
chk(22.7, "利用回（日）数の表に合計の列を置いていないこと",
    "画面に合計の列がない（見出しと値の行の両方を見る）",
    "見出し%d列（%s）／見出しより右に値のある行 %s"
    % (_n08, "・".join(str(x) for x in (_h08[0] if _h08 else []) if x),
       _x08[:3] or "なし"),
    bool(_h08) and not _x08
    and all("合計" not in h and "計" not in h for h in _h08))

_kg, _kgn = [], 0
for _nm, _v in YS.KAISU_GAMEN_R7.items():
    _t = kaisu_tanka_do(_BY_NAME[_nm])
    for _i, _d in enumerate(DO_COLS):
        if _d not in _v:                # 画面の写しで読み取れていない欄
            continue
        _kgn += 1
        _a = None if _t[_i] is None else round(_t[_i], 1)
        if _a != _v[_d]:
            _kg.append("%s %s（当方%s／画面%s）" % (_nm, _d, _a, _v[_d]))
chk(23, "令和7年度の1人1月あたりが画面の表示値と一致すること",
    "画面の写し（訪問介護・訪問看護）と当方の年報の要介護度別明細による算定",
    "%d区分・%d欄／合わない %s"
    % (len(YS.KAISU_GAMEN_R7), _kgn, _kg or "なし"), not _kg and _kgn >= 10)

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
     "本表の値が暫定値であることを変えるものではありません。",
     span=5, height=26)

# ============================================================ 印刷・保存
for _w in wb.worksheets:
    _w.page_setup.orientation = "landscape"
    _w.page_setup.fitToWidth = 1
    _w.page_setup.fitToHeight = 0
    _w.sheet_properties.pageSetUpPr.fitToPage = True
    _w.print_title_rows = "4:4"
    _w.oddFooter.left.text = "見える化システムに入力する箇所と入力する値（%s）" \
        % KIJUNBI
    _w.oddFooter.center.text = "【暫定値】必要資料未受領による暫定値"
    _w.oddFooter.right.text = "&P / &N"

wb.save(OUT)
print("書き出しました:", OUT)
for _w in wb.worksheets:
    print("  -", _w.title, _w.max_row, "rows")
print("変更する箇所 %d件／設定の選択 %d件" % (len(ICHIRAN), len(SETTEI)))
print("人口の入力欄 %d／サービス 施設居住系%d・在宅%d／回（日）数%d区分"
      % (_inp, N_SHISETSU, N_ZAITAKU, N_KAISU))
print("地域支援事業 サービス事業%d・通いの場%d／地域支援事業費 %s円／年"
      % (N_SOGO, N_KAYOI, "{:,}".format(CHIIKI_R6)))
print("所得段階 %d段階（弾力化）／補正後被保険者数（3か年）%.1f人" % (NDAN, _HOSEI3))
print("算定上の月額 %d円（保険料基準額 %s円）"
      % (GETSU, "{:,}".format(KIJUN_GAKU)))
print("**本表の値は必要資料が未受領であることによる暫定値です。**")
print("自己点検 %d件：適合%d件・不適合%d件"
      % (len(CHECKS), len(CHECKS) - _NG, _NG))
if _NG:
    print("不適合があります。")
    sys.exit(1)
print("すべての点検に適合しました。")
