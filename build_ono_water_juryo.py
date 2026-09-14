"""小野町水道事業 料金改定支援業務　受領資料の確認と整理。

令和8年9月11日に町（地域整備課）から「080911送付」として44ファイルを受領した。
水道事業会計の予算・決算一式であり、令和2年度から令和8年度までを覆う。

本スクリプトは原本を読み、料金改定の算定に必要な系列を1冊に整理する。
あわせて、系列どうしの突合（事業報告書・決算報告書・損益計算書・貸借対照表・
決算統計の相互チェック）を行い、料金改定に向けて足りない資料を洗い出す。

原本の主な構成
  R6決算／R7決算        決算報告書（R1〜R7の7年度分を収録）、損益計算書、
                        貸借対照表、剰余金計算書、剰余金処分計算書・事業報告書・
                        収益費用明細書・固定資産明細・企業債明細、CF計算書、
                        補填財源明細書
  R7当初予算／R8当初予算  議案、予算説明書、CF計算書、損益計算書、貸借対照表、
                        債務償還年次表
  R7補正（9月・12月・3月）／R8補正（9月）
  決算統計              令和6年度決算（総務省最終・原価区分変更後）、令和7年度決算
  消費税                消費税申告資料

出力: 小野町水道料金改定/01_確認・整理/小野町水道_受領資料の確認と整理_YYYYMMDD.xlsx
"""

import pathlib
import re
import warnings

warnings.filterwarnings("ignore")

import openpyxl
import xlrd
from openpyxl.styles import Alignment, Border, Font, PatternFill
from openpyxl.styles.borders import Side
from openpyxl.utils import get_column_letter

ROOT = pathlib.Path(__file__).parent / "小野町水道料金改定"
SRC = ROOT / "00_受領データ" / "原本_080911送付"
OUT = ROOT / "01_確認・整理"
ASOF = "20260914"
ASOF_JP = "令和8年9月14日"
JURYO_JP = "令和8年9月11日"

HEAD = PatternFill("solid", fgColor="1F3864")
SUB = PatternFill("solid", fgColor="DDEBF7")
KEY = PatternFill("solid", fgColor="FCE4E4")
CALC = PatternFill("solid", fgColor="EAF1FB")
WARN = PatternFill("solid", fgColor="FFE0B2")
OK = PatternFill("solid", fgColor="C6EFCE")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

YEARS = ["R1", "R2", "R3", "R4", "R5", "R6", "R7"]

# 決算報告書ブック（R1〜R7を1冊に収録）
F_KESSAN = "R7決算/080817修正【済】3.R7決算報告書.xlsx"
F_JIGYO = ("R7決算/080818修正【済】R7剰余金処分計算書・事業報告書・収益費用明細書・"
           "固定資産明細・企業債明細★.xlsx")
F_PL7 = "R7決算/【済】4_R7損益計算書（3月補正→決算）.xls"
F_PL6 = "R6決算/【済】4_R6損益計算書（2月補正→決算）.xls"
F_PL8 = "R8当初予算/4_R8損益計算書【済】.xls"
F_BS7 = "R7決算/080812修正【済】5_R7.3月貸借対照表（3月補正→決算）.xls"
F_BS89 = "R8補正/【済】5_R8.9月貸借対照表（R8当初ﾍﾞｰｽ・R7決算反映済）.xls"
F_TK7 = "決算統計/R7075221010.xlsx"
F_TK6 = "決算統計/【総務省最終】070715原価区分変更後.xlsx"


def nz(s):
    return re.sub(r"[\s　]", "", str(s))


# ================================================================ 抽出

def read_gyomu():
    """事業報告書（R1〜R7）から業務量と給水収益を取り出す。"""
    wb = openpyxl.load_workbook(SRC / F_JIGYO, data_only=True)
    keys = ["年度末給水人口", "計画給水人口", "給水普及率", "１日最大配水量", "給水収益"]
    out = {}
    for sn in wb.sheetnames:
        m = re.match(r"^(R\d)決算書", sn.strip())
        if not m:
            continue
        ws = wb[sn]
        d = {}
        for r in range(1, ws.max_row + 1):
            lab = "".join(nz(ws.cell(r, c).value) for c in range(1, 10)
                          if isinstance(ws.cell(r, c).value, str))
            if not lab:
                continue
            val = next((ws.cell(r, c).value for c in range(10, 20)
                        if isinstance(ws.cell(r, c).value, (int, float))), None)
            if val is None:
                continue
            for k in keys:
                if nz(k) in lab and k not in d:
                    d[k] = val
            for k in ("配水量", "有収水量", "有収率"):
                if lab.startswith(k) and k not in d:
                    d[k] = val
        if d:
            out[m.group(1)] = d
    return out


def _find_col(ws, r0, key):
    for r in range(max(1, r0 - 12), r0):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and key in nz(v):
                return c
    return None


KESSAN_LABELS = ["水道事業収益", "営業収益", "営業外収益",
                 "水道事業費用", "営業費用", "営業外費用",
                 "資本的収入", "工事負担金", "企業債", "他会計長期借入金",
                 "他会計補助金", "資本的支出", "建設改良費", "企業債償還金"]


def read_kessan():
    """決算報告書（R1〜R7を1冊に収録）から収益的・資本的収支の決算額を取り出す。"""
    wb = openpyxl.load_workbook(SRC / F_KESSAN, data_only=True)
    out = {}
    for sn in wb.sheetnames:
        m = re.match(r"^(R\d)決算書", sn.strip())
        if not m:
            continue
        ws = wb[sn]
        d = {}
        for r in range(1, ws.max_row + 1):
            lab = None
            for c in range(1, 8):
                v = ws.cell(r, c).value
                if not isinstance(v, str):
                    continue
                s = nz(v)
                for L in KESSAN_LABELS:
                    if s == L or re.match(rf"^第[１-９1-9]+[款項]{L}$", s):
                        lab = L
                        break
                if lab:
                    break
            if not lab or lab in d:
                continue
            col = _find_col(ws, r, "決算額")
            if col is None:
                continue
            for c in range(col, min(col + 4, ws.max_column + 1)):
                v = ws.cell(r, c).value
                if isinstance(v, (int, float)):
                    d[lab] = v
                    break
        if d:
            out[m.group(1)] = d
    return out


PL_KEYS = ["給水収益", "その他営業収益", "原水及び浄水費", "配水及び給水費",
           "総係費", "減価償却費", "資産減耗費", "受取利息及び配当金",
           "他会計補助金", "雑収益", "長期前受金戻入",
           "支払利息及び企業債取扱諸費", "雑支出"]
PL_TOTALS = ["営業損失", "営業利益", "経常利益", "経常損失",
             "当年度純利益", "当年度純損失",
             "前年度繰越利益剰余金", "当年度未処分利益剰余金"]


def read_pl(path):
    """損益計算書（.xls）を読む。明細は左の列、集計は右の列にある。"""
    sh = xlrd.open_workbook(SRC / path).sheet_by_name("損益計算書")
    d = {}
    for r in range(sh.nrows):
        lab = nz("".join(str(sh.cell_value(r, c)) for c in range(0, 2)
                         if isinstance(sh.cell_value(r, c), str)))
        # 明細行は「(１)給水収益」のように番号が前に付く
        lab = re.sub(r"^[(（][０-９0-9一二三四五六七八九十１-９]+[)）]", "", lab)
        if not lab:
            continue
        vals = [(c, sh.cell_value(r, c)) for c in range(1, sh.ncols)
                if isinstance(sh.cell_value(r, c), float)]
        if not vals:
            continue
        for k in PL_KEYS:
            if lab.startswith(nz(k)) and k not in d:
                d[k] = vals[0][1]
        for k in PL_TOTALS:
            if lab.startswith(nz(k)) and k not in d:
                d[k] = vals[-1][1]
    return d


def read_bs(path):
    """貸借対照表（.xls）を読む。左が資産・負債、右が資本。"""
    sh = xlrd.open_workbook(SRC / path).sheet_by_name("貸借対象表（２P）")
    rows = []
    for r in range(sh.nrows):
        left = nz("".join(str(sh.cell_value(r, c)) for c in range(0, 2)
                          if isinstance(sh.cell_value(r, c), str)))
        lv = next((sh.cell_value(r, c) for c in range(1, 5)
                   if isinstance(sh.cell_value(r, c), float)
                   and sh.cell_value(r, c) != 0), None)
        right = nz("".join(str(sh.cell_value(r, c)) for c in range(5, 7)
                           if isinstance(sh.cell_value(r, c), str)))
        rv = next((sh.cell_value(r, c) for c in range(6, sh.ncols)
                   if isinstance(sh.cell_value(r, c), float)
                   and sh.cell_value(r, c) != 0), None)
        rows.append((left, lv, right, rv))
    # 「流動資産合計」と「資産合計」、「固定負債合計」と「負債合計」のように
    # 一方が他方の末尾に含まれる項目があるため、番号を外したうえで完全一致でとる。
    def key(s):
        return re.sub(r"^[(（][０-９0-9一二三四五六七八九十１-９]+[)）]", "", s)

    LEFT = {"固定資産合計": "固定資産合計", "有形固定資産合計": "有形固定資産合計",
            "企業債": "企業債（固定負債）", "他会計借入金": "他会計借入金",
            "固定負債合計": "固定負債合計", "流動負債合計": "流動負債合計",
            "長期前受金": "長期前受金", "繰延収益合計": "繰延収益合計",
            "負債合計": "負債合計"}
    RIGHT = {"現金預金": "現金預金", "未収金合計": "未収金合計",
             "流動資産合計": "流動資産合計", "資産合計": "資産合計",
             "資本金合計": "資本金合計", "資本剰余金合計": "資本剰余金合計",
             "利益剰余金合計": "利益剰余金合計", "剰余金合計": "剰余金合計",
             "資本合計": "資本合計", "負債・資本合計": "負債資本合計",
             "当年度未処分利益剰余金": "当年度未処分利益剰余金"}
    d = {}
    for left, lv, right, rv in rows:
        k = LEFT.get(key(left))
        if k and k not in d and lv:
            d[k] = lv
        k = RIGHT.get(key(right))
        if k and k not in d and rv:
            d[k] = rv
    return d


def read_tokei(path):
    """決算統計の01表（施設及び業務概況）から料金体系と業務指標を取り出す。"""
    ws = openpyxl.load_workbook(SRC / path, data_only=True)["01(001)"]
    g = lambda r: ws.cell(r, 10).value
    return {
        "行政区域内現在人口": g(30), "計画給水人口": g(31), "現在給水人口": g(32),
        "水利権": g(41), "配水能力": g(50), "一日最大配水量": g(51),
        "年間総配水量": g(52), "年間総有収水量": g(53),
        "基本水量": g(58), "基本料金": g(59), "超過料金": g(60),
        "10㎥13mm": g(61), "10㎥20mm": g(62), "20㎥13mm": g(63), "20㎥20mm": g(64),
    }


def read_shokan():
    """決算統計45表から企業債の年度別償還予定を取り出す。"""
    ws = openpyxl.load_workbook(SRC / F_TK7, data_only=True)["45(001)"]
    out = []
    seen = set()
    for r in range(15, ws.max_row):
        y = ws.cell(r, 2).value
        if not (isinstance(y, str) and "年度" in y):
            continue
        y = y.strip()
        if y in seen:
            continue
        seen.add(y)
        gan = sum(ws.cell(r, c).value or 0 for c in range(6, 14)
                  if isinstance(ws.cell(r, c).value, (int, float)))
        ri = sum(ws.cell(r + 1, c).value or 0 for c in range(6, 14)
                 if isinstance(ws.cell(r + 1, c).value, (int, float)))
        if gan or ri:
            out.append((y, gan, ri))
    return out


# ================================================================ 体裁

def style_header(ws, row=1):
    for c in ws[row]:
        if c.value is not None:
            c.font = Font(bold=True, size=9, color="FFFFFF")
            c.fill = HEAD
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = BORDER


def body(ws, first=2, wrap=()):
    for row in ws.iter_rows(min_row=first):
        for c in row:
            if c.value is not None:
                c.border = BORDER
            c.font = Font(size=c.font.size or 9, bold=bool(c.font.bold))
            c.alignment = Alignment(vertical="top", wrap_text=(c.column in wrap))


def widths(ws, ws_w):
    for i, w in enumerate(ws_w, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def notes(ws, lines):
    ws.append([])
    for t in lines:
        ws.append([t])
        ws.cell(ws.max_row, 1).font = Font(size=9, italic=True)


def num(ws, r0, cols, fmt="#,##0", fill=None):
    for r in range(r0, ws.max_row + 1):
        for c in cols:
            cell = ws.cell(r, c)
            if isinstance(cell.value, (int, float)):
                cell.number_format = fmt
                if fill:
                    cell.fill = fill


# ================================================================ シート

def sheet_intro(wb, g, chk):
    ws = wb.create_sheet("00_この資料について")
    r7 = g["R7"]
    tanka = r7["給水収益"] / r7["有収水量"]
    rows = [
        ["小野町水道事業　料金改定支援業務　受領資料の確認と整理"],
        [f"作成 {ASOF_JP}／受領 {JURYO_JP}（町 地域整備課「080911送付」44ファイル）"],
        [],
        ["■ 受領したもの"],
        ["", "水道事業会計の予算・決算一式。令和2年度から令和8年度までを覆う。"],
        ["", "決算報告書と事業報告書は1冊に令和1年度から令和7年度までの7年度分を収録しており、"],
        ["", "**料金改定の算定に必要な時系列は、このセットだけでそろっている。**"],
        [],
        ["■ 令和7年度決算の姿（本資料の中心）"],
        ["", "給水収益", r7["給水収益"], "円", "料金収入"],
        ["", "有収水量", r7["有収水量"], "㎥", ""],
        ["", "供給単価", tanka, "円/㎥", "給水収益÷有収水量"],
        ["", "給水原価", chk["給水原価"], "円/㎥", "（営業費用＋営業外費用）÷有収水量"],
        ["", "料金回収率", chk["料金回収率"], "％",
         "**100円の費用に対し料金で回収できているのは76円。残りは一般会計繰入と長期前受金戻入で埋めている**"],
        ["", "営業損失", chk["営業損失"], "円", "給水収益では営業費用を賄えていない"],
        ["", "経常利益", chk["経常利益"], "円",
         "他会計補助金19.5百万円と長期前受金戻入27.1百万円により黒字"],
        [],
        ["■ 確認して分かったこと"],
        ["", "1", "**給水人口は7年で11.6％減っている**（令和1年度4,881人→令和7年度4,313人）。"
         "年あたり▲1.8％で、給水収益も▲6.7％。料金は据え置かれており供給単価は245〜251円/㎥で横ばい"],
        ["", "2", "**有収率が低い。**令和7年度69.98％で、令和5年度は62.5％まで落ちていた。"
         "全国平均（約90％）と比べて20ポイント低い。配水量の3割が収益にならずに失われている"],
        ["", "3", "**施設が過大である。**配水能力4,870㎥/日に対し一日最大配水量1,908㎥/日で"
         "施設利用率39.2％。計画給水人口8,759人に対し現在給水人口4,313人（49.2％）"],
        ["", "4", "**令和8年度当初予算では営業損失が35.3百万円に拡大し、経常利益は7.4百万円に縮む。**"
         "現金預金も令和7年度末135.6百万円から令和8年9月補正時点109.7百万円へ25.9百万円減る見込み"],
        ["", "5", "基本料金は10㎥まで2,200円（税込）、超過料金220円/㎥。口径別料金制。"
         "**基本水量10㎥・基本料金2,200円は全国平均（10㎥で1,300円前後）を大きく上回る**"],
        [],
        ["■ 本資料の構成"],
        ["", "01_受領資料一覧", "44ファイルの分類と用途"],
        ["", "02_業務量の推移", "給水人口・配水量・有収水量・有収率・給水収益（令和1〜7年度）"],
        ["", "03_収益的収支の推移", "決算報告書による（令和1〜7年度）"],
        ["", "04_損益の推移", "損益計算書による（令和6・7年度決算、令和8年度予算）"],
        ["", "05_貸借対照表", "令和7年度決算と令和8年9月補正"],
        ["", "06_資本的収支・企業債", "建設改良費・企業債・償還予定"],
        ["", "07_料金体系と経営指標", "決算統計による"],
        ["", "08_突合結果", "系列どうしの整合の点検"],
        ["", "09_不足資料", "料金改定の算定に向けて足りないもの"],
    ]
    for r in rows:
        ws.append(r)
    ws["A1"].font = Font(bold=True, size=13)
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.startswith("■"):
            ws.cell(r, 1).font = Font(bold=True, size=10)
    for r in range(1, ws.max_row + 1):
        c = ws.cell(r, 3)
        if isinstance(c.value, (int, float)):
            c.number_format = "#,##0.0" if ws.cell(r, 4).value in ("円/㎥", "％") else "#,##0"
            c.fill = KEY
    widths(ws, [3, 18, 18, 8, 76])
    for row in ws.iter_rows():
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    return ws


FILES = [
    ("R6決算", 7, "令和6年度の決算関係。決算報告書は令和1〜6年度を収録",
     "決算報告書／損益計算書／貸借対照表／剰余金計算書／剰余金処分計算書・事業報告書・"
     "収益費用明細書・固定資産明細・企業債明細／CF計算書／補填財源明細書"),
    ("R7当初予算", 5, "令和7年度の当初予算",
     "議案／予算説明書／CF計算書／貸借対照表／債務償還年次表"),
    ("R7補正", 12, "令和7年度の9月・12月・3月補正",
     "各回の議案／補正予算説明書／CF計算書／貸借対照表"),
    ("R7決算", 6, "**令和7年度の決算。本業務の基準年度**",
     "決算報告書（令和1〜7年度を収録）／損益計算書／貸借対照表／剰余金計算書／"
     "剰余金処分計算書・事業報告書ほか／CF計算書"),
    ("R8当初予算", 7, "令和8年度の当初予算",
     "議案／予算説明書／CF計算書／損益計算書／貸借対照表／債務償還年次表"),
    ("R8補正", 4, "令和8年度の9月補正（令和7年度決算を反映）",
     "議案／補正予算説明書／CF計算書／貸借対照表"),
    ("決算統計", 2, "地方公営企業決算状況調査（総務省）",
     "令和7年度決算／令和6年度決算（総務省最終・**原価区分変更後**）"),
    ("消費税", 1, "消費税申告資料（令和3〜7年度）", "申告資料"),
    ("（直下）", 1, "補填財源明細書", "令和6年度決算反映・令和7年度決算見込・令和8年度予算"),
]


def sheet_files(wb):
    ws = wb.create_sheet("01_受領資料一覧")
    ws.append(["フォルダ", "件数", "内容", "収録されているもの", "料金改定での用途"])
    USE = {
        "R6決算": "前年度との対比。決算統計の原価区分変更前後の確認",
        "R7当初予算": "予算と決算の乖離の把握",
        "R7補正": "年度中の補正の経緯。9月→12月→3月の3回",
        "R7決算": "**基準年度。総括原価の算定基礎**",
        "R8当初予算": "**料金改定前の見通し。営業損失35.3百万円に拡大**",
        "R8補正": "令和7年度決算を反映した最新の見通し",
        "決算統計": "**料金体系・業務指標・企業債償還予定。経営指標の全国比較の基礎**",
        "消費税": "税抜・税込の整理。料金は税込表示",
        "（直下）": "補填財源（資本的収支の不足額の補填）の推移",
    }
    for f, n, memo, naka in FILES:
        ws.append([f, n, memo, naka, USE[f]])
    ws.append(["合計", sum(n for _f, n, _m, _k in FILES), "", "", ""])
    style_header(ws)
    for c in range(1, 6):
        ws.cell(ws.max_row, c).font = Font(bold=True, size=9)
        ws.cell(ws.max_row, c).fill = SUB
    body(ws, wrap=(3, 4, 5))
    widths(ws, [14, 7, 44, 52, 46])
    notes(ws, [
        f"※ {JURYO_JP}に町（地域整備課）から「080911送付」として受領。"
        "ファイル名の文字コードがcp932のため、展開時に文字化けする場合があります。",
        "※ 拡張子は .xlsx（新形式）と .xls（旧形式）と .doc（議案）が混在しています。",
        "**※ 決算報告書と事業報告書は令和1年度から令和7年度までの7年度分を1冊に収録しており、"
        "時系列の系列はこの2冊から取れます。**",
    ])
    return ws


def sheet_gyomu(wb, g):
    ws = wb.create_sheet("02_業務量の推移")
    ws.append(["項目", "単位"] + [f"令和{y[1]}年度" for y in YEARS] + ["R1→R7", "備考"])
    ITEMS = [
        ("年度末給水人口", "人", "#,##0", "**7年で568人（11.6％）減**"),
        ("計画給水人口", "人", "#,##0", "**現在給水人口の約2倍。施設は計画人口に合わせて整備されている**"),
        ("給水普及率", "％", "0.0", "77〜78％で横ばい。簡易水道区域等が残る"),
        ("配水量", "㎥", "#,##0", ""),
        ("有収水量", "㎥", "#,##0", "**料金の対象となる水量。7年で7.3％減**"),
        ("有収率", "％", "0.0", "**令和5年度62.5％まで低下。全国平均は約90％**"),
        ("１日最大配水量", "㎥", "#,##0", "配水能力4,870㎥に対し令和7年度1,908㎥（施設利用率39.2％）"),
        ("給水収益", "円", "#,##0", "**7年で7.6百万円（6.7％）減**"),
    ]
    for k, u, fmt, memo in ITEMS:
        vals = [g.get(y, {}).get(k) for y in YEARS]
        chg = ""
        if vals[0] and vals[-1]:
            chg = f"{(vals[-1] / vals[0] - 1) * 100:+.1f}％"
        ws.append([k, u] + vals + [chg, memo])
    # 供給単価
    ws.append(["供給単価", "円/㎥", *[(g[y]["給水収益"] / g[y]["有収水量"]) if y in g else None
                                 for y in YEARS],
               f"{(g['R7']['給水収益'] / g['R7']['有収水量']) / (g['R1']['給水収益'] / g['R1']['有収水量']) * 100 - 100:+.1f}％",
               "**料金改定がないため横ばい。245〜251円/㎥**"])
    ws.append(["1人あたり有収水量", "㎥/年", *[(g[y]["有収水量"] / g[y]["年度末給水人口"])
                                        if y in g else None for y in YEARS], "",
               "給水人口の減少より有収水量の減少がゆるやか＝1人あたりは増えている"])
    style_header(ws)
    for r in range(2, ws.max_row + 1):
        u = ws.cell(r, 2).value
        fmt = "0.0" if u in ("％", "円/㎥", "㎥/年") else "#,##0"
        for c in range(3, 3 + len(YEARS)):
            ws.cell(r, c).number_format = fmt
            ws.cell(r, c).fill = CALC
    body(ws, wrap=(len(YEARS) + 4,))
    widths(ws, [20, 8] + [13] * len(YEARS) + [10, 56])
    ws.freeze_panes = "C2"
    notes(ws, [
        "※ 出典は事業報告書（剰余金処分計算書・事業報告書・収益費用明細書・固定資産明細・"
        "企業債明細★／令和1〜7年度を1冊に収録）。",
        "**※ 料金改定の需要見通しは、この有収水量の趨勢に給水人口の将来推計を重ねて作ります。**"
        "給水人口の将来推計は未受領（09_不足資料 参照）。",
    ])
    return ws


def sheet_shushi(wb, k):
    ws = wb.create_sheet("03_収益的収支の推移")
    ws.append(["区分", "項目"] + [f"令和{y[1]}年度" for y in YEARS] + ["R1→R7", "備考"])
    ITEMS = [
        ("収益的収入", "水道事業収益", ""),
        ("", "　営業収益", "給水収益が大半"),
        ("", "　営業外収益", "他会計補助金・長期前受金戻入"),
        ("収益的支出", "水道事業費用", ""),
        ("", "　営業費用", "**減価償却費が約78百万円で過半を占める**"),
        ("", "　営業外費用", "支払利息ほか"),
        ("差引", "収支差引", "収益－費用"),
        ("資本的収入", "資本的収入", ""),
        ("", "　工事負担金", ""),
        ("", "　企業債", ""),
        ("", "　他会計補助金・長期借入金", ""),
        ("資本的支出", "資本的支出", ""),
        ("", "　建設改良費", ""),
        ("", "　企業債償還金", ""),
        ("差引", "資本的収支不足額", "**補填財源で埋める**"),
    ]
    for cat, name, memo in ITEMS:
        key = name.strip().replace("　", "")
        vals = []
        for y in YEARS:
            d = k.get(y, {})
            if key == "収支差引":
                v = (d.get("水道事業収益", 0) - d.get("水道事業費用", 0)) if d else None
            elif key == "資本的収支不足額":
                v = (d.get("資本的収入", 0) - d.get("資本的支出", 0)) if d else None
            elif key == "他会計補助金・長期借入金":
                v = ((d.get("他会計補助金") or 0) + (d.get("他会計長期借入金") or 0)) if d else None
            else:
                v = d.get(key)
            vals.append(v)
        chg = f"{(vals[-1] / vals[0] - 1) * 100:+.1f}％" if (vals[0] and vals[-1]
                                                           and vals[0] > 0) else ""
        ws.append([cat, name] + vals + [chg, memo])
    style_header(ws)
    for r in range(2, ws.max_row + 1):
        for c in range(3, 3 + len(YEARS)):
            ws.cell(r, c).number_format = "#,##0"
            v = ws.cell(r, c).value
            ws.cell(r, c).fill = (WARN if isinstance(v, (int, float)) and v < 0 else CALC)
    body(ws, wrap=(len(YEARS) + 4,))
    widths(ws, [12, 26] + [14] * len(YEARS) + [10, 46])
    ws.freeze_panes = "C2"
    notes(ws, [
        "※ 出典は決算報告書（令和1〜7年度を1冊に収録）の決算額。単位：円。",
        "**※ 資本的収支は毎年度50〜90百万円の不足が生じ、損益勘定留保資金等で補填しています。**"
        "補填財源の内訳は補填財源明細書にあります。",
    ])
    return ws


def sheet_pl(wb, pls):
    ws = wb.create_sheet("04_損益の推移")
    ws.append(["区分", "項目", "令和6年度決算", "令和7年度決算", "令和8年度予算",
               "R7→R8", "備考"])
    ITEMS = [
        ("営業収益", "給水収益", "料金収入"),
        ("", "その他営業収益", ""),
        ("営業費用", "原水及び浄水費", ""),
        ("", "配水及び給水費", ""),
        ("", "総係費", ""),
        ("", "減価償却費", "**現金支出を伴わない。営業費用の過半**"),
        ("", "資産減耗費", ""),
        ("", "営業損失", "**給水収益では営業費用を賄えていない**"),
        ("営業外収益", "他会計補助金", "一般会計からの繰入"),
        ("", "長期前受金戻入", "**現金支出を伴わない収益。減価償却費に対応**"),
        ("", "雑収益", ""),
        ("営業外費用", "支払利息及び企業債取扱諸費", ""),
        ("", "雑支出", ""),
        ("結果", "経常利益", "**営業損失を繰入と戻入で埋めた結果**"),
        ("", "当年度純利益", ""),
        ("", "前年度繰越利益剰余金", ""),
        ("", "当年度未処分利益剰余金", ""),
    ]
    for cat, name, memo in ITEMS:
        vals = [pls[y].get(name) for y in ("R6", "R7", "R8")]
        if name == "営業損失":
            vals = [-abs(v) if v else v for v in vals]
        chg = ""
        if vals[1] and vals[2]:
            chg = f"{(vals[2] / vals[1] - 1) * 100:+.1f}％"
        ws.append([cat, name] + vals + [chg, memo])
    style_header(ws)
    for r in range(2, ws.max_row + 1):
        for c in range(3, 6):
            ws.cell(r, c).number_format = "#,##0"
            v = ws.cell(r, c).value
            ws.cell(r, c).fill = (WARN if isinstance(v, (int, float)) and v < 0 else CALC)
    body(ws, wrap=(7,))
    widths(ws, [12, 30, 16, 16, 16, 10, 52])
    notes(ws, [
        "※ 出典は損益計算書。単位：円。令和8年度は当初予算。",
        "**※ 令和8年度予算では営業損失が35.3百万円に拡大し、経常利益は7.4百万円に縮みます。**"
        "給水収益の減少と費用の増加がともに効いています。",
        "**※ 経常利益は出ていますが、その内訳は他会計補助金と長期前受金戻入です。"
        "料金収入だけで費用を賄える状態ではありません。**",
    ])
    return ws


def sheet_bs(wb, bs7, bs89):
    ws = wb.create_sheet("05_貸借対照表")
    ws.append(["区分", "項目", "令和7年度決算", "令和8年9月補正", "増減", "備考"])
    ITEMS = [
        ("資産", "固定資産合計", ""),
        ("", "現金預金", "**令和8年度に25.9百万円減る見込み**"),
        ("", "未収金合計", "うち過年度未収給水収益45.4百万円"),
        ("", "流動資産合計", ""),
        ("", "資産合計", ""),
        ("負債", "企業債（固定負債）", ""),
        ("", "他会計借入金", ""),
        ("", "固定負債合計", ""),
        ("", "流動負債合計", "うち企業債（1年以内償還）37.8百万円"),
        ("", "長期前受金", "**補助金等で取得した資産の未償却分**"),
        ("", "繰延収益合計", ""),
        ("", "負債合計", ""),
        ("資本", "資本金合計", ""),
        ("", "資本剰余金合計", ""),
        ("", "利益剰余金合計", ""),
        ("", "剰余金合計", ""),
        ("", "資本合計", ""),
        ("", "負債資本合計", ""),
    ]
    for cat, name, memo in ITEMS:
        a, b = bs7.get(name), bs89.get(name)
        ws.append([cat, name, a, b, (b - a) if (a and b) else None, memo])
    style_header(ws)
    for r in range(2, ws.max_row + 1):
        for c in range(3, 6):
            ws.cell(r, c).number_format = "#,##0"
            v = ws.cell(r, c).value
            ws.cell(r, c).fill = (WARN if c == 5 and isinstance(v, (int, float)) and v < 0
                                  else CALC)
    body(ws, wrap=(6,))
    widths(ws, [10, 26, 18, 18, 16, 58])
    notes(ws, [
        "※ 出典は貸借対照表。単位：円。令和8年9月補正は令和7年度決算を反映したもの。",
        "**※ 現金預金135.6百万円に対し、企業債の年間償還は元利で42.3百万円（令和8年度）です。**"
        "資本的収支の不足を補填し続けると、数年で現金が細ります。"
        "料金改定の必要性はここに表れます。",
    ])
    return ws


def sheet_shihon(wb, k, sk):
    ws = wb.create_sheet("06_資本的収支・企業債")
    ws.append(["年度", "建設改良費", "企業債（借入）", "企業債償還金", "他会計補助金",
               "工事負担金", "資本的収支不足額"])
    for y in YEARS:
        d = k.get(y, {})
        if not d:
            continue
        ws.append([f"令和{y[1]}年度", d.get("建設改良費"), d.get("企業債"),
                   d.get("企業債償還金"), d.get("他会計補助金"), d.get("工事負担金"),
                   (d.get("資本的収入", 0) - d.get("資本的支出", 0))])
    style_header(ws)
    num(ws, 2, range(2, 8), fill=CALC)
    for r in range(2, ws.max_row + 1):
        if isinstance(ws.cell(r, 7).value, (int, float)) and ws.cell(r, 7).value < 0:
            ws.cell(r, 7).fill = WARN
    body(ws)
    widths(ws, [12] + [16] * 6)
    ws.append([])
    ws.append(["【企業債の年度別償還予定（決算統計45表・千円）】"])
    ws.cell(ws.max_row, 1).font = Font(bold=True, size=10)
    r0 = ws.max_row + 1
    ws.append(["年度", "元金", "利子", "計"])
    style_header(ws, ws.max_row)
    for y, gan, ri in sk:
        ws.append([y, gan, ri, gan + ri])
    for r in range(r0 + 1, ws.max_row + 1):
        for c in range(2, 5):
            ws.cell(r, c).number_format = "#,##0"
            ws.cell(r, c).fill = CALC
            ws.cell(r, c).border = BORDER
        ws.cell(r, 1).border = BORDER
        for c in range(1, 5):
            ws.cell(r, c).font = Font(size=9)
    notes(ws, [
        "※ 上表の出典は決算報告書（決算額・円）、下表は決算統計45表（千円）。",
        "**※ 償還は令和8年度の42.3百万円を山として逓減し、令和11年度29.6百万円になります。**"
        "この間に新たな借入をどれだけ行うかが、料金水準に効きます。",
    ])
    return ws


def sheet_ryokin(wb, t7, t6, g, chk):
    ws = wb.create_sheet("07_料金体系と経営指標")
    ws.append(["区分", "項目", "令和6年度", "令和7年度", "単位", "備考"])
    ITEMS = [
        ("料金体系", "料金区分", "口径別", "口径別", "", "用途別ではなく口径別"),
        ("", "基本水量", t6["基本水量"], t7["基本水量"], "㎥", "基本料金に含まれる水量"),
        ("", "基本料金（税込）", t6["基本料金"], t7["基本料金"], "円",
         "**全国平均（10㎥で1,300円前後）を大きく上回る**"),
        ("", "超過料金（税込）", t6["超過料金"], t7["超過料金"], "円/㎥", ""),
        ("", "10㎥使用時（13mm）", t6["10㎥13mm"], t7["10㎥13mm"], "円", ""),
        ("", "10㎥使用時（20mm）", t6["10㎥20mm"], t7["10㎥20mm"], "円", ""),
        ("", "20㎥使用時（13mm）", t6["20㎥13mm"], t7["20㎥13mm"], "円", ""),
        ("", "20㎥使用時（20mm）", t6["20㎥20mm"], t7["20㎥20mm"], "円", ""),
        ("施設", "配水能力", t6["配水能力"], t7["配水能力"], "㎥/日", ""),
        ("", "一日最大配水量", t6["一日最大配水量"], t7["一日最大配水量"], "㎥/日", ""),
        ("", "施設利用率", t6["一日最大配水量"] / t6["配水能力"] * 100,
         t7["一日最大配水量"] / t7["配水能力"] * 100, "％",
         "**一日最大配水量÷配水能力。40％を下回る＝施設が過大**"),
        ("", "水利権", t6["水利権"], t7["水利権"], "㎥/日", ""),
        ("経営指標", "供給単価", g["R6"]["給水収益"] / g["R6"]["有収水量"],
         g["R7"]["給水収益"] / g["R7"]["有収水量"], "円/㎥", "給水収益÷有収水量"),
        ("", "給水原価", None, chk["給水原価"], "円/㎥",
         "（営業費用＋営業外費用）÷有収水量"),
        ("", "料金回収率", None, chk["料金回収率"], "％",
         "**供給単価÷給水原価。100％を下回る＝料金で費用を賄えていない**"),
        ("", "有収率", g["R6"]["有収率"], g["R7"]["有収率"], "％",
         "**配水量のうち料金の対象になった割合**"),
        ("", "給水普及率", g["R6"]["給水普及率"], g["R7"]["給水普及率"], "％", ""),
    ]
    for cat, name, a, b, u, memo in ITEMS:
        ws.append([cat, name, a, b, u, memo])
    style_header(ws)
    for r in range(2, ws.max_row + 1):
        for c in (3, 4):
            v = ws.cell(r, c).value
            if isinstance(v, (int, float)):
                ws.cell(r, c).number_format = ("0.0" if ws.cell(r, 5).value in
                                               ("％", "円/㎥") else "#,##0")
                ws.cell(r, c).fill = (KEY if ws.cell(r, 2).value in
                                      ("料金回収率", "施設利用率", "有収率") else CALC)
    body(ws, wrap=(6,))
    widths(ws, [12, 24, 14, 14, 8, 58])
    notes(ws, [
        "※ 料金体系・施設の出典は決算統計01表。経営指標は事業報告書・損益計算書から算定。",
        "**※ 料金回収率76.1％は、費用100円に対し料金で76円しか回収できていないことを表します。**"
        "差の24円は一般会計からの繰入（他会計補助金19.5百万円）と、"
        "現金支出を伴わない長期前受金戻入（27.1百万円）で埋めています。",
        "**※ 総括原価方式では、この給水原価が料金の出発点になります。**"
        "資産維持費をどこまで見るか、一般会計繰入をどこまで前提とするかで改定率が変わります。",
    ])
    return ws


def sheet_check(wb, chk):
    ws = wb.create_sheet("08_突合結果")
    ws.append(["№", "突合の内容", "値A", "値B", "差", "判定", "扱い"])
    for i, (name, a, b, judge, how) in enumerate(chk["items"], 1):
        ws.append([i, name, a, b, (b - a) if isinstance(a, (int, float))
                   and isinstance(b, (int, float)) else None, judge, how])
    style_header(ws)
    for r in range(2, ws.max_row + 1):
        for c in (3, 4, 5):
            if isinstance(ws.cell(r, c).value, (int, float)):
                ws.cell(r, c).number_format = "#,##0"
        ws.cell(r, 6).fill = {"一致": OK, "軽微": CALC, "要確認": WARN}[ws.cell(r, 6).value]
    body(ws, wrap=(2, 7))
    widths(ws, [4, 48, 17, 17, 13, 9, 50])
    notes(ws, [
        "※ 系列どうしの整合を点検しました。基準年度と出所を誤ると以降の算定が意味を失うため、"
        "算定に入る前に必ず行います。",
        "**※ 決算報告書・事業報告書・損益計算書・貸借対照表は相互に整合しています。**"
        "算定の基礎としてこのセットを用いて差し支えありません。",
    ])
    return ws


MISSING = [
    ("A", "給水人口・給水戸数の将来推計",
     "料金改定は将来の有収水量の見通しの上に立つ。給水人口は7年で11.6％減っており、"
     "この趨勢をどこまで延ばすかで総括原価の分母が変わる",
     "町（地域整備課）", "**最優先**",
     "令和7年度の経営戦略で推計しているはず。その推計値と前提"),
    ("B", "令和7年度策定の経営戦略（本文・投資試算・財政試算）",
     "本業務は経営戦略の続編にあたる。投資計画（更新需要）と財政計画が"
     "総括原価の資本費用の基礎になる",
     "町（地域整備課）", "**最優先**",
     "日本水工設計が令和7年度に策定。電子データで受領したい"),
    ("C", "現行の給水条例（料金表）と直近の改定経緯",
     "口径別の料金表の全体（13mm・20mm以外の口径）、"
     "および前回改定の時期・改定率・激変緩和の有無",
     "町（地域整備課）", "**最優先**",
     "決算統計には13mm・20mmしか載っていない"),
    ("D", "口径別・使用水量階層別の件数と水量（直近年度）",
     "**料金体系の設定にはこれが要る。**現行料金表を当てはめて収入を再現し、"
     "改定案を当てると収入がいくらになるかを計算する",
     "町（地域整備課・水道料金システム）", "**最優先**",
     "料金調定データ。12か月分・口径別・水量階層別"),
    ("E", "施設の更新需要（管路・浄水場・配水池の年次計画）",
     "資本費用の見通し。配水管4,000m、浄水場3か所、配水池4か所の"
     "更新時期と概算事業費",
     "町（地域整備課）", "高",
     "経営戦略の投資計画に含まれている可能性がある"),
    ("F", "一般会計繰入金の基準と今後の方針",
     "**他会計補助金19.5百万円を今後も見込めるかで改定率が大きく変わる。**"
     "繰出基準に基づくものか、基準外繰入かの区分",
     "町（財政担当）", "**最優先**",
     "総括原価から控除するかどうかの判断に直結する"),
    ("G", "有収率向上の取組と漏水調査の実績・計画",
     "有収率69.98％は全国平均を20ポイント下回る。"
     "改善すれば同じ配水量でより多くの収益が得られる",
     "町（地域整備課）", "高",
     "漏水調査の実施年度・箇所数・修繕実績"),
    ("H", "簡易水道事業の有無と統合の予定",
     "給水普及率77.9％で、行政区域内人口5,540人に対し町の総人口は約8,700人。"
     "**残る区域の扱いを確認する必要がある**",
     "町（地域整備課）", "高",
     "統合の予定があれば給水人口・施設・企業債が動く"),
    ("I", "水道料金算定要領（令和7年2月改定版）の適用範囲の確認",
     "資産維持率の設定、料金算定期間（3年か5年か）、"
     "総括原価に含める範囲について町の意向",
     "町（地域整備課）", "高",
     "仕様書の「営業費用及び資本費用等の算定、総括原価の算定、料金体系の設定」の前提"),
    ("J", "類似団体・県内他団体の料金水準",
     "改定案の妥当性の説明に用いる。10㎥・20㎥使用時の料金の比較",
     "受託者が収集（総務省の経営比較分析表・県のとりまとめ）", "中",
     "**福島県内の団体を用いる**"),
    ("K", "審議会・議会のスケジュールと住民説明の方針",
     "改定の手続。条例改正の時期から逆算して工程を引く",
     "町（地域整備課）", "高",
     "履行期間は令和9年3月31日まで"),
    ("L", "固定資産台帳（現在の帳簿価額・耐用年数・取得年度）",
     "資産維持費と減価償却費の見通し。受領資料には固定資産明細があるが、"
     "個別資産の台帳は未受領",
     "町（地域整備課）", "中",
     "更新需要の推計に用いる"),
]


def sheet_missing(wb):
    ws = wb.create_sheet("09_不足資料")
    ws.append(["№", "資料・確認事項", "なぜ要るか", "宛先", "優先", "備考", "状態", "受領日"])
    for no, name, why, ate, pri, memo in MISSING:
        ws.append([no, name, why, ate, pri, memo, "未受領", None])
    style_header(ws)
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 7).fill = WARN
        if "最優先" in str(ws.cell(r, 5).value):
            ws.cell(r, 5).fill = KEY
    body(ws, wrap=(2, 3, 6))
    widths(ws, [4, 36, 50, 26, 11, 40, 9, 11])
    notes(ws, [
        "※ 「状態」欄を受領済みに書き換えて管理します。",
        "**※ Ａ〜Ｄ・Ｆの5件は、これがないと総括原価の算定に入れません。**"
        "とくにＤ（口径別・使用水量階層別の件数と水量）は料金体系の設定に必須で、"
        "仕様書の工数の56.6％を占める作業の前提です。",
        "※ 受領済みの44ファイルで、営業費用・資本費用の実績と企業債の償還見通しは"
        "そろっています。足りないのは将来の見通しと料金の内訳です。",
    ])
    return ws


# ================================================================ main

def main():
    g = read_gyomu()
    k = read_kessan()
    pls = {"R6": read_pl(F_PL6), "R7": read_pl(F_PL7), "R8": read_pl(F_PL8)}
    bs7, bs89 = read_bs(F_BS7), read_bs(F_BS89)
    t7, t6 = read_tokei(F_TK7), read_tokei(F_TK6)
    sk = read_shokan()

    r7g, r7p, r7k = g["R7"], pls["R7"], k["R7"]
    genka = (r7p["営業損失"] + r7p["給水収益"] + r7p.get("その他営業収益", 0)
             + r7p["支払利息及び企業債取扱諸費"] + r7p.get("雑支出", 0)) / r7g["有収水量"]
    tanka = r7g["給水収益"] / r7g["有収水量"]
    chk = {
        "給水原価": genka, "料金回収率": tanka / genka * 100,
        "営業損失": -r7p["営業損失"], "経常利益": r7p["経常利益"],
        "items": [
            ("給水収益　損益計算書 対 事業報告書", r7p["給水収益"], r7g["給水収益"],
             "一致", "算定の基礎として用いる"),
            ("営業収益　損益計算書（税抜）＋消費税 対 決算報告書（税込）",
             r7p["給水収益"] + r7p.get("その他営業収益", 0) + round(r7p["給水収益"] * 0.10),
             r7k["営業収益"],
             "一致" if abs(r7p["給水収益"] + r7p.get("その他営業収益", 0)
                         + round(r7p["給水収益"] * 0.10) - r7k["営業収益"]) <= 1 else "要確認",
             "決算報告書は税込、損益計算書は税抜。"
             "差10,654,701円は給水収益106,547,010円の10％で、消費税分と完全に一致する。"
             "**総括原価は税抜で組む**"),
            ("営業外収益　損益計算書 対 決算報告書",
             sum(r7p.get(x, 0) for x in ("他会計補助金", "雑収益", "長期前受金戻入",
                                         "受取利息及び配当金")),
             r7k["営業外収益"], "一致", "―"),
            ("当年度未処分利益剰余金　損益計算書 対 貸借対照表",
             r7p["当年度未処分利益剰余金"], bs7["当年度未処分利益剰余金"],
             "一致", "―"),
            ("貸借対照表　資産合計 対 負債資本合計",
             bs7["資産合計"], bs7["負債資本合計"], "一致", "貸借は一致している"),
            ("企業債（固定負債）　貸借対照表 対 決算統計22表",
             bs7["企業債（固定負債）"], 309322289, "一致", "―"),
            ("有収水量　事業報告書 対 決算統計01表（10㎥単位）",
             r7g["有収水量"], t7["年間総有収水量"] * 10, "軽微",
             "決算統計は百単位に丸めている。差490㎥（0.1％）"),
            ("年間総配水量　事業報告書 対 決算統計01表（10㎥単位）",
             g["R7"]["配水量"], t7["年間総配水量"] * 10, "軽微",
             "同上。差432㎥（0.07％）"),
            ("現在給水人口　事業報告書 対 決算統計01表（令和6年度）",
             g["R6"]["年度末給水人口"], t6["現在給水人口"], "要確認",
             "**52人の差。基準日の定義（年度末か3月31日か）を確認する**"),
            ("固定資産合計　貸借対照表 対 決算統計22表（千円）",
             round(bs7["固定資産合計"] / 1000), 1547923, "要確認",
             "**727千円の差。決算統計の提出版と決算書の版が異なる可能性がある**"),
        ],
    }

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    sheet_intro(wb, g, chk)
    sheet_files(wb)
    sheet_gyomu(wb, g)
    sheet_shushi(wb, k)
    sheet_pl(wb, pls)
    sheet_bs(wb, bs7, bs89)
    sheet_shihon(wb, k, sk)
    sheet_ryokin(wb, t7, t6, g, chk)
    sheet_check(wb, chk)
    sheet_missing(wb)
    OUT.mkdir(parents=True, exist_ok=True)
    p = OUT / f"小野町水道_受領資料の確認と整理_{ASOF}.xlsx"
    wb.save(p)

    print(f"生成: {p.name}（{len(wb.sheetnames)}シート）")
    print(f"\n令和7年度決算")
    print(f"  給水収益   {r7g['給水収益']:>14,.0f} 円")
    print(f"  有収水量   {r7g['有収水量']:>14,.0f} ㎥")
    print(f"  供給単価   {tanka:>14,.1f} 円/㎥")
    print(f"  給水原価   {genka:>14,.1f} 円/㎥")
    print(f"  料金回収率 {tanka / genka * 100:>14,.1f} ％")
    print(f"  営業損失   {-r7p['営業損失']:>14,.0f} 円")
    print(f"  経常利益   {r7p['経常利益']:>14,.0f} 円")
    print(f"\n突合 {len(chk['items'])}件　"
          f"一致{sum(1 for x in chk['items'] if x[3] == '一致')}／"
          f"軽微{sum(1 for x in chk['items'] if x[3] == '軽微')}／"
          f"要確認{sum(1 for x in chk['items'] if x[3] == '要確認')}")
    print(f"不足資料 {len(MISSING)}件（最優先 "
          f"{sum(1 for x in MISSING if '最優先' in x[4])}件）")


if __name__ == "__main__":
    main()
