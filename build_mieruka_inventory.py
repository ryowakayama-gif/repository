# -*- coding: utf-8 -*-
"""
地域包括ケア「見える化」システム ダウンロードデータ 棚卸し表ジェネレータ
（音威子府村高齢者福祉計画・第10期介護保険事業計画 策定業務）

入力：見える化システムからDLしたExcel群（カテゴリ1〜17のフォルダ構成のまま）
出力：output/見える化システムDLデータ棚卸し_音威子府村.xlsx ＋ 同名.md

使い方：
    python3 build_mieruka_inventory.py [<DLデータのルートディレクトリ>]
"""

import os
import re
import sys
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SRC = sys.argv[1] if len(sys.argv) > 1 else os.environ.get("MIERUKA_SRC", "")
OUT_DIR = "/home/user/repository/output"
FILENAME = "見える化システムDLデータ棚卸し_音威子府村.xlsx"
TARGET = "音威子府村"
PEERS = ["音威子府村", "中川町", "美深町", "名寄市", "北海道", "全国"]

FONT = "游ゴシック"
COLORS = {
    "header": "1F3864", "subhead": "2E75B6", "band": "DDEBF7", "alt": "F7FAFC",
    "ok": "70AD47", "warn": "ED7D31", "ng": "C00000", "gray": "7F7F7F",
}
THIN = Side(border_style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

BLANK = {"", "-", "‐", "－", "ー", "***", "*", "…", "N/A"}
is_blank = lambda v: v is None or str(v).strip() in BLANK
is_ph = lambda v: v is not None and re.fullmatch(r"(行|列)\d+", str(v).strip()) is not None


# ============================================================
# スタイル
# ============================================================
def style_title(cell, text, fill=COLORS["header"], size=14):
    cell.value = text
    cell.font = Font(name=FONT, size=size, bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor=fill)
    cell.alignment = Alignment(vertical="center", horizontal="left", indent=1)


def style_subhead(cell, text, fill=COLORS["subhead"]):
    cell.value = text
    cell.font = Font(name=FONT, size=11, bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor=fill)
    cell.alignment = Alignment(vertical="center", horizontal="left", indent=1)


def style_header_row(ws, row, headers, fill=COLORS["header"]):
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=fill)
        c.alignment = Alignment(vertical="center", horizontal="center", wrap_text=True)
        c.border = BORDER


def style_data_cell(cell, alt=False, center=False):
    cell.font = Font(name=FONT, size=10)
    cell.alignment = Alignment(vertical="center",
                               horizontal="center" if center else "left",
                               wrap_text=True, indent=0 if center else 1)
    cell.border = BORDER
    if alt:
        cell.fill = PatternFill("solid", fgColor=COLORS["alt"])


def set_col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ============================================================
# 見える化システムのブックを読む共通処理
# ============================================================
def table_sheet(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = next((wb[s] for s in wb.sheetnames if s.startswith("表形式")), wb[wb.sheetnames[-1]])
    grid = [list(r) for r in ws.iter_rows(max_row=320, max_col=400, values_only=True)]
    out_date = None
    if grid and grid[0] and grid[0][0] and "出力日" in str(grid[0][0]):
        out_date = str(grid[0][0]).replace("出力日:", "").strip()
    wb.close()
    return grid, out_date


def find_period_row(grid):
    for i, row in enumerate(grid[:12]):
        n = sum(1 for v in row
                if v is not None and re.search(r"(年度|年|末|^\s*(19|20)\d{2}\s*$|R\d|H\d)", str(v)))
        if n >= 3:
            return i
    return 1


def read_timeseries(path, region=TARGET):
    """時系列ファイル → [(系列名, [(時点, 値), ...]), ...]"""
    grid, _ = table_sheet(path)
    hr = find_period_row(grid)
    per = ["" if v is None or is_ph(v) else str(v).replace("\n", "").strip() for v in grid[hr]]
    out = []
    for row in grid[hr + 1:]:
        if len(row) < 2 or row[1] is None or not str(row[1]).strip().startswith(region):
            continue
        name = str(row[2]).strip() if len(row) > 2 and row[2] else ""
        out.append((name, [(per[j], row[j]) for j in range(len(per)) if per[j] and j < len(row)]))
    return out


def read_by_region(path, regions=PEERS):
    """地域別ファイル → [(指標名, {地域: 値}), ...]"""
    grid, _ = table_sheet(path)
    hr = None
    for i, row in enumerate(grid[:10]):
        if any(v is not None and str(v).strip() == TARGET for v in row):
            hr = i
            break
    if hr is None:
        return []
    idx = {}
    for j, v in enumerate(grid[hr]):
        if v is None:
            continue
        s = str(v).strip()
        for rg in regions:
            if s == rg or s.startswith(rg):
                idx.setdefault(rg, j)
    out = []
    for row in grid[hr + 1:]:
        lab = row[1] if len(row) > 1 else None
        if lab is None or is_ph(lab) or str(lab).startswith("（"):
            continue
        out.append((str(lab).strip(),
                    {k: (row[j] if len(row) > j else None) for k, j in idx.items()}))
    return out


# ============================================================
# 全ファイル監査
# ============================================================
def audit_one(path):
    fn = os.path.basename(path)
    kind = "時系列" if "時系列" in fn else ("地域別" if "地域別" in fn else "不明")
    rec = dict(file=fn, cat=os.path.basename(os.path.dirname(path)), kind=kind,
               marked_x=any(c in fn for c in "✕×✖"), dup=bool(re.search(r"\(\d+\)", fn)))
    m = re.search(r"_((?:19|20)\d{2,6})_", fn)
    rec["vintage"] = m.group(1) if m else ""
    m = re.match(r"([A-Z]\d+(?:-[a-z])?)_", fn)
    rec["code"] = m.group(1) if m else ""
    nm = fn[len(rec["code"]) + 1:] if rec["code"] else fn
    rec["indicator"] = re.sub(r"_((?:19|20)\d{2,6})$", "", re.sub(r"_(?:時系列|地域別).*$", "", nm))

    try:
        grid, out_date = table_sheet(path)
    except Exception as e:
        rec.update(status="読込エラー", note=f"{type(e).__name__}", n_cells=0, n_value=0, fill=0.0)
        return rec
    rec["out_date"] = out_date or ""

    labels, vals = [], []
    if kind == "地域別":
        hr = ci = None
        exact = False
        for i, row in enumerate(grid[:12]):
            for j, v in enumerate(row):
                if v is None:
                    continue
                s = str(v).strip()
                if s == TARGET:
                    hr, ci, exact = i, j, True
                    break
                if s.startswith(TARGET) and hr is None:
                    hr, ci = i, j
            if exact:
                break
        if hr is None:
            has_axis = any(not is_blank(v) and not is_ph(v) for row in grid[:6] for v in row[3:])
            rec.update(status="対象自治体なし" if has_axis else "空DL（地域軸なし）",
                       n_cells=0, n_value=0, fill=0.0)
            return rec
        rec["n_regions"] = sum(1 for v in grid[hr][3:] if not is_blank(v) and not is_ph(v))
        cols = [j for j, v in enumerate(grid[hr])
                if v is not None and str(v).strip().startswith(TARGET)]
        for row in grid[hr + 1:]:
            lab = row[1] if len(row) > 1 else None
            if lab is None or is_ph(lab) or str(lab).startswith("（"):
                continue
            for j in cols:
                labels.append(f"{lab}｜{str(grid[hr][j]).strip()}")
                vals.append(row[j] if len(row) > j else None)
        if not labels:
            rec.update(status="空DL（データ行なし）", n_cells=0, n_value=0, fill=0.0)
            return rec
    else:
        hr = find_period_row(grid)
        per = ["" if v is None or is_ph(v) else str(v).replace("\n", "").strip() for v in grid[hr]]
        named = [r for r in grid[hr + 1:] if len(r) > 1 and r[1] is not None and not is_ph(r[1])]
        rec["n_regions"] = len({str(r[1]).strip() for r in named})
        if not named:
            rec.update(status="空DL（データ行なし）", n_cells=0, n_value=0, fill=0.0)
            return rec
        tr = [r for r in named if str(r[1]).strip().startswith(TARGET)]
        if not tr:
            rec.update(status="対象自治体なし", n_cells=0, n_value=0, fill=0.0)
            return rec
        for row in tr:
            sname = str(row[2]).strip() if len(row) > 2 and row[2] else ""
            for j, p in enumerate(per):
                if not p or j >= len(row):
                    continue
                labels.append(f"{sname}｜{p}" if sname else p)
                vals.append(row[j])

    nb = [(L, v) for L, v in zip(labels, vals) if not is_blank(v)]
    rec.update(n_cells=len(labels), n_value=len(nb),
               fill=round(len(nb) / len(labels), 3) if labels else 0.0,
               status="データあり" if nb else "全欠測（‐表記）")
    if nb:
        rec["first"], rec["last"] = nb[0][0], nb[-1][0]
    return rec


def run_audit(src):
    files = []
    for d in sorted(os.listdir(src), key=lambda x: int(x.split(".")[0]) if x.split(".")[0].isdigit() else 999):
        p = os.path.join(src, d)
        if os.path.isdir(p):
            files += [os.path.join(p, f) for f in sorted(os.listdir(p))
                      if f.endswith(".xlsx") and not f.startswith("~")]
    recs = []
    for i, f in enumerate(files, 1):
        recs.append(audit_one(f))
        if i % 100 == 0:
            print(f"    監査 {i}/{len(files)}", file=sys.stderr)
    return recs


# ============================================================
# 主要指標の抽出定義
#   (分野, 相対パス, 形式, 系列フィルタ, 単位, 備考)
# ============================================================
KEY_TS = [
    ("人口", "2.人口/A1_総人口_時系列.xlsx", "総人口", "人",
     "実績（2000〜）＋社人研推計（〜2050）"),
    ("人口", "2.人口/A2_高齢化率_時系列.xlsx", "高齢化率", "％", "65歳以上人口割合"),
    ("被保険者", "4.第1号被保険者/B1_第１号被保険者数_時系列.xlsx", "第１号被保険者数", "人", ""),
    ("認定", "5.要介護（要支援）認定/B3-a_要支援・要介護認定者数（要介護度別）_時系列.xlsx",
     "合計認定者数", "人", "第1号被保険者分"),
    ("認定", "5.要介護（要支援）認定/B4-a_認定率（要介護度別）_時系列.xlsx",
     "合計認定率", "％", ""),
    ("保険料", "6.介護保険料/C1_第１号被保険者１人あたり保険給付月額・第１号保険料月額・必要保険料月額_時系列✕.xlsx",
     None, "円", "保険給付月額／第1号保険料月額／必要保険料月額"),
    ("財政", "14.介護保険特別会計経理状況/D47_歳入_時系列.xlsx", "合計", "円", "歳入合計"),
    ("財政", "14.介護保険特別会計経理状況/D47-k_歳入（繰越金）_時系列.xlsx", None, "円", ""),
    ("財政", "14.介護保険特別会計経理状況/D48-b_歳出（保険給付費）_時系列.xlsx", None, "円", ""),
    ("資源", "9.入所（利用）定員/D27_定員（通所系サービス別）_時系列.xlsx",
     "定員合計（通所系サービス）", "人", "施設・居住系は村内になく全欠測"),
    ("包括", "10.地域包括支援センター/F16_センター設置数_時系列.xlsx", "センター数（総数）", "か所", ""),
    ("包括", "10.地域包括支援センター/F16_センター設置数_時系列.xlsx", "日常生活圏域数", "圏域", ""),
]

KEY_REGION = [
    ("認定", "5.要介護（要支援）認定/B4-a_認定率（要介護度別）_2026_地域別.xlsx",
     ["合計認定率"], "％", "令和8年4月末"),
    ("認定", "5.要介護（要支援）認定/B3-a_要支援・要介護認定者数（要介護度別）_2026_地域別.xlsx",
     ["認定者数（要支援１）", "認定者数（要支援２）", "認定者数（要介護１）", "認定者数（要介護２）",
      "認定者数（要介護３）", "認定者数（要介護４）", "認定者数（要介護５）", "合計認定者数"],
     "人", "令和8年4月末"),
    ("人口", "2.人口/A2_高齢化率_2020_地域別.xlsx", ["高齢化率"], "％", "2020年"),
    ("総合事業", "11.介護予防・日常生活支援総合事業/F3_週1回以上の通いの場の箇所数_2020_地域別.xlsx",
     None, "か所", "令和2年度（システム最新値）"),
    ("総合事業", "11.介護予防・日常生活支援総合事業/F1_週１回以上の通いの場の参加率_2020_地域別.xlsx",
     None, "％／人", "令和2年度（システム最新値）"),
]


def extract_key(src):
    """主要指標を tidy 形式で返す：(分野, コード, 指標, 系列, 時点/地域, 値, 単位, 出典, 備考)"""
    rows = []
    for field, rel, want, unit, note in KEY_TS:
        path = os.path.join(src, rel)
        if not os.path.exists(path):
            continue
        code = re.match(r"([A-Z]\d+(?:-[a-z])?)_", os.path.basename(rel))
        code = code.group(1) if code else ""
        ind = re.sub(r"_(?:時系列|地域別).*$", "", os.path.basename(rel).split("_", 1)[1])
        try:
            for sname, vals in read_timeseries(path):
                if want and want not in sname:
                    continue
                for p, v in vals:
                    if is_blank(v):
                        continue
                    rows.append((field, code, ind, sname, p, v, unit, "時系列", note))
        except Exception as e:
            print(f"    ! 抽出失敗 {rel}: {e}", file=sys.stderr)
    return rows


def extract_peer(src):
    """近隣比較：(分野, 指標, 系列, {地域: 値}, 単位, 時点)"""
    rows = []
    for field, rel, wants, unit, when in KEY_REGION:
        path = os.path.join(src, rel)
        if not os.path.exists(path):
            continue
        try:
            for lab, d in read_by_region(path):
                if wants and lab not in wants:
                    continue
                if not wants and lab.startswith("行"):
                    continue
                u = unit
                if "／" in unit:                      # 単位が混在する指標は行ラベルから判定
                    u = "％" if "率" in lab else ("か所" if "箇所" in lab else "人")
                rows.append((field, os.path.basename(rel), lab, d, u, when))
        except Exception as e:
            print(f"    ! 比較抽出失敗 {rel}: {e}", file=sys.stderr)
    return rows


# ============================================================
# 確認結果・要対応事項（監査から機械的に導けない判断部分）
# ============================================================
FINDINGS = [
    ("データ全般", "重要",
     "741ファイルすべて音威子府村を対象自治体として出力済み。うち521ファイルで値を取得できており、"
     "現状分析（WBS 2）と推計（WBS 3・4）の基礎データは揃っている。",
     "そのまま利用可"),
    ("データ全般", "注意",
     "741ファイル中162ファイルは同一指標の重複配置（『1.地域分析』が他カテゴリの指標を再掲）。"
     "実質のユニーク指標×形式は579種。",
     "集計時は指標コードで名寄せする"),
    ("データ全般", "注意",
     "『(1)』付き19ファイルは再ダウンロード分。データ内容は同一で出力日時のみ相違。",
     "いずれか一方を採用し、他は削除して混乱を防ぐ"),
    ("在宅医療", "要対応",
     "L16〜L28（訪問診療・訪問看護・往診・退院支援・看取り等のレセプト系指標）90ファイルは、"
     "地域軸は出力されているがデータ行が存在しない空振りDL。市町村単位では出力されないと考えられる。",
     "二次医療圏（上川北部）単位で再取得するか、北海道の医療計画・NDBオープンデータで代替（WBS 5.3.5 在宅医療・介護連携の記述に必要）"),
    ("認知症", "要対応",
     "13.認知症対策の14ファイルはすべて『-』。認知症サポーター・各種研修修了者数は"
     "都道府県／政令市単位でしか公表されていない。",
     "村の実績データで補完（WBS 1.3.1 貸与資料リストに追加）"),
    ("総合事業", "要対応",
     "F1〜F14（通いの場）の最新値は令和2年度、F28〜F40（総合事業実施件数）は令和2年度末で更新が止まっている。"
     "第10期計画の現状分析には5年以上古い。",
     "村の総合事業実績（令和3〜7年度）で補完（WBS 2.2.2）"),
    ("サービス資源", "重要",
     "施設サービス・居住系サービスの定員はすべて欠測＝村内に施設・居住系サービスが存在しない。"
     "村内資源は地域密着型通所介護（定員15人）と直営の地域包括支援センター1か所のみ。",
     "見込量推計では圏域外（近隣市町村）利用を明示的に分解して推計（WBS 4.1.2〜4.1.4）"),
    ("認定", "重要",
     "合計認定率9.4%（令和8年4月末）は全国20.2%・北海道21.7%・近隣（美深20.0%／中川21.0%）の半分以下。"
     "かつ要支援1・2の認定者が0人で、認定者19人中16人が要介護2以上と重度に偏っている。",
     "『健康だから低い』のか『サービスがないため申請に至っていない』のかを"
     "ニーズ調査結果と突合して見極める（WBS 2.3.1／2.4.2）。認定率固定での推計は需要を過小評価する恐れ"),
    ("保険料", "重要",
     "必要保険料月額はR7時点5,162円で、実際の第1号保険料月額3,600円を1,562円上回る（R6も4,145円対3,600円）。"
     "第1号被保険者1人あたり保険給付月額もR5 13,498円→R7 20,654円と1.5倍に急増。",
     "第10期は相応の引上げ圧力がある前提で複数パターンを試算（WBS 4.3.2）"),
    ("財政", "注意",
     "介護保険特別会計の繰越金がR4年3月末6,166千円→R6年3月末24,104千円と3年で約4倍。"
     "一方、見える化システムに介護給付費準備基金の残高そのものは収録されていない"
     "（歳出『基金積立金』は利子相当の少額のみ）。",
     "基金残高の推移は村から入手（WBS 1.3.1／4.3.2）"),
    ("推計", "重要",
     "A1総人口は2000年実績から2050年推計まで収録（社人研準拠）。"
     "総人口は2026年584人→2050年328人。高齢化率は2026年31.2%をピークに2050年23.5%まで低下する見通し。",
     "人口推計（WBS 3.1）のベースラインとして利用可。高齢者数自体が減る点を計画に明記"),
    ("運用", "確認",
     "ファイル名の「✕」印31件（うち30件が時系列）は、データ内容としては欠測が多いわけではなく"
     "（データ取得率90%／全体69%）、意図が判別できない。",
     "印の意味をご教示ください（不要＝除外指標なのか、確認済みマークなのか）"),
]

NEXT_ACTIONS = [
    ("A-1", "在宅医療系（L16〜L28）の再取得または代替データの確保", "受託者", "第1回打合せまで",
     "二次医療圏単位での再出力を試行。取得できない場合は北海道の医療計画等で代替", "WBS 5.3.5"),
    ("A-2", "認知症関連の実績を村から入手", "村", "令和8年9月末", "サポーター養成数、初期集中支援チーム、"
     "認知症カフェ等の実施状況", "WBS 2.2.3／1.3.1"),
    ("A-3", "総合事業・通いの場の令和3〜7年度実績を村から入手", "村", "令和8年9月末",
     "見える化システムは令和2年度で更新停止", "WBS 2.2.2"),
    ("A-4", "介護給付費準備基金の残高推移を村から入手", "村", "令和8年10月上旬",
     "保険料算定・財政シミュレーションの前提", "WBS 2.3.4／4.3.2"),
    ("A-5", "認定率が全国の半分である要因の分析", "受託者", "令和8年11月",
     "ニーズ調査結果・認定申請の状況と突合。潜在需要の有無を確認", "WBS 2.3.1／2.4.2"),
    ("A-6", "圏域外サービス利用の実態把握", "受託者・村", "令和8年10月",
     "村内に施設・居住系がないため、近隣市町村の施設利用状況を給付実績から分解", "WBS 2.1.2／4.1.2"),
    ("A-7", "重複ファイルの整理（(1)付き19件・カテゴリ間重複162件）", "受託者", "令和8年9月中",
     "指標コードで名寄せしたマスタを作成", "WBS 1.3.2"),
    ("A-8", "「✕」印の意味を村／社内で確認", "受託者", "第1回打合せまで",
     "除外指標か確認済みマークかで取扱いが変わる", "WBS 1.3.2"),
]


# ============================================================
# シート生成
# ============================================================
STATUS_COLOR = {
    "データあり": COLORS["ok"],
    "全欠測（‐表記）": COLORS["warn"],
    "空DL（データ行なし）": COLORS["ng"],
    "空DL（地域軸なし）": COLORS["ng"],
    "対象自治体なし": COLORS["gray"],
    "読込エラー": COLORS["ng"],
}


def add_summary(wb, recs, src):
    import collections
    ws = wb.active
    ws.title = "00_サマリ"
    set_col_widths(ws, [16, 10, 74, 52, 14])
    ws.row_dimensions[1].height = 34
    ws.merge_cells("A1:E1")
    style_title(ws["A1"], "地域包括ケア「見える化」システム　ダウンロードデータ棚卸し　／　音威子府村")
    ws.merge_cells("A2:E2")
    ws["A2"] = (f"対象：{len(recs)}ファイル（カテゴリ1〜17）　"
                f"出力日：{'／'.join(sorted({r.get('out_date','') for r in recs if r.get('out_date')}))}　"
                "／　本表は自動監査の結果に所見を加えたもの")
    ws["A2"].font = Font(name=FONT, size=9, italic=True, color="595959")
    ws["A2"].alignment = Alignment(vertical="center", indent=1)

    st = collections.Counter(r["status"] for r in recs)
    r0 = 4
    ws.merge_cells(f"A{r0}:E{r0}")
    style_subhead(ws.cell(row=r0, column=1), "1. 取得状況の全体像")
    r0 += 1
    style_header_row(ws, r0, ["区分", "件数", "内容", "対応", "割合"])
    r0 += 1
    rows = [
        ("データあり", st["データあり"], "音威子府村の値が1つ以上入っている。現状分析・推計にそのまま利用可", "利用可"),
        ("全欠測（‐表記）", st["全欠測（‐表記）"],
         "地域軸に音威子府村はあるが値がすべて『-』。村内に当該サービスがない／市町村単位で非公表", "要因確認"),
        ("空DL（データ行なし）", st["空DL（データ行なし）"],
         "地域軸は出力されているがデータ行が無い。L16〜L28のレセプト系指標", "再取得・代替"),
        ("対象自治体なし", st["対象自治体なし"],
         "地域軸に音威子府村が現れない（L系時系列）。市町村単位で提供されない指標", "代替データ"),
    ]
    total = len(recs)
    for name, n, desc, act in rows:
        ws.cell(row=r0, column=1, value=name)
        ws.cell(row=r0, column=2, value=n)
        ws.cell(row=r0, column=3, value=desc)
        ws.cell(row=r0, column=4, value=act)
        ws.cell(row=r0, column=5, value=n / total if total else 0)
        for c in range(1, 6):
            style_data_cell(ws.cell(row=r0, column=c), alt=(r0 % 2 == 0), center=(c in (2, 4, 5)))
        c1 = ws.cell(row=r0, column=1)
        c1.fill = PatternFill("solid", fgColor=STATUS_COLOR.get(name, COLORS["gray"]))
        c1.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
        ws.cell(row=r0, column=5).number_format = "0.0%"
        ws.row_dimensions[r0].height = 30
        r0 += 1
    ws.cell(row=r0, column=1, value="合計")
    ws.cell(row=r0, column=2, value=f"=SUM(B{r0-len(rows)}:B{r0-1})")
    ws.cell(row=r0, column=5, value=f"=SUM(E{r0-len(rows)}:E{r0-1})")
    for c in range(1, 6):
        cc = ws.cell(row=r0, column=c)
        cc.fill = PatternFill("solid", fgColor=COLORS["header"])
        cc.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
        cc.border = BORDER
        cc.alignment = Alignment(vertical="center", horizontal="center")
    ws.cell(row=r0, column=5).number_format = "0.0%"
    r0 += 2

    ws.merge_cells(f"A{r0}:E{r0}")
    style_subhead(ws.cell(row=r0, column=1), "2. 確認結果（所見）")
    r0 += 1
    style_header_row(ws, r0, ["分野", "区分", "確認した内容", "計画策定上の取扱い", ""])
    r0 += 1
    lvl_color = {"重要": COLORS["ng"], "要対応": COLORS["warn"], "注意": COLORS["subhead"], "確認": COLORS["gray"]}
    for field, lvl, body, act in FINDINGS:
        ws.cell(row=r0, column=1, value=field)
        ws.cell(row=r0, column=2, value=lvl)
        ws.cell(row=r0, column=3, value=body)
        ws.merge_cells(start_row=r0, start_column=4, end_row=r0, end_column=5)
        ws.cell(row=r0, column=4, value=act)
        for c in range(1, 6):
            style_data_cell(ws.cell(row=r0, column=c), alt=(r0 % 2 == 0), center=(c == 2))
        lc = ws.cell(row=r0, column=2)
        lc.fill = PatternFill("solid", fgColor=lvl_color[lvl])
        lc.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
        ws.row_dimensions[r0].height = 46
        r0 += 1
    r0 += 1

    ws.merge_cells(f"A{r0}:E{r0}")
    style_subhead(ws.cell(row=r0, column=1), "3. 次のアクション")
    r0 += 1
    style_header_row(ws, r0, ["No.", "内容", "担当", "期限目安", "WBS"])
    r0 += 1
    for no, body, who, due, detail, wbs in NEXT_ACTIONS:
        ws.cell(row=r0, column=1, value=no)
        ws.cell(row=r0, column=2, value=f"{body}\n（{detail}）")
        ws.cell(row=r0, column=3, value=who)
        ws.cell(row=r0, column=4, value=due)
        ws.cell(row=r0, column=5, value=wbs)
        for c in range(1, 6):
            style_data_cell(ws.cell(row=r0, column=c), alt=(r0 % 2 == 0), center=(c in (1, 3, 4, 5)))
        ws.row_dimensions[r0].height = 34
        r0 += 1

    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"
    return ws


def add_file_list(wb, recs):
    ws = wb.create_sheet("01_ファイル一覧")
    headers = ["No.", "カテゴリ", "指標コード", "指標名", "形式", "データ年次",
               "ステータス", "取得セル数", "値あり", "充足率", "最初の時点／系列",
               "最後の時点／系列", "重複", "✕印", "ファイル名"]
    set_col_widths(ws, [5, 26, 9, 46, 8, 11, 17, 9, 8, 9, 26, 26, 7, 6, 62])
    ws.row_dimensions[1].height = 28
    ws.merge_cells("A1:O1")
    style_title(ws["A1"], "ダウンロードファイル一覧（全件監査結果）")
    ws.merge_cells("A2:O2")
    ws["A2"] = "充足率＝音威子府村の値が入っているセル ÷ 取得対象セル。オートフィルタでステータス別の絞り込みが可能。"
    ws["A2"].font = Font(name=FONT, size=9, color="595959")
    ws["A2"].alignment = Alignment(vertical="center", indent=1)
    style_header_row(ws, 4, headers)
    ws.row_dimensions[4].height = 30

    r = 5
    for i, x in enumerate(recs, 1):
        vals = [i, x["cat"], x["code"], x["indicator"], x["kind"], x["vintage"],
                x["status"], x.get("n_cells", 0), x.get("n_value", 0), x.get("fill", 0.0),
                x.get("first", ""), x.get("last", ""),
                "○" if x["dup"] else "", "✕" if x["marked_x"] else "", x["file"]]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            style_data_cell(cell, alt=(r % 2 == 0), center=(c in (1, 3, 5, 6, 7, 8, 9, 10, 13, 14)))
        ws.cell(row=r, column=10).number_format = "0.0%"
        sc = ws.cell(row=r, column=7)
        sc.fill = PatternFill("solid", fgColor=STATUS_COLOR.get(x["status"], COLORS["gray"]))
        sc.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
        r += 1

    ws.freeze_panes = "D5"
    ws.auto_filter.ref = f"A4:O{r-1}"
    ws.sheet_view.showGridLines = False
    ws.print_title_rows = "4:4"
    return ws


def add_gap_sheet(wb, recs):
    ws = wb.create_sheet("02_欠測・要対応")
    set_col_widths(ws, [26, 9, 50, 8, 17, 44, 16])
    ws.row_dimensions[1].height = 28
    ws.merge_cells("A1:G1")
    style_title(ws["A1"], "取得できていない指標と対応方針")
    n_gap = sum(1 for x in recs if x["status"] != "データあり")
    ws.merge_cells("A2:G2")
    ws["A2"] = f"『データあり』以外の{n_gap}ファイルを原因別に整理。代替データの手当てが必要なものを明示。"
    ws["A2"].font = Font(name=FONT, size=9, color="595959")
    ws["A2"].alignment = Alignment(vertical="center", indent=1)

    def cause(x):
        if x["status"] == "全欠測（‐表記）":
            if x["cat"].startswith("13."):
                return "都道府県／政令市単位のみ公表", "村の実績で補完"
            if any(k in x["indicator"] for k in ("定員", "施設", "居住系", "認知症対応型共同生活")):
                return "村内に当該サービスなし", "圏域外利用として推計"
            return "村内に利用実績なし（該当者0）", "見込量は0または近隣実績で検討"
        if x["status"].startswith("空DL"):
            return "市町村単位では出力されない（レセプト系）", "二次医療圏単位で再取得／代替"
        if x["status"] == "対象自治体なし":
            return "地域軸に市町村が現れない（レセプト系）", "二次医療圏単位で再取得／代替"
        return "読込エラー", "再ダウンロード"

    style_header_row(ws, 4, ["カテゴリ", "コード", "指標名", "形式", "ステータス", "推定原因", "対応方針"])
    ws.row_dimensions[4].height = 28
    r = 5
    for x in recs:
        if x["status"] == "データあり":
            continue
        c1, c2 = cause(x)
        for c, v in enumerate([x["cat"], x["code"], x["indicator"], x["kind"], x["status"], c1, c2], 1):
            cell = ws.cell(row=r, column=c, value=v)
            style_data_cell(cell, alt=(r % 2 == 0), center=(c in (2, 4, 5)))
        sc = ws.cell(row=r, column=5)
        sc.fill = PatternFill("solid", fgColor=STATUS_COLOR.get(x["status"], COLORS["gray"]))
        sc.font = Font(name=FONT, size=9, bold=True, color="FFFFFF")
        r += 1

    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:G{r-1}"
    ws.sheet_view.showGridLines = False
    return ws


def add_key_sheet(wb, keyrows):
    ws = wb.create_sheet("03_主要指標")
    set_col_widths(ws, [10, 9, 44, 34, 22, 16, 8, 30])
    ws.row_dimensions[1].height = 28
    ws.merge_cells("A1:H1")
    style_title(ws["A1"], "音威子府村の主要指標（見える化システムから抽出）")
    ws.merge_cells("A2:H2")
    ws["A2"] = "1行＝1時点の値。ピボットテーブルでそのまま集計・作図できる形式（tidy形式）。"
    ws["A2"].font = Font(name=FONT, size=9, color="595959")
    ws["A2"].alignment = Alignment(vertical="center", indent=1)
    style_header_row(ws, 4, ["分野", "コード", "指標名", "系列", "時点", "値", "単位", "備考"])
    ws.row_dimensions[4].height = 26
    r = 5
    for field, code, ind, sname, period, val, unit, kind, note in keyrows:
        for c, v in enumerate([field, code, ind, sname, period, val, unit, note], 1):
            cell = ws.cell(row=r, column=c, value=v)
            style_data_cell(cell, alt=(r % 2 == 0), center=(c in (1, 2, 5, 6, 7)))
        vc = ws.cell(row=r, column=6)
        if isinstance(val, (int, float)):
            vc.number_format = "#,##0.0" if unit == "％" else "#,##0"
        r += 1
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:H{r-1}"
    ws.sheet_view.showGridLines = False
    return ws


def add_peer_sheet(wb, peerrows):
    ws = wb.create_sheet("04_近隣比較")
    set_col_widths(ws, [10, 42, 12, 14] + [13] * len(PEERS))
    ws.row_dimensions[1].height = 28
    last = get_column_letter(4 + len(PEERS))
    ws.merge_cells(f"A1:{last}1")
    style_title(ws["A1"], "近隣市町村・北海道・全国との比較")
    ws.merge_cells(f"A2:{last}2")
    ws["A2"] = "比較群は上川管内の近隣自治体（中川町・美深町・名寄市）と北海道・全国。"
    ws["A2"].font = Font(name=FONT, size=9, color="595959")
    ws["A2"].alignment = Alignment(vertical="center", indent=1)
    style_header_row(ws, 4, ["分野", "指標", "単位", "時点"] + PEERS)
    ws.row_dimensions[4].height = 26
    r = 5
    for field, srcf, lab, d, unit, when in peerrows:
        ws.cell(row=r, column=1, value=field)
        ws.cell(row=r, column=2, value=lab)
        ws.cell(row=r, column=3, value=unit)
        ws.cell(row=r, column=4, value=when)
        for k, rg in enumerate(PEERS):
            v = d.get(rg)
            cell = ws.cell(row=r, column=5 + k, value=None if is_blank(v) else v)
            if isinstance(v, (int, float)):
                cell.number_format = "#,##0.0" if unit == "％" else "#,##0"
        for c in range(1, 5 + len(PEERS)):
            style_data_cell(ws.cell(row=r, column=c), alt=(r % 2 == 0), center=(c >= 3))
        tc = ws.cell(row=r, column=5)
        tc.font = Font(name=FONT, size=10, bold=True, color="1F3864")
        tc.fill = PatternFill("solid", fgColor=COLORS["band"])
        r += 1
    ws.freeze_panes = "E5"
    ws.sheet_view.showGridLines = False
    return ws


WBS_MAP = [
    ("2.1.1", "人口・世帯の動向分析", "2.人口（A1〜A4）／3.世帯", "12＋14", "利用可", "A1は2050年までの推計値を収録"),
    ("2.1.2", "地域資源・社会基盤の整理", "9.入所（利用）定員／15.医療機関（G6-b, G7）", "12＋46", "一部要補完",
     "施設・居住系は村内になく全欠測。医療は診療所数・医師数のみ取得可"),
    ("2.1.3", "北海道・全国との比較分析", "各カテゴリの地域別ファイル", "－", "利用可", "比較群は上川管内＋北海道＋全国"),
    ("2.2.2", "総合事業の実績分析", "11.介護予防・日常生活支援総合事業", "54", "要補完",
     "最新値が令和2年度。村の実績で補完が必要"),
    ("2.2.3", "包括的支援事業・任意事業の分析", "10.地域包括支援センター／12.包括的支援事業", "25", "利用可", ""),
    ("2.3.1", "認定者数・認定率の推移分析", "5.要介護（要支援）認定（B3〜B6）", "32", "利用可",
     "平成19年〜令和8年4月末の長期時系列あり"),
    ("2.3.2", "サービス種類別給付実績の分析", "7.受給者数・利用回数／8.介護給付費・単位数", "302", "利用可",
     "村内に無いサービスは欠測（＝利用実績なし）"),
    ("2.3.4", "介護保険財政の分析", "14.介護保険特別会計経理状況", "50", "一部要補完",
     "基金残高そのものは収録なし。村から入手が必要"),
    ("3.1", "人口推計", "2.人口 A1（2000〜2050）", "2", "利用可", "社人研準拠の推計値をベースラインに使用"),
    ("3.2", "要支援・要介護認定者数の推計", "5.要介護（要支援）認定 B4・B5", "－", "利用可",
     "調整済み認定率（B5・B6）で全国比較が可能"),
    ("4.1", "サービス見込量の推計", "7.受給者数／8.給付費／9.定員", "314", "利用可",
     "圏域外利用の分解が必要"),
    ("4.3", "介護保険料の算定", "6.介護保険料 C1／14.特別会計", "52", "利用可",
     "必要保険料月額と実際の保険料の乖離が把握できる"),
    ("5.3.5", "第5章 施策の展開（在宅医療・介護連携）", "17.在宅医療・介護連携推進事業（L16〜L28）", "76", "要対応",
     "レセプト系90ファイルが空振り。代替データの手当てが必要"),
    ("5.3.5", "第5章 施策の展開（認知症施策）", "13.認知症対策", "14", "要対応",
     "全て欠測。村の実績で補完"),
]


def add_wbs_map(wb):
    ws = wb.create_sheet("05_WBS対応")
    set_col_widths(ws, [10, 40, 44, 12, 14, 48])
    ws.row_dimensions[1].height = 28
    ws.merge_cells("A1:F1")
    style_title(ws["A1"], "WBSの作業と見える化システムデータの対応")
    ws.merge_cells("A2:F2")
    ws["A2"] = "『WBS_音威子府村_高齢者福祉計画・第10期介護保険事業計画.xlsx』の各作業で、どのカテゴリを使うかの対応表。"
    ws["A2"].font = Font(name=FONT, size=9, color="595959")
    ws["A2"].alignment = Alignment(vertical="center", indent=1)
    style_header_row(ws, 4, ["WBS No.", "作業項目", "使用カテゴリ", "ファイル数", "判定", "留意点"])
    ws.row_dimensions[4].height = 26
    jc = {"利用可": COLORS["ok"], "一部要補完": COLORS["subhead"], "要補完": COLORS["warn"], "要対応": COLORS["ng"]}
    r = 5
    for wbs, task, cats, n, judge, note in WBS_MAP:
        for c, v in enumerate([wbs, task, cats, n, judge, note], 1):
            cell = ws.cell(row=r, column=c, value=v)
            style_data_cell(cell, alt=(r % 2 == 0), center=(c in (1, 4, 5)))
        jcell = ws.cell(row=r, column=5)
        jcell.fill = PatternFill("solid", fgColor=jc[judge])
        jcell.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
        ws.row_dimensions[r].height = 30
        r += 1
    ws.freeze_panes = "A5"
    ws.sheet_view.showGridLines = False
    return ws


# ============================================================
# Markdown 版
# ============================================================
def build_markdown(recs, keyrows, peerrows):
    import collections
    st = collections.Counter(r["status"] for r in recs)
    L = ["# 見える化システム ダウンロードデータ 棚卸し（音威子府村）", ""]
    L.append(f"- 対象：**{len(recs)}ファイル**（カテゴリ1〜17）")
    dates = sorted({r.get("out_date", "") for r in recs if r.get("out_date")})
    L.append(f"- システム出力日：{'／'.join(dates)}")
    uniq = len({(r["code"], r["indicator"], r["kind"]) for r in recs})
    L.append(f"- ユニーク指標×形式：{uniq}種（重複配置 {len(recs)-uniq}件）")
    L.append("")
    L.append("## 1. 取得状況")
    L.append("")
    L.append("| ステータス | 件数 | 割合 | 意味 |")
    L.append("|---|---:|---:|---|")
    meaning = {
        "データあり": "音威子府村の値が入っている。そのまま利用可",
        "全欠測（‐表記）": "地域軸に村はあるが値が全て『-』",
        "空DL（データ行なし）": "地域軸はあるがデータ行が無い（空振りDL）",
        "対象自治体なし": "地域軸に村が現れない",
    }
    for k, v in st.most_common():
        L.append(f"| {k} | {v} | {v/len(recs)*100:.1f}% | {meaning.get(k,'')} |")
    L.append("")

    L.append("## 2. カテゴリ別")
    L.append("")
    L.append("| カテゴリ | 計 | データあり | 全欠測 | 空DL | 対象外 |")
    L.append("|---|---:|---:|---:|---:|---:|")
    cat = collections.defaultdict(collections.Counter)
    for x in recs:
        cat[x["cat"]][x["status"]] += 1
    for c in sorted(cat, key=lambda x: int(x.split(".")[0]) if x.split(".")[0].isdigit() else 999):
        d = cat[c]
        empt = d["空DL（データ行なし）"] + d["空DL（地域軸なし）"]
        L.append(f"| {c} | {sum(d.values())} | {d['データあり']} | {d['全欠測（‐表記）']} | {empt} | {d['対象自治体なし']} |")
    L.append("")

    L.append("## 3. 確認結果（所見）")
    L.append("")
    for field, lvl, body, act in FINDINGS:
        L.append(f"### [{lvl}] {field}")
        L.append("")
        L.append(body)
        L.append("")
        L.append(f"→ **{act}**")
        L.append("")

    L.append("## 4. 次のアクション")
    L.append("")
    L.append("| No. | 内容 | 担当 | 期限目安 | WBS |")
    L.append("|---|---|---|---|---|")
    for no, body, who, due, detail, wbs in NEXT_ACTIONS:
        L.append(f"| {no} | {body}<br>{detail} | {who} | {due} | {wbs} |")
    L.append("")

    L.append("## 5. 近隣比較（主要指標）")
    L.append("")
    L.append("| 指標 | 単位 | 時点 | " + " | ".join(PEERS) + " |")
    L.append("|---|---|---|" + "---:|" * len(PEERS))
    for field, srcf, lab, d, unit, when in peerrows:
        cells = []
        for rg in PEERS:
            v = d.get(rg)
            cells.append("-" if is_blank(v) else (f"{v:,.1f}" if isinstance(v, float) else f"{v:,}"))
        L.append(f"| {lab} | {unit} | {when} | " + " | ".join(cells) + " |")
    L.append("")
    L.append(f"詳細は Excel 版 `output/{FILENAME}` を参照（全6シート）。")
    L.append("")

    path = os.path.join(OUT_DIR, FILENAME.replace(".xlsx", ".md"))
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(L) + "\n")
    print(f"  ✓ 作成: {path}")
    return path


# ============================================================
# main
# ============================================================
def main():
    if not SRC or not os.path.isdir(SRC):
        print("使い方: python3 build_mieruka_inventory.py <見える化システムDLデータのルート>", file=sys.stderr)
        print("  （カテゴリ『1.地域分析』〜『17.在宅医療・介護連携推進事業』を含むディレクトリ）", file=sys.stderr)
        sys.exit(1)
    os.makedirs(OUT_DIR, exist_ok=True)

    print(f"【1】全ファイル監査： {SRC}")
    recs = run_audit(SRC)
    print(f"    {len(recs)} ファイルを監査")

    print("【2】主要指標を抽出")
    keyrows = extract_key(SRC)
    peerrows = extract_peer(SRC)
    print(f"    主要指標 {len(keyrows)} 行 ／ 近隣比較 {len(peerrows)} 行")

    print("【3】棚卸し表を作成")
    wb = Workbook()
    add_summary(wb, recs, SRC)
    add_file_list(wb, recs)
    add_gap_sheet(wb, recs)
    add_key_sheet(wb, keyrows)
    add_peer_sheet(wb, peerrows)
    add_wbs_map(wb)
    path = os.path.join(OUT_DIR, FILENAME)
    wb.save(path)
    print(f"  ✓ 作成: {path}")
    print(f"    シート数: {len(wb.sheetnames)}")
    build_markdown(recs, keyrows, peerrows)
    print("完了。出力先: " + OUT_DIR)


if __name__ == "__main__":
    main()
