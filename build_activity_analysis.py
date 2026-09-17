# -*- coding: utf-8 -*-
"""
営業活動ログ 分析ブック ジェネレータ

入力: 訪問活動ログ（SharePointリスト書き出し / Book1.xlsx シート「query」）
出力: 以下を集計した分析ブック
  1. 今週の活動量（週次推移）
  2. 個人ごとの案件数
  3. 再訪不要の件数と要因分析

集計はすべて「データ」シートに対する COUNTIFS で組んでいるため、
元データを差し替えれば全シートが再計算される。
"""

import os
import datetime as dt

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import activity_factor_rules

SRC = "/root/.claude/uploads/2d4ea02e-a396-5b8d-9782-a60ae676556d/e00ac175-Book1.xlsx"
SRC_SHEET = "query"
OUT_DIR = "/home/user/repository/output"
OUT_PATH = os.path.join(OUT_DIR, "営業活動ログ_分析_20260917.xlsx")

AS_OF = dt.date(2026, 9, 17)          # 分析基準日
THIS_WEEK = dt.date(2026, 9, 14)      # 今週（月曜起算）

# 移動圏（活動の止まり方がこの単位で揃うため、エリアを3圏に束ねる）
ZONE = {
    "北海道": "北海道",
    "青森県": "北東北", "秋田県": "北東北",
    "岩手県": "南東北・岩手", "宮城県": "南東北・岩手",
    "山形県": "南東北・岩手", "福島県": "南東北・岩手",
}
ZONES = ["北海道", "北東北", "南東北・岩手"]

# 北海道の担当者（営業会議ブック 20260904㉜!J162「※基本は長戸・村田・林・吉野」）
HOKKAIDO_MEMBERS = ["長戸", "村田", "林", "吉野"]

# 営業会議ブック 20260904㉜ からの転記値（比較・裏取り用）
BOOK = {
    "見積今期":   {"北海道": ("H172", 1),  "東北": ("H186", 2)},
    "見積来期":   {"北海道": ("I172", 44), "東北": ("I186", 57)},
    "累計営業":   {"北海道": ("G172", 288), "東北": ("G186", 275)},
    "9月入札見込": {"北海道": ("E63", 3),  "東北": ("E75", 3)},
    "9月入札確定": {"北海道": ("E64", 0),  "東北": ("E76", 1)},
    "公会計現状":  {"北海道": ("E240", 138), "東北": ("E241", 75)},
    "精緻化":     {"北海道": ("E259", 39), "東北": ("E260", 11)},
    "会計支援対象": {"北海道": ("D270", 29), "東北": ("D271", 24)},
    "決算報告書":  {"北海道": ("I270", 10), "東北": ("I271", 6)},
    "消費税計算":  {"北海道": ("G270", 4),  "東北": ("G271", 8)},
    "自治体数":   {"北海道": ("D52", 179), "東北": ("D53", 227)},
}

FONT = "游ゴシック"
COLORS = {
    "header": "1F3864", "subhead": "2E75B6", "band": "DDEBF7",
    "alt": "F7FAFC", "total": "FFF2CC", "warn": "FCE4E4", "good": "E2EFDA",
    "note": "FFF9E6",
}
THIN = Side(border_style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
NUM = '#,##0;(#,##0);-'
PCT1 = '0.0"%"'
DEC1 = '0.0'

DATA_SHEET = "データ"
# データシートの列配置（1始まり）
COL = {
    "入力者": 1, "訪問日": 2, "週ラベル": 3, "エリア": 4, "団体名": 5, "テーマ": 6,
    "結果": 7, "次回アクション": 8, "前回策定": 9, "改定時期": 10,
    "アプローチ時期": 11, "状態": 12, "契約予定": 13, "要因": 14, "圏": 15, "備考": 16,
}
HDR_ROW = 1
BLANK = "(未記入)"


# ============================================================
# データ読み込み
# ============================================================
def load() -> pd.DataFrame:
    df = pd.read_excel(SRC, sheet_name=SRC_SHEET)
    df["訪問日"] = pd.to_datetime(df["訪問日"], errors="coerce")
    monday = df["訪問日"] - pd.to_timedelta(df["訪問日"].dt.weekday, unit="D")
    df["週ラベル"] = monday.dt.strftime("%Y-%m-%d").fillna(BLANK)
    df["圏"] = df["エリア"].map(ZONE).fillna(BLANK)
    df = activity_factor_rules.add_factors(df)
    for c in ("入力者", "エリア", "団体名", "テーマ", "結果", "次回アクション",
              "前回策定", "改定時期", "アプローチ時期", "状態", "契約予定", "圏"):
        df[c] = df[c].fillna(BLANK).astype(str).str.strip().replace("", BLANK)
    df["要因"] = df["要因"].replace("", "")
    df["備考"] = df["備考"].fillna("").astype(str).str.replace("\n", "／", regex=False)
    return df


# ============================================================
# 書式ヘルパ
# ============================================================
def title(ws, cell, text, size=14):
    ws[cell] = text
    ws[cell].font = Font(name=FONT, size=size, bold=True, color=COLORS["header"])


def note(ws, cell, text, size=9, color="595959"):
    ws[cell] = text
    ws[cell].font = Font(name=FONT, size=size, color=color)


def head(ws, row, col, text, fill="subhead", width=None):
    c = ws.cell(row=row, column=col, value=text)
    c.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=COLORS[fill])
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = BORDER
    if width:
        ws.column_dimensions[get_column_letter(col)].width = width
    return c


def body(ws, row, col, value, fmt=None, bold=False, fill=None, align="right", src=False):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name=FONT, size=10, bold=bold, color="0000FF" if src else "000000")
    c.border = BORDER
    c.alignment = Alignment(horizontal=align, vertical="center")
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = PatternFill("solid", fgColor=COLORS[fill])
    return c


def label(ws, row, col, text, bold=False, fill=None):
    return body(ws, row, col, text, bold=bold, fill=fill, align="left")


def page_setup(ws, freeze=None, one_page=False):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1 if one_page else 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    if freeze:
        ws.freeze_panes = freeze


# ============================================================
# COUNTIFS 組み立て
# ============================================================
class Ref:
    """データシートの列を絶対参照で返す。"""

    def __init__(self, nrows):
        self.last = HDR_ROW + nrows

    def rng(self, name):
        c = get_column_letter(COL[name])
        return f"'{DATA_SHEET}'!${c}${HDR_ROW + 1}:${c}${self.last}"

    def count_all(self):
        """全データ行数。COUNTIFS は条件ゼロだと #VALUE! になるため COUNTA を使う。"""
        return f'=COUNTA({self.rng("入力者")})'

    def countifs(self, **cond):
        """cond の値は文字列（等値条件）。セル参照は '=' 付きで渡す。"""
        parts = []
        for k, v in cond.items():
            crit = v[1:] if isinstance(v, str) and v.startswith("=") else f'"{v}"'
            parts.append(f"{self.rng(k)},{crit}")
        return "=COUNTIFS(" + ",".join(parts) + ")"


def pct(num_cell, den_cell):
    return f"=IFERROR({num_cell}/{den_cell}*100,0)"


# ============================================================
# Sheet: データ（集計の参照元）
# ============================================================
def build_data(wb, df):
    ws = wb.create_sheet(DATA_SHEET)
    order = sorted(COL, key=lambda k: COL[k])
    for name in order:
        head(ws, HDR_ROW, COL[name], name,
             width=11 if name not in ("備考", "団体名", "要因", "圏") else
                   (60 if name == "備考" else (14 if name == "圏" else 24)))
    for i, rec in enumerate(df.itertuples(index=False), start=HDR_ROW + 1):
        d = rec._asdict()
        for name in order:
            v = d[name]
            if name == "訪問日":
                v = None if pd.isna(v) else v.date()
                c = body(ws, i, COL[name], v, "yyyy/mm/dd", src=True, align="center")
            else:
                c = body(ws, i, COL[name], v, src=True,
                         align="left" if name in ("団体名", "備考", "要因") else "center")
            c.font = Font(name=FONT, size=9, color="0000FF")
    ws.auto_filter.ref = f"A{HDR_ROW}:{get_column_letter(len(COL))}{HDR_ROW + len(df)}"
    ws.freeze_panes = "B2"
    return ws


# ============================================================
# Sheet: ① 今週の活動量（週次推移）
# ============================================================
def build_weekly(ws, df, ref):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 22
    for col in "CDEFGHIJKL":
        ws.column_dimensions[col].width = 12

    title(ws, "B2", "① 週次の活動量")
    note(ws, "B3", f"基準日：{AS_OF:%Y/%m/%d}（木）　週は月曜起算　"
                   f"今週＝{THIS_WEEK:%-m/%d}週は {AS_OF:%-m/%d} 時点で未了（最終入力 9/16）")

    weeks = [w for w in sorted(df["週ラベル"].unique()) if w != BLANK]
    weeks = weeks[-10:]

    r = 5
    for col, text in [(2, "週開始\n(月曜)"), (3, "活動件数"), (4, "団体数"), (5, "稼働者数"),
                      (6, "訪問日数"), (7, "1人1日\n件数"), (8, "再アプローチ"), (9, "見積提示"),
                      (10, "再訪不要"), (11, "再訪不要率"), (12, "前週比\n件数")]:
        head(ws, r, col, text, fill="header" if col == 3 else "subhead")
    ws.row_dimensions[r].height = 30

    first = r + 1
    for i, wk in enumerate(weeks):
        rr = first + i
        sub = df[df["週ラベル"] == wk]
        is_tw = wk == f"{THIS_WEEK:%Y-%m-%d}"
        fill = "band" if is_tw else ("alt" if i % 2 else None)
        label(ws, rr, 2, wk + ("　←今週" if is_tw else ""), bold=is_tw, fill=fill)
        body(ws, rr, 3, ref.countifs(週ラベル=wk), NUM, bold=is_tw, fill=fill)
        # 団体数・稼働者数・訪問日数は一意カウントのため Python 側で算出（注記参照）
        body(ws, rr, 4, sub["団体名"].nunique(), NUM, fill=fill, src=True)
        body(ws, rr, 5, sub["入力者"].nunique(), NUM, fill=fill, src=True)
        body(ws, rr, 6, sub["訪問日"].dt.date.nunique(), NUM, fill=fill, src=True)
        body(ws, rr, 7, f"=IFERROR(C{rr}/F{rr}/E{rr},0)", DEC1, fill=fill)
        body(ws, rr, 8, ref.countifs(週ラベル=wk, 次回アクション="再アプローチ"), NUM, fill=fill)
        body(ws, rr, 9, ref.countifs(週ラベル=wk, 次回アクション="見積提示"), NUM, fill=fill)
        body(ws, rr, 10, ref.countifs(週ラベル=wk, 次回アクション="再訪不要"), NUM, fill=fill)
        body(ws, rr, 11, pct(f"J{rr}", f"C{rr}"), PCT1, fill=fill)
        body(ws, rr, 12, "-" if i == 0 else f"=C{rr}-C{rr-1}", NUM, fill=fill)

    last = first + len(weeks) - 1
    rr = last + 1
    label(ws, rr, 2, "掲載週 計", bold=True, fill="total")
    for col in (3, 8, 9, 10):
        L = get_column_letter(col)
        body(ws, rr, col, f"=SUM({L}{first}:{L}{last})", NUM, bold=True, fill="total")
    body(ws, rr, 11, pct(f"J{rr}", f"C{rr}"), PCT1, bold=True, fill="total")

    # 今週の内訳
    r = rr + 3
    title(ws, f"B{r}", f"今週（{THIS_WEEK:%-m/%d}〜）の内訳", size=12)
    tw = f"{THIS_WEEK:%Y-%m-%d}"
    r += 1
    blocks = [
        ("入力者別", "入力者", ["相澤", "高橋", "角張"]),
        ("エリア別", "エリア", ["岩手県", "福島県"]),
        ("テーマ別", "テーマ", ["公会計", "その他計画", "福祉計画", "公共施設マネジメント", "経営戦略"]),
        ("状態別", "状態", ["セカンドアプローチ", "初回訪問", "見積提示", "見積＋仕様書"]),
    ]
    for bname, field, keys in blocks:
        head(ws, r, 2, bname)
        head(ws, r, 3, "件数")
        for j, k in enumerate(keys):
            label(ws, r + 1 + j, 2, k, fill="alt" if j % 2 else None)
            body(ws, r + 1 + j, 3, ref.countifs(週ラベル=tw, **{field: k}), NUM,
                 fill="alt" if j % 2 else None)
        tr = r + 1 + len(keys)
        label(ws, tr, 2, "計", bold=True, fill="total")
        body(ws, tr, 3, f"=SUM(C{r+1}:C{tr-1})", NUM, bold=True, fill="total")
        r = tr + 2

    note(ws, f"B{r}", "※ 活動件数・各アクション件数は「データ」シートへの COUNTIFS。団体数／稼働者数／訪問日数は一意カウントのため生成時点の実数（青字）。")
    note(ws, f"B{r+1}", "※ 1人1日件数＝活動件数÷訪問日数÷稼働者数。稼働者数の増減を除いた1人あたりの活動密度を見るための指標。")
    note(ws, f"B{r+2}", f"※ 今週（{THIS_WEEK:%-m/%d}週）は 9/14・9/15・9/16 の3日分のみ。週の途中であり、他週との単純比較はできない。")


# ============================================================
# Sheet: ② 個人ごとの案件数
# ============================================================
def build_person(ws, df, ref):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 10
    for col in "CDEFGHIJKLMNO":
        ws.column_dimensions[col].width = 11

    title(ws, "B2", "② 個人ごとの案件数")
    note(ws, "B3", "活動件数＝ログ1行を1件として計上（同一団体でもテーマが異なれば別件）")

    people = (df[df["入力者"] != BLANK].groupby("入力者").size()
              .sort_values(ascending=False).index.tolist())
    weeks = [w for w in sorted(df["週ラベル"].unique()) if w != BLANK][-4:]

    r = 5
    cols = [(2, "入力者"), (3, "活動件数"), (4, "団体数"), (5, "訪問日数"), (6, "1日あたり\n件数"),
            (7, "再アプローチ"), (8, "見積提示"), (9, "同行依頼、\n再訪"), (10, "再訪不要"),
            (11, "再訪不要率")]
    for j, wk in enumerate(weeks):
        cols.append((12 + j, wk[5:].replace("-", "/") + "週"))
    for col, text in cols:
        head(ws, r, col, text, fill="header" if col == 3 else "subhead")
    ws.row_dimensions[r].height = 30

    first = r + 1
    for i, p in enumerate(people):
        rr = first + i
        sub = df[df["入力者"] == p]
        fill = "alt" if i % 2 else None
        label(ws, rr, 2, p, bold=True, fill=fill)
        body(ws, rr, 3, ref.countifs(入力者=p), NUM, bold=True, fill=fill)
        body(ws, rr, 4, sub["団体名"].nunique(), NUM, fill=fill, src=True)
        body(ws, rr, 5, sub["訪問日"].dt.date.nunique(), NUM, fill=fill, src=True)
        body(ws, rr, 6, f"=IFERROR(C{rr}/E{rr},0)", DEC1, fill=fill)
        body(ws, rr, 7, ref.countifs(入力者=p, 次回アクション="再アプローチ"), NUM, fill=fill)
        body(ws, rr, 8, ref.countifs(入力者=p, 次回アクション="見積提示"), NUM, fill=fill)
        body(ws, rr, 9, ref.countifs(入力者=p, 次回アクション="同行依頼、再訪"), NUM, fill=fill)
        body(ws, rr, 10, ref.countifs(入力者=p, 次回アクション="再訪不要"), NUM, fill=fill)
        body(ws, rr, 11, f"=IFERROR(J{rr}/SUM(G{rr}:J{rr})*100,0)", PCT1, fill=fill)
        for j, wk in enumerate(weeks):
            body(ws, rr, 12 + j, ref.countifs(入力者=p, 週ラベル=wk), NUM, fill=fill)

    last = first + len(people) - 1
    rr = last + 1
    label(ws, rr, 2, "計", bold=True, fill="total")
    for col in [3] + list(range(7, 11)) + [12 + j for j in range(len(weeks))]:
        L = get_column_letter(col)
        body(ws, rr, col, f"=SUM({L}{first}:{L}{last})", NUM, bold=True, fill="total")
    body(ws, rr, 4, None, fill="total")
    body(ws, rr, 5, None, fill="total")
    body(ws, rr, 6, None, fill="total")
    body(ws, rr, 11, f"=IFERROR(J{rr}/SUM(G{rr}:J{rr})*100,0)", PCT1, bold=True, fill="total")

    r = rr + 2
    note(ws, f"B{r}", "※ 再訪不要率＝再訪不要÷（次回アクション記入分の合計）。アクション未記入の行は母数から除外。")
    note(ws, f"B{r+1}", "※ 団体数・訪問日数は一意カウントのため生成時点の実数（青字）。活動件数・アクション別件数は COUNTIFS。")
    note(ws, f"B{r+2}", "※ 入力者が未記入の行3件は本表の対象外。")
    return first, last


# ============================================================
# Sheet: ③ 再訪不要の要因分析
# ============================================================
# クロス集計する属性。選択肢はデータから全件導出するため、計は必ず総件数に一致する。
CROSS_FIELDS = ["結果", "状態", "テーマ", "エリア", "アプローチ時期", "前回策定"]
# 段階を表す属性は出現頻度ではなく業務上の順序で並べる
CROSS_ORDER = {
    "結果": ["自力", "他社随契", "その他", "見積提示", "不在"],
    "状態": ["初回訪問", "セカンドアプローチ", "見積提示", "見積＋仕様書"],
    "アプローチ時期": ["今年度", "来年度", "再来年度"],
    "前回策定": ["自前", "委託"],
}


def cross_keys(df, field):
    """当該列の全カテゴリを返す。既定順があればそれを先に、残りは件数降順。(未記入)は末尾。"""
    counts = df[field].value_counts()
    fixed = [k for k in CROSS_ORDER.get(field, []) if k in counts.index]
    rest = [k for k in counts.index if k not in fixed and k != BLANK]
    tail = [BLANK] if BLANK in counts.index else []
    return fixed + rest + tail


def build_factor(ws, df, ref):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 34
    for col in "CDEFGH":
        ws.column_dimensions[col].width = 13

    title(ws, "B2", "③ 再訪不要の件数と要因分析")
    note(ws, "B3", "要因は備考テキストのキーワード判定＋結果区分で自動分類（ロジックは activity_factor_rules.py）")

    n_total = int((df["次回アクション"] == "再訪不要").sum())
    act_total = int((df["次回アクション"] != BLANK).sum())

    r = 5
    title(ws, f"B{r}", "全体像", size=12)
    r += 1
    for col, text in [(2, "指標"), (3, "件数"), (4, "構成比")]:
        head(ws, col=col, row=r, text=text)
    kpis = [
        ("活動ログ総件数", ref.count_all(), None),
        ("うち次回アクション記入", f'=COUNTA({ref.rng("次回アクション")})-COUNTIFS({ref.rng("次回アクション")},"{BLANK}")', None),
        ("再訪不要", ref.countifs(次回アクション="再訪不要"), True),
        ("再アプローチ", ref.countifs(次回アクション="再アプローチ"), True),
        ("見積提示", ref.countifs(次回アクション="見積提示"), True),
        ("同行依頼、再訪", ref.countifs(次回アクション="同行依頼、再訪"), True),
    ]
    kfirst = r + 1
    for i, (name, formula, ratio) in enumerate(kpis):
        rr = kfirst + i
        emph = name == "再訪不要"
        fill = "warn" if emph else ("alt" if i % 2 else None)
        label(ws, rr, 2, name, bold=emph, fill=fill)
        body(ws, rr, 3, formula, NUM, bold=emph, fill=fill)
        body(ws, rr, 4, pct(f"C{rr}", f"$C${kfirst+1}") if ratio else None, PCT1,
             bold=emph, fill=fill)
    note(ws, f"E{kfirst+2}", "← アクション記入分に対する比率")

    # 要因別
    r = kfirst + len(kpis) + 2
    title(ws, f"B{r}", "要因別の内訳", size=12)
    r += 1
    for col, text in [(2, "要因"), (3, "件数"), (4, "構成比"), (5, "大分類")]:
        head(ws, col=col, row=r, text=text)
    factors = (df[df["次回アクション"] == "再訪不要"]["要因"]
               .value_counts().index.tolist())
    ffirst = r + 1
    for i, f in enumerate(factors):
        rr = ffirst + i
        fill = "alt" if i % 2 else None
        label(ws, rr, 2, f, fill=fill)
        body(ws, rr, 3, ref.countifs(次回アクション="再訪不要", 要因=f), NUM, fill=fill)
        body(ws, rr, 4, pct(f"C{rr}", f"$C${ffirst+len(factors)}"), PCT1, fill=fill)
        label(ws, rr, 5, f.split("：")[0], fill=fill)
    rr = ffirst + len(factors)
    label(ws, rr, 2, "計", bold=True, fill="total")
    body(ws, rr, 3, f"=SUM(C{ffirst}:C{rr-1})", NUM, bold=True, fill="total")
    body(ws, rr, 4, pct(f"C{rr}", f"C{rr}"), PCT1, bold=True, fill="total")
    body(ws, rr, 5, None, fill="total")

    # 属性別の再訪不要率
    r = rr + 3
    title(ws, f"B{r}", "属性別の再訪不要率（どこで案件が落ちているか）", size=12)
    r += 1
    for field in CROSS_FIELDS:
        keys = cross_keys(df, field)
        head(ws, r, 2, field)
        head(ws, r, 3, "再訪不要")
        head(ws, r, 4, "全件数")
        head(ws, r, 5, "再訪不要率")
        cfirst = r + 1
        for i, k in enumerate(keys):
            rr = cfirst + i
            fill = "alt" if i % 2 else None
            label(ws, rr, 2, k, fill=fill)
            body(ws, rr, 3, ref.countifs(次回アクション="再訪不要", **{field: k}), NUM, fill=fill)
            body(ws, rr, 4, ref.countifs(**{field: k}), NUM, fill=fill)
            body(ws, rr, 5, pct(f"C{rr}", f"D{rr}"), PCT1, fill=fill)
        rr = cfirst + len(keys)
        label(ws, rr, 2, "計", bold=True, fill="total")
        body(ws, rr, 3, f"=SUM(C{cfirst}:C{rr-1})", NUM, bold=True, fill="total")
        body(ws, rr, 4, f"=SUM(D{cfirst}:D{rr-1})", NUM, bold=True, fill="total")
        body(ws, rr, 5, pct(f"C{rr}", f"D{rr}"), PCT1, bold=True, fill="total")
        r = rr + 2

    note(ws, f"B{r}", "※ 「全件数」は当該属性の全活動件数。再訪不要率が高い属性は、初回接触で案件化しないまま終わっている領域。")
    note(ws, f"B{r+1}", f"※ 再訪不要 {n_total}件／次回アクション記入 {act_total}件。"
                        f"アクション未記入{len(df) - act_total}件は母数から除外している。")
    return ffirst


# ============================================================
# Sheet: 明細（再訪不要の全件・要因の検証用）
# ============================================================
def build_detail(ws, df):
    ws.sheet_view.showGridLines = False
    cols = [("入力者", 9), ("訪問日", 11), ("エリア", 9), ("団体名", 13), ("テーマ", 17),
            ("結果", 10), ("要因", 30), ("備考", 80)]
    title(ws, "B2", "④ 再訪不要 明細（要因判定の検証用）")
    note(ws, "B3", "要因の自動判定が実態と合わない行は、この明細で確認して activity_factor_rules.py のルールを調整する。")

    r = 5
    for j, (name, w) in enumerate(cols):
        head(ws, r, 2 + j, name, width=w)
    n = df[df["次回アクション"] == "再訪不要"]
    for i, rec in enumerate(n.itertuples(index=False)):
        rr = r + 1 + i
        d = rec._asdict()
        fill = "alt" if i % 2 else None
        for j, (name, _) in enumerate(cols):
            v = d[name]
            if name == "訪問日":
                v = None if pd.isna(v) else v.date()
                c = body(ws, rr, 2 + j, v, "yyyy/mm/dd", fill=fill, align="center")
            else:
                c = body(ws, rr, 2 + j, v, fill=fill,
                         align="left" if name in ("団体名", "要因", "備考") else "center")
            c.font = Font(name=FONT, size=9)
    ws.auto_filter.ref = f"B{r}:{get_column_letter(1 + len(cols))}{r + len(n)}"
    ws.freeze_panes = f"B{r+1}"


# ============================================================
# Sheet: サマリー
# ============================================================
def build_summary(ws, df, ref, sheets):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 26
    for col in "CDEFGH":
        ws.column_dimensions[col].width = 14

    title(ws, "B2", "営業活動ログ 分析サマリー")
    note(ws, "B3", f"対象：訪問活動ログ {len(df)}件（{df['訪問日'].min():%Y/%m/%d}〜{df['訪問日'].max():%Y/%m/%d}）　"
                   f"基準日 {AS_OF:%Y/%m/%d}")

    wkly, per, fac = sheets["週次"], sheets["個人"], sheets["要因"]
    tw = f"{THIS_WEEK:%Y-%m-%d}"
    lw = f"{THIS_WEEK - dt.timedelta(days=7):%Y-%m-%d}"

    # ① 活動量
    r = 5
    title(ws, f"B{r}", "① 今週の活動量", size=12)
    r += 1
    for col, text in [(2, "区分"), (3, "活動件数"), (4, "団体数"), (5, "稼働者数"), (6, "再アプローチ"),
                      (7, "見積提示"), (8, "再訪不要")]:
        head(ws, r, col, text, fill="header" if col == 3 else "subhead")
    rows = [
        (f"今週（{THIS_WEEK:%-m/%d}〜、3日分）", tw, True),
        (f"先週（{THIS_WEEK - dt.timedelta(days=7):%-m/%d}〜）", lw, False),
        (f"前々週（{THIS_WEEK - dt.timedelta(days=14):%-m/%d}〜）",
         f"{THIS_WEEK - dt.timedelta(days=14):%Y-%m-%d}", False),
    ]
    afirst = r + 1
    for i, (name, wk, emph) in enumerate(rows):
        rr = afirst + i
        sub = df[df["週ラベル"] == wk]
        fill = "band" if emph else ("alt" if i % 2 else None)
        label(ws, rr, 2, name, bold=emph, fill=fill)
        body(ws, rr, 3, ref.countifs(週ラベル=wk), NUM, bold=emph, fill=fill)
        body(ws, rr, 4, sub["団体名"].nunique(), NUM, fill=fill, src=True)
        body(ws, rr, 5, sub["入力者"].nunique(), NUM, fill=fill, src=True)
        body(ws, rr, 6, ref.countifs(週ラベル=wk, 次回アクション="再アプローチ"), NUM, fill=fill)
        body(ws, rr, 7, ref.countifs(週ラベル=wk, 次回アクション="見積提示"), NUM, fill=fill)
        body(ws, rr, 8, ref.countifs(週ラベル=wk, 次回アクション="再訪不要"), NUM, fill=fill)
    note(ws, f"B{afirst+3}", "→ 8/24週109件をピークに 85→35→17件と減少。稼働者数も9→7→5→3名と連動しており、"
                             "1人1日あたりの件数（約2件）はほぼ横ばい。活動量の減少は「人が動いていない」ことが主因。")

    # ② 個人
    r = afirst + 5
    title(ws, f"B{r}", "② 個人ごとの案件数（上位）", size=12)
    r += 1
    for col, text in [(2, "入力者"), (3, "活動件数"), (4, "1日あたり"), (5, "再訪不要率"), (6, "直近2週")]:
        head(ws, r, col, text, fill="header" if col == 3 else "subhead")
    pfirst, plast = per["first"], per["last"]
    top = (df[df["入力者"] != BLANK].groupby("入力者").size()
           .sort_values(ascending=False).index.tolist())
    sfirst = r + 1
    for i, p in enumerate(top):
        rr = sfirst + i
        prow = pfirst + i
        fill = "alt" if i % 2 else None
        label(ws, rr, 2, p, fill=fill)
        body(ws, rr, 3, f"='{per['name']}'!C{prow}", NUM, fill=fill)
        body(ws, rr, 4, f"='{per['name']}'!F{prow}", DEC1, fill=fill)
        body(ws, rr, 5, f"='{per['name']}'!K{prow}", PCT1, fill=fill)
        body(ws, rr, 6, f"=COUNTIFS({ref.rng('入力者')},\"{p}\",{ref.rng('週ラベル')},\"{tw}\")"
                        f"+COUNTIFS({ref.rng('入力者')},\"{p}\",{ref.rng('週ラベル')},\"{lw}\")",
             NUM, fill=fill)
    rr = sfirst + len(top)
    label(ws, rr, 2, "計", bold=True, fill="total")
    body(ws, rr, 3, f"=SUM(C{sfirst}:C{rr-1})", NUM, bold=True, fill="total")
    body(ws, rr, 4, None, fill="total")
    body(ws, rr, 5, None, fill="total")
    body(ws, rr, 6, f"=SUM(F{sfirst}:F{rr-1})", NUM, bold=True, fill="total")
    note(ws, f"B{rr+1}", "→ 活動量は林・相澤・村田の3名で272件（全体の6割）。ただし林は直近2週ゼロ。"
                         "相澤は1日7.1件・再訪不要率7.7%で量と質が両立。村田は再訪不要率52.5%で突出。")

    # ③ 要因
    r = rr + 3
    title(ws, f"B{r}", "③ 再訪不要の要因（大分類）", size=12)
    r += 1
    for col, text in [(2, "大分類"), (3, "件数"), (4, "構成比"), (5, "主な内容")]:
        head(ws, r, col, text, fill="header" if col == 3 else "subhead")
    n = df[df["次回アクション"] == "再訪不要"]
    grp = n["要因"].str.split("：").str[0].value_counts()
    desc = {
        "他社受託": "システム保守に抱き合わせ22件／随契・委託継続14件／入札実績要件1件",
        "自前作成方針（庁内対応）": "長年自前・担当者が内製、前回数値の置き換えで足りる",
        "策定済み・改定時期が先": "R8改定直後、計画期間がR10〜R11まで。時期を待てば再浮上",
        "必要性・意欲が低い": "罰則なし・メリット感じない・長の方針で実施しない",
        "予算がつかない・否決": "例年予算つかず、予算否決により自前化",
        "他計画へ一体化・包含され単独案件消滅": "こども計画・健康増進計画・総合計画に内包",
        "庁内事情・外部要因による遅延": "庁内調整負担、地番整理待ち",
        "要因不明（備考・結果とも未記入）": "備考・結果とも空欄で判定不能",
    }
    gfirst = r + 1
    factors_all = n["要因"].value_counts().index.tolist()
    ffirst = fac["ffirst"]
    for i, (g, _) in enumerate(grp.items()):
        rr2 = gfirst + i
        fill = "warn" if g in ("他社受託", "自前作成方針（庁内対応）") else ("alt" if i % 2 else None)
        label(ws, rr2, 2, g, bold=True, fill=fill)
        terms = [f"'{fac['name']}'!C{ffirst + factors_all.index(f)}"
                 for f in factors_all if f.split("：")[0] == g]
        body(ws, rr2, 3, "=" + "+".join(terms), NUM, bold=True, fill=fill)
        body(ws, rr2, 4, pct(f"C{rr2}", f"$C${gfirst + len(grp)}"), PCT1, fill=fill)
        label(ws, rr2, 5, desc.get(g, ""), fill=fill)
    rr2 = gfirst + len(grp)
    label(ws, rr2, 2, "計", bold=True, fill="total")
    body(ws, rr2, 3, f"=SUM(C{gfirst}:C{rr2-1})", NUM, bold=True, fill="total")
    body(ws, rr2, 4, pct(f"C{rr2}", f"C{rr2}"), PCT1, bold=True, fill="total")
    label(ws, rr2, 5, None, fill="total")
    note(ws, f"B{rr2+2}", "→ 「他社受託」と「自前作成方針」で各36%、合計7割超。"
                          "うちシステム保守への抱き合わせ22件は公会計・会計支援・福祉計画に集中しており、"
                          "単発の訪問では崩せない構造要因。")
    note(ws, f"B{rr2+3}", "→ 一方「策定済み・改定時期が先」10件は案件が消えたのではなく時期待ち。"
                          "改定時期を管理して再アプローチすれば戻せる母数。")


# ============================================================
# Sheet: ⑤ 北海道の活動停止 要因分析
# ============================================================
def build_hokkaido(ws, df, ref):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 30
    for col in "CDEFG":
        ws.column_dimensions[col].width = 14
    ws.column_dimensions["H"].width = 62

    title(ws, "B2", "⑤ 北海道の活動が止まっている要因")
    note(ws, "B3", "活動ログで特定できるのは「いつ・何が」まで。「なぜ」は営業会議ブックの構造要因と突き合わせた推論を含む。支持度欄を参照。")

    # --- 1. 止まり方は圏単位で揃っている ---
    r = 5
    title(ws, f"B{r}", "1. 止まったのは北海道だけではない（移動圏単位で揃っている）", size=12)
    r += 1
    weeks = [w for w in sorted(df["週ラベル"].unique()) if w != BLANK][-5:]
    head(ws, r, 2, "移動圏")
    for j, wk in enumerate(weeks):
        head(ws, r, 3 + j, wk[5:].replace("-", "/") + "週",
             fill="header" if wk == f"{THIS_WEEK:%Y-%m-%d}" else "subhead")
    head(ws, r, 3 + len(weeks), "最終活動日")
    ws.row_dimensions[r].height = 28
    zfirst = r + 1
    for i, z in enumerate(ZONES):
        rr = zfirst + i
        sub = df[df["圏"] == z]
        fill = "warn" if z != "南東北・岩手" else "good"
        label(ws, rr, 2, z, bold=True, fill=fill)
        for j, wk in enumerate(weeks):
            body(ws, rr, 3 + j, ref.countifs(圏=z, 週ラベル=wk), NUM, fill=fill)
        body(ws, rr, 3 + len(weeks), sub["訪問日"].max().date(), "yyyy/mm/dd",
             fill=fill, align="center", src=True)
    r = zfirst + len(ZONES) + 1
    note(ws, f"B{r}", "→ 北東北（青森・秋田）は9/07週から、北海道は9/14週からゼロ。南東北・岩手の3名（相澤・角張・高橋）のみ継続。"
                      "個人の問題ではなく圏単位で止まっている。")

    # --- 2. 担当者ごとの最終活動 ---
    r += 2
    title(ws, f"B{r}", "2. 北海道担当4名の最終活動（営業会議ブック J162「※基本は長戸・村田・林・吉野」）", size=12)
    r += 1
    for col, text in [(2, "担当者"), (3, "活動件数"), (4, "最終活動日"), (5, "経過日数"),
                      (6, "最終日の件数"), (7, "最終日のアクション")]:
        head(ws, r, col, text)
    mfirst = r + 1
    for i, p in enumerate(HOKKAIDO_MEMBERS):
        rr = mfirst + i
        sub = df[(df["入力者"] == p) & (df["圏"] == "北海道")]
        last_day = sub["訪問日"].max()
        lastrows = sub[sub["訪問日"] == last_day]
        fill = "alt" if i % 2 else None
        label(ws, rr, 2, p, bold=True, fill=fill)
        body(ws, rr, 3, ref.countifs(入力者=p, 圏="北海道"), NUM, fill=fill)
        body(ws, rr, 4, last_day.date(), "yyyy/mm/dd", fill=fill, align="center", src=True)
        body(ws, rr, 5, (AS_OF - last_day.date()).days, NUM, fill=fill)
        body(ws, rr, 6, len(lastrows), NUM, fill=fill, src=True)
        label(ws, rr, 7, "／".join(f"{k}{v}件" for k, v in
                                  lastrows["次回アクション"].value_counts().items()), fill=fill)
    r = mfirst + len(HOKKAIDO_MEMBERS) + 1
    note(ws, f"B{r}", "→ 4名とも最終記録は「見積提示」または「再アプローチ」。投げ出しではなく、訪問から見積・提案フェーズへ移った形で記録が途絶えている。")

    # --- 3. 北海道 vs 東北 の指標比較 ---
    r += 2
    title(ws, f"B{r}", "3. 北海道と東北の違い（なぜ北海道から先に枯れたか）", size=12)
    r += 1
    for col, text in [(2, "指標"), (3, "北海道"), (4, "東北"), (5, "単位"), (6, "出所"), (7, ""), (8, "読み方")]:
        head(ws, r, col, text)
    TOH = [z for z in ZONES if z != "北海道"]
    PIPE_ACTIONS = ["再アプローチ", "同行依頼、再訪"]

    def zsum(zones, **c):
        """指定した圏の COUNTIFS を合算した式を返す。"""
        return "=" + "+".join(ref.countifs(圏=z, **c)[1:] for z in zones)

    def pipe(zones, **c):
        """残パイプライン（再アプローチ＋同行依頼、再訪）の合算式。"""
        terms = [ref.countifs(圏=z, 次回アクション=a, **c)[1:]
                 for z in zones for a in PIPE_ACTIONS]
        return "=" + "+".join(terms)

    tsum = lambda **c: zsum(TOH, **c)
    metrics = [
        ("活動件数", ref.countifs(圏="北海道"), tsum(), "件", "活動ログ",
         "母数はほぼ同等。活動量そのものの差ではない"),
        ("訪問済み団体数", None, None, "団体", "活動ログ",
         "北海道108／道内179自治体（60%）。未訪問は約71団体残っている"),
        ("再訪不要", ref.countifs(圏="北海道", 次回アクション="再訪不要"),
         tsum(次回アクション="再訪不要"), "件", "活動ログ", ""),
        ("再訪不要率（対 活動件数）", "=IFERROR(C{r0}/C{r1}*100,0)",
         "=IFERROR(D{r0}/D{r1}*100,0)", "%", "活動ログ",
         "北海道は3件に1件が案件化せず終了。東北の2.4倍"),
        ("残パイプライン（再アプローチ＋同行依頼）",
         pipe(["北海道"]), pipe(TOH), "件", "活動ログ",
         "次がある前提の案件。件数では東北のほうが多い"),
        ("うちアプローチ時期＝今年度",
         pipe(["北海道"], アプローチ時期="今年度"),
         pipe(TOH, アプローチ時期="今年度"), "件", "活動ログ",
         "今すぐ訪ねる理由のある案件。北海道は残パイプの40%、東北は66%"),
        ("アプローチ時期＝再来年度",
         ref.countifs(圏="北海道", アプローチ時期="再来年度"),
         tsum(アプローチ時期="再来年度"), "件", "活動ログ",
         "先送り案件。北海道に偏って積み上がっている"),
    ]
    xfirst = r + 1
    cnt_row, base_row = xfirst + 2, xfirst
    for i, (name, hk, tk, unit, srcname, read) in enumerate(metrics):
        rr = xfirst + i
        fill = "alt" if i % 2 else None
        label(ws, rr, 2, name, fill=fill)
        if name == "訪問済み団体数":
            body(ws, rr, 3, df[df["圏"] == "北海道"]["団体名"].nunique(), NUM, fill=fill, src=True)
            body(ws, rr, 4, df[df["圏"].isin(TOH)]["団体名"].nunique(), NUM, fill=fill, src=True)
        elif "率" in name:
            body(ws, rr, 3, hk.format(r0=cnt_row, r1=base_row), PCT1, fill=fill)
            body(ws, rr, 4, tk.format(r0=cnt_row, r1=base_row), PCT1, fill=fill)
        else:
            body(ws, rr, 3, hk, NUM, fill=fill)
            body(ws, rr, 4, tk, NUM, fill=fill)
        label(ws, rr, 5, unit, fill=fill)
        label(ws, rr, 6, srcname, fill=fill)
        body(ws, rr, 7, None, fill=fill)
        label(ws, rr, 8, read, fill=fill)

    # 営業会議ブック側の指標
    r = xfirst + len(metrics)
    book_rows = [
        ("見積提示済（今期分）", "見積今期", "件", "成果が今期に落ちていない"),
        ("見積提示済（来期分）", "見積来期", "件", "北海道の成果は44件すべて来期。今期に刈る玉がない"),
        ("9月入札見込", "9月入札見込", "件", ""),
        ("9月入札確定", "9月入札確定", "件", "北海道は9月の確定ゼロ。5〜8月は毎月2〜4件確定していた"),
        ("公会計 現状件数（納品）", "公会計現状", "件", "全国最多。納品側の母数が大きい"),
        ("固定資産台帳 精緻化", "精緻化", "件", "全国55件中39件が北海道。9月決算期の作業負荷が集中"),
        ("会計支援 対象", "会計支援対象", "件", ""),
        ("うち決算報告書", "決算報告書", "件", "9月は自治体の決算期。納品と営業が同じ人員なら営業は後回しになる"),
    ]
    for i, (name, key, unit, read) in enumerate(book_rows):
        rr = r + i
        fill = "band"
        label(ws, rr, 2, name, fill=fill)
        body(ws, rr, 3, BOOK[key]["北海道"][1], NUM, fill=fill, src=True)
        body(ws, rr, 4, BOOK[key]["東北"][1], NUM, fill=fill, src=True)
        label(ws, rr, 5, unit, fill=fill)
        label(ws, rr, 6, "会議ブック", fill=fill)
        label(ws, rr, 7, BOOK[key]["北海道"][0], fill=fill)
        label(ws, rr, 8, read, fill=fill)
    r = r + len(book_rows) + 1
    note(ws, f"B{r}", "※ 「会議ブック」出所の値は営業会議ブック 20260904㉜ からの転記（青字）。G列に参照セルを記載。")
    return r + 2


FACTORS_HOK = [
    ("① 今期に動かせる案件の枯渇", "強い（ログ・会議ブック双方が一致）",
     "残パイプ109件中 今年度分は44件（40%）。東北は141件中93件（66%）。"
     "再来年度案件19件が北海道に偏在。8/31週は38件訪問して22件（57.9%）が再訪不要。"
     "会議ブックでも見積提示済は今期1件・来期44件、9月入札は見込3件→確定0件。"),
    ("② 納品業務のピークと人員の重複", "中（構造は事実／投入の直接記録はなし）",
     "9月は自治体の決算期。北海道は公会計138件・精緻化39件（全国55件の7割）・"
     "決算報告書10件を抱え全国最多。札幌14名で最大拠点だが採用は一旦中止。"
     "営業と納品を同じ人員が担う場合、9月は納品が優先される。"),
    ("③ 広域ゆえの移動コスト", "中（効率差はログで確認できる）",
     "1実訪問日あたり 北海道2.4団体 vs 南東北・岩手3.9団体。"
     "未訪問の約71団体は道内に散在。今期案件が薄い先へ長距離移動する費用対効果は低く、①と重なって足が止まる。"),
    ("④ テーマ構成の偏り", "中（ログで確認できる）",
     "北海道は福祉計画が71件（31%）で東北11件（5%）と大きく異なる。"
     "福祉計画は全体でも再訪不要率28%と高く、自前作成・他計画への一体化が起きやすい領域。"),
    ("⑤ セミナー準備", "弱い（時期は一致するが担当が合わない）",
     "全国セミナーは9/25〜10/23開催、テキスト期限9/18〜9/22で今週が締切直前。"
     "ただし講師は川崎・北城・南澤・若山・山田・さいとう/おの・大窪で、"
     "北海道担当4名（長戸・村田・林・吉野）は含まれない。若山の停止（最終8/24）は説明できる。"),
]

UNKNOWNS = [
    ("記録漏れか活動停止かは区別できない",
     "活動ログに作成日時・更新日時の列がないため、「訪問していない」のか「訪問したが未入力」なのかを判定できない。"
     "9/4時点で会議ブックの累計営業件数は北海道288件・東北275件に対し、活動ログは北海道214件・東北176件。"
     "ログは北海道で26%、東北で36%少ない。記録漏れは全社的に存在し、むしろ東北のほうが大きい。"),
    ("個人の事情は一切データにない",
     "休暇・異動・体調・他案件へのアサインなどは活動ログにも会議ブックにも現れない。"
     "4名が同時期に止まっている点は、個別事情より①〜③の共通要因を示唆するが、確認は本人・上長への確認が必要。"),
    ("②③④は相関であって因果ではない",
     "納品件数・移動効率・テーマ構成はいずれも北海道の特徴だが、"
     "これらが今回の停止を引き起こしたことをログだけでは証明できない。"),
]


def build_hokkaido_factors(ws, r):
    title(ws, f"B{r}", "4. 要因の候補と根拠", size=12)
    r += 1
    head(ws, r, 2, "要因")
    head(ws, r, 3, "データでの支持度")
    head(ws, r, 5, "根拠")
    for c in (4, 6, 7, 8):
        head(ws, r, c, "")
    ffirst = r + 1
    for i, (name, strength, basis) in enumerate(FACTORS_HOK):
        rr = ffirst + i
        fill = "warn" if strength.startswith("強い") else ("alt" if i % 2 else None)
        label(ws, rr, 2, name, bold=True, fill=fill)
        c = label(ws, rr, 3, strength, fill=fill)
        c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        for cc in (4,):
            body(ws, rr, cc, None, fill=fill)
        c = label(ws, rr, 5, basis, fill=fill)
        c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        for cc in (6, 7, 8):
            body(ws, rr, cc, None, fill=fill)
        ws.row_dimensions[rr].height = 14 * (1 + len(basis) // 60)

    r = ffirst + len(FACTORS_HOK) + 1
    title(ws, f"B{r}", "5. このデータでは判定できないこと", size=12)
    r += 1
    head(ws, r, 2, "論点")
    head(ws, r, 5, "内容")
    for c in (3, 4, 6, 7, 8):
        head(ws, r, c, "")
    ufirst = r + 1
    for i, (name, text) in enumerate(UNKNOWNS):
        rr = ufirst + i
        label(ws, rr, 2, name, bold=True, fill="note")
        for cc in (3, 4):
            body(ws, rr, cc, None, fill="note")
        c = label(ws, rr, 5, text, fill="note")
        c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        for cc in (6, 7, 8):
            body(ws, rr, cc, None, fill="note")
        ws.row_dimensions[rr].height = 14 * (1 + len(text) // 60)


# ============================================================
# Sheet: 前提・注記
# ============================================================
def build_notes(ws, df):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 100

    title(ws, "B2", "前提・出所・注記")
    miss = {c: int((df[c] == BLANK).sum()) for c in
            ("結果", "契約予定", "状態", "アプローチ時期", "前回策定")}
    rows = [
        ("元データ", f"訪問活動ログ（Book1.xlsx／シート「{SRC_SHEET}」）{len(df)}行。"
                     "SharePointリストの書き出しと見られ、1行＝1団体1テーマの接触記録。"),
        ("期間", f"訪問日 {df['訪問日'].min():%Y/%m/%d}〜{df['訪問日'].max():%Y/%m/%d}（欠損8件）。"
                 "8/17週より入力者が7〜9名に増え、それ以前は実質1〜2名のみの記録。"
                 "したがって8月以前との時系列比較はできない。"),
        ("「今週」の定義", f"月曜起算で {THIS_WEEK:%Y/%m/%d}〜。基準日 {AS_OF:%Y/%m/%d}（木）時点の"
                          "最終入力は 9/16 で、9/14・9/15・9/16 の3日分のみ。週の途中の数値である。"),
        ("案件数の数え方", "ログ1行を1件として計上。同一団体でもテーマが異なれば別件。"
                          "団体数は重複を除いた実数を併記している。"),
        ("要因分類の方法", "備考テキストのキーワード判定を優先し、該当しない場合は結果区分"
                          "（自力→自前作成方針、他社随契→他社受託、不在→接触不可）で補完。"
                          "判定ロジックは activity_factor_rules.py に切り出しており、"
                          "「④再訪不要 明細」で1件ずつ検証・調整できる。102件中101件を分類、1件は備考・結果とも未記入。"),
        ("システム抱き合わせの判定", "備考に 行政研／ぎょうせい／TKC／MJS／アチカ／オーレンス、"
                                    "または「システム保守」「保守範囲」等が現れる行。"
                                    "会計システムの保守料の中で会計支援・消費税監査まで提供されており、"
                                    "「切替の意思なし」と記録されているケース。"),
        ("データ品質（未記入）", "　".join(f"{k} {v}件" for k, v in miss.items()) +
                                "。結果・契約予定の未記入が多く、受注確度の定量評価には使えない。"
                                "再訪不要率の母数は「次回アクション」記入分に限定している。"),
        ("集計方式", "すべての件数は「データ」シートへの COUNTIFS。元データを差し替えれば全シートが再計算される。"
                    "ただし団体数・稼働者数・訪問日数は一意カウントのため生成時点の実数を青字で埋め込んでいる。"),
        ("色の意味", "青字＝元データからの転記値および一意カウントの実数／黒字＝本ブックの計算式。"),
    ]
    r = 4
    for k, v in rows:
        label(ws, r, 2, k, bold=True, fill="band")
        c = body(ws, r, 3, v, align="left")
        c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        ws.row_dimensions[r].height = 14 * (1 + len(v) // 52)
        r += 1


# ============================================================
def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    df = load()
    ref = Ref(len(df))

    wb = Workbook()
    ws_sum = wb.active
    ws_sum.title = "サマリー"
    ws_w = wb.create_sheet("①週次活動量")
    ws_p = wb.create_sheet("②個人別案件数")
    ws_f = wb.create_sheet("③再訪不要_要因分析")
    ws_d = wb.create_sheet("④再訪不要_明細")
    ws_h = wb.create_sheet("⑤北海道_活動停止要因")
    ws_n = wb.create_sheet("前提・注記")
    build_data(wb, df)

    build_weekly(ws_w, df, ref)
    pfirst, plast = build_person(ws_p, df, ref)
    ffirst = build_factor(ws_f, df, ref)
    build_detail(ws_d, df)
    build_hokkaido_factors(ws_h, build_hokkaido(ws_h, df, ref))
    build_summary(ws_sum, df, ref, {
        "週次": {"name": ws_w.title},
        "個人": {"name": ws_p.title, "first": pfirst, "last": plast},
        "要因": {"name": ws_f.title, "ffirst": ffirst},
    })
    build_notes(ws_n, df)

    for ws in (ws_sum, ws_w, ws_p, ws_n):
        page_setup(ws, one_page=True)
    page_setup(ws_h)
    # 要因分析は行数が多く1ページ圧縮すると縮小率が大きいため自然な改ページにする
    page_setup(ws_f)
    page_setup(ws_d, freeze="B6")

    wb.active = 0
    wb.save(OUT_PATH)
    print(f"saved: {OUT_PATH}  rows={len(df)}")


if __name__ == "__main__":
    main()
