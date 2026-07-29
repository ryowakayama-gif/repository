# -*- coding: utf-8 -*-
"""大雪地区広域連合 第10期介護保険事業計画　調査結果のクロス集計・分析.

業務仕様書「４．業務内容（3）実施済み調査結果の集計・分析」に対応する。
対象4調査のうち、健康とくらしの調査（JAGES調査・令和7年度）の個票データにより
単純集計及びクロス集計を行い、地域包括ケア「見える化」システム、介護給付実績及び
人口推計と組み合わせて分析する。

個票データは個人情報を含むため本リポジトリには格納しない。
SRC で指定するディレクトリに配置して実行する。

シート構成
  00_目次と分析の前提
  01_地区別の主要指標
  02_年齢調整による地域比較
  03_外出手段と移動制約
  04_外出の抑制と生活動作の困りごと
  05_社会参加とリスクの関連
  06_世帯構成とリスク
  07_経済状況とリスク
  08_在宅生活の課題と解決手段
  09_住み続ける意向と転居理由
  10_介護の担い手と介護の意向
  11_KPIのデータ源の再検討
  12_主要所見
"""

import os
import math
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SRC = ("/tmp/claude-0/-home-user-repository/54f527c9-842b-534d-b822-e5a6c91f837c/scratchpad/"
       "z4/3-19大雪地区広域連合_2025健康とくらしの調査_個票データ・集計表_030")
CSV = os.path.join(SRC, "KK_2025_CSV_3-19大雪地区広域連合.csv")
OUT = "/home/user/repository/output/第10期計画_調査クロス集計・分析.xlsx"

FONT = "游ゴシック"
NAVY, HEAD = "1F4E78", "5B9BD5"
IN_Y, OK_G, NG_O, MID_B, GRAY = "FFF2CC", "E2EFDA", "FCE4D6", "DEEBF7", "F2F2F2"
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

# ============================================================ データの読込み
df = pd.read_csv(CSV, encoding="cp932", low_memory=False)
A = df[df["uncounted25"].isna()].copy()          # 集計対象 4,729票

TOWN = {1453: "東神楽町", 1458: "東川町", 1459: "美瑛町"}
AREA = {1453001: ("東神楽町", "東聖小学校区"), 1453002: ("東神楽町", "東神楽小学校区"),
        1453004: ("東神楽町", "志比内小学校区"),
        1458001: ("東川町", "東川第1小学校区"), 1458002: ("東川町", "東川第2小学校区"),
        1458003: ("東川町", "東川第3小学校区"), 1458004: ("東川町", "東川小学校区"),
        1459001: ("美瑛町", "美瑛小学校区"), 1459003: ("美瑛町", "美馬牛小学校区"),
        1459004: ("美瑛町", "美沢小学校区"), 1459005: ("美瑛町", "明徳小学校区"),
        1459007: ("美瑛町", "美瑛東小学校区")}
AGES = ["65～69歳", "70～74歳", "75～79歳", "80～84歳", "85歳以上"]

A["町"] = A["kcode25"].map(TOWN)
A["地区"] = A["scode1_25"].map(lambda v: AREA.get(v, ("―", "―"))[1])
A["年齢階級"] = pd.cut(pd.to_numeric(A["age_m25"], errors="coerce"),
                       bins=[65, 70, 75, 80, 85, 200], labels=AGES, right=False)


def num(col):
    return pd.to_numeric(A[col], errors="coerce")


def flag_in(col, vals):
    """単一選択の設問について、指定した選択肢に該当するかの0/1系列を返す。"""
    s = num(col)
    return s.isin(vals).astype(float).where(s.notna())


def any_of(cols, vals=(1,)):
    """複数回答の設問群について、いずれかに該当するかの0/1系列を返す。
    分母は当該設問群に1つでも回答のあった者とする。"""
    sub = A[list(cols)].apply(pd.to_numeric, errors="coerce")
    ok = sub.notna().any(axis=1)
    return sub.isin(vals).any(axis=1).astype(float).where(ok)


# 派生変数
SOC6 = ["cmnt6vl25", "cmnt6sp25", "cmnt6hb25", "cmnt6le25", "cmnt6sl25", "cmnt6sk25"]
SOC6 = [c for c in SOC6 if c in A.columns]
A["社会参加"] = any_of(SOC6, vals=(1, 2, 3, 4))          # 月1回以上
A["通いの場"] = flag_in("cmnt6sl25", [1, 2, 3, 4])
A["運転あり"] = flag_in("gout2dv25", [1])
A["同乗のみ"] = ((flag_in("gout2dv25", [1]) == 0) & (flag_in("gout2rd25", [1]) == 1)).astype(float)
GOUT = [c for c in A.columns if c.startswith("gout2") and c != "gout2ref25"]
A["自力移動なし"] = ((num("gout2dv25") != 1) & (num("gout2bi25") != 1)
                     & (num("gout2mo25") != 1)).astype(float).where(A[GOUT].notna().any(axis=1))
A["外出低頻度"] = flag_in("gout7fq25", [5, 6, 7])          # 月1～3回以下
A["外出減少"] = flag_in("gout4la25", [1, 2])
A["独居"] = pd.to_numeric(A["alone2gp_25"], errors="coerce")
A["夫婦のみ"] = ((num("hous2sp25") == 1) & (pd.to_numeric(A["mebr2nb25"], errors="coerce") == 2)
                 ).astype(float).where(num("hous2no25").notna())
A["暮らし苦しい"] = flag_in("sfs5_25", [1, 2])
A["介助不要"] = flag_in("adl3ra25", [1])
A["要介助・未受給"] = flag_in("adl3ra25", [2])
A["要介助・受給中"] = flag_in("adl3ra25", [3])
A["住み続けたい"] = flag_in("taisetsu_q3_1", [1])
A["転居も考える"] = flag_in("taisetsu_q3_1", [2])
Q21 = [c for c in A.columns if c.startswith("taisetsu_q2_1_s")]
Q21_NEED = [c for c in Q21 if not c.endswith("_s21")]          # 「特になし」を除く
A["困りごとあり"] = any_of(Q21_NEED)
A["解決できず困っている"] = flag_in("taisetsu_q2_2_s10", [1])
A["世話をする人なし"] = flag_in("taisetsu_q4_1_s9", [1])

IND = [
    ("frail2gp_25", "フレイルあり割合"),
    ("undo2gp_25", "運動機能低下者割合"),
    ("tojikomori2gp_25", "閉じこもり者割合"),
    ("utsucheck2gp_25", "うつ割合"),
    ("ninchicheck2gp_25", "認知機能低下者割合"),
    ("koukukinou2gp_25", "口腔機能低下者割合"),
    ("iadl1_2gp_25", "IADL低下者割合"),
    ("tento2gp_25", "1年間の転倒あり割合"),
    ("alone2gp_25", "独居者割合"),
    ("社会参加", "社会参加あり割合（月1回以上）"),
    ("通いの場", "通いの場参加割合（月1回以上）"),
]


def series(name):
    return pd.to_numeric(A[name], errors="coerce")


def rate(mask, name):
    """割合（％）と分母を返す。"""
    s = series(name)[mask]
    n = int(s.notna().sum())
    if n == 0:
        return None, 0
    return round(100 * s.sum() / n, 1), n


def wilson(p, n, z=1.96):
    """Wilson法による95％信頼区間（p は0〜1）。"""
    if n == 0:
        return None, None
    c = p + z * z / (2 * n)
    d = z * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n))
    return (c - d) / (1 + z * z / n), (c + d) / (1 + z * z / n)


def ci_text(mask, name):
    r, n = rate(mask, name)
    if r is None:
        return "―"
    lo, hi = wilson(r / 100, n)
    return "%.1f〜%.1f" % (100 * lo, 100 * hi)


def ztest(m1, m2, name):
    """2標本の母比率の差の検定（両側・正規近似）。判定文字列を返す。"""
    s1, s2 = series(name)[m1].dropna(), series(name)[m2].dropna()
    n1, n2 = len(s1), len(s2)
    if n1 < 10 or n2 < 10:
        return "判定不能"
    p1, p2 = s1.mean(), s2.mean()
    p = (s1.sum() + s2.sum()) / (n1 + n2)
    se = math.sqrt(p * (1 - p) * (1 / n1 + 1 / n2))
    if se == 0:
        return "判定不能"
    z = abs(p1 - p2) / se
    return "有意" if z >= 1.96 else "有意でない"


# 年齢調整（直接法・基準人口＝広域連合全体の年齢階級別回答者数）
STD = A["年齢階級"].value_counts().reindex(AGES).fillna(0)
STD_TOTAL = STD.sum()


def adjusted(mask, name):
    """直接法による年齢調整済み割合（％）。"""
    s = series(name)
    tot, wt = 0.0, 0.0
    for g in AGES:
        sub = s[mask & (A["年齢階級"] == g)].dropna()
        if len(sub) < 5:
            continue
        tot += sub.mean() * STD[g]
        wt += STD[g]
    if wt < STD_TOTAL * 0.8:
        return None
    return round(100 * tot / wt, 1)


# ============================================================ 出力
wb = Workbook()


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
    ws.row_dimensions[2].height = 50
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = freeze
    return ws


def header(ws, row, cols, height=32):
    for i, h in enumerate(cols, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = Font(name=FONT, size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=HEAD)
        c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
        c.border = BORDER
    ws.row_dimensions[row].height = height
    return row + 1


def body(ws, row, vals, fills=None, height=18, align=None, bold=False):
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


def note(ws, row, text, span=8):
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=9, italic=True)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    ws.row_dimensions[row].height = max(28, 12 * (len(text) // (span * 14) + 1))
    return row + 2


def lead(ws, row, text, span=8):
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=10, bold=True)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    ws.row_dimensions[row].height = 20
    return row + 1


ALL = pd.Series(True, index=A.index)
N_ALL = len(A)

# ============================================================ 00 目次
ws = sheet("00_目次と分析の前提", "調査結果のクロス集計・分析",
           "業務仕様書４（3）実施済み調査結果の集計・分析に対応する成果物。"
           "健康とくらしの調査（JAGES調査）令和7年度の個票データ4,729票により、"
           "単純集計、クロス集計及び見える化システム・給付実績・人口推計と組み合わせた分析を行う。"
           "作成日：令和8年7月29日。",
           [4, 26, 58, 34, 16], freeze="A4")

r = header(ws, 4, ["No.", "シート", "内容", "計画への反映先", "分母"])
for i, (nm, cont, dst, den) in enumerate([
    ("01_地区別の主要指標", "小学校区12地区別に主要11指標を集計。美瑛町の美馬牛・美沢・"
     "明徳・美瑛東の各小学校区が独立して取得できる。", "第1章第7節／第2章第4節", "地区別"),
    ("02_年齢調整による地域比較", "町間・地区間の差が年齢構成の違いによるものか、"
     "地域固有の差かを直接法の年齢調整により判別する。", "第2章第4節", "町・地区別"),
    ("03_外出手段と移動制約", "外出時の交通手段14区分の利用状況を町・地区別に集計。"
     "自分で運転できない層の分布を示す。", "第1章第7節／第5章基本目標2", "4,680票"),
    ("04_外出の抑制と生活動作の困りごと", "外出を控えている者の割合とその理由、"
     "外出頻度の低い者の割合。", "第2章第4節", "設問により異なる"),
    ("05_社会参加とリスクの関連", "社会参加の有無別のリスク指標。年齢階級で層化し、"
     "年齢の交絡を除いて比較する。", "第5章基本目標1", "4,611票"),
    ("06_世帯構成とリスク", "独居・夫婦のみ・その他の別によるリスク指標の差。",
     "第2章第1節／第5章基本目標2", "4,652票"),
    ("07_経済状況とリスク", "暮らし向き5段階別のリスク指標。経済的理由による外出抑制。",
     "第2章第2節3", "4,596票"),
    ("08_在宅生活の課題と解決手段", "大雪独自設問。生活動作の困りごと20項目と"
     "解決手段10項目。「解決できず、困っている」を含む。", "第2章第5節／第5章基本目標2", "4,128票"),
    ("09_住み続ける意向と転居理由", "大雪独自設問。住み続ける意向と、転居を考える理由"
     "（通院・買い物の困難、施設に入所できない、生活支援サービスがない）。",
     "第2章第6節／第6章第2節", "4,446票"),
    ("10_介護の担い手と介護の意向", "大雪独自設問。世話をしてくれる家族の有無、"
     "介護を受けたい場所、人生会議・死後事務委任の認知。", "第5章基本目標2・5", "4,599票"),
    ("11_KPIのデータ源の再検討", "算定不可としていた代表KPIのうち、本調査の大雪独自設問"
     "により算定できるものを整理する。", "第4章第3節／資料1", "―"),
    ("12_主要所見", "本分析から得られた所見と、計画への反映方針。", "各章", "―"),
], start=1):
    r = body(ws, r, [i, nm, cont, dst, den], height=40)

r += 1
r = lead(ws, r, "【分析の前提】")
r = header(ws, r, ["項目", "内容", "", "", ""])
for k, vtxt in [
    ("調査名", "令和7年度 健康とくらしの調査（日常生活圏域ニーズ調査／JAGES調査）"),
    ("実施期間", "令和7年11月17日〜12月8日"),
    ("対象", "大雪地区広域連合に居住する65歳以上の方 7,121人"),
    ("回収", "4,798票（回収率67.4％）。うち集計対象4,729票"),
    ("除外", "ID切り取り及び対象者名簿65歳未満の69票（uncounted25）を除外している"),
    ("分母の扱い", "指標ごとに有効回答者を分母とする。複数回答の設問は、当該設問群に"
     "1つ以上の回答があった者を分母とする。すべての表に分母を併記する"),
    ("調査票の版", "本調査は8種類の調査票（A〜H）を無作為に配布しており、"
     "各版に固有の設問（サブコア・セクションA〜H）の分母は550〜616票となる。"
     "共通設問及び大雪独自設問の分母は4,100〜4,700票である"),
    ("信頼区間", "母比率の95％信頼区間はWilson法による"),
    ("差の検定", "2群の比率の差は正規近似による両側検定（有意水準5％）。"
     "いずれかの群の分母が10未満の場合は判定不能とする"),
    ("年齢調整", "直接法。基準人口は広域連合全体の年齢階級別回答者数"
     "（65〜69歳968・70〜74歳1,204・75〜79歳1,246・80〜84歳784・85歳以上527）"),
    ("突合の確認", "本分析による広域連合全体及び町別の値は、"
     "調査報告書の公表値と完全に一致することを確認している"),
]:
    r = body(ws, r, [k, vtxt, "", "", ""], height=32)
    ws.merge_cells(start_row=r - 1, start_column=2, end_row=r - 1, end_column=5)

note(ws, r + 1,
     "注1）個票データは個人情報を含むため、成果物には集計値のみを収録している。"
     "注2）他の3調査（在宅生活改善調査、居所変更実態調査、介護人材実態調査）については、"
     "第10期分の調査票・回答データ・集計資料が未受領のため、本表ではクロス集計を行っていない。"
     "第9期計画に掲載された令和5年調査分の再整理は、図表集11〜13シートに収録している。"
     "注3）本分析により、大雪独自設問を用いれば、これまで算定不可としていた代表KPIの一部が"
     "算定できることが判明した。詳細は11シートに示す。", 5)

# ============================================================ 01 地区別
ws = sheet("01_地区別の主要指標", "小学校区別の主要指標",
           "分析地域1（小学校区）別に主要11指標を集計する。"
           "美瑛町については美瑛・美馬牛・美沢・明徳・美瑛東の5小学校区が独立して取得できるため、"
           "日常生活圏域別の分析及び訪問・通所困難地域の判定の基礎資料となる。"
           "分母が100票未満の地区は網掛けとし、値の変動が大きいことに留意する。",
           [16, 18, 8, 9] + [9] * 11)

cols = ["町", "地区", "回答者数", "75歳以上\n構成比"] + [t for _, t in IND]
r = header(ws, 4, cols, height=44)

rows = [("大雪地区広域連合", "全体", ALL)]
for t in ["東川町", "美瑛町", "東神楽町"]:
    rows.append((t, "町計", A["町"] == t))
    for code, (tn, an) in sorted(AREA.items()):
        if tn == t:
            rows.append(("", an, A["地区"] == an))

for tname, aname, mask in rows:
    n = int(mask.sum())
    o75 = A.loc[mask, "年齢階級"].isin(["75～79歳", "80～84歳", "85歳以上"]).mean()
    vals = [tname, aname, n, round(100 * o75, 1)]
    for key, _ in IND:
        v, _n = rate(mask, key)
        vals.append(v)
    small = n < 100
    fills = {3: NG_O} if small else ({2: MID_B} if aname == "町計" else None)
    if aname == "全体":
        fills = {i: GRAY for i in range(1, 16)}
    r = body(ws, r, vals, fills, height=18, align={1: "left", 2: "left"},
             bold=(aname in ("全体", "町計")))

note(ws, r + 1,
     "注1）単位は％。回答者数のみ人。「75歳以上構成比」は回答者に占める75歳以上の割合で、"
     "地区間の年齢構成の差を示す。地区間の比較は年齢調整後の値（02シート）による。"
     "注2）分母が100票未満の地区（志比内小学校区21票、明徳小学校区49票、美沢小学校区57票）は"
     "回答者数を網掛けとした。これらの地区の値は95％信頼区間が広く、単独では判断材料としない。"
     "注3）小学校区と第9期計画の日常生活圏域（美瑛町の旭・北西地区、美馬牛地区、朗根内地区、"
     "市街地・周辺地区）との対応関係は確定していない。"
     "美馬牛小学校区は美馬牛地区に対応すると考えられるが、"
     "他の小学校区との対応は美瑛町への確認を要する（確認事項No.12）。"
     "注4）東川小学校区は1,040票と東川町の回答の73.5％を占める。"
     "東川第1〜第3小学校区は旧小学校区であり、統合後の区域との関係の確認を要する。", 15)

# ============================================================ 02 年齢調整
ws = sheet("02_年齢調整による地域比較", "年齢調整による町間・地区間の比較",
           "町別・地区別の指標の差が年齢構成の違いによるものか、地域固有の差かを判別する。"
           "直接法により、広域連合全体の年齢階級別回答者数を基準人口として調整した。"
           "粗率と調整率の差が大きい地域は、年齢構成の影響が大きいことを示す。",
           [16, 18, 8] + [11] * 10)

SEL = ["frail2gp_25", "tojikomori2gp_25", "iadl1_2gp_25", "utsucheck2gp_25", "社会参加"]
SELN = ["フレイル", "閉じこもり", "IADL低下", "うつ", "社会参加あり"]
hdr = ["町", "地区", "回答者数"]
for nm in SELN:
    hdr += [nm + "\n粗率", nm + "\n年齢調整"]
r = header(ws, 4, hdr, height=40)

for tname, aname, mask in rows:
    n = int(mask.sum())
    vals = [tname, aname, n]
    for key in SEL:
        v, _ = rate(mask, key)
        vals += [v, adjusted(mask, key)]
    fills = {i: GRAY for i in range(1, 14)} if aname == "全体" else (
        {2: MID_B} if aname == "町計" else None)
    r = body(ws, r, vals, fills, height=18, align={1: "left", 2: "left"},
             bold=(aname in ("全体", "町計")))

r += 1
r = lead(ws, r, "【町間の差の検定（粗率）】")
r = header(ws, r, ["指標", "東川町", "美瑛町", "東神楽町", "最大差",
                   "美瑛町と東神楽町の差の検定", "", "", "", "", "", "", ""])
mB, mH, mHG = A["町"] == "美瑛町", A["町"] == "東川町", A["町"] == "東神楽町"
for key, nm in IND:
    vH, _ = rate(mH, key)
    vB, _ = rate(mB, key)
    vG, _ = rate(mHG, key)
    if None in (vH, vB, vG):
        continue
    r = body(ws, r, [nm, vH, vB, vG, round(max(vH, vB, vG) - min(vH, vB, vG), 1),
                     ztest(mB, mHG, key), "", "", "", "", "", "", ""], height=18)

note(ws, r + 1,
     "注1）単位は％。年齢調整は直接法（基準人口＝広域連合全体の年齢階級別回答者数）による。"
     "いずれかの年齢階級の分母が5票未満の場合は当該階級を除外し、"
     "基準人口の8割に満たない場合は算定していない。"
     "注2）美瑛町のフレイル該当割合21.6％は東神楽町17.0％より4.6ポイント高いが、"
     "美瑛町は回答者に占める75歳以上の割合が高い。"
     "年齢調整後の値を比較することで、年齢構成の影響を除いた地域差を確認できる。"
     "注3）検定は2標本の母比率の差の検定（両側・有意水準5％）による。", 13)

# ============================================================ 03 外出手段
ws = sheet("03_外出手段と移動制約", "外出時の交通手段と移動制約",
           "外出時に利用している交通手段（複数回答）を町別・地区別に集計する。"
           "美瑛町では国鉄バスの路線廃止によりスクールバスが交通弱者の公共交通となっており、"
           "町道414路線346.6kmの除雪率は52.9％である。"
           "自分で運転できない層の分布は、訪問・通所困難地域の判定及び"
           "通所系サービスの利用可能性の検討に直結する。",
           [16, 18, 8] + [9] * 12)

MEANS = [("gout2wk25", "徒歩"), ("gout2bi25", "自転車"), ("gout2dv25", "自動車\n（自分で運転）"),
         ("gout2rd25", "自動車\n（乗せてもらう）"), ("gout2pb25", "路線バス"),
         ("gout2cb25", "コミュニティ\nバス"), ("gout2tx25", "タクシー"),
         ("gout2fc25", "病院や施設\nのバス"), ("gout2tr25", "電車"),
         ("gout2sc25", "歩行器・\nシルバーカー")]
r = header(ws, 4, ["町", "地区", "回答者数"] + [t for _, t in MEANS]
           + ["自力移動の\n手段なし", "外出頻度が\n月1〜3回以下"], height=44)

for tname, aname, mask in rows:
    n = int(mask.sum())
    vals = [tname, aname, n]
    for key, _ in MEANS:
        v, _ = rate(mask, key)
        vals.append(v)
    vals.append(rate(mask, "自力移動なし")[0])
    vals.append(rate(mask, "外出低頻度")[0])
    fills = {i: GRAY for i in range(1, 16)} if aname == "全体" else (
        {2: MID_B} if aname == "町計" else None)
    r = body(ws, r, vals, fills, height=18, align={1: "left", 2: "left"},
             bold=(aname in ("全体", "町計")))

r += 1
r = lead(ws, r, "【自分で運転しない層のリスク指標】")
r = header(ws, r, ["区分", "回答者数"] + [t for _, t in IND[:8]] + ["", "", "", "", ""])
for lb, m in [("自分で運転する", series("運転あり") == 1),
              ("自分では運転しない", series("運転あり") == 0),
              ("（再掲）運転せず、他者の運転で移動する", (series("運転あり") == 0)
               & (num("gout2rd25") == 1)),
              ("（再掲）自力移動の手段なし（運転・自転車・バイクのいずれもなし）",
               series("自力移動なし") == 1)]:
    vals = [lb, int(m.sum())] + [rate(m, k)[0] for k, _ in IND[:8]] + ["", "", "", "", ""]
    r = body(ws, r, vals, height=18)
r = body(ws, r, ["差の検定（運転する／しない）", ""]
         + [ztest(series("運転あり") == 1, series("運転あり") == 0, k) for k, _ in IND[:8]]
         + ["", "", "", "", ""], {1: IN_Y}, height=18)

note(ws, r + 1,
     "注1）単位は％。交通手段は複数回答のため合計は100％を超える。"
     "分母は交通手段の設問に1つ以上回答した4,680票。"
     "注2）「自力移動の手段なし」は、自分での自動車運転・自転車・バイクのいずれも"
     "利用していない者の割合である。"
     "注3）自分で運転しない層はフレイル・閉じこもり・IADL低下の割合が高いが、"
     "この差には年齢の交絡が含まれる。運転の可否は年齢と強く相関するため、"
     "因果関係の解釈には年齢調整又は年齢階級別の比較を要する。"
     "注4）路線バス・コミュニティバスの利用割合が低い地区は、"
     "公共交通による通所が現実的でないことを示唆する。"
     "第10期基本指針案の別表 一が求める「訪問・通所困難地域」の判定においては、"
     "事業所の通常の事業の実施地域（確認事項No.13）と本表を併せて用いる。"
     "注5）「（再掲）」の2行は「自分では運転しない」の内訳であるが、"
     "他者の運転による移動と自力移動の手段なしは重複するため、合計は一致しない。"
     "注6）志比内小学校区は21票であり、徒歩0％・路線バス0％等の値は標本の偏りによる。"
     "同地区単独では判断材料としない。", 15)

# ============================================================ 04 外出の抑制
ws = sheet("04_外出の抑制と生活動作の困りごと", "外出の抑制とその理由",
           "外出を控えているかどうか、及びその理由を集計する。"
           "外出の抑制に関する設問はセクションAの調査票にのみ含まれるため分母が小さい。"
           "外出頻度及び前年比の変化は共通設問であり分母が大きい。",
           [30, 10, 10, 10, 10, 12, 12, 12])

r = lead(ws, 4, "【外出頻度（共通設問・全数）】")
r = header(ws, r, ["区分", "広域連合", "東川町", "美瑛町", "東神楽町", "回答者数", "", ""])
FQ = [("週に5回以上", [1]), ("週4回", [2]), ("週2〜3回", [3]), ("週1回", [4]),
      ("月1〜3回", [5]), ("年に数回", [6]), ("していない", [7])]
for lb, vs in FQ:
    A["_t"] = flag_in("gout7fq25", vs)
    r = body(ws, r, [lb, rate(ALL, "_t")[0], rate(mH, "_t")[0], rate(mB, "_t")[0],
                     rate(mHG, "_t")[0], rate(ALL, "_t")[1], "", ""], height=18)
A["_t"] = flag_in("gout7fq25", [5, 6, 7])
r = body(ws, r, ["【再掲】月1〜3回以下", rate(ALL, "_t")[0], rate(mH, "_t")[0],
                 rate(mB, "_t")[0], rate(mHG, "_t")[0], rate(ALL, "_t")[1], "", ""],
         {1: IN_Y}, height=18, bold=True)

r += 1
r = lead(ws, r, "【前年と比べた外出回数（共通設問・全数）】")
r = header(ws, r, ["区分", "広域連合", "東川町", "美瑛町", "東神楽町", "回答者数", "", ""])
for lb, vs in [("とても減っている", [1]), ("減っている", [2]),
               ("あまり減っていない", [3]), ("減っていない", [4])]:
    A["_t"] = flag_in("gout4la25", vs)
    r = body(ws, r, [lb, rate(ALL, "_t")[0], rate(mH, "_t")[0], rate(mB, "_t")[0],
                     rate(mHG, "_t")[0], rate(ALL, "_t")[1], "", ""], height=18)

r += 1
r = lead(ws, r, "【外出を控えているか（セクションAの調査票のみ）】")
r = header(ws, r, ["区分", "割合", "該当数", "分母", "", "", "", ""])
for lb, vs in [("控えている", [1]), ("控えていない", [2])]:
    A["_t"] = flag_in("gout2ref25", vs)
    v, n = rate(ALL, "_t")
    r = body(ws, r, [lb, v, int(series("_t").sum()), n, "", "", "", ""], height=18)

r += 1
r = lead(ws, r, "【外出を控えている理由（外出を控えている者が分母）】")
r = header(ws, r, ["理由", "割合", "該当数", "分母", "計画上の意味", "", "", ""])
GORE = [("gore2pa25", "足腰などの痛み", "介護予防・リハビリテーション"),
        ("gore2dis25", "病気", "医療との連携"),
        ("gore2tr25", "交通手段がない", "移動支援・訪問通所困難地域"),
        ("gore2fu25", "外での楽しみがない", "通いの場・社会参加"),
        ("gore2re25", "トイレの心配（失禁など）", "在宅生活の支援"),
        ("gore2ec25", "経済的に出られない", "低所得層の支援"),
        ("gore2ea25", "耳の障害", "コミュニケーション支援"),
        ("gore2ey25", "目の障害", "コミュニケーション支援"),
        ("gore2st25", "障害（脳卒中の後遺症など）", "医療との連携"),
        ("gore2ot25", "その他", "―")]
mref = series("gout2ref25_flag") if "gout2ref25_flag" in A.columns else (num("gout2ref25") == 1)
for key, lb, mean in GORE:
    s = pd.to_numeric(A[key], errors="coerce")[mref]
    n = int(s.notna().sum())
    v = round(100 * s.sum() / n, 1) if n else None
    r = body(ws, r, [lb, v, int(s.sum()) if n else 0, n, mean, "", "", ""], height=18)

note(ws, r + 1,
     "注1）単位は％。"
     "注2）外出の抑制に関する設問（Ａ【問19】6）7））はセクションAの調査票にのみ含まれ、"
     "外出を控えていると回答した者は111人、理由の回答は106票である。"
     "分母が小さいため、理由別の割合は参考値として扱い、"
     "町別・地区別の内訳は算定していない。"
     "注3）外出頻度及び前年比の変化は全数の共通設問であり、"
     "町別・地区別の比較に用いることができる。"
     "注4）「交通手段がない」を理由とする者の割合は、"
     "移動支援施策及び通所系サービスの送迎体制の検討に用いる。", 8)

# ============================================================ 05 社会参加
ws = sheet("05_社会参加とリスクの関連", "社会参加の有無とリスク指標の関連",
           "社会参加（ボランティア・スポーツの会・趣味の会・学習教養サークル・通いの場・"
           "特技を伝える活動の6区分のいずれかに月1回以上参加）の有無別にリスク指標を比較する。"
           "社会参加は年齢と相関するため、年齢階級で層化した比較を併せて示す。",
           [26, 10, 10, 10, 10, 10, 10, 12, 12])

msoc1, msoc0 = series("社会参加") == 1, series("社会参加") == 0
r = lead(ws, 4, "【社会参加の有無別（全年齢）】")
r = header(ws, r, ["指標", "参加あり", "参加なし", "差\n（ポイント）", "検定",
                   "参加あり\n分母", "参加なし\n分母", "参加あり\n95％信頼区間",
                   "参加なし\n95％信頼区間"], height=40)
for key, nm in [("frail2gp_25", "フレイルあり割合"), ("undo2gp_25", "運動機能低下者割合"),
                ("tojikomori2gp_25", "閉じこもり者割合"), ("utsucheck2gp_25", "うつ割合"),
                ("ninchicheck2gp_25", "認知機能低下者割合"),
                ("koukukinou2gp_25", "口腔機能低下者割合"),
                ("iadl1_2gp_25", "IADL低下者割合"), ("tento2gp_25", "1年間の転倒あり割合"),
                ("happy2gp_25", "幸福感がある者の割合"), ("eatalone2gp_25", "孤食者割合")]:
    v1, n1 = rate(msoc1, key)
    v0, n0 = rate(msoc0, key)
    if v1 is None or v0 is None:
        continue
    r = body(ws, r, [nm, v1, v0, round(v1 - v0, 1), ztest(msoc1, msoc0, key), n1, n0,
                     ci_text(msoc1, key), ci_text(msoc0, key)], height=18)
A["_rs"] = num("riskscore_25")
r = body(ws, r, ["要支援・要介護リスク点数（平均点）",
                 round(series("_rs")[msoc1].mean(), 1), round(series("_rs")[msoc0].mean(), 1),
                 round(series("_rs")[msoc1].mean() - series("_rs")[msoc0].mean(), 1),
                 "―", int(series("_rs")[msoc1].notna().sum()),
                 int(series("_rs")[msoc0].notna().sum()), "―", "―"], {1: IN_Y}, height=18)

r += 1
r = lead(ws, r, "【年齢階級で層化したフレイル該当割合】")
r = header(ws, r, ["年齢階級", "参加あり", "参加なし", "差\n（ポイント）", "検定",
                   "参加あり\n分母", "参加なし\n分母", "", ""])
for g in AGES:
    m1 = msoc1 & (A["年齢階級"] == g)
    m0 = msoc0 & (A["年齢階級"] == g)
    v1, n1 = rate(m1, "frail2gp_25")
    v0, n0 = rate(m0, "frail2gp_25")
    r = body(ws, r, [g, v1, v0, round(v1 - v0, 1) if None not in (v1, v0) else None,
                     ztest(m1, m0, "frail2gp_25"), n1, n0, "", ""], height=18)

r += 1
r = lead(ws, r, "【参加の区分別の参加率（月1回以上）】")
r = header(ws, r, ["区分", "広域連合", "東川町", "美瑛町", "東神楽町", "分母", "", "", ""])
SOCN = [("cmnt6vl25", "ボランティアのグループ"), ("cmnt6le25", "学習・教養サークル"),
        ("cmnt6sl25", "介護予防のための通いの場"), ("cmnt6sk25", "特技や経験を伝える活動")]
for key, nm in SOCN:
    if key not in A.columns:
        continue
    A["_t"] = flag_in(key, [1, 2, 3, 4])
    r = body(ws, r, [nm, rate(ALL, "_t")[0], rate(mH, "_t")[0], rate(mB, "_t")[0],
                     rate(mHG, "_t")[0], rate(ALL, "_t")[1], "", "", ""], height=18)
r = body(ws, r, ["【合成】6区分のいずれかに参加", rate(ALL, "社会参加")[0],
                 rate(mH, "社会参加")[0], rate(mB, "社会参加")[0], rate(mHG, "社会参加")[0],
                 rate(ALL, "社会参加")[1], "", "", ""], {1: IN_Y}, height=18, bold=True)

note(ws, r + 1,
     "注1）単位は％（リスク点数のみ点）。"
     "注2）社会参加ありの群はフレイル・閉じこもり・うつ・IADL低下の割合がいずれも低い。"
     "ただし、これは社会参加が心身の状態を保っている効果と、"
     "心身の状態が良いから参加できているという逆方向の関係の両方を含む。"
     "本調査は横断調査であるため、因果関係を示すものではない。"
     "注3）年齢階級で層化しても同じ方向の差がみられることは、"
     "年齢の交絡だけでは説明できないことを示す。"
     "注4）代表KPI H05（社会参加率）の基準値としては、"
     "第9期の通いの場参加率7.3％を継承する場合は「通いの場」の値を、"
     "社会参加の広がりを捉える場合は合成指標を用いる。"
     "注5）要支援・要介護リスク点数のみ参加あり群の方がわずかに高いが、"
     "これは社会参加の割合が65〜69歳28.9％に対し80〜84歳43.1％と高齢層で高く、"
     "参加あり群の年齢構成が高いことによる。"
     "同点数は年齢と強く相関するため（65〜69歳5.1点、85歳以上30.0点）、"
     "年齢調整を行わない群間比較には適さない。"
     "年齢階級で層化した比較では、他の指標と同様に参加あり群が良好である。", 9)

# ============================================================ 06 世帯構成
ws = sheet("06_世帯構成とリスク", "世帯構成別のリスク指標",
           "独居・夫婦のみ・その他の別にリスク指標を比較する。"
           "令和2年国勢調査では、高齢者を含む世帯に占める高齢者単身世帯の割合は"
           "美瑛町が30.3％と最も高く、高齢者夫婦世帯の割合は東神楽町が37.1％と最も高い。",
           [26, 10, 10, 10, 10, 10, 12, 12, 12])

mal, mcp = series("独居") == 1, series("夫婦のみ") == 1
moth = (series("独居") == 0) & (series("夫婦のみ") != 1)
r = lead(ws, 4, "【世帯構成別のリスク指標】")
r = header(ws, r, ["指標", "独居", "夫婦のみ", "その他", "独居とその他の差",
                   "検定", "独居\n分母", "夫婦のみ\n分母", "その他\n分母"], height=40)
for key, nm in [("frail2gp_25", "フレイルあり割合"), ("tojikomori2gp_25", "閉じこもり者割合"),
                ("utsucheck2gp_25", "うつ割合"), ("ninchicheck2gp_25", "認知機能低下者割合"),
                ("iadl1_2gp_25", "IADL低下者割合"), ("eatalone2gp_25", "孤食者割合"),
                ("社会参加", "社会参加あり割合"), ("通いの場", "通いの場参加割合"),
                ("happy2gp_25", "幸福感がある者の割合"),
                ("困りごとあり", "生活動作の困りごとあり割合"),
                ("世話をする人なし", "世話をしてくれる人がいない割合")]:
    v1, n1 = rate(mal, key)
    v2, n2 = rate(mcp, key)
    v3, n3 = rate(moth, key)
    if v1 is None or v3 is None:
        continue
    r = body(ws, r, [nm, v1, v2, v3, round(v1 - v3, 1), ztest(mal, moth, key),
                     n1, n2, n3], height=18)

r += 1
r = lead(ws, r, "【町別・地区別の独居者割合】")
r = header(ws, r, ["町", "地区", "回答者数", "独居者割合", "夫婦のみ割合",
                   "世話をしてくれる\n人がいない割合", "", "", ""], height=36)
for tname, aname, mask in rows:
    r = body(ws, r, [tname, aname, int(mask.sum()), rate(mask, "独居")[0],
                     rate(mask, "夫婦のみ")[0], rate(mask, "世話をする人なし")[0],
                     "", "", ""], height=18,
             bold=(aname in ("全体", "町計")))

note(ws, r + 1,
     "注1）単位は％。「独居」はリスク指標58（独居者割合）による。"
     "「夫婦のみ」は同居者が配偶者のみで世帯人員が2人の者とした。"
     "注2）独居者はフレイル・閉じこもり・うつ・孤食の割合が高く、"
     "社会参加の割合も低い傾向がある。"
     "注3）「世話をしてくれる人がいない割合」は大雪独自設問【問4】1）による。"
     "この指標は、在宅生活の継続可能性及び身寄りのない高齢者への支援の"
     "検討に用いることができる（第5章 基本目標2・3）。", 9)

# ============================================================ 07 経済状況
ws = sheet("07_経済状況とリスク", "経済状況とリスク指標",
           "暮らし向き（経済的にみた現在の暮らしの状況）の5段階別にリスク指標を比較する。"
           "認定を受けながらサービスを利用していない方が504人（認定者の25.7％）おり、"
           "そのうちに経済的理由によるものが含まれるかどうかは第10期の論点である。",
           [26, 11, 11, 11, 11, 11, 11, 12, 12])

SFS = [("大変苦しい", [1]), ("やや苦しい", [2]), ("ふつう", [3]),
       ("ややゆとりがある", [4]), ("大変ゆとりがある", [5])]
r = lead(ws, 4, "【暮らし向き別のリスク指標】")
r = header(ws, r, ["指標"] + [n for n, _ in SFS] + ["苦しい計\n（1・2）", "ゆとり計\n（4・5）",
                                                     "苦しい計と\nふつうの差の検定"], height=40)
msfs = {n: flag_in("sfs5_25", v) == 1 for n, v in SFS}
mhard = series("暮らし苦しい") == 1
mnorm = flag_in("sfs5_25", [3]) == 1
measy = flag_in("sfs5_25", [4, 5]) == 1
for key, nm in [("frail2gp_25", "フレイルあり割合"), ("tojikomori2gp_25", "閉じこもり者割合"),
                ("utsucheck2gp_25", "うつ割合"), ("iadl1_2gp_25", "IADL低下者割合"),
                ("teieiyo2gp_25", "低栄養者割合"), ("eatalone2gp_25", "孤食者割合"),
                ("社会参加", "社会参加あり割合"), ("通いの場", "通いの場参加割合"),
                ("happy2gp_25", "幸福感がある者の割合"), ("独居", "独居者割合"),
                ("困りごとあり", "生活動作の困りごとあり割合"),
                ("解決できず困っている", "解決できず困っている割合")]:
    if key not in A.columns:
        continue
    vs = [rate(msfs[n], key)[0] for n, _ in SFS]
    if all(v is None for v in vs):
        continue
    r = body(ws, r, [nm] + vs + [rate(mhard, key)[0], rate(measy, key)[0],
                                 ztest(mhard, mnorm, key)], height=18)
r = body(ws, r, ["回答者数（人）"] + [int(msfs[n].sum()) for n, _ in SFS]
         + [int(mhard.sum()), int(measy.sum()), ""], {1: GRAY}, height=18)

r += 1
r = lead(ws, r, "【町別の経済状況】")
r = header(ws, r, ["町", "大変苦しい", "やや苦しい", "苦しい計", "ふつう",
                   "ゆとり計", "回答者数", "", ""])
for tname, m in [("大雪地区広域連合", ALL), ("東川町", mH), ("美瑛町", mB), ("東神楽町", mHG)]:
    A["_h"] = flag_in("sfs5_25", [1])
    A["_h2"] = flag_in("sfs5_25", [2])
    A["_n"] = flag_in("sfs5_25", [3])
    A["_e"] = flag_in("sfs5_25", [4, 5])
    r = body(ws, r, [tname, rate(m, "_h")[0], rate(m, "_h2")[0], rate(m, "暮らし苦しい")[0],
                     rate(m, "_n")[0], rate(m, "_e")[0], rate(m, "暮らし苦しい")[1],
                     "", ""], height=18, bold=(tname == "大雪地区広域連合"))

note(ws, r + 1,
     "注1）単位は％。分母は暮らし向きの設問に回答した4,596票。"
     "注2）暮らし向きが「苦しい」層はフレイル・閉じこもり・うつ・孤食の割合が高く、"
     "社会参加の割合が低い。経済状況と心身の状態及び社会参加は関連している。"
     "注3）認定を受けながらサービスを利用していない504人について、"
     "経済的理由によるものがあるかどうかは本調査からは直接判別できない。"
     "本調査の対象は認定の有無を問わない65歳以上であり、"
     "未利用認定者を特定する情報を含まないためである。"
     "未利用の要因の判別には、居宅サービス計画未作成者の抽出等の代理指標を用いる"
     "（素案 第2章第2節3）。", 9)

# ============================================================ 08 在宅生活の課題
ws = sheet("08_在宅生活の課題と解決手段", "生活動作の困りごとと解決手段（大雪独自設問）",
           "【大雪－問2】生活動作の中で不安や困っていると感じていることと、"
           "その解決手段を集計する。本設問は当広域連合が独自に追加したものであり、"
           "在宅生活改善調査が把握しようとしている在宅生活の継続困難要因に対応する。",
           [30, 10, 10, 10, 10, 10, 12, 12, 12])

Q21LAB = [("taisetsu_q2_1_s12", "除雪"), ("taisetsu_q2_1_s4", "庭の手入れ"),
          ("taisetsu_q2_1_s13", "災害時の避難"), ("taisetsu_q2_1_s15", "布団干し"),
          ("taisetsu_q2_1_s3", "風呂やトイレの掃除"), ("taisetsu_q2_1_s2", "部屋の掃除や片付け"),
          ("taisetsu_q2_1_s14", "簡単な修繕や電球替え"), ("taisetsu_q2_1_s16", "季節の衣服の入れ替え"),
          ("taisetsu_q2_1_s1", "食事の準備や片付け"), ("taisetsu_q2_1_s7", "買い物"),
          ("taisetsu_q2_1_s8", "通院"), ("taisetsu_q2_1_s11", "外出"),
          ("taisetsu_q2_1_s6", "ゴミの分別やゴミ出し"), ("taisetsu_q2_1_s5", "衣服の洗濯や片付け"),
          ("taisetsu_q2_1_s9", "薬の管理"), ("taisetsu_q2_1_s10", "預貯金の管理"),
          ("taisetsu_q2_1_s17", "犬の散歩などペットの世話"),
          ("taisetsu_q2_1_s18", "話し相手がいない"), ("taisetsu_q2_1_s19", "趣味や役割がない"),
          ("taisetsu_q2_1_s20", "その他"), ("taisetsu_q2_1_s21", "特になし")]
Q21LAB = [(k, n) for k, n in Q21LAB if k in A.columns]

r = lead(ws, 4, "【生活動作の困りごと（複数回答・全数）】")
r = header(ws, r, ["項目", "広域連合", "東川町", "美瑛町", "東神楽町", "該当数", "独居者",
                   "75歳以上", "美瑛町と東神楽町\nの差の検定"], height=40)
m75 = A["年齢階級"].isin(["75～79歳", "80～84歳", "85歳以上"])
res = []
for key, nm in Q21LAB:
    if nm == "特になし":
        continue
    A["_t"] = pd.to_numeric(A[key], errors="coerce")
    v, n = rate(ALL, "_t")
    res.append((v if v is not None else -1, key, nm))
res = sorted(res, reverse=True)
res += [(-1, k, n) for k, n in Q21LAB if n == "特になし"]
for _, key, nm in res:
    A["_t"] = pd.to_numeric(A[key], errors="coerce")
    v, n = rate(ALL, "_t")
    fill = {1: IN_Y} if nm == "特になし" else None
    r = body(ws, r, [nm, v, rate(mH, "_t")[0], rate(mB, "_t")[0], rate(mHG, "_t")[0],
                     int(series("_t").sum()), rate(mal, "_t")[0], rate(m75, "_t")[0],
                     ztest(mB, mHG, "_t")], fill, height=18)
r = body(ws, r, ["【再掲】いずれかの困りごとあり", rate(ALL, "困りごとあり")[0],
                 rate(mH, "困りごとあり")[0], rate(mB, "困りごとあり")[0],
                 rate(mHG, "困りごとあり")[0], int(series("困りごとあり").sum()),
                 rate(mal, "困りごとあり")[0], rate(m75, "困りごとあり")[0],
                 ztest(mB, mHG, "困りごとあり")], {1: OK_G}, height=18, bold=True)

r += 1
r = lead(ws, r, "【困りごとの解決手段（困りごとがある者が分母）】")
r = header(ws, r, ["解決手段", "割合", "該当数", "分母", "独居者", "75歳以上",
                   "計画上の意味", "", ""], height=32)
Q22 = [("taisetsu_q2_2_s8", "自力で何とかしている", "支援につながっていない可能性"),
       ("taisetsu_q2_2_s1", "家族や親族の手助け", "家族介護力への依存"),
       ("taisetsu_q2_2_s2", "近所に住む方々の手助け", "地域の支え合い"),
       ("taisetsu_q2_2_s3", "友人や知人のサポート", "地域の支え合い"),
       ("taisetsu_q2_2_s6", "介護保険などのサービス", "保険給付でのカバー"),
       ("taisetsu_q2_2_s7", "民間業者のサービス", "インフォーマルサービス"),
       ("taisetsu_q2_2_s5", "シルバー人材センターなどのサービス", "生活支援体制整備"),
       ("taisetsu_q2_2_s4", "ボランティアのサポート", "生活支援体制整備"),
       ("taisetsu_q2_2_s10", "解決できず、困っている", "未充足ニーズ（H06の候補）"),
       ("taisetsu_q2_2_s9", "その他", "―")]
mneed = series("困りごとあり") == 1
for key, nm, mean in Q22:
    if key not in A.columns:
        continue
    s = pd.to_numeric(A[key], errors="coerce")[mneed]
    n = int(s.notna().sum())
    v = round(100 * s.sum() / n, 1) if n else None
    A["_t"] = pd.to_numeric(A[key], errors="coerce")
    fill = {1: NG_O} if nm == "解決できず、困っている" else None
    r = body(ws, r, [nm, v, int(s.sum()) if n else 0, n,
                     rate(mneed & mal, "_t")[0], rate(mneed & m75, "_t")[0], mean, "", ""],
             fill, height=18)

note(ws, r + 1,
     "注1）単位は％。複数回答のため合計は100％を超える。"
     "困りごとの分母は当該設問群に回答した4,128票、"
     "解決手段の分母は困りごとがあると回答した者である。"
     "注2）「除雪」が最上位となることは、当広域連合の3町がいずれも豪雪地帯対策特別措置法の"
     "特別豪雪地帯又は豪雪地帯に指定されていること、"
     "美瑛町の町道除雪率が52.9％であることと整合する。"
     "除雪は介護保険の給付対象外であり、生活支援体制整備事業及び3町の施策で対応する領域である。"
     "注3）「解決できず、困っている」と回答した者の割合は、"
     "在宅生活改善調査が把握しようとしている在宅生活の継続困難要因に対応しており、"
     "代表KPI H06（在宅生活継続困難割合）の基準値として用いることができる（11シート）。"
     "注4）「自力で何とかしている」の割合が高いことは、"
     "支援につながっていない層が一定数存在することを示唆する。", 9)

# ============================================================ 09 住み続ける意向
ws = sheet("09_住み続ける意向と転居理由", "住み続ける意向と転居を考える理由（大雪独自設問）",
           "【大雪－問3】現在住んでいる地域に住み続けたいかどうかと、転居を考える理由を集計する。"
           "転居理由には「今住んでいる地域では通院や買い物などが困難なため」"
           "「希望する介護施設に入所できないため」"
           "「希望する生活支援サービスが受けることができないため」が含まれ、"
           "居所変更実態調査が把握しようとしている供給不足を理由とする住替えに対応する。",
           [30, 10, 10, 10, 10, 10, 12, 12, 12])

r = lead(ws, 4, "【住み続ける意向】")
r = header(ws, r, ["区分", "広域連合", "東川町", "美瑛町", "東神楽町", "分母",
                   "独居者", "75歳以上", "自力移動の\n手段なし"], height=36)
mnodrive = series("自力移動なし") == 1
for lb, vs in [("住み続けたい", [1]), ("転居も少し考えている", [2]), ("わからない", [3])]:
    A["_t"] = flag_in("taisetsu_q3_1", vs)
    r = body(ws, r, [lb, rate(ALL, "_t")[0], rate(mH, "_t")[0], rate(mB, "_t")[0],
                     rate(mHG, "_t")[0], rate(ALL, "_t")[1], rate(mal, "_t")[0],
                     rate(m75, "_t")[0], rate(mnodrive, "_t")[0]], height=18)

r += 1
r = lead(ws, r, "【地区別の「転居も少し考えている」割合】")
r = header(ws, r, ["町", "地区", "回答者数", "住み続けたい", "転居も少し\n考えている",
                   "わからない", "", "", ""], height=32)
A["_stay"] = flag_in("taisetsu_q3_1", [1])
A["_move"] = flag_in("taisetsu_q3_1", [2])
A["_unk"] = flag_in("taisetsu_q3_1", [3])
for tname, aname, mask in rows:
    r = body(ws, r, [tname, aname, int(mask.sum()), rate(mask, "_stay")[0],
                     rate(mask, "_move")[0], rate(mask, "_unk")[0], "", "", ""],
             height=18, bold=(aname in ("全体", "町計")))

r += 1
r = lead(ws, r, "【転居を考える理由（転居も少し考えていると回答した者が分母）】")
r = header(ws, r, ["理由", "割合", "該当数", "分母", "計画上の意味", "", "", "", ""], height=32)
mmove = series("_move") == 1
Q32 = [("taisetsu_q3_2_s1", "子などと一緒に暮らすため", "家族との同居（供給とは無関係）"),
       ("taisetsu_q3_2_s2", "今住んでいる地域では通院や買い物などが困難なため",
        "移動支援・訪問通所困難地域"),
       ("taisetsu_q3_2_s3", "希望する介護施設に入所できないため", "施設整備（H11の候補）"),
       ("taisetsu_q3_2_s4", "希望する生活支援サービスが受けることができないため",
        "サービス供給不足（H11の候補）"),
       ("taisetsu_q3_2_s5", "その他", "―")]
for key, nm, mean in Q32:
    if key not in A.columns:
        continue
    s = pd.to_numeric(A[key], errors="coerce")[mmove]
    n = int(s.notna().sum())
    v = round(100 * s.sum() / n, 1) if n else None
    fill = {1: NG_O} if "希望する" in nm or "困難" in nm else None
    r = body(ws, r, [nm, v, int(s.sum()) if n else 0, n, mean, "", "", "", ""],
             fill, height=18)
SUP3 = [c for c in ["taisetsu_q3_2_s2", "taisetsu_q3_2_s3", "taisetsu_q3_2_s4"]
        if c in A.columns]
A["_supply"] = any_of(SUP3)                       # 分母＝転居理由に回答した者
# 分母を住み続ける意向の設問に回答した者（4,446票）に広げた系列
A["_supply_all"] = (series("_supply") == 1).astype(float).where(num("taisetsu_q3_1").notna())
s = series("_supply")[mmove]
n = int(s.notna().sum())
r = body(ws, r, ["【再掲】供給・アクセスを理由とする転居意向（2〜4のいずれか）",
                 round(100 * s.sum() / n, 1) if n else None, int(s.sum()) if n else 0, n,
                 "分母＝転居理由の設問に回答した者", "", "", "", ""],
         {1: MID_B}, height=18, bold=True)
sa = series("_supply_all")
na = int(sa.notna().sum())
r = body(ws, r, ["【再掲】回答者全体に対する割合",
                 round(100 * sa.sum() / na, 1) if na else None, int(sa.sum()) if na else 0,
                 na, "H11（供給不足を理由とする住替え割合）の候補", "", "", "", ""],
         {1: OK_G}, height=18, bold=True)

note(ws, r + 1,
     "注1）単位は％。転居理由は複数回答のため合計は100％を超える。"
     "注2）転居を考える理由のうち、「通院や買い物などが困難」"
     "「希望する介護施設に入所できない」「希望する生活支援サービスが受けることができない」の"
     "3つは、いずれもサービス供給又はアクセスの不足を理由とするものであり、"
     "居所変更実態調査が把握しようとしている供給不足を理由とする住替えに対応する。"
     "注3）代表KPI H11の基準値としては、転居意向がある者を分母とする割合と、"
     "回答者全体を分母とする割合の両方を示した。"
     "計画では母集団の解釈が明確な後者を用いることを提案する。"
     "注4）本設問は意向であり、実際の住替えの実績ではない。"
     "実績の把握は居所変更実態調査（第10期分）による。"
     "注5）地区別では美馬牛小学校区の「転居も少し考えている」が23.4％と高い。"
     "同地区は自分で運転する者の割合が83.8％と高く、徒歩22.3％・タクシー2.0％と低いことから、"
     "運転できなくなった後の生活の見通しが転居意向に表れている可能性がある。"
     "志比内小学校区（21票）の値は標本が小さく、単独では判断材料としない。", 9)

# ============================================================ 10 介護の担い手
ws = sheet("10_介護の担い手と介護の意向", "介護の担い手と介護に関する意向（大雪独自設問）",
           "【大雪－問4】日常的に支障が生じた場合に世話をしてくれる家族等の有無、"
           "介護を受けたい場所、人生会議（ACP）及び死後事務委任契約の認知度を集計する。"
           "在宅介護実態調査が把握しようとしている家族の介護力に対応する。",
           [30, 10, 10, 10, 10, 10, 12, 12, 12])

r = lead(ws, 4, "【世話をしてくれる家族等（複数回答）】")
r = header(ws, r, ["区分", "広域連合", "東川町", "美瑛町", "東神楽町", "分母",
                   "独居者", "75歳以上", "85歳以上"], height=32)
m85 = A["年齢階級"] == "85歳以上"
Q41 = [("taisetsu_q4_1_s1", "配偶者"), ("taisetsu_q4_1_s2", "子"),
       ("taisetsu_q4_1_s3", "子の配偶者"), ("taisetsu_q4_1_s4", "孫"),
       ("taisetsu_q4_1_s5", "兄弟・姉妹"), ("taisetsu_q4_1_s6", "近所の人"),
       ("taisetsu_q4_1_s7", "友人"), ("taisetsu_q4_1_s8", "その他の人"),
       ("taisetsu_q4_1_s9", "そのような人はいない")]
for key, nm in Q41:
    if key not in A.columns:
        continue
    A["_t"] = pd.to_numeric(A[key], errors="coerce")
    fill = {1: NG_O} if nm == "そのような人はいない" else None
    r = body(ws, r, [nm, rate(ALL, "_t")[0], rate(mH, "_t")[0], rate(mB, "_t")[0],
                     rate(mHG, "_t")[0], rate(ALL, "_t")[1], rate(mal, "_t")[0],
                     rate(m75, "_t")[0], rate(m85, "_t")[0]], fill, height=18)

r += 1
r = lead(ws, r, "【介護が必要になった場合に介護を受けたい場所】")
r = header(ws, r, ["区分", "広域連合", "東川町", "美瑛町", "東神楽町", "分母",
                   "独居者", "75歳以上", "計画上の意味"], height=32)
Q42 = [("なるべく家族のみで、自宅で介護してほしい", [1], "家族介護力への依存"),
       ("介護保険のサービスや福祉サービスを利用しながら、自宅で介護してほしい", [2],
        "在宅サービスの需要"),
       ("老人ホームなどの施設に入所したい", [3], "施設サービスの需要"),
       ("その他", [4], "―"), ("わからない", [5], "意思決定支援・ACP")]
for nm, vs, mean in Q42:
    A["_t"] = flag_in("taisetsu_q4_2", vs)
    r = body(ws, r, [nm, rate(ALL, "_t")[0], rate(mH, "_t")[0], rate(mB, "_t")[0],
                     rate(mHG, "_t")[0], rate(ALL, "_t")[1], rate(mal, "_t")[0],
                     rate(m75, "_t")[0], mean], height=30)
A["_home"] = flag_in("taisetsu_q4_2", [1, 2])
r = body(ws, r, ["【再掲】自宅での介護を希望（1・2の計）", rate(ALL, "_home")[0],
                 rate(mH, "_home")[0], rate(mB, "_home")[0], rate(mHG, "_home")[0],
                 rate(ALL, "_home")[1], rate(mal, "_home")[0], rate(m75, "_home")[0],
                 "在宅サービスの見込量（第6章第2節）"], {1: OK_G}, height=18, bold=True)

r += 1
r = lead(ws, r, "【意思決定支援に関する認知度】")
r = header(ws, r, ["区分", "既に実施\nしている", "知っている", "言葉だけは\n知っている",
                   "知らない", "分母", "75歳以上で\n「知らない」", "", ""], height=32)
for key, nm in [("taisetsu_q4_4", "人生会議（ACP）"), ("taisetsu_q4_5", "死後事務委任契約")]:
    if key not in A.columns:
        continue
    vs = []
    for v in [1, 2, 3, 4]:
        A["_t"] = flag_in(key, [v])
        vs.append(rate(ALL, "_t")[0])
    A["_t"] = flag_in(key, [4])
    r = body(ws, r, [nm] + vs + [rate(ALL, "_t")[1], rate(m75, "_t")[0], "", ""], height=18)

r += 1
r = lead(ws, r, "【おひとりになった場合の終のすみか】")
r = header(ws, r, ["区分", "広域連合", "東川町", "美瑛町", "東神楽町", "分母",
                   "独居者", "", ""], height=32)
Q46 = [("現在のまま、自宅に住み続けたい", [1]), ("改築の上、自宅に住み続けたい", [6]),
       ("高齢者用住宅へ引っ越したい", [2]), ("老人ホームへ入居したい", [3]),
       ("子どもの住宅へ引っ越したい", [5]), ("病院に入院したい", [7]),
       ("実家（田舎）に引っ越したい", [4]), ("その他", [8])]
for nm, vs in Q46:
    A["_t"] = flag_in("taisetsu_q4_6", vs)
    r = body(ws, r, [nm, rate(ALL, "_t")[0], rate(mH, "_t")[0], rate(mB, "_t")[0],
                     rate(mHG, "_t")[0], rate(ALL, "_t")[1], rate(mal, "_t")[0], "", ""],
             height=18)

note(ws, r + 1,
     "注1）単位は％。世話をしてくれる家族等は複数回答のため合計は100％を超える。"
     "注2）「そのような人はいない」と回答した者の割合は、独居者及び85歳以上で高い。"
     "身寄りのない高齢者の入居・入所、保証、死後事務等の支援体制"
     "（第5章 基本目標3）の対象規模を示す指標となる。"
     "注3）介護を受けたい場所について「介護保険のサービスや福祉サービスを利用しながら、"
     "自宅で介護してほしい」が最も多いことは、"
     "在宅サービスの見込量（第6章第2節）の前提となる。"
     "一方、当広域連合には定期巡回・随時対応型訪問介護看護、夜間対応型訪問介護、"
     "看護小規模多機能型居宅介護の事業所が制度創設以来存在せず、"
     "重度者の在宅生活を24時間支える体制がない。"
     "注4）人生会議（ACP）を「知らない」と回答した者が多数を占めることは、"
     "第10期における意思決定支援の普及啓発の必要性を示す。", 9)

# ============================================================ 11 KPIのデータ源
ws = sheet("11_KPIのデータ源の再検討", "本分析により算定可能となった代表KPI",
           "第10稿では、事業所実態調査及び在宅介護実態調査を当広域連合が実施していないことから、"
           "H07・H08・H12・H16の4項目を「算定不可」又は「データ源の確保が前提」としていた。"
           "本分析の結果、大雪独自設問により一部の指標が算定できることが判明したため、"
           "データ源の再検討結果を示す。",
           [7, 24, 22, 26, 30, 12, 12, 12])

r = header(ws, 4, ["ID", "指標", "第10稿での扱い", "本分析による代替データ源",
                   "算定方法", "基準値\n（令和7年度）", "分母", "判定"], height=40)


def val_of(name, mask=None):
    v, n = rate(mask if mask is not None else ALL, name)
    return v, n


kpi_rows = []
A["_unsolved_all"] = ((series("解決できず困っている") == 1).astype(float)
                      .where(series("困りごとあり").notna()))
sv = series("解決できず困っている")[mneed]
n_sub = int(sv.notna().sum())
v_sub = round(100 * sv.sum() / n_sub, 1) if n_sub else None
sa2 = series("_unsolved_all")
n_all = int(sa2.notna().sum())
v_all = round(100 * sa2.sum() / n_all, 1) if n_all else None
kpi_rows.append(("H06", "在宅生活継続困難割合", "第10期調査で算定（未受領）",
                 "健康とくらしの調査【大雪－問2】1）2）",
                 "生活動作の困りごとについて「解決できず、困っている」と回答した者を"
                 "分子とする。分母は回答者全体（%d票）とする案と、"
                 "困りごとがある者（%d票）とする案がある。前者では%.1f％、後者では%.1f％となる"
                 % (n_all, n_sub, v_all, v_sub), v_all, n_all, "算定可能"))
sup_all = series("_supply_all")
n_h = int(sup_all.notna().sum())
v_h = round(100 * sup_all.sum() / n_h, 1) if n_h else None
sup_sub = series("_supply")[mmove]
n_h2 = int(sup_sub.notna().sum())
v_h2 = round(100 * sup_sub.sum() / n_h2, 1) if n_h2 else None
kpi_rows.append(("H11", "供給不足を理由とする住替え意向割合", "第10期調査で算定（未受領）",
                 "健康とくらしの調査【大雪－問3】1）2）",
                 "転居を考える理由として「通院や買い物などが困難」「希望する介護施設に"
                 "入所できない」「希望する生活支援サービスが受けることができない」の"
                 "いずれかを選んだ者を分子とする。分母を回答者全体（%d票）とすると%.1f％、"
                 "転居理由に回答した者（%d票）とすると%.1f％となる"
                 % (n_h, v_h, n_h2, v_h2), v_h, n_h, "算定可能（意向）"))
A["_care"] = flag_in("care3fam25", [1])
v_c, n_c = val_of("_care")
kpi_rows.append(("H07", "主介護者の高負担割合", "算定不可（在宅介護実態調査が未実施）",
                 "健康とくらしの調査 サブコア1【問17】4）",
                 "「主に介護をしている」と回答した者÷回答者×100。"
                 "負担感そのものではなく主介護者の割合であるため、指標の定義変更を要する",
                 v_c, n_c, "代替指標として算定可能"))
A["_nohelp"] = flag_in("taisetsu_q4_1_s9", [1])
v_x, n_x = val_of("_nohelp")
kpi_rows.append(("H08", "家族介護者の支援充足割合", "算定不可（在宅介護実態調査が未実施）",
                 "健康とくらしの調査【大雪－問4】1）",
                 "「日常的に支障が生じた場合に世話をしてくれる人がいない」と回答した者÷"
                 "回答者×100。支援の充足ではなく担い手の不在を測る指標に変更する",
                 v_x, n_x, "代替指標として算定可能"))
kpi_rows.append(("H12", "必要サービス未充足率", "算定不可（事業所実態調査が未実施）",
                 "―（本調査では算定できない）",
                 "本調査は利用者側の調査であり、事業所側の受入れ可否を把握できない。"
                 "居宅サービス計画未作成者の抽出等の代理指標による",
                 None, None, "引き続き代理指標を検討"))
kpi_rows.append(("H16", "災害・感染症時の必須サービス継続率", "算定不可（事業所実態調査が未実施）",
                 "―（本調査では算定できない）",
                 "事業所のBCP策定状況及び発動実績を要する。"
                 "指定台帳及び集団指導の機会を通じた把握による",
                 None, None, "引き続き代替手段を検討"))
A["_frail"] = pd.to_numeric(A["frail2gp_25"], errors="coerce")
v_f, n_f = val_of("_frail")
kpi_rows.append(("H04", "フレイル該当割合", "19.1％（R7）で確定済み",
                 "健康とくらしの調査（基本チェックリスト8項目以上）",
                 "基本チェックリスト該当項目数8以上の者÷判定可能な有効回答者×100",
                 v_f, n_f, "確定済み"))
v_s, n_s = val_of("社会参加")
kpi_rows.append(("H05", "社会参加率", "8.8％（通いの場）で確定済み",
                 "健康とくらしの調査【問5】1）",
                 "6区分（ボランティア・スポーツの会・趣味の会・学習教養サークル・"
                 "通いの場・特技を伝える活動）のいずれかに月1回以上参加した者÷回答者×100",
                 v_s, n_s, "合成指標も選択可能"))
A["_need_nouse"] = flag_in("adl3ra25", [2])
v_nn, n_nn = val_of("_need_nouse")
kpi_rows.append(("―", "介護・介助が必要だが受けていない者の割合（新規提案）", "第10稿では設定なし",
                 "健康とくらしの調査【問1】2）",
                 "「何らかの介護・介助が必要だが、現在は受けていない」と回答した者÷"
                 "回答者×100。未利用認定者504人（認定者の25.7％）の論点に対応する",
                 v_nn, n_nn, "新規に設定を提案"))

for row in kpi_rows:
    fill = {8: OK_G} if "算定可能" in str(row[7]) or row[7] == "確定済み" else {8: NG_O}
    r = body(ws, r, list(row), fill, height=64, align={1: "center", 6: "center", 8: "center"})

note(ws, r + 1,
     "注1）単位は％。基準値は令和7年度 健康とくらしの調査による。"
     "注2）H06・H11については、当初は在宅生活改善調査及び居所変更実態調査を"
     "データ源としていたが、本調査の大雪独自設問により基準値を設定できる。"
     "第10期分の各調査を受領した後に、両者の値を突合して整合を確認する。"
     "注3）H07・H08については、指標の定義を変更したうえで代替データ源を用いる案である。"
     "定義変更の可否は発注者の決定による（確認事項No.4）。"
     "注4）H12・H16は事業所側の情報を要するため、本調査では算定できない。"
     "注5）「介護・介助が必要だが受けていない者の割合」は、"
     "認定を受けながらサービスを利用していない504人（認定者の25.7％）という"
     "第10期の中心論点に対応する指標として新規に設定することを提案する。"
     "ただし本調査の対象は認定の有無を問わない65歳以上であり、"
     "認定者の未利用率とは母集団が異なることに留意する。", 8)

# ============================================================ 12 主要所見
ws = sheet("12_主要所見", "クロス集計・分析の主要所見",
           "本分析から得られた所見と、計画素案への反映方針を示す。"
           "所見は根拠となる数値とシートを併記する。",
           [4, 26, 44, 30, 20, 14])

r = header(ws, 4, ["No.", "所見", "根拠となる数値", "計画への反映", "反映先", "シート"])


def g(name, mask=None):
    v, n = rate(mask if mask is not None else ALL, name)
    return "―" if v is None else "%.1f％（n=%d）" % (v, n)


FIND = [
    ("小学校区別の分析が可能",
     "個票データの分析地域1により、東神楽町3・東川町4・美瑛町5の計12小学校区別の集計ができる。"
     "美瑛町では美馬牛小学校区（149票）、美沢小学校区（57票）、明徳小学校区（49票）、"
     "美瑛東小学校区（533票）が独立して取得できる。",
     "日常生活圏域別の分析を、見える化システムのデータを待たずに着手できる。"
     "小学校区と第9期の6圏域の対応関係の確認を要する。",
     "第1章第7節\n第2章第4節", "01・02"),
    ("美瑛町の指標の差には年齢構成の影響がある",
     "美瑛町のフレイル該当割合は21.6％で東神楽町17.0％より4.6ポイント高いが、"
     "美瑛町は回答者に占める75歳以上の割合が高い。年齢調整後の値を02シートに示した。",
     "町間比較は粗率ではなく年齢調整後の値によることとし、"
     "本文の記述を「高齢化の進行度の差を含む」と明示する。", "第2章第4節", "02"),
    ("除雪が生活動作の困りごとの最上位",
     "生活動作の困りごとのうち除雪が最も多い（%s）。"
     "美瑛町の町道除雪率は52.9％であり、3町はいずれも豪雪地帯又は特別豪雪地帯である。"
     % g("taisetsu_q2_1_s12"),
     "除雪は介護保険の給付対象外であり、生活支援体制整備事業及び3町の施策で対応する。"
     "第5章 基本目標2の主な事業に位置づける。", "第5章 基本目標2", "08"),
    ("困りごとを解決できない層が存在する",
     "生活動作の困りごとがある者のうち「解決できず、困っている」と回答した者は%s、"
     "回答者全体に対する割合は%s である。"
     % (g("解決できず困っている", mneed), g("_unsolved_all")),
     "代表KPI H06（在宅生活継続困難割合）の基準値として設定できる。"
     "在宅生活改善調査（第10期分）の受領後に突合する。", "第4章第3節\n第5章 基本目標2", "08・11"),
    ("供給・アクセスを理由とする転居意向が把握できる",
     "転居を考える理由として、通院・買い物の困難、希望する介護施設に入所できない、"
     "希望する生活支援サービスが受けられないのいずれかを挙げた者は、"
     "住み続ける意向の設問に回答した者に対して%s である。"
     % g("_supply_all"),
     "代表KPI H11（供給不足を理由とする住替え）の基準値として設定できる。"
     "居所変更実態調査（第10期分）の実績と突合する。", "第4章第3節\n第6章第2節", "09・11"),
    ("介護・介助が必要だが受けていない層がいる",
     "「何らかの介護・介助が必要だが、現在は受けていない」と回答した者は%s である。"
     % g("_need_nouse"),
     "認定を受けながらサービスを利用していない504人（認定者の25.7％）の論点に対応する"
     "指標として、新規に設定することを提案する。", "第2章第2節3\n第4章第3節", "11"),
    ("世話をしてくれる人がいない層がいる",
     "日常的に支障が生じた場合に世話をしてくれる人が「いない」と回答した者は%s であり、"
     "独居者では%s と高い。"
     % (g("世話をする人なし"), g("世話をする人なし", mal)),
     "身寄りのない高齢者の入居・入所、保証、死後事務等の支援体制の"
     "対象規模を示す指標として用いる。", "第5章 基本目標3", "06・10"),
    ("在宅での介護を希望する者が多数",
     "介護が必要になった場合に自宅での介護を希望する者は%s である。" % g("_home"),
     "在宅サービスの見込量の前提となる。"
     "一方、24時間対応の3サービスが区域内に存在しないことと合わせて論じる。",
     "第6章第2節・第4節", "10"),
    ("社会参加とリスク指標に一貫した関連",
     "社会参加あり群はフレイル・閉じこもり・うつ・IADL低下の割合がいずれも低く、"
     "年齢階級で層化しても同じ方向の差がみられる。",
     "介護予防・社会参加の施策（基本目標1）の根拠として用いる。"
     "横断調査であり因果関係を示すものではないことを注記する。", "第5章 基本目標1", "05"),
    ("経済状況による差が最も大きい",
     "フレイル該当割合は暮らし向きが「大変苦しい」層で42.7％、"
     "「大変ゆとりがある」層で6.8％と6.3倍の差がある。"
     "うつ割合（49.0％対12.6％）、生活動作の困りごと（60.9％対20.0％）も同様であり、"
     "本分析で確認した属性のうち経済状況による差が最も大きい。",
     "低所得層への配慮を保険料段階の設定、相談支援及び介護予防の"
     "対象者の把握に反映する。", "第2章第2節3\n第5章 基本目標1\n第6章第6節", "07"),
    ("独居者の社会参加は高い",
     "独居者の社会参加あり割合は41.5％で、その他の世帯30.5％より11.0ポイント高く、"
     "通いの場参加割合も14.8％対7.1％と高い。"
     "一方、独居者は孤食16.9％、生活動作の困りごと57.7％、"
     "世話をしてくれる人がいない32.4％といずれも高い。",
     "独居であることが直ちに社会的孤立を意味しないため、"
     "施策の対象は世帯類型ではなく、困りごとの有無及び支援者の有無で捉える。"
     "通いの場は独居者の参加が多く、既存の資源が機能している。",
     "第5章 基本目標1・2", "06"),
    ("自力移動の手段がない層の分布",
     "自分での運転・自転車・バイクのいずれも利用していない者の割合を町別・地区別に算定した。"
     "路線バス及びコミュニティバスの利用割合は低い。",
     "訪問・通所困難地域の判定において、事業所の通常の事業の実施地域と併せて用いる。",
     "第1章第7節", "03"),
    ("人生会議（ACP）の認知度が低い",
     "人生会議（ACP）を「知らない」と回答した者が多数を占める。",
     "意思決定支援の普及啓発を基本目標5の主な事業に位置づける。", "第5章 基本目標5", "10"),
]
for i, (t, ev, act, dst, sh) in enumerate(FIND, start=1):
    r = body(ws, r, [i, t, ev, act, dst, sh], height=64, align={1: "center", 6: "center"})

note(ws, r + 1,
     "注1）本シートの所見は、健康とくらしの調査（令和7年度）の個票データの分析によるものである。"
     "他の3調査（在宅生活改善調査、居所変更実態調査、介護人材実態調査）の第10期分は未受領であり、"
     "受領後に本分析の結果と突合して整合を確認する。"
     "注2）所見2〜12は計画素案（第10稿）への反映を要する。"
     "反映は図表レイアウトの確認結果と併せて次稿で行う。"
     "注3）業務工程管理表の01シート（3）実施済み調査結果の集計・分析について、"
     "本分析の完了により作業項目①⑤⑥が完了する。", 6)

del wb["Sheet"]
wb.save(OUT)
print("saved:", OUT, "sheets=%d" % len(wb.sheetnames))
for s in wb.sheetnames:
    print("  -", s, wb[s].max_row, "rows")
