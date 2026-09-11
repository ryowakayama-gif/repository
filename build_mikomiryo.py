# -*- coding: utf-8 -*-
"""大雪地区広域連合 第10期介護保険事業計画　サービス見込量の算定.

令和8年9月10日のご指示
  「今までの整理を踏まえて、見込み量算定を進めて下さい。
    各初期設定で基準年度をどこでとるか、
    給付費や保険料試算として妥当性があるパターンを特定して下さい。
    特にエラーチェック時にエラーがでないよう留意願います」

本表は、基準年度をどこに置くかによって見込量・給付費・保険料が
どこまで動くかを示し、妥当性のあるパターンを特定するものである。
受託者の分析であり確定値ではない。

━━ 判明したこと ━━

1 数量（利用者数）の基準年度に使えるのは令和7年度だけである。
  令和8年度は約4か月分の実績しかなく、
  計画作成支援ツールの「月平均利用者数」は年度累計を12で除しているため
  令和7年度の約0.35倍の値になっている。そのまま基準年度に用いられない。
  令和6年度は完結しているが、令和7年度との間に
  予防訪問看護の計上の変化がある。

2 単価（1人1月あたり給付費・利用回数）は令和8年度でも用いられる。
  月数に中立な指標であり、令和7年度比0.99〜1.08の範囲にある。

3 出所により水準が異なる。
  見える化システムの総括表は保険者（広域連合）単位、
  計画作成支援ツールは3町合計であり、
  同じ令和7年度の利用者数でも +1.4％〜+243％の開きがある。
  給付費でも令和7年度は3町合計が広域連合単位を4.0％上回る（確認事項No.87）。
  保険料は保険者単位で算定するため、総括表を基礎とするのが筋である。

4 見える化システムの総括表は内的整合が取れている。
  利用者数×1人あたり給付費＝給付費が全行で成立し、
  給付費の合計は介護保険事業状況報告（年報）と円単位で一致する。

5 見える化システムのD32（受給率）は小数第1位までしか公表されない。
  分母が第1号被保険者数（令和7年度9,082人）であるため、
  0.1％＝約9.1人／月の刻みになる。
  実数（受給者数）が併載されるのは訪問介護・訪問看護・通所介護の3種別のみで、
  他の13種別は丸めた率しか得られない。
  月5.2人の訪問入浴介護のような小口のサービスは0人になる。

シート構成
  00_この算定について
  01_基準年度の判定
  02_出所の比較
  03_算定パターンと妥当性
  04_サービス見込量
  05_給付費と保険料への影響
  06_自己点検
  07_確認事項

出力
  output/第10期計画_サービス見込量の算定_基準年度とパターンの比較.xlsx

自己点検で1件でも不適合があると終了コード1で終わる。
"""

import io
import os
import runpy
import sys

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import data_kessan_r6 as KS
import data_kyufu_jisseki as KJ
import data_mieru_soukatu as M
import data_nenpo as N
import data_shien_tool as T
import repo_paths as RP

ODIR = RP.ROOT + "/output"
OUT = os.path.join(ODIR,
                   "第10期計画_サービス見込量の算定_基準年度とパターンの比較.xlsx")

KIJUNBI = "令和8年9月10日"
TOWNS3 = ["東川町", "美瑛町", "東神楽町"]

FONT = "游ゴシック"
NAVY, HEAD = "1F3864", "4472C4"
IN_Y, OK_G, NG_O, MID_B, GRAY = "FFF2CC", "E2EFDA", "FCE4D6", "DEEBF7", "F2F2F2"
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = Workbook()
wb.remove(wb.active)

CHECKS = []          # 自己点検の結果（06シート）


def chk(no, naiyo, shiki, kekka, ok):
    CHECKS.append((no, naiyo, shiki, kekka, "適合" if ok else "不適合"))
    return ok


# ============================================================ 前段の推計
_buf, _old = io.StringIO(), sys.stdout
sys.stdout = _buf
try:
    _G = runpy.run_path(RP.ROOT + "/build_projection.py")       # 第1段階（人口・認定者数）
    _P = runpy.run_path(RP.ROOT + "/build_projection3.py")      # 第3段階（給付費・保険料）
finally:
    sys.stdout = _old

total, SC = _G["total"], _G["SC"]
gaku = _P["gaku"]                    # 保険料基準額（算定上の月額）を返す
KYUFU_GENKO = _P["KYUFU"]            # 現行の給付費（見える化の自然体推計×案C補正）

BASE_Y, YEARS_N = "2025", ["2027", "2028", "2029"]
YEARS = ["令和9年度", "令和10年度", "令和11年度"]
BASE_NINTEI = total(BASE_Y, 2)                       # 令和7年度の認定者数
NOBI = {sc: [total(y, sc) / BASE_NINTEI for y in YEARS_N] for sc in (1, 2, 3)}

# ============================================================ 総括表
# 見える化システムの総括表（保険者単位）。
#   詳細（１）利用者数（年間の延べ人数）
#   詳細（４）1人あたり給付費（円。延べ1人あたり）
#   詳細（５）給付費（円）
D1 = {l: d for l, u, d in M.SOUKATU["総括表詳細（１）"]["行"] if u == "（人）"}
D4 = {l: d for l, u, d in M.SOUKATU["総括表詳細（４）"]["行"] if u == "（円）"}
D5 = {l: d for l, u, d in M.SOUKATU["総括表詳細（５）"]["行"] if u == "（円）"}
SVC = [l for l in D1 if "小計" not in l]          # 小計行を除いたサービスの行


def jisseki(lab, year, key="実績値"):
    return D1[lab][year][key]


def tanka(lab, year, key="実績値"):
    return D4[lab][year][key]


def tsumiage(year):
    """総括表の実績値による総給付費（円）。利用者数×1人あたり給付費の積上げ。"""
    s = 0
    for l in SVC:
        n, p = jisseki(l, year), tanka(l, year)
        if n and p:
            s += n * p
    return s


R6_TSUMI, R7_TSUMI = tsumiage("R6"), tsumiage("R7")

# ============================================================ 計画作成支援ツール
def tool_tsuki(year, name):
    """計画作成支援ツールの月平均利用者数（3町合計）。要介護度別の合計。"""
    for r in KJ.KYUFU[("大雪", year)]["月平均利用者数"]:
        if r[0] == name:
            return sum(v for v in r[1:] if isinstance(v, (int, float)))
    return None


# 総括表の行 → 計画作成支援ツールの行（介護と予防を合わせる）
TOOLMAP = {
    "施設サービス 介護老人福祉施設": ["介護老人福祉施設サービス"],
    "施設サービス 地域密着型介護老人福祉施設入所者生活介護":
        ["地域密着型介護老人福祉施設"],
    "施設サービス 介護老人保健施設": ["介護老人保健施設サービス"],
    "施設サービス 介護医療院": ["介護医療院サービス"],
    "居住系サービス 特定施設入居者生活介護":
        ["特定施設生活介護（短期以外）", "予防特定施設入居者生活介護"],
    "居住系サービス 認知症対応型共同生活介護":
        ["認知症対応型共同生活", "予防認知症型共同生活"],
    "在宅サービス 訪問介護": ["訪問介護"],
    "在宅サービス 訪問入浴介護": ["訪問入浴介護"],
    "在宅サービス 訪問看護": ["訪問看護", "予防訪問看護"],
    "在宅サービス 訪問リハビリテーション":
        ["訪問リハビリテーション", "予防訪問リハビリテーション"],
    "在宅サービス 居宅療養管理指導": ["居宅療養管理指導", "予防居宅療養管理指導"],
    "在宅サービス 通所介護": ["通所介護"],
    "在宅サービス 地域密着型通所介護": ["地域密着型通所介護"],
    "在宅サービス 通所リハビリテーション":
        ["通所リハビリテーション", "予防通所リハビリテーション"],
    "在宅サービス 短期入所生活介護": ["短期入所生活介護", "予防短期入所生活介護"],
    "在宅サービス 短期入所療養介護（老健）":
        ["短期入所療養介護（老健施設）", "予防短期入所療養介護（老健）"],
    "在宅サービス 福祉用具貸与": ["福祉用具貸与", "予防福祉用具貸与"],
    "在宅サービス 特定福祉用具販売": ["特定福祉用具購入費"],
    "在宅サービス 住宅改修": ["住宅改修費"],
    "在宅サービス 定期巡回・随時対応型訪問介護看護":
        ["定期巡回・随時対応型訪問介護"],
    "在宅サービス 認知症対応型通所介護": ["認知症対応型通所介護"],
    "在宅サービス 小規模多機能型居宅介護":
        ["小規模多機能型居宅（短期以外）", "予防小規模多機能型（短期以外）"],
    "在宅サービス 介護予防支援・居宅介護支援": ["居宅介護支援", "予防支援"],
}


def tool_gokei(year, lab):
    ns = TOOLMAP.get(lab)
    if not ns:
        return None
    v = [tool_tsuki(year, n) for n in ns]
    return sum(x for x in v if x is not None) if any(
        x is not None for x in v) else None


def short(lab):
    return lab.split(" ", 1)[1] if " " in lab else lab


# ============================================================ 算定パターン
# 採用パターンは P3（令和7年度基準・総括表・認定者数シナリオ②）とする。
SAIYO = "P3"

PATTERNS = [
    dict(no="P1", kijun="令和7年度（見える化の自然体推計の基礎）",
         demae="見える化システムの自然体推計", nobi="社人研を案C（総合戦略）で置換え",
         kyufu=list(KYUFU_GENKO), hantei="採り得る",
         riyu="現行の第3段階が用いている置き方。"
              "利用率及び単価の趨勢が自然体推計に織り込まれているため、"
              "認定者数の伸び（案A 令和9年度＋0.6％）を超えて"
              "給付費が伸びる（同＋4.9％）。"
              "趨勢の中身が当方で分解できない点が難である。"),
    dict(no="P2", kijun="同上", demae="見える化システムの自然体推計",
         nobi="置換えを行わない（社人研のまま）",
         kyufu=[3056.7, 3090.2, 3120.5], hantei="採らない",
         riyu="人口の基礎を総合戦略に置くという令和8年8月31日のご指示"
              "（将来推計 第1段階の案C）と合わない。"),
    dict(no="P3", kijun="令和7年度（12か月・完結）",
         demae="見える化システムの総括表（保険者単位）",
         nobi="認定者数シナリオ②トレンド継続（採用シナリオ）",
         kyufu=None, sc=2, base="R7", hantei="妥当",
         riyu="基準年度が完結しており、給付費の合計が年報と円単位で一致する。"
              "保険者単位であるため保険料算定の単位と合う。"
              "利用率と単価を令和7年度で固定し、認定者数の伸びのみを反映するため、"
              "伸びの中身が説明できる。"),
    dict(no="P4", kijun="令和7年度（12か月・完結）",
         demae="見える化システムの総括表（保険者単位）",
         nobi="認定者数シナリオ①現状固定",
         kyufu=None, sc=1, base="R7", hantei="感度",
         riyu="認定率を令和7年度で固定した場合の上側。感度分析に用いる。"),
    dict(no="P5", kijun="令和7年度（12か月・完結）",
         demae="見える化システムの総括表（保険者単位）",
         nobi="認定者数シナリオ③令和元年水準へ回帰",
         kyufu=None, sc=3, base="R7", hantei="感度",
         riyu="認定率が令和元年の水準へ戻る場合の上側。感度分析に用いる。"),
    dict(no="P6", kijun="令和6年度（12か月・完結）",
         demae="見える化システムの総括表（保険者単位）",
         nobi="認定者数シナリオ②トレンド継続",
         kyufu=None, sc=2, base="R6", hantei="採らない",
         riyu="令和7年度の実績が得られている以上、"
              "1年古い年度を基準年度に置く理由がない。"
              "また令和6年度と令和7年度の間に予防訪問看護の計上の変化がある。"),
    dict(no="P7", kijun="令和7年度", demae="計画作成支援ツール（3町合計）",
         nobi="認定者数シナリオ②トレンド継続",
         kyufu=None, sc=2, base="R7", tool=True, hantei="採らない",
         riyu="3町合計は町をまたぐ計上の重複を含み、"
              "令和7年度の給付費で広域連合単位を4.0％上回る（確認事項No.87）。"
              "保険料は保険者単位で算定するため、算定の基礎には用いられない。"
              "町別の実績を掲げる用途には用いる。"),
    dict(no="P8", kijun="令和8年度", demae="計画作成支援ツール（3町合計）",
         nobi="―", kyufu=None, hantei="用いられない",
         riyu="令和8年度は約4か月分の実績しかない。"
              "月平均利用者数は年度累計を12で除しているため"
              "令和7年度の約0.35倍の値になっており、基準年度に用いられない。"
              "単価（1人1月あたり給付費・利用回数）は月数に中立であり用いられる。"),
]


def pattern_kyufu(p):
    """パターンの3か年の給付費（百万円）を返す。算定できないものはNone。"""
    if p.get("kyufu") is not None:
        return p["kyufu"]
    if p["no"] == "P8":
        return None
    f = NOBI[p["sc"]]
    if p.get("tool"):
        # 3町合計の給付費（計画作成支援ツール）を基準にした場合。
        # 広域連合の行ではなく、3町の行を実際に足し上げる。
        b = sum(KJ.KYUFU_TOTAL[(t, p["base"])][-1] for t in TOWNS3) / 1e6
    else:
        b = (R7_TSUMI if p["base"] == "R7" else R6_TSUMI) / 1e6
    return [b * x for x in f]


for _p in PATTERNS:
    _p["kyufu3"] = pattern_kyufu(_p)
    _p["gaku"] = (gaku(kyufu_oku=sum(_p["kyufu3"]) * 1e6)
                  if _p["kyufu3"] else None)

BASIC = [p for p in PATTERNS if p["no"] == SAIYO][0]
GENKO = [p for p in PATTERNS if p["no"] == "P1"][0]


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
    ws.row_dimensions[2].height = 62
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


def body(ws, row, vals, fills=None, height=26, align=None, bold=False,
         fmt=None):
    for i, v in enumerate(vals, start=1):
        c = ws.cell(row=row, column=i, value=v)
        c.font = Font(name=FONT, size=9, bold=bold)
        c.border = BORDER
        ha = (align or {}).get(i, "left" if isinstance(v, str) else "right")
        c.alignment = Alignment(wrap_text=True, vertical="top", horizontal=ha)
        if fills and fills.get(i):
            c.fill = PatternFill("solid", fgColor=fills[i])
        if fmt and fmt.get(i):
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


def yen(v):
    return "{:,}".format(int(round(v)))


HAN_F = {"妥当": OK_G, "採り得る": MID_B, "感度": IN_Y,
         "採らない": GRAY, "用いられない": NG_O}

# ============================================================ 00
ws = sheet("00_この算定について", "サービス見込量の算定　基準年度とパターンの比較",
           "基準年度をどこに置くかによって見込量・給付費・保険料がどこまで動くかを示し、"
           "妥当性のあるパターンを特定するものです。受託者の分析であり確定値ではありません。"
           "採用パターンは%s（令和7年度基準・見える化システムの総括表・"
           "認定者数シナリオ②トレンド継続）としています。" % SAIYO,
           [3, 26, 66, 24], freeze="A6")
r = 4
r = lead(ws, r, "1　判明したこと", span=4)
r = header(ws, r, ["", "項目", "内容", "根拠"])
for i, (k, v, b) in enumerate([
    ("数量の基準年度は令和7年度に限られる",
     "令和8年度は約4か月分の実績しかありません。"
     "計画作成支援ツールの「月平均利用者数」は年度累計を12で除しているため、"
     "令和7年度の約0.35倍の値になっています。"
     "令和6年度は完結していますが、令和7年度の実績が得られている以上、"
     "1年古い年度を基準年度に置く理由がありません。",
     "01シート"),
    ("単価は令和8年度でも用いられる",
     "1人1月あたり給付費及び1人1月あたり利用回数は月数に中立な指標であり、"
     "令和8年度の値は令和7年度比0.99〜1.08の範囲にあります。"
     "報酬改定の影響を見るには令和8年度の単価が使えます。",
     "01シート"),
    ("出所により水準が異なる",
     "見える化システムの総括表は保険者（広域連合）単位、"
     "計画作成支援ツールは3町合計です。"
     "同じ令和7年度の利用者数でも、"
     "3町合計が保険者単位を+1.4％から+243％上回ります。"
     "給付費でも令和7年度は3町合計が4.0％上回ります（確認事項No.87）。",
     "02シート"),
    ("総括表は内的整合が取れている",
     "利用者数×1人あたり給付費＝給付費が全行で成立し、"
     "給付費の合計は介護保険事業状況報告（年報）と円単位で一致します。"
     "給付費から保険料までの経路が一本でつながります。",
     "06シート（自己点検）"),
    ("見える化システムのD32は小口のサービスに使えない",
     "受給率は小数第1位までしか公表されません。"
     "分母が第1号被保険者数（令和7年度9,082人）であるため、"
     "0.1％が約9.1人／月の刻みになります。"
     "実数が併載されるのは訪問介護・訪問看護・通所介護の3種別のみで、"
     "月5.2人の訪問入浴介護のような小口のサービスは0人になります。",
     "02シート"),
], start=1):
    r = body(ws, r, [i, k, v, b], height=76)
r += 1

r = lead(ws, r, "2　算定の式", span=4)
r = header(ws, r, ["", "段階", "式", "備考"])
for i, (k, v, b) in enumerate([
    ("見込量（人／月）",
     "令和7年度の延べ利用者数 ÷ 12 × 認定者数の伸び率",
     "伸び率は将来推計 第1段階のシナリオ②による"),
    ("給付費（円）",
     "見込量（人／月）× 12 × 1人あたり給付費（令和7年度実績）",
     "報酬改定率は未公表のため単価は令和7年度で固定。［要協議］"),
    ("標準給付費見込額（A）",
     "総給付費 × 割増率1.05909（令和6年度決算による）",
     "将来推計 第3段階 01シートと同じ"),
    ("保険料収納必要額（J）", "C＋D－E＋F＋G±H－I",
     "F・G・Hはいずれも0"),
    ("算定上の月額基準額", "J ÷ 予定収納率 ÷ 補正後被保険者数 ÷ 12",
     "予定収納率99.0％、補正後被保険者数は将来推計 第3段階による"),
], start=1):
    r = body(ws, r, [i, k, v, b], height=32)
r += 1
r = note(ws, r,
         "注1）保険料は介護保険法第3条及び地方自治法第284条により"
         "保険者（広域連合）を単位として算定します。町別には算定しません。\n"
         "注2）第10期の第1号被保険者負担割合、報酬改定率、"
         "令和7年度末の基金残高、所得段階の政令改正はいずれも未確定です。\n"
         "注3）本表は保険料の採用値を定めるものではありません。"
         "採用値のご判断をお願いする材料です。", span=4, height=64)

# ============================================================ 01
ws = sheet("01_基準年度の判定", "基準年度の判定（各出所の月数の実測）",
           "どの年度が基準年度に使えるかを、出所ごとに実測した月数で判定しています。"
           "月数は、延べの第1号被保険者数を1か月あたりの被保険者数で除して求めました。",
           [30, 14, 14, 14, 16, 34], freeze="A6")
r = 4
r = lead(ws, r, "1　出所ごとの年度の完結の状況", span=6)
r = header(ws, r, ["出所", "令和6年度", "令和7年度", "令和8年度",
                   "基準年度への可否", "判定の根拠"])

# 令和8年度が何か月分かを給付費から実測する
_r7g = KJ.KYUFU_TOTAL[("大雪", "R7")][-1]
_r8g = KJ.KYUFU_TOTAL[("大雪", "R8")][-1]
R8_RATIO = _r8g / _r7g
R8_TSUKI = R8_RATIO * 12

# 月平均利用者数の年度比（数量が年度途中であることの実測）
_qs = []
for lab in SVC:
    a, b = tool_gokei("R7", lab), tool_gokei("R8", lab)
    if a and b and a >= 20:            # 月20人以上の種別で比較する
        _qs.append(b / a)
Q_RATIO = sum(_qs) / len(_qs)

for src, a, b, c, kahi, konkyo in [
    ("見える化システム 総括表詳細（１）〜（５）", "12か月", "12か月",
     "実績値なし（計画値のみ）", "令和7年度が最良",
     "給付費の合計が年報と円単位で一致する（06シート 点検2）"),
    ("見える化システム D32（受給率・受給者数）", "12か月", "11か月", "―",
     "率は使えるが小口は不可",
     "延べ第1号被保険者数 令和6年度109,848÷9,154＝12.00、"
     "令和7年度99,902÷9,082＝11.00。"
     "率は分子分母が同じ月数のため月数に中立"),
    ("計画作成支援ツール 月平均利用者数", "12か月", "12か月",
     "約%.1fか月分" % R8_TSUKI, "令和8年度は不可",
     "令和8年度の月平均利用者数は令和7年度の%.2f倍。"
     "年度累計を12で除しているため月数に比例する" % Q_RATIO),
    ("計画作成支援ツール 1人1月あたり給付費", "12か月", "12か月",
     "約%.1fか月分だが値は使える" % R8_TSUKI, "3か年とも可",
     "月数に中立な指標。令和8年度は令和7年度比0.99〜1.08"),
    ("計画作成支援ツール 1人1月あたり利用回数", "12か月", "12か月",
     "約%.1fか月分だが値は使える" % R8_TSUKI, "3か年とも可", "同上"),
    ("介護保険事業状況報告（年報）", "確定", "確定", "未公表", "令和7年度が最良",
     "総括表詳細（５）の合計と一致する"),
]:
    f = {5: OK_G if "可" in kahi and "不可" not in kahi else
         (NG_O if "不可" in kahi else IN_Y)}
    r = body(ws, r, [src, a, b, c, kahi, konkyo], fills=f, height=46)
r = note(ws, r,
         "注1）令和8年度の給付費は令和7年度の%.1f％であり、"
         "月数に換算するとおおむね%.1fか月分に当たります。"
         "月数そのものは公表されていないため［要確認］です。\n"
         "注2）令和6年度と令和8年度は3町合計と広域連合単位が完全に一致し、"
         "令和7年度のみ4.0％の差があります（確認事項No.87）。"
         % (R8_RATIO * 100, R8_TSUKI), span=6, height=48)
r += 1

r = lead(ws, r, "2　令和8年度が基準年度に使えないことの実測", span=6)
r = header(ws, r, ["サービス", "R6（人／月）", "R7（人／月）", "R8（人／月）",
                   "R8÷R7", "備考"])
for lab in SVC:
    a, b, c = (tool_gokei(y, lab) for y in ("R6", "R7", "R8"))
    if not (a and b) or b < 20:
        continue
    r = body(ws, r, [short(lab), a, b, c if c is not None else "―",
                     round(c / b, 3) if (c and b) else "―",
                     "年度途中のため約3分の1"], fills={4: NG_O}, height=22)
r = body(ws, r, ["平均", "―", "―", "―", round(Q_RATIO, 3),
                 "月数に換算して約%.1fか月分" % (Q_RATIO * 12)],
         fills={i: MID_B for i in range(1, 7)}, height=22, bold=True)
r = note(ws, r,
         "注）月20人以上の種別のみを掲げています。"
         "月20人未満の種別は実人数が小さく、比が安定しません。\n"
         "注）計画作成支援ツールの「月平均利用者数」は年度累計を12で除した値です。"
         "令和8年度は年度の途中であるため、"
         "月平均としては約3分の1の値になっています。"
         "年度換算するには12÷月数を乗じることになりますが、"
         "月数が公表されていないため当方では換算しません。", span=6, height=62)

# ============================================================ 02
ws = sheet("02_出所の比較", "出所の比較（保険者単位と3町合計）",
           "同じ令和7年度の利用者数でも、見える化システムの総括表（保険者単位）と"
           "計画作成支援ツール（3町合計）で水準が異なります。"
           "保険料は保険者単位で算定するため、算定の基礎には総括表を用います。",
           [30, 13, 13, 13, 13, 11, 30], freeze="A6")
r = 4
r = lead(ws, r, "1　令和7年度の利用者数（人／月）", span=7)
r = header(ws, r, ["サービス", "総括表 R6", "ツール R6", "総括表 R7",
                   "ツール R7", "R7の差", "備考"])
_diffs = []
for lab in SVC:
    m6 = jisseki(lab, "R6") / 12 if jisseki(lab, "R6") is not None else None
    m7 = jisseki(lab, "R7") / 12 if jisseki(lab, "R7") is not None else None
    t6, t7 = tool_gokei("R6", lab), tool_gokei("R7", lab)
    if m7 is None or t7 is None or m7 == 0:
        d = "―"
    else:
        dv = (t7 - m7) / m7 * 100
        d = "%+.1f％" % dv
        _diffs.append(dv)
    bk = ""
    if lab.endswith("訪問看護"):
        bk = ("ツールの予防訪問看護が令和7年度0人。"
              "令和6年度は18人であり、計上の変化とみられる")
    elif lab.endswith("居宅療養管理指導"):
        bk = "令和6年度も+43％で一貫している。町をまたぐ計上とみられる"
    r = body(ws, r, [short(lab),
                     round(m6, 1) if m6 is not None else "―",
                     t6 if t6 is not None else "―",
                     round(m7, 1) if m7 is not None else "―",
                     t7 if t7 is not None else "―", d, bk], height=24)
r = note(ws, r,
         "注1）総括表は年間の延べ利用者数を12で除したものです。"
         "ツールは月平均利用者数の要介護度別の合計です。"
         "ツールは介護と予防を合わせています。\n"
         "注2）差が大きい種別ほど実人数が小さく、"
         "町をまたぐ計上の重複と端数の影響が相対的に大きくなります。\n"
         "注3）訪問看護のみ差が負になります。"
         "計画作成支援ツールの予防訪問看護が令和7年度0人であることによるものとみられ、"
         "［要確認］です。", span=7, height=62)
r += 1

r = lead(ws, r, "2　給付費の出所と関係", span=7)
r = header(ws, r, ["出所", "令和6年度（円）", "令和7年度（円）", "単位",
                   "", "", "内容"])
_k2 = KS.SAISHUTSU["2 保険給付費"][1]
_sogo_kessan = KS.SAISHUTSU[
    "4 地域支援事業費／1 介護予防・生活支援サービス事業費"][1]
for src, v6, v7, tan, naiyo in [
    ("介護保険事業状況報告（年報）の給付費", N.KYUFU["給付費　合計"]["R6"],
     N.KYUFU["給付費　合計"]["R7"], "広域連合",
     "総給付費。総括表詳細（５）の合計と一致する"),
    ("見える化システム 総括表詳細（５）の合計", R6_TSUMI, R7_TSUMI, "広域連合",
     "利用者数×1人あたり給付費の積上げ"),
    ("計画作成支援ツールの給付費（広域連合の行）",
     KJ.KYUFU_TOTAL[("大雪", "R6")][-1], KJ.KYUFU_TOTAL[("大雪", "R7")][-1],
     "広域連合",
     "令和6年度は決算の保険給付費＋介護予防・生活支援サービス事業費と"
     "おおむね一致する。年報の給付費とは範囲が異なる"),
    ("計画作成支援ツールの給付費（3町の行の合計）",
     sum(KJ.KYUFU_TOTAL[(t, "R6")][-1] for t in TOWNS3),
     sum(KJ.KYUFU_TOTAL[(t, "R7")][-1] for t in TOWNS3), "3町合計",
     "令和6年度・令和8年度は広域連合の行と完全に一致し、"
     "令和7年度のみ%s円（%.1f％）上回る（確認事項No.87）"
     % (yen(sum(KJ.KYUFU_TOTAL[(t, "R7")][-1] for t in TOWNS3)
            - KJ.KYUFU_TOTAL[("大雪", "R7")][-1]),
        (sum(KJ.KYUFU_TOTAL[(t, "R7")][-1] for t in TOWNS3)
         / KJ.KYUFU_TOTAL[("大雪", "R7")][-1] - 1) * 100)),
    ("令和6年度決算 保険給付費（款2）", _k2, None, "広域連合",
     "標準給付費見込額に当たる。割増率1.05909の分子"),
    ("同＋介護予防・生活支援サービス事業費", _k2 + _sogo_kessan, None, "3町合計",
     "計画作成支援ツールの給付費合計との差は%s円"
     % yen(KJ.KYUFU_TOTAL[("大雪", "R6")][-1] - (_k2 + _sogo_kessan))),
]:
    r = body(ws, r, [src, yen(v6) if v6 else "―", yen(v7) if v7 else "―",
                     tan, "", "", naiyo], height=32)
r = note(ws, r,
         "注）計画作成支援ツールの給付費合計は、"
         "令和6年度において決算の保険給付費（款2）に"
         "介護予防・生活支援サービス事業費を加えた額と%s円の差で一致します。"
         "ツールの給付費は総合事業費を含む範囲であるとみられます。［要確認］"
         % yen(abs(KJ.KYUFU_TOTAL[("大雪", "R6")][-1] - (_k2 + _sogo_kessan))),
         span=7, height=34)
r += 1

r = lead(ws, r, "3　見える化システムD32を用いない理由", span=7)
r = header(ws, r, ["論点", "内容", "", "", "", "", "影響"])
_pop1 = 9082
for k, v, e in [
    ("公表の精度", "受給率は小数第1位（0.1％）までしか公表されない。"
     "分母は第1号被保険者数（令和7年度%s人）" % "{:,}".format(_pop1),
     "0.1％＝%.1f人／月の刻みになる" % (_pop1 * 0.001)),
    ("実数の併載", "受給者数の実数が併載されるのは"
     "訪問介護（D32-a）・訪問看護（D32-c）・通所介護（D32-f）の3種別のみ",
     "他の13種別は丸めた率しか得られない"),
    ("小口のサービス", "訪問入浴介護は令和7年度で月5.2人。"
     "第1号被保険者数に対する率は0.06％で、0.1％未満",
     "率が0.0％となり、見込量が0人になる"),
    ("要介護度別の値", "要介護度別に率を掛けると、"
     "要介護度ごとに同じ刻みが生じる",
     "短期入所療養介護（月16.6人）のような種別で誤差が実数を上回る"),
]:
    r = body(ws, r, [k, v, "", "", "", "", e], height=34)
r = note(ws, r,
         "注）このため、サービス種別の見込量は"
         "見える化システムの総括表（実数）から算定します。"
         "D32は要介護度別の構成を見る用途に限って用います。", span=7, height=26)

# ============================================================ 03
ws = sheet("03_算定パターンと妥当性", "算定パターンと妥当性",
           "基準年度・出所・伸ばし方の組合せごとに、給付費と算定上の月額基準額を示します。"
           "「妥当」としたものを採用パターンとしています。"
           "月額はいずれも暫定であり、確定値ではありません。",
           [5, 22, 24, 22, 11, 12, 11, 44], freeze="A6")
r = 4
r = header(ws, r, ["No", "基準年度", "出所", "伸ばし方", "判定",
                   "給付費計\n（百万円）", "月額\n（円）", "判定の理由"])
for p in PATTERNS:
    g = p["gaku"]
    r = body(ws, r, [p["no"], p["kijun"], p["demae"], p["nobi"], p["hantei"],
                     round(sum(p["kyufu3"]), 1) if p["kyufu3"] else "―",
                     round(g["月額"]) if g else "―", p["riyu"]],
             fills={5: HAN_F[p["hantei"]]}, height=74)
r += 1
_ok = [p for p in PATTERNS if p["gaku"]]
_lo = min(_ok, key=lambda p: p["gaku"]["月額"])
_hi = max(_ok, key=lambda p: p["gaku"]["月額"])
r = note(ws, r,
         "注1）月額は算定上の月額基準額です。条例で定める基準額ではありません。\n"
         "注2）算定できたパターンは%d件で、月額は%s（%s）から%s（%s）まで"
         "%s円の幅があります。基準年度と出所の置き方だけでこれだけ動きます。\n"
         "注3）採用パターンは%sです。"
         "現行の将来推計 第3段階が用いているのは%sであり、"
         "両者の差は月額%s円です。\n"
         "注4）報酬改定率は未公表であるため、いずれのパターンも単価を固定しています。"
         "［要協議］"
         % (len(_ok),
            "{:,}円".format(round(_lo["gaku"]["月額"])), _lo["no"],
            "{:,}円".format(round(_hi["gaku"]["月額"])), _hi["no"],
            "{:,}".format(round(_hi["gaku"]["月額"] - _lo["gaku"]["月額"])),
            SAIYO, GENKO["no"],
            "{:+,}".format(round(BASIC["gaku"]["月額"]
                                 - GENKO["gaku"]["月額"]))),
         span=8, height=76)

# ============================================================ 04
ws = sheet("04_サービス見込量", "サービス見込量（採用パターン%s）" % SAIYO,
           "令和7年度の延べ利用者数を12で除した月平均に、"
           "認定者数の伸び率（シナリオ②トレンド継続）を乗じたものです。"
           "1人あたり給付費は令和7年度の実績で固定しています。",
           [30, 13, 12, 12, 12, 14, 15, 15], freeze="A6")
r = 4
r = lead(ws, r, "1　認定者数の伸び率", span=8)
r = header(ws, r, ["シナリオ", "令和7年度\n（基準）", "令和9年度", "令和10年度",
                   "令和11年度", "採否", "", ""])
for sc in (1, 2, 3):
    v = [BASE_NINTEI * x for x in NOBI[sc]]
    r = body(ws, r, [SC[sc], round(BASE_NINTEI), round(v[0]), round(v[1]),
                     round(v[2]),
                     "採用" if sc == 2 else "感度", "", ""],
             fills={6: OK_G if sc == 2 else IN_Y}, height=22)
for sc in (1, 2, 3):
    r = body(ws, r, ["　伸び率　" + SC[sc], 1.0, round(NOBI[sc][0], 4),
                     round(NOBI[sc][1], 4), round(NOBI[sc][2], 4), "", "", ""],
             height=20)
r += 1

r = lead(ws, r, "2　サービス種別の見込量（人／月）と給付費", span=8)
r = header(ws, r, ["サービス", "R7実績\n（人／月）", "令和9年度", "令和10年度",
                   "令和11年度", "1人あたり\n給付費（円）",
                   "第10期の給付費\n（百万円）", "備考"])
_sum_k = [0.0, 0.0, 0.0]
_n_zero = 0
for lab in SVC:
    n7, p7 = jisseki(lab, "R7"), tanka(lab, "R7")
    m7 = n7 / 12 if n7 is not None else 0.0
    vs = [m7 * f for f in NOBI[2]]
    if p7:
        for i, v in enumerate(vs):
            _sum_k[i] += v * 12 * p7
    ks = sum(v * 12 * (p7 or 0) for v in vs) / 1e6
    bk = ""
    if m7 == 0:
        _n_zero += 1
        bk = "区域内・区域外とも実績なし"
    r = body(ws, r, [short(lab), round(m7, 1), round(vs[0], 1),
                     round(vs[1], 1), round(vs[2], 1),
                     yen(p7) if p7 else "―", round(ks, 1), bk],
             fills={2: GRAY} if m7 == 0 else None, height=22)
r = body(ws, r, ["計", round(sum(jisseki(l, "R7") or 0 for l in SVC) / 12, 1),
                 "―", "―", "―", "―", round(sum(_sum_k) / 1e6, 1),
                 "延べ数のため利用者の実人数ではない"],
         fills={i: MID_B for i in range(1, 9)}, height=24, bold=True)
r = note(ws, r,
         "注1）利用者数は延べ数です。"
         "1人が複数のサービスを利用する場合はそれぞれに計上されます。"
         "したがって計の欄は利用者の実人数ではありません。\n"
         "注2）実績が0のサービスが%d種別あります。"
         "区域内に事業所がないサービスでも区域外の事業所の利用により"
         "実績が生じる場合があるため、事業所数とは別に見ています。\n"
         "注3）1人あたり給付費は延べ1人あたりの月額です。"
         "報酬改定率が未公表であるため令和7年度で固定しています。［要協議］\n"
         "注4）定期巡回・随時対応型訪問介護看護は令和7年度に延べ%s人"
         "（月%.1f人）の実績があります。区域内に事業所はありません。"
         % (_n_zero,
            "{:,}".format(int(jisseki(
                "在宅サービス 定期巡回・随時対応型訪問介護看護", "R7"))),
            jisseki("在宅サービス 定期巡回・随時対応型訪問介護看護", "R7") / 12),
         span=8, height=78)

# ============================================================ 05
ws = sheet("05_給付費と保険料", "給付費と保険料への影響",
           "採用パターンと現行の置き方で、標準給付費見込額から"
           "算定上の月額基準額までがどのように変わるかを示します。"
           "第10期の第1号被保険者負担割合・報酬改定率・基金残高は未確定です。",
           [5, 26, 22, 22, 14, 34], freeze="A6")
r = 4
r = lead(ws, r, "1　採用パターン%s と現行（%s）の対比" % (SAIYO, GENKO["no"]),
         span=6)
r = header(ws, r, ["記号", "項目", "採用%s（円）" % SAIYO,
                   "現行%s（円）" % GENKO["no"], "差", "置き方"])
_b, _g = BASIC["gaku"], GENKO["gaku"]
for sym, nm, key, oki in [
    ("", "総給付費（3か年計）", None,
     "%sは令和7年度実績×認定者数の伸び。%sは見える化の自然体推計×案C補正"
     % (SAIYO, GENKO["no"])),
    ("A", "標準給付費見込額", "A", "総給付費×割増率1.05909"),
    ("B", "地域支援事業費", "B", "令和6年度決算額×3年"),
    ("①", "合計（A＋B）", "①", "第1号被保険者負担分の算定基礎"),
    ("②", "調整交付金の算定基礎", "②", "A＋総合事業費"),
    ("C", "第1号被保険者負担分相当額", "C", "①×23％"),
    ("D", "調整交付金相当額", "D", "②×5％"),
    ("E", "調整交付金見込額", "E", "②×7.375％"),
    ("J", "保険料収納必要額", "J", "C＋D－E＋F＋G±H－I"),
]:
    if key is None:
        a, b = sum(BASIC["kyufu3"]) * 1e6, sum(GENKO["kyufu3"]) * 1e6
    else:
        a, b = _b[key], _g[key]
    r = body(ws, r, [sym, nm, yen(a), yen(b), yen(a - b), oki], height=26)
r = body(ws, r, ["③", "補正後被保険者数", "{:,}人".format(int(_b["③"])),
                 "{:,}人".format(int(_g["③"])), "―",
                 "推計上の被保険者数×令和6年度の実績係数。パターンによらない"],
         height=24)
r = body(ws, r, ["", "算定上の月額基準額",
                 "{:,}円".format(round(_b["月額"])),
                 "{:,}円".format(round(_g["月額"])),
                 "{:+,}円".format(round(_b["月額"] - _g["月額"])),
                 "J÷予定収納率99.0％÷③÷12"],
         fills={i: IN_Y for i in range(1, 7)}, height=24, bold=True)
r += 1

r = lead(ws, r, "2　年度別の給付費（採用パターン%s）" % SAIYO, span=6)
r = header(ws, r, ["", "区分", "令和9年度", "令和10年度", "令和11年度",
                   "第10期計"])
for i, (nm, vs) in enumerate([
    ("総給付費（百万円）", BASIC["kyufu3"]),
    ("（参考）現行%s" % GENKO["no"], GENKO["kyufu3"]),
], start=1):
    r = body(ws, r, [i, nm, round(vs[0], 1), round(vs[1], 1), round(vs[2], 1),
                     round(sum(vs), 1)], height=22)
r += 1

r = lead(ws, r, "3　月額基準額の幅", span=6)
r = header(ws, r, ["No", "パターン", "給付費計（百万円）", "月額（円）",
                   "採用との差", "判定"])
for i, p in enumerate([x for x in PATTERNS if x["gaku"]], start=1):
    r = body(ws, r, [i, "%s %s／%s" % (p["no"], p["kijun"], p["nobi"]),
                     round(sum(p["kyufu3"]), 1), round(p["gaku"]["月額"]),
                     "{:+,}".format(round(p["gaku"]["月額"] - _b["月額"])),
                     p["hantei"]],
             fills={6: HAN_F[p["hantei"]]}, height=30)
r = note(ws, r,
         "注1）月額は算定上の月額基準額です。"
         "第9期の条例で定める基準額は6,400円、算定上の月額基準額は6,428円です。\n"
         "注2）基金の取崩しは行わないものとしています。"
         "取崩しを行う場合は月額が下がります（将来推計 第3段階 07シート）。\n"
         "注3）第10期の第1号被保険者負担割合は政令によります。"
         "23％として置いています。［要確認］", span=6, height=48)

# ============================================================ 06 自己点検
# ------- 点検の実施
_bad_rows = []
for lab in list(D1):
    if lab not in D4 or lab not in D5:
        continue
    for y in ("R6", "R7"):
        n, p, g = (D1[lab][y]["実績値"], D4[lab][y]["実績値"],
                   D5[lab][y]["実績値"])
        if not n or not p or not g:
            continue
        if abs(n * p - g) / g > 0.001:
            _bad_rows.append((lab, y))
chk(1, "総括表の内的整合（利用者数×1人あたり給付費＝給付費）",
    "全%d行×2か年" % len(D1), "不一致%d件" % len(_bad_rows), not _bad_rows)

_e6 = abs(R6_TSUMI - N.KYUFU["給付費　合計"]["R6"])
_e7 = abs(R7_TSUMI - N.KYUFU["給付費　合計"]["R7"])
# 積上げは浮動小数の演算であるため、円未満の誤差（0.01円未満）を許容する。
chk(2, "総括表の給付費の合計と年報の給付費の一致（円単位）",
    "R6 %s円／R7 %s円" % (yen(R6_TSUMI), yen(R7_TSUMI)),
    "差 R6 %.4f円・R7 %.4f円（円未満）" % (_e6, _e7), _e6 < 0.5 and _e7 < 0.5)

_rebuild = sum((jisseki(l, "R7") or 0) / 12 * 12 * (tanka(l, "R7") or 0)
               for l in SVC)
chk(3, "見込量の基準値からの再構成（月平均×12×単価＝R7総給付費）",
    "%s円" % yen(_rebuild), "差 %s円" % yen(abs(_rebuild - R7_TSUMI)),
    abs(_rebuild - R7_TSUMI) < 1)

_g0 = gaku()
chk(4, "将来推計 第3段階の基本ケースの再現",
    "現行パターン%s の月額" % GENKO["no"],
    "%.0f円（第3段階 %.0f円）" % (_g["月額"], _g0["月額"]),
    abs(_g["月額"] - _g0["月額"]) < 0.5)

_k9 = _P["K9"]
_j9 = _k9["C"] + _k9["D"] - _k9["E"] - _k9["I"]
_g9 = _j9 / _k9["shuno"] / _k9["hosei"] / 12
chk(5, "第9期の算定式の再現（J÷収納率÷補正後被保険者数÷12）",
    "J＝%s円" % yen(_j9), "%.1f円（公表値%d円）" % (_g9, _k9["gaku"]),
    abs(_g9 - _k9["gaku"]) < 1)

_ratio = [BASIC["kyufu3"][i] / (R7_TSUMI / 1e6) for i in range(3)]
chk(6, "見込量の伸び率が認定者数の伸び率と一致すること",
    "給付費の伸び %.4f／%.4f／%.4f" % tuple(_ratio),
    "認定者数の伸び %.4f／%.4f／%.4f" % tuple(NOBI[2]),
    all(abs(a - b) < 1e-9 for a, b in zip(_ratio, NOBI[2])))

_uses_r8 = [p["no"] for p in PATTERNS
            if p["kyufu3"] and "令和8年度" in p["kijun"]]
chk(7, "令和8年度を基準年度に用いたパターンが算定に含まれないこと",
    "算定したパターン%d件" % len([p for p in PATTERNS if p["kyufu3"]]),
    "令和8年度基準の算定%d件" % len(_uses_r8), not _uses_r8)

_neg = [short(l) for l in SVC
        if (jisseki(l, "R7") or 0) < 0 or (tanka(l, "R7") or 0) < 0]
chk(8, "見込量及び単価に負の値がないこと", "全%d種別" % len(SVC),
    "負の値%d件" % len(_neg), not _neg)

_miss = [short(l) for l in SVC
         if jisseki(l, "R7") is None or
         (jisseki(l, "R7") and tanka(l, "R7") is None)]
chk(9, "実績のある種別に単価が収録されていること", "全%d種別" % len(SVC),
    "欠測%d件" % len(_miss), not _miss)

_sum05 = sum(_sum_k) / 1e6
chk(10, "04シートの給付費の合計と03シートの採用パターンの一致",
    "04シート %.1f百万円" % _sum05,
    "03シート %.1f百万円" % sum(BASIC["kyufu3"]),
    abs(_sum05 - sum(BASIC["kyufu3"])) < 0.05)

# 3町合計と広域連合単位の差は令和7年度のみに生じる（確認事項No.87）。
_gap = {y: sum(KJ.KYUFU_TOTAL[(t, y)][-1] for t in TOWNS3)
        - KJ.KYUFU_TOTAL[("大雪", y)][-1] for y in ("R6", "R7", "R8")}
chk(11, "3町合計と広域連合単位の給付費の差が令和7年度のみに生じること",
    "R6 %s円／R7 %s円／R8 %s円" % tuple(yen(_gap[y]) for y in
                                    ("R6", "R7", "R8")),
    "令和6年度・令和8年度は一致、令和7年度のみ%.1f％の差"
    % (_gap["R7"] / KJ.KYUFU_TOTAL[("大雪", "R7")][-1] * 100),
    _gap["R6"] == 0 and _gap["R8"] == 0 and _gap["R7"] != 0)

NG_WORDS = ["に由来する", "と整合する", "1件も", "有意差がないため", "全国トップ級"]

ws = sheet("06_自己点検", "自己点検（エラーチェック）",
           "算定の内的整合と、前段の推計との接続を点検した結果です。"
           "1件でも不適合があると、本表を作るスクリプトは終了コード1で終わります。",
           [5, 44, 34, 34, 12], freeze="A6")
r = 4
r = header(ws, r, ["No", "点検の内容", "式・対象", "結果", "判定"])
for no, naiyo, shiki, kekka, han in CHECKS:
    r = body(ws, r, [no, naiyo, shiki, kekka, han],
             fills={5: OK_G if han == "適合" else NG_O}, height=32)
_ng = [c for c in CHECKS if c[4] != "適合"]
r = body(ws, r, ["", "計", "%d件" % len(CHECKS),
                 "適合%d件・不適合%d件" % (len(CHECKS) - len(_ng), len(_ng)),
                 "適合" if not _ng else "不適合"],
         fills={i: (MID_B if not _ng else NG_O) for i in range(1, 6)},
         height=24, bold=True)
r += 1
r = lead(ws, r, "点検の対象としていないもの", span=5)
r = header(ws, r, ["", "項目", "理由", "", ""])
for i, (k, v) in enumerate([
    ("令和8年度の月数", "公表されていないため、給付費の比から推計するにとどめた。"
     "［要確認］"),
    ("3町合計と広域連合単位の給付費の差（令和7年度4.0％）",
     "確認事項No.87として照会中である。本表では広域連合単位を用いている"),
    ("計画作成支援ツールの予防訪問看護（令和7年度0人）",
     "計上の変化か欠測かを確認できていない。［要確認］"),
    ("報酬改定率", "未公表であるため単価を令和7年度で固定した。［要協議］"),
    ("第10期の第1号被保険者負担割合", "政令によるため23％として置いた。［要確認］"),
], start=1):
    r = body(ws, r, [i, k, v, "", ""], height=32)

# ============================================================ 07
ws = sheet("07_確認事項", "確認事項",
           "見込量の採用値を定めるために、ご判断又はご確認をお願いする事項です。"
           "業務工程管理表 03_確認事項一覧にも登録します。",
           [5, 16, 46, 40, 14], freeze="A6")
r = 4
r = header(ws, r, ["", "区分", "確認事項", "理由・背景", "希望時期"])
KAKUNIN = [
    ("ご判断",
     "サービス見込量及び給付費の算定を、"
     "採用パターン%s（令和7年度基準・見える化システムの総括表・"
     "認定者数シナリオ②）に改めてよいか" % SAIYO,
     "現行の将来推計 第3段階は%s（見える化システムの自然体推計×案C補正）を"
     "用いており、算定上の月額基準額は%s円である。"
     "採用パターンに改めると%s円となり、差は%s円である。"
     "%sは利用率と単価の趨勢が自然体推計に織り込まれており、"
     "当方でその中身を分解できない。"
     "採用パターンは認定者数の伸びのみを反映するため、伸びの中身を説明できる。"
     % (GENKO["no"], "{:,}".format(round(_g["月額"])),
        "{:,}".format(round(_b["月額"])),
        "{:+,}".format(round(_b["月額"] - _g["月額"])), GENKO["no"]),
     "令和8年9月26日まで"),
    ("ご判断",
     "計画素案 第6章第2節2の見込量の表の数値を、"
     "本表の算定結果に差し替えてよいか",
     "計画素案の同表の数値は、見える化システムの総括表・"
     "計画作成支援ツール・将来推計 第2段階のいずれとも一致せず、"
     "当方で出所を特定できていない。"
     "本表の算定結果に差し替えれば、"
     "見込量・給付費・保険料が一本の経路でつながる。",
     "令和8年9月26日まで"),
    ("ご確認",
     "計画作成支援ツールの令和8年度が何か月分の実績か",
     "給付費が令和7年度の%.1f％であることから、"
     "おおむね%.1fか月分と見ているが、月数そのものは公表されていない。"
     "令和8年度を実績として掲げる場合、月数の明記が要る。"
     % (R8_RATIO * 100, R8_TSUKI),
     "令和8年9月"),
    ("ご確認",
     "計画作成支援ツールの予防訪問看護が令和7年度0人であること",
     "令和6年度は18人であり、令和7年度に0人となっている。"
     "同じ年度の見える化システムの訪問看護（月135.2人）と"
     "ツールの訪問看護（月117人）の差18.2人にほぼ一致するため、"
     "計上の変化とみられる。欠測か、訪問看護への統合かをご確認いただきたい。",
     "令和8年9月"),
    ("ご確認",
     "報酬改定率の取扱い",
     "第10期の初年度（令和9年度）に報酬改定が見込まれるが率は未公表である。"
     "本表は1人あたり給付費を令和7年度実績で固定している。"
     "改定率が示された時点で単価に乗じる方法でよいか。",
     "令和8年10月"),
    ("ご確認",
     "第10期の第1号被保険者負担割合",
     "第9期と同じ23％として置いている。政令により定まる。"
     "1ポイントの違いで月額が約290円動く。",
     "令和8年10月"),
]
for i, (k, a, b, c) in enumerate(KAKUNIN, start=1):
    f = {2: IN_Y if k == "ご判断" else MID_B}
    r = body(ws, r, [i, k, a, b, c], fills=f, height=62)
r += 1
r = note(ws, r,
         "注）確認事項は業務工程管理表 03_確認事項一覧で一元管理します。"
         "計画素案の本文には確認事項の注記を書きません。", span=5, height=20)

# ============================================================ 出力
os.makedirs(ODIR, exist_ok=True)
wb.save(OUT)

# 禁止表現の点検（本ファイルの出力に対して）
_ngw = []
for _ws in wb.worksheets:
    for _row in _ws.iter_rows(values_only=True):
        for _v in _row:
            if isinstance(_v, str):
                for _w in NG_WORDS:
                    if _w in _v:
                        _ngw.append((_ws.title, _w))
print("書き出しました:", OUT)
for s in wb.sheetnames:
    print("  -", s, wb[s].max_row, "rows")
print("パターン%d件（算定できたもの%d件）／サービス%d種別／確認事項%d件"
      % (len(PATTERNS), len(_ok), len(SVC), len(KAKUNIN)))
print("採用%s 月額%.0f円　現行%s 月額%.0f円　差%+.0f円"
      % (SAIYO, _b["月額"], GENKO["no"], _g["月額"], _b["月額"] - _g["月額"]))
print("自己点検 %d件：適合%d件・不適合%d件"
      % (len(CHECKS), len(CHECKS) - len(_ng), len(_ng)))
if _ngw:
    print("禁止表現:", _ngw)
if _ng or _ngw:
    for c in _ng:
        print("  不適合:", c[0], c[1], c[3])
    sys.exit(1)
print("すべての点検に適合しました。")
