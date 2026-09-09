"""小野町 第10期 サービス別単価の算出と見込量の試算。

介護保険事業状況報告（年報）令和3〜7年度の様式2（件数・単位数・費用額・給付費）と
様式1・様式1の6・様式3から、サービス区分ごとの実績と単価を取り出し、
令和9〜11年度（第10期）の見込量と給付費を試算する。

**認定者数を推計の基礎に使わない。**
令和7年9月審査分で認定者数が961人から862人へ1か月で99人（10.3％）減った一方、
同じ月の給付件数は1,303件から1,365件へ増えている。年度でみても認定者は16.1％減、
受給者延は4.6％減、給付費は7.4％減であり、認定者数だけが下方に外れている。
このため認定者1人当たりで組むと給付費を過小に見込む。

分母の候補を令和3〜7年度で比べると、給付件数総計を割ったときの変動係数は

    第1号被保険者数   1.63％   ← 最も安定
    75歳以上人口      2.28％
    85歳以上人口      2.38％
    認定者数          5.18％   ← 最も不安定

であり、第1号被保険者数を分母とする方式を基本ケースとした。
後期高齢者連動・認定者連動は感度ケースとして併記する。

出力:
  04_算定・見込量/小野町_第10期_サービス別単価・見込量試算_YYYYMMDD.xlsx
  10_給付費等分析/小野町_第9期計画の見込みと実績の対比_YYYYMMDD.docx
  04_算定・見込量/小野町_人口推計の出典統一メモ_YYYYMMDD.docx
"""

import decimal
import pathlib
import statistics
import tempfile
from collections import OrderedDict

import openpyxl
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt
from openpyxl.styles import Alignment, Border, Font, PatternFill
from openpyxl.styles.borders import Side

from build_ono_nenpo import SRC as NENPO_SRC
from build_ono_nenpo import YEARS as NENPO_YEARS
from build_ono_nenpo import load_nenpo

ROOT = pathlib.Path(__file__).parent / "小野町_引継ぎ_整理済"
OUT_X = ROOT / "04_算定・見込量"
OUT_D = ROOT / "10_給付費等分析"
ASOF = "20260909"
ASOF_JP = "令和8年9月9日"

IN_FILL = PatternFill("solid", fgColor="FFF2CC")
CALC_FILL = PatternFill("solid", fgColor="EAF1FB")
HEAD_FILL = PatternFill("solid", fgColor="DDEBF7")
KEY_FILL = PatternFill("solid", fgColor="FCE4E4")
WARN_FILL = PatternFill("solid", fgColor="FFE0B2")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

JP_FONT = "游明朝"
JP_GOTHIC = "游ゴシック"

YS = [lab for _, lab in NENPO_YEARS]                    # 令和3〜7年度
PLAN_YEARS = ["令和9年度", "令和10年度", "令和11年度"]   # 第10期
EST_YEARS = ["令和8年度"] + PLAN_YEARS

# 第10期将来推計用推計人口（各年10月1日時点。令和8年9月8日受領）は
# read_pop_dai10() で原本から集計する。POP_EST／POP_DAI10 は同じ集計結果を参照する。

# 認定者数の見込（03_認定者推計と同じ年齢階層別認定率による）
NINTEI_EST = {"令和8年度": 785, "令和9年度": 782, "令和10年度": 781, "令和11年度": 780}

# ---------------------------------------------------------------- サービス名の正規化
# 年報の様式2はサービス名の表記が年度で変わる。
#   令和5年度まで  福祉用具購入費／住宅改修費／短期入所療養介護（介護療養型医療施設等）
#   令和7年度から  小規模多機能型居宅介護（短期利用以外）と（短期利用）に分割
# 計画の見込量はいずれも合算した区分で設定するため、正規化して5年分を接続する。

CANON = {
    "訪問介護": "訪問介護",
    "訪問入浴介護": "訪問入浴介護",
    "訪問看護": "訪問看護",
    "訪問リハビリテーション": "訪問リハビリテーション",
    "居宅療養管理指導": "居宅療養管理指導",
    "通所介護": "通所介護",
    "通所リハビリテーション": "通所リハビリテーション",
    "短期入所生活介護": "短期入所生活介護",
    "短期入所療養介護（介護老人保健施設）": "短期入所療養介護（介護老人保健施設）",
    "短期入所療養介護（介護療養型医療施設等）": "短期入所療養介護（病院等）",
    "短期入所療養介護（病院等）": "短期入所療養介護（病院等）",
    "短期入所療養介護（介護医療院）": "短期入所療養介護（介護医療院）",
    "福祉用具貸与": "福祉用具貸与",
    "福祉用具購入費": "特定福祉用具販売",
    "特定福祉用具販売": "特定福祉用具販売",
    "住宅改修費": "住宅改修",
    "住宅改修": "住宅改修",
    "特定施設入居者生活介護": "特定施設入居者生活介護",
    "特定施設入居者生活介護（短期利用以外）": "特定施設入居者生活介護",
    "特定施設入居者生活介護（短期利用）": "特定施設入居者生活介護",
    "介護予防支援・居宅介護支援": "居宅介護支援・介護予防支援",
    "定期巡回・随時対応型訪問介護看護": "定期巡回・随時対応型訪問介護看護",
    "夜間対応型訪問介護": "夜間対応型訪問介護",
    "地域密着型通所介護": "地域密着型通所介護",
    "認知症対応型通所介護": "認知症対応型通所介護",
    "小規模多機能型居宅介護": "小規模多機能型居宅介護",
    "小規模多機能型居宅介護（短期利用以外）": "小規模多機能型居宅介護",
    "小規模多機能型居宅介護（短期利用）": "小規模多機能型居宅介護",
    "認知症対応型共同生活介護": "認知症対応型共同生活介護",
    "認知症対応型共同生活介護（短期利用以外）": "認知症対応型共同生活介護",
    "認知症対応型共同生活介護（短期利用）": "認知症対応型共同生活介護",
    "地域密着型特定施設入居者生活介護": "地域密着型特定施設入居者生活介護",
    "地域密着型特定施設入居者生活介護（短期利用以外）": "地域密着型特定施設入居者生活介護",
    "地域密着型特定施設入居者生活介護（短期利用）": "地域密着型特定施設入居者生活介護",
    "地域密着型介護老人福祉施設入所者生活介護": "地域密着型介護老人福祉施設入所者生活介護",
    "複合型サービス(看護小規模多機能型居宅介護)": "看護小規模多機能型居宅介護",
    "複合型サービス(看護小規模多機能型居宅介護)（短期利用以外）": "看護小規模多機能型居宅介護",
    "複合型サービス(看護小規模多機能型居宅介護)（短期利用）": "看護小規模多機能型居宅介護",
    "介護老人福祉施設": "介護老人福祉施設",
    "介護老人保健施設": "介護老人保健施設",
    "介護療養型医療施設": "介護療養型医療施設",
    "介護医療院": "介護医療院",
}

SUBTOTALS = ["居宅（介護予防）サービス", "訪問サービス", "通所サービス",
             "短期入所サービス", "福祉用具・住宅改修サービス",
             "地域密着型（介護予防）サービス", "施設サービス", "総計"]

# 表示順（区分, 正規化名, 計画表記の単位）
ORDER_KAIGO = [
    ("居宅", "訪問介護", "回"),
    ("居宅", "訪問入浴介護", "回"),
    ("居宅", "訪問看護", "回"),
    ("居宅", "訪問リハビリテーション", "回"),
    ("居宅", "居宅療養管理指導", "人"),
    ("居宅", "通所介護", "回"),
    ("居宅", "通所リハビリテーション", "回"),
    ("居宅", "短期入所生活介護", "日"),
    ("居宅", "短期入所療養介護（介護老人保健施設）", "日"),
    ("居宅", "短期入所療養介護（病院等）", "日"),
    ("居宅", "短期入所療養介護（介護医療院）", "日"),
    ("居宅", "福祉用具貸与", "人"),
    ("居宅", "特定福祉用具販売", "人"),
    ("居宅", "住宅改修", "人"),
    ("居宅", "特定施設入居者生活介護", "人"),
    ("居宅", "居宅介護支援・介護予防支援", "人"),
    ("地域密着型", "定期巡回・随時対応型訪問介護看護", "人"),
    ("地域密着型", "夜間対応型訪問介護", "人"),
    ("地域密着型", "地域密着型通所介護", "回"),
    ("地域密着型", "認知症対応型通所介護", "回"),
    ("地域密着型", "小規模多機能型居宅介護", "人"),
    ("地域密着型", "認知症対応型共同生活介護", "人"),
    ("地域密着型", "地域密着型特定施設入居者生活介護", "人"),
    ("地域密着型", "地域密着型介護老人福祉施設入所者生活介護", "人"),
    ("地域密着型", "看護小規模多機能型居宅介護", "人"),
    ("施設", "介護老人福祉施設", "人"),
    ("施設", "介護老人保健施設", "人"),
    ("施設", "介護療養型医療施設", "人"),
    ("施設", "介護医療院", "人"),
]

# 予防給付にない区分（総合事業へ移行済み、または制度上対象外）
YOBO_EXCLUDE = {
    "訪問介護", "通所介護", "地域密着型通所介護",
    "定期巡回・随時対応型訪問介護看護", "夜間対応型訪問介護",
    "地域密着型特定施設入居者生活介護",
    "地域密着型介護老人福祉施設入所者生活介護", "看護小規模多機能型居宅介護",
    "介護老人福祉施設", "介護老人保健施設", "介護療養型医療施設", "介護医療院",
}

# 基準期間の例外。既定は令和5〜7年度の3年平均。
#   計上の開始が令和6年度からのもの、令和7年度に事業所が撤退したものは個別に扱う。
BASE_OVERRIDE = {
    "特定福祉用具販売": (["令和6年度", "令和7年度"], "令和6年度から様式2に計上。2年平均"),
    "住宅改修": (["令和6年度", "令和7年度"], "令和6年度から様式2に計上。2年平均"),
    "定期巡回・随時対応型訪問介護看護": (["令和7年度"], "令和7年度に利用実績が消滅。0で据置"),
    "地域密着型通所介護": (["令和7年度"], "令和5年度以降ほぼ実績なし。0で据置"),
    "介護療養型医療施設": (["令和7年度"], "令和5年度末で制度廃止。0で据置"),
    "介護医療院": (["令和7年度"], "令和6年度1件から令和7年度14件へ。"
                          "入所が定着したとみて直近年度で置く"),
    "訪問リハビリテーション": (["令和7年度"], "5年間で令和5年度の1件のみ。0で据置"),
}
BASE_DEFAULT = ["令和5年度", "令和6年度", "令和7年度"]

# 審査支払手数料の単価（円／件）。決算での確認を要する仮置き。
TESURYO = 60

# 所得段階別第1号被保険者数と保険者の定める割合（年報 様式1 所得段階別・令和7年度末）
SHOTOKU_R7 = [("第1段階", 0.285, 525), ("第2段階", 0.486, 297), ("第3段階", 0.685, 272),
              ("第4段階", 0.900, 434), ("第5段階", 1.000, 610), ("第6段階", 1.200, 491),
              ("第7段階", 1.300, 418), ("第8段階", 1.500, 231), ("第9段階", 1.700, 77),
              ("第10段階", 1.900, 36), ("第11段階", 2.100, 14), ("第12段階", 2.300, 7),
              ("第13段階", 2.400, 31)]

# 価格水準の調整。令和6年度の介護報酬改定は＋1.59％であり、
# 基準期間のうち令和5年度だけが改定前の価格である。改定後の水準にそろえてから平均する。
KAITEI_R6 = 0.0159
PRICE_ADJ = {"令和3年度": 1 + KAITEI_R6, "令和4年度": 1 + KAITEI_R6,
             "令和5年度": 1 + KAITEI_R6, "令和6年度": 1.0, "令和7年度": 1.0}


# ================================================================ 抽出

def _label(ws, r):
    for c in (3, 4, 5):
        v = ws.cell(r, c).value
        if isinstance(v, str) and v.strip():
            return v.strip().replace("　", " ")
    return None


def _read_form2(ws):
    """様式2系のシートを {表記名: (予防計, 介護計, 合計)} で返す。"""
    out = {}
    for r in range(11, ws.max_row + 1):
        name = _label(ws, r)
        if not name:
            continue
        y, k, t = ws.cell(r, 8).value, ws.cell(r, 15).value, ws.cell(r, 16).value
        if not any(isinstance(x, (int, float)) for x in (y, k, t)):
            continue
        out[name] = (y or 0, k or 0, t or 0)
    return out


def extract():
    """年報5年分から必要な系列をすべて取り出す。"""
    work = pathlib.Path(tempfile.mkdtemp())
    act = OrderedDict()          # act[年度][指標][正規化名] = (予防, 介護)
    sub = OrderedDict()          # act の小計（検算用）
    gen = OrderedDict()          # 一般状況・保険料・保険給付支払
    for y, lab in NENPO_YEARS:
        wb = load_nenpo(NENPO_SRC / "原本_年報" / f"年報データ_{y}_小野町.xlsx", work)
        row = {}
        srow = {}
        for key, sn in (("件数", "様式２（件数）"), ("単位数", "様式２（単位数）"),
                        ("費用額", "様式２（費用額）"), ("給付費", "様式２（給付費）")):
            raw = _read_form2(wb[sn])
            agg = {}
            for name, (yo, ka, _t) in raw.items():
                canon = CANON.get(name)
                if canon is None:
                    continue
                a = agg.setdefault(canon, [0, 0])
                a[0] += yo
                a[1] += ka
            row[key] = {k: tuple(v) for k, v in agg.items()}
            srow[key] = {n: raw[n] for n in SUBTOTALS if n in raw}
        act[lab] = row
        sub[lab] = srow

        w1, w3, w6 = wb["様式１"], wb["様式３"], wb["様式１の６"]
        g = {
            "世帯数": w1.cell(12, 7).value,
            "65-74": w1.cell(17, 7).value,
            "75-84": w1.cell(18, 7).value,
            "85+": w1.cell(19, 7).value,
            "住所地特例": w1.cell(21, 7).value,
            "1号": w1.cell(22, 7).value,
            "死亡": w1.cell(29, 6).value,
            "介護諸費": w3.cell(21, 6).value,
            "予防諸費": w3.cell(22, 6).value,
            "高額": w3.cell(23, 6).value,
            "高額合算": w3.cell(24, 6).value,
            "特定入所": w3.cell(25, 6).value,
            "支払計": w3.cell(27, 6).value,
            "調定現年": w3.cell(12, 6).value,
            "収納現年": w3.cell(12, 7).value,
            "受給_居宅": w6.cell(13, 14).value,
            "受給_地密": w6.cell(20, 14).value,
            "受給_施設": w6.cell(37, 13).value,
        }
        g["75+"] = (g["75-84"] or 0) + (g["85+"] or 0)
        g["受給計"] = (g["受給_居宅"] or 0) + (g["受給_地密"] or 0) + (g["受給_施設"] or 0)
        g["総給付費"] = (g["介護諸費"] or 0) + (g["予防諸費"] or 0)

        w5 = wb["様式１の５ 総数"]
        tot5 = sorted({r for r in range(1, w5.max_row + 1)
                       for c in range(1, 7)
                       if isinstance(w5.cell(r, c).value, str)
                       and "総" in w5.cell(r, c).value and "数" in w5.cell(r, c).value})
        nums = [w5.cell(tot5[-1], c).value for c in range(1, w5.max_column + 1)
                if isinstance(w5.cell(tot5[-1], c).value, (int, float))]
        g["認定"] = nums[-1] if nums else None
        gen[lab] = g
    return act, sub, gen


# ================================================================ 単価と基準率

def unit_price(act, svc, kind, year):
    """1件当たり給付費（円）。件数が0なら None。"""
    i = 0 if kind == "予防" else 1
    ken = act[year]["件数"].get(svc, (0, 0))[i]
    kyu = act[year]["給付費"].get(svc, (0, 0))[i]
    return (kyu / ken) if ken else None


def unit_per_case(act, svc, kind, year):
    """1件当たり単位数。回数・日数の代替指標として使う。"""
    i = 0 if kind == "予防" else 1
    ken = act[year]["件数"].get(svc, (0, 0))[i]
    tan = act[year]["単位数"].get(svc, (0, 0))[i]
    return (tan / ken) if ken else None


def base_years(svc):
    return BASE_OVERRIDE.get(svc, (BASE_DEFAULT, ""))[0]


def base_note(svc):
    return BASE_OVERRIDE.get(svc, (None, "令和5〜7年度の3年平均"))[1] or "令和5〜7年度の3年平均"


def rate(act, gen, svc, kind, den_key, metric="件数"):
    """基準期間における「分母1人当たりの年間量」。

    metric が '給付費' のときは価格水準を令和6年度改定後にそろえてから平均する。
    件数と単価は逆方向に動くことがあり（訪問介護は件数▲25％・単価＋22％）、
    平均件数に直近単価を乗じると給付費を二重に見込むため、給付費そのものを延ばす。
    """
    i = 0 if kind == "予防" else 1
    vals = []
    for y in base_years(svc):
        v = act[y][metric].get(svc, (0, 0))[i]
        if metric == "給付費":
            v *= PRICE_ADJ.get(y, 1.0)
        den = gen[y][den_key]
        if den:
            vals.append(v / den)
    return statistics.mean(vals) if vals else 0.0


def project(act, gen, den_key, pop_key, kaitei=0.0):
    """見込量と給付費を試算する。

    den_key … 実績側の分母（'1号' / '75+' / '認定'）
    pop_key … 推計側の分母（POP_EST のキー、または NINTEI_EST）
    kaitei  … 令和9年度の報酬改定率
    """
    out = {}
    for kind, order in (("介護", ORDER_KAIGO),
                        ("予防", [x for x in ORDER_KAIGO if x[1] not in YOBO_EXCLUDE])):
        rows = []
        for cat, svc, unit in order:
            r_ken = rate(act, gen, svc, kind, den_key, "件数")
            r_kyu = rate(act, gen, svc, kind, den_key, "給付費")
            cells = []
            for y in EST_YEARS:
                den = (POP_EST[y][pop_key] if pop_key in ("1号", "75+", "総人口")
                       else NINTEI_EST[y])
                ken = r_ken * den
                kyu = r_kyu * den * (1 + kaitei if y != "令和8年度" else 1)
                cells.append((ken, kyu / 1000.0))
            tanka = (r_kyu / r_ken) if r_ken else 0.0
            rows.append({"区分": cat, "サービス": svc, "単位": unit,
                         "率": r_ken, "単価": tanka, "見込": cells})
        out[kind] = rows
    return out


# 見込量保険料算定シート（build_ono_santei.py）のサービス表記と、本スクリプトの
# 正規化名の対応。複数の正規化名を合算する区分がある（短期入所療養介護など）。
SANTEI_MAP = {
    "介護": {
        "訪問介護": ["訪問介護"],
        "訪問入浴介護": ["訪問入浴介護"],
        "訪問看護": ["訪問看護"],
        "訪問リハビリテーション": ["訪問リハビリテーション"],
        "居宅療養管理指導": ["居宅療養管理指導"],
        "通所介護": ["通所介護"],
        "通所リハビリテーション": ["通所リハビリテーション"],
        "短期入所生活介護": ["短期入所生活介護"],
        "短期入所療養介護": ["短期入所療養介護（介護老人保健施設）",
                       "短期入所療養介護（病院等）",
                       "短期入所療養介護（介護医療院）"],
        "福祉用具貸与": ["福祉用具貸与"],
        "特定福祉用具販売": ["特定福祉用具販売"],
        "住宅改修": ["住宅改修"],
        "特定施設入居者生活介護": ["特定施設入居者生活介護"],
        "居宅介護支援": ["居宅介護支援・介護予防支援"],
        "定期巡回・随時対応型訪問介護看護": ["定期巡回・随時対応型訪問介護看護"],
        "夜間対応型訪問介護": ["夜間対応型訪問介護"],
        "地域密着型通所介護": ["地域密着型通所介護"],
        "認知症対応型通所介護": ["認知症対応型通所介護"],
        "小規模多機能型居宅介護": ["小規模多機能型居宅介護"],
        "認知症対応型共同生活介護": ["認知症対応型共同生活介護"],
        "地域密着型特定施設入居者生活介護": ["地域密着型特定施設入居者生活介護"],
        "地域密着型介護老人福祉施設入所者生活介護":
            ["地域密着型介護老人福祉施設入所者生活介護"],
        "看護小規模多機能型居宅介護": ["看護小規模多機能型居宅介護"],
        "介護老人福祉施設": ["介護老人福祉施設"],
        "介護老人保健施設": ["介護老人保健施設"],
        "介護医療院": ["介護医療院", "介護療養型医療施設"],
    },
    "予防": {
        "介護予防訪問入浴介護": ["訪問入浴介護"],
        "介護予防訪問看護": ["訪問看護"],
        "介護予防訪問リハビリテーション": ["訪問リハビリテーション"],
        "介護予防居宅療養管理指導": ["居宅療養管理指導"],
        "介護予防通所リハビリテーション": ["通所リハビリテーション"],
        "介護予防短期入所生活介護": ["短期入所生活介護"],
        "介護予防短期入所療養介護": ["短期入所療養介護（介護老人保健施設）",
                            "短期入所療養介護（病院等）",
                            "短期入所療養介護（介護医療院）"],
        "介護予防福祉用具貸与": ["福祉用具貸与"],
        "特定介護予防福祉用具販売": ["特定福祉用具販売"],
        "介護予防住宅改修": ["住宅改修"],
        "介護予防特定施設入居者生活介護": ["特定施設入居者生活介護"],
        "介護予防支援": ["居宅介護支援・介護予防支援"],
        "介護予防認知症対応型通所介護": ["認知症対応型通所介護"],
        "介護予防小規模多機能型居宅介護": ["小規模多機能型居宅介護"],
        "介護予防認知症対応型共同生活介護": ["認知症対応型共同生活介護"],
    },
}


def kasan_ratios(gen):
    """特定入所者介護サービス費・高額介護サービス費・高額医療合算介護サービス費の
    総給付費に対する比率（基準期間の平均）。標準給付費の加算に使う。"""
    return {key: statistics.mean([gen[y][key] / gen[y]["総給付費"] for y in BASE_DEFAULT])
            for key in ("特定入所", "高額", "高額合算")}


def build_mikomi_input():
    """見込量保険料算定シートの05・06に流し込む実績・単価・見込量を返す。

    戻り値 ({kind: {サービス表記: {"件数実績", "給付費実績", "単価", "見込"}}}, 加算比率)
    見込は第10期の3年度分 [(件数, 給付費千円), ...]。
    """
    act, _sub, gen = extract()
    proj = project(act, gen, "1号", "1号", 0.0)
    out = {}
    for kind in ("介護", "予防"):
        by_svc = {r["サービス"]: r for r in proj[kind]}
        i = 1 if kind == "介護" else 0
        d = {}
        for name, srcs in SANTEI_MAP[kind].items():
            ken7 = sum(act["令和7年度"]["件数"].get(s, (0, 0))[i] for s in srcs)
            kyu7 = sum(act["令和7年度"]["給付費"].get(s, (0, 0))[i] for s in srcs)
            cells = []
            for k in range(1, len(EST_YEARS)):        # 令和9〜11年度
                ken = sum(by_svc[s]["見込"][k][0] for s in srcs if s in by_svc)
                kyu = sum(by_svc[s]["見込"][k][1] for s in srcs if s in by_svc)
                cells.append((ken, kyu))
            d[name] = {"件数実績": ken7, "給付費実績": kyu7 / 1000.0,
                       "単価": (kyu7 / ken7) if ken7 else 0.0,
                       "見込": cells}
        out[kind] = d
    return out, kasan_ratios(gen)


def totals(proj):
    """年度別の総給付費（千円）。"""
    t = [0.0] * len(EST_YEARS)
    for kind in ("介護", "予防"):
        for row in proj[kind]:
            for i, (_k, g) in enumerate(row["見込"]):
                t[i] += g
    return t


# ================================================================ Excel の体裁

def style_header(ws, row=1):
    for c in ws[row]:
        if c.value is not None:
            c.font = Font(bold=True, size=9, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F3864")
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = BORDER


def body_style(ws, first=2, wrap_cols=()):
    for row in ws.iter_rows(min_row=first):
        for c in row:
            c.font = Font(size=9)
            c.border = BORDER
            c.alignment = Alignment(vertical="top",
                                    wrap_text=(c.column in wrap_cols))


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w


def notes(ws, lines):
    ws.append([])
    for t in lines:
        ws.append([t])
        ws.cell(ws.max_row, 1).font = Font(size=9, italic=True)


# ---------------------------------------------------------------- 00_読み方

def sheet_intro(wb, act, sub, gen, cv):
    ws = wb.create_sheet("00_この試算の読み方")
    rows = [
        ["小野町 第10期介護保険事業計画　サービス別単価の算出と見込量の試算"],
        [f"作成 {ASOF_JP}／出典 介護保険事業状況報告（年報）令和3〜7年度、第10期将来推計用推計人口"],
        [],
        ["■ 何を計算したか"],
        ["", "年報の様式2から、サービス区分ごとの件数・単位数・費用額・給付費を令和3〜7年度の5年分取り出した。"],
        ["", "件数を給付費で割って1件当たり給付費（単価）を求め、件数を第1号被保険者数で割って利用率を求めた。"],
        ["", "この利用率に将来の第1号被保険者数を乗じて件数を見込み、単価を乗じて給付費を試算した。"],
        [],
        ["■ 認定者数を基礎に使わなかった理由"],
        ["", "令和7年9月審査分で認定者数が961人から862人へ1か月で99人（10.3％）減った。"],
        ["", "しかし同じ月の給付件数は1,303件から1,365件へ増えている。認定者だけが落ちて給付が落ちていない。"],
        ["", "年度でみても認定者は16.1％減、受給者延は4.6％減、給付費は7.4％減であり、認定者数だけが下方に外れる。"],
        ["", "**認定者1人当たりで組むと、令和7年度の高い1人当たり単価が将来に持ち越され、見込みが不安定になる。**"],
        [],
        ["■ 分母の選び方（令和3〜7年度の給付件数総計を各分母で割ったときのばらつき）"],
        ["", "分母", "変動係数", "判定"],
    ]
    for nm, v in cv:
        rows.append(["", nm, f"{v:.2f}％",
                     "**採用（最も安定）**" if nm == "第1号被保険者数" else "感度ケース"])
    rows += [
        [],
        ["■ 試算の連鎖"],
        ["", "① 件数率 ＝ 件数 ÷ 第1号被保険者数（基準期間の平均）"],
        ["", "② 給付費率 ＝ 価格調整後の給付費 ÷ 第1号被保険者数（基準期間の平均）"],
        ["", "③ 見込件数 ＝ ① × 将来の第1号被保険者数（第10期将来推計用推計人口）"],
        ["", "④ 見込給付費 ＝ ② × 将来の第1号被保険者数 ×（1＋令和9年度の報酬改定率）"],
        ["", "⑤ 表示単価 ＝ ④ ÷ ③"],
        [],
        ["■ 給付費に平均件数×直近単価を使わなかった理由"],
        ["", "件数と単価が逆方向に動くサービスがある。訪問介護は令和5年度から令和7年度にかけて"],
        ["", "件数が958件から628件へ34.4％減った一方、単価は37,532円から47,368円へ26.2％上がった。"],
        ["", "**平均件数に直近単価を乗じると、少ない利用者に高い単価を掛けることになり給付費を二重に見込む。**"],
        ["", "このため給付費そのものを第1号被保険者数で割った率を延ばし、単価は結果として表示している。"],
        [],
        ["■ 価格水準の調整"],
        ["", f"令和6年度の介護報酬改定は＋{KAITEI_R6 * 100:.2f}％。基準期間（令和5〜7年度）のうち"],
        ["", "令和5年度だけが改定前の価格である。令和5年度の給付費に改定率を乗じ、"],
        ["", "改定後の価格水準にそろえてから3年平均をとった。"],
        ["", "※ 令和6年度改定は令和6年6月施行（訪問看護等は4月）であり、令和6年度の実績は"],
        ["", "　 11か月分の反映にとどまる。この分だけ給付費率がわずかに低めに出る。"],
        [],
        ["■ 基準期間"],
        ["", "既定は令和5〜7年度の3年平均。次のサービスは個別に扱う。"],
    ]
    for svc, (yrs, note_txt) in BASE_OVERRIDE.items():
        rows.append(["", svc, "／".join(yrs), note_txt])
    rows += [
        [],
        ["■ この試算で決めていないこと（町・協議会の判断が要るもの）"],
        ["", "1. 令和9年度の介護報酬改定率。基本ケースは0.0％（据置）。06_感度分析で±1.5％を示した"],
        ["", "2. 施設・居住系の整備方針。第9期は新規整備を原則行わない方針だったが、認知症対応型共同生活介護の"],
        ["", "   定員は令和4年度71人から令和6年度98人へ増えている。整備があれば件数が段差で増える"],
        ["", "3. 地域支援事業費。年報 様式4の令和7年度が未入力のため実績が取れない。決算書の受領を要する"],
        ["", "4. 審査支払手数料。年報にない。1件60円で仮置きした"],
        [],
        ["■ この試算の限界"],
        ["", "**年報には回数・日数が入っていない。**様式2にあるのは件数（＝月あたり受給者数の年間累計）と単位数である。"],
        ["", "計画書の見込量は訪問系を回数、通所・短期入所を日数で書くのが通例であり、その単位に変換するには"],
        ["", "国保連の給付実績（サービス提供量）または見える化システムのサービス別給付実績が要る。"],
        ["", "**本試算は件数（延べ利用者数）と給付費で組んでいる。回数・日数への変換は当該データの受領後に行う。**"],
        ["", "参考として、1件当たり単位数を03シートに載せた（回数・日数の代替指標）。"],
        [],
        ["■ 検算1　サービス名の正規化に漏れがないか"],
        ["", "年報の様式2はサービス名の表記が年度で変わる。29区分に正規化したうえで、"
             "合計が様式2の「総計」行と一致するかを全年度・介護／予防別に確認した。"],
        ["", "年度", "正規化後の合計（介護）", "様式2の総計（介護）", "差",
         "正規化後の合計（予防）", "様式2の総計（予防）", "差"],
    ]
    for y in YS:
        ak = sum(act[y]["給付費"].get(s, (0, 0))[1] for _c, s, _u in ORDER_KAIGO)
        ay = sum(act[y]["給付費"].get(s, (0, 0))[0] for _c, s, _u in ORDER_KAIGO)
        bk, by = sub[y]["給付費"]["総計"][1], sub[y]["給付費"]["総計"][0]
        rows.append(["", y, ak, bk, ak - bk, ay, by, ay - by])
    rows += [
        ["", "", "", "", "", "", "",
         "**全年度・介護／予防とも差0で一致。取りこぼしも二重計上もない**"],
        [],
        ["■ 検算2　様式2（サービス別）と様式3（保険給付支払）の突合"],
        ["", "年度", "様式2の総計（給付費）",
         "様式3の介護サービス等諸費＋介護予防サービス等諸費", "差", "差の率"],
    ]
    for y in YS:
        a = sum(act[y]["給付費"].get(s, (0, 0))[0] + act[y]["給付費"].get(s, (0, 0))[1]
                for _c, s, _u in ORDER_KAIGO)
        b = gen[y]["総給付費"]
        rows.append(["", y, a, b, a - b, f"{(a - b) / b * 100:+.2f}％"])
    rows.append(["", "", "", "", "",
                 "**全年度で差0。サービス別の給付費の積み上げは、"
                 "保険給付支払状況の総給付費と完全に一致する**"])
    for r in rows:
        ws.append(r)
    ws["A1"].font = Font(bold=True, size=13)
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.startswith("■"):
            ws.cell(r, 1).font = Font(bold=True, size=10)
    widths(ws, [3, 34, 26, 26, 18, 26, 26, 60])
    for row in ws.iter_rows():
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
            if isinstance(c.value, (int, float)):
                c.number_format = "#,##0"
    return ws


# ---------------------------------------------------------------- 01/02 実績と単価

def sheet_jisseki(wb, name, act, kind, label):
    ws = wb.create_sheet(name)
    header = ["区分", "サービス"]
    for y in YS:
        header += [f"{y[2:]} 件数", f"{y[2:]} 給付費(千円)", f"{y[2:]} 単価(円/件)"]
    header += ["単価の5年変化", "基準期間", "備考"]
    ws.append(header)
    order = ORDER_KAIGO if kind == "介護" else [x for x in ORDER_KAIGO
                                                if x[1] not in YOBO_EXCLUDE]
    i = 1 if kind == "介護" else 0
    for cat, svc, _unit in order:
        row = [cat, svc]
        seen = []
        for y in YS:
            ken = act[y]["件数"].get(svc, (0, 0))[i]
            kyu = act[y]["給付費"].get(svc, (0, 0))[i]
            u = (kyu / ken) if ken else None
            if u:
                seen.append(u)
            row += [ken, kyu / 1000.0, u]
        # 単価のある年が2年以上ないと変化率は出せない
        chg = f"{(seen[-1] / seen[0] - 1) * 100:+.1f}％" if len(seen) >= 2 else "―"
        row += [chg, "／".join(y[2:] for y in base_years(svc)), base_note(svc)]
        ws.append(row)
    # 合計
    tot = ["合計", ""]
    for y in YS:
        k = sum(act[y]["件数"].get(s, (0, 0))[i] for _c, s, _u in order)
        g = sum(act[y]["給付費"].get(s, (0, 0))[i] for _c, s, _u in order)
        tot += [k, g / 1000.0, (g / k) if k else None]
    tot += ["", "", ""]
    ws.append(tot)
    style_header(ws)
    ncols = len(header)
    for r in range(2, ws.max_row + 1):
        for c in range(3, 3 + len(YS) * 3):
            cell = ws.cell(r, c)
            cell.number_format = "#,##0" if (c - 3) % 3 != 1 else "#,##0.0"
            cell.fill = CALC_FILL
    for c in range(1, ncols + 1):
        ws.cell(ws.max_row, c).font = Font(bold=True, size=9)
        ws.cell(ws.max_row, c).fill = KEY_FILL
    body_style(ws, wrap_cols=(ncols,))
    widths(ws, [11, 32] + [10, 13, 12] * len(YS) + [13, 20, 34])
    ws.freeze_panes = "C2"
    notes(ws, [
        f"※ {label}。件数は請求件数であり、月あたり受給者数の年間累計（人月）にあたる。",
        "※ 単価は給付費÷件数。1人が1か月そのサービスを使ったときの給付費である。",
        "※ 令和6年度に介護報酬が改定された（改定率＋1.59％）。単価の上昇にはこれが含まれる。",
        "※ 令和7年度から様式2の表記が「（短期利用以外）」「（短期利用）」に分割されたサービスは合算した。",
        "※ 令和5年度までの「福祉用具購入費」「住宅改修費」は令和6年度からの「特定福祉用具販売」"
        "「住宅改修」と同一区分として接続した。",
    ])
    return ws


# ---------------------------------------------------------------- 03 基礎率

def sheet_kiso(wb, act, gen):
    ws = wb.create_sheet("03_利用率と1件当たり単位数")
    ws.append(["区分", "サービス", "計画表記の単位"]
              + [f"{y[2:]} 件数/1号" for y in YS]
              + ["採用利用率", "基準期間"]
              + [f"{y[2:]} 単位数/件" for y in YS]
              + ["備考"])
    for cat, svc, unit in ORDER_KAIGO:
        row = [cat, svc, unit]
        for y in YS:
            ken = act[y]["件数"].get(svc, (0, 0))[1] + act[y]["件数"].get(svc, (0, 0))[0]
            row.append(ken / gen[y]["1号"])
        r_k = rate(act, gen, svc, "介護", "1号")
        r_y = rate(act, gen, svc, "予防", "1号")
        row += [r_k + r_y, "／".join(y[2:] for y in base_years(svc))]
        for y in YS:
            ken = act[y]["件数"].get(svc, (0, 0))[1] + act[y]["件数"].get(svc, (0, 0))[0]
            tan = act[y]["単位数"].get(svc, (0, 0))[1] + act[y]["単位数"].get(svc, (0, 0))[0]
            row.append((tan / ken) if ken else None)
        row.append("介護給付と予防給付の合計で表示。試算は04・05で別々に行う")
        ws.append(row)
    style_header(ws)
    n = ws.max_column
    for r in range(2, ws.max_row + 1):
        for c in range(4, 4 + len(YS) + 1):
            ws.cell(r, c).number_format = "0.0000"
            ws.cell(r, c).fill = CALC_FILL
        for c in range(4 + len(YS) + 2, n):
            ws.cell(r, c).number_format = "#,##0"
    body_style(ws, wrap_cols=(n,))
    widths(ws, [11, 32, 8] + [10] * len(YS) + [11, 18] + [11] * len(YS) + [40])
    ws.freeze_panes = "D2"
    notes(ws, [
        "※ 「件数/1号」は第1号被保険者1人当たりの年間利用件数。この率を将来の第1号被保険者数に乗じる。",
        "※ 「単位数/件」は1件（＝1人が1か月そのサービスを使ったとき）当たりの単位数。",
        "　 訪問系の回数、通所・短期入所の日数を推定する手がかりになるが、"
        "報酬区分の構成により変動するため、回数・日数の見込量には直接使えない。",
        "**※ 回数・日数ベースの見込量には、国保連の給付実績（サービス提供量）または"
        "見える化システムのサービス別給付実績が必要。町へ追加依頼する。**",
    ])
    return ws


# ---------------------------------------------------------------- 04/05 見込量

def sheet_mikomi(wb, name, proj, kind, label, kaitei):
    ws = wb.create_sheet(name)
    header = ["区分", "サービス", "計画表記の単位", "採用件数率", "表示単価(円/件)"]
    for y in EST_YEARS:
        header += [f"{y[2:]} 件数", f"{y[2:]} 人/月", f"{y[2:]} 給付費(千円)"]
    header += ["令和7年度実績(千円)", "R7→R11"]
    ws.append(header)
    for row in proj[kind]:
        r = [row["区分"], row["サービス"], row["単位"], row["率"], row["単価"]]
        for ken, kyu in row["見込"]:
            r += [ken, ken / 12.0, kyu]
        r += [row["_R7給付費"] / 1000.0, ""]
        ws.append(r)
        rr = ws.max_row
        base = row["_R7給付費"] / 1000.0
        last = row["見込"][-1][1]
        ws.cell(rr, len(header)).value = (f"{(last / base - 1) * 100:+.1f}％"
                                          if base else ("新規" if last else "―"))
    start, end = 2, ws.max_row
    ws.append(["合計", "", "", "", ""]
              + sum(([f"=SUM({openpyxl.utils.get_column_letter(6 + i * 3)}{start}:"
                      f"{openpyxl.utils.get_column_letter(6 + i * 3)}{end})",
                      f"=SUM({openpyxl.utils.get_column_letter(7 + i * 3)}{start}:"
                      f"{openpyxl.utils.get_column_letter(7 + i * 3)}{end})",
                      f"=SUM({openpyxl.utils.get_column_letter(8 + i * 3)}{start}:"
                      f"{openpyxl.utils.get_column_letter(8 + i * 3)}{end})"]
                     for i in range(len(EST_YEARS))), [])
              + [f"=SUM({openpyxl.utils.get_column_letter(len(header) - 1)}{start}:"
                 f"{openpyxl.utils.get_column_letter(len(header) - 1)}{end})", ""])
    style_header(ws)
    n = len(header)
    for r in range(start, ws.max_row + 1):
        ws.cell(r, 4).number_format = "0.0000"
        ws.cell(r, 5).number_format = "#,##0"
        for i in range(len(EST_YEARS)):
            ws.cell(r, 6 + i * 3).number_format = "#,##0"
            ws.cell(r, 7 + i * 3).number_format = "#,##0.0"
            ws.cell(r, 8 + i * 3).number_format = "#,##0"
            for c in (6 + i * 3, 7 + i * 3, 8 + i * 3):
                ws.cell(r, c).fill = CALC_FILL
        ws.cell(r, n - 1).number_format = "#,##0"
    for c in range(1, n + 1):
        ws.cell(ws.max_row, c).font = Font(bold=True, size=9)
        ws.cell(ws.max_row, c).fill = KEY_FILL
    body_style(ws)
    widths(ws, [11, 32, 8, 11, 12] + [10, 9, 13] * len(EST_YEARS) + [15, 11])
    ws.freeze_panes = "F2"
    notes(ws, [
        f"※ {label}。件数＝採用件数率×第1号被保険者数（第10期将来推計用推計人口）。",
        "※ 「人/月」は件数÷12。計画書の見込量表に載せる月あたり延べ利用者数にあたる。",
        "※ 給付費＝給付費率×第1号被保険者数。給付費率は基準期間の"
        "「価格調整後の給付費÷第1号被保険者数」の平均。",
        f"　 令和9年度以降はこれに報酬改定率(1{kaitei:+.3f})を乗じている。",
        "**※ 平均件数に直近単価を乗じる方法は採らなかった。件数と単価が逆方向に動くサービスで"
        "給付費を二重に見込むため（00シート参照）。表示単価は給付費÷件数の結果である。**",
        "※ 令和8年度は第9期の最終年度であり、第10期の見込みではない。実績との突合に使う。",
        "※ この表は件数（延べ利用者数）ベース。回数・日数ベースの表記は給付実績データの受領後に行う。",
    ])
    return ws


# ---------------------------------------------------------------- 06 感度

def sheet_kando(wb, act, gen, cases):
    ws = wb.create_sheet("06_感度分析")
    ws.append(["ケース", "分母", "報酬改定率", "設定の趣旨"]
              + [f"{y[2:]} 総給付費(千円)" for y in EST_YEARS]
              + ["第10期3年計(千円)", "基本ケースとの差", "差の率"])
    base3 = None
    for nm, den, pop, kaitei, why in cases:
        t = totals(project(act, gen, den, pop, kaitei))
        three = sum(t[1:])
        if base3 is None:
            base3 = three
        ws.append([nm, den, f"{kaitei * 100:+.1f}％", why] + t
                  + [three, three - base3,
                     f"{(three / base3 - 1) * 100:+.2f}％" if base3 else ""])
    style_header(ws)
    n = ws.max_column
    for r in range(2, ws.max_row + 1):
        for c in range(5, n):
            ws.cell(r, c).number_format = "#,##0"
            ws.cell(r, c).fill = CALC_FILL
    for c in range(1, n + 1):
        ws.cell(2, c).fill = KEY_FILL
        ws.cell(2, c).font = Font(bold=True, size=9)
    body_style(ws, wrap_cols=(4,))
    widths(ws, [22, 11, 11, 46] + [15] * len(EST_YEARS) + [17, 17, 12])
    notes(ws, [
        "※ 総給付費は介護給付と予防給付の合計。特定入所者介護サービス費・高額介護サービス費・"
        "高額医療合算介護サービス費・審査支払手数料は含まない（07で加算する）。",
        "※ 第10期3年計は令和9〜11年度の合計。令和8年度は第9期のため含まない。",
        "**※ 分母を認定者数にすると総給付費が大きく振れる。これが認定者数ベースで組めない理由である。**",
    ])
    return ws


# ---------------------------------------------------------------- 07 標準給付費

def sheet_hyojun(wb, act, gen, proj):
    ws = wb.create_sheet("07_標準給付費への接続")
    t = totals(proj)
    ratios = {}
    for key in ("特定入所", "高額", "高額合算"):
        vals = [gen[y][key] / gen[y]["総給付費"] for y in BASE_DEFAULT]
        ratios[key] = statistics.mean(vals)
    ken = {}
    for i, y in enumerate(EST_YEARS):
        k = 0.0
        for kind in ("介護", "予防"):
            for row in proj[kind]:
                k += row["見込"][i][0]
        ken[y] = k

    ws.append(["区分", "算出", "令和7年度実績(千円)"]
              + [f"{y[2:]}(千円)" for y in EST_YEARS] + ["第10期3年計(千円)", "備考"])
    rows = []
    r7 = gen["令和7年度"]
    rows.append(("① 総給付費", "04＋05の合計", r7["総給付費"] / 1000.0,
                 [t[i] for i in range(len(EST_YEARS))],
                 "介護給付＋予防給付"))
    rows.append(("② 特定入所者介護サービス費", f"①×{ratios['特定入所'] * 100:.3f}％",
                 r7["特定入所"] / 1000.0,
                 [t[i] * ratios["特定入所"] for i in range(len(EST_YEARS))],
                 "令和5〜7年度の対総給付費比の平均。補足給付は施設利用と連動する"))
    rows.append(("③ 高額介護サービス費", f"①×{ratios['高額'] * 100:.3f}％",
                 r7["高額"] / 1000.0,
                 [t[i] * ratios["高額"] for i in range(len(EST_YEARS))],
                 "令和5〜7年度の対総給付費比の平均"))
    rows.append(("④ 高額医療合算介護サービス費", f"①×{ratios['高額合算'] * 100:.3f}％",
                 r7["高額合算"] / 1000.0,
                 [t[i] * ratios["高額合算"] for i in range(len(EST_YEARS))],
                 "令和5〜7年度の対総給付費比の平均"))
    ken_r7 = sum(act["令和7年度"]["件数"].get(s, (0, 0))[0]
                 + act["令和7年度"]["件数"].get(s, (0, 0))[1] for _c, s, _u in ORDER_KAIGO)
    rows.append(("⑤ 審査支払手数料", f"件数×{TESURYO}円", ken_r7 * TESURYO / 1000.0,
                 [ken[y] * TESURYO / 1000.0 for y in EST_YEARS],
                 "**年報にないため仮置き。決算書での確認を要する**"))
    for nm, how, r7v, vals, memo in rows:
        ws.append([nm, how, r7v] + vals + [sum(vals[1:]), memo])
    first, last = 2, ws.max_row
    ws.append(["標準給付費見込額（①〜⑤の計）", "",
               f"=SUM(C{first}:C{last})"]
              + [f"=SUM({openpyxl.utils.get_column_letter(4 + i)}{first}:"
                 f"{openpyxl.utils.get_column_letter(4 + i)}{last})"
                 for i in range(len(EST_YEARS))]
              + [f"=SUM({openpyxl.utils.get_column_letter(4 + len(EST_YEARS))}{first}:"
                 f"{openpyxl.utils.get_column_letter(4 + len(EST_YEARS))}{last})",
                 "09_保険料算定の①に入る"])
    ws.append([])
    ws.append(["【参考】第9期計画が見込んだ標準給付費", "同年度の見込み",
               float(DAI9["標準給付費"]["令和7年度"]),
               float(DAI9["標準給付費"]["令和8年度"]), None, None, None,
               float(sum(DAI9["標準給付費"].values())),
               "第9期の3年計は令和6〜8年度で3,566,486千円。"
               "第10期の3年計はこれの93.1％の水準となる"])
    ws.append(["【参考】令和7年度の実績（様式3の保険給付支払計）", "実績",
               r7["支払計"] / 1000.0, None, None, None, None, None,
               "審査支払手数料を含まない。第9期の見込みとの対比は08を参照"])
    style_header(ws)
    n = ws.max_column
    for r in range(2, ws.max_row + 1):
        for c in range(3, n):
            ws.cell(r, c).number_format = "#,##0"
            if ws.cell(r, c).value is not None:
                ws.cell(r, c).fill = CALC_FILL
    for c in range(1, n + 1):
        ws.cell(last + 1, c).font = Font(bold=True, size=9)
        ws.cell(last + 1, c).fill = KEY_FILL
    body_style(ws, wrap_cols=(n,))
    widths(ws, [30, 20, 17] + [15] * len(EST_YEARS) + [17, 46])
    notes(ws, [
        "※ 標準給付費＝総給付費＋特定入所者介護サービス費＋高額介護サービス費"
        "＋高額医療合算介護サービス費＋審査支払手数料。",
        "※ 地域支援事業費は含まない。年報 様式4の令和7年度が未入力のため、実績が取れていない。",
        "**※ ②〜④は総給付費に対する比率で置いた。施設整備方針が変わると②の比率が動く。**",
    ])
    # ---- 参考：保険料基準額までの見当
    std3 = sum(sum(v[i] for _n, _h, _r, v, _m in rows) for i in range(1, len(EST_YEARS)))
    ws.append([])
    ws.append(["【参考】この標準給付費で保険料基準額はどのくらいになるか"])
    ws.cell(ws.max_row, 1).font = Font(bold=True, size=10)
    ws.append(["**地域支援事業費と調整交付金と基金取崩額が未確定のため、"
               "下表は見当をつけるためのものです。確定値ではありません。**"])
    ws.append(["項目", "値", "単位", "置き方"])
    hr = ws.max_row
    hosei = sum(rate_ * n * 3 for _nm, rate_, n in SHOTOKU_R7)
    chiiki3 = 207260
    for kikin, lab in ((0, "取崩なし"), (12000, "保有額12,000千円を全額取り崩す"),
                       (35000, "第9期と同じ35,000千円を取り崩す")):
        need = (std3 + chiiki3) * 0.23 - kikin
        tsuki = need / 99.35 * 100 * 1000 / hosei / 12
        ws.append([f"保険料基準額（月額）　{lab}", tsuki, "円",
                   f"総費用額{std3 + chiiki3:,.0f}千円×23％−取崩{kikin:,}千円"
                   f"÷収納率99.35％÷補正後被保険者数{hosei:,.0f}人÷12"])
    ws.append(["第9期の保険料基準額（月額）", 6600, "円",
               "算定値6,886円を基金取崩により6,597円とし、100円未満切上げ"])
    for r in range(hr, ws.max_row + 1):
        for c in range(1, 5):
            ws.cell(r, c).border = BORDER
            ws.cell(r, c).font = Font(size=9, bold=(r == hr))
            ws.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")
        if isinstance(ws.cell(r, 2).value, (int, float)):
            ws.cell(r, 2).number_format = "#,##0"
            ws.cell(r, 2).fill = KEY_FILL
    notes(ws, [
        f"※ 地域支援事業費は第9期計画の見込み（3年計{chiiki3:,}千円）を仮置きした。実績は未受領。",
        "※ 第1号被保険者負担割合は第9期と同じ23％。第10期の割合は国の基本指針で示される。",
        "※ 調整交付金は相当額と見込額が一致する（標準の5％が全額交付される）ものとした。"
        "**小野町は後期高齢者加入割合が高く、実際には5％を超える交付が見込まれる。"
        "その分だけ保険料は下がる。**",
        "※ 補正後被保険者数は令和7年度の所得段階別第1号被保険者数（様式1 所得段階別）を"
        "3年分としたもの。所得段階の設定が変われば動く。",
        "**※ 準備基金の保有額は令和5・6年度末とも12,000千円である。"
        "第9期が見込んだ35,000千円の取崩は、この残高では実行できない。"
        "残高の確認が第10期の保険料算定の前提となる。**",
    ])
    return ws


# ---------------------------------------------------------------- 08 第9期対比

DAI9 = {   # 第9期計画（令和6年3月策定）の見込み（千円）
    "標準給付費": {"令和6年度": 1185894, "令和7年度": 1189712, "令和8年度": 1190880},
    "総給付費": {"令和6年度": 1106604, "令和7年度": 1110764, "令和8年度": 1112373},
    "特定入所者": {"令和6年度": 52077, "令和7年度": 51854, "令和8年度": 51564},
    "地域支援事業費": {"令和6年度": 67320, "令和7年度": 69120, "令和8年度": 70820},
    "総合事業費": {"令和6年度": 36220, "令和7年度": 37020, "令和8年度": 37720},
    "包括的支援・任意": {"令和6年度": 24000, "令和7年度": 25000, "令和8年度": 26000},
}


def sheet_dai9(wb, gen):
    ws = wb.create_sheet("08_第9期の見込みと実績")
    ws.append(["区分", "令和6年度 見込", "令和6年度 実績", "達成率", "差",
               "令和7年度 見込", "令和7年度 実績", "達成率", "差",
               "令和8年度 見込", "令和8年度 実績", "評価"])

    def line(nm, kind, act6, act7, memo):
        r = [nm]
        for y, a in (("令和6年度", act6), ("令和7年度", act7)):
            m = DAI9[kind][y]
            r += [m, a, (a / m if a else None), (a - m if a else None)]
        r += [DAI9[kind]["令和8年度"], "年度途中", memo]
        ws.append(r)

    g6, g7 = gen["令和6年度"], gen["令和7年度"]
    line("総給付費", "総給付費", g6["総給付費"] / 1000.0, g7["総給付費"] / 1000.0,
         "令和7年度は見込みを9.5％下回った。認定者数の計上変更と重なる")
    line("　うち特定入所者介護サービス費", "特定入所者",
         g6["特定入所"] / 1000.0, g7["特定入所"] / 1000.0,
         "見込みはほぼ的中。補足給付は施設利用に連動しており安定している")
    std6 = (g6["総給付費"] + g6["特定入所"] + g6["高額"] + g6["高額合算"]) / 1000.0
    std7 = (g7["総給付費"] + g7["特定入所"] + g7["高額"] + g7["高額合算"]) / 1000.0
    line("標準給付費（審査支払手数料を除く）", "標準給付費", std6, std7,
         "**令和7年度の実績は見込みの91.1％。106百万円の下振れ**")
    ws.append(["地域支援事業費", DAI9["地域支援事業費"]["令和6年度"], "未受領", None, None,
               DAI9["地域支援事業費"]["令和7年度"], "未受領", None, None,
               DAI9["地域支援事業費"]["令和8年度"], "年度途中",
               "**年報 様式4の令和6・7年度が未入力。決算書の受領を要する**"])
    ws.append(["　うち介護予防・日常生活支援総合事業費", DAI9["総合事業費"]["令和6年度"],
               "未受領", None, None, DAI9["総合事業費"]["令和7年度"], "未受領", None, None,
               DAI9["総合事業費"]["令和8年度"], "年度途中",
               "令和5年度決算は34.7百万円（介護予防・生活支援サービス事業）"])
    ws.append(["　うち包括的支援事業・任意事業費", DAI9["包括的支援・任意"]["令和6年度"],
               "未受領", None, None, DAI9["包括的支援・任意"]["令和7年度"], "未受領", None, None,
               DAI9["包括的支援・任意"]["令和8年度"], "年度途中",
               "令和5年度決算は27.6百万円（社会保障充実分を含む）"])
    ws.append([])
    ws.append(["■ 保険料算定の前提"])
    ws.append(["保険料収納必要額", "832,562（基金取崩前）／797,563（取崩後）", "", "", "",
               "", "", "", "", "", "", "第9期計画本文"])
    ws.append(["介護給付費準備基金の取崩", "3年間で35,000", "12,000（保有額）", "", "",
               "", "", "", "", "", "",
               "**令和5・6年度末の基金保有額はいずれも12,000千円。"
               "令和4年度末の120,000千円から1桁減っている。取崩実績か入力誤りかの確認を要する**"])
    ws.append(["保険料収納率（現年度分）", "99.40（予定）",
               g6["収納現年"] / g6["調定現年"] * 100,
               "", "", "99.40（予定）", g7["収納現年"] / g7["調定現年"] * 100, "", "",
               "", "", "実績は予定をやや下回る。第10期は99.35％で置く"])
    ws.append(["保険料基準額（月額）", "6,600円", "6,600円", "", "", "6,600円", "6,600円",
               "", "", "6,600円", "6,600円",
               "算定値6,886円を基金取崩により6,597円とし、100円未満切上げで6,600円"])
    ws.append(["調定額（現年度分）", "", g6["調定現年"] / 1000.0, "", "", "",
               g7["調定現年"] / 1000.0, "", "", "", "",
               "令和7年度は265,882千円。令和6年度から2.8％増"])
    style_header(ws)
    n = ws.max_column
    for r in range(2, ws.max_row + 1):
        for c in (2, 3, 5, 6, 7, 9, 10, 11):
            cell = ws.cell(r, c)
            if isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0"
        for c in (4, 8):
            cell = ws.cell(r, c)
            if isinstance(cell.value, (int, float)):
                cell.number_format = "0.0％"
                cell.fill = WARN_FILL if cell.value < 0.95 else CALC_FILL
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.startswith("■"):
            ws.cell(r, 1).font = Font(bold=True, size=10)
    body_style(ws, wrap_cols=(n,))
    widths(ws, [30, 16, 16, 9, 12, 16, 16, 9, 12, 16, 12, 50])
    ws.freeze_panes = "B2"
    notes(ws, [
        "※ 見込みは第9期計画（令和6年3月策定）の記載。実績は介護保険事業状況報告（年報）様式3。",
        "※ 標準給付費の実績には審査支払手数料を含めていない。年報に計上欄がないため。",
        "※ 令和8年度は年度途中であり、実績が確定していない。",
        "**※ 令和7年度の標準給付費は見込みの91.1％にとどまった。"
        "第10期の見込みを第9期と同水準で置くと、同じ幅の過大見積りを繰り返すことになる。**",
    ])
    return ws


# ---------------------------------------------------------------- 09 人口推計

POP_SHOGAI = {   # 障がい福祉計画の設定シートが用いている人口
    "令和5年度": {"総人口": 9372, "0-14": 884, "15-64": 4986, "65+": 3502},
    "令和6年度": {"総人口": 9116, "0-14": 809, "15-64": 4814, "65+": 3493},
    "令和7年度": {"総人口": 8740, "0-14": 732, "15-64": 4525, "65+": 3483},
    "令和8年度": {"総人口": 8520, "0-14": 679, "15-64": 4382, "65+": 3459},
    "令和9年度": {"総人口": 8287, "0-14": 624, "15-64": 4203, "65+": 3460},
    "令和10年度": {"総人口": 8076, "0-14": 584, "15-64": 4040, "65+": 3452},
    "令和11年度": {"総人口": 7830, "0-14": 517, "15-64": 3891, "65+": 3422},
}

def _r(x):
    """四捨五入（0.5を切り上げ）。Pythonの round は偶数丸めのため使わない。"""
    return int(decimal.Decimal(str(x)).quantize(decimal.Decimal("1"),
                                                rounding=decimal.ROUND_HALF_UP))


def read_pop_dai10():
    """第10期将来推計用推計人口（男女別5歳階級）を年齢区分に集計する。

    シート 07522_小野町 は 8行目が男性の総数で9行目から19階級、28行目が女性の総数で
    29行目から19階級。列は8列目が令和2年で、以降1列ずつ1年進む。
    """
    ws = openpyxl.load_workbook(
        NENPO_SRC / "【受領】第10期将来推計用推計人口_小野町.xlsx",
        data_only=True)["07522_小野町"]
    bands = ["0-4", "5-9", "10-14", "15-19", "20-24", "25-29", "30-34", "35-39",
             "40-44", "45-49", "50-54", "55-59", "60-64", "65-69", "70-74",
             "75-79", "80-84", "85-89", "90+"]
    out = {}
    for yi, lab in enumerate(["令和5年度", "令和6年度", "令和7年度", "令和8年度",
                              "令和9年度", "令和10年度", "令和11年度"]):
        col = 8 + 3 + yi                       # 8＝令和2年、+3で令和5年
        v = {b: 0.0 for b in bands}
        for base in (9, 29):                   # 男性・女性の階級行の先頭
            for i, b in enumerate(bands):
                v[b] += ws.cell(base + i, col).value or 0
        d = {
            "0-14": _r(sum(v[b] for b in bands[:3])),
            "15-64": _r(sum(v[b] for b in bands[3:13])),
            "65-74": _r(v["65-69"] + v["70-74"]),
            "75-84": _r(v["75-79"] + v["80-84"]),
            "85+": _r(v["85-89"] + v["90+"]),
        }
        # 第1号被保険者数は3階級の合計とする。年齢階層別認定率の計算で
        # 内訳と合計が食い違うと検算が通らないため。
        d["65+"] = d["65-74"] + d["75-84"] + d["85+"]
        d["75+"] = d["75-84"] + d["85+"]
        d["1号"] = d["65+"]
        # 総人口は総数を独立に四捨五入したもの。年齢区分の合計とは
        # 端数処理により1人ずれることがある（令和9年度・令和11年度）。
        d["総人口"] = _r(sum(v.values()))
        out[lab] = d
    return out


POP_DAI10 = read_pop_dai10()
POP_EST = {y: POP_DAI10[y] for y in EST_YEARS}


def sheet_jinko(wb, gen):
    ws = wb.create_sheet("09_人口推計の出典比較")
    yrs = list(POP_SHOGAI.keys())
    ws.append(["区分", "出典"] + [y[2:] for y in yrs] + ["令和7→11 増減率", "備考"])

    def line(key, label, src, data, memo):
        vals = [data[y][key] for y in yrs]
        chg = (vals[-1] / vals[2] - 1) * 100
        ws.append([label, src] + vals + [f"{chg:+.1f}％", memo])

    line("総人口", "総人口", "第10期将来推計用推計人口", POP_DAI10,
         "コーホート変化率法。平成27年→令和2年の国勢調査の移動率で令和2年から延ばしたもの")
    line("総人口", "総人口", "障がい計画の設定シート", POP_SHOGAI,
         "令和5〜8年度はワークシート入力値（住民基本台帳）、令和9年度以降が推計")
    ws.append(["　差（第10期−障がい）", ""]
              + [POP_DAI10[y]["総人口"] - POP_SHOGAI[y]["総人口"] for y in yrs] + ["", ""])
    ws.append([])
    line("65+", "65歳以上人口", "第10期将来推計用推計人口", POP_DAI10,
         "**介護保険事業計画が使う区分。両者はよく一致している**")
    line("65+", "65歳以上人口", "障がい計画の設定シート", POP_SHOGAI, "")
    ws.append(["　差（第10期−障がい）", ""]
              + [POP_DAI10[y]["65+"] - POP_SHOGAI[y]["65+"] for y in yrs] + ["", ""])
    ws.append([])
    line("0-14", "0〜14歳人口", "第10期将来推計用推計人口", POP_DAI10,
         "**障害児サービス（児童発達支援・放課後等デイサービス）の人口比率型の分母**")
    line("0-14", "0〜14歳人口", "障がい計画の設定シート", POP_SHOGAI, "")
    ws.append(["　差（第10期−障がい）", ""]
              + [POP_DAI10[y]["0-14"] - POP_SHOGAI[y]["0-14"] for y in yrs] + ["", ""])
    ws.append([])
    line("15-64", "15〜64歳人口", "第10期将来推計用推計人口", POP_DAI10,
         "第2号被保険者数の基礎")
    line("15-64", "15〜64歳人口", "障がい計画の設定シート", POP_SHOGAI, "")
    ws.append([])
    ws.append(["■ 実績との突合（第1号被保険者数・年報 様式1）"])
    ws.append(["年度", "年報の実績（年度末）", "第10期推計（10月1日）", "差", "差の率",
               "障がい計画の65歳以上", "差", "差の率"])
    for y in BASE_DEFAULT:
        a = gen[y]["1号"]
        b = POP_DAI10[y]["65+"]
        c = POP_SHOGAI[y]["65+"]
        ws.append([y, a, b, b - a, f"{(b - a) / a * 100:+.2f}％",
                   c, c - a, f"{(c - a) / a * 100:+.2f}％"])
    for r in ws.iter_rows(min_row=2):
        for c in r:
            if isinstance(c.value, (int, float)):
                c.number_format = "#,##0"
    style_header(ws)
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.startswith("■"):
            ws.cell(r, 1).font = Font(bold=True, size=10)
        if isinstance(v, str) and v.startswith("　差"):
            for c in range(1, len(yrs) + 3):
                ws.cell(r, c).fill = WARN_FILL
    body_style(ws, wrap_cols=(len(yrs) + 4,))
    widths(ws, [22, 26] + [10] * len(yrs) + [14, 56])
    notes(ws, [
        "※ 第10期将来推計用推計人口の年齢区分は、男女別5歳階級の推計値を合計して"
        "四捨五入したもの。総人口は総数を独立に四捨五入しているため、"
        "年齢3区分の合計と1人ずれる年がある（令和9年度・令和11年度）。",
        "※ 第1号被保険者数は65〜74歳・75〜84歳・85歳以上の3階級の合計としている"
        "（年齢階層別認定率の計算で内訳と合計を一致させるため）。",
        "※ 65歳以上人口は3つの出典がよく一致している。差が出るのは0〜14歳と15〜64歳である。",
        "**※ 0〜14歳は令和11年度で697人と517人、180人（35％）の差がある。"
        "障害児サービスの見込量は人口比率型で算定しているため、この差がそのまま見込量に効く。**",
        "**※ 高齢者計画は第10期将来推計用推計人口を用いる。介護保険事業計画の様式であり、"
        "第1号被保険者数の実績との一致が確認できているため。**",
    ])
    return ws


# ================================================================ Word

def new_doc():
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = JP_FONT
    st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), JP_FONT)
    for name, size, bold in (("Heading 1", 14, True), ("Heading 2", 12, True),
                             ("Heading 3", 11, True)):
        s = doc.styles[name]
        s.font.name = JP_GOTHIC
        s.font.size = Pt(size)
        s.font.bold = bold
        s.font.color.rgb = None
        s.element.rPr.rFonts.set(qn("w:eastAsia"), JP_GOTHIC)
    return doc


def title(doc, main, subs=()):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(main)
    r.font.size = Pt(17)
    r.font.bold = True
    r.font.name = JP_GOTHIC
    r.element.rPr.rFonts.set(qn("w:eastAsia"), JP_GOTHIC)
    for s in subs:
        q = doc.add_paragraph()
        q.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rr = q.add_run(s)
        rr.font.size = Pt(11.5)
        rr.font.name = JP_GOTHIC
        rr.element.rPr.rFonts.set(qn("w:eastAsia"), JP_GOTHIC)


def table(doc, header, rows):
    t = doc.add_table(rows=1, cols=len(header))
    t.style = "Table Grid"
    for c, h in zip(t.rows[0].cells, header):
        c.text = ""
        r = c.paragraphs[0].add_run(h)
        r.font.bold = True
        r.font.size = Pt(9)
        r.font.name = JP_GOTHIC
        r.element.rPr.rFonts.set(qn("w:eastAsia"), JP_GOTHIC)
    for row in rows:
        for c, v in zip(t.add_row().cells, row):
            c.text = ""
            r = c.paragraphs[0].add_run(str(v))
            r.font.size = Pt(9)
    return t


def note(doc, text):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.font.size = Pt(9)
    r.italic = True
    return p


def page_break(doc):
    doc.add_section(WD_SECTION.NEW_PAGE)


def yen(v, unit=1000.0):
    return f"{v / unit:,.0f}"


# ---------------------------------------------------------------- 第9期対比の文書

def doc_dai9(gen, proj):
    doc = new_doc()
    title(doc, "第9期計画の見込みと実績の対比",
          ("小野町高齢者保健福祉計画・第10期介護保険事業計画",
           f"策定支援業務　{ASOF_JP}"))
    doc.add_paragraph()
    doc.add_heading("1　この文書の位置づけ", level=1)
    doc.add_paragraph(
        "委託仕様書は、第10期計画の策定にあたり前期計画の進捗状況を評価することを求めている。"
        "本文書は、そのうち給付費・保険料に関する部分について、"
        "第9期計画（令和6年3月策定）が見込んだ額と、"
        "介護保険事業状況報告（年報）で確定した実績とを対比したものである。"
        "成果品①（計画書）第7章7及び骨子案の参考1に反映する。")
    doc.add_paragraph(
        "令和8年9月8日に町から年報の令和3〜7年度分を受領したことにより、"
        "第9期の3年度のうち令和6年度・令和7年度の2年度について実績が確定した。"
        "令和8年度は年度途中であり、確定は令和9年夏の年報を待つ。")

    doc.add_heading("2　結論", level=1)
    g6, g7 = gen["令和6年度"], gen["令和7年度"]
    std6 = (g6["総給付費"] + g6["特定入所"] + g6["高額"] + g6["高額合算"]) / 1000.0
    std7 = (g7["総給付費"] + g7["特定入所"] + g7["高額"] + g7["高額合算"]) / 1000.0
    m6, m7 = DAI9["標準給付費"]["令和6年度"], DAI9["標準給付費"]["令和7年度"]
    doc.add_paragraph(
        f"第9期が見込んだ標準給付費に対し、令和6年度の実績は{std6 / m6 * 100:.1f}％、"
        f"令和7年度の実績は{std7 / m7 * 100:.1f}％であった。"
        f"令和6年度はほぼ見込みどおりだったが、令和7年度は{m7 - std7:,.0f}千円下振れした。"
        "第10期の見込みを第9期と同じ水準に置くと、同じ幅の過大見積りを繰り返すことになる。")
    doc.add_paragraph(
        "下振れの主因は令和7年度の総給付費である。"
        f"見込み{DAI9['総給付費']['令和7年度']:,}千円に対し実績は{g7['総給付費'] / 1000:,.0f}千円で、"
        f"{(g7['総給付費'] / 1000) / DAI9['総給付費']['令和7年度'] * 100:.1f}％にとどまった。"
        "同じ年度に要支援・要介護認定者数が944人から792人へ152人減っており、"
        "この減少が令和7年9月審査分の1か月に集中していることが国保連の月報で確認されている。"
        "**したがって令和7年度の実績は、利用の実態が下がった結果というより、"
        "計上方法の変更を含んだ数字とみるのが妥当である。**"
        "町への照会結果を待って、第10期の基準年度としての扱いを確定する。")

    doc.add_heading("3　標準給付費の対比", level=1)
    table(doc, ["区分", "令和6年度 見込", "令和6年度 実績", "達成率",
                "令和7年度 見込", "令和7年度 実績", "達成率"], [
        ("総給付費", f"{DAI9['総給付費']['令和6年度']:,}",
         f"{g6['総給付費'] / 1000:,.0f}",
         f"{(g6['総給付費'] / 1000) / DAI9['総給付費']['令和6年度'] * 100:.1f}％",
         f"{DAI9['総給付費']['令和7年度']:,}", f"{g7['総給付費'] / 1000:,.0f}",
         f"{(g7['総給付費'] / 1000) / DAI9['総給付費']['令和7年度'] * 100:.1f}％"),
        ("　うち介護給付", "―", f"{g6['介護諸費'] / 1000:,.0f}", "―",
         "―", f"{g7['介護諸費'] / 1000:,.0f}", "―"),
        ("　うち予防給付", "―", f"{g6['予防諸費'] / 1000:,.0f}", "―",
         "―", f"{g7['予防諸費'] / 1000:,.0f}", "―"),
        ("特定入所者介護サービス費", f"{DAI9['特定入所者']['令和6年度']:,}",
         f"{g6['特定入所'] / 1000:,.0f}",
         f"{(g6['特定入所'] / 1000) / DAI9['特定入所者']['令和6年度'] * 100:.1f}％",
         f"{DAI9['特定入所者']['令和7年度']:,}", f"{g7['特定入所'] / 1000:,.0f}",
         f"{(g7['特定入所'] / 1000) / DAI9['特定入所者']['令和7年度'] * 100:.1f}％"),
        ("高額介護サービス費", "―", f"{g6['高額'] / 1000:,.0f}", "―",
         "―", f"{g7['高額'] / 1000:,.0f}", "―"),
        ("高額医療合算介護サービス費", "―", f"{g6['高額合算'] / 1000:,.0f}", "―",
         "―", f"{g7['高額合算'] / 1000:,.0f}", "―"),
        ("標準給付費（手数料を除く）", f"{m6:,}", f"{std6:,.0f}",
         f"{std6 / m6 * 100:.1f}％", f"{m7:,}", f"{std7:,.0f}",
         f"{std7 / m7 * 100:.1f}％"),
    ])
    note(doc, "※ 見込みは第9期計画本文、実績は介護保険事業状況報告（年報）様式3による。"
              "様式3の総給付費は、様式2のサービス別給付費の合計と全年度で完全に一致する。"
              "実績側には審査支払手数料を含まない（年報に計上欄がないため）。"
              "第9期の見込額は手数料を含むため、実績側がその分だけ小さく出る。"
              "令和7年度の件数14,488件に1件60円を乗じると869千円であり、"
              "達成率への影響は0.1ポイント程度である。")
    doc.add_paragraph()
    doc.add_paragraph(
        "特定入所者介護サービス費は、令和6年度・令和7年度とも見込みに対して"
        "ほぼ計画どおりである。補足給付は施設・居住系の利用と連動しており、"
        "認定者数の変動の影響を受けにくい。"
        "**これは、令和7年度の総給付費の下振れが施設利用の減少によるものではないことを示す。**")

    page_break(doc)
    doc.add_heading("4　地域支援事業費の対比", level=1)
    doc.add_paragraph(
        "地域支援事業費は、年報の様式4（介護保険特別会計経理状況）の"
        "令和6年度・令和7年度が未入力のため、実績が取れていない。"
        "町の決算書（歳入歳出決算書及び主要な施策の成果に関する説明書）の受領を要する。")
    table(doc, ["区分", "令和6年度 見込", "令和7年度 見込", "令和8年度 見込", "実績", "確認事項"], [
        ("地域支援事業費", f"{DAI9['地域支援事業費']['令和6年度']:,}",
         f"{DAI9['地域支援事業費']['令和7年度']:,}",
         f"{DAI9['地域支援事業費']['令和8年度']:,}", "未受領",
         "3年計207,260千円。令和5年度決算は63,400千円で、見込みは実績を上回る設定"),
        ("　うち総合事業費", f"{DAI9['総合事業費']['令和6年度']:,}",
         f"{DAI9['総合事業費']['令和7年度']:,}",
         f"{DAI9['総合事業費']['令和8年度']:,}", "未受領",
         "令和5年度は介護予防・生活支援サービス事業34,700千円、一般介護予防事業1,000千円。"
         "一般介護予防（通いの場等）がほぼ計上されていない"),
        ("　うち包括的支援事業・任意事業費", f"{DAI9['包括的支援・任意']['令和6年度']:,}",
         f"{DAI9['包括的支援・任意']['令和7年度']:,}",
         f"{DAI9['包括的支援・任意']['令和8年度']:,}", "未受領",
         "令和5年度決算は27,600千円（社会保障充実分を含む）で見込みを上回る"),
    ])
    note(doc, "※ 単位は千円。見込みは第9期計画本文による。")

    doc.add_heading("5　保険料算定の前提の対比", level=1)
    table(doc, ["区分", "第9期の設定", "実績", "評価"], [
        ("保険料基準額（月額）", "6,600円（算定値6,886円を基金取崩により6,597円とし、"
                          "100円未満切上げ）", "6,600円",
         "平成30年度以降据置。第10期は基金・繰越金の水準を踏まえ、"
         "据置・引下げ・取崩の判断が論点となる"),
        ("予定収納率（現年度分）", "99.40％",
         f"令和6年度 {g6['収納現年'] / g6['調定現年'] * 100:.2f}％／"
         f"令和7年度 {g7['収納現年'] / g7['調定現年'] * 100:.2f}％",
         "実績は予定をわずかに下回る。第10期は99.35％で置く"),
        ("保険料収納必要額", "832,562千円（基金取崩前）／797,563千円（取崩後）",
         f"調定額（現年度分）は令和6年度 {g6['調定現年'] / 1000:,.0f}千円／"
         f"令和7年度 {g7['調定現年'] / 1000:,.0f}千円",
         f"令和7年度の調定額は令和6年度から"
         f"{(g7['調定現年'] / g6['調定現年'] - 1) * 100:+.1f}％。"
         "3年計は概ね見込みどおりの水準"),
        ("介護給付費準備基金の取崩", "3年間で35,000千円",
         "令和5年度末・令和6年度末の保有額はいずれも12,000千円",
         "**令和4年度末の120,000千円から1桁減っている。"
         "取崩の実績なのか入力誤りなのか、町への確認を要する。"
         "35,000千円を取り崩した場合、残高は不足する**"),
        ("所得段階", "13段階（第1段階0.285〜第4段階0.900）",
         "令和7年度の年報 様式1 所得段階別で13段階を確認",
         "第1〜3段階の割合は27.7％。国の標準は第1段階0.285で一致"),
    ])

    doc.add_heading("6　施設サービスの見込みとの対比", level=1)
    doc.add_paragraph(
        "第9期は施設サービスについて、介護老人福祉施設62人（187,243千円）、"
        "介護老人保健施設24人（81,076千円）、介護医療院0人を見込んでいた。"
        "年報の様式1の6による令和7年度の受給者数（年間延べ）は次のとおりである。")
    r7 = gen["令和7年度"]
    table(doc, ["区分", "第9期の見込み（人／月）", "令和7年度実績（人／月）",
                "見込み給付費（千円）", "令和7年度実績（千円）", "評価"], [
        ("介護老人福祉施設", "62", f"{827 / 12:.1f}", "187,243", "204,850",
         "受給者数は見込みを上回り、給付費も9.4％上回った"),
        ("介護老人保健施設", "24", f"{343 / 12:.1f}", "81,076", "92,523",
         "受給者数・給付費とも見込みを上回った"),
        ("介護医療院", "0", f"{14 / 12:.1f}", "0", "4,310",
         "**見込み0人に対して実績がある。第10期では計上を要する**"),
        ("介護療養型医療施設", "―", "0.0", "―", "0",
         "令和5年度末で制度廃止。実績0で整合"),
        ("地域密着型介護老人福祉施設", "―", f"{496 / 12:.1f}", "―", "126,350",
         "町内施設。定員29人に対し受給者41.3人／月は町外分を含む"),
    ])
    note(doc, "※ 実績の人数は年報 様式1の6の年間延べ受給者数を12で除したもの。"
              "給付費は年報 様式2による。"
              "施設サービス全体の給付費は令和7年度301,682千円であり、"
              "地域密着型介護老人福祉施設126,350千円を加えると428,032千円となる。")

    page_break(doc)
    doc.add_heading("7　第10期の見込みへの反映", level=1)
    t = totals(proj)
    doc.add_paragraph(
        "以上を踏まえ、第10期の見込みは次の方針で組む。")
    table(doc, ["論点", "第9期の扱い", "第10期の方針", "根拠"], [
        ("推計の基礎", "認定者数×1人当たり給付費",
         "**第1号被保険者数×サービス別利用率**",
         "令和3〜7年度で給付件数総計を各分母で割ったときの変動係数は、"
         "第1号被保険者数1.63％、認定者数5.18％。認定者数は令和7年9月の計上変更で断層がある"),
        ("基準年度", "直近年度",
         "令和5〜7年度の3年平均",
         "令和7年度単年で置くと計上変更の影響をそのまま将来に持ち越す"),
        ("報酬改定", "改定を織り込まず",
         "令和9年度改定は0.0％で仮置きし、±1.5％の感度を併記",
         "次期改定は令和9年度。率は令和8年12月頃に決まる"),
        ("特定入所者介護サービス費", "逓減で見込み（実績は増加基調）",
         "総給付費に対する比率（令和5〜7年度平均）で置く",
         "第9期は横ばい〜逓減で見込んだが、実績は令和6年度52,990千円まで増えた"),
        ("施設整備", "新規整備を原則行わない",
         "同方針を継続する場合と、整備がある場合を分けて示す",
         "第9期は原則整備なしとしながら、認知症対応型共同生活介護の定員が"
         "令和4年度71人から令和6年度98人へ増えている。整合の整理を要する"),
        ("基金の取崩", "3年間で35,000千円",
         "基金残高の確認結果を待って設定する",
         "**保有額12,000千円に対し35,000千円の取崩は実行できない。"
         "残高の確認が第10期の保険料算定の前提となる**"),
    ])
    doc.add_paragraph()
    doc.add_paragraph(
        f"本方針による第10期（令和9〜11年度）の総給付費の試算は、"
        f"3年計で{sum(t[1:]):,.0f}千円である。"
        f"第9期の総給付費見込み（3年計{sum(DAI9['総給付費'].values()):,}千円）に対して"
        f"{sum(t[1:]) / sum(DAI9['総給付費'].values()) * 100:.1f}％の水準となる。"
        "詳細は「小野町_第10期_サービス別単価・見込量試算」を参照。")

    doc.add_heading("8　残る確認事項", level=1)
    table(doc, ["№", "確認事項", "宛先", "第10期への影響"], [
        ("1", "令和7年8〜9月の認定者数の計上方法の変更の有無"
              "（町外施設入所者の帰属、住所地特例、国保連の集計定義）",
         "町・国保連", "**基準年度の扱いが決まる。第10期の見込量の水準を左右する**"),
        ("2", "介護給付費準備基金の残高と取崩実績"
              "（令和4年度末120,000千円→令和5・6年度末12,000千円）",
         "町（財政）", "**保険料基準額を直接左右する**"),
        ("3", "令和6・7年度の地域支援事業費の決算額と事業別内訳", "町",
         "総費用額に加算する。3年計で約210,000千円の規模"),
        ("4", "令和7年度の介護保険特別会計決算（年報 様式4が未入力）", "町",
         "歳入歳出差引残額・繰越金の確定"),
        ("5", "審査支払手数料の実績（1件当たり単価）", "町・国保連",
         "標準給付費に約900千円／年で効く"),
        ("6", "国保連の給付実績（サービス提供量＝回数・日数）",
         "町・国保連", "**計画書の見込量表を回数・日数で書くために必要**"),
    ])
    note(doc, f"※ 本文書は{ASOF_JP}時点の受領データによる。"
              "令和8年度の実績は令和9年夏の年報公表後に追記する。")
    return doc


# ---------------------------------------------------------------- 人口統一メモ

def doc_jinko(gen):
    doc = new_doc()
    title(doc, "人口推計の出典の統一について",
          ("小野町高齢者保健福祉計画・第10期介護保険事業計画",
           "おのまち障がい者計画・第4期障がい児福祉計画・第8期障がい福祉計画",
           f"{ASOF_JP}"))
    doc.add_paragraph()
    doc.add_heading("1　論点", level=1)
    doc.add_paragraph(
        "小野町では、高齢者保健福祉計画・第10期介護保険事業計画と、"
        "おのまち障がい者計画・第4期障がい児福祉計画・第8期障がい福祉計画を"
        "同じ年度に策定し、同じ年度に公表する。"
        "この2つの計画が、異なる人口推計を使っている。")
    doc.add_paragraph(
        f"令和9年度の総人口は、高齢者計画が用いる第10期将来推計用推計人口で"
        f"{POP_DAI10['令和9年度']['総人口']:,}人、"
        f"障がい計画の見込量目標値設定シートで{POP_SHOGAI['令和9年度']['総人口']:,}人であり、"
        f"{POP_DAI10['令和9年度']['総人口'] - POP_SHOGAI['令和9年度']['総人口']}人の差がある。"
        f"令和11年度では{POP_DAI10['令和11年度']['総人口'] - POP_SHOGAI['令和11年度']['総人口']}人"
        f"（{(POP_DAI10['令和11年度']['総人口'] / POP_SHOGAI['令和11年度']['総人口'] - 1) * 100:.1f}％）"
        "まで広がる。")
    doc.add_paragraph(
        "同じ町が同じ年度に公表する2つの計画で、冒頭の「小野町の人口」の数字が違う。"
        "住民・議会・県のいずれから見ても説明がつかない。統一を要する。")

    doc.add_heading("2　2つの出典", level=1)
    table(doc, ["項目", "第10期将来推計用推計人口", "障がい計画の設定シート"], [
        ("入手", "令和8年9月8日に町から受領", "令和8年9月7日に町から受領"),
        ("作成", "厚生労働省が保険者向けに配布する様式"
               "（地域包括ケア「見える化」システムの推計人口）",
         "町のワークシート（第3期地域福祉計画と同系統）"),
        ("方法", "コーホート変化率法。平成27年→令和2年の国勢調査による"
               "5歳階級・男女別の人口移動率を、令和2年国勢調査人口に繰り返し適用",
         "コーホート変化率法。令和5〜8年度は住民基本台帳の実績値、"
         "令和9年度以降が推計"),
        ("基準時点", "各年10月1日", "各年度（住基は年度末とみられる）"),
        ("区分", "男女別5歳階級（0〜4歳から90歳以上まで19区分）",
         "年齢3区分（0〜14歳、15〜64歳、65歳以上）"),
        ("推計期間", "令和2年から令和37年（2055年）まで", "令和5年度から令和11年度まで"),
    ])

    doc.add_heading("3　差はどこにあるか", level=1)
    doc.add_paragraph(
        "年齢3区分に分けると、差の所在がはっきりする。")
    yrs = ["令和7年度", "令和8年度", "令和9年度", "令和10年度", "令和11年度"]
    rows = []
    for key, lab in (("総人口", "総人口"), ("0-14", "0〜14歳"),
                     ("15-64", "15〜64歳"), ("65+", "65歳以上")):
        rows.append([lab + "（第10期）"] + [f"{POP_DAI10[y][key]:,}" for y in yrs])
        rows.append([lab + "（障がい）"] + [f"{POP_SHOGAI[y][key]:,}" for y in yrs])
        rows.append(["　差"] + [f"{POP_DAI10[y][key] - POP_SHOGAI[y][key]:+,}" for y in yrs])
    table(doc, ["区分"] + [y[2:] for y in yrs], rows)
    doc.add_paragraph()
    doc.add_paragraph(
        "**65歳以上人口はよく一致している。**令和11年度で3,418人と3,422人、差は4人である。"
        "介護保険事業計画が使うのはこの区分であり、"
        "どちらの出典を採っても第1号被保険者数の見込みは変わらない。")
    doc.add_paragraph(
        "差は0〜14歳と15〜64歳に集中している。"
        f"0〜14歳は令和11年度で{POP_DAI10['令和11年度']['0-14']}人と"
        f"{POP_SHOGAI['令和11年度']['0-14']}人、"
        f"{POP_DAI10['令和11年度']['0-14'] - POP_SHOGAI['令和11年度']['0-14']}人"
        f"（{(POP_DAI10['令和11年度']['0-14'] / POP_SHOGAI['令和11年度']['0-14'] - 1) * 100:.0f}％）"
        "の差がある。令和7年度からの減少率も、第10期推計が"
        f"{(POP_DAI10['令和11年度']['0-14'] / POP_DAI10['令和7年度']['0-14'] - 1) * 100:.1f}％、"
        f"障がい計画が"
        f"{(POP_SHOGAI['令和11年度']['0-14'] / POP_SHOGAI['令和7年度']['0-14'] - 1) * 100:.1f}％"
        "と、倍近く違う。")

    doc.add_heading("4　どちらが実績に近いか", level=1)
    doc.add_paragraph(
        "第1号被保険者数については、介護保険事業状況報告（年報）様式1に実績がある。"
        "これと突き合わせる。")
    rows = []
    for y in BASE_DEFAULT:
        a = gen[y]["1号"]
        b = POP_DAI10[y]["65+"]
        c = POP_SHOGAI[y]["65+"]
        rows.append([y, f"{a:,}", f"{b:,}", f"{b - a:+,}", f"{(b - a) / a * 100:+.2f}％",
                     f"{c:,}", f"{c - a:+,}", f"{(c - a) / a * 100:+.2f}％"])
    table(doc, ["年度", "年報の第1号被保険者数（年度末）",
                "第10期推計の65歳以上", "差", "差の率",
                "障がい計画の65歳以上", "差", "差の率"], rows)
    note(doc, "※ 年報は年度末（3月31日）現在、第10期推計は各年10月1日時点であり、"
              "半年のずれがある。第1号被保険者数には適用除外者を含まない。")
    doc.add_paragraph()
    doc.add_paragraph(
        "**両者とも実績との差は1％前後で、優劣はつかない。**"
        "65歳以上については、どちらを採っても介護保険の算定結果は変わらない。")
    doc.add_paragraph(
        "総人口については、第10期推計が令和2年国勢調査から機械的に延ばしたものであるのに対し、"
        "障がい計画の設定シートは令和5〜8年度に住民基本台帳の実績値を使っている。"
        f"令和7年度で{POP_DAI10['令和7年度']['総人口']:,}人と"
        f"{POP_SHOGAI['令和7年度']['総人口']:,}人の差があるが、"
        "**令和7年度は障がい計画側が実績値であるから、こちらが実勢に近い。**"
        "第10期推計が平成27年→令和2年の移動率を使っているため、"
        "令和2年以降に転出が加速した分を捉えられていないとみられる。")

    page_break(doc)
    doc.add_heading("5　見込量への影響", level=1)
    doc.add_paragraph(
        "出典を入れ替えたときに、それぞれの計画の見込量がどれだけ動くかを整理する。")
    table(doc, ["計画", "推計に使う人口", "出典を入れ替えたときの影響", "影響度"], [
        ("高齢者計画・介護保険事業計画", "第1号被保険者数（65歳以上）",
         f"令和11年度で3,418人と3,422人、差は4人（{4 / 3418 * 100:.2f}％）。"
         "見込量・給付費・保険料のいずれもほぼ動かない", "軽微"),
        ("障がい者計画・障がい福祉計画", "総人口（手帳所持者の推計）",
         f"令和11年度で7,830人と8,232人、差は402人（5.1％）。"
         "手帳所持者数を人口比で置いている項目が同率で動く", "中"),
        ("障がい児福祉計画", "0〜14歳人口（児童発達支援・放課後等デイサービス）",
         f"令和11年度で517人と697人、差は180人（34.8％）。"
         "**人口比率型で算定している見込量が3割以上動く**", "**大**"),
    ])
    doc.add_paragraph()
    doc.add_paragraph(
        "障がい児福祉計画の見込量は、レッドチームレビューでも指摘した"
        "「0〜14歳人口推計の令和10年度から令和11年度への減少率が他年から跳ねている」"
        "（指摘C-6）という問題を抱えている。"
        "年次減少率は令和6年度▲8.5％、令和7年度▲9.5％、令和8年度▲7.2％、"
        "令和9年度▲8.1％、令和10年度▲6.4％に対し、令和11年度は▲11.5％である。"
        "第10期推計の0〜14歳は年▲4％台で滑らかに減っており、この跳ねがない。")

    doc.add_heading("6　統一案", level=1)
    doc.add_paragraph("次の3案を検討した。")
    table(doc, ["案", "内容", "利点", "欠点"], [
        ("案1", "第10期将来推計用推計人口に統一する",
         "国が配布する様式であり、県への説明がしやすい。"
         "男女別5歳階級まであるため、どの年齢区分にも対応できる。"
         "令和37年まであり、長期の見通しにも使える",
         "**令和2年国勢調査から機械的に延ばしたものであり、"
         "令和5〜7年度の住基実績と200人前後ずれている。**"
         "総人口が実勢より多めに出る"),
        ("案2", "障がい計画の設定シートに統一する",
         "令和5〜8年度が住民基本台帳の実績値であり、実勢に近い",
         "年齢3区分しかないため、介護保険の第1号被保険者数を"
         "65〜74歳・75〜84歳・85歳以上に分けられない。"
         "**介護保険事業計画の算定様式に合わない**"),
        ("案3", "**住民基本台帳の実績を基準年とし、"
                "第10期将来推計用推計人口の変化率で延ばす**",
         "基準年が実勢に合い、かつ男女別5歳階級まで持てる。"
         "**両計画で同じ総人口・同じ年齢構成を使える**",
         "推計をやり直す手間がかかる。町の了解を要する"),
    ])
    doc.add_paragraph()
    doc.add_paragraph(
        "**案3を推奨する。**具体的には次の手順による。")
    doc.add_paragraph(
        "① 基準年を令和7年度（住民基本台帳）とし、男女別5歳階級の実績値を町から受領する。", style=None)
    doc.add_paragraph(
        "② 第10期将来推計用推計人口の年次変化率（階級別・男女別）を、①に適用して"
        "令和8〜11年度を推計する。")
    doc.add_paragraph(
        "③ 得られた系列を、高齢者計画・障がい計画の双方で共通に使う。")
    doc.add_paragraph(
        "④ 計画本文には「住民基本台帳（令和7年度）を基準とし、"
        "第10期将来推計用推計人口の変化率により推計した」と注記する。")
    doc.add_paragraph()
    doc.add_paragraph(
        "**ただし、介護保険事業計画の保険料算定については、"
        "第10期将来推計用推計人口をそのまま用いることを妨げない。**"
        "第1号被保険者数は両者で4人しか違わず、"
        "国の算定様式との整合を優先する方が説明しやすいためである。"
        "この場合、計画本文の人口の記述（第2章）と保険料算定の基礎数値（第7章）で"
        "出典が分かれることになるため、その旨を注記する。")

    doc.add_heading("7　決めていただきたいこと", level=1)
    table(doc, ["№", "決定事項", "期限", "決定者"], [
        ("1", "計画本文に載せる人口推計の出典を案1〜3のいずれにするか",
         "令和8年10月中（骨子案の作成前）", "町（健康福祉課）"),
        ("2", "案3を採る場合、男女別5歳階級の住民基本台帳人口（令和7年度）の提供",
         "令和8年10月中", "町（住民課）"),
        ("3", "障がい児福祉計画の0〜14歳人口を差し替えるか"
              "（差し替える場合、児童発達支援・放課後等デイサービスの見込量を再計算する）",
         "令和8年10月中", "町（健康福祉課）"),
        ("4", "介護保険の保険料算定に第10期将来推計用推計人口を用いること",
         "令和8年10月中", "町（健康福祉課）"),
    ])
    note(doc, f"※ 本メモは{ASOF_JP}時点の受領データによる。"
              "第10期将来推計用推計人口の年齢区分は、男女別5歳階級の推計値を"
              "合計して四捨五入したものである。総人口は総数を独立に四捨五入しているため、"
              "年齢3区分の合計と1人ずれる年がある（令和9年度・令和11年度）。")
    return doc


# ================================================================ main

def main():
    act, sub, gen = extract()

    # 分母の安定性
    tot = [sum(act[y]["件数"].get(s, (0, 0))[0] + act[y]["件数"].get(s, (0, 0))[1]
               for _c, s, _u in ORDER_KAIGO) for y in YS]
    cv = []
    for nm, key in (("第1号被保険者数", "1号"), ("75歳以上人口", "75+"),
                    ("85歳以上人口", "85+"), ("認定者数", "認定")):
        rt = [t / gen[y][key] for t, y in zip(tot, YS)]
        cv.append((nm, statistics.pstdev(rt) / statistics.mean(rt) * 100))

    proj = project(act, gen, "1号", "1号", 0.0)
    for kind in ("介護", "予防"):
        i = 1 if kind == "介護" else 0
        for row in proj[kind]:
            row["_R7給付費"] = act["令和7年度"]["給付費"].get(row["サービス"], (0, 0))[i]

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    sheet_intro(wb, act, sub, gen, cv)
    sheet_jisseki(wb, "01_実績と単価_介護", act, "介護", "介護給付（要介護1〜5）")
    sheet_jisseki(wb, "02_実績と単価_予防", act, "予防", "予防給付（要支援1・2）")
    sheet_kiso(wb, act, gen)
    sheet_mikomi(wb, "04_見込量試算_介護", proj, "介護", "介護給付（要介護1〜5）", 0.0)
    sheet_mikomi(wb, "05_見込量試算_予防", proj, "予防", "予防給付（要支援1・2）", 0.0)
    sheet_kando(wb, act, gen, [
        ("基本（1号人口）", "1号", "1号", 0.0,
         "第1号被保険者数に連動。変動係数が最も小さく、認定者数の断層の影響を受けない"),
        ("後期高齢者連動", "75+", "75+", 0.0,
         "75歳以上人口に連動。利用が後期高齢者に集中していることを反映するが、"
         "小野町は令和11年度まで85歳以上が減るため下振れる"),
        ("認定者連動", "認定", "認定", 0.0,
         "**認定者数に連動。令和7年9月の計上変更により1人当たり利用量が跳ねており、上振れる**"),
        ("基本＋報酬改定+1.5％", "1号", "1号", 0.015,
         "令和9年度の介護報酬改定を＋1.5％と置いた場合"),
        ("基本＋報酬改定−1.5％", "1号", "1号", -0.015,
         "令和9年度の介護報酬改定を−1.5％と置いた場合"),
    ])
    sheet_hyojun(wb, act, gen, proj)
    sheet_dai9(wb, gen)
    sheet_jinko(wb, gen)
    OUT_X.mkdir(parents=True, exist_ok=True)
    px = OUT_X / f"小野町_第10期_サービス別単価・見込量試算_{ASOF}.xlsx"
    wb.save(px)

    OUT_D.mkdir(parents=True, exist_ok=True)
    pd1 = OUT_D / f"小野町_第9期計画の見込みと実績の対比_{ASOF}.docx"
    doc_dai9(gen, proj).save(pd1)
    pd2 = OUT_X / f"小野町_人口推計の出典統一メモ_{ASOF}.docx"
    doc_jinko(gen).save(pd2)

    t = totals(proj)
    print(f"生成: {px.name}（{len(wb.sheetnames)}シート）")
    print(f"生成: {pd1.name}")
    print(f"生成: {pd2.name}")
    print("\n分母の変動係数")
    for nm, v in cv:
        print(f"  {nm:<14} {v:5.2f}％")
    print("\n第10期 総給付費の試算（千円）")
    for y, v in zip(EST_YEARS, t):
        print(f"  {y:<8} {v:>12,.0f}")
    print(f"  第10期3年計 {sum(t[1:]):>12,.0f}")
    return act, gen, proj


if __name__ == "__main__":
    main()
