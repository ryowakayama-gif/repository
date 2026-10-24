# -*- coding: utf-8 -*-
"""川崎町 第10期 計画見込量の独立算定（令和9〜11年度）

方法
  基準 ： 令和7年度（完結年度）の実績。介護保険事業状況報告（年報・令和7年度）
          の要介護度別の受給者数・利用回（日）数・給付費をそのまま用いる。
  伸び ： 要介護度別の認定者数の伸びのみ。
          伸び率(年,d) ＝ 認定者数(年,d) ÷ 認定者数(令和8年度,d)
          （令和8年度の実績見込み値は令和7年度実績に置き換え済みのため、
            令和7年度実績にこの比を乗じることで令和9〜11年度が得られる）
  据置 ： サービス利用率・1人1月あたり利用回（日）数・1人1月あたり給付費は
          令和7年度の水準で据え置く。

出力
  04_調査・入力・分析/R8.9.15計画見込量/ に txt、
  05_試算・管理シート/ に xlsx
"""
import os
import warnings
from decimal import Decimal, ROUND_HALF_UP

import openpyxl

warnings.filterwarnings("ignore")

BASE = "/home/user/repository/kawasaki_project"
NEN = f"{BASE}/09_元資料/R7実績データ/年報データ_2025_川崎町.xlsx"
SOK = (f"{BASE}/09_元資料/R8実績データ/R8.9.11受領版/"
       "【川崎町】第10期_将来推計総括表_R8.9.11出力_在宅回数入力後.xlsx")

D7 = ["要支援1", "要支援2", "要介護1", "要介護2", "要介護3", "要介護4", "要介護5"]
YEARS = ["令和9年度", "令和10年度", "令和11年度"]

wb = openpyxl.load_workbook(NEN, data_only=True)
sok = openpyxl.load_workbook(SOK, data_only=True)


# ══════════════════════════════════ 1. 認定者数と伸び率
def nintei(col):
    """総括表 1_推計値サマリ の要介護度別認定者数（総数）"""
    ws = sok["1_推計値サマリ"]
    return [float(ws.cell(r, col).value) for r in range(20, 27)]


NIN = {
    "令和6年度": nintei(4), "令和7年度": nintei(5), "令和8年度": nintei(6),
    "令和9年度": nintei(7), "令和10年度": nintei(8), "令和11年度": nintei(9),
    "令和12年度": nintei(11), "令和17年度": nintei(13),
    "令和22年度": nintei(15),
}
RATIO = {y: [NIN[y][i] / NIN["令和8年度"][i] if NIN["令和8年度"][i] else 0.0
             for i in range(7)]
         for y in ["令和9年度", "令和10年度", "令和11年度", "令和12年度",
                   "令和17年度", "令和22年度"]}


# ══════════════════════════════════ 2. 年報からの実績抽出
def row7(ws, r, c_s1, c_s2, c_k1):
    out = [ws.cell(r, c_s1).value or 0, ws.cell(r, c_s2).value or 0]
    out += [ws.cell(r, c_k1 + i).value or 0 for i in range(5)]
    return [float(v) for v in out]


w16 = wb["様式１の７（１６）居宅介護"]
w17 = wb["様式１の７（１７）居宅介護"]
w18 = wb["様式１の７（１８）地域密着型"]
w19 = wb["様式１の７（１９）地域密着型（２０）施設介護"]
w16f = wb["様式１の６"]
wkyu = wb["様式２（給付費）"]
wken = wb["様式２（件数）"]

w16r = {str(w16.cell(r, 3).value or "").strip(): r for r in range(12, 27)}
w17r = {str(w17.cell(r, 3).value or "").strip(): r for r in range(12, 23)}
w18r = {str(w18.cell(r, 3).value or "").strip().replace("\n", "").replace("　", ""): r
        for r in range(12, 25)}
w19r = {str(w19.cell(r, 3).value or "").strip(): r for r in range(12, 18)}


def kyu(rows):
    """様式２（給付費）の行番号（複数可）を合算。円／年"""
    if isinstance(rows, int):
        rows = [rows]
    out = [0.0] * 7
    for r in rows:
        v = row7(wkyu, r, 6, 7, 10)
        out = [a + b for a, b in zip(out, v)]
    return out


def n16(name):
    return row7(w16, w16r[name], 4, 5, 8)


def n18(name):
    return row7(w18, w18r[name], 4, 5, 8)


def nfac(r):
    return [0.0, 0.0] + [float(w16f.cell(r, 7 + i).value or 0) for i in range(5)]


def nken(name):
    for r in range(11, 53):
        nm = None
        for c in (3, 4, 5):
            v = wken.cell(r, c).value
            if v and str(v).strip():
                nm = str(v).strip()
        if nm == name:
            return row7(wken, r, 6, 7, 10)
    raise KeyError(name)


def q17(name):
    return row7(w17, w17r[name], 4, 5, 8)


def q19(name):
    return row7(w19, w19r[name], 4, 5, 8)


def add(*a):
    return [sum(x) for x in zip(*a)]


Z = [0.0] * 7

# (計画上の名称, 給付費行, 年間受給者数, 年間回（日）数 or None, 単位, 区分)
SV = [
    ("訪問介護", 13, n16("訪問介護"), q17("訪問介護（回）"), "回", "在宅"),
    ("訪問入浴介護", 14, n16("訪問入浴介護"), q17("訪問入浴介護（回）"), "回", "在宅"),
    ("訪問看護", 15, n16("訪問看護"), q17("訪問看護（回）"), "回", "在宅"),
    ("訪問リハビリテーション", 16, n16("訪問リハビリテーション"),
     q17("訪問リハビリテーション（回）"), "回", "在宅"),
    ("居宅療養管理指導", 17, n16("居宅療養管理指導"), None, "", "在宅"),
    ("通所介護", 19, n16("通所介護"), q17("通所介護（回）"), "回", "在宅"),
    ("通所リハビリテーション", 20, n16("通所リハビリテーション"),
     q17("通所リハビリテーション（回）"), "回", "在宅"),
    ("短期入所生活介護", 22, n16("短期入所生活介護"), q17("短期入所生活介護（日）"), "日", "在宅"),
    ("短期入所療養介護（老健）", 23, n16("短期入所療養介護（介護老人保健施設）"),
     q17("短期入所療養介護（介護老人保健施設）（日）"), "日", "在宅"),
    ("短期入所療養介護（病院等）", 24, n16("短期入所療養介護（病院等）"),
     q17("短期入所療養介護（病院等）（日）"), "日", "在宅"),
    ("短期入所療養介護（介護医療院）", 25, n16("短期入所療養介護（介護医療院）"),
     q17("短期入所療養介護（介護医療院）（日）"), "日", "在宅"),
    ("福祉用具貸与", 27, n16("福祉用具貸与"), None, "", "在宅"),
    ("特定福祉用具購入費", 28, nken("特定福祉用具販売"), None, "", "在宅"),
    ("住宅改修費", 29, nken("住宅改修"), None, "", "在宅"),
    ("特定施設入居者生活介護", [30, 31],
     add(n16("特定施設入居者生活介護（短期利用以外）"),
         n16("特定施設入居者生活介護（短期利用）")), None, "", "居住系"),
    ("定期巡回・随時対応型訪問介護看護", 34, n18("定期巡回・随時対応型訪問介護看護"),
     None, "", "在宅"),
    ("夜間対応型訪問介護", 35, n18("夜間対応型訪問介護"), None, "", "在宅"),
    ("地域密着型通所介護", 36, n18("地域密着型通所介護"), q19("地域密着型通所介護（回）"),
     "回", "在宅"),
    ("認知症対応型通所介護", 37, n18("認知症対応型通所介護"),
     q19("認知症対応型通所介護（回）"), "回", "在宅"),
    ("小規模多機能型居宅介護", [38, 39],
     add(n18("小規模多機能型居宅介護（短期利用以外）"),
         n18("小規模多機能型居宅介護（短期利用）")), None, "", "在宅"),
    ("認知症対応型共同生活介護", [40, 41],
     add(n18("認知症対応型共同生活介護（短期利用以外）"),
         n18("認知症対応型共同生活介護（短期利用）")), None, "", "居住系"),
    ("地域密着型特定施設入居者生活介護", [42, 43],
     add(n18("地域密着型特定施設入居者生活介護（短期利用以外）"),
         n18("地域密着型特定施設入居者生活介護（短期利用）")), None, "", "居住系"),
    ("地域密着型介護老人福祉施設入所者生活介護", 44,
     n18("地域密着型介護老人福祉施設入所者生活介護"), None, "", "施設"),
    ("看護小規模多機能型居宅介護", [45, 46],
     add(n18("複合型サービス(看護小規模多機能型居宅介護)（短期利用以外）"),
         n18("複合型サービス(看護小規模多機能型居宅介護)（短期利用）")), None, "", "在宅"),
    ("介護老人福祉施設", 48, nfac(25), None, "", "施設"),
    ("介護老人保健施設", 49, nfac(28), None, "", "施設"),
    ("介護医療院", 51, nfac(34), None, "", "施設"),
    ("介護予防支援・居宅介護支援", 32, n16("介護予防支援・居宅介護支援"), None, "", "在宅"),
]

PRE = {  # 介護予防（要支援1・2）としての表示名
    "訪問入浴介護": "介護予防訪問入浴介護",
    "訪問看護": "介護予防訪問看護",
    "訪問リハビリテーション": "介護予防訪問リハビリテーション",
    "居宅療養管理指導": "介護予防居宅療養管理指導",
    "通所リハビリテーション": "介護予防通所リハビリテーション",
    "短期入所生活介護": "介護予防短期入所生活介護",
    "短期入所療養介護（老健）": "介護予防短期入所療養介護（老健）",
    "短期入所療養介護（病院等）": "介護予防短期入所療養介護（病院等）",
    "短期入所療養介護（介護医療院）": "介護予防短期入所療養介護（介護医療院）",
    "福祉用具貸与": "介護予防福祉用具貸与",
    "特定福祉用具購入費": "特定介護予防福祉用具購入費",
    "住宅改修費": "介護予防住宅改修",
    "特定施設入居者生活介護": "介護予防特定施設入居者生活介護",
    "認知症対応型通所介護": "介護予防認知症対応型通所介護",
    "小規模多機能型居宅介護": "介護予防小規模多機能型居宅介護",
    "認知症対応型共同生活介護": "介護予防認知症対応型共同生活介護",
    "介護予防支援・居宅介護支援": "介護予防支援",
}
S_IDX = (0, 1)          # 要支援1・2
K_IDX = (2, 3, 4, 5, 6)  # 要介護1〜5


def proj(vals7, ratio, idx):
    """令和7年度の要介護度別実績に伸び率を乗じ、対象区分だけ合計する"""
    return sum(vals7[i] * ratio[i] for i in idx)


def build():
    """サービス×区分×年度の見込量をつくる"""
    rows = []
    for nm, krow, nyear, qyear, unit, kubun in SV:
        g7 = kyu(krow)                       # 給付費（円／年）
        for idx, label in ((K_IDX, nm), (S_IDX, PRE.get(nm))):
            if label is None:
                continue
            if sum(g7[i] for i in idx) == 0 and sum(nyear[i] for i in idx) == 0:
                continue
            rec = {"name": label, "unit": unit, "kubun": kubun,
                   "予防" if idx == S_IDX else "介護": True}
            rec["区分"] = "予防" if idx == S_IDX else "介護"
            rec["R7給付費"] = sum(g7[i] for i in idx) / 1000.0
            rec["R7人数"] = sum(nyear[i] for i in idx) / 12.0
            rec["R7量"] = (sum(qyear[i] for i in idx) / 12.0) if qyear else None
            for y in ["令和9年度", "令和10年度", "令和11年度", "令和12年度",
                      "令和17年度", "令和22年度"]:
                rt = RATIO[y]
                rec[y + "給付費"] = proj(g7, rt, idx) / 1000.0
                rec[y + "人数"] = proj(nyear, rt, idx) / 12.0
                rec[y + "量"] = (proj(qyear, rt, idx) / 12.0) if qyear else None
            rows.append(rec)
    return rows


ROWS = build()

# ══════════════════════════════════ 3. その他給付費
TOKUTEI7 = row7(wb["様式２の５"], 52, 6, 7, 10)           # 特定入所者（要介護度別）
KOGAKU7 = 30468013.0                                     # 高額介護サービス費
GASSAN7 = 2512792.0                                      # 高額医療合算
KENSU7 = 10929.0                                         # 給付費決定件数
TESURYO_TANKA = 60.0

SOK5 = sok["5_保険料推計"]
CHIIKI = {  # 地域支援事業費（現行値。令和11年度は入力漏れ3項目を補完）
    "令和9年度": {"総合": 6786184, "包括任意": 38647166, "充実": 599744},
    "令和10年度": {"総合": 6505887, "包括任意": 42317402, "充実": 772924},
    "令和11年度": {"総合": 2227310 + 780000 + 45144 + 3173135,
                   "包括任意": 45987639, "充実": 946104},
}
KOFU = {"令和9年度": 0.0614, "令和10年度": 0.0593, "令和11年度": 0.0569}
HOKENSHA = {"令和9年度": 3231, "令和10年度": 3203, "令和11年度": 3177}
SHUNORITSU = 0.96

# 所得段階別第1号被保険者数（13段階）
# 出所：介護保険事業状況報告（年報）様式1「所得段階別」列20「年度末現在被保険者数」
RYORITSU = [0.455, 0.685, 0.690, 0.90, 1.00, 1.20, 1.30, 1.50, 1.70,
            1.90, 2.10, 2.30, 2.40]
KEIGEN = {0: 0.285, 1: 0.485, 2: 0.685}          # 公費による軽減後の乗率
NINZU_R7 = [383, 298, 305, 343, 679, 410, 427, 232, 85, 26, 17, 7, 28]
NINZU_R6 = [412, 308, 304, 383, 699, 375, 419, 218, 67, 25, 13, 7, 31]


def hosei_n(ninzu):
    """所得段階別加入割合による補正係数（＝加重計÷被保険者数計）。"""
    return sum(n * r for n, r in zip(ninzu, RYORITSU)) / sum(ninzu)


HOSEI_SYS = 0.964047      # 見える化システムに登録されている値（第10〜13段階が0人）
HOSEI_R6 = hosei_n(NINZU_R6)
HOSEI = hosei_n(NINZU_R7)  # 採用値（令和7年度末の実績分布）

G7_TOTAL = sum(sum(kyu(s[1])) for s in SV)


def zaisei(y):
    sogo = sum(r[y + "給付費"] for r in ROWS) * 1000.0      # 円
    scale = sogo / G7_TOTAL
    tok = proj(TOKUTEI7, RATIO[y], (0, 1, 2, 3, 4, 5, 6))
    kog = KOGAKU7 * scale
    gas = GASSAN7 * scale
    ken = round(KENSU7 * scale)
    tes = ken * TESURYO_TANKA
    hyojun = sogo + tok + kog + gas + tes
    c = CHIIKI[y]
    chiiki = c["総合"] + c["包括任意"] + c["充実"]
    dai1 = (hyojun + chiiki) * 0.23
    chosei_so = hyojun * 0.05 + c["総合"] * 0.05
    chosei_mi = (hyojun + c["総合"]) * KOFU[y]
    shuno = dai1 + chosei_so - chosei_mi
    return dict(総給付費=sogo, 特定入所者=tok, 高額介護=kog, 高額医療合算=gas,
                審査件数=ken, 審査手数料=tes, 標準給付費=hyojun,
                地域支援事業費=chiiki, 総合事業費=c["総合"],
                第1号負担分=dai1, 調整交付金相当額=chosei_so,
                調整交付金見込額=chosei_mi, 保険料収納必要額=shuno)


FIN = {y: zaisei(y) for y in YEARS}
shuno3 = sum(FIN[y]["保険料収納必要額"] for y in YEARS)
hihoken3 = sum(HOKENSHA[y] for y in YEARS)
hokenryo = shuno3 / SHUNORITSU / 12.0 / (hihoken3 * HOSEI)

# ══════════════════════════════════ 4. 出力
OUTDIR = f"{BASE}/04_調査・入力・分析/R8.9.15計画見込量"
os.makedirs(OUTDIR, exist_ok=True)
L = []
P = L.append
P("川崎町 第10期 計画見込量の独立算定（令和9〜11年度）")
P("作成：ビズアップ公共コンサルティング 札幌事業所／令和8年9月15日")
P("")
P("■ 基礎となる認定者数（総括表・見える化システムの推計）")
P(f"{'年度':<10}" + "".join(f"{d:>9}" for d in D7) + f"{'計':>9}")
for y in ["令和7年度", "令和8年度"] + YEARS + ["令和12年度", "令和17年度",
                                              "令和22年度"]:
    P(f"{y:<10}" + "".join(f"{v:>9,.0f}" for v in NIN[y]) + f"{sum(NIN[y]):>9,.0f}")
P("")
P("■ 令和8年度を1.000とした要介護度別の伸び率")
for y in YEARS:
    P(f"{y:<10}" + "".join(f"{v:>9.3f}" for v in RATIO[y]))
P("")
P("■ サービス別 見込量（給付費：千円／年、人数・量：1月当たり）")
P(f"{'サービス':<32}{'区分':<5}{'項目':<8}"
  + "".join(f"{y:>12}" for y in ["R7実績"] + YEARS))
for r in ROWS:
    for key, lab in (("給付費", "給付費"), ("人数", "人数"), ("量", r["unit"] + "数")):
        if key == "量" and r["R7量"] is None:
            continue
        v7 = r["R7" + key]
        vv = [r[y + key] for y in YEARS]
        fmt = "{:>12,.1f}" if key != "給付費" else "{:>12,.0f}"
        P(f"{r['name']:<32}{r['区分']:<5}{lab:<8}"
          + fmt.format(v7) + "".join(fmt.format(x) for x in vv))
P("")
P("■ 区分別の給付費（千円）")
P(f"{'区分':<14}{'R7実績':>14}" + "".join(f"{y:>14}" for y in YEARS))
for kb in ["在宅", "居住系", "施設"]:
    for ku in ["予防", "介護"]:
        sel = [r for r in ROWS if r["kubun"] == kb and r["区分"] == ku]
        if not sel:
            continue
        P(f"{ku + ' ' + kb:<14}{sum(r['R7給付費'] for r in sel):>14,.0f}"
          + "".join(f"{sum(r[y + '給付費'] for r in sel):>14,.0f}" for y in YEARS))
P(f"{'合計':<14}{sum(r['R7給付費'] for r in ROWS):>14,.0f}"
  + "".join(f"{sum(r[y + '給付費'] for r in ROWS):>14,.0f}" for y in YEARS))
P("")
P("■ 保険料算定")
KEYS = ["総給付費", "特定入所者", "高額介護", "高額医療合算", "審査手数料",
        "標準給付費", "地域支援事業費", "第1号負担分", "調整交付金相当額",
        "調整交付金見込額", "保険料収納必要額"]
P(f"{'項目':<20}" + "".join(f"{y:>18}" for y in YEARS) + f"{'3か年計':>18}")
for k in KEYS:
    vals = [FIN[y][k] for y in YEARS]
    P(f"{k:<20}" + "".join(f"{v:>18,.0f}" for v in vals) + f"{sum(vals):>18,.0f}")
P("")
P(f"第1号被保険者数（3か年計）　{hihoken3:,.0f} 人")
P(f"所得段階別加入割合による補正後　{hihoken3 * HOSEI:,.1f} 人")
P(f"予定保険料収納率　{SHUNORITSU:.2%}")
P(f"★ 保険料基準額（月額）　{hokenryo:,.2f} 円")

open(f"{OUTDIR}/計画見込量_独立算定_R8.9.15.txt", "w").write("\n".join(L))
print("\n".join(L))


# ══════════════════════════════════ 5. 感度分析（所得段階・準備基金）
def premium(h, torikuzushi=0.0):
    return (shuno3 / SHUNORITSU / 12.0 / (hihoken3 * h)
            - torikuzushi / (12.0 * hihoken3 * h))


CASES = [("見える化システムの現在の登録値（第10〜13段階が0人）", HOSEI_SYS),
         ("令和6年度末の実績分布（13段階）", HOSEI_R6),
         ("令和7年度末の実績分布（13段階）★採用", HOSEI)]

L2 = []
Q = L2.append
Q("")
Q("■ 所得段階別第1号被保険者数（令和7年度末・年報 様式1 所得段階別）")
Q(f"{'段階':<10}{'乗率':>8}{'人数':>8}{'構成比':>9}{'加重':>10}")
_tot = sum(NINZU_R7)
for i, (n, rt) in enumerate(zip(NINZU_R7, RYORITSU), 1):
    Q(f"{'第' + str(i) + '段階':<10}{rt:>8.3f}{n:>8,}"
      f"{n / _tot * 100:>8.1f}%{n * rt:>10.1f}")
Q(f"{'計':<10}{'':>8}{_tot:>8,}{100.0:>8.1f}%"
  f"{sum(n * r for n, r in zip(NINZU_R7, RYORITSU)):>10.1f}")
Q("")
Q("■ 所得段階別の分布のとり方による保険料基準額")
Q(f"{'分布':<44}{'補正係数':>10}{'補正後被保険者数':>18}{'保険料基準額':>14}")
base = premium(HOSEI)
for nm, h in CASES:
    Q(f"{nm:<44}{h:>10.6f}{hihoken3 * h:>18,.1f}{premium(h):>14,.2f}")
Q(f"　登録値→令和7年度末の実績に改めた場合の差　{base - premium(HOSEI_SYS):+,.1f} 円／月")
Q("")
Q("■ 準備基金を取り崩す場合の換算")
Q(f"　取崩額1,000万円あたり　月額 ▲{10_000_000 / (12 * hihoken3 * HOSEI):,.1f} 円")
for t, nm in ((76_250_000, "50％取崩し"), (152_500_000, "全額取崩し"),
              (78_000_000, "第9期と同額")):
    Q(f"　{nm}（{t // 1000:,}千円）　月額 ▲"
      f"{t / (12 * hihoken3 * HOSEI):,.1f} 円 → {premium(HOSEI, t):,.0f} 円")
Q("")
Q("■ 見える化システムの現行推計（令和8年9月11日出力）との比較")
Q(f"{'項目':<24}{'当社の独立算定':>18}{'見える化システム':>18}{'差':>14}")
for lab, a, b in [
    ("総給付費（3か年・千円）", sum(FIN[y]["総給付費"] for y in YEARS) / 1000, 3_075_699),
    ("標準給付費見込額（千円）", sum(FIN[y]["標準給付費"] for y in YEARS) / 1000, 3_333_015),
    ("地域支援事業費（千円）", sum(FIN[y]["地域支援事業費"] for y in YEARS) / 1000, 144_790),
    ("保険料基準額（月額・円）", base, 7_200.16),
]:
    Q(f"{lab:<24}{a:>18,.2f}{b:>18,.2f}{a - b:>14,.2f}")

open(f"{OUTDIR}/計画見込量_独立算定_R8.9.15.txt", "a").write("\n".join(L2))
print("\n".join(L2))


# ══════════════════════════════════ 6. Excel 成果品
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

XL = (f"{BASE}/05_試算・管理シート/川崎町_第10期_計画見込量_R8.9.15.xlsx")
TT = Font(bold=True, size=12)
HD = Font(bold=True, color="FFFFFF")
HF = PatternFill("solid", fgColor="365F91")
SF = PatternFill("solid", fgColor="DCE6F1")
YF = PatternFill("solid", fgColor="FFF2CC")
GF = PatternFill("solid", fgColor="D9D9D9")
TH = Side(style="thin", color="A6A6A6")
BX = Border(left=TH, right=TH, top=TH, bottom=TH)
YS = ["令和9年度", "令和10年度", "令和11年度"]
YL = YS + ["令和12年度", "令和17年度", "令和22年度"]

bk = openpyxl.Workbook()
bk.remove(bk.active)


def hrow(ws, r, labels, w=None):
    for i, v in enumerate(labels, 1):
        c = ws.cell(r, i, v)
        c.font, c.fill, c.border = HD, HF, BX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if w:
        for i, x in enumerate(w, 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = x


def brow(ws, r, vals, fmt=None, fill=None, bold=False):
    for i, v in enumerate(vals, 1):
        c = ws.cell(r, i, v)
        c.border = BX
        if bold:
            c.font = Font(bold=True)
        if fill:
            c.fill = fill
        if fmt and i >= 3 and isinstance(v, (int, float)):
            c.number_format = fmt


# ── ① 算定の考え方
ws = bk.create_sheet("①算定の考え方")
ws.column_dimensions["A"].width = 100
ws["A1"] = "川崎町 第10期介護保険事業計画　サービス見込量の算定方法"
ws["A1"].font = TT
txt = [
    "",
    "作成：ビズアップ公共コンサルティング 札幌事業所／令和8年9月15日　保険者番号 04324",
    "",
    "【1】基準年度",
    "　令和7年度（完結年度）の実績を基準とする。出所は介護保険事業状況報告（年報・令和7年度）。",
    "　・受給者数　　　　様式1の6(13)(14)(15)、様式1の7(16)(18)、様式2（件数）",
    "　・利用回（日）数　様式1の7(17)(19)",
    "　・給付費　　　　　様式2（給付費）",
    "　いずれも要介護度別に取得している。",
    "",
    "【2】伸びの取り方",
    "　要介護度別の認定者数の伸びのみを乗じる。",
    "　　見込量(年,要介護度) ＝ 令和7年度実績(要介護度) × 認定者数(年,要介護度) ÷ 認定者数(令和8年度,要介護度)",
    "　認定者数は地域包括ケア「見える化」システムの推計値（性別・年齢5歳階級別・要介護度別）を用いる。",
    "　令和8年度の実績見込み値は令和7年度実績に置き換え済みであるため、この比を令和7年度実績に乗じることで",
    "　令和9〜11年度の見込量が得られる。",
    "",
    "【3】据え置くもの（過年度数値据え置き）",
    "　・サービス利用率（認定者数に対する利用者数の割合）",
    "　・1人1月あたりの利用回（日）数",
    "　・1人1月あたりの給付費（単価）",
    "　いずれも令和7年度の水準で据え置く。",
    "",
    "【4】据え置きとした理由",
    "　(1) 令和6年度→令和7年度の伸びは、利用者数の少ないサービスで振れが大きい。",
    "　　　例：訪問看護の1人1月あたり回数 令和6年度5.40回→令和7年度8.17回（年＋51％）",
    "　　　　　訪問介護 要介護4 令和6年度42.1回→令和7年度17.7回（年▲58％）",
    "　　　これを3年間複利で効かせると、計画値として説明できない水準になる。",
    "　(2) 第9期計画は実績を5〜6％上回って推移しており、伸びを重ねることは適当でない。",
    "　(3) 見える化システムの令和9年度以降の在宅サービス利用回（日）数は、令和8年5月ベースの値が",
    "　　　固定されたまま解除できない状態にある（別紙のとおり）。当社の独立算定はこれを回避している。",
    "",
    "【5】その他給付費",
    "　・特定入所者介護サービス費等　令和7年度実績52,967,156円を要介護度別に伸ばす",
    "　・高額介護サービス費等　　　　令和7年度実績30,468,013円を総給付費の伸びで按分",
    "　・高額医療合算介護サービス費　令和7年度実績2,512,792円を同上",
    "　・審査支払手数料　　　　　　　令和7年度の給付費決定件数10,929件を同上、単価60円",
    "",
    "【6】地域支援事業費",
    "　見える化システムに現在入力されている値を用いる（町のデータ未受領のため）。",
    "　ただし令和11年度に入力漏れの3項目（介護予防ケアマネジメント・審査支払手数料・",
    "　一般介護予防事業評価事業、計約4,247千円）があるため、令和9→10年度の傾向を延長して補完した。",
    "",
    "【7】保険料の算定",
    "　保険料収納必要額 ＝ 第1号被保険者負担分相当額 ＋ 調整交付金相当額 － 調整交付金見込額",
    "　　第1号被保険者負担分相当額 ＝（標準給付費見込額＋地域支援事業費）×23％",
    "　　調整交付金相当額 ＝ 標準給付費見込額×5％ ＋ 総合事業費×5％",
    "　　調整交付金見込額 ＝（標準給付費見込額＋総合事業費）× 見込交付割合",
    "　　　　　　　　　　　　（令和9年度6.14％・令和10年度5.93％・令和11年度5.69％）",
    "　保険料基準額（月額）＝ 保険料収納必要額 ÷ 予定収納率96％ ÷ 36月 ÷ 所得段階補正後被保険者数",
    "",
    "【8】留意事項",
    "　・年度間に凹凸があるのは、認定者数の推計が性別・年齢5歳階級別のグリッドで行われ、",
    "　　各セルが整数に丸められるためである。3か年の合計には影響しない。",
    "　・所得段階別第1号被保険者数の第10〜13段階が0人のままであり、保険料は過大に出ている。",
    "　　実際の13段階分布を入手しだい、再算定が必要（⑦シートの感度分析を参照）。",
]
for i, t in enumerate(txt, 2):
    ws.cell(i + 1, 1, t)

# ── ② 認定者数
ws = bk.create_sheet("②認定者数")
hrow(ws, 1, ["年度", "区分"] + D7 + ["計"], [14, 10] + [10] * 8)
r = 2
for y in ["令和6年度", "令和7年度", "令和8年度"] + YL:
    kb = "実績" if y in ("令和6年度", "令和7年度") else ("実績見込み" if y == "令和8年度" else "推計")
    brow(ws, r, [y, kb] + [round(v) for v in NIN[y]] + [round(sum(NIN[y]))], "#,##0",
         SF if y in ("令和6年度", "令和7年度", "令和8年度") else None)
    r += 1
r += 1
ws.cell(r, 1, "令和8年度を1.000とした伸び率").font = Font(bold=True)
r += 1
hrow(ws, r, ["年度", ""] + D7 + [""])
r += 1
for y in YL:
    brow(ws, r, [y, ""] + [round(v, 3) for v in RATIO[y]] + [""], "0.000")
    r += 1

# ── ③④ サービス見込量
for sheet, ku in (("③介護予防サービス", "予防"), ("④介護サービス", "介護")):
    ws = bk.create_sheet(sheet)
    hrow(ws, 1, ["サービス", "項目", "令和7年度実績"] + YS + ["第10期計"],
         [38, 14, 14, 14, 14, 14, 14])
    r = 2
    sel = [x for x in ROWS if x["区分"] == ku]
    for rec in sel:
        items = [("給付費", "給付費（千円）", "#,##0")]
        if rec["R7量"] is not None:
            items.append(("量", f"回（日）数（{rec['unit']}／月）", "#,##0.0"))
        items.append(("人数", "利用者数（人／月）", "#,##0.0"))
        for k, lab, fm in items:
            vv = [rec[y + k] for y in YS]
            tot = sum(vv) if k == "給付費" else None
            brow(ws, r, [rec["name"], lab, rec["R7" + k]] + vv + [tot], fm,
                 SF if k == "給付費" else None, bold=(k == "給付費"))
            r += 1
    r += 1
    brow(ws, r, [f"{ku} 給付費 合計", "給付費（千円）",
                 sum(x["R7給付費"] for x in sel)]
         + [sum(x[y + "給付費"] for x in sel) for y in YS]
         + [sum(sum(x[y + "給付費"] for x in sel) for y in YS)], "#,##0", YF, True)

# ── ⑤ 総給付費
ws = bk.create_sheet("⑤総給付費")
hrow(ws, 1, ["区分", "令和7年度実績"] + YS + ["第10期計"], [24, 16, 16, 16, 16, 16])
r = 2
for kb in ["在宅", "居住系", "施設"]:
    for ku in ["予防", "介護"]:
        sel = [x for x in ROWS if x["kubun"] == kb and x["区分"] == ku]
        if not sel:
            continue
        vv = [sum(x[y + "給付費"] for x in sel) for y in YS]
        brow(ws, r, [f"{ku} {kb}サービス", sum(x["R7給付費"] for x in sel)] + vv + [sum(vv)],
             "#,##0")
        r += 1
vv = [sum(x[y + "給付費"] for x in ROWS) for y in YS]
brow(ws, r, ["総給付費", sum(x["R7給付費"] for x in ROWS)] + vv + [sum(vv)], "#,##0", YF, True)

# ── ⑥ 地域支援事業費
ws = bk.create_sheet("⑥地域支援事業費")
hrow(ws, 1, ["区分"] + YS + ["第10期計", "備考"], [40, 16, 16, 16, 16, 40])
r = 2
for k, lab, note in [("総合", "介護予防・日常生活支援総合事業費",
                      "令和11年度は入力漏れ3項目（約4,247千円）を補完"),
                     ("包括任意", "包括的支援事業（包括センター運営）及び任意事業費", ""),
                     ("充実", "包括的支援事業（社会保障充実分）", "")]:
    vv = [CHIIKI[y][k] for y in YS]
    brow(ws, r, [lab] + vv + [sum(vv), note], "#,##0")
    r += 1
vv = [FIN[y]["地域支援事業費"] for y in YS]
brow(ws, r, ["地域支援事業費 計"] + vv + [sum(vv), ""], "#,##0", YF, True)

# ── ⑦ 保険料算定
ws = bk.create_sheet("⑦保険料算定")
hrow(ws, 1, ["項目"] + YS + ["第10期計"], [34, 20, 20, 20, 20])
r = 2
for k in KEYS:
    vv = [FIN[y][k] for y in YEARS]
    brow(ws, r, [k] + vv + [sum(vv)], "#,##0",
         YF if k in ("標準給付費", "保険料収納必要額") else None,
         bold=k in ("標準給付費", "保険料収納必要額"))
    r += 1
r += 1
for lab, v, fm in [("第1号被保険者数（3か年計）", hihoken3, "#,##0"),
                   ("所得段階別加入割合による補正係数（令和7年度末実績・13段階）",
                    HOSEI, "0.000000"),
                   ("補正後被保険者数（3か年計）", hihoken3 * HOSEI, "#,##0.0"),
                   ("予定保険料収納率", SHUNORITSU, "0.0%"),
                   ("保険料基準額（月額）", premium(HOSEI), "#,##0.00")]:
    brow(ws, r, [lab, v], fm, YF if "基準額" in lab else None, bold="基準額" in lab)
    ws.cell(r, 2).number_format = fm
    r += 1
r += 2
ws.cell(r, 1, "【内訳】所得段階別第1号被保険者数"
              "（令和7年度末・年報 様式1 所得段階別 列20）").font = Font(bold=True)
r += 1
hrow(ws, r, ["段階", "基準額に対する乗率", "被保険者数（人）", "構成比",
             "加重（人×乗率）"])
r += 1
_tot = sum(NINZU_R7)
for i, (n, rt) in enumerate(zip(NINZU_R7, RYORITSU), 1):
    lab = f"第{i}段階"
    if i - 1 in KEIGEN:
        lab += f"（軽減後 {KEIGEN[i - 1]:.3f}）"
    brow(ws, r, [lab, rt, n, n / _tot, n * rt], "#,##0.0")
    ws.cell(r, 2).number_format = "0.000"
    ws.cell(r, 3).number_format = "#,##0"
    ws.cell(r, 4).number_format = "0.0%"
    r += 1
brow(ws, r, ["計", "", _tot, 1.0,
             sum(n * rt for n, rt in zip(NINZU_R7, RYORITSU))], "#,##0.0",
     bold=True)
ws.cell(r, 3).number_format = "#,##0"
ws.cell(r, 4).number_format = "0.0%"
r += 2

ws.cell(r, 1, "【比較】所得段階別の分布のとり方による保険料基準額").font = Font(bold=True)
r += 1
hrow(ws, r, ["分布", "補正係数", "補正後被保険者数", "保険料基準額（月額）",
             "採用との差"])
r += 1
for nm, h in CASES:
    brow(ws, r, [nm, h, hihoken3 * h, premium(h), premium(h) - premium(HOSEI)],
         "#,##0.00")
    ws.cell(r, 2).number_format = "0.000000"
    ws.cell(r, 3).number_format = "#,##0.0"
    r += 1
r += 1
ws.cell(r, 1, "※ 見える化システムには第10〜13段階が0人として登録されているため、"
              "補正係数が実績より 0.0497 低く出て、保険料が 352円／月 高く算定される。")
r += 2

ws.cell(r, 1, "【感度分析】介護給付費準備基金の取崩し").font = Font(bold=True)
r += 1
hrow(ws, r, ["取崩額（千円）", "月額換算（円）", "保険料基準額（月額）",
             "第9期6,500円との差"])
r += 1
for t in [0, 30000, 50000, 76250, 78000, 100000, 152500]:
    v = premium(HOSEI, t * 1000)
    brow(ws, r, [t, v - premium(HOSEI), v, v - 6500], "#,##0.00")
    ws.cell(r, 1).number_format = "#,##0"
    r += 1
r += 2
ws.cell(r, 1, "※ 基金残高は令和8年度末見込み152,500千円（町財政担当に確定値を照会中）。"
              "76,250千円は50％取崩し、152,500千円は全額取崩し。")
r += 1
ws.cell(r, 1, "※ 第9期は78,000千円（月額698.59円相当）を取り崩し、"
              "条例上の基準額を6,508.33円としている。")

bk.save(XL)
print("\n保存しました：", XL)
print("シート：", bk.sheetnames)
