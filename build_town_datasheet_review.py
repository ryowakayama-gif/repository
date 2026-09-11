# -*- coding: utf-8 -*-
"""町別データシートのレッドチームレビュー（機械点検）.

町別データシート（第10期計画_町別データシート.xlsx）の数値と記述を、
出典から独立に再計算して突き合わせる。

「作った側の言い分」を採らず、成果品の実物だけを読んで、
出典のデータモジュール及び他の成果品と照合する。
指摘が1件でも残る場合は終了コード1で終わる。

点検の観点
  A 合計の一致　　町別の合計が広域連合・保険者の値に一致するか
  B 転記の一致　　他の成果品からの転記が原本と一致するか
  C 算定方法　　　計画素案・国の様式・見える化との相違がないか
  D 越権　　　　　保険者単位でしか算定できないものを町別に掲げていないか
  E 体裁　　　　　出典と時点の明記、禁止表現、個人情報の非収録

使い方
  python3 build_town_datasheet_review.py
"""

import os
import repo_paths as RP
import re
import sys

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import data_juki as DJ                       # noqa: E402
import data_shien_tool as T                  # noqa: E402
import data_kyufu_jisseki as KJ              # noqa: E402
import data_kofukin_zenkoku as KZ            # noqa: E402

ODIR = RP.ROOT + "/output"
SHEET = os.path.join(ODIR, "第10期計画_町別データシート.xlsx")
CROSS = os.path.join(ODIR, "第10期計画_調査クロス集計・分析.xlsx")
RONTEN = os.path.join(ODIR, "第10期計画_3町別の論点整理.xlsx")

TOWNS = ["東川町", "美瑛町", "東神楽町"]
TSHEET = {"東川町": "01_東川町", "美瑛町": "02_美瑛町",
          "東神楽町": "03_東神楽町"}

FINDINGS = []


def ng(kubun, item, expected, actual, why):
    FINDINGS.append((kubun, item, str(expected), str(actual), why))


def cells(ws):
    """シートの全セルを (行番号, 値の並び) で返す。"""
    return list(ws.iter_rows(values_only=True))


def flat(ws):
    """シートの全文字列を1本につなげる。"""
    out = []
    for row in ws.iter_rows(values_only=True):
        for v in row:
            if v is not None:
                out.append(str(v))
    return "\n".join(out)


def block(rows, lead_text):
    """【…】見出しの次の行から、次の見出し又は「資料：」までを返す。"""
    got, on = [], False
    for r in rows:
        head = "" if r[0] is None else str(r[0])
        if on:
            if head.startswith("【") or head.startswith("資料："):
                break
            got.append(r)
            continue
        if head.startswith(lead_text):
            on = True
    return got


def nums(rows, label, col_from=1, col_to=None):
    for r in rows:
        if r[0] is not None and str(r[0]).strip() == label:
            v = list(r[col_from:col_to])
            return [x for x in v if x is not None and x != ""]
    return None


wb = openpyxl.load_workbook(SHEET, data_only=True)
ALL = {s: cells(wb[s]) for s in wb.sheetnames}
TEXT = {s: flat(wb[s]) for s in wb.sheetnames}
WHOLE = "\n".join(TEXT.values())
TOWN_TEXT = "\n".join(TEXT[TSHEET[t]] for t in TOWNS)


# ============================================================ A 合計の一致
def juki(town, year, band):
    return sum(DJ.JUKI[town][year][DJ.BAND[band]])


# A1 要介護認定者数
for y, yn in [("R6", "令和6年度"), ("R7", "令和7年度"), ("R8", "令和8年度")]:
    got = []
    for t in TOWNS:
        rows = block(ALL[TSHEET[t]], "【3　要介護認定者数】")
        v = nums(rows, yn)
        if v is None:
            ng("A", "認定者数 %s %s" % (t, yn), "行がある", "行がない",
               "町別シートの3の表に年度の行が見当たらない")
            got.append(None)
            continue
        got.append(v[0])
        if v[0] != T.NINTEI[t][y][0]:
            ng("A", "認定者数 %s %s" % (t, yn), T.NINTEI[t][y][0], v[0],
               "計画作成支援ツールの値と一致しない")
        # 要介護度別の内訳の合計＝計
        if len(v) >= 8 and sum(v[1:8]) != v[0]:
            ng("A", "認定者数の内訳 %s %s" % (t, yn), v[0], sum(v[1:8]),
               "要介護度別の内訳の合計が計に一致しない")
    if all(x is not None for x in got):
        z = sum(T.NINTEI[t][y][0] for t in TOWNS)
        if sum(got) != z:
            ng("A", "認定者数の3町計 %s" % yn, z, sum(got),
               "町別の合計が3町計に一致しない")

# A2 施設・居住系サービスの月平均利用者数
SHISETSU_ORDER = [
    "介護老人福祉施設", "地域密着型介護老人福祉施設入所者生活介護",
    "介護老人保健施設", "介護医療院", "介護療養型医療施設",
    "特定施設入居者生活介護", "地域密着型特定施設入居者生活介護",
    "認知症対応型共同生活介護"]
for t in TOWNS:
    rows = block(ALL[TSHEET[t]], "【5　施設・居住系サービスの月平均利用者数")
    for s in SHISETSU_ORDER:
        v = nums(rows, s)
        exp = T.SHISETSU[t].get(s, 0)
        if v is None:
            ng("A", "施設利用者数 %s %s" % (t, s), exp, "行がない",
               "サービスの行が見当たらない")
        elif v[0] != exp:
            ng("A", "施設利用者数 %s %s" % (t, s), exp, v[0],
               "計画作成支援ツールの値と一致しない")
    v = nums(rows, "計")
    exp = sum(T.SHISETSU[t].get(s, 0) for s in SHISETSU_ORDER)
    if v is None or v[0] != exp:
        ng("A", "施設利用者数の計 %s" % t, exp, v[0] if v else "行がない",
           "町内の利用者数の合計が一致しない")

# A2b 比較シートの3町計
rows = ALL["04_3町の比較"]
for s in SHISETSU_ORDER:
    v = nums(rows, s)
    exp = [T.SHISETSU[t].get(s, 0) for t in TOWNS]
    if v is None:
        ng("A", "比較シート 施設利用者数 %s" % s, exp, "行がない",
           "行が見当たらない")
    elif list(v[:3]) != exp or v[3] != sum(exp):
        ng("A", "比較シート 施設利用者数 %s" % s, exp + [sum(exp)], list(v[:4]),
           "町別の値又は3町計が一致しない")

# A3 地域支援事業費
for k in ["介護予防・日常生活支援総合事業費", "包括的支援事業及び任意事業費",
          "地域支援事業費"]:
    v = nums(ALL["04_3町の比較"], k)
    exp = [T.HOKENRYO[t][k + "（円）"] for t in TOWNS]
    if v is None:
        ng("A", "地域支援事業費 %s" % k, exp, "行がない", "行が見当たらない")
    elif list(v[:3]) != exp or v[3] != sum(exp):
        ng("A", "地域支援事業費 %s" % k, exp + [sum(exp)], list(v[:4]),
           "町別の値又は3町計が一致しない")

# A4 交付金
for t in TOWNS:
    rows = block(ALL[TSHEET[t]], "【9　保険者機能強化推進交付金")
    for k in ["推進合計", "支援合計", "推進・支援合計"]:
        v = nums(rows, k)
        exp = [KZ.MACHI[n][t][k] for n in
               ["令和6年度", "令和7年度", "令和8年度"]]
        if v is None:
            ng("A", "交付金 %s %s" % (t, k), exp, "行がない", "行が見当たらない")
        elif list(v[:3]) != exp:
            ng("A", "交付金 %s %s" % (t, k), exp, list(v[:3]),
               "公表資料の値と一致しない")
        elif v[3] != exp[2] - exp[0]:
            ng("A", "交付金の増減 %s %s" % (t, k), exp[2] - exp[0], v[3],
               "令和6→8年度の増減が一致しない")

# A5 将来人口（見える化システム A系列）の転記と合計
import io                                     # noqa: E402
import runpy                                  # noqa: E402
import data_mieruka_a as MA                   # noqa: E402

_buf, _old = io.StringIO(), sys.stdout
sys.stdout = _buf
try:
    _G = runpy.run_path(RP.ROOT + "/build_projection.py")
finally:
    sys.stdout = _old
BANDS = ["65-74", "75-84", "85+"]
PROJ3_65_2029 = sum(_G["pop_juki"](b, "2029") for b in BANDS)

MA_Y = [2025, 2026, 2027, 2028, 2029, 2030, 2040, 2045, 2050]
MA_ROWS = [("A1", "総人口", "総人口（人）"),
           ("A3", "高齢者数", "65歳以上（人）"),
           ("A2", "高齢化率", "高齢化率（％）"),
           ("A9", "高齢者１人あたり現役世代数",
            "高齢者1人あたり現役世代数（％）")]
for t in TOWNS:
    rows = block(ALL[TSHEET[t]], "【2　高齢者人口の推移と将来推計")
    if not rows:
        ng("A", "将来人口の表 %s" % t, "表がある", "ない",
           "見える化システムA系列の表が見当たらない")
        continue
    for series, shihyo, label in MA_ROWS:
        v = nums(rows, label)
        exp = [MA.val(series, t, shihyo, y) for y in MA_Y]
        exp = [x for x in exp if x is not None]
        if v is None:
            ng("A", "将来人口 %s %s" % (t, label), exp, "行がない",
               "行が見当たらない")
        elif [float(x) for x in v[:len(exp)]] != [float(x) for x in exp]:
            ng("A", "将来人口 %s %s" % (t, label), exp, list(v[:len(exp)]),
               "見える化システムA系列の値と一致しない")

# 3町の合計＝保険者の値
for y in MA_Y:
    z = MA.val("A3", "大雪地区広域連合", "高齢者数", y)
    s3 = sum(MA.val("A3", t, "高齢者数", y) for t in TOWNS)
    if z is None or abs(s3 - z) > 0.5:
        ng("A", "将来人口の3町計 %d" % y, z, s3,
           "町別の合計が保険者の値に一致しない")

# 世帯（A5〜A8）
MA_SETAI_Y = [2000, 2005, 2010, 2015, 2020]
SETAI_ROWS = [("A5", "一般世帯数", "一般世帯数（世帯）"),
              ("A6", "高齢者を含む世帯数", "高齢者を含む世帯数（世帯）"),
              ("A6-a", "高齢者を含む世帯の割合",
               "高齢者を含む世帯の割合（％）"),
              ("A7", "高齢独居世帯数", "高齢独居世帯数（世帯）"),
              ("A7-a", "高齢独居世帯の割合", "高齢独居世帯の割合（％）"),
              ("A8", "高齢夫婦世帯数", "高齢夫婦世帯数（世帯）"),
              ("A8-a", "高齢夫婦世帯の割合", "高齢夫婦世帯の割合（％）")]
for t in TOWNS:
    rows = block(ALL[TSHEET[t]], "【2の2　世帯の状況")
    if not rows:
        ng("A", "世帯の表 %s" % t, "表がある", "ない",
           "世帯の表が見当たらない")
        continue
    for series, shihyo, label in SETAI_ROWS:
        v = nums(rows, label)
        exp = [MA.val(series, t, shihyo, y) for y in MA_SETAI_Y]
        exp = [x for x in exp if x is not None]
        if v is None:
            ng("A", "世帯 %s %s" % (t, label), exp, "行がない",
               "行が見当たらない")
        elif [float(x) for x in v[:len(exp)]] != [float(x) for x in exp]:
            ng("A", "世帯 %s %s" % (t, label), exp, list(v[:len(exp)]),
               "見える化システムA5〜A8の値と一致しない")

# 04シートの1の2・1の3
for y in [2025, 2029, 2040, 2045]:
    lbl = "%s（%d）　65歳以上" % ({2025: "令和7年", 2029: "令和11年",
                                   2040: "令和22年", 2045: "令和27年"}[y], y)
    v = nums(ALL["04_3町の比較"], lbl)
    exp = [MA.val("A3", t, "高齢者数", y) for t in TOWNS]
    exp.append(MA.val("A3", "大雪地区広域連合", "高齢者数", y))
    if v is None:
        ng("A", "比較シート 将来人口 %d" % y, exp, "行がない",
           "行が見当たらない")
    elif [float(x) for x in v[:4]] != [float(x) for x in exp]:
        ng("A", "比較シート 将来人口 %d" % y, exp, list(v[:4]),
           "見える化システムA3の値と一致しない")

# ============================================================ B 転記の一致
# B1 健康とくらしの調査
wbc = openpyxl.load_workbook(CROSS, read_only=True, data_only=True)
wsc = wbc["01_地区別の主要指標"]
crows = list(wsc.iter_rows(values_only=True))
chead, cdata, cur = None, {}, None
started = False
for r in crows:
    if r and r[0] == "町":
        chead = [("" if v is None else str(v).replace("\n", "")) for v in r]
        started = True
        continue
    if not started or not r:
        continue
    if r[0] and str(r[0]).startswith("注"):
        break
    if r[0]:
        cur = r[0]
    if cur is None or r[1] is None:
        continue
    cdata.setdefault(cur, {})[str(r[1])] = list(r[2:])
wbc.close()

for t in TOWNS:
    rows = block(ALL[TSHEET[t]], "【8　健康とくらしの調査")
    if not rows:
        ng("B", "健康とくらしの調査 %s" % t, "表がある", "表がない",
           "8の表が見当たらない")
        continue
    hdr = rows[0]
    for r in rows[1:]:
        chiku = None if r[0] is None else str(r[0])
        if chiku is None or chiku not in cdata.get(t, {}):
            if chiku:
                ng("B", "健康とくらしの調査 %s %s" % (t, chiku),
                   "原本にある地区", "原本にない",
                   "調査クロス集計・分析01シートに同じ地区がない")
            continue
        orig = cdata[t][chiku]
        for j in range(1, len(hdr)):
            if hdr[j] is None:
                continue
            name = str(hdr[j]).replace("\n", "")
            if name not in chead:
                ng("B", "健康とくらしの調査の列名 %s" % name,
                   "原本の列名", name, "原本に同じ列名がない")
                continue
            oi = chead.index(name) - 2
            if r[j] is None or orig[oi] is None:
                continue
            if abs(float(r[j]) - float(orig[oi])) > 1e-9:
                ng("B", "健康とくらしの調査 %s %s %s" % (t, chiku, name),
                   orig[oi], r[j], "原本の値と一致しない")

# B2 給付費の実績
for t in TOWNS:
    rows = block(ALL[TSHEET[t]], "【4　給付費の実績】")
    for y, yn in [("R6", "令和6年度"), ("R7", "令和7年度"),
                  ("R8", "令和8年度（年度途中）")]:
        v = nums(rows, yn)
        exp = KJ.KYUFU_TOTAL[(t, y)]
        if v is None:
            ng("B", "給付費 %s %s" % (t, yn), exp[7], "行がない",
               "年度の行が見当たらない")
        elif v[0] != exp[7]:
            ng("B", "給付費 %s %s" % (t, yn), exp[7], v[0],
               "給付実績データの総額と一致しない")
        elif len(v) >= 8 and sum(v[1:8]) != v[0]:
            ng("B", "給付費の内訳 %s %s" % (t, yn), v[0], sum(v[1:8]),
               "要介護度別の内訳の合計が総額に一致しない")

# B3 論点の転記
wbr = openpyxl.load_workbook(RONTEN, read_only=True, data_only=True)
wsr = wbr["04_町ごとの論点"]
ron, started = {}, False
for r in wsr.iter_rows(values_only=True):
    if r and r[0] == "No.":
        started = True
        continue
    if not started or not r or r[0] is None:
        continue
    if isinstance(r[0], str) and r[0].startswith("注"):
        break
    ron.setdefault(r[1], []).append(str(r[2]))
wbr.close()
for t in TOWNS:
    rows = block(ALL[TSHEET[t]], "【10　")
    got = [str(r[0]) for r in rows[1:] if r[0]]
    if got != ron.get(t, []):
        ng("B", "論点の転記 %s" % t, ron.get(t, []), got,
           "3町別の論点整理04シートの論点と一致しない")

# B4 住民基本台帳
for t in TOWNS:
    rows = block(ALL[TSHEET[t]], "【1　人口と高齢化")
    for label, band in [("総人口", "tot"), ("65〜74歳", "65-74"),
                        ("75〜84歳", "75-84"), ("85歳以上", "85+"),
                        ("65歳以上", "65+"), ("75歳以上", "75+")]:
        v = nums(rows, label)
        exp = [juki(t, y, band) for y in sorted(DJ.JUKI[t])]
        if v is None:
            ng("B", "住基 %s %s" % (t, label), exp, "行がない",
               "行が見当たらない")
        elif list(v[:len(exp)]) != exp:
            ng("B", "住基 %s %s" % (t, label), exp, list(v[:len(exp)]),
               "住民基本台帳の値と一致しない")

# B5 比較シートの健康とくらしの調査
rows4 = ALL["04_3町の比較"]
for key in ["フレイルあり割合", "社会参加あり割合（月1回以上）",
            "通いの場参加割合（月1回以上）", "独居者割合"]:
    if key not in chead:
        continue
    oi = chead.index(key) - 2
    v = nums(rows4, key)
    exp = [cdata[t]["町計"][oi] for t in TOWNS]
    exp.append(cdata["大雪地区広域連合"]["全体"][oi])
    if v is None:
        ng("B", "比較シート 健康とくらし %s" % key, exp, "行がない",
           "行が見当たらない")
    elif [float(x) for x in v[:4]] != [float(x) for x in exp]:
        ng("B", "比較シート 健康とくらし %s" % key, exp, list(v[:4]),
           "調査クロス集計・分析01シートの値と一致しない")

# ============================================================ C 算定方法
NEED = ["サービス見込量", "給付費の見込み", "保険料", "第1号被保険者数",
        "要介護認定率", "将来人口", "地域支援事業費",
        "施設・居住系サービスの利用者数", "施設・居住系サービスの定員",
        "総合事業", "交付金の評価"]
rows5 = ALL["05_算定方法の整理"]
have5 = [str(r[1]) for r in rows5 if r[1] is not None]
for k in NEED:
    if k not in have5:
        ng("C", "算定方法の対照 %s" % k, "05シートに項目がある", "ない",
           "計画素案・国の様式・見える化との対照が欠けている")
for r in rows5:
    if r[1] in NEED:
        for ci, cn in [(2, "国の様式・見える化での扱い"),
                       (3, "計画素案での扱い"), (4, "本表での扱い")]:
            if not r[ci] or len(str(r[ci])) < 10:
                ng("C", "算定方法の対照 %s の%s" % (r[1], cn),
                   "記載がある", "空欄又は短すぎる",
                   "4つの扱いのうち1つが埋まっていない")

for kw, why in [
        ("J＝C＋D－E＋F＋G±H－I", "保険料収納必要額の式が示されていない"),
        ("1.05909", "標準給付費見込額への割増率が示されていない"),
        ("介護保険事業計画作成支援ツール", "国の様式の名称が示されていない"),
        ("見える化", "見える化システムへの言及がない"),
        ("第117条", "計画の根拠条文が示されていない"),
        ("第284条", "広域連合が保険者となる根拠が示されていない")]:
    if kw not in TEXT["05_算定方法の整理"]:
        ng("C", "算定方法の根拠 %s" % kw, "記載がある", "ない", why)

# ============================================================ D 越権
BAN_TOWN = [
    ("保険料基準額", "保険料は保険者単位で条例により定めるものであり、"
     "町別の保険料は存在しない"),
    ("標準給付費見込額", "標準給付費見込額は保険者単位で算定する"),
    ("6,740円", "保険料の暫定額を町別シートに載せている"),
    ("収納必要額", "保険料収納必要額は保険者単位で算定する"),
    ("補正後被保険者数", "補正後被保険者数は保険者単位で算定する"),
]
for kw, why in BAN_TOWN:
    if kw in TOWN_TEXT:
        ng("D", "町別シートの記載 %s" % kw, "町別シートに現れない",
           "現れる", why)

for t in TOWNS:
    txt = TEXT[TSHEET[t]]
    if "見込量" in txt and "掲げていません" not in txt and \
            "本シートには掲げていません" not in txt:
        ng("D", "町別シートの記載 見込量 %s" % t,
           "掲げない旨の注記がある", "注記がない",
           "見込量という語を使うなら町別には掲げない旨を明示する")

# 04シートにも見込量・保険料の数値がないこと
for kw in ["保険料基準額", "標準給付費見込額", "6,740円"]:
    if kw in TEXT["04_3町の比較"]:
        ng("D", "比較シートの記載 %s" % kw, "現れない", "現れる",
           "保険者単位の値を町別の比較表に載せている")

# ============================================================ E 体裁
# E1 出典の明記
for s in [TSHEET[t] for t in TOWNS]:
    n_lead = TEXT[s].count("【")
    n_src = TEXT[s].count("資料：")
    if n_src < n_lead - 1:
        ng("E", "出典の明記 %s" % s, "表の数（%d）に見合う資料欄" % n_lead,
           "%d件" % n_src, "「資料：」の欄がない表がある")

# E2 禁止表現
#   06シートは点検の方法として禁止表現そのものを引用するため、
#   その1行だけを除いて走査する。
BODY_SHEETS = [s for s in wb.sheetnames if s != "06_自己点検の結果"]
BODY_TEXT = "\n".join(TEXT[s] for s in BODY_SHEETS)
for w in ["に由来する", "と整合する", "1件も", "有意差がないため",
          "全国トップ級"]:
    if w in BODY_TEXT:
        ng("E", "禁止する表現 %s" % w, "現れない", "現れる",
           "受託者が用いないと定めた表現である")
if "禁止する表現" not in TEXT["06_自己点検の結果"]:
    ng("E", "自己点検の項目 禁止する表現", "06シートに点検項目がある",
       "ない", "禁止する表現の点検が自己点検に掲げられていない")

# E3 個人情報
#   語の有無ではなく、個人が特定される形の値が現れるかを見る。
#   00シートと06シートは「収録していない」旨を述べるために
#   語そのものを用いるため、語の有無では判定できない。
PII = [
    (re.compile(r"\d{2,4}-\d{2,4}-\d{4}"), "電話番号の形をした文字列"),
    (re.compile(r"[\w.+-]+@[\w-]+\.[\w.]+"), "メールアドレス"),
    (re.compile(r"(明治|大正|昭和)\d+年\d+月\d+日"), "生年月日の形をした日付"),
    (re.compile(r"(?<!\d)\d{10}(?!\d)"), "被保険者番号・事業所番号の形をした数字"),
]
for s in wb.sheetnames:
    for row in wb[s].iter_rows(values_only=True):
        for v in row:
            if not isinstance(v, str):
                continue
            for pat, why in PII:
                m = pat.search(v)
                if m:
                    ng("E", "個人情報 %s（%s）" % (why, s), "現れない",
                       m.group(0), "個人が特定される形の値は収録しない")
# 収録しない旨の記載があること
if "収録していません" not in TEXT["00_この資料について"]:
    ng("E", "個人情報の取扱いの記載", "00シートに記載がある", "ない",
       "個人情報を収録していない旨を明記する")

# E4 時点の明記
for kw in ["令和8年6月30日現在", "各年1月1日現在", "各年度末",
           "令和7年11月17日"]:
    if kw not in WHOLE:
        ng("E", "時点の明記 %s" % kw, "記載がある", "ない",
           "出典の時点が示されていない")

# E5 令和8年度の給付費が年度途中である旨
if "年度途中" not in TOWN_TEXT:
    ng("E", "令和8年度の給付費", "年度途中である旨の注記がある", "ない",
       "通年の値と誤読されるおそれがある")

# E6 認定率の定義の違いの注記
for t in TOWNS:
    if "第1号被保険者数を分母" not in TEXT[TSHEET[t]]:
        ng("E", "認定率の定義 %s" % t, "分母の違いの注記がある", "ない",
           "広域連合の認定率と分母が異なることを示していない")

# E7 06シートの判定
rows6 = ALL["06_自己点検の結果"]
n_hantei = sum(1 for r in rows6
               if r and len(r) > 5 and r[5] == "適" and r[0] is not None)
if n_hantei < 20:
    ng("E", "自己点検の結果", "20件以上の判定がある", "%d件" % n_hantei,
       "06シートの点検項目が実際の点検を覆っていない")
_kubun6 = {str(r[1]).split()[0] for r in rows6
           if r and r[1] is not None and str(r[1])[:1] in "ABCDEFGH"}
for k in "ABCDEFGH":
    if k not in _kubun6:
        ng("E", "自己点検の区分 %s" % k, "06シートに区分がある", "ない",
           "実際に行っている点検の区分が自己点検に掲げられていない")

# ============================================================ F 出所の差
# 同じ事項に複数の出所があり値が異なるものについて、
# 差に触れているか、触れているなら数値が正しいかを見る。
import data_kessan_r6 as KK                    # noqa: E402
import data_nenpo as NN                        # noqa: E402

TOOL_CHIIKI = sum(T.HOKENRYO[t]["地域支援事業費（円）"] for t in TOWNS)
KESSAN_CHIIKI = KK.SAISHUTSU["4 地域支援事業費"][1]
TOOL_NIN_R7 = sum(T.NINTEI[t]["R7"][0] for t in TOWNS)
NENPO_NIN1_R7 = NN.NINTEI["第1号被保険者\u3000合計"]["R7"]
NENPO_NIN_R7 = NN.NINTEI["総数（第1号＋第2号）"]["R7"]
TOOL_GH = sum(T.SHISETSU[t].get("認知症対応型共同生活介護", 0) for t in TOWNS)
KYUFU_R6 = KJ.KYUFU_TOTAL[("大雪", "R6")][7]
KESSAN_KYUFU = KK.SAISHUTSU["2 保険給付費"][1] \
    if "2 保険給付費" in KK.SAISHUTSU else None

F = [
    ("地域支援事業費の出所の差",
     [format(TOOL_CHIIKI, ","), format(KESSAN_CHIIKI, ",")],
     "計画作成支援ツールの3町計と介護保険特別会計決算とで額が異なる。"
     "将来推計は決算を用いているため、差に触れる必要がある"),
    ("要介護認定者数の出所の差",
     [str(TOOL_NIN_R7), str(NENPO_NIN1_R7), str(NENPO_NIN_R7)],
     "計画作成支援ツールと介護保険事業状況報告とで人数が異なる。"
     "第1号のみか第2号を含むかでも異なる"),
    ("認知症対応型共同生活介護の定員の出所の差",
     ["99", "108"],
     "見える化D26（5事業所99人）と北海道の名簿（7事業所108人＋要確認）とで"
     "定員が異なる。計画作成支援ツールの値は月平均利用者数であり"
     "定員ではない"),
    ("将来人口の推計の基礎の差",
     ["{:,.0f}".format(MA.val("A3", "大雪地区広域連合", "高齢者数", 2029)),
      "{:,.0f}".format(PROJ3_65_2029)],
     "見える化システム（社人研が基礎）と計画素案の案C"
     "（総合戦略・住民基本台帳の実績趨勢が基礎）とで"
     "3町計の令和11年度の65歳以上人口が異なる"),
    ("給付費の実績の出所の差",
     [format(KYUFU_R6, ",")],
     "給付費データ集計の総計と決算の保険給付費とは範囲が異なる。"
     "代表KPI H15の分母の定義は確定していない"),
]
for nm, needs, why in F:
    hit = [x for x in needs if x in WHOLE]
    if len(hit) < len(needs):
        ng("F", nm, "%s のすべてが記載されている" % "／".join(needs),
           "記載は %s" % ("／".join(hit) if hit else "なし"), why)

# F5 決算・年報・ツールの3つの名称が本表に現れること
for kw, why in [
        ("介護保険事業計画作成支援ツール", "国の様式の名称"),
        ("介護保険事業状況報告", "年報・月報の正式名称"),
        ("介護保険特別会計", "決算の名称"),
        ("地域包括ケア「見える化」システム", "見える化システムの正式名称")]:
    if kw not in WHOLE:
        ng("F", "出所の名称 %s" % kw, "記載がある", "ない",
           "%s が示されていない" % why)

# F6 表題の語と出所の意味の一致
# 令和8年9月10日の点検で、5の表題を「施設・居住系の定員」としながら
# 収録していた値が計画作成支援ツールの月平均利用者数であったことが分かった。
# 同種の取り違えを検出するため、表題に「定員」を含む表の
# 「資料：」の欄が実績値シート（利用者数の出所）を指していないことを見る。
RIYOSHA_SRC = "介護保険事業計画作成支援ツール"
for sname, rows in ALL.items():
    head = None
    for r in rows:
        a = "" if r[0] is None else str(r[0])
        if a.startswith("【"):
            head = a
        elif a.startswith("資料：") and head and "定員" in head:
            if RIYOSHA_SRC in a and "定員" not in a:
                ng("F", "%s %s" % (sname, head),
                   "定員の出所（北海道の名簿・見える化システム等）",
                   a,
                   "表題は定員だが、出所が計画作成支援ツールの実績値"
                   "（月平均利用者数）である")
            head = None

# ============================================================ G 施策
# 施策は町ではなく広域連合が定めるものであるため、
# 町別データシートに施策の体系を掲げていないこと。
for kw, why in [
        ("基本目標1", "施策の体系は計画本文で定めるものである"),
        ("代表KPI H", "代表KPIは広域連合の指標である")]:
    if kw in TOWN_TEXT:
        ng("G", "町別シートの記載 %s" % kw, "現れない", "現れる", why)
if "施策" not in TEXT["05_算定方法の整理"]:
    ng("G", "施策の扱い", "05シートに施策の扱いの記載がある", "ない",
       "施策を町別に掲げない理由が示されていない")

# ============================================================ H 深掘り
# H1 保険料の算定の記述が実際の算定と一致すること
_row5 = {str(r[1]): r for r in ALL["05_算定方法の整理"] if r[1] is not None}
for nm, needs in [
    ("保険料", ["J＝C＋D－E＋F＋G±H－I", "予定収納率", "6,740円"]),
    ("給付費の見込み", ["標準給付費見込額", "特定入所者介護サービス費",
                        "高額介護サービス費", "高額医療合算",
                        "算定対象審査支払手数料", "市町村特別給付費"]),
    ("サービス見込量", ["介護保険事業計画作成支援ツール", "見える化",
                        "保険者"]),
]:
    if nm not in _row5:
        continue
    txt = "".join(str(x) for x in _row5[nm] if x is not None)
    for k in needs:
        if k not in txt:
            ng("H", "05シート %s の記述" % nm, "%s に触れている" % k,
               "触れていない",
               "国の様式の算定の構成要素が示されていない")

# H2 04シートの3町計が3町別の論点整理と一致すること
wbr2 = openpyxl.load_workbook(RONTEN, read_only=True, data_only=True)
wsr2 = wbr2["01_町別の現在地"]
r01 = {}
for r in wsr2.iter_rows(values_only=True):
    if r and r[0] and isinstance(r[0], str):
        r01[r[0].strip()] = list(r[1:5])
wbr2.close()
for label, key in [("介護老人福祉施設", "介護老人福祉施設"),
                   ("介護老人保健施設", "介護老人保健施設"),
                   ("認知症対応型共同生活介護", "認知症対応型共同生活介護"),
                   ("特定施設入居者生活介護", "特定施設入居者生活介護")]:
    if key not in r01:
        continue
    exp = [v for v in r01[key][:4]]
    got = nums(ALL["04_3町の比較"], label)
    if got is None:
        continue
    if [x for x in got[:4]] != exp:
        ng("H", "3町別の論点整理との一致 %s" % label, exp, list(got[:4]),
           "他の成果品の同じ表と値が食い違う")

# H3 交付金の得点が3町別の論点整理と一致すること
for n in ["令和6年度", "令和7年度", "令和8年度"]:
    key = "%s　合計得点" % n
    if key not in r01:
        continue
    exp = list(r01[key][:4])
    got = nums(ALL["04_3町の比較"], n + "　推進・支援合計")
    if got is None:
        ng("H", "交付金の得点 %s" % n, exp, "行がない", "行が見当たらない")
    elif [x for x in got[:4]] != exp:
        ng("H", "交付金の得点 %s" % n, exp, list(got[:4]),
           "3町別の論点整理01シートと食い違う")

# H4 保険者単位であることの説明が町別シートにあること
for t in TOWNS:
    if "保険者単位" not in TEXT[TSHEET[t]]:
        ng("H", "保険者単位の説明 %s" % t, "記載がある", "ない",
           "町別に掲げない理由が町別シートに書かれていない")

# H5 将来人口の基礎の違いの注記
for t in TOWNS:
    txt = TEXT[TSHEET[t]]
    for k in ["見える化", "社人研", "案C"]:
        if k not in txt:
            ng("H", "将来人口の注記 %s %s" % (t, k), "記載がある", "ない",
               "推計の基礎と計画素案との違いが示されていない")
    if "按分" in txt:
        ng("H", "将来人口の注記 %s" % t, "按分の記載が残っていない",
           "残っている",
           "見える化システムの町別の値に置き換えたため按分は行っていない")

# H6 00シートの一覧が実際のシート構成と一致すること
listed = [str(r[1]) for r in ALL["00_この資料について"]
          if r[1] is not None and re.match(r"^0[1-6]_", str(r[1]))]
actual = [s for s in wb.sheetnames if re.match(r"^0[1-6]_", s)]
if listed != actual:
    ng("H", "00シートのシート一覧", actual, listed,
       "実際のシート構成と一覧が食い違う")

# H7 06シートの点検項目数が実際の点検の区分を覆っていること
t6 = TEXT["06_自己点検の結果"]
for k in ["合計", "転記", "算定方法", "出典", "個人情報", "禁止"]:
    if k not in t6:
        ng("H", "自己点検の項目 %s" % k, "06シートに項目がある", "ない",
           "実際に行っている点検が自己点検に掲げられていない")

# H8 令和8年度の認定者数が年度途中でないこと（ツールは年度末の値）
for t in TOWNS:
    rows = block(ALL[TSHEET[t]], "【3　要介護認定者数】")
    v = nums(rows, "令和8年度")
    if v is not None and v[0] != T.NINTEI[t]["R8"][0]:
        ng("H", "認定者数 %s 令和8年度" % t, T.NINTEI[t]["R8"][0], v[0],
           "計画作成支援ツールの値と一致しない")

# ============================================================ 出力
print("=" * 72)
print("町別データシート　レッドチームレビュー")
print("=" * 72)
if not FINDINGS:
    print("指摘 0件。レビューを通過した。")
    print()
    print("点検した項目")
    print("  A 合計の一致　認定者数・施設利用者数・地域支援事業費・交付金・"
          "将来人口・世帯")
    print("  B 転記の一致　健康とくらしの調査・給付費・論点・住民基本台帳")
    print("  C 算定方法　　10項目の4欄・根拠6件")
    print("  D 越権　　　　保険者単位の値を町別に掲げていないこと")
    print("  E 体裁　　　　出典・時点・禁止表現・個人情報・注記")
    print("  F 出所の差　　値が異なる4件の記載と数値")
    print("  G 施策　　　　施策を町別に掲げていないこと")
    print("  H 深掘り　　　他の成果品との一致・05シートの記述の中身")
    sys.exit(0)

print("指摘 %d件" % len(FINDINGS))
print()
for i, (k, item, exp, act, why) in enumerate(FINDINGS, start=1):
    print("[%s] %d %s" % (k, i, item))
    print("     あるべき値： %s" % exp[:160])
    print("     実際の値　： %s" % act[:160])
    print("     指摘　　　： %s" % why)
sys.exit(1)
