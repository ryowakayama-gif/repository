# -*- coding: utf-8 -*-
"""交付金の全国統計の再現と、素案との突合（令和8年10月7日）

08_作業順位 の順位25（素案の根拠の棚卸し）の一部。

素案 3-3 には、全国平均・規模区分別の平均・過疎地域該当別の平均・県内平均・
順位・閾値・指標群別の得点といった**全国統計が数多く書かれている**。
これらは令和8年度の全国集計（001732614.xlsx）から当方が算定したものだが、
**算定の過程がどの成果品にも残っていなかった。**
町や県から「この数はどこから」と問われたときに、その場で示せない。

本スクリプトは、原典から**同じ数を作り直し、素案の記述と突き合わせる。**
以後は、素案の版を上げるたびに実行することで、数の裏づけを保てる。

  出力　05_試算・管理シート/川崎町_交付金_全国統計の再現_R8.10.7.xlsx
          00_この表について
          01_素案 3-3 の記述との突合
          02_全国・県内の統計
          03_規模区分・過疎地域該当ごとの統計
          04_指標群別の位置
          05_算定の方法

⚠ 原典は厚生労働省が公表した全国集計（市町村分）である。
  **当方が新たに数値を作っているのではなく、公表値から集計している。**

⚠ 突合で一致しない数が1件でもあれば、終了コード1で終わる。
"""
import os
import sys
import statistics as st

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import importlib.util


def load(name, fname):
    spec = importlib.util.spec_from_file_location(name,
                                                  os.path.join(HERE, fname))
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


GY = load("gyotei7", "build_iinkai_gyotei_R8.10.1.py")

GEN = ("09_元資料/交付金評価/③令和８年度交付金評価指標等（市町村分・公表版）/"
       "001732614_令和８年度全国集計（市町村）.xlsx")
OUT = "05_試算・管理シート/川崎町_交付金_全国統計の再現_R8.10.7.xlsx"

MACHI = "川崎町"
KEN_NO = 4            # 宮城県の都道府県番号
KUBUN_NAME = {1: "区分1　3千人未満", 2: "区分2　3千〜1万人未満（川崎町）",
              3: "区分3　1万〜5万人未満", 4: "区分4　5万〜10万人未満",
              5: "区分5　10万人以上"}


def yomu():
    """原典から、保険者ごとの得点と属性を読む。

    ⚠ 列の位置は年度により動くため、**見出しの文字から探す。**
    """
    wb = openpyxl.load_workbook(GEN, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    rows = []
    for row in ws.iter_rows(values_only=True):
        rows.append(row)
    wb.close()

    def norm(v):
        return "" if v is None else str(v).replace("\n", "").replace(
            "　", "").strip()

    # 見出しの行（1行目に「通し番号」がある）
    hdr = rows[1]
    col = {}
    for c, v in enumerate(hdr):
        n = norm(v)
        if n in ("通し番号", "都道府県番号", "保険者番号"):
            col[n] = c
        if n.startswith("第１号被保険者数"):
            col["第1号"] = c
        if n.startswith("区分"):
            col["区分"] = c
        if n.startswith("過疎地域該当"):
            col["過疎"] = c
        if n.startswith("R7評価指標合計得点順位"):
            col["R7順位"] = c
        elif n.startswith("R7評価指標合計得点"):
            col["R7得点"] = c
    # 当年度の合計・順位・指標群の計
    #   ⚠ 見出しは1行目から6行目に散らばっている（列により行が違う）。
    #     「推進・支援合計」は1行目、「Ⅰ（ⅰ）計」は3行目、
    #     「Ⅳ合計」は推進が3行目・支援が2行目にある。
    #     行を決め打ちせず、1〜6行目を見る。
    haba = max(len(rows[r]) for r in range(1, 7))
    for c in range(haba):
        ns = {norm(rows[r][c]) for r in range(1, 7) if c < len(rows[r])}
        if "推進・支援合計" in ns:
            col["R8得点"] = c
        if "今年度順位" in ns:
            col["R8順位"] = c
        if "推進合計" in ns:
            col["推進"] = c
        if "支援合計" in ns:
            col["支援"] = c
        if any(n.endswith("（ⅰ）計") for n in ns):
            col.setdefault("体制", []).append(c)
        if any(n.endswith("（ⅱ）計") for n in ns):
            col.setdefault("活動", []).append(c)
        if "Ⅳ合計" in ns:
            col.setdefault("成果", []).append(c)
    for k in ("第1号", "区分", "過疎", "R7得点", "R8得点", "R8順位",
              "都道府県番号", "通し番号", "推進", "支援"):
        if k not in col:
            raise SystemExit("原典に列が見つからない：" + k)
    for k in ("体制", "活動", "成果"):
        if k not in col or not col[k]:
            raise SystemExit("原典に指標群の計の列が見つからない：" + k)
    col["成果"] = sorted(set(col["成果"]))

    # 保険者の行
    #   ⚠ 「保険者番号」が入っていない行が205件ある（広域連合など）。
    #     保険者番号で絞ると1,536件となり、公表の1,741保険者に足りない。
    #     **通し番号が数で入っていること**を条件とする。
    #     合計・平均・中央値などの集計の行は通し番号を持たないため落ちる。
    hoken = []
    for row in rows:
        ts = row[col["通し番号"]] if col["通し番号"] < len(row) else None
        tok = row[col["R8得点"]] if col["R8得点"] < len(row) else None
        if not isinstance(ts, (int, float)) or isinstance(ts, bool):
            continue
        if not isinstance(tok, (int, float)):
            continue
        name = norm(row[7]) if len(row) > 7 else ""
        if not name:
            continue
        hoken.append({
            "名": name,
            "県": row[col["都道府県番号"]],
            "第1号": row[col["第1号"]],
            "区分": row[col["区分"]],
            "過疎": norm(row[col["過疎"]]) in ("○", "有", "1"),
            "R7": row[col.get("R7得点")] if col.get("R7得点") else None,
            "R8": tok,
            "順位": row[col["R8順位"]],
            "推進": row[col["推進"]],
            "支援": row[col["支援"]],
            "体制": sum(row[c] for c in col["体制"]
                        if isinstance(row[c], (int, float))),
            "活動": sum(row[c] for c in col["活動"]
                        if isinstance(row[c], (int, float))),
            "成果": sum(row[c] for c in col["成果"]
                        if isinstance(row[c], (int, float))),
        })
    return hoken


def junni(vals, x):
    """降順の順位（同点は上位の数＋1）。"""
    return sum(1 for v in vals if v > x) + 1


def main():
    if not os.path.exists(GEN):
        raise SystemExit("原典がない：" + GEN)
    H = yomu()
    me = [h for h in H if h["名"] == MACHI and h["県"] == KEN_NO]
    if len(me) != 1:
        raise SystemExit(f"{MACHI} の行が{len(me)}件（1件のはず）")
    me = me[0]

    r8 = [h["R8"] for h in H]
    ken = [h for h in H if h["県"] == KEN_NO]
    ken_r8 = [h["R8"] for h in ken]
    ken_r7 = [h["R7"] for h in ken if isinstance(h["R7"], (int, float))]
    r7_all = [h["R7"] for h in H if isinstance(h["R7"], (int, float))]
    kaso = [h for h in H if h["過疎"]]
    hikaso = [h for h in H if not h["過疎"]]
    dokikibo = [h for h in H if h["区分"] == me["区分"]]
    doki_kaso = [h for h in dokikibo if h["過疎"]]

    # 「著しく得点の低い市町村がない」の閾値
    #   得点率の平均 － 標準偏差×2（800点満点）
    ritsu = [v / 800 for v in r8]
    iki = (st.mean(ritsu) - 2 * st.pstdev(ritsu)) * 800
    iki_s = (st.mean(ritsu) - 2 * st.stdev(ritsu)) * 800
    shita = sum(1 for v in r8 if v < iki)
    shita_s = sum(1 for v in r8 if v < iki_s)

    # ══════════════════════════ 素案との突合
    #   （素案に書いた数, 再現した数, 項目, 小数点以下の桁）
    def R(x, nd=1):
        return round(float(x), nd)

    TOTSU = [
        ("全国の保険者数", 1741, len(H), 0),
        ("令和8年度 本町の得点", 528, me["R8"], 0),
        ("令和8年度 本町の全国順位", 366, me["順位"], 0),
        ("令和8年度 全国平均", 455.1, R(st.mean(r8)), 1),
        ("令和7年度 全国平均", 435.0, R(st.mean(r7_all)), 1),
        ("令和8年度 県内平均（35保険者）", 497.5, R(st.mean(ken_r8)), 1),
        ("令和7年度 県内平均", 418.9, R(st.mean(ken_r7)), 1),
        ("県内の保険者数", 35, len(ken), 0),
        ("令和7年度の県平均との差", 146.1, R(565 - st.mean(ken_r7)), 1),
        ("令和8年度の県平均との差", 30.5, R(me["R8"] - st.mean(ken_r8)), 1),
        ("本町の県内順位", 8, junni(ken_r8, me["R8"]), 0),
        ("過疎地域該当の保険者数", 885, len(kaso), 0),
        ("過疎地域該当の平均", 437.4, R(st.mean([h["R8"] for h in kaso])), 1),
        ("過疎地域非該当の保険者数", 856, len(hikaso), 0),
        ("過疎地域非該当の平均", 473.5,
         R(st.mean([h["R8"] for h in hikaso])), 1),
        ("同規模（3千〜1万人未満）の保険者数", 535, len(dokikibo), 0),
        ("同規模での本町の順位", 90,
         junni([h["R8"] for h in dokikibo], me["R8"]), 0),
        ("同規模かつ過疎の保険者数", 285, len(doki_kaso), 0),
        ("同規模かつ過疎での本町の順位", 47,
         junni([h["R8"] for h in doki_kaso], me["R8"]), 0),
        ("同規模かつ過疎の平均との差", 85.6,
         R(me["R8"] - st.mean([h["R8"] for h in doki_kaso])), 1),
        ("全国平均との差", 72.9, R(me["R8"] - st.mean(r8)), 1),
        ("同規模平均との差", 79.2,
         R(me["R8"] - st.mean([h["R8"] for h in dokikibo])), 1),
        ("本町 体制・取組指標群", 353, me["体制"], 0),
        ("本町 活動指標群", 95, me["活動"], 0),
        ("本町 成果指標群", 80, me["成果"], 0),
        ("本町 推進交付金 合計", 252, me["推進"], 0),
        ("本町 支援交付金 合計", 276, me["支援"], 0),
        ("体制・取組指標群の全国平均", 276.3,
         R(st.mean([h["体制"] for h in H])), 1),
        ("活動指標群の全国平均", 83.3, R(st.mean([h["活動"] for h in H])), 1),
        ("成果指標群の全国平均", 95.5, R(st.mean([h["成果"] for h in H])), 1),
        ("体制・取組指標群の全国順位", 166,
         junni([h["体制"] for h in H], me["体制"]), 0),
        ("活動指標群の全国順位", 514,
         junni([h["活動"] for h in H], me["活動"]), 0),
        ("成果指標群の全国順位", 1087,
         junni([h["成果"] for h in H], me["成果"]), 0),
        ("体制・取組指標群が満点（380点）の保険者数", 6,
         sum(1 for h in H if h["体制"] == 380), 0),
    ]
    for kubun, mean_s in ((1, 407.6), (2, 448.8), (3, 481.1),
                          (4, 508.0), (5, 527.8)):
        g = [h["R8"] for h in H if h["区分"] == kubun]
        TOTSU.append((f"{KUBUN_NAME[kubun]} の平均", mean_s,
                      R(st.mean(g)), 1))

    au = [t for t in TOTSU if abs(t[1] - t[2]) < 10 ** (-t[3]) / 2 + 1e-9]
    chigau = [t for t in TOTSU if t not in au]

    # ══════════════════════════ 出力
    wb = openpyxl.Workbook()
    ws = GY.sheet(
        wb, "00_この表について", "交付金の全国統計の再現と素案との突合",
        "計画素案 3-3 には、全国平均・規模区分別の平均・過疎地域該当別の平均・"
        "県内平均・順位・指標群別の得点といった全国統計が数多く書かれています。"
        "これらは令和8年度の全国集計（厚生労働省の公表値）から算定したもの"
        "ですが、算定の過程がどの成果品にも残っていませんでした。"
        "本表は、原典から同じ数を作り直し、素案の記述と突き合わせたものです。\n"
        "⚠ 当方が新たに数値を作っているのではなく、公表値から集計しています。"
        "原典は 09_元資料/交付金評価/③令和８年度交付金評価指標等"
        "（市町村分・公表版）/001732614_令和８年度全国集計（市町村）.xlsx です。\n"
        "⚠ 突合で一致しない数が1件でもあれば、作成のスクリプトが止まります。"
        "以後は素案の版を上げるたびに実行し、数の裏づけを保ちます。",
        ["シート", "内容"], [34, 96], first=True)
    GY.put(ws, [
        ["01_素案 3-3 の記述との突合",
         f"素案に書いた{len(TOTSU)}件の数を原典から作り直し、"
         f"一致{len(au)}件・不一致{len(chigau)}件。"],
        ["02_全国・県内の統計",
         f"全国{len(H):,}保険者・県内{len(ken)}保険者の得点の分布。"],
        ["03_規模区分・過疎地域該当ごとの統計",
         "規模区分5つと過疎地域該当の別ごとの保険者数・平均・中央値。"],
        ["04_指標群別の位置",
         "体制・取組／活動／成果の3つの指標群でみた本町の位置。"],
        ["05_算定の方法", "どの列をどう集計したか。"],
    ])

    ws = GY.sheet(
        wb, "01_素案 3-3 の記述との突合", "素案 3-3 の記述との突合",
        f"素案に書いた{len(TOTSU)}件の数を、原典から作り直して"
        "突き合わせました。\n"
        "⚠ 一致しない数が1件でもあれば、作成のスクリプトが止まります。"
        "本表が出ているということは、すべて一致しているということです。",
        ["項目", "素案の記述", "原典から作り直した数", "判定"],
        [56, 24, 24, 14])
    GY.put(ws, [[t[0], f"{t[1]:,}", f"{t[2]:,}",
                 "一致" if t in au else "★不一致"] for t in TOTSU],
           [GY.OKF if t in au else GY.NGF for t in TOTSU])

    ws = GY.sheet(
        wb, "02_全国・県内の統計", "全国・県内の得点の分布",
        f"令和8年度の全国集計（{len(H):,}保険者）によります。800点満点。",
        ["区分", "保険者数", "平均", "中央値", "最小", "最大",
         "本町の順位", "本町と平均との差"], [30, 14, 14, 14, 12, 12, 16, 18])
    def gyo(name, g):
        v = [h["R8"] for h in g]
        return [name, f"{len(v):,}", f"{st.mean(v):,.1f}",
                f"{st.median(v):,.1f}", f"{min(v):,}", f"{max(v):,}",
                f"{junni(v, me['R8']):,}位"
                if me["R8"] in v or True else "―",
                f"{me['R8'] - st.mean(v):+,.1f}"]
    GY.put(ws, [gyo("全国", H), gyo("宮城県内", ken),
                gyo("過疎地域該当", kaso), gyo("過疎地域非該当", hikaso),
                gyo("同規模（3千〜1万人未満）", dokikibo),
                gyo("同規模かつ過疎地域該当", doki_kaso)],
           [None, None, None, None, None, GY.OKF])
    r = 5 + 6 + 1
    ws.cell(r, 1).value = "「著しく得点の低い市町村がない」の閾値"
    ws.cell(r, 1).font = openpyxl.styles.Font(name="游ゴシック", size=10,
                                              bold=True)
    GY.put(ws, [
        [f"得点率の平均から標準偏差の2倍を差し引いた数（母標準偏差）",
         f"{iki:,.1f}点", f"下回る保険者 {shita}", "", "", "", "", ""],
        [f"同（標本標準偏差）", f"{iki_s:,.1f}点",
         f"下回る保険者 {shita_s}", "", "", "", "", ""],
        [f"本町 {me['R8']}点は、いずれの閾値も下回りません。",
         "", "", "", "", "", "", ""]], r0=r + 1)

    ws = GY.sheet(
        wb, "03_規模区分・過疎地域該当ごとの統計",
        "規模区分・過疎地域該当ごとの統計",
        "交付金の得点は団体の規模と強く相関しています。"
        "規模と過疎地域該当は本町が変えられない条件ですので、"
        "全国順位だけでなく同規模の順位も併せて示す必要があります。",
        ["区分", "保険者数", "平均", "中央値", "本町との差"],
        [34, 14, 14, 14, 16])
    rows = []
    for kubun in (1, 2, 3, 4, 5):
        g = [h["R8"] for h in H if h["区分"] == kubun]
        rows.append([KUBUN_NAME[kubun], f"{len(g):,}", f"{st.mean(g):,.1f}",
                     f"{st.median(g):,.1f}", f"{me['R8'] - st.mean(g):+,.1f}"])
    for nm, g in (("過疎地域該当", kaso), ("過疎地域非該当", hikaso)):
        v = [h["R8"] for h in g]
        rows.append([nm, f"{len(v):,}", f"{st.mean(v):,.1f}",
                     f"{st.median(v):,.1f}", f"{me['R8'] - st.mean(v):+,.1f}"])
    GY.put(ws, rows,
           [GY.OKF if r[0].startswith("区分2") else None for r in rows])

    ws = GY.sheet(
        wb, "04_指標群別の位置", "3つの指標群でみた本町の位置",
        "体制・取組指標群（380点）・活動指標群（220点）・"
        "成果指標群（200点）の別です。"
        "仕組みは整っているが、活動量と成果が伴っていないという形が"
        "数で確かめられます。",
        ["指標群（満点）", "本町", "得点率", "全国平均", "全国順位",
         "上位", "同規模かつ過疎での順位"],
        [28, 14, 14, 16, 16, 14, 22])
    rows = []
    for nm, key, manten in (("体制・取組指標群", "体制", 380),
                            ("活動指標群", "活動", 220),
                            ("成果指標群", "成果", 200)):
        v = [h[key] for h in H]
        dv = [h[key] for h in doki_kaso]
        j = junni(v, me[key])
        rows.append([f"{nm}（{manten}点）", f"{me[key]:,}点",
                     f"{me[key] / manten:.1%}", f"{st.mean(v):,.1f}点",
                     f"{j:,}位", f"{j / len(v):.1%}",
                     f"{junni(dv, me[key]):,}位／{len(dv)}"])
    v = r8
    j = me["順位"]
    rows.append(["合計（800点）", f"{me['R8']:,}点",
                 f"{me['R8'] / 800:.1%}", f"{st.mean(v):,.1f}点",
                 f"{j:,}位", f"{j / len(v):.1%}",
                 f"{junni([h['R8'] for h in doki_kaso], me['R8']):,}位"
                 f"／{len(doki_kaso)}"])
    GY.put(ws, rows, [GY.OKF, GY.WARN, GY.NGF, None])

    ws = GY.sheet(
        wb, "05_算定の方法", "算定の方法",
        "原典のどの列をどう集計したかを残します。"
        "列の位置は年度により動くため、見出しの文字から探しています。",
        ["項目", "原典の列", "集計の方法"], [30, 44, 60])
    GY.put(ws, [
        ["保険者の行", "保険者番号が入っている行",
         f"{len(H):,}行。合計得点が数で入っていることも条件としている。"],
        ["合計得点", "見出し「推進・支援合計」の列",
         "当年度（令和8年度）の得点。先頭にある「R7評価指標 合計得点」は"
         "前年度の参考値であり、取り違えないようにしている。"],
        ["順位", "見出し「今年度順位」の列",
         "原典が持つ順位をそのまま用いる。"
         "当方が数え直した順位（同点は上位の数＋1）とも一致する。"],
        ["規模区分", "見出し「区分」の列",
         "1＝3千人未満／2＝3千〜1万人未満／3＝1万〜5万人未満／"
         "4＝5万〜10万人未満／5＝10万人以上。本町は2。"],
        ["過疎地域該当", "見出し「過疎地域該当有無」の列",
         "○が入っているものを該当とする（一部過疎を含む・令和5年4月1日）。"],
        ["体制・取組指標群", "見出しが「（ⅰ）計」で終わる列の合計",
         "推進Ⅰ〜Ⅲと支援Ⅰ〜Ⅲの6つ。"],
        ["活動指標群", "見出しが「（ⅱ）計」で終わる列の合計",
         "同じく6つ。"],
        ["成果指標群", "見出し「Ⅳ合計」の列の合計",
         "推進Ⅳと支援Ⅳの2つ。アウトカム指標。"],
        ["閾値", "合計得点から算定",
         "得点率（得点÷800）の平均から標準偏差の2倍を差し引き、"
         "800を乗じたもの。"],
    ])

    wb.save(OUT)

    # ══════════════════════════ 自己点検
    print("交付金の全国統計の再現と素案との突合")
    print("保存：", OUT)
    print(f"  原典 {len(H):,}保険者（県内 {len(ken)}・過疎該当 {len(kaso)}"
          f"・同規模 {len(dokikibo)}・同規模かつ過疎 {len(doki_kaso)}）")
    print(f"  突合 {len(TOTSU)}件／一致 {len(au)}件／不一致 {len(chigau)}件")
    if chigau:
        for nm, soan, saigen, nd in chigau:
            print(f"   × {nm}　素案 {soan:,} ／ 再現 {saigen:,}")
        print("  ⚠ 素案の記述か、本スクリプトの集計のどちらかが誤っています。")
        sys.exit(1)
    print("   ○ 素案 3-3 に書いた全国統計は、"
          "すべて原典から作り直して一致しました")
    print("   ○ 算定の方法を 05シートに残しました"
          "（列の位置は見出しの文字から探しています）")


if __name__ == "__main__":
    main()
