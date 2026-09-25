"""小野町 第10期 令和8年度を自然体推計で置き、それを基点に第10期を見込む。

**なぜ必要か**

見える化システムのワークシート（将来推計1）は、1人1月あたり給付費の実績値を
令和8年度に置いている。ところが令和8年度の列は1か月分であり、
その1か月の水準が第10期の3年間に効いてしまう。
令和8年度を自然体推計に置き換える必要がある。

**令和8年度をどう置くか**

置き方は3つある。いずれも令和7年度の実績（年報 様式2・1,005,144千円）に対する比で示す。

  ① 据え置き（変化0）                     1.0000
  ② 令和6年度→令和7年度の動きを1年延長      0.9255
  ③ 当方の現行方式（令和5〜7年度の利用率平均） 1.0171

このうち②は、令和6年度から令和7年度にかけての総給付費の落ち（▲7.45％）が
そのまま続くとみるものである。令和7年9月の認定者数の断層を含んだ落ちであり、
そのまま延ばすと過小になる。

**国保連月報で確かめられる。**
令和8年度の提供分は審査202605〜202607の3か月あり、前年の同じ提供月と比べられる。

  ④ 月報3か月の前年同期比（実績）           0.9598

実績は①と②の間にあり、**当方の現行方式（③）は実績より6.0％高い。**
令和5〜7年度の3年平均が、下がり続けている系列を上に引いているためである。

したがって令和8年度は④で置く。これが「令和6年度→令和7年度の動きを勘案し、
その動きが令和8年度も続いているかを月報で確かめたうえでの自然体推計」にあたる。

**第10期の置き方**

令和8年度を基点とし、第1号被保険者1人当たりの給付費を据え置いて人口で延ばす。
見える化システムの自然体推計（利用率・1人1月あたり給付費の変化を0とする）と
同じ構造である。

  見込量(y) ＝ 令和8年度 × 第1号被保険者数(y) ÷ 第1号被保険者数(令和8年度)

**留意点**

  月報は審査月の集計であり、過誤調整や返戻の再請求は後の審査月に現れる。
  したがって1つの提供月の額が翌月の審査分にすべて収まるとは限らない。
  ただしこの影響は令和7年度側にも同じように及ぶため、比としては大きく歪まない。
  過誤調整の状況は町への確認事項（依頼票B11）に挙げている。

  令和8年度は3か月分しかなく、完結年度ではない。
  他町村の手引きは「基準年度は完結した最新年度とする」としており、
  この意味では現行（令和7年度を基準年度とする）のほうが原則に沿う。
  どちらを採るかは協議会の判断による。

出力はない。build_ono_kaisu.py・build_ono_mikomi_kakutei.py・
build_ono_soan_elderly.py から参照する。
"""

import build_ono_tanka as T

# 令和8年度の提供分（審査202605〜202607）と、前年の同じ提供月（審査202505〜202507）
R8_SHINSA = ["202605", "202606", "202607"]
R7_SHINSA = ["202505", "202506", "202507"]

# 令和8年度の置き方の候補。値は令和7年度の実績に対する比。
#   ① 据え置き ② 令和6→令和7の動きを1年延長 ③ 当方の現行方式 ④ 月報の実績
#   ⑤ 認定者数の伸び＋サービス別趨勢の半分（月報を使わない独立の置き方）
# ④を標準とする。⑤は④の検算にあたり、出所が重ならない。
R8_CASES = ["据え置き", "令和6→令和7を延長", "現行方式", "月報の実績",
            "認定者数＋趨勢半分"]

_CACHE = {}


def geppo_ratio():
    """令和8年度の前年同期比を国保連月報から求める。

    総合事業（Ａ系）と小計・合計の行を除いた給付額の合計で、
    提供月をそろえて比べる。返り値は (比, 令和7年度の額, 令和8年度の額)。
    """
    if "geppo" in _CACHE:
        return _CACHE["geppo"]
    from build_ono_kaisu import read_geppo
    d = read_geppo()

    def tot(months):
        t = 0.0
        for m in months:
            for code, v in (d.get((m, "給付額"), {}) or {}).items():
                if code.startswith("A") or "計" in code:
                    continue
                t += (v or {}).get("計", 0) or 0
        return t

    a, b = tot(R7_SHINSA), tot(R8_SHINSA)
    _CACHE["geppo"] = (b / a if a else 1.0, a, b)
    return _CACHE["geppo"]


def r8_options(act=None, gen=None, proj=None):
    """令和8年度の置き方5通りを、令和7年度に対する比で返す。

    返り値 {名称: (比, 令和8年度の総給付費（千円）, 根拠)}

    年報の読み取りと推計が重いため、引数なしで呼ばれたときは結果を残す。
    素案・見込量・点検のいずれも同じ値を使うので、1回で足りる。
    """
    if act is None and gen is None and proj is None and "r8" in _CACHE:
        return _CACHE["r8"]
    cached = act is None and gen is None and proj is None
    if act is None:
        act, _sub, gen = T.extract()
    r7 = sum(sum(act["令和7年度"]["給付費"].get(s, (0, 0)))
             for _c, s, _u in T.ORDER_KAIGO) / 1000
    r6 = sum(sum(act["令和6年度"]["給付費"].get(s, (0, 0)))
             for _c, s, _u in T.ORDER_KAIGO) / 1000
    if proj is None:
        proj = T.project(act, gen, "1号", "1号", 0.0)
    i8 = T.EST_YEARS.index("令和8年度")
    cur = sum(r["見込"][i8][1] for k in ("介護", "予防") for r in proj[k])
    g, ga, gb = geppo_ratio()
    out = ({
        "据え置き": (1.0, r7,
                 "令和7年度の実績をそのまま置く（変化0）"),
        "令和6→令和7を延長": (r7 / r6, r7 * r7 / r6,
                        f"令和6年度{r6:,.0f}千円→令和7年度{r7:,.0f}千円の比を"
                        f"もう1年延ばす。**断層を含んだ落ちであり過小になる**"),
        "現行方式": (cur / r7, cur,
                 "令和5〜7年度の利用率の平均×令和8年度の推計人口。"
                 "**3年平均が下がり続けている系列を上に引いている**"),
        "月報の実績": (g, r7 * g,
                  f"国保連月報の令和8年度4〜6月提供分{gb:,.0f}円と"
                  f"前年の同じ月{ga:,.0f}円の比。**3か月分の実績による**"),
        "認定者数＋趨勢半分": blend(act, r7, r6),
    }, r7)
    if cached:
        _CACHE["r8"] = out
    return out


# 他町村の案件で用いている置き方。サービス別の趨勢をそのまま延ばすと外れやすいため
# 半分だけ織り込み、一律の伸びからの乖離に上限を置く。
BLEND_HANEI = 0.50      # 趨勢の反映割合
BLEND_JOGEN = 0.20      # 一律の伸びからの乖離の上限（ポイント）
NINTEI_R7_1GO = 783     # 年報 様式1の5 第1号・令和8年3月末


def blend(act, r7, r6, hanei=BLEND_HANEI, jogen=BLEND_JOGEN):
    """一律の伸び（認定者数）にサービス別の趨勢を半分だけ織り込む置き方。

    一律の伸び ＝ 第1号認定者数 令和8年度÷令和7年度
    採用伸び ＝ MEDIAN(一律−上限, 一律＋反映割合×(サービス別趨勢−一律), 一律＋上限)
    """
    ichiritsu = T.NINTEI_EST["令和8年度"] / NINTEI_R7_1GO
    tot = 0.0
    for _c, s, _u in T.ORDER_KAIGO:
        v7 = sum(act["令和7年度"]["給付費"].get(s, (0, 0))) / 1000
        v6 = sum(act["令和6年度"]["給付費"].get(s, (0, 0))) / 1000
        if v7 <= 0:
            continue
        susei = v7 / v6 if v6 > 0 else ichiritsu
        cand = ichiritsu + hanei * (susei - ichiritsu)
        tot += v7 * min(max(cand, ichiritsu - jogen), ichiritsu + jogen)
    return (tot / r7, tot,
            f"第1号認定者数の伸び{(ichiritsu - 1) * 100:+.2f}％に、"
            f"サービス別の趨勢を{hanei * 100:.0f}％だけ織り込む"
            f"（乖離の上限{jogen * 100:.0f}ポイント）。"
            "**他町村の案件で用いている置き方。月報を使わずに置ける**")


def plan(base="月報の実績"):
    """令和8年度を基点とする第10期の総給付費（千円）。

    返り値 {"令和8年度": v, "令和9年度": v, …, "比": 令和7年度に対する比}
    """
    opts, r7 = r8_options()
    ratio = opts[base][0]
    r8 = r7 * ratio
    pop = T.POP_DAI10
    out = {"令和8年度": r8, "比": ratio, "令和7年度": r7}
    for y in T.PLAN_YEARS:
        out[y] = r8 * pop[y]["1号"] / pop["令和8年度"]["1号"]
    return out


def premium(base="月報の実績", chi3=None, sogo3=None, **kw):
    """自然体推計（令和8年度基点）による保険料。

    標準給付費は plan_rows の系列をそのまま合計する。
    **素案・見込量の表に載る値と、保険料の基数を同じものにするため。**
    （従前は現行の算定の標準給付費に総給付費の比を当てていたが、
    審査支払手数料は給付費でなく件数に比例するため21千円ずれていた。）
    """
    from build_ono_kaisu import chiiki_plan, plan_rows as base_rows
    A = plan_rows(base)["集計"]
    cp = chiiki_plan()
    if chi3 is None:
        chi3 = sum(v[3] for v in cp) * 3 / 1000
    if sogo3 is None:
        sogo3 = sum(v[3] for v in cp if v[0] == "総合事業") * 3 / 1000
    pop3 = sum(T.POP_DAI10[y]["1号"] for y in T.PLAN_YEARS)
    new3 = sum(A["計"])
    std3 = sum(A["標準給付費"])
    res = T.premium(std3, chi3, sogo3, pop3, **kw)
    res["総給付費3年計"] = new3
    res["倍率"] = new3 / sum(base_rows()["集計"]["計"])
    return res


def factors(base="月報の実績"):
    """第10期の各年度に乗じる係数（令和7年度の実績に対する比）。

    令和8年度 ＝ 令和7年度 × 比、第10期 ＝ 令和8年度 × 人口比、なので
        係数(y) ＝ 比 × 第1号被保険者数(y) ÷ 第1号被保険者数(令和8年度)
    """
    opts, _r7 = r8_options()
    ratio = opts[base][0]
    p8 = T.POP_DAI10["令和8年度"]["1号"]
    return [ratio * T.POP_DAI10[y]["1号"] / p8 for y in T.PLAN_YEARS]


def plan_rows(base="月報の実績"):
    """自然体推計によるサービス別の見込量。build_ono_kaisu.plan_rows と同じ形。

    令和7年度の実績（各行の先頭）に係数を乗じたもので置き換える。
    **人数・量・給付費のすべてに同じ係数を当てる。**
    総額を月報で裏づけるのが目的であり、サービスの構成と1人当たりの水準は
    令和7年度のまま動かさない（見える化システムの自然体推計と同じ置き方）。

    なお月報でみると、給付額の前年同期比0.9598に対し受給者計は0.9648で、
    0.5ポイントの差がある。すべてに給付額の比を当てているため、
    1人1月あたり給付費は令和7年度から0.5％下がる形になる。
    """
    if base in _CACHE.get("rows", {}):
        return _CACHE["rows"][base]
    from build_ono_kaisu import plan_rows as base_rows
    P = base_rows()
    f = factors(base)
    act, _sub, gen = T.extract()
    out = {}
    for kind in ("介護", "予防"):
        i = 1 if kind == "介護" else 0
        rows = []
        for cat, name, unit, nin, ryo, _kyu in P[kind]:
            from build_ono_kaisu import _svc_of
            svc = _svc_of(name, kind)
            k7 = act["令和7年度"]["給付費"].get(svc, (0, 0))[i] / 1000
            rows.append((cat, name, unit,
                         [nin[0]] + [nin[0] * x for x in f],
                         ([ryo[0]] + [ryo[0] * x for x in f]) if ryo else None,
                         [k7 * x for x in f]))
        out[kind] = rows
    # 集計。加算4項目の比は現行と同じものを使う
    n = len(T.PLAN_YEARS)
    agg = {"区分": {}, "給付": {}, "計": [0.0] * n}
    for kind in ("介護", "予防"):
        v = [0.0] * n
        for cat, _nm, _u, _nin, _ryo, kyu in out[kind]:
            for k in range(n):
                v[k] += kyu[k]
                agg["区分"].setdefault(cat, [0.0] * n)[k] += kyu[k]
        agg["給付"][kind] = v
        for k in range(n):
            agg["計"][k] += v[k]
    ratios = T.kasan_ratios(gen)
    for key in ("特定入所", "高額", "高額合算"):
        agg[key] = [agg["計"][k] * ratios[key] for k in range(n)]
    # 審査支払手数料は件数に比例する。件数も同じ係数で動かす
    ken7 = sum(act["令和7年度"]["件数"].get(s, (0, 0))[0]
               + act["令和7年度"]["件数"].get(s, (0, 0))[1]
               for _c, s, _u in T.ORDER_KAIGO)
    from build_ono_kaisu import TESURYO_YEN
    agg["手数料"] = [ken7 * f[k] * TESURYO_YEN / 1000 for k in range(n)]
    agg["標準給付費"] = [agg["計"][k] + agg["特定入所"][k] + agg["高額"][k]
                    + agg["高額合算"][k] + agg["手数料"][k] for k in range(n)]
    out["集計"] = agg
    _CACHE.setdefault("rows", {})[base] = out
    return out


def summary():
    """比較表のもと。現行と自然体推計（令和8年度基点）を並べる。"""
    from build_ono_kaisu import chiiki_plan, plan_rows
    A = plan_rows()["集計"]
    cp = chiiki_plan()
    chi3 = sum(v[3] for v in cp) * 3 / 1000
    sogo3 = sum(v[3] for v in cp if v[0] == "総合事業") * 3 / 1000
    pop3 = sum(T.POP_DAI10[y]["1号"] for y in T.PLAN_YEARS)
    cur = T.premium(sum(A["標準給付費"]), chi3, sogo3, pop3)
    cur["総給付費3年計"] = sum(A["計"])
    cur["倍率"] = 1.0
    return {"現行": cur, "自然体": premium()}


if __name__ == "__main__":
    opts, r7 = r8_options()
    print(f"令和7年度の総給付費（年報 様式2） {r7:,.0f}千円\n")
    print(f"{'令和8年度の置き方':<22}{'対令和7年度':>12}{'令和8年度':>14}")
    for k in R8_CASES:
        v, amt, _memo = opts[k]
        print(f"  {k:<20}{v:>12.4f}{amt:>14,.0f}千円")
    print()
    s = summary()
    print(f"{'':<26}{'現行':>16}{'自然体（令和8年度基点）':>26}")
    for k in ("総給付費3年計", "標準給付費", "月額"):
        a, b = s["現行"][k], s["自然体"][k]
        f = ",.0f" if k != "月額" else ",.0f"
        print(f"  {k:<24}{a:>16{f}}{b:>26{f}}   差 {b - a:+,.0f}")
    p = plan()
    print()
    for y in ["令和8年度"] + T.PLAN_YEARS:
        print(f"  {y:<10}{p[y]:>12,.0f}千円")
