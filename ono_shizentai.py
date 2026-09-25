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
R8_CASES = ["据え置き", "令和6→令和7を延長", "現行方式", "月報の実績"]

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
    """令和8年度の置き方4通りを、令和7年度に対する比で返す。

    返り値 {名称: (比, 令和8年度の総給付費（千円）, 根拠)}
    """
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
    return {
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
    }, r7


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

    標準給付費は、現行の算定の総給付費に対する比をそのまま当てる
    （特定入所者・高額・高額医療合算・審査支払手数料は総給付費に比例する）。
    """
    from build_ono_kaisu import chiiki_plan, plan_rows
    A = plan_rows()["集計"]
    cp = chiiki_plan()
    if chi3 is None:
        chi3 = sum(v[3] for v in cp) * 3 / 1000
    if sogo3 is None:
        sogo3 = sum(v[3] for v in cp if v[0] == "総合事業") * 3 / 1000
    pop3 = sum(T.POP_DAI10[y]["1号"] for y in T.PLAN_YEARS)
    p = plan(base)
    new3 = sum(p[y] for y in T.PLAN_YEARS)
    scale = new3 / sum(A["計"])
    std3 = sum(A["標準給付費"]) * scale
    res = T.premium(std3, chi3, sogo3, pop3, **kw)
    res["総給付費3年計"] = new3
    res["倍率"] = scale
    return res


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
