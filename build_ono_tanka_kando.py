"""小野町 第10期　単価の置き方が保険料にどう効くかの感度分析。

**なぜ要るか。**

本計画のサービス見込量は、単価を令和7年度で固定して組んでいる。
報酬改定率が示されていないためで、置き方としては筋が通っている。
ただし**固定は過小側の置き方である。**単価は改定のない年でも上がっており、
固定したままだと給付費を少なく見込むことになる。

保険料基準額を6,100円とするか6,600円とするか（決定1）をお決めいただく際、
**「単価の置き方を変えると、いくら動くのか」**が分からないと判断しにくい。
その幅を測るのが本資料である。

**単価の趨勢をどう測るか。**

年報 様式2 の給付費を件数で割った額を単価とする。
ただし全サービスをまとめた平均単価は、どのサービスが使われたか
（構成）の変化に引きずられる。令和7年度は平均単価が3.2％下がっているが、
これは価格が下がったのではなく、安いサービスの割合が増えたことによる。

そこで**サービスの構成を基準年で固定した価格指数**（ラスパイレス型）で
価格だけの動きを取り出す。構成の変化による分も別に示す。

出力:
  04_算定・見込量/小野町_第10期_単価の置き方の感度分析_YYYYMMDD.docx
  04_算定・見込量/小野町_第10期_単価の置き方の感度分析_YYYYMMDD.xlsx
"""

import pathlib
import warnings

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

warnings.filterwarnings("ignore")

import build_ono_tanka as T           # noqa: E402
import ono_shizentai as SZ            # noqa: E402
import ono_style as S                 # noqa: E402

ROOT = pathlib.Path(__file__).parent
OUT = ROOT / "小野町_引継ぎ_整理済" / "04_算定・見込量"
ASOF = "20261005"
ASOF_JP = "令和8年10月5日"

BASE_YEAR = "令和3年度"        # 価格指数の基準年
# 基準年度（令和7年度）から各計画年度までの年数
NEN_SA = (2, 3, 4)


def series():
    """年度別の平均単価と、構成を固定した価格指数を返す。"""
    act, _sub, _nen = T.extract()
    years = list(act)
    svcs = sorted({s for y in years for s in act[y]["件数"]})

    def unit(y, svc, i):
        ken = act[y]["件数"].get(svc, (0, 0))[i]
        kyu = act[y]["給付費"].get(svc, (0, 0))[i]
        return (kyu / ken) if ken else None

    out = []
    for y in years:
        ken = sum(sum(v) for v in act[y]["件数"].values())
        kyu = sum(sum(v) for v in act[y]["給付費"].values())
        # ラスパイレス型の価格指数。基準年の件数構成で固定する。
        num = den = 0.0
        for s in svcs:
            for i in (0, 1):
                kb = act[BASE_YEAR]["件数"].get(s, (0, 0))[i]
                pb, pt = unit(BASE_YEAR, s, i), unit(y, s, i)
                if kb and pb and pt:
                    num += pt * kb
                    den += pb * kb
        out.append({"年度": y, "件数": ken, "給付費": kyu,
                    "平均単価": kyu / ken, "指数": num / den * 100})
    for i, d in enumerate(out):
        d["平均単価の前年比"] = (None if i == 0 else
                           d["平均単価"] / out[i - 1]["平均単価"] - 1)
        d["指数の前年比"] = (None if i == 0 else
                        d["指数"] / out[i - 1]["指数"] - 1)
        # 平均単価の動きのうち、価格で説明できない分＝構成の変化
        d["構成の寄与"] = (None if i == 0 else
                      d["平均単価の前年比"] - d["指数の前年比"])
    return out


def nenritsu(ser):
    """価格指数の年率。基準年から直近年までの幾何平均。"""
    n = len(ser) - 1
    return (ser[-1]["指数"] / ser[0]["指数"]) ** (1 / n) - 1


def kaitei_nashi(ser):
    """報酬改定のなかった年度だけの、価格指数の前年比の平均。

    令和6年度は介護報酬改定（＋1.59％）があった年度である。
    改定のない年にも単価が上がることを示すために分けて見る。
    """
    nashi = [d["指数の前年比"] for d in ser[1:] if d["年度"] != "令和6年度"]
    ari = [d["指数の前年比"] for d in ser[1:] if d["年度"] == "令和6年度"]
    return sum(nashi) / len(nashi), (ari[0] if ari else None)


def cases(ser):
    """単価の置き方ごとの標準給付費3年計と保険料基準額。"""
    plan = SZ.plan_rows()
    std = plan["集計"]["標準給付費"]
    chiiki3 = T.CHIIKI_JISSEKI["令和6年度"]["計"] * 3 / 1000
    sogo3 = T.CHIIKI_JISSEKI["令和6年度"]["総合事業"] * 3 / 1000
    pop3 = sum(T.POP_DAI10[y]["1号"]
               for y in ("令和9年度", "令和10年度", "令和11年度"))
    r = nenritsu(ser)
    nashi, _ari = kaitei_nashi(ser)

    def calc(rate, ippatsu=False):
        """rate で伸ばした標準給付費3年計と月額。

        ippatsu=True は令和9年度に一度だけ改定し、以後据え置く置き方。
        """
        if ippatsu:
            s3 = sum(v * (1 + rate) for v in std)
        else:
            s3 = sum(v * (1 + rate) ** n for v, n in zip(std, NEN_SA))
        return s3, T.premium(s3, chiiki3, sogo3, pop3)["月額"]

    out = []
    s3, m = calc(0.0)
    out.append(("**本計画（単価を令和7年度で固定）**", "―", s3, m,
                "**報酬改定率が示されていないため、動かさない置き方**"))
    s3, m = calc(r)
    out.append(("価格指数の趨勢で伸ばす", f"年率{r * 100:+.2f}％", s3, m,
                f"令和3→令和7年度の幾何平均。**構成の変化を除いた価格だけの動き**"))
    s3, m = calc(nashi)
    out.append(("報酬改定のない年の趨勢で伸ばす", f"年率{nashi * 100:+.2f}％", s3, m,
                "令和6年度（改定のあった年度）を除いた平均"))
    for k in (0.01, 0.02, 0.03):
        s3, m = calc(k, ippatsu=True)
        out.append((f"令和9年度に報酬改定 ＋{k * 100:.0f}％（以後据え置き）",
                    f"＋{k * 100:.0f}％", s3, m,
                    "改定率が示された場合の目安"))
    return out, r, nashi


# ---------------------------------------------------------------- 出力
def build_docx(ser, cs, r, nashi):
    rep = S.Report("小野町 第10期　単価の置き方の感度分析")
    rep.cover(["小野町高齢者保健福祉計画・第10期介護保険事業計画",
               "単価の置き方の感度分析"],
              ["", ASOF_JP, "受託者"])
    rep.toc_here()
    rep.body_here()

    base = cs[0][3]
    rep.h1("1　要点")
    rep.p("**本計画は、サービスの単価を令和7年度で固定して見込量を組んでいます。**"
          "報酬改定率が示されていないためで、置き方としては筋が通っています。"
          "**ただし固定は過小側の置き方です。**"
          "単価は報酬改定のない年にも上がっており、"
          "固定したままだと給付費を少なく見込むことになります。")
    rep.p(f"**構成の変化を除いた価格は、令和3年度から令和7年度まで"
          f"年率{r * 100:+.2f}％で上がっています。**"
          f"報酬改定のあった令和6年度を除いても年率{nashi * 100:+.2f}％で、"
          "**改定の有無で大きな差はありません。**")
    rep.p(f"**この趨勢で単価を伸ばすと、保険料基準額は月額{base:,.0f}円から"
          f"{cs[1][3]:,.0f}円へ、{cs[1][3] - base:+,.0f}円動きます。**"
          "保険料を6,100円とするか6,600円とするか（決定1）をお決めいただく際の"
          "幅としてご覧ください。")
    rep.tbl([["単価の置き方", "伸び", "標準給付費（3年計）", "保険料基準額（月額）",
              "本計画との差"]]
            + [[nm, rate, f"{s3:,.0f}千円", f"{m:,.0f}円",
                "―" if abs(m - base) < 0.5 else f"**{m - base:+,.0f}円**"]
               for nm, rate, s3, m, _memo in cs],
            widths=[5.4, 2.0, 3.4, 3.0, 2.4], right=(1, 2, 3, 4))
    rep.src("資料：介護保険事業状況報告（年報）様式2、及び本業務による算定。"
            "**保険料基準額は100円未満を切り上げる前の値です。**")

    rep.h1("2　単価の趨勢をどう測ったか")
    rep.p("**全サービスをまとめた平均単価は、価格の動きを表しません。**"
          "どのサービスが使われたか（構成）の変化に引きずられるためです。"
          "**令和7年度は平均単価が下がっていますが、"
          "これは価格が下がったのではなく、"
          "単価の低いサービスの割合が増えたことによります。**")
    rep.p("そこで、**サービスの構成を令和3年度で固定した価格指数**"
          "（ラスパイレス型）で価格だけの動きを取り出しました。"
          "平均単価の動きのうち、価格指数で説明できない分を"
          "「構成の変化による分」として分けています。")
    rep.tbl([["年度", "件数", "給付費（円）", "平均単価（円）", "平均単価の前年比",
              "価格指数", "価格指数の前年比", "うち構成の変化による分"]]
            + [[d["年度"], f"{d['件数']:,}", f"{d['給付費']:,.0f}",
                f"{d['平均単価']:,.0f}",
                "―" if d["平均単価の前年比"] is None
                else f"{d['平均単価の前年比'] * 100:+.2f}％",
                f"{d['指数']:.2f}",
                "―" if d["指数の前年比"] is None
                else f"**{d['指数の前年比'] * 100:+.2f}％**",
                "―" if d["構成の寄与"] is None
                else f"{d['構成の寄与'] * 100:+.2f}％"]
               for d in ser],
            widths=[1.8, 1.8, 2.6, 2.1, 2.4, 1.7, 2.4, 2.4],
            right=(1, 2, 3, 4, 5, 6, 7))
    rep.src(f"資料：介護保険事業状況報告（年報）様式2。"
            f"価格指数は{BASE_YEAR}＝100。"
            "**単価＝給付費÷件数。**"
            "基準年に実績のないサービスは指数に算入していません。")
    rep.p("**読み取れることが2つあります。**")
    rep.bul(f"**価格は改定の有無によらず上がっています。**"
            f"報酬改定のあった令和6年度は"
            f"{[d for d in ser if d['年度'] == '令和6年度'][0]['指数の前年比'] * 100:+.2f}％、"
            f"改定のなかった年の平均は{nashi * 100:+.2f}％で、ほとんど差がありません。"
            "**改定率が0％でも、単価は上がるとみておく必要があります。**")
    rep.bul("**令和7年度の平均単価の下落は、構成の変化によるものです。**"
            "価格指数は上がっており、価格が下がったわけではありません。"
            "**平均単価をそのまま将来に延ばすと、"
            "構成の変化を価格の変化と取り違えることになります。**")

    rep.h1("3　この幅をどう扱うか")
    rep.p("**本計画は引き続き、単価を令和7年度で固定した値を本体とします。**"
          "報酬改定率が示されていない段階で、"
          "当方の見立てで単価を動かすのは適当でないためです。")
    rep.p("**そのうえで、固定が過小側であることを計画本文に注記します。**"
          f"趨勢で置いた場合に月額{cs[1][3] - base:+,.0f}円動くことを"
          "示したうえで、改定率が示された時点で差し替えます。")
    rep.tbl([["場面", "どう扱うか"],
             ["**協議会（決定1）**",
              "**保険料基準額の2案（6,100円・6,600円）に加え、"
              "単価の置き方で動く幅をあわせてお示しする**"],
             ["計画本文（第6章5）",
              "単価を固定したこと、それが過小側であること、"
              "趨勢で置いた場合の差を注記する"],
             ["**報酬改定率が示されたとき**",
              "**本資料の算式に改定率を入れて差し替える。"
              "令和8年12月から令和9年1月に示されるのが通例であり、"
              "11月の協議会には間に合わない**"],
             ["成果品②", "算定の前提として本資料の内容を収める"]],
            widths=[3.6, 13.6])

    rep.build_toc()
    p = OUT / f"小野町_第10期_単価の置き方の感度分析_{ASOF}.docx"
    rep.save(p)
    return p


HEAD = PatternFill("solid", fgColor="1F3864")


def build_xlsx(ser, cs, r, nashi):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "01_単価の置き方と保険料"
    ws.append(["単価の置き方", "伸び", "標準給付費（3年計・千円）",
               "保険料基準額（月額・円）", "本計画との差（円）", "備考"])
    base = cs[0][3]
    for nm, rate, s3, m, memo in cs:
        ws.append([nm.replace("**", ""), rate, round(s3), round(m),
                   round(m - base), memo.replace("**", "")])
    ws2 = wb.create_sheet("02_単価の趨勢")
    ws2.append(["年度", "件数", "給付費（円）", "平均単価（円）",
                "平均単価の前年比", "価格指数", "価格指数の前年比",
                "うち構成の変化による分"])
    for d in ser:
        ws2.append([d["年度"], d["件数"], round(d["給付費"]),
                    round(d["平均単価"], 1),
                    None if d["平均単価の前年比"] is None
                    else round(d["平均単価の前年比"], 5),
                    round(d["指数"], 2),
                    None if d["指数の前年比"] is None
                    else round(d["指数の前年比"], 5),
                    None if d["構成の寄与"] is None
                    else round(d["構成の寄与"], 5)])
    ws3 = wb.create_sheet("03_前提")
    for row in [
        ["小野町 第10期　単価の置き方の感度分析"],
        [f"基準日：{ASOF_JP}"],
        [],
        ["単価＝給付費÷件数（介護保険事業状況報告（年報）様式2）"],
        [f"価格指数＝サービスの構成を{BASE_YEAR}で固定したラスパイレス型。"
         f"{BASE_YEAR}＝100"],
        [f"価格指数の年率（令和3→令和7年度の幾何平均）　{r * 100:+.3f}％"],
        [f"報酬改定のない年の平均（令和6年度を除く）　{nashi * 100:+.3f}％"],
        [],
        ["標準給付費は、基準年度（令和7年度）から各計画年度までの年数"
         f"{NEN_SA} 乗で伸ばしている"],
        ["保険料基準額は100円未満を切り上げる前の値"],
        ["地域支援事業費は年報 様式4 の令和6年度決算を3年度とも据え置いている"],
    ]:
        ws3.append(row)
    for w in (wb["01_単価の置き方と保険料"], ws2):
        for c in w[1]:
            c.font = Font(bold=True, size=9, color="FFFFFF")
            c.fill = HEAD
            c.alignment = Alignment(wrap_text=True, horizontal="center",
                                    vertical="center")
    for w, widths in ((wb["01_単価の置き方と保険料"], (38, 12, 22, 20, 16, 58)),
                      (ws2, (12, 10, 16, 14, 16, 10, 16, 20)),
                      (ws3, (96,))):
        for i, x in enumerate(widths, 1):
            w.column_dimensions[chr(64 + i)].width = x
    for r_ in ws2.iter_rows(min_row=2):
        for i in (4, 6, 7):
            r_[i].number_format = "0.00%"
    p = OUT / f"小野町_第10期_単価の置き方の感度分析_{ASOF}.xlsx"
    wb.save(p)
    return p


def selfcheck(ser, cs, r, nashi):
    bad = []
    # 1 価格指数の基準年が100か
    if abs(ser[0]["指数"] - 100) > 1e-6:
        bad.append(f"基準年の価格指数が100でない {ser[0]['指数']}")
    # 2 平均単価の前年比＝価格指数の前年比＋構成の寄与
    for d in ser[1:]:
        if abs(d["平均単価の前年比"] - (d["指数の前年比"] + d["構成の寄与"])) > 1e-9:
            bad.append(f"{d['年度']} の分解が合わない")
    # 3 本計画（伸び0）の月額が、素案の算定値と一致するか
    plan = SZ.plan_rows()
    std3 = sum(plan["集計"]["標準給付費"])
    chiiki3 = T.CHIIKI_JISSEKI["令和6年度"]["計"] * 3 / 1000
    sogo3 = T.CHIIKI_JISSEKI["令和6年度"]["総合事業"] * 3 / 1000
    pop3 = sum(T.POP_DAI10[y]["1号"]
               for y in ("令和9年度", "令和10年度", "令和11年度"))
    m0 = T.premium(std3, chiiki3, sogo3, pop3)["月額"]
    if abs(cs[0][3] - m0) > 0.5:
        bad.append(f"本計画の月額が算定値と違う {cs[0][3]:.1f} 対 {m0:.1f}")
    # 4 伸びを大きくするほど保険料が上がるか
    ms = [cs[0][3], cs[2][3], cs[1][3]]      # 0％ → 改定なしの年 → 全体の趨勢
    if not (ms[0] <= ms[1] or ms[0] <= ms[2]):
        bad.append("伸びを与えても保険料が上がらない")
    # 5 年率が常識の範囲か（年±10％を超えたら計算違いを疑う）
    for v, nm in ((r, "価格指数の年率"), (nashi, "改定なしの年の平均")):
        if abs(v) > 0.10:
            bad.append(f"{nm}が {v * 100:.1f}％ で大きすぎる")
    if bad:
        for b in bad:
            print(f"  不適合: {b}")
        raise SystemExit(1)
    print("  自己点検 5項目 適合")


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    ser = series()
    cs, r, nashi = cases(ser)
    selfcheck(ser, cs, r, nashi)
    p1 = build_docx(ser, cs, r, nashi)
    p2 = build_xlsx(ser, cs, r, nashi)
    print(f"出力: {p1}")
    print(f"     {p2}")
    print(f"  価格指数の年率 {r * 100:+.3f}％"
          f"（改定のない年の平均 {nashi * 100:+.3f}％）")
    base = cs[0][3]
    for nm, rate, _s3, m, _memo in cs:
        print(f"  {nm.replace('**', ''):<34}{rate:>10}　"
              f"月額 {m:>8,.0f}円　本計画との差 {m - base:+,.0f}円")


if __name__ == "__main__":
    main()
