"""小野町 年報の要介護度別の系列を取り出す。

川崎町の第10期で用いている算定方式を小野町のデータで再現するために作った。
同町の方式は次のとおりである（project_mikomi_R8.9.15.py の冒頭による）。

  基準  令和7年度（完結年度）の実績。年報の要介護度別の受給者数・
        利用回（日）数・給付費をそのまま用いる
  伸び  要介護度別の認定者数の伸びのみ
          伸び率(年, d) ＝ 認定者数(年, d) ÷ 認定者数(基準年度, d)
  据置  サービス利用率・1人1月あたり利用回（日）数・1人1月あたり給付費は
        基準年度の水準で据え置く

当方（小野町）の方式は第1号被保険者1人当たりの利用率を延ばすもので、
要介護度の構成が変わることを織り込まない。川崎町の方式は逆に、
要介護度の構成の変化だけを織り込み、1人当たりの水準は据え置く。
どちらが小野町の実績をよく再現するかは build_ono_backtest.py で測る。

読み取る様式
  様式2（給付費）   要介護度別の給付費（列6〜16）。サービス別
  様式1の5 総数     要介護度別の認定者数。第1号と総数の2系列
"""

import pathlib
import tempfile

import openpyxl  # noqa: F401  （load_nenpo が用いる）

from build_ono_nenpo import SRC as NENPO_SRC
from build_ono_nenpo import YEARS as NENPO_YEARS
from build_ono_nenpo import load_nenpo
from build_ono_tanka import CANON, _label

# 要介護度の区分。様式2（給付費）の列番号と対応させる。
# 列9の「経過的要介護」は制度上すでに発生しないため区分に含めない
# （値が入っていれば KEIKATEKI に拾い、合計の点検で用いる）。
KAIGODO = ["要支援1", "要支援2", "要介護1", "要介護2", "要介護3", "要介護4", "要介護5"]
FORM2_COL = {"要支援1": 6, "要支援2": 7, "要介護1": 10, "要介護2": 11,
             "要介護3": 12, "要介護4": 13, "要介護5": 14}
FORM2_KEIKA = 9          # 経過的要介護
FORM2_TOTAL = 16         # 合計

# 様式1の5「計」ブロックの行と列。
#   行31 第1号被保険者、行38 第2号被保険者、行39 総数
NINTEI_ROW = {"1号": 31, "2号": 38, "総数": 39}
NINTEI_COL = {"要支援1": 5, "要支援2": 6, "要介護1": 9, "要介護2": 10,
              "要介護3": 11, "要介護4": 12, "要介護5": 13}
NINTEI_TOTAL = 15

YS = [lab for _c, lab in NENPO_YEARS]


def _sheet_nintei(wb):
    """様式1の5 総数のシート。年度により表記が揺れるため名前で探す。"""
    for sn in wb.sheetnames:
        if "１の５" in sn and "総数" in sn:
            return wb[sn]
    raise KeyError("様式１の５ 総数 が見つからない")


def read():
    """要介護度別の給付費と認定者数を年度ごとに返す。

    戻り値
      kyufu[年度][正規化サービス名][要介護度] = 給付費（円）
      nintei[年度][区分][要介護度]            = 認定者数（人）。区分は 1号／2号／総数
      chk[年度] = {"様式2合計": …, "度別の和": …, "経過的要介護": …}
    """
    work = pathlib.Path(tempfile.mkdtemp())
    kyufu, nintei, chk = {}, {}, {}
    for y, lab in NENPO_YEARS:
        wb = load_nenpo(NENPO_SRC / "原本_年報" / f"年報データ_{y}_小野町.xlsx", work)

        ws = wb["様式２（給付費）"]
        agg, tot2, sum2, keika = {}, 0.0, 0.0, 0.0
        for r in range(11, ws.max_row + 1):
            name = _label(ws, r)
            canon = CANON.get(name) if name else None
            if canon is None:
                continue
            row = {}
            for d, c in FORM2_COL.items():
                row[d] = float(ws.cell(r, c).value or 0)
            cur = agg.setdefault(canon, {d: 0.0 for d in KAIGODO})
            for d in KAIGODO:
                cur[d] += row[d]
            tot2 += float(ws.cell(r, FORM2_TOTAL).value or 0)
            sum2 += sum(row.values())
            keika += float(ws.cell(r, FORM2_KEIKA).value or 0)
        kyufu[lab] = agg

        wn = _sheet_nintei(wb)
        n = {}
        for kind, r in NINTEI_ROW.items():
            n[kind] = {d: float(wn.cell(r, c).value or 0)
                       for d, c in NINTEI_COL.items()}
        nintei[lab] = n

        chk[lab] = {"様式2合計": tot2, "度別の和": sum2, "経過的要介護": keika}
    return kyufu, nintei, chk


def growth(nintei, t0, tgt, kind="総数"):
    """要介護度別の伸び率 認定者数(tgt, d) ÷ 認定者数(t0, d)。

    基準年度に0人の度は伸ばしようがないため1.0とする。
    """
    a, b = nintei[t0][kind], nintei[tgt][kind]
    return {d: (b[d] / a[d] if a[d] else 1.0) for d in KAIGODO}


def project(kyufu, nintei, t0, tgt, kind="総数"):
    """川崎町方式の予測。{サービス名: 給付費（円）} を返す。

    サービスごとに、基準年度の要介護度別給付費にその度の認定者数の伸びを
    乗じて足し上げる。1人当たりの水準は基準年度で据え置かれる。
    """
    g = growth(nintei, t0, tgt, kind)
    return {svc: sum(v[d] * g[d] for d in KAIGODO)
            for svc, v in kyufu[t0].items()}


def selfcheck():
    """様式2の合計と要介護度別の和が一致することを確かめる。

    差は経過的要介護の列に限られるはずである。1円でも合わなければ
    AssertionError を投げる（大雪の手引き 資料2 の型）。
    """
    _k, _n, chk = read()
    for lab, c in chk.items():
        gap = c["様式2合計"] - c["度別の和"] - c["経過的要介護"]
        assert abs(gap) < 0.5, (
            f"{lab}: 様式2の合計{c['様式2合計']:,.0f}と"
            f"要介護度別の和{c['度別の和']:,.0f}＋経過的要介護"
            f"{c['経過的要介護']:,.0f}が{gap:,.0f}円合わない")
    return chk


if __name__ == "__main__":
    kyufu, nintei, chk = read()
    print("年度別の点検（様式2の合計＝要介護度別の和＋経過的要介護）")
    for lab in YS:
        c = chk[lab]
        print(f"  {lab}  合計{c['様式2合計']:>14,.0f}  度別の和{c['度別の和']:>14,.0f}"
              f"  経過的{c['経過的要介護']:>8,.0f}"
              f"  差{c['様式2合計'] - c['度別の和'] - c['経過的要介護']:>6,.0f}")
    selfcheck()
    print("\n要介護度別の認定者数（総数・年度末）")
    print("  年度      " + "".join(f"{d:>9s}" for d in KAIGODO) + "      計")
    for lab in YS:
        n = nintei[lab]["総数"]
        print(f"  {lab:<9s}" + "".join(f"{n[d]:>9,.0f}" for d in KAIGODO)
              + f"{sum(n.values()):>9,.0f}")
