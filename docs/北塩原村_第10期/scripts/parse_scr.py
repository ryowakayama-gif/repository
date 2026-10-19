# -*- coding: utf-8 -*-
"""介護ＳＣＲ（標準化レセプト出現比）を読み取る。

   受領（令和8年10月19日）
     data/原典_SCR/介護SCRの概説と利用上の留意点_令和8年6月.pdf
     data/原典_SCR/SCR_住所地ベース_北塩原村_2021-2023.csv
     data/原典_SCR/xSCR_事業所所在地ベース_北塩原村_2021-2023.csv

   出どころ
     厚生労働科学研究費補助金〈長寿科学政策研究〉
     「2040年を見据えた地域特性に応じた介護サービス提供体制の最適化と
       科学的介護の推進に関する調査研究」（研究代表者：陣内裕成）
     データソースは匿名介護保険等関連情報データベース（介護保険総合データベース）

   【名称について】　概説（令和8年6月）により、介護ＳＣＲの分子は
   **レセプトコード出現数（レセプト件数）**であって給付費ではない。
   したがって「標準化給付費比」ではなく「標準化レセプト出現比」である。
   当方はこれを受領前に「標準化給付費比」と記していたが、誤りであった。

     ＳＣＲ ＝ Σ性・年齢別レセプトコード出現数
              ÷ Σ（性・年齢別人口 × 全国の性・年齢別レセプトコード出現率）× 100

   2系列ある。
     ＳＣＲ  利用者の住所地でみる。本村の被保険者が**どこで利用したかを問わず**数える
     ｘＳＣＲ 事業所の所在地でみる。村内の事業所が**誰に提供したかを問わず**数える

   分母が同じ（村の人口×全国出現率）であるため、
     ｘＳＣＲ ÷ ＳＣＲ ＝ 村内事業所のレセプト件数 ÷ 本村被保険者のレセプト件数
   となる。これは**自給率とは別のもの**である。自給率（2-5）は給付月額を用い、
   分子は村内事業所が**本村の被保険者に**提供した分であるのに対し、
   ｘＳＣＲの分子は村外の方への提供も含む。村内事業所が村外の方を受け入れていれば
   ｘＳＣＲ÷ＳＣＲは自給率より高く出る。read_jikyu() との突合せは行わない。

   python3 scripts/parse_scr.py        読み取って CSV に落とす
"""
import csv
import io
import os
import sys
sys.dont_write_bytecode = True
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import paths as P

SRC = os.path.join(P.DATA, "原典_SCR")
F_SCR = os.path.join(SRC, "SCR_住所地ベース_北塩原村_2021-2023.csv")
F_XSCR = os.path.join(SRC, "xSCR_事業所所在地ベース_北塩原村_2021-2023.csv")

NEN = ("2021", "2022", "2023")          # 年度。2023年度＝令和5年度

# 概説 V(1)　小規模市町村では単年度の振れ幅が大きいため、3か年の推移で読む。
# 概説 V(2)　事業所所在地は介護ＤＢの事業所台帳による。実際の所在地と異なる
#            市町村で数えられる例が一部で確認されているとされる。
CHUI = (
 "介護ＳＣＲは各年度の実績による単年度集計であり、被保険者数の少ない市町村や"
 "利用者が限られるサービス区分では年度ごとの振れが生じやすい。"
 "単年度では読まず3か年の推移で見る（概説Ⅴ(1)）。")

# 本村で読み取る大区分。概説別添2のサービス区分表による
#   区分1   介護給付と予防給付
#   区分2-1 居宅系／施設系
#   区分2-2 居宅系／施設居住系（居住系を施設と同じ側に置く）
#   区分2-3 提供形態（訪問・通所など）とサービス特性
DAI = {"10": "区分1", "21": "区分2-1", "22": "区分2-2", "23": "区分2-3"}
SHU = "30"                               # 区分3　基本的なサービス種別


def _load(path):
    """（区分番号, サービス区分名）→ (2021, 2022, 2023) の辞書を返す。

       市区町村の列がすべて北塩原村であることを確かめる。
       （抽出の際に他団体の行が混じっていないことの確認。
         他団体の数値は参照値であり、成果品に掲げない。）
    """
    rows = list(csv.reader(io.open(path, encoding="utf-8-sig")))
    hd = [c.replace("\n", "").strip() for c in rows[0]]
    # 列の位置を見出しで引く（版で動くため）
    i_shi = hd.index("市区町村")
    i_ku = [i for i, c in enumerate(hd) if c.startswith("サービス区分") and "番号" in c]
    i_ku = i_ku[0] if i_ku else 8
    i_na = i_ku + 1
    i_v = [i for i, c in enumerate(hd) if c.endswith(NEN)]
    assert len(i_v) == 3, hd
    d = {}
    for r in rows[1:]:
        r = [c.strip() for c in r]
        if not r or not r[i_shi]:
            continue
        assert r[i_shi] == "北塩原村", "本村以外の行がある: %s" % r[i_shi]
        d[(r[i_ku], r[i_na])] = tuple(r[i] for i in i_v)
    return d


def scr():
    """住所地ベース（本村の被保険者の利用）"""
    return _load(F_SCR)


def xscr():
    """事業所所在地ベース（村内の事業所の供給）"""
    return _load(F_XSCR)


def _f(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def table(kubun):
    """大区分または区分3について、[(名, SCR3か年, xSCR3か年, 2023の比)] を返す。

       比は ｘＳＣＲ÷ＳＣＲ（村内事業所の件数÷本村被保険者の件数）。
       ＳＣＲが0のときは None。
    """
    a, x = scr(), xscr()
    out = []
    for (k, na) in sorted(a):
        if k != kubun:
            continue
        va, vx = a[(k, na)], x.get((k, na), ("0", "0", "0"))
        fa, fx = _f(va[2]), _f(vx[2])
        hi = (fx / fa) if (fa and fx is not None) else None
        out.append((na, va, vx, hi))
    return out


def kyokyu_nashi():
    """区分3のうち、本村に利用があり（ＳＣＲ>0）村内に供給がない（ｘＳＣＲ=0）もの"""
    return [t[0] for t in table(SHU)
            if (_f(t[1][2]) or 0) > 0 and (_f(t[2][2]) or 0) == 0]


def kyokyu_ari():
    """区分3のうち、村内に供給があるもの（ｘＳＣＲ>0）"""
    return [t[0] for t in table(SHU) if (_f(t[2][2]) or 0) > 0]


def main():
    rows = [("区分", "サービス区分", "ＳＣＲ2021", "ＳＣＲ2022", "ＳＣＲ2023",
             "ｘＳＣＲ2021", "ｘＳＣＲ2022", "ｘＳＣＲ2023", "ｘＳＣＲ÷ＳＣＲ(2023)")]
    for k in ("10", "21", "22", "23", SHU):
        for na, va, vx, hi in table(k):
            if va == ("0", "0", "0") and vx == ("0", "0", "0"):
                continue
            rows.append((DAI.get(k, "区分3"), na) + va + vx +
                        ("%.3f" % hi if hi is not None else "―",))
    p = os.path.join(P.DATA, "第10期_介護SCR.csv")
    with io.open(p, "w", encoding="utf-8-sig", newline="") as fh:
        csv.writer(fh).writerows(rows)
    print("書いた %s（%d行）" % (p, len(rows) - 1))
    print()
    print("■ 村内に供給がある（ｘＳＣＲ>0）サービス種別 %d" % len(kyokyu_ari()))
    for n in kyokyu_ari():
        print("   ", n)
    print("■ 利用はあるが村内に供給がない（ＳＣＲ>0 かつ ｘＳＣＲ=0）%d"
          % len(kyokyu_nashi()))
    for n in kyokyu_nashi():
        print("   ", n)
    print()
    print(CHUI)


if __name__ == "__main__":
    main()
