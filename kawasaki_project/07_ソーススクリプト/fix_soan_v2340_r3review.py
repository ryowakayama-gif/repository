# -*- coding: utf-8 -*-
"""協議用素案 Ver.2.33 → Ver.2.34（Red Team 再レビュー第3回（令和8年10月24日）の指摘の是正）

  python3 07_ソーススクリプト/fix_soan_v2340_r3review.py

再レビュー原本　09_元資料/レビュー/川崎町_第10期計画_RedTeam再レビュー第3回_R8.10.24受領.md

  R3-01 3-3　「在宅から施設・居住系への移行が進んでいます」の断定。計画値と実績の乖離だけでは個人の
        移行は示せない。乖離の事実と、影響した可能性・原因が未確定であることに分けて書く
  R3-02 6-1　推計値と自立度Ⅱ以上の実数の「差が、認定を受けておられない認知症の方のおおよその規模」
        は、把握の方法と対象が違うため言えない。単純な差を未認定の人数と読まない旨に改める
  R3-04 10-4　「認知症基本法第12条第5項」「第12条第6項」を、市町村計画に当たる
        「第13条第3項において準用する第12条第5項・第6項」に揃える
  R3-08 5-4・10-4　通いの場の参加率15.9％（n=439）は一般高齢者調査の回答者のうちの割合。KPI の名称に
        分母を書き、行政実績の参加率（参加者数÷65歳以上人口）は別の指標とする
  R3-09 6-1　年齢階級別の認知症の人の数の内訳の和（372人）と計（374人）の差を、注で具体的に書く
        （各行を四捨五入して表示し、計は小数のまま合計してから四捨五入）

終了コード　0＝検査すべて適合／1＝1件でも不適合
"""
import importlib.util
import os
import re
import shutil
import sys

import docx

MOTO = "01_第10期_最新版成果品/川崎町_計画書素案_v2.33_協議用.docx"
SAKI = "01_第10期_最新版成果品/川崎町_計画書素案_v2.34_協議用.docx"
HERE = os.path.dirname(os.path.abspath(__file__))


def kakikae(p, s):
    p.runs[0].text = s
    for r in p.runs[1:]:
        r.text = ""


def zenbun(d):
    return "\n".join([p.text for p in d.paragraphs]
                     + [c.text for t in d.tables for r in t.rows for c in r.cells])


def subete(d):
    """本文と表（入れ子を含む）の段落を1度ずつ。"""
    seen = set()
    for p in list(d.paragraphs) + [p for t in d.tables for r in t.rows for c in r.cells for p in c.paragraphs]:
        if id(p._p) in seen:
            continue
        seen.add(id(p._p))
        yield p


def main():
    futeki = []

    def judge(ok, mi, setsu=""):
        print(("   ○ " if ok else "   × ") + mi + (("　" + setsu) if setsu else ""))
        if not ok:
            futeki.append(mi)

    print("協議用素案 Ver.2.33 → Ver.2.34（Red Team 再レビュー第3回の指摘の是正）")
    shutil.copy2(MOTO, SAKI)
    d = docx.Document(SAKI)
    mae = zenbun(d)

    def rep(kyu, shin, n_min=1, mi=""):
        n = 0
        for p in subete(d):
            if kyu in p.text and p.runs:
                kakikae(p, p.text.replace(kyu, shin))
                n += 1
        judge(n >= n_min, f"{mi}「{kyu[:24]}」（{n}か所）")
        return n

    print("  ── R3-01 3-3 在宅から施設・居住系への移行の断定")
    rep("総額が一致していることの内側で、在宅から施設・居住系への移行が進んでいます。居宅サービスが計画を下回ったのは、"
        "町内事業所の縮小（訪問介護の登録ヘルパーの減少）と、住所地特例24人（令和8年3月末）に表れる町外施設の利用の"
        "常態化が同時に進んだことが影響している可能性があります（因果関係は、町内外別の利用の推移等による追加の検証を"
        "要します）。第10期のサービス見込量は、この移行が続くことを前提としています。",
        "総額は計画に近い一方、居宅サービスは計画を下回り、施設サービスは上回りました。この乖離には、町内事業所の縮小"
        "（訪問介護の登録ヘルパーの減少）と、住所地特例24人（令和8年3月末）に表れる町外施設の利用の常態化が影響した"
        "可能性がありますが、在宅から施設等への移行の実態及び原因は確かめられていません（町内外別の利用者数や経年の"
        "推移等による追加の検証を要します）。第10期のサービス見込量は、令和7年度の利用実績を出発点としているため、"
        "この利用の構成が続くことを前提としています。", mi="3-3")

    print("  ── R3-02 6-1 推計値と自立度Ⅱ以上の実数の差")
    rep("両者の差が、認定を受けておられない認知症の方のおおよその規模に当たります。",
        "両者は把握の方法と対象が異なるため、単純な差を、要介護認定を受けておられない認知症の方の数と読むことは"
        "できません。地域の状況を検討する参考として並べて掲げます。", mi="6-1")

    print("  ── R3-04 10-4 認知症基本法の条項（市町村計画は第13条第3項による準用）")
    rep("（認知症基本法第12条第5項）", "（認知症基本法第13条第3項において準用する第12条第5項）", mi="10-4")
    rep("認知症基本法第12条第6項により", "認知症基本法第13条第3項において準用する第12条第6項により", mi="10-4")

    print("  ── R3-08 通いの場の参加率の分母")
    SHIN_KPI = "通いの場への参加割合（一般高齢者調査の回答者。65歳以上）"
    n = rep("通いの場参加率（65歳以上）", SHIN_KPI, mi="KPI")
    n += rep("通いの場への参加率（65歳以上）", SHIN_KPI, mi="KPI")
    judge(n == 2, f"KPI の名称 {n}か所")
    rep("実際の参加率が15.9％（n=439）にとどまる",
        "一般高齢者調査の回答者のうち実際に参加している割合が15.9％（n=439）にとどまる", mi="5-4")
    rep("参加意向55.2％に対し参加率15.9％という開きを埋める取組です。",
        "参加意向55.2％に対し参加割合15.9％（いずれも一般高齢者調査の回答者）という開きを埋める取組です。", mi="5-4 表")
    # 出典の欄に、行政実績の参加率は別の指標である旨
    n = 0
    for t in d.tables:
        for r in t.rows:
            cs = r.cells
            if cs and cs[0].text.strip() == SHIN_KPI:
                last = cs[-1]
                if "行政実績" not in last.text:
                    kakikae(last.paragraphs[0], last.text.strip() + "（調査の回答者のうちの割合。通いの場の参加者数÷"
                            "65歳以上人口による参加率は、地域支援事業のデータの受領後に別の指標として置きます）")
                    for q in last.paragraphs[1:]:
                        q._p.getparent().remove(q._p)
                    n += 1
    judge(n == 2, f"KPI の出典の欄に分母と別の指標の旨（{n}か所）")

    print("  ── R3-09 6-1 認知症の人の数の内訳の和と計")
    spec = importlib.util.spec_from_file_location("calc_ninchisho", os.path.join(HERE, "calc_ninchisho_R8.10.19.py"))
    cn = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(cn)
    r = cn.santei(cn.machi_kaikyu(), "認知症")
    wa = sum(round(m[6]) for m in r["meisai"])
    kei = round(r["sei"])
    judge(kei == 374 and wa == 372, f"算定の正本から：内訳の和 {wa}人・計 {kei}人（小数のまま {r['sei']:.2f}人）")
    rep("四捨五入により、内訳の合計と計とは一致しない場合があります。",
        f"各年齢階級の人数は1人未満を四捨五入して表示し、計は小数のまま合計してから四捨五入しています。このため"
        f"認知症の人の数は、内訳の和（{wa}人）と計（{kei}人）が{kei - wa}人一致しません。", mi="6-1 注")

    d.save(SAKI)
    ato = zenbun(docx.Document(SAKI))
    print("  ── 検査1　古い記述")
    for go in ("在宅から施設・居住系への移行が進んで", "この移行が続く", "おおよその規模に当たります",
               "（認知症基本法第12条第5項）", "認知症基本法第12条第6項により", "通いの場参加率（65歳以上）",
               "通いの場への参加率（65歳以上）", "実際の参加率が15.9％"):
        judge(go not in ato, f"「{go}」0件")
    print("  ── 検査2　未確定の印・他団体名・受託者の語は増減なし")
    for w in ("【町確認】", "【委員会協議】", "［要確認］", "［要協議］", "［要内訳］", "大雪", "北塩原", "金ヶ崎",
              "川崎市", "受託者", "当社", "ビズアップ"):
        judge(ato.count(w) == mae.count(w), f"「{w}」{mae.count(w)}→{ato.count(w)}")
    print(f"  {SAKI}")
    if futeki:
        print(f"   × 不適合 {len(futeki)}件")
        sys.exit(1)
    print("   ○ 不適合0件")


if __name__ == "__main__":
    main()
