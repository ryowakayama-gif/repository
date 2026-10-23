# -*- coding: utf-8 -*-
"""第2回策定委員会資料 v16 → v17（Red Team レビュー（令和8年10月24日）の指摘の是正）

  python3 07_ソーススクリプト/build_2kai_v17.py

レビュー原本　09_元資料/レビュー/川崎町_第10期計画_RedTeamレビュー_R8.10.24受領.md
内容の正本　　協議用素案 Ver.2.32（fix_soan_v2320_redteam.py）

  A03 選択２の表　5,959円・6,382円は算定Ｂが6,170円だった時の値（6,170∓211）が残っていた。
       素案 9-3 の算定Ｂ と 5-5 の「10人では約211円」から起こす（6,143∓211）
  A04 ６「保険料を動かす要因（算定Ｂ 6,143円を基準）」の差額は、素案 9-3 の感度の表と同じく
       算定Ａ（取崩しなし）を基準とした値だった。見出しを算定Ａに改め、報酬改定の行に
       算定Ｂ・Ｃの差を表26 から計算して併記する
  A06 ７「ご意見をいただくこと自体が…要件イを満たします」「そのまま要件を満たすことになります」
       → 要件に当たる取組の一つとなり得る（評価は該当状況調査への回答と記録による）
  B03 選択５「国は基金残高の概ね半分以下を取崩水準として示しています」は根拠文書を確かめられない。
       国の基準とは書かず、事務局案の考え方として書く
  B12 選択１　素案がＣ案で仮に組まれていることと、Ａ・Ｂとなった場合の組み替えを書く
  B15 「町内14の指定」→「町内８事業所の14の指定」。選択２の「入所・居住系」の範囲を素案と同じに
"""
import importlib.util
import os
import re
import shutil
import sys

import docx

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)


def load(name, fn):
    spec = importlib.util.spec_from_file_location(name, os.path.join(HERE, fn))
    m = importlib.util.module_from_spec(spec)
    sys.modules[name] = m
    spec.loader.exec_module(m)
    return m


V9 = load("build_2kai_v9", "build_2kai_v9.py")
SRC = "03_委員会・説明資料/川崎町_第2回策定委員会資料_R8.11_v16.docx"
DST = "03_委員会・説明資料/川崎町_第2回策定委員会資料_R8.11_v17.docx"
SOAN = "01_第10期_最新版成果品/川崎町_計画書素案_v2.32_協議用.docx"


def kakikae(p, s):
    p.runs[0].text = V9.clean(s)
    for r in p.runs[1:]:
        r.text = ""


def yen(s):
    return int(re.sub(r"[^\d]", "", s))


def main():
    futeki = []
    KYU = []

    def judge(ok, mi, setsu=""):
        print(("   ○ " if ok else "   × ") + mi + (("　" + setsu) if setsu else ""))
        if not ok:
            futeki.append(mi)

    print("第2回策定委員会資料 v16 → v17（Red Team レビューの指摘の是正）")
    S = docx.Document(SOAN)
    SP = [p.text for p in S.paragraphs]
    sz = "\n".join(SP + [c.text for t in S.tables for r in t.rows for c in r.cells])
    b = yen(re.search(r"パターンB ([\d,]+)円", sz).group(1))
    j10 = int(re.search(r"10人では約(\d+)円動きます", sz).group(1))
    teigi = re.search(r"入所・居住系（(施設サービス、認知症対応型共同生活介護、地域密着型介護老人福祉施設"
                      r"入所者生活介護及び特定施設入居者生活介護)）", sz).group(1)
    judge(True, f"素案から：算定Ｂ {b:,}円・10人 約{j10}円・入所・居住系＝{teigi[:20]}…")

    shutil.copy2(SRC, DST)
    d = docx.Document(DST)
    mae = "\n".join([p.text for p in d.paragraphs] + [c.text for t in d.tables for r in t.rows for c in r.cells])

    def p_rep(atama, kyu, shin, n=1):
        h = [p for p in d.paragraphs if p.text.strip().startswith(atama) and kyu in p.text]
        judge(len(h) == n, f"『{atama[:18]}』の「{kyu[:16]}」（{len(h)}か所）")
        for p in h:
            KYU.append(p.text)
            kakikae(p, p.text.replace(kyu, shin))

    def c_rep(kyu, shin):
        n = 0
        for t in d.tables:
            seen = []
            for r in t.rows:
                for c in r.cells:
                    if c._tc in seen:
                        continue
                    seen.append(c._tc)
                    if kyu in c.text:
                        KYU.append(c.text)
                        V9.set_cell(c, c.text.replace(kyu, shin))
                        n += 1
        return n

    print("  ── A03 選択２の表")
    t2 = [t for t in d.tables if [c.text.strip() for c in t.rows[0].cells][:3] == ["選択肢", "内　容", "保険料（月額）"]
          and "定員を維持" in t.rows[1].cells[1].text]
    judge(len(t2) == 1, "選択２の表を特定")
    if t2:
        KYU.append("\n".join(c.text for r in t2[0].rows for c in r.cells))
        V9.set_cell(t2[0].rows[1].cells[2], f"{b:,}円")
        V9.set_cell(t2[0].rows[2].cells[2], f"{b - j10:,}円（▲{j10}円）")
        V9.set_cell(t2[0].rows[3].cells[2], f"{b + j10:,}円（＋{j10}円）")
    p_rep("いずれも算定Ｂ（介護給付費準備基金の50％取崩し）の場合の額です。",
          "10人で月額約211円動きます。",
          f"10人で月額約{j10}円動きます（基金の取崩しの額によらず同じ幅です）。")

    print("  ── A04 保険料を動かす要因の基準")
    p_rep("◆ 保険料を動かす要因（算定Ｂ 6,143円を基準）", "（算定Ｂ 6,143円を基準）",
          "（算定Ａ 6,822円＝基金の取崩しなしを基準）")
    t26 = [t for t in d.tables if t.rows[0].cells[0].text.strip() == "介護報酬改定率"][0]
    g = {r.cells[0].text.strip()[:2]: [yen(c.text) for c in r.cells[1:4]] for r in t26.rows[1:4]}
    z, i1, i3 = g["０％"], g["１％"], g["３％"]
    d1 = [a - b_ for a, b_ in zip(i1, z)]
    d3 = [a - b_ for a, b_ in zip(i3, z)]
    judge(d1[0] == 66 and d3[0] == 196, f"表26 から 1％：Ａ＋{d1[0]}・Ｂ＋{d1[1]}・Ｃ＋{d1[2]}／3％：Ａ＋{d3[0]}・Ｂ＋{d3[1]}・Ｃ＋{d3[2]}")
    t27 = [t for t in d.tables if [c.text.strip() for c in t.rows[0].cells] == ["要　因", "影　響"]][0]
    for r in t27.rows:
        k = r.cells[0].text.strip()
        if k == "介護報酬改定率が１％だった場合":
            KYU.append(r.cells[1].text)
            V9.set_cell(r.cells[1], f"＋{d1[0]}円（算定Ｂでは＋{d1[1]}円、算定Ｃでは＋{d1[2]}円）")
        elif k == "同 ３％だった場合":
            KYU.append(r.cells[1].text)
            V9.set_cell(r.cells[1], f"＋{d3[0]}円（算定Ｂでは＋{d3[1]}円、算定Ｃでは＋{d3[2]}円）")
    c = [t for t in d.tables if t.rows[0].cells[0].text.startswith("最も大きな変動要因は介護給付費準備基金の取崩額")]
    if c:
        cc = c[0].rows[0].cells[0]
        KYU.append(cc.text)
        V9.set_cell(cc, cc.text + "上の差額は、基金の取崩し・入所・居住系の人数を除き、算定Ａを基準とした"
                    "値です。取崩しを行う算定Ｂ・Ｃでは、保険料に掛かる割合で動く要因（報酬改定・"
                    "負担割合・収納率）の差がこれと異なります。")
        judge(True, "表28 に基準の注記")

    print("  ── A06 交付金の要件の言い切り")
    n = c_rep("ご意見をいただくこと自体が、「評価結果の活用」の要件イ（庁外の関係者の参画）を満たします。"
              "評価指標は「地域ケア会議や計画策定委員会等既存の会議体に、議題を追加して報告」することで足りると定めています。",
              "ご意見をいただくことは、「評価結果の活用」の要件イ（庁外の関係者の参画）に当たる取組の一つと"
              "なり得ます。評価指標は「地域ケア会議や計画策定委員会等既存の会議体に、議題を追加して報告」する"
              "ことで足りるとしています。評価は、国の該当状況調査への回答と、本日の議事の記録・反映の結果に"
              "よります。")
    judge(n == 1, f"表36（{n}か所）")
    n = c_rep("本日ご報告し、ご意見をいただくことが、そのまま要件を満たすことになります。",
              "本日ご報告し、ご意見をいただくことは、この要件に当たる取組の一つとなり得ます（評価は国の"
              "該当状況調査への回答と記録によります）。")
    judge(n == 1, f"表38（{n}か所）")

    print("  ── B03 基金の取崩しの国の水準")
    p_rep("リスク：Ａは第11期に基金を残せますが", "国は基金残高の概ね半分以下を取崩水準として示しています。",
          "事務局案Ｂは、基金残高の見通しと第11期以降の保険料の平準化を踏まえ、比較案の一つとして"
          "50％取崩しを示すものです。")

    print("  ── B12 選択１（素案はＣ案で仮に組んでいる）")
    p_rep("リスク：Ａは第９期からの継続性が最も高く", "事務局はＣをお勧めします。",
          "事務局はＣをお勧めします。なお、計画素案は町のご回答（令和８年９月25日）によりＣ案で仮に"
          "組んでおり、委員会のご承認の前のものです。Ａ又はＢとなった場合は、第５章から第８章の節・項を"
          "組み替えます。")

    print("  ── B15 集計の単位と範囲")
    n = c_rep("町内14の指定すべてについて把握しました", "町内８事業所の14の指定すべてについて把握しました")
    n += c_rep("公表情報により町内14の指定を把握", "公表情報により町内８事業所の14の指定を把握")
    p_rep("第１回で最も議論が集中した論点です。", "町内14の指定すべてについて把握しました",
          "町内８事業所の14の指定すべてについて把握しました")
    judge(n == 2, f"表6・表39（{n}か所）")
    p_rep("令和７年度の給付費998,222千円のうち",
          "入所・居住系（施設サービス・地域密着型サービス・特定施設入居者生活介護）",
          f"入所・居住系（{teigi}）")

    print("  ── 段落の書式の要素の順（v16 と同じ）")
    from docx.oxml.ns import qn
    JUN = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
           "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
           "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE",
           "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind",
           "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
           "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
           "sectPr", "pPrChange"]
    for pPr in d.element.body.iter(qn("w:pPr")):
        ko = list(pPr)
        jun = sorted(ko, key=lambda c: JUN.index(c.tag.split("}")[1]) if c.tag.split("}")[1] in JUN else len(JUN))
        if jun != ko:
            for c_ in jun:
                pPr.append(c_)
    d.save(DST)

    d2 = docx.Document(DST)
    ato = "\n".join([p.text for p in d2.paragraphs] + [c.text for t in d2.tables for r in t.rows for c in r.cells])
    print("  ── 検査1　古い記述")
    for go in ("5,959", "6,382", "（算定Ｂ 6,143円を基準）", "を満たします", "そのまま要件を満たす",
               "概ね半分以下", "町内14の指定", "地域密着型サービス・特定施設"):
        judge(go not in ato, f"「{go}」0件")
    print("  ── 検査2　素案と同じ値")
    judge(f"{b - j10:,}円" in ato and f"{b + j10:,}円" in ato, f"選択２ {b - j10:,}円・{b + j10:,}円（素案の算定Ｂ±{j10}円）")
    judge(teigi in ato and teigi in sz, "入所・居住系の範囲が素案と同じ")
    print("  ── 検査3　消えた数は旧の文に現れる")
    kyu = "\n".join(KYU)
    kieta = sorted({m.group(0) for m in re.finditer(r"[0-9０-９][0-9０-９,.]*(?:人|名|％|千円|件|頁|日|点|円|位)", mae)
                    if mae.count(m.group(0)) > ato.count(m.group(0)) and m.group(0) not in kyu})
    judge(not kieta, "消えた数はすべて旧の文に現れる", "／".join(kieta[:5]))
    print("  ── 検査4　事業所名・他団体名・⚠")
    for na in ("釜房", "アルパイン", "鶴寿会", "グループホームさわやか", "グループホームふれあい",
               "北塩原", "金ヶ崎", "大雪", "川崎市"):
        judge(na not in ato, f"「{na}」0件")
    judge("⚠" not in ato and "**" not in ato, "⚠・** 0件（第1回の書式）")
    print(f"  {DST}")
    print(f"  ⚠ 町のご確認が要る印【町確認】が {ato.count('【町確認】')}件（No.189）")
    if futeki:
        print(f"   × 不適合 {len(futeki)}件")
        sys.exit(1)
    print("   ○ 検査1〜4　不適合0件")


if __name__ == "__main__":
    main()
