# -*- coding: utf-8 -*-
"""協議用素案 Ver.2.30 → Ver.2.31（第2回策定委員会資料の確定に向けた素案と資料の整合）
（令和8年10月23日）

  python3 07_ソーススクリプト/fix_soan_v2310_kakutei.py

ご指示「第2回策定委員会資料の確定に向けて進めて下さい」による。委員会資料は素案から
起こす（CLAUDE.md §3）ため、資料 v15 と素案 Ver.2.30 の食い違いを素案の側で先に解く。

  ① 3-4 に重点課題（5本）を置く
       資料は「第１０期の重点課題（５本）」を諮るが、素案 3-4 は11の課題の表だけで、
       5本への集約がなかった。11の課題を5本に集約した表を加える。根拠となる事実の数は
       いずれも素案の他の箇所にあるものに限る（資料 v15 の「85歳以上の認定率58.67％」は
       素案になく、第1回資料の値であるため用いない）。
  ② 4-3 の4つの重点と5本の重点課題の対応を書く
  ③ 交付金の27点（22点）の言い回し
       「本計画に書くことで得点できます」「取り戻せます」は言い過ぎ。資料と同じく
       「評価の対象となり得る。得点は国の該当状況調査への回答と実績による」に揃える（4か所）
  ④ 3-3 図3-2 と本文の令和7年度の給付費
       図3-2 の実績は年報の額を10千円単位に丸めた値（計998,220千円）で、2-4 の表の
       総額998,222千円と2つの値が並んでいた。千円単位に揃え、図を作り直して差し替える
       （data_zuhyo.py を直し build_plan_figures.fig32 で作図）
  ⑤ 1-4 の法令が定める手続の項番号
       条文を受領していなかったため項を書いていなかった。介護保険法（e-Gov・令和8年10月1日
       施行時点。09_元資料/法令/ に退避）により、被保険者の意見＝第117条第12項、都道府県の
       意見＝第13項、都道府県知事への提出＝第14項。計画の公表は認知症基本法第13条第3項が
       準用する第12条第4項による。資料 v15 の「第117条第５項」は誤り（資料 v16 で是正）
  ⑥ 5-5 の注記「近隣市町村における事業所の有無は確かめていません［要確認］」
       介護サービス情報公表システム オープンデータ（令和8年6月末）で仙南圏域を確かめた
       （fetch_kinrin_chiiki_R8.10.23.py）。件数は CSV から数える。本町による指定（区域外の
       指定）の状況は公表資料で分からないため、［要確認］はそこに残す

終了コード　0＝検査1〜6すべて適合／1＝1件でも不適合
"""
import copy
import json
import os
import re
import shutil
import sys

import docx
import docx.table as dt
import docx.text.paragraph as dp
from docx.oxml.ns import qn

MOTO = "01_第10期_最新版成果品/川崎町_計画書素案_v2.30_協議用.docx"
SAKI = "01_第10期_最新版成果品/川崎町_計画書素案_v2.31_協議用.docx"
HOREI = "09_元資料/法令/介護保険法_eGov_令和8年10月1日施行時点.json"
FIG32 = "08_図表/fig3-2_keikaku_jisseki.png"
KINRIN = "09_元資料/その他データ/公表システム_仙南圏域_在宅系地域密着型_R8.6末.csv"
KYU_KINRIN = "近隣市町村における事業所の有無は確かめていません［要確認］。"

JUTEN_HYO = [
    ["重点課題", "根拠となる事実", "集約した課題（上表）・接続する施策"],
    ["1　サービス提供体制の維持",
     "第1回策定委員会で報告された1ユニット分の人員の削減、訪問介護の従業者が常勤換算4.5人、"
     "施設・居住系の介護職員（常勤換算）82.8人→80.5人、給付費の平坦化（令和7年度0.8％減）",
     "②・⑧／見込量確保の方策、介護人材の確保、事業所との対話（第7章）"],
    ["2　家族介護者が不在の在宅高齢者",
     "家族介護もサービス利用もない世帯50件（36.2％）。「単身世帯かつ家族介護なし」では"
     "24件中20件（83.3％）が施設入所を申込済み",
     "⑤・⑩／家族介護者への支援、見守り、ケアマネジメント（第5章5-3）"],
    ["3　後期高齢者の増加と重度化",
     "75歳以上が4年で＋15.4％（令和7年度末1,742人）。要介護1・2の平均要介護度の伸びは全国1,580位",
     "①・⑪／介護予防、重度化予防、KPIによる成果の確認（第5章5-1・5-4、第10章）"],
    ["4　移動・外出と社会参加",
     "外出時の移動手段は自動車（自分で運転）74.1％・路線バス5.4％。外出を控えている方19.1％",
     "③／移動支援3制度、通いの場（第5章5-3・5-4）"],
    ["5　相談が届いていない・認知症",
     "家族・友人以外に相談できる相手が「いない」38.8％、認知症相談窓口を知らない52.3％",
     "④・⑥・⑨／相談体制、アウトリーチ、認知症施策（第5章5-4、第6章）"],
]
JUTEN_MAE = ("これら11の課題を、①調査又は実績で裏付けのある事実があること、②第10期の3年間で"
             "打ち手があること、の2つを基準に、次の5本の重点課題に集約します。⑦（3計画の整合）は"
             "特定の重点課題に限らず、計画全体の進め方（第10章）で扱います。")
JUTEN_ATO = ("5本の重点課題に対する施策のうち、とくに力を入れるものを4-3の4つの重点として定めます。")
SHITEN_43 = ("4つの重点と3-4の重点課題（5本）との対応は、重点1が重点課題5（相談が届いていない・"
             "認知症）、重点2が重点課題4（移動・外出と社会参加）、重点3が重点課題2（家族介護者が"
             "不在の在宅高齢者）、重点4が重点課題1（サービス提供体制の維持）です。重点課題3"
             "（後期高齢者の増加と重度化）は、柱1・柱4の介護予防と重度化予防の施策で受けます。")

TOKUTEN = [
    ("仕組みを定めることで取り戻せます（本節の後段）。",
     "仕組みを定めて実施することにより評価の対象となり得ます（本節の後段）。"),
    ("いずれも新たな事業も予算も要さず、仕組みを定めて本計画に書くことで得点できます。"
     "失点の額としては小さいものの、取り戻せる確実性は最も高い部分です。",
     "いずれも新たな事業も予算も要さず、仕組みを定めて実施することにより評価の対象となり得る"
     "ものです。ただし、仕組みを定めることのみで得点が確定するものではなく、実際の評価は国の"
     "該当状況調査への回答と実績によります。失点の額としては小さいものの、取り戻せる見込みが"
     "最も高い部分です。"),
    ("様式を定めて継続することで得点できます。",
     "様式を定めて継続することにより評価の対象となり得ます。"),
    ("仕組みを定めて実施することで得点できます。",
     "仕組みを定めて実施することにより評価の対象となり得ます。"),
]

KYU_CHUI_14 = ("⚠ 条文を受領していないため、条番号のみを記載し、項番号は記載していません。"
               "国の告示及び関係政省令の公布を受けて照合します。")
SHIN_CHUI_14 = ("⚠ 介護保険法の項番号は、令和8年10月1日施行時点の条文（e-Gov法令検索）によります。"
                "令和8年法律第51号による改正で項が移る場合は、施行の時期を確かめて照合し直します。")


def kakikae(p, s):
    p.runs[0].text = s
    for r in p.runs[1:]:
        r.text = ""


def zenbun(d):
    return "\n".join([p.text for p in d.paragraphs]
                     + [c.text for t in d.tables for r in t.rows for c in r.cells])


def jobun(num, kou):
    j = json.load(open(HOREI, encoding="utf-8"))

    def text(n):
        if isinstance(n, str):
            return n
        if isinstance(n, dict):
            return "".join(text(c) for c in n.get("children", []))
        return "".join(text(c) for c in n) if isinstance(n, list) else ""

    def find(n):
        if isinstance(n, dict):
            if n.get("tag") == "Article" and n.get("attr", {}).get("Num") == num:
                return n
            for c in n.get("children", []):
                r = find(c)
                if r:
                    return r
        elif isinstance(n, list):
            for c in n:
                r = find(c)
                if r:
                    return r
    a = find(j["law_full_text"])
    p = [c for c in a["children"] if isinstance(c, dict) and c.get("attr", {}).get("Num") == kou]
    return text(p[0]) if p else ""


def set_cell(c, s):
    kakikae(c.paragraphs[0], s)
    for q in c.paragraphs[1:]:
        q._p.getparent().remove(q._p)


def main():
    futeki = []

    def judge(ok, mi, setsu=""):
        print(("   ○ " if ok else "   × ") + mi + (("　" + setsu) if setsu else ""))
        if not ok:
            futeki.append(mi)

    print("協議用素案 Ver.2.30 → Ver.2.31（第2回策定委員会資料の確定に向けた整合）")
    for f in (MOTO, HOREI, FIG32, KINRIN):
        if not os.path.exists(f):
            print(f"   × 入力がありません：{f}")
            sys.exit(1)
    shutil.copy2(MOTO, SAKI)
    d = docx.Document(SAKI)
    mae = zenbun(d)

    print("  ── ⓪ 根拠となる事実の数が素案に既にある")
    for g in JUTEN_HYO[1:]:
        for x in re.findall(r"[\d,.]+(?:％|人|件|位)", g[1]):
            judge(x in mae, f"「{x}」が素案にある")
    judge("1ユニット" in mae and "0.8％減" in mae, "「1ユニット」「0.8％減」が素案にある")

    print("  ── ① 3-4 に重点課題（5本）")
    t11 = [t for t in d.tables if [c.text for c in t.rows[0].cells][:2] == ["No.", "課題"]]
    judge(len(t11) == 1 and len(t11[0].rows) == 12, "3-4 の11の課題の表を特定")
    hon = [p for p in d.paragraphs if p.text.startswith("キックオフ会議における自治体課題の確認")][0]
    if t11:
        anchor = t11[0]._tbl
        p1 = copy.deepcopy(hon._p)
        anchor.addnext(p1)
        kakikae(dp.Paragraph(p1, hon._parent), JUTEN_MAE)
        tb = copy.deepcopy(t11[0]._tbl)
        p1.addnext(tb)
        trs = tb.findall(qn("w:tr"))
        for tr in trs[len(JUTEN_HYO):]:
            tb.remove(tr)
        haba = (26, 44, 30)
        grid = tb.find(qn("w:tblGrid"))
        cols = grid.findall(qn("w:gridCol"))
        zentai = sum(int(c.get(qn("w:w"))) for c in cols)
        w = [int(zentai * x / 100) for x in haba]
        w[-1] += zentai - sum(w)
        for c, x in zip(cols, w):
            c.set(qn("w:w"), str(x))
        for tr in tb.findall(qn("w:tr")):
            for tc, x in zip(tr.findall(qn("w:tc")), w):
                tc.find(qn("w:tcPr")).find(qn("w:tcW")).set(qn("w:w"), str(x))
        t = dt.Table(tb, hon._parent)
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        for ri, g in enumerate(JUTEN_HYO):
            for ci, v in enumerate(g):
                set_cell(t.rows[ri].cells[ci], v)
                if ri:   # 表中の文言は左揃え（雛形の1列目は番号のため中央揃えだった）
                    t.rows[ri].cells[ci].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        p2 = copy.deepcopy(hon._p)
        tb.addnext(p2)
        kakikae(dp.Paragraph(p2, hon._parent), JUTEN_ATO)
        judge(True, "本文・5本の表・つなぎの文を置いた")

    print("  ── ② 4-3 と重点課題の対応")
    p43 = [p for p in d.paragraphs if p.text.startswith("第10期計画では、上記7つの柱に沿って施策を展開する")]
    judge(len(p43) == 1, "4-3 の導入の段落を特定")
    if p43:
        kakikae(p43[0], p43[0].text + SHITEN_43)

    print("  ── ③ 交付金の得点の言い回し")
    for kyu, shin in TOKUTEN:
        h = [p for p in d.paragraphs if kyu in p.text]
        judge(len(h) == 1, f"「{kyu[:24]}…」")
        for p in h:
            kakikae(p, p.text.replace(kyu, shin))

    print("  ── ④ 図3-2 と本文（令和7年度の給付費 998,222千円）")
    h = [p for p in d.paragraphs if "計画1,011,020千円に対し実績998,220千円" in p.text]
    judge(len(h) == 1, "3-3 の本文")
    for p in h:
        kakikae(p, p.text.replace("実績998,220千円", "実績998,222千円"))
    kids = list(d.element.body.iterchildren())
    cap = [i for i, el in enumerate(kids) if el.tag == qn("w:p")
           and "".join(n.text or "" for n in el.iter(qn("w:t"))).startswith("図3-2　")]
    blip = kids[cap[0] - 1].findall(".//" + qn("a:blip")) if cap else []
    judge(len(blip) == 1, "図3-2 の画像を特定（キャプションの直前の段落）")
    if blip:
        part = d.part.related_parts[blip[0].get(qn("r:embed"))]
        part._blob = open(FIG32, "rb").read()
        judge(True, f"図3-2 の画像を差し替えた（{len(part._blob):,}バイト）")

    print("  ── ⑤ 1-4 法令が定める手続の項番号")
    j12, j13, j14 = jobun("117", "12"), jobun("117", "13"), jobun("117", "14")
    judge("被保険者の意見を反映させるために必要な措置" in j12, "第117条第12項＝被保険者の意見")
    judge("都道府県の意見を聴かなければならない" in j13, "第117条第13項＝都道府県の意見")
    judge("都道府県知事に提出しなければならない" in j14, "第117条第14項＝知事への提出")
    t14 = [t for t in d.tables if [c.text.replace("　", "") for c in t.rows[0].cells][:2] == ["手続", "根拠"]]
    judge(len(t14) == 1, "1-4 の表を特定")
    if t14:
        tb = t14[0]
        for r in tb.rows:
            k = r.cells[0].text
            if k.startswith("被保険者の意見を反映"):
                set_cell(r.cells[1], "介護保険法第117条第12項")
            elif k.startswith("都道府県の意見の聴取"):
                set_cell(r.cells[1], "介護保険法第117条第13項")
            elif k.startswith("計画の公表"):
                set_cell(r.cells[1], "認知症基本法第13条第3項において準用する第12条第4項")
                tr = copy.deepcopy(r._tr)
                r._tr.addnext(tr)
                nr = dt._Row(tr, tb)
                for c, v in zip(nr.cells, ["都道府県知事への提出", "介護保険法第117条第14項",
                                           "計画を定めたのち", "遅滞なく宮城県知事に提出します"]):
                    set_cell(c, v)
        judge(True, "項番号を入れ、知事への提出の行を加えた")
    h = [p for p in d.paragraphs if p.text == KYU_CHUI_14]
    judge(len(h) == 1, "1-4 の注記")
    for p in h:
        kakikae(p, SHIN_CHUI_14)

    print("  ── ⑥ 5-5 近隣市町村の事業所（仙南圏域）")
    import collections
    import csv
    with open(KINRIN, encoding="utf-8-sig", newline="") as f:
        kr = list(csv.DictReader(f))
    ken = collections.defaultdict(list)
    for r in kr:
        ken[r["サービスの種類"]].append(r["市町村"])

    def ari(na):
        m = ken.get(na, [])
        jun = sorted(set(m), key=lambda x: m.index(x))
        return f"{len(m)}事業所（{'・'.join(jun)}）" if m else "なし"
    shin = ("厚生労働省「介護サービス情報公表システム」オープンデータ（令和8年6月末時点）でみると、"
            "宮城県の高齢者福祉圏域である仙南圏域（2市7町）には、定期巡回・随時対応型訪問介護看護が"
            f"{ari('定期巡回・随時対応型訪問介護看護')}、看護小規模多機能型居宅介護が"
            f"{ari('看護小規模多機能型居宅介護')}、小規模多機能型居宅介護が"
            f"{ari('小規模多機能型居宅介護')}あり、夜間対応型訪問介護は"
            + ("ありません" if not ken.get("夜間対応型訪問介護") else ari("夜間対応型訪問介護") + "あります")
            + "。地域密着型サービスを町外の事業所で利用するには本町に"
            "よる指定が要り、その状況は確かめています［要確認］。")
    h = [p for p in d.paragraphs if KYU_KINRIN in p.text]
    judge(len(h) == 1 and len(kr) > 0, f"5-5 の旧の文（仙南圏域 {len(kr)}事業所）")
    for p in h:
        kakikae(p, p.text.replace(KYU_KINRIN, shin))

    d.save(SAKI)
    ato = zenbun(docx.Document(SAKI))
    print("  ── 検査1　古い記述が残っていない")
    for go in ("得点できます", "取り戻せます", "998,220", "条文を受領していないため"):
        judge(go not in ato, f"「{go}」0件")
    print("  ── 検査2　重点課題（5本）")
    judge(all(g[0] in ato for g in JUTEN_HYO[1:]), "5本の名称がある")
    judge(ato.count("998,222") >= mae.count("998,222") + 1, "998,222 が本文に増えた")
    print("  ── 検査3　未確定の印")
    for w in ("［要確認］", "［要協議］", "［要内訳］", "【町確認】"):
        judge(ato.count(w) == mae.count(w), f"{w} 増減なし")
    print("  ── 検査4　他団体名・受託者の語")
    for w in ("大雪", "北塩原", "金ヶ崎", "川崎市", "受託者", "当社", "ビズアップ"):
        judge(ato.count(w) == mae.count(w), f"「{w}」が増えていない")
    print("  ── 検査5　消えた数は旧の文の中のものだけ")
    kyu = "".join(k for k, _ in TOKUTEN) + "実績998,220千円" + KYU_CHUI_14 + "介護保険法第117条" + KYU_KINRIN
    kieta = sorted({m.group(0) for m in re.finditer(r"[0-9][0-9,.]*(?:人|％|%|千円|件|頁|点|円)", mae)
                    if mae.count(m.group(0)) > ato.count(m.group(0)) and m.group(0) not in kyu})
    judge(not kieta, "消えた数 0件", "／".join(kieta[:5]))
    print("  ── 検査6　「第117条第５項」0件")
    judge("第117条第５項" not in ato and "第117条第5項" not in ato, "誤った項番号 0件")
    print(f"  {SAKI}")
    if futeki:
        print(f"   × 不適合 {len(futeki)}件")
        sys.exit(1)
    print("   ○ 不適合0件")


if __name__ == "__main__":
    main()
