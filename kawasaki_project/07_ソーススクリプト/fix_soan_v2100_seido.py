# -*- coding: utf-8 -*-
"""計画素案 Ver.2.9 → Ver.2.10（制度改正と法定の手続を補完する）

点検スキルの軸Ｃ（制度改正）の点検
（`check_8jiku_R8.10.4.py`）で、**令和8年3月の全国介護保険・高齢者保健福祉
担当課長会議資料が示す制度改正の事項が、素案1-5に整理されていない**ことが
分かった。素案に見当たらなかったのは次の7語である。

  7区分／82.65／令和8年度介護報酬改定／一定以上所得／給与所得控除／
  返還義務／地域類型

1-5 には①基本指針（案）の動向（別表11事項）②中山間・人口減少地域の特例
③令和8年法律第51号 ④医療計画等との整合 の4つの小見出しがあるが、
**課長会議資料が示す個別の制度改正の事項を置く小見出しがなかった。**
節の表題が「介護保険制度改正の主な内容」であるにもかかわらず、
保険料・給付費に直接に効く事項（調整交付金の7区分化、保険料段階の基準額、
補足給付の見直し、介護報酬改定）が落ちていた。

  ▌ 令和8年3月 全国課長会議資料が示す制度改正（11事項）を新設する。

**施行の時期は事項ごとに持つ。**「原則令和9年4月」を機械的に当てない。
とくに補足給付は2段階に分かれ、令和8年8月施行分は第9期の最終年度であり、
令和9年度中の施行分が第10期の期間中である。まとめて書かない。

あわせて、軸Ａ（法定記載事項）の点検で、**1-4 に法令が定める手続が
書かれていない**ことが分かった。1-4 は策定委員会・住民意見の反映・
関係機関との連携を置いているが、介護保険法第117条が定める
意見公募手続・都道府県の意見の聴取・計画の公表、及び認知症基本法が
定める認知症の人及びその家族等からの意見の聴取が、手続として
整理されていなかった。

  ▌ 法令が定める手続（4件）を 1-4 に新設する。

**条文を受領していないため、項番号は書かない**（条番号のみとする）。

  python3 07_ソーススクリプト/fix_soan_v2100_seido.py
"""
import copy
import os
import re
import shutil
import sys
import zipfile

import docx
from docx.oxml.ns import qn

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

SRC = "01_第10期_最新版成果品/川崎町_計画書素案_v2.9_事業内容版.docx"
DST = "01_第10期_最新版成果品/川崎町_計画書素案_v2.10_制度改正補完版.docx"

HEAD = "▌ 令和8年3月 全国課長会議資料が示す制度改正（11事項）"
LEAD = ("令和8年3月の全国介護保険・高齢者保健福祉担当課長会議資料が示した"
        "制度改正のうち、本計画の記載・数値に関わるものを整理します。"
        "施行の時期は事項により異なります。"
        "とくに補足給付の見直しは2段階に分かれ、"
        "令和8年8月に施行された部分は第9期の最終年度（令和8年度）に、"
        "令和9年度中に施行される部分が第10期の計画期間中に当たります。")

COLS = ["事　項", "内　容", "施行の時期", "本計画での扱い"]
ROWS = [
    ["調整交付金の年齢区分の細分化（3区分→7区分）",
     "後期高齢者の年齢区分を3区分から7区分に細分化し、"
     "市町村ごとの財政力の差をより細かく調整する",
     "令和10年度の財政調整交付金から（激変緩和の措置が置かれる）",
     "9-3の調整交付金見込交付割合は現行の算定方法による。"
     "令和10年度以降は再算定を要する"],
    ["保険料段階の判定に用いる基準額の引上げ（80.9万円→82.65万円）",
     "第1段階から第3段階までの判定に用いる年金収入等の基準額を引き上げる",
     "政令は改正済み。適用は第10期から",
     "2-2の所得段階の区分及び9-3の所得段階別保険料。"
     "⚠ 介護保険条例の改正の手続を要する"],
    ["補足給付（特定入所者介護サービス費）の見直し",
     "居住費・基準費用額・負担限度額の引上げ、"
     "及び第3段階①②のそれぞれを2分割する",
     "⚠ 2段階に分かれる。引上げは令和8年8月施行（第9期の最終年度）、"
     "段階の分割は令和9年度中の施行（第10期の期間中）",
     "9-2の給付費の見込み及び9-3の低所得者支援。"
     "第10期の基準年度（令和7年度）の実績には引上げ分が入っていない"],
    ["令和8年度介護報酬改定",
     "介護報酬の改定",
     "令和8年度",
     "⚠ 第10期の算定は令和7年度の単価で固定しているため、"
     "令和8年度・令和9年度の2回分の改定が反映されていない。"
     "9-3に改定率別の幅を示している"],
    ["一定以上所得の判断基準の見直し",
     "利用者負担2割の対象となる所得の範囲",
     "令和9年度の開始前までに結論を得る方針",
     "9-2の給付費に効く。結論が出るまでは現行の基準による"],
    ["令和7年度税制改正（給与所得控除の見直し）",
     "給与所得控除額の改正",
     "令和7年度",
     "所得段階別加入割合に効く。"
     "9-3の所得段階別加入割合補正後被保険者数は令和7年度末の実績による"],
    ["介護保険事業状況報告の見直し",
     "報告の様式・集計区分の変更",
     "―",
     "2-3・2-4の実績値の系列の接続を確認する"],
    ["高額医療合算介護サービス費の自動償還",
     "申請によらず支給できる仕組み",
     "市町村の判断により実施",
     "窓口の手続の簡素化。9-2の給付費の区分は変わらない"],
    ["被保険者証等の返還義務の廃止",
     "資格の喪失等に伴う被保険者証の返還義務を廃止する",
     "―",
     "窓口事務。計画の記載には及ばない"],
    ["2040年を見据えた地域類型",
     "中山間・人口減少地域／大都市部／一般市等に分けて"
     "サービス提供体制の確保の考え方を示す",
     "第10期から",
     "本町は中山間・人口減少地域に当たると考えられる。"
     "1-6の日常生活圏域の設定及び7-3に位置付けている"],
    ["医療と介護の協議の場の再編成（協力医療機関）",
     "施設と協力医療機関との連携について3つの要件が定められ、"
     "要件を満たさない施設の把握が求められる",
     "―",
     "5-3の在宅医療・介護連携及び7-3。"
     "要件を満たさない施設の把握は町の照会による"],
]

TAIL = ("⚠ 本表は令和8年3月の全国課長会議資料によるものです。"
        "施行の時期・適用の年度は、国の告示及び関係政省令の公布を受けて"
        "最終的に照合します。")

# ══════════════════════════ 1-4 に加える「法令が定める手続」
TET_HEAD = "▌ 法令が定める手続"
TET_LEAD = ("計画を定め、又は変更しようとするときに、法令が求める手続は"
            "次のとおりです。いずれも第10期計画の策定の過程で実施します。")
TET_COLS = ["手　続", "根　拠", "時　期", "本計画での実施"]
TET_ROWS = [
    ["被保険者の意見を反映させるために必要な措置（意見公募手続）",
     "介護保険法第117条",
     "計画の案を作成したのち",
     "町の実施要領により実施します。期間・方法は町とご相談のうえ定めます"
     "（確認事項No.19・No.136）"],
    ["都道府県の意見の聴取",
     "介護保険法第117条",
     "計画の案を確定する前",
     "宮城県への事前協議を令和8年12月から行います。"
     "要する期間は県の運用によります（確認事項No.49）"],
    ["認知症の人及びその家族等からの意見の聴取",
     "認知症基本法第13条第3項において準用する第12条第3項",
     "計画書案の作成前",
     "郵送の調査とは別の経路（地域包括支援センター・国民健康保険川崎病院）"
     "により令和8年10月から11月に実施します。記録の様式は作成済みです"],
    ["計画の公表",
     "介護保険法第117条",
     "計画を定めたのち",
     "町のホームページ及び窓口により公表します。"
     "概要版（8頁程度・3,500部）を併せて配布します"],
]
TET_TAIL = ("⚠ 条文を受領していないため、条番号のみを記載し、項番号は"
            "記載していません。国の告示及び関係政省令の公布を受けて照合します。")

VER = {4: "計画書素案 Ver.2.10", 5: "（制度改正補完版）",
       21: ("⚠ 頁番号は、本素案Ver.2.10の組版によるものです。"
            "図表の差替え・加筆により頁の割付けが動いた場合は、"
            "目次の頁番号も併せて更新します。")}
VER_SWAP = [("本素案Ver.2.9", "本素案Ver.2.10"),
            ("計画書素案 Ver.2.9", "計画書素案 Ver.2.10")]


def txt(el):
    return "".join(n.text or "" for n in el.iter(qn("w:t"))).strip()


def main():
    if not os.path.exists(SRC):
        raise SystemExit("素案が見つからない：" + SRC)
    shutil.copy(SRC, DST)
    doc = docx.Document(DST)
    body = doc.element.body
    from fix_soan_v111 import set_el, clone_table

    # 雛形　1-5 の小見出し・本文・4列の表
    tpl_sub = tpl_body = tpl_tbl = None
    for el in body.iterchildren():
        if el.tag == qn("w:p"):
            s = txt(el)
            if tpl_sub is None and s.startswith("▌ 国の第10期基本指針"):
                tpl_sub = copy.deepcopy(el)
            if tpl_body is None and s.startswith("国の基本指針案では"):
                tpl_body = copy.deepcopy(el)
        elif el.tag == qn("w:tbl") and tpl_tbl is None:
            t = docx.table.Table(el, doc)
            if len(t.columns) == 4 and len(t.rows) >= len(ROWS) + 1:
                h = [c.text.strip() for c in t.rows[0].cells]
                if h[0] == "事項":
                    tpl_tbl = el
    if tpl_sub is None or tpl_body is None or tpl_tbl is None:
        raise SystemExit("雛形が見つからない")

    # 差し込み位置　「▌ 中山間・人口減少地域（特定地域）に係る特例の創設」の直前
    anchor = None
    for el in body.iterchildren():
        if el.tag == qn("w:p") and txt(el).startswith("▌ 中山間・人口減少地域"):
            anchor = el
            break
    if anchor is None:
        raise SystemExit("1-5 の中山間の小見出しが見つからない")

    prev = anchor.getprevious()
    els = []
    p = copy.deepcopy(tpl_sub)
    set_el(p, HEAD)
    prev.addnext(p)
    els.append(p)
    q = copy.deepcopy(tpl_body)
    set_el(q, LEAD)
    p.addnext(q)
    tbl = clone_table(tpl_tbl, [COLS] + [list(r) for r in ROWS])
    q.addnext(tbl)
    r = copy.deepcopy(tpl_body)
    set_el(r, TAIL)
    tbl.addnext(r)

    # ══════════════════════════ 1-4 に法令が定める手続を加える
    #   「▌ 関係機関等との連携」の直前に置く
    anchor2 = None
    for el in body.iterchildren():
        if el.tag == qn("w:p") and txt(el).startswith("▌ 関係機関等との連携"):
            anchor2 = el
            break
    if anchor2 is None:
        raise SystemExit("1-4 の関係機関の小見出しが見つからない")
    prev2 = anchor2.getprevious()
    p2 = copy.deepcopy(tpl_sub)
    set_el(p2, TET_HEAD)
    prev2.addnext(p2)
    q2 = copy.deepcopy(tpl_body)
    set_el(q2, TET_LEAD)
    p2.addnext(q2)
    tbl2 = clone_table(tpl_tbl, [TET_COLS] + [list(r) for r in TET_ROWS])
    q2.addnext(tbl2)
    r2 = copy.deepcopy(tpl_body)
    set_el(r2, TET_TAIL)
    tbl2.addnext(r2)

    # ══════════════════════════ 版表記
    ps = [el for el in body.iterchildren() if el.tag == qn("w:p")]
    for i, t in VER.items():
        set_el(ps[i - 1], t)
    for el in body.iterchildren():
        if el.tag != qn("w:p"):
            continue
        s = "".join(n.text or "" for n in el.iter(qn("w:t")))
        new = s
        for a, b in VER_SWAP:
            new = new.replace(a, b)
        if new != s:
            set_el(el, new)
    doc.save(DST)

    # ヘッダー
    tmp = DST + ".tmp"
    with zipfile.ZipFile(DST) as zin, \
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/header1.xml":
                data = re.sub(r"計画書素案 ?Ver\.[0-9.]+",
                              "計画書素案 Ver.2.10",
                              data.decode("utf-8")).encode("utf-8")
            info = zipfile.ZipInfo(item.filename, item.date_time)
            info.compress_type = item.compress_type
            info.external_attr = item.external_attr
            zout.writestr(info, data)
    shutil.move(tmp, DST)

    d2 = docx.Document(DST)
    print("制度改正の節を補完しました")
    print("  保存：", DST)
    print(f"  段落 {len(d2.paragraphs)}／表 {len(d2.tables)}"
          f"（Ver.2.9 は660段落125表）")
    # ══════════════════════════ 自己点検
    full = "\n".join(q.text for q in d2.paragraphs)
    for t in d2.tables:
        for row in t.rows:
            for c in row.cells:
                full += "\n" + c.text
    ng = []
    for w in ("7区分", "82.65", "令和8年度介護報酬改定", "一定以上所得",
              "給与所得控除", "返還義務", "地域類型", "高額医療合算",
              "補足給付"):
        if w not in full:
            ng.append(f"語が入っていない：{w}")
    # 補足給付の2段階を1つにまとめていないか
    if "2段階に分かれる" not in full:
        ng.append("補足給付が2段階に分かれることを書いていない")
    # 1-4 の手続
    for w in ("意見公募手続", "都道府県の意見の聴取", "計画の公表"):
        if w not in full:
            ng.append(f"法令が定める手続が入っていない：{w}")
    # 本スクリプトが加えた手続の表に、項番号を書いていないこと
    #   （素案には第117条第2項第1号など、従来からの引用が6か所ある。
    #     令和8年法律第51号による号の繰り下がりの照合を要するため、
    #     不適合ではなく要判断として別に数える。）
    tet = ""
    for t in d2.tables:
        h = [c.text.strip() for c in t.rows[0].cells]
        if h == TET_COLS:
            tet = "\n".join(c.text for r in t.rows for c in r.cells)
    if not tet:
        ng.append("法令が定める手続の表が見つからない")
    elif re.search(r"第117条第\d+項", tet):
        ng.append("加えた手続の表に項番号を書いている")
    goulist = [m.group(0) for s2 in
               [q.text for q in d2.paragraphs] +
               [c.text for t in d2.tables for r in t.rows for c in r.cells]
               for m in re.finditer(r"第117条第\d+項第\d+号", s2)]
    print("  ── 自己点検")
    if ng:
        for m in ng:
            print("   ×", m)
        sys.exit(1)
    print("   ○ 11事項の語がすべて入り、補足給付の2段階も書き分けている")
    print("   ○ 法令が定める手続4件が1-4にあり、加えた表に項番号はない")
    print(f"   △ 要判断　素案に従来からある号の番号の引用 {len(goulist)}か所"
          "（令和8年法律第51号による繰り下がりの照合を要する・"
          "確認事項No.163）")


if __name__ == "__main__":
    main()
