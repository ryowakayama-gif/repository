# -*- coding: utf-8 -*-
"""第2回策定委員会 想定問答集 v5 → v6（宮城県の入所状況調の反映）
（令和8年10月22日）

  python3 07_ソーススクリプト/build_2kai_qa_v6.py

v5 からの変更の理由
  計画素案を Ver.2.23、委員会資料を v10 に上げた。
  (1) QB-3・QB-47 の「公表の項目にない」の整理を改める
      （稼働率・休止床は県の入所状況調で推移が分かった）
  (2) 3問を加える（全61問 → 全64問）
      QB-62 軽費老人ホームの廃止　QB-63 入所前の居所と保険者の関与
      QB-64 他の市町村も同じ方法で把握しているのか
      ⚠ QB-64 は**他団体の計画を確かめていないこと**を正直に答える
      （ご指示により方針Ｃ。根拠は国・県の原典であり他団体に依存しない）
  (3) 計画素案の参照 Ver.2.22 → Ver.2.23

⚠ 宣言した件数と実際の件数が合っていることを機械で確かめる。
"""
import collections
import copy
import importlib.util
import os
import re
import shutil
import sys

import docx
from docx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
spec = importlib.util.spec_from_file_location(
    "fix_soan_v2230", os.path.join(HERE, "fix_soan_v2230_kenchosa.py"))
S23 = importlib.util.module_from_spec(spec)
sys.modules["fix_soan_v2230"] = S23
spec.loader.exec_module(S23)
A = S23.A

SRC = "03_委員会・説明資料/川崎町_第2回策定委員会_想定問答集_R8.11_v5.docx"
DST = "03_委員会・説明資料/川崎町_第2回策定委員会_想定問答集_R8.11_v6.docx"
VER_OLD, VER_NEW = "計画素案Ver.2.22", "計画素案Ver.2.23"
KEN = "宮城県「老人福祉施設等入所状況調」"

NAOSU = [
    ("A．介護サービス情報公表システムにより把握しました。町内14の",
     "A．介護サービス情報公表システムにより把握しました。町内14の"
     "指定すべてについて、定員・在所者数・入所率・待機者数・職種別の"
     "従業者数・前年度の採用者数及び退職者数・指定の年月日・退所先を"
     "採録しています（令和8年10月）。あわせて、入所率と退所理由の推移を"
     f"{KEN}（3か月ごと・{A['時点数']}時点）で把握しました。"
     "事業所にお手数をおかけせず、また町のご回答をお待ちすることなく"
     "把握できました。⚠ いずれの資料にもないものは4件（入所前の居所・"
     "経営の状況・町への要望・短期入所の稼働）です。入所前の居所は県の"
     "入所指針が定める保険者の協力の場で集計値を求めることを検討し、"
     "経営の状況は計画に書かず、町への要望は地域ケア会議・地域密着型"
     "サービス運営協議会という既存の場で聴くこととし、計画期間を通じた"
     "仕組みとして第7章に位置付けます。"),
    ("A．⚠ 入所率は「記入日前月の在所者数÷定員」であり、",
     "A．⚠ 入所率は「記入日前月の在所者数÷定員」であり、延べ利用日数に"
     "よる稼働率ではありません。また、町外の方の利用も含みますので、"
     "第9章の本町の被保険者の利用実績と割ることはできません。内わけは、"
     "介護老人福祉施設が定員30人に対し在所30人（100.0％）、介護老人保健"
     "施設が86人に対し82人（95.3％）、地域密着型介護老人福祉施設が30人に"
     "対し28人（93.3％）、認知症対応型共同生活介護が27人に対し26人"
     f"（96.3％）です。{KEN}で3か月ごとの推移をみると、介護老人福祉施設は"
     f"{S23.haba(A['特養'])}で、1施設には恒常的に空きがあります。"
     "⚠ 施設単位の休止は県の調査で分かりますが、床単位で休止なのか"
     "空床なのかの別は、いずれの資料にもありません。"),
]

ADD = [
    ("カテゴリ B",
     "QB-62　【民生委員】　町内にあった軽費老人ホームはどうなったのか",
     f"A．{KEN}によると、町内にあった軽費老人ホーム（定員"
     f"{A['軽費定員']:.0f}人）は、{S23.wa(S23.KYUSHI)}時点で休止中となり、"
     f"{S23.wa(S23.NAI)}以降は掲載されていません。第9期計画期間が"
     "始まる直前に、町内の高齢者向けの住まい"
     f"{A['軽費定員']:.0f}人分が失われたことになります。"
     "⚠ 軽費老人ホームは老人福祉法に基づく施設で、介護保険の指定事業所"
     "ではないため、介護サービス情報公表システムには載りません。"
     "休止・廃止の経緯と、入居されていた方の行き先を確認しており、"
     "結果は第3回でお示しします。",
     f"【根拠データ】 {KEN}（{S23.wa(S23.KYUSHI)}「休止中」・"
     f"{S23.wa(S23.NAI)}以降掲載なし）",
     "【関連資料】 計画素案Ver.2.23 第2章2-7、"
     "川崎町_施設の入所状況の推移_県調査_R8.10.22.xlsx 05シート、"
     "確認事項No.154"),
    ("カテゴリ B",
     "QB-63　【有識者】　施設に入る前に在宅にいたのか病院にいたのかは、"
     "分からないのか",
     "A．⚠ いずれの公表資料にもありません。ただし、宮城県指定介護老人"
     "福祉施設入所指針（令和5年7月25日）により、施設は入所希望者の"
     "置かれている状況を面談等で確認し、入所希望者名簿を5年間保管する"
     "こととされています。また同指針は、保険者（町）が施設の入所規程の"
     "作成に積極的に協力し必要な援助を行うことを定めています。新たな"
     "調査ではなく、この協力の場で、名簿の集計値（在宅・医療機関・"
     "他の施設の別の人数）の提供を求めることを検討しています。"
     "個人の情報はいただきません。提供を受けられない場合は、退所先と"
     "入所希望者数により整備の必要量を判断します。",
     "【根拠データ】 宮城県指定介護老人福祉施設入所指針 第3・第4(1)(5)",
     "【関連資料】 計画素案Ver.2.23 第7章7-1、確認事項No.159"),
    ("カテゴリ B",
     "QB-64　【有識者】　他の市町村も、同じ方法で施設の状況を"
     "把握しているのか",
     "A．他の市町村の計画との比較は行っていません。本町の把握の方法は、"
     "国の介護サービス情報公表システムと、県が県内全施設について3か月"
     "ごとに公表している入所状況調という、どの市町村も使える公表資料に"
     "よっています。他の市町村の書き方に合わせるのではなく、国・県の"
     "原典に基づいて把握していますので、比較の有無によって本町の数値や"
     "方針は変わりません。ご要望があれば、近隣市町の計画との比較を"
     "第3回までにお示しします。",
     f"【根拠データ】 介護サービス情報公表システム／{KEN}",
     "【関連資料】 川崎町_残る5件の確認方針_R8.10.22.md"),
]


def cat_of(s):
    m = re.match(r"^カテゴリ ([A-H])　", s)
    return "カテゴリ " + m.group(1) if m else None


def txt(el):
    return "".join(n.text or "" for n in el.iter(qn("w:t"))).strip()


def main():
    futeki = []

    def judge(ok, mi, setsu=""):
        print(("   ○ " if ok else "   × ") + mi + (("　" + setsu)
                                                   if setsu else ""))
        if not ok:
            futeki.append(mi)

    print("第2回策定委員会 想定問答集 v5 → v6")
    print()
    shutil.copy(SRC, DST)
    doc = docx.Document(DST)
    body = doc.element.body
    from fix_soan_v111 import set_el, set_cell

    n = 0
    for atama, atarashii in NAOSU:
        hit = [el for el in body.iterchildren()
               if el.tag == qn("w:p") and txt(el).startswith(atama)]
        if len(hit) != 1:
            judge(False, f"『{atama[:20]}…』に当たる段落が{len(hit)}件")
            continue
        set_el(hit[0], atarashii)
        n += 1
    judge(n == len(NAOSU), f"{len(NAOSU)}段落を書き改めた（実際 {n}）")

    n_ver = 0
    for el in body.iterchildren():
        if el.tag != qn("w:p"):
            continue
        s = "".join(x.text or "" for x in el.iter(qn("w:t")))
        if VER_OLD in s or "素案Ver.2.22" in s:
            set_el(el, s.replace("Ver.2.22", "Ver.2.23"))
            n_ver += 1
    judge(n_ver >= 30, f"素案の版の参照を {n_ver}か所改めた")

    kids = list(body.iterchildren())
    tpl_q = tpl_a = None
    for i, el in enumerate(kids):
        if el.tag == qn("w:p") and txt(el).startswith("QB-1　"):
            tpl_q, tpl_a = copy.deepcopy(el), copy.deepcopy(kids[i + 1])
            break

    def last_of(cat):
        cur = last = None
        for el in body.iterchildren():
            if el.tag != qn("w:p"):
                continue
            s = txt(el)
            c = cat_of(s)
            if c:
                cur = c
            if cur == cat and s.startswith("【関連資料】"):
                last = el
        return last

    for cat, q, a, kon, kan in ADD:
        after = last_of(cat)
        for text, tpl in ((kan, tpl_a), (kon, tpl_a), (a, tpl_a), (q, tpl_q)):
            el = copy.deepcopy(tpl)
            set_el(el, text)
            after.addnext(el)
    judge(True, f"{len(ADD)}問を加えた")

    cat_cnt = collections.Counter()
    cur = None
    for el in body.iterchildren():
        if el.tag != qn("w:p"):
            continue
        s = txt(el)
        c = cat_of(s)
        if c:
            cur = c
        if re.match(r"^QB-\d+　", s):
            cat_cnt[cur] += 1
    total = sum(cat_cnt.values())
    for el in body.iterchildren():
        if el.tag != qn("w:p"):
            continue
        s = txt(el)
        m = re.match(r"^(カテゴリ [A-H]　.+?)（\d+問）$", s)
        if m:
            set_el(el, f"{m.group(1)}（{cat_cnt[cat_of(s)]}問）")
        elif re.search(r"全\d+問", s):
            set_el(el, re.sub(r"全\d+問", f"全{total}問", s))
        if "第５版" in s or "v5" in s:
            set_el(el, txt(el).replace("第５版", "第６版").replace("v5", "v6"))
    for t in doc.tables:
        if [c.text.strip() for c in t.rows[0].cells][:3] == \
                ["カテゴリ", "テーマ", "問数"]:
            for r in t.rows[1:]:
                k = "カテゴリ " + r.cells[0].text.strip()
                if k in cat_cnt:
                    set_cell(r.cells[2], str(cat_cnt[k]))
    doc.save(DST)

    d2 = docx.Document(DST)
    zen = "\n".join(p.text for p in d2.paragraphs) + "\n" + "\n".join(
        c.text for t in d2.tables for r in t.rows for c in r.cells)
    print("  ── 検査1　件数")
    qs = [int(x) for x in re.findall(r"QB-(\d+)　", zen)]
    judge(len(qs) == total == 64, f"全64問（実際 {len(qs)}・宣言 {total}）")
    judge(sorted(qs) == list(range(1, 65)), "番号に欠番・重複がない")
    judge(f"全{total}問" in zen, f"「全{total}問」と宣言している")
    print("  ── 検査2　古い前提")
    judge("公表の項目にないものは5件" not in zen,
          "「公表の項目にないものは5件」が0件")
    judge("Ver.2.22" not in zen, "「Ver.2.22」が0件")
    print("  ── 検査3　加えた内容")
    for go in ("軽費老人ホーム", "入所希望者名簿を5年間保管",
               "他の市町村の計画との比較は行っていません", KEN):
        judge(go in zen, f"「{go[:20]}」がある")
    print("  ── 検査4　⚠ 事業所名・他団体名・強調記号")
    for na in ("釜房みどりの園", "アルパイン", "蔵王グリーンホーム",
               "グループホームさわやか", "グループホームふれあい",
               "ながさか", "リリーハイム"):
        judge(na not in zen, f"「{na}」が入っていない")
    judge("**" not in zen, "強調記号（**）が0件")

    print()
    print(f"  {DST}")
    print(f"  全{total}問・" + "／".join(
        f"{k[-1]}{v}問" for k, v in sorted(cat_cnt.items())))
    if futeki:
        print(f"   × 不適合 {len(futeki)}件")
        for m in futeki:
            print(f"      {m}")
        sys.exit(1)
    print("   ○ 検査1〜4　不適合0件")


if __name__ == "__main__":
    main()
