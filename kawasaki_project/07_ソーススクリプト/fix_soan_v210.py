# -*- coding: utf-8 -*-
"""計画素案 Ver.2.0 → Ver.2.1（新規・拡充42事業の説明と、交付金27点の回復管理）

Ver.2.0で143事業を書き下ろしたが、表は事業名・区分・所管の一覧であり、
「なぜ新たに立てるのか」「何を広げるのか」が読み取れない。
第10期で内容が変わる42事業（新規23・拡充19）について、
根拠となる事実と第10期での内容を章ごとの表で示す。

  ① 5-1〜5-4・6-4・第8章に「新たに立てる事業・広げる事業」の表を置く
     第1章3件・第2章1件・第3章7件・第4章9件・第5章17件・第7章5件＝42件
  ② 10-4に「交付金の失点27点の回復管理」の表を置く
     第3章3-3で特定した体制・取組指標群の失点3か所を、
     KPI管理表の体制指標として管理できる形にする
  ③ 目次の頁番号の注記を現在の版に改める

根拠：`05_試算・管理シート/川崎町_第10期_施策事業統合体系表_R8.9.25b.xlsx`
      03_新規拡充一覧シート

原本の段落・表の要素を deepcopy して文言を差し替えており、新規に生成していない。
"""
import copy
import sys
from collections import OrderedDict

import docx
import openpyxl
from docx.oxml.ns import qn

sys.path.insert(0, "07_ソーススクリプト")
from fix_soan_v111 import clone_table, set_el  # noqa: E402

SRC = "01_第10期_最新版成果品/川崎町_計画書素案_v2.0_施策体系書下ろし版.docx"
DST = "01_第10期_最新版成果品/川崎町_計画書素案_v2.1_新規拡充説明版.docx"
TAIKEI = "05_試算・管理シート/川崎町_第10期_施策事業統合体系表_R8.9.25b.xlsx"

# ══════════════════════════ 交付金の失点27点の回復管理
T_KOFU = [
    ["交付金の評価項目（配点）", "現状", "第10期の目標", "本計画での位置／集計担当"],
    ["評価結果の活用　ア　庁内で共有する場（4点）", "0点（令和7年度は4点）",
     "年1回以上開催", "8-4①／保健福祉課・年1回"],
    ["同　イ　庁外の関係者の参画（4点）", "0点（令和7年度は4点）",
     "策定委員会等に議題として報告", "8-4②／保健福祉課・年1回"],
    ["同　ウ　意見の施策への反映（4点）", "0点（令和7年度は4点）",
     "翌年度の予算編成に反映", "8-4③／保健福祉課・年1回"],
    ["同　エ　評価結果の公表（4点）", "0点（令和7年度は4点）",
     "小項目ごとの得点まで公表", "8-4④／保健福祉課・年1回"],
    ["給付費適正化事業　ア　3事業の全てを実施（6点）", "0点（要確認）",
     "3事業すべてを実施し、該当状況調査で「実施」と回答",
     "8-2／保健福祉課・年1回"],
    ["認知症　カ　計画の策定に着手（5点）", "0点",
     "意見を反映させる仕組みを整えたうえで策定",
     "6-1①／保健福祉課・地域包括支援センター"],
    ["【小計】体制・取組指標群の失点", "27点", "27点の回復", "―"],
    ["後期高齢者数と給付費の伸び率の比較（12点）", "0点",
     "様式を定めて毎年度比較し公表", "8-3③／保健福祉課・年1回"],
    ["生活支援コーディネーターの地域ケア会議への参加割合（4点）", "0点",
     "出席を運用で確保する", "5-4第2節／社会福祉協議会・年1回"],
]

P = {}
P[4] = "計画書素案 Ver.2.1"
P[5] = "（新規・拡充事業の説明版）"
P[21] = (
    "⚠ 頁番号は、本素案Ver.2.1の組版によるものです。"
    "図表の差替え・加筆により頁の割付けが動いた場合は、"
    "目次の頁番号も併せて更新します。")
P[23] = (
    "※本素案Ver.2.1は、令和8年6〜7月に実施した2調査"
    "（一般高齢者576件・要支援要介護認定者142件）の結果、"
    "令和8年9月1日に町からご提供いただいた実績データ、"
    "保険者機能強化推進交付金等の評価結果、"
    "及び令和8年9月2日の第1回策定委員会でのご意見を反映した版です。"
    "施策体系は、町からご提示いただいた「施策・事業の体系（案）」の章立てを基本とし"
    "認知症施策のみを独立章に引き上げるＣ案により、"
    "全7章17節37項143事業で確定しています。"
    "Ver.2.0で143事業を第5章・第6章・第8章に書き下ろし、"
    "本版では、第10期で内容が変わる42事業（新規23・拡充19）について"
    "根拠となる事実と第10期での内容を章ごとに示すとともに、"
    "交付金の失点27点の回復を第10章10-4のKPI管理表に組み込みました。"
    "施策体系は第2回策定委員会（令和8年11月）の選択1のご判断により確定し、"
    "介護給付費準備基金の取崩方針と令和9年度の介護報酬改定を踏まえ、"
    "第3回策定委員会（令和9年1月）で計画書（案）として確定します。")


def load_shinki():
    ws = openpyxl.load_workbook(TAIKEI, data_only=True)["03_新規拡充一覧"]
    by = OrderedDict()
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r[0]:
            continue
        kbn, ch, biz, kon, bikou = r[0], r[1], r[2], r[3] or "", r[4] or ""
        by.setdefault(ch, []).append([kbn, str(biz), str(kon), str(bikou)])
    return by


def main():
    by = load_shinki()
    doc = docx.Document(SRC)
    body = doc.element.body
    ps = [el for el in body.iterchildren() if el.tag == qn("w:p")]
    ts = [el for el in body.iterchildren() if el.tag == qn("w:tbl")]

    for i, txt in P.items():
        set_el(ps[i - 1], txt)

    TPL4 = copy.deepcopy(ts[47 - 1])
    P_SUB = copy.deepcopy(ps[303 - 1])
    P_BODY = copy.deepcopy(ps[301 - 1])

    def find(pred):
        for el in body.iterchildren():
            if el.tag != qn("w:p"):
                continue
            t = "".join(n.text or "" for n in el.iter(qn("w:t"))).strip()
            if pred(t):
                return el
        raise KeyError("見つからない")

    def put(anchor, ch, note=None):
        rows = by[ch]
        n_new = sum(1 for r in rows if r[0] == "新規")
        n_exp = len(rows) - n_new
        h = copy.deepcopy(P_SUB)
        set_el(h, f"　▌ 新たに立てる事業・広げる事業"
                  f"（新規{n_new}件・拡充{n_exp}件）")
        anchor.addprevious(h)
        tbl = [["区分", "事業", "根拠となる事実", "第10期での内容"]]
        tbl += rows
        anchor.addprevious(clone_table(TPL4, tbl, widths=(6, 24, 36, 34)))
        if note:
            e = copy.deepcopy(P_BODY)
            set_el(e, note)
            anchor.addprevious(e)

    C = {
        1: "第１章 健康づくりの推進",
        2: "第２章 高齢者が安心して暮らせるまちづくりの推進",
        3: "第３章 地域生活を支援する取り組みの充実",
        4: "第４章 地域支援事業の充実",
        5: "第５章 認知症施策の推進",
        7: "第７章 地域マネジメントに関する事項",
    }

    # 5-1 → 「成果指標（KPI）」の前
    put(find(lambda t: t == "▌ 成果指標（KPI）"), C[1])
    # 5-2 → 5-2の「成果指標（KPI）」の前（1つ目は上で使われたので次を探す）
    seen = []
    for el in body.iterchildren():
        if el.tag == qn("w:p"):
            t = "".join(n.text or "" for n in el.iter(qn("w:t"))).strip()
            if t == "▌ 成果指標（KPI）":
                seen.append(el)
    put(seen[1], C[2])
    # 5-3 → 「▌ 家族介護者への支援」の前
    put(find(lambda t: t.startswith("▌ 家族介護者への支援")), C[3])
    # 5-4 → 「5-5　介護給付」の前
    put(find(lambda t: t.startswith("5-5　介護給付")), C[4])
    # 6-4 → 「第 7 章」の前
    put(find(lambda t: t == "第 7 章"), C[5])
    # 柱7 → 「8-1　介護サービスの質の向上」の前
    put(find(lambda t: t.startswith("8-1　介護サービスの質の向上")), C[7],
        note="⚠ 柱7の事業は、本章のほか第7章（介護人材・生産性向上）及び"
             "第10章（目標と指標の管理）に分かれて位置付けています。"
             "区分「拡充」は、第9期から引き続き実施するものの、"
             "交付金の評価結果を踏まえて内容を広げる事業です。")

    # ══════════ 10-4に交付金の回復管理を置く
    anchor = find(lambda t: t.startswith(
        "第9期計画のKPIのうち、現時点で現状値が確定しているのは"))
    h = copy.deepcopy(P_SUB)
    set_el(h, "　▌ 交付金の失点の回復管理")
    anchor.addprevious(h)
    e = copy.deepcopy(P_BODY)
    set_el(e, "第3章3-3で特定した交付金の失点のうち、"
              "体制・取組指標群の3か所27点は、"
              "新たな事業も予算も要さず、仕組みを定めて実施することで得点できます。"
              "KPI管理表の体制指標として、次のとおり管理します。"
              "活動指標群のうち町の取組により直接動かせる2項目も併せて掲げます。")
    anchor.addprevious(e)
    anchor.addprevious(clone_table(TPL4, T_KOFU, widths=(30, 16, 26, 28)))
    e2 = copy.deepcopy(P_BODY)
    set_el(e2, "⚠ 「給付費適正化事業 ア」の現状は、3事業のいずれかが未実施なのか、"
               "実施していても国の該当状況調査で「実施」と回答していないのかを"
               "確認中です（確認事項No.127）。"
               "「PFS委託事業数」（12点）と「短期的な要介護度の変化」（40点）は、"
               "3-3に記したとおり本計画では追わない／一部取得を目標とするものであり、"
               "本表には掲げていません。")
    anchor.addprevious(e2)

    doc.save(DST)
    print("保存：", DST)
    d2 = docx.Document(DST)
    print(f"  段落 {len(d2.paragraphs)}／表 {len(d2.tables)}")
    n = 0
    for t in d2.tables:
        h = [c.text.strip() for c in t.rows[0].cells]
        if h[:4] == ["区分", "事業", "根拠となる事実", "第10期での内容"]:
            n += len(t.rows) - 1
    print(f"  新規・拡充事業の説明 {n}件")


if __name__ == "__main__":
    main()
