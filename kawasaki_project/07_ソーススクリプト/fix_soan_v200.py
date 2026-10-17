# -*- coding: utf-8 -*-
"""計画素案 v1.25 → Ver.2.0（統合体系表の143事業を計画書に書き下ろす）

Ver.1.21以降、素案は「施策の内容は変えずに体系上の位置付けのみを確定する」
状態にとどめ、組み替えはVer.2.0で行うこととしていた。本版でこれを行う。

  ① 第5章を、第9期からの6基本目標による5節から、
     Ｃ案の柱（柱1〜柱4・柱6）による5節に組み替える。
     統合体系表の節・項の単位で事業の表を置き、106事業をすべて掲げる。
       5-1 健康づくりの推進（柱1・2節5項17事業）
       5-2 高齢者が安心して暮らせるまちづくりの推進（柱2・2節3項8事業）
       5-3 地域生活を支援する取り組みの充実（柱3・2節4項25事業）
       5-4 地域支援事業の充実（柱4・2節5項30事業）
       5-5 介護給付・介護予防給付サービスの充実（柱6・3節3項26事業）
  ② 第6章に6-4を新設し、柱5の6節8項28事業を掲げる。
  ③ 第8章の章頭に、柱7の9事業と本章の節との対応表を置く。
  ④ 6-1にあったKPIの3層構造の説明（6-3と同一文）を削除する。
  ⑤ 章頭の説明・注記・版の表記を改める。

  ①〜③により、統合体系表の143事業（106＋28＋9）がすべて計画書に載る。

根拠：`05_試算・管理シート/川崎町_第10期_施策事業統合体系表_R8.9.25b.xlsx`

原本の段落・表の要素を deepcopy して文言を差し替えており、新規に生成していない。
"""
import copy
import re
import sys
from collections import OrderedDict

import docx
import openpyxl
from docx.oxml.ns import qn

sys.path.insert(0, "07_ソーススクリプト")
from fix_soan_v111 import clone_table, set_el, set_text  # noqa: E402

SRC = "01_第10期_最新版成果品/川崎町_計画書素案_v1.25_体系数値是正版.docx"
DST = "01_第10期_最新版成果品/川崎町_計画書素案_v2.0_施策体系書下ろし版.docx"
TAIKEI = "05_試算・管理シート/川崎町_第10期_施策事業統合体系表_R8.9.25b.xlsx"


def load_taikei():
    ws = openpyxl.load_workbook(TAIKEI, data_only=True)["01_統合体系表"]
    tree = OrderedDict()
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r[0]:
            continue
        ch, se, ko, biz, kbn, _kon, shokan = r[0], r[1], r[2], r[3], r[4], r[5], r[6]
        tree.setdefault(ch, OrderedDict()).setdefault(
            se or "―", OrderedDict()).setdefault(ko, []).append(
                (biz, kbn, shokan or "町"))
    return tree


def _num(s):
    """「（１）」の番号を取り出す。統合体系表の行の並びが番号順でない箇所が
    2か所あるため（柱4第1節2(7)・柱5第2節1(3)）、表に出すときに並べ替える。"""
    m = re.match(r"（([０-９\d]+)）", str(s))
    if not m:
        return 999
    return int(m.group(1).translate(str.maketrans("０１２３４５６７８９",
                                                  "0123456789")))


def table_rows(sec):
    """節のなかの項・事業を、表の行（項／事業／区分／所管）にする。"""
    rows = [["項", "事業", "区分", "所管"]]
    for ko, items in sec.items():
        for i, (biz, kbn, shokan) in enumerate(sorted(items,
                                                     key=lambda x: _num(x[0]))):
            rows.append([ko if i == 0 else "", biz, kbn, shokan])
    return rows


def count(sec):
    return sum(len(v) for v in sec.values()), len(sec)


P = {}
P[4] = "計画書素案 Ver.2.0"
P[5] = "（施策体系 書下ろし版）"
P[23] = (
    "※本素案Ver.2.0は、令和8年6〜7月に実施した2調査"
    "（一般高齢者576件・要支援要介護認定者142件）の結果、"
    "令和8年9月1日に町からご提供いただいた実績データ、"
    "保険者機能強化推進交付金等の評価結果、"
    "及び令和8年9月2日の第1回策定委員会でのご意見を反映した版です。"
    "施策体系は、町からご提示いただいた「施策・事業の体系（案）」の章立てを基本とし"
    "認知症施策のみを独立章に引き上げるＣ案により、"
    "全7章17節37項143事業で確定しています。"
    "本版では、Ver.1.21以降「体系上の位置付けのみ」としていた第5章を"
    "Ｃ案の柱により組み替え、"
    "統合体系表の143事業を第5章・第6章・第8章にすべて書き下ろしました"
    "（第5章106事業・第6章28事業・第8章9事業）。"
    "施策体系は第2回策定委員会（令和8年11月）の選択1のご判断により確定し、"
    "介護給付費準備基金の取崩方針と令和9年度の介護報酬改定を踏まえ、"
    "第3回策定委員会（令和9年1月）で計画書（案）として確定します。")

P[296] = (
    "本章では、第4章で示した7つの柱のうち、"
    "認知症施策（柱5）と地域マネジメント（柱7）を除く5つの柱の施策を、"
    "施策の方向性・事業の一覧・成果指標（KPI）の3層で整理します。"
    "事業は、別冊「施策・事業の体系（Ｃ案）」の節・項の単位で掲げており、"
    "本章に106事業を収めています。"
    "認知症施策（柱5・28事業）は第6章、"
    "地域マネジメント（柱7・9事業）は第8章に掲げます。"
    "区分の「継続」は第9期から引き続き実施する事業、"
    "「拡充」は内容を広げる事業、「新規」は第10期で新たに立てる事業、"
    "「再掲」は他の節に掲げた事業を参照のために再び掲げたものです。")
P[298] = (
    "⚠ 本表の組み替えは本版（Ver.2.0）で行いました。"
    "第9期の6つの基本方針・本素案Ver.1.25までの6つの基本目標と、"
    "第10期の7つの柱との対応は第4章の表をご参照ください。"
    "体系（Ｃ案）で新たに立てた事業23件"
    "（家族介護者への支援6件・認知症施策13件・町外医療機関との広域連携・"
    "重層的支援体制整備事業との連動・計画値と実績値の乖離のモニタリング・"
    "住所地特例への対応）は、本章・第6章・第8章の該当する節に"
    "「新規」として掲げています。")

# ── 第5章の節の見出しと方向性の書き出し
P[299] = "5-1　健康づくりの推進（柱1）"
P[305] = "5-2　高齢者が安心して暮らせるまちづくりの推進（柱2）"
P[312] = "5-3　地域生活を支援する取り組みの充実（柱3）"
P[326] = "5-4　地域支援事業の充実（柱4）"
P[332] = "5-5　介護給付・介護予防給付サービスの充実（柱6）"

P[328] = (
    "介護予防・日常生活支援総合事業と包括的支援事業・任意事業により、"
    "要支援者と事業対象者への支援、地域包括支援センターの運営、"
    "生活支援体制の整備、家族介護者への支援等を行います。"
    "令和8年度の保険者努力支援交付金では、"
    "目標Ⅰ（介護予防・日常生活支援の推進）が100点満点中87点で全国5位であり、"
    "本町が最も高い評価を得ている分野です。"
    "一方、活動指標群のうち「生活支援コーディネーターの地域ケア会議への"
    "参加割合」は0点であり（配置数は満点）、"
    "会議への出席を運用で確保することを第2節に位置付けます。")
P[329] = (
    "また、交付金で評価されながら第9期計画に記載のなかった"
    "「介護予防ポイント事業」（支援 目標Ⅰ・県内1位）と"
    "「アウトリーチ等の取組」（支援 目標Ⅰ・9点満点・県内1位）を、"
    "第1節の事業として明記します。"
    "とりわけアウトリーチは、通いの場への参加意向55.2％（n=543）に対し"
    "実際の参加率が15.9％（n=439）にとどまるという開きを埋める手段として、"
    "施策上の意味が大きいものです。")
P[330] = (
    "⚠ 地域支援事業の事業別の実績（費用・実施回数・対象者数）は"
    "町からのご提供を待っています。"
    "上限額の管理と事業費の見込みは第9章9-2に整理しています。")

P[334] = (
    "介護給付・介護予防給付のサービスは、"
    "居宅サービス14種別・地域密着型サービス9種別・施設サービス3種別です。"
    "本節では体系上の位置付けを示し、"
    "種別ごとの見込量・給付費及び必要利用定員総数は第9章9-1に定めます。")
P[335] = (
    "本町の給付構造は、令和7年度で施設サービスが給付費の49.4％、"
    "地域密着型が18.3％、居宅サービスが32.3％です。"
    "入所・居住系（施設・地域密着型・特定施設）が680,226千円・68.1％を占め、"
    "1人1月あたり給付費は約29.9万円で、"
    "居宅サービス（特定施設を除く）の約10.2万円の約2.9倍です。"
    "入所・居住系の利用者が1人増減すると保険料基準額は月額約21円動きます。")
P[336] = (
    "⚠ 在宅系の地域密着型サービス（定期巡回・随時対応型訪問介護看護、"
    "夜間対応型訪問介護、看護小規模多機能型居宅介護）は、"
    "在宅介護実態調査での利用がほとんど確認できませんでした。"
    "町内及び近隣市町村に指定事業所が存在するかを確認できておらず"
    "［要確認］、存在しない場合、利用が0件であることは供給側の事実であり、"
    "利用勧奨では対応できません。第2回策定委員会の選択3でお諮りしています。")


def main():
    tree = load_taikei()
    doc = docx.Document(SRC)
    body = doc.element.body

    # ── 位置の索引（段落番号・表番号 → 要素）
    ps, ts = [], []
    for el in body.iterchildren():
        if el.tag == qn("w:p"):
            ps.append(el)
        elif el.tag == qn("w:tbl"):
            ts.append(el)

    def P_(n):
        return ps[n - 1]

    def T_(n):
        return ts[n - 1]

    for i, txt in P.items():
        set_el(P_(i), txt)

    TPL4 = copy.deepcopy(T_(47))       # 4列の事業表
    TPL3 = copy.deepcopy(T_(46))       # 3列の対応表
    P_SUB = copy.deepcopy(P_(303))     # 「　▌ …」の小見出し
    P_BODY = copy.deepcopy(P_(301))    # 本文

    def sub(anchor, text):
        e = copy.deepcopy(P_SUB)
        set_el(e, text)
        anchor.addprevious(e)

    def bodyp(anchor, text):
        e = copy.deepcopy(P_BODY)
        set_el(e, text)
        anchor.addprevious(e)

    def swap(old_tbl, rows, widths=(22, 46, 8, 24)):
        """既存の表を、新しい表に差し替える。"""
        new = clone_table(TPL4, rows, widths=widths)
        old_tbl.addprevious(new)
        old_tbl.getparent().remove(old_tbl)
        return new

    def put_sec(anchor, ch, se, label):
        n, k = count(tree[ch][se])
        sub(anchor, f"　▌ {label}（{k}項{n}事業）")
        anchor.addprevious(clone_table(TPL4, table_rows(tree[ch][se]),
                                       widths=(22, 46, 8, 24)))

    C1 = "第１章 健康づくりの推進"
    C2 = "第２章 高齢者が安心して暮らせるまちづくりの推進"
    C3 = "第３章 地域生活を支援する取り組みの充実"
    C4 = "第４章 地域支援事業の充実"
    C5 = "第５章 認知症施策の推進"
    C6 = "第６章 介護給付・介護予防給付サービスの充実"
    C7 = "第７章 地域マネジメントに関する事項"

    # ══════════ 5-1 柱1（第1節9事業・第2節8事業）
    se1, se2 = list(tree[C1].keys())
    set_el(P_(303), f"　▌ {se1}（{count(tree[C1][se1])[1]}項"
                    f"{count(tree[C1][se1])[0]}事業）")
    swap(T_(47), table_rows(tree[C1][se1]))
    put_sec(P_(304), C1, se2, se2)           # 成果指標の見出しの前に挿入

    # ══════════ 5-2 柱2
    se1, se2 = list(tree[C2].keys())
    set_el(P_(310), f"　▌ {se1}（{count(tree[C2][se1])[1]}項"
                    f"{count(tree[C2][se1])[0]}事業）")
    swap(T_(49), table_rows(tree[C2][se1]))
    put_sec(P_(311), C2, se2, se2)

    # ══════════ 5-3 柱3（第1節19事業・第2節6事業）
    se1, se2 = list(tree[C3].keys())
    set_el(P_(317), f"　▌ {se1}（{count(tree[C3][se1])[1]}項"
                    f"{count(tree[C3][se1])[0]}事業）")
    swap(T_(51), table_rows(tree[C3][se1]))
    set_el(P_(318), "　▌ 家族介護者への支援（第1節3・新設）")
    set_el(P_(323), f"　▌ {se2}（{count(tree[C3][se2])[1]}項"
                    f"{count(tree[C3][se2])[0]}事業）")
    swap(T_(52), table_rows(tree[C3][se2]))

    # ══════════ 5-4 柱4
    se1, se2 = list(tree[C4].keys())
    set_el(P_(331), f"　▌ {se1}（{count(tree[C4][se1])[1]}項"
                    f"{count(tree[C4][se1])[0]}事業）")
    swap(T_(54), table_rows(tree[C4][se1]))
    put_sec(P_(332), C4, se2, se2)           # 5-5の見出しの前に挿入

    # ══════════ 5-5 柱6（3節）
    se1, se2, se3 = list(tree[C6].keys())
    set_el(P_(337), f"　▌ {se1}（{count(tree[C6][se1])[0]}事業）")
    swap(T_(55), table_rows(tree[C6][se1]))
    for se in (se2, se3):
        n, _k = count(tree[C6][se])
        sub(P_(338), f"　▌ {se}（{n}事業）")
        P_(338).addprevious(clone_table(TPL4, table_rows(tree[C6][se]),
                                        widths=(22, 46, 8, 24)))

    # ══════════ 第6章　6-4を新設（柱5の6節8項28事業）
    anchor6 = P_(368)                        # 「第 7 章」の前に入れる
    sub_h = copy.deepcopy(P_(360))           # 「6-3　重点施策とKPI」の見出し
    h = copy.deepcopy(sub_h)
    set_el(h, "6-4　認知症施策の体系（6節8項28事業）")
    anchor6.addprevious(h)
    bodyp(anchor6,
          "本章の施策を、別冊「施策・事業の体系（Ｃ案）」の柱5"
          "（認知症施策の推進）の節・項の単位で掲げます。"
          "認知症基本法第14条から第21条までの8つの基本的施策のすべてに"
          "対応する事業を置いたうえで、"
          "第9条（認知症の日・認知症月間）及び第13条（計画の策定手続）への"
          "対応を加えており、6節8項28事業（ほかに再掲4）で構成されます。")
    for se in tree[C5]:
        n, k = count(tree[C5][se])
        sub(anchor6, f"　▌ {se}（{k}項{n}事業）")
        anchor6.addprevious(clone_table(TPL4, table_rows(tree[C5][se]),
                                        widths=(24, 44, 8, 24)))
    bodyp(anchor6,
          "⚠ 「再掲」は、他の柱に掲げた事業のうち"
          "認知症施策としても位置付けるものです。"
          "事業数（28事業）には再掲4件を含めていません。")

    # ══════════ 第8章　章頭に柱7の対応表
    anchor8 = P_(417)                        # 「8-1　介護サービスの質の向上」
    bodyp(anchor8,
          "別冊「施策・事業の体系（Ｃ案）」の柱7（地域マネジメントに関する事項）"
          "の9事業と、本章及び第7章の節との対応は次のとおりです。")
    rows = [["柱7の事業", "区分", "本計画での位置"]]
    IDX = {
        "１　介護サービスの質の向上": "8-1",
        "２　介護費用適正化の推進": "8-2",
        "３　助言・指導の充実": "8-1",
        "４　情報提供・相談体制の充実": "8-5",
        "５　福祉・介護人材の確保・定着支援【重点】": "第7章7-4",
        "６　介護現場の生産性向上（介護ロボット・ＩＣＴの導入支援）": "第7章7-5",
        "７　自立支援・重度化防止に向けた目標と指標": "第10章10-4",
        "８　計画値と実績値の乖離のモニタリングと公表": "8-3",
        "９　町外施設利用・住所地特例への対応": "第2章2-3・第9章9-1",
    }
    for ko, items in tree[C7]["―"].items():
        for biz, kbn, _sh in items:
            rows.append([f"{ko}　{biz}", kbn, IDX.get(ko, "―")])
    anchor8.addprevious(clone_table(TPL3, rows, widths=(56, 10, 34)))

    # ══════════ 6-1のKPI3層の説明（6-3と同一文）を削除する
    dup = P_(353)
    if "KPIは、プロセス（活動量）" in "".join(
            n.text or "" for n in dup.iter(qn("w:t"))):
        dup.getparent().remove(dup)
        print("  6-1の重複段落（KPIの3層構造・6-3と同一文）を削除しました。")

    # ══════════ 目次の節名を改める（頁番号は fill_toc_pages で入れ直す）
    TOCMAP = {
        "5-1．健康づくり・介護予防の推進（基本目標1）":
            "5-1．健康づくりの推進（柱1）",
        "5-2．高齢者が安心して暮らせるまちづくり（基本目標2）":
            "5-2．高齢者が安心して暮らせるまちづくりの推進（柱2）",
        "5-3．在宅生活継続の支援（基本目標3）":
            "5-3．地域生活を支援する取り組みの充実（柱3）",
        "5-4．介護サービスの質の確保と人材確保（基本目標4）":
            "5-4．地域支援事業の充実（柱4）",
        "5-5．地域包括ケアシステムの深化（基本目標5）":
            "5-5．介護給付・介護予防給付サービスの充実（柱6）",
    }
    from fix_soan_v111 import set_tc          # noqa: PLC0415
    done = 0
    add6 = []
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                key = cell.text.strip()
                if key in TOCMAP:
                    set_tc(cell._tc, TOCMAP[key])
                    done += 1
                if key == "6-3．重点施策とKPI":
                    add6.append((t, row))
    for t, row in add6:
        new_tr = copy.deepcopy(row._tr)
        row._tr.addnext(new_tr)
        import docx.table as _dt              # noqa: PLC0415
        nr = _dt._Row(new_tr, t)
        set_tc(nr.cells[1]._tc, "6-4．認知症施策の体系（6節8項28事業）")
        set_tc(nr.cells[2]._tc, "―")
        done += 1
    print(f"  目次の節名を{done}件改めました。")

    doc.save(DST)
    print("保存：", DST)
    d2 = docx.Document(DST)
    print(f"  段落 {len(d2.paragraphs)}／表 {len(d2.tables)}")
    # 書き下ろした事業数を数える
    n = 0
    for t in d2.tables:
        h = [c.text.strip() for c in t.rows[0].cells]
        if h[:4] == ["項", "事業", "区分", "所管"]:
            n += len(t.rows) - 1
    print(f"  事業の表に掲げた事業 {n}件（＋柱7の9件＝{n + 9}件）")


if __name__ == "__main__":
    main()
