# -*- coding: utf-8 -*-
"""計画素案 Ver.2.11（見込量が0のサービスの是正と、施策との関係の追記）

08_作業順位 の順位22。Ver.2.10 に当てて Ver.2.11 を作る。

  入力　01_第10期_最新版成果品/川崎町_計画書素案_v2.10_制度改正補完版.docx
  出力　01_第10期_最新版成果品/川崎町_計画書素案_v2.11_見込量是正版.docx

是正1　9-1「見込量を0としたサービスの区分」の表
  Ver.2.10 は5件すべてを「実績あり → 0」としているが、
  **5件とも令和6年度の実績も0である。**
    ・将来推計総括表（R8.9.11出力・④訂正後）の令和6・7・8年度はいずれも0.0
    ・町提供実績データ（R8.9.1受領）13・14シートは、夜間対応型訪問介護を除き
      見込量・実績とも空欄
    ・夜間対応型訪問介護は全年度同値の仮置き（延利用者数500人／給付費8,000千円）
      が入っているのみで、確認事項No.41で0でのご記入をお願いしている
  「実績がない」と「実績があったが消えた」では第10期の扱いの筋道が変わるため、
  欄の見出しごと改める。

是正2　9-1 に「施策と見込量の関係」の小見出しを新設
  見込量の算定式で動く変数は**要介護度別の認定者数1つだけ**であり、
  利用率・1人1月あたりの回（日）数・単価は令和7年度で据え置いている。
  この区別を示さずに「施策を進めるので見込量を増やす／減らす」とすると、
  給付費と保険料が根拠なく動く。**当方の算定は自然体である**ことを本文に置く。

  詳細は 05_試算・管理シート/川崎町_認知症施策_位置づけと見込量_R8.10.6.xlsx。

⚠ 数値は一切変えていない。見込量・給付費・保険料はVer.2.10のままである。
"""
import copy
import os
import re
import shutil
import sys
import zipfile

import docx
from docx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from fix_soan_v111 import clone_table, set_el, set_tc        # noqa: E402

SRC = "01_第10期_最新版成果品/川崎町_計画書素案_v2.10_制度改正補完版.docx"
DST = "01_第10期_最新版成果品/川崎町_計画書素案_v2.11_見込量是正版.docx"

# ══════════════════════════ 是正1　0のサービスの表
T_ZERO = [
    ["サービス種類", "給付実績（令和6〜8年度）", "0の区分",
     "第10期の扱い", "確認事項"],
    ["夜間対応型訪問介護", "いずれの年度も実績なし",
     "［要確認］供給なし又は利用なし",
     "見込量0。⚠ 町提供実績データには全年度同値の仮置き"
     "（延利用者数500人・給付費8,000千円）が入っているため、"
     "0でのご記入をお願いしている。"
     "⚠ 令和8年法律第51号により定期巡回・随時対応型訪問介護看護へ"
     "統合される（1-5）",
     "No.41・No.151"],
    ["認知症対応型通所介護（介護予防を含む）", "いずれの年度も実績なし",
     "供給なし（町内に事業所がなく、町外の事業所による給付もない）",
     "見込量0。認知症施策としては第5章5-5に継続の位置付けとし、"
     "必要となった場合は近隣市町村の事業所の利用を支援する",
     "No.29"],
    ["小規模多機能型居宅介護（介護予防を含む）", "いずれの年度も実績なし",
     "［要確認］町内に事業所があるとの整理があり、実績との食い違いを確認中",
     "見込量0。第2章2-7の事業所一覧と突き合わせる",
     "No.29"],
    ["地域密着型特定施設入居者生活介護", "いずれの年度も実績なし",
     "供給なし（町内の定員0人。必要利用定員総数の表に記載）",
     "見込量0。必要利用定員総数も0とする", "No.29"],
    ["介護医療院", "いずれの年度も実績なし",
     "［要確認］供給なし又は利用なし",
     "見込量0。介護療養型医療施設は令和6年3月末で廃止済み", "No.29"],
]

ZERO_NOTE_OLD = "⚠ 令和6年度に実績があり令和7年度に実績がなかったサービス"
ZERO_NOTE_NEW = (
    "⚠ 上表の5つのサービスは、令和6年度・令和7年度・令和8年度のいずれにも"
    "給付の実績がありません。出所は地域包括ケア「見える化」システムの"
    "将来推計総括表（令和8年9月11日出力・訂正後）及び町提供実績データ"
    "（令和8年9月1日受領）です。実績がないため、第10期の見込量は0と"
    "しています。整備により供給を設ける場合は見込量が立ちますので、"
    "第10期中の整備の予定を策定委員会で確認します。")

# ══════════════════════════ 是正3　受託者を主語とする語の削除
#
# ⚠ 計画は川崎町が定めるものである。本文にも見出しにも受託者を主語とする語
#   （受託者・当社・ビズアップ・確認事項No・スクリプト等）を置かない
#   （点検スキル 3「守るべき制約の走査」）。Ver.2.10 には2か所残っていた。
#   いずれも当方の作業の事情であり、計画書に書くべきものではない。
JUTAKU = [
    ("当社の作業環境のネットワーク制限により",
     "⚠ 本節の制度改正の内容は、令和8年6月以降に国が示した資料に"
     "よるものです。原典（『基本指針の構成について』"
     "第135回社会保障審議会介護保険部会 資料2-1・令和8年6月29日、"
     "『社会福祉法等の一部を改正する法律の概要について』"
     "社会保障審議会 生活困窮者自立支援及び生活保護部会（第30回）"
     "資料2・令和8年7月17日、介護保険最新情報Vol.1525・"
     "令和8年7月14日）との照合を行ったうえで、本計画の確定時に"
     "記述を確定します。"),
    ("保険料試算の精緻化（受託者＋町担当課）",
     "④　令和8年12月〜令和9年1月：保険料試算の精緻化"),
]

# ══════════════════════════ 是正4　未確定箇所の括弧
#
# ⚠ 未確定箇所は［　］、町のご確認を待つものは【町確認】、
#   委員会のご協議を待つものは【委員会協議】の3つに限る定めである。
#   Ver.2.10 には【町確認・最新値】という4つめの表記が1か所あった。
KAKKO = [("【町確認・最新値】", "【町確認】")]

# ══════════════════════════ 是正2　施策と見込量の関係
SEISAKU = [
    ("見出し", "▌ 施策と見込量の関係"),
    ("本文",
     "本節の見込量は、令和7年度の実績に要介護度別の認定者数の伸びを"
     "乗じて算定しています。この算定式で年度により動く要素は"
     "「要介護度別の認定者数」の一つだけであり、サービスの利用率、"
     "1人1月あたりの利用回（日）数、1人1月あたりの給付費（単価）は、"
     "いずれも令和7年度の水準で据え置いています。"),
    ("本文",
     "このため、本計画に掲げる施策の目標（KPI）が見込量に及ぶかどうかは、"
     "その目標がどの要素に効くかによって分かれます。"),
    ("T3", [
        ["目標（KPI）が効く要素", "見込量への反映", "本計画の目標の例"],
        ["要介護度別の認定者数",
         "反映できる。目標値を置くと見込量・給付費・保険料が動く",
         "調整済み認定率（要介護2以上）、要介護1・2の平均要介護度の伸び"],
        ["認定者数に間接的に効くもの",
         "⚠ 効く向きは説明できるが、目標値と認定者数を結ぶ係数が"
         "本町にないため、数値としては反映していない",
         "通いの場への参加率、外出を控えている高齢者の割合、"
         "自立支援型地域ケア会議での個別事例の協議件数、"
         "介護予防事業の参加者数"],
        ["サービスの利用率・利用回（日）数・単価",
         "反映できない。算定式で据え置いており、変数として存在しない",
         "―（反映するには算定の方法そのものを改める必要があります）"],
        ["供給体制・体制整備",
         "量としては効かない。見込量を満たせるかどうか"
         "（供給の実現性）の確認に用いる",
         "生活支援コーディネーター数、事業所ヒアリングの実施、"
         "計画値と実績値の乖離のモニタリングと公表"],
    ], [26, 34, 40]),
    ("本文",
     "以上により、本計画の見込量は、施策の達成を数値として織り込まない"
     "自然体の推計としています。認知症施策についても同様です。"
     "認知症施策を進めると、早期発見・早期対応により認知症対応型サービスの"
     "利用が早く始まる向きと、地域での支えが充実して在宅での生活が続く"
     "向きの双方が働くため、量としての向きが定まらないためです。"),
    ("囲み",
     "⚠ 施策の達成を前提に見込量を動かす場合は、動かす向きと大きさの根拠を"
     "示す必要があります。第9期は、総給付費で計画が実績を1.3％上回り、"
     "居宅サービスは8.8％下回る一方、施設サービスは3.3％上回りました"
     "（3-3・図3-2）。計画が実績を上回る側に偏っているため、"
     "施策の達成を前提に見込量をさらに下げると、給付費と保険料を"
     "過小に見積もることになります［要協議］。"),
]


def main():
    if not os.path.exists(SRC):
        raise SystemExit("入力がない：" + SRC)
    shutil.copy(SRC, DST)
    doc = docx.Document(DST)
    body = doc.element.body
    kids = list(body.iterchildren())

    # 雛形を採る
    tpl = {}
    for el in kids:
        if el.tag == qn("w:p"):
            if el.findall(".//" + qn("w:drawing")):
                continue
            s = "".join(n.text or "" for n in el.iter(qn("w:t"))).strip()
            if "見出し" not in tpl and s.startswith("▌"):
                tpl["見出し"] = copy.deepcopy(el)
            if "本文" not in tpl and s.startswith("令和7年度から令和8年度への"):
                tpl["本文"] = copy.deepcopy(el)
            if "囲み" not in tpl and s.startswith("⚠ 宮城県自身の評価"):
                tpl["囲み"] = copy.deepcopy(el)
        elif el.tag == qn("w:tbl"):
            # ⚠ 列数だけで雛形を選ぶと、**目次の表**（3列・10pt）を
            #   拾ってしまう。用途の合う本文の表を名指しで採る
            #   （点検スキル 8「雛形として複製する段落・表を、
            #   用途の合うものから採る」）。
            tb = docx.table.Table(el, doc)
            head = [c.text.replace("\n", "").strip()
                    for c in tb.rows[0].cells]
            if head == ["要素", "採用データ・方針", "根拠・出典"]:
                tpl.setdefault("T3", copy.deepcopy(el))
            if head[:2] == ["サービス種類", "実績（令和6年度→令和7年度）"]:
                tpl.setdefault("T5", copy.deepcopy(el))
    for k in ("見出し", "本文", "囲み", "T3", "T5"):
        if k not in tpl:
            raise SystemExit("雛形が見つからない：" + k)
    # 雛形の表の文字の大きさが本文の表と同じであること（10ptを拾っていないこと）
    for k in ("T3", "T5"):
        szs = {int(n.get(qn("w:val"))) / 2
               for n in tpl[k].iter(qn("w:sz"))}
        if szs - {8.5}:
            raise SystemExit(f"雛形 {k} に本文の表でない大きさがある：{szs}")

    # ── 是正1　0のサービスの表を差し替える
    n_zero = 0
    for el in kids:
        if el.tag != qn("w:tbl"):
            continue
        tb = docx.table.Table(el, doc)
        head = [c.text.replace("\n", "").strip() for c in tb.rows[0].cells]
        if head[:2] == ["サービス種類", "実績（令和6年度→令和7年度）"]:
            new = clone_table(tpl["T5"], T_ZERO, [26, 20, 24, 44, 12])
            el.addnext(new)
            body.remove(el)
            n_zero += 1
    if n_zero != 1:
        raise SystemExit(f"0のサービスの表が{n_zero}件（1件のはず）")

    # ── 是正1b　注記を差し替える
    n_note = 0
    for el in body.iterchildren():
        if el.tag != qn("w:p"):
            continue
        s = "".join(n.text or "" for n in el.iter(qn("w:t"))).strip()
        if s.startswith(ZERO_NOTE_OLD):
            set_el(el, ZERO_NOTE_NEW)
            n_note += 1
    if n_note != 1:
        raise SystemExit(f"0のサービスの注記が{n_note}件（1件のはず）")

    # ── 是正2　「施策と見込量の関係」を 9-1 の末尾（地域支援事業の前）に置く
    anchor = None
    for el in body.iterchildren():
        if el.tag != qn("w:p"):
            continue
        s = "".join(n.text or "" for n in el.iter(qn("w:t"))).strip()
        if s.startswith("▌ 地域支援事業の量の見込み"):
            anchor = el
            break
    if anchor is None:
        raise SystemExit("差し込む場所（▌ 地域支援事業の量の見込み）がない")
    prev = None
    for item in SEISAKU:
        kind, v = item[0], item[1]
        if kind.startswith("T"):
            el = clone_table(tpl[kind], v, item[2])
        else:
            el = copy.deepcopy(tpl[kind])
            set_el(el, v)
        if prev is None:
            anchor.addprevious(el)
        else:
            prev.addnext(el)
        prev = el
    # 表のあとに空の段落（表が続くと Word で見分けづらい）
    e = copy.deepcopy(tpl["本文"])
    set_el(e, "")
    prev.addnext(e)

    # ── 是正3　受託者を主語とする語を取り除く
    n_ju = 0
    for el in body.iterchildren():
        if el.tag != qn("w:p"):
            continue
        s = "".join(n.text or "" for n in el.iter(qn("w:t")))
        for sagasu, atarashii in JUTAKU:
            if sagasu in s:
                set_el(el, atarashii)
                n_ju += 1
    if n_ju != len(JUTAKU):
        raise SystemExit(f"受託者を主語とする語の差し替えが{n_ju}件"
                         f"（{len(JUTAKU)}件のはず）")

    # ── 是正4　未確定箇所の括弧を定めの3つに揃える
    n_kak = 0
    for el in body.iter(qn("w:t")):
        for furui, atarashii in KAKKO:
            if el.text and furui in el.text:
                el.text = el.text.replace(furui, atarashii)
                n_kak += 1
    if n_kak != len(KAKKO):
        raise SystemExit(f"未確定箇所の括弧の差し替えが{n_kak}件"
                         f"（{len(KAKKO)}件のはず）")

    # ── 版表記（本文）を Ver.2.11 に改める
    #    ⚠ ヘッダーだけ直すと check_font の検査8（本文の版とヘッダーの版）に
    #      落ちる。本文・表の双方を見る。
    n_ver = 0
    for el in body.iter(qn("w:t")):
        if el.text and "Ver.2.10" in el.text:
            el.text = el.text.replace("Ver.2.10", "Ver.2.11")
            n_ver += 1
    if n_ver == 0:
        raise SystemExit("本文に Ver.2.10 の表記が見つからない")

    doc.save(DST)
    naoshi_header(DST)

    # ══════════════════════════ 自己点検
    d2 = docx.Document(DST)
    d0 = docx.Document(SRC)
    text = "\n".join(q.text for q in d2.paragraphs)
    tbl_text = "\n".join(c.text for t in d2.tables
                         for r in t.rows for c in r.cells)
    zen = text + "\n" + tbl_text
    ng = []
    # 誤りが残っていないこと
    if "実績あり → 0" in zen:
        ng.append("「実績あり → 0」が残っている")
    if "令和6年度に実績があり令和7年度に実績がなかった" in zen:
        ng.append("古い注記が残っている")
    # 加えたものが入っていること
    for w in ("▌ 施策と見込量の関係", "自然体の推計", "反映できない。算定式で据え置",
              "いずれの年度も実績なし"):
        if w not in zen:
            ng.append(f"加えるはずの記述がない：{w}")
    # 数値を変えていないこと（見込量の表の数が同じであること）
    import collections
    def nums(d):
        s = "\n".join(c.text for t in d.tables for r in t.rows for c in r.cells)
        return collections.Counter(re.findall(r"\d[\d,]*\.?\d*", s))
    n0, n2 = nums(d0), nums(d2)
    # 差は本バージョンで加えた表と差し替えた表の分のみ
    kieta = [k for k in n0 if n0[k] > n2.get(k, 0)
             and k in ("80,059", "81,659", "24.5", "25.0", "998,220")]
    if kieta:
        ng.append("見込量の数値が減っている：" + "／".join(kieta))
    # 強調記号・他団体
    for w in ("**", "__"):
        if w in zen:
            ng.append(f"強調記号が残っている：{w}")
    for w in ("大雪", "東川", "東神楽", "上川", "美瑛", "金ヶ崎", "川崎市"):
        if w in zen:
            ng.append(f"他団体の名称が入っている：{w}")
    # 受託者を主語とする語
    #   ⚠ 奥付の「（策定支援：…）」は、計画書に策定支援者を表示する
    #     慣行によるものであり、本文の記述ではない。残すかどうかは
    #     町のご判断による（確認事項No.167）。ここでは本文のみを見る。
    #   ⚠ 「確認事項No.」の参照は、協議用の版では**意図して残している**。
    #     公表版では取り除く（確認事項No.162・公表版への整理 R8.10.2）。
    #     ここでは数が増えていないことだけを見る。
    for w in ("受託者", "当社", "スクリプト"):
        if w in zen:
            ng.append(f"受託者を主語とする語が本文に入っている：{w}")
    n_ref0 = len(re.findall(r"確認事項No", honbun_zen(d0)))
    n_ref2 = len(re.findall(r"確認事項No", zen))
    if n_ref2 > n_ref0:
        ng.append(f"確認事項No.の参照が増えている（{n_ref0}→{n_ref2}）")
    okuzuke = [q.text for q in d2.paragraphs if "策定支援：" in q.text]
    if len(okuzuke) != 1:
        ng.append(f"奥付の策定支援の表示が{len(okuzuke)}件（1件のはず）")
    # 版表記
    z = zipfile.ZipFile(DST)
    hdr = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>",
                             z.read("word/header1.xml").decode("utf-8")))
    if "Ver.2.11" not in hdr:
        ng.append(f"ヘッダーの版表記が更新されていない：{hdr}")

    print("計画素案 Ver.2.11（見込量是正版）")
    print("保存：", DST)
    print(f"  段落 {len(d2.paragraphs)}（Ver.2.10 は {len(d0.paragraphs)}）"
          f"／表 {len(d2.tables)}（同 {len(d0.tables)}）")
    print("  ── 自己点検")
    if ng:
        for m in ng:
            print("   ×", m)
        sys.exit(1)
    print("   ○ 「実績あり → 0」を是正した（5件とも令和6〜8年度に実績なし）")
    print("   ○ 「施策と見込量の関係」を 9-1 に新設した")
    print("   ○ 見込量・給付費・保険料の数値は変えていない")
    print("  ⚠ 目次の頁番号は fill_toc_pages_v200.py で作り直してください。")


def honbun_zen(d):
    """本文と表の文字をひとつにする。"""
    return ("\n".join(q.text for q in d.paragraphs) + "\n"
            + "\n".join(c.text for t in d.tables
                        for r in t.rows for c in r.cells))


def naoshi_header(path):
    z = zipfile.ZipFile(path)
    names = z.namelist()
    data = {n: z.read(n) for n in names}
    z.close()
    n_fix = 0
    for n in names:
        if not re.match(r"word/(header|footer)\d*\.xml", n):
            continue
        s = data[n].decode("utf-8")
        s2 = re.sub(r"Ver\.2\.10", "Ver.2.11", s)
        if s2 != s:
            data[n] = s2.encode("utf-8")
            n_fix += 1
    if n_fix == 0:
        raise SystemExit("ヘッダーの版表記（Ver.2.10）が見つからない")
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, data[n])
    shutil.move(tmp, path)


if __name__ == "__main__":
    main()
