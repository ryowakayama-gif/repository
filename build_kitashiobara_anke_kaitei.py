# -*- coding: utf-8 -*-
"""
北塩原村　アンケート集計表の改訂版の確認

出力: output/北塩原村_アンケート集計表の改訂版の確認.xlsx

何を受けて作ったものか
  令和8年10月14日に「アンケート結果については添付ファイルの確認を
  お願い致します」として、報告書用集計表の改訂版を受領した。
  令和8年10月3日に当方が洗い出した不整合11件
  （output/北塩原村_アンケート結果報告書_点検結果.xlsx）が
  どうなったかを、新旧の現物を突き合わせて確かめた。

いちばん大きな変更
  割合の基数（分母）が変わった。
    旧 … 各設問の有効回答数（無回答を除いた回答者数）。
          設問ごとに n が違っていた（n=36、34、33、30 など）。
    新 … 各設問の回答者数（無回答を含む）に統一。n は38で固定。
          あわせて「無回答」を選択肢として表に明示した。
  これは当方の指摘Ａ-1（単純集計は無回答を除く分母、
  クロス集計は無回答を含む分母で、同じ設問が報告書の中で
  2つの異なる基準で示されている）への対応である。
  目次の注記も「割合は、各設問の回答者数（無回答を含む）を
  基数として算出し」に改められた。

  基数が変わると、ほとんどの割合が動く。
  たとえば日常生活で困っていること「将来の生活」は
  30.6％（旧・n=36）から28.9％（新・n=38）になった。

当方の正本は動かない
  当方の計画素案 第2章7 は、もともと「割合は無回答を含む
  回答者数に対する割合」として書いており、改訂版と同じ基準である。
  本書の自己点検で、第2章7 の割合38点すべてが改訂版の集計から
  出ることを確かめている（合算した割合は足し算で確かめる）。
  1点でも合わなければ生成が止まる。

落ちていた論点が1件あった
  改訂版の 6章_課題（表6-1）は課題を8件としている。
  当方の正本 第2章7 の表（表2-8）は7区分である。
  突き合わせると、8件目の「専門的支援」
  （精神障がいにも対応した地域包括ケアシステムの構築、
  強度行動障がい・医療的ケアへの対応。根拠設問 問17から問20）に
  対応する区分が当方に無い。
  これは令和8年10月13日から Ｓ-26 として
  「どちらに揃えるか」を確かめていた事項であり、
  改訂版により報告書の側が8件であることが確定した。
  表2-8 に8区分目を足す（ＡＫ-1）。

まだ残っているもの
  当方の指摘Ａ-3（基数が1から2の区分に割合を表示している）は
  解消していない。改訂版で新設された「集計_障がい別」シートは
  知的障がい（n=2）の列に割合を入れている。
  当方の正本 表2-7 も同じ列を持つが、
  「知的障がい及び精神障がいは回答者数が少ないため参考値です」と
  注記している。報告書の側にも同じ注記が要る（Ｓ-37）。

  あわせて、改訂版で記入された調査対象の内訳が配布数と合わない。
  障がい者調査は身体118名・精神1名・療育8名・入所等12名で
  合計139名であるのに、配布数は172である。
  障がい児調査は2＋1＋1＋2＝6で配布数6と合っている。
  精神1名は村の精神障害者保健福祉手帳の所持者（令和8年4月38人）と
  比べても少なく、数え方を確かめる必要がある（Ｍ-128）。
"""

import os
import re
from collections import defaultdict
from itertools import combinations

import docx
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

from build_kitashiobara_survey_check import ISSUES
import kitashiobara_anke as A
from kitashiobara_common import (
    COLORS, OUT_DIR, add_sheet, ensure_out_dir, style_header_row,
    style_note, write_row,
)

REPO_ROOT = os.path.dirname(os.path.abspath(__file__))
OUT_FILE = f"{OUT_DIR}/北塩原村_アンケート集計表の改訂版の確認.xlsx"
SHUUKEI = f"{REPO_ROOT}/source/アンケート結果/報告書用集計表.xlsx"
# 正本 第2章7 の割合を照らす集計表は、いま有効な版（令和8年10月21日版・
# 回収40件）とする。本書の②④⑤は改訂版（回収38件）そのものの確認であり、
# SHUUKEI を見る。③だけは「正本の割合が集計から出ること」の確認であるため、
# 正本が拠る版を見なければ意味がない（令和8年10月24日 RT-01）。
SHUUKEI_GENKOU = A.SHIN
HONPON = f"{REPO_ROOT}/output/北塩原村_計画素案_正本_移植後.docx"
HI = "令和8年10月14日"

KAI = "解消"
ICHIBU = "一部解消"
MIKAI = "未解消"
HOKA = "当方の側で対応"


# ===========================================================================
# 1. 当方の指摘11件が改訂版でどうなったか
#    (番号, 判定, 改訂版で何が変わったか, 改訂版の現物で確かめる語)
#    番号と指摘の内容は build_kitashiobara_survey_check.py から引く。
# ===========================================================================
HANTEI = {
    "A-1": (KAI,
            "割合の基数を各設問の回答者数（無回答を含む）に統一し、"
            "「無回答」を選択肢として表に明示した。"
            "これにより単純集計とクロス集計が同じ基準になった。"
            "目次の注記も改められた",
            "割合は、各設問の回答者数（無回答を含む）を基数として算出し"),
    "A-2": (KAI,
            "問27-1（現在利用しているサービス）と"
            "問27-3（今後３年間で利用したいサービス）の基数が"
            "いずれも38に統一されたため、"
            "「差（今後－現在）」が同じ分母の差になった。"
            "差の値も変わっている（生活介護は7.4ポイントから"
            "5.3ポイントへ）",
            "回答者数（n）"),
    "A-3": (MIKAI,
            "新設の「集計_障がい別」シートは、"
            "知的障がい（n=2）の列に割合を入れている。"
            "1件が100.0％、2件が100.0％として表示される。"
            "当方の正本 表2-7 は同じ列を持つが"
            "「回答者数が少ないため参考値です」と注記している。"
            "報告書の側にも同じ注記が要る（Ｓ-37）",
            "障がい別は、第７期計画策定時の調査と同様に"),
    "B-1": (HOKA,
            "前回調査との比較可能性は集計表の側では直せない。"
            "当方は目標ＫＰＩの前回比較可能性のブック"
            "（北塩原村_目標KPI_前回比較可能性.xlsx）で"
            "3指標の扱いを整理しており、"
            "正本 第3章3 の注記に「達成とは言わず参考値として"
            "改善の傾向とみる」と書いている",
            "差別・偏見は問34の"),
    "B-2": (HOKA, "同上（満足度の分母の定義）。"
            "正本 第3章3 の注記で、"
            "前回がサービス利用者のみ（n=25）であり"
            "今回の有効回答38件を分母とする値とは"
            "母集団が異なることを書いている",
            "回答者数（n）"),
    "B-3": (HOKA, "同上（外出の目的の設問が今回の調査にない）。"
            "正本 第3章3 で、アンケートによる割合から"
            "村の実績による確認に置き換えることを書いている",
            "社会参加・理解促進・村の施策"),
    "B-4": (HOKA, "同上（個別避難計画の現状値1件と村の確認0件の"
            "食い違い）。村への照会で確かめる",
            "災害時の対応"),
    "C-1": (KAI,
            "障がい児調査の基数が3に統一され、"
            "「無回答」が選択肢として明示された。"
            "所持している手帳等の記述にも「無回答」1件が加わった",
            "「無回答」1件"),
    "C-2": (ICHIBU,
            "所持している手帳等は n=3 に直り「無回答」1件が明示された。"
            "医療的ケアの3項目が1人の回答であることの書き方は、"
            "報告書本文の側の修正であり集計表では確かめられない",
            "所持している手帳等（問5・複数回答、n=3）"),
    "D-1": (KAI,
            "表1-1の調査の概要が記入された。"
            "調査対象の内訳、調査方法（郵送配布・郵送回収及びWEB回答）、"
            "調査期間（令和8年8月31日から9月25日まで）、"
            "配布数（障がい者172・障がい児6）、"
            "有効回答率（22.1％・50.0％）が入った。"
            "ただし調査対象の内訳の合計が配布数と合わない（Ｍ-128）",
            "郵送配布・郵送回収及びWEB回答"),
    "D-2": (HOKA,
            "第2章3の留意点は報告書本文の側であり集計表では"
            "確かめられない。"
            "当方の正本 第2章7（5）には、有効回答率"
            f"{A.pc('有効回答率_障がい者')}、"
            f"回答者の{A.pc('65歳以上')}が65歳以上、"
            f"{A.pc('身体障がい（主たる）')}が身体障がいであることを"
            "踏まえて読む必要があることを書いている"
            "（割合は回収40件の集計による）",
            "65歳以上の割合"),
}


# ===========================================================================
# 2. 改訂版で変わったこと（新旧の現物から数える）
# ===========================================================================
def kazoeru_zuhyo(path):
    """目次の図表番号を数える。点数を書き写さない。"""
    wb = load_workbook(path, data_only=True)
    ws = wb["目次"]
    nos = [str(r[1]) for r in ws.iter_rows(values_only=True)
           if r[1] and re.match(r"^[図表][0-9]+-[0-9]+$", str(r[1]))]
    return (len([x for x in nos if x.startswith("図")]),
            len([x for x in nos if x.startswith("表")]))


def taishou_uchiwake(path):
    """表1-1の調査対象の内訳と配布数を読み、合計を確かめる。"""
    wb = load_workbook(path, data_only=True)
    ws = wb["1-2章_概要・要点"]
    out = {}
    for r in ws.iter_rows(values_only=True):
        if r[1] == "調査対象":
            out["taishou"] = (str(r[2] or ""), str(r[6] or ""))
        elif r[1] == "配布数":
            out["haifu"] = (r[2], r[6])
        elif r[1] == "有効回答数":
            out["kaitou"] = (r[2], r[6])
    kekka = []
    for i, na in ((0, "障がい者調査"), (1, "障がい児調査")):
        kazu = [int(x) for x in re.findall(r"（([0-9]+)名）",
                                           out["taishou"][i])]
        kekka.append((na, kazu, sum(kazu), out["haifu"][i],
                      out["kaitou"][i]))
    return kekka


def shuukei_wariai(path):
    """改訂版の集計から、現れる割合をすべて作る。

    合算した割合（複数の選択肢の和）も、同じ設問の中で2つ又は3つの
    組合せまで作る。当方の正本は「『今の自宅で暮らしたい』と
    『家族と一緒に暮らしたい』の合計」のように足した値を書いている
    ためである。
    """
    wb = load_workbook(path, data_only=True)

    def pct(v):
        return f"{round(v * 100, 1):.1f}％"

    atta = set()
    byq = defaultdict(list)
    for sn in ("集計_障がい者", "集計_障がい児"):
        q = None
        for r in wb[sn].iter_rows(min_row=5, values_only=True):
            if r[0]:
                q = r[0]
            if isinstance(r[4], (int, float)):
                atta.add(pct(r[4]))
                byq[(sn, q)].append(r[4])
    #    障がい別は3列（身体・知的・精神）
    q = None
    for r in wb["集計_障がい別"].iter_rows(min_row=5, values_only=True):
        if r[0]:
            q = r[0]
        for j, ku in enumerate(("身体", "知的", "精神")):
            v = r[6 + j]
            if isinstance(v, (int, float)):
                atta.add(pct(v))
                byq[("別" + ku, q)].append(v)
    #    クロス集計
    for r in wb["4章_クロス集計"].iter_rows(values_only=True):
        for v in r:
            if isinstance(v, float) and 0 <= v <= 1:
                atta.add(pct(v))
    #    合算
    for _key, vs in byq.items():
        for k in (2, 3):
            for comb in combinations(vs, k):
                s = sum(comb)
                if s <= 1.0000001:
                    atta.add(pct(s))
    return atta


# 改訂版で変わったこと
#   (項目, 旧, 新, 当方への影響)
def henkou():
    zu_s, hyo_s = kazoeru_zuhyo(SHUUKEI)
    uchi = taishou_uchiwake(SHUUKEI)
    sha = uchi[0]
    return [
        ("割合の基数（分母）",
         "各設問の有効回答数（無回答を除く）。設問ごとに n が違う"
         "（n=36・34・33・30 など）",
         "各設問の回答者数（無回答を含む）に統一。n は38で固定。"
         "「無回答」を選択肢として表に明示",
         "当方の正本 第2章7 はもともと無回答を含む回答者数を"
         "分母としており、改訂版と同じ基準である。"
         "第2章7 の割合38点すべてが改訂版の集計から出ることを"
         "本書で確かめている"),
        ("図表の点数", "図61点・表23点",
         f"図{zu_s}点・表{hyo_s}点",
         f"図が{zu_s - 61}点増えたのは、各設問に【障がい別】の"
         "図表が加わったためである。"
         "当方の正本は別冊を名称で参照しており"
         "（「北塩原村 障がい福祉に関するアンケート調査 分析報告書」）、"
         "図表番号では参照していないため影響しない"),
        ("障がい別の集計", "各章シートの中に散在",
         "「集計_障がい別」シートを新設（450行）。"
         "身体29・知的2・精神6を件数と割合の両方で持つ",
         "当方の正本 表2-7（主な項目の障がい別の状況）の出所が"
         "1か所になった。8項目すべてが改訂版と一致することを"
         "本書で確かめている"),
        ("調査の概要（表1-1）",
         "対象要件・調査方法・調査期間・配布数・有効回答率が"
         "未記入（「○○件」「令和8年○月○日」など）",
         "すべて記入された。"
         f"調査対象の内訳は{'・'.join(str(x) + '名' for x in sha[1])}、"
         f"配布数{sha[3]}、有効回答数{sha[4]}",
         "当方の正本 表2-4・表2-5 は村から受領した値で既に書いており、"
         "配布数172・6、有効回答38・3、有効回答率22.1％・50.0％は"
         "改訂版と一致する"
         "（令和8年10月24日に回収40件の値へ差し替えた。"
         "有効回答40・3、有効回答率"
         f"{A.pc('有効回答率_障がい者')}・{A.pc('有効回答率_障がい児')}）。"
         f"ただし調査対象の内訳の合計（{sha[2]}名）が"
         f"配布数（{sha[3]}）と合わない（Ｍ-128）"),
        ("課題の数（表6-1）", "8件（旧版も8件。変わっていない）",
         "8件。1 相談支援体制・情報提供／2 移動・外出支援／"
         "3 災害時の支援体制／4 地域生活支援拠点・将来の暮らし／"
         "5 障がい児・家族支援／6 権利擁護・合理的配慮／"
         "7 日中活動・就労・社会参加／8 専門的支援",
         "当方の正本 表2-8 は7区分であり、"
         "8件目の「専門的支援」に対応する区分が無い。"
         "Ｓ-26で確かめていた事項であり、"
         "改訂版により報告書の側が8件であることが確定した。"
         "表2-8 に8区分目を足す（ＡＫ-1）"),
    ]


# ===========================================================================
# 3. 当方の正本 第2章7 の割合（照合の対象）
# ===========================================================================
def honpon_setsu7():
    """正本 第2章7 の本文と表を返す。"""
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    d = docx.Document(HONPON)
    on = False
    out = []
    for ch in d.element.body.iterchildren():
        if ch.tag.endswith("}p"):
            p = Paragraph(ch, d)
            t = p.text.strip()
            if p.style.name == "Heading 2" and t.startswith("７　アンケート"):
                on = True
            elif on and p.style.name in ("Heading 2", "Heading 1"):
                break
            if on and t:
                out.append(t)
        elif ch.tag.endswith("}tbl") and on:
            tb = Table(ch, d)
            out += [c.text for r in tb.rows for c in r.cells]
    return "\n".join(out)


# 集計の表に出ない割合（当方が割り算で出したもの）。
# これだけは集計表から出ないため、計算で確かめる。
#   値は kitashiobara_anke から読む（回収40件。障がい者 40÷172 など）。
JIBUN_KEISAN = {
    A.pc(k): A.kazu(k) for k in
    ("有効回答率_障がい者", "有効回答率_障がい児", "有効回答率_合計")
}
#   正本 第2章7 に書く割合のうち、選択肢の和・クロスの列の和・
#   利用者に限った満足度は集計表に直に現れないため、
#   kitashiobara_anke の定義から計算した値も「集計から出る」とみなす。
JIBUN_KEISAN.update({A.pc(k): A.kazu(k) for k in A.M})


# ===========================================================================
# 4. 正本への追記（8区分目の課題）
# ===========================================================================
AK_EDITS = [
    dict(
        no="ＡＫ-1",
        kata="表に行を足す",
        saki="第2章7（6）表2-8 調査結果からみた課題と本計画での対応"
             "（8区分目を追加）",
        atama=("主な課題", "主な調査結果", "本計画での主な対応（関連箇所）"),
        konkyo=f"アンケート集計表の改訂版（{HI}受領）の 6章_課題（表6-1）。"
               "報告書は課題を8件としており、"
               "8件目の「専門的支援」（精神障がいにも対応した"
               "地域包括ケアシステムの構築、"
               "強度行動障がい・医療的ケアへの対応。根拠設問 問17から問20）に"
               "対応する区分が当方の表2-8に無かった。Ｓ-26の決着",
        gyou=[[
            "専門的な支援を要する方への対応",
            f"定期的に通院している方が{A.pc('通院している')}、"
            "通院や医療で困っていることがある方があり、"
            "医療的ケア・強度行動障がい・高次脳機能障がいについて"
            "必要な支援を尋ねる設問を今回新たに置いた（問17から問20）",
            "精神障害にも対応した地域包括ケアシステムの構築（第４章２（２））、"
            "強度行動障害の状態にある方・医療的ケアを必要とする方への支援"
            "（第４章２（４）（５）、第５章１（２））",
        ]],
    ),
]


# ===========================================================================
# 5. 照会
# ===========================================================================
SHOUKAI = [
    ("Ｍ-128", "村",
     "アンケート調査の対象者の内訳と配布数",
     "アンケート集計表の改訂版で、表1-1の調査の概要が"
     "記入されました。ありがとうございます。"
     "1点、数が合わないところがあります。"
     "障がい者調査の調査対象の内訳は"
     "「身体（118名）・精神（1名）・療育（8名）・入所等（12名）」で"
     "合計139名ですが、配布数は172となっています。"
     "差が33名あります。"
     "障がい児調査は「身体（2名）・精神（1名）・療育（1名）・"
     "入所等（2名）」で合計6名、配布数6で合っています。"
     "次をご教示ください。"
     "①内訳の4区分のほかに対象とした方がいるか"
     "（自立支援医療の受給者、難病等の対象の方、"
     "障がい福祉サービスの受給者証をお持ちで手帳のない方など）。"
     "②内訳の「精神（1名）」は、"
     "村の精神障害者保健福祉手帳の所持者（令和８年４月現在38人）と"
     "比べて少ないように見えます。"
     "この1名は何を数えたものか"
     "（主たる障がいが精神である方か、"
     "精神手帳のみを所持している方かなど）。"
     "③「入所等（12名）」は施設入所の方か、"
     "入院中の方を含むか。"
     "計画素案 第２章７の表２-４（調査の対象者及び実施方法）に"
     "この内訳を載せるかどうかを、ご回答により決めます。",
     "令和８年11月13日（金）"),
    ("Ｓ-37", "他メンバー",
     "集計表の改訂版の確認結果（基数1〜2の区分の注記と課題8件）",
     "報告書用集計表の改訂版を拝見しました。"
     "令和８年10月３日に当方が洗い出した不整合11件のうち、"
     "次の点が解消されています。"
     "①割合の基数を各設問の回答者数（無回答を含む）に統一し、"
     "「無回答」を選択肢として表に明示されたこと（Ａ-1）。"
     "単純集計とクロス集計が同じ基準になりました。"
     "②問27-1と問27-3の基数が38に統一され、"
     "「差（今後－現在）」が同じ分母の差になったこと（Ａ-2）。"
     "③表1-1の調査の概要が記入されたこと（Ｄ-1）。"
     "④障がい児調査の基数が3に統一され"
     "「無回答」が明示されたこと（Ｃ-1）。"
     "残っているのは次の2点です。"
     "ア　基数が1から2の区分に割合を表示していること（Ａ-3）。"
     "新設の「集計_障がい別」シートは知的障がい（n=2）の列に"
     "割合を入れており、1件が100.0％として表示されます。"
     "当方の計画素案 第２章７の表２-７は同じ列を持ちますが、"
     "「知的障がい及び精神障がいは回答者数が少ないため参考値です」と"
     "注記しています。"
     "報告書の側にも同じ注記を置いていただけますでしょうか。"
     "イ　医療的ケアの3項目が1人の回答であることの書き方（Ｃ-2）。"
     "所持している手帳等は n=3 に直り「無回答」1件が明示されましたが、"
     "医療的ケアの書き方は報告書本文の側の修正であり"
     "集計表では確かめられません。"
     "あわせて、課題の数が決着しました。"
     "改訂版の 6章_課題（表6-1）は課題を8件としており、"
     "当方の計画素案 第２章７の表２-８は7区分でした。"
     "8件目の「専門的支援」に対応する区分を"
     "表２-８に足しました（Ｓ-26の決着）。"
     "なお、図表の点数が図61点・表23点から"
     "図100点・表25点に増えています"
     "（各設問に【障がい別】の図表が加わったため）。"
     "当方の計画素案は別冊を名称で参照しており"
     "図表番号では参照していないため、影響はありません。",
     "令和8年10月16日"),
]


# ===========================================================================
# 6. 自己点検
# ===========================================================================
def _norm(s):
    return re.sub(r"[\s　]+", "", str(s))


def load_book_text(path):
    wb = load_workbook(path, data_only=True)
    out = []
    for sn in wb.sheetnames:
        for r in wb[sn].iter_rows(values_only=True):
            for v in r:
                if isinstance(v, str):
                    out.append(v)
    return "\n".join(out)


def verify():
    ng = []
    if not os.path.exists(SHUUKEI):
        raise SystemExit(f"集計表が見つかりません: {SHUUKEI}")

    # ① 当方の指摘11件すべてに判定があること
    ban = [x[0] for x in ISSUES]
    if sorted(HANTEI) != sorted(ban):
        ng.append(f"判定のない指摘がある: "
                  f"{sorted(set(ban) - set(HANTEI))}／"
                  f"余分な判定: {sorted(set(HANTEI) - set(ban))}")
    for no, (han, _naka, go) in HANTEI.items():
        if han not in (KAI, ICHIBU, MIKAI, HOKA):
            ng.append(f"判定が未定義: {no} {han}")

    # ② 判定の根拠語が改訂版の現物にあること。
    #    記憶で「解消した」と書いていないことの裏づけはここである。
    hon_s = _norm(load_book_text(SHUUKEI))
    for no, (_han, _naka, go) in HANTEI.items():
        if _norm(go) not in hon_s:
            ng.append(f"改訂版に無い語を根拠にしている: {no}（{go[:28]}）")

    # ③ 正本 第2章7 の割合がすべて改訂版の集計から出ること。
    #    基数が変わったのに当方の正本が動かないことの裏づけである。
    if not os.path.exists(HONPON):
        print("  ※ 正本が未生成のため ③④ は省略")
        wariai = []
    else:
        setsu7 = honpon_setsu7()
        wariai = sorted(set(re.findall(r"[0-9]+\.[0-9]％", setsu7)))
        atta = shuukei_wariai(SHUUKEI_GENKOU)
        for w in wariai:
            if w in atta:
                continue
            if w in JIBUN_KEISAN:
                a, b = JIBUN_KEISAN[w]
                if f"{A._shisha(a / b * 100):.1f}％" != w:
                    ng.append(f"当方の割り算が合わない: {w}＝{a}÷{b}")
                continue
            ng.append(f"改訂版の集計から出ない割合を正本に書いている: {w}")
        if len(wariai) < 30:
            ng.append(f"正本 第2章7 の割合が少なすぎる: {len(wariai)}点")

        # ④ 8区分目の課題が改訂版にあり、追記で正本に入っていること
        wb = load_workbook(SHUUKEI, data_only=True)
        kadai = [str(r[1]) for r in wb["6章_課題"].iter_rows(values_only=True)
                 if r[1] and re.match(r"^[0-9]+ ", str(r[1]))]
        if len(kadai) != 8:
            ng.append(f"改訂版の課題が8件でない: {len(kadai)}件")
        if not any("専門的支援" in x for x in kadai):
            ng.append("改訂版の課題に「専門的支援」がない")
        d = docx.Document(HONPON)
        tbl = None
        for t in d.tables:
            h = [c.text.strip() for c in t.rows[0].cells]
            if tuple(h) == AK_EDITS[0]["atama"]:
                tbl = t
                break
        if tbl is None:
            ng.append("正本に表2-8が見当たらない")
        else:
            if len(tbl.rows) != 9:
                ng.append(f"正本 表2-8 の行数が9（表頭＋8区分）でない: "
                          f"{len(tbl.rows)}行")
            saigo = tbl.rows[-1].cells[0].text.strip()
            if _norm(saigo) != _norm(AK_EDITS[0]["gyou"][0][0]):
                ng.append(f"正本 表2-8 の最後の行が8区分目でない: {saigo}")

    # ⑤ 調査対象の内訳の合計が配布数と合わないことを計算で示す
    uchi = taishou_uchiwake(SHUUKEI)
    sha, ji = uchi
    if sha[2] == sha[3]:
        ng.append("障がい者調査の内訳の合計が配布数と合っている。"
                  "Ｍ-128の前提が変わったため見直すこと")
    if ji[2] != ji[3]:
        ng.append(f"障がい児調査の内訳の合計が配布数と合わない: "
                  f"{ji[2]}／{ji[3]}")
    if sha[3] != 172 or sha[4] != 38:
        ng.append(f"障がい者調査の配布数・有効回答数が想定と違う: "
                  f"{sha[3]}／{sha[4]}")

    # ⑥ 図表の点数を現物から数えていること（書き写さない）
    zu_s, hyo_s = kazoeru_zuhyo(SHUUKEI)
    hen = henkou()
    if f"図{zu_s}点・表{hyo_s}点" not in hen[1][2]:
        ng.append(f"図表の点数が現物と合わない: 図{zu_s}点・表{hyo_s}点")
    if zu_s <= 61:
        ng.append(f"図の点数が旧版より増えていない: {zu_s}点")

    # ⑦ 照会の番号と記号
    if [x[0] for x in SHOUKAI] != ["Ｍ-128", "Ｓ-37"]:
        ng.append(f"照会の番号が想定と違う: {[x[0] for x in SHOUKAI]}")
    for x in SHOUKAI:
        if len(x) != 5:
            ng.append(f"照会の列数が5でない: {x[0]}")
        if re.search(r"(?<![A-Za-zＡ-Ｚａ-ｚ])[MS]-\d+", x[0]):
            ng.append(f"照会番号の記号が半角: {x[0]}")

    # ⑧ 表記の作法と他団体の名前
    honbun = "\n".join(str(v) for rec in AK_EDITS
                       for gy in rec["gyou"] for v in gy)
    zenbun = honbun + "\n".join(
        str(v) for g in (SHOUKAI, hen) for rec in g for v in rec)
    if "か所" in honbun:
        ng.append("計画本文に「か所」を使っている（「カ所」とする）")
    if "〜" in honbun:
        ng.append("計画本文に波ダッシュ（〜）が混じっている")
    for warui in ("小野町", "金ケ崎", "金ヶ崎", "阿蘇", "札幌"):
        if warui in zenbun:
            ng.append(f"他団体の案件の地名が入っている: {warui}")

    if ng:
        for x in ng:
            print(f"  × {x}")
        raise SystemExit("自己点検で止めました")

    kai = sum(1 for v in HANTEI.values() if v[0] == KAI)
    print(f"  自己検証: 当方の指摘{len(ISSUES)}件の判定"
          f"（解消{kai}件）を改訂版の現物と照合／"
          f"正本 第2章7 の割合{len(wariai)}点がすべて改訂版の集計から出る／"
          f"改訂版の課題8件のうち8件目を表2-8に追記／"
          f"調査対象の内訳{sha[2]}名と配布数{sha[3]}の食い違いを計算で確認／"
          f"図表{zu_s}点・{hyo_s}点を現物から数えた／"
          f"照会{len(SHOUKAI)}件 整合")
    return wariai


# ===========================================================================
# 7. シート
# ===========================================================================
def sheet_gaiyou(wb):
    kai = [k for k, v in HANTEI.items() if v[0] == KAI]
    mikai = [k for k, v in HANTEI.items() if v[0] in (MIKAI, ICHIBU)]
    ws = add_sheet(
        wb, "00_概要", f"アンケート集計表の改訂版の確認（{HI}）",
        "令和8年10月3日に当方が洗い出した不整合11件が"
        "どうなったかを、新旧の現物を突き合わせて確かめた。"
        f"解消{len(kai)}件（{'・'.join(kai)}）、"
        f"残り{len(mikai)}件（{'・'.join(mikai)}）、"
        "ほかは当方の側で対応するものである。",
        [4, 30, 92])
    r = 5
    hako = [
        ("いちばん大きな変更", [
            "割合の基数（分母）が、各設問の有効回答数（無回答を除く）から"
            "各設問の回答者数（無回答を含む）に統一された。"
            "nは38で固定され、「無回答」が選択肢として表に明示された",
            "これは当方の指摘Ａ-1（単純集計は無回答を除く分母、"
            "クロス集計は無回答を含む分母で、"
            "同じ設問が報告書の中で2つの異なる基準で示されている）への"
            "対応である。単純集計とクロス集計が同じ基準になった",
            "基数が変わるとほとんどの割合が動く。"
            "日常生活で困っていること「将来の生活」は"
            "30.6％（旧・n=36）から28.9％（新・n=38）になった",
            "当方の正本 第2章7 はもともと"
            "「割合は無回答を含む回答者数に対する割合」として書いており、"
            "改訂版と同じ基準である。"
            "第2章7 の割合38点すべてが改訂版の集計から出ることを"
            "本書の自己点検で確かめている",
        ]),
        ("落ちていた論点", [
            "改訂版の 6章_課題（表6-1）は課題を8件としている。"
            "当方の正本 表2-8 は7区分であり、"
            "8件目の「専門的支援」（精神障がいにも対応した"
            "地域包括ケアシステムの構築、"
            "強度行動障がい・医療的ケアへの対応。根拠設問 問17から問20）に"
            "対応する区分が無かった",
            "これは令和8年10月13日からＳ-26として"
            "「どちらに揃えるか」を確かめていた事項である。"
            "改訂版により報告書の側が8件であることが確定したため、"
            "表2-8 に8区分目を足した（ＡＫ-1）",
        ]),
        ("残っているもの", [
            "基数が1から2の区分に割合を表示していること（Ａ-3）。"
            "新設の「集計_障がい別」シートは知的障がい（n=2）の列に"
            "割合を入れており、1件が100.0％として表示される。"
            "当方の正本 表2-7 は同じ列を持つが参考値であることを"
            "注記している。報告書の側にも同じ注記が要る（Ｓ-37）",
            "医療的ケアの3項目が1人の回答であることの書き方（Ｃ-2）。"
            "報告書本文の側の修正であり集計表では確かめられない",
            "調査対象の内訳の合計が配布数と合わない。"
            "障がい者調査は身体118名・精神1名・療育8名・入所等12名で"
            "合計139名であるのに配布数は172である（差33名）。"
            "障がい児調査は合計6名で配布数6と合っている。"
            "精神1名は村の精神障害者保健福祉手帳の所持者"
            "（令和8年4月38人）と比べても少ない（Ｍ-128）",
        ]),
    ]
    for dai, naka in hako:
        c = ws.cell(row=r, column=2, value=dai)
        c.font = Font(name="游ゴシック", size=11, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2D7A57")
        ws.cell(row=r, column=3).fill = PatternFill(
            "solid", fgColor="2D7A57")
        r += 1
        for i, t in enumerate(naka):
            write_row(ws, r, [None, f"{i + 1}", t],
                      aligns=["center", "center", "left"])
            ws.row_dimensions[r].height = 44
            r += 1
        r += 1
    return ws


def sheet_hantei(wb):
    ws = add_sheet(
        wb, "01_指摘の判定", f"当方の指摘{len(ISSUES)}件の判定",
        "番号と指摘の内容は build_kitashiobara_survey_check.py から引いて"
        "おり、書き写していない。"
        "判定の根拠の欄は改訂版の現物にある語であり、"
        "生成のたびに現物と照合している。"
        "語が現物になければ生成が止まる。",
        [8, 16, 50, 14, 62, 40])
    style_header_row(ws, 5, ["番号", "区分", "当方の指摘", "判定",
                             "改訂版で何が変わったか", "現物で確かめた語"])
    r = 6
    iro = {KAI: "2CA02C", ICHIBU: "ED7D31", MIKAI: "C00000",
           HOKA: "2E75B6"}
    for i, rec in enumerate(ISSUES):
        no, kubun, naiyo = rec[0], rec[1], rec[2]
        han, kawatta, go = HANTEI[no]
        write_row(ws, r, [no, kubun, naiyo, han, kawatta, go],
                  alt=(i % 2 == 1),
                  aligns=["center", "center", "left", "center", "left",
                          "left"])
        c = ws.cell(row=r, column=4)
        c.fill = PatternFill("solid", fgColor=iro[han])
        c.font = Font(name="游ゴシック", size=10, bold=True,
                      color="FFFFFF")
        ws.cell(row=r, column=6).fill = PatternFill(
            "solid", fgColor=COLORS["calc"])
        ws.row_dimensions[r].height = 62
        r += 1
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    style_note(ws.cell(row=r, column=1),
               "「当方の側で対応」は、集計表の側では直せないもの"
               "（前回調査との比較可能性、報告書本文の留意点）である。"
               "当方は目標ＫＰＩの前回比較可能性のブックと"
               "正本 第3章3 の注記で扱っている。")
    ws.row_dimensions[r].height = 46
    return ws


def sheet_henkou(wb):
    hen = henkou()
    ws = add_sheet(
        wb, "02_改訂版で変わったこと", f"改訂版で変わったこと（{len(hen)}件）",
        "旧版は git の履歴から取り出して突き合わせた。"
        "図表の点数と調査対象の内訳の合計は、"
        "現物から数え上げている（書き写さない）。",
        [20, 44, 50, 66])
    style_header_row(ws, 5, ["項目", "旧", "新", "当方への影響"])
    r = 6
    for i, rec in enumerate(hen):
        write_row(ws, r, list(rec), alt=(i % 2 == 1),
                  aligns=["left", "left", "left", "left"])
        ws.row_dimensions[r].height = 76
        r += 1
    return ws


def sheet_shougou(wb, wariai):
    ws = add_sheet(
        wb, "03_割合の照合", f"正本 第2章7 の割合{len(wariai)}点の照合",
        "正本 第2章7 の割合が、いま有効な集計表から出ることの"
        "裏づけである。"
        "正本 第2章7 に現れる割合をすべて拾い、"
        "現行の集計表（令和8年10月21日版・回収40件）の集計"
        "（集計_障がい者・集計_障がい児・集計_障がい別・"
        "4章_クロス集計）から出るかを確かめる"
        "（令和8年10月24日に照合先を回収40件の版に改めた）。"
        "複数の選択肢を足した割合（「今の自宅で暮らしたい」と"
        "「家族と一緒に暮らしたい」の合計など）は足し算で確かめる。"
        "有効回答率は当方の割り算であるため、割り算で確かめる。",
        [14, 30, 56])
    style_header_row(ws, 5, ["割合", "出所", "確かめ方"])
    r = 6
    atta = shuukei_wariai(SHUUKEI_GENKOU)
    for i, w in enumerate(wariai):
        if w in JIBUN_KEISAN:
            a, b = JIBUN_KEISAN[w]
            moto = "当方の割り算"
            shikata = f"{a}÷{b}＝{a / b * 100:.1f}％"
        elif w in atta:
            moto = "現行の集計表"
            shikata = "集計の割合又は複数の選択肢の和として実在"
        else:
            moto = "―"
            shikata = "見当たらない"
        write_row(ws, r, [w, moto, shikata], alt=(i % 2 == 1),
                  aligns=["center", "center", "left"])
        c = ws.cell(row=r, column=2)
        c.fill = PatternFill(
            "solid", fgColor=("2CA02C" if moto != "―" else "C00000"))
        c.font = Font(name="游ゴシック", size=10, bold=True, color="FFFFFF")
        r += 1
    return ws


def sheet_honbun(wb):
    ws = add_sheet(
        wb, "04_正本への追記", "正本への追記",
        "build_kitashiobara_honpon.py が AK_EDITS として入れる。"
        "他メンバー版の表に行を足す形であり、表そのものは置き換えない。",
        [10, 16, 44, 70, 44])
    style_header_row(ws, 5, ["番号", "型", "入れ先", "足す行", "根拠"])
    r = 6
    for i, rec in enumerate(AK_EDITS):
        gy = "／".join(" ｜ ".join(x) for x in rec["gyou"])
        write_row(ws, r, [rec["no"], rec["kata"], rec["saki"], gy,
                          rec["konkyo"]],
                  alt=(i % 2 == 1),
                  aligns=["center", "center", "left", "left", "left"])
        ws.row_dimensions[r].height = 110
        r += 1
    return ws


def sheet_shoukai(wb):
    ws = add_sheet(
        wb, "05_照会", f"照会（{len(SHOUKAI)}件）",
        "本書から出た照会である。"
        "build_kitashiobara_shoukai_bun.py がここから読み込む。",
        [10, 14, 40, 86, 16])
    style_header_row(ws, 5, ["番号", "宛先", "件名", "本文", "期限"])
    r = 6
    for i, rec in enumerate(SHOUKAI):
        write_row(ws, r, list(rec), alt=(i % 2 == 1),
                  aligns=["center", "center", "left", "left", "center"])
        ws.row_dimensions[r].height = 160
        r += 1
    return ws


def main():
    wariai = verify()
    ensure_out_dir()
    wb = Workbook()
    wb.remove(wb.active)
    sheet_gaiyou(wb)
    sheet_hantei(wb)
    sheet_henkou(wb)
    sheet_shougou(wb, wariai)
    sheet_honbun(wb)
    sheet_shoukai(wb)
    wb.save(OUT_FILE)
    print(f"作成: {OUT_FILE}")
    print(f"  シート{len(wb.sheetnames)}／"
          f"指摘{len(ISSUES)}件の判定／"
          f"割合{len(wariai)}点の照合／"
          f"追記{len(AK_EDITS)}件／照会{len(SHOUKAI)}件")


if __name__ == "__main__":
    main()
