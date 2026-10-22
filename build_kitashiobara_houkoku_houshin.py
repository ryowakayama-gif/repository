"""報告書の修正方針と計画への影響ブック ジェネレータ

ご指示（令和8年10月22日）
  「報告書の修正方針の検討と計画への影響範囲を再度確認お願い致します」

二つを分けて答える。
  ① 報告書（アンケート調査結果報告書 令和8年10月22日版）を
     どう直すか。誰が、何を、どの順で直すか。
  ② 直した結果が計画素案（正本）のどこに及ぶか。

① 修正方針 ― 置換文をお渡しする形を採る
  直し残し11件はいずれも機械的に直せるものである。
  指摘だけをお返しすると直し方が伝わらず往復が増えるため、
  11件すべてについて「置換前（報告書の現物）→置換後（当方の案）」を
  用意し、そのまま適用できる形でお渡しする。
  当方が報告書そのものを直すことはしない（他メンバーの成果品であり、
  版が分かれる）。

  令和8年10月5日の当方のレビューで挙げた20件（Ｒ-1〜Ｒ-20）の
  反映状況も原典に当てて確かめた。解消2件・一部対応2件のほかは
  いずれも手が付いていない。これらを「報告書に入れるもの」と
  「計画（正本）で受けるもの」に分けた。報告書は調査の分析であり、
  計画の施策を書く場ではないためである。

  いちばん重いのはＲ-4（自由記述に個人が特定され得る記述が残っている）
  である。報告書は策定委員会に出す資料であり公表の対象となるため、
  村と要約の程度を協議する必要がある（Ｍ-132）。

② 計画への影響範囲 ― 順序が変わる
  前回（令和8年10月21日）は「Ｓ-39で第2章7を回収40件に差し替えて
  いただく」ことを先に置いていた。
  しかし直し残し11件のうち7件は、同じ文が正本の第2章7にある。
  壊れた報告書から写すと壊れた文が計画に入るため、
  Ｓ-40（直し残しの修正）がＳ-39（差し替え）より前に来る。

出力: output/北塩原村_報告書の修正方針と計画への影響.xlsx
"""

import os
import re
import sys
import warnings

import pdfplumber
from docx import Document
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from kitashiobara_common import (
    COLORS, FONT, OUT_DIR, SRC_DIR, add_sheet, ensure_out_dir,
    style_header_row, style_title, write_row,
)
from build_kitashiobara_anke_kaishuu import wakeru
from build_kitashiobara_houkoku_hanei import FUBI, KYU, SHIN
from build_kitashiobara_houkokusho_review import (
    HI, LO, MID, OK, SHITEKI, ZENKAI,
)

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

OUT_FILE = f"{OUT_DIR}/北塩原村_報告書の修正方針と計画への影響.xlsx"
HONPON = f"{OUT_DIR}/北塩原村_計画素案_正本_移植後.docx"
ZENKAI_PDF = (f"{SRC_DIR}/現行計画/"
              "第7期障がい福祉計画等策定に係るアンケート調査_"
              "調査結果報告書.pdf")
SHUKEI = (f"{SRC_DIR}/アンケート結果/"
          "報告書用集計表_令和8年10月21日版_回収40.xlsx")
KIJUNBI = "基準日：令和8年10月22日"


# ===========================================================================
# 1. 修正方針の3案
# ===========================================================================
SAIYO = "採る"
HUSAIYO = "採らない"

HOUSHIN = [
    ("案①", "指摘だけをお返しする", HUSAIYO,
     "箇所と分類をお伝えし、直し方は他メンバーに委ねる",
     "直し方が伝わらないと再度崩れる。11件はいずれも機械的に直せる"
     "ものであり、往復を増やす理由がない"),
    ("案②", "置換文をお渡しする", SAIYO,
     "11件すべてについて「置換前（報告書の現物）→置換後（当方の案）」"
     "を用意し、そのまま適用できる形でお渡しする",
     "置換前は報告書の現物であり、置換後は旧版の文又は報告書の別の"
     "箇所にある正しい形を用いる。当方の判断が入る2件は理由を添える"),
    ("案③", "当方が報告書を直す", HUSAIYO,
     "当方が報告書そのものを直して版を差し替える",
     "報告書は他メンバーの成果品である。当方が直すと版が分かれ、"
     "どちらが正本か分からなくなる。本文を書き換えるのは事実として"
     "誤っているものに限るという取決めにも反する"),
]


# ===========================================================================
# 2. 直し残し11件の置換文
#
#    置換前は FUBI（build_kitashiobara_houkoku_hanei.py）が持つ
#    壊れている語を含む形で書く。壊れている語そのものを二重に持たない。
#    置換後は旧版の文、又は報告書の別の箇所にある正しい形を用いる。
#    自前で言い回しを作るのは、どちらも無い場合に限る。
# ===========================================================================
CHIKAN = {
    "Ｒ-1": dict(
        mae="福祉サービスの利用状況・今後の相談支援、"
            "福祉サービス利用意向、、就労・日中活動",
        ato="福祉サービスの利用状況・今後の利用意向、相談支援、"
            "就労・日中活動",
        betsu="",
        memo="旧版の語順に戻す。正本 第2章7（1）も"
             "「福祉サービスの利用状況・今後の利用意向、相談支援、就労」"
             "としている。「福祉サービス利用意向」を明示する意図であれば"
             "「今後の利用意向」がすでにこれを含む"),
    "Ｒ-2": dict(
        mae="それぞれ22.5％ありとなっており、サービスの周知が課題です",
        ato="それぞれ22.5％となっており、サービスの周知が課題です",
        betsu="それぞれ22.5％あり、サービスの周知が課題です",
        memo="「あり」と「となっており」のどちらでもよい。"
             "置換後は改訂の意図（「となっており」）を残した形、"
             "別案は旧版のままの形"),
    "Ｒ-3": dict(
        mae="「制度を知らない」もという回答も17.5％あります",
        ato="「制度を知らない」という回答も17.5％あります",
        betsu="",
        memo="「も」と「という回答も」の重なり。"
             "改訂の意図（「という回答も」）を残す"),
    "Ｒ-4": dict(
        mae="今後利用したいサービスでも3件全てが"
            "「移動支援」「移動に関する支援」を挙げています",
        ato="今後利用したいサービスでも3件全てが「移動支援」を"
            "挙げています",
        betsu="",
        memo="問29の選択肢名は「移動支援」であり、"
             "「移動に関する支援」という選択肢は集計表にない。"
             "報告書の第6章 課題2 も「移動支援」と書いており、"
             "意図の確認は要らない"),
    "Ｒ-5": dict(
        mae="高次脳機能障がい支援に関してのある方に対する"
            "必要な支援（問20）27.5％",
        ato="高次脳機能障がいのある方に対する必要な支援（問20）27.5％",
        betsu="",
        memo="改訂の意図（「のある方に対する」）を残し、"
             "「支援に関して」を削る。"
             "正本 第2章7（2）も「高次脳機能障がいのある方に必要な支援」"
             "としている"),
    "Ｒ-6": dict(
        mae="これらの設問は無回答の多さも踏まえてが"
            "含まれることを踏まえて解釈する必要があります",
        ato="これらの設問は無回答が含まれることを踏まえて"
            "解釈する必要があります",
        betsu="これらの設問は無回答の多さも踏まえて"
              "解釈する必要があります",
        memo="置換後は改訂の意図（「無回答が含まれることを踏まえて」）"
             "を残した形、別案は旧版のままの形"),
    "Ｒ-7": dict(
        mae="サービス利用実績等の実績データ統計データを総合して勘案し、"
            "組み合わせて、ニーズを多面的に把握する",
        ato="サービス利用実績等の実績データ・統計データを総合して"
            "勘案し、ニーズを多面的に把握する",
        betsu="",
        memo="「実績データ」「統計データ」を中点で並べ、"
             "「総合して勘案し」と「組み合わせて」の重なりを解く"),
    "Ｒ-8": dict(
        mae="作成を進めます。、災害時を想定した図上訓練や",
        ato="作成を進めます。あわせて、災害時を想定した図上訓練や",
        betsu="作成を進めるとともに、災害時を想定した図上訓練や",
        memo="句点の後に接続の語を置く。"
             "別案は一文にまとめる形"),
    "Ｒ-9": dict(
        mae="避難所における安全性と合理的配慮のに基づく"
            "支え合いの提供体制を整えます",
        ato="避難所における安全性の確保と合理的配慮の提供体制を"
            "整えます",
        betsu="避難所における安全性の確保と、"
              "合理的配慮に基づく支え合いの体制を整えます",
        memo="【意図の確認が要る唯一の箇所】"
             "「安全性」と「支え合い」でそれぞれ何を指すかによる。"
             "置換後は旧版（合理的配慮の提供体制）に安全性を加えた形、"
             "別案は「支え合い」を残した形。"
             "確認が取れない場合は置換後を推奨する"
             "（当方の正本に「避難所の安全性」に当たる記述がないため、"
             "別案を採ると計画に受け先のない記述が残る）"),
    "Ｒ-10": dict(
        mae="機能を充実させる、とともに、制度の周知を図ります",
        ato="機能を充実させるとともに、制度の周知を図ります",
        betsu="",
        memo="読点を削る"),
    "Ｒ-11": dict(
        mae="電話リレーサービス(日本財団電話リレーサービス)など",
        ato="電話リレーサービス（日本財団電話リレーサービス）など",
        betsu="電話リレーサービスなど",
        memo="報告書の他の箇所はいずれも全角のかっこである。"
             "なお提供機関の正式名称（法人格の表記）は当方で"
             "確かめられていない。確かめられない場合は別案"
             "（かっこ書きを置かない形）でも足りる"),
}


# ===========================================================================
# 3. 当方のレビュー指摘（Ｒ-1〜Ｒ-20）の反映状況
#
#    指摘の本文は build_kitashiobara_houkokusho_review.py の SHITEKI から
#    読み込む。ここには照合語と、照合の結果から定めた扱いだけを持つ。
#    状態は報告書の現物を照合して出し、記録と食い違えば生成を止める。
# ===========================================================================
KAISHO = "解消"
ICHIBU = "一部対応"
MITAIOU = "未対応"

HOUKOKU = "報告書に入れる"
SEIHON_UKE = "計画で受けている"
RYOUHOU = "報告書に1文・計画で受ける"
MURA_KYOUGI = "村と協議する"

#   照合語は、ふつうは「有＝対応の証拠」である。
#   Ｒ-4（自由記述に個人が特定され得る記述が残っている）だけは逆で、
#   語が残っていること自体が未対応の証拠である。向きを明示する。
MUKI_GYAKU = {"R-4"}

RJOUKYOU = [
    ("R-1", ["経年"], MITAIOU, HOUKOKU,
     "前回調査（発送164人・回収95人・回収率57.9％）との対比が"
     "1件もない。正本にも対比の記述はなく、両方の宿題である"),
    ("R-2", ["57.9"], MITAIOU, HOUKOKU,
     "回収率が57.9％から23.3％へ34.6ポイント下がったことに"
     "触れていない。結果の読み方の根幹であり、委員会資料には要る"),
    ("R-3", ["給付実績"], MITAIOU, RYOUHOU,
     "問27-1の自己申告が村の給付実績と食い違う点。"
     "正本 第4章2（3）・第5章1（2）で受けているが、"
     "報告書にも自己申告である旨の注記が要る"),
    ("R-4", ["経管栄養"], MITAIOU, MURA_KYOUGI,
     "自由記述の要旨（表7）に個人が特定され得る記述が残っている。"
     "正本は自由記述を引いていないため、報告書だけの論点である"),
    ("R-5", ["人材確保"], MITAIOU, SEIHON_UKE,
     "成果目標の七（人材確保・生産性向上）・八（質の向上）。"
     "住民調査では把握できない。正本 第4章2（7）（8）で受けている。"
     "報告書は表6-1の脚注で足りる"),
    ("R-6", ["保育所等", "受入"], MITAIOU, SEIHON_UKE,
     "保育所等における障がい児の受入れの体制。"
     "行政データであり調査では把握できない。"
     "正本 第4章2（4）・第5章1（5）で受けている"),
    ("R-7", ["引継ぎ"], MITAIOU, SEIHON_UKE,
     "前回計画課題の引継ぎ10件との対応。"
     "正本 第3章・表3-3で受けている。報告書は表6-1に列を足す形でよい"),
    ("R-8", ["外出の目的"], MITAIOU, SEIHON_UKE,
     "第4次障がい者計画 施策目標⑦の指標。"
     "正本 第3章3で測り方を村の実績に置き換えており、受けている"),
    ("R-9", ["市町村必須事業"], ICHIBU, SEIHON_UKE,
     "課題2の方向性。令和8年10月22日版で「地域生活支援事業"
     "（市町村必須事業）の範囲で実施する移動支援の概要を正しく周知」"
     "が入り、制度未周知の整理が加わった。"
     "実績が0人であることへの言及はない。正本 第5章3（2）で受けている"),
    ("R-10", ["削減"], MITAIOU, SEIHON_UKE,
     "施設入所支援の利用意向が国基準の方向と逆である点"
     "（12.5％→22.5％で＋10.0ポイント。回収40件でも最大の増加）。"
     "正本 第2章7（3）・第4章2（1）で受けている"),
    ("R-11", ["分母"], MITAIOU, SEIHON_UKE,
     "問27-3を見込量の参考にする際の分母。"
     "正本 第2章7（3）が「回答の割合に対象となる方の人数を掛けて"
     "見込量を求めることはしない」と定めており、受けている"),
    ("R-12", ["近隣市町村と連携した確保", "移動支援"], MITAIOU, HOUKOKU,
     "障がい児の利用意向（移動支援3件＝全員）が課題5の方向性に"
     "入っていない。第2章の本文には現れるが方向性に挙がっていない。"
     "課題2には移動支援が入ったため、課題5にも1行要る"),
    ("R-13", ["複数の区分"], MITAIOU, HOUKOKU,
     "複数回答を区分とするクロスでは1人が複数の区分に計上され"
     "合計が全体を超える旨の注記。報告書の見方の注記は"
     "不足側のみを説明している"),
    ("R-14", ["100.0％（"], MITAIOU, HOUKOKU,
     "第3章の本文が「100.0％」を実数なしで引用している（20件）。"
     "第4章のクロス表は「100.0%(2)」と実数を併記しており"
     "表記が非対称である"),
    ("R-15", ["調査の概要"], KAISHO, HOUKOKU,
     "受領版のZIP構造の欠落。令和8年10月21日版・22日版は"
     "いずれも正常に開ける"),
    ("R-16", ["ポイント"], KAISHO, HOUKOKU,
     "表5の「差」が丸めた割合の差で計算されていた点。"
     "回収数が40件になったことで1件＝2.5％と割り切れるため"
     "丸めの差が生じない。表頭の単位（ポイント）の注記もある"),
    ("R-17", ["-15.0"], MITAIOU, HOUKOKU,
     "表5の「特にない」「無回答」行の差が空欄。"
     "「特にない」は45.0％→30.0％（-15.0ポイント）で、"
     "利用意向の広がりを示す情報である"),
    ("R-18", ["医療的ケア者"], MITAIOU, SEIHON_UKE,
     "「医療的ケア者」（成人）の用語。"
     "正本 第4章2（4）が法改正（令和8年法律第77号）を書いており"
     "受けている。報告書は用語に1文で足りる"),
    ("R-19", ["参考値"], ICHIBU, HOUKOKU,
     "基数5未満の区分（65件中28件）。"
     "報告書の見方に一般の注記はあるが、"
     "該当する区分への記号と表ごとの脚注はない"),
    ("R-20", ["問27-1", "問27-3"], KAISHO, HOUKOKU,
     "現在の利用×今後の利用希望のクロス。"
     "表5と図3-50が同じ比較を提供しており、"
     "クロス分析のテーマに立てる必要はない"),
]


# ===========================================================================
# 4. 公表前に要約を要する自由記述（Ｒ-4）
#
#    村の人口は約2,400人、障がい者調査の有効回答は40件である。
#    状態像を書くと村内で特定され得る。
#    要約の程度は村と協議する事項であり、当方が決めない。
# ===========================================================================
JIYU = [
    ("経管栄養を行っているが、口から食べたいという思いがある。",
     "医療的ケアを必要とする方から、本人の希望を踏まえた支援を"
     "求める声がありました。",
     "本村の医療的ケア児は1人。経管栄養は村内で極めて限られる"),
    ("村役場に手話通訳者を継続して配置してほしい。",
     "村の窓口における意思疎通支援について、継続的な体制と"
     "障がいへの理解の促進を求める声がありました。",
     "手話を必要とする方は村内で限られる。"
     "施策としては第5章2（1）で受けている"),
    ("村外への通院時の手話通訳者・付き添いは事前の予約が必要であり、",
     "通院時の意思疎通支援について、緊急時にも対応できる体制を"
     "求める声がありました。",
     "同上。緊急時の対応は第5章2（15）で受けている"),
    ("家族が突然障がいを負った際、相談後の対応に時間がかかり、",
     "家族に障がいが生じた際の相談から支援に至るまでの時間について、"
     "改善を求める声がありました。",
     "事案が特定され得る。相談支援の経路は第5章2（2）で受けている"),
]


# ===========================================================================
# 5. 段取り（順序が変わった）
# ===========================================================================
DANDORI = [
    (1, "Ｓ-40", "他メンバー",
     "直し残し11件を置換文でお戻しする（本書02シート）",
     "令和８年10月31日（金）",
     "報告書は策定委員会に出す資料である。"
     "ここが直らないうちに第2章7を差し替えると、"
     "壊れた文が計画に入る"),
    (2, "Ｍ-132", "村",
     "自由記述の要約の程度を協議する（本書04シート）",
     "令和８年10月31日（金）",
     "報告書は公表の対象となる。"
     "村の人口約2,400人・有効回答40件の規模では状態像で特定され得る"),
    (3, "Ｓ-42", "他メンバー",
     "当方のレビュー指摘20件のうち手の付いていないものの扱いを"
     "合わせる（本書03シート）",
     "令和８年10月31日（金）",
     "報告書に入れるもの（5件）と、計画で受けており脚注で足りるもの"
     "（8件）に分かれる。二重に書かないよう先に決める"),
    (4, "Ｓ-39", "他メンバー",
     "第2章7の割合を回収40件に差し替えていただく",
     "令和８年11月14日（金）",
     "Ｓ-40・Ｓ-42の後に行う。前回はこれを先に置いていたが、"
     "直し残しのうち正本 第2章7 に同じ文があるものがあるため"
     "順序を改めた（件数は05シート）"),
    (5, "―", "当方",
     "集計表を令和8年10月21日版に差し替え、"
     "当方の記述8件（ＡＫ-2〜ＡＫ-9）を直す",
     "差し替えの受領後",
     "いま差し替えると第2章7の割合39点の照合が落ち、"
     "正本が生成できなくなる"),
    (6, "―", "当方",
     "再生成し、第2章7の割合39点すべてが新しい集計から出ることと、"
     "表2-8の数値12点が新しい集計と合うことを確かめる",
     "上記の後",
     "検査で確かめる。目で見て直さない"),
]


# ===========================================================================
# 6. 現物を読む
# ===========================================================================
def paras(path):
    return [p.text for p in Document(path).paragraphs]


def honbun(path):
    """全文（段落＋表のセル）。所在が分かる形で返す。"""
    d = Document(path)
    out = [("P%d" % i, p.text) for i, p in enumerate(d.paragraphs)]
    for ti, t in enumerate(d.tables):
        for ri, r in enumerate(t.rows):
            for ci, c in enumerate(r.cells):
                out.append(("T%d.%d.%d" % (ti, ri, ci), c.text))
    return out


def terasu(rows, keys):
    if not keys:
        return []
    return [loc for loc, t in rows if all(k in t for k in keys)]


def hyou28(path):
    """正本の表2-8（主な課題）の行と、主な調査結果欄の割合の数。

    課題の区分（行）は回収数が変わっても動かないが、
    主な調査結果欄の数値は動く。誰が直すかを分けるために数える。
    """
    d = Document(path)
    for t in d.tables:
        if not t.rows:
            continue
        if t.rows[0].cells[0].text.strip() == "主な課題":
            gyou = len(t.rows) - 1
            n = 0
            for r in t.rows[1:]:
                n += len(re.findall(r"\d+\.\d＋?％", r.cells[1].text)) \
                    or len(re.findall(r"\d+\.\d％", r.cells[1].text))
            return gyou, n
    return 0, 0


def hyou5_sa_kara():
    """表5の「特にない（各設問）」行の差の欄が空かどうか。

    None は表そのものが見つからない場合。
    """
    d = Document(SHIN)
    for t in d.tables:
        if not t.rows:
            continue
        if t.rows[0].cells[0].text.strip() != "サービス":
            continue
        if "差" not in t.rows[0].cells[-1].text:
            continue
        for row in t.rows[1:]:
            if row.cells[0].text.strip().startswith("特にない"):
                return row.cells[-1].text.strip() == ""
    return None


#   直し残しが正本でどう受けられているか。3つに分かれる。
#     ONAJI   … 正本 第2章7 に同じ文がある。差し替えの対象であり、
#                壊れた報告書から写すと壊れた文が計画に入る
#     YOUYAKU … 表2-8 の対応欄又は当方の文で趣旨を受けている。
#                文そのものは写していない
#     NASHI   … 正本に受け先がない
#   手がかりは数値を含まない語にする（回収数が違うため数値では当たらない）。
ONAJI = "第2章7に同じ文がある"
YOUYAKU = "趣旨を受けている（文は写していない）"
NASHI = "正本に受け先がない"

UKEKATA = {
    "Ｒ-1": (["生活実態", "今後の利用意向"], ONAJI,
            "第2章7（1）調査の目的"),
    "Ｒ-2": (["内容をよく知らない", "周知"], ONAJI,
            "第2章7（5）サービスの認知"),
    "Ｒ-3": (["地域生活支援拠点", "緊急時の受入れ"], ONAJI,
            "第2章7（5）地域生活支援拠点等"),
    "Ｒ-4": (["今後利用したいサービス", "移動支援"], ONAJI,
            "第2章7（5）障がい児調査"),
    "Ｒ-5": (["無回答", "高次脳機能障がい"], NASHI,
            "無回答の多い設問の記述が正本にない"),
    "Ｒ-6": (["無回答", "踏まえ"], NASHI,
            "無回答の解釈の記述が正本にない"),
    "Ｒ-7": (["多面的"], NASHI,
            "多面的な把握の記述が正本にない"),
    "Ｒ-8": (["個別避難計画の作成推進"], YOUYAKU,
            "表2-8 の対応欄"),
    "Ｒ-9": (["福祉避難所"], YOUYAKU,
            "第5章2（11）災害時にサービスを止めないための備え"),
    "Ｒ-10": (["地域生活支援拠点", "機能の充実"], YOUYAKU,
             "表2-8 の対応欄・第3章3"),
    "Ｒ-11": (["意思疎通支援", "手話通訳"], YOUYAKU,
             "第5章2（1）住民への情報提供等"),
}


def utsushi(rows):
    """直し残しが正本でどう受けられているか。(手がかり, 所在) を返す。"""
    return {no: (UKEKATA[no][0], terasu(rows, UKEKATA[no][0]))
            for no in UKEKATA}


def zenkai_genten():
    """前回調査の発送・回収・回収率を原典から取る。"""
    with pdfplumber.open(ZENKAI_PDF) as pdf:
        t = "".join((p.extract_text() or "") for p in pdf.pages)
    return re.sub(r"[\s　]+", "", t)


def shukei_sentakushi(toi):
    """集計表にある選択肢の名の集合。"""
    from openpyxl import load_workbook
    wb = load_workbook(SHUKEI, read_only=True, data_only=True)
    out = set()
    for sh in ("集計_障がい者", "集計_障がい児"):
        for r in wb[sh].iter_rows(values_only=True):
            if r and r[0] and str(r[0]).strip() == toi and len(r) > 2:
                out.add(str(r[2]).strip())
    wb.close()
    return out


# ===========================================================================
# 7. シート
# ===========================================================================
def _kazu_jou():
    k = {}
    for rec in RJOUKYOU:
        k[rec[2]] = k.get(rec[2], 0) + 1
    return k


def _kazu_uke():
    k = {}
    for rec in RJOUKYOU:
        if rec[2] == KAISHO:
            continue
        k[rec[3]] = k.get(rec[3], 0) + 1
    return k


def sheet_ketsuron(wb):
    jou, uke = _kazu_jou(), _kazu_uke()
    ws = add_sheet(
        wb, "00_結論", "報告書の修正方針と計画への影響",
        f"直し残し{len(FUBI)}件は、置換文をお渡しする形で直していただき"
        "ます（案②）。"
        f"令和8年10月5日のレビューで挙げた{len(RJOUKYOU)}件のうち、"
        f"解消は{jou.get(KAISHO, 0)}件、一部対応が{jou.get(ICHIBU, 0)}件で、"
        f"{jou.get(MITAIOU, 0)}件は手が付いていません。"
        "計画への影響では、差し替えの順序が変わります。"
        f"{KIJUNBI}。",
        [30, 20, 20, 62])
    r = 5
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    style_title(ws.cell(row=r, column=1), "① 報告書の修正方針",
                fill=COLORS["subhead"], size=11)
    r += 1
    style_header_row(ws, r, ["", "件数", "誰が直すか", "いつ"])
    r += 1
    for nm, n, dare, itsu in (
            ("直し残し（日本語が崩れている箇所）", f"{len(FUBI)}件",
             "他メンバー（置換文を当方が用意）", "10月31日（Ｓ-40）"),
            ("自由記述の要約（公表前）", f"{len(JIYU)}件",
             "村と協議のうえ他メンバー", "10月31日（Ｍ-132）"),
            ("レビュー指摘のうち報告書に入れるもの",
             f"{uke.get(HOUKOKU, 0)}件",
             "他メンバー", "10月31日に方針（Ｓ-42）"),
            ("レビュー指摘のうち計画で受けているもの",
             f"{uke.get(SEIHON_UKE, 0)}件",
             "当方（正本に既にある）", "対応済み"),
            ("レビュー指摘のうち両方に要るもの",
             f"{uke.get(RYOUHOU, 0)}件",
             "報告書に1文・計画は対応済み", "10月31日に方針（Ｓ-42）"),
    ):
        write_row(ws, r, [nm, n, dare, itsu],
                  aligns=["left", "center", "left", "left"])
        ws.cell(row=r, column=2).fill = PatternFill("solid", fgColor="FFF2CC")
        ws.row_dimensions[r].height = 24
        r += 1
    r += 1

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    style_title(ws.cell(row=r, column=1), "② 計画への影響範囲",
                fill=COLORS["subhead"], size=11)
    r += 1
    rows = honbun(HONPON) if os.path.exists(HONPON) else []
    u = utsushi(rows) if rows else {}
    aru = sum(1 for k in u if u[k][1])
    gyou, suuchi = hyou28(HONPON) if os.path.exists(HONPON) else (0, 0)
    style_header_row(ws, r, ["", "件数", "", "どうするか"])
    r += 1
    kk = _kazu_uke_kata()
    kare, ware = wakeru()
    for nm, n, dou in (
            ("直し残しのうち正本 第2章7 に同じ文があるもの",
             f"{kk.get(ONAJI, 0)}件／{len(FUBI)}件",
             "正本の文はいまは壊れていない（旧版の文である）。"
             "壊れた報告書から写すと壊れた文が計画に入るため、"
             "Ｓ-40をＳ-39より前に置く"),
            ("同じく、趣旨だけを受けているもの",
             f"{kk.get(YOUYAKU, 0)}件",
             "表2-8の対応欄又は当方の文で受けている。"
             "文は写していないため計画は動かない"),
            ("同じく、正本に受け先がないもの",
             f"{kk.get(NASHI, 0)}件",
             "報告書だけの論点"),
            ("表2-8（主な課題）の行", f"{gyou}行",
             "回収数が変わっても動かない（区分は同じ）"),
            ("表2-8 主な調査結果欄の割合", f"{suuchi}点",
             "動く。行1〜7は他メンバー版、行8（専門的な支援）は"
             "当方が足したもの（ＡＫ-1）"),
            ("第2章7 の割合（他メンバー版の現物）", f"{len(kare)}件",
             "他メンバーが差し替える（Ｓ-39）"),
            ("第2章7 の割合（当方が足した・書き換えた）", f"{len(ware)}件",
             "うちアンケート由来で値が動く8件（ＡＫ-2〜ＡＫ-9）を当方が直す"),
            ("自由記述", "0件",
             "正本は自由記述を引いていない。Ｒ-4は報告書だけの論点"),
    ):
        write_row(ws, r, [nm, n, "", dou],
                  aligns=["left", "center", "left", "left"])
        ws.row_dimensions[r].height = 34
        r += 1
    r += 2

    ws.merge_cells(start_row=r, start_column=1, end_row=r + 4, end_column=4)
    c = ws.cell(row=r, column=1)
    c.value = (
        "※ 前回（令和８年10月21日）は「Ｓ-39で第２章７を回収40件に"
        "差し替えていただく」ことを先に置いていました。"
        f"今回、直し残し{len(FUBI)}件のうち{kk.get(ONAJI, 0)}件は"
        "同じ文が正本の第２章７にあることが分かりました。正本の文はいまは壊れていません（旧版の文です）が、"
        "壊れた報告書から写すと壊れた文が計画に入ります。"
        "このためＳ-40（直し残しの修正）をＳ-39（差し替え）より前に"
        "置きます。順序だけの違いですが、入れ替えると"
        "計画の本文に手が入ってしまいます。")
    c.alignment = Alignment(wrap_text=True, vertical="top")
    c.font = Font(name=FONT, size=9)
    return ws


def sheet_houshin(wb):
    ws = add_sheet(
        wb, "01_修正方針",
        "どの形で直していただくか（3案）",
        "直し残し11件はいずれも機械的に直せるものです。"
        "指摘だけをお返しすると直し方が伝わらず往復が増えるため、"
        "置換文をお渡しする案②を採ります。"
        "当方が報告書そのものを直すことはしません。",
        [8, 30, 12, 50, 62])
    style_header_row(ws, 5, ["", "案", "採否", "内容", "理由"])
    r = 6
    for i, (no, na, sai, naka, riyu) in enumerate(HOUSHIN):
        write_row(ws, r, [no, na, sai, naka, riyu],
                  alt=(i % 2 == 1),
                  aligns=["center", "left", "center", "left", "left"])
        ws.cell(row=r, column=3).fill = PatternFill(
            "solid", fgColor=("E2EFDA" if sai == SAIYO else "F2F2F2"))
        ws.row_dimensions[r].height = 56
        r += 1
    r += 2
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    style_title(ws.cell(row=r, column=1), "置換後をどこから取るか",
                fill=COLORS["subhead"], size=11)
    r += 1
    style_header_row(ws, r, ["", "取り方", "", "件数", "考え方"])
    r += 1
    shin = "\n".join(paras(SHIN))
    kyu = "\n".join(paras(KYU))
    hoka = sum(1 for no in CHIKAN if CHIKAN[no]["ato"] in shin)
    moto = sum(1 for no in CHIKAN
               if CHIKAN[no]["ato"] not in shin and CHIKAN[no]["ato"] in kyu)
    jisaku = len(CHIKAN) - hoka - moto
    for na, n, kangae in (
            ("報告書の別の箇所にある正しい形", hoka,
             "同じ文が報告書の中に正しい形で存在する。"
             "これがいちばん確かである"),
            ("旧版（10月21日版）の文", moto,
             "改訂の前の形。改訂の意図が読める場合は意図を残した形を"
             "置換後とし、旧版の形を別案に置く"),
            ("当方が作った形", jisaku,
             "上の二つがいずれも無い場合に限る。"
             "理由を備考に書いている"),
    ):
        write_row(ws, r, [na, "", "", f"{n}件", kangae],
                  aligns=["left", "left", "left", "center", "left"])
        ws.row_dimensions[r].height = 32
        r += 1
    return ws


def sheet_chikan(wb):
    shin = "\n".join(paras(SHIN))
    ws = add_sheet(
        wb, "02_置換文",
        f"そのまま適用できる形の置換文 {len(CHIKAN)}件",
        "「置換前」は報告書（令和8年10月22日版）の現物です。"
        "生成のたびに、置換前が報告書にちょうど1つあることと、"
        "置換後が新たな不備を生まないことを確かめています。"
        "別案がある箇所は、どちらでも差し支えありません。",
        [8, 24, 50, 50, 46, 40])
    style_header_row(ws, 5, ["番号", "箇所", "置換前（報告書の現物）",
                             "置換後（当方の案）", "別案", "備考"])
    r = 6
    fubi = {x[0]: x for x in FUBI}
    for i, no in enumerate(sorted(CHIKAN, key=lambda s: int(s.split("-")[1]))):
        rec = CHIKAN[no]
        basho = fubi[no][1]
        onaji = "報告書の別の箇所にある形" if rec["ato"] in shin else ""
        memo = rec["memo"] + (("／" + onaji) if onaji else "")
        write_row(ws, r, [no, basho, rec["mae"], rec["ato"],
                          rec["betsu"] or "―", memo],
                  alt=(i % 2 == 1),
                  aligns=["center", "left", "left", "left", "left", "left"])
        ws.cell(row=r, column=3).fill = PatternFill("solid", fgColor="FCE4D6")
        ws.cell(row=r, column=4).fill = PatternFill("solid", fgColor="E2EFDA")
        ws.row_dimensions[r].height = 62
        r += 1
    return ws


def sheet_review(wb):
    jou = _kazu_jou()
    rows = honbun(SHIN)
    moto = {x[0]: x for x in SHITEKI}
    ws = add_sheet(
        wb, "03_レビュー指摘の反映状況",
        f"令和8年10月5日のレビュー {len(RJOUKYOU)}件",
        f"解消{jou.get(KAISHO, 0)}件・一部対応{jou.get(ICHIBU, 0)}件・"
        f"未対応{jou.get(MITAIOU, 0)}件。"
        "状態は報告書の現物を照合して出しており、書き写していません。"
        "指摘の本文は当方のレビューのブックから読み込んでいます。",
        [8, 10, 14, 50, 14, 10, 24, 50])
    style_header_row(ws, 5, ["番号", "重み", "観点", "指摘",
                             "状態", "照合", "どこで受けるか", "備考"])
    r = 6
    for i, (no, keys, jotai, uke, memo) in enumerate(RJOUKYOU):
        src = moto[no]
        hit = terasu(rows, keys)
        write_row(ws, r, [no, src[1], src[2], src[3], jotai,
                          (hit[0] if hit else "無"), uke, memo],
                  alt=(i % 2 == 1),
                  aligns=["center", "center", "left", "left",
                          "center", "center", "left", "left"])
        iro = {KAISHO: "E2EFDA", ICHIBU: "FFF2CC", MITAIOU: "FCE4D6"}[jotai]
        ws.cell(row=r, column=5).fill = PatternFill("solid", fgColor=iro)
        ws.row_dimensions[r].height = 56
        r += 1
    return ws


def sheet_jiyu(wb):
    ws = add_sheet(
        wb, "04_自由記述の要約",
        f"公表前に要約の程度を村と協議する記述 {len(JIYU)}件",
        "本村の人口は約2,400人、障がい者調査の有効回答は40件です。"
        "状態像をそのまま書くと村内で特定され得ます。"
        "要約の程度は村とご協議いただく事項であり、当方が決めません"
        "（Ｍ-132）。下の要約案は協議の材料としてお使いください。",
        [50, 56, 50])
    style_header_row(ws, 5, ["いまの記述（表7）", "要約案（村と協議のうえ）",
                             "なぜ特定され得るか"])
    r = 6
    for i, (ima, an, naze) in enumerate(JIYU):
        write_row(ws, r, [ima, an, naze],
                  alt=(i % 2 == 1), aligns=["left", "left", "left"])
        ws.cell(row=r, column=1).fill = PatternFill("solid", fgColor="FCE4D6")
        ws.row_dimensions[r].height = 46
        r += 1
    return ws


def _kazu_uke_kata():
    k = {}
    for no in UKEKATA:
        k[UKEKATA[no][1]] = k.get(UKEKATA[no][1], 0) + 1
    return k


def sheet_eikyou(wb):
    rows = honbun(HONPON) if os.path.exists(HONPON) else []
    u = utsushi(rows) if rows else {}
    fubi = {x[0]: x for x in FUBI}
    k = _kazu_uke_kata()
    ws = add_sheet(
        wb, "05_計画への影響範囲",
        f"直し残し{len(FUBI)}件を正本がどう受けているか",
        f"同じ文がある{k.get(ONAJI, 0)}件／"
        f"趣旨を受けている{k.get(YOUYAKU, 0)}件／"
        f"受け先がない{k.get(NASHI, 0)}件。"
        "正本の文はいずれも旧版（壊れる前）の形であり、"
        "いまは壊れていません。"
        "手がかりは数値を含まない語にしています"
        "（回収数が違うため数値では当たりません）。",
        [8, 24, 28, 30, 14, 46])
    style_header_row(ws, 5, ["番号", "箇所（報告書）", "受け方",
                             "正本での箇所", "照合", "どうするか"])
    r = 6
    dou = {
        ONAJI: "差し替えの対象。壊れた報告書から写すと"
               "壊れた文が計画に入る（Ｓ-40をＳ-39より前に置く）",
        YOUYAKU: "文は写していないため、報告書の直しが"
                 "そのまま計画に及ぶことはない",
        NASHI: "報告書だけの論点。計画に受け先がないため"
               "計画は動かない",
    }
    for i, no in enumerate(sorted(u, key=lambda s: int(s.split("-")[1]))):
        keys, hit = u[no]
        kata, basho = UKEKATA[no][1], UKEKATA[no][2]
        write_row(ws, r, [no, fubi[no][1], kata, basho,
                          (hit[0] if hit else "無"), dou[kata]],
                  alt=(i % 2 == 1),
                  aligns=["center", "left", "left", "left",
                          "center", "left"])
        iro = {ONAJI: "FCE4D6", YOUYAKU: "DDEBF7", NASHI: "F2F2F2"}[kata]
        ws.cell(row=r, column=3).fill = PatternFill("solid", fgColor=iro)
        ws.row_dimensions[r].height = 40
        r += 1
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    style_title(ws.cell(row=r, column=1),
                "回収数40件で動くもの（誰が直すか）",
                fill=COLORS["subhead"], size=11)
    r += 1
    style_header_row(ws, r, ["", "件数", "", "", "", "誰が直すか"])
    r += 1
    kare, ware = wakeru()
    gyou, suuchi = hyou28(HONPON) if os.path.exists(HONPON) else (0, 0)
    for nm, n, dare in (
            ("第2章7 の割合（他メンバー版の現物）", f"{len(kare)}件",
             "他メンバーが差し替える（Ｓ-39）"),
            ("第2章7 の割合（当方が足した・書き換えた）", f"{len(ware)}件",
             "うちアンケート由来で値が動く8件（ＡＫ-2〜ＡＫ-9）を当方"),
            ("表2-8（主な課題）の行", f"{gyou}行",
             "動かない（課題の区分は回収数によらない）"),
            ("表2-8 主な調査結果欄の割合", f"{suuchi}点",
             "行1〜7は他メンバー版、行8（専門的な支援）は当方（ＡＫ-1）"),
            ("自由記述", "0件",
             "正本は自由記述を引いていない。Ｒ-4は報告書だけの論点"),
    ):
        write_row(ws, r, [nm, n, "", "", "", dare],
                  aligns=["left", "center", "left", "left", "left", "left"])
        ws.row_dimensions[r].height = 26
        r += 1
    return ws


def sheet_dandori(wb):
    ws = add_sheet(
        wb, "06_段取り",
        f"直す順序 {len(DANDORI)}段",
        "順序が前回と変わっています。"
        "Ｓ-40（直し残しの修正）がＳ-39（第2章7の差し替え）より"
        "前に来ます。",
        [8, 12, 14, 52, 22, 56])
    style_header_row(ws, 5, ["順", "番号", "相手", "すること", "期限",
                             "なぜその順か"])
    r = 6
    for i, (jun, no, aite, suru, kigen, naze) in enumerate(DANDORI):
        write_row(ws, r, [str(jun), no, aite, suru, kigen, naze],
                  alt=(i % 2 == 1),
                  aligns=["center", "center", "center", "left",
                          "center", "left"])
        ws.row_dimensions[r].height = 46
        r += 1
    return ws


# ===========================================================================
# 8. 照会
# ===========================================================================
SHOUKAI = [
    ("Ｓ-42", "他メンバー",
     "当方のレビュー指摘20件のうち手の付いていないものの扱い",
     "令和8年10月5日に当方からお出ししたレビューについて、"
     "令和8年10月22日版での反映の状況を1件ずつ確かめました"
     "（「北塩原村_報告書の修正方針と計画への影響.xlsx」の03シート）。"
     "すべてを報告書に入れる必要はないと考えております。"
     "報告書は調査の分析であり、計画の施策を書く場ではないためです。"
     "当方の整理では、報告書に入れていただきたいものが次の6項目です。"
     "①前回調査（発送164人・回収95人・回収率57.9％）との対比と、"
     "今回23.3％まで下がったことの読み方（Ｒ-1・Ｒ-2）。"
     "結果の読み方の根幹であり、委員会資料には要ると考えます。"
     "②障がい児の移動支援を課題5の方向性に加える（Ｒ-12。"
     "課題2には加えていただきました）"
     "③複数回答を区分とするクロスでは1人が複数の区分に計上され"
     "合計が全体を超える旨の注記（Ｒ-13）"
     "④第3章の本文の「100.0％」に実数を併記する（Ｒ-14。20か所）"
     "⑤表5の「特にない」「無回答」行の差（Ｒ-17）。"
     "「特にない」は45.0％から30.0％へ15.0ポイント下がっており、"
     "利用意向の広がりを示す情報です。"
     "⑥基数5未満の区分（65件中28件）への記号と表ごとの脚注（Ｒ-19）。"
     "このほかは当方の計画素案で受けておりますので、"
     "報告書は表6-1の脚注で足ります。"
     "二重に書かないよう、どこまで報告書に入れるかを"
     "先に合わせていただけますでしょうか。",
     "令和８年10月31日（金）"),
    ("Ｍ-132", "村",
     "アンケート調査結果報告書の自由記述の要約の程度",
     "アンケート調査結果報告書の自由記述の要旨（表7）に、"
     "本村の規模では個人が特定され得る記述が残っております。"
     "本村の人口は約2,400人、障がい者調査の有効回答は40件であり、"
     "経管栄養や手話の利用といった状態像は村内で極めて限られます。"
     "報告書は策定委員会にお出しし公表の対象となりますので、"
     "どの程度まで要約するかをご相談させてください。"
     "当方の要約案を「北塩原村_報告書の修正方針と計画への影響.xlsx」の"
     "04シートにお示ししております。"
     "分野別の要約にとどめる形でよろしいか、"
     "それとも個々の声を残す形をお望みかをお聞かせください。"
     "なお、計画素案の本文は自由記述を引いておりませんので、"
     "この論点は報告書に限られます。",
     "令和８年10月31日（金）"),
]


# ===========================================================================
# 9. 自己点検
# ===========================================================================
def verify():
    ng = []
    for p in (KYU, SHIN, ZENKAI_PDF, SHUKEI):
        if not os.path.exists(p):
            raise SystemExit(f"資料が見つかりません: {p}")

    shin = "\n".join(paras(SHIN))
    kyu = "\n".join(paras(KYU))
    shin_all = "\n".join(t for _l, t in honbun(SHIN))

    # ① 置換文が直し残しの一覧と1対1であること
    fubi_no = {x[0] for x in FUBI}
    if set(CHIKAN) != fubi_no:
        ng.append(f"置換文と直し残しが1対1でない: "
                  f"{sorted(set(CHIKAN) ^ fubi_no)}")

    # ② 置換前が壊れている語を含み、報告書にちょうど1つあること。
    #    置換後が置換前と違い、報告書の現物にある形か旧版の形か、
    #    いずれでもなければ備考に理由があること。
    fubi = {x[0]: x for x in FUBI}
    for no, rec in CHIKAN.items():
        bad = fubi[no][3]
        if bad not in rec["mae"]:
            ng.append(f"{no} 置換前が壊れている語を含まない: {bad}")
        n = shin.count(rec["mae"])
        if n != 1:
            ng.append(f"{no} 置換前が報告書に{n}件（1件でない）: "
                      f"{rec['mae'][:40]}")
        if rec["ato"] == rec["mae"]:
            ng.append(f"{no} 置換後が置換前と同じ")
        if bad in rec["ato"]:
            ng.append(f"{no} 置換後に壊れている語が残っている: {bad}")
        if rec["betsu"] and bad in rec["betsu"]:
            ng.append(f"{no} 別案に壊れている語が残っている: {bad}")
        if rec["ato"] not in shin and rec["ato"] not in kyu \
                and not rec["memo"].strip():
            ng.append(f"{no} 当方が作った形なのに理由が空")

    # ③ 置換後・別案が新たな不備を生んでいないこと
    patt = [("読点の重複", r"、、"), ("句点＋読点", r"。、"),
            ("助詞の重なり", r"の(に基づく|による)"),
            ("連体形＋読点＋とともに", r"(せる|する|める|げる)、とともに"),
            ("半角かっこ", r"[^\x00-\x7F]\([^)]*\)"),
            ("述語の二重", r"(あり|です|ます)となって")]
    for no, rec in CHIKAN.items():
        for bun in (rec["ato"], rec["betsu"]):
            if not bun:
                continue
            for na, pa in patt:
                if re.search(pa, bun):
                    ng.append(f"{no} 置換後・別案に{na}がある: {bun[:40]}")

    # ④ Ｒ-4 の置換後の選択肢名が集計表に実在し、
    #    置換前にあるもう一方の語が集計表に無いこと。
    #    「どちらを残すか」を当方で決められる根拠である。
    sel = shukei_sentakushi("問29")
    if "移動支援" not in sel:
        ng.append("問29の選択肢に「移動支援」が無い（Ｒ-4の根拠が崩れる）")
    if "移動に関する支援" in sel:
        ng.append("問29の選択肢に「移動に関する支援」がある"
                  "（Ｒ-4はどちらを残すか確認が要る）")

    # ⑤ レビュー指摘の一覧が当方のレビューのブックと1対1であること。
    #    （Ｒ-21〜Ｒ-23 は適合の記録であり対象外）
    moto = {x[0]: x for x in SHITEKI}
    tekigou = {k for k, v in moto.items() if v[1] == OK}
    naka = {x[0] for x in RJOUKYOU}
    if naka != set(moto) - tekigou:
        ng.append(f"レビュー指摘の一覧が1対1でない: "
                  f"{sorted(naka ^ (set(moto) - tekigou))}")
    for no, _keys, _jotai, uke, memo in RJOUKYOU:
        if uke not in (HOUKOKU, SEIHON_UKE, RYOUHOU, MURA_KYOUGI):
            ng.append(f"{no} 受け先が未定義: {uke}")
        if not memo.strip():
            ng.append(f"{no} 備考が空")
        if moto[no][1] not in (HI, MID, LO):
            ng.append(f"{no} 重みが未定義: {moto[no][1]}")

    # ⑥ 状態と報告書の照合の結果が食い違わないこと。
    #    解消・一部対応としたものは語が現れ、
    #    未対応としたものは現れない。
    rows = honbun(SHIN)
    for no, keys, jotai, _uke, _memo in RJOUKYOU:
        aru = bool(terasu(rows, keys))
        if no in MUKI_GYAKU:
            #   語が残っていること自体が未対応の証拠であるもの
            if jotai == MITAIOU and not aru:
                ng.append(f"{no} 状態は「未対応」だが報告書に語が無い"
                          f"（残っていないなら解消である）: "
                          f"{'＋'.join(keys)}")
            if jotai in (KAISHO, ICHIBU) and aru:
                ng.append(f"{no} 状態は「{jotai}」だが語が残っている: "
                          f"{'＋'.join(keys)}")
            continue
        if jotai in (KAISHO, ICHIBU) and not aru:
            ng.append(f"{no} 状態は「{jotai}」だが報告書に語が無い: "
                      f"{'＋'.join(keys)}")
        if jotai == MITAIOU and aru:
            ng.append(f"{no} 状態は「未対応」だが報告書に語がある: "
                      f"{'＋'.join(keys)}")

    # ⑥b Ｒ-17 の前提 ― 表5の「特にない（各設問）」行の差の欄が
    #     実際に空であること。照合語（-15.0）だけでは、
    #     別の値で埋められた場合に気づけない。
    kara = hyou5_sa_kara()
    if kara is None:
        ng.append("表5（現在の利用と今後の利用意向）が報告書に無い")
    elif not kara:
        ng.append("表5の「特にない（各設問）」行の差が埋まっている"
                  "（Ｒ-17は解消である）")

    # ⑦ 前回調査の発送・回収・回収率が原典に実在すること
    z = zenkai_genten()
    for v in (str(ZENKAI["発送"]) + "人", str(ZENKAI["回収"]) + "人",
              "57.9"):
        if v not in z:
            ng.append(f"前回調査の値が原典に無い: {v}")
    #    低下の幅を書き写していないこと（原典と新版から出す）
    if "23.3" not in shin_all:
        ng.append("新版に有効回答率23.3％が無い（低下の幅が出せない）")

    # ⑧ 要約を要する記述が報告書の現物にあること。
    #    要約案に状態像が残っていないこと。
    for ima, an, naze in JIYU:
        if ima not in shin_all:
            ng.append(f"要約を要するとした記述が報告書に無い: {ima[:30]}")
        if not naze.strip():
            ng.append(f"特定され得る理由が空: {ima[:30]}")
    jotai_go = ("経管栄養", "手話通訳者を継続", "膝")
    for ima, an, _naze in JIYU:
        for w in jotai_go:
            if w in an:
                ng.append(f"要約案に状態像が残っている: {w}")

    # ⑨ 段取りが順に並び、Ｓ-40 がＳ-39 より前であること
    jun = [x[0] for x in DANDORI]
    if jun != sorted(jun):
        ng.append(f"段取りの順が並んでいない: {jun}")
    ban = {x[1]: x[0] for x in DANDORI}
    if ban.get("Ｓ-40", 99) >= ban.get("Ｓ-39", 0):
        ng.append("Ｓ-40（直し残しの修正）がＳ-39（差し替え）より"
                  "後になっている")

    # ⑩ 受け方の一覧が直し残しと1対1で、手がかりに数値が入っておらず、
    #    記録した受け方と正本の照合の結果が食い違わないこと
    if set(UKEKATA) != fubi_no:
        ng.append(f"受け方の一覧が直し残しと1対1でない: "
                  f"{sorted(set(UKEKATA) ^ fubi_no)}")
    for no, (keys, kata, basho) in UKEKATA.items():
        for k in keys:
            if re.search(r"\d", k):
                ng.append(f"{no} 手がかりに数値が入っている: {k}")
        if kata not in (ONAJI, YOUYAKU, NASHI):
            ng.append(f"{no} 受け方が未定義: {kata}")
        if not basho.strip():
            ng.append(f"{no} 正本での箇所が空")
    rows_h = honbun(HONPON) if os.path.exists(HONPON) else []
    if rows_h:
        for no, (keys, kata, _b) in UKEKATA.items():
            aru = bool(terasu(rows_h, keys))
            if kata in (ONAJI, YOUYAKU) and not aru:
                ng.append(f"{no} 受け方は「{kata}」だが正本に無い: "
                          f"{'＋'.join(keys)}")
            if kata == NASHI and aru:
                ng.append(f"{no} 受け方は「受け先がない」だが正本にある: "
                          f"{'＋'.join(keys)}")
        #   第2章7 に同じ文があるとしたものは、
        #   正本の文がいまは壊れていないこと（旧版の形であること）
        zen_h = "\n".join(t for _l, t in rows_h)
        for no in (n for n in UKEKATA if UKEKATA[n][1] == ONAJI):
            bad = fubi[no][3]
            if bad in zen_h:
                ng.append(f"{no} 壊れている語が正本に入っている: {bad}")
    else:
        print("  ※ 正本が無いため⑩の照合は省略")

    # ⑩b Ｓ-42 に列挙した番号が、「報告書に入れる」と判定したもの
    #     （解消を除く）と一致すること。
    #     文章と判定が離れると、方針を取り違えたままお送りしてしまう。
    s42 = [x for x in SHOUKAI if x[0] == "Ｓ-42"]
    if not s42:
        ng.append("Ｓ-42が無い")
    else:
        kaku = set(re.findall(r"Ｒ-(\d+)", s42[0][3]))
        beki = {no.split("-")[1] for no, _k, jo, u, _m in RJOUKYOU
                if u == HOUKOKU and jo != KAISHO}
        if kaku != beki:
            ng.append(f"Ｓ-42の列挙と「報告書に入れる」の判定が合わない: "
                      f"文章{sorted(kaku, key=int)} ≠ "
                      f"判定{sorted(beki, key=int)}")

    # ⑪ 照会の番号が重複せず、宛先と期限が入っていること
    sno = [x[0] for x in SHOUKAI]
    if len(sno) != len(set(sno)):
        ng.append(f"照会の番号に重複がある: {sno}")
    for no, ate, ken, naka, kigen in SHOUKAI:
        if ate not in ("村", "他メンバー"):
            ng.append(f"{no} 宛先が未定義: {ate}")
        if not ken.strip() or not naka.strip() or not kigen.strip():
            ng.append(f"{no} 件名・内容・期限のいずれかが空")

    # ⑫ 他団体の名が入っていないこと
    zen = "\n".join(str(v) for rec in HOUSHIN for v in rec)
    zen += "\n" + "\n".join(
        str(v) for rec in CHIKAN.values() for v in rec.values())
    zen += "\n" + "\n".join(
        str(v) for rec in RJOUKYOU for v in rec if isinstance(v, str))
    zen += "\n" + "\n".join(str(v) for rec in JIYU for v in rec)
    zen += "\n" + "\n".join(str(v) for rec in DANDORI for v in rec
                            if isinstance(v, str))
    zen += "\n" + "\n".join(str(v) for rec in SHOUKAI for v in rec)
    for hoka in ("小野町", "金ヶ崎", "金ケ崎", "阿蘇", "札幌", "川崎町"):
        if hoka in zen:
            ng.append(f"他団体の名が入っている: {hoka}")

    if ng:
        for x in ng:
            print(f"  × {x}")
        raise SystemExit("修正方針の自己点検で止めました")

    jou, uke = _kazu_jou(), _kazu_uke()
    kk = _kazu_uke_kata()
    gyou, suuchi = hyou28(HONPON) if os.path.exists(HONPON) else (0, 0)
    teika = round((ZENKAI["回収率"] * 100) - 23.3, 1)
    return (f"置換文 {len(CHIKAN)}件（報告書の別の箇所から"
            f"{sum(1 for n in CHIKAN if CHIKAN[n]['ato'] in shin)}件）／"
            f"レビュー指摘 解消{jou.get(KAISHO, 0)}・"
            f"一部{jou.get(ICHIBU, 0)}・未対応{jou.get(MITAIOU, 0)}／"
            f"報告書に入れる{uke.get(HOUKOKU, 0)}・"
            f"計画で受ける{uke.get(SEIHON_UKE, 0)}／"
            f"要約を要する記述 {len(JIYU)}件／"
            f"回収率 57.9→23.3％（{teika}ポイント低下）／"
            f"正本での受け方 同じ文{kk.get(ONAJI, 0)}・"
            f"趣旨{kk.get(YOUYAKU, 0)}・受け先なし{kk.get(NASHI, 0)}／"
            f"表2-8 {gyou}行・割合{suuchi}点")


def main():
    ensure_out_dir()
    msg = verify()
    wb = Workbook()
    wb.remove(wb.active)
    sheet_ketsuron(wb)
    sheet_houshin(wb)
    sheet_chikan(wb)
    sheet_review(wb)
    sheet_jiyu(wb)
    sheet_eikyou(wb)
    sheet_dandori(wb)
    wb.save(OUT_FILE)
    print(f"  自己検証: {msg}")
    print(f"作成: {OUT_FILE}")
    print(f"  シート{len(wb.sheetnames)}／置換文{len(CHIKAN)}件／"
          f"照会{len(SHOUKAI)}件")


if __name__ == "__main__":
    sys.exit(main())
