# -*- coding: utf-8 -*-
"""
北塩原村　村への照会文（6通）

出力: output/北塩原村_村への照会_6通.docx

何のために作るものか
  令和8年10月10日までの各点検・レビューで立てた村への照会は66件に
  なったが、村に出す文書そのものは1通も作っていなかった。
  照会を出さない限り、見込量も成果目標も委員会資料も確定しない。
  令和8年10月15日（木）の打合せで手渡せる形にする。

便の分け方（令和8年11月の委員会から逆算する）
  66件を1通にすると回答が遅れる。
  回答がいつ要るかで6通に分ける。
    第1便 委員会（令和8年11月）に出す資料に直に要るもの
    第2便 見込量の確定に要るもの（第2次算定の入力）
    第3便 高齢者施策との接続（対象となるかどうか）
    第4便 児童福祉との接続
    第5便 既存施策の実績（計画の確定には間に合わなくてよい）
    第6便 会津北部4町村及び県との広域整合

  Ｍ-40（工程の確認）は文書にせず電話で確かめる。
  工程が分からないと、この6通の期限そのものが決められないため。

単一出所
  照会の中身は各点検・レビューの生成器に置いてある。
  本書はそれを読み込んで文書に組むだけで、文言を書き足さない。
  番号の抜け・重複・分類漏れは自己点検で落ちる。
"""

import os
import re
import sys

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_kitashiobara_kaigi import (  # noqa: E402
    h2, new_doc, note, para, table,
)
from build_kitashiobara_murashiryo import SHOUKAI as S_MURA  # noqa: E402
from build_kitashiobara_juten12 import SHOUKAI as S_JUTEN  # noqa: E402
from build_kitashiobara_hp_shisaku import SHOUKAI as S_HP  # noqa: E402
from build_kitashiobara_hp_3bunya import SHOUKAI as S_HP3  # noqa: E402
from build_kitashiobara_redteam_kyukyu import SHOUKAI as S_RT  # noqa: E402
from build_kitashiobara_mece_houshu import SHOUKAI as S_ME  # noqa: E402
from build_kitashiobara_mikomi_redteam import SHOUKAI as S_SA  # noqa: E402
from build_kitashiobara_zenkai_hyouka import SHOUKAI as S_ZEN  # noqa: E402
from build_kitashiobara_zuhyo import KAKUNIN as K_ZUHYO  # noqa: E402
from build_kitashiobara_zuhyo_bangou import KAKUNIN as K_BAN  # noqa: E402
from build_kitashiobara_anke_mikomi import KAKUNIN as K_ANKE  # noqa: E402
from build_kitashiobara_kaigo_seigo import SHOUKAI as S_KG  # noqa: E402
from build_kitashiobara_kouiki_redteam import SHOUKAI as S_KO  # noqa: E402
from build_kitashiobara_jouhou_kouhyou import SHOUKAI as S_KJ  # noqa: E402
from build_kitashiobara_yosan_mece import SHOUKAI as S_YS  # noqa: E402
from build_kitashiobara_kyoseigata import SHOUKAI as S_KY  # noqa: E402
from build_kitashiobara_kojino import SHOUKAI as S_KN  # noqa: E402
from build_kitashiobara_chiikishien import SHOUKAI as S_CS  # noqa: E402
from build_kitashiobara_ishi_kettei import KAKUNIN as K_IK  # noqa: E402
from build_kitashiobara_review2_taiou import (  # noqa: E402
    KAKUNIN as K_R2,
)
from build_kitashiobara_houteikisai import (  # noqa: E402
    KAKUNIN as K_HT, SHOUKAI as S_HT,
)
from build_kitashiobara_shihyou_sokutei import (  # noqa: E402
    SHOUKAI as S_SS,
)
from build_kitashiobara_progress import (  # noqa: E402
    SHOUKAI as S_PR,
)
from build_kitashiobara_review3_taiou import (  # noqa: E402
    SHOUKAI as S_RV,
)
from build_kitashiobara_houkaisei_r8 import (  # noqa: E402
    SHOUKAI as S_HK,
)
from build_kitashiobara_uchiawase_ronten import (  # noqa: E402
    SHOUKAI as S_UR,
)
from build_kitashiobara_kanren_hourei import (  # noqa: E402
    SHOUKAI as S_KH,
)
from build_kitashiobara_taikei_kousei import (  # noqa: E402
    SHOUKAI as S_TK,
)
from build_kitashiobara_taikei_hanei import (  # noqa: E402
    SHOUKAI as S_TH,
)
from build_kitashiobara_shogaiji_ronten import (  # noqa: E402
    SHOUKAI as S_SJ,
)
from build_kitashiobara_gaiyou_2hon import (  # noqa: E402
    SHOUKAI as S_G2,
)
from build_kitashiobara_anke_kaitei import (  # noqa: E402
    SHOUKAI as S_AK,
)
from build_kitashiobara_ken_iryo import (  # noqa: E402
    SHOUKAI as S_KI,
)
from build_kitashiobara_houkoku_hanei import (  # noqa: E402
    SHOUKAI as S_HR,
)
from build_kitashiobara_houkoku_houshin import (  # noqa: E402
    SHOUKAI as S_HS,
)
from build_kitashiobara_anke_kaishuu import (  # noqa: E402
    SHOUKAI as _S_KS,
)

#   回収数の変更の照会のうち、村あてのＭ-128 は解消の記録であり、
#   台帳の現物は build_kitashiobara_anke_kaitei.py の側にある。
#   二重に登録すると宛先が「村（解消）」で上書きされ、
#   村あての番号に抜けができる。他メンバーあてのものだけを取る。
S_KS = [r for r in _S_KS if r[1] == "他メンバー"]


def _kubun():
    """確認事項の区分（Ａ〜Ｄ）を読み込む。

    Ａ（解決済み）とＢ（当方で確認）は相手にお尋ねしない。
    文書に載せるのはＣとＤだけにし、ＡとＢは「当方で確認する分」として
    末尾に一覧で掲げる（ご回答不要）。
    割付そのものは build_kitashiobara_kakunin_kubun.py に置いてある。
    """
    import importlib
    m = importlib.import_module("build_kitashiobara_kakunin_kubun")
    return {r[0]: (r[1], r[2], r[3]) for r in m.KUBUN}, m.A, m.B


KUBUN_MAP, KU_A, KU_B = _kubun()


def tazuneru(nos):
    """便・節の番号のうち、相手にお尋ねするもの（Ｃ・Ｄ）だけを返す。"""
    return [no for no in nos
            if KUBUN_MAP.get(no, ("", "", ""))[0] not in (KU_A, KU_B)]


def jibun(nos):
    """便・節の番号のうち、当方で確認するもの（Ａ・Ｂ）を返す。"""
    return [no for no in nos
            if KUBUN_MAP.get(no, ("", "", ""))[0] in (KU_A, KU_B)]
# 図表の点数は番号一覧の定義から取る。数を直書きすると、
# 図表が増減したときに照会文だけが古くなる（Ｓ-21で一度そうなった）。
from build_kitashiobara_zuhyo_bangou import (  # noqa: E402
    HYO as _HYO, ZU as _ZU,
)

REPO_ROOT = os.path.dirname(os.path.abspath(__file__))
OUT_FILE = f"{REPO_ROOT}/output/北塩原村_村への照会_6通.docx"
OUT_MEMBER = (f"{REPO_ROOT}/output/北塩原村_他メンバーへの確認_1通.docx")

UCHIAWASE = "令和８年10月15日（木）"
IINKAI = "令和８年11月"


# ===========================================================================
# 1. 照会を集めて形を揃える
#    古い3本は (番号, 件名, 内容, 使い道, 期限)
#    新しい5本は (番号, 宛先, 件名, 内容, 期限)
# ===========================================================================
def _zen(no):
    """Ｍ-40 と Ｍ-40 が混ざっているため全角に揃える。"""
    return no.replace("M-", "Ｍ-").replace("S-", "Ｓ-")


def _su(n):
    """本文の数字。1桁は全角、2桁以上は半角（表記の作法）。"""
    n = int(n)
    if n < 10:
        return "０１２３４５６７８９"[n]
    return str(n)


def atsumeru():
    out = {}
    for recs in (S_MURA, S_JUTEN, S_HP):
        for r in recs:
            out[_zen(r[0])] = dict(no=_zen(r[0]), atesaki="村", kenmei=r[1],
                                   naiyo=r[2], tsukaimichi=r[3], kigen=r[4])
    for recs in (S_HP3, S_RT, S_ME, S_SA, S_ZEN, S_KG, S_KO, S_KJ,
                 S_YS, S_KY, S_KN, S_CS, K_IK, S_HT, K_HT,
                 S_SS, S_PR, S_RV, S_HK, S_UR, S_KH, S_TK, S_TH,
                 S_SJ, S_G2, S_AK, S_KI, S_HR, S_HS, S_KS,
                 K_ZUHYO, K_BAN, K_ANKE):
        for r in recs:
            out[_zen(r[0])] = dict(no=_zen(r[0]), atesaki=r[1], kenmei=r[2],
                                   naiyo=r[3], tsukaimichi="", kigen=r[4])
    return out


# ===========================================================================
# 2. 便の分け方（番号を明示する。漏れたら自己点検で落ちる）
# ===========================================================================
BIN = [
    dict(
        no=1,
        midashi="委員会に出す資料に直に要るもの",
        kigen="令和８年10月24日（金）",
        riyu=f"{IINKAI}の委員会に、前回計画の評価と計画の素案を出します。"
             "評価の表のうち、いまの値が分からない欄が５つあります。"
             "ここが埋まらないと、基準値から目標までの進み具合を"
             "お示しできません。"
             "計画の名称と数値の食い違いも、資料に載せる前に"
             "確かめておく必要があります。"
             "あわせて、計画策定の工程（Ｍ-125）を"
             "この便に入れています。"
             "策定委員会の開催が11月となったことにより、"
             "策定委員会資料のスケジュール表のままでは"
             "成果品の履行期限（令和９年３月31日）を"
             "約１週間超えることになります。"
             "工程は委員会の資料にも載るものですので、"
             "委員会の前に決めておく必要があります。"
             "また、第８次福島県医療計画（令和６年３月）を拝見し、"
             "本計画との調和を図るべき事項を整理しました（Ｍ-129）。"
             "国の基本指針は、障害福祉計画を医療計画と調和が"
             "保たれたものとすることを求めており（第三の一の７）、"
             "精神障害にも対応した地域包括ケアシステムの構築の目標に"
             "ついては特に医療計画との関係に留意することとしています"
             "（第二の二）。手元に無い章の御提供と、"
             "村内の医療資源についてお伺いします。",
        nos=["Ｍ-41", "Ｍ-42", "Ｍ-43", "Ｍ-44", "Ｍ-69", "Ｍ-70",
             "Ｍ-95", "Ｍ-96", "Ｍ-104", "Ｍ-110", "Ｍ-113",
             "Ｍ-123", "Ｍ-124", "Ｍ-125", "Ｍ-129"],
    ),
    dict(
        no=2,
        midashi="サービスの見込量を確定するために要るもの",
        kigen="令和８年10月31日（金）",
        riyu="見込量は、過去の実績を延ばすのではなく、"
             "お一人ずつの状況を積み上げて定めます。"
             "とくに、重度訪問介護・同行援護・行動援護を"
             "令和９年度から令和11年度まで０人としている根拠は、"
             "いまは「利用の希望がない」ことだけです。"
             "対象となる状態の方がいらっしゃるのに"
             "制度を知らずに申し出ていない場合と、"
             "本当に必要がない場合とを分けられていません。",
        nos=["Ｍ-87", "Ｍ-88", "Ｍ-89", "Ｍ-90", "Ｍ-91",
             "Ｍ-92", "Ｍ-93", "Ｍ-94", "Ｍ-97", "Ｍ-109",
             "Ｍ-115", "Ｍ-118", "Ｍ-130", "Ｍ-132"],
    ),
    dict(
        no=3,
        midashi="高齢者福祉の施策との接続",
        kigen="令和８年11月14日（金）",
        riyu="障がいのある方の暮らしを支える仕組みは、"
             "障がい福祉の制度の中だけにあるわけではありません。"
             "村の高齢者福祉の事業に障がいのある方が含まれるのであれば、"
             "新しい事業を起こす前にそちらで対応できます。"
             "逆に含まれないのであれば、そこが本計画で埋めるべき"
             "すき間になります。",
        nos=["Ｍ-45", "Ｍ-46", "Ｍ-47", "Ｍ-48", "Ｍ-49", "Ｍ-50",
             "Ｍ-51", "Ｍ-52", "Ｍ-53", "Ｍ-54", "Ｍ-55", "Ｍ-56",
             "Ｍ-57", "Ｍ-58", "Ｍ-64", "Ｍ-65", "Ｍ-71", "Ｍ-72",
             "Ｍ-73", "Ｍ-81", "Ｍ-112", "Ｍ-116",
             "Ｍ-119", "Ｍ-120", "Ｍ-126", "Ｍ-128"],
    ),
    dict(
        no=4,
        midashi="児童福祉の施策との接続",
        kigen="令和８年11月14日（金）",
        riyu="本村には短期入所の事業所がありません。"
             "保護者の方が病気になられたときや出産のときに、"
             "障がいのある子どもをどこで預かるかは、"
             "障がい福祉の制度だけでは答えが出ません。"
             "村が令和８年度から始められた子育て短期支援事業で"
             "受け入れていただけるかどうかが、"
             "家族の支援の実際の分かれ目になります。",
        nos=["Ｍ-74", "Ｍ-75", "Ｍ-76", "Ｍ-77"],
    ),
    dict(
        no=5,
        midashi="既にある施策の実績",
        kigen="令和８年11月28日（金）",
        riyu="既にある施策については、"
             "「あること」ではなく「使われていること」を書きます。"
             "実績が分かれば、計画の評価と次の３年間の目標を"
             "数字で示せます。"
             "村内の事業所についても同じで、"
             "何カ所あるかではなく、"
             "どの指定の形で提供しているか"
             "（通常の指定か、共生型の指定か、"
             "基準該当障害福祉サービスとしての取扱いか）が分からないと、"
             "村内の介護保険の事業所の力を障がい福祉の供給に"
             "生かせるかどうかを判断できません。"
             "介護保険の地域支援事業についても同じです。"
             "国の実施要綱は任意事業の対象者を"
             "「被保険者、要介護被保険者を現に介護する者"
             "その他個々の事業の対象者として市町村が認める者」"
             "としており、"
             "障がいのある方を対象に含めるかどうかは"
             "村の要綱の定め方によります。"
             "Ｍ-50・Ｍ-71・Ｍ-81では"
             "「障がいのある方が対象となるか」と"
             "お尋ねしていましたが、"
             "お尋ねの仕方を改めます（Ｍ-108）。"
             "計画の確定にすぐ要るものではありませんので、"
             "他の便より後で差し支えありません。",
        nos=["Ｍ-59", "Ｍ-60", "Ｍ-61", "Ｍ-62", "Ｍ-63", "Ｍ-66",
             "Ｍ-67", "Ｍ-68", "Ｍ-78", "Ｍ-79", "Ｍ-80",
             "Ｍ-82", "Ｍ-83", "Ｍ-84", "Ｍ-85", "Ｍ-86",
             "Ｍ-103", "Ｍ-105", "Ｍ-106", "Ｍ-108", "Ｍ-111",
             "Ｍ-114", "Ｍ-121", "Ｍ-122", "Ｍ-127", "Ｍ-131"],
    ),
    dict(
        no=6,
        midashi="会津北部４町村及び県との広域整合",
        kigen="令和８年12月12日（金）",
        riyu="本計画は、地域生活支援拠点、基幹相談支援センター、"
             "医療的ケア児等支援、インクルージョン推進及び"
             "強度行動障害支援を、猪苗代町、磐梯町及び湯川村との"
             "広域で確保することとしています。"
             "しかし、本村の計画に「広域で行う」と書いても、"
             "相手方の計画で同じ連携単位、同じ会議体、"
             "同じ数え方になっていなければ、"
             "広域で確保したことになりません。"
             "とくに医療的ケア児等支援については、"
             "磐梯町及び湯川村の現行計画が協議の場を"
             "「圏域１カ所」としており、"
             "この「圏域」が会津北部の４町村を指すのか"
             "会津障がい保健福祉圏域を指すのかが分かりません。"
             "単位が違えば、設置すべき会議体そのものが変わります。"
             "３町村と県への照会を挟むため、"
             "他の便より期限を遅く置いています。"
             "ただしＭ-101（成年後見センターの構成市町村）は"
             "村内で確かめられるものですので、"
             "分かり次第お知らせいただければ助かります。"
             "また、高次脳機能障害者支援法が令和８年４月１日に"
             "施行されたことに伴い、"
             "県の支援センターの指定、地域協議会への参画、"
             "及び実施状況の公表の様式について"
             "県に確かめていただく事項を加えています"
             "（Ｍ-107）。"
             "このうちＭ-107の⑤と⑥は村内で分かるものです。",
        nos=["Ｍ-98", "Ｍ-99", "Ｍ-100", "Ｍ-101", "Ｍ-102",
             "Ｍ-107", "Ｍ-117"],
    ),
]

# 文書にせず電話で確かめるもの
DENWA = ["Ｍ-40"]

# 村以外に出すもの（村の照会文には入れない。他メンバーへの確認）
HOKA = ["Ｓ-19", "Ｓ-20", "Ｓ-21", "Ｓ-22", "Ｓ-23", "Ｓ-24",
        "Ｓ-25", "Ｓ-26", "Ｓ-27", "Ｓ-28", "Ｓ-29", "Ｓ-30",
        "Ｓ-31", "Ｓ-32", "Ｓ-33", "Ｓ-34", "Ｓ-35", "Ｓ-36",
        "Ｓ-37", "Ｓ-38", "Ｓ-39", "Ｓ-40", "Ｓ-41", "Ｓ-42"]

# 県あてだが他メンバーを通じてお願いするもの
KEN_KEIYU = ["Ｓ-11", "Ｓ-12"]

# 法令の原典の入手をお願いするもの（当方は取得できない）
GENTEN_IRAI = ["Ｓ-28", "Ｓ-31", "Ｓ-35"]


# ===========================================================================
# 2b. 他メンバーへの確認（1通の中の節）
#     村の照会と混ぜない。宛先が違うものを1通にすると、
#     どちらが答えるのかが分からなくなる。
# ===========================================================================
MEMBER_SETSU = [
    dict(
        no=1,
        midashi="図表について",
        kigen="令和８年10月16日（金）",
        riyu="図はカラーで進めることとし、"
             f"令和８年10月10日に当方で図{len(_ZU)}点・"
             f"表{len(_HYO)}点に番号と表題を付け、"
             "目次の後に図表目次を置きました。"
             "これに伴って、既存の図についてお願いしたいことが"
             "６件あります。"
             "いずれも最終版の体裁に関わるもので、"
             "村の回答を待たずに決められるものです。",
        nos=["Ｓ-19", "Ｓ-20", "Ｓ-21", "Ｓ-22", "Ｓ-23", "Ｓ-24",
             "Ｓ-30"],
    ),
    dict(
        no=2,
        midashi="制度名称の表記の統一と策定経過の日付",
        kigen="令和８年10月16日（金）",
        riyu="計画の冒頭の注記は、法律用語や施設名等の固有名称を除き"
             "「害」の字をひらがなで表記するとしています。"
             "ところが制度の名称について、"
             "法令の表記と村の表記が混在している箇所があります。"
             "同じ並びの中で食い違っている箇所もあるため、"
             "どちらに揃えるかを決めていただきたく存じます。"
             "当方の追記は決まった方に合わせます。"
             "決まり次第、置換の一覧を作ってお送りします。"
             "あわせて、第１章５（４）のパブリックコメントの記述の"
             "時制について確かめさせていただきたく存じます。"
             "日程は工程の引き直し（令和８年10月14日）により"
             "令和９年１月26日から２月９日まで（15日間）へ"
             "書き換えておりますが、"
             "「意見を募りました」と過去形で書かれている点は"
             "そのままにしております。",
        nos=["Ｓ-25", "Ｓ-27"],
    ),
    dict(
        no=3,
        midashi="アンケート分析報告書のクロス集計",
        kigen="令和８年10月16日（金）",
        riyu="アンケート調査から見えた潜在のニーズを"
             "見込量につなぐために、"
             "報告書のクロス集計をお願いしたい事項があります。"
             "本調査は無記名であり回答者本人は特定できないため、"
             "報告書で示していただくのはアンケート内のクロスまでとし、"
             "そこから先の個別の確認は当方が行政のデータで行います。"
             "あわせて、前回計画の評価とアンケート結果報告の"
             "概要版（各Ａ４・２ページ）を作成しましたので、"
             "アンケートの概要版に載せた割合を"
             "分析報告書の確定値と照合させてください（Ｓ-36）。"
             "あわせて、令和８年10月14日に報告書用集計表の改訂版を"
             "受領しました。"
             "令和８年10月３日に当方が洗い出した不整合11件のうち、"
             "割合の基数の統一（Ａ-1）、"
             "問27-1と問27-3の基数の統一（Ａ-2）、"
             "表1-1の調査の概要の記入（Ｄ-1）、"
             "障がい児調査の基数の統一（Ｃ-1）が解消されています。"
             "残っているのは基数が１から２の区分に割合を"
             "表示していること（Ａ-3）と、"
             "医療的ケアの３項目が１人の回答であることの"
             "書き方（Ｃ-2）です（Ｓ-37）。"
             "課題の数は改訂版により８件で確定しましたので、"
             "当方の計画素案 第２章７の表２-８に"
             "８区分目を足しました。",
        nos=["Ｓ-26", "Ｓ-36", "Ｓ-37"],
    ),
    dict(
        no=4,
        midashi="計画の体系と成果品の構え、目次の更新",
        kigen="令和８年10月16日（金）",
        riyu="村の方針資料では計画の基本項目が７章、"
             "当方の正本と策定委員会資料は８章で食い違っていました。"
             "令和８年10月14日に折衷の案で決着し、"
             "当方の８章を保ったまま、"
             "第３章の章題を"
             "「第４次北塩原村障がい者計画の推進状況と"
             "本計画で新たに加える施策」へ改め、"
             "同章の末尾に節４と新たに加える施策の一覧（表３-３）を"
             "置きました。"
             "第１章４にも項（４）を本文で置いています。"
             f"図表は図{len(_ZU)}点・表{len(_HYO)}点で番号が確定しており、"
             "章の途中に表を足すと以降の番号がすべてずれるため、"
             "いずれも章の末尾に足す形にしています。"
             "本文の目次は Ｗｏｒｄ の目次フィールドですので、"
             "目次を更新していただければ章題と頁数が入れ替わります"
             "（当方の環境では更新できません）。"
             "あわせて、概要版の体裁についてもご相談させてください。",
        nos=["Ｓ-32", "Ｓ-33", "Ｓ-34"],
    ),
    dict(
        no=5,
        midashi="県へのお取次ぎ",
        kigen="令和８年10月31日（金）",
        riyu="県にご確認いただきたい事項のうち、"
             "計画の記述に直に関わるものです。"
             "村を通じた県への照会（Ｍ-99）とは別に、"
             "研修の実施の予定に関わるものであるため、"
             "お取次ぎいただけますと幸いです。",
        nos=KEN_KEIYU,
    ),
    dict(
        no=6,
        midashi="介護保険事業計画との中長期の推計の揃え方",
        kigen="令和８年10月31日（金）",
        riyu="社会福祉法等の一部を改正する法律"
             "（令和８年法律第51号）により、"
             "介護保険事業（支援）計画には"
             "サービス量の中長期的な推計と"
             "将来に向けたサービス提供体制の確保に関する施策を"
             "必ず記載することとされ、"
             "厚生労働省の概要資料には"
             "「障害福祉分野においても同様の見直しを行う」と"
             "明記されています（11頁。施行日 令和９年４月１日）。"
             "障がい福祉の側でも中長期の推計を載せることとなった場合、"
             "両計画で年次が食い違うと並べて読めなくなります。"
             "貴計画の素案 5-11 が令和22年度（2040年度）を"
             "見据えた見通しを置いておられるため、"
             "年次の揃え方をご相談させてください。",
        nos=["Ｓ-29"],
    ),
    dict(
        no=7,
        midashi="法令の原典の入手のお願い",
        kigen="令和８年10月31日（金）",
        riyu="当方の環境からは、官報及び法令のページに"
             "接続できません（組織のネットワーク方針によります）。"
             f"下記の{_su(len(GENTEN_IRAI))}件について、"
             "原典（官報又は所管省庁の通知・告示）を"
             "お送りいただけますと、照合して本文に反映します。"
             "原典が無い間は、本文に法律番号や月数を書きません。"
             "これまでにお送りいただいた原典により、"
             "高次脳機能障害者支援法（施行通知・施行令）、"
             "社会福祉法等の一部を改正する法律"
             "（令和８年法律第51号。Ｓ-11の一部）、"
             "及び民法等の一部を改正する法律"
             "（令和８年法律第45号。令和８年６月24日公布。Ｓ-13）の"
             "法律番号・公布の日・施行の期日を"
             "本文に書くことができました。"
             "Ｓ-13は解消し、Ｓ-11は原典ではなく"
             "県への照会（上記５）に移しています。"
             "同じ形でお願いできますと幸いです。"
             "このうちＳ-35は、令和８年10月に告示される予定の"
             "基本指針の追加改正です。"
             "当方は現在、障害者部会第157回（令和８年７月10日）及び"
             "障害児支援部会第21回・障害者部会第158回"
             "（令和８年８月19日）の会議資料（新旧対照表）を"
             "原典として用いていますが、"
             "いずれも「改正案」であり告示そのものではありません。"
             "別表第二の六の項の新設（市町村の努力義務事項）、"
             "医療的ケア児等支援地域協議会、"
             "医療的ケア者の定義の二つの類型、"
             "及び重度の障がいのある方の内数の対象サービスは、"
             "いずれもこの改正案を根拠としているため、"
             "告示の現物で最終の照合を行う必要があります。",
        nos=GENTEN_IRAI,
    ),
    dict(
        no=8,
        midashi="県医療計画と県障がい福祉計画の指標の食い違い",
        kigen="令和８年10月31日（金）",
        riyu="第８次福島県医療計画（令和６年３月）を拝見したところ、"
             "精神疾患対策の指標について、第７期福島県障がい福祉計画と"
             "基準年・目標値が違っておりました。"
             "出典はいずれも地域精神保健福祉資源分析データベースです。"
             "正本 第４章２（２）はこれまで県の障がい福祉計画の目標値と"
             "突き合わせておりましたが、県医療計画の目標年（令和11年）が"
             "本計画の終期と同じであるため、どちらに沿うかを"
             "決めておく必要があります。"
             "県にお取次ぎいただく事項を含みます。",
        nos=["Ｓ-38"],
    ),
    dict(
        no=9,
        midashi="アンケート調査結果報告書の直し残しとレビューの反映",
        kigen="令和８年10月31日（金）",
        riyu="レビューの結果を反映いただいた報告書"
             "（令和８年10月22日版）を拝見しました。"
             "前の文を消さずに"
             "新しい文を重ねたとみられる箇所が11か所あり、"
             "日本語が崩れております。"
             "内容の誤りではなく機械的に直せるものですが、"
             "報告書は策定委員会に出す資料ですので、"
             "お戻しいたします（Ｓ-40）。"
             "第６章に加えていただいた方向性のうち、"
             "「第二種運送事業者」「福祉運送事業等の事業者指定」は"
             "道路運送法上の制度の名と合わないように思われますので、"
             "どの制度を指しておられるかをお教えください（Ｓ-41）。"
             "移動の手段の確保は本村の重い課題であり、"
             "担い手を増やす手立てを書くことには賛成です。"
             "あわせて、令和８年10月５日に当方からお出ししたレビューの"
             "反映の状況を1件ずつ確かめました。"
             "すべてを報告書に入れる必要はないと考えており、"
             "報告書に入れていただきたいものと、"
             "当方の計画素案で受けているものに分けております。"
             "二重に書かないよう、どこまで報告書に入れるかを"
             "先に合わせていただけますでしょうか（Ｓ-42）。",
        nos=["Ｓ-40", "Ｓ-41", "Ｓ-42"],
    ),
    dict(
        no=10,
        midashi="第2章7の割合の差し替え（回収40件）",
        kigen="令和８年11月14日（金）",
        riyu="報告書用集計表の令和８年10月21日版（有効回答数40件）に"
             "基づき、正本 第２章７の割合の差し替えをお願いいたします。"
             "回収数が38件から40件になったことによる影響を"
             "洗い出しましたところ、設問ごとに最も多い選択肢は"
             "1件も入れ替わっておらず、課題の８区分・障がい児調査・"
             "障がい別の集計も動いておりません。"
             "結論は変わらず、動いたのは数値だけです。"
             "第２章７の割合は素案の本文にあるものですので、"
             "新しい集計に合わせて差し替えていただけますでしょうか。"
             "当方が足した記述（８件）は当方で直します。"
             "この節を第９節より後に置いているのは、"
             "直し残し11か所のうち４か所は同じ文が正本の第２章７に"
             "あるためです。直る前に差し替えますと、"
             "崩れた文が計画の本文に入ります。"
             "また、村が力を入れるべきこと（問44）は、"
             "新しい集計では相談支援体制37.5％・災害時支援35.0％となり"
             "同率でなくなります。「各34.2％」という書き方をされている"
             "箇所は言い回しの修正が要ります。",
        nos=["Ｓ-39"],
    ),
]


# ===========================================================================
# 3. 自己点検
# ===========================================================================
def verify(zenbu):
    ng = []
    mura = {k: v for k, v in zenbu.items() if v["atesaki"] == "村"}

    # ① 便に入れた番号が実在し、重複しないこと
    ireta = []
    for b in BIN:
        for no in b["nos"]:
            if no not in zenbu:
                ng.append(f"第{b['no']}便に実在しない番号: {no}")
            elif zenbu[no]["atesaki"] != "村":
                ng.append(f"第{b['no']}便に村以外あての番号: {no}")
            ireta.append(no)
    if len(ireta) != len(set(ireta)):
        dup = [x for x in set(ireta) if ireta.count(x) > 1]
        ng.append(f"便をまたいで重複している番号: {sorted(dup)}")

    # ② 村あての照会が、いずれかの便か電話に必ず入っていること
    nokori = set(mura) - set(ireta) - set(DENWA)
    if nokori:
        ng.append(f"どの便にも入っていない照会: {sorted(nokori)}")

    # ③ 村以外あてのものを村の照会文に入れていないこと
    for no in HOKA:
        if no in ireta:
            ng.append(f"村以外あてを村の照会文に入れている: {no}")

    # ④ 番号に抜けがないこと（Ｍ-40からＭ-112まで）
    ban = sorted(int(k.split("-")[1]) for k in mura)
    nuke = [n for n in range(ban[0], ban[-1] + 1) if n not in ban]
    if nuke:
        ng.append(f"村あての番号に抜けがある: {nuke}")

    # ⑤ 期限が便の順に遅くなっていること
    kigen = [b["kigen"] for b in BIN]
    if kigen != sorted(kigen, key=lambda s: (
            int(re.search(r"(\d+)月", s).group(1)),
            int(re.search(r"月(\d+)日", s).group(1)))):
        ng.append(f"便の期限が順に並んでいない: {kigen}")

    # ⑥ 中身が空でないこと
    for no, v in mura.items():
        for key in ("kenmei", "naiyo"):
            if not str(v[key]).strip():
                ng.append(f"{no}の{key}が空")

    # ⑥b 他メンバーへの1通の点検
    ireta_m = []
    for b in MEMBER_SETSU:
        for no in b["nos"]:
            if no in zenbu and zenbu[no]["atesaki"] == "村":
                ng.append(f"他メンバーの文書に村あての番号がある: {no}")
            ireta_m.append(no)
    if len(ireta_m) != len(set(ireta_m)):
        dup = [x for x in set(ireta_m) if ireta_m.count(x) > 1]
        ng.append(f"他メンバーの節をまたいで重複している番号: {sorted(dup)}")
    nokori_m = set(HOKA) - set(ireta_m)
    if nokori_m:
        ng.append(f"他メンバーあてでどの節にも入っていない: "
                  f"{sorted(nokori_m)}")
    # 第2回レビューの確認事項（K_R2）は atsumeru() の対象外
    # （宛先の形が違う）ため、実在の確認はここで併せて引く。
    hoka_moto = {}
    for r in K_R2:
        hoka_moto[_zen(r[0])] = r[2] + r[3]
    for no in HOKA + KEN_KEIYU:
        if no not in zenbu and no not in hoka_moto:
            ng.append(f"他メンバーあてに実在しない番号: {no}")
    # ⑦ 区分（Ａ〜Ｄ）がすべての番号に付いており、
    #    文書に載るのがＣ・Ｄだけであること。
    #    Ａ・Ｂを相手に尋ねたままにしない、Ｃ・Ｄを落とさない。
    waritsuke = {no for b in BIN for no in b["nos"]}
    waritsuke |= {no for b in MEMBER_SETSU for no in b["nos"]}
    waritsuke |= set(DENWA)
    nai = sorted(waritsuke - set(KUBUN_MAP))
    if nai:
        ng.append(f"区分が付いていない番号がある: {nai}")
    for no in sorted(waritsuke & set(KUBUN_MAP)):
        ku = KUBUN_MAP[no][0]
        noru = no in tazuneru(sorted(waritsuke))
        if ku in (KU_A, KU_B) and noru:
            ng.append(f"Ａ・Ｂを相手に尋ねている: {no}（{ku}）")
        if ku not in (KU_A, KU_B) and not noru:
            ng.append(f"Ｃ・Ｄを落としている: {no}（{ku}）")
    #    便・節のいずれにもＣ・Ｄが残っていること（空の文書にしない）
    if not any(tazuneru(b["nos"]) for b in BIN):
        ng.append("村の照会に尋ねるものが1件も残っていない")
    if not any(tazuneru(b["nos"]) for b in MEMBER_SETSU):
        ng.append("他メンバーへの確認に尋ねるものが1件も残っていない")

    # 原典の入手をお願いするものは、第2回レビューの確認事項（K_R2）と
    # 再レビューの照会（S_RV）の2か所に置いてある。
    genten_moto = dict(hoka_moto)
    for recs in (S_RV, S_KH, S_SJ):
        for r in recs:
            genten_moto[_zen(r[0])] = r[2] + r[3]
    for no in GENTEN_IRAI:
        if no not in genten_moto:
            ng.append(f"原典の入手をお願いする番号が実在しない: {no}")
        # 原典の入手の節に、原典の入手でないものを混ぜていないこと。
        # 原典そのものを求めるものか、法律番号・公布日・施行期日のように
        # 原典でしか確かめられない事項を求めるものに限る。
        elif not any(g in genten_moto[no]
                     for g in ("原典", "法律番号", "公布", "施行期日",
                               "告示")):
            ng.append(f"原典の入手でない番号が節5に入っている: {no}")
    #    他メンバーの節の期限が節の順に遅くなっていること
    kigen_m = [b["kigen"] for b in MEMBER_SETSU]
    if kigen_m != sorted(kigen_m, key=lambda s2: (
            int(re.search(r"(\d+)月", s2).group(1)),
            int(re.search(r"月(\d+)日", s2).group(1)))):
        ng.append(f"他メンバーの節の期限が順に並んでいない: {kigen_m}")

    # ⑥c 個々の照会が持っている期限が、入れた便・節の期限より
    #     後になっていないこと。
    #     照会は生成器ごとに起票しており、各記録が自分の期限を持っている。
    #     便・節に入れるときに期限を合わせ忘れると、
    #     文書には便の期限（例 10月16日）が印字されるのに
    #     台帳には別の期限（例 10月21日）が残る。
    def _hi(s3):
        m3 = re.search(r"(\d+)年(\d+)月(\d+)日", str(s3))
        return tuple(int(x) for x in m3.groups()) if m3 else None

    for grp, na in ((BIN, "便"), (MEMBER_SETSU, "節")):
        for b in grp:
            bk = _hi(b["kigen"])
            for no in b["nos"]:
                v = zenbu.get(no)
                if not v:
                    continue
                k = _hi(v.get("kigen", ""))
                if k and bk and k > bk:
                    ng.append(f"{no}の期限（{v['kigen']}）が"
                              f"第{b['no']}{na}の期限（{b['kigen']}）より後")

    # ⑦ 他団体の名前を出していないこと
    s = "\n".join(str(v) for v in zenbu.values()) + "\n".join(
        b["riyu"] for b in BIN)
    for g in ["小野町", "金ケ崎", "金ヶ崎", "阿蘇", "札幌"]:
        if g in s:
            ng.append(f"他団体の名前が入っている: {g}")

    if ng:
        print("自己点検 不合格:")
        for x in ng:
            print("   -", x)
        raise SystemExit(1)
    print(f"  自己点検: 村あて{len(mura)}件が{len(BIN)}通{len(ireta)}件と"
          f"電話{len(DENWA)}件に過不足なく収まる／"
          f"番号Ｍ-{ban[0]}〜Ｍ-{ban[-1]}に抜けなし／"
          f"便をまたぐ重複なし／村以外あて{len(HOKA)}件を含めていない／"
          f"期限が便の順／"
          f"他メンバーあて{len(HOKA)}件が"
          f"{len(MEMBER_SETSU)}節{len(ireta_m)}件に収まる／"
          f"区分が{len(waritsuke)}件すべてに付き、"
          f"文書に載るのはＣ・Ｄの"
          f"{len(tazuneru(sorted(waritsuke)))}件"
          f"（Ａ・Ｂの{len(jibun(sorted(waritsuke)))}件は"
          f"当方で確認する分として掲げる）")
    return mura


# ===========================================================================
# 4. 文書
# ===========================================================================
def midashi_bin(doc, b, n):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(f"第{b['no']}便　{b['midashi']}（{n}件）")
    r.font.size = Pt(13)
    r.font.bold = True
    return p


def build(zenbu):
    doc = new_doc("北塩原村　第８期障がい福祉計画・第４期障がい児福祉計画",
                  f"村へのご照会（{UCHIAWASE}　ビズアップ公共コンサルティング"
                  "株式会社）")

    h2(doc, "このご照会について")
    para(doc,
         "計画の素案を作る中で、村にご確認いただきたい事項が"
         f"{len([v for v in zenbu.values() if v['atesaki'] == '村'])}件に"
         "なりました。"
         "一度にお送りするとご負担が大きく、"
         "お急ぎでないものに引きずられて急ぐものの回答が遅れます。"
         f"このため、回答がいつ必要かによって{_su(len(BIN))}通に"
         "分けました。"
         "第１便から順にお願いできますと幸いです。")
    para(doc,
         f"{IINKAI}の委員会に前回計画の評価と計画の素案をお出しする"
         "予定とうかがっています。"
         "第１便は、その資料に直に載る数値です。"
         "第２便は、サービスの見込量をお一人ずつ積み上げて"
         "確定するために要るものです。"
         "第３便以降は、計画の確定には間に合えば足りるものです。")
    para(doc,
         "人数が少なく個人が特定されるおそれがある事項については、"
         "実数ではなく「該当あり」「該当なし」のみでも差し支えありません。"
         "その場合は計画にも実数を載せず、"
         "確かめた結果があることだけを書きます。")

    note(doc,
         "【お電話でお願いしたいこと】"
         "計画策定方針（令和８年５月25日）の工程について、"
         "第２回策定委員会の開催日、パブリックコメントの時期、"
         "４町村広域の協議会の開催日、庁議と議会への説明の時期が"
         "現在どうなっているかをお聞かせください。"
         "工程が決まらないと、このご照会の期限そのものを"
         "決められないためです。"
         "書面では第１便のＭ-125にまとめております"
         "（従前はＭ-40・Ｍ-123として別に掲げておりましたが、"
         "同じ事項であるため１件にまとめました）。")

    h2(doc, f"{_su(len(BIN))}通の構成")
    rows = [["便", "内容", "ご回答をお願いする件数",
             "回答の希望期限", "なぜその期限か"]]
    for b in BIN:
        t = tazuneru(b["nos"])
        j = jibun(b["nos"])
        ken = f"{len(t)}件"
        if j:
            ken += f"（ほか{len(j)}件は当方で確認）"
        rows.append([f"第{b['no']}便", b["midashi"], ken,
                     b["kigen"], b["riyu"][:60] + "…"])
    table(doc, rows, widths=[1100, 3800, 1900, 2200, 6400])

    for b in BIN:
        t = tazuneru(b["nos"])
        if not t:
            continue
        midashi_bin(doc, b, len(t))
        para(doc, b["riyu"])
        rows = [["番号", "区分", "ご確認いただきたいこと", "内容",
                 "計画のどこに使うか"]]
        for no in t:
            v = zenbu[no]
            rows.append([v["no"], KUBUN_MAP[no][0].split()[0],
                         v["kenmei"], v["naiyo"],
                         v["tsukaimichi"] or "―"])
        table(doc, rows, widths=[900, 700, 3000, 7500, 3100])

    h2(doc, "回答をいただいた後に当方が行うこと")
    table(doc, [
        ["便", "回答後に進めること"],
        ["第１便",
         "前回計画の評価の表を埋め、委員会資料として確定します。"
         "計画の名称の食い違いは、第１章２の関連計画の記載と"
         "図（計画の位置づけ・他の計画との関係）の両方を直します。"],
        ["第２便",
         "サービスの見込量を第２次算定として組み直します。"
         "０人としているサービスについては、"
         "確かめた結果として０人とするか、見込みを立てるかを決めます。"],
        ["第３便",
         "高齢者福祉の事業のうち障がいのある方が対象となるものを"
         "計画に書き、対象とならないものは本計画で埋めるべき"
         "すき間として整理します。"],
        ["第４便",
         "一般の子育て施策で受けられる範囲と、"
         "障がい福祉で受ける範囲の切り分けを計画に書きます。"],
        ["第５便",
         "既にある施策の実績を計画に書き、"
         "次の３年間の目標の基準値とします。"
         "村内の事業所の指定の形が分かれば、"
         "村内の介護保険の事業所の力を障がい福祉の供給に"
         "生かせるかどうかを判断します。"],
        ["第６便",
         "会津北部４町村及び県との広域整合を確かめます。"
         "地域生活支援拠点、基幹相談支援センター、"
         "医療的ケア児等支援、インクルージョン推進及び"
         "強度行動障害支援について、"
         "相手方の計画と同じ連携単位・同じ会議体・同じ数え方に"
         "なっているかを突き合わせ、"
         "食い違う場合は本計画の記述を直します。"
         "高次脳機能障害者支援法の施行に伴う県の体制も"
         "ここで確かめます。"],
    ], widths=[1200, 14200])

    note(doc,
         "【ご回答の様式について】"
         "様式は定めておりません。"
         "番号を添えてメール本文にご記入いただく形で差し支えありません。"
         "既存の資料で足りるものは、その資料をお送りいただくだけで結構です。"
         "分からない事項は「分からない」「記録していない」と"
         "お答えいただければ、それ自体を計画に書きます"
         "（記録がないことも、次の３年間で記録を始める根拠になります）。")

    _jibun_setsu(doc, zenbu, {},
                 [no for b in BIN for no in b["nos"]] + list(DENWA), "村")

    doc.save(OUT_FILE)
    return doc


def _jibun_setsu(doc, zenbu, r2, nos_all, aite):
    """当方で確認する分（Ａ・Ｂ）を末尾に掲げる。ご回答は不要である。

    従前はこれらも一緒にお尋ねしていた。当方で解消すべき疑問を
    相手の確認責任に転嫁しないため、当方の側に引き取ったものである。
    何を引き取ったかが分からないと、相手は従前の照会との差が
    追えないため、落としたのではなく当方で確認することを明記する。
    """
    #   番号の順に並べる。便・節の順のままでは引きにくい。
    j = sorted((no for no in nos_all
                if KUBUN_MAP.get(no, ("",))[0] in (KU_A, KU_B)),
               key=lambda x: (0 if x.startswith("Ｍ") else 1,
                              int(x.split("-")[1])))
    if not j:
        return
    h2(doc, f"当方で確認する分（{len(j)}件・ご回答は不要です）")
    para(doc,
         "従前のご照会に掲げておりました事項のうち、次の"
         f"{len(j)}件は当方で確認いたします。ご回答は不要です。"
         "内訳は、ほかのご照会に１件としてまとめたもの、"
         "既に決着したもの、及び受領済みの資料で当方が確かめられる"
         "ものです。"
         "当方で解消すべき計算上・資料解釈上の疑問を"
         f"{aite}の確認責任に移さないという考え方によるものです。")
    rows = [["番号", "区分", "ご確認いただきたいとしていたこと",
             "当方での取扱い"]]
    for no in j:
        v = zenbu.get(no) or r2.get(no) or {}
        ku, riyu, togo = KUBUN_MAP[no]
        rows.append([no, ku.split()[0], v.get("kenmei", "―"),
                     (f"{togo}にまとめました。" if togo else "") + riyu])
    table(doc, rows, widths=[900, 700, 3600, 10000])


def build_member(zenbu):
    """他メンバーへの確認を1通に組む。村の照会とは別の文書にする。"""
    r2 = {_zen(r[0]): dict(no=_zen(r[0]), atesaki=r[1], kenmei=r[2],
                           naiyo=r[2], tsukaimichi=r[3])
          for r in K_R2}
    doc = new_doc("北塩原村　第８期障がい福祉計画・第４期障がい児福祉計画",
                  f"素案作成のご担当者様へのご確認（{UCHIAWASE}　"
                  "ビズアップ公共コンサルティング株式会社）")

    h2(doc, "このご確認について")
    para(doc,
         "素案の作成と当方の追記を一つの計画書にまとめるにあたり、"
         f"ご担当者様にご確認・ご相談したい事項が{_su(len(HOKA))}件、"
         "県へのお取次ぎをお願いしたい事項が"
         f"{_su(len(KEN_KEIYU))}件、"
         f"法令の原典の入手をお願いしたい事項が"
         f"{_su(len(GENTEN_IRAI))}件"
         "ございます。"
         "村へのご照会とは宛先が違うため、別の文書にしております。")
    para(doc,
         "いずれも村からの回答を待たずに決められるものです。"
         "最終版の体裁と用語の統一に関わるため、"
         "先に決めておきますと後戻りが生じません。")
    para(doc,
         "なお、当方が作成した追記は、"
         "ご担当者様の素案の記述を書き換えたものではありません。"
         "事実の誤りを直した箇所は３件あり、"
         "いずれも修正履歴の文書"
         "（北塩原村_計画素案_修正履歴.docx）に"
         "どこをどう直したかを残しております。")

    ikiru = [b for b in MEMBER_SETSU if tazuneru(b["nos"])]
    h2(doc, f"{_su(len(ikiru))}節の構成")
    rows = [["節", "内容", "ご確認をお願いする件数",
             "ご回答の希望期限"]]
    for b in ikiru:
        t = tazuneru(b["nos"])
        j = jibun(b["nos"])
        ken = f"{_su(len(t))}件"
        if j:
            ken += f"（ほか{_su(len(j))}件は当方で確認）"
        rows.append([_su(b["no"]), b["midashi"], ken, b["kigen"]])
    table(doc, rows, widths=[800, 4400, 1800, 3000])

    for b in ikiru:
        pp = doc.add_paragraph()
        pp.paragraph_format.space_before = Pt(18)
        pp.paragraph_format.space_after = Pt(6)
        rr = pp.add_run(f"{_su(b['no'])}　{b['midashi']}"
                        f"（{_su(len(tazuneru(b['nos'])))}件）")
        rr.font.size = Pt(13)
        rr.font.bold = True
        para(doc, b["riyu"])
        rows = [["番号", "区分", "ご確認いただきたいこと", "内容"]]
        for no in tazuneru(b["nos"]):
            v = zenbu.get(no) or r2[no]
            rows.append([v["no"], KUBUN_MAP[no][0].split()[0],
                         v["kenmei"], v["naiyo"]])
        table(doc, rows, widths=[900, 700, 3400, 10200])

    h2(doc, "ご回答をいただいた後に当方が行うこと")
    table(doc, [
        ["節", "ご回答後に進めること"],
        ["１",
         "元のデータをいただければ、村の実績が届いた段階で"
         "図を作り直します。"
         "白黒で見分けられない配色は、青を橙に替えるか"
         "網かけを足す形で直します。"
         "画像の中の表題は、図の下の番号付きの表題と"
         "二重になるため外します。"],
        ["２",
         "決まった方に合わせて、制度名称の置換の一覧を作り、"
         "正本の全文に当てます。"
         "当方の追記も同じ表記に揃えます。"],
        ["３",
         "クロス集計の結果を見込量の第２次算定の材料にします。"
         "個別の確認は当方が行政のデータで行いますので、"
         "報告書ではアンケート内のクロスまでで結構です。"],
        ["４",
         "県研修の予定が分かれば、"
         "第４章２（７）の意思決定支援に関する研修への参加を"
         "第２版を用いた研修の参加として位置づけます。"],
        ["５",
         "原典をいただければ照合し、法律番号と施行期日を"
         "本文に書きます。"
         "いただけない間は、公布の月までの書き方にとどめます。"],
    ], widths=[800, 14600])

    note(doc,
         "【ご回答の様式について】"
         "様式は定めておりません。"
         "番号を添えてメール本文にご記入いただく形で差し支えありません。"
         "図の元データは、作図に使った道具が分かるだけでも助かります。"
         "表記の統一については、どちらに揃えるかのご判断だけで結構です。"
         "置換の一覧は当方で作ります。")

    _jibun_setsu(doc, zenbu, r2,
                 [no for b in MEMBER_SETSU for no in b["nos"]],
                 "ご担当者様")

    doc.save(OUT_MEMBER)
    return doc


def main():
    zenbu = atsumeru()
    mura = verify(zenbu)
    doc = build(zenbu)
    doc_m = build_member(zenbu)
    print(f"作成: {OUT_FILE}")
    print(f"  段落{len(doc.paragraphs)}・表{len(doc.tables)}")
    bin_t = sum(len(tazuneru(b["nos"])) for b in BIN)
    bin_j = sum(len(jibun(b["nos"])) for b in BIN) + len(jibun(DENWA))
    print(f"  村あて{len(mura)}件のうち、ご回答をお願いするのは{bin_t}件"
          f"（当方で確認{bin_j}件）")
    for b in BIN:
        t, j = tazuneru(b["nos"]), jibun(b["nos"])
        print(f"  第{b['no']}便 {b['midashi']}　"
              f"尋ねる{len(t)}件・当方{len(j)}件　{b['kigen']}")
    print(f"作成: {OUT_MEMBER}")
    print(f"  段落{len(doc_m.paragraphs)}・表{len(doc_m.tables)}")
    m_t = sum(len(tazuneru(b["nos"])) for b in MEMBER_SETSU)
    m_j = sum(len(jibun(b["nos"])) for b in MEMBER_SETSU)
    print(f"  他メンバー・県あて{len(HOKA) + len(KEN_KEIYU)}件のうち、"
          f"ご確認をお願いするのは{m_t}件（当方で確認{m_j}件）")
    for b in MEMBER_SETSU:
        t, j = tazuneru(b["nos"]), jibun(b["nos"])
        nokeru = "（節を落とす）" if not t else ""
        print(f"  {b['no']} {b['midashi']}　"
              f"尋ねる{len(t)}件・当方{len(j)}件{nokeru}　{b['kigen']}")


if __name__ == "__main__":
    main()
