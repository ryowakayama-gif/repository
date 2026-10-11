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
                 S_YS, S_KY, S_KN, S_CS, K_IK,
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
             "確かめておく必要があります。",
        nos=["Ｍ-41", "Ｍ-42", "Ｍ-43", "Ｍ-44", "Ｍ-69", "Ｍ-70",
             "Ｍ-95", "Ｍ-96", "Ｍ-104"],
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
             "Ｍ-92", "Ｍ-93", "Ｍ-94", "Ｍ-97", "Ｍ-109"],
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
             "Ｍ-73", "Ｍ-81"],
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
             "Ｍ-103", "Ｍ-105", "Ｍ-106", "Ｍ-108"],
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
             "Ｍ-107"],
    ),
]

# 文書にせず電話で確かめるもの
DENWA = ["Ｍ-40"]

# 村以外に出すもの（村の照会文には入れない。他メンバーへの確認）
HOKA = ["Ｓ-19", "Ｓ-20", "Ｓ-21", "Ｓ-22", "Ｓ-23", "Ｓ-24",
        "Ｓ-25", "Ｓ-26"]

# 県あてだが他メンバーを通じてお願いするもの
KEN_KEIYU = ["Ｓ-12"]

# 法令の原典の入手をお願いするもの（当方は取得できない）
GENTEN_IRAI = ["Ｓ-11", "Ｓ-13"]


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
             "令和８年10月10日に当方で図26点・表92点に番号と表題を付け、"
             "目次の後に図表目次を置きました。"
             "これに伴って、既存の図についてお願いしたいことが"
             "６件あります。"
             "いずれも最終版の体裁に関わるもので、"
             "村の回答を待たずに決められるものです。",
        nos=["Ｓ-19", "Ｓ-20", "Ｓ-21", "Ｓ-22", "Ｓ-23", "Ｓ-24"],
    ),
    dict(
        no=2,
        midashi="制度名称の表記の統一",
        kigen="令和８年10月16日（金）",
        riyu="計画の冒頭の注記は、法律用語や施設名等の固有名称を除き"
             "「害」の字をひらがなで表記するとしています。"
             "ところが制度の名称について、"
             "法令の表記と村の表記が混在している箇所があります。"
             "同じ並びの中で食い違っている箇所もあるため、"
             "どちらに揃えるかを決めていただきたく存じます。"
             "当方の追記は決まった方に合わせます。"
             "決まり次第、置換の一覧を作ってお送りします。",
        nos=["Ｓ-25"],
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
             "そこから先の個別の確認は当方が行政のデータで行います。",
        nos=["Ｓ-26"],
    ),
    dict(
        no=4,
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
        no=5,
        midashi="法令の原典の入手のお願い",
        kigen="令和８年10月31日（金）",
        riyu="当方の環境からは、官報及び法令のページに"
             "接続できません（組織のネットワーク方針によります）。"
             "下記の２件について、法律番号・公布日・施行期日の"
             "原典（官報又は所管省庁の通知）をお送りいただけますと、"
             "照合して本文に反映します。"
             "原典が無い間は、本文に法律番号を書かず"
             "「令和８年６月に公布された」までの書き方に"
             "とどめています。"
             "さきに高次脳機能障害者支援法については"
             "施行通知と施行令をお送りいただき、"
             "公布と施行の日を本文に書くことができました。"
             "同じ形でお願いできますと幸いです。",
        nos=GENTEN_IRAI,
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

    # ④ 番号に抜けがないこと（Ｍ-40からＭ-109まで）
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
    for no in HOKA + KEN_KEIYU:
        if no not in zenbu:
            ng.append(f"他メンバーあてに実在しない番号: {no}")
    for no in GENTEN_IRAI:
        if no not in {_zen(r[0]) for r in K_R2}:
            ng.append(f"原典の入手をお願いする番号が実在しない: {no}")
    #    他メンバーの節の期限が節の順に遅くなっていること
    kigen_m = [b["kigen"] for b in MEMBER_SETSU]
    if kigen_m != sorted(kigen_m, key=lambda s2: (
            int(re.search(r"(\d+)月", s2).group(1)),
            int(re.search(r"月(\d+)日", s2).group(1)))):
        ng.append(f"他メンバーの節の期限が順に並んでいない: {kigen_m}")

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
          f"{len(MEMBER_SETSU)}節{len(ireta_m)}件に収まる")
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
         "第２回協議会の開催日、パブリックコメントの時期、"
         "第３回協議会（４町村広域）の開催日、庁議と議会への説明の時期が"
         "現在どうなっているかをお聞かせください。"
         "工程が決まらないと、このご照会の期限そのものを"
         "決められないためです（Ｍ-40）。")

    h2(doc, f"{_su(len(BIN))}通の構成")
    rows = [["便", "内容", "件数", "回答の希望期限", "なぜその期限か"]]
    for b in BIN:
        rows.append([f"第{b['no']}便", b["midashi"], f"{len(b['nos'])}件",
                     b["kigen"], b["riyu"][:60] + "…"])
    table(doc, rows, widths=[1100, 4200, 900, 2400, 6800])

    for b in BIN:
        midashi_bin(doc, b, len(b["nos"]))
        para(doc, b["riyu"])
        rows = [["番号", "ご確認いただきたいこと", "内容",
                 "計画のどこに使うか"]]
        for no in b["nos"]:
            v = zenbu[no]
            rows.append([v["no"], v["kenmei"], v["naiyo"],
                         v["tsukaimichi"] or "―"])
        table(doc, rows, widths=[900, 3200, 8200, 3100])

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

    doc.save(OUT_FILE)
    return doc


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

    h2(doc, f"{_su(len(MEMBER_SETSU))}節の構成")
    rows = [["節", "内容", "件数", "ご回答の希望期限"]]
    for b in MEMBER_SETSU:
        rows.append([_su(b["no"]), b["midashi"],
                     f"{_su(len(b['nos']))}件", b["kigen"]])
    table(doc, rows, widths=[800, 5200, 1000, 3000])

    for b in MEMBER_SETSU:
        pp = doc.add_paragraph()
        pp.paragraph_format.space_before = Pt(18)
        pp.paragraph_format.space_after = Pt(6)
        rr = pp.add_run(f"{_su(b['no'])}　{b['midashi']}"
                        f"（{_su(len(b['nos']))}件）")
        rr.font.size = Pt(13)
        rr.font.bold = True
        para(doc, b["riyu"])
        rows = [["番号", "ご確認いただきたいこと", "内容"]]
        for no in b["nos"]:
            v = zenbu.get(no) or r2[no]
            rows.append([v["no"], v["kenmei"], v["naiyo"]])
        table(doc, rows, widths=[900, 3600, 10900])

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

    doc.save(OUT_MEMBER)
    return doc


def main():
    zenbu = atsumeru()
    mura = verify(zenbu)
    doc = build(zenbu)
    doc_m = build_member(zenbu)
    print(f"作成: {OUT_FILE}")
    print(f"  段落{len(doc.paragraphs)}・表{len(doc.tables)}")
    print(f"  村あて{len(mura)}件＝{len(BIN)}通{sum(len(b['nos']) for b in BIN)}件"
          f"＋電話{len(DENWA)}件")
    for b in BIN:
        print(f"  第{b['no']}便 {b['midashi']}　{len(b['nos'])}件　"
              f"{b['kigen']}")
    print(f"作成: {OUT_MEMBER}")
    print(f"  段落{len(doc_m.paragraphs)}・表{len(doc_m.tables)}")
    print(f"  他メンバーあて{len(HOKA)}件＋県への取次ぎ"
          f"{len(KEN_KEIYU)}件＋原典の入手{len(GENTEN_IRAI)}件")
    for b in MEMBER_SETSU:
        print(f"  {b['no']} {b['midashi']}　{len(b['nos'])}件　"
              f"{b['kigen']}")


if __name__ == "__main__":
    main()
