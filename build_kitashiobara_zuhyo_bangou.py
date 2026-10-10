# -*- coding: utf-8 -*-
"""
北塩原村　図表番号・表題・図表目次の付与

出力: output/北塩原村_図表番号一覧.xlsx
      （番号と表題そのものは build_kitashiobara_honpon.py が正本に入れる）

何をするものか
  図表一覧（build_kitashiobara_zuhyo.py）の06シートで定めた規約に従い、
  正本の図24点・表84点に番号と表題を付け、目次の後に図表目次を置く。

規約（06シートによる）
  番号の形       図は「図１-１」、表は「表１-１」。章ごとの通し番号
  番号を振る対象 データを示す図と表のすべて。注記ボックス（1×1）は除く
  置き場所       表題は表の上、図の表題は図の下
  出所           図表の下に「資料：○○（基準日）」

本書で足した規約の補足（1件）
  出所を付けるのは、村の実績・アンケート・外部資料から数値を引いた
  図表に限る。手順・区分・確認事項を整理しただけの表
  （「確かめること／確かめ方」等）は本計画の記述そのものであり、
  出所を書くとかえって出どころが二重になる。

他メンバー版の図について分かったこと
  16点のうち14点は、表題が画像の中に焼き込まれている。
  本書はその表題を読み取って、番号付きの表題を図の下に置く。
  このため最終版では表題が画像内と図の下で二重になる。
  画像を作り直すときに画像内の表題を外すことを提案する（Ｓ-21）。

  表題が画像内の語のままでは何の図か分からないものが2点ある。
    ＜障がい種別の内訳＞   → 身体障害者手帳所持者の障がい種別の内訳
    ＜障がい程度の内訳＞   → 精神障害者保健福祉手帳所持者の障がい程度の内訳
  図表目次に並べたときに読めるよう、章節が分かる表題に改めている。

  図２-８（現在利用しているサービスと今後３年間で利用したいサービス）は、
  画像内の表題が右端で切れており、括弧の中が読み取れない。
  原データでの確認が要る（Ｓ-22）。
"""

import os
import re
import sys

import docx
from openpyxl import Workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from kitashiobara_common import (  # noqa: E402
    add_sheet, ensure_out_dir, style_header_row, write_row,
)

REPO_ROOT = os.path.dirname(os.path.abspath(__file__))
OUT_FILE = f"{REPO_ROOT}/output/北塩原村_図表番号一覧.xlsx"
HONPON = f"{REPO_ROOT}/output/北塩原村_計画素案_正本_移植後.docx"

HI = "FFF2CC"      # 他メンバー版の図（画像内に表題がある）
MI = "FFC7CE"      # 確認が要る
OK = "E2EFDA"      # 当方が作った図

A_MURA = "資料：北塩原村"
A_ANKE = "資料：障がい福祉に関するアンケート調査（令和８年）"
A_KYUFU = "資料：村提供の障がいサービス給付実績"


# ===========================================================================
# 1. 図（本文に現れる順。章ごとに1から振る）
#    dare … 他メンバー版か当方か
#    midashi_zo … 画像の中に焼き込まれている表題（無いものは None）
# ===========================================================================
ZU = [
    dict(sho=1, midashi="計画の位置づけ", dare="他メンバー", midashi_zo=None,
         moto="", basho="第1章2 計画の位置づけ"),
    dict(sho=1, midashi="障がい福祉計画におけるＰＤＣＡサイクル",
         dare="他メンバー", midashi_zo=None, moto="",
         basho="第1章6 計画の推進体制"),
    dict(sho=2, midashi="年齢３区分別人口の推移", dare="他メンバー",
         midashi_zo="年齢3区分別人口の推移",
         moto="資料：住民基本台帳（各年４月１日現在）",
         basho="第2章1（1）人口の推移"),
    dict(sho=2, midashi="障がい者手帳所持者数の推移", dare="他メンバー",
         midashi_zo="障がい者手帳所持者数の推移",
         moto=A_MURA + "（各年４月１日現在）",
         basho="第2章2 障がい者手帳の所持者数"),
    dict(sho=2, midashi="身体障がい者の推移", dare="他メンバー",
         midashi_zo="＜身体障がい者の推移＞",
         moto=A_MURA + "（各年４月１日現在）",
         basho="第2章3 身体障がい者の現状"),
    dict(sho=2, midashi="身体障害者手帳所持者の障がい種別の内訳",
         dare="他メンバー", midashi_zo="＜障がい種別の内訳＞",
         moto=A_MURA + "（令和８年４月１日現在）",
         basho="第2章3 身体障がい者の現状"),
    dict(sho=2, midashi="知的障がい者の推移", dare="他メンバー",
         midashi_zo="＜知的障がい者の推移＞",
         moto=A_MURA + "（各年４月１日現在）",
         basho="第2章4 知的障がい者の現状"),
    dict(sho=2, midashi="精神障がい者の推移", dare="他メンバー",
         midashi_zo="＜精神障がい者の推移＞",
         moto=A_MURA + "（各年４月１日現在）",
         basho="第2章5 精神障がい者の現状"),
    dict(sho=2, midashi="精神障害者保健福祉手帳所持者の障がい程度の内訳",
         dare="他メンバー", midashi_zo="＜障がい程度の内訳＞",
         moto=A_MURA + "（令和８年４月１日現在）",
         basho="第2章5 精神障がい者の現状"),
    dict(sho=2, midashi="外出するときに困ること", dare="他メンバー",
         midashi_zo="外出するときに困ること", moto=A_ANKE,
         basho="第2章7（5）調査結果の概要"),
    dict(sho=2,
         midashi="現在利用しているサービスと今後３年間で利用したいサービス",
         dare="他メンバー",
         midashi_zo="現在利用しているサービスと今後3年間で利用したいサービス",
         moto=A_ANKE, basho="第2章7（5）調査結果の概要"),
    dict(sho=2, midashi="災害時に不安なこと", dare="他メンバー",
         midashi_zo="災害時に不安なこと", moto=A_ANKE,
         basho="第2章7（5）調査結果の概要"),
    dict(sho=2, midashi="今後、村が特に力を入れるべきこと（３つまで）",
         dare="他メンバー", midashi_zo="今後、村が特に力を入れるべきこと（3つまで）",
         moto=A_ANKE, basho="第2章7（5）調査結果の概要"),
    dict(sho=2, midashi="主な項目の障がい別の状況（障がい者調査）",
         dare="他メンバー", midashi_zo="主な項目の障がい別の状況（障がい者調査）",
         moto=A_ANKE, basho="第2章7（5）調査結果の概要"),
    dict(sho=3, midashi="計画の体系", dare="当方", midashi_zo=None,
         moto="資料：第４次北塩原村障がい者計画及び本計画により作成",
         basho="第3章3 基本施策の目標値と本計画期間の主な取組み"),
    dict(sho=4, midashi="地域生活支援拠点等の５つの機能", dare="当方",
         midashi_zo=None,
         moto="資料：国の基本指針 第二 五及び北塩原村の資料により作成",
         basho="第4章2（5）地域生活支援の充実"),
    dict(sho=5, midashi="北塩原村の障がい福祉に関する圏域（概念図）",
         dare="他メンバー",
         midashi_zo="北塩原村の障がい福祉に関する圏域（概念図）", moto="",
         basho="第5章2（4）広域連携によるサービス提供体制の確保"),
    dict(sho=5, midashi="緊急時に働く仕組み（４つの場面）", dare="当方",
         midashi_zo=None, moto="資料：北塩原村（第５章２（11）により作成）",
         basho="第5章2（11）災害時にサービスを止めないための備え"),
    dict(sho=5, midashi="ライフコースと制度の移行", dare="当方",
         midashi_zo=None,
         moto="資料：北塩原村（第５章５及び第７章により作成）",
         basho="第5章5 年齢到達に伴うサービスの移行"),
    dict(sho=5, midashi="65歳到達時の判定の流れ", dare="当方",
         midashi_zo=None, moto="資料：障害者総合支援法第７条及び本計画により作成",
         basho="第5章5（2）65歳到達時のサービスの区分"),
    dict(sho=5, midashi="関係機関との連携", dare="当方", midashi_zo=None,
         moto="資料：北塩原村（第５章６により作成）",
         basho="第5章6 関係機関との連携"),
    dict(sho=5, midashi="障がい福祉サービス等に係る給付費の推移",
         dare="当方", midashi_zo=None, moto=A_KYUFU,
         basho="第5章7（1）給付費の推移"),
    dict(sho=5, midashi="サービス別の給付費（令和７年度）", dare="当方",
         midashi_zo=None, moto=A_KYUFU,
         basho="第5章7（3）サービス別の構成"),
    dict(sho=6, midashi="高齢化率の推移", dare="当方", midashi_zo=None,
         moto=A_MURA + "（各年４月１日現在）", basho="第6章1（1）高齢化率"),
    dict(sho=7, midashi="他の計画との関係", dare="他メンバー",
         midashi_zo="図表　他の計画との関係", moto="",
         basho="第7章 他計画との連携（章の冒頭）"),
]


# ===========================================================================
# 2. 表（本文に現れる順。章ごとに1から振る）
#    atama … 表頭の先頭2セル。正本と突き合わせるための目印
# ===========================================================================
def _h(*cells):
    return list(cells)


HYO = [
    # --- 第1章 -----------------------------------------------------------
    dict(sho=1, midashi="本計画を構成する計画と根拠法令",
         atama=_h("計画名", "項目"), moto=""),
    dict(sho=1, midashi="本計画と計画期間が重なる村の計画",
         atama=_h("計画", "計画期間"), moto=""),
    dict(sho=1, midashi="福島県の計画との整合",
         atama=_h("県の計画", "計画期間"), moto=""),
    dict(sho=1, midashi="本計画でいう圏域",
         atama=_h("本計画での呼び方", "範囲"), moto=""),
    dict(sho=1, midashi="県障がい者計画の重点施策との対応",
         atama=_h("本計画の基本施策", "対応する県の重点施策"), moto=""),
    dict(sho=1, midashi="県の成果目標と本計画の目標設定",
         atama=_h("成果目標", "県の第７期目標値"), moto=""),
    dict(sho=1, midashi="計画の期間",
         atama=_h("年度", "R6\n(2024)"), moto=""),
    dict(sho=1, midashi="第８期計画における成果目標の構成の見直し",
         atama=_h("第８期計画の成果目標", "主な見直し内容"), moto=""),
    dict(sho=1, midashi="年次の点検評価の項目",
         atama=_h("項目", "書くこと"), moto=""),
    dict(sho=1, midashi="既にある資源を評価する順序",
         atama=_h("順", "確かめること"), moto=""),
    # --- 第2章 -----------------------------------------------------------
    dict(sho=2, midashi="本計画期間の人口の見通し",
         atama=_h("区分", "令和８年度\n（現状）"), moto=A_MURA),
    # この2表は「（単位：人、令和８年４月１日現在）」が表に添えられている。
    # 出所に基準日を重ねて書かない。
    dict(sho=2, midashi="特別支援学校在籍者数",
         atama=_h("学校名", "小学部"), moto=A_MURA),
    dict(sho=2, midashi="特別支援学級在籍者数",
         atama=_h("学校名", "在籍者数"), moto=A_MURA),
    dict(sho=2, midashi="アンケート調査の対象者及び実施方法",
         atama=_h("項目", "障がい者調査"), moto=""),
    dict(sho=2, midashi="アンケート調査の回収状況",
         atama=_h("調査区分", "配布数"), moto=A_ANKE),
    dict(sho=2, midashi="アンケート調査結果の概要",
         atama=_h("分野", "主な調査結果"), moto=A_ANKE),
    dict(sho=2, midashi="主な項目の障がい別の状況",
         atama=_h("項目", "全体\n（n=38）"), moto=A_ANKE),
    dict(sho=2, midashi="調査結果からみた課題と本計画での対応",
         atama=_h("主な課題", "主な調査結果"), moto=""),
    # --- 第3章 -----------------------------------------------------------
    dict(sho=3, midashi="第４次北塩原村障がい者計画の基本施策ごとの目標値",
         atama=_h("基本施策", "指標項目"),
         moto="資料：第４次北塩原村障がい者計画及び本計画により作成"),
    dict(sho=3, midashi="実績がない場合の評価の六つの区分",
         atama=_h("区分", "意味"), moto=""),
    # --- 第4章 -----------------------------------------------------------
    dict(sho=4, midashi="第７期計画の成果目標の達成状況",
         atama=_h("成果目標", "項目"), moto=A_MURA),
    dict(sho=4, midashi="施設入所者の地域生活への移行の目標",
         atama=_h("項目", "数値"), moto=""),
    dict(sho=4, midashi="地域移行の可能性を確かめる事項",
         atama=_h("確認する事項", "確認の方法"), moto=""),
    dict(sho=4,
         midashi="精神障害にも対応した地域包括ケアシステムの構築の目標",
         atama=_h("項目", "数値"), moto=""),
    dict(sho=4, midashi="精神障がい者のサービス利用の見込み（活動指標）",
         atama=_h("サービス種別", "単位"), moto=""),
    dict(sho=4, midashi="精神病床からの退院に関する県の算定との関係",
         atama=_h("項目", "値"), moto=""),
    dict(sho=4, midashi="福祉施設から一般就労への移行等の目標",
         atama=_h("項目", "数値"), moto=""),
    dict(sho=4, midashi="就労移行支援等から一般就労への移行者数の目標",
         atama=_h("事業種別", "令和６年度実績"), moto=""),
    dict(sho=4, midashi="障がい児支援の提供体制の整備等の目標",
         atama=_h("項目", "現状\n（令和８年度見込み）"), moto=""),
    dict(sho=4, midashi="早期発見から支援につながる経路の記録",
         atama=_h("記録する事項", "ねらい"), moto=""),
    dict(sho=4, midashi="地域生活支援の充実の目標",
         atama=_h("項目", "数値"), moto=""),
    dict(sho=4, midashi="地域生活支援拠点等の５つの機能と記録する実績",
         atama=_h("機能", "本村での確保の方法"), moto=""),
    dict(sho=4, midashi="相談支援体制の充実・強化等の目標",
         atama=_h("項目", "数値"), moto=""),
    dict(sho=4, midashi="障害福祉人材の確保・定着等の目標",
         atama=_h("項目", "数値"), moto=""),
    dict(sho=4,
         midashi="障害福祉サービス等の質を向上させる体制の構築の目標",
         atama=_h("項目", "数値"), moto=""),
    dict(sho=4, midashi="発達障がい者等に対する支援の活動指標",
         atama=_h("活動指標", "令和９年度"), moto=""),
    # --- 第5章 -----------------------------------------------------------
    dict(sho=5, midashi="障がい福祉サービス等の体系（参考）",
         atama=_h("区分", "主なサービス・事業"), moto=""),
    dict(sho=5, midashi="見込量を定めるまでの段階",
         atama=_h("段階", "何を見るか"), moto=""),
    dict(sho=5, midashi="訪問系サービスの実績",
         atama=_h("サービス種別", "単位"), moto=A_MURA),
    dict(sho=5, midashi="訪問系サービスの見込量",
         atama=_h("サービス種別", "単位"), moto=""),
    dict(sho=5, midashi="訪問系サービスの利用がない場合に確かめる事項",
         atama=_h("確かめる側", "確かめる事項"), moto=""),
    dict(sho=5, midashi="日中活動系サービスの実績",
         atama=_h("サービス種別", "単位"), moto=A_MURA),
    dict(sho=5, midashi="日中活動系サービスの見込量",
         atama=_h("サービス種別", "単位"), moto=""),
    dict(sho=5, midashi="強度行動障害を有する方等の生活介護の利用の見込み",
         atama=_h("区分", "現在の人数"), moto=""),
    dict(sho=5, midashi="居住系サービスの実績",
         atama=_h("サービス種別", "単位"), moto=A_MURA),
    dict(sho=5, midashi="居住系サービスの見込量",
         atama=_h("サービス種別", "単位"), moto=""),
    dict(sho=5, midashi="相談支援の実績",
         atama=_h("サービス種別", "単位"), moto=A_MURA),
    dict(sho=5, midashi="相談支援の見込量",
         atama=_h("サービス種別", "単位"), moto=""),
    dict(sho=5, midashi="障がい児支援の実績",
         atama=_h("サービス種別", "単位"), moto=A_MURA),
    dict(sho=5, midashi="障がい児支援の見込量",
         atama=_h("サービス種別", "単位"), moto=""),
    dict(sho=5, midashi="一般の子育て施策と障がい福祉の切り分け",
         atama=_h("場面", "まず確かめる一般の子育て施策"), moto=""),
    dict(sho=5, midashi="事業所の受入余力について照会する事項",
         atama=_h("照会する事項", "確保できる量との関係"), moto=""),
    dict(sho=5, midashi="緊急時に働く仕組み（４つの場面）",
         atama=_h("場面", "想定される事態"), moto=""),
    dict(sho=5, midashi="災害時にサービスを止めないための備え",
         atama=_h("備える事項", "本村の取組み"), moto=""),
    dict(sho=5, midashi="災害時の情報の伝達について確かめること",
         atama=_h("確かめること", "確かめ方"), moto=""),
    dict(sho=5, midashi="虐待の予防から再発の防止までの流れ",
         atama=_h("段階", "本村の取組み"), moto=""),
    dict(sho=5, midashi="経営基盤に関して把握する事項",
         atama=_h("把握する事項", "見込量への影響"), moto=""),
    dict(sho=5, midashi="冬季の支援を検討する三つの段階",
         atama=_h("段階", "確かめること"), moto=""),
    dict(sho=5, midashi="冬季の通行規制による中断について記録する事項",
         atama=_h("記録する事項", "記録の仕方"), moto=""),
    dict(sho=5, midashi="医療へのアクセス",
         atama=_h("場面", "主な医療資源"), moto=""),
    dict(sho=5, midashi="地域生活支援事業の実績",
         atama=_h("サービスの種類", "令和６年度"), moto=A_MURA),
    dict(sho=5, midashi="地域生活支援事業の見込量",
         atama=_h("サービスの種類", "令和９年度"), moto=""),
    dict(sho=5, midashi="実績がない事業の扱い",
         atama=_h("実績がない理由", "確かめ方"), moto=""),
    dict(sho=5, midashi="地域生活支援事業の実施に必要な事項",
         atama=_h("事項", "本村の取扱い"), moto=""),
    dict(sho=5, midashi="本人に代わって伝えるための道具の使い分け",
         atama=_h("道具", "伝える場面"), moto=""),
    dict(sho=5, midashi="年齢到達に伴うサービスの移行",
         atama=_h("年齢", "制度上の内容"), moto=""),
    dict(sho=5, midashi="学校卒業に伴う移行について確認する事項",
         atama=_h("確認する事項", "見込量への反映"), moto=""),
    dict(sho=5, midashi="65歳到達時のサービスの区分",
         atama=_h("区分", "該当するサービス"), moto=""),
    dict(sho=5, midashi="65歳到達時に確認する事項",
         atama=_h("確認する事項", "確認の内容"), moto=""),
    dict(sho=5, midashi="65歳到達時の引継ぎの時期と担当",
         atama=_h("時期", "行うこと"), moto=""),
    dict(sho=5, midashi="関係機関と連携の内容",
         atama=_h("分野", "関係機関"), moto=""),
    dict(sho=5, midashi="障がい福祉サービス等に係る給付費の推移",
         atama=_h("年度", "介護給付費等"), moto=A_KYUFU),
    dict(sho=5, midashi="財源構成（令和７年度・法定負担割合による試算）",
         atama=_h("区分", "金額"), moto=A_KYUFU),
    dict(sho=5,
         midashi="サービス別の給付費と計画上の位置づけ（令和７年度）",
         atama=_h("サービス", "給付費（構成比）"), moto=A_KYUFU),
    # --- 第6章 -----------------------------------------------------------
    # この2表も「（各年４月１日現在）」が表に添えられている。
    dict(sho=6, midashi="高齢化率の推移", atama=_h("年度", "高齢化率"),
         moto=A_MURA),
    dict(sho=6, midashi="精神障がい者及び知的障がい者の人数の推移",
         atama=_h("年度", "精神障がい者数（※１）"), moto=A_MURA),
    dict(sho=6, midashi="成年後見制度利用支援事業の利用状況",
         atama=_h("区分", "令和５年度"), moto=A_MURA),
    dict(sho=6, midashi="成年後見制度の利用促進に係る指標",
         atama=_h("指標", "何を数えるか"), moto=""),
    # --- 第7章 -----------------------------------------------------------
    dict(sho=7, midashi="健康21・グッドヘルスプランとの対応",
         atama=_h("項目", "健康21・グッドヘルスプラン"), moto=""),
    dict(sho=7, midashi="第五次総合振興計画における本計画の位置づけ",
         atama=_h("総合振興計画の位置づけ", "内容"), moto=""),
    dict(sho=7,
         midashi="高齢者福祉計画・介護保険事業計画と重ねて数えないための確認事項",
         atama=_h("事項", "両計画での扱い"), moto=""),
    dict(sho=7, midashi="こども・子育て計画との整合を確かめる事項",
         atama=_h("確かめる事項", "確かめ方"), moto=""),
    # --- 第8章 -----------------------------------------------------------
    dict(sho=8, midashi="北塩原村障がい者自立支援協議会委員名簿",
         atama=_h("区　分", "氏　名"), moto=""),
    dict(sho=8, midashi="計画の策定経過", atama=_h("年月日", "内容"), moto=""),
    dict(sho=8, midashi="近年の障がい者施策の動向",
         atama=_h("年", "主な動向"), moto=""),
    dict(sho=8, midashi="用語解説", atama=_h("用語", "説明"), moto=""),
]


# ===========================================================================
# 3. 番号を振る（章ごとの通し番号）
# ===========================================================================
def _bangou(items, kigou):
    cnt = {}
    for it in items:
        cnt[it["sho"]] = cnt.get(it["sho"], 0) + 1
        it["no"] = f"{kigou}{it['sho']}-{cnt[it['sho']]}"
    return items


_bangou(ZU, "図")
_bangou(HYO, "表")


# ===========================================================================
# 4. 新たに要る図表（今回の点検で挙げたもの）
# ===========================================================================
TARINAI = [
    ("ライフコースと制度の移行", "第5章5 年齢到達に伴うサービスの移行",
     "児童期→成人期→高齢期を一つの流れとして、"
     "こども施策→障がい福祉→介護保険・高齢者福祉の制度間の移行を描く。"
     "18歳到達（第5章5（1））と65歳到達（同（2））が"
     "別々の項に分かれているため、全体像を示す図がない",
     "作成済み（図５-３）。入れ先は第7章4としていたが、"
     "18歳・65歳の到達を扱う第5章5の冒頭に置いた", "高"),
    ("災害時の情報の伝達と個別避難計画の接続", "第5章2（11）",
     "ＲＴ-2で「手段があるか／届いているか／特性に応じた配慮／"
     "個別避難計画との接続」の4つを表にした。"
     "情報が届いてから避難を支えるまでの流れは"
     "時間軸のある図のほうが読める",
     "村の回答（Ｍ-83）の後に判断", "中"),
    ("計画の位置づけ図の重複", "第1章2 と 第7章冒頭",
     "図１-１（計画の位置づけ）と図７-１（他の計画との関係）は、"
     "いずれも国・県・村の計画の関係を示しており内容が重なる。"
     "どちらかに寄せるか、第1章は国・県との関係、"
     "第7章は村の関連計画との関係、と役割を分ける",
     "他メンバーと相談（Ｓ-23）", "高"),
    ("手帳所持者数の将来推計", "第2章2",
     "仕様書が定める記載事項に「手帳所持者数の将来推計」がある。"
     "現状は推移（図２-２）のみで、推計の図がない",
     "村の実績（Ｍ-28）の確定後", "高"),
    ("サービス見込量の推移", "第5章1",
     "11表のうち主なサービスについて、"
     "実績3年と見込3年をつないだ図があると増減が読める",
     "第2次算定の確定後", "中"),
    ("成年後見制度の利用の流れ", "第6章",
     "相談→村長申立て→選任→後見開始までの流れと、"
     "中核機関・家庭裁判所・村の役割分担を示す図。"
     "第6章は本文と表のみで図がない",
     "作れる（優先度は低い）", "低"),
]


# ===========================================================================
# 5. 他メンバーへの確認（Ｓ-22・Ｓ-23）
# ===========================================================================
KAKUNIN = [
    ("Ｓ-22", "他メンバー", "図２-８の表題",
     "図２-８（現在利用しているサービスと今後３年間で利用したいサービス）は、"
     "画像内の表題が右端で切れており、括弧の中が読み取れない。"
     "また「回答者数：38人（複数回答）」の注記が表題に重なっている。"
     "原データでの表題の確認と、画像の作り直しをお願いしたい",
     "令和8年10月16日"),
    ("Ｓ-23", "他メンバー", "計画の位置づけ図の重複",
     "図１-１（計画の位置づけ。第1章2）と"
     "図７-１（他の計画との関係。第7章冒頭）は、"
     "いずれも国・県・村の計画の関係を示しており内容が重なる。"
     "第1章は国・県との関係、第7章は村の関連計画との関係、"
     "と役割を分けることを提案したい。"
     "あわせて、両図とも村のこども計画を"
     "「子ども・子育て支援事業計画」と書いているが、"
     "村ホームページは「こども・子育て計画」としており、"
     "村への照会中である（Ｍ-70）",
     "令和8年10月16日"),
    ("Ｓ-24", "他メンバー", "画像内の表題の扱い",
     "他メンバー版の図16点のうち14点は、表題が画像の中に焼き込まれている。"
     "規約により図の下に番号付きの表題を置いたため、"
     "最終版では表題が二重になる。"
     "画像を作り直すときに画像内の表題を外すことを提案したい",
     "令和8年10月16日"),
]


# ===========================================================================
# 6. 自己点検
# ===========================================================================
def _norm(s):
    return re.sub(r"\s+", "", str(s))


_ZEN = str.maketrans("０１２３４５６７８９", "0123456789")


def _hikaku(s):
    """画像内の表題と比べるための形。

    数字の全角・半角（表記の作法で1桁は全角にしている）と、
    囲みの記号（＜＞）は、言い回しの違いではないため落とす。
    """
    return _norm(s).translate(_ZEN).replace("＜", "").replace("＞", "")


_SUJI = {}
for _i in range(1, 10):
    _SUJI[str(_i)] = _i
    _SUJI["０１２３４５６７８９"[_i]] = _i
    _SUJI["一二三四五六七八九"[_i - 1]] = _i


def survey(path):
    """正本から図と表を本文の順に読み取る。"""
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    from docx.oxml.ns import qn
    d = docx.Document(path)
    tag = lambda e: e.tag.split("}")[-1]      # noqa: E731
    sho = 0
    zu, hyo = [], []
    for e in d.element.body.iterchildren():
        if tag(e) == "p":
            p = Paragraph(e, d)
            t = p.text.strip()
            m = re.match(r"^第([０-９0-9一二三四五六七八九])章", t)
            if p.style.name == "Heading 1" and m:
                sho = _SUJI[m.group(1)]
            blips = e.findall(".//" + qn("a:blip"))
            # 表紙と奥付はページいっぱいの画像であり、図ではない。
            # 幅が本文の幅を大きく超えるものがあれば、その段落は数えない。
            ookii = any(int(x.get("cx")) > 19 * 360000
                        for x in e.findall(".//" + qn("wp:extent")))
            if blips and sho and not ookii:
                zu.append(dict(sho=sho, ato=""))
            elif zu and not zu[-1]["ato"] and t:
                zu[-1]["ato"] = t
        elif tag(e) == "tbl" and sho:
            tb = Table(e, d)
            if len(tb.rows) == 1 and len(tb.columns) == 1:
                continue
            head = [c.text.strip() for c in tb.rows[0].cells]
            uniq = []
            for h in head:
                if h not in uniq:
                    uniq.append(h)
            hyo.append(dict(sho=sho, atama=uniq[:2]))
    return zu, hyo


def verify():
    ng = []
    if not os.path.exists(HONPON):
        raise SystemExit(f"正本が見つかりません: {HONPON}")
    zu, hyo = survey(HONPON)

    # 表紙と奥付の画像は図ではない。正本の末尾にある分を落とす。
    zu = [z for z in zu if z["sho"]]

    if len(zu) != len(ZU):
        ng.append(f"図の数が合わない: 正本{len(zu)}点／本書{len(ZU)}点")
    if len(hyo) != len(HYO):
        ng.append(f"表の数が合わない: 正本{len(hyo)}点／本書{len(HYO)}点")

    for i, (a, b) in enumerate(zip(zu, ZU)):
        if a["sho"] != b["sho"]:
            ng.append(f"{b['no']}の章が合わない: 正本第{a['sho']}章")
    for i, (a, b) in enumerate(zip(hyo, HYO)):
        if a["sho"] != b["sho"]:
            ng.append(f"{b['no']}の章が合わない: 正本第{a['sho']}章")
            continue
        want = [_norm(x) for x in b["atama"]]
        got = [_norm(x) for x in a["atama"]]
        if got[:len(want)] != want:
            ng.append(f"{b['no']}の表頭が合わない: 正本{a['atama']}／"
                      f"本書{b['atama']}")

    # 番号が章ごとに1から通っていること
    for items, kigou in ((ZU, "図"), (HYO, "表")):
        cnt = {}
        for it in items:
            cnt[it["sho"]] = cnt.get(it["sho"], 0) + 1
            if it["no"] != f"{kigou}{it['sho']}-{cnt[it['sho']]}":
                ng.append(f"番号が通っていない: {it['no']}")

    # 表題が空でないこと／図の表題が重複していないこと
    for it in ZU + HYO:
        if not it["midashi"].strip():
            ng.append(f"表題が空: {it['no']}")
    midashi = [z["midashi"] for z in ZU]
    if len(midashi) != len(set(midashi)):
        ng.append("図の表題が重複している")

    # 画像内の表題と図の表題が食い違うものを把握していること
    chigau = [z["no"] for z in ZU
              if z["midashi_zo"]
              and _hikaku(z["midashi_zo"]) != _hikaku(z["midashi"])]
    # 図2-4・図2-7 は画像内が「＜障がい種別の内訳＞」等で何の図か分からず、
    # 図7-1 は画像内に「図表」の語が入っている。いずれも承知のうえで改めた。
    if sorted(chigau) != ["図2-4", "図2-7", "図7-1"]:
        ng.append(f"画像内の表題と違うものの一覧が変わっている: {sorted(chigau)}")

    # 他団体の名前を出していないこと
    zenbu = "\n".join([str(ZU), str(HYO), str(TARINAI), str(KAKUNIN)])
    for g in ["小野町", "金ケ崎", "金ヶ崎", "阿蘇", "札幌"]:
        if g in zenbu:
            ng.append(f"他団体の名前が入っている: {g}")

    if ng:
        print("自己点検 不合格:")
        for x in ng:
            print("   -", x)
        raise SystemExit(1)
    n_moto = sum(1 for it in ZU + HYO if it["moto"])
    print(f"  自己点検: 正本の図{len(zu)}点・表{len(hyo)}点と"
          f"本書の番号が1対1／表頭{len(HYO)}件が一致／"
          f"章ごとの通し番号／表題に空なし・図の表題に重複なし／"
          f"出所を付けるもの{n_moto}点")
    return zu, hyo


# ===========================================================================
# 7. シート
# ===========================================================================
def sheet_matome(wb):
    ws = add_sheet(
        wb, "00_結論", "図表番号・表題・図表目次の付与",
        "図24点・表84点に番号と表題を付け、目次の後に図表目次を置いた。",
        [8, 112])
    n_zo = sum(1 for z in ZU if z["midashi_zo"])
    rows = [
        ("1", f"図{len(ZU)}点・表{len(HYO)}点に章ごとの通し番号を振った"
         "（図１-１・表１-１の形）。"
         "注記ボックス（【村へ確認】等の1×1の枠）45点には振っていない。"
         "確定時に削るため、番号を振ると欠番が出るからである。"),
        ("2", "表題は表の上、図の表題は図の下に置いた。"
         "目次の後に図表目次を置き、図と表を分けて章順に並べた。"),
        ("3", f"他メンバー版の図16点のうち{n_zo}点は、"
         "表題が画像の中に焼き込まれている。"
         "本書はその表題を読み取って図の下に置いたため、"
         "最終版では表題が二重になる。"
         "画像を作り直すときに画像内の表題を外すことを提案する（Ｓ-24）。"),
        ("4", "画像内の表題のままでは何の図か分からないものが2点あった。"
         "＜障がい種別の内訳＞（身体）と＜障がい程度の内訳＞（精神）である。"
         "図表目次に並べたときに読めるよう、"
         "手帳の種類が分かる表題に改めた（図２-６・図２-９）。"),
        ("5", "図２-８は画像内の表題が右端で切れており、"
         "括弧の中が読み取れない。"
         "「回答者数：38人（複数回答）」の注記も表題に重なっている。"
         "原データでの確認と画像の作り直しが要る（Ｓ-22）。"),
        ("6", "図１-１（計画の位置づけ）と図７-１（他の計画との関係）は、"
         "いずれも国・県・村の計画の関係を示しており内容が重なる。"
         "役割を分けることを提案する（Ｓ-23）。"
         "両図とも村のこども計画を「子ども・子育て支援事業計画」と"
         "書いており、村ホームページの「こども・子育て計画」と"
         "食い違っている（Ｍ-70で照会中）。"
         "名称が変われば両図とも作り直しになる。"),
        ("7", "出所は、村の実績・アンケート・外部資料から数値を引いた"
         f"図表{sum(1 for it in ZU + HYO if it['moto'])}点にのみ付けた。"
         "手順・区分・確認事項を整理しただけの表は本計画の記述そのもので"
         "あり、出所を書くと出どころが二重になるためである"
         "（規約への補足1件）。"),
        ("8", f"新たに要る図表を{len(TARINAI)}件挙げた。"
         "いま作れて効くのは「ライフコースと制度の移行」で、"
         "18歳到達（第5章5（1））と65歳到達（同（2））が"
         "別々の節に分かれているため全体像を示す図がない。"
         "第7章4（こども・子育て計画との連携）に置くのが収まりがよい。"),
    ]
    r = 6
    for i, rec in enumerate(rows):
        r = write_row(ws, r, list(rec), alt=(i % 2 == 1))
    return ws


def sheet_zu(wb):
    ws = add_sheet(
        wb, "01_図の番号と表題", f"図 {len(ZU)}点",
        "図の表題は図の下に置く。出所はその下。",
        [10, 40, 34, 12, 34, 32])
    style_header_row(ws, 5, ["番号", "表題", "置き場所", "作成",
                             "画像内の表題", "出所"])
    r = 6
    for i, z in enumerate(ZU):
        zo = z["midashi_zo"] or "（画像内に表題なし）"
        chigau = (z["midashi_zo"]
                  and _hikaku(z["midashi_zo"]) != _hikaku(z["midashi"]))
        r = write_row(ws, r, [z["no"], z["midashi"], z["basho"], z["dare"],
                              zo, z["moto"] or "―"],
                      alt=(i % 2 == 1),
                      fills=[None, None, None,
                             OK if z["dare"] == "当方" else HI,
                             MI if chigau else None, None])
    return ws


def sheet_hyo(wb):
    ws = add_sheet(
        wb, "02_表の番号と表題", f"表 {len(HYO)}点",
        "表題は表の上に置く。注記ボックス45点には番号を振らない。",
        [10, 52, 8, 40, 34])
    style_header_row(ws, 5, ["番号", "表題", "章", "表頭（照合用）", "出所"])
    r = 6
    for i, h in enumerate(HYO):
        r = write_row(ws, r, [h["no"], h["midashi"], f"第{h['sho']}章",
                              " / ".join(h["atama"]), h["moto"] or "―"],
                      alt=(i % 2 == 1))
    return ws


def sheet_tarinai(wb):
    ws = add_sheet(
        wb, "03_新たに要る図表", "今回の点検で挙げたもの",
        "図を増やすこと自体が目的ではない。"
        "いまの構成で全体像が見えなくなっているところだけを挙げる。",
        [28, 24, 58, 28, 8])
    style_header_row(ws, 5, ["図表", "入れる先", "なぜ要るか", "扱い", "優先"])
    r = 6
    for i, rec in enumerate(TARINAI):
        r = write_row(ws, r, list(rec), alt=(i % 2 == 1),
                      fills=[None] * 4 + [MI if rec[4] == "高" else None])
    return ws


def sheet_kakunin(wb):
    ws = add_sheet(
        wb, "04_確認事項", "他メンバーへの確認（Ｓ-22〜Ｓ-24）",
        "図表一覧のＳ-19〜Ｓ-21に続けて出す。",
        [8, 12, 26, 78, 16])
    style_header_row(ws, 5, ["番号", "宛先", "件名", "内容", "期限"])
    r = 6
    for i, rec in enumerate(KAKUNIN):
        r = write_row(ws, r, list(rec), alt=(i % 2 == 1))
    return ws


def main():
    verify()
    ensure_out_dir()
    wb = Workbook()
    wb.remove(wb.active)
    for f in (sheet_matome, sheet_zu, sheet_hyo, sheet_tarinai,
              sheet_kakunin):
        f(wb)
    wb.save(OUT_FILE)
    print(f"作成: {OUT_FILE}")
    print(f"  シート数: {len(wb.sheetnames)}　{wb.sheetnames}")
    print(f"  図{len(ZU)}点（当方{sum(1 for z in ZU if z['dare'] == '当方')}・"
          f"他メンバー{sum(1 for z in ZU if z['dare'] == '他メンバー')}）／"
          f"表{len(HYO)}点")
    print(f"  新たに要る図表{len(TARINAI)}件／確認事項{len(KAKUNIN)}件")


if __name__ == "__main__":
    main()
