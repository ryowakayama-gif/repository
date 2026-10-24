# -*- coding: utf-8 -*-
"""成果物とエビデンスを分けてZIPにまとめる.

【成果物】仕様書４が求める成果品。納品対象。
  D1 計画本体と管理表　　　計画素案・骨子案・図表集・管理表・会議資料
  D2 第9期の評価と分析　　 評価、調査・地域差・推計
  D3 設計と他団体比較　　　概要版の構成案、修正指示書、他計画の参考資料

【エビデンス】成果物の数値の根拠となるデータ。納品対象ではないが、
              数値の検証と第11期への引継ぎのために添える。
  E1 データ集　　　　　　　統計・事業所・調査の集計値（Excel 4冊）
  E2 図表の画像　　　　　　計画素案及び報告書に掲載した図表（PNG 79点）

成果物のZIPには図表を入れない（エビデンス側に集約する）。
ZIP名及びファイル名は日本語のため、UTF-8フラグを立てて格納する。
"""

import os
import zipfile
import repo_paths as RP

ROOT = RP.ROOT + "/output"
DEST = os.path.join(ROOT, "zip")
EDIR = os.path.join(ROOT, "evidence")

# ---------------------------------------------------------------- 成果物
SEIKA = [
    ("第10期計画_成果物_1_計画本体と管理表.zip", [
        "第10期計画_送付資料一覧.xlsx",
        "第10期計画_成果品一覧.xlsx",
        "第10期計画_成果品の送付区分表.xlsx",
        "第10期介護保険事業計画_骨子案.docx",
        "第10期介護保険事業計画_協議用素案_令和8年8月.docx",
        "第10期介護保険事業計画_公表版.docx",
        "第10期介護保険事業計画_概要版.docx",
        "第10期計画_図表集_白黒.xlsx",
        "第10期計画_必要事項の一覧.xlsx",
        "第10期計画_資料提供依頼_第9期の施策事業実績.xlsx",
        "第10期計画_確認依頼書.docx",
        "第10期計画_業務工程管理表.xlsx",
        "第10期計画_内部確認結果への対応整理.xlsx",
        "第10期計画_業務進捗報告書_令和8年8月分.docx",
        "第10期計画_計画素案の別管理表.xlsx",
        "第10期計画_計画素案の課題整理.xlsx",
        "第10期計画_管理の統合と成果品の棚卸し.xlsx",
        "第10期計画_3町ヒアリング資料.docx",
        "第10期計画_3町別の論点整理.xlsx",
        "第10期計画_構成町の意見への回答と工程の前倒し案.docx",
        "第10期計画_策定委員会資料_交付金評価結果.docx",
        "第10期計画_ケアプラン点検実施要領（案）.docx",
        "第10期計画_給付適正化計画の位置づけと追加検討事項.xlsx",
        "第10期計画_見える化データの受領点検_福祉用具住宅改修とリハ.xlsx",
        "第10期計画_地域密着型サービス基準省令の論点整理.xlsx",
        "第10期計画_サービス見込量の算定_基準年度とパターンの比較.xlsx",
        "第10期計画_見込量の季節性の検証と施策反映の設計.xlsx",
        "第10期計画_対計画比の偏りの補正試算.xlsx",
        "第10期計画_趨勢に現れない要因の分析.xlsx",
        "第10期計画_サービス見込量算定_統合報告書.xlsx",
        "第10期計画_サービス見込量の算定.xlsx",
        "第10期計画_サービス見込量の算定_第1次概算.xlsx",
        "第10期計画_サービス見込量の算定_第2次概算.xlsx",
        "第10期計画_3町合同意見交換会資料.docx",
        "第10期計画_計画策定の進捗と今後の確認事項.docx",
        "第10期計画_10月 地域課題の整理と3町協議の確認事項.xlsx",
        "第10期計画_資料を受領できない場合のサービス見込量.docx",
        "第10期計画_認知症施策推進計画の位置づけと見込量への反映.xlsx",
        "第10期計画_地域ケア会議_受領点検とKPI設計案.xlsx",
        "第10期計画_見える化_将来推計の設定資料の受領点検.xlsx",
        "第10期計画_見える化_令和8年度実績見込み値の入力案.xlsx",
        "第10期計画_見える化_見込量の入力値（暫定）.xlsx",
        "第10期計画_所得段階別第1号被保険者数の将来推計.xlsx",
        "介護保険事業計画_サービス見込量算定の手引き.docx",
        "第10期計画_介護高齢者福祉サービスガイド_構成案.docx",
        "第10期計画_キックオフ会議資料_点検反映版.docx",
        "第10期計画_キックオフ会議ヒアリングシート.docx",
        "第10期計画_キックオフ会議資料_令和8年8月.docx",
        "第10期計画_キックオフ会議資料（更新版）の校正結果.xlsx",
        "第10期計画_キックオフ会議_トークスクリプト_60分.docx",
        "第10期計画_キックオフ会議_ファシリテーションの分岐.xlsx",
        "第10期計画_キックオフ資料の点検結果.xlsx",
        "第10期計画_キックオフ会議議事録_令和8年8月6日.odt",
        "第10期計画_キックオフ会議議事録の校正結果.xlsx",
    ], []),
    ("第10期計画_成果物_2_第9期の評価と分析.zip", [
        "第10期計画_妥当性検証報告書.xlsx",
        "第10期計画_第9期施策と調査・KPIの紐付けレビュー.xlsx",
        "第10期計画_第9期計画の評価・検証_中間報告.docx",
        "第10期計画_中間報告の根拠対照表.xlsx",
        "第10期計画_第9期施策別評価表と暫定評価ルール.xlsx",
        "第10期計画_事業所調査の照会票と確定値管理表.xlsx",
        "第10期計画_9月作業_アンケート分析の作業状況.xlsx",
        "第10期計画_町別データシート.xlsx",
        "第10期計画_調査クロス集計・分析.xlsx",
        "第10期計画_調査結果と年報実績の突合クロス集計.xlsx",
        "第10期計画_認定率の年齢調整分析.xlsx",
        "第10期計画_地域差の分析.xlsx",
        "第10期計画_世帯構成の突合.xlsx",
        "第10期計画_3町の社会資源一覧との突合.xlsx",
        "第10期計画_従業員数の重複計上の整理.xlsx",
        "第10期計画_給付実績データの受領点検.xlsx",
        "第10期計画_令和8年8月28日受領資料の点検結果.xlsx",
        "第10期計画_見える化総括表の受領点検.xlsx",
        "第10期計画_総合事業実施状況調査の受領点検.xlsx",
        "第10期計画_年報月報の受領点検.xlsx",
        "第10期計画_令和6年度決算書の受領点検.xlsx",
        "第10期計画_基金条例の確認.xlsx",
        "第10期計画_人口推計の基礎の検証.xlsx",
        "第10期計画_人口推計の補正_65歳以上75歳以上の突合.xlsx",
        "第10期計画_人口推計の基礎の変更_総合戦略ベース.xlsx",
        "第10期計画_算定方法の整合確認_国と見える化.xlsx",
        "第10期計画_交付金評価指標と全国集計の受領点検.xlsx",
        "第10期計画_広域連合において行うべき事業の優先順位.xlsx",
        "第10期計画_社人研推計の町別データの受領点検.xlsx",
        "第10期計画_将来推計_人口と認定者数.xlsx",
        "第10期計画_将来推計_第2段階_サービス見込量.xlsx",
        "第10期計画_将来推計_第3段階_給付費と保険料.xlsx",
        "第10期計画_将来推計_需要3シナリオの感度表.xlsx",
        "第10期計画_実施済み3調査の受領点検と集計.xlsx",
        "第10期計画_要介護認定データの確認.xlsx",
        "第10期計画_サービス受給者数データの確認.xlsx",
        "第10期計画_住まいと施設の公表名簿との突合.xlsx",
        "第10期計画_アンケート調査の集計分析報告書.xlsx",
        "第10期計画_実施済み調査_結果報告書.docx",
        "第10期計画_アンケート調査の分析_令和8年9月業務報告.docx",
    ], []),
    ("第10期計画_成果物_3_設計と他団体比較.zip", [
        "第10期計画_保険料と施策評価の他団体比較.xlsx",
        "第10期計画_保険料の所得段階と低所得者軽減の検証.xlsx",
        "第10期計画_追加調査報告書.xlsx",
        "第10期計画_施策体系新旧対照表.xlsx",
        "第10期計画_代表KPIの振替案と確認事項の精査.xlsx",
        "第10期計画_概要版の構成案.xlsx",
        "第10期計画素案_修正指示書_令和8年7月.xlsx",
        "第10期計画素案_修正指示書_令和8年8月.xlsx",
        "第10期計画_発注者確認事項一覧.xlsx",
        "第10期計画_レビュー対応記録.xlsx",
        "第10期計画_入力前段支援アプリの点検.xlsx",
        "第10期計画_第9期評価の自己点検記録.xlsx",
        "第10期計画_主張の根拠水準の棚卸しと再レビュー.xlsx",
        "00_全計画マスター管理表.xlsx",
        "01_共通_基本コラム部品.xlsx",
        "02_高齢者介護保険事業計画.xlsx",
        "03_障がい福祉計画.xlsx",
        "04_こども計画.xlsx",
    ], []),
]

# ------------------------------------------------ 第9期評価（工程提出用の抜粋）
# 仕様書５の令和8年8月の工程末に提出する「第9期計画の評価・検証の中間報告」と、
# その根拠となる成果品をまとめたもの。成果物ZIPとは別に、提出単位で束ねる。
TEISHUTSU = [
    ("第10期計画_第9期計画の評価・検証_令和8年8月.zip", ROOT, [
        "第10期計画_第9期計画の評価・検証_中間報告.docx",
        "第10期計画_中間報告の根拠対照表.xlsx",
        "第10期計画_妥当性検証報告書.xlsx",
        "第10期計画_第9期施策と調査・KPIの紐付けレビュー.xlsx",
        "第10期計画_第9期施策別評価表と暫定評価ルール.xlsx",
        "第10期計画_施策体系新旧対照表.xlsx",
        "第10期計画_業務進捗報告書_令和8年8月分.docx",
    ], []),
]

# ------------------------------------------------ 計画素案（読むための一式）
# 素案は［要協議］等90箇所を本文に残す運用のため、
# 決定事項一覧・別管理表・図表集と併せて読む必要がある。
SOAN = [
    ("第10期計画_計画素案一式_協議用素案（令和8年8月時点）.zip", ROOT, [
        "第10期介護保険事業計画_協議用素案_令和8年8月.docx",
        "第10期計画_発注者確認事項一覧.xlsx",
        "第10期計画_計画素案の別管理表.xlsx",
        "第10期計画_計画素案の課題整理.xlsx",
        "第10期計画_管理の統合と成果品の棚卸し.xlsx",
        "第10期計画_図表集_白黒.xlsx",
        "第10期計画素案_修正指示書_令和8年8月.xlsx",
        "第10期介護保険事業計画_骨子案.docx",
        "第10期計画_施策体系新旧対照表.xlsx",
        "第10期計画_代表KPIの振替案と確認事項の精査.xlsx",
    ], []),
]

# ------------------------------------------------ 3町のヒアリングと打合せ
# 構成3町（東川町・美瑛町・東神楽町）の町別ヒアリング及び打合せで用いる資料を、
# 使う場面ごとに4つに分けて1つにまとめたもの。
# 町ごとの個別協議（令和8年10月中旬〜下旬）に持参する一式である。
# 送付区分が内部保管・対象外のものは入れない（build_machi が確かめる）。
MACHI_NAME = "第10期計画_3町ヒアリング・打合せ資料一式_令和8年10月.zip"
MACHI_GROUPS = [
    ("00_事前打合せ", [
        ("第10期計画_ヒアリング前事前打合せ_確認事項.xlsx",
         "最低限の確認事項と時間配分（A群・B群・C群）、"
         "ヒアリングの前に広域連合でお決めいただく事項と、"
         "3町へ伺う事項の最終確認。月額を動かす前提とヒアリングの進め方"),
    ]),
    ("01_ヒアリング", [
        ("第10期計画_3町ヒアリング資料.docx",
         "町別ヒアリングで用いる資料。伺う事項と当方の整理を章立てで示す"),
        ("第10期計画_キックオフ会議ヒアリングシート.docx",
         "ヒアリングの記録用紙。伺った内容をその場で書き留める"),
        ("第10期計画_3町別の論点整理.xlsx",
         "町ごとの現在地と論点、3町共通の論点、協議の進め方"),
    ]),
    ("02_協議（令和8年10月）", [
        ("第10期計画_3町合同意見交換会資料.docx",
         "3町合同の意見交換会（10月上旬）で用いる資料。"
         "次第と時間配分の案・策定の状況・保険料の月額を動かす前提・"
         "令和8年9月の制度の動向・地域課題・共通の確認事項（A群・B群・C群）"),
        ("第10期計画_10月 地域課題の整理と3町協議の確認事項.xlsx",
         "地域課題15件と、町ごとの確認事項（東川町・美瑛町・東神楽町の各シート）"),
        ("第10期計画_構成町の意見への回答と工程の前倒し案.docx",
         "構成町からいただいた意見への回答と、工程を前倒しする案"),
    ]),
    ("03_町別の基礎データ", [
        ("第10期計画_町別データシート.xlsx",
         "町ごとの人口・認定者数・サービスの利用状況・総合事業（町別の3シート）"),
        ("第10期計画_3町の社会資源一覧との突合.xlsx",
         "3町の社会資源一覧と公表名簿との突合"),
        ("第10期計画_社人研推計の町別データの受領点検.xlsx",
         "町別の将来人口の基礎データの点検"),
        ("第10期計画_交付金 市町村分評価指標と3町の評価.xlsx",
         "保険者機能強化推進交付金等の評価指標と3町の得点"),
        ("第10期計画_交付金 市町村分評価指標と3町の評価.pdf",
         "同じ内容を紙面で読む形にしたもの"),
        ("第10期計画_交付金評価の取りまとめと第10期への反映.xlsx",
         "3か年の推移、全国との比較、一部の町のみ得点している項目"),
        ("第10期計画_交付金評価の取りまとめと第10期への反映.pdf",
         "同じ内容を紙面で読む形にしたもの"),
    ]),
    ("04_協議の前提となる算定", [
        ("第10期計画_サービス見込量の算定_第2次概算.xlsx",
         "第2次概算（令和8年10月30日に提示する予定の版）の組立て。"
         "第1次概算からの差分と、月額を動かす前提"),
        ("第10期計画_サービス見込量の算定.xlsx",
         "サービス見込量・給付費・保険料の算定"),
        ("第10期計画_資料を受領できない場合のサービス見込量.docx",
         "資料のご提供がない場合に確定する値と、確定しない値"),
        ("第10期計画_認知症施策推進計画の位置づけと見込量への反映.xlsx",
         "市町村認知症施策推進計画を3町が策定することの整理と、見込量への反映"),
    ]),
]


def build_machi(missing):
    """3町のヒアリングと打合せの資料を、使う場面ごとに分けて1つにまとめる。"""
    from data_dispatch import dispatch_of

    lines = ["第10期介護保険事業計画　3町ヒアリング・打合せ資料一式",
             "",
             "大雪地区広域連合の構成3町（東川町・美瑛町・東神楽町）の",
             "町別ヒアリング及び打合せで用いる資料をまとめたものです。",
             "使う場面ごとに4つのフォルダに分けています。",
             "",
             "「お諮りする内容を含む」とあるものは、ご判断・ご協議をお願いする",
             "事項や当方の仮置きを含みます。ご決定の内容により置き直します。",
             ""]
    n, bad = 0, []
    with zipfile.ZipFile(os.path.join(DEST, MACHI_NAME), "w",
                         zipfile.ZIP_DEFLATED) as z:
        for sub, items in MACHI_GROUPS:
            lines.append("■ " + sub)
            for fn, yoto in items:
                src = os.path.join(ROOT, fn)
                if not os.path.exists(src):
                    missing.append(fn)
                    continue
                kb = dispatch_of(fn)[0]
                if kb in ("内部保管", "対象外"):
                    bad.append(fn)
                    continue
                add(z, "%s/%s" % (sub, fn), src)
                n += 1
                lines.append("  ・%s" % fn)
                lines.append("    %s" % yoto)
                lines.append("    区分：%s"
                             % ("そのまま送付できるもの" if kb == "送付"
                                else "お諮りする内容を含むもの"))
            lines.append("")
        txt = "\n".join(lines).rstrip() + "\n"
        info = zipfile.ZipInfo("00_この一式について.txt")
        info.flag_bits |= 0x800
        info.compress_type = zipfile.ZIP_DEFLATED
        z.writestr(info, txt.encode("utf-8"))
        n += 1

    size = os.path.getsize(os.path.join(DEST, MACHI_NAME)) / 1024
    print("  %-52s %3d件 %6.0f KB" % (MACHI_NAME, n, size))
    if bad:
        print("  内部保管・対象外が混ざっています:", "、".join(bad))
    return n, txt, bad


# ------------------------------------ 見込量算定と保険料の検証用一式
# 第三者（他のレビューを含む）が、サービス見込量・給付費・保険料の算定を
# 検証するための一式。算定の本体・前提の検証・人口と認定者数の推計・
# 保険料・見える化システムへの入力・交付金と認知症施策・原データを束ねる。
# 検証の依頼（人口推計の案Cの妥当性）をMarkdownで同梱する。
# 個票（個人情報）は収録しない。
MIKOMI_NAME = "第10期計画_見込量算定と保険料の検証用一式_令和8年10月.zip"
MIKOMI_GROUPS = [
    ("00_検証の依頼", [
        ("検証依頼_案Cの妥当性.md",
         "人口推計の案Cの置き方・数値・他の置き方との対比・"
         "保険料に及ぶ経路と大きさ。確かめていただきたい論点13件"),
        ("交付金の獲得状況と第10期への反映.md",
         "保険者機能強化推進交付金等の3町の得点の推移、"
         "全国及び北海道との比較、同様の構成の団体との比較、"
         "第10期で取り組む候補、調整交付金の減額"),
        ("認知症施策の位置づけと取組の状況.md",
         "市町村認知症施策推進計画を構成3町が策定することの整理、"
         "認知症高齢者の状況、見込量への反映、3町の取組の状況"),
    ]),
    ("01_算定の本体", [
        ("第10期計画_サービス見込量の算定.xlsx",
         "サービス見込量・利用回（日）数・給付費・保険料の算定（正本）。"
         "要介護度別の伸びによる。中長期は令和32年度まで"),
        ("第10期計画_サービス見込量の算定_第2次概算.xlsx",
         "第2次概算（令和8年10月30日に提示する予定の版）の組立て。"
         "第1次概算からの差分と、月額を動かす前提"),
        ("第10期計画_サービス見込量の算定_第1次概算.xlsx",
         "令和8年9月26日にお示しした版。以後作り直していない"),
        ("第10期計画_サービス見込量算定_統合報告書.xlsx",
         "算定の全体像・採用案・月額への効きの統合一覧・"
         "記載事項の網羅点検"),
        ("第10期計画_資料を受領できない場合のサービス見込量.docx",
         "資料のご提供がない場合に確定する値と、確定しない値。"
         "据え置きを区分A（当方で確定できる）・B（資料待ち）・"
         "C（国の告示待ち）に分けたもの"),
        ("第10期介護保険事業計画_協議用素案_令和8年8月.docx",
         "計画素案。第6章が見込量・給付費・保険料に当たる"),
    ]),
    ("02_算定の前提の検証", [
        ("第10期計画_サービス見込量の算定_基準年度とパターンの比較.xlsx",
         "基準年度の月数の実測と、算定できた7パターンの比較"),
        ("第10期計画_見込量の季節性の検証と施策反映の設計.xlsx",
         "季節性の測定と年度換算係数。"
         "完結年度を基準とする限り効きが消えること"),
        ("第10期計画_対計画比の偏りの補正試算.xlsx",
         "第7期からの8年度の対計画比と、"
         "サービス別の趨勢を反映する場合の補正の試算"),
        ("第10期計画_趨勢に現れない要因の分析.xlsx",
         "1人1月あたり給付費の趨勢、介護報酬改定、"
         "趨勢に現れない要因12件"),
        ("第10期計画_算定方法の整合確認_国と見える化.xlsx",
         "国の様式と見える化システムの算定式との対照"),
    ]),
    ("03_人口と認定者数の推計", [
        ("第10期計画_人口推計の基礎の検証.xlsx",
         "人口の基礎として用いる統計の検証"),
        ("第10期計画_人口推計の基礎の変更_総合戦略ベース.xlsx",
         "人口の基礎を地方創生総合戦略とするご指示への対応。"
         "総合戦略の目標人口の制約、案Cの組立て、3案の比較"),
        ("第10期計画_人口推計の補正_65歳以上75歳以上の突合.xlsx",
         "推計人口と第1号被保険者数の突合と補正"),
        ("第10期計画_人口推計コーホート分析の妥当性検証.xlsx",
         "コーホート変化率法による置き方の点検と、"
         "同じ資料で組み直した場合の対比"),
        ("第10期計画_社人研推計の町別データの受領点検.xlsx",
         "町別の将来人口の基礎データの点検"),
        ("第10期計画_将来推計_人口と認定者数.xlsx",
         "将来推計 第1段階。人口と認定者数"),
        ("第10期計画_将来推計_第2段階_サービス見込量.xlsx",
         "将来推計 第2段階。サービス見込量"),
        ("第10期計画_将来推計_第3段階_給付費と保険料.xlsx",
         "将来推計 第3段階。給付費と保険料"),
        ("第10期計画_将来推計_需要3シナリオの感度表.xlsx",
         "認定者数のシナリオ別の感度"),
        ("第10期計画_認定率の年齢調整分析.xlsx",
         "認定率の年齢調整と全国・北海道との比較"),
    ]),
    ("04_保険料", [
        ("第10期計画_所得段階別第1号被保険者数の将来推計.xlsx",
         "16段階×8年度の推計と、国の標準13段階へ束ねた場合の対比"),
        ("第10期計画_保険料の所得段階と低所得者軽減の検証.xlsx",
         "所得段階の区分と乗率、低所得者に対する公費軽減の検証"),
        ("第10期計画_保険料と施策評価の他団体比較.xlsx",
         "保険料の水準と施策評価の比較"),
        ("第10期計画_基金条例の確認.xlsx",
         "介護給付費準備基金の条例の確認"),
        ("第10期計画_令和6年度決算書の受領点検.xlsx",
         "令和6年度の決算額。上乗せ率と地域支援事業費の基礎"),
    ]),
    ("05_見える化システムへの入力", [
        ("第10期計画_見える化_入力箇所と入力値.xlsx",
         "画面の順に、値を変える箇所と入れる値をまとめたもの"),
        ("第10期計画_見える化_見込量の入力値（暫定）.xlsx",
         "要介護度別・全8年度の入力値"),
        ("第10期計画_見える化_令和8年度実績見込み値の入力案.xlsx",
         "令和8年度の実績見込み値を令和7年度の月平均に差し替える案"),
        ("第10期計画_見える化_将来推計の設定資料の受領点検.xlsx",
         "推計方法の設定画面の参考シートの点検"),
        ("第10期計画_見える化_推計結果の受領点検.xlsx",
         "実際に推計した結果と当方の算定との対照。"
         "手当てを要する箇所とワーニングチェック"),
        ("第10期計画_見える化の利用マニュアル等の受領点検.xlsx",
         "利用マニュアル・地域分析の手引き・事業所分析データの点検。"
         "実績値の期間と地域支援事業費の額の定義"),
        ("第10期計画_見える化の指標（B5・B6・D49）の受領点検.xlsx",
         "調整済み認定率・重度認定率・地域差指数の指標の定義"),
        ("第10期計画_見える化総括表の受領点検.xlsx",
         "総括表及び総括表詳細。基準年度の実績の出所"),
        ("第10期計画_見える化システムからご出力をお願いする資料.xlsx",
         "保険者でなければ操作できない画面と、ご出力をお願いする資料"),
    ]),
    ("06_交付金と認知症施策", [
        ("第10期計画_交付金 市町村分評価指標と3町の評価.xlsx",
         "評価指標の構造、大項目53件・明細列299件の配点と3町の得点"),
        ("第10期計画_交付金 市町村分評価指標と3町の評価.pdf",
         "同じ内容を紙面で読む形にしたもの"),
        ("第10期計画_交付金評価の取りまとめと第10期への反映.xlsx",
         "3か年の推移、全国との比較、同様の構成の団体との比較、"
         "必要項目の洗い出しと対応可否"),
        ("第10期計画_交付金評価の取りまとめと第10期への反映.pdf",
         "同じ内容を紙面で読む形にしたもの"),
        ("第10期計画_交付金評価指標と全国集計の受領点検.xlsx",
         "評価指標及び全国集計の受領点検"),
        ("第10期計画_認知症施策推進計画の位置づけと見込量への反映.xlsx",
         "市町村認知症施策推進計画の位置づけ3案の検討と、"
         "認知症に係るサービスの見込量"),
    ]),
    ("07_実績と原典の点検", [
        ("第10期計画_年報月報の受領点検.xlsx",
         "介護保険事業状況報告（年報・月報）の点検。"
         "見込量の数量の基礎"),
        ("第10期計画_給付実績データの受領点検.xlsx",
         "給付費データの点検"),
        ("第10期計画_給付費データ・地域支援事業実施報告の受領点検.xlsx",
         "令和7年度の給付費データの再提供と、"
         "3町の地域包括支援センター事業実施報告書の点検"),
        ("第10期計画_総合事業実施状況調査の受領点検.xlsx",
         "総合事業の実施状況に関する調査。地域支援事業の量の基礎"),
        ("第10期計画_利用状況統計表（R8.8月審査分）の受領点検.xlsx",
         "審査月別・町別・サービス種類別・要介護度別の給付費"),
    ]),
    ("08_原データ", [
        ("第10期計画_エビデンス_0_索引.xlsx", "原データの索引"),
        ("第10期計画_エビデンス_1_統計データ.xlsx",
         "人口・認定・給付の統計データ"),
        ("第10期計画_エビデンス_2_事業所と公表情報.xlsx",
         "事業所の一覧と公表情報"),
        ("第10期計画_エビデンス_3_調査の集計値.xlsx",
         "実施済み調査の集計値"),
    ]),
]


def build_mikomi(missing):
    """見込量算定と保険料の検証用一式を、検証の順に分けて1つにまとめる。"""
    from data_dispatch import dispatch_of

    bases = [ROOT, EDIR, os.path.join(RP.ROOT, "docs")]
    lines = ["第10期介護保険事業計画　見込量算定と保険料の検証用一式",
             "",
             "サービス見込量・給付費・保険料の算定を検証していただくための",
             "一式です。検証の順に9つのフォルダに分けています。",
             "",
             "はじめに 00_検証の依頼 の「検証依頼_案Cの妥当性」をお読みください。",
             "人口推計の案Cの置き方と、確かめていただきたい論点を示しています。",
             "",
             "「お諮りする内容を含む」とあるものは、ご判断・ご協議をお願いする",
             "事項や当方の仮置きを含みます。ご決定の内容により置き直します。",
             "「内部の検証」とあるものは、当方が算定の妥当性を確かめるために",
             "作成したもので、発注者への送付を想定していません。",
             "",
             "本一式に個人を特定できる情報は含まれていません。",
             "いずれも集計した値です。",
             ""]
    KB = {"送付": "そのまま送付できるもの",
          "条件付き": "お諮りする内容を含むもの",
          "内部保管": "内部の検証",
          "未分類": "検証の依頼"}
    n, bad = 0, []
    with zipfile.ZipFile(os.path.join(DEST, MIKOMI_NAME), "w",
                         zipfile.ZIP_DEFLATED) as z:
        for sub, items in MIKOMI_GROUPS:
            lines.append("■ " + sub)
            for fn, yoto in items:
                src = next((os.path.join(b, fn) for b in bases
                            if os.path.exists(os.path.join(b, fn))), None)
                if src is None:
                    missing.append(fn)
                    continue
                kb = dispatch_of(fn)[0]
                if kb == "対象外":
                    bad.append(fn)
                    continue
                add(z, "%s/%s" % (sub, fn), src)
                n += 1
                lines.append("  ・%s" % fn)
                lines.append("    %s" % yoto)
                if kb in KB:
                    lines.append("    区分：%s" % KB[kb])
            lines.append("")
        txt = "\n".join(lines).rstrip() + "\n"
        info = zipfile.ZipInfo("00_この一式について.txt")
        info.flag_bits |= 0x800
        info.compress_type = zipfile.ZIP_DEFLATED
        z.writestr(info, txt.encode("utf-8"))
        n += 1

    size = os.path.getsize(os.path.join(DEST, MIKOMI_NAME)) / 1024
    print("  %-52s %3d件 %6.0f KB" % (MIKOMI_NAME, n, size))
    if bad:
        print("  対象外が混ざっています:", "、".join(bad))
    return n, txt, bad


# ------------------------------------------------ 中間報告の検証用エビデンス
# 第三者が中間報告の記述を検証するための一式。
# 報告本体・根拠対照表・根拠となる成果品・原データを1つにまとめる。
# 個票（個人情報）及び見える化δは収録しない（根拠対照表 05シートに明示）。
KENSHO_ROOT = [
    "第10期計画_第9期計画の評価・検証_中間報告.docx",
    "第10期計画_中間報告の根拠対照表.xlsx",
    "第10期計画_妥当性検証報告書.xlsx",
    "第10期計画_第9期施策と調査・KPIの紐付けレビュー.xlsx",
    "第10期計画_第9期施策別評価表と暫定評価ルール.xlsx",
    "第10期計画_施策体系新旧対照表.xlsx",
    "第10期計画_認定率の年齢調整分析.xlsx",
    "第10期計画_地域差の分析.xlsx",
    "第10期計画_調査クロス集計・分析.xlsx",
    "第10期計画_調査結果と年報実績の突合クロス集計.xlsx",
    "第10期計画_アンケート調査の集計分析報告書.xlsx",
    "第10期計画_サービス受給者数データの確認.xlsx",
    "第10期計画_要介護認定データの確認.xlsx",
    "第10期計画_業務進捗報告書_令和8年8月分.docx",
]
KENSHO_EVID = [
    "第10期計画_エビデンス_0_索引.xlsx",
    "第10期計画_エビデンス_1_統計データ.xlsx",
    "第10期計画_エビデンス_2_事業所と公表情報.xlsx",
    "第10期計画_エビデンス_3_調査の集計値.xlsx",
]

# ---------------------------------------------------------------- エビデンス
EVID = [
    ("第10期計画_エビデンス_1_データ集.zip", EDIR, [
        "第10期計画_エビデンス_0_索引.xlsx",
        "第10期計画_エビデンス_1_統計データ.xlsx",
        "第10期計画_エビデンス_2_事業所と公表情報.xlsx",
        "第10期計画_エビデンス_3_調査の集計値.xlsx",
    ], []),
    ("第10期計画_エビデンス_2_図表の画像.zip", ROOT, [],
     ["figures", "figures_report", "images_basic"]),
]


def build_by_dispatch(missing):
    """送付区分により、送付用と内部保管用のZIPを分けて作る。"""
    from data_dispatch import DISPATCH
    groups = {"送付": [], "条件付き": [], "内部保管": []}
    for fn, (kb, _why) in DISPATCH.items():
        if kb in groups and os.path.exists(os.path.join(ROOT, fn)):
            groups[kb].append(fn)

    n_sofu = 0
    nm = "第10期計画_送付用_令和8年8月.zip"
    with zipfile.ZipFile(os.path.join(DEST, nm), "w",
                         zipfile.ZIP_DEFLATED) as z:
        for sub, kb in [("01_送付可能（留保明記済み）", "送付"),
                        ("02_お諮りする内容を含むもの", "条件付き")]:
            for fn in sorted(groups[kb]):
                add(z, "%s/%s" % (sub, fn), os.path.join(ROOT, fn))
                n_sofu += 1
    print("  %-52s %3d件 %6.0f KB"
          % (nm, n_sofu, os.path.getsize(os.path.join(DEST, nm)) / 1024))

    # 計画本文に掲載する図表の画像を同送する
    fig = os.path.join(ROOT, "figures")
    n_png = 0
    if os.path.isdir(fig):
        with zipfile.ZipFile(os.path.join(DEST, nm), "a",
                             zipfile.ZIP_DEFLATED) as z:
            for f in sorted(os.listdir(fig)):
                add(z, "03_図表の画像/" + f, os.path.join(fig, f))
                n_png += 1
                n_sofu += 1
        print("  %-52s 図表の画像 %d点を同梱" % ("", n_png))

    doc = RP.ROOT + "/docs/令和8年8月_送付状.md"
    if os.path.exists(doc):
        with zipfile.ZipFile(os.path.join(DEST, nm), "a",
                             zipfile.ZIP_DEFLATED) as z:
            add(z, "00_送付状_令和8年8月.md", doc)
            n_sofu += 1
        print("  %-52s 送付状を同梱" % "")

    nm2 = "第10期計画_内部保管_作業記録.zip"
    n_naibu = 0
    with zipfile.ZipFile(os.path.join(DEST, nm2), "w",
                         zipfile.ZIP_DEFLATED) as z:
        for fn in sorted(groups["内部保管"]):
            add(z, fn, os.path.join(ROOT, fn))
            n_naibu += 1
    print("  %-52s %3d件 %6.0f KB"
          % (nm2, n_naibu, os.path.getsize(os.path.join(DEST, nm2)) / 1024))
    return n_sofu, n_naibu


def add(z, arc, src):
    info = zipfile.ZipInfo.from_file(src, arc)
    info.flag_bits |= 0x800                      # ファイル名がUTF-8である旨
    info.compress_type = zipfile.ZIP_DEFLATED
    with open(src, "rb") as f:
        z.writestr(info, f.read())


def build(name, base, files, dirs, missing):
    n = 0
    with zipfile.ZipFile(os.path.join(DEST, name), "w",
                         zipfile.ZIP_DEFLATED) as z:
        for fn in files:
            src = os.path.join(base, fn)
            if not os.path.exists(src):
                missing.append(fn)
                continue
            add(z, fn, src)
            n += 1
        for d in dirs:
            p = os.path.join(base, d)
            if not os.path.isdir(p):
                missing.append(d + "/")
                continue
            for fn in sorted(os.listdir(p)):
                add(z, "%s/%s" % (d, fn), os.path.join(p, fn))
                n += 1
    size = os.path.getsize(os.path.join(DEST, name)) / 1024
    print("  %-52s %3d件 %6.0f KB" % (name, n, size))
    return n


if __name__ == "__main__":
    os.makedirs(DEST, exist_ok=True)
    # 旧名のZIPを残さない
    for fn in os.listdir(DEST):
        if fn.startswith("第10期計画_成果品_"):
            os.remove(os.path.join(DEST, fn))

    missing = []
    print("【成果物】")
    ns = sum(build(nm, ROOT, f, d, missing) for nm, f, d in SEIKA)
    print("【エビデンス】")
    ne = sum(build(nm, base, f, d, missing) for nm, base, f, d in EVID)
    print("【計画素案（読むための一式）】")
    nso = sum(build(nm, base, f, d, missing) for nm, base, f, d in SOAN)
    print("【工程提出用（第9期評価）】")
    nt = sum(build(nm, base, f, d, missing) for nm, base, f, d in TEISHUTSU)

    print("【3町ヒアリング・打合せ資料一式】")
    nm_, machi_txt, machi_bad = build_machi(missing)

    print("【見込量算定と保険料の検証用一式】")
    nmi_, mikomi_txt, mikomi_bad = build_mikomi(missing)

    print("【送付区分別】")
    n_so, n_na = build_by_dispatch(missing)

    # 検証用は1つのZIPに2つの階層で格納する
    print("【中間報告の検証用エビデンス】")
    _nm = "第10期計画_中間報告の検証用エビデンス.zip"
    nk = 0
    with zipfile.ZipFile(os.path.join(DEST, _nm), "w",
                         zipfile.ZIP_DEFLATED) as z:
        for fn in KENSHO_ROOT:
            src = os.path.join(ROOT, fn)
            if not os.path.exists(src):
                missing.append(fn)
                continue
            add(z, "01_報告と根拠/" + fn, src)
            nk += 1
        for fn in KENSHO_EVID:
            src = os.path.join(EDIR, fn)
            if not os.path.exists(src):
                missing.append(fn)
                continue
            add(z, "02_原データ/" + fn, src)
            nk += 1
        d = os.path.join(ROOT, "figures")
        for fn in sorted(os.listdir(d)):
            add(z, "03_図表/" + fn, os.path.join(d, fn))
            nk += 1
    _sz = os.path.getsize(os.path.join(DEST, _nm)) / 1024
    print("  %-52s %3d件 %6.0f KB" % (_nm, nk, _sz))
    print()
    print("成果物 %d件／エビデンス %d件／素案一式 %d件／"
          "工程提出用 %d件／検証用 %d件" % (ns, ne, nso, nt, nk))
    print("3町ヒアリング・打合せ %d件（%d群）"
          % (nm_, len(MACHI_GROUPS)))
    print("見込量算定と保険料の検証用 %d件（%d群）"
          % (nmi_, len(MIKOMI_GROUPS)))
    print("送付用 %d件／内部保管 %d件" % (n_so, n_na))
    print("欠落: %s" % ("なし" if not missing else "、".join(missing)))

    # 成果品一覧との突合（一覧にあってZIPに入っていないものを検出する）
    import io as _io
    import runpy as _runpy
    import sys as _sys
    _buf, _old = _io.StringIO(), _sys.stdout
    _sys.stdout = _buf
    try:
        _G = _runpy.run_path(RP.ROOT + "/build_deliverable_index.py")
    finally:
        _sys.stdout = _old
    idx = {fn for _c, fn, _a, _b, _d in _G["ITEMS"]}
    zipped = set()
    for _nm, _f, _d2 in SEIKA:
        zipped |= set(_f)
    gap = sorted(idx - zipped)
    print("成果品一覧との突合: %s"
          % ("全件がZIPに入っている" if not gap else "ZIP未収録 " + "、".join(gap)))

    # 送付用ZIPに内部保管のものが混ざっていないことを確かめる
    from data_dispatch import dispatch_of
    with zipfile.ZipFile(
            os.path.join(DEST, "第10期計画_送付用_令和8年8月.zip")) as z:
        bad = [a for a in z.namelist()
               if dispatch_of(os.path.basename(a))[0] in ("内部保管", "対象外")]
    print("送付用ZIPの点検: %s"
          % ("内部保管・対象外は含まれていない" if not bad
             else "混入 " + "、".join(bad)))

    # 3町ヒアリング・打合せ資料一式の点検
    import re as _re
    with zipfile.ZipFile(os.path.join(DEST, MACHI_NAME)) as z:
        _names = z.namelist()
    _mbad = [a for a in _names if not a.endswith(".txt")
             and dispatch_of(os.path.basename(a))[0] in ("内部保管", "対象外")]
    _msub = sorted({a.split("/")[0] for a in _names if "/" in a})
    _NAIBU = ["固定値", "実物から", "章節ごとに", "判定している", "読んで判定",
              "スクリプト", "runpy", ".py", "再実行", "書き写して",
              "個票がなくても", "モジュール", "終了コード"]
    _NG = ["に由来する", "と整合する", "1件も", "有意差がないため関係がない",
           "全国トップ級"]
    _mn = [w for w in _NAIBU if w in machi_txt]
    _mg = [w for w in _NG if w in machi_txt]
    _mp = _re.search(r"0\d{1,3}-\d{2,4}-\d{4}|[\w.+-]+@[\w-]+\.[\w.]+",
                     machi_txt)
    print("3町ヒアリング・打合せZIPの点検: 群%d（%s）／%s／%s／%s／%s"
          % (len(_msub), "・".join(_msub),
             "内部保管・対象外なし" if not _mbad and not machi_bad
             else "混入 " + "、".join(_mbad + machi_bad),
             "内部の仕組みの語なし" if not _mn else "検出 " + "・".join(_mn),
             "禁止表現なし" if not _mg else "検出 " + "・".join(_mg),
             "個人情報の形なし" if not _mp else "検出あり"))
    # 見込量算定と保険料の検証用一式の点検
    with zipfile.ZipFile(os.path.join(DEST, MIKOMI_NAME)) as z:
        _knames = z.namelist()
    _kbad = [a for a in _knames if not a.endswith(".txt")
             and dispatch_of(os.path.basename(a))[0] == "対象外"]
    _ksub = sorted({a.split("/")[0] for a in _knames if "/" in a})
    _kmd = [a for a in _knames if a.endswith(".md")]
    _kn = [w for w in _NAIBU if w in mikomi_txt]
    _kg = [w for w in _NG if w in mikomi_txt]
    _kp = _re.search(r"0\d{1,3}-\d{2,4}-\d{4}|[\w.+-]+@[\w-]+\.[\w.]+",
                     mikomi_txt)
    print("見込量算定と保険料の検証用ZIPの点検: 群%d／検証の依頼%d件／"
          "%s／%s／%s／%s"
          % (len(_ksub), len(_kmd),
             "対象外なし" if not _kbad and not mikomi_bad
             else "混入 " + "、".join(_kbad + mikomi_bad),
             "内部の仕組みの語なし" if not _kn else "検出 " + "・".join(_kn),
             "禁止表現なし" if not _kg else "検出 " + "・".join(_kg),
             "個人情報の形なし" if not _kp else "検出あり"))

    if (_mbad or machi_bad or _mn or _mg or _mp
            or _kbad or mikomi_bad or _kn or _kg or _kp or len(_kmd) != 3):
        import sys as _sys2
        _sys2.exit(1)
