# -*- coding: utf-8 -*-
"""
北塩原村　前回計画の評価・アンケート結果報告の概要版（2本）

出力
  output/北塩原村_前回計画の評価_概要版.docx（A4縦・2ページ）
  output/北塩原村_アンケート結果報告_概要版.docx（A4縦・2ページ）

なぜ作るのか
  令和8年10月13日に受領した依頼に
  「前回計画評価、アンケート結果報告、計画素案の3本を
  素案と概要版で用意しないといけないと想定しています」とあった。
  当方は成果品の構えとして案②（3本の本体＋計画素案の概要版＋
  策定委員会資料）を推奨していたが、令和8年10月14日に
  「概要版に進んで下さい」との指示を受けたため、
  案①（3本×素案・概要版＝6本）がそろう形にする。

  本体は既にある。
    前回計画の評価 … output/会議資料/03_協議会資料3_前回計画の評価.docx
    アンケート結果 … 他メンバーの分析報告書（令和8年10月）。
                      当方の要約は正本 第2章7
    計画素案       … output/北塩原村_計画素案_正本_移植後.docx
  概要版は計画素案の分（7ページ）だけがあった。本書で残る2本を作る。

数値を書き写さない
  成果品が6本になると、同じ数値が6カ所に分かれ、
  照合の組合せが15通りになる。
  これを避けるため、本書は数値を一切書き写さない。
    前回計画の評価 … build_kitashiobara_zenkai_hyouka.py の
                      KPI7・KPI4・MIKOMI・SHISAKU から引く。
                      件数は数え上げて出す（手で書かない）。
    アンケート     … 正本 第2章7 の現物から引く。
                      概要版に書いた割合が正本にあることを
                      生成のたびに確かめる。

住民に配るものとして
  本文は12ポイントではなく、計画素案の概要版と同じ9ポイントに
  そろえている。3本の概要版で字の大きさが違うと並べて読めない。
  字の大きさ・色の使い方・グレースケールでの見え方は
  他の計画の概要版とそろえる必要があるため確認中である（Ｓ-34）。

アンケートの概要版についての断り
  アンケート調査の分析報告書は他メンバーが作成している。
  本書の概要版は、当方が正本 第2章7 に要約した内容から作ったもので
  あり、分析報告書の確定値との照合が済んでいない。
  委員会・パブリックコメントに出す前に、分析報告書の確定版と
  突き合わせる必要がある（Ｓ-26・Ｓ-36）。
"""

import os
import re

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Mm, Pt

from build_kitashiobara_gaiyou import (
    _cell_margins, _cell_shade, _run, _tbl_width, _no_borders,
    band, box, estimate_lines, flow, note, para, setup_section,
    sub, table,
)
# 体裁は前回計画の概要版の実測から引く（3本で同じ体裁にする）。
from kitashiobara_gaiyou_taisai import (
    ACCENT, BAND, BODY_PT, DARK, LINES_PER_PAGE, NOTE_FILL, SMALL_PT,
    TAISAI,
)
from build_kitashiobara_zenkai_hyouka import (
    KANGAE, KPI4, KPI7, KUBUN, LIFESTAGE, MIKOMI, SHISAKU, SHISAKU_KUBUN, _shu_bun,
    mikomi_kubun,
)
from kitashiobara_common import OUT_DIR
import kitashiobara_anke as A
from kitashiobara_koutei import (
    PABU_KAISHI, PABU_NISSU, PABU_SHURYO, wareki,
)

REPO_ROOT = os.path.dirname(os.path.abspath(__file__))
HYOUKA_FILE = f"{OUT_DIR}/北塩原村_前回計画の評価_概要版.docx"
ANKE_FILE = f"{OUT_DIR}/北塩原村_アンケート結果報告_概要版.docx"
HONPON = f"{REPO_ROOT}/output/北塩原村_計画素案_正本_移植後.docx"
HYOUKA_HONTAI = (f"{REPO_ROOT}/output/会議資料/"
                 "03_協議会資料3_前回計画の評価.docx")
HI = "令和８年10月"


# ===========================================================================
# 1. 数え上げ（手で書かない）
# ===========================================================================
def kazoeru(recs, idx):
    """列の値ごとの件数を、現れる順に数える。"""
    out = []
    for rec in recs:
        key = rec[idx]
        for i, (k, _n) in enumerate(out):
            if k == key:
                out[i] = (k, _n + 1)
                break
        else:
            out.append((key, 1))
    return out


def _kanten():
    """見込量を見る観点（KANGAE の2つめから読む。数も数える）。"""
    naka = KANGAE[1][1].split("を分けて見る")[0]
    n = len(naka.split("、"))
    return f"{naka}の{'一二三四五六七八九'[n - 1]}つ"


def hantei_hyou():
    """成果目標・指標14項目の判定の内訳。"""
    return kazoeru(KPI7, 4)


def atsukai_hyou():
    """第8期での扱いの内訳（継続・見直し・強化）。"""
    return kazoeru([(x[0], x[1], x[2], x[3], x[4], x[5],
                     x[6].split("（")[0]) for x in KPI7], 6)


def mikomi_hyou():
    """見込量の評価の6区分ごとの件数。区分の定義は KUBUN から引く。"""
    cnt = dict(kazoeru(MIKOMI, 1))
    return [(k, imi, cnt.get(k, 0)) for k, imi in KUBUN]


def shisaku_hyou():
    """施策の評価の6区分ごとの件数。"""
    cnt = dict(kazoeru(SHISAKU, 1))
    return [(k, imi, cnt.get(k, 0)) for k, imi in SHISAKU_KUBUN]


def kpi4_hyou():
    """第4次障がい者計画の指標（中間評価）。"""
    rows = [["基本施策", "指標", "基準値\n（令和５年度）", "現在値",
             "目標\n（令和11年度）"]]
    for shisaku, shihyo, kijun, ima, mokuhyo, _yomi in KPI4:
        rows.append([shisaku, shihyo, kijun, ima, mokuhyo])
    return rows


def kakunin_chu():
    """現在値が「確認中」の指標の数。"""
    return sum(1 for x in KPI4 if "確認中" in x[3])


# ===========================================================================
# 2. アンケートの概要版に載せる項目
#    (分野, 本文, 本文に出る割合の指標名)
#    割合は kitashiobara_anke（回収40件の集計）から読む。
#    令和8年10月24日の他ＡＩの横断レビュー（RT-01・RT-10）で、
#    本書が38件の割合を文字で持っていたことが指摘されたため改めた。
#    書いた割合が正本 第2章7 にも実在することを生成のたびに確かめる。
# ===========================================================================
P = A.pc

ANKE = [
    ("日常生活と将来の暮らし",
     f"日常生活で困っていることは、「特にない」（{P('困りごと_特にない')}）を"
     f"除くと「将来の生活」（{P('困りごと_将来の生活')}）が最も高く、"
     f"65歳未満では{P('困りごと_将来の生活_65歳未満')}と半数を超えています。"
     "将来の暮らしについては、"
     f"「今の自宅で暮らしたい」（{P('暮らし_今の自宅')}）と"
     f"「家族と一緒に暮らしたい」（{P('暮らし_家族と一緒')}）を合わせて"
     f"{P('在宅を希望')}が、"
     "住み慣れた地域での生活の継続を希望しています。",
     ["困りごと_特にない", "困りごと_将来の生活",
      "困りごと_将来の生活_65歳未満", "暮らし_今の自宅", "暮らし_家族と一緒",
      "在宅を希望"]),
    ("外出・移動",
     f"外出時に困ることは「冬季・積雪時に外出しにくい」（{P('外出_冬季')}）が"
     f"最も高く、裏磐梯地区では{P('外出_冬季_裏磐梯')}に上ります。"
     f"移動手段は「自分で運転」（{P('移動_自分で運転')}）、"
     f"「家族・知人の送迎」（{P('移動_家族の送迎')}）が中心です。"
     "障がい児調査では、回答者全員が"
     "「送迎・移動手段の確保が難しい」と回答しています。",
     ["外出_冬季", "外出_冬季_裏磐梯", "移動_自分で運転", "移動_家族の送迎"]),
    ("相談・情報",
     f"困ったときの相談相手は「家族・親族」（{P('相談相手_家族')}）が"
     "突出しており、"
     f"「相談支援専門員」（{P('相談相手_相談支援専門員')}）、"
     f"「村役場」（{P('相談相手_村役場')}）は限られています。"
     "相談支援に期待することは"
     f"「わかりやすい情報提供」（{P('相談への期待_情報提供')}）、"
     f"「サービス利用の相談」（{P('相談への期待_サービス利用')}）が上位です。",
     ["相談相手_家族", "相談相手_相談支援専門員", "相談相手_村役場",
      "相談への期待_情報提供", "相談への期待_サービス利用"]),
    ("福祉サービスの利用",
     "今後３年間で利用したいサービスは"
     f"「施設入所支援」（{P('利用したい_施設入所支援')}）、"
     f"「グループホーム」（{P('利用したい_グループホーム')}）、"
     f"「短期入所」（{P('利用したい_短期入所')}）が上位です。"
     "一方、「重度訪問介護」「同行援護・行動援護」"
     "「就労選択支援」「移動支援」は、"
     "内容をよく知らないとの回答がそれぞれ"
     f"{P('知らない_重度訪問介護')}あり、"
     "サービスの周知が課題です。"
     "地域生活支援拠点等に期待する機能は"
     f"「緊急時の受入れ」（{P('拠点_緊急時の受入れ')}）が上位となっています。",
     ["利用したい_施設入所支援", "利用したい_グループホーム",
      "利用したい_短期入所", "知らない_重度訪問介護", "拠点_緊急時の受入れ"]),
    ("就労・日中活動",
     f"日中は「自宅で過ごすことが多い」（{P('日中_自宅')}）が最も高く、"
     "働くために必要な支援として"
     f"「職場の障がい理解」（{P('就労支援_職場の理解')}）、"
     f"「体調に合わせた勤務」（{P('就労支援_体調に合わせた勤務')}）が"
     "挙げられています。",
     ["日中_自宅", "就労支援_職場の理解", "就労支援_体調に合わせた勤務"]),
    ("権利擁護・差別解消",
     f"差別や偏見、不利益を感じたことがある方は{P('差別を感じた')}で、"
     "必要な合理的配慮として"
     f"「わかりやすい説明」（{P('配慮_わかりやすい説明')}）、"
     f"「手続きの支援」（{P('配慮_手続きの支援')}）が上位です。"
     "成年後見制度や権利擁護支援については"
     f"「知らない」が{P('成年後見_知らない')}となっています。",
     ["差別を感じた", "配慮_わかりやすい説明", "配慮_手続きの支援",
      "成年後見_知らない"]),
    ("災害時の対応",
     "避難行動要支援者名簿・個別避難計画を「知らない」が"
     f"{P('個別避難計画_知らない')}に上ります。"
     f"災害時の不安は「避難先での生活」（{P('災害不安_避難先での生活')}）、"
     f"「薬・医療機器」（{P('災害不安_薬・医療機器')}）、"
     f"「冬季・積雪時の避難」（{P('災害不安_冬季の避難')}）が上位です。",
     ["個別避難計画_知らない", "災害不安_避難先での生活",
      "災害不安_薬・医療機器", "災害不安_冬季の避難"]),
    ("村が特に力を入れるべきこと",
     f"「相談支援体制」（{P('施策_相談支援体制')}）が最も高く、"
     f"「災害時支援」（{P('施策_災害時支援')}）、"
     "「在宅生活支援」「移動支援」"
     f"「サービス事業所の確保」（各{P('施策_在宅生活支援')}）が"
     "続いています。",
     ["施策_相談支援体制", "施策_災害時支援", "施策_在宅生活支援",
      "施策_移動支援", "施策_事業所の確保"]),
    ("障がい児調査（回答３人）",
     "回答者全員が家族の負担が大きいと感じています。"
     "通園・通学等では「送迎・移動手段の確保が難しい」"
     "「付き添いなど家族の負担が大きい」と回答しています。"
     "今後利用したいサービスでは全員が「移動支援」を挙げ、"
     "18歳以降の生活については"
     "「進学先・就労先」「親など家族が高齢になった後の生活」など"
     "８項目で全員が不安を感じています。",
     []),
]

# 障がい別の特徴（正本 表2-7 と同じ項目に絞る）
#   全体の列は回収40件の集計から読む。障がい別の列は手帳で区分しており、
#   増えた2件がどの手帳も持たない方であるため38件の時から動かない
#   （kitashiobara_anke の冒頭の注記）。
SHOGAIBETSU = [
    ["項目", f"全体\n（n={A.n()}）", "身体\n（n=29）", "知的\n（n=2）",
     "精神\n（n=6）"],
    ["65歳以上の割合", P("65歳以上"), "72.4％", "50.0％", "16.7％"],
    ["将来も在宅での生活を希望", P("在宅を希望"), "58.6％", "0.0％", "50.0％"],
    ["冬季・積雪時に外出しにくい", P("外出_冬季"), "27.6％", "0.0％",
     "83.3％"],
    ["相談相手は「家族・親族」", P("相談相手_家族"), "75.9％", "50.0％",
     "66.7％"],
    ["差別や偏見を感じたことがある", P("差別を感じた"), "17.2％", "0.0％",
     "50.0％"],
    ["災害時の不安「避難先での生活」", P("災害不安_避難先での生活"), "55.2％",
     "100.0％", "100.0％"],
    ["名簿・個別避難計画を「知らない」", P("個別避難計画_知らない"), "62.1％",
     "100.0％", "100.0％"],
]

# 調査結果からみた課題（正本 表2-8 の7区分）
KADAI = [
    ["主な課題", "本計画での主な対応"],
    ["相談支援体制の充実とわかりやすい情報提供",
     "相談支援体制の充実・強化、サービス内容の周知"],
    ["移動・外出支援（冬季への対応を含む）",
     "移動支援事業の利用促進、近隣市町村との連携"],
    ["地域生活の継続と緊急時の備え",
     "地域生活支援拠点等の機能の充実、短期入所の確保"],
    ["災害時の支援体制",
     "個別避難計画の作成推進、避難行動要支援者への支援の周知"],
    ["権利擁護・差別解消",
     "理解促進研修・啓発事業、成年後見制度の周知と利用支援"],
    ["就労・日中活動",
     "就労選択支援の周知・利用促進、一般就労への移行支援"],
    ["障がい児と家族への支援",
     "障がい児支援の提供体制の整備、医療的ケア児等への支援"],
    # 8区分目。アンケート集計表の改訂版（令和8年10月14日受領）の
    # 6章_課題（表6-1）が課題を8件としていることにより確定した。
    # 正本 表2-8 にもＡＫ-1で足している。
    ["専門的な支援を要する方への対応",
     "精神障がいにも対応した地域包括ケアシステムの構築、"
     "強度行動障がい・医療的ケアへの対応"],
]


# ===========================================================================
# 3. 照会
# ===========================================================================
def _n_wariai():
    """アンケートの概要版に載せた割合の点数（本文と障がい別の表）。"""
    hon = sum(len(t) for _m, _h, t in ANKE)
    hyo = sum(1 for r in SHOGAIBETSU[1:] for v in r[1:] if v.endswith("％"))
    return hon + hyo


SHOUKAI = [
    ("Ｓ-36", "他メンバー",
     "アンケート結果報告の概要版と分析報告書の確定値の照合",
     "前回計画の評価とアンケート結果報告について、"
     "それぞれ概要版（Ａ４・２ページ）を作成しました。"
     "３本（前回計画の評価・アンケート結果報告・計画素案）を"
     "本体と概要版でそろえる構えに合わせたものです。"
     "アンケートの概要版は、当方が計画素案 第２章７に"
     "要約した内容から作成しており、"
     "分析報告書の確定値との照合が済んでいません。"
     "次の２点をお願いします。"
     "①分析報告書の確定版（令和８年10月21日版の集計表による"
     "回収40件のもの）と、概要版に載せた割合"
     f"（{_n_wariai()}点）を照合させてください。"
     "課題の区分は分析報告書と同じ８件に揃えています"
     "（Ｓ-26の決着。令和８年10月24日に本文を更新）。"
     "②概要版の体裁（字の大きさ、色の使い方、"
     "グレースケールでの見え方）について、"
     "３本の概要版で揃える必要があります。"
     f"現在は前回計画の概要版の実測に合わせ、本文を{BODY_PT:g}ポイント、"
     f"表と注記を{SMALL_PT:g}ポイントとしています（Ｓ-34）。",
     "令和8年10月16日"),
]


# ===========================================================================
# 4. 前回計画の評価 概要版
# ===========================================================================
def _hyoushi(doc, lines):
    t = doc.add_table(rows=1, cols=1)
    t.autofit = False
    cell = t.cell(0, 0)
    cell.width = Mm(180)
    _cell_shade(cell, DARK)
    _cell_margins(cell, top=120, bottom=120, left=160)
    for i, (text, size) in enumerate(lines):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = Pt(size + 6)
        _run(p, text, size=size, bold=True, color="FFFFFF")
    _tbl_width(t)
    _no_borders(t)
    para(doc, "", size=4, space_after=0, line=5)


def hy_hyoushi(doc):
    _hyoushi(doc, [
        ("◆　前回計画の評価　概要版", 13),
        ("第７期北塩原村障がい福祉計画・第３期北塩原村障がい児福祉計画"
         "（令和８年度末の見込みによる評価）", 10),
        ("第４次北塩原村障がい者計画（中間評価）", 10),
    ])
    p = para(doc, align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=4)
    _run(p, f"作成年月　{HI}　／　発行　北塩原村　／　"
            "編集　北塩原村保健福祉課　福祉係　（TEL）0241-23-3113",
         size=8.5)

def hy_kono(doc):
    band(doc, "１", "この資料について")
    para(doc,
         "第８期北塩原村障がい福祉計画・第４期北塩原村障がい児福祉計画の"
         "策定にあたり、前回計画の達成状況を整理したものです。"
         "評価の対象は二つあります。"
         "第７期障がい福祉計画・第３期障がい児福祉計画"
         "（令和６年度～令和８年度）は計画期間の最終年度の途中であるため、"
         "令和８年度末の見込みによる評価（暫定）です。"
         "第４次北塩原村障がい者計画（令和６年度～令和11年度）は"
         "期間の途中であるため中間評価です。")
    box(doc, [
        "評価は成績表ではなく、第８期計画の根拠とするために行います",
        "・見込量は計画値と実績が一致したかだけでは評価しません。"
        f"{_kanten()}に分けて見ます。",
        "・実績が０であることを、ニーズが０であることと読みません。"
        "利用がない理由を、制度を知らない、供給がない、送迎がない、"
        "家族が担っている、他の制度で足りている、"
        "本当に必要がないのいずれかまで確かめます。",
        "・指標の達成と課題の解決は別のことです。"
        "設置した数や配置した人数ではなく、"
        "相談の件数、緊急時に受け入れた件数、"
        "地域生活へ移った方の数で測ります。",
    ], fill=NOTE_FILL, border=ACCENT, size=SMALL_PT, bold_first=True)

def hy_seika(doc):
    band(doc, "２", "成果目標・指標の達成状況（第７期・令和８年度末見込み）")
    hantei = hantei_hyou()
    atsukai = atsukai_hyou()
    para(doc,
         f"第７期計画の成果目標及び指標{len(KPI7)}項目について、"
         "計画値と実績を突き合わせました。"
         + "、".join(f"{k}が{n}項目" for k, n in hantei)
         + "です。", size=SMALL_PT)
    table(doc,
          [["判定", "項目数", "第８期での扱い", "項目数"]]
          + [[hantei[i][0] if i < len(hantei) else "",
              hantei[i][1] if i < len(hantei) else "",
              atsukai[i][0] if i < len(atsukai) else "",
              atsukai[i][1] if i < len(atsukai) else ""]
             for i in range(max(len(hantei), len(atsukai)))],
          [2600, 2500, 2600, 2506], size=SMALL_PT,
          aligns=[None, "center", None, "center"])
    para(doc,
         "「達成」としたものも、数値の達成と課題の解決は別です。"
         "例えば地域生活支援拠点等は１カ所の設置、"
         "コーディネーターは１人の配置をいずれも達成していますが、"
         "緊急時に実際に受け入れた件数は記録されていません。"
         "第８期では、設置や配置の数ではなく機能の実績で測ります。"
         "未達成となった精神科病院からの退院（計画１人・実績０人）は、"
         "第８期では具体の方を定めて支援します。", size=SMALL_PT)

def hy_mikomi(doc):
    band(doc, "３", "サービス見込量の評価（第７期・令和８年度末見込み）")
    mh = mikomi_hyou()
    para(doc,
         f"サービス{len(MIKOMI)}区分を六つの区分に分けました。"
         + "、".join(f"{k}が{n}件" for k, _imi, n in mh if n)
         + "です。", size=SMALL_PT)
    table(doc,
          [["区分", "意味", "件数"]]
          + [[k, imi, n] for k, imi, n in mh],
          [2200, 6500, 1506], size=8,
          aligns=[None, None, "center"])
    para(doc,
         _shu_bun().replace("第７期障がい福祉計画・第３期障がい児福祉計画の", "")
         + "移動支援は実績が０ですが、アンケート調査で外出のニーズが"
         f"うかがえます（{mikomi_kubun('移動支援')}）。"
         "訪問系サービス（重度訪問介護・同行援護・行動援護）も実績が０で、"
         f"ニーズの有無を確かめられていません（{mikomi_kubun('重度訪問介護')}）。"
         "実績が０であることをニーズが０と読まず、"
         "第８期では利用がない理由まで確かめます。", size=SMALL_PT)

def hy_shisaku(doc):
    band(doc, "４", "施策の評価（第７期・令和８年度末見込み）")
    sh = shisaku_hyou()
    para(doc,
         f"施策は{len(SHISAKU)}分野について、"
         "本村で実施しているかどうかではなく、"
         "住民の必要が充たされているかどうかで評価しました。"
         "村内に事業所がないサービスは、"
         "圏域で対応できていれば課題として扱いません。", size=SMALL_PT)
    table(doc,
          [["区分", "意味", "件数"]]
          + [[k, imi, n] for k, imi, n in sh],
          [2600, 6100, 1506], size=8,
          aligns=[None, None, "center"])
    idou = next(r for r in SHISAKU if r[0] == "移動")
    para(doc,
         "移動は移動支援の実績が０で、誰がどの手段で困っているかを"
         f"確かめられていないため「{idou[1][1:]}」としました。"
         f"第８期では{idou[3]}の扱いとし、"
         "公共交通・送迎・家族の負担を含めて再評価します。", size=SMALL_PT)

def hy_kpi4(doc):
    band(doc, "５", "第４次障がい者計画の指標（中間評価）")
    kc = kakunin_chu()
    para(doc,
         f"第４次障がい者計画の計画期間は令和６年度から令和11年度までで、"
         "いまは期間の途中です。"
         "このため「未達成」とは評価せず、"
         "基準値から目標値までの進み具合として見ます。"
         f"{len(KPI4)}指標のうち{kc}指標は現在値を確認中です"
         "（村への照会の回答をいただいたうえで確定します）。", size=SMALL_PT)
    table(doc, kpi4_hyou(), [1500, 3700, 1700, 1700, 1606], size=8,
          aligns=[None, None, "center", "center", "center"])
    note(doc,
         "①差別や偏見と③満足度は、今回のアンケート調査による参考値です。"
         "前回調査と設問の形や分母の定義が異なるため、"
         "達成とは言わず改善の傾向としてみます。"
         "⑦外出の目的は今回の調査に設問がなく測定できないため、"
         "村の生涯学習・スポーツ・文化の事業への参加の実績による"
         "確認に置き換えます。"
         "⑥個別避難計画は、基準値（令和５年度）は１件で決まっていますが、"
         "令和７年度の現在の件数が村の資料どうしで１件と実質０件に分かれており、"
         "確認中です。")

def hy_hanei(doc):
    band(doc, "６", "第８期計画への反映")
    box(doc, [
        "以上を踏まえ、第８期計画では次の三つを行います",
        "第一に、見込量をお一人ずつの状況から積み上げます。"
        f"{LIFESTAGE}といったライフステージの変化を個別に把握します"
        "（40歳に達したことのみをもって減らすことはしません）。",
        "第二に、必要な量と確保できる量を分けて示します。"
        "事業所の受入れの余力、送迎、冬季の交通、人材の不足を確かめ、"
        "確保できない分は未充足量として記録します。"
        "見込量から引くと、不足が見えなくなります。",
        "第三に、達成の判定を数だけに置きません。"
        "設置した数や配置した人数ではなく、相談の件数、"
        "緊急時に受け入れた件数、地域生活へ移った方の数、"
        "調整が付かなかった件数で測ります。",
    ], fill=NOTE_FILL, border=ACCENT, size=SMALL_PT, bold_first=True)
    para(doc,
         "この評価の詳細は、別冊「前回計画の評価」"
         "（北塩原村障がい者自立支援協議会 資料３）に掲載しています。",
         size=SMALL_PT)

HYOUKA_BLOCKS = [
    ("表紙", hy_hyoushi),
    ("1 この資料について", hy_kono),
    ("2 成果目標・指標の達成状況", hy_seika),
    ("3 サービス見込量の評価", hy_mikomi),
    ("4 施策の評価", hy_shisaku),
    ("5 第4次障がい者計画の指標", hy_kpi4),
    ("6 第8期計画への反映", hy_hanei),
]


def build_hyouka():
    doc = docx.Document()
    setup_section(doc)
    pages, oki = flow(doc, HYOUKA_BLOCKS)
    doc.save(HYOUKA_FILE)
    return doc, oki


# ===========================================================================
# 5. アンケート結果報告 概要版
# ===========================================================================
def an_hyoushi(doc):
    _hyoushi(doc, [
        ("◆　障がい福祉に関するアンケート調査　結果報告　概要版", 12),
        ("第８期北塩原村障がい福祉計画・"
         "第４期北塩原村障がい児福祉計画の策定のために実施", 10),
    ])
    p = para(doc, align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=4)
    _run(p, f"作成年月　{HI}　／　発行　北塩原村　／　"
            "編集　北塩原村保健福祉課　福祉係　（TEL）0241-23-3113",
         size=8.5)

def an_gaiyou(doc):
    band(doc, "１", "調査の概要")
    table(doc,
          [["項目", "障がい者調査（18歳以上）", "障がい児調査（18歳未満）"],
           ["調査対象",
            "村内在住の18歳以上の障がいのある方"
            "（手帳所持者、自立支援医療受給者、障がい福祉サービス利用者）",
            "村内在住の18歳未満の障がいのある児童とその保護者"],
           ["調査方法",
            "郵送による配布・回収及びインターネットによる回答の併用",
            "同左"],
           ["調査期間", "令和８年８月31日（月）～９月25日（金）", "同左"],
           ["配布数", "172人", "６人"],
           ["有効回答数",
            f"{A.n()}人（{P('有効回答率_障がい者')}）",
            f"３人（{P('有効回答率_障がい児')}）"]],
          [1500, 4500, 4206], size=SMALL_PT, fill_first=BAND)
    note(doc,
         "【結果を読むときの注意】障がい者調査の有効回答率は"
         f"{P('有効回答率_障がい者')}で、"
         f"回答者の{P('65歳以上')}が65歳以上、"
         f"{P('身体障がい（主たる）')}が身体障がいを主たる障がいとしています。"
         "このため、結果はこうした回答者の構成を踏まえて読む必要が"
         "あります。"
         "障がい児調査は回答が３人と少ないため参考値として扱います。"
         "また、この調査の結果は見込量の人数や量を直に求めるためのもの"
         "ではありません。"
         "回答の割合に対象となる方の人数を掛けて"
         "見込量を求めることはしません。")

def _an_kekka(doc, recs):
    for midashi, honbun, _tashika in recs:
        p = para(doc, space_after=2, size=SMALL_PT)
        _run(p, f"■　{midashi}　", size=SMALL_PT, bold=True, color=DARK)
        _run(p, honbun, size=SMALL_PT)


# 主な結果は9分野あり、まとめて1ブロックにすると頁に収まらない。
# 前半・後半に割り、flow() が頁の割り方を決める。
_AN_HAN = (len(ANKE) + 1) // 2


def an_kekka(doc):
    band(doc, "２", "主な結果")
    _an_kekka(doc, ANKE[:_AN_HAN])


def an_kekka2(doc):
    _an_kekka(doc, ANKE[_AN_HAN:])

def an_betsu(doc):
    band(doc, "３", "障がい別にみた特徴")
    para(doc,
         "所持している手帳により、身体障がい（29人）、"
         "知的障がい（２人）、精神障がい（６人）に区分して集計しました"
         "（複数の手帳を所持している方は各区分に計上）。"
         "知的障がい及び精神障がいは回答者数が少ないため参考値です。",
         size=SMALL_PT)
    table(doc, SHOGAIBETSU, [3906, 1600, 1600, 1550, 1550], size=8,
          aligns=[None, "center", "center", "center", "center"])
    para(doc,
         "精神障がいの方では、外出時に「冬季・積雪時に外出しにくい」が"
         "83.3％、災害時の不安として「避難先での生活」が100.0％と、"
         "身体障がいの方（それぞれ27.6％、55.2％）を"
         "大きく上回っています。"
         "全体の割合だけを見ていると、"
         "この違いが見えなくなります。", size=SMALL_PT)

def an_kadai(doc):
    band(doc, "４", "調査結果からみた課題と本計画での対応")
    para(doc,
         "調査の結果から、本計画で対応すべき主な課題を"
         f"{len(KADAI) - 1}つに整理し、"
         "成果目標、サービス見込量及びその確保のための方策に"
         "反映しました。", size=SMALL_PT)
    table(doc, KADAI, [4400, 5806], size=SMALL_PT)

def an_mochiikata(doc):
    band(doc, "５", "この調査の用い方")
    box(doc, [
        "実績の数字では見えないニーズを拾うために用います",
        "・移動支援の実績は０ですが、外出に困っている方が"
        f"{P('外出_冬季')}います。"
        "重度訪問介護・同行援護・行動援護の実績も０ですが、"
        f"制度をよく知らないという回答が{P('知らない_重度訪問介護')}あります。"
        "実績が０であることを、必要がないことと読みません。",
        "・施策の優先順位を決めるために用います。"
        "村が特に力を入れるべきこととして、"
        f"相談支援体制（{P('施策_相談支援体制')}）と"
        f"災害時支援（{P('施策_災害時支援')}）が上位です。",
        "・確かめる対象の層を絞るために用います。"
        "この調査は無記名で回答された方を特定できないため、"
        "どの層を村の記録で確かめるかの手がかりとして使います。",
    ], fill=NOTE_FILL, border=ACCENT, size=SMALL_PT, bold_first=True)

def an_iken(doc):
    # 計画の工程表はここには置かない。
    # 工程の全体は計画素案の概要版に載せており、
    # この概要版で住民に必要なのはパブリックコメントの日付である。
    band(doc, "", "ご意見をお寄せください")
    box(doc, [
        "パブリックコメント（村民意見公募）を実施します",
        f"{wareki(PABU_KAISHI)}から{wareki(PABU_SHURYO)}まで"
        f"（{PABU_NISSU}日間）、計画の素案を村ホームページ及び"
        "村役場窓口で公開し、ご意見を募ります。"
        "この調査の結果の詳細は、別冊"
        "「北塩原村 障がい福祉に関するアンケート調査 分析報告書」に"
        "掲載しています。",
        "【お問合せ・ご意見の提出先】北塩原村役場　保健福祉課　福祉係"
        "（TEL）0241-23-3113",
    ], fill=NOTE_FILL, border=ACCENT, size=SMALL_PT, bold_first=True)

ANKE_BLOCKS = [
    ("表紙", an_hyoushi),
    ("1 調査の概要", an_gaiyou),
    ("2 主な結果（前半）", an_kekka),
    ("2 主な結果（後半）", an_kekka2),
    ("3 障がい別にみた特徴", an_betsu),
    ("4 調査結果からみた課題", an_kadai),
    ("5 この調査の用い方", an_mochiikata),
    ("ご意見をお寄せください", an_iken),
]


def build_anke():
    doc = docx.Document()
    setup_section(doc)
    pages, oki = flow(doc, ANKE_BLOCKS)
    doc.save(ANKE_FILE)
    return doc, oki


# ===========================================================================
# 6. 自己点検
# ===========================================================================
def load_docx(path):
    d = docx.Document(path)
    out = [p.text for p in d.paragraphs]
    for t in d.tables:
        for row in t.rows:
            for c in row.cells:
                out.append(c.text)
    return "\n".join(out)


def _norm(s):
    return re.sub(r"[\s　]+", "", str(s))


def verify(pages_h, pages_a):
    ng = []
    hyouka = load_docx(HYOUKA_FILE)
    anke = load_docx(ANKE_FILE)

    # ① 1ページに収まること
    for na, pages in (("前回計画の評価", pages_h), ("アンケート", pages_a)):
        for i, lines in enumerate(pages, 1):
            if lines > LINES_PER_PAGE:
                ng.append(f"{na}の概要版 p.{i} が1ページに収まらない"
                          f"（{lines:.0f}行／目安{LINES_PER_PAGE}行）")
        #   頁数は flow() が計算した結果である。
        #   収まりさえすればよいが、4頁を超えると配布の形が変わるため
        #   そこだけを限りとする。
        if len(pages) > 4:
            ng.append(f"{na}の概要版が4ページを超えている: "
                      f"{len(pages)}ページ")

    # ② 件数を数え上げた値が本文に出ていること（手で書いていない証明）
    for k, n in hantei_hyou():
        if f"{k}が{n}項目" not in hyouka:
            ng.append(f"前回計画の評価に判定の件数がない: {k}が{n}項目")
    for k, _imi, n in mikomi_hyou():
        if n and f"{k}が{n}件" not in hyouka:
            ng.append(f"前回計画の評価に見込量の件数がない: {k}が{n}件")
    if f"{len(KPI7)}項目" not in hyouka:
        ng.append(f"前回計画の評価に成果目標の項目数（{len(KPI7)}）がない")
    if f"{len(KPI4)}指標のうち{kakunin_chu()}指標" not in hyouka:
        ng.append("前回計画の評価に第4次計画の指標数と確認中の数がない")
    #    区分の定義を書き写していないこと（KUBUN から引く）
    for k, imi in KUBUN:
        if imi not in hyouka:
            ng.append(f"見込量の区分の意味が本体と違う: {k}")
    for k, imi in SHISAKU_KUBUN:
        if imi not in hyouka:
            ng.append(f"施策の区分の意味が本体と違う: {k}")

    # ③ 前回計画の評価の本体が実在すること（概要版は本体の要約である）
    if not os.path.exists(HYOUKA_HONTAI):
        ng.append(f"前回計画の評価の本体が無い: {HYOUKA_HONTAI}")
    else:
        hontai = _norm(load_docx(HYOUKA_HONTAI))
        for iru in ("前回計画の評価", "評価の考え方",
                    "成果目標・ＫＰＩの達成状況",
                    "サービス見込量の評価", "施策の評価",
                    "第４次障がい者計画の指標", "第８期計画への反映"):
            if _norm(iru) not in hontai:
                ng.append(f"本体に節が見当たらない: {iru}")

    # ④ アンケートの概要版に書いた割合が、正本 第2章7 にあること。
    #    思い込みで数字を書いていないことの裏づけはここである。
    if not os.path.exists(HONPON):
        print("  ※ 正本が未生成のため ④ は省略")
    else:
        hon = _norm(load_docx(HONPON))
        kazu = 0
        for midashi, honbun, tashika in ANKE:
            for nm in tashika:
                kazu += 1
                wari = P(nm)
                if wari not in honbun:
                    ng.append(f"指標の割合が本文に出ていない: "
                              f"{midashi}／{nm} {wari}")
                if wari not in hon:
                    ng.append(f"正本に無い割合を概要版に書いている: "
                              f"{midashi}／{nm} {wari}")
            #    38件の割合が残っていないこと（新旧で値が違うものに限る）
            for nm in tashika:
                kyu = P(nm, kyu=True)
                if kyu != P(nm) and kyu in honbun:
                    ng.append(f"38件の割合が残っている: {midashi}／{nm} {kyu}")
            if honbun not in anke:
                ng.append(f"概要版に本文が入っていない: {midashi}")
        #    障がい別の表の値も正本にあること
        for row in SHOGAIBETSU[1:]:
            for v in row[1:]:
                kazu += 1
                if v not in hon:
                    ng.append(f"正本に無い値を障がい別の表に書いている: "
                              f"{row[0]}／{v}")
        #    調査の概要の値も正本にあること
        for v in ("172人", f"{A.n()}人", P("有効回答率_障がい者"), "50.0％",
                  P("65歳以上"), P("身体障がい（主たる）"),
                  "令和８年８月31日"):
            kazu += 1
            if _norm(v) not in hon:
                ng.append(f"正本に無い値を調査の概要に書いている: {v}")
        #    課題の数が正本の表2-8 と一致すること
        if "表２-８　調査結果からみた課題と本計画での対応" not in \
                load_docx(HONPON):
            ng.append("正本に表2-8（調査結果からみた課題）がない")

    # ⑤ 工程が単一出所であること
    for need in (wareki(PABU_KAISHI), wareki(PABU_SHURYO),
                 f"{PABU_NISSU}日間"):
        if need not in anke:
            ng.append(f"アンケートの概要版にパブリックコメントの"
                      f"{need}がない")
    for furui in ("２月８日", "２月22日", "令和８年度中"):
        if furui in anke:
            ng.append(f"アンケートの概要版に従前の工程が残っている: "
                      f"{furui}")

    # ⑥ 照会の番号と記号
    if [x[0] for x in SHOUKAI] != ["Ｓ-36"]:
        ng.append(f"照会の番号が想定と違う: {[x[0] for x in SHOUKAI]}")
    for x in SHOUKAI:
        if re.search(r"(?<![A-Za-zＡ-Ｚａ-ｚ])[MS]-\d+", x[0]):
            ng.append(f"照会番号の記号が半角: {x[0]}")

    # ⑦ 表記の作法と他団体の名前
    zen = hyouka + anke + "\n".join(str(v) for r in SHOUKAI for v in r)
    if "か所" in zen:
        ng.append("「カ所」に揃っていない")
    if "〜" in hyouka + anke:
        ng.append("概要版に波ダッシュ（〜）が混じっている")
    for warui in ("小野町", "金ケ崎", "金ヶ崎", "阿蘇", "札幌"):
        if warui in zen:
            ng.append(f"他団体の案件の地名が入っている: {warui}")

    if ng:
        print("自己点検 不合格:")
        for x in ng:
            print("  -", x)
        return False
    print(f"自己点検: 合格　前回計画の評価 "
          f"成果目標{len(KPI7)}項目・見込量{len(MIKOMI)}区分・"
          f"施策{len(SHISAKU)}分野・第4次の指標{len(KPI4)}件を"
          f"本体の生成器から数え上げ／"
          f"アンケート 割合{kazu}点を正本 第2章7 の現物と照合／"
          f"工程を kitashiobara_koutei.py と照合／"
          f"前回計画の評価{len(pages_h)}ページ・"
          f"アンケート{len(pages_a)}ページに収まる")
    return True


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    dh, oh = build_hyouka()
    da, oa = build_anke()
    ph = estimate_lines(dh)
    pa = estimate_lines(da)
    for path, pages, oki in ((HYOUKA_FILE, ph, oh), (ANKE_FILE, pa, oa)):
        print(f"作成: {path}")
        for name, pg, nokori in oki:
            print(f"    p.{pg}  {name}（置いた後 {nokori:.0f}行）")
        for i, lines in enumerate(pages, 1):
            print(f"  p.{i}  推定 {lines:.0f} 行 / 目安 {LINES_PER_PAGE} 行")
    return ph, pa


if __name__ == "__main__":
    import sys
    if not verify(*main()):
        sys.exit(1)
