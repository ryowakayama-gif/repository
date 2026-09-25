# -*- coding: utf-8 -*-
"""大雪地区広域連合 第10期介護保険事業計画
アンケート調査の分析　点検箇所と手順（社内チェック用）.

令和8年9月24日のご依頼
  「9月業務上のアンケート分析について先方に提出する資料と
    社内チェック用の点検箇所を整理した資料をワードでアウトプットして下さい」

本書は**社内チェック用**である。発注者提出用は
`build_survey_monthly_r8_09.py`（別冊）とする。

この業務に携わっていない者でも点検できるよう、
**何を・どこで・どうやって確かめるのか**を書く。

━━ 数値はすべて実物・ソースから読む ━━

成果品のシート数は実物の xlsx を開いて数える。
自己点検の件数はソースの `chk(` の呼出しを `ast` で数える。
留保・所見の件数は集計分析報告書の実物から読む。
突合の測定値は `build_survey_jisseki_cross.py` を `runpy` で読む。

構成
  第1節　この資料の使い方と点検の3層
  第2節　点検の対象（成果品とスクリプト）
  第3節　第1層　機械で確かめる
  第4節　第2層　数値の出所を辿る
  第5節　第3層　記述の当否と守るべき制約
  第6節　特に見ていただきたい箇所
  第7節　この作業で実際に起きた誤り
  第8節　個人情報・禁止表現の探し方
  第9節　再実行できないもの
  第10節　点検の記録様式と分担

出力
  output/第10期計画_アンケート分析_点検箇所と手順.docx

自己点検で1件でも不適合があると終了コード1で終わる。
"""

import ast
import os
import re
import runpy
import sys

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from openpyxl import load_workbook

import repo_paths as RP

OUT = (RP.ROOT + "/output/"
       "第10期計画_アンケート分析_点検箇所と手順.docx")
KIJUNBI = "令和8年9月24日"
FONT = "游ゴシック"
NAVY = RGBColor(0x1F, 0x38, 0x64)

CHECKS = []


def chk(no, naiyo, kekka, ok):
    CHECKS.append((no, naiyo, kekka, "適合" if ok else "不適合"))
    return ok


# ==================================================== 実物・ソースから読む
def _load(name):
    """スクリプトを runpy で読む（固定値を書き写さない）。"""

    class Sink(object):
        closed = False
        encoding = "utf-8"
        errors = "strict"
        newlines = None
        line_buffering = False
        name = "<sink>"
        mode = "w"

        def __init__(self):
            self.buffer = self

        def write(self, *_a, **_k):
            return 0

        def writelines(self, _l):
            pass

        def flush(self):
            pass

        def close(self):
            pass

        def fileno(self):
            raise OSError("sink")

        def isatty(self):
            return False

        def readable(self):
            return False

        def writable(self):
            return True

        def seekable(self):
            return False

        def detach(self):
            return self

        def reconfigure(self, *_a, **_k):
            pass

    old, sys.stdout = sys.stdout, Sink()
    try:
        return runpy.run_path(os.path.join(RP.ROOT, name), run_name="_loaded")
    finally:
        sys.stdout = old


def _chk_count(name):
    """スクリプトの自己点検の件数をソースから数える。"""
    path = os.path.join(RP.ROOT, name)
    if not os.path.exists(path):
        return 0
    src = open(path, encoding="utf-8").read()
    n = 0
    for node in ast.walk(ast.parse(src)):
        if isinstance(node, ast.Call) and getattr(node.func, "id", "") == "chk":
            n += 1
    return n


def _sheets(fn):
    """成果品（xlsx）のシート数を実物から数える。"""
    path = os.path.join(RP.OUTPUT, fn)
    if not os.path.exists(path):
        return 0
    return len(load_workbook(path, read_only=True).sheetnames)


def _paras(fn):
    """成果品（docx）の段落数・表数を実物から数える。"""
    path = os.path.join(RP.OUTPUT, fn)
    if not os.path.exists(path):
        return (0, 0)
    d = Document(path)
    return (len(d.paragraphs), len(d.tables))


SRC_XLSX = "第10期計画_アンケート調査の集計分析報告書.xlsx"
WB = load_workbook(os.path.join(RP.OUTPUT, SRC_XLSX), data_only=True)


def _rows(sheet, head_row=4):
    ws = WB[sheet]
    out = []
    for row in ws.iter_rows(min_row=head_row + 1, values_only=True):
        if row[0] is None or (isinstance(row[0], str)
                              and row[0].startswith("注")):
            break
        out.append(row)
    return out


RYUHO = _rows("11_調査結果の限界と留保")
SHOKEN = _rows("12_主要所見と計画本文への反映")
SHOKEN_ST = {}
for r in SHOKEN:
    SHOKEN_ST[r[5]] = SHOKEN_ST.get(r[5], 0) + 1

_X = _load("build_survey_jisseki_cross.py")
CU_CHI, CHI_1, CU_N = _X["CU_CHI"], _X["CHI_1"], _X["CU_N"]

DRAFT = RP.draft_label()

# ==================================================== 点検の対象
# （成果品名, スクリプト, 区分）シート数・自己点検の件数は実物・ソースから数える。
SEIKA = [
    (SRC_XLSX, "build_survey_report.py", "送付"),
    ("第10期計画_実施済み調査_結果報告書.docx",
     "build_survey_report_doc.py", "送付"),
    ("第10期計画_調査クロス集計・分析.xlsx",
     "build_survey_crosstab.py", "送付"),
    ("第10期計画_実施済み3調査の受領点検と集計.xlsx",
     "build_survey_review.py", "送付"),
    ("第10期計画_調査結果と年報実績の突合クロス集計.xlsx",
     "build_survey_jisseki_cross.py", "条件付き"),
    ("第10期計画_9月作業_アンケート分析の作業状況.xlsx",
     "build_survey_status.py", "送付"),
    ("第10期計画_事業所調査の照会票と確定値管理表.xlsx",
     "build_survey_inquiry.py", "送付"),
    ("第10期計画_アンケート調査の分析_令和8年9月業務報告.docx",
     "build_survey_monthly_r8_09.py", "送付"),
]

# 数値の出所の対照（値, どこに書いてあるか, どこから来ているか）
DEDOKORO = [
    ("3調査（①②③）の集計値", "data_survey2025.py",
     "受領した様式（事業所票・施設等票・職員票）。個票は収録しない"),
    ("3調査のクロス集計（度数）", "data_survey_cross.py",
     "同上。度数のみを収める"),
    ("照会に要する記載内容", "data_survey_entry.py",
     "提出元・ファイル名・票番号・所在地区欄の記載文字列のみ"),
    ("④健康とくらしの調査の集計値", "data_survey_jages.py",
     "個票4,729票の集計結果（`build_survey_crosstab.py` の出力から写した値）"),
    ("年報（令和7年度）の要介護度別明細", "data_nenpo_meisai.py",
     "様式1の6・1の7・様式2"),
    ("見込量・給付費・保険料", "build_mikomiryo_santei.py",
     "第1次概算（runpy で読む。報告書に書き写さない）"),
    ("調査と実績の突合", "build_survey_jisseki_cross.py",
     "年報の要介護度別明細と調査の分布（runpy で読む）"),
    ("計画素案の規模・反映の判定", "repo_paths.py／素案の実物",
     "`RP.draft_label()` と `_draft_sections()` が実物の docx を数える"),
]

# 特に見てほしい箇所
MIRU = [
    ("母集団の書き方",
     "①在宅生活改善調査の結果を「在宅の認定者の割合」として書いていないか",
     "χ²＝%.1f（自由度6・1％点%.2f）で在宅の認定者を代表しない。"
     "「調査対象となった%d人のうち」と母集団を明記する（確認事項No.140）"
     % (CU_CHI, CHI_1, CU_N)),
    ("在籍と受給の書き分け",
     "②居所変更実態調査の在籍者を年報の受給者と同じものとして扱っていないか",
     "年報は施設の所在地を問わず当連合の被保険者を数え、"
     "調査は保険者を問わず区域内の施設の在籍者を数える。数える対象が違う"
     "（確認事項No.141）"),
    ("率と実数の併記",
     "割合だけを書いて実数を落としていないか",
     "母数が小さい項目は率だけでは読み違える。実数を併記する"),
    ("見込量のじか書き",
     "報告書の見込量が計画素案 第6章と食い違っていないか",
     "`build_mikomiryo_santei.py` を runpy で読む形になっていること。"
     "固定値で書くと素案を改めたときにずれる"),
    ("反映済みの判定",
     "計画本文への反映を語の有無だけで判定していないか",
     "`in_draft(章節, 語...)` で章節を限って判定する。"
     "「除雪」は第1章第7節に、「同規模保険者」は第3章第4節にも現れる"),
    ("留保の落とし",
     "限界と留保が報告書から落ちていないか",
     "11シートの留保は%d件。解消したものは14シートで解消の根拠まで示す"
     % len(RYUHO)),
    ("個票の非収録",
     "個票（個人情報）が収録されていないか",
     "④の個票CSV（4,729票）は収録しない。集計値は "
     "`data_survey_jages.py` に収める"),
    ("単位と月数",
     "月平均の値をもう一度12で除していないか",
     "`SHISETSU_M`（年報）は既に月平均である。"
     "在籍と受給の表が12分の1になっていた（第7節の1件目）。"
     "**桁が1つ違う値は単位を疑う**"),
    ("仕様書の作業項目",
     "仕様書４（3）の作業項目7件が漏れなく挙がっているか",
     "作業項目は業務工程管理表の `WORK` から `ast` で読む。"
     "「課題整理」と「第10期計画への反映」まで求められている"),
    ("経年比較の出所",
     "前回調査との比較を固定値で書いていないか",
     "前回の規模は図表集 00シートから、④の指標の比較は"
     "計画素案 第2章第4節の表から読む。"
     "①②③の内容の比較は、第9期計画の本文は確認しているが"
     "前回調査の指標別の分母と集計条件が復元できないためできない"
     "（No.90）。「本文が未受領」と書かない"),
    ("発注者提供文の書き方",
     "当方の内部の仕組み・作業経過が発注者提出用の文章に出ていないか",
     "スクリプト名・モジュール名、「固定値では書いていません」"
     "「実物から読んでいます」、個票の格納・再現の可否、"
     "受託者の内部の課題整理、作業日の細かい日付は書かない。"
     "**「収める」は当方の作業の語であり、発注者に対しては「確認する」**"
     "（CLAUDE.md §1。令和8年9月24日 ご指示）"),
    ("原典との突合",
     "報告書の数値が調査票の集計値（原典）から再現できるか",
     "9月業務報告の自己点検No.22〜29が、"
     "`data_survey2025`・`data_survey_cross` を直接読んで突き合わせる。"
     "報告書の集計を経由しないため、途中で写し違えていれば不適合になる"),
    ("割合の分母（n）",
     "割合に分母を併せて示しているか。分母を取り違えていないか",
     "④は設問ごとに分母が異なる（4,128票から4,729票）。"
     "**「除雪24.3％」の分母は回答者4,128票であり、"
     "困りごとがある者1,426〜1,679ではない**（該当数1,002人）。"
     "表題に「〜のうち」と書くときは、その分母が実際に使われているかを確かめる"),
    ("同じものに2つの数値がないか",
     "②の入所者数のように、同じ項目に複数の数え方がないか",
     "入所欄の合計536人（記入のない1施設あり）と、"
     "その1施設を要介護度別の内訳58人で補った594人（入所率91.1％）、"
     "要介護度別内訳の合計599人の3通りがある。"
     "**集計分析報告書 03シートの【入所者数の数え方の照合】で"
     "3通りと採用値の用いどころを対照している**。"
     "老健は入所欄201人・内訳206人で5人違う（点検事項No.32）"),
    ("推計と実績の書き分け",
     "推計値を実績のように書いていないか",
     "認定者の利用状況（在宅1,016人ほか）は"
     "認定者数に利用率（見える化D45系列）を乗じた**推計**であり、"
     "年報の受給者実人数ではない。"
     "定員への到達も算定による**見込み**である。"
     "「求めました」「達します」ではなく"
     "「推計しました」「達する見込みです」と書く"),
    ("①の分母は設問ごとの有効回答か",
     "利用者票99票の割合を回収票数で割っていないか",
     "**H06（在宅生活継続困難割合）の分母は"
     "「在宅生活の維持の見通し」の有効回答98票**"
     "（72人÷98票＝73.5％）。回収票数99票で割った72.7％は使わない。"
     "**H12の分母は「より適切と思われるサービス」の有効回答99票**"
     "（20票÷99票＝20.2％）。"
     "要介護度の有効回答98票を別の設問の分母に流用しない"
     "（確認事項No.158）"),
    ("代表KPIの定義が1か所から引かれているか",
     "同じ代表KPIの名称・定義・算式・基準値が表ごとに違っていないか",
     "第5章第3節の一覧・各基本目標の達成目標・資料1の3か所に"
     "同じKPIが現れる。"
     "**H07・H08・H12・H16は代理指標へ振り替えたため、"
     "当初の定義と「算定不可」が残っていないことを確かめる**"),
    ("目標年度と未達幅の対応",
     "「○ポイント未達」がその年度の目標との差になっているか",
     "通いの場参加率は令和7年度の中間目標9.0％・"
     "令和8年度の期末目標10.0％である。"
     "実績8.8％との差は0.2ポイント（R7）・1.2ポイント（R8）。"
     "**実績の年度と同じ年度の目標と比べる**"),
    ("反映先の節が存在するか",
     "反映先として挙げた章節が計画素案に実際にあるか",
     "所見6・10・11・12の反映先を「第3章第5節」としていたが、"
     "第3章は第4節までであり当該節は存在しない。"
     "介護人材の記述は第2章第6節にある。"
     "**反映先は素案の章節の一覧と突き合わせる**"),
]

# この作業で実際に起きた誤り（気づいた経緯と手立て）
AYAMARI = [
    ("報告書の見込量が計画素案と食い違っていた",
     "10シートの見込量が将来推計 第2段階のままで、"
     "計画素案 第6章（第1次概算）と別の答えが並んでいた",
     "第6章を差し替えたときに報告書を見直さなかった",
     "`build_mikomiryo_santei.py` を runpy で読む形に改めた。"
     "算定を改めれば報告書が追随する"),
    ("12シートが反映状況を持っていなかった",
     "第2章第4節（9/15反映）・第6章（9/16反映）が着手前のように見えていた",
     "状態が「着手可」のままで、素案の実物を見ていなかった",
     "状態に「反映済」を加え、素案を章節ごとに読んで判定する形にした"),
    ("語の有無だけの判定で偽陽性が出た",
     "所見7の反映先を第2章第4節としていたが実際は第2章第5節であった",
     "素案全体で語を検索していた",
     "`_draft_sections()`／`in_draft(章節, 語...)` で章節を限る"),
    ("過去の時点の報告が現時点の値を読んでいた",
     "業務進捗報告書 令和8年8月分の素案の規模が再出力のたびに"
     "667→806→843→896段落と書き換わっていた",
     "`RP.draft_label()`（現時点を数える）を過去の資料が読んでいた",
     "`RP.draft_label_at(基準日)` と `data_progress.snapshot(基準日)` を新設した"),
    ("禁止表現が残っていた",
     "確認事項No.84の「北海道の指定事業所一覧に掲載がなく休止と整合する」",
     "根拠として「整合する」を使っていた",
     "「北海道の指定事業所一覧にも掲載がない」に改めた"),
    ("docx が Markdown を解釈せずアスタリスクが本文に出た",
     "強調のつもりで書いたアスタリスク2つがそのまま印字されていた",
     "python-docx は Markdown を解釈しない",
     "`_runs()` でアスタリスク2つを分割して太字の run にする。"
     "段落・箇条書き・表のセルの全てを通す"),
    ("割合の分母を取り違えていた",
     "計画素案 第2章第4節の表が「生活動作の困りごとがある者のうち"
     "除雪を挙げた者 24.3％ 4,128票」となっていた",
     "24.3％の分母は回答者4,128票（該当数1,002人）であり、"
     "「困りごとがある者」（1,426〜1,679）ではない。"
     "表題と分母が食い違っていた",
     "表題を「生活動作の困りごととして除雪を挙げた者」に改め、"
     "回答者に占める割合であること・該当数1,002人を注記した。"
     "**表題に「〜のうち」と書くときは、その分母が実際に使われているかを"
     "原典の該当数から逆算して確かめる**"),
    ("推計を実績のように書いていた",
     "9月業務報告が「認定者1,984人がどのサービスを使っているかを"
     "年報から求めました」と書いていた",
     "実際は認定者数に利用率（見える化D45系列）を乗じた推計であり、"
     "年報の受給者実人数ではない。集計分析報告書 10シートの注3は"
     "そのことを書いていたが、業務報告に写す際に落ちた",
     "「推計しました」に改め、月平均の受給者実績とは算定の方法が異なり"
     "一致しないことを注記した。**注記も併せて写す**"),
    ("同じ項目に2つの数値があった",
     "②回答18施設の入所者数が、00シートでは536人、"
     "所見13では594人（入所率91.1％）となっていた",
     "入所欄の記入がない1施設があり、"
     "要介護度別の内訳58人で補うかどうかで2通りになる。"
     "どちらを使うかを書いていなかった",
     "両方を併記し、どちらを用いるかの決定を要することを明記した"
     "（点検事項No.32）。業務報告の前回比較の表からは入所者数を外した"),
    ("存在しない章節を反映先としていた",
     "所見6・10・11・12の反映先を「第3章第5節」としていたが、"
     "計画素案の第3章は第4節までであり当該節は存在しない",
     "反映先を素案の章節の一覧と突き合わせていなかった。"
     "介護人材の記述は第2章第6節にある",
     "4件の反映先を第2章第6節に改めた。"
     "**所見6は既に第2章第6節に反映済みであったにもかかわらず、"
     "存在しない節を見に行っていたため「着手可」のまま残っていた**"),
    ("発注者提出用に当方の内部の仕組みが出ていた",
     "9月業務報告に「集計値は data_survey_jages.py に収めています」"
     "「固定値では書いていません」「個票がなくても再現できます」"
     "「計画素案の記述が薄かった3節のうち1節が埋まりました」が出ていた",
     "社内向けの書き方をそのまま発注者提出用に持ち込んだ",
     "いずれも落とした。"
     "**「収める」は当方の作業の語であり、発注者に対しては「確認する」**。"
     "作業日の細かい日付も本文からは落とした（CLAUDE.md §1）"),
    ("在籍と受給の表が12分の1になっていた",
     "9月業務報告の「施設の在籍者と保険者の受給者」の表が、"
     "受給を月平均の12分の1（老健134.1→11.2人／月）で出しており、"
     "在籍の列そのものが落ちていた",
     "`SHISETSU_M` が既に月平均であるのに、もう一度12で除していた。"
     "在籍を組み立てる途中の行が使われないまま残っていた",
     "集計分析報告書 10シートの実物（在籍・受給・差・見方）を"
     "読む形に改めた。"
     "**表の値が算定と一致することを自己点検で確かめる**"
     "（一致しなければ終了コード1）"),
    ("14シートの行が1件も読めなかった",
     "本書と同時に作った9月業務報告で、"
     "補完と解消の件数が0件になり自己点検が不適合になった",
     "No.の列が文字列であるのに `isinstance(row[0], int)` で拾っていた",
     "`str(row[0]).isdigit()` に改めた。**自己点検が働いて気づいた**"),
    ("割合の分母を回収票数のままにしていた",
     "在宅生活の維持が困難な者の割合を72人÷99票＝72.7％と書いていた"
     "箇所と、72人÷有効回答98票＝73.5％と書いていた箇所が"
     "同じ報告書の中に併存していた",
     "02シートの本表は有効回答98票を分母としていたのに、"
     "00シートの要約・注記・11シート・12シートへ写すときに"
     "回収票数99票で割った値を書いていた",
     "**設問ごとの有効回答から割合を計算する形に改めた**"
     "（無回答1票は分母から除く）。"
     "回収票数を分母とする値が本文に残っていないことを"
     "自己点検で確かめる"),
    ("要介護度の有効回答を別の設問の分母に流用していた",
     "代表KPI H12（区域内で提供できないサービスが必要と判断された"
     "利用者の割合）を20票÷98票＝20.4％としていた",
     "98票は要介護度の有効回答であり、"
     "H12が用いる問3の有効回答は99票である。"
     "要介護度はH12の分母の条件ではない",
     "20票÷99票＝20.2％に改めた。"
     "**分母は「その設問に回答した票」であることを確かめる**"),
    ("代表KPIの定義が計画素案の中で新旧併存していた",
     "第5章第3節では代理指標へ振り替えた定義（算定可）を掲げる一方、"
     "各基本目標の達成目標と資料1には当初の定義と"
     "「算定不可」が残っていた（H06・H07・H08・H12・H16）",
     "KPIの名称・定義・算式・基準値を表ごとに書いていた",
     "**代表KPIの定義を1つの辞書にまとめ、"
     "第5章第3節・各基本目標の達成目標・資料1のすべてが"
     "そこから引く形に改めた**"),
    ("目標年度と未達幅が対応していなかった",
     "通いの場参加率について、第9期目標を"
     "「10.0％（令和8年度）」と掲げながら"
     "「0.2ポイント未達」と書いていた（差は1.2ポイント）",
     "0.2ポイントは令和7年度の中間目標9.0％との差である。"
     "同じ素案の別の表（第3章第1節）は正しく令和7年度目標と"
     "比べていた",
     "実績の年度と同じ年度の目標と比べる形に統一した。"
     "**「未達○ポイント」は目標値と実績の差で検算する**"),
    ("同じ「入所者」の語で3通りの数が現れていた",
     "18施設の入所者数が、入所欄の合計536人・"
     "要介護度別内訳の合計599人・"
     "記入のない1施設を内訳58人で補った594人の3通りあり、"
     "老健は入所欄201人と内訳206人で食い違っていた",
     "欠測の埋め方と数える範囲が集計ごとに違うのに、"
     "同じ名称で並べていた",
     "集計分析報告書 03シートに"
     "**【入所者数の数え方の照合】**を新設し、"
     "3通りの数と採用値の用いどころを1表で対照した"),
]

# 守るべき制約（この分野に効くもの）
SEIYAKU = [
    ("個票データはリポジトリに格納しない。集計値のみ収録する",
     "④の個票CSV 4,729票・①の利用者票99票・③の職員個票317人"),
    ("担当者名・電話番号・メールアドレスは収録しない",
     "事業所票の記入者欄・照会票の宛先。〔　〕の記入枠にする"),
    ("二次情報を出典に使わない。原典のみ",
     "調査の結果は受領した様式、実績は年報・見える化による"),
    ("禁止表現5件",
     "「〜に由来する」「〜と整合する（根拠として）」「1件も〜ない」"
     "「有意差がないため関係がない」「全国トップ級」"),
    ("計画素案の本文に確認事項の注記を書かない",
     "確認事項は業務工程管理表と別管理表で一元管理する。"
     "本文の注記は出典・単位・欠測・数値の性質・記号の意味に限る"),
    ("［要協議］［要確認］［要内訳］は本文に残す",
     "確定しない値を推測しない"),
]

# 記録様式
KIROKU_SHIKI = [
    ("重大", "数値が誤っている／母集団を取り違えている／個人情報が入っている",
     "直ちに差止め。作成者へ連絡し、直してから再出力する"),
    ("中", "出所が辿れない／2つの成果品で値が食い違う／留保が落ちている",
     "当日中に作成者へ連絡する。確認事項として起票するかを判断する"),
    ("軽", "表記のゆれ／単位の書き落とし／体裁",
     "まとめて連絡する"),
]

# 分担の案
BUNTAN = [
    ("第1層", "どなたでも", "30分",
     "スクリプトを再実行して終了コード0であること・"
     "docx が validate.py に適合すること"),
    ("第2層", "算定を見た方", "2〜3時間",
     "数値の出所を辿る。じか書き・固定値・出所の食い違い"),
    ("第3層", "業務を見ている方", "半日",
     "記述の当否と守るべき制約。母集団・率と実数・留保の落とし"),
]


# ==================================================== docx の体裁
doc = Document()
st = doc.styles["Normal"]
st.font.name = FONT
st.font.size = Pt(10.5)
st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
st.paragraph_format.space_after = Pt(4)
st.paragraph_format.line_spacing = 1.15
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
sec.top_margin = sec.bottom_margin = Cm(2.0)
sec.left_margin = sec.right_margin = Cm(1.9)
TEXTW = 21.0 - 1.9 * 2


def _shade(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color)
    # shd は vAlign 等より前に置く（CLAUDE.md §4）。
    after = None
    for tag in ("w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText",
                "w:vAlign", "w:hideMark"):
        after = tcPr.find(qn(tag))
        if after is not None:
            break
    if after is None:
        tcPr.append(shd)
    else:
        after.addprevious(shd)


def _cellmargin(t):
    mar = OxmlElement("w:tblCellMar")
    for tag, v in (("top", 0.06), ("left", 0.12), ("bottom", 0.06),
                   ("right", 0.12)):
        e = OxmlElement("w:" + tag)
        e.set(qn("w:w"), str(int(v * 567)))
        e.set(qn("w:type"), "dxa")
        mar.append(e)
    # tblCellMar は tblLook より前に置く（CLAUDE.md §4）。
    tblPr = t._tbl.tblPr
    after = None
    for tag in ("w:tblLook", "w:tblCaption", "w:tblDescription"):
        after = tblPr.find(qn(tag))
        if after is not None:
            break
    if after is None:
        tblPr.append(mar)
    else:
        after.addprevious(mar)


def _runs(p, text, size, bold=False, color=None):
    """** で囲んだ部分を太字にする。docx は Markdown を解釈しない。"""
    for i, part in enumerate(str(text).split("**")):
        if not part:
            continue
        r = p.add_run(part)
        r.font.name = FONT
        r.font.size = Pt(size)
        r.font.bold = bold or (i % 2 == 1)
        if color is not None:
            r.font.color.rgb = color
        r.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    if not p.runs:
        r = p.add_run("")
        r.font.name = FONT
        r.font.size = Pt(size)
        r.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    return p


def P(text="", size=10.5, bold=False, align=None, indent=0.0, after=4):
    p = doc.add_paragraph()
    _runs(p, text, size, bold)
    if align:
        p.alignment = align
    if indent:
        p.paragraph_format.left_indent = Cm(indent)
    p.paragraph_format.space_after = Pt(after)
    return p


def H1(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.font.name = FONT
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = NAVY
    r.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)


def H2(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(9)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.font.name = FONT
    r.font.size = Pt(11)
    r.font.bold = True
    r.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)


def BUL(text, size=10.5, mark="・"):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.6)
    p.paragraph_format.first_line_indent = Cm(-0.6)
    p.paragraph_format.space_after = Pt(3)
    _runs(p, mark + text, size)
    return p


def NOTE(text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.4)
    _runs(p, "※ " + text, 9)
    return p


def CODE(text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.8)
    p.paragraph_format.space_after = Pt(1)
    r = p.add_run(text)
    r.font.name = "Consolas"
    r.font.size = Pt(9.5)
    r.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    return p


TBLNO = [0]


def CAP(text):
    TBLNO[0] += 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    r = p.add_run("表%d　%s" % (TBLNO[0], text))
    r.font.name = FONT
    r.font.size = Pt(10)
    r.font.bold = True
    r.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)


def TBL(head, rows, widths, size=9, center=None, first_bold=False):
    center = center or set()
    t = doc.add_table(rows=1 + len(rows), cols=len(head))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    s = sum(widths)
    if s > TEXTW:
        widths = [w * TEXTW / s for w in widths]
    _cellmargin(t)
    tr = t.rows[0]
    tr._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
    for j, v in enumerate(head):
        c = tr.cells[j]
        c.width = Cm(widths[j])
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        _shade(c, "1F3864")
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(v)
        r.font.name = FONT
        r.font.size = Pt(size)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        r.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    for i, row in enumerate(rows):
        tr = t.rows[i + 1]
        tr._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
        for j, v in enumerate(row):
            c = tr.cells[j]
            c.width = Cm(widths[j])
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if i % 2 == 1:
                _shade(c, "F2F5FA")
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if j in center:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _runs(p, "" if v is None else v, size,
                  bold=(first_bold and j == 0))
    return t


# ============================================================ 表紙
P("大雪地区広域連合　第10期介護保険事業計画策定支援業務", size=12,
  align=WD_ALIGN_PARAGRAPH.CENTER)
p = P("アンケート調査の分析　点検箇所と手順", size=17, bold=True,
      align=WD_ALIGN_PARAGRAPH.CENTER, after=2)
for r in p.runs:
    r.font.color.rgb = NAVY
P("社内チェック用　　基準日　%s" % KIJUNBI, size=10,
  align=WD_ALIGN_PARAGRAPH.CENTER)
P("")

P("本書は**社内の点検のための資料**です。発注者への提出は想定していません。"
  "発注者へ提出する報告は別冊"
  "「アンケート調査の分析　令和8年9月　業務報告」です。")
P("**この業務に携わっていない方でも点検できるように**、"
  "何を・どこで・どうやって確かめるのかを書いています。"
  "件数・シート数・測定値はいずれも実物又はソースから数えており、"
  "本書に固定値は書いていません。")

# ============================================================ 第1節
H1("第1節　この資料の使い方と点検の3層")

P("点検を3層に分けています。**上の層ほど短時間で・機械的にでき、"
  "下の層ほど業務の理解を要します。**"
  "時間が取れない場合は第1層だけでも行ってください。"
  "第1層は誤りの型のうち「出力が壊れている」「自己点検に反している」を"
  "すべて拾います。")

CAP("点検の3層")
TBL(["層", "何を確かめるか", "だれが", "目安", "できること"],
    [(g[0], g[3], g[1], g[2],
      {"第1層": "出力が壊れていないこと・自己点検に適合すること",
       "第2層": "数値がどこから来ているかを辿れること",
       "第3層": "書いてある内容が業務として正しいこと"}[g[0]])
     for g in BUNTAN],
    [1.4, 5.6, 2.2, 1.4, 6.6], center={0, 2, 3})

P("")
P("**点検の順序は第1層→第2層→第3層です。**"
  "第1層で不適合が出た場合、第2層以降は直してから行ってください。"
  "出力が壊れている状態で内容を読んでも意味がありません。")

# ============================================================ 第2節
H1("第2節　点検の対象（成果品とスクリプト）")

P("アンケート調査の分析に関わる成果品は次のとおりです。"
  "**規模（シート数・段落数）は実物を開いて数えた値です。**")

_rows_seika = []
for fn, sc, kb in SEIKA:
    path = os.path.join(RP.OUTPUT, fn)
    if fn.endswith(".xlsx"):
        kibo = "%dシート" % _sheets(fn)
    else:
        pa, ta = _paras(fn)
        kibo = "%d段落・%d表" % (pa, ta)
    n = _chk_count(sc)
    _rows_seika.append((fn, sc, kibo, ("%d件" % n) if n else "―", kb,
                        "○" if os.path.exists(path) else "×"))

CAP("点検の対象（成果品・スクリプト・自己点検の件数）")
TBL(["成果品", "スクリプト", "規模", "自己点検", "区分", "実体"],
    _rows_seika, [5.6, 5.0, 2.2, 1.4, 1.5, 1.0],
    center={2, 3, 4, 5})
NOTE("「自己点検」は、スクリプトの中に `chk(` として書かれた点検の件数です"
     "（ソースを `ast` で数えています）。"
     "「―」は自己点検を備えていないもので、この場合は第2層・第3層で見ます。")
NOTE("「実体」は成果品のファイルが output/ にあるかどうかです。")

P("")
P("集計値を収める側（データモジュール）は次の4件です。"
  "**個票（個人情報）はいずれにも収めていません。**")
BUL("`data_survey2025.py`　　①②③の集計値")
BUL("`data_survey_cross.py`　①②のクロス集計（度数）")
BUL("`data_survey_entry.py`　照会に要する記載内容（所在地区欄の文字列まで）")
BUL("`data_survey_jages.py`　④健康とくらしの調査の集計値")

# ============================================================ 第3節
H1("第3節　第1層　機械で確かめる（30分）")

H2("1　スクリプトを再実行し、終了コードが0であること")

P("自己点検を備えたスクリプトは、**1件でも不適合があると終了コード1で"
  "終わります。** 点検する側は終了コードを見ればよく、"
  "出力の中身を読む必要はありません。")

CODE("cd /home/user/repository")
CODE("python3 build_survey_report.py            && echo OK")
CODE("python3 build_survey_jisseki_cross.py     && echo OK")
CODE("python3 build_survey_monthly_r8_09.py     && echo OK")
CODE("python3 build_survey_report_doc.py        && echo OK")
CODE("python3 build_survey_review.py            && echo OK")
CODE("python3 build_survey_status.py            && echo OK")
CODE("python3 build_survey_inquiry.py           && echo OK")
CODE("python3 build_survey_check.py             && echo OK   ← 本書")

NOTE("`build_survey_crosstab.py` は個票CSVを読むため再実行できません"
     "（第9節）。")

H2("2　docx が検証に適合すること")

P("docx は python-docx が生の要素を並べる作りであるため、"
  "**要素の順序や必須属性を欠くと Word で開けなくなります。**"
  "作った docx は必ず検証します。")

CODE("python3 /mnt/skills/public/docx/scripts/office/validate.py \\")
CODE('  "output/第10期計画_実施済み調査_結果報告書.docx"')
CODE("python3 /mnt/skills/public/docx/scripts/office/validate.py \\")
CODE('  "output/第10期計画_アンケート調査の分析_令和8年9月業務報告.docx"')
CODE("python3 /mnt/skills/public/docx/scripts/office/validate.py \\")
CODE('  "output/第10期計画_アンケート分析_点検箇所と手順.docx"')

P("「All validations PASSED!」と出れば適合です。")

H2("3　出力に差分が出ていないこと")

P("再実行して内容が変わらないはずのものが変わっていれば、"
  "**どこかが現時点の値を読んでいます**（第7節の4件目）。")

CODE("git status --short output/")

# ============================================================ 第4節
H1("第4節　第2層　数値の出所を辿る（2〜3時間）")

P("**同じ数値を2つの成果品に別々に書かない**というのが当方の決めごとです。"
  "報告書に書いてある数値が、どこから来ているのかを辿ります。")

CAP("数値の出所の対照")
TBL(["値", "収めている場所", "どこから来ているか"],
    DEDOKORO, [4.6, 4.4, 8.2])

P("")
H2("1　固定値が書かれていないか")

P("次の書き方であれば固定値ではありません。")
BUL("`runpy.run_path(...)` で算定スクリプトを読んでいる")
BUL("`ast.literal_eval` でソースの定数を読んでいる")
BUL("`load_workbook(...)` で成果品の実物を開いて数えている")
BUL("`RP.draft_label()` で計画素案の実物を数えている")

P("**逆に、次の書き方があれば固定値です。**"
  "その値を改めたときに追随しないため、点検の対象になります。")
BUL("数字が文字列としてそのまま書かれている（「20件」「15シート」など）")
BUL("同じ数字が2つ以上のスクリプトに現れる")

NOTE("探し方の例：`grep -n \"件\\|シート\" build_survey_*.py` で"
     "数字を伴う語を拾い、その値がどこから来ているかを見ます。")

H2("2　2つの成果品で値が食い違っていないか")

P("特に見る組み合わせは次の3つです。")
CAP("突き合わせる組み合わせ")
TBL(["①", "②", "一致していること"],
    [("集計分析報告書 10シートの見込量", "計画素案 第6章",
      "いずれも `build_mikomiryo_santei.py` の値であること"),
     ("集計分析報告書 11シートの留保", "同 14シートの解消",
      "留保の番号が対応していること。解消したものに根拠があること"),
     ("集計分析報告書 12シートの反映状況", "計画素案の実物",
      "「反映済」としたものが素案の当該章節に実際にあること")],
    [5.0, 4.4, 7.8])

P("")
H2("3　突合の測定値が再現できるか")

P("突合（`build_survey_jisseki_cross.py`）の測定値は次のとおりです。"
  "**本書はこの値を書き写しておらず、スクリプトを読んで表示しています。**")
BUL("①在宅生活改善調査の要介護度分布　χ²＝%.1f（自由度6・1％点%.2f・"
    "n＝%d）" % (CU_CHI, CHI_1, CU_N))
BUL("②居所変更実態調査の在籍者と年報の受給者は一致しない"
    "（介護老人保健施設だけが在籍＞受給）")

# ============================================================ 第5節
H1("第5節　第3層　記述の当否と守るべき制約（半日）")

P("守るべき制約は発注者の指示によるものです。"
  "**言い換えや要約ではなく、そのまま守ります。**")

CAP("この分野に効く守るべき制約")
TBL(["制約", "この業務でどこに効くか"], SEIYAKU, [7.2, 10.0])

P("")
H2("1　計画素案の記述")

P("計画素案は現在 %s です。" % DRAFT)
P("アンケートの結果が入っているのは**第2章第4節（高齢者の生活実態）**です。"
  "④健康とくらしの調査の結果を令和8年9月15日に反映しました。"
  "①②③の結果は第2章第5節・第2章第6節・第4章第3節に入っています。")

BUL("調査の対象者の範囲が書かれていること"
    "（④は要支援者・要介護者を含まない。他調査と合算しない）")
BUL("町間の差を述べるときに年齢調整をしていること")
BUL("地区の定義（小学校区とするか）が［要確認］のまま残っていること"
    "（確認事項No.9）")
BUL("本文に確認事項の注記が書かれていないこと"
    "（注記は出典・単位・欠測・数値の性質・記号の意味に限る）")

# ============================================================ 第6節
H1("第6節　特に見ていただきたい箇所")

P("**過去に誤りが出た型、又は出やすい型です。**"
  "第3層の中でも、ここだけは必ず見てください。")

CAP("特に見ていただきたい箇所")
TBL(["箇所", "何を確かめるか", "なぜ"],
    MIRU, [3.0, 5.6, 8.6], first_bold=True)

P("")
P("主要所見は%d件で、反映の状況は%sです。"
  % (len(SHOKEN),
     "・".join("%s%d件" % (k, v)
               for k, v in sorted(SHOKEN_ST.items(), key=lambda x: -x[1]))))
P("**「反映済」としたものは、計画素案の実物を章節ごとに読んで判定しています。**"
  "語の有無だけで判定すると偽陽性が出ます（第7節の3件目）。")

# ============================================================ 第7節
H1("第7節　この作業で実際に起きた誤り")

P("**いずれも実際に起きたものです。**"
  "同じ型の誤りが残っていないかという見方で点検してください。")

CAP("実際に起きた誤りと手立て")
TBL(["誤り", "どうなっていたか", "なぜ起きたか", "手立て"],
    AYAMARI, [3.4, 4.6, 3.6, 5.6], first_bold=True)

P("")
P("**末尾の7件は本書と同時に作った9月業務報告及び計画素案で"
  "起きたものです。**"
  "1件は自己点検（「測定済み」の区分があること）が不適合になったことで、"
  "他はご指摘と再レビューで気づきました。"
  "**割合の分母の取違いは、原典の該当数（1,002人）から逆算して分かりました。**"
  "**自己点検が置かれていなかった表に誤りが残っていました。**"
  "その後、表の値が算定と一致することを確かめる点検を加えています。"
  "自己点検は、点検する側の手間を減らすだけでなく、"
  "**作る側が誤りに気づく仕掛けでもあります。**")

# ============================================================ 第8節
H1("第8節　個人情報・禁止表現の探し方")

P("**語の有無だけで探すと見落とします。**"
  "個人情報は値の形（正規表現）で、禁止表現は語で探します。")

H2("1　個人情報")

CODE("grep -nE '0[0-9]{1,4}-[0-9]{2,4}-[0-9]{3,4}' data_survey*.py "
     "build_survey*.py")
CODE("grep -nE '[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+' data_survey*.py "
     "build_survey*.py")

P("電話番号・メールアドレスのほか、**施設の管理者名・記入者名**も収録しません。"
  "調査票には記入者欄がありますが、"
  "`data_survey_entry.py` には提出元（事業所名）・ファイル名・票番号・"
  "所在地区欄の記載文字列しか収めていません。")

NOTE("個票（①利用者票99票・③職員個票317人・④個票4,729票）は"
     "リポジトリに格納していません。"
     "④の集計値は `data_survey_jages.py` に収めています。")

H2("2　禁止表現")

CODE("grep -n 'に由来する\\|と整合する\\|1件も\\|"
     "有意差がないため関係がない\\|全国トップ級' \\")
CODE("  build_survey*.py data_survey*.py")

P("**点検方法そのものを書いたシート（本書を含む）は判定から除きます。**"
  "禁止表現の一覧を掲げている箇所が引っ掛かるためです。"
  "それ以外で見つかったものは言い換えます。")

CAP("禁止表現と言い換えの例")
TBL(["禁止表現", "なぜ", "言い換えの例"],
    [("〜に由来する", "因果を断定している",
      "〜と同じ向きである／〜の後に生じている"),
     ("〜と整合する（根拠として）", "別の資料の一致を根拠にしている",
      "〜にも同じ記載がある"),
     ("1件も〜ない", "網羅を断定している", "確認した範囲では見当たらない"),
     ("有意差がないため関係がない", "検定の結果を誤って読んでいる",
      "有意な差は認められなかった（関係がないことを示すものではない）"),
     ("全国トップ級", "根拠のない序列", "同規模保険者40のうち上位である")],
    [3.4, 4.6, 9.2], first_bold=True)

# ============================================================ 第9節
H1("第9節　再実行できないもの")

P("アンケート関係のスクリプトのうち、**再実行できないのは1件だけです。**")

CAP("再実行できないスクリプト")
TBL(["スクリプト", "読むもの", "なぜできないか", "扱い"],
    [("build_survey_crosstab.py",
      "④健康とくらしの調査の個票CSV（4,729票）",
      "個票は発注者指示により格納せず、作業領域にも残っていない",
      "成果品（%dシート）は収録済みで内容は確定している。"
      "作り直しが要るときだけ個票の再提供を依頼する（確認事項No.139）"
      % _sheets("第10期計画_調査クロス集計・分析.xlsx"))],
    [4.0, 3.6, 4.4, 5.2])

P("")
P("**計画素案 第2章第4節の側は `data_survey_jages.py` により"
  "個票なしで再現できます。** 個票が要るのは"
  "クロス集計・分析（%dシート）を作り直すときだけです。"
  % _sheets("第10期計画_調査クロス集計・分析.xlsx"))

# ============================================================ 第10節
H1("第10節　点検の記録様式と分担")

H2("1　記録様式")

CAP("見つけたものの区分と扱い")
TBL(["区分", "あてはまるもの", "扱い"], KIROKU_SHIKI,
    [1.4, 8.0, 7.8], center={0}, first_bold=True)

P("")
P("記録には次の4つを書いてください。"
  "**「おかしい気がする」だけでは直せません。**")
BUL("どの成果品の・どのシート（節）の・どの行か")
BUL("いま何と書いてあるか（そのまま写す）")
BUL("何が問題か（数値／出所／記述／制約のどれか）")
BUL("どうあるべきと考えるか（分からなければ「分からない」と書く）")

H2("2　分担の案")

CAP("分担の案")
TBL(["層", "だれが", "目安", "何を確かめるか"],
    [(g[0], g[1], g[2], g[3]) for g in BUNTAN],
    [1.4, 2.6, 1.4, 11.8], center={0, 2})

P("")
P("**第1層はどなたでもできます。**"
  "コマンドを順に実行し、終了コードが0であること・"
  "検証が「All validations PASSED!」であることを見るだけです。")

# ============================================================ 自己点検
H1("自己点検")

P("本書は自らの記述を次のとおり点検しています。"
  "**1件でも不適合があれば出力せずに終わります。**")

_all_text = []


def _collect():
    out = []
    for p in doc.paragraphs:
        out.append(p.text)
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                out.append(c.text)
    return "\n".join(out)


_txt = _collect()

chk(1, "成果品の規模を実物から数えていること",
    "%d件のうち実体のあるもの%d件"
    % (len(SEIKA),
       sum(1 for fn, _s, _k in SEIKA
           if os.path.exists(os.path.join(RP.OUTPUT, fn)))),
    all(os.path.exists(os.path.join(RP.OUTPUT, fn))
        for fn, _s, _k in SEIKA))
chk(2, "自己点検の件数をソースから数えていること",
    "突合 %d件・9月業務報告 %d件"
    % (_chk_count("build_survey_jisseki_cross.py"),
       _chk_count("build_survey_monthly_r8_09.py")),
    _chk_count("build_survey_jisseki_cross.py") > 0
    and _chk_count("build_survey_monthly_r8_09.py") > 0)
chk(3, "留保・所見を集計分析報告書の実物から読んでいること",
    "留保%d件・所見%d件" % (len(RYUHO), len(SHOKEN)),
    len(RYUHO) > 0 and len(SHOKEN) > 0)
chk(4, "所見の状態の内訳が総数と一致すること",
    "%d件＝%s" % (len(SHOKEN),
                  "＋".join("%s%d" % (k, v) for k, v in SHOKEN_ST.items())),
    len(SHOKEN) == sum(SHOKEN_ST.values()))
chk(5, "突合の測定値をスクリプトから読んでいること",
    "χ²＝%.1f（1％点%.2f・n＝%d）" % (CU_CHI, CHI_1, CU_N),
    CU_CHI > CHI_1)
chk(6, "計画素案の規模を実物から数えていること", DRAFT, "段落" in DRAFT)
chk(7, "強調の記号が本文に残っていないこと",
    "** の残り %d件" % _txt.count("**"), _txt.count("**") == 0)
# 個人情報は語ではなく値の形（正規表現）で判定する（CLAUDE.md §4）。
_tel = re.findall(r"0\d{1,4}-\d{2,4}-\d{3,4}", _txt)
_mail = re.findall(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", _txt)
chk(8, "電話番号・メールアドレスを書いていないこと",
    "電話%d件・メール%d件" % (len(_tel), len(_mail)),
    not _tel and not _mail)
chk(9, "点検の3層がすべて節として立っていること",
    "第1層・第2層・第3層",
    all(("第%d層" % i) in _txt for i in (1, 2, 3)))
chk(10, "実際に起きた誤りを掲げていること",
    "%d件" % len(AYAMARI), len(AYAMARI) >= 5)
chk(11, "再実行できないものを明示していること",
    "1件（build_survey_crosstab.py）",
    "build_survey_crosstab.py" in _txt)
chk(12, "本書が社内保管であることを明示していること",
    "表紙に記載", "社内の点検のための資料" in _txt)

CAP("自己点検")
TBL(["No.", "点検の内容", "結果", "判定"],
    [(str(n), naiyo, kekka, hantei) for n, naiyo, kekka, hantei in CHECKS],
    [1.0, 7.6, 5.0, 1.6], center={0, 3})

_ng = [c for c in CHECKS if c[3] != "適合"]

# ============================================================ 出力
_z = doc.settings.element.find(qn("w:zoom"))
if _z is None:
    _z = OxmlElement("w:zoom")
    doc.settings.element.insert(0, _z)
_z.set(qn("w:percent"), "100")

if _ng:
    print("自己点検 %d件：適合%d件・不適合%d件"
          % (len(CHECKS), len(CHECKS) - len(_ng), len(_ng)))
    for c in _ng:
        print("  不適合: %s 「%s」 %s" % (c[0], c[1], c[2]))
    sys.exit(1)

doc.save(OUT)
print("書き出しました: %s" % OUT)
print("段落 %d ／ 表 %d" % (len(doc.paragraphs), len(doc.tables)))
print("成果品 %d件／留保 %d件／主要所見 %d件"
      % (len(SEIKA), len(RYUHO), len(SHOKEN)))
print("自己点検 %d件：すべて適合" % len(CHECKS))
