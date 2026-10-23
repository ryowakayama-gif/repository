# -*- coding: utf-8 -*-
"""Red Team レビュー（令和8年10月24日）への対応：成果物とエビデンスの一式を ZIP にする

  python3 07_ソーススクリプト/build_redteam_teishutsu_R8.10.24.py

ご指示「添付レビュー確認の上、必要な修正を行った上で成果物とエビデンスをまとめてジップファイルに
してアウトプットして下さい」（令和8年10月24日）による。

対応表の正本は本スクリプトの TAIOU（md と xlsx はここから書き出す。2か所に持たない）。
点検のログ・検算・差分・PDF は、その場で成果物から作る（手で書かない）。

出力
  11_提出一式/川崎町_RedTeamレビュー対応_成果物とエビデンス_R8.10.24.zip
  01_第10期_最新版成果品/川崎町_RedTeamレビュー対応表_R8.10.24.md
  05_試算・管理シート/川崎町_RedTeamレビュー対応表_R8.10.24.xlsx
中間の組立ての場所はスクラップパッド（ZIP に入れた後は要らない）。
終了コード　0＝点検ログがすべて終了コード0・ZIP の中身の MD5 が一覧と一致／1＝それ以外
"""
import difflib
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

import docx
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

HI = "R8.10.24"
ZIPNAME = f"11_提出一式/川崎町_RedTeamレビュー対応_成果物とエビデンス_{HI}.zip"
MD_OUT = f"01_第10期_最新版成果品/川崎町_RedTeamレビュー対応表_{HI}.md"
XLSX_OUT = f"05_試算・管理シート/川崎町_RedTeamレビュー対応表_{HI}.xlsx"
REVIEW = "09_元資料/レビュー/川崎町_第10期計画_RedTeamレビュー_R8.10.24受領.md"

SEIKA = [  # (区分, ファイル, 旧の版（レビューの対象）)
    ("協議用素案", "01_第10期_最新版成果品/川崎町_計画書素案_v2.32_協議用.docx",
     "01_第10期_最新版成果品/川崎町_計画書素案_v2.29_協議用.docx"),
    ("公表用の計画書（案）", "01_第10期_最新版成果品/川崎町_計画書案_公表用_v2.32.docx",
     "01_第10期_最新版成果品/川崎町_計画書案_公表用_v2.29.docx"),
    ("概要版", "01_第10期_最新版成果品/川崎町_計画書案_概要版_v2.32.docx",
     "01_第10期_最新版成果品/川崎町_計画書案_概要版_v2.30.docx"),
    ("第2回策定委員会資料", "03_委員会・説明資料/川崎町_第2回策定委員会資料_R8.11_v17.docx",
     "03_委員会・説明資料/川崎町_第2回策定委員会資料_R8.11_v16.docx"),
    ("想定問答集", "03_委員会・説明資料/川崎町_第2回策定委員会_想定問答集_R8.11_v14.docx",
     "03_委員会・説明資料/川崎町_第2回策定委員会_想定問答集_R8.11_v13.docx"),
]
S = SEIKA[0][1]
P = SEIKA[1][1]
SC = "07_ソーススクリプト/"
TENKEN = [  # (名前, コマンド)
    ("01_素案_自己点検", [SC + "check_soan_selfcheck_R8.9.28.py", S]),
    ("02_素案_保険料の検算", [SC + "check_hokenryo_R8.9.30.py", S]),
    ("03_素案_書体と紙面", [SC + "check_font_R8.9.30.py", S]),
    ("04_素案_8軸の網羅性", [SC + "check_8jiku_R8.10.4.py", S]),
    ("05_素案_章節の参照の機械検索", [SC + "check_sansho_R8.10.24.py", S]),
    ("06_図表データ管理台帳", [SC + "build_zuhyo_daicho.py"]),
    ("07_事業所の受入れ様式", [SC + "check_jigyosho_uke_R8.10.15.py"]),
    ("08_見込量の漏れ", [SC + "check_mikomi_morasi_R8.10.16.py"]),
    ("09_認知症の人の数の整合", [SC + "check_ninchisho_seigo_R8.10.19.py"]),
    ("10_認定者数", [SC + "check_nintei_R8.10.16.py"]),
    ("11_公表用_書体と紙面", [SC + "check_font_R8.9.30.py", P]),
    ("12_委員会資料_自己点検", [SC + "check_2kai_selfcheck_R8.9.28.py"]),
    ("13_守るべき制約の走査", [SC + "check_seiyaku_R8.10.6.py"]),
    ("14_リポジトリの健全性", [SC + "check_repo_kenzensei_R8.10.4.py"]),
]
SCRIPTS = ["fix_soan_v2310_kakutei.py", "fix_soan_v2320_redteam.py", "build_kouhyou_soan.py",
           "build_gaiyo_v230.py", "build_2kai_v16.py", "build_2kai_v17.py", "build_2kai_qa_v13.py",
           "build_2kai_qa_v14.py", "check_sansho_R8.10.24.py", "fill_toc_pages_v200.py",
           "fetch_kinrin_chiiki_R8.10.23.py", "data_zuhyo.py", "build_redteam_teishutsu_R8.10.24.py"]

# ══════════════════════════ 対応表（正本）
# 区分：是正済／一部是正／既に対応／未実施（資料未受領）／町のご判断
TAIOU = [
    ("A01", "重大", "添付の素案が v2.29 で、点検メモの v2.31 と食い違い、修正結果を検証できない",
     "是正済", "提出の一式を 協議用素案 Ver.2.32・公表用 v2.32・概要版 v2.32・委員会資料 v17・想定問答集 v14 に"
     "そろえた。成果品一覧（MD5・頁数）と、レビュー対象版からの差分を同梱", "全5冊",
     "00_README.md／03_エビデンス/差分_*.txt", ""),
    ("A02", "重大", "給付費総額が998,220千円と998,222千円で並ぶ",
     "是正済", "年報の円単位（998,222,421円）を千円未満四捨五入した998,222千円に統一。図3-2 の実績は10千円単位に"
     "丸めた値だったため千円単位（322,676・182,194・493,352）に改め図を作り直した（v2.31）",
     "素案 3-3・図3-2／概要版 3／資料 表37", "03_エビデンス/検算.md §1・旧表現の残存検索.txt", ""),
    ("A03", "高", "選択２の5,959円・6,382円と「±211円」が合わない",
     "是正済", "5,959円・6,382円は算定Ｂが6,170円だった時の値（6,170∓211）が残っていた。現在の算定Ｂ6,143円∓211円＝"
     "5,932円・6,354円に改めた。211円は表101 と調整交付金の見込交付割合から検算（211.7円）し、基金の取崩しの額に"
     "よらないことを素案 5-5 に書いた", "資料 v17 選択２／素案 5-5", "03_エビデンス/検算.md §2", ""),
    ("A04", "高", "＋66円・＋196円は算定Ａの差で、算定Ｂを基準とする見出しと合わない",
     "是正済", "表27 の見出しを「算定Ａ（取崩しなし）を基準」に改め、報酬改定の行に算定Ｂ（＋92円・＋223円）・"
     "算定Ｃ（＋119円・＋250円）を併記。表28 に基準の注記。問答集3問も同じ",
     "資料 v17 表27・表28／問答集 v14", "03_エビデンス/検算.md §3", ""),
    ("A05", "高", "11事項の件数が表（対応済1・一部8・未対応2）と本文（一部7・未対応3）で違う",
     "是正済", "本文を表から数えた件数に改めた。未対応2事項は確認事項No.156・No.158、一部対応のうち4事項は"
     "No.155・No.157・No.138 に紐づけ済み（表の右列）", "素案 1-5", "03_エビデンス/検算.md §4",
     "一部対応の一・六・九は町のデータ待ち"),
    ("A06", "高", "交付金の要件を言い切る表現（得点できます・要件を満たします）",
     "是正済", "「評価の対象となり得る。評価は国の該当状況調査への回答と、実施・記録・反映の結果による」に統一。"
     "素案8か所（v2.31で4・v2.32で4）、資料2か所、問答集3問", "素案 3-3・5-4・7-6・10-4／資料 表36・表38／問答集",
     "03_エビデンス/旧表現の残存検索.txt", "実際の評価は町の該当状況調査への回答による（当方では確かめられない）"),
    ("A07", "高", "11の課題と重点課題5本を橋渡しする整理がない",
     "是正済", "素案 3-4 に11→5の集約表（根拠となる事実・集約した課題番号・接続する施策の章節）、4-3 に5本と"
     "4つの重点の対応を置いた（v2.31）。概要版にも5本の表を加えた（v2.32）。7つの柱との対応は 4-2・3-4 の表の章節",
     "素案 3-4・4-3／概要版 3／資料 表13", "01_成果物", ""),
    ("A08", "中", "資料 v16 を PDF にすると22頁（点検メモは21頁）",
     "一部是正", "当方の環境（HG丸ｺﾞｼｯｸM-PRO を IPAゴシックで代用）で PDF を作り、頁数を成果品一覧に記録した。"
     "頁数は書体の環境で変わるため、配布用の PDF は町のパソコン（Word）で作り、頁送りを確かめていただく",
     "資料 v17", "01_成果物/*.pdf・00_README.md", "配布用 PDF の確定は町の環境で（No.53）"),
    ("A09", "中", "目次の「ご挨拶」が3頁（実際は2頁）",
     "是正済", "目次の頁番号を入れるスクリプトが、見出し「ご　挨　拶」でなく目次の行の語で探し、目次の頁を拾っていた。"
     "見出しの語で探すよう直し、2頁となった", "素案・公表用 目次", "03_エビデンス/点検ログ（fill_toc）", ""),
    ("A10", "中", "「第1期から第9期までの25年間」（27年間）、概要版「11つ」",
     "是正済", "「第1期（平成12年度）から第9期（令和8年度）までの27年間」。概要版は「11の課題を整理しました。"
     "これらを5つの重点課題に集約して取り組みます」", "素案 1-1／概要版 3", "03_エビデンス/差分", ""),
    ("A11", "中", "「町内の8の事業所」、介護職員が常勤換算であることが分からない",
     "是正済", "「町内8事業所」（素案3か所・概要版）。概要版は「施設・居住系の介護職員数（常勤換算）」",
     "素案 2-7・7-1／概要版 3", "03_エビデンス/差分", ""),
    ("A12", "中", "所得段階の「詳細は第7章で論じます」が詳述先（9-3）と合わない。他の参照先も機械検索",
     "是正済", "「乗率の検討は第9章9-3に示します」。機械検索（check_sansho）で、実在しない参照「9-4の感度表」"
     "2か所を新たに見つけ 9-3 に改めた。実在しない参照・章と節の食い違いは0件",
     "素案 2-2・9-3", "03_エビデンス/点検ログ/05_素案_章節の参照の機械検索.txt",
     "参照の内容の当否（12件の章名指し）は一覧を人が確かめる"),
    ("B01", "最優先", "見込量に供給制約が入っていない可能性",
     "未実施（資料未受領）", "稼働率・空き床・介護職員の推移は素案 2-7 に示したが、町内外別の利用・断り件数・待機の重複除去は"
     "未受領。供給の維持・縮小・改善の見込量の感度は第3回に向けて作る", "素案 2-7・7-2・9-1",
     "―", "確認事項No.159・147・184"),
    ("B02", "最優先", "基金残高は見込み。取崩し・収納率・調整交付金・負担割合の審議が不十分",
     "町のご判断", "資料は「選択５は方向性のみ・第３回で決定」とし、確定していない事項（表29）に並べている。"
     "基金の確定残高・調整交付金・交付金実績は第3回までに確定", "資料 選択５・表29",
     "―", "確認事項No.85・78・39・11"),
    ("B03", "高", "「国は基金残高の概ね半分以下を取崩水準として示している」の根拠文書を確認できない",
     "是正済", "根拠文書を確かめられないため、国の基準としての記述を削り、事務局案の考え方（基金残高の見通しと"
     "第11期以降の平準化を踏まえた比較案の一つ）に改めた", "資料 v17 選択５／問答集 v14",
     "03_エビデンス/旧表現の残存検索.txt", ""),
    ("B04", "高", "報酬改定・制度改正をいつ・どこで反映するか不明瞭",
     "一部是正", "令和7年度単価で固定していることと報酬改定の感度（算定Ａ・Ｂ・Ｃ）を明記。入力セルを分けた試算は"
     "令和9年度改定の公表後（令和9年1月頃）に作る", "素案 9-3／資料 表26〜28", "03_エビデンス/検算.md §3",
     "確認事項No.148"),
    ("B05", "高", "空床約6人分と在所28/30人は測定方法・時点が違う",
     "既に対応", "資料・素案とも「時点の在所者数による入所率」と「延べ利用日数による稼働率」を書き分けている。"
     "受入れの制限・町民と町外の利用の別は未把握", "資料 ３ 表6〜表10／素案 2-7", "―", "確認事項No.159"),
    ("B06", "高", "人材施策の案Ｂ・Ｃに事業費の概算・財源の確度がない",
     "未実施", "選択４のご判断の後に積算する。資料は「町長部局・財政担当との調整が必要」と記載済み",
     "資料 選択４", "―", "選択４のご判断後"),
    ("B07", "高", "27点は評価点であり交付金額ではない。委員会開催のみで要件適合とは言えない",
     "是正済", "A06 と同じ。素案 3-3 は「得点は交付額に直接に換算できません」と記載済み",
     "素案 3-3／資料 ７", "03_エビデンス/旧表現の残存検索.txt", ""),
    ("B08", "高", "法定・重点記載事項の一部未確定",
     "一部是正", "1-5 の件数を是正。国の基本指針の告示と県の第10期計画（中間案）を受けて照合し直す",
     "素案 1-2・1-5", "―", "確認事項No.163"),
    ("B09", "中", "認知症の意見聴取の実施結果・ケアパスが未確認",
     "未実施", "意見聴取は令和8年10〜11月に実施予定。ケアパスは町のご確認待ち", "素案 第6章",
     "―", "確認事項No.90・137"),
    ("B10", "中", "KPI の目標値が「増加」「委員会協議」のまま",
     "町のご判断", "目標値は策定委員会の協議事項。概要版には確定しているものだけを載せている",
     "素案 10-4", "―", "確認事項No.70・34"),
    ("B11", "中", "生活支援コーディネーター25人と生活支援サポーター25名の定義確認",
     "是正済", "原票で確かめた。町の実績データ（生活支援ｻﾎﾟｰﾀｰ(SC)_人数：令和6年度・令和7年度見込とも26人）と、"
     "国の評価（令和8年度・該当状況調査26人）が一致。素案の25名は地域福祉計画の資料の内訳（生活支援サポーター"
     "25名）の値であったため、生活支援コーディネーターは26人に改めた", "素案 3-3・5-4・10-4／概要版",
     "03_エビデンス/原票抜粋_生活支援コーディネーター.csv", ""),
    ("B12", "中", "Ｃ案で組み込んだ素案のまま選択を諮ると結論先取りに見える",
     "是正済", "選択１に「素案は町のご回答（令和8年9月25日）によりＣ案で仮に組んでおり、委員会のご承認の前のもの。"
     "Ａ又はＢとなった場合は第5章〜第8章の節・項を組み替える」と明記", "資料 v17 選択１", "03_エビデンス/差分", ""),
    ("B13", "中", "仕様書は4回予定、事務局案は3回。工程が逼迫",
     "町のご判断", "資料８に4案と納品日を並べ、委託仕様書11による協議をお願いしている", "資料 ８",
     "―", "確認事項No.49・19"),
    ("B14", "中", "軽費老人ホーム廃止前の町民入所・移った先が未確定",
     "未実施（資料未受領）", "廃止届の記載内容を県・町に照会中", "資料 表6 注／素案 2-7", "―", "確認事項No.154"),
    ("B15", "中", "8事業所と14の指定、49.4％と68.1％の集計単位・定義",
     "是正済", "資料・問答集を「町内8事業所の14の指定」に統一。選択２の入所・居住系の範囲を素案と同じ定義に改めた"
     "（49.4％は施設サービスの給付費の構成比、68.1％は入所・居住系）", "資料 v17／問答集 v14",
     "03_エビデンス/旧表現の残存検索.txt", ""),
]
ZANSON = [  # 旧表現の残存検索（レビューで挙がった語）
    "998,220", "5,959", "6,382", "（算定Ｂ 6,143円を基準）", "25年間", "11つ", "8の事業所", "第7章で論じ",
    "得点できます", "得点できる", "要件を満たします", "そのまま要件", "概ね半分以下", "対応済1・一部7",
    "生活支援コーディネーター25", "SC25名", "9-4の感度表", "町内14の指定", "58.67", "第117条第５項",
]


def md5(p):
    return hashlib.md5(open(p, "rb").read()).hexdigest()


def zenbun(path):
    d = docx.Document(path)
    out = [p.text for p in d.paragraphs]
    for t in d.tables:
        for r in t.rows:
            seen = []
            for c in r.cells:
                if c._tc in seen:
                    continue
                seen.append(c._tc)
                out.append(c.text)
    return out


def to_pdf(src, outdir):
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", outdir, src],
                   capture_output=True, timeout=600)
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(src))[0] + ".pdf")
    import pymupdf
    return pdf, (len(pymupdf.open(pdf)) if os.path.exists(pdf) else 0)


def kensan():
    """検算.md の本文（成果物から読む）。"""
    s = zenbun(S)
    sz = "\n".join(s)
    L = ["# 検算（成果物から読み直した値）", "", f"対象：{os.path.basename(S)}・{os.path.basename(SEIKA[3][1])}", ""]
    d = docx.Document(S)
    t24 = [t for t in d.tables if [c.text for c in t.rows[0].cells][:4] == ["サービス区分", "要支援計", "要介護計", "合計（円）"]][0]
    en = {r.cells[0].text.split("\n")[0]: int(r.cells[3].text.replace(",", "")) for r in t24.rows[1:]}
    ks = [round(en[k] / 1000) for k in ("居宅サービス計", "地域密着型計", "施設サービス計")]
    L += ["## §1 令和7年度の給付費（A02）", "",
          f"- 素案 2-4 の表（円）：居宅 {en['居宅サービス計']:,}／地域密着型 {en['地域密着型計']:,}／施設 {en['施設サービス計']:,}／総計 {en['総計']:,}",
          f"- 千円未満四捨五入：{ks[0]:,}＋{ks[1]:,}＋{ks[2]:,}＝{sum(ks):,}千円（総計 {round(en['総計'] / 1000):,}千円）",
          f"- 達成率：{round(en['総計'] / 1000):,}÷1,011,020＝{round(en['総計'] / 1000) / 1011020 * 100:.1f}％", ""]
    tanka = float(re.search(r"1人1月あたり給付費は約([\d.]+)万円", sz).group(1)) * 10
    rates = [float(x) / 100 for x in re.findall(r"令和\d+年度([\d.]+)％", [x for x in s if "見込交付割合は全国の仮置値" in x][0])]
    nn = float(re.search(r"([\d,.]+)人としています", sz).group(1).replace(",", ""))
    dj = sum(10 * tanka * 12 * (0.23 + 0.05 - e) for e in rates)
    dm = dj * 1000 / 0.96 / nn / 12
    b = int(re.search(r"パターンB ([\d,]+)円", sz).group(1).replace(",", ""))
    L += ["## §2 入所・居住系±10人の月額（A03）", "",
          f"- 1人1月あたり給付費 約{tanka / 10:.1f}万円（素案 5-5）、調整交付金の見込交付割合 "
          + "・".join(f"{e * 100:.2f}％" for e in rates) + "（素案 9-3）",
          f"- 保険料収納必要額の増減＝10人×{tanka:.0f}千円×12か月×Σ(23％＋5％－見込交付割合)＝{dj:,.0f}千円（3か年）",
          f"- 月額＝{dj:,.0f}千円÷予定収納率96.0％÷補正後被保険者数{nn:,.1f}人÷12＝**{dm:.1f}円**（約211円）",
          f"- 基金の取崩しは差し引く定額であり、この幅は算定Ａ・Ｂ・Ｃで同じ",
          f"- 選択２：算定Ｂ {b:,}円 → 縮小 {b - 211:,}円（▲211円）／増床 {b + 211:,}円（＋211円）", ""]
    dd = docx.Document(SEIKA[3][1])
    t26 = [t for t in dd.tables if t.rows[0].cells[0].text.strip() == "介護報酬改定率"][0]
    g = {r.cells[0].text.strip()[:2]: [int(re.sub(r"\D", "", c.text)) for c in r.cells[1:4]] for r in t26.rows[1:4]}
    L += ["## §3 介護報酬改定の差額（A04）", "", "| 改定率 | 算定Ａ | 算定Ｂ | 算定Ｃ |", "|---|---|---|---|"]
    for k in ("０％", "１％", "３％"):
        L.append(f"| {k} | {g[k][0]:,}円 | {g[k][1]:,}円 | {g[k][2]:,}円 |")
    L += ["", f"- 1％：Ａ＋{g['１％'][0] - g['０％'][0]}・Ｂ＋{g['１％'][1] - g['０％'][1]}・Ｃ＋{g['１％'][2] - g['０％'][2]}円"
          f"／3％：Ａ＋{g['３％'][0] - g['０％'][0]}・Ｂ＋{g['３％'][1] - g['０％'][1]}・Ｃ＋{g['３％'][2] - g['０％'][2]}円", ""]
    t15 = [t for t in d.tables if [c.text.strip() for c in t.rows[0].cells][:3] == ["事項", "確認すべき指標・状況", "対応"]][0]
    k = [r.cells[2].text.strip() for r in t15.rows[1:]]
    L += ["## §4 基本指針案の別表11事項（A05）", "",
          f"- 表の対応欄：対応済{k.count('対応済')}・一部{k.count('一部')}・未対応{k.count('未対応')}（計{len(k)}）",
          "- 本文：" + [x for x in s if x.startswith("対応済")][0][:40], ""]
    L += ["## §5 生活支援コーディネーター（B11）", "", "- 原票抜粋は 原票抜粋_生活支援コーディネーター.csv", ""]
    return "\n".join(L)


def sc_genpyo(path):
    import csv
    rows = [["出所", "区分", "令和6年度", "令和7年度見込", "令和8年度", "備考"]]
    ws = openpyxl.load_workbook("09_元資料/R7実績データ/★【川崎町様】データ入力フォーマット.xlsx",
                                read_only=True, data_only=True)["02_総合事業_介護予防"]
    rr = list(ws.iter_rows(values_only=True))
    h = rr[3]
    r = dict(zip(h, [x for x in rr if x and any(v and "SC" in str(v) for v in x)][0]))
    rows.append(["町の実績データ（データ入力フォーマット 02シート）", r["事業名"], r["R6"], r["R7見込"],
                 r["R8当初・見込"], "町入力"])
    f = ("09_元資料/交付金評価/①令和８年度保険者機能強化推進交付金等（市町村）に係る全国集計結果/"
         "05-2_（支援）市町村評価指標（ⅱ活動指標群）.xlsx")
    ws = openpyxl.load_workbook(f, read_only=True, data_only=True)["目標Ⅰ"]
    rr = list(ws.iter_rows(values_only=True))
    hd = rr[4]
    col = [j for j, x in enumerate(hd) if x and str(x).startswith("生活支援コーディネーター数\n（該当状況調査）")][0]
    k = [x for x in rr if x and str(x[1]) == "043240"][0]
    rows.append(["国 令和8年度評価（05-2 活動指標群 目標Ⅰ）", "生活支援コーディネーター数（該当状況調査）",
                 k[col], "", "", f"市町村コード043240 {k[2]}（福岡県川崎町406058とは別）"])
    with open(path, "w", encoding="utf-8-sig", newline="") as fp:
        csv.writer(fp).writerows(rows)


def jobun_117(path):
    j = json.load(open("09_元資料/法令/介護保険法_eGov_令和8年10月1日施行時点.json", encoding="utf-8"))

    def text(n):
        if isinstance(n, str):
            return n
        if isinstance(n, dict):
            return "".join(text(c) for c in n.get("children", []))
        return "".join(text(c) for c in n) if isinstance(n, list) else ""

    def find(n):
        if isinstance(n, dict):
            if n.get("tag") == "Article" and n.get("attr", {}).get("Num") == "117":
                return n
            for c in n.get("children", []):
                r = find(c)
                if r:
                    return r
        elif isinstance(n, list):
            for c in n:
                r = find(c)
                if r:
                    return r
    a = find(j["law_full_text"])
    L = [f"介護保険法第117条（e-Gov法令検索 {j['revision_info']['law_revision_id']}）", ""]
    for c in a["children"]:
        if isinstance(c, dict) and c.get("attr", {}).get("Num") in ("12", "13", "14"):
            L.append(text(c))
    open(path, "w", encoding="utf-8").write("\n".join(L) + "\n")


def main():
    work = tempfile.mkdtemp(prefix="redteam_")
    root = os.path.join(work, f"川崎町_RedTeamレビュー対応_{HI}")
    d_seika, d_taiou, d_ev, d_sc = (os.path.join(root, x) for x in
                                    ("01_成果物", "02_レビュー対応", "03_エビデンス", "04_スクリプト"))
    d_log = os.path.join(d_ev, "点検ログ")
    for x in (d_seika, d_taiou, d_ev, d_sc, d_log):
        os.makedirs(x)
    ng = []
    print(f"Red Team レビュー対応の一式（{HI}）")

    print("  ── 1 成果物と PDF")
    ichiran = []
    for kubun, f, kyu in SEIKA:
        shutil.copy2(f, d_seika)
        pdf, n = to_pdf(f, d_seika)
        ichiran.append((kubun, os.path.basename(f), n, md5(f), os.path.basename(kyu)))
        print(f"   ○ {kubun} {os.path.basename(f)}（PDF {n}頁）")
        if not n:
            ng.append(f"PDF {f}")

    print("  ── 2 点検ログ")
    logs = []
    for na, cmd in TENKEN:
        r = subprocess.run([sys.executable] + cmd, capture_output=True, text=True, timeout=900)
        out = "\n".join(x for x in (r.stdout + r.stderr).splitlines()
                        if "FutureWarning" not in x and "warnings.warn" not in x)
        open(os.path.join(d_log, f"{na}.txt"), "w", encoding="utf-8").write(
            f"$ python3 {' '.join(cmd)}\n終了コード {r.returncode}\n\n{out}\n")
        logs.append((na, r.returncode))
        print(("   ○ " if r.returncode == 0 else "   × ") + f"{na}（終了コード {r.returncode}）")
        if r.returncode:
            ng.append(na)
    # 目次の頁番号（A09）はファイルを書き換えるため写しに当てる
    tmp = os.path.join(work, "toc.docx")
    shutil.copy2(S, tmp)
    r = subprocess.run([sys.executable, SC + "fill_toc_pages_v200.py", tmp], capture_output=True, text=True, timeout=900)
    toc = docx.Document(tmp).tables[0].rows[0].cells
    open(os.path.join(d_log, "15_目次の頁番号（写しに当てた）.txt"), "w", encoding="utf-8").write(
        r.stdout + f"\n目次の先頭行：{[c.text for c in toc]}\n")
    print(f"   ○ 15_目次の頁番号（ご挨拶 {toc[2].text}頁）")

    print("  ── 3 検算・原票抜粋・法令・旧表現の残存検索・差分")
    open(os.path.join(d_ev, "検算.md"), "w", encoding="utf-8").write(kensan())
    sc_genpyo(os.path.join(d_ev, "原票抜粋_生活支援コーディネーター.csv"))
    jobun_117(os.path.join(d_ev, "原典抜粋_介護保険法第117条第12〜14項.txt"))
    shutil.copy2("09_元資料/その他データ/公表システム_仙南圏域_在宅系地域密着型_R8.6末.csv", d_ev)
    zan = ["語\t" + "\t".join(k for k, _, _ in SEIKA)]
    zan_ng = 0
    texts = {k: "\n".join(zenbun(f)) for k, f, _ in SEIKA}
    for w in ZANSON:
        c = [texts[k].count(w) for k, _, _ in SEIKA]
        zan.append(w + "\t" + "\t".join(str(x) for x in c))
        zan_ng += sum(c)
    open(os.path.join(d_ev, "旧表現の残存検索.txt"), "w", encoding="utf-8").write(
        "レビューで挙がった旧の表現が、提出の5冊に残っていないか（件数）\n\n" + "\n".join(zan) +
        f"\n\n合計 {zan_ng}件\n")
    print(("   ○ " if zan_ng == 0 else "   × ") + f"旧表現の残存 {zan_ng}件（{len(ZANSON)}語×5冊）")
    if zan_ng:
        ng.append("旧表現の残存")
    for kubun, f, kyu in SEIKA:
        if not os.path.exists(kyu):
            continue
        a, b = zenbun(kyu), zenbun(f)
        diff = list(difflib.unified_diff(a, b, os.path.basename(kyu), os.path.basename(f), n=0, lineterm=""))
        open(os.path.join(d_ev, f"差分_{kubun}.txt"), "w", encoding="utf-8").write(
            f"{kubun}：{os.path.basename(kyu)}（レビュー対象又は前の版）→ {os.path.basename(f)}\n"
            "段落・表のセルの単位の差分（-＝旧、+＝新）\n\n" + "\n".join(diff) + "\n")
        print(f"   ○ 差分_{kubun}（{sum(1 for x in diff if x.startswith('+') and not x.startswith('+++'))}行が新）")

    print("  ── 4 対応表（md・xlsx）")
    md = [f"# Red Team レビュー（令和8年10月24日）への対応表", "",
          f"> レビュー原本：`{REVIEW}`（素案 v2.29・概要版 v2.30・資料 v16・問答集 v13 が対象）",
          "> 対応後の版：協議用素案 Ver.2.32／公表用 v2.32／概要版 v2.32／委員会資料 v17／想定問答集 v14",
          "> 本表の正本は `07_ソーススクリプト/build_redteam_teishutsu_R8.10.24.py` の TAIOU", "",
          "| ID | 判定 | 指摘 | 対応 | 対応の内容 | 箇所 | エビデンス | 残る事項 |", "|---|---|---|---|---|---|---|---|"]
    for row in TAIOU:
        md.append("| " + " | ".join(x.replace("|", "／") for x in row) + " |")
    import collections
    c = collections.Counter(r[3] for r in TAIOU)
    md += ["", "対応の区分：" + "／".join(f"{k} {v}件" for k, v in c.items()), "",
           "## C　委員会・町への確認事項について", "",
           "C-1 の6項目は確認事項No.189（【町確認】2件）・No.127・No.154・No.159・No.137 で台帳に載っています。"
           "回答がない場合に文を落とすのは**資料の文言の扱い**に限り、台帳の確認事項は閉じません"
           "（供給・認知症の実施状況の欠けは、第3回までの確定事項として残します）。",
           "C-2 の3群（施策体系・基金の取崩し・回数とパブリックコメント・頁数）は、"
           "`川崎町_第2回策定委員会資料の確定に向けた点検_R8.10.23.md` §3 でご了承をお願いしています。", ""]
    open(MD_OUT, "w", encoding="utf-8").write("\n".join(md))
    shutil.copy2(MD_OUT, d_taiou)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "対応表"
    head = ["ID", "判定", "指摘", "対応", "対応の内容", "箇所", "エビデンス", "残る事項"]
    ws.append(head)
    for row in TAIOU:
        ws.append(list(row))
    for i, w in enumerate((6, 8, 40, 12, 70, 28, 30, 26), 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    for r in ws.iter_rows():
        for cc in r:
            cc.alignment = Alignment(wrap_text=True, vertical="top")
            cc.font = Font(name="游ゴシック", size=9, bold=(cc.row == 1))
            if cc.row == 1:
                cc.fill = PatternFill("solid", fgColor="D9E1F2")
    ws.freeze_panes = "A2"
    os.makedirs(os.path.dirname(XLSX_OUT), exist_ok=True)
    wb.save(XLSX_OUT)
    shutil.copy2(XLSX_OUT, d_taiou)
    shutil.copy2(REVIEW, d_taiou)
    print(f"   ○ {len(TAIOU)}件（" + "／".join(f"{k}{v}" for k, v in c.items()) + "）")

    print("  ── 5 スクリプト")
    for f in SCRIPTS:
        shutil.copy2(SC + f, d_sc)

    print("  ── 6 README（成果品一覧・MD5）")
    rd = [f"# 川崎町 第10期計画　Red Team レビュー対応の一式（{HI}）", "",
          "川崎町（宮城県柴田郡）高齢者保健福祉計画・第10期介護保険事業計画", "",
          "## 中身", "",
          "| フォルダ | 内容 |", "|---|---|",
          "| 01_成果物 | 是正後の5冊（docx）と、当方の環境で作った PDF |",
          "| 02_レビュー対応 | 対応表（md・xlsx）とレビュー原本 |",
          "| 03_エビデンス | 点検ログ（各スクリプトの出力と終了コード）・検算・旧表現の残存検索・差分・原票と法令の抜粋 |",
          "| 04_スクリプト | 是正と点検に用いたスクリプト（リポジトリの 07_ソーススクリプト と同じもの） |", "",
          "## 成果品一覧", "",
          "| 区分 | ファイル | PDF頁数（当方の環境） | MD5 | 比べた旧の版 |", "|---|---|---|---|---|"]
    for kubun, f, n, h, kyu in ichiran:
        rd.append(f"| {kubun} | {f} | {n} | `{h}` | {kyu} |")
    rd += ["", "⚠ 頁数は当方の環境（HG丸ｺﾞｼｯｸM-PRO・游ゴシックを IPAゴシックで代用した LibreOffice）による。"
           "Word と町のパソコンの書体では変わり得るため、配布用の PDF は町の環境で作り、頁送りを確かめていただく（A08）。",
           "", "## 点検ログ", "", "| 点検 | 終了コード |", "|---|---|"]
    for na, rc in logs:
        rd.append(f"| {na} | {rc} |")
    rd += ["", "## 再現の手順", "", "リポジトリの `kawasaki_project/` で次を実行すると、同じ一式ができます。", "",
           "```", "python3 07_ソーススクリプト/fix_soan_v2320_redteam.py", "python3 07_ソーススクリプト/build_kouhyou_soan.py",
           "python3 07_ソーススクリプト/build_gaiyo_v230.py", "python3 07_ソーススクリプト/build_2kai_v17.py",
           "python3 07_ソーススクリプト/build_2kai_qa_v14.py", "python3 07_ソーススクリプト/build_redteam_teishutsu_R8.10.24.py", "```", ""]
    open(os.path.join(root, "00_README.md"), "w", encoding="utf-8").write("\n".join(rd))

    print("  ── 7 ZIP")
    os.makedirs(os.path.dirname(ZIPNAME), exist_ok=True)
    with zipfile.ZipFile(ZIPNAME, "w", zipfile.ZIP_DEFLATED) as z:
        for dp, _, fs in os.walk(root):
            for f in sorted(fs):
                full = os.path.join(dp, f)
                z.write(full, os.path.relpath(full, work))
    with zipfile.ZipFile(ZIPNAME) as z:
        names = z.namelist()
        ok = all(hashlib.md5(z.read(f"{os.path.basename(root)}/01_成果物/{f}")).hexdigest() == h
                 for _, f, _, h, _ in ichiran)
    print(("   ○ " if ok else "   × ") + f"{ZIPNAME}（{len(names)}ファイル・{os.path.getsize(ZIPNAME):,}バイト）"
          "・成果物の MD5 が一覧と一致")
    if not ok:
        ng.append("MD5")
    shutil.rmtree(work)
    if ng:
        print(f"   × 不適合 {len(ng)}件：" + "／".join(ng))
        sys.exit(1)
    print("   ○ 不適合0件")


if __name__ == "__main__":
    main()
