# -*- coding: utf-8 -*-
"""Red Team 再レビュー第3回（令和8年10月24日）：影響範囲の再点検と是正の一式を ZIP にする

  python3 07_ソーススクリプト/build_redteam_r3eikyou_teishutsu_R8.10.24.py

ご指示「レビュー結果の確認と修正進めて下さい。先程同様影響範囲確認の上、成果物とエビデンスをまとめてアウトプットして下さい」
（令和8年10月24日）。受領したレビューは第3回と同じもの（MD5 一致）であったため、v2.34 での是正を前提に、指摘の語だけでなく
同じ趣旨の記述を5冊で走査し（影響範囲の走査）、残っていたものを素案 Ver.2.35・資料 v20 で是正した。

ご指示「再レビュー結果確認後、修正をお願い致します。成果物とエビデンスをジップにまとめてアウトプットして」
（令和8年10月24日・第3回）による。組み立ては前回（build_redteam_saireview_teishutsu_R8.10.24.py）と同じ。
対応表の正本は本スクリプトの TAIOU・SAIHANTEI（md と xlsx はここから書き出す）。
点検のログ・検算・差分・PDF は、その場で成果物と原票から作る（手で書かない）。

出力
  11_提出一式/川崎町_RedTeam再レビュー第3回対応_成果物とエビデンス_R8.10.24.zip
  01_第10期_最新版成果品/川崎町_RedTeam再レビュー第3回対応表_R8.10.24.md
  05_試算・管理シート/川崎町_RedTeam再レビュー第3回対応表_R8.10.24.xlsx
終了コード　0＝点検ログがすべて終了コード0・検算が成果物と原票に一致・旧表現0件・ZIP の MD5 が一覧と一致／1＝それ以外
"""
import collections
import csv
import difflib
import hashlib
import importlib.util
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

HERE = os.path.dirname(os.path.abspath(__file__))


def load(name, fn):
    spec = importlib.util.spec_from_file_location(name, os.path.join(HERE, fn))
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


RT1 = load("rt1", "build_redteam_teishutsu_R8.10.24.py")            # md5・zenbun・to_pdf・sc_genpyo
RT2 = load("rt2", "build_redteam_saireview_teishutsu_R8.10.24.py")  # 検算 §1・§2（報酬改定・入所・居住系）
MIKOMI = load("mikomi", "check_mikomi_morasi_R8.10.16.py")         # 素案 9-1 の表の読み取り
NINCHI = load("ninchi", "calc_ninchisho_R8.10.19.py")              # 認知症の人の数の算定の正本

HI = "R8.10.24"
ZIPNAME = f"11_提出一式/川崎町_RedTeam再レビュー第3回_影響範囲再点検_成果物とエビデンス_{HI}.zip"
MD_OUT = f"01_第10期_最新版成果品/川崎町_RedTeam再レビュー第3回_影響範囲再点検の対応表_{HI}.md"
XLSX_OUT = f"05_試算・管理シート/川崎町_RedTeam再レビュー第3回_影響範囲再点検の対応表_{HI}.xlsx"
REVIEW = "09_元資料/レビュー/川崎町_第10期計画_RedTeam再レビュー第3回_R8.10.24受領.md"
NENPO = "09_元資料/R7実績データ/年報データ_2025_川崎町.xlsx"
RYUHO = "05_試算・管理シート/川崎町_公表用の留保一覧_v2.35.xlsx"
OTOSHI = "05_試算・管理シート/川崎町_公表用で落とした文_v2.35.xlsx"

SEIKA = [  # (区分, ファイル, 第3回レビューの対象の版)　差分は第3回レビューの対象の版から
    ("協議用素案", "01_第10期_最新版成果品/川崎町_計画書素案_v2.35_協議用.docx",
     "01_第10期_最新版成果品/川崎町_計画書素案_v2.33_協議用.docx"),
    ("公表用の計画書（案）", "01_第10期_最新版成果品/川崎町_計画書案_公表用_v2.35.docx",
     "01_第10期_最新版成果品/川崎町_計画書案_公表用_v2.33.docx"),
    ("概要版", "01_第10期_最新版成果品/川崎町_計画書案_概要版_v2.35.docx",
     "01_第10期_最新版成果品/川崎町_計画書案_概要版_v2.33.docx"),
    ("第2回策定委員会資料", "03_委員会・説明資料/川崎町_第2回策定委員会資料_R8.11_v20.docx",
     "03_委員会・説明資料/川崎町_第2回策定委員会資料_R8.11_v18.docx"),
    ("想定問答集", "03_委員会・説明資料/川崎町_第2回策定委員会_想定問答集_R8.11_v16.docx",
     "03_委員会・説明資料/川崎町_第2回策定委員会_想定問答集_R8.11_v15.docx"),
]
S = SEIKA[0][1]
P = SEIKA[1][1]
SC = "07_ソーススクリプト/"
TENKEN = [  # (名前, コマンド)
    ("01_素案_自己点検", [SC + "check_soan_selfcheck_R8.9.28.py", S]),
    ("02_素案_保険料の検算", [SC + "check_hokenryo_R8.9.30.py", S]),
    ("03_素案_書体と紙面", [SC + "check_font_R8.9.30.py", S]),
    ("04_素案_8軸の網羅性（記載の網羅性。未実施10事項を含む）", [SC + "check_8jiku_R8.10.4.py", S]),
    ("05_素案_章節の参照の機械検索", [SC + "check_sansho_R8.10.24.py", S]),
    ("06_図表データ管理台帳（素案v2.35）", [SC + "build_zuhyo_daicho.py"]),
    ("07_事業所の受入れ様式（様式の検査。実データの全数照合ではない）", [SC + "check_jigyosho_uke_R8.10.15.py"]),
    ("08_見込量の漏れ", [SC + "check_mikomi_morasi_R8.10.16.py"]),
    ("09_認知症の人の数の整合", [SC + "check_ninchisho_seigo_R8.10.19.py"]),
    ("10_認定者数", [SC + "check_nintei_R8.10.16.py"]),
    ("11_公表用_書体と紙面", [SC + "check_font_R8.9.30.py", P]),
    ("12_委員会資料_自己点検", [SC + "check_2kai_selfcheck_R8.9.28.py"]),
    ("13_守るべき制約の走査", [SC + "check_seiyaku_R8.10.6.py"]),
    ("14_リポジトリの健全性", [SC + "check_repo_kenzensei_R8.10.4.py"]),
]
SCRIPTS = ["fix_soan_v2340_r3review.py", "fix_soan_v2350_eikyou.py", "build_kouhyou_soan.py", "build_gaiyo_v230.py",
           "build_2kai_v19.py", "build_2kai_v20.py",
           "build_2kai_qa_v16.py", "calc_ninchisho_R8.10.19.py", "check_mikomi_morasi_R8.10.16.py",
           "build_redteam_saireview_teishutsu_R8.10.24.py", "build_redteam_r3_teishutsu_R8.10.24.py",
           "build_redteam_r3eikyou_teishutsu_R8.10.24.py"]

# ══════════════════════════ 対応表（正本）
TAIOU = [
    ("R3-01", "高", "計画比の乖離から「在宅から施設・居住系への移行が進んだ」と断定（素案39頁・資料18頁・QB-44）",
     "是正済", "「居宅は計画を下回り施設は上回った。供給体制や町外施設の利用が影響した可能性はあるが、移行の実態及び原因は"
     "確かめられていない（町内外別の利用者数・経年の推移等の追加の検証を要する）」に改めた。見込量の前提は「令和7年度の"
     "利用実績を出発点としているため、この利用の構成が続くこと」と書き分けた",
     "素案 3-3／公表用／資料 v19・v20 ８／問答集 v16 QB-44", "03_エビデンス/旧表現の残存検索.txt", "町内外別の利用の推移（No.159）"),
    ("R3-02", "高", "推計値374人と自立度Ⅱ以上の実数の「差が認定を受けていない認知症の方の規模」と断定（素案68頁）",
     "是正済", "「両者は把握の方法と対象が異なるため、単純な差を要介護認定を受けておられない認知症の方の数と読むことは"
     "できない。地域の状況を検討する参考として並べて掲げる」に改めた", "素案 6-1／公表用",
     "03_エビデンス/差分_協議用素案.txt", "未認定の方の把握は別の調査の設計による（No.157）"),
    ("R3-03", "高", "公表用の留保台帳103行の「公表の前の扱い」がすべて空欄",
     "一部是正", "台帳に確認先・処理の期限・既定の扱い（回答・決定がない場合）を103件すべてに置き、協議用素案の同じ段落に"
     "ある確認事項No を引いた。処理結果・決裁日は町が記入する列とし「未決定」と明示した。各件の確定・注記・取下げの"
     "決定は町・委員会による。公表用の表紙は留保が残る間「体裁確認用」のまま",
     "05_試算・管理シート/川崎町_公表用の留保一覧_v2.35.xlsx", "03_エビデンス/川崎町_公表用の留保一覧_v2.35.xlsx",
     "103件の処理の決定（公表の確定の条件）"),
    ("R3-04", "高", "資料19頁「第13条第3項が準用する第12条第4〜7項」。意見聴取は第12条第3項",
     "是正済", "「第12条第3項から第7項まで」に改めた。素案 10-4 の「第12条第5項」「第12条第6項」も、市町村計画に当たる"
     "「第13条第3項において準用する第12条第5項・第6項」に揃えた（素案 第6章の整理と同じ）",
     "資料 v19・v20 表（第10章の行）／素案 10-4", "03_エビデンス/差分", ""),
    ("R3-05", "高", "全3回とする根拠に委託仕様書11（定めのない事項の協議）を援用（資料21〜22頁・QB-58）",
     "是正済（手続は町）", "6(3)に定めのある「4回開催の予定」の変更であるため、町の契約担当で契約上の手続（変更契約又は書面による"
     "合意の要否）を確かめ、町が書面で決定する形に改めた。委員会には工程の見通しについてご意見をいただく（⑦の文言も改めた）",
     "資料 v19・v20 ８・本日のご審議事項⑦／問答集 v16 QB-58", "03_エビデンス/差分",
     "契約書・仕様書の全文による手続の判断は町の契約担当（確認事項No.49）"),
    ("R3-06", "中〜高", "「第10期の算定（令和7年度の年報による確定値）」は保険料が確定したと読める",
     "是正済", "見出しを「第１０期保険料の現時点の試算（令和７年度の介護保険事業状況報告等を基礎に算定。基金取崩額・令和９年度の"
     "報酬改定等は未確定）」に、本文を「年報（確定値）を基礎として算定。保険料は要因が定まるまでは試算」に改めた",
     "資料 v19・v20 ６", "03_エビデンス/旧表現の残存検索.txt", ""),
    ("R3-07", "中", "資料21頁冒頭が前頁からの文の続き（頁数の見通しの囲みが分割）",
     "是正済", "頁数の見通しの囲みを行で分割しない設定とし、21頁の冒頭から囲み全体が載るようにした。22頁は構成の案・審議事項・"
     "次回予定（当方の環境で22頁）", "資料 v19・v20", "01_成果物/川崎町_第2回策定委員会資料_R8.11_v20.pdf",
     "配布用 PDF は町の環境（Word）で頁送りを確かめていただく"),
    ("R3-08", "中", "通いの場参加率15.9％（n=439）は一般高齢者調査の回答者のうちの割合。分母が不明瞭",
     "是正済", "KPI の名称を「通いの場への参加割合（一般高齢者調査の回答者。65歳以上）」とし、出典の欄に「通いの場の参加者数÷"
     "65歳以上人口による参加率は、地域支援事業のデータの受領後に別の指標として置く」と書いた。5-4 の本文・表も同じ",
     "素案 5-4・10-4", "03_エビデンス/差分_協議用素案.txt", "地域支援事業のデータ（未受領）"),
    ("R3-09", "中", "認知症の人の数の年齢階級別の和372人と計374人",
     "是正済", "算定の正本（calc_ninchisho）で確かめた：各行の小数は7.25・23.45・52.49・58.20・115.09・117.46人、計373.93人。"
     "注を「各年齢階級は1人未満を四捨五入して表示し、計は小数のまま合計してから四捨五入。内訳の和（372人）と計（374人）が"
     "2人一致しない」と具体化した", "素案 6-1", "03_エビデンス/検算.md §4", ""),
    ("R3-10", "高・未確認", "見込量の照合は32サービスの一致まで。原票との突合・需要供給3ケースが未完",
     "一部是正", "令和7年度の37サービスの給付費を年報（様式２ 給付費）の円単位の値と突合し、37件すべてが千円未満四捨五入で"
     "素案 9-1 の値に一致し、円単位の和が総計998,222,421円に一致することを確かめた（検算 §3）。将来推計総括表・算定ブック・"
     "交付金の回答原票との突合、需要・供給の3ケース比較は第3回に向けて行う",
     "素案 9-1", "03_エビデンス/検算.md §3・突合_令和7年度給付費_年報様式2.csv", "需要・供給3ケース（No.159・147）"),
    ("R3-06 波及", "中", "（影響範囲の再点検で検出）素案 9-3 の「本算定」5か所・資料 表26 の「本算定」",
     "是正済", "保険料の「本算定」は確定の算定と読めるため、R3-06 と同じ趣旨で「本試算」に改めた（素案 Ver.2.35・資料 v20）",
     "素案 9-3／資料 v20 表26", "03_エビデンス/影響範囲の走査.txt", ""),
    ("R3-08 波及", "中", "（影響範囲の再点検で検出）KPI 以外の箇所の「通いの場への参加率」「参加率は15.9％」",
     "是正済", "15.9％と同じ指標を指す箇所を「参加割合（調査の回答者）」に揃えた。KPI 候補の表は「参加割合（調査）・参加者数"
     "（町実績）」と分けた（素案 Ver.2.35）", "素案 2-6・10-4", "03_エビデンス/影響範囲の走査.txt", ""),
    ("1-2 留保", "低〜中", "入所・居住系±10人の前提「29.9万円/人月」は概数",
     "是正済", "出所を素案 5-5 に書いた：令和7年度の入所・居住系の給付費680,226千円 ÷ 年報の受給者数 延べ2,274人月"
     "（施設1,612・認知症対応型共同生活介護311・地域密着型介護老人福祉施設319・特定施設32）＝299.13千円。丸める前の値でも"
     "10人で211.84円となり約212円は変わらない", "素案 5-5", "03_エビデンス/検算.md §6", ""),
    ("R3-11", "高・未確認", "計画118頁＋資料編13頁＝131頁。仕様書の「100頁程度」の対象範囲",
     "町のご判断", "資料８でご判断をお願いしている。確認事項No.83 に、対象範囲（本文・資料編・概要版）を町と書面で確かめる"
     "ことを追記", "資料 v19・v20 ８", "―", "確認事項No.83"),
]
SAIHANTEI = [  # 第3回レビュー 1-1 の再判定のうち、今回手を入れたもの
    ("N07", "概ね解消", "R3-07 のとおり頁数の見通しの囲みを分割しない"),
    ("N10", "一部是正・再開", "R3-01 のとおり素案・資料・問答集の断定を改めた"),
    ("N11", "一部是正", "R3-10 のとおり37サービスを年報の円単位で突合し、2千円の差が丸めであることを確かめた"),
    ("N05", "一部是正", "R3-03 のとおり留保の台帳に扱いの列を置いた。処理の決定は町"),
    ("N09", "解消（前提の29.9万円は概数）", "1-2 留保 のとおり29.9万円の出所を年報から示し、丸める前の値でも212円であることを確かめた"),
    ("N03", "未確定（適切に留保）", "変更なし。確認事項No.190"),
    ("N01・N02・N04・N06・N08・N09", "解消", "変更なし"),
]
ZANSON = [  # 旧表現の残存検索（第3回と前回までのレビューで挙がった語）
    "移行が進んでいます", "この移行が続く", "この移行を前提", "おおよその規模に当たります", "第12条第4〜7項",
    "（認知症基本法第12条第5項）", "認知症基本法第12条第6項により", "年報による確定値", "算定により確定しました",
    "委託仕様書11（本仕様の定めのない", "仕様書11は", "通いの場参加率（65歳以上）", "通いの場への参加率（65歳以上）",
    "本算定", "参加率は15.9％", "通いの場への参加率、", "通いの場参加率／",
    # 前回まで
    "6,235円", "5,583円", "6,366円", "5,714円", "約211円", "＋211円", "▲211円", "差がこれと異なります",
    "要件となります", "残高の確定", "結果と考えられます", "1,741保険者", "35保険者", "285保険者", "保険者中",
    "998,220", "5,959", "6,382", "25年間", "11つ", "8の事業所", "要件を満たします", "概ね半分以下",
    "生活支援コーディネーター25", "9-4の感度表", "町内14の指定", "第117条第５項",
]


def nenpo_totsugo(ng, csv_path):
    """令和7年度の給付費：素案 9-1 の37サービスの千円の値を、年報 様式２（給付費）の円単位の値と突き合わせる。"""
    so = MIKOMI.yomu_soan(S)
    ws = openpyxl.load_workbook(NENPO, data_only=True)["様式２（給付費）"]
    GROUP = {11, 12, 18, 21, 26, 33, 47, 52}
    rows = {}
    for r in range(13, 52):
        if r in GROUP:
            continue
        nm = next((ws.cell(r, c).value for c in (5, 4, 3) if ws.cell(r, c).value), None)
        if nm:
            rows[str(nm).strip()] = r
    soukei = ws.cell(52, 16).value

    def yen(na):
        if na == "介護予防支援":
            return ws.cell(rows["介護予防支援・居宅介護支援"], 8).value, "介護予防支援・居宅介護支援 H（予防）"
        if na == "居宅介護支援":
            return ws.cell(rows["介護予防支援・居宅介護支援"], 15).value, "介護予防支援・居宅介護支援 O（介護）"
        col, kubun = 15, "O（介護）"
        base = na
        if "（介護予防を含む）" in na:
            base, col, kubun = na.replace("（介護予防を含む）", ""), 16, "P（合計）"
        elif na.startswith("特定介護予防"):
            base, col, kubun = na[len("特定介護予防"):], 8, "H（予防）"
            base = "特定" + base
        elif na.startswith("介護予防"):
            base, col, kubun = na[len("介護予防"):], 8, "H（予防）"
        base = base.replace("特定福祉用具購入費", "特定福祉用具販売")
        hit = [n for n in rows if n == base or n.startswith(base + "（") or n.startswith("複合型サービス(" + base)]
        if not hit:
            return None, "年報に該当の行なし"
        return sum(ws.cell(rows[n], col).value or 0 for n in hit), "＋".join(hit) + " " + kubun

    out = [["サービス（素案 9-1）", "素案の千円", "年報の円", "年報の行・列", "円÷1000 を四捨五入", "一致"]]
    wa_en, wa_sen, itchi = 0, 0, 0
    for na, v in so.items():
        y, w = yen(na)
        ok = y is not None and round(y / 1000) == round(v)
        itchi += ok
        wa_en += y or 0
        wa_sen += round(v)
        out.append([na, f"{round(v):,}", f"{y:,}" if y is not None else "", w,
                    f"{round(y / 1000):,}" if y is not None else "", "○" if ok else "×"])
    out.append(["計", f"{wa_sen:,}", f"{wa_en:,}", f"総計 {soukei:,}円", f"{round(wa_en / 1000):,}", ""])
    with open(csv_path, "w", encoding="utf-8-sig", newline="") as fp:
        csv.writer(fp).writerows(out)
    ok_all = itchi == len(so) and wa_en == soukei
    if not ok_all:
        ng.append("検算§3 年報との突合")
    return [f"## §3 令和7年度の給付費：素案 9-1 の37サービスと年報（様式２ 給付費）の円単位の突合（N11・R3-10）", "",
            f"- 原票：`{NENPO}` の「様式２（給付費）」（介護給付は O 列・予防給付は H 列・介護予防を含むものは P 列）",
            f"- 37サービスのうち、年報の円を千円未満四捨五入した値が素案と一致：{itchi}件／{len(so)}件",
            f"- 37サービスの円の和：{wa_en:,}円（年報の総計 {soukei:,}円と{'一致' if wa_en == soukei else '不一致'}）",
            f"- 素案の千円の和 {wa_sen:,}千円と、総計を丸めた {round(soukei / 1000):,}千円の差 {wa_sen - round(soukei / 1000):+,}千円は、"
            "37の値を別々に丸めたことによる差である（円の和が総計と一致するため、漏れ・重複ではない）",
            "- 明細は 突合_令和7年度給付費_年報様式2.csv", ""], ok_all


EIKYOU = [  # (指摘, 同じ趣旨の記述を拾う形, 残っていてはならない形)
    ("R3-01 施設への移行の断定", r"施設[・等]?[^。]{0,10}(への)?移行|在宅から施設|施設へ移|常態化が", r"移行が進ん|移行が続く|この移行を前提"),
    ("R3-02 推計値と実数の差", r"未認定|認定を受けて(い|おられ)ない", r"おおよその規模"),
    ("R3-04 認知症基本法の条項", r"第12条第[3-7]項|第12条第[3-7]〜", r"第12条第4〜7項|（認知症基本法第12条第[5-6]項）"),
    ("R3-05 委員会回数の根拠", r"仕様書11|仕様書１１|定めのない事項", r"仕様書11|仕様書１１"),
    ("R3-06 保険料が確定と読める語", r"確定値|本算定|本試算|確定しました", r"年報による確定値|本算定|算定により確定しました"),
    ("R3-08 通いの場の参加率の分母", r"参加率|参加割合", r"通いの場参加率（65歳以上）|通いの場への参加率（65歳以上）|参加率は15\.9|実際の参加率が15\.9"),
    ("R3-09 認知症の人の数の内訳と計", r"372|374人", None),
]


def eikyou_sousa(ng, path):
    """第3回レビューの指摘ごとに、同じ趣旨の記述を5冊で拾い、残っていてはならない形が0件であることを確かめる。"""
    L = ["影響範囲の走査（第3回レビューの指摘と同じ趣旨の記述を、提出の5冊の本文・表から拾ったもの）", "",
         "各行：［冊］前後の文。「残っていてはならない形」が0件であることを確かめ、拾った記述は是正後の文言として目で確かめる。", ""]
    texts = {}
    for k, f, _ in SEIKA:
        texts[k] = list(dict.fromkeys(RT1.zenbun(f)))
    kei_ng = 0
    for mi, hiro, dame in EIKYOU:
        L += [f"## {mi}", f"拾う形：{hiro}", f"残っていてはならない形：{dame or '（なし）'}"]
        for k, _, _ in SEIKA:
            hits = [s_ for s_ in texts[k] if re.search(hiro, s_)]
            n_dame = sum(len(re.findall(dame, s_)) for s_ in texts[k]) if dame else 0
            kei_ng += n_dame
            L.append(f"［{k}］拾った記述 {len(hits)}件・残っていてはならない形 {n_dame}件")
            for s_ in hits:
                m = re.search(hiro, s_)
                L.append("    …" + s_[max(0, m.start() - 50):m.end() + 50].replace("\n", " ") + "…")
        L.append("")
    L.append(f"残っていてはならない形の合計 {kei_ng}件")
    open(path, "w", encoding="utf-8").write("\n".join(L) + "\n")
    if kei_ng:
        ng.append("影響範囲の走査")
    return kei_ng


def kensan(ng, d_ev):
    RT2.S = S
    RT2.SEIKA = SEIKA
    moto = RT2.kensan(ng)
    head = moto.split("## §3")[0]
    L = [head.rstrip(), ""]
    sec3, _ = nenpo_totsugo(ng, os.path.join(d_ev, "突合_令和7年度給付費_年報様式2.csv"))
    L += sec3
    r = NINCHI.santei(NINCHI.machi_kaikyu(), "認知症")
    gyo = [m[6] for m in r["meisai"]]
    wa, kei = sum(round(x) for x in gyo), round(r["sei"])
    sz = "\n".join(RT1.zenbun(S))
    ok4 = f"内訳の和（{wa}人）と計（{kei}人）" in sz
    if not ok4:
        ng.append("検算§4 認知症の注")
    L += ["## §4 認知症の人の数の内訳と計（R3-09）", "",
          "- 算定の正本 `07_ソーススクリプト/calc_ninchisho_R8.10.19.py`（性・年齢階級別の有病率×町の性・年齢階級別人口）",
          "- 年齢階級別（小数）：" + "・".join(f"{x:.2f}" for x in gyo) + f"人、計 {r['sei']:.2f}人",
          f"- 各行を四捨五入した和 {wa}人／計を四捨五入 {kei}人（差 {kei - wa}人）",
          f"- 素案 6-1 の注にこの説明があること：{'○' if ok4 else '×'}", "",
          "## §5 生活支援コーディネーター（N03）", "",
          "- 原票抜粋は 原票抜粋_生活支援コーディネーター.csv。確認事項No.190（町のご確認待ち）", ""]
    sz = "\n".join(RT1.zenbun(S))
    m = re.search(r"給付費([\d,]+)千円÷介護保険事業状況報告の受給者数延べ([\d,]+)人月", sz)
    kyufu, nin = int(m.group(1).replace(",", "")), int(m.group(2).replace(",", ""))
    tanka = kyufu / nin
    rates = [float(x) / 100 for x in re.findall(r"令和\d+年度([\d.]+)％",
                                                [x for x in RT1.zenbun(S) if "見込交付割合は全国の仮置値" in x][0])]
    dm = sum(10 * tanka * 12 * (0.23 + 0.05 - e) for e in rates) * 1000 / 0.96 / 9743.1 / 12
    ok6 = round(dm) == 212 and "10人では約212円" in sz
    if not ok6:
        ng.append("検算§6 29.9万円")
    L += ["## §6 入所・居住系の1人1月あたり給付費（1-2 の留保「29.9万円は概数」）", "",
          f"- 素案 5-5：令和7年度の入所・居住系の給付費 {kyufu:,}千円 ÷ 年報の受給者数 延べ{nin:,}人月 ＝ {tanka:.2f}千円（約{tanka / 10:.1f}万円）",
          "- 受給者数は 様式１の６（15）施設介護サービス受給者数の総数、様式１の７（18）認知症対応型共同生活介護（短期利用以外）・"
          "地域密着型介護老人福祉施設入所者生活介護、様式１の７（16）特定施設入居者生活介護（短期利用以外）の合計（年度の延べ人月）",
          f"- 丸める前の {tanka:.2f}千円で10人の月額を計算すると {dm:.2f}円 → 約{round(dm)}円（29.9万円による211.75円と同じく212円）",
          f"- 素案 5-5 の「約212円」と一致：{'○' if ok6 else '×'}", ""]
    return "\n".join(L)


def main():
    work = tempfile.mkdtemp(prefix="redteam3_")
    root = os.path.join(work, f"川崎町_RedTeam再レビュー第3回_影響範囲再点検_{HI}")
    d_seika, d_taiou, d_ev, d_sc = (os.path.join(root, x) for x in
                                    ("01_成果物", "02_レビュー対応", "03_エビデンス", "04_スクリプト"))
    d_log = os.path.join(d_ev, "点検ログ")
    for x in (d_seika, d_taiou, d_ev, d_sc, d_log):
        os.makedirs(x)
    ng = []
    print(f"Red Team 再レビュー第3回：影響範囲の再点検の一式（{HI}）")

    print("  ── 1 成果物と PDF")
    ichiran = []
    for kubun, f, kyu in SEIKA:
        shutil.copy2(f, d_seika)
        pdf, n = RT1.to_pdf(f, d_seika)
        ichiran.append((kubun, os.path.basename(f), n, RT1.md5(f), os.path.basename(kyu)))
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

    print("  ── 3 検算・原票との突合・留保一覧・旧表現の残存検索・差分")
    n_e = eikyou_sousa(ng, os.path.join(d_ev, "影響範囲の走査.txt"))
    print(("   ○ " if n_e == 0 else "   × ") + f"影響範囲の走査（{len(EIKYOU)}指摘×5冊）残っていてはならない形 {n_e}件")
    open(os.path.join(d_ev, "検算.md"), "w", encoding="utf-8").write(kensan(ng, d_ev))
    print(("   ○ " if not [x for x in ng if x.startswith("検算")] else "   × ") + "検算 §1〜§6（成果物・年報・算定の正本と突合）")
    RT1.sc_genpyo(os.path.join(d_ev, "原票抜粋_生活支援コーディネーター.csv"))
    for f in (RYUHO, OTOSHI):
        shutil.copy2(f, d_ev)
    wb = openpyxl.load_workbook(RYUHO, read_only=True)
    rr = list(wb["留保の一覧"].iter_rows(values_only=True))
    kara = sum(1 for x in rr[1:] if not (x[5] and x[6] and x[7]))
    print(("   ○ " if kara == 0 else "   × ") + f"留保の一覧 {len(rr) - 1}件・確認先／期限／既定の扱いが空の行 {kara}件")
    if kara:
        ng.append("留保の一覧の空欄")
    texts = {k: "\n".join(RT1.zenbun(f)) for k, f, _ in SEIKA}
    zan = ["語\t" + "\t".join(k for k, _, _ in SEIKA)]
    zan_ng = 0
    for w in ZANSON:
        c = [texts[k].count(w) for k, _, _ in SEIKA]
        zan.append(w + "\t" + "\t".join(str(x) for x in c))
        zan_ng += sum(c)
    zan += ["", "公表用・概要版に残さない内部の記号"]
    for w in ("【委員会協議】", "【町確認】", "⚠", "確認事項No"):
        c = [texts[k].count(w) for k in ("公表用の計画書（案）", "概要版")]
        zan.append(f"{w}\t公表用 {c[0]}\t概要版 {c[1]}")
        zan_ng += sum(c)
    ryuho = {w: texts["公表用の計画書（案）"].count(w) for w in ("［要確認］", "［要協議］", "［要内訳］")}
    zan += ["", "公表用に残る留保（扱いは 川崎町_公表用の留保一覧_v2.35.xlsx）"] + [f"{k}\t{v}" for k, v in ryuho.items()]
    open(os.path.join(d_ev, "旧表現の残存検索.txt"), "w", encoding="utf-8").write(
        "第3回と前回までのレビューで挙がった旧の表現が、提出の5冊に残っていないか（件数）\n\n" + "\n".join(zan) +
        f"\n\n旧表現・内部の記号の合計 {zan_ng}件\n")
    print(("   ○ " if zan_ng == 0 else "   × ") + f"旧表現・内部の記号の残存 {zan_ng}件（{len(ZANSON)}語×5冊＋4語×2冊）")
    if zan_ng:
        ng.append("旧表現の残存")
    for kubun, f, kyu in SEIKA:
        a, b = RT1.zenbun(kyu), RT1.zenbun(f)
        diff = list(difflib.unified_diff(a, b, os.path.basename(kyu), os.path.basename(f), n=0, lineterm=""))
        open(os.path.join(d_ev, f"差分_{kubun}.txt"), "w", encoding="utf-8").write(
            f"{kubun}：{os.path.basename(kyu)}（第3回レビューの対象）→ {os.path.basename(f)}\n"
            "段落・表のセルの単位の差分（-＝旧、+＝新）\n\n" + "\n".join(diff) + "\n")
        print(f"   ○ 差分_{kubun}（{sum(1 for x in diff if x.startswith('+') and not x.startswith('+++'))}行が新）")

    print("  ── 4 対応表（md・xlsx）")
    c = collections.Counter(r[3] for r in TAIOU)
    md = ["# Red Team 再レビュー第3回（令和8年10月24日）への対応表（影響範囲の再点検を含む）", "",
          f"> レビュー原本：`{REVIEW}`（素案 v2.33・公表用 v2.33・概要版 v2.33・資料 v18・問答集 v15 が対象）",
          "> 本日2度目の受領は同じ内容（MD5 一致）。v2.34 での是正に加え、指摘の語と同じ趣旨の記述を5冊で走査し"
          "（03_エビデンス/影響範囲の走査.txt）、残っていたものを是正した",
          "> 対応後の版：協議用素案 Ver.2.35／公表用 v2.35／概要版 v2.35／委員会資料 v20／想定問答集 v16（v16 は走査で是正対象なし）",
          "> 本表の正本は `07_ソーススクリプト/build_redteam_r3eikyou_teishutsu_R8.10.24.py` の TAIOU・SAIHANTEI", "",
          "## 1　第3回の指摘（R3-01〜R3-11）と影響範囲の再点検で是正したもの", "",
          "| ID | 重要度 | 指摘 | 対応 | 対応の内容 | 箇所 | エビデンス | 残る事項 |", "|---|---|---|---|---|---|---|---|"]
    for row in TAIOU:
        md.append("| " + " | ".join(x.replace("|", "／") for x in row) + " |")
    md += ["", "対応の区分：" + "／".join(f"{k} {v}件" for k, v in c.items()), "",
           "## 2　先行指摘（N01〜N11）の再判定への対応", "", "| 前回ID | 第3回の判定 | 今回の対応 |", "|---|---|---|"]
    for row in SAIHANTEI:
        md.append("| " + " | ".join(row) + " |")
    md += ["", "## 3　第3回レビュー §6 の優先順位との対応", "",
           "| 時期 | 項目 | 状態 |", "|---|---|---|",
           "| 第2回の配付前 | R3-04 法令／R3-06 試算の表現／R3-01 の委員会・問答集の断定／R3-07 組版 | 是正済 |",
           "| 第2回の配付前 | R3-05 委員会回数の契約手続 | 資料・問答集の書き方は是正済。手続の判断は町の契約担当（No.49） |",
           "| 第3回まで | No.190・供給実態・基金見込額・報酬改定・所得段階・地域支援事業の費用・KPI の目標と分母・認知症本人の意見聴取 | 業務工程管理表の確認事項で継続（No.190・159・85・148・34・70・90 ほか） |",
           "| 公表前 | 留保103件の処理の決定・頁数と納品条件・原票突合と組版の検査 | 留保の台帳に扱いの列を置いた（処理は町）。頁数は No.83。37サービスの年報突合は済 |",
           "", "## 4　当方で確かめられないもの（未実施として残す）", "",
           "- 見える化の将来推計総括表・保険料の算定ブック・基金の決算・交付金の該当状況調査の原本と成果物の全数の突合は未実施"
           "（令和7年度の給付費37サービスは年報の円単位で突合済み）。",
           "- 委託契約書は手元にない。全3回への変更の契約上の扱いは町の契約担当のご判断による。",
           "- 「8軸51/51適合」は記載の網羅性の点検であり、同じログの未実施10事項は資料の未受領による。", ""]
    open(MD_OUT, "w", encoding="utf-8").write("\n".join(md))
    shutil.copy2(MD_OUT, d_taiou)
    wbo = openpyxl.Workbook()
    ws = wbo.active
    ws.title = "第3回の指摘"
    ws.append(["ID", "重要度", "指摘", "対応", "対応の内容", "箇所", "エビデンス", "残る事項"])
    for row in TAIOU:
        ws.append(list(row))
    for i, w in enumerate((7, 9, 40, 14, 80, 28, 34, 30), 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws2 = wbo.create_sheet("先行指摘の再判定")
    ws2.append(["前回ID", "第3回の判定", "今回の対応"])
    for row in SAIHANTEI:
        ws2.append(list(row))
    for i, w in enumerate((24, 22, 80), 1):
        ws2.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    for w_ in (ws, ws2):
        for r in w_.iter_rows():
            for cc in r:
                cc.alignment = Alignment(wrap_text=True, vertical="top")
                cc.font = Font(name="游ゴシック", size=9, bold=(cc.row == 1))
                if cc.row == 1:
                    cc.fill = PatternFill("solid", fgColor="D9E1F2")
        w_.freeze_panes = "A2"
    wbo.save(XLSX_OUT)
    shutil.copy2(XLSX_OUT, d_taiou)
    shutil.copy2(REVIEW, d_taiou)
    print(f"   ○ {len(TAIOU)}件（" + "／".join(f"{k}{v}" for k, v in c.items()) + f"）・再判定 {len(SAIHANTEI)}行")

    print("  ── 5 スクリプト")
    for f in SCRIPTS:
        shutil.copy2(SC + f, d_sc)

    print("  ── 6 README（成果品一覧・MD5）")
    rd = [f"# 川崎町 第10期計画　Red Team 再レビュー第3回対応の一式（影響範囲の再点検を含む・{HI}）", "",
          "川崎町（宮城県柴田郡）高齢者保健福祉計画・第10期介護保険事業計画", "",
          "## 中身", "", "| フォルダ | 内容 |", "|---|---|",
          "| 01_成果物 | 是正後の5冊（docx）と、当方の環境で作った PDF |",
          "| 02_レビュー対応 | 対応表（md・xlsx）と第3回レビュー原本 |",
          "| 03_エビデンス | 影響範囲の走査・点検ログ・検算（成果物・年報・算定の正本と突合）・年報との突合の明細・公表用の留保一覧・旧表現の残存検索・差分（第3回レビューの対象の版から）・原票抜粋 |",
          "| 04_スクリプト | 是正と点検に用いたスクリプト（リポジトリの 07_ソーススクリプト と同じもの） |", "",
          "## 成果品一覧", "", "| 区分 | ファイル | PDF頁数（当方の環境） | MD5 | 比べた旧の版 |", "|---|---|---|---|---|"]
    for kubun, f, n, h, kyu in ichiran:
        rd.append(f"| {kubun} | {f} | {n} | `{h}` | {kyu} |")
    rd += ["", "⚠ 頁数は当方の環境（HG丸ｺﾞｼｯｸM-PRO・游ゴシックを IPAゴシックで代用した LibreOffice）による。"
           "配布用の PDF は町の環境（Word）で作り、頁送りを確かめていただく。",
           "", "## 点検ログ", "", "| 点検 | 終了コード |", "|---|---|"]
    for na, rc in logs:
        rd.append(f"| {na} | {rc} |")
    rd += ["", "## 再現の手順", "", "リポジトリの `kawasaki_project/` で次を実行すると、同じ一式ができます。", "", "```",
           "python3 07_ソーススクリプト/fix_soan_v2340_r3review.py", "python3 07_ソーススクリプト/fix_soan_v2350_eikyou.py",
           "python3 07_ソーススクリプト/fill_toc_pages_v200.py 01_第10期_最新版成果品/川崎町_計画書素案_v2.35_協議用.docx",
           "python3 07_ソーススクリプト/build_kouhyou_soan.py",
           "python3 07_ソーススクリプト/fill_toc_pages_v200.py 01_第10期_最新版成果品/川崎町_計画書案_公表用_v2.35.docx",
           "python3 07_ソーススクリプト/build_gaiyo_v230.py", "python3 07_ソーススクリプト/build_2kai_v19.py",
           "python3 07_ソーススクリプト/build_2kai_v20.py",
           "python3 07_ソーススクリプト/build_2kai_qa_v16.py",
           "python3 07_ソーススクリプト/build_redteam_r3eikyou_teishutsu_R8.10.24.py", "```", ""]
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
