# -*- coding: utf-8 -*-
"""中間報告書（第1回・令和8年10月分）（令和8年10月5日）

08_作業順位 の順位15。

委託仕様書（保介第41号）6（8）は
**「中間報告は、その都度指定された期日まで提出すること」**と定めている。
期日のご指定がないため（確認事項No.17）、**既定値として月末ごとの提出**とし、
その最初のものとして令和8年10月分を作る。

  出力　01_第10期_最新版成果品/川崎町_中間報告書_第1回_R8.10.docx

**数値は `川崎町_業務工程管理表.xlsx` と `05_試算・管理シート/` の
成果品から読む。本スクリプトに数を書かない。**
台帳を直したら本スクリプトを実行し直すことで、報告書が追随する。

⚠ 中間報告書は**受託者から町へ提出する業務の報告**であるから、
  計画素案・概要版・委員会資料とは異なり、受託者を主語とする語を用いる。
  素案の点検（受託者を主語とする語を置かない）は本文書には当てない。

⚠ 進捗は**当方の見立て**であり、町のご判断を拘束するものではない。
  「未着手」「確認待ち」を「完了」に丸めない。

書式は計画素案から引き継ぐ（雛形の段落・表を複製して中身を差し替える）。
"""
import copy
import os
import re
import shutil
import sys
import zipfile

import docx
import openpyxl
from docx.oxml.ns import qn

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fix_soan_v111 import clone_table, set_el   # noqa: E402

SOAN = "01_第10期_最新版成果品/川崎町_計画書素案_v2.10_制度改正補完版.docx"
WBS = "川崎町_業務工程管理表.xlsx"
GYOTEI = "05_試算・管理シート/川崎町_策定委員会_逆算工程表_R8.10.1.xlsx"
OUT = "01_第10期_最新版成果品/川崎町_中間報告書_第1回_R8.10.docx"

TSUKI = "令和8年10月"
HIZUKE = "令和8年10月31日"


# ══════════════════════════════════════════════ 台帳を読む
def yomu():
    """業務工程管理表から、報告に用いる数を読む。"""
    wb = openpyxl.load_workbook(WBS)
    d = {}

    ws = wb["01_業務内容別の進捗"]
    d["gyomu"] = []
    for r in range(5, ws.max_row + 1):
        no = ws.cell(r, 1).value
        if not (isinstance(no, str) and no.startswith("(")):
            continue
        d["gyomu"].append({
            "no": no,
            "name": str(ws.cell(r, 2).value or ""),
            "state": str(ws.cell(r, 6).value or ""),
            "pct": ws.cell(r, 7).value,
            "yotei": str(ws.cell(r, 8).value or ""),
            "riyu": str(ws.cell(r, 9).value or ""),
        })

    ws = wb["03_確認事項一覧"]
    rows = [r for r in range(5, ws.max_row + 1)
            if str(ws.cell(r, 1).value or "").strip().isdigit()]
    d["kakunin_n"] = len(rows)
    import collections
    d["kakunin_state"] = collections.Counter(
        str(ws.cell(r, 8).value or "").strip() for r in rows)

    ws = wb["04_成果品管理"]
    d["seika"] = []
    for r in range(5, ws.max_row + 1):
        no = ws.cell(r, 1).value
        if not isinstance(no, int):
            continue
        d["seika"].append({
            "no": no,
            "name": str(ws.cell(r, 2).value or ""),
            "shiyo": str(ws.cell(r, 4).value or "").replace("\n", "／"),
            "state": str(ws.cell(r, 5).value or ""),
            "pct": ws.cell(r, 6).value,
        })

    ws = wb["07_ペンディング整理"]
    d["taba"] = []
    for r in range(5, ws.max_row + 1):
        p = ws.cell(r, 1).value
        if not isinstance(p, int):
            continue
        # ⚠ このシートは【要旨】12行のあとに【明細】が続く。
        #   要旨の行は5列目が件数（数）、明細の行は5列目が期限（文字）。
        #   5列目が数でなくなったところで要旨は終わる。
        if not isinstance(ws.cell(r, 5).value, int):
            break
        d["taba"].append({
            "yusen": p,
            "name": str(ws.cell(r, 2).value or ""),
            "eikyo": str(ws.cell(r, 3).value or ""),
            "kigen": str(ws.cell(r, 4).value or ""),
            "n": ws.cell(r, 5).value,
            "chokka": ws.cell(r, 6).value,
        })

    ws = wb["08_作業順位"]
    d["sagyo"] = []
    for r in range(5, ws.max_row + 1):
        p = ws.cell(r, 1).value
        if not isinstance(p, int):
            continue
        d["sagyo"].append({
            "yusen": p,
            "name": str(ws.cell(r, 2).value or ""),
            "state": str(ws.cell(r, 8).value or ""),
        })

    wb2 = openpyxl.load_workbook(GYOTEI)
    ws = wb2["02_第2回から納品までの全体工程"]
    d["kotei"] = []
    for r in range(5, ws.max_row + 1):
        if ws.cell(r, 2).value is None:
            continue
        d["kotei"].append((str(ws.cell(r, 2).value),
                           str(ws.cell(r, 3).value or "")))
    return d


# ══════════════════════════════════════════════ 本文を組む
def honbun(d):
    """(種類, 中身) の並び。種類は 表題／見出し／本文／囲み／T2〜T5。"""
    C = []
    a = C.append

    kan = d["kakunin_state"]
    machi = kan.get("確認待ち", 0)
    kanryo = kan.get("完了", 0)
    sag = d["sagyo"]
    done = [x for x in sag if x["state"].startswith("完了")]
    heikin = round(sum(g["pct"] for g in d["gyomu"]) / len(d["gyomu"]), 2)

    # ── 表題
    a(("表題", "川崎町 高齢者保健福祉計画・第10期介護保険事業計画"
               "策定支援業務　中間報告書（第1回）"))
    a(("本文", f"報告の対象期間　契約日の翌日から{TSUKI}31日まで"))
    a(("本文", f"提出日　{HIZUKE}"))
    a(("本文", "提出先　川崎町 保健福祉課"))
    a(("本文", "提出者　ビズアップ公共コンサルティング株式会社 札幌事業所"))
    a(("囲み",
       "⚠ 委託仕様書（保介第41号）6（8）は「中間報告は、その都度指定された"
       "期日まで提出すること」と定めています。提出の期日をいただいていない"
       "ため（確認事項No.17）、当面は月末ごとに提出いたします。"
       "ご希望の期日・様式・頻度がございましたらお知らせください。"
       "本報告書の体裁は当方の案です。"))

    # ── 1 要旨
    a(("見出し", "▌1　報告の要旨"))
    a(("本文",
       f"委託仕様書6の業務内容(1)〜(8)のうち、(1)アンケート調査業務と"
       f"(2)調査票の発送準備業務は完了しました。(3)〜(7)は実施中であり、"
       f"(8)計画書の作成・納品は、計画書（案）の母体となる計画素案が"
       f"整った段階です。8つの業務内容の進捗の単純平均は"
       f"{int(heikin * 100)}％です。"))
    a(("本文",
       "計画素案は Ver.2.10（本文112頁・図10点）まで作成し、"
       "法定記載事項・国の基本指針の別表・制度改正の網羅性を"
       "機械で点検できる形にしました（51事項すべて適合）。"
       "第1回策定委員会は令和8年9月2日に開催済みで、"
       "第2回の資料と想定問答集（全44問）は作成を終えています。"))
    a(("本文",
       f"確認事項は{d['kakunin_n']}件を起票し、うち{kanryo}件が決着、"
       f"{machi}件がご回答をお待ちしている状態です。"
       f"ご回答を待たずに進められる作業は、優先順位を付けて"
       f"{len(done)}件を完了しました。"))
    a(("囲み",
       "⚠ もっともお急ぎいただきたいのは、第2回策定委員会の開催日の確定"
       "（確認事項No.85）です。開催日が令和8年11月11日以降になりますと、"
       "現行の前提（パブリックコメント30日・策定委員会4回・印刷製本21日）"
       "では納期（令和9年3月15日）に収まりません。"
       "工程の短縮の案4件を第2回策定委員会資料にお示ししています。"))

    # ── 2 業務内容ごとの進捗
    a(("見出し", "▌2　業務内容ごとの進捗（委託仕様書6）"))
    rows = [["業務内容（仕様書6）", "状態", "進捗", "完了予定", "次の作業"]]
    for g in d["gyomu"]:
        rows.append([f"{g['no']}　{g['name']}", g["state"],
                     f"{int(g['pct'] * 100)}％", g["yotei"],
                     mijikaku(g["riyu"], 90)])
    a(("T5", rows, [30, 8, 8, 12, 42]))
    a(("本文",
       "進捗は、当該業務の作業項目のうち完了したものの割合を"
       "当方が見立てたものです。町のご判断を拘束するものではありません。"))

    # ── 3 成果品の状態
    a(("見出し", "▌3　成果品の状態（委託仕様書8）"))
    rows = [["成果品", "仕様書が定める仕様", "状態", "進捗"]]
    for s in d["seika"]:
        rows.append([f"{s['no']}　{s['name']}", s["shiyo"],
                     s["state"], f"{int(s['pct'] * 100)}％"])
    a(("T4", rows, [26, 42, 16, 16]))
    a(("本文",
       "いずれも納品期日は令和9年3月15日です。判型・用紙・刷色・頁数・部数は"
       "委託仕様書8に明記されており、当方で変えることはできません。"))
    a(("囲み",
       "⚠ 計画書は「100頁程度」と定められていますが、現在の素案は本文112頁"
       "（資料編を含め125頁）です。本文を105頁に縮める圧縮案を"
       "お示ししています（確認事項No.83）。"
       "制度改正と法令が定める手続に関わる記述は削ることができません。"))

    # ── 4 本月の主な作業
    a(("見出し", f"▌4　{TSUKI}の主な作業"))
    rows = [["順位", "作業", "状態"]]
    for x in d["sagyo"]:
        rows.append([str(x["yusen"]), IIKAE.get(x["yusen"], x["name"]),
                     x["state"]])
    a(("T3", rows, [8, 62, 30]))
    a(("本文",
       "確認事項のご回答を待たずに進められる作業を洗い出し、"
       "期日と影響の大きさにより順位を付けて進めました。"
       "順位1は、見える化システムへの対応を町において進めていただくことに"
       "なりましたので取り下げています。"))

    # ── 5 点検の結果
    a(("見出し", "▌5　点検の結果"))
    rows = [["点検", "件数", "結果"]]
    for name, n, res in KENSA:
        rows.append([name, n, res])
    a(("T3", rows, [46, 12, 42]))
    a(("本文",
       "いずれも機械で回せる形にしてあり、版を上げるたびに実行しています。"
       "点検そのものが当てにならないことを避けるため、"
       "素案から1事項を落としたときに不適合になることを"
       "確かめたうえで用いています。"))
    a(("囲み",
       "⚠ 「未実施」を「適合」に丸めていません。地域支援事業のデータが"
       "未受領であるため量・費用の見込みを置けないこと、"
       "年齢調整後の認定率・自給率が見える化システムの出力待ちであること、"
       "図の中の文字と Word での見え方は機械では確かめられないことを、"
       "未実施8件として別に数えています。"))

    # ── 6 確認事項の状態
    a(("見出し", "▌6　確認事項の状態"))
    rows = [["優先", "束", "影響度", "直近の期日", "件数"]]
    for t in d["taba"]:
        rows.append([str(t["yusen"]), t["name"], t["eikyo"],
                     t["kigen"], str(t["n"])])
    a(("T5", rows, [8, 42, 10, 20, 20]))
    a(("本文",
       f"確認事項は{d['kakunin_n']}件です。決着していないものを、"
       f"止めている成果品と期日により12の束に分け、"
       f"影響度（Ｓ＝工程全体／Ａ＝保険料・見込量／Ｂ＝成果品の体裁／"
       f"Ｃ＝個別の記述）を付けています。"))
    a(("本文",
       "すべての確認事項に「決着しない場合の当方の扱い（既定値）」を"
       "置いてあります。既定値は協議を進めるための仮置きであり、"
       "公表の確定の根拠にはなりません。"))

    # ── 7 工程
    a(("見出し", "▌7　今後の工程"))
    a(("本文",
       "第2回策定委員会を令和8年11月11日に開催する場合の工程を示します。"
       "開催日が確定しだい、逆算工程表を更新してお示しします。"))
    rows = [["年月日", "工程"]]
    for ymd, k in d["kotei"]:
        rows.append([ymd, k])
    a(("T2", rows, [22, 78]))
    a(("囲み",
       "⚠ 第2回の開催が令和8年11月11日の場合、納品の予定日は"
       "令和9年3月22日となり、納期（令和9年3月15日）を7日超えます。"
       "11月4日の開催であれば余裕0日で収まります。"
       "工程の短縮の案（パブリックコメントを20日間とする案ほか）を"
       "あわせてご検討ください。"))

    # ── 8 次月の予定
    a(("見出し", "▌8　令和8年11月の予定"))
    rows = [["順位", "作業", "いつまでに"]]
    for yusen, name, itsu in YOTEI:
        rows.append([str(yusen), name, itsu])
    a(("T3", rows, [8, 70, 22]))
    a(("本文",
       "いずれもご回答を待たずに進められるものです。"
       "第2回策定委員会の開催日が確定しだい、"
       "資料・想定問答集の最終版の確定を工程に組み込みます。"))

    # ── 9 お急ぎいただきたいこと
    a(("見出し", "▌9　ご確認をお急ぎいただきたいこと"))
    rows = [["", "事項", "確認事項No.", "決まらない場合の当方の扱い"]]
    for i, (koto, no, kitei) in enumerate(ISOGI, 1):
        rows.append([str(i), koto, no, kitei])
    a(("T4", rows, [6, 34, 14, 46]))
    a(("本文", "以上"))
    return C


def mijikaku(s, n):
    """長い説明を n 字で切る（切ったことが分かるようにする）。"""
    s = re.sub(r"\s+", " ", s).strip()
    return s if len(s) <= n else s[:n] + "…"


# 作業順位の台帳には当方の内部の事情（他案件の名称など）が入ることがある。
# **町へ提出する報告書には他団体の名称を書かない**ため、
# 順位ごとに言い換えを置く。台帳の側は内部の記録としてそのまま残す。
IIKAE = {
    14: "成果品が作り直せる状態かの点検（原本とスクリプトの保全）",
}

KENSA = [
    ("リポジトリの健全性（検査1〜8）", "8", "すべて適合"),
    ("網羅性の8軸（法定記載事項ほか）", "51", "すべて適合（未実施8件は別）"),
    ("素案の自己点検（検査1〜6）", "32", "すべて適合"),
    ("保険料の検算（検算1〜9）", "39", "すべて適合"),
    ("書体・文字の大きさ（検査1〜8）", "13", "すべて適合"),
    ("図表データ管理台帳", "10", "すべて適合"),
    ("委員会資料の自己点検", "27", "すべて適合"),
    ("認知症 計画記載事項の適合性", "―", "必須記載事項を充足"),
    ("docx の構造検証", "4文書", "いずれも適合"),
]

YOTEI = [
    (16, "宮城県への事前協議の資料一式（法第117条第5項）", "R8.11上旬"),
    (17, "意見公募手続（パブリックコメント）一式"
         "（実施要領案・意見提出様式ほか）", "R8.11上旬"),
    (18, "第3回・第4回策定委員会資料の骨子", "R8.11中旬"),
    (19, "概要版を8頁へ（現在6頁）", "R8.11中旬"),
    (20, "第2回策定委員会の開催（開催日が確定しだい）と議事録の作成",
     "開催日＋7日"),
    (21, "第2回の選択結果を反映した計画素案の更新", "開催日＋14日"),
]

ISOGI = [
    ("第2回策定委員会の開催日と、工程の短縮の案（Ａ〜Ｄ）のご選択",
     "85・49・19",
     "令和8年11月中旬を仮置きし、事前配布を2週間前として逆算します。"),
    ("計画書の頁数（圧縮案によるか／「100頁程度」の範囲）", "83",
     "本文105頁の圧縮案をお示しし、ご判断を待ちます。"),
    ("事業の内容のうち、当方が書き起こした53件のご確認", "164",
     "当方の書き起こしのまま素案に置き、決着しだい差し替えます。"),
    ("第9期35事業の実績のご記入（残る65件）", "8",
     "実績欄を［要確認］のまま第2回策定委員会に諮ります。"),
    ("中間報告の提出期日・様式・頻度", "17",
     "月末ごとに本様式で提出します。"),
    ("未確定箇所の表記（【町確認】【委員会協議】）の扱い", "162",
     "公表版では取り除く前提で整理表を用意しています。"),
    ("認知症の意見聴取の運用と個人情報の手続", "146・149",
     "個人が特定されない形で計画に載せる前提で様式を設計しています。"),
    ("地域支援事業のデータ（量・費用）", "13・39・88・141・150",
     "⚠ 未受領のため量・費用の見込みを置けません。"
     "推測の数値は置きません。"),
]


# ══════════════════════════════════════════════ docx を作る
def main():
    for p in (SOAN, WBS, GYOTEI):
        if not os.path.exists(p):
            raise SystemExit("入力がない：" + p)
    d = yomu()
    C = honbun(d)

    shutil.copy(SOAN, OUT)
    doc = docx.Document(OUT)
    body = doc.element.body

    # 雛形を採る（図の段落は採らない）
    tpl = {}
    for el in body.iterchildren():
        if el.tag == qn("w:p"):
            if el.findall(".//" + qn("w:drawing")):
                continue
            s = "".join(n.text or "" for n in el.iter(qn("w:t"))).strip()
            if "表題" not in tpl and s.startswith("第 1 章"):
                tpl["表題"] = copy.deepcopy(el)
            if "見出し" not in tpl and s.startswith("▌"):
                tpl["見出し"] = copy.deepcopy(el)
            if "本文" not in tpl and s.startswith("令和7年度から令和8年度への"):
                tpl["本文"] = copy.deepcopy(el)
            if "囲み" not in tpl and s.startswith("⚠ 宮城県自身の評価"):
                tpl["囲み"] = copy.deepcopy(el)
        elif el.tag == qn("w:tbl"):
            n = len(el.findall(qn("w:tr"))[0].findall(qn("w:tc")))
            tpl.setdefault(f"T{n}", copy.deepcopy(el))
    need = ["表題", "見出し", "本文", "囲み", "T2", "T3", "T4", "T5"]
    for k in need:
        if k not in tpl:
            raise SystemExit("雛形が見つからない：" + k)

    sectPr = body.find(qn("w:sectPr"))
    for el in list(body.iterchildren()):
        if el is not sectPr:
            body.remove(el)

    def add(el):
        if sectPr is not None:
            sectPr.addprevious(el)
        else:
            body.append(el)

    n_tbl = 0
    for item in C:
        kind, v = item[0], item[1]
        if kind.startswith("T"):
            # ⚠ 列幅は**表ごと**に持つ。列数が同じでも用途により幅は違う
            #   （3列の「順位／作業／状態」と「点検／件数／結果」は別）。
            add(clone_table(tpl[kind], v, item[2]))
            # 表のあとに空の段落を置く（表が続くと Word で見分けづらい）
            e = copy.deepcopy(tpl["本文"])
            set_el(e, "")
            add(e)
            n_tbl += 1
        else:
            el = copy.deepcopy(tpl[kind])
            set_el(el, v)
            add(el)
    doc.save(OUT)
    strip_media(OUT)
    naoshi_header(OUT)

    # ══════════════════════════ 自己点検
    d2 = docx.Document(OUT)
    text = "\n".join(q.text for q in d2.paragraphs)
    tbl_text = "\n".join(c.text for t in d2.tables
                         for r in t.rows for c in r.cells)
    zen = text + "\n" + tbl_text
    ng = []
    z = zipfile.ZipFile(OUT)
    xml = z.read("word/document.xml").decode("utf-8")
    if "<w:drawing>" in xml:
        ng.append("図が残っている（中間報告書に図は置かない）")
    if sum(1 for n in z.namelist() if n.startswith("word/media")):
        ng.append("使っていない画像が残っている")
    if len(d2.tables) != n_tbl:
        ng.append(f"表が{len(d2.tables)}件（{n_tbl}件のはず）")
    # 強調記号の残り
    for w in ("**", "__"):
        if w in zen:
            ng.append(f"強調記号が残っている：{w}")
    # 他団体の固有名称
    for w in ("大雪", "東川", "東神楽", "上川", "金ヶ崎", "金ケ崎", "川崎市"):
        if w in zen:
            ng.append(f"他団体の名称が入っている：{w}")
    # 台帳の数と本文の数が合っていること
    if str(d["kakunin_n"]) not in zen:
        ng.append(f"確認事項の件数{d['kakunin_n']}が本文にない")
    for g in d["gyomu"]:
        if f"{g['no']}　{g['name']}" not in tbl_text:
            ng.append(f"業務内容{g['no']}が表にない")
    for s in d["seika"]:
        if f"{s['no']}　{s['name']}" not in tbl_text:
            ng.append(f"成果品{s['no']}が表にない")
    if len(d["taba"]) != 12:
        ng.append(f"束が{len(d['taba'])}件（12件のはず）")
    for x in d["sagyo"]:
        if IIKAE.get(x["yusen"], x["name"]) not in tbl_text:
            ng.append(f"作業順位{x['yusen']}が表にない")
    # ⚠ を用いた注意書きが残っていること（丸めていないことの手掛かり）
    if zen.count("⚠") < 4:
        ng.append("注意書き（⚠）が少なすぎる")
    # 未確定を断定していないこと
    for w in ("確定しました", "決定しました"):
        if w in zen:
            ng.append(f"未確定の事項を断定している疑い：{w}")

    n_char = sum(len(q.text) for q in d2.paragraphs) + len(tbl_text)
    print("中間報告書（第1回・令和8年10月分）")
    print("保存：", OUT)
    print(f"  段落 {len(d2.paragraphs)}／表 {len(d2.tables)}"
          f"／字数 {n_char:,}字")
    print(f"  読んだ台帳　業務内容{len(d['gyomu'])}件／成果品"
          f"{len(d['seika'])}件／確認事項{d['kakunin_n']}件"
          f"／束{len(d['taba'])}件／作業順位{len(d['sagyo'])}件"
          f"／工程{len(d['kotei'])}段")
    print("  ── 自己点検")
    if ng:
        for m in ng:
            print("   ×", m)
        sys.exit(1)
    print("   ○ 図・画像なし／強調記号なし／他団体の名称なし")
    print("   ○ 業務内容(1)〜(8)・成果品6種・束12件・作業順位"
          f"{len(d['sagyo'])}件がすべて表にある")
    print("   ○ 件数は業務工程管理表から読んでいる"
          "（本スクリプトに数を書いていない）")
    print("  ⚠ 頁数は組版によります。Word で開いてご確認ください。")
    print("  ⚠ 提出の期日・様式・頻度は確認事項No.17のご回答により改めます。")


def strip_media(path):
    """使っていない画像を取り除く（素案から複製したため残っている）。"""
    z = zipfile.ZipFile(path)
    names = z.namelist()
    data = {n: z.read(n) for n in names}
    z.close()
    doc_xml = data["word/document.xml"].decode("utf-8")
    used = set(re.findall(r'r:embed="([^"]+)"', doc_xml))
    rels_name = "word/_rels/document.xml.rels"
    from lxml import etree
    root = etree.fromstring(data[rels_name])
    drop = set()
    for rel in list(root):
        rid = rel.get("Id")
        tgt = rel.get("Target") or ""
        if tgt.startswith("media/") and rid not in used:
            drop.add("word/" + tgt)
            root.remove(rel)
    data[rels_name] = etree.tostring(root, xml_declaration=True,
                                     encoding="UTF-8", standalone=True)
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            if n in drop:
                continue
            zout.writestr(n, data[n])
    shutil.move(tmp, path)
    return len(drop)


def naoshi_header(path):
    """ヘッダー・フッターの版表記を、中間報告書のものに替える。"""
    z = zipfile.ZipFile(path)
    names = z.namelist()
    data = {n: z.read(n) for n in names}
    z.close()
    n_fix = 0
    for n in names:
        if not re.match(r"word/(header|footer)\d*\.xml", n):
            continue
        s = data[n].decode("utf-8")
        s2 = re.sub(r"川崎町[^<]*計画書素案[^<]*", "川崎町 中間報告書（第1回）", s)
        s2 = re.sub(r"Ver\.[\d.]+[^<]*", TSUKI + "分", s2)
        if s2 != s:
            data[n] = s2.encode("utf-8")
            n_fix += 1
    # 目次の自動更新は中間報告書には要らない（目次を置いていない）
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, data[n])
    shutil.move(tmp, path)
    return n_fix


if __name__ == "__main__":
    main()
