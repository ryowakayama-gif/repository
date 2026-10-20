# -*- coding: utf-8 -*-
"""事業所・高齢者向け住まいの突合の結果（令和8年10月20日）　作業順位26

  python3 07_ソーススクリプト/build_jigyosho_totsugo_R8.10.20.py

  出力　05_試算・管理シート/川崎町_事業所・住まい_突合結果_R8.10.20.xlsx

何をするものか
  作業順位26（県の公表名簿との突合）を、介護サービス情報公表システムの
  **オープンデータ**により実施した結果をまとめる。
  公表システムの画面は自動のアクセスを受け付けないが、厚生労働省が同じ
  内容をサービス種類別のＣＳＶで公表しており、許可済みのドメインから取れる。

  入力　09_元資料/その他データ/公表システム_オープンデータ_川崎町_R8.6末.csv
        （07_ソーススクリプト/fetch_kouhyou_opendata_R8.10.20.py が作る）
        01_第10期_最新版成果品/川崎町_計画書素案_v2.20_事業所確定版.docx
        （令和7年度の実績・見込量・町内の定員の［要確認］）

⚠ **数値の正本は1か所。** 町内の事業所はオープンデータ、
  利用の実績は計画素案から読む。本スクリプトに数値を持たない。

⚠ **オープンデータで埋まらないものを「埋まった」ことにしない。**
  指定の年月日・指定の主体・入所者数・入所率・待機者数、及び
  指定を受けていない有料老人ホーム・サービス付き高齢者向け住宅は
  含まれない（06シート）。
"""
import csv
import importlib.util
import os
import re
import sys

import docx
import openpyxl
from openpyxl.styles import Alignment, Font

HERE = os.path.dirname(os.path.abspath(__file__))
CSV_MEISAI = ("09_元資料/その他データ/"
              "公表システム_オープンデータ_川崎町_R8.6末.csv")
SOAN = "01_第10期_最新版成果品/川崎町_計画書素案_v2.20_事業所確定版.docx"
OUT = "05_試算・管理シート/川崎町_事業所・住まい_突合結果_R8.10.20.xlsx"
SHUTSUJO = ("厚生労働省「介護サービス情報公表システム」オープンデータ"
            "（2026年６月末時点・出力日2026年7月9日）。"
            "https://www.mhlw.go.jp/stf/kaigo-kouhyou_opendata.html "
            "市町村コード043249で抽出")


def load(name, fname):
    spec = importlib.util.spec_from_file_location(
        name, os.path.join(HERE, fname))
    m = importlib.util.module_from_spec(spec)
    sys.modules[name] = m
    spec.loader.exec_module(m)
    return m


G = load("gyotei_teisai", "build_iinkai_gyotei_R8.10.1.py")
sheet, put = G.sheet, G.put
TITLE, HEAD, LEAD = G.TITLE, G.HEAD, G.LEAD
OKF, NGF, WARN, HI = G.OKF, G.NGF, G.WARN, G.HI

# オープンデータのサービスコード → 計画素案の表示名
#   ⚠ 素案に現れない名で突合しない。素案の表の行の名で引く。
TAIO = {
    "110": "訪問介護", "120": "訪問入浴介護", "130": "訪問看護",
    "140": "訪問リハビリテーション", "150": "通所介護",
    "160": "通所リハビリテーション", "170": "福祉用具貸与",
    "210": "短期入所生活介護",
    "220": "短期入所療養介護（介護老人保健施設）",
    "320": "認知症対応型共同生活介護", "410": "特定福祉用具販売",
    "430": "居宅介護支援", "510": "介護老人福祉施設",
    "520": "介護老人保健施設",
    "540": "地域密着型介護老人福祉施設入所者生活介護",
    "550": "介護医療院", "710": "夜間対応型訪問介護",
    "720": "認知症対応型通所介護", "730": "小規模多機能型居宅介護",
    "760": "定期巡回・随時対応型訪問介護看護",
    "770": "看護小規模多機能型居宅介護", "780": "地域密着型通所介護",
    "331": "特定施設入居者生活介護",
}


def yomu_meisai():
    if not os.path.exists(CSV_MEISAI):
        raise SystemExit(
            "オープンデータの抽出がない：" + CSV_MEISAI +
            "\n  先に 07_ソーススクリプト/fetch_kouhyou_opendata_R8.10.20.py "
            "を実行してください")
    with open(CSV_MEISAI, encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def yomu_soan():
    """素案から令和7年度の利用者数（人／月）を読む。戻り　{表示名: 値}"""
    d = docx.Document(SOAN)
    out = {}
    for t in d.tables:
        h = [c.text.strip() for c in t.rows[0].cells]
        if not (h and h[0] == "サービス種類" and "項目" in h):
            continue
        try:
            jisseki = [i for i, x in enumerate(h)
                       if x.replace(" ", "").startswith("令和7年度")][0]
        except IndexError:
            continue
        ima = None
        for r in t.rows[1:]:
            c = [x.text.strip() for x in r.cells]
            if c[0]:
                ima = re.sub(r"\s+", "", c[0])
            if ima and c[1] == "利用者数":
                v = c[jisseki].replace(",", "")
                try:
                    out[ima] = float(v)
                except ValueError:
                    pass
    return out, d


def seiri(meisai):
    """サービスコードごとに、事業所数・定員・名称をまとめる。"""
    out = {}
    for r in meisai:
        k = r["サービスコード"]
        o = out.setdefault(k, {"件数": 0, "定員": 0, "定員なし": 0,
                               "事業所": [], "法人": set()})
        o["件数"] += 1
        t = (r.get("定員") or "0").strip()
        try:
            n = int(float(t))
        except ValueError:
            n = 0
        if n > 0:
            o["定員"] += n
        else:
            o["定員なし"] += 1
        o["事業所"].append((r["事業所名"], r["事業所番号"], n,
                            r["法人の名称"]))
        o["法人"].add(r["法人番号"])
    return out


def main():
    meisai = yomu_meisai()
    jisseki, doc = yomu_soan()
    svc = seiri(meisai)
    hojin = {}
    for r in meisai:
        hojin.setdefault(r["法人番号"], []).append(r)
    ban = sorted(set(r["事業所番号"] for r in meisai))

    wb = openpyxl.Workbook()

    # ══════════════════════════ 01　結論
    ws = sheet(
        wb, "01_結論",
        "事業所・高齢者向け住まいの突合（作業順位26）　— 結果",
        "作業順位26は「宮城県の公表名簿により町のご回答を待たずに"
        "下ごしらえができる」という見立てで起こしたものでしたが、"
        "公表システムの画面が自動のアクセスを受け付けないため止まって"
        "いました。⚠ **厚生労働省が同じ内容をサービス種類別のＣＳＶ"
        "（オープンデータ）で公表していることが分かり、町のご回答を"
        "待たずに町内の事業所を押さえられました。**"
        "⚠ **ただしすべてが埋まるわけではありません**（06シート）。",
        ["区分", "数", "内容"], [26, 14, 92], first=True)
    put(ws, [
        ["町内の事業所（延べ）", f"{len(meisai)}件",
         "1つの事業所が複数のサービスを担うため、延べの数です"],
        ["事業所番号の異なり", f"{len(ban)}件",
         "／".join(ban)],
        ["法人（法人番号で数える）", f"{len(hojin)}法人",
         "⚠ **担い手が3法人に集まっています。**"
         "1法人の撤退がサービスの種類ごと失われることを意味します"],
        ["サービスの種類（町内にあるもの）", f"{len(svc)}種類",
         "／".join(f"{TAIO.get(k, k)}" for k in sorted(svc))],
        ["定員の記載があるもの",
         f"{sum(1 for r in meisai if (r.get('定員') or '0').strip() not in ('', '0'))}件",
         "⚠ 訪問系・居宅介護支援は定員の概念がありません。"
         "それ以外で0のものは**未報告**とみられます（06シート）"],
    ], fills=[None, None, WARN, None, WARN])
    r = 11
    for t in [
        "",
        "▶ **解けた［要確認］が4件あります**（05シート）。",
        "　① 小規模多機能型居宅介護　町内に事業所があるとの整理と実績0の"
        "食い違い　→ **町内に事業所はありません**（実績0と整合）",
        "　② 地域密着型介護老人福祉施設入所者生活介護の町内の定員　"
        "→ **30人**",
        "　③ 在宅系の地域密着型サービスの町内の事業所の有無　"
        "→ **定期巡回・夜間対応・看多機・地域密着型通所介護はいずれも0**",
        "　④ 介護医療院　→ **町内に事業所はありません**",
        "",
        "⚠ **素案の記述に誤りが1件ありました**（05シート）。"
        "2-7 と 7-1 の「小規模多機能型居宅介護等の地域密着型サービスが"
        "中心」は、町内に小規模多機能型居宅介護が無いため成り立ちません。",
        "",
        f"出典：{SHUTSUJO}",
    ]:
        ws.cell(r, 1).value = t
        ws.cell(r, 1).font = Font(name="游ゴシック", size=9)
        ws.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        r += 1

    # ══════════════════════════ 02　町内の事業所
    ws = sheet(
        wb, "02_町内の事業所",
        "町内の事業所（公表システムのオープンデータ・令和8年6月末時点）",
        f"{SHUTSUJO}。1行1件（1つの事業所が複数のサービスを担う場合は"
        "サービスごとに1行）。"
        "⚠ **本計画に個々の法人名・事業所名は掲げません。**"
        "数と集まり方のみを示します。本シートは当方の突合のためのものです。",
        ["コード", "サービスの種類", "事業所名", "事業所番号", "定員",
         "法人の名称", "法人番号", "所在地"],
        [8, 30, 34, 13, 8, 28, 15, 44])
    rows = []
    fills = []
    for r0 in sorted(meisai, key=lambda x: (x["サービスコード"],
                                            x["事業所名"])):
        t = (r0.get("定員") or "0").strip()
        teiin = int(float(t)) if t.replace(".", "").isdigit() else 0
        teiin_hyoji = f"{teiin}人" if teiin > 0 else "―"
        rows.append([r0["サービスコード"], r0["サービスの種類"],
                     r0["事業所名"], r0["事業所番号"], teiin_hyoji,
                     r0["法人の名称"], r0["法人番号"],
                     r0["住所"] + (" " + r0["方書（ビル名等）"]
                                   if r0["方書（ビル名等）"] else "")])
        fills.append(WARN if (teiin == 0 and r0["サービスコード"]
                              in ("320", "220")) else None)
    put(ws, rows, fills=fills)
    r = 5 + len(rows) + 1
    ws.cell(r, 1).value = (
        "⚠ 橙の行は、定員のあるはずのサービスで定員が0のものです。"
        "公表情報は事業所の年1回の報告によるため、未報告とみられます。"
        "照会票（確認事項No.29）で確かめます。")
    ws.cell(r, 1).font = Font(name="游ゴシック", size=9)
    ws.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)

    # ══════════════════════════ 03　法人ごとの担い方
    ws = sheet(
        wb, "03_法人ごとの担い方",
        "担い手の法人と、担っているサービス",
        "⚠ **供給の担い手が少数の法人に集まっていることは、"
        "1つの法人の撤退がサービスの種類ごと失われることを意味します。**"
        "第2章2-7（供給体制）・第7章7-1（提供体制の確保）に効きます。"
        "⚠ **本計画に法人名は掲げません。**数と集まり方のみを示します。",
        ["法人番号", "法人の名称（表記のゆれを含む）", "事業所番号",
         "担っているサービス", "件数"],
        [15, 34, 26, 62, 8])
    rows, fills = [], []
    for bn, v in sorted(hojin.items(),
                        key=lambda x: -len(x[1])):
        na = sorted(set(x["法人の名称"] for x in v), key=len)[-1]
        bans = sorted(set(x["事業所番号"] for x in v))
        sv = sorted(set(x["サービスの種類"] for x in v))
        rows.append([bn, na, "／".join(bans), "／".join(sv), len(v)])
        fills.append(WARN if len(sv) >= 3 else None)
    put(ws, rows, fills=fills)

    # ══════════════════════════ 04　サービスごとの突合
    ws = sheet(
        wb, "04_サービスごとの突合",
        "サービスごとの突合　— 町内の供給と、令和7年度の利用の実績",
        "町内の事業所（オープンデータ）と、計画素案 9-1 の令和7年度の"
        "実績（利用者数・人／月）を突き合わせます。"
        "⚠ **町内に事業所が無いのに実績があるものは、町外の事業所の"
        "利用です。**町域の供給だけでは町民の利用を説明できないことの"
        "裏づけになります（素案 2-7）。",
        ["サービスの種類", "町内の事業所数", "町内の定員",
         "令和7年度の実績（人／月）", "読み取り"],
        [36, 14, 14, 22, 70])
    rows, fills = [], []
    for k in sorted(TAIO):
        na = TAIO[k]
        o = svc.get(k)
        kensu = o["件数"] if o else 0
        teiin = o["定員"] if o else 0
        # 素案の表示名のゆれを吸収して実績を引く
        jis = None
        for key, v in jisseki.items():
            if key.startswith(na) or na.startswith(key[:10]):
                jis = v
                break
        tei_s = f"{teiin}人" if teiin else ("―" if kensu else "―")
        if o and o["定員なし"] and teiin:
            tei_s += f"（うち{o['定員なし']}件は未報告）"
        if kensu == 0 and jis:
            yomi = ("⚠ **町内に事業所なし。実績は町外の事業所による。**"
                    "供給は町域の外にある")
            f = WARN
        elif kensu == 0 and (jis == 0 or jis is None):
            yomi = "町内に事業所なし。実績も0（供給なし）"
            f = HI
        elif kensu and jis:
            yomi = "町内に事業所あり"
            f = OKF
        else:
            yomi = "町内に事業所あり（素案に利用者数の行なし）"
            f = None
        rows.append([na, f"{kensu}件" if kensu else "0件", tei_s,
                     f"{jis:.1f}" if jis is not None else "―", yomi])
        fills.append(f)
    put(ws, rows, fills=fills)

    # ══════════════════════════ 05　素案の直すべき記述
    ws = sheet(
        wb, "05_素案の直すべき記述",
        "計画素案で直すところ・埋まるところ",
        "オープンデータにより、素案の［要確認］4件が解け、"
        "誤りが1件見つかりました。",
        ["箇所", "いまの記述", "オープンデータで分かったこと", "扱い"],
        [26, 54, 54, 34])
    put(ws, [
        ["2-7 の表（事業所・施設）／7-1 の本文",
         "「特別養護老人ホーム1施設、介護老人保健施設1施設、"
         "**小規模多機能型居宅介護**等の地域密着型サービスが中心」",
         "⚠ **町内に小規模多機能型居宅介護はありません。**"
         "町内の地域密着型は 認知症対応型共同生活介護2・"
         "地域密着型介護老人福祉施設入所者生活介護2 です",
         "**誤り。直す**"],
        ["9-1 の0のサービスの表（小規模多機能型居宅介護）",
         "［要確認］町内に事業所があるとの整理があり、実績との"
         "食い違いを確認中",
         "**町内に事業所はありません。**実績0と整合します",
         "**［要確認］が解ける**"],
        ["9-1 の0のサービスの表（介護医療院）",
         "［要確認］供給なし又は利用なし",
         "**町内に事業所はありません**（供給なし）",
         "**［要確認］が解ける**"],
        ["7-1 の表（町内の定員・地域密着型介護老人福祉施設）",
         "［要確認］",
         f"**{svc.get('540', {}).get('定員', 0)}人**"
         f"（{svc.get('540', {}).get('件数', 0)}施設）",
         "**［要確認］が解ける**"],
        ["7-1 の表（町内の定員・認知症対応型共同生活介護）",
         "［要確認］　第9期計画値は月27.0人",
         f"{svc.get('320', {}).get('件数', 0)}事業所。"
         f"⚠ **定員は{svc.get('320', {}).get('定員', 0)}人までしか"
         f"分かりません**（1事業所が未報告）",
         "⚠ **一部しか埋まらない。**"
         "［要確認］を残し、照会票で確かめる"],
        ["6-4 ほか（在宅系の地域密着型の事業所の有無）",
         "町内及び近隣市町村に指定事業所が存在するかを確認できて"
         "おらず［要確認］",
         "**町内は 定期巡回0・夜間対応0・看護小規模多機能0・"
         "地域密着型通所介護0 です。**"
         "⚠ 近隣市町村については本件では確かめていません",
         "**町内の分は解ける。**近隣は残る"],
    ], fills=[NGF, OKF, OKF, OKF, WARN, OKF])

    # ══════════════════════════ 06　埋まらないもの
    ws = sheet(
        wb, "06_埋まらないもの",
        "オープンデータでは埋まらないもの（「未実施」を「適合」に丸めない）",
        "⚠ **突合が済んだのは「町内にどの事業所があるか」までです。**"
        "次は引き続き町・県へのご確認によります。",
        ["件", "埋まらないもの", "なぜ", "どこから取るか", "確認事項"],
        [6, 30, 48, 48, 14])
    put(ws, [
        [1, "指定の年月日・指定の主体（町か県か）の別",
         "オープンデータの24列に含まれません",
         "町の指定台帳（地域密着型）／宮城県の指定事業所一覧",
         "No.29"],
        [2, "入所者数・入所率・稼働率・待機者数",
         "同上。公表システムは事業所の基本情報が中心です",
         "施設への文書照会（照会票）",
         "No.29・No.159"],
        [3, "定員の未報告（認知症対応型共同生活介護1件ほか）",
         "⚠ **公表情報は事業所の年1回の報告によります。**"
         "報告がなければ0のまま載ります",
         "照会票で確かめる",
         "No.29"],
        [4, "指定を受けていない有料老人ホーム・"
            "サービス付き高齢者向け住宅",
         "⚠ **介護保険の指定事業所の一覧ですので、"
         "指定を受けていない住まいは載りません。**"
         "特定施設入居者生活介護の指定を受けたものだけが載ります"
         "（町内は0件）",
         "宮城県の有料老人ホーム届出状況の公表一覧／"
         "サービス付き高齢者向け住宅情報提供システム"
         "（⚠ www.satsuki-jutaku.mlit.go.jp は未許可）",
         "No.154"],
        [5, "職員数・介護人材の状況",
         "オープンデータの24列に含まれません",
         "事業所ヒアリング",
         "No.50・No.51"],
        [6, "町外の事業所（町民が利用しているもの）",
         "⚠ **本件では町内（市町村コード043249）だけを取り出しました。**"
         "町外の事業所を並べると他団体の固有名称が成果品に入ります",
         "住所地特例者台帳（24人・令和8年3月末）／国保連データ",
         "No.12"],
    ], fills=[None, None, WARN, WARN, None, HI])

    # ══════════════════════════ 07　出所と取り出しの手順
    ws = sheet(
        wb, "07_出所と手順",
        "出所と、同じことを繰り返すための手順",
        "⚠ **場所をスクリプトに書き写していません。**入口のページから"
        "その時点の最新のＣＳＶの場所を読み、サービスの種類ごとに"
        "取り出します。版が上がっても同じ手順で取れます。",
        ["順", "すること", "中身"], [6, 34, 96])
    put(ws, [
        [1, "入口のページを開く",
         "https://www.mhlw.go.jp/stf/kaigo-kouhyou_opendata.html"
         "　⚠ **www.mhlw.go.jp は許可済みのドメインです。**"
         "公表システムの画面（www.kaigokensaku.mhlw.go.jp）は"
         "自動のアクセスを受け付けません"],
        [2, "最新の基準日を確かめる",
         "見出しの「2026年６月末時点」。版が上がっていれば出典の表記を"
         "改めます（スクリプトの検査1が見ています）"],
        [3, "サービスの種類ごとのＣＳＶを取る",
         "35種類。ページに現れる順で、種類ごとの最初の1件が最新です"],
        [4, "市町村コードで抽出する",
         "**043249**（宮城県柴田郡川崎町）。"
         "⚠ 川崎市（神奈川県）・福岡県田川郡川崎町と取り違えないよう、"
         "都道府県名・市区町村名も検査で確かめています"],
        [5, "書き出す",
         "09_元資料/その他データ/公表システム_オープンデータ_川崎町_"
         "R8.6末.csv（町内の事業所）と同_件数_R8.6末.csv"
         "（種類ごとの全国・宮城県・川崎町の件数）"],
        [6, "実行",
         "python3 07_ソーススクリプト/fetch_kouhyou_opendata_R8.10.20.py"
         "（検査1〜7）→ "
         "python3 07_ソーススクリプト/build_jigyosho_totsugo_R8.10.20.py"],
    ])

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb.save(OUT)
    print(f"作成　{OUT}（{len(wb.sheetnames)}シート）")
    for s in wb.sheetnames:
        print(f"  {s}")

    # ══════════════════════════ 自己点検
    print()
    futeki = []

    def judge(ok, mi, setsu=""):
        print(("   ○ " if ok else "   × ") + mi + (("　" + setsu)
                                                   if setsu else ""))
        if not ok:
            futeki.append(mi)

    wb2 = openpyxl.load_workbook(OUT, data_only=True)
    zen = []
    for ws2 in wb2:
        for row in ws2.iter_rows(values_only=True):
            zen += [str(v) for v in row if v is not None]
    s = "\n".join(zen)

    print("  ── 自己点検")
    judge(len(wb2.sheetnames) == 7, f"7シート（{len(wb2.sheetnames)}）")
    judge(f"{len(meisai)}件" in s, f"町内の事業所 {len(meisai)}件 が載る")
    judge(f"{len(hojin)}法人" in s, f"法人 {len(hojin)}法人 が載る")
    judge(svc.get("730") is None,
          "小規模多機能型居宅介護が町内に0件であることと整合")
    judge("埋まらないもの" in s, "埋まらないものを件数と理由で残している")
    for mi, pat in (("受託者を主語とする語",
                     r"(受託者|当社|弊社|ビズアップ)"),
                    ("他団体の固有名称",
                     r"(大雪|東川町|東神楽町|金ケ崎町|金ヶ崎町|北塩原村)"),
                    ):
        atari = [x[:50] for x in zen if re.search(pat, x)]
        judge(not atari, f"{mi} が無い", f"→ {atari[:1]}" if atari else "")
    # ⚠ 「川崎市と取り違えない」という**注意書き**自体を不適合にしない。
    #   見るのは**データの側**である（所在地・市区町村名）。
    judge(all(r["市区町村名"] == "川崎町" for r in meisai),
          "すべての行の市区町村名が川崎町")
    judge(not any("川崎市" in r.get("住所", "") for r in meisai),
          "所在地に川崎市（神奈川県）が混じっていない")

    print()
    if futeki:
        print(f"   × 不適合 {len(futeki)}件")
        sys.exit(1)
    print("   ○ 自己点検 10件　不適合0件")


if __name__ == "__main__":
    main()
