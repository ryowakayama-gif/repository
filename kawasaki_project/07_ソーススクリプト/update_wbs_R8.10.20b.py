# -*- coding: utf-8 -*-
"""業務工程管理表の更新（作業順位26の実体を実施・令和8年10月20日）

  python3 07_ソーススクリプト/update_wbs_R8.10.20b.py

なぜ更新するか
  作業順位26（県の公表名簿との突合）の**実体**を、介護サービス情報
  公表システムのオープンデータにより実施した。
  町内の事業所が確定し、計画素案の［要確認］3件が解け、誤り1件が直った。

更新するもの
  03_確認事項一覧　No.29（事業所の一覧）・No.154（高齢者向け住まい）
  04_成果品管理　　成果品3 に突合結果を登録／成果品4 を Ver.2.20 に
  08_作業順位　　　順位26 を「完了」に／順位34 として記録

⚠ **確認事項の番号は増やさない。**
⚠ **埋まらなかったものを「埋まった」ことにしない。**
  指定の年月日・指定の主体・入所者数・入所率・待機者数、及び
  指定を受けていない有料老人ホーム・サ高住は、引き続き確認待ちである。
"""
import csv
import os
import sys

import openpyxl
from openpyxl.styles import Alignment, Font

WBS = "川崎町_業務工程管理表.xlsx"
CSV_MEISAI = ("09_元資料/その他データ/"
              "公表システム_オープンデータ_川崎町_R8.6末.csv")
SOAN = "川崎町_計画書素案_v2.20_事業所確定版.docx"
TOTSUGO = "川崎町_事業所・住まい_突合結果_R8.10.20.xlsx"
FONT = Font(name="游ゴシック", size=9)
# ⚠ 同じ日に2度目の更新であるため、印を分ける。
#   1度目（update_wbs_R8.10.20.py）は「【令和8年10月20日 更新】」。
HIZUKE = "【令和8年10月20日 更新②（突合の実施）】"


def yomu():
    with open(CSV_MEISAI, encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    ban = sorted(set(r["事業所番号"] for r in rows))
    hojin = sorted(set(r["法人番号"] for r in rows))
    svc = {}
    for r in rows:
        k = r["サービスコード"]
        t = (r.get("定員") or "0").strip()
        n = int(float(t)) if t.replace(".", "").isdigit() else 0
        o = svc.setdefault(k, {"件数": 0, "定員": 0, "未報告": 0})
        o["件数"] += 1
        if n > 0:
            o["定員"] += n
        else:
            o["未報告"] += 1
    return rows, ban, hojin, svc


def iro(ws, r, n):
    for c in range(1, n + 1):
        ws.cell(r, c).font = FONT
        ws.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")


def main():
    if not os.path.exists(WBS):
        raise SystemExit("業務工程管理表がない：" + WBS)
    rows, ban, hojin, svc = yomu()
    g540 = svc.get("540", {"件数": 0, "定員": 0})
    g320 = svc.get("320", {"件数": 0, "定員": 0, "未報告": 0})

    TSUIKI = {
        "29": (10,
               f"{HIZUKE}⚠ **町内の事業所を当方で押さえました。**"
               f"介護サービス情報公表システムのオープンデータ"
               f"（令和8年6月末時点）を市町村コード043249で抽出した結果、"
               f"町内には**{len(ban)}の事業所**があり、延べ{len(rows)}の"
               f"サービスを担っています。運営しているのは**{len(hojin)}法人**"
               f"です。計画素案を Ver.2.20 に上げ、2-7・7-1・9-1 に反映"
               f"しました。\n"
               f"⚠ **素案の記述に誤りが1件ありました。**"
               f"「小規模多機能型居宅介護等の地域密着型サービスが中心」と"
               f"していましたが、**町内に小規模多機能型居宅介護はありません。**"
               f"直しました。\n"
               f"⚠ **引き続きご確認をお願いするもの。**"
               f"①指定の年月日と指定の主体（町か県か）の別"
               f"②入所者数・入所率・稼働率・待機者数"
               f"③認知症対応型共同生活介護の定員"
               f"（{g320['未報告']}事業所が未報告で、分かっているのは"
               f"{g320['定員']}人までです）。"
               f"いずれもオープンデータの24列に含まれません。"
               f"照会票（順位27）でお尋ねする内容は変わりません。　"),
        "154": (10,
                f"{HIZUKE}⚠ **一部だけ確かめられました。**"
                f"公表システムのオープンデータにより、"
                f"**特定施設入居者生活介護の指定を受けた住まいは"
                f"町内に0件**であることを確かめました"
                f"（有料老人ホーム・軽費老人ホーム・サービス付き高齢者向け"
                f"住宅のいずれの類型も0件。地域密着型特定施設も0件）。"
                f"素案 7-1 の「0人（事業所なし）」の裏づけになります。\n"
                f"⚠ **指定を受けていない有料老人ホーム・サービス付き"
                f"高齢者向け住宅は、オープンデータには載りません。**"
                f"介護保険の指定事業所の一覧であるためです。"
                f"宮城県の有料老人ホーム届出状況の公表一覧と、"
                f"サービス付き高齢者向け住宅情報提供システムが要ります"
                f"（⚠ 後者は www.satsuki-jutaku.mlit.go.jp で、"
                f"当方の環境からは未だ届きません）。"
                f"戸数・入居者数・要介護者等である入居者の状況も"
                f"引き続き確認待ちです。　"),
    }

    wb = openpyxl.load_workbook(WBS)

    # ── 03_確認事項一覧
    ws = wb["03_確認事項一覧"]
    n = sumi = 0
    for x in range(5, ws.max_row + 1):
        no = str(ws.cell(x, 1).value or "").strip()
        if no not in TSUIKI:
            continue
        retsu, moji = TSUIKI[no]
        mae = str(ws.cell(x, retsu).value or "")
        if HIZUKE in mae:
            sumi += 1
            continue
        ws.cell(x, retsu).value = moji + mae
        iro(ws, x, 10)
        n += 1
    if n + sumi != len(TSUIKI):
        raise SystemExit(f"該当行が{n + sumi}件（{len(TSUIKI)}件のはず）")

    # ── 04_成果品管理
    ws = wb["04_成果品管理"]
    n_seika = 0
    for x in range(5, ws.max_row + 1):
        no = str(ws.cell(x, 1).value or "").strip()
        na = str(ws.cell(x, 3).value or "")
        if no == "3" and TOTSUGO not in na:
            ws.cell(x, 3).value = na.rstrip() + "\n" + TOTSUGO + "（7シート）"
            ws.cell(x, 8).value = (
                f"{HIZUKE}事業所・住まいの突合の結果を登録した。"
                f"町内{len(ban)}事業所・{len(hojin)}法人。"
                f"⚠ **指定の年月日・入所率・未指定の住まいは"
                f"含まれない**（確認事項No.29・No.154）。　"
                + str(ws.cell(x, 8).value or ""))
            iro(ws, x, 9)
            n_seika += 1
        elif no == "4" and "v2.19_認知症推計版" in na:
            ws.cell(x, 3).value = na.replace(
                "川崎町_計画書素案_v2.19_認知症推計版.docx", SOAN)
            ws.cell(x, 8).value = (
                f"{HIZUKE}素案を Ver.2.20 に上げた"
                f"（2-7・7-1・9-1 に町内の事業所を反映。全116頁）。"
                f"⚠ **誤りを1件直した**（町内にない小規模多機能型居宅介護を"
                f"「中心」としていた）。［要確認］は18件→15件。　"
                + str(ws.cell(x, 8).value or ""))
            iro(ws, x, 9)
            n_seika += 1
    if n_seika != 2:
        raise SystemExit(f"04シートの該当行が{n_seika}件（2件のはず）")

    # ── 08_作業順位
    ws2 = wb["08_作業順位"]
    n26 = 0
    for x in range(5, ws2.max_row + 1):
        if str(ws2.cell(x, 1).value or "").strip() == "26":
            ws2.cell(x, 8).value = "完了"
            mae = str(ws2.cell(x, 9).value or "")
            if HIZUKE not in mae:
                ws2.cell(x, 9).value = (
                    mae + f"　{HIZUKE}⚠ **突合の実体を実施した。**"
                    f"公表システムのオープンデータにより町内"
                    f"{len(ban)}事業所・{len(hojin)}法人を確定し、"
                    f"素案 Ver.2.20 に反映した。"
                    f"地域密着型介護老人福祉施設入所者生活介護の町内の定員"
                    f"{g540['定員']}人（{g540['件数']}施設）が埋まり、"
                    f"小規模多機能型居宅介護・介護医療院・在宅系の地域密着型が"
                    f"町内にないことが確定した。"
                    f"⚠ 指定の年月日・入所率・未指定の住まいは残る。")
            iro(ws2, x, 9)
            n26 += 1
    if n26 != 1:
        raise SystemExit("順位26 の行が見つからない")

    jun = [str(ws2.cell(x, 1).value or "").strip()
           for x in range(5, ws2.max_row + 1)]
    if "34" not in jun:
        r = ws2.max_row + 1
        while str(ws2.cell(r - 1, 1).value or "").strip() != "33":
            r -= 1
            if r < 6:
                raise SystemExit("順位33 の行が見つからない")
        ws2.insert_rows(r)
        for c, v in enumerate([
            "34", "事業所の突合（公表システムのオープンデータによる）",
            "R8.10.20",
            "ご指示による。許可済みのドメインから取れることが分かったため、"
            "環境の設定の変更も町のご回答も待たずに行える。",
            "束4のうちNo.29／束10のうちNo.154",
            "事業所・住まいの突合結果（xlsx・7シート）と素案 Ver.2.20",
            "半日", "完了",
            f"令和8年10月20日　町内{len(ban)}事業所・{len(hojin)}法人を確定。"
            f"⚠ **素案の誤りを1件直した**（町内にない小規模多機能型居宅介護を"
            f"「中心」としていた）。［要確認］18件→15件。"
            f"⚠ **町民が利用しているサービスのうち7種類は町内に事業所がなく、"
            f"町外の事業所の利用である**ことが分かり、素案2-7に加えた。"
            f"担い手が{len(hojin)}法人に集まっていることも加えた。"
            f"⚠ 指定の年月日・入所率・未指定の住まいはオープンデータに"
            f"含まれず、確認待ちのまま。",
        ], start=1):
            ws2.cell(r, c).value = v
        iro(ws2, r, 9)

    wb.save(WBS)
    print(f"更新　{WBS}")

    # ══════════════════════════ 自己点検
    print()
    futeki = []

    def judge(ok, mi):
        print(("   ○ " if ok else "   × ") + mi)
        if not ok:
            futeki.append(mi)

    wb2 = openpyxl.load_workbook(WBS, data_only=True)
    ws = wb2["03_確認事項一覧"]
    nos = [str(ws.cell(x, 1).value or "").strip()
           for x in range(5, ws.max_row + 1)]
    nos = [x for x in nos if x.isdigit()]
    judge(len(nos) == len(set(nos)), f"重複がない（{len(nos)}件）")
    saidai = max(int(x) for x in nos)
    judge(not [i for i in range(1, saidai + 1) if str(i) not in nos],
          f"欠番がない（最大 {saidai}）")
    judge(saidai == 186, "確認事項を増やしていない")

    hon = {}
    for x in range(5, ws.max_row + 1):
        no = str(ws.cell(x, 1).value or "").strip()
        if no.isdigit():
            hon[no] = "\n".join(str(ws.cell(x, c).value or "")
                                for c in range(1, 11))
    judge(f"{len(ban)}の事業所" in hon.get("29", ""),
          f"No.29 に町内の事業所 {len(ban)} が入っている")
    judge("小規模多機能型居宅介護はありません" in hon.get("29", ""),
          "No.29 に誤りを直した件が入っている")
    judge("0件" in hon.get("154", ""),
          "No.154 に特定施設が町内0件であることが入っている")
    for no in ("29", "154"):
        x = [r for r in range(5, ws.max_row + 1)
             if str(ws.cell(r, 1).value or "").strip() == no][0]
        if str(ws.cell(x, 8).value or "").strip() != "確認待ち":
            judge(False, f"No.{no} が「確認待ち」でない")
    judge(True, "残るご確認は「確認待ち」のまま（適合に丸めていない）")

    ws = wb2["04_成果品管理"]
    na3 = na4 = ""
    for x in range(5, ws.max_row + 1):
        no = str(ws.cell(x, 1).value or "").strip()
        if no == "3":
            na3 = str(ws.cell(x, 3).value or "")
        if no == "4":
            na4 = str(ws.cell(x, 3).value or "")
    judge(TOTSUGO in na3, "成果品3 に突合結果が登録されている")
    judge(SOAN in na4 and "v2.19" not in na4,
          "成果品4 が Ver.2.20 を指している")

    ws2 = wb2["08_作業順位"]
    jun = [str(ws2.cell(x, 1).value or "").strip()
           for x in range(5, ws2.max_row + 1)]
    jun = [x for x in jun if x.isdigit()]
    judge(jun == [str(i) for i in range(1, len(jun) + 1)],
          f"順位が1から{len(jun)}まで連なっている")
    for x in range(5, ws2.max_row + 1):
        if str(ws2.cell(x, 1).value or "").strip() == "26":
            judge(str(ws2.cell(x, 8).value or "").strip() == "完了",
                  "順位26 が「完了」になっている")

    import re
    zen = []
    for sh in ("03_確認事項一覧", "04_成果品管理", "08_作業順位"):
        for row in wb2[sh].iter_rows(values_only=True):
            zen += [str(v) for v in row if v is not None]
    judge(not [s for s in zen
               if re.search(r"(北塩原村|金ケ崎町|金ヶ崎町|大雪)", s)],
          "台帳に他団体の固有名称を書いていない")
    judge(not [s for s in zen if re.search(r"(鶴寿会|光友会)", s)],
          "台帳に個々の法人名を書いていない")

    print()
    if futeki:
        print(f"   × 不適合 {len(futeki)}件")
        sys.exit(1)
    print("   ○ 自己点検 12件　不適合0件")


if __name__ == "__main__":
    main()
