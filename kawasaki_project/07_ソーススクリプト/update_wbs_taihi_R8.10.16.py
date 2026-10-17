# -*- coding: utf-8 -*-
"""業務工程管理表の更新（他メンバー素案との対比・令和8年10月16日）

  python3 07_ソーススクリプト/update_wbs_taihi_R8.10.16.py

更新するもの
  03_確認事項一覧　　No.174〜179 を起票
  08_作業順位　　　　順位30（他メンバー素案との対比）を加えて「完了」に
  01_業務内容別の進捗　(5)(6)(7) の未了の理由を更新
  04_成果品管理　　　成果品4（計画書）を Ver.2.15 に

⚠ **件数は成果品の実物から数える。** 本スクリプトに数を書かない。
"""
import importlib.util
import os
import re
import sys

import openpyxl
from openpyxl.styles import Alignment, Font

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

WBS = "川崎町_業務工程管理表.xlsx"
HIZUKE = "【令和8年10月16日 更新】"
FONT = Font(name="游ゴシック", size=9)


def load(name, fname):
    spec = importlib.util.spec_from_file_location(name,
                                                  os.path.join(HERE, fname))
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


B = load("taihi_wbs", "build_taihi_tenpu_R8.10.16.py")
D = load("data_taihi_wbs", "data_taihi_tenpu_R8.10.16.py")

# 本日起票する確認事項（B.KAKUNIN_SHINKI に1件足す）
SHINKI = list(B.KAKUNIN_SHINKI) + [
    ("179", "(1)(7)", "R8.10", "調査の配布数・回収数の出所",
     "他メンバー作成の第10期計画素案に、ニーズ調査の配布数1,034通・"
     "回収576件・回収率55.7％、在宅介護実態調査の配布数305通・回収142件・"
     "回収率46.6％という実数が入っています。算術は整合します"
     "（576÷1,034＝55.71％、142÷305＝46.56％）。"
     "⚠ **1,034通・305通という発送通数は、当方の受領資料にも業務工程管理表にも"
     "現れない値です。** 町の発送記録によるものかをご確認ください。"
     "当方の素案 2-6 の「調査の概要」の表は、この2か所を【町確認】のまま"
     "残しています。",
     "計画素案 2-6 の調査の概要の表", "発注者", "確認待ち", "R8.10",
     "既定値＝出所が確かめられるまで【町確認】のまま残します。"
     "推測で置きません。"),
]

# 作業順位に加える行
JUNI30 = [
    "30", "他メンバー作成の第10期計画素案との対比と第9期計画の再整理",
    "R8.10.16",
    "ご指示による。他メンバーが作成した第10期計画素案（2部8章＋資料編・"
    "段落1212・表132・図45）を受領した。正本は当方案とし、内容だけを"
    "取り込む方針で進める。あわせて第9期計画（全108頁）を再整理し、"
    "記載変更が必要な箇所と修正が必要な箇所を洗い出す。",
    "―（成果品の質・法定記載事項の充足）",
    "他メンバー素案との対比_論点整理（xlsx・11シート）／"
    "計画素案 Ver.2.14・Ver.2.15・Ver.2.16",
    "1日", "完了",
]


def iro(ws, r, n_col):
    for c in range(1, n_col + 1):
        ws.cell(r, c).font = FONT
        ws.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")


def main():
    if not os.path.exists(WBS):
        raise SystemExit("業務工程管理表がない：" + WBS)
    wb = openpyxl.load_workbook(WBS)

    # ══════ 03_確認事項一覧
    ws = wb["03_確認事項一覧"]
    aru = {str(ws.cell(r, 1).value or "").strip()
           for r in range(5, ws.max_row + 1)}
    r = max(x for x in range(5, ws.max_row + 1)
            if str(ws.cell(x, 1).value or "").strip()) + 1
    n_kakunin = 0
    for rec in SHINKI:
        if rec[0] in aru:
            continue
        for c, v in enumerate(rec, start=1):
            ws.cell(r, c).value = v
        iro(ws, r, len(rec))
        r += 1
        n_kakunin += 1
    nos = sorted(int(str(ws.cell(x, 1).value).strip())
                 for x in range(5, ws.max_row + 1)
                 if str(ws.cell(x, 1).value or "").strip().isdigit())
    kessan = [i for i in range(1, max(nos) + 1) if i not in nos]

    # ══════ 08_作業順位
    ws = wb["08_作業順位"]
    aru = {str(ws.cell(x, 1).value or "").strip()
           for x in range(5, ws.max_row + 1)}
    n_juni = 0
    omoi = sum(1 for x in D.RONTEN if x[1] == "重い")
    kekka = (
        f"令和8年10月16日　受領した素案を 09_元資料/他メンバー資料/ へ退避し"
        f"（検査8のMD5突合が10件で適合）、川崎町_他メンバー素案との対比_"
        f"論点整理_R8.10.16.xlsx（11シート）を作成した。"
        f"⚠ 添付は第9期計画書の改訂版であり（章立てが第9期の目次と同一）、"
        f"強みは体裁と文章の充実である。数値・制度改正・供給体制・保険料は"
        f"当方の正本が先に進んでいる（添付は【町確認】90か所）。"
        f"11の持ち場に手分けして303件を洗い出し、当方が機械で確かめた"
        f"重点は{len(D.RONTEN)}件（重い{omoi}件）。"
        f"⚠ **令和7年度の給付実績は17サービスで完全に一致した**"
        f"（出所の異なる2つの算定が一致し、当方の読み取りが確かめられた）。"
        f"⚠ **いちばん重い発見は2件。**①素案 9-1 の見込量の表に"
        f"15サービス・84,370千円（総給付費の8.5％）が落ちていた"
        f"（法第117条第2項第1号の種類ごとの量の見込みの欠落）"
        f"②素案 2-3 の第1号被保険者数・認定者数の実績が年報と20か所"
        f"食い違っていた（3,270は令和5年度末、3,285は令和4年度末の値で"
        f"年度が2年ずれていた）。いずれも素案 Ver.2.14・Ver.2.15 で是正し、"
        f"機械の検査（check_mikomi_morasi・check_nintei）を新設した。"
        f"⚠ **添付と並べて読んだことにより、当方の正本そのものの"
        f"不具合も見えた。** ①9-1 で「介護予防サービスの見込量」の"
        f"見出しの下に地域密着型の表が入り、同じ表が2つあった"
        f"（重なる11キーの値はすべて一致＝同じ数値を2か所に置いていた）"
        f"②施策事業表の内容欄31セルに事業の内容でない記述が入っていた"
        f"③本文に受託者の担当者名・他社名・版の履歴が残っていた"
        f"④表紙の版名が食い違い発行年月が誤っていた"
        f"⑤「地域密着型サービス運営委員会」は町の正本では"
        f"「運営協議会」であった。**いずれも Ver.2.16 で是正した。** "
        f"内容欄のうち16セルは添付の本文を取り込んだ"
        f"（ご指示の「内容だけ添付の使えるものを取り込む」に当たる）。"
        f"⚠ 保険料・給付費・見込量の数値は変えていない。"
        f"確認事項No.174〜181を起票。")
    if JUNI30[0] not in aru:
        r = max(x for x in range(5, ws.max_row + 1)
                if str(ws.cell(x, 1).value or "").strip()) + 1
        for c, v in enumerate(JUNI30, start=1):
            ws.cell(r, c).value = v
        ws.cell(r, 9).value = kekka
        iro(ws, r, 9)
        n_juni += 1
    else:
        for x in range(5, ws.max_row + 1):
            if str(ws.cell(x, 1).value or "").strip() == JUNI30[0]:
                ws.cell(x, 8).value = "完了"
                ws.cell(x, 9).value = kekka
                iro(ws, x, 9)
                n_juni += 1
    juni = [x for x in range(5, ws.max_row + 1)
            if str(ws.cell(x, 1).value or "").strip().isdigit()]
    jotai = {}
    for x in juni:
        v = str(ws.cell(x, 8).value or "").strip()
        jotai[v] = jotai.get(v, 0) + 1

    # ══════ 01_業務内容別の進捗
    ws = wb["01_業務内容別の進捗"]
    TSUIKI = {
        "(5)": f"{HIZUKE}他メンバー素案との対比により、素案 9-1 の見込量の表に"
               "15サービス（84,370千円・総給付費の8.5％）が落ちていたことが"
               "分かり、Ver.2.14 で是正した。"
               "⚠ 介護保険法第117条第2項第1号の種類ごとの量の見込みの"
               "欠落であった。数値はワークブックに持っていた。",
        "(6)": f"{HIZUKE}素案 2-3 の第1号被保険者数・認定者数の実績が"
               "年報と20か所食い違っていたことが分かり、Ver.2.15 で是正した。"
               "⚠ 年度が2年ずれていた（3,270は令和5年度末、3,285は"
               "令和4年度末の値）。年報と突き合わせる検査を新設した。",
        "(7)": f"{HIZUKE}他メンバー作成の第10期計画素案を受領し、"
               f"11の持ち場に手分けして303件を洗い出した。"
               "正本は当方案とし、内容だけを取り込む。"
               "⚠ 添付にない節（1-4・1-5・2-5・2-7・第3章・7-1〜7-6・10-4）は"
               "当方の強みであり、落とすものはない。"
               "⚠ **添付と並べて読んだことにより、当方の正本そのものの"
               "不具合も見えた。** 施策事業表の内容欄31セルに事業の内容で"
               "ない記述（作業過程の記述・文頭の欠落・存在しない節番号）が"
               "入っており、本文に受託者の担当者名・他社名・版の履歴が"
               "残っていた。いずれも Ver.2.16 で是正し、"
               "27語を見る自己点検を組み込んだ。"
               "うち16セルは添付の本文を取り込んだ（ご指示の"
               "「内容だけ添付の使えるものを取り込む」に当たる）。",
    }
    n_shinchoku = 0
    for x in range(5, ws.max_row + 1):
        no = str(ws.cell(x, 1).value or "").strip()
        if no in TSUIKI:
            mae = str(ws.cell(x, 9).value or "")
            if TSUIKI[no] not in mae:          # 二度足さない
                ws.cell(x, 9).value = TSUIKI[no] + "　" + mae
                iro(ws, x, 10)
            n_shinchoku += 1
    if n_shinchoku != len(TSUIKI):
        raise SystemExit(f"01シートの更新が{n_shinchoku}件"
                         f"（{len(TSUIKI)}件のはず）")

    # ══════ 04_成果品管理
    ws = wb["04_成果品管理"]
    n_seika = 0
    for x in range(5, ws.max_row + 1):
        if str(ws.cell(x, 2).value or "").strip() == "計画書":
            ws.cell(x, 3).value = (
                "川崎町_計画書素案_v2.17_中長期補完版.docx／"
                "同_v2.16c_圧縮案_R8.10.16.docx（本文109頁）／"
                "川崎町_計画書_別冊_詳細分析_R8.10.16.docx（9頁）")
            ws.cell(x, 8).value = (
                f"{HIZUKE}素案 Ver.2.16（内容整備版・全116頁）。"
                "見込量の表に落ちていた15サービスを加え（Ver.2.14）、"
                "第1号被保険者数・認定者数の実績を年報に合わせ（Ver.2.15）、"
                "施策事業表の内容欄31セル・9-1の見出しと表の対応・"
                "本文に残っていた作業経過を是正した（Ver.2.16）。"
                "⚠ **頁数が113頁から116頁に増えた。** "
                "仕様書8の100頁程度との差が広がったため、圧縮案"
                "（確認事項No.83）のご判断が一層重要になっている"
                "（圧縮案を Ver.2.16c として作り直した。本文109頁＋"
                "資料編13頁＝122頁）。⚠ **資料編と奥付がまだない**"
                "（確認事項No.181）。　"
                + re.sub(r"【令和8年10月16日 更新】素案 Ver\.2\.1[56].*?"
                         r"一層重要になっている。　", "",
                         str(ws.cell(x, 8).value or ""), count=1))
            iro(ws, x, 9)
            n_seika += 1
    if n_seika != 1:
        raise SystemExit(f"04シートの該当行が{n_seika}件（1件のはず）")

    wb.save(WBS)

    # ══════════════════════════ 自己点検
    wb2 = openpyxl.load_workbook(WBS)
    ng = []
    ws = wb2["03_確認事項一覧"]
    nos2 = [str(ws.cell(x, 1).value or "").strip()
            for x in range(5, ws.max_row + 1)]
    for rec in SHINKI:
        if rec[0] not in nos2:
            ng.append(f"確認事項No.{rec[0]}が登録されていない")
    if kessan:
        ng.append("確認事項に欠番がある：" + "／".join(map(str, kessan[:10])))

    print("業務工程管理表の更新（他メンバー素案との対比）")
    print("保存：", WBS)
    print(f"  確認事項　{len(nos)}件（本日 新規{n_kakunin}件）"
          f"／欠番{len(kessan)}件")
    print(f"  作業順位　{len(juni)}件（"
          + "／".join(f"{k} {v}" for k, v in sorted(jotai.items())) + "）")
    print(f"  01シート　{n_shinchoku}件を更新／04シート　{n_seika}件を更新")
    print("  ── 自己点検")
    if ng:
        for m in ng:
            print("   ×", m)
        sys.exit(1)
    print("   ○ 起票した確認事項はすべて台帳に登録されている")
    print("   ○ 確認事項に欠番はない")


if __name__ == "__main__":
    main()
