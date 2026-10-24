# -*- coding: utf-8 -*-
"""業務工程管理表の更新（県の入所状況調による残る5件の整理）
（令和8年10月22日）

  python3 07_ソーススクリプト/update_wbs_R8.10.22.py

なぜ更新するか
  ⚠ **宮城県が3か月ごとに施設別の入所状況を公表していた。**
  残る5件のうち**稼働率と休止床**が、照会を行わずに埋まった。
  ⚠ あわせて、**町内にあった軽費老人ホーム（定員70人）が令和5年10月に
  休止し、令和6年4月までに一覧から消えている**ことが分かった。

更新するもの
  03_確認事項一覧　No.159・No.154・No.170・No.50・No.51
  04_成果品管理　　成果品3（推移シート）
  08_作業順位　　　順位37

⚠ **列8（状態）には書き込まない。** 回答・対応は列10に書く。
⚠ **「埋まった」ことにしない。** 入所前の居所・床単位の休止・
  短期入所の稼働・経営の状況・町への要望は残る。
"""
import os
import sys

import openpyxl
from openpyxl.styles import Alignment, Font

WBS = "川崎町_業務工程管理表.xlsx"
SUII = "川崎町_施設の入所状況の推移_県調査_R8.10.22.xlsx"
FONT = Font(name="游ゴシック", size=9)
HIZUKE = "【令和8年10月22日 更新（県の入所状況調）】"


def iro(ws, r, n):
    for c in range(1, n + 1):
        ws.cell(r, c).font = FONT
        ws.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")


TSUIKI = {
    "159": (10,
            f"{HIZUKE}⚠ **稼働率と休止床が、照会を行わずに埋まりました。**"
            f"宮城県が3か月ごとに施設別の入所状況を公表しています"
            f"（老人福祉施設等入所状況調）。令和5年4月から令和8年7月まで"
            f"14時点を取得しました。\n"
            f"・入所率　介護老人福祉施設97.5〜100.0％と70.0〜100.0％"
            f"（⚠ 後者は恒常的に2〜5床の空き）／介護老人保健施設"
            f"90.7〜97.7％／認知症対応型共同生活介護88.9〜100.0％と"
            f"94.4〜100.0％\n"
            f"・入所希望者　介護老人福祉施設2施設の計が令和5年4月69人→"
            f"令和8年7月24人（⚠ 複数施設への申込みが重複計上）\n"
            f"・退所理由「在宅」14時点計　介護老人保健施設136人に対し"
            f"**介護老人福祉施設0人**。在宅復帰と終のすみかの別が"
            f"3年半の実績で分かれました\n"
            f"⚠ **残るのは3件**です。①入所前の居所（調査の項目にない）"
            f"②床単位の休止（施設単位までしか分からない）"
            f"③**短期入所の稼働（新たに分かった空白。県調査も公表システムも"
            f"長期（短期分を除く））**。\n"
            f"⚠ **入所前の居所には筋が立ちました。** 宮城県指定介護老人"
            f"福祉施設入所指針（令和5年7月25日）第4(1)(5)により、施設は"
            f"入所希望者の置かれている状況を確認し名簿を5年保管します。"
            f"同指針 第3 は**保険者が入所規程の作成に協力する責務**を"
            f"定めています。新たな照会ではなく、この関与の中で名簿の"
            f"**集計値**（在宅・医療機関・他の施設の別の人数）の提供を"
            f"求めるかをご判断ください。　"),
    "154": (10,
            f"{HIZUKE}⚠ **町内にあった軽費老人ホーム（定員70人）が"
            f"廃止されていました。**県の入所状況調によると、"
            f"令和5年4月1日 現員13人 → 令和5年7月1日 7人 → "
            f"**令和5年10月1日「（休止中）」現員0人** → 令和6年1月1日は"
            f"定員・現員とも空欄 → **令和6年4月1日以降は一覧に"
            f"載っていません**。\n"
            f"⚠ **第9期計画期間（令和6〜8年度）が始まる直前に、町内の"
            f"高齢者向け住まい70人分が失われています。**\n"
            f"⚠ 軽費老人ホームは老人福祉法の施設であり、介護サービス情報"
            f"公表システム（介護保険の指定事業所の仕組み）には載りません。"
            f"素案の「特定施設入居者生活介護の指定を受けた住まいは町内に"
            f"0件」は介護保険の指定の話であり、別のものです。\n"
            f"⚠ **休止・廃止の経緯と、入居されていた方の行き先をご確認"
            f"ください。** 基本指針 別表 事項八（高齢者向け住まいの状況）に"
            f"直接に効きます。素案 2-7・7-1 への反映をお待ちしています。　"),
    "170": (10,
            f"{HIZUKE}⚠ **照会を行わない整理がさらに固まりました。**"
            f"照会でお尋ねする予定であった項目のうち、"
            f"**稼働率と休止床も県の公表資料で代えられる**ことが"
            f"分かりました（確認事項No.159）。\n"
            f"⚠ 残るのは、入所前の居所・床単位の休止・短期入所の稼働・"
            f"経営の状況・町への要望の5件です。このうち入所前の居所は"
            f"県の入所指針による保険者の関与の中で、経営の状況は計画に"
            f"書かない扱いとし、町への要望は既存の場で聴きます。"
            f"**新たな照会票は要しません。**　"),
    "50": (10,
           f"{HIZUKE}⚠ **町への要望・必要な支援は、県の入所状況調にも"
           f"ありません。** 地域ケア会議・地域密着型サービス運営協議会と"
           f"いう既存の場で聴くという整理（素案 第7章7-6）を変えません。　"),
    "51": (10,
           f"{HIZUKE}⚠ **経営の状況（収支・人件費率）は、県の入所状況調に"
           f"もありません。** 公表されている仕組みがありませんので、"
           f"**計画には書かず**、介護事業経営実態調査（全国・サービス種類"
           f"別の収支差率）の水準を参照するにとどめます。　"),
}


def main():
    if not os.path.exists(WBS):
        raise SystemExit("業務工程管理表がない：" + WBS)
    wb = openpyxl.load_workbook(WBS)

    ws = wb["03_確認事項一覧"]
    n = sumi = 0
    atatta = set()
    for x in range(5, ws.max_row + 1):
        no = str(ws.cell(x, 1).value or "").strip()
        if no not in TSUIKI:
            continue
        atatta.add(no)
        retsu, moji = TSUIKI[no]
        mae = str(ws.cell(x, retsu).value or "")
        if HIZUKE in mae:
            sumi += 1
            continue
        ws.cell(x, retsu).value = moji + mae
        iro(ws, x, 10)
        n += 1
    nai = set(TSUIKI) - atatta
    if nai:
        raise SystemExit("確認事項の行が見つからない：No."
                         + "・No.".join(sorted(nai)))

    # ── 04_成果品管理　成果品3
    ws = wb["04_成果品管理"]
    n_seika = 0
    for x in range(5, ws.max_row + 1):
        if str(ws.cell(x, 1).value or "").strip() != "3":
            continue
        na = str(ws.cell(x, 3).value or "")
        if SUII in na:
            n_seika = 1
            break
        ws.cell(x, 3).value = na.rstrip() + "\n" + SUII + "（8シート）"
        ws.cell(x, 8).value = (
            f"{HIZUKE}県の入所状況調（3か月ごと・14時点）から町内施設の"
            f"推移を登録した。⚠ **稼働率と休止床が照会によらず埋まった。**"
            f"⚠ **軽費老人ホーム（定員70人）の休止から廃止までの経過**も"
            f"分かった（05シート）。残るものは07シートに6件示している。　"
            + str(ws.cell(x, 8).value or ""))
        iro(ws, x, 9)
        n_seika = 1
        break
    if n_seika != 1:
        raise SystemExit("04シートの成果品3 が見つからない")

    # ── 08_作業順位　順位37
    ws2 = wb["08_作業順位"]
    jun = [str(ws2.cell(x, 1).value or "").strip()
           for x in range(5, ws2.max_row + 1)]
    n_jun = 0
    if "37" not in jun:
        r = ws2.max_row + 1
        gyo = ["37", "残る5件の確認方針（県の公表資料の洗い出し）",
               "R8.10.22",
               "ご指示による。県内他団体の前期計画を確かめるにあたり、"
               "まず届く範囲（宮城県）に何があるかを当たったところ、"
               "県が3か月ごとに施設別の入所状況を公表していた。",
               "束のうちNo.159（稼働率・休止床が埋まる）／"
               "No.154（軽費老人ホームの廃止）／No.170・No.50・No.51",
               "施設の入所状況の推移（xlsx・8シート）／"
               "残る5件の確認方針（md）／原典18件",
               "1日", "完了",
               "⚠ **残る5件のうち稼働率と休止床が、照会を行わずに"
               "埋まった。** 令和5年4月から令和8年7月まで14時点。"
               "⚠ **町内にあった軽費老人ホーム（定員70人）が令和5年10月に"
               "休止し令和6年4月までに廃止されていた**ことが分かった"
               "（No.154）。⚠ **入所前の居所は、県の入所指針が定める"
               "保険者の責務（第3）の中で集計値を求める筋を立てた。**"
               "⚠ **県内他団体の前期計画は読めていない**（環境の通信の"
               "制限により県内市町村のサイトが403。13ホストを実測）。"
               "未実施として残し、確かめる3案をメモに示した。"
               "⚠ 途中で所在地の読み取りを誤り、別の市町村の同名施設を"
               "拾った。検査を加えて是正した。"]
        for c, v in enumerate(gyo, start=1):
            ws2.cell(r, c).value = v
        iro(ws2, r, len(gyo))
        n_jun = 1

    wb.save(WBS)

    # ══════ 自己点検
    futeki = []

    def judge(ok, mi, setsu=""):
        print(("   ○ " if ok else "   × ") + mi + (("　" + setsu)
                                                   if setsu else ""))
        if not ok:
            futeki.append(mi)

    print("業務工程管理表の更新（県の入所状況調）")
    print(f"  {WBS}")
    print()
    print("  ── 自己点検")
    judge(n + sumi == len(TSUIKI),
          f"確認事項 {len(TSUIKI)}件に書いた（新 {n}・既 {sumi}）")
    judge(n_seika == 1, "成果品管理の成果品3 を更新した")
    judge(n_jun == 1 or "37" in jun, "作業順位37 を記録した")

    wb2 = openpyxl.load_workbook(WBS, data_only=True)
    ws3 = wb2["03_確認事項一覧"]
    jotai = [str(ws3.cell(x, 8).value or "")
             for x in range(5, ws3.max_row + 1)
             if str(ws3.cell(x, 1).value or "").strip() in TSUIKI]
    judge(all(s in ("確認待ち", "決着", "保留", "回答済") for s in jotai),
          "⚠ 列8（状態）を動かしていない", "／".join(sorted(set(jotai))))
    no_all = sorted(int(s) for s in
                    (str(ws3.cell(x, 1).value or "").strip()
                     for x in range(5, ws3.max_row + 1)) if s.isdigit())
    kesu = [i for i in range(1, max(no_all) + 1) if i not in no_all]
    judge(not kesu, f"欠番0（No.1〜{max(no_all)}・{len(no_all)}件）")

    zen = []
    for sh in wb2:
        for row in sh.iter_rows(values_only=True):
            zen += [str(v) for v in row if v is not None]
    s = "\n".join(zen)
    judge("川崎市" not in s and "福岡" not in s,
          "⚠ 川崎市（神奈川県）・福岡県川崎町が混じらない")
    # ⚠ 他の市町村の施設名が混じっていないか
    for na in ("ながさか", "リリーハイム", "ケアハウスやまぶき", "八宮荘"):
        judge(na not in s, f"⚠ 他の市町村の施設「{na}」が入っていない")
    # ⚠ 町内の施設名も入れない（WBS は町へお出しするもの）
    for na in ("釜房みどりの園", "アルパイン", "蔵王グリーンホーム",
               "グループホームさわやか", "グループホームふれあい"):
        judge(na not in s, f"事業所の固有名称「{na}」が入っていない")
    judge("埋まった" in s and "残る" in s,
          "⚠ 埋まったものと残るものを書き分けている")

    print()
    if futeki:
        print(f"   × 不適合 {len(futeki)}件")
        for m in futeki:
            print(f"      {m}")
        sys.exit(1)
    print(f"   ○ 自己点検 不適合0件（確認事項 {len(no_all)}件）")


if __name__ == "__main__":
    main()
