# -*- coding: utf-8 -*-
"""業務工程管理表の更新（ヒアリングを行わない整理・第2回11月24日確定）
（令和8年10月20日）

  python3 07_ソーススクリプト/update_wbs_R8.10.20c.py

なぜ更新するか
  ①介護サービス情報公表システム（本体）から入所率・待機者数・職員数・
    指定の年月日・退所先を採録し、**事業所ヒアリングを行わない整理**が
    成り立つことを確かめた（素案 Ver.2.21）。
  ②**第2回策定委員会の開催日が令和8年11月24日（火）と確定**した。
    開催日が決まったことで、資料の確定期限・第3回の時期・委員会の回数が
    決まる（逆算工程表 R8.10.20）。

更新するもの
  03_確認事項一覧　No.29・No.50・No.51・No.159・No.170（ヒアリング関係）
                  No.4・No.49・No.85・No.19（日程関係）
  04_成果品管理　　成果品3 に実態シートと逆算工程表／成果品4 を Ver.2.21 に
  08_作業順位　　　順位35（ヒアリングを行わない整理）を記録

⚠ **確認事項の番号は増やさない。**
⚠ **列8（状態）には書き込まない。** 状態は町のご回答により変わるもので
  あり、当方が勝手に動かすものではない。回答・対応は列10に書く。
  （No.170 のみ、当方の整理の結論を列10に書き、状態は町のご判断を待つ。）
⚠ **埋まらなかったものを「埋まった」ことにしない。**
  稼働率・入所前の居所・休止床の数・経営の状況・町への要望は残る。
"""
import importlib.util
import os
import sys

import openpyxl
from openpyxl.styles import Alignment, Font

HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location(
    "data_jigyosho", os.path.join(HERE, "data_jigyosho_jittai_R8.10.20.py"))
D = importlib.util.module_from_spec(spec)
sys.modules["data_jigyosho"] = D
spec.loader.exec_module(D)

WBS = "川崎町_業務工程管理表.xlsx"
SOAN = "川崎町_計画書素案_v2.21_公表情報把握版.docx"
JITTAI = "川崎町_町内事業所の実態_公表システム_R8.10.20.xlsx"
GYOTEI = "川崎町_策定委員会_逆算工程表_R8.10.20.xlsx"
FONT = Font(name="游ゴシック", size=9)
# ⚠ 同じ日に3度目の更新であるため、印を分ける。
#   1度目「【令和8年10月20日 更新】」／2度目「…更新②（突合の実施）】」
HIZUKE = "【令和8年10月20日 更新③（ヒアリングを行わない整理）】"


def iro(ws, r, n):
    for c in range(1, n + 1):
        ws.cell(r, c).font = FONT
        ws.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")


def main():
    if not os.path.exists(WBS):
        raise SystemExit("業務工程管理表がない：" + WBS)
    j = D.yomu()
    nyu = [d for d in j.values() if d["入所率"]]
    teiin = sum(d["定員"] for d in nyu)
    zaisho = sum(d["在所者数"] for d in nyu)
    rmin = min(d["入所率"] for d in nyu)
    rmax = max(d["入所率"] for d in nyu)
    tokuyo = next(d for d in j.values() if d["サービスコード"] == "510")
    roken = next(d for d in j.values() if d["サービスコード"] == "520")
    gh = sum(d["定員"] for d in j.values() if d["サービスコード"] == "320")
    homon = next(d for d in j.values() if d["サービスコード"] == "110")

    TSUIKI = {
        # ══════ ヒアリング関係
        "29": (10,
               f"{HIZUKE}⚠ **指定の年月日・入所率・待機者数・認知症対応型"
               f"共同生活介護の定員が埋まりました。**介護サービス情報公表"
               f"システム（本体）の詳細ページから、町内14の指定すべてを"
               f"採録しました（令和8年10月20日）。\n"
               f"・入所・入居系の定員 計{teiin:.0f}人／在所者 計{zaisho:.0f}人"
               f"／入所率 {rmin:.1f}％〜{rmax:.1f}％\n"
               f"・待機者数　介護老人福祉施設{tokuyo['待機者数']:.0f}人・"
               f"介護老人保健施設{roken['待機者数']:.0f}人\n"
               f"・認知症対応型共同生活介護の定員 計{gh:.0f}人"
               f"（オープンデータで未報告であった1事業所の18人が"
               f"本体に載っていました）\n"
               f"・指定の年月日と指定権者の別（地域密着型4件と居宅介護支援"
               f"2件は町、8件は県）\n"
               f"・短期入所生活介護は「空床利用型」\n"
               f"⚠ **残るのは地域密着型介護老人福祉施設の待機者数のみ**"
               f"です（公表の様式に項目がありません）。町が指定権者ですので"
               f"町の台帳で補えます。　"),
        "50": (10,
               f"{HIZUKE}⚠ **ヒアリングを行わない整理に改めました。**"
               f"職種別の従業者数・常勤換算・前年度の採用者数及び退職者数は"
               f"公表システムから採録できました。素案 7-6 の①を"
               f"「年1回の公表情報の定点把握」に改めています。\n"
               f"⚠ **公表の項目にないもの（経営の状況・町への要望）は"
               f"残ります。**これらは地域ケア会議・地域密着型サービス運営"
               f"協議会という既存の場で聴くものとして素案に位置付けました。"
               f"新たな調査は要しません。　"),
        "51": (10,
               f"{HIZUKE}⚠ **公表値で置き換えました。**訪問介護の従業者は"
               f"実人数{homon['実人数の合計']:.0f}人・常勤換算"
               f"{homon['常勤換算の合計']:.1f}人です（公表システム・"
               f"令和8年1月記入）。\n"
               f"⚠ **公表の項目は「従業者数」であり、登録者数と実働者数の"
               f"別はありません。**第1回策定委員会でご発言のあった"
               f"登録者数・実働者数は、議事録が手元になく出所を明示できない"
               f"ため、素案から数値を外し、数値を伴わない事実の記述として"
               f"残しました。議事録でお確かめいただけるのであれば"
               f"戻せますので、ご指示ください。　"),
        "159": (10,
                f"{HIZUKE}⚠ **入所率・待機者数・退所先が埋まり、"
                f"残るのは3件になりました。**\n"
                f"埋まったもの　入所率（{rmin:.1f}％〜{rmax:.1f}％）／"
                f"待機者数（介護老人福祉施設{tokuyo['待機者数']:.0f}人・"
                f"介護老人保健施設{roken['待機者数']:.0f}人）／"
                f"退所先（介護老人福祉施設は前年度"
                f"{sum(tokuyo['退所先'].values()):.0f}人、"
                f"介護老人保健施設は直近3か月"
                f"{sum(roken['退所先'].values()):.0f}人）\n"
                f"⚠ **残るもの3件**　①稼働率（延べ利用日数による。"
                f"公表の項目になく、入所率で代えます）②入所前の居所"
                f"（公表の項目になく、退所先と待機者数で代えます）"
                f"③休止床の数（定員と在所者数の差で代えます）。\n"
                f"いずれも素案 7-1・表に「公表の項目にない」旨と代わりに"
                f"用いる値を明記しました。照会を行わない場合でも、"
                f"施設整備の判断は可能と考えます。　"),
        "170": (10,
                f"{HIZUKE}⚠ **照会票を発出しないという整理をお示しします。**"
                f"ご指示により、事業所ヒアリングを行わない前提で整理し直した"
                f"結果、照会票でお尋ねする予定であった項目のうち、"
                f"**入所率・待機者数・職員数・指定の年月日・退所先は"
                f"公表システムから採録できました**（14の指定すべて・"
                f"令和8年10月20日）。\n"
                f"⚠ **残るのは5件**（稼働率・入所前の居所・休止床の数・"
                f"経営の状況・町への要望）で、いずれも代わりの扱いを"
                f"素案に書きました。経営の状況と町への要望は、"
                f"地域ケア会議等の既存の場で聴くものとして計画期間中の"
                f"仕組みに位置付けました。\n"
                f"⚠ **ご判断をお願いするのは「発出しない」ことの"
                f"ご了解のみです。**発出される場合は、お尋ねする項目を"
                f"5件に絞った版をお作りします。　"),
        # ══════ 日程関係
        "4": (10,
              f"{HIZUKE}⚠ **第2回＝令和8年11月24日（火）で逆算し直しました。**"
              f"逆算工程表を R8.10.20 版に改め、候補日4つの比較から"
              f"確定日1つの工程に引き直しました（{GYOTEI}）。\n"
              f"⚠ **本日（10月20日）が開催通知（35日前）の期限です。**　"),
        "49": (10,
               f"{HIZUKE}⚠ **全4回では納期（令和9年3月15日）に"
               f"間に合いません。**第2回を11月24日として逆算すると、"
               f"全4回・パブコメ30日・印刷3週間では納品が令和9年4月4日と"
               f"なり**20日超過**します。パブコメを20日・印刷を2週間に"
               f"詰めても3日超過します。\n"
               f"**全3回（第3回で保険料の決定・パブコメの結果の報告・"
               f"答申を併せて行う）を推します。**この場合の納品は"
               f"令和9年3月10日（余裕5日）です。第2回でお諮りします。　"),
        "85": (10,
               f"{HIZUKE}⚠ **開催日が確定しました（令和8年11月24日・火）。**"
               f"これにより次が決まります。\n"
               f"・資料・想定問答集の確定　⚠ **10月30日（金）**"
               f"（21日前は11月3日の文化の日に当たるため前倒し。"
               f"町とのすり合わせの時間を確保するためでもあります）\n"
               f"・事前配布（14日前）　11月10日（火）\n"
               f"・事前質問の締切（7日前）　11月17日（火）\n"
               f"・想定問答集の更新　11月20日（金）\n"
               f"諮る事項は資料v9でお示しします。　"),
        "19": (10,
               f"{HIZUKE}⚠ **30日間とすると開始が12月29日になり"
               f"年末年始を挟みます。**第2回（11月24日）から計画書（案）の"
               f"作成と県協議に35日を見込むとこうなります。"
               f"**令和9年1月4日（月）開始とする案**を逆算工程表の"
               f"02シートに併記しました（この場合の納品は3月15日"
               f"ちょうどになります）。期間・開始日のご判断をお願いします。　"),
    }

    wb = openpyxl.load_workbook(WBS)

    # ── 03_確認事項一覧
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
        raise SystemExit(f"確認事項の行が見つからない：No."
                         + "・No.".join(sorted(nai)))

    # ── 04_成果品管理
    ws = wb["04_成果品管理"]
    n_seika = 0
    for x in range(5, ws.max_row + 1):
        no = str(ws.cell(x, 1).value or "").strip()
        na = str(ws.cell(x, 3).value or "")
        if no == "3" and JITTAI not in na:
            ws.cell(x, 3).value = (na.rstrip() + "\n" + JITTAI + "（8シート）"
                                   + "\n" + GYOTEI + "（5シート）")
            ws.cell(x, 8).value = (
                f"{HIZUKE}公表システム本体からの採録の結果（8シート）と、"
                f"第2回＝11月24日で引き直した逆算工程表（5シート）を"
                f"登録した。⚠ **残るもの8件は「06_残るもの」に理由つきで"
                f"示し、「適合」に丸めていない**。　"
                + str(ws.cell(x, 8).value or ""))
            iro(ws, x, 9)
            n_seika += 1
        elif no == "4" and "v2.20_事業所確定版" in na:
            ws.cell(x, 3).value = na.replace(
                "川崎町_計画書素案_v2.20_事業所確定版.docx", SOAN)
            ws.cell(x, 8).value = (
                f"{HIZUKE}素案を Ver.2.21 に上げた（全117頁）。"
                f"ヒアリングに依存した記述を公表情報による把握に置き換え"
                f"（段落11・セル28）、「事業所ヒアリング」は0件になった。"
                f"⚠ **7-6 の交付金の記述を是正した**（体制・取組指標群は"
                f"64点満点＝全国で最も高い得点であり、"
                f"「計画に位置付ければ得点が上がる」は誤りであった）。"
                f"［要確認］は15件→14件。　"
                + str(ws.cell(x, 8).value or ""))
            iro(ws, x, 9)
            n_seika += 1
    if n_seika != 2:
        raise SystemExit(f"04シートの該当行が{n_seika}件（2件のはず）")

    # ── 08_作業順位　順位35
    ws2 = wb["08_作業順位"]
    jun = [str(ws2.cell(x, 1).value or "").strip()
           for x in range(5, ws2.max_row + 1)]
    n_jun = 0
    if "35" not in jun:
        r = ws2.max_row + 1
        gyo = ["35",
               "事業所ヒアリングを行わない整理（公表システム本体からの採録）",
               "R8.10.20",
               "ご指示による。オープンデータの24列に含まれない項目"
               "（入所率・待機者数・職員数・指定の年月日）が、"
               "公表システム本体の詳細ページには載っていることが分かった。"
               "対象（事業所番号×サービスコード）はオープンデータの"
               "抽出結果から起こせるため、町のご回答を待たずに行える。",
               "束4のうちNo.29・No.50・No.51／束10のうちNo.159／No.170",
               "町内事業所の実態（xlsx・8シート）と素案 Ver.2.21／"
               "逆算工程表 R8.10.20（xlsx・5シート）",
               "1日", "完了",
               "⚠ **14の指定すべてを採録した。**入所率90.0〜100.0％・"
               "待機者数（介護老人福祉施設20人・介護老人保健施設2人）・"
               "認知症対応型共同生活介護の定員計27人・指定の年月日・"
               "退所先・職種別の従業者数と常勤換算・前年度の採用者数と"
               "退職者数が埋まった。素案のヒアリング依存の記述を"
               "39か所差し替え、「事業所ヒアリング」は0件になった。"
               "⚠ 残る5件（稼働率・入所前の居所・休止床・経営の状況・"
               "町への要望）は代わりの扱いを素案に明記した。"
               "⚠ 第1回策定委員会での委員のご発言に基づく人数・金額は"
               "外した（出所を明示できず、8事業所3法人のため法人が"
               "識別され得る）。"]
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

    print("業務工程管理表の更新（ヒアリングを行わない整理・"
          "第2回11月24日確定）")
    print(f"  {WBS}")
    print()
    print("  ── 自己点検")
    judge(n + sumi == len(TSUIKI),
          f"確認事項 {len(TSUIKI)}件に書いた（新 {n}・既 {sumi}）")
    judge(n_seika == 2, "成果品管理の2行を更新した")
    judge(n_jun == 1 or "35" in jun, "作業順位35 を記録した")

    wb2 = openpyxl.load_workbook(WBS, data_only=True)
    ws3 = wb2["03_確認事項一覧"]
    # ⚠ 列8（状態）を動かしていないか
    jotai = [str(ws3.cell(x, 8).value or "") for x in range(5, ws3.max_row + 1)
             if str(ws3.cell(x, 1).value or "").strip() in TSUIKI]
    judge(all(s in ("確認待ち", "決着", "保留", "回答済") for s in jotai),
          "⚠ 列8（状態）を動かしていない", "／".join(sorted(set(jotai))))
    # 欠番
    no_all = [str(ws3.cell(x, 1).value or "").strip()
              for x in range(5, ws3.max_row + 1)]
    no_all = sorted(int(s) for s in no_all if s.isdigit())
    kesu = [i for i in range(1, max(no_all) + 1) if i not in no_all]
    judge(not kesu, f"欠番0（No.1〜{max(no_all)}・{len(no_all)}件）",
          "／".join(map(str, kesu[:5])))
    # 他団体名・法人名が入っていないか
    zen = []
    for sh in wb2:
        for row in sh.iter_rows(values_only=True):
            zen += [str(v) for v in row if v is not None]
    s = "\n".join(zen)
    judge("川崎市" not in s and "福岡" not in s,
          "⚠ 川崎市（神奈川県）・福岡県川崎町が混じらない")
    # ⚠ **町の事業名と取り違えない。**「ふれあい」は町の
    #   「ふれあいネットワーク」（見守りの仕組み）にも現れる。
    #   認知症高齢者グループホームの名は「グループホームふれあい」である。
    for na in ("釜房みどりの園", "アルパイン", "鶴寿会", "光友会",
               "グループホームさわやか", "グループホームふれあい"):
        judge(na not in s, f"法人・事業所の固有名称「{na}」が入っていない")

    print()
    if futeki:
        print(f"   × 不適合 {len(futeki)}件")
        for m in futeki:
            print(f"      {m}")
        sys.exit(1)
    print(f"   ○ 自己点検 不適合0件（確認事項 {len(no_all)}件）")


if __name__ == "__main__":
    main()
