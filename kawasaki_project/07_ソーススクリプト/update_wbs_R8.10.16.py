# -*- coding: utf-8 -*-
"""業務工程管理表（WBS）の更新（令和8年10月16日）

作業順位27（事業所への照会票の様式）の結果を反映する。

  python3 07_ソーススクリプト/update_wbs_R8.10.16.py

更新するもの
  03_確認事項一覧　　No.170〜173 を起票
  06_資料提供依頼一覧　No.39 を起票（居宅介護支援事業所の給付管理の相手方）
  08_作業順位　　　　順位27 を「完了」に
  01_業務内容別の進捗　(4)(5)(6)(7) の未了の理由を更新
  04_成果品管理　　　成果品6（電子データ）に照会票一式を加える

⚠ **件数は成果品の実物から数える。** 本スクリプトに数を書かない。

⚠ 進捗は**当方の見立て**であり、町のご判断を拘束するものではない。
  「未着手」「確認待ち」を「完了」に丸めない。
"""
import importlib.util
import os
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


B = load("shokai_wbs", "build_jigyosho_shokai_R8.10.16.py")

# ══════════════════════════ 起票する確認事項
#   （No., 業務内容, 工程, 確認事項, 内容, 止めている成果物, 確認先,
#     状態, 回答期限, 既定値）
SHINKI = [
    ("170", "(4)(5)", "R8.10", "事業所照会票の発出の可否と発送日",
     "作業順位27により、事業所照会票（様式1から様式9・設問58問）と"
     "電子記入様式・確定値管理表を作成しました。素案 2-7 は供給側のデータを"
     "「事業所ヒアリング（令和8年9〜10月）により確定する」とし、"
     "素案 7-6 ①は年1回のヒアリングを計画期間を通じた仕組みとして"
     "位置付けています。発出の可否と発送日をご判断ください。"
     "⚠ **発送日と第2回策定委員会の開催日は同時にお決めいただく必要が"
     "あります。** 当方の逆算では、令和8年10月19日に発送できても、"
     "第2回を11月4日又は11月11日に開催する場合は入所率・稼働率の実数が"
     "間に合いません（回答期間が3日以下になります）。11月18日で10日、"
     "11月25日で17日です。",
     "素案 2-7・7-1・7-2・9-1 の供給側のデータ\n"
     "第2回策定委員会の選択2（施設整備）", "発注者", "確認待ち", "R8.10",
     "既定値＝発送しない場合、素案 2-7 の表は【町確認】のまま残し、"
     "9-1 の施設整備に係る基本方針は「既存資源の稼働の状況を把握できて"
     "いない」旨を明記します。第2回策定委員会の選択2は第3回に送ります。"),
    ("171", "(4)(8)", "R8.10", "照会票の［　］4か所のご記入",
     "照会票には、町にご記入いただく箇所を［　］で4か所残しています。"
     "①文書番号と発出の年月日 ②ご提出の期限 ③ご提出先・お問い合わせ先"
     "（電話・ファクシミリ・電子メール）④個人情報の取扱いの3点"
     "（計画の策定の支援の委託先の名称・保管の期間・廃棄の方法）。"
     "⚠ **とくに④が空欄のままでは、個人情報の取扱いを示せないまま"
     "発送することになります。**",
     "事業所照会票（発送できない）", "発注者", "確認待ち", "R8.10",
     "既定値＝④は「町の個人情報保護条例の定めによる」と書き、"
     "委託先の名称は伏せた形とします。ただし条例上の扱いは町でご確認を"
     "お願いします。"),
    ("172", "(4)(6)", "R8.10", "町外の事業所の宛先の作り方",
     "本町の被保険者が現に利用している町外の事業所にも照会票をお送りする"
     "必要がありますが、⚠ **国保連の給付実績（事業所別）が未受領であり"
     "（確認事項No.12）、年報では事業所別に落ちないため、現状では町外の"
     "事業所の宛先を作れません。** 当方の案は次の2経路による代替です。"
     "①町内の居宅介護支援事業所に、給付管理の相手方の事業所の一覧を"
     "お願いする（資料提供依頼No.39）②住所地特例者台帳の入居先施設から"
     "入所・入居系の事業所を拾う。この代替によることのご了解と、"
     "①の依頼の可否をご判断ください。",
     "照会票の送付先の一覧\n素案 9-1 の広域調整", "発注者", "確認待ち",
     "R8.10",
     "既定値＝町内の事業所にのみ発送し、素案 9-1 の広域調整に"
     "「町外施設の受入余地を把握できていない」旨を明記します。"),
    ("173", "(4)(7)", "R8.11", "照会票が回収できなかった場合の素案の書き方",
     "素案 2-7 は「確認が取れない数値及び特定の事業所が識別され得る"
     "数値は、計画書（案）には掲載しません」と定めています。"
     "したがって、回収できなかった設問を素案でどう書くかを"
     "発送の前に決めておく必要があります。確定値管理表の"
     "09_未回答のときの扱い に当方の案を7件示しました"
     "（①回答のあった事業所数と全事業所数を併記する"
     "②1事業所しかない種類で未回答の場合は【町確認】のまま残す"
     "③「0件」と「記録なし」を区別する ほか）。ご判断をお願いします。",
     "素案 2-7・7-1・7-2・9-1", "発注者", "確認待ち", "R8.11",
     "既定値＝確定値管理表 09シートの当方の案によります。"),
]

# ══════════════════════════ 起票する資料提供依頼
IRAI = [
    ("39", "D", "居宅介護支援事業所の給付管理の相手方の事業所の一覧",
     "町内の居宅介護支援事業所が給付管理を行っている相手方の事業所"
     "（事業所番号・名称・所在市町村・サービスの種類）の一覧。"
     "⚠ **国保連の給付実績（事業所別）が未受領であるため、"
     "本町の被保険者が現に利用している町外の事業所の宛先を作れません。**"
     "居宅介護支援事業所は給付管理のために相手方を把握しているので、"
     "この経路が最も早い代替になります。",
     "事業所照会票の送付先（町外の事業所）。"
     "素案 9-1 の広域調整（町外施設の受入余地）",
     "発注者（町内の居宅介護支援事業所を通じて）", "R8.10下旬", "高",
     "未着手", "確認事項No.172と対応。国保連データ（依頼No.1・"
     "確認事項No.12）が届けば不要になります。"),
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
    r = max(r for r in range(5, ws.max_row + 1)
            if str(ws.cell(r, 1).value or "").strip()) + 1
    n_kakunin = 0
    for rec in SHINKI:
        if rec[0] in aru:
            continue
        for c, v in enumerate(rec, start=1):
            ws.cell(r, c).value = v
        iro(ws, r, len(rec))
        r += 1
        n_kakunin += 1
    kakunin_zen = sorted(
        int(str(ws.cell(x, 1).value).strip())
        for x in range(5, ws.max_row + 1)
        if str(ws.cell(x, 1).value or "").strip().isdigit())
    kessan = [i for i in range(1, max(kakunin_zen) + 1)
              if i not in kakunin_zen]

    # ══════ 06_資料提供依頼一覧
    ws = wb["06_資料提供依頼一覧"]
    aru = {str(ws.cell(x, 1).value or "").strip()
           for x in range(5, ws.max_row + 1)}
    r = max(x for x in range(5, ws.max_row + 1)
            if str(ws.cell(x, 1).value or "").strip()) + 1
    n_irai = 0
    for rec in IRAI:
        if rec[0] in aru:
            continue
        for c, v in enumerate(rec, start=1):
            ws.cell(r, c).value = v
        iro(ws, r, len(rec))
        r += 1
        n_irai += 1
    irai_zen = [x for x in range(5, ws.max_row + 1)
                if str(ws.cell(x, 1).value or "").strip().isdigit()]

    # ══════ 08_作業順位
    ws = wb["08_作業順位"]
    n_yoshiki = len(B.YOSHIKI)
    n_toi = len(B.SETSUMON)
    n_kakutei = len(B.KAKUTEI)
    k_min, k_max = B.kaito_kikan()
    kekka = (
        f"令和8年10月16日　川崎町_事業所照会票_R8.10.16.docx"
        f"（様式{n_yoshiki}種・設問{n_toi}問・全27頁）、"
        f"同_電子記入様式_R8.10.16.xlsx（{n_yoshiki}シート）、"
        f"川崎町_事業所照会_確定値管理表_R8.10.16.xlsx（11シート）。"
        f"⚠ 当初の見立て（第2回策定委員会で事業所ヒアリングのご了解を"
        f"いただいた日から動ける）は誤りであった。素案 2-7 は"
        f"「事業所ヒアリング（令和8年9〜10月）により確定する」と既に書いて"
        f"おり、第2回策定委員会の選択2（施設整備）は入所率・稼働率が"
        f"なければ諮れない。**照会は第2回の前に出す必要がある。**"
        f"照会票を令和8年10月19日に発送できても、第2回を11月4日・11日に"
        f"開催する場合は間に合わない（回答期間が3日以下）。"
        f"11月18日で{k_min}日、11月25日で{k_max}日。"
        f"⚠ 票の単位を2層にした（様式1＝事業所票・様式2＝サービス票）。"
        f"職員・採用・離職・経営は事業所の単位でしか数えられず、"
        f"1枚にすると併設の多い法人が同じ職員の数を何度も書くことになり"
        f"合計が実数を超える。⚠ 7地区の名称は当方の仮置きが誤っており、"
        f"素案 1-6 の表で確かめて直した。"
        f"確定する素案の箇所{n_kakutei}件。確認事項No.170〜173を起票。"
        f"資料提供依頼No.39を起票（国保連データが未受領で町外の事業所の"
        f"宛先が作れないため、居宅介護支援事業所の給付管理の相手方で代替）。"
        f"回答の受入検査（検査1〜10）を先に用意した。")
    n_juni = 0
    for x in range(5, ws.max_row + 1):
        if str(ws.cell(x, 1).value or "").strip() == "27":
            ws.cell(x, 8).value = "完了"
            ws.cell(x, 9).value = kekka
            iro(ws, x, 9)
            n_juni += 1
    if n_juni != 1:
        raise SystemExit(f"作業順位27の行が{n_juni}件（1件のはず）")
    juni = [x for x in range(5, ws.max_row + 1)
            if str(ws.cell(x, 1).value or "").strip().isdigit()]
    jotai = {}
    for x in juni:
        v = str(ws.cell(x, 8).value or "").strip()
        jotai[v] = jotai.get(v, 0) + 1

    # ══════ 01_業務内容別の進捗
    ws = wb["01_業務内容別の進捗"]
    TSUIKI = {
        "(4)": f"{HIZUKE}事業所照会票（様式{n_yoshiki}種・設問{n_toi}問）と"
               "電子記入様式・確定値管理表を作成した。"
               "⚠ 供給側のデータは依然として未受領であり、"
               "発出のご判断（確認事項No.170）を待っている。",
        "(5)": f"{HIZUKE}見込量の確からしさに関わる供給可能量を"
               "照会する様式を用意した（様式2のお受けできる上限・"
               "令和9〜11年度の各年度の見通し）。"
               "⚠ 自然体推計は供給の制約を表現できないため、"
               "回収後に見込量の調整の方針を第2回策定委員会にお諮りする。",
        "(6)": f"{HIZUKE}素案 2-7 の供給側のデータ（定員と稼働・職員数・"
               "訪問系の供給余力・町外施設の利用）を確定させる照会の"
               "様式を用意した。⚠ 町外の事業所の宛先が作れないため、"
               "居宅介護支援事業所への照会で代替する"
               "（確認事項No.172・資料提供依頼No.39）。",
        "(7)": f"{HIZUKE}照会で確定する素案の箇所を{n_kakutei}件まで"
               "洗い出した（2-7・5-1・5-3・5-4・5-5・6-3・7-1〜7-6・"
               "9-1・9-2・1-5）。回収後にこの順で直す。",
    }
    n_shinchoku = 0
    for x in range(5, ws.max_row + 1):
        no = str(ws.cell(x, 1).value or "").strip()
        if no in TSUIKI:
            mae = str(ws.cell(x, 9).value or "")
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
        if str(ws.cell(x, 2).value or "").strip() == \
                "計画書概要及び計画書案等の電子データ":
            mae = str(ws.cell(x, 3).value or "")
            ws.cell(x, 3).value = (
                mae + "\n川崎町_事業所照会票_R8.10.16.docx（全27頁）"
                "\n川崎町_事業所照会票_電子記入様式_R8.10.16.xlsx"
                f"（{n_yoshiki}シート）")
            ws.cell(x, 8).value = (
                f"{HIZUKE}事業所照会票一式を加えた。"
                "⚠ 発出のご判断（確認事項No.170）と［　］4か所のご記入"
                "（確認事項No.171）を待っている。　"
                + str(ws.cell(x, 8).value or ""))
            iro(ws, x, 9)
            n_seika += 1
    if n_seika != 1:
        raise SystemExit(f"04シートの該当行が{n_seika}件（1件のはず）")

    wb.save(WBS)

    # ══════════════════════════ 自己点検
    wb2 = openpyxl.load_workbook(WBS, data_only=False)
    ng = []
    ws = wb2["03_確認事項一覧"]
    nos = [str(ws.cell(x, 1).value or "").strip()
           for x in range(5, ws.max_row + 1)]
    for rec in SHINKI:
        if rec[0] not in nos:
            ng.append(f"確認事項No.{rec[0]}が登録されていない")
    ws = wb2["06_資料提供依頼一覧"]
    nos = [str(ws.cell(x, 1).value or "").strip()
           for x in range(5, ws.max_row + 1)]
    for rec in IRAI:
        if rec[0] not in nos:
            ng.append(f"資料提供依頼No.{rec[0]}が登録されていない")
    if kessan:
        ng.append("確認事項に欠番がある：" + "／".join(map(str, kessan[:10])))

    print("業務工程管理表の更新（作業順位27）")
    print("保存：", WBS)
    print(f"  確認事項　{len(kakunin_zen)}件（本日 新規{n_kakunin}件）"
          f"／欠番{len(kessan)}件")
    print(f"  資料提供依頼　{len(irai_zen)}件（本日 新規{n_irai}件）")
    print(f"  作業順位　{len(juni)}件（"
          + "／".join(f"{k} {v}" for k, v in sorted(jotai.items())) + "）")
    print(f"  01シート　{n_shinchoku}件を更新／04シート　{n_seika}件を更新")
    print("  ── 自己点検")
    if ng:
        for m in ng:
            print("   ×", m)
        sys.exit(1)
    print("   ○ 起票した確認事項・資料提供依頼はすべて台帳に登録されている")
    print("   ○ 確認事項に欠番はない")
    print("  ⚠ 「未着手」「確認待ち」を「完了」に丸めていません。")


if __name__ == "__main__":
    main()
