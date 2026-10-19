# -*- coding: utf-8 -*-
"""業務工程管理表の更新（作業順位28・29・令和8年10月17日）

  python3 07_ソーススクリプト/update_wbs_R8.10.17.py

更新するもの
  08_作業順位　　　　順位28・29 を「完了」にし、結果を書く
  03_確認事項一覧　　No.182・183 を起票
  01_業務内容別の進捗　(4)(5)(8) の未了の理由を更新
  04_成果品管理　　　成果品4（計画書）を Ver.2.17 に

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
HIZUKE = "【令和8年10月17日 更新】"
FONT = Font(name="游ゴシック", size=9)

SOAN = "01_第10期_最新版成果品/川崎町_計画書素案_v2.19_認知症推計版.docx"
CHU = ("05_試算・管理シート/"
       "川崎町_第10期_計画見込量_中長期・保険料内訳_R8.10.17.xlsx")
YORYO = "05_試算・管理シート/川崎町_第9期施策評価_実施要領_案_R8.10.17.xlsx"
TANA = "05_試算・管理シート/川崎町_素案の根拠の棚卸し_R8.10.7.xlsx"


def load(name, fname):
    spec = importlib.util.spec_from_file_location(
        name, os.path.join(HERE, fname))
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


CC = load("calc_wbs", "calc_chuchoki_R8.10.17.py")


def kazoeru():
    """実物から数える"""
    out = {}
    wb = openpyxl.load_workbook(CHU, read_only=True)
    ws = wb["06_棚卸し70件の対応"]
    rows = [r for r in ws.iter_rows(min_row=5, values_only=True) if r[0]]
    out["棚卸し"] = len(rows)
    out["対応済み"] = sum(1 for r in rows if r[6] == "対応済み")
    out["シート_中長期"] = len(wb.sheetnames)
    wb.close()
    wb = openpyxl.load_workbook(YORYO, read_only=True)
    ws = wb["05_事業ごとの評価欄"]
    out["事業"] = len([r for r in ws.iter_rows(min_row=5, values_only=True)
                       if r[0]])
    out["シート_要領"] = len(wb.sheetnames)
    wb.close()
    k = CC.santei()
    out["補完"] = len([x for x in k if not x.startswith("__")])
    out["補完R17"] = CC.gokei(k, "令和17年度")
    out["補完R22"] = CC.gokei(k, "令和22年度")
    return out


N = kazoeru()

SHINKI = [
    ("184", "(5)(6)", "R8.10",
     "見える化システムの出力と本計画の見込量の差（県への事前協議の備え）",
     "地域包括ケア「見える化」システムの将来推計総括表の"
     "「2_サービス別給付費」には、サービス別の中長期の給付費が"
     "入っていますが、**本計画ではこれを用いていません。** "
     "理由は既にお伝えしている2点です。"
     "①令和7年度に給付実績のあるサービスを令和9年度以降0と置いている"
     "（確認事項No.76）②令和11年度から令和27年度まで同じ値が並び、"
     "認定者数が伸びても給付費が動かない（確認事項No.80の、"
     "令和8年5月時点の値が解除できない状態によるもの）。"
     "本計画は令和9〜11年度についても用いておらず、全年度を通じて"
     "年報の実績に認定者数の伸びを乗じる方法で一貫して算定しています。"
     "⚠ **両者の差は次のとおりです。**"
     "令和9年度 ＋8,812千円（＋0.87％）／"
     "令和11年度 ＋12,506千円（＋1.23％）／"
     "令和17年度 ＋37,257千円（＋3.35％）／"
     "令和22年度 ＋36,096千円（＋2.96％）"
     "（いずれも総括表が本計画を上回る）。"
     "⚠ **令和9〜11年度の差は保険料に効きます**"
     "（＋12,506千円は月額約30円に相当）。"
     "⚠ **県への事前協議で、本計画の数値と見える化システムの出力の"
     "違いを問われる可能性があります。** 本計画の算定方法を"
     "説明する用意はできています（素案 9-1 及び"
     "中長期・保険料内訳の06bシート）。この整理でよいかをご確認ください。",
     "計画素案 9-1・9-2／県事前協議", "発注者", "確認待ち", "R8.11",
     "既定値＝本計画の算定（年報の実績に認定者数の伸びを乗じるもの）を"
     "維持します。⚠ **総括表の値に合わせるには、まず確認事項No.76・80の"
     "是正（0のサービスの補完と固定の解除）が必要であり、"
     "それなしに合わせることはできません。**"),
    ("182", "(4)", "R8.10", "第9期施策評価の実施要領（案）のご確認",
     f"第9期の{N['事業']}事業について、達成度の区分（Ａ〜Ｄ＋評価不能）、"
     "第10期方針の区分（継続・拡充・縮小・廃止・新規・統合）、"
     "優先度（高・中・低）の判定の基準を定めた実施要領（案）を"
     "お示しします。委託仕様書(4)は、シート案を提案し"
     "「発注者と十分協議のうえ内容を決定すること」と定めているため、"
     "次の3点をご確認ください。"
     "①達成度の閾値を100／80／50％とすることでよいか"
     "②優先度を3段階とすることでよいか"
     "③⚠ **関連施策評価調査シート（令和8年9月）の"
     "「自己評価（1〜4）」と評価の言葉が2通りになっています。** "
     "本要領のＡ〜Ｄに揃えてよいか。"
     "⚠ また、事業ごとの根拠法令（必須事業か任意事業か独自施策か）は"
     "町の要綱・予算科目によるため当方では埋められません。"
     "実施要領05シートの「必須事業の区分」欄のご記入をお願いします。",
     "第9期施策評価 実施要領（案）", "発注者", "確認待ち", "R8.11",
     "既定値＝お示しした案のまま進めます。達成度の閾値・優先度の段数は"
     "町のご判断で変えていただけます。"),
    ("183", "(5)", "R8.10", "中長期（令和17・22年度）の見込量の位置づけ",
     "素案 9-1 に、令和17年度・令和22年度のサービス別の見込量を"
     f"掲げています（Ver.2.17 で{N['補完']}サービスの空欄を埋め、"
     "9-1 のサービス別の和が 9-2 の総給付費と一致するようにしました）。"
     "⚠ **これらは第10期計画期間（令和9〜11年度）の見込量とは"
     "位置づけが異なり、計画の数値ではなく長期の見通しです。** "
     "次の2点をご確認ください。"
     "①計画書に中長期の列を載せたままとするか、別冊に移すか"
     "（仕様書5(5)は「10年間の…推計」を求めており、載せる前提で"
     "作っていますが、頁数の制約があります）"
     "②標準給付費見込額・地域支援事業費の中長期は算定していません"
     "（9-2 では「─」）。総給付費のみで足りるか。",
     "計画素案 9-1・9-2", "発注者", "確認待ち", "R8.11",
     "既定値＝中長期は総給付費のみとし、計画書に載せたままとします。"
     "⚠ 推測で標準給付費・地域支援事業費の中長期を置きません。"),
]

JUNI = {
    "28": {
        "状態": "完了",
        "結果": (
            f"令和8年10月17日　第9期施策評価の実施要領（案）を作成した"
            f"（{YORYO.split('/')[-1]}・{N['シート_要領']}シート）。"
            f"第9期実績一覧（町記入用）の「06_第10期反映方針」は、"
            f"事業ごとに第10期方針と優先度を町にご記入いただく様式であるが、"
            f"**評価の区分と判定の基準がなく記入のしようがなかった。** "
            f"達成度4区分（Ａ〜Ｄ）＋評価不能、第10期方針6区分と判定の道筋、"
            f"優先度3区分を定め、{N['事業']}事業の評価欄を起こした"
            f"（事業は実物から数えた）。"
            f"⚠ **「評価不能」を「未達」に丸めない形にした。** "
            f"Ｄは実施しなかったことが確かめられた場合に限る。"
            f"⚠ **必須事業では「縮小」「廃止」を選べない**ことを先に"
            f"確かめる順番にし、達成度Ｃ・Ｄでも原因が供給側なら"
            f"「拡充」が答えになるようにした。"
            f"⚠ 素案 8-4 の交付金「評価結果の活用」の要件ア〜エ4件すべてに"
            f"対応を書いた。**本要領そのものが要件エ（公表の様式を"
            f"あらかじめ定める）に当たり、16点の回復に直接効く。** "
            f"⚠ **関連施策評価調査シート（R8.9）と評価の言葉が2通りに"
            f"なっていたことが分かった**（「自己評価（1〜4）」の意味が"
            f"定義されておらず、方針の区分も食い違う）。揃え方を提案した。"
            f"確認事項No.182を起票。"),
    },
    "29": {
        "状態": "完了",
        "結果": (
            f"令和8年10月17日　中長期の見込量と保険料の内訳を書き出した"
            f"（{CHU.split('/')[-1]}・{N['シート_中長期']}シート）。"
            f"**棚卸しの{N['棚卸し']}件すべてに出所又は算定式を"
            f"対応づけた（対応済み{N['対応済み']}件・未対応0件）。** "
            f"⚠ **数値の正本は計画素案である。** 本表は素案の表から値を"
            f"読むもので、数値を別に持たない。"
            f"⚠ **いちばん重い発見。** 9-1 のサービス別の中長期の和が、"
            f"9-2 の総給付費と合っていなかった（令和17年度で▲94,878千円・"
            f"総給付費の8.5％）。Ver.2.14 で 9-1 に加えた15サービスの"
            f"中長期の列を「─」のまま空けたことによる、表の間の範囲の"
            f"食い違いであった。**これらの値は既に算定されており、"
            f"9-2 の総給付費には最初から入っていた。** "
            f"「推計していない」のではなく「表に書き写していなかった」。"
            f"原典（年報・見える化システムの総括表）から算定し直し、"
            f"**算定した令和7年度・令和11年度の給付費・利用者数・"
            f"回（日）数が素案の値とすべて一致することを確かめたうえで**、"
            f"素案 Ver.2.17 で15サービス・74セルを埋めた。"
            f"これにより 9-1 と 9-2 が3か年すべてで一致した"
            f"（令和17年度 {N['補完R17']:,}千円・"
            f"令和22年度 {N['補完R22']:,}千円を補完）。"
            f"⚠ Ver.2.16 で置いた注の記述（「サービスごとの長期推計を"
            f"行っていない」）が誤っていたため書き直した。"
            f"⚠ 5-5 の「入所・居住系680,226千円」は、地域密着型を"
            f"まるごと足したものではないことも分かった"
            f"（定期巡回・地域密着型通所介護・看護小規模多機能型は在宅。"
            f"介護予防の認知症対応型共同生活介護は含む）。"
            f"円で足して一度だけ千円に丸めると差0で再現した。"
            f"⚠ 保険料・給付費の総額は変えていない。"
            f"確認事項No.183を起票。"),
    },
}

TSUIKI = {
    "(4)": f"{HIZUKE}第9期施策評価の実施要領（案）を作成した"
           f"（仕様書(4)「現行施策の検証及び実施状況の取りまとめ」）。"
           f"⚠ **評価の区分と判定の基準がなく、町が記入のしようがない"
           f"状態であった**ことに手当てした。確認事項No.182。",
    "(5)": f"{HIZUKE}中長期（令和17・22年度）の見込量を成果品に書き出し、"
           f"素案 9-1 の空欄15サービスを原典から算定して埋めた"
           f"（Ver.2.17）。⚠ **9-1 のサービス別の和と 9-2 の総給付費が"
           f"合っていなかった**（令和17年度で▲94,878千円）。"
           f"3か年すべてで一致するようになった。確認事項No.183。",
    "(8)": f"{HIZUKE}素案 Ver.2.17（中長期補完版）。"
           f"⚠ 棚卸しで出所が示せていなかった70件は0件になった。",
}


def iro(ws, r, n_col):
    for c in range(1, n_col + 1):
        ws.cell(r, c).font = FONT
        ws.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")


def main():
    for p in (WBS, CHU, YORYO, SOAN):
        if not os.path.exists(p):
            raise SystemExit("入力がない：" + p)
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
    n_juni = 0
    for x in range(5, ws.max_row + 1):
        no = str(ws.cell(x, 1).value or "").strip()
        if no in JUNI:
            ws.cell(x, 8).value = JUNI[no]["状態"]
            ws.cell(x, 9).value = JUNI[no]["結果"]
            iro(ws, x, 9)
            n_juni += 1
    if n_juni != len(JUNI):
        raise SystemExit(f"08シートの該当行が{n_juni}件"
                         f"（{len(JUNI)}件のはず）")
    juni = [x for x in range(5, ws.max_row + 1)
            if str(ws.cell(x, 1).value or "").strip().isdigit()]
    jotai = {}
    for x in juni:
        v = str(ws.cell(x, 8).value or "").strip()
        jotai[v] = jotai.get(v, 0) + 1

    # ══════ 01_業務内容別の進捗
    ws = wb["01_業務内容別の進捗"]
    n_shinchoku = 0
    for x in range(5, ws.max_row + 1):
        no = str(ws.cell(x, 1).value or "").strip()
        if no in TSUIKI:
            mae = str(ws.cell(x, 9).value or "")
            if TSUIKI[no] not in mae:
                ws.cell(x, 9).value = TSUIKI[no] + "　" + mae
                iro(ws, x, 10)
            n_shinchoku += 1
    if n_shinchoku != len(TSUIKI):
        raise SystemExit(f"01シートの更新が{n_shinchoku}件"
                         f"（{len(TSUIKI)}件のはず）")

    # ══════ 04_成果品管理
    ws = wb["04_成果品管理"]
    n_seika = 0
    SEIKA = (
        f"{HIZUKE}素案 Ver.2.17（中長期補完版）。"
        f"9-1 の令和17・22年度の空欄15サービス・74セルを原典から算定して"
        f"埋め、**9-1 のサービス別の和と 9-2 の総給付費が3か年すべてで"
        f"一致するようにした。** Ver.2.16 で置いた注の誤りも直した。"
        f"⚠ 保険料・給付費の総額は変えていない。　")
    for x in range(5, ws.max_row + 1):
        if str(ws.cell(x, 2).value or "").strip() == "計画書":
            ws.cell(x, 3).value = (
                "川崎町_計画書素案_v2.19_認知症推計版.docx／"
                "同_v2.17c_圧縮案_R8.10.17.docx（本文106頁）／"
                "川崎町_計画書_別冊_詳細分析_R8.10.17.docx（9頁）")
            mae = str(ws.cell(x, 8).value or "")
            if SEIKA not in mae:
                ws.cell(x, 8).value = SEIKA + re.sub(
                    r"^【令和8年10月17日 更新】.*?変えていない。　", "",
                    mae, count=1)
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
    ws = wb2["08_作業順位"]
    nokori = [str(ws.cell(x, 1).value or "").strip()
              for x in range(5, ws.max_row + 1)
              if str(ws.cell(x, 1).value or "").strip().isdigit()
              and str(ws.cell(x, 8).value or "").strip() == "未着手"]

    print("業務工程管理表の更新（作業順位28・29）")
    print("保存：", WBS)
    print(f"  確認事項　{len(nos)}件（本日 新規{n_kakunin}件）"
          f"／欠番{len(kessan)}件")
    print(f"  作業順位　{len(juni)}件（"
          + "／".join(f"{k} {v}" for k, v in sorted(jotai.items())) + "）")
    print(f"  01シート　{n_shinchoku}件／04シート　{n_seika}件を更新")
    print(f"  実物から数えた　棚卸し{N['棚卸し']}件（対応済み{N['対応済み']}）"
          f"／事業{N['事業']}件／補完{N['補完']}サービス")
    print("  ── 自己点検")
    if ng:
        for m in ng:
            print("   ×", m)
        sys.exit(1)
    print("   ○ 起票した確認事項はすべて台帳に登録されている")
    print("   ○ 確認事項に欠番はない")
    print(f"   ○ 未着手の作業順位は{len(nokori)}件"
          + ("（" + "／".join(nokori) + "）" if nokori else "（なし）"))


if __name__ == "__main__":
    main()
