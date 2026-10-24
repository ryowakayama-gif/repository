# -*- coding: utf-8 -*-
"""業務工程管理表の更新（基準日の更新・令和8年10月19日）

  python3 07_ソーススクリプト/update_wbs_R8.10.19.py

なぜ更新するか
  ⚠ **確認事項No.170 の逆算は「令和8年10月19日に発送できても」を
    前提に書いてあるが、その日が本日である。** ご回答がないまま
    前提の日が来たため、数え直した内容に書き換える。
  ⚠ あわせて、**照会票は町の台帳が埋まってはじめて発送できる**という
    連なり（No.29・No.154 → 突合台帳03・04 → 照会票）が
    No.170 に書かれていなかったため、書き足す。

更新するもの
  03_確認事項一覧　No.170 の「確認をお願いする内容」を本日基準に
  08_作業順位　　　順位26 の実施の結果に、線が止まっている箇所を書き足す
  04_成果品管理　　逆算工程表を R8.10.19 版に

⚠ **確認事項の番号は増やさない。** 同じご判断（発送の可否と発送日）を
  問うものであり、No.170 の中で数え直す。
"""
import datetime as dt
import importlib.util
import os
import sys

import openpyxl
from openpyxl.styles import Alignment, Font

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

WBS = "川崎町_業務工程管理表.xlsx"
FONT = Font(name="游ゴシック", size=9)
HIZUKE = "【令和8年10月19日 更新】"


def load(name, fname):
    spec = importlib.util.spec_from_file_location(
        name, os.path.join(HERE, fname))
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


G = load("gyotei_moto_u", "build_iinkai_gyotei_R8.10.1.py")
KYO = dt.date(2026, 10, 19)


def hantei():
    """候補日ごとの判定を、逆算の定数から機械で作る（手で書かない）"""
    T35 = -35
    out = []
    for k in G.KOUHO:
        d35 = k + dt.timedelta(days=T35)
        kikan = (k - KYO).days
        out.append((G.ymd(k), d35 < KYO, kikan))
    return out


H = hantei()
IKITERU = [x[0] for x in H if not x[1]]


def no170():
    bun = []
    for ymd, sugita, kikan in H:
        if kikan <= 3:
            hyo = f"回答期間{kikan}日（⚠ 実数は間に合いません）"
        elif kikan <= 10:
            hyo = f"回答期間{kikan}日（督促の余地がありません）"
        else:
            hyo = f"回答期間{kikan}日"
        bun.append(f"{ymd}＝"
                   + ("⚠ 開催通知の期限（35日前）を過ぎています／"
                      if sugita else "開催通知の期限内／")
                   + hyo)
    return (
        "作業順位27により、事業所照会票（様式1から様式9・設問58問）と"
        "電子記入様式・確定値管理表を作成しました。素案 2-7 は供給側の"
        "データを「事業所ヒアリング（令和8年9〜10月）により確定する」とし、"
        "素案 7-6 ①は年1回のヒアリングを計画期間を通じた仕組みとして"
        "位置付けています。発出の可否と発送日をご判断ください。"
        "⚠ **発送日と第2回策定委員会の開催日は同時にお決めいただく"
        "必要があります。**"
        f"\n{HIZUKE}**本日（令和8年10月19日）時点で数え直しました。** "
        "従前お示しした逆算は「令和8年10月19日に発送できても」を前提に"
        "しており、**その日が本日です。** 候補日ごとの状況は次のとおりです。"
        "　" + "／".join(bun) + "。"
        + (f"⚠ **開催通知の期限を過ぎていない候補日は "
           f"{'・'.join(IKITERU)} のみです。**"
           if IKITERU else
           "⚠ **開催通知の期限を過ぎていない候補日はありません。**")
        + "\n⚠ **さらに、照会票は町の台帳が埋まってはじめて発送できます。** "
        "照会票は事業所の名称・事業所番号・所在地・定員・施設の類型を"
        "「［町がお写しします］」としており、作業順位26で用意した記入様式"
        "（事業所・住まい 突合台帳の03・04シート）が前提です。"
        "その記入には確認事項No.29（町内の介護保険事業所一覧）及び"
        "No.154（高齢者向け住まいの状況）のご回答を要し、"
        "いずれも確認待ちです。**この連なりは従前お示しした内容に"
        "書かれていませんでした。** 逆算工程表（R8.10.19）の"
        "06シートに段を分けて示しています。"
        "\n⚠ 当方の成果品（照会票・電子記入様式・確定値管理表・"
        "ご回答の受入検査）は**すべて作成済み**です。"
        "止まっているのは町のご回答の段です。")


TSUIKI_26 = (
    f"{HIZUKE}⚠ **当初の見立てが外れた点（外向きの通信が403で遮断されて"
    f"いること）は、本日時点でも変わっていません。** 介護サービス情報"
    f"公表システム・宮城県・川崎町・厚生労働省・e-Stat のいずれにも"
    f"接続できないことを改めて確かめました（プロキシが CONNECT に対して"
    f"403 を返します）。**県の公表名簿との突合そのものは、当方の環境では"
    f"依然として行えません。**"
    f"\n⚠ **この線（事業所・住まいのデータ）は、第2回策定委員会の"
    f"選択2（入所・居住系の整備量）につながる経路の上にあります。** "
    f"連なりは No.29・No.154 のご回答 → 突合台帳03・04シートへのご記入 → "
    f"照会票の発送（No.170）→ 事業所のご回答 → 素案 2-7・7-1・7-2・9-1 の"
    f"確定、です。**先頭の2件が確認待ちのため、線全体が止まっています。**"
    f"逆算工程表（R8.10.19）の06シートに段を分けて示し、No.170 に"
    f"書き足しました。"
    f"\n⚠ 当方の側の用意（記入様式・取り出しの手順10件・照会票・"
    f"受入検査）は**すべて済んでいます。**")


def iro(ws, r, n_col):
    for c in range(1, n_col + 1):
        ws.cell(r, c).font = FONT
        ws.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")


def main():
    if not os.path.exists(WBS):
        raise SystemExit("業務工程管理表がない：" + WBS)
    wb = openpyxl.load_workbook(WBS)

    # ══════ 03_確認事項一覧　No.170 を本日基準に書き換える
    ws = wb["03_確認事項一覧"]
    n_170 = 0
    for x in range(5, ws.max_row + 1):
        if str(ws.cell(x, 1).value or "").strip() == "170":
            ws.cell(x, 5).value = no170()
            iro(ws, x, 10)
            n_170 += 1
    if n_170 != 1:
        raise SystemExit(f"No.170 の行が{n_170}件（1件のはず）")

    # ══════ 08_作業順位　順位26 に書き足す
    ws = wb["08_作業順位"]
    n_26 = 0
    for x in range(5, ws.max_row + 1):
        if str(ws.cell(x, 1).value or "").strip() == "26":
            mae = str(ws.cell(x, 9).value or "")
            if HIZUKE not in mae:
                ws.cell(x, 9).value = mae + "　" + TSUIKI_26
            iro(ws, x, 9)
            n_26 += 1
    if n_26 != 1:
        raise SystemExit(f"順位26 の行が{n_26}件（1件のはず）")

    # ══════ 04_成果品管理　記載が古くなっている2件を直す
    #   ⚠ 04シートは**仕様書8の成果品6件だけの台帳**である
    #     （逆算工程表のような内部の管理シートは載せない）。
    #   ⚠ **成果品3に中長期・保険料内訳を登録していなかった。**
    #   ⚠ **成果品4が v2.17c／別冊 R8.10.17 を指したままで、
    #     現に作った v2.18c／別冊 R8.10.17b と食い違っていた。**
    ws = wb["04_成果品管理"]
    NAOSU = {
        "3": [("川崎町_第10期_計画見込量_R8.9.15.xlsx（7シート）",
               "川崎町_第10期_計画見込量_R8.9.15.xlsx（7シート）\n"
               "川崎町_第10期_計画見込量_中長期・保険料内訳_R8.10.17.xlsx"
               "（9シート）")],
        "4": [("同_v2.17c_圧縮案_R8.10.17.docx", "同_v2.18c_圧縮案_R8.10.17.docx"),
              ("川崎町_計画書_別冊_詳細分析_R8.10.17.docx",
               "川崎町_計画書_別冊_詳細分析_R8.10.17b.docx")],
    }
    TSUIKI_04 = {
        "3": (f"{HIZUKE}中長期（令和17・22年度）の見込量と保険料の内訳を"
              f"書き出した版を登録した（⚠ **登録が漏れていた**）。"
              f"棚卸しで出所が示せていなかった70件はすべて対応済み。　"),
        "4": (f"{HIZUKE}⚠ **圧縮案と別冊の記載が古くなっていた**"
              f"（v2.17c／別冊R8.10.17 を指したまま。現に作ったのは"
              f"v2.18c／別冊R8.10.17b）。素案 Ver.2.18 に合わせて直した。　"),
    }
    n_seika = 0
    for x in range(5, ws.max_row + 1):
        no = str(ws.cell(x, 1).value or "").strip()
        if no not in NAOSU:
            continue
        na = str(ws.cell(x, 3).value or "")
        for furui, atarashii in NAOSU[no]:
            if furui in na and atarashii not in na:
                na = na.replace(furui, atarashii)
        ws.cell(x, 3).value = na
        mae_t = str(ws.cell(x, 8).value or "")
        if HIZUKE not in mae_t:
            ws.cell(x, 8).value = TSUIKI_04[no] + mae_t
        iro(ws, x, 9)
        n_seika += 1
    if n_seika != len(NAOSU):
        raise SystemExit(f"04シートの該当行が{n_seika}件"
                         f"（{len(NAOSU)}件のはず）")
    # 直した結果、古い稿を指していないことを確かめる
    furui_nokori = []
    for x in range(5, ws.max_row + 1):
        na = str(ws.cell(x, 3).value or "")
        for w in ("v2.17c_圧縮案", "別冊_詳細分析_R8.10.17.docx"):
            if w in na:
                furui_nokori.append(w)

    wb.save(WBS)

    # ══════════════════════════ 自己点検
    wb2 = openpyxl.load_workbook(WBS)
    ok, ng = [], []
    ws = wb2["03_確認事項一覧"]
    for x in range(5, ws.max_row + 1):
        if str(ws.cell(x, 1).value or "").strip() == "170":
            t = str(ws.cell(x, 5).value or "")
            if HIZUKE in t and "その日が本日です" in t:
                ok.append("検査1　No.170 が本日基準に書き換わっている")
            else:
                ng.append("検査1　No.170 が書き換わっていない")
            if "［町がお写しします］" in t and "No.29" in t and \
                    "No.154" in t:
                ok.append("検査2　No.170 に照会票の前提の連なりが"
                          "書かれている")
            else:
                ng.append("検査2　No.170 に前提の連なりが書かれていない")
    nos = sorted(int(str(ws.cell(x, 1).value).strip())
                 for x in range(5, ws.max_row + 1)
                 if str(ws.cell(x, 1).value or "").strip().isdigit())
    kessan = [i for i in range(1, max(nos) + 1) if i not in nos]
    if kessan:
        ng.append("確認事項に欠番がある：" + "／".join(map(str, kessan[:10])))
    else:
        ok.append(f"検査3　確認事項 {len(nos)}件・欠番なし"
                  f"（⚠ 番号は増やしていない）")
    ws = wb2["04_成果品管理"]
    zen04 = " ".join(str(ws.cell(x, 3).value or "")
                     for x in range(5, ws.max_row + 1))
    furui = [w for w in ("v2.17c_圧縮案", "別冊_詳細分析_R8.10.17.docx")
             if w in zen04]
    if furui:
        ng.append("検査5　成果品管理が古い稿を指している：" + "／".join(furui))
    else:
        ok.append("検査5　成果品管理は古い稿（v2.17c・別冊R8.10.17）を"
                  "指していない")
    if "中長期・保険料内訳_R8.10.17.xlsx" in zen04:
        ok.append("検査6　成果品3に中長期・保険料内訳が登録されている")
    else:
        ng.append("検査6　成果品3に中長期・保険料内訳が登録されていない")

    ws = wb2["08_作業順位"]
    for x in range(5, ws.max_row + 1):
        if str(ws.cell(x, 1).value or "").strip() == "26":
            if HIZUKE in str(ws.cell(x, 9).value or ""):
                ok.append("検査4　順位26 に本日時点の状況が書き足されている")
            else:
                ng.append("検査4　順位26 が書き足されていない")

    print("業務工程管理表の更新（基準日の更新）")
    print("保存：", WBS)
    print(f"  03シート　No.170 を本日基準に書き換え（{n_170}件）")
    print(f"  08シート　順位26 に書き足し（{n_26}件）")
    print(f"  04シート　記載が古くなっていた成果品を直した（{n_seika}件）")
    print(f"  開催通知の期限内の候補日　"
          + ("・".join(IKITERU) if IKITERU else "なし"))
    print("  ── 自己点検")
    for s in ok:
        print("   ○", s)
    if ng:
        for s in ng:
            print("   ×", s)
        sys.exit(1)


if __name__ == "__main__":
    main()
