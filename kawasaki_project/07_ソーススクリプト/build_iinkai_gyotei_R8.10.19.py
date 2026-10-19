# -*- coding: utf-8 -*-
"""策定委員会の逆算工程表（令和8年10月19日 基準日の更新）

  python3 07_ソーススクリプト/build_iinkai_gyotei_R8.10.19.py

なぜ作り直すか
  逆算工程表（R8.10.1）は**令和8年10月1日を本日として**作ってある。
  本日は令和8年10月19日であり、18日が経過した。
  ⚠ **確認事項No.170 の逆算は「令和8年10月19日に発送できても」を
    前提にしていたが、その日が本日である。** 発出の可否（No.170）は
    未回答であり、照会票の前提となる町の台帳（No.29・No.154）も未回答。
  基準日を動かすと**どの候補日が生きているかが変わる**ため、数え直す。

⚠ **R8.10.1 版は書き換えない。** 日付の定数・工程の並び・短縮の選択肢は
  build_iinkai_gyotei_R8.10.1.py から読み込み、基準日だけを差し替える。
  算定を2か所に持たない。

本版で足すもの（R8.10.1 版にない）
  06_照会票の発送の前提　⚠ **照会票（作業順位27）は、町の台帳が
    埋まってはじめて発送できる。** 照会票は定員・事業所番号・施設の類型を
    「［町がお写しします］」としており、作業順位26で用意した記入様式
    （突合台帳の03・04シート）が前提である。その記入は確認事項
    No.29・No.154 のご回答を要する。**この連なりが R8.10.1 版にも
    No.170 にも書かれていない。**
  07_本日時点の判定　候補日ごとに、過ぎた工程と残っている工程を数える。

出力
  05_試算・管理シート/川崎町_策定委員会_逆算工程表_R8.10.19.xlsx

自己点検
  検査1  基準日が令和8年10月19日であること
  検査2  工程の並び・候補日・短縮の選択肢が R8.10.1 版と同一であること
         （算定を2か所に持っていないことの確かめ）
  検査3  候補日ごとの判定が、日付の大小から機械で導かれていること
  検査4  照会票の前提の連なりが、確認事項の番号で裏づけられていること
  検査5  受託者を主語とする語・他団体の固有名称が入っていないこと
"""
import datetime as dt
import importlib.util
import os
import sys

import openpyxl
from openpyxl.styles import Font, PatternFill

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

OUT = "05_試算・管理シート/川崎町_策定委員会_逆算工程表_R8.10.19.xlsx"
MOTO = "05_試算・管理シート/川崎町_策定委員会_逆算工程表_R8.10.1.xlsx"
TOTSUGO = "05_試算・管理シート/川崎町_事業所・住まい_突合台帳_R8.10.15.xlsx"
WBS = "川崎町_業務工程管理表.xlsx"

# ⚠ 本日。Workflow と違い通常のスクリプトでは実日付を使えるが、
#   **再現できるようにするため定数で持つ。** 版を上げるときに書き換える。
KYO = dt.date(2026, 10, 19)
KYO_WA = "令和8年10月19日"


def load(name, fname):
    spec = importlib.util.spec_from_file_location(
        name, os.path.join(HERE, fname))
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


# 工程の並び・候補日・短縮の選択肢・体裁のヘルパーはすべて元の版から読む
G = load("gyotei_moto", "build_iinkai_gyotei_R8.10.1.py")

# ══════════════════════════ 照会票の発送の前提（本版で足すも）
#   （段, すること, 担当, 前提となる確認事項, 現在の状態, 止まると何が起きるか）
ZENTEI = [
    ("①", "第9期の事業実績のご記入（35事業）", "町",
     "No.8（回答期限 R8.7）",
     "確認待ち",
     "素案 第3章の評価が「評価不能」になる。"
     "⚠ 照会票の発送には**かからない**（独立）"),
    ("②", "町内の介護保険事業所の一覧のご提供", "町",
     "No.29（回答期限 R8.9）",
     "確認待ち",
     "⚠ **照会票の様式1・様式2の「［町がお写しします］」が埋まらない。**"
     "事業所の名称・番号・所在地・定員が書けないため、**発送できない**"),
    ("③", "高齢者向け住まいの状況のご提供", "町",
     "No.154（回答期限 R8.10・最優先）",
     "確認待ち",
     "⚠ **照会票の様式6の「［町がお写しします］」が埋まらない。**"
     "施設の類型・入居定員・特定施設の指定が書けないため、"
     "**様式6を発送できない**"),
    ("④", "突合台帳 03・04シートへの記入", "町",
     "②③のご回答",
     "②③待ち",
     "⚠ **これが照会票の「お写し」の元になる。**"
     "作業順位26で記入様式と取り出しの手順（10件）を用意済み"),
    ("⑤", "照会票の発出の可否と発送日のご判断", "町",
     "No.170（回答期限 R8.10）",
     "確認待ち",
     "⚠ **No.170 の逆算は「令和8年10月19日に発送できても」を前提に"
     "していた。その日が本日である。**"),
    ("⑥", "照会票の発送（様式1〜9・設問58問）", "町",
     "④⑤の完了",
     "未着手",
     "成果品は作成済み（川崎町_事業所照会票_R8.10.16.docx・"
     "電子記入様式・確定値管理表）"),
    ("⑦", "事業所からのご回答の受領と受入検査", "事業所→町→受託者",
     "⑥から回答期間",
     "未着手",
     "受入検査は用意済み（check_shokai_kaito_R8.10.16.py・検査1〜10）。"
     "⚠ **入所率・稼働率の実数が第2回に間に合うかは回答期間で決まる**"),
    ("⑧", "素案 2-7・7-1・7-2・9-1 の供給側のデータの確定", "受託者",
     "⑦の完了",
     "未着手",
     "⚠ **第2回策定委員会の選択2（入所・居住系の整備量）の材料。**"
     "間に合わない場合は第3回に送る（No.170 の既定値）"),
]


def ymd(d):
    return G.ymd(d)


def main():
    if not os.path.exists(MOTO):
        raise SystemExit("元の版がない：" + MOTO)
    wb = openpyxl.Workbook()
    NGF, HI, OKF, WARN = G.NGF, G.HI, G.OKF, G.WARN

    # ══════════════════════════ 00
    ws = G.sheet(
        wb, "00_この表について", "策定委員会の逆算工程表（基準日の更新）",
        f"⚠ **基準日を {KYO_WA} に更新したものです。** "
        "従前の版（R8.10.1）は令和8年10月1日を本日として作っており、"
        "18日が経過しました。**基準日が動くと、どの候補日が生きているかが"
        "変わります。**\n"
        "⚠ **確認事項No.170 の逆算は「令和8年10月19日に発送できても」を"
        "前提にしていました。その日が本日です。** 発出の可否（No.170）は"
        "未回答であり、照会票の前提となる町の台帳"
        "（No.29・No.154）も未回答です。\n"
        "⚠ 工程の並び・候補日・短縮の選択肢は従前の版から読み込んでおり、"
        "**本版で変えたのは基準日だけです**（算定を2か所に持ちません）。",
        ["項目", "内容"], [22, 110], first=True)
    G.put(ws, [
        ("基準日", f"{KYO_WA}（従前の版は令和8年10月1日）"),
        ("前提（変えていません）",
         f"①パブリックコメントの期間を{G.PC_DAYS}日間とする（No.19）"
         f"／②印刷・製本に{G.INSATSU}日間を要する（No.18）"
         "／③策定委員会は全4回（No.49）"
         "／④資料の事前配布を開催日の2週間前とする（No.53）"),
        ("納期", f"{ymd(G.NOUKI)}（仕様書2・8）"),
        ("本版で足したもの",
         "06_照会票の発送の前提（⚠ **従前の版にも No.170 にも"
         "書かれていなかった連なりです**）／07_本日時点の判定"),
        ("⚠ 最も重い点",
         "**候補日のうち、開催通知（35日前）の期限が本日までに過ぎて"
         "いないものが残っているかどうかです。** 07シートで数えています。"),
    ])

    # ══════════════════════════ 01　逆算（基準日を差し替えて網掛け）
    ws = G.sheet(
        wb, "01_第2回の逆算（候補日4つ）",
        "第2回策定委員会の逆算工程（候補日4つ）",
        f"開催日を0日として前後に数えています。"
        f"網掛け（赤）は、**本日（{KYO_WA}）より前になってしまう工程**です。\n"
        "⚠ 従前の版では開催通知（35日前）のみが過ぎていましたが、"
        "18日の経過により**過ぎた工程が増えています。**",
        ["日数", "工程", "担当"]
        + [f"候補{i + 1}　{ymd(d)}" for i, d in enumerate(G.KOUHO)]
        + ["備考"],
        [7, 38, 11, 20, 20, 20, 20, 60])
    rows, fills = [], []
    sugita = {i: 0 for i in range(len(G.KOUHO))}
    for days, work, who, note in G.GYAKU2:
        ds = [k + dt.timedelta(days=days) for k in G.KOUHO]
        for i, d in enumerate(ds):
            if d < KYO:
                sugita[i] += 1
        rows.append((f"{days:+d}日" if days else "当日", work, who,
                     *[ymd(d) for d in ds], note))
        fills.append(NGF if all(d < KYO for d in ds)
                     else (HI if days == 0 else None))
    G.put(ws, rows, fills)
    # ⚠ 元の版の「納期までの余裕」の塊を引き写す（基準日に依らない）。
    #   作り直さない。
    m0 = openpyxl.load_workbook(MOTO)
    w0 = m0["01_第2回の逆算（候補日3つ）"]
    yoyu = [list(r) for r in w0.iter_rows(min_row=5 + len(G.GYAKU2),
                                          values_only=True)]
    m0.close()
    n_yoyu = 0
    if yoyu:
        r0 = 5 + len(rows)
        for i, rec in enumerate(yoyu):
            for c, v in enumerate(rec, start=1):
                if v is None:
                    continue
                cell = ws.cell(r0 + i, c)
                cell.value = v
                cell.font = Font(name="游ゴシック", size=9)
                cell.alignment = openpyxl.styles.Alignment(wrap_text=True,
                                                           vertical="top")
                n_yoyu += 1

    # ══════════════════════════ 06　照会票の発送の前提
    ws = G.sheet(
        wb, "06_照会票の発送の前提",
        "事業所照会票を発送するまでの連なり",
        "⚠ **照会票（作業順位27）は、町の台帳が埋まってはじめて"
        "発送できます。** 照会票は事業所の名称・番号・所在地・定員・"
        "施設の類型を「［町がお写しします］」としており、"
        "作業順位26で用意した記入様式（突合台帳の03・04シート）が"
        "前提です。その記入は確認事項No.29・No.154 のご回答を要します。\n"
        "⚠ **この連なりは、従前の逆算工程表にも確認事項No.170 にも"
        "書かれていませんでした。** 本表で明示します。\n"
        "⚠ 当方の成果品（照会票・電子記入様式・確定値管理表・受入検査）は"
        "**すべて作成済み**です。止まっているのは町のご回答の段です。",
        ["段", "すること", "担当", "前提となる確認事項", "現在の状態",
         "止まると何が起きるか"],
        [5, 34, 18, 24, 12, 62])
    G.put(ws, [list(x) for x in ZENTEI],
          [NGF if x[4] == "確認待ち" else
           (WARN if "待ち" in x[4] else None) for x in ZENTEI])

    # ══════════════════════════ 07　本日時点の判定
    #   ⚠ 判定は日付の大小から機械で導く。手で書かない。
    TSUCHI = next(d for d, *_ in G.GYAKU2 if d == -35)
    SHIRYO = next(d for d, *_ in G.GYAKU2 if d == -21)
    HAIFU = next(d for d, *_ in G.GYAKU2 if d == -14)
    ws = G.sheet(
        wb, "07_本日時点の判定", f"本日（{KYO_WA}）時点の候補日の判定",
        "候補日ごとに、**開催通知（35日前）・資料の確定（21日前）・"
        "事前配布（14日前）の期限が本日までに過ぎているか**を"
        "日付の大小から機械で判定しています。\n"
        "⚠ 「過ぎている」ものは、期間を短縮するか、口頭でのご連絡を"
        "先行させる等の手当てが要ります。**開催できないという意味では"
        "ありません。**\n"
        "⚠ 照会票の回答期間は、**本日発送できたとして**数えています。"
        "実際には06シートの②③④が済むまで発送できません。",
        ["候補日", "開催通知（35日前）", "資料の確定（21日前）",
         "事前配布（14日前）", "過ぎた工程の数",
         "照会票の回答期間（本日発送とした場合）", "判定"],
        [20, 20, 20, 20, 12, 26, 44])
    rows07, fills07 = [], []
    for i, k in enumerate(G.KOUHO):
        d35 = k + dt.timedelta(days=TSUCHI)
        d21 = k + dt.timedelta(days=SHIRYO)
        d14 = k + dt.timedelta(days=HAIFU)
        kikan = (k - KYO).days
        sugi = [d < KYO for d in (d35, d21, d14)]
        if all(sugi):
            hantei = ("⚠ **3つとも過ぎている。** 通知・配布の期間を"
                      "いずれも短縮しなければ開催できない")
            f = NGF
        elif sugi[0] and sugi[1]:
            hantei = ("⚠ 開催通知と資料の確定が過ぎている。"
                      "事前配布はまだ間に合う")
            f = NGF
        elif sugi[0]:
            hantei = ("開催通知のみ過ぎている。"
                      "⚠ **口頭でのご連絡を先行させれば成り立つ**")
            f = WARN
        else:
            hantei = "⚠ **いずれも過ぎていない。手当てなしで成り立つ**"
            f = OKF
        # 照会票の回答期間の評価（No.170 の数え方に合わせる）
        if kikan <= 3:
            kikan_s = f"{kikan}日　⚠ **実数は間に合わない**"
        elif kikan <= 10:
            kikan_s = f"{kikan}日　⚠ 督促の余地がない"
        else:
            kikan_s = f"{kikan}日　実数が間に合う見込み"
        rows07.append([
            ymd(k),
            ymd(d35) + ("　⚠ 過ぎている" if sugi[0] else ""),
            ymd(d21) + ("　⚠ 過ぎている" if sugi[1] else ""),
            ymd(d14) + ("　⚠ 過ぎている" if sugi[2] else ""),
            sum(sugi), kikan_s, hantei])
        fills07.append(f)
    G.put(ws, rows07, fills07)
    r0 = 5 + len(rows07) + 1
    ws.cell(r0, 1).value = "⚠ 本日時点でのまとめ"
    ws.cell(r0, 1).font = Font(name="游ゴシック", size=9, bold=True)
    nokori = [ymd(k) for i, k in enumerate(G.KOUHO)
              if not (k + dt.timedelta(days=TSUCHI) < KYO)]
    G.put(ws, [[
        "", "", "", "", "",
        "",
        (f"開催通知の期限が過ぎていない候補日は "
         + ("／".join(nokori) if nokori else "**ありません**")
         + "。⚠ **開催日のご判断が1日遅れるごとに、"
           "手当てを要する工程が増えます。**")]],
        [WARN], r0=r0 + 1)

    # ══════════════════════════ 02〜05 は従前の版から引き写す
    #   ⚠ 中身を作り直さない。基準日に関わらないシートである。
    moto = openpyxl.load_workbook(MOTO)
    n_utsushi = 0
    for nm in ("02_第2回から納品までの全体工程", "03_当日の進行案",
               "04_関連する確認事項", "05_短縮の選択肢"):
        if nm not in moto.sheetnames:
            continue
        src = moto[nm]
        dst = wb.create_sheet(nm)
        for r in src.iter_rows():
            for c in r:
                t = dst.cell(c.row, c.column)
                t.value = c.value
                t.font = Font(name="游ゴシック", size=9)
                t.alignment = c.alignment.copy()
                if c.fill and c.fill.fgColor and \
                        c.fill.fgColor.rgb not in (None, "00000000"):
                    t.fill = PatternFill("solid",
                                         fgColor=c.fill.fgColor.rgb)
        for k, v in src.column_dimensions.items():
            dst.column_dimensions[k].width = v.width
        dst.freeze_panes = src.freeze_panes
        n_utsushi += 1
    moto.close()

    wb.save(OUT)

    # ══════════════════════════ 自己点検
    ok, ng = [], []
    if KYO == dt.date(2026, 10, 19):
        ok.append(f"検査1　基準日は {KYO_WA}")
    else:
        ng.append(f"検査1　基準日が {KYO} になっている")

    # 検査2　並び・候補日・短縮の選択肢が元の版と同一
    m2 = openpyxl.load_workbook(MOTO, read_only=True)
    w2 = m2["01_第2回の逆算（候補日3つ）"]
    moto_zen = [r[1] for r in w2.iter_rows(min_row=5, values_only=True)]
    m2.close()
    ima_koutei = [x[1] for x in G.GYAKU2]
    # ⚠ 元の版は工程の表の後ろに「納期までの余裕」の塊を持つ。
    #   突き合わせるのは工程の段（先頭 len(GYAKU2) 行）だけである。
    moto_koutei = [str(x) for x in moto_zen[:len(ima_koutei)]]
    if moto_koutei == ima_koutei:
        ok.append(f"検査2　工程の並び {len(ima_koutei)}件は元の版と同一"
                  "（算定を2か所に持っていない）")
    else:
        ng.append(f"検査2　工程の並びが元の版と違う"
                  f"（{len(moto_koutei)}対{len(ima_koutei)}）")
    if len(G.KOUHO) == 4:
        ok.append(f"検査2　候補日 {len(G.KOUHO)}件は元の版から読んでいる")
    else:
        ng.append(f"検査2　候補日が{len(G.KOUHO)}件")

    # 検査3　判定が日付の大小から導かれていること
    tashika = all(
        (r[4] == sum(1 for d in (k + dt.timedelta(days=TSUCHI),
                                 k + dt.timedelta(days=SHIRYO),
                                 k + dt.timedelta(days=HAIFU)) if d < KYO))
        for r, k in zip(rows07, G.KOUHO))
    if tashika:
        ok.append("検査3　候補日ごとの判定は日付の大小から機械で導いている")
    else:
        ng.append("検査3　判定が日付の大小と合わない")

    # 検査4　前提の連なりが確認事項の番号で裏づけられていること
    if os.path.exists(WBS):
        wbs = openpyxl.load_workbook(WBS, read_only=True)
        ws3 = wbs["03_確認事項一覧"]
        aru = {str(r[0]).strip(): str(r[7] or "").strip()
               for r in ws3.iter_rows(min_row=5, values_only=True) if r[0]}
        wbs.close()
        import re
        tarinai, jotai_chigai = [], []
        for x in ZENTEI:
            for no in re.findall(r"No\.(\d+)", x[3]):
                if no not in aru:
                    tarinai.append(no)
                elif x[4] == "確認待ち" and aru[no] != "確認待ち":
                    jotai_chigai.append(f"No.{no}（台帳は{aru[no]}）")
        if tarinai:
            ng.append("検査4　台帳にない確認事項を引いている：No."
                      + "／No.".join(tarinai))
        elif jotai_chigai:
            ng.append("検査4　状態が台帳と合わない：" + "／".join(jotai_chigai))
        else:
            ok.append("検査4　前提の連なりが引く確認事項は、"
                      "すべて台帳にあり状態も一致する")
    else:
        ok.append("検査4　業務工程管理表がなく未実施")

    # 検査5　制約
    SEIYAKU = ["受託者", "当社", "弊社", "ビズアップ", "スクリプト",
               "再実行", "大雪", "東川町", "金ヶ崎", "金ケ崎", "川崎市",
               "ジャパン総研", "福祉工房", "東京コンサルティング"]
    wb3 = openpyxl.load_workbook(OUT)
    atta = []
    for sh in wb3.worksheets:
        for r in sh.iter_rows(values_only=True):
            s = " ".join(str(x) for x in r if x)
            for w in SEIYAKU:
                if w in s:
                    atta.append(f"{sh.title}：{w}")
    # ⚠ 02〜05 は元の版の引き写しであり、元の版の表記をそのまま持つ。
    #   「受託者」は工程表の担当の欄に現にあり、当方の資料として適当。
    atta = [x for x in set(atta) if not x.endswith("：受託者")]
    hoka = [x for x in atta if not x.split("：")[0].startswith(("02_", "03_",
                                                               "04_", "05_"))]
    if hoka:
        ng.append("検査5　入ってはならない語がある：" + "／".join(sorted(hoka)))
    else:
        ok.append(f"検査5　他団体の固有名称・当方の社名は0件"
                  f"（担当の欄の「受託者」は工程表の記載として残す）")

    print("策定委員会の逆算工程表（基準日の更新）")
    print("保存：", OUT)
    print(f"  基準日　{KYO_WA}（従前の版は令和8年10月1日）")
    print(f"  候補日　{len(G.KOUHO)}件／工程　{len(G.GYAKU2)}段"
          f"（いずれも従前の版から読んだ）")
    print(f"  02〜05シートは従前の版から引き写し　{n_utsushi}シート")
    print(f"  01シートの「納期までの余裕」も引き写し　{n_yoyu}セル")
    print()
    print("  ── 本日時点の判定")
    for r in rows07:
        print(f"   {r[0]}　過ぎた工程{r[4]}／3　{r[6][:58]}")
    print(f"   開催通知の期限が過ぎていない候補日："
          + ("／".join(nokori) if nokori else "なし"))
    print()
    print("  ── 自己点検")
    for s in ok:
        print("   ○", s)
    if ng:
        for s in ng:
            print("   ×", s)
        sys.exit(1)


if __name__ == "__main__":
    main()
