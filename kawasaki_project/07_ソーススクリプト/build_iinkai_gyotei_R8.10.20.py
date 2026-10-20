# -*- coding: utf-8 -*-
"""策定委員会の逆算工程表（第2回＝令和8年11月24日 確定・令和8年10月20日）

  python3 07_ソーススクリプト/build_iinkai_gyotei_R8.10.20.py

  R8.10.1 版は**開催日が未確定**であったため候補日4つの逆算を並べていた。
  本版は**第2回を令和8年11月24日（火）と確定**したものである。
  確定したことにより、次が決まる。

    ①資料・想定問答集を**いつまでに固めるか**（開催の21日前）
    ②事前配布と事前質問の締切
    ③第3回の時期とパブリックコメントの期間
    ④⚠ **委員会を何回にするか**（4回構成では納期に間に合わない）

  ⚠ **本日（令和8年10月20日）が開催通知（35日前）の期限である。**

  出力　05_試算・管理シート/川崎町_策定委員会_逆算工程表_R8.10.20.xlsx

関係する確認事項　No.4（開催日程・審議事項）／No.49（回数と答申の時期）／
No.53（資料の様式と事前配布の期限）／No.85（第2回の開催日と諮る事項）／
No.19（パブリックコメントの取扱い）／No.18（成果品の入稿期日）
"""
import datetime as dt
import sys

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUT = "05_試算・管理シート/川崎町_策定委員会_逆算工程表_R8.10.20.xlsx"
W = "月火水木金土日"

TITLE = PatternFill("solid", fgColor="1F4E78")
HEAD = PatternFill("solid", fgColor="5B9BD5")
LEAD = PatternFill("solid", fgColor="F2F2F2")
OKF = PatternFill("solid", fgColor="E2EFDA")
NGF = PatternFill("solid", fgColor="FFC7CE")
WARN = PatternFill("solid", fgColor="FCE4D6")

# ══════════════════════════ 前提
KAISAI2 = dt.date(2026, 11, 24)     # ⚠ 第2回の開催日（確定）
NOUKI = dt.date(2027, 3, 15)        # 成果品の納期（仕様書8）
INSATSU = 21                        # 印刷・製本に要する日数（当方の見込み）
PC_DAYS = 30                        # パブリックコメントの期間（既定値）
T_PC_START = 35      # 第2回から計画書（案）の作成・県協議を経てパブコメ開始
T_3 = 10             # パブコメ終了から第3回
T_4 = 25             # 第3回から第4回（4回構成の場合）
T_NYUKO = 10         # 最終回から入稿

# ══════════════════════════ 構成の案（回数, パブコメ, 印刷, 説明）
AN = [
    ("案Ⅰ　全3回・パブコメ30日・印刷3週間", 3, PC_DAYS, INSATSU,
     "第3回で保険料の決定・パブリックコメントの結果の報告・答申を"
     "併せて行う。⚠ **当方はこれを推します。**"),
    ("案Ⅱ　全3回・パブコメ30日・印刷2週間", 3, PC_DAYS, 14,
     "印刷・製本を2週間で行えるなら、パブリックコメントの開始を"
     "年明けに回しても間に合う。"),
    ("案Ⅲ　全4回・パブコメ30日・印刷3週間", 4, PC_DAYS, INSATSU,
     "⚠ R8.10.1 版の前提。第4回（答申）を設ける。"),
    ("案Ⅳ　全4回・パブコメ20日・印刷2週間", 4, 20, 14,
     "4回構成を維持する場合に、パブリックコメントと印刷の双方を"
     "詰めたもの。"),
]


def nissu(kaisu, pc, ins):
    g = T_4 if kaisu >= 4 else 0
    return T_PC_START + pc + T_3 + g + T_NYUKO + ins


# ══════════════════════════ 第2回に向けた逆算
#   (開催日からの日数, 作業, 担当, 備考)
GYAKU2 = [
    (-35, "開催日の確定と委員への開催通知", "町",
     "⚠ **本日（令和8年10月20日）が期限です。** 開催日が確定しましたので、"
     "本日中にお出しいただく必要があります。"),
    (-25, "資料v9・想定問答集v5の素案を当方からお送りする", "受託者",
     "⚠ **11月初旬のすり合わせに間に合わせるため、"
     "10月30日（金）までにお送りします。**"),
    (-21, "資料・想定問答集の最終版の確定", "町・受託者",
     "⚠ 21日前は11月3日（火・文化の日）に当たるため、"
     "**実務上は10月30日（金）**とします。"),
    (-14, "委員への事前配布（資料・計画骨子）", "町",
     "確認事項No.53。紙・電子のいずれによるか、"
     "当日配布のみとするかをご指示ください。"),
    (-7, "委員からの事前質問の締切", "町",
     "事前質問があれば想定問答集に反映します。"),
    (-4, "想定問答集の更新（事前質問の反映）", "受託者",
     "⚠ 3日前は11月21日（土）に当たるため、"
     "**実務上は11月20日（金）**とします。"),
    (0, "第2回策定委員会の開催", "町・受託者",
     "諮る事項。選択1〜5のご選択をいただきます。"),
    (3, "議事録（速報版）の提出", "受託者", "当日の録音又はメモによります。"),
    (7, "議事録（確定版）の提出・選択結果の素案への反映に着手", "受託者",
     "選択2・3は見込量に、選択5は保険料に直結します。"),
    (14, "素案Ver.3.0（選択結果の反映）", "受託者",
     "選択1は第5章の章立て、選択2・3は第9章9-1の見込量、"
     "選択4は第7章7-4、選択5は第9章9-3に反映します。"),
    (21, "計画書（案）の作成・宮城県への事前協議の開始", "受託者・町",
     "法第117条第5項の都道府県の意見の聴取。"),
]

# ══════════════════════════ 11月初旬までに固めるもの
#   (件, 固めるもの, 誰が, いつまでに, いま何が足りないか)
KATAMERU = [
    ("1", "第2回策定委員会資料 v9", "受託者", "10月30日（金）",
     "v7（自己点検27件 適合）から、①素案Ver.2.21の反映"
     "（公表情報による供給側の把握・認知症の人の数の推計）"
     "②ヒアリングを行わない整理③推計パターンの確定（No.186）"
     "④第3回を1回にする工程の提示を加えます。"),
    ("2", "想定問答集 v5", "受託者", "10月30日（金）",
     "v4（全39問）に、①ヒアリングを行わない理由と代わりの把握の方法"
     "②交付金の27点の所在（研修の実施）③認知症の人の数の推計の置き方"
     "④保険料の3パターンの問を加えます。"),
    ("3", "計画骨子（素案）の確定", "町・受託者", "11月10日（火）",
     "素案Ver.2.21（全117頁）を骨子として事前配布します。"
     "⚠ ［要確認］14件が残りますので、委員会でお諮りする箇所と"
     "そのまま残す箇所を分けてお示しします。"),
    ("4", "⚠ 推計パターンの決定（確認事項No.186）", "町", "10月30日（金）",
     "将来推計総括表の「0911」を正本とするかどうか。"
     "⚠ **これが決まらないと保険料の3パターンを資料に載せられません。**"),
    ("5", "⚠ 令和8年度の実績見込み値（確認事項No.79）", "町",
     "10月30日（金）",
     "R8.10.19出力の総括表で、差し替え前の値に戻っています。"
     "伸び率の基準年が動くため、保険料に直接影響します。"),
    ("6", "委員会の回数のご判断（確認事項No.49）", "町",
     "11月24日（第2回でお諮りする）",
     "⚠ **4回構成では納期（令和9年3月15日）に間に合いません。**"
     "本表02シートで日数を示します。"),
    ("7", "パブリックコメントの期間と開始日（確認事項No.19）", "町",
     "11月24日（第2回でお諮りする）",
     "⚠ 30日間とすると開始が12月29日になり年末年始を挟みます。"
     "令和9年1月4日開始とする案を02シートに併記します。"),
    ("8", "照会票を発出しないことの確認（確認事項No.170）", "町",
     "10月30日（金）",
     "⚠ **公表情報で代替できることを確かめましたので、"
     "照会票は発出しないという整理をお示しします。**"
     "残る5件（稼働率・入所前の居所・休止床・経営の状況・町への要望）の"
     "扱いもあわせてご判断ください。"),
]

KAKUNIN = [
    ("4", "策定委員会の開催日程・審議事項", "決着",
     "第2回＝令和8年11月24日（火）。第3回は02シートの案による。"),
    ("49", "策定委員会の回数と答申の確定時期", "⚠ 急ぐ",
     "⚠ 4回構成では納期に間に合いません。全3回を推します。"),
    ("53", "第2回委員会資料の様式と事前配布の期限", "R8.11.10",
     "事前配布を開催日の14日前（11月10日）としています。"),
    ("85", "第2回策定委員会の開催日と諮る事項", "決着",
     "開催日が確定しました。諮る事項は資料v9でお示しします。"),
    ("19", "パブリックコメント及び条例改正に係る支援の取扱い", "⚠ 急ぐ",
     "⚠ 30日間とすると年末年始を挟みます。開始日のご判断が要ります。"),
    ("18", "成果品の入稿期日", "R8.12",
     f"印刷・製本に{INSATSU}日間を見込んでいます。"
     "2週間で行えるなら工程に余裕が生じます。"),
    ("170", "事業所照会票の発出の可否と発送日", "⚠ 決着の見込み",
     "⚠ 公表情報で代替できることを確かめました。"
     "発出しないという整理をお示しします。"),
    ("186", "将来推計のパターン「0911」を正本とするか", "⚠ 急ぐ",
     "⚠ 保険料の3パターンの前提です。10月30日までにご判断ください。"),
]


def ymd(d):
    return "令和%d年%d月%d日（%s）" % (d.year - 2018, d.month, d.day,
                                      W[d.weekday()])


def sheet(wb, name, title, lead, head, width, first=False):
    ws = wb.active if first else wb.create_sheet()
    ws.title = name
    ws.cell(1, 1).value = title
    ws.cell(1, 1).font = Font(name="游ゴシック", size=14, bold=True,
                              color="FFFFFF")
    ws.cell(1, 1).fill = TITLE
    ws.row_dimensions[1].height = 24
    ws.cell(2, 1).value = lead
    ws.cell(2, 1).font = Font(name="游ゴシック", size=9)
    ws.cell(2, 1).fill = LEAD
    ws.cell(2, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=2, start_column=1, end_row=2,
                   end_column=max(len(head), 2))
    ws.row_dimensions[2].height = 60
    for c, (h, w) in enumerate(zip(head, width), start=1):
        cell = ws.cell(4, c)
        cell.value = h
        cell.font = Font(name="游ゴシック", size=9, bold=True, color="FFFFFF")
        cell.fill = HEAD
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "A5"
    return ws


def put(ws, rows, fills=None, r0=5):
    for i, r in enumerate(rows):
        for j, v in enumerate(r, start=1):
            cell = ws.cell(r0 + i, j)
            cell.value = v
            cell.font = Font(name="游ゴシック", size=9)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        f = fills(r) if fills else None
        if f:
            for j in range(1, len(r) + 1):
                ws.cell(r0 + i, j).fill = f


def main():
    futeki = []

    def judge(ok, mi, setsu=""):
        print(("   ○ " if ok else "   × ") + mi + (("　" + setsu)
                                                   if setsu else ""))
        if not ok:
            futeki.append(mi)

    kyou = dt.date(2026, 10, 20)
    wb = openpyxl.Workbook()

    # ══════ 00
    ws = sheet(wb, "00_この表について",
               "策定委員会の逆算工程表（第2回＝令和8年11月24日 確定）",
               "⚠ 第2回策定委員会の開催日が令和8年11月24日（火）と"
               "確定しました。開催日が決まったことにより、資料をいつまでに"
               "固めるか、第3回をいつ置くか、委員会を何回にするかが"
               "決まります。⚠ **本日（令和8年10月20日）が開催通知"
               "（35日前）の期限です。**",
               ["項", "内容"], [26, 118], first=True)
    put(ws, [
        ["第2回の開催日", ymd(KAISAI2) + "（確定・確認事項No.85）"],
        ["⚠ 本日の位置", f"開催日の{(KAISAI2 - kyou).days}日前。"
                     "開催通知（35日前）の期限に当たります。"],
        ["資料を固める期限",
         "⚠ **10月30日（金）。** 21日前は11月3日（文化の日）に"
         "当たるため前倒しします。11月初旬のすり合わせに間に合わせる"
         "ためでもあります。"],
        ["⚠ 最も重い判断",
         "**委員会を何回にするか。** 4回構成では納期"
         f"（{ymd(NOUKI)}）に間に合いません（02シート）。"],
        ["次に重い判断",
         "⚠ パブリックコメントの開始日。30日間とすると開始が"
         "12月29日になり年末年始を挟みます（02シート）。"],
        ["R8.10.1 版との違い",
         "R8.10.1 版は開催日が未確定であったため候補日4つを並べて"
         "いました。本版は確定日1つで引き直したものです。"],
    ])

    # ══════ 01　第2回に向けた逆算
    ws = sheet(wb, "01_第2回に向けた逆算",
               f"第2回策定委員会（{ymd(KAISAI2)}）に向けた逆算",
               "開催日からの日数で引いたものです。⚠ 土曜・日曜・祝日に"
               "当たる期限は前倒ししています。",
               ["開催日からの日数", "日付", "作業", "担当", "状態", "備考"],
               [12, 20, 40, 11, 12, 62])
    rows = []
    for nd, sagyo, tan, biko in GYAKU2:
        hi = KAISAI2 + dt.timedelta(days=nd)
        if hi < kyou:
            jotai = "⚠ 期限を過ぎた"
        elif hi == kyou:
            jotai = "⚠ 本日が期限"
        else:
            jotai = f"あと{(hi - kyou).days}日"
        rows.append([f"{nd:+d}日", ymd(hi), sagyo, tan, jotai, biko])
    put(ws, rows,
        lambda r: NGF if "過ぎた" in r[4] else
        (WARN if "本日" in r[4] else None))
    n_gyaku = len(rows)

    # ══════ 02　納品までの全体工程（案の比較）
    ws = sheet(wb, "02_納品までの全体工程",
               "納品までの全体工程　— 構成の案の比較",
               "⚠ **4回構成では納期（令和9年3月15日）に間に合いません。**"
               "日数は「第2回から納品まで」の合計です。"
               "余裕が負のものは納期を超過します。",
               ["案", "回数", "パブコメ", "印刷", "第2回から納品までの日数",
                "納品予定日", "納期までの余裕", "説明"],
               [34, 7, 10, 8, 20, 20, 14, 60])
    rows = []
    for name, kaisu, pc, ins, setsu in AN:
        n = nissu(kaisu, pc, ins)
        nohin = KAISAI2 + dt.timedelta(days=n)
        yoyu = (NOUKI - nohin).days
        rows.append([name, f"全{kaisu}回", f"{pc}日", f"{ins}日",
                     f"{n}日", ymd(nohin), f"{yoyu:+d}日", setsu])
    put(ws, rows, lambda r: NGF if r[6].startswith("-") else OKF)

    # ── 案Ⅰの工程表（推す案）
    r0 = 5 + len(rows) + 2
    ws.cell(r0 - 1, 1).value = "【案Ⅰ　全3回・パブコメ30日・印刷3週間】の工程"
    ws.cell(r0 - 1, 1).font = Font(name="游ゴシック", size=10, bold=True)
    kotei = [
        (0, "第2回策定委員会", "町・受託者", "選択1〜5の方向を決める"),
        (21, "計画書（案）の作成／宮城県への事前協議の開始", "受託者・町",
         "県の意見の聴取に要する期間は県の運用による"),
        (T_PC_START, "パブリックコメントの開始", "町",
         "⚠ **12月29日に当たります。** 年末年始を挟むため、"
         "令和9年1月4日（月）開始とする案を併記します"
         "（その場合の納品は3月15日ちょうどになります）"),
        (T_PC_START + PC_DAYS, "パブリックコメントの終了", "町", ""),
        (T_PC_START + PC_DAYS + T_3,
         "第3回策定委員会（保険料の決定・意見の反映・答申）", "町・受託者",
         "⚠ 日曜に当たります。前後の平日に置いてください"),
        (T_PC_START + PC_DAYS + T_3 + T_NYUKO, "成果品の入稿", "受託者",
         f"印刷・製本に{INSATSU}日間（確認事項No.18）"),
        (nissu(3, PC_DAYS, INSATSU), "納品（仕様書8）", "受託者",
         "調査報告書5部・計画書・概要版3,500部"),
    ]
    put(ws, [[f"{nd:+d}日", "", "", "", ymd(KAISAI2 + dt.timedelta(days=nd)),
              k, tan, biko] for nd, k, tan, biko in kotei], r0=r0)

    # ══════ 03　11月初旬までに固めるもの
    ws = sheet(wb, "03_11月初旬までに固める",
               "11月初旬までに固めるもの",
               "⚠ 第2回の21日前（資料の確定）は11月3日（文化の日）に"
               "当たるため、**10月30日（金）**を期限とします。"
               "町とのすり合わせの時間を確保するためでもあります。",
               ["件", "固めるもの", "誰が", "いつまでに",
                "いま何が足りないか"], [5, 36, 11, 20, 78])
    put(ws, [list(x) for x in KATAMERU],
        lambda r: WARN if str(r[1]).startswith("⚠") else None)

    # ══════ 04　確認事項
    ws = sheet(wb, "04_確認事項",
               "本表に関わる確認事項",
               "⚠ 急ぐものを上に置いています。"
               "No.49（回数）とNo.19（パブコメ）は第2回でお諮りしますが、"
               "No.186・No.79・No.170は10月30日までにご判断が要ります。",
               ["No", "確認事項", "期限・状態", "本表での扱い"],
               [6, 46, 16, 86])
    put(ws, [list(x) for x in KAKUNIN],
        lambda r: WARN if "⚠" in str(r[2]) else OKF)

    wb.save(OUT)

    # ══════ 自己点検
    print("策定委員会の逆算工程表（第2回＝令和8年11月24日 確定）")
    print(f"  {OUT}")
    print()
    print(f"  第2回　{ymd(KAISAI2)}　（本日の{(KAISAI2 - kyou).days}日後）")
    print("  ── 構成の案")
    ng = []
    for name, kaisu, pc, ins, _ in AN:
        n = nissu(kaisu, pc, ins)
        nohin = KAISAI2 + dt.timedelta(days=n)
        yoyu = (NOUKI - nohin).days
        print(f"     {name:<38} {n:>3}日　納品 {ymd(nohin)}　"
              f"余裕 {yoyu:+d}日")
        if yoyu < 0:
            ng.append((name, yoyu))
    print()
    print("  ── 自己点検")
    judge(len(wb.sheetnames) == 5, f"5シート（実際 {len(wb.sheetnames)}）")
    judge(n_gyaku == len(GYAKU2), f"逆算 {len(GYAKU2)}行")
    judge(len(KATAMERU) == 8, f"固めるもの8件（実際 {len(KATAMERU)}）")
    # ⚠ 4回構成が納期を超えることを数で示す
    yon = [n for n, y in ng if "全4回" in n or "Ⅲ" in n or "Ⅳ" in n]
    judge(bool(ng), "⚠ 納期を超過する案がある（＝回数の判断が要る）",
          "／".join(f"{n}（{y:+d}日）" for n, y in ng))
    # 案Ⅰが間に合うこと
    n1 = nissu(3, PC_DAYS, INSATSU)
    yoyu1 = (NOUKI - (KAISAI2 + dt.timedelta(days=n1))).days
    judge(yoyu1 >= 0, f"案Ⅰ（全3回）は納期に間に合う（余裕 {yoyu1:+d}日）")
    # 資料の確定期限が本日より後であること
    kigen = dt.date(2026, 10, 30)
    judge(kigen > kyou,
          f"資料の確定期限（{ymd(kigen)}）は本日より後"
          f"（あと{(kigen - kyou).days}日）")
    # ⚠ 開催通知の期限
    tsuchi = KAISAI2 - dt.timedelta(days=35)
    judge(tsuchi == kyou,
          f"⚠ 開催通知（35日前）の期限は本日である（{ymd(tsuchi)}）")

    print()
    if futeki:
        print(f"   × 不適合 {len(futeki)}件")
        for m in futeki:
            print(f"      {m}")
        sys.exit(1)
    print("   ○ 自己点検 不適合0件")
    print("   ⚠ パブリックコメントの開始が12月29日に当たります。"
          "令和9年1月4日開始とする案を02シートに併記しています。")


if __name__ == "__main__":
    main()
