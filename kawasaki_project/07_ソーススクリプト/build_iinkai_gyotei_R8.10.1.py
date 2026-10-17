# -*- coding: utf-8 -*-
"""策定委員会の逆算工程表（令和8年10月1日）

08_作業順位 の順位2。第2回策定委員会の開催日が確定していない
（確認事項No.85）ため、資料の事前配布日・質問の締切・第3回の日程・
保険料をいつ諮るかがいずれも決まらない。

**開催日が未確定でも、事前配布を2週間前とする逆算の形なら作れる。**
候補日を3つ置き、それぞれについて逆算した工程を示す。
決定しだい日付のみ差し替える。

  出力　05_試算・管理シート/川崎町_策定委員会_逆算工程表_R8.10.1.xlsx

**最も重いのは、第2回が11月下旬にずれるとパブリックコメントの30日間が
確保できなくなること。** 成果品の納期（令和9年3月15日）から逆算すると、
印刷・製本に3週間、第4回（答申）に2月中旬が要る。本表はこれを
候補日ごとに数えて示す。

関係する確認事項　No.4（開催日程・審議事項）／No.49（回数と答申の時期）／
No.53（資料の様式と事前配布の期限）／No.85（第2回の開催日と諮る事項）／
No.19（パブリックコメントの取扱い）／No.18（成果品の入稿期日）
"""
import datetime as dt
import sys

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUT = "05_試算・管理シート/川崎町_策定委員会_逆算工程表_R8.10.1.xlsx"
W = "月火水木金土日"

TITLE = PatternFill("solid", fgColor="1F4E78")
HEAD = PatternFill("solid", fgColor="5B9BD5")
LEAD = PatternFill("solid", fgColor="F2F2F2")
OKF = PatternFill("solid", fgColor="E2EFDA")
NGF = PatternFill("solid", fgColor="FFC7CE")
WARN = PatternFill("solid", fgColor="FCE4D6")
HI = PatternFill("solid", fgColor="DAE3F3")

# ══════════════════════════ 前提
NOUKI = dt.date(2027, 3, 15)        # 成果品の納期（仕様書8）
INSATSU = 21                        # 印刷・製本に要する日数（当方の見込み）
PC_DAYS = 30                        # パブリックコメントの期間（当方の既定値）
KOUHO = [dt.date(2026, 11, 4),
         dt.date(2026, 11, 11),
         dt.date(2026, 11, 18),
         dt.date(2026, 11, 25)]

# ══════════════════════════ 工程を短縮する選択肢
#   納品日 ＝ 開催日 ＋ T_PC_START ＋ パブコメ日数 ＋ T_3 ＋ 第3回から第4回
#            ＋ T_NYUKO ＋ 印刷・製本
T_PC_START = 35      # 第2回から計画書（案）の作成・県協議を経てパブコメ開始
T_3 = 10             # パブコメ終了から第3回（結果の報告・保険料の決定）
T_NYUKO = 10         # 最終回から入稿
AN = [
    ("案Ａ（現行の前提）", PC_DAYS, 25, INSATSU,
     "策定委員会を全4回とし、パブリックコメントを30日間とする。"),
    ("案Ｂ（パブコメを20日間）", 20, 25, INSATSU,
     "意見公募手続の期間を20日間に短縮する。"
     "国の指針は30日間以上を原則とするが、市町村の計画は"
     "実施要領の定めによる（確認事項No.19・No.136）。"),
    ("案Ｃ（委員会を全3回）", PC_DAYS, 0, INSATSU,
     "第3回で保険料の決定と答申を併せて行い、第4回を設けない。"
     "確認事項No.49のご判断による。"),
    ("案Ｄ（Ｂ＋Ｃ）", 20, 0, INSATSU,
     "パブリックコメントを20日間とし、委員会を全3回とする。"),
    ("案Ｅ（Ｂ＋印刷2週間）", 20, 25, 14,
     "パブリックコメントを20日間とし、印刷・製本を2週間で行う"
     "（確認事項No.18で入稿期日を定める際にご相談ください）。"),
]

# 第2回の逆算（開催日Dからの日数, 作業, 担当）
GYAKU2 = [
    (-35, "開催日の確定と委員への開催通知", "町",
     "⚠ 本日（令和8年10月1日）は、候補日のいずれについても"
     "この時点を過ぎています。通知の期間を短縮するか、"
     "口頭でのご連絡を先行させる必要があります。"),
    (-21, "資料・想定問答集の最終版の確定", "受託者",
     "資料v6・想定問答集v3（全39問）は作成済み。自己点検27件に適合。"
     "確定するのは日付と、第1回でいただいたご意見の反映のみ。"),
    (-14, "委員への事前配布（資料・計画骨子）", "町",
     "確認事項No.53。紙・電子のいずれによるか、"
     "当日配布のみとするかをご指示ください。"),
    (-7, "委員からの事前質問の締切", "町",
     "事前質問があれば想定問答集に反映します。"),
    (-3, "想定問答集の更新（事前質問の反映）", "受託者",
     "事前質問がない場合はv3のまま。"),
    (0, "第2回策定委員会の開催", "町・受託者",
     "諮る事項6件（03シート）。選択1〜5のご選択をいただきます。"),
    (3, "議事録（速報版）の提出", "受託者", "当日の録音又はメモによります。"),
    (7, "議事録（確定版）の提出・選択結果の素案への反映に着手", "受託者",
     "選択2・3は見込量に、選択5は保険料に直結します。"),
    (14, "素案Ver.2.8（選択結果の反映）", "受託者",
     "選択1は第5章の章立て、選択2・3は第9章9-1の見込量、"
     "選択4は第7章7-4、選択5は第9章9-3に反映します。"),
    (21, "計画書（案）の作成・宮城県への事前協議の開始", "受託者・町",
     "法第117条第5項の都道府県の意見の聴取。"),
]

# 第2回から納品までの全体工程（開催日Dからの日数, 工程, 担当, 備考）
ZENTAI = [
    (0, "第2回策定委員会", "町・受託者", "選択1〜5の方向を決める"),
    (21, "計画書（案）の作成／宮城県への事前協議の開始", "受託者・町",
     "県の意見の聴取に要する期間は県の運用による（確認事項No.49）"),
    (35, "パブリックコメントの開始", "町",
     f"期間{PC_DAYS}日間を当方の既定値とする（確認事項No.19）。"
     "実施要領は資料編25（確認事項No.136）"),
    (35 + PC_DAYS, "パブリックコメントの終了", "町", ""),
    (35 + PC_DAYS + 10, "第3回策定委員会（保険料の決定・意見の反映）",
     "町・受託者",
     "選択5（基金の取崩額）を決定。パブリックコメントの結果を報告"),
    (35 + PC_DAYS + 35, "第4回策定委員会（計画書（案）の確定・答申）",
     "町・受託者", "確認事項No.49・No.131（答申書を掲げるか）"),
    (35 + PC_DAYS + 45, "成果品の入稿", "受託者",
     f"印刷・製本に{INSATSU}日間（確認事項No.18）"),
    (35 + PC_DAYS + 45 + INSATSU, "納品（仕様書8）", "受託者",
     "調査報告書5部・計画書・概要版3,500部"),
]

# 当日の進行案（次第, 分, 内容, 諮る事項）
SHINKO = [
    ("開会・委員紹介", 5, "事務局", ""),
    ("① 第1回でいただいたご意見と訂正事項", 10,
     "訂正事項6件。誤った前提のままご議論いただかないよう冒頭に置く",
     "①訂正事項のご確認"),
    ("② 追加調査の結果－介護人材とサービス提供体制", 15,
     "第1回で最も議論が集中した論点。事業所の状況", ""),
    ("③ 第10期の重点課題（5本）", 10,
     "第1回の5分野・課題11本を5本に集約", "②追加すべき視点の有無"),
    ("④ ご選択いただきたい事項（5件）", 40,
     "選択1 施策の体系／選択2 入所・居住系の整備量／"
     "選択3 在宅の供給／選択4 介護人材／選択5 保険料（方向性のみ）",
     "③5件のご選択"),
    ("⑤ 介護保険料の見通し", 15,
     "3パターン（6,822円／6,143円／5,464円）。"
     "選択2・3・4の帰結として示す", "③のうち選択5"),
    ("⑥ 交付金の評価と取り戻せる27点", 10,
     "⚠ ご報告し、ご意見をいただくこと自体が"
     "「評価結果の活用」（16点）の要件イを満たします",
     "⑥交付金の評価結果のご報告"),
    ("⑦ 計画骨子（案）と今後の予定", 10,
     "章構成と、第3回・第4回の日程",
     "④章構成のご確認／⑤事業所ヒアリングのご意見"),
    ("閉会", 5, "事務局", ""),
]

KAKUNIN = [
    ("4", "策定委員会の開催日程・審議事項", "R8.10",
     "本表の01・02シートで候補日3つの逆算をお示しします。"),
    ("49", "策定委員会の回数と答申の確定時期", "R8.9",
     "全4回（第2回 11月・第3回 1月・第4回 2月）を前提としています。"
     "答申を行うかは確認事項No.131。"),
    ("53", "第2回委員会資料の様式と事前配布の期限", "R8.10",
     "事前配布を開催日の2週間前としています。"
     "紙・電子の別と部数をご指示ください。"),
    ("85", "第2回策定委員会の開催日と諮る事項", "R8.10",
     "諮る事項6件を03シートに整理しました。"
     "⚠ 本表で最も急ぐものです。"),
    ("19", "パブリックコメント及び条例改正に係る支援の取扱い", "R8.9",
     f"期間{PC_DAYS}日間を当方の既定値としています。"
     "実施の有無・期間・当方の支援の範囲をご指示ください。"),
    ("18", "成果品の入稿期日", "R8.12",
     f"印刷・製本に{INSATSU}日間を見込んでいます。"),
    ("136", "パブリックコメントの実施要領", "R8.11",
     "資料編25に掲げる予定です。"),
    ("131", "策定委員会が答申を行うか。諮問書・答申書を掲げるか", "R8.11",
     "資料編5に掲げる予定です。第4回の位置付けに関わります。"),
]


def ymd(d):
    return "%d年%d月%d日（%s）" % (d.year - 2018, d.month, d.day,
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
    for i, rec in enumerate(rows):
        r = r0 + i
        for c, v in enumerate(rec, start=1):
            cell = ws.cell(r, c)
            cell.value = v
            cell.font = Font(name="游ゴシック", size=9)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if fills is not None and fills[i] is not None:
                cell.fill = fills[i]
    return r0 + len(rows) - 1


def main():
    today = dt.date(2026, 10, 1)
    wb = openpyxl.Workbook()

    # ══════════════════════════ 00
    ws = sheet(wb, "00_この表について", "策定委員会の逆算工程表",
               "第2回策定委員会の開催日が確定していないため"
               "（確認事項No.85）、資料の事前配布日・質問の締切・"
               "第3回の日程・保険料をいつ諮るかがいずれも決まりません。"
               "本表は、候補日を3つ置いてそれぞれ逆算した工程を示すものです。"
               "**開催日が決まりしだい、日付のみを差し替えます。**"
               "本表の日付は当方の案であり、確定したものではありません。",
               ["項目", "内容"], [22, 110], first=True)
    rows = [
        ("作成", "ビズアップ公共コンサルティング株式会社 札幌事業所／"
                 "令和8年10月1日"),
        ("目的", "第2回策定委員会の開催日のご判断に要する材料を示す。"
                 "あわせて、開催日が後ろにずれた場合に何が成り立たなくなるかを"
                 "数えて示す。"),
        ("前提（当方の既定値）",
         f"①パブリックコメントの期間を{PC_DAYS}日間とする（確認事項No.19）"
         f"／②印刷・製本に{INSATSU}日間を要する（確認事項No.18）"
         "／③策定委員会は全4回（確認事項No.49）"
         "／④資料の事前配布を開催日の2週間前とする（確認事項No.53）"),
        ("納期", f"{ymd(NOUKI)}（仕様書2・8）"),
        ("第1回の実績", "令和8年9月2日（水）に開催済み。"
                        "議事録・想定問答集を提出済み。"),
        ("第2回の資料", "川崎町_第2回策定委員会資料_R8.11_v6.docx／"
                        "同_想定問答集_R8.11_v3.docx（全39問）。"
                        "自己点検27件はすべて適合。**作成済みです。**"),
        ("⚠ 最も重い点",
         "**第2回が11月下旬にずれると、パブリックコメントの30日間と"
         "印刷・製本の3週間を納期（令和9年3月15日）までに収められません。**"
         "01シートで候補日ごとに数えています。"),
        ("⚠ 本表で仮置きしたもの",
         "候補日（11月11日・18日・25日の各水曜日）は、"
         "第1回が水曜日に開催されたことによる当方の仮置きです。"
         "ご都合に合わせて差し替えてください。"),
    ]
    put(ws, rows)

    # ══════════════════════════ 01
    ws = sheet(wb, "01_第2回の逆算（候補日3つ）",
               "第2回策定委員会の逆算工程（候補日3つ）",
               "開催日を0日として前後に数えています。"
               "網掛けは、本日（令和8年10月1日）より前になってしまう工程です。"
               "**開催通知（35日前）は候補日のいずれについても既に過ぎており、"
               "通知の期間を短縮するか口頭でのご連絡を先行させる必要があります。**",
               ["日数", "工程", "担当"]
               + [f"候補{i+1}　{ymd(d)}" for i, d in enumerate(KOUHO)]
               + ["備考"],
               [7, 38, 11, 20, 20, 20, 60])
    rows, fills = [], []
    for days, work, who, note in GYAKU2:
        ds = [k + dt.timedelta(days=days) for k in KOUHO]
        rows.append((f"{days:+d}日" if days else "当日", work, who,
                     *[ymd(d) for d in ds], note))
        fills.append(NGF if all(d < today for d in ds)
                     else (HI if days == 0 else None))
    last = put(ws, rows, fills)

    # 納期までの余裕
    r = last + 2
    ws.cell(r, 1).value = "【納期までの余裕】"
    ws.cell(r, 1).font = Font(name="游ゴシック", size=10, bold=True)
    r += 1
    head = ["候補", "開催日", "パブコメ終了", "第4回（答申）", "入稿",
            "納品の予定日", f"納期 {ymd(NOUKI)} までの余裕"]
    for c, h in enumerate(head, start=1):
        cell = ws.cell(r, c)
        cell.value = h
        cell.font = Font(name="游ゴシック", size=9, bold=True, color="FFFFFF")
        cell.fill = HEAD
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    r += 1
    yoyu = []
    for i, k in enumerate(KOUHO, start=1):
        pc_end = k + dt.timedelta(days=T_PC_START + PC_DAYS)
        k4 = k + dt.timedelta(days=T_PC_START + PC_DAYS + T_3 + 25)
        nyuko = k4 + dt.timedelta(days=T_NYUKO)
        nouhin = nyuko + dt.timedelta(days=INSATSU)
        margin = (NOUKI - nouhin).days
        yoyu.append(margin)
        vals = (f"候補{i}", ymd(k), ymd(pc_end), ymd(k4), ymd(nyuko),
                ymd(nouhin), f"{margin:+d}日")
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(r, c)
            cell.value = v
            cell.font = Font(name="游ゴシック", size=9)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.fill = OKF if margin >= 7 else (WARN if margin >= 0 else NGF)
        r += 1
    ws.cell(r + 1, 1).value = (
        "※ 余裕が0日を下回る候補は、納期に間に合いません。"
        "パブリックコメントの期間を短縮するか、第4回を前倒しする必要があります。"
        "※ 本表の日数は当方の見込みであり、"
        "宮城県の意見の聴取に要する期間（確認事項No.49）は含んでいません。")
    ws.cell(r + 1, 1).font = Font(name="游ゴシック", size=9)
    ws.cell(r + 1, 1).alignment = Alignment(wrap_text=True, vertical="top")

    # ══════════════════════════ 02
    ws = sheet(wb, "02_第2回から納品までの全体工程",
               "第2回から納品までの全体工程",
               "候補2（令和8年11月18日）を例に、納品までの全体を示します。"
               "法令が定める手続（都道府県の意見の聴取・意見公募手続）と、"
               "印刷・製本の期間を含みます。"
               "第3回で保険料を決定し、第4回で計画書（案）を確定する前提です。",
               ["開催日からの日数", "年月日", "工程", "担当", "備考"],
               [14, 24, 44, 14, 66])
    base = KOUHO[1]
    rows = [(f"{d:+d}日" if d else "当日", ymd(base + dt.timedelta(days=d)),
             w, who, note) for d, w, who, note in ZENTAI]
    fills = [HI if d == 0 else (WARN if "パブリック" in w else None)
             for d, w, who, note in ZENTAI]
    put(ws, rows, fills)

    # ══════════════════════════ 03
    ws = sheet(wb, "03_当日の進行案",
               "第2回策定委員会 当日の進行案（2時間）",
               "資料v6の次第によります。"
               "ご選択いただきたい事項（選択1〜5）に40分を充てる配分です。"
               "⚠ 交付金の評価結果をご報告し、ご意見をいただくこと自体が"
               "「評価結果の活用」（16点・本町は令和8年度0点）の要件イを"
               "満たします。",
               ["次第", "分", "内容", "諮る事項"], [40, 6, 66, 34])
    tot = sum(x[1] for x in SHINKO)
    rows = [(a, b, c, d) for a, b, c, d in SHINKO]
    last = put(ws, rows)
    ws.cell(last + 1, 1).value = "合計"
    ws.cell(last + 1, 2).value = tot
    for c in (1, 2):
        ws.cell(last + 1, c).font = Font(name="游ゴシック", size=9, bold=True)
        ws.cell(last + 1, c).fill = HI

    # ══════════════════════════ 04
    ws = sheet(wb, "04_関連する確認事項",
               "本表に関係する確認事項",
               "業務工程管理表の03_確認事項一覧と対応します。"
               "本表はこれらが決着していない状態で作っており、"
               "決着しだい日付・前提を差し替えます。",
               ["No.", "確認事項", "回答期限", "本表での扱い"],
               [6, 48, 10, 86])
    put(ws, KAKUNIN)

    # ══════════════════════════ 05_短縮の選択肢
    ws = sheet(wb, "05_短縮の選択肢",
               "納期に収めるための短縮の選択肢",
               "案Ａ（現行の前提）では、いずれの候補日でも"
               f"納期（{ymd(NOUKI)}）に間に合いません。"
               "どこを短縮すれば収まるかを、候補日ごとに数えたものです。"
               "緑は余裕7日以上、橙は0〜6日、赤は間に合わないことを示します。"
               "**どの案によるかは町のご判断です。"
               "当方の既定値は案Ｂ（パブリックコメントを20日間）です。**",
               ["案", "パブコメ", "第3回→第4回", "印刷・製本",
                "開催日から納品までの日数"]
               + [f"候補{i+1}　{ymd(d)}" for i, d in enumerate(KOUHO)]
               + ["内容"],
               [26, 10, 13, 11, 18, 17, 17, 17, 17, 60])
    rows, fills = [], []
    best = {}
    for name, pc, g34, ins, note in AN:
        total = T_PC_START + pc + T_3 + g34 + T_NYUKO + ins
        ms = [(NOUKI - (k + dt.timedelta(days=total))).days for k in KOUHO]
        for i, m in enumerate(ms):
            if m >= 7:
                best.setdefault(i, name)
        rows.append((name, f"{pc}日間",
                     "なし（全3回）" if g34 == 0 else f"{g34}日",
                     f"{ins}日間", f"{total}日",
                     *[f"{m:+d}日" for m in ms], note))
        fills.append(None)
    last = put(ws, rows, fills)
    # 余裕の欄だけ色を付ける
    for i, (name, pc, g34, ins, note) in enumerate(AN):
        total = T_PC_START + pc + T_3 + g34 + T_NYUKO + ins
        for j, k in enumerate(KOUHO):
            m = (NOUKI - (k + dt.timedelta(days=total))).days
            cell = ws.cell(5 + i, 6 + j)
            cell.fill = OKF if m >= 7 else (WARN if m >= 0 else NGF)
    r = last + 2
    ws.cell(r, 1).value = (
        "※ 日数は当方の見込みです。宮城県の意見の聴取に要する期間"
        "（確認事項No.49）は含んでいません。\n"
        "※ 介護保険料は条例で定めるため、"
        "令和9年3月の町議会に条例改正案を諮る日程との整合も要します"
        "（確認事項No.19）。本表はこれを含んでいません。\n"
        "※ 余裕が0日を下回る組合せは、納期に間に合いません。")
    ws.cell(r, 1).font = Font(name="游ゴシック", size=9)
    ws.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(OUT)
    print("保存：", OUT)
    print("  シート", len(wb.sheetnames))
    for i, (k, m) in enumerate(zip(KOUHO, yoyu), start=1):
        mark = "○" if m >= 7 else ("△" if m >= 0 else "×")
        print(f"  {mark} 候補{i} {ymd(k)}　納期までの余裕 {m:+d}日")
    ng = [i for i, m in enumerate(yoyu, start=1) if m < 0]
    if ng:
        print("  ⚠ 案Ａ（現行の前提）で納期に間に合わない候補："
              + "、".join(f"候補{i}" for i in ng))
    print("  ── 短縮の選択肢")
    for name, pc, g34, ins, note in AN:
        total = T_PC_START + pc + T_3 + g34 + T_NYUKO + ins
        ms = [(NOUKI - (k + dt.timedelta(days=total))).days for k in KOUHO]
        print("   %-22s %3d日　" % (name, total)
              + "／".join("候補%d %+d日" % (i + 1, m)
                          for i, m in enumerate(ms)))
    # 自己点検　当日の進行の合計が2時間に収まっているか
    print(f"  当日の進行 {tot}分", "（適合）" if tot <= 120 else "（2時間を超過）")
    if tot > 120:
        sys.exit(1)


if __name__ == "__main__":
    main()
