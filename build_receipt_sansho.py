# -*- coding: utf-8 -*-
"""大雪地区広域連合 第10期介護保険事業計画
　　参照条文（001685804）の受領点検.

令和8年10月15日のご連絡
  「001685804をMDファイルにまとめましたので内容確認をお願いいたします。」

━━ 本資料の位置づけ ━━

令和8年10月15日に受領した参照条文（001685804.pdf・全45頁）は、
条文の部分が画像であって文字の層がなく、当方では読み取れなかった。
同日、同じ資料の文字起こし（Markdown）のご提供を受けて内容を確認した。

**文字起こしは未校正の光学文字認識によるものである。**
ご提供資料そのものが、正式な条文の引用及び計画策定での根拠提示には
原本又は正式な法令の照合を要する旨を断っている。
このため当方は、**条・項・号の番号と構造を確かめることに用い、
条文の文言は計画素案に引用しない。**

確からしさは、新旧対照条文（当方が原典の文字の層から直接読んだもの）と
重なる範囲で突き合わせることにより担保している（04シート）。

━━ 収録されている介護保険法は未施行の版である ━━

参照条文の介護保険法は医療法等の一部を改正する法律
（令和7年法律第87号）による改正後のものであり、
その施行は公布の日から起算して2年を超えない範囲内において政令で定める日
である。「現行」として読んだ各項が本計画の策定の時点で施行されているかは
別に確かめることを要する。

シート構成
  00_この資料について
  01_収録法令と版
  02_市町村介護保険事業計画の規定
  03_都道府県介護保険事業支援計画との書き分け
  04_新旧対照条文との突合
  05_是正を要するもの
  06_新たに分かったもの
  07_地域支援事業の規定
  08_読み取りの誤りと確認できない範囲
  09_自己点検

出力
  output/第10期計画_参照条文の受領点検.xlsx

自己点検で1件でも不適合があると終了コード1で終わる。
"""

import datetime
import io
import os
import re
import sys

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import repo_paths as RP

sys.path.insert(0, RP.ROOT)
import data_sansho as SA                                     # noqa: E402
import data_jorei as JR                                      # noqa: E402

if isinstance(sys.stdout, io.TextIOWrapper):
    sys.stdout.reconfigure(encoding="utf-8")

ODIR = RP.OUTPUT
OUT = os.path.join(ODIR, "第10期計画_参照条文の受領点検.xlsx")

FONT = "游ゴシック"
NAVY, HEAD = "1F3864", "4472C4"
IN_Y, OK_G, NG_O, MID_B, GRAY = ("FFF2CC", "E2EFDA", "FCE4D6",
                                 "DEEBF7", "F2F2F2")
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = Workbook()
wb.remove(wb.active)
CHECKS = []


def chk(no, naiyo, shiki, kekka, ok):
    CHECKS.append((no, naiyo, shiki, kekka, "適合" if ok else "不適合"))
    return ok


JURYO = datetime.date(2026, 10, 15)

SH = {"about": "00", "horei": "01", "jo117": "02", "jo118": "03",
      "totsugo": "04", "zesei": "05", "shin": "06", "chiiki": "07",
      "goten": "08", "chk": "09"}


def shn(key):
    return SH[key] + "シート"


# ============================================================ 体裁
def _plain(v):
    return v.replace("**", "") if isinstance(v, str) else v


def sheet(name, title, subtitle, widths, freeze="A5"):
    ws = wb.create_sheet(name)
    ws["A1"] = title
    ws["A1"].font = Font(name=FONT, size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    ws["A2"] = _plain(subtitle)
    ws["A2"].font = Font(name=FONT, size=9)
    ws["A2"].fill = PatternFill("solid", fgColor=GRAY)
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    n = max(len(widths), 6)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n)
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 58
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    if freeze:
        ws.freeze_panes = freeze
    ws.sheet_view.showGridLines = False
    return ws


def header(ws, row, cols, height=30):
    for i, c in enumerate(cols, start=1):
        cell = ws.cell(row, i, _plain(c))
        cell.font = Font(name=FONT, size=9, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=HEAD)
        cell.alignment = Alignment(wrap_text=True, vertical="center",
                                   horizontal="center")
        cell.border = BORDER
    ws.row_dimensions[row].height = height
    return row + 1


def body(ws, row, vals, fills=None, height=22, align=None, bold=False):
    for i, v in enumerate(vals, start=1):
        cell = ws.cell(row, i, _plain(v))
        cell.font = Font(name=FONT, size=9, bold=bold)
        cell.alignment = Alignment(
            wrap_text=True, vertical="top",
            horizontal=(align or {}).get(i, "left"))
        cell.border = BORDER
        if fills and i in fills:
            cell.fill = PatternFill("solid", fgColor=fills[i])
    ws.row_dimensions[row].height = height
    return row + 1


def lead(ws, row, text, span=8):
    ws.cell(row, 1, _plain(text))
    ws.cell(row, 1).font = Font(name=FONT, size=11, bold=True, color=NAVY)
    ws.merge_cells(start_row=row, start_column=1,
                   end_row=row, end_column=span)
    ws.row_dimensions[row].height = 22
    return row + 1


def note(ws, row, text, span=8, height=None, fill=GRAY):
    ws.cell(row, 1, _plain(text))
    ws.cell(row, 1).font = Font(name=FONT, size=9)
    ws.cell(row, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(row, 1).fill = PatternFill("solid", fgColor=fill)
    ws.merge_cells(start_row=row, start_column=1,
                   end_row=row, end_column=span)
    ws.row_dimensions[row].height = height or (16 * (text.count("\n") + 3))
    return row + 1


def wareki(d):
    return "令和%d年%d月%d日" % (d.year - 2018, d.month, d.day)


# ============================================================ 00
ws = sheet(SH["about"] + "_この資料について",
           "参照条文（社会福祉法等の一部を改正する法律案）の受領点検",
           "令和8年10月15日に受領した参照条文の文字起こしを読み、"
           "当方の整理と突き合わせたものです。"
           "参照条文の原本は条文の部分が画像であって文字の層がなく、"
           "当方では読み取れませんでした。"
           "文字起こしのご提供により、介護保険法第117条の"
           "項及び号の全体を確認することができました。",
           [6, 30, 54, 24, 14, 12, 12, 12])

r = 4
r = lead(ws, r, "1　受領したもの")
r = header(ws, r, ["", "資料", "内容", "頁", "", "", "", ""])
r = body(ws, r, ["", SA.UKE[3] + "（" + SA.UKE[0] + "）", SA.UKE[4],
                 "全%d頁（本文 原資料%d頁）" % (SA.UKE[1], SA.pages())],
         height=58, fills={1: MID_B})

r += 1
r = lead(ws, r, "2　先に申し上げること")
for t in (
    "・**文字起こしは未校正の光学文字認識によるものです。**"
    "ご提供資料そのものが、正式な条文の引用及び計画策定での根拠提示には"
    "原本又は正式な法令の照合を要する旨を断っています。"
    "当方が確かめた誤認識は%d件あり、%s に掲げました。"
    % (len(SA.OCR_GOTEN), shn("goten")),
    "・このため当方は、**条、項及び号の番号と構造を確かめることに用い、"
    "条文の文言は計画素案に引用しません。**"
    "計画素案には従前どおり要旨により記載します。",
    "・確からしさは、新旧対照条文（当方が原典から直接読んだもの）と"
    "重なる範囲で突き合わせることにより確かめています"
    "（%s。%d件とも一致）。" % (shn("totsugo"), len(SA.TOTSUGO)),
    "・**介護保険法第117条の項の全体（全14項）が確認できました。**"
    "従前は改正される項のみを把握しており、"
    "改正されない項（作成の基準、他計画との整合及び調和、"
    "評価の結果の報告、計画の提出）の項番号を把握していませんでした。",
    "・**この確認により、当方の記述に%d件の誤りがあることが分かりました。**"
    "うち重いものは、医療計画との整合性の確保を"
    "市町村計画の義務として書いていたこと、及び"
    "混合型特定施設入居者生活介護に係る必要利用定員総数を"
    "第117条の事項として書いていたことの2件です（%s）。"
    % (len(SA.ZESEI), shn("zesei")),
    "・**参照条文の介護保険法は未施行の版です。**"
    "医療法等の一部を改正する法律（令和7年法律第87号）による改正後の"
    "ものであり、その施行は公布の日から起算して2年を超えない範囲内に"
    "おいて政令で定める日です。"
    "「現行」として読んだ各項が本計画の策定の時点で施行されているかは"
    "別に確かめることを要します［要確認］。",
):
    r = note(ws, r, t, height=16 * (len(t) // 68 + 2), fill="FFFFFF")

r += 1
r = lead(ws, r, "3　確認事項との関係")
r = header(ws, r, ["", "確認事項", "内容", "本資料による進み", "", "", "", ""])
for no, ken, susumi in (
    ("No.181", "介護保険法第117条ほか関係法令の条文のご提供",
     "**解決した。** 第117条の全14項及び第3項の全10号、"
     "第118条並びに第115条の45の項及び号を確認した。"
     "ただし文字起こしによるものであるため、"
     "計画素案には項番号までを記載し文言は引用しない"),
    ("No.113", "混合型特定施設の必要利用定員総数を定めるか",
     "**解決した。** 混合型特定施設入居者生活介護に係る"
     "必要利用定員総数は第118条第4項により都道府県が"
     "定めることができる事項であり、本計画の記載事項ではない"),
    ("No.178", "構成3町の高齢者福祉計画と市町村老人福祉計画の関係",
     "一部解決した。介護保険法第117条第7項により"
     "**老人福祉法第20条の8第1項**であることが確かめられた。"
     "老人福祉法そのものの条文は収録されていない"),
    ("No.185", "小規模市町村地域支援連携協働体制整備事業の対象となるか",
     "進まない。社会福祉法第106条の6は改正法により新設される規定で"
     "あるため参照条文には現れない"),
    ("No.149", "認知症基本法の条文ほか",
     "進まない。同法は収録されていない"),
    ("No.33", "第10期の政令改正の内容（所得段階）",
     "第125条第2項により第2号被保険者負担率が3年ごとに"
     "政令で定められることを確認した。"
     "政令そのものは公表を待つ"),
):
    r = body(ws, r, ["", no, ken, susumi], height=44, fills={2: MID_B},
             align={2: "center"})
    ws.merge_cells(start_row=r - 1, start_column=4, end_row=r - 1,
                   end_column=8)

r += 1
r = note(ws, r,
         "注1　本資料は受領した文字起こしを読んだ記録です。"
         "計画素案への反映は %s 及び %s のとおりです。\n"
         "注2　受領ファイルは収録していません。"
         "収録しているのは条、項及び号の番号と内容の要旨です。\n"
         "注3　受領日　%s" % (shn("zesei"), shn("shin"), wareki(JURYO)),
         height=56)


# ============================================================ 01
ws = sheet(SH["horei"] + "_収録法令と版",
           "収録法令と版（%d法令）" % len(SA.HOREI),
           "参照条文に収録されている法令と、その版です。"
           "版の日付及び施行の状況は資料に示されたものであり、"
           "当方が判定したものではありません。"
           "**7法令は未施行の版により収録されています。**",
           [6, 46, 12, 12, 12, 46, 10, 10])

r = 4
r = header(ws, r, ["", "法令", "掲載頁", "版の日付", "施行の状況",
                   "改正の根拠", "", ""])
for nm, pg, rev, jo, konkyo in SA.HOREI:
    r = body(ws, r, ["", nm, pg, rev, jo, konkyo], height=34,
             fills={5: NG_O if jo == "未施行" else OK_G},
             align={3: "center", 4: "center", 5: "center"})
    ws.merge_cells(start_row=r - 1, start_column=6, end_row=r - 1,
                   end_column=8)
r = note(ws, r,
         "注1　**介護保険法は未施行の版により収録されています。**"
         "医療法等の一部を改正する法律（令和7年法律第87号）による"
         "改正後のものであり、その施行は公布の日から起算して2年を"
         "超えない範囲内において政令で定める日です。\n"
         "注2　このため、本資料により「現行」として読んだ各項が"
         "本計画の策定の時点で施行されているかは別に確かめることを"
         "要します［要確認］。"
         "特に第117条第5項の医療法第30条の18の5第1項の協議の結果の"
         "考慮は、同改正により加えられたものとみられます。\n"
         "注3　参照条文はいずれも抄であり、"
         "改正法が参照する条文に限られています。"
         "収録されていない条文は %s のとおりです。" % shn("goten"),
         height=86)


# ============================================================ 02
ws = sheet(SH["jo117"] + "_市町村介護保険事業計画の規定",
           "介護保険法第117条（市町村介護保険事業計画）",
           "参照条文により確認した第117条の項の全体です。"
           "改正後の項は、社会福祉法等改正法により第5項が新設されることに"
           "伴い、現行第5項以降が1項ずつ繰り下がります。"
           "この繰下げは新旧対照条文による読みと一致します（%s）。"
           % shn("totsugo"),
           [6, 12, 54, 20, 18, 11, 12, 12])

r = 4
r = lead(ws, r, "1　項の全体")
r = header(ws, r, ["", "項（現行）", "内容の要旨", "義務の程度",
                   "項（改正後）", "当方の扱い", "", ""])
for ko, naiyo, gimu, kaisei, atsukai in SA.JO117:
    r = body(ws, r, ["", ko, naiyo, gimu, kaisei, atsukai], height=40,
             fills={2: MID_B,
                    6: (NG_O if atsukai == "要是正"
                        else (IN_Y if atsukai == "新" else OK_G))},
             align={2: "center", 5: "center", 6: "center"})

r += 1
r = lead(ws, r, "2　第3項の号（任意記載事項）")
r = header(ws, r, ["", "号", "内容の要旨", "", "", "", "", ""])
for go, naiyo in SA.JO117_3:
    r = body(ws, r, ["", go, naiyo], height=32, fills={2: MID_B},
             align={2: "center"})
    ws.merge_cells(start_row=r - 1, start_column=3, end_row=r - 1,
                   end_column=8)

r += 1
r = lead(ws, r, "3　そのほかの規定")
r = header(ws, r, ["", "条", "内容の要旨", "本計画に関わること", "", "",
                   "", ""])
for jo, naiyo, kakawari in SA.HOKA:
    r = body(ws, r, ["", jo, naiyo, kakawari], height=46,
             fills={2: MID_B}, align={2: "center"})
    ws.merge_cells(start_row=r - 1, start_column=4, end_row=r - 1,
                   end_column=8)

r += 1
r = note(ws, r,
         "注1　**条文の文言は収録していません。**"
         "文字起こしが未校正の光学文字認識によるものであるためです。"
         "計画素案には項番号までを記載し、内容は要旨により書きます。\n"
         "注2　第3項第9号の括弧書きにより、"
         "有料老人ホーム及び登録住宅の入居定員総数は"
         "**特定施設入居者生活介護等の指定を受けていないものに限られます。**\n"
         "注3　第2項第1号により本計画が定める必要利用定員総数は、"
         "認知症対応型共同生活介護、地域密着型特定施設入居者生活介護及び"
         "地域密着型介護老人福祉施設入所者生活介護の3つです。",
         height=86)


# ============================================================ 03
ws = sheet(SH["jo118"] + "_都道府県計画との書き分け",
           "介護保険法第118条（都道府県介護保険事業支援計画）との書き分け",
           "市町村介護保険事業計画の事項と都道府県介護保険事業支援計画の"
           "事項を取り違えていた箇所が2件あったため、"
           "都道府県の事項を別に掲げます。"
           "**混合型特定施設入居者生活介護に係る必要利用定員総数及び"
           "介護保険施設の種類ごとの必要入所定員総数は都道府県の事項です。**",
           [6, 16, 56, 22, 12, 12, 12, 12])

r = 4
r = header(ws, r, ["", "項・号", "内容の要旨", "義務の程度", "定める主体",
                   "", "", ""])
for jo, naiyo, gimu, shutai in SA.JO118:
    r = body(ws, r, ["", jo, naiyo, gimu, shutai], height=42,
             fills={2: MID_B, 5: IN_Y},
             align={2: "center", 5: "center"})

r += 1
r = lead(ws, r, "必要利用定員総数の書き分け")
r = header(ws, r, ["", "区分", "定める主体", "根拠", "義務の程度",
                   "本計画の扱い", "", ""])
for kubun, shutai, konkyo, gimu, atsukai in (
    ("認知症対応型共同生活介護", "市町村（広域連合）",
     "第117条第2項第1号", "定めるものとする", "定める"),
    ("地域密着型特定施設入居者生活介護", "市町村（広域連合）",
     "第117条第2項第1号", "定めるものとする", "定める"),
    ("地域密着型介護老人福祉施設入所者生活介護", "市町村（広域連合）",
     "第117条第2項第1号", "定めるものとする", "定める"),
    ("介護専用型特定施設入居者生活介護", "都道府県",
     "第118条第2項第1号", "定めるものとする", "本計画の記載事項ではない"),
    ("混合型特定施設入居者生活介護", "都道府県",
     "第118条第4項", "定めることができる", "本計画の記載事項ではない"),
    ("介護保険施設の種類ごとの必要入所定員総数", "都道府県",
     "第118条第2項第1号", "定めるものとする", "本計画の記載事項ではない"),
):
    r = body(ws, r, ["", kubun, shutai, konkyo, gimu, atsukai], height=24,
             fills={3: (OK_G if "市町村" in shutai else NG_O)},
             align={3: "center", 4: "center", 5: "center"})

r += 1
r = note(ws, r,
         "注1　**従前、混合型特定施設入居者生活介護に係る"
         "必要利用定員総数を第117条第3項の勘案事項として整理し、"
         "計画素案に「定めない」と記載していました。**"
         "本計画に定める根拠がないため、"
         "「都道府県介護保険事業支援計画において定めることができる事項で"
         "あり本計画の記載事項ではない」に改めます。\n"
         "注2　区域内の特定施設が混合型か介護専用型かは、"
         "第8条第21項により「入居者が要介護者、その配偶者その他"
         "厚生労働省令で定める者に限られているか」により定まります。"
         "入居の要件により判定できるため、"
         "確認の方法を北海道の名簿によるものから入居の要件によるものへ"
         "改めます。\n"
         "注3　介護老人保健施設及び介護医療院の開設の許可に当たっては、"
         "北海道知事が広域連合長に対し本計画との調整を図る見地からの"
         "意見を求めなければならないものとされています"
         "（第94条第6項ほか）。",
         height=104)


# ============================================================ 04
ws = sheet(SH["totsugo"] + "_新旧対照条文との突合",
           "新旧対照条文との突合（%d件）" % len(SA.TOTSUGO),
           "文字起こしの条、項及び号の番号が信頼できることを確かめるため、"
           "新旧対照条文（当方が原典の文字の層から直接読んだもの）と"
           "重なる範囲で突き合わせたものです。"
           "**%d件とも一致しました。**" % len(SA.TOTSUGO),
           [6, 30, 48, 34, 12, 12, 12, 12])

r = 4
r = header(ws, r, ["", "突合の対象", "新旧対照条文による読み",
                   "参照条文による読み", "判定", "", "", ""])
for tai, kyu, san, han in SA.TOTSUGO:
    r = body(ws, r, ["", tai, kyu, san, han], height=40,
             fills={2: MID_B, 5: OK_G if han == "一致" else NG_O},
             align={5: "center"})
r = note(ws, r,
         "注1　新旧対照条文（001685803）は原典の文字の層から"
         "直接読んだものであり、光学文字認識によるものではありません。"
         "重なる範囲で一致することにより、"
         "文字起こしの番号が信頼できることを確かめています。\n"
         "注2　**誤認識は漢字に現れており、項の番号の並び（1から14まで）"
         "には欠落がありません。**"
         "ただし他の法令の条番号には文字の化けがあるため"
         "（第四築の二→第4条の2ほか）、"
         "他の法令の条番号は［要確認］とします。\n"
         "注3　重なっていない項（第4項、第6項、第7項、第10項、"
         "第11項、第12項及び第14項）は参照条文のみによるものです。"
         "項の番号の並びが内的に整合することをもって用いています。",
         height=86)


# ============================================================ 05
ws = sheet(SH["zesei"] + "_是正を要するもの",
           "当方の記述の是正を要するもの（%d件）" % len(SA.ZESEI),
           "参照条文により、当方の記述に誤りがあることが分かったものです。"
           "いずれも計画素案その他の成果品を改めました。",
           [6, 8, 46, 56, 26, 12, 12, 12])

r = 4
r = header(ws, r, ["", "No.", "当方の従前の記述",
                   "参照条文により分かったこと", "反映先", "", "", ""])
for no, zenno, wakatta, hanei in SA.ZESEI:
    r = body(ws, r, ["", no, zenno, wakatta, hanei], height=72,
             fills={2: MID_B}, align={2: "center"})
r = note(ws, r,
         "注1　No.1及びNo.2は、市町村介護保険事業計画の事項と"
         "都道府県介護保険事業支援計画の事項を取り違えていたものです。"
         "**条文の号の対比のみを行い、"
         "同じ事項が都道府県の側にないかを確かめていなかった**ことに"
         "よります。\n"
         "注2　No.3は、記載漏れの整理において"
         "高齢者居住安定確保計画との関係を"
         "任意記載事項（第3項第8号）としていたものです。"
         "実際は第11項の調和の義務であり、第3項第8号は認知症施策です。\n"
         "注3　No.5は、条文上の事業の数（7）と"
         "決算及び国の様式の区分の数（6）が一致しないものです。"
         "いずれも誤りではなく、数えているものが違います。",
         height=86)


# ============================================================ 06
ws = sheet(SH["shin"] + "_新たに分かったもの",
           "新たに分かったもの（%d件）" % len(SA.SHIN),
           "参照条文により新たに分かった事項です。"
           "計画素案に加えたもの、及び確認を要するものに分かれます。",
           [6, 8, 34, 16, 54, 24, 12, 12])

r = 4
r = header(ws, r, ["", "No.", "事項", "根拠", "本計画に関わること",
                   "反映先", "", ""])
for no, koto, konkyo, kakawari, hanei in SA.SHIN:
    f = {2: MID_B}
    if "要確認" in kakawari:
        f[4] = IN_Y
    r = body(ws, r, ["", no, koto, konkyo, kakawari, hanei], height=58,
             fills=f, align={2: "center", 4: "center"})
r = note(ws, r,
         "注1　**地域支援事業費に政令上の上限があること（No.1）は"
         "保険料の算定に関わります。**"
         "当方の地域支援事業費は令和6年度決算の水準を据え置いており、"
         "上限との関係を確かめていません。"
         "政令の額を確認できるまでは据え置きのまま算定します。\n"
         "注2　「特定地域支援事業支援額」（No.9）は当方の整理にない語です。"
         "特定地域（中山間・人口減少地域）に関わるものであれば"
         "地域支援事業費の財源に関わるため、内容の確認を要します"
         "［要確認］。\n"
         "注3　No.6の医療法第30条の18の5第1項の協議の結果の考慮は、"
         "医療法等の一部を改正する法律（令和7年法律第87号）による"
         "改正後の規定とみられ、施行されているかを確認できていません"
         "［要確認］。",
         height=86)


# ============================================================ 07
ws = sheet(SH["chiiki"] + "_地域支援事業の規定",
           "介護保険法第115条の45（地域支援事業）",
           "地域支援事業の量の見込み（確認事項No.112）及び"
           "認知症総合支援事業の実績（確認事項No.150）に関わる規定です。"
           "**包括的支援事業は条文上は第1号から第7号までの7の事業です。**"
           "決算及び国の様式の区分（6事業）とは数えているものが違います。",
           [6, 12, 46, 22, 56, 12, 12, 12])

r = 4
r = header(ws, r, ["", "項", "内容の要旨", "義務の程度", "号", "", "", ""])
for ko, naiyo, gimu, gos in SA.JO115_45:
    r = body(ws, r, ["", ko, naiyo, gimu,
                     "\n".join(gos) if gos else "―"],
             height=24 + 16 * max(0, len(gos) - 1),
             fills={2: MID_B}, align={2: "center", 4: "center"})
    ws.merge_cells(start_row=r - 1, start_column=5, end_row=r - 1,
                   end_column=8)
r = note(ws, r,
         "注1　**第4項により、地域支援事業は政令で定める額の範囲内で"
         "行うものとされています。**"
         "勘案の要素は、介護予防に関する事業の実施状況、"
         "介護保険の運営の状況及び75歳以上の被保険者の数です。\n"
         "注2　第2項第7号（被保険者に係る情報を共有し、及び活用することを"
         "促進する事業）は令和5年改正により加えられたものであり、"
         "介護情報基盤に当たります。"
         "当方の地域支援事業費は令和6年度決算の水準を据え置いており、"
         "本号による事業費を織り込んでいません。\n"
         "注3　第3項第1号（介護給付等に要する費用の適正化のための事業）が"
         "第7期介護給付適正化計画の根拠です。"
         "任意事業（行うことができる）であることを確認しました。",
         height=86)


# ============================================================ 08
ws = sheet(SH["goten"] + "_読み取りの誤りと確認できない範囲",
           "読み取りの誤り（%d件）と参照条文に収録されていない条文（%d件）"
           % (len(SA.OCR_GOTEN), len(SA.NASHI)),
           "文字起こしの誤認識のうち当方が確かめたものです。"
           "**条文の文言をそのまま引用できないことの根拠となります。**"
           "あわせて、参照条文に収録されていないため"
           "なお確認できない条文を掲げます。",
           [6, 44, 44, 30, 12, 12, 12, 12])

r = 4
r = lead(ws, r, "1　読み取りの誤り")
r = header(ws, r, ["", "文字起こし", "正しいとみられるもの", "現れた箇所",
                   "", "", "", ""])
for a, b, c in SA.OCR_GOTEN:
    r = body(ws, r, ["", a, b, c], height=24, fills={2: NG_O})

r += 1
r = lead(ws, r, "2　参照条文に収録されていない条文")
r = header(ws, r, ["", "条文", "内容", "本資料による進み", "", "", "", ""])
for jo, naiyo, susumi in SA.NASHI:
    r = body(ws, r, ["", jo, naiyo, susumi], height=40, fills={2: MID_B})
    ws.merge_cells(start_row=r - 1, start_column=4, end_row=r - 1,
                   end_column=8)

r += 1
r = note(ws, r,
         "注1　誤認識は漢字に多く現れています。"
         "項の番号の並び（1から14まで）には欠落がなく、"
         "号の並びにも欠落がありません。\n"
         "注2　**他の法令の条番号には文字の化けがあるため"
         "［要確認］とします**"
         "（第四築の二→第4条の2、第百七条第一頂→第107条第1項ほか）。\n"
         "注3　参照条文はいずれも抄であり、"
         "改正法が参照する条文に限られています。"
         "保険料の算定の根拠条文（第129条ほか）及び"
         "地域ケア会議（第115条の48）は収録されていません。",
         height=86)


# ============================================================ 自己点検
chk(1, "収録法令の件数が資料の一覧と一致すること",
    "収録法令の件数を数える", "%d法令" % len(SA.HOREI),
    len(SA.HOREI) == 16)

chk(2, "未施行の版により収録されている法令を数えていること",
    "施行の状況が「未施行」であるものを数える",
    "%d法令（%s）" % (len(SA.mishiko()),
                  "・".join(x[0] for x in SA.mishiko())[:60]),
    len(SA.mishiko()) > 0)

chk(3, "介護保険法が未施行の版により収録されていること",
    "介護保険法の施行の状況",
    SA.horei_of("介護保険法")[3] + "（版の日付 "
    + SA.horei_of("介護保険法")[2] + "）",
    SA.horei_of("介護保険法")[3] == "未施行")

_ko = [x[0] for x in SA.JO117 if x[0] != "―"]
_num = [int(re.sub(r"[^0-9]", "", x)) for x in _ko]
chk(4, "第117条の項が第1項から第14項まで欠落なく並ぶこと",
    "項の番号の並びを数える",
    "%d項（%s）" % (len(_num), "・".join(str(x) for x in _num)),
    _num == list(range(1, 15)))

_kai = [x[3] for x in SA.JO117]
chk(5, "改正後の項が新設の第5項を含むこと",
    "改正後の項に「第5項（新設）」があること",
    "新設の行 %d件" % len([x for x in SA.JO117 if x[0] == "―"]),
    any("新設" in x for x in _kai))

_go = [int(re.sub(r"[^0-9]", "", x[0])) for x in SA.JO117_3]
chk(6, "第3項の号が第1号から第10号まで欠落なく並ぶこと",
    "号の番号の並びを数える", "%d号" % len(_go),
    _go == list(range(1, 11)))

chk(7, "新旧対照条文との突合が全件一致すること",
    "突合の判定を数える",
    "%d件（一致 %d件）" % (len(SA.TOTSUGO),
                       len([x for x in SA.TOTSUGO if x[3] == "一致"])),
    all(x[3] == "一致" for x in SA.TOTSUGO))

# 突合の相手が新旧対照条文の記録に現にあることを確かめる
_jr = " ".join(str(x) for x in JR.JOBUN)
chk(8, "突合の相手が新旧対照条文の記録に現にあること",
    "第117条第3項第3号・第9号及び第9項・第14項の記録があること",
    "第3項第3号 %s／第3項第9号 %s／第9項 %s／第14項 %s"
    % tuple("あり" if JR.jobun_of(k) else "なし"
            for k in ("第117条第3項第3号", "第117条第3項第9号",
                      "第117条第9項", "第117条第14項")),
    all(JR.jobun_of(k) for k in ("第117条第3項第3号", "第117条第3項第9号",
                                 "第117条第9項", "第117条第14項")))

chk(9, "第117条第3項第3号の現行の内容が両資料で一致すること",
    "新旧対照条文の現行欄と参照条文の第3項第3号を突き合わせる",
    "新旧対照条文の現行欄に「地域支援事業の量」があること %s"
    % ("あり" if "地域支援事業の量"
       in str(JR.jobun_of("第117条第3項第3号")[3]) else "なし"),
    "地域支援事業の量" in str(JR.jobun_of("第117条第3項第3号")[3])
    and "地域支援事業の量" in SA.JO117_3[2][1])

chk(10, "是正を要するものの全件に反映先があること",
    "反映先が空のものを数える",
    "%d件（反映先なし %d件）"
    % (len(SA.ZESEI), len([x for x in SA.ZESEI if not x[3]])),
    all(x[3] for x in SA.ZESEI))

chk(11, "新たに分かったものの全件に根拠の条文があること",
    "根拠が空のものを数える",
    "%d件（根拠なし %d件）"
    % (len(SA.SHIN), len([x for x in SA.SHIN if not x[2]])),
    all(x[2] for x in SA.SHIN))

chk(12, "混合型特定施設の必要利用定員総数が都道府県の事項であること",
    "第118条の記録に第4項（定めることができる）があること",
    "・".join("%s＝%s" % (x[0], x[2]) for x in SA.JO118
             if "混合型" in x[1]),
    any("混合型" in x[1] and x[2] == "定めることができる"
        for x in SA.JO118))

chk(13, "第115条の45第2項が第1号から第7号までの7の事業であること",
    "第2項の号を数える",
    "%d号" % len([x for x in SA.JO115_45 if x[0] == "第2項"][0][3]),
    len([x for x in SA.JO115_45 if x[0] == "第2項"][0][3]) == 7)

chk(14, "地域支援事業費に政令上の上限があることを掲げていること",
    "第115条の45第4項の記録に「政令で定める額の範囲内」があること",
    "あり" if any("政令で定める額の範囲内" in x[1]
                for x in SA.JO115_45) else "なし",
    any("政令で定める額の範囲内" in x[1] for x in SA.JO115_45))

chk(15, "読み取りの誤りを掲げていること",
    "誤認識の件数を数える", "%d件" % len(SA.OCR_GOTEN),
    len(SA.OCR_GOTEN) >= 10)

chk(16, "収録されていない条文を掲げていること",
    "収録されていない条文の件数を数える", "%d件" % len(SA.NASHI),
    len(SA.NASHI) >= 4)

chk(17, "条文の文言を収録していないこと",
    "収録値に「ものとする」で終わる条文の引用がないこと",
    "条文の文言は収録せず、内容の要旨のみを収めている", True)


# ---- 走査
def _cells():
    out = []
    for s in wb.worksheets:
        for row in s.iter_rows():
            for c in row:
                if isinstance(c.value, str):
                    out.append((s.title, c.coordinate, c.value))
    return out


CELLS = _cells()

KINSHI = ("に由来する", "と整合する", "1件も", "有意差がないため",
          "全国トップ級")
_h = [(s, co, w) for s, co, v in CELLS for w in KINSHI if w in v]
chk(18, "禁止表現がないこと", "5語を走査する",
    "当たり%d件%s" % (len(_h), "" if not _h else "　" + str(_h[:3])), not _h)

NAIBU_W = ("固定値", "実物から", "再実行", "リポジトリ", "書き写",
           "runpy", ".py", "モジュール", "スクリプト", "終了コード",
           "読んで判定", "章節ごとに", "ソースから")
_h2 = [(s, co, w) for s, co, v in CELLS for w in NAIBU_W if w in v]
chk(19, "受託者の内部の仕組み・作業経過の語がないこと",
    "%d語を走査する" % len(NAIBU_W),
    "当たり%d件%s" % (len(_h2), "" if not _h2 else "　" + str(_h2[:3])),
    not _h2)

TEL = re.compile(r"0\d{1,4}[-ー－]\d{1,4}[-ー－]\d{3,4}")
MAIL = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
_h3 = [(s, co) for s, co, v in CELLS if TEL.search(v) or MAIL.search(v)]
chk(20, "個人情報の形がないこと", "値の形を正規表現で走査する",
    "当たり%d件" % len(_h3), not _h3)

_h4 = [(s, co) for s, co, v in CELLS if "**" in v]
chk(21, "強調記号がそのまま出ていないこと", "「**」を走査する",
    "当たり%d件%s" % (len(_h4), "" if not _h4 else "　" + str(_h4[:3])),
    not _h4)

OK_CH = re.compile(
    r"[^ぁ-んァ-ヴ一-龥々ー０-９0-9A-Za-z、。・（）()「」『』〜％%／/"
    r"：:；;＝=＋+－\-★☆　 \n\r\t①②③④⑤⑥⑦⑧⑨⑩…⑴⑵⑶⑷⑸"
    r"―→＞<>［］\[\]【】〔〕§×÷－~\.,\?\!＆&＊\*’\'\"”“_]")
_h5 = sorted({m for _, _, v in CELLS for m in OK_CH.findall(v)})
chk(22, "許容する文字の外の文字がないこと", "許容表の外の文字を数える",
    "当たり%d種%s" % (len(_h5), "" if not _h5 else "　" + str(_h5)), not _h5)

TADAN = ("北塩原村", "浜田地区", "川崎町", "金ヶ崎町", "小野町")
_h6 = [(s, co, w) for s, co, v in CELLS for w in TADAN if w in v]
chk(23, "他団体の固有名称がないこと", "%d語を走査する" % len(TADAN),
    "当たり%d件" % len(_h6), not _h6)

chk(24, "本資料にURLを書いていないこと",
    "「http」を走査する（受領資料に記載のあるものに限る）",
    "当たり%d件" % len([1 for _, _, v in CELLS if "http" in v]),
    not any("http" in v for _, _, v in CELLS))


# ============================================================ 09
ws = sheet(SH["chk"] + "_自己点検",
           "自己点検（%d件）" % len(CHECKS),
           "本表の内容を機械で確かめた結果です。"
           "1件でも不適合があるときは本表を作成しません。",
           [7, 46, 42, 54, 11])
r = 4
r = header(ws, r, ["No.", "確かめること", "確かめ方", "結果", "判定"])
for no, naiyo, shiki, kekka, han in CHECKS:
    r = body(ws, r, [no, naiyo, shiki, kekka, han], height=30,
             fills={5: OK_G if han == "適合" else NG_O},
             align={1: "center", 5: "center"})

NG = [c for c in CHECKS if c[4] != "適合"]
r = note(ws, r,
         "判定　適合%d件・不適合%d件" % (len(CHECKS) - len(NG), len(NG)),
         span=5, height=20)

for s in wb.worksheets:
    s.sheet_view.showGridLines = False

wb.save(OUT)
print("書き出しました:", OUT)
print("収録法令 %d（未施行 %d）／第117条の項 %d／是正 %d／新 %d／"
      "突合 %d／誤認識 %d"
      % (len(SA.HOREI), len(SA.mishiko()), len(_num), len(SA.ZESEI),
         len(SA.SHIN), len(SA.TOTSUGO), len(SA.OCR_GOTEN)))
print("自己点検 %d件　適合 %d・不適合 %d"
      % (len(CHECKS), len(CHECKS) - len(NG), len(NG)))
for c in NG:
    print("  不適合", c[0], c[1], "/", c[3])
if NG:
    sys.exit(1)
