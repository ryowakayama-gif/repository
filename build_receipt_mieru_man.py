# -*- coding: utf-8 -*-
"""大雪地区広域連合 第10期介護保険事業計画
　　見える化システムの利用マニュアル及び事業所分析データの受領点検.

令和8年10月23日のご連絡
  「確認事項No.191・No.192も含めて受託者案で進めて下さい
　　添付ファイル順次送りますので確認して下さい。」

━━ 受領したもの ━━

  ① 利用マニュアル【システム操作編①】第17.0版
     （【見える化】老健定員関連誤データ正誤表を同梱）
  ② 利用マニュアル【システム操作編② 将来推計】第11.0版（276頁）
  ③ 利用マニュアル【システム操作編③ 将来推計（都道府県機能）】第5.0版
  ④ 利用マニュアル【システム操作編④ 需要推計（都道府県機能）】第1.0版
  ⑤ 利用マニュアル【システム管理編】第4.3版
  ⑥ 地域分析の手引き（令和8年9月改訂版）
  ⑦ 事業所分析データ　4項目×47都道府県
  ⑧ 推計結果（総括表。推計名 1006）
  ⑨ 介護保険事業状況報告の設定の画面の写し（推計名 1006）

①から⑤は資料提供依頼No.26、⑦は同No.25に当たる。いずれも受領済に改めた。

━━ 確かめたこと ━━

1 確認事項No.191が解決した。システムは令和8年度を
  令和8年5月報告（令和8年3月サービス分）以降により置く。
  受領した画面では5月・6月・7月の3か月が設定されている。

2 確認事項No.192が解決した。入力する事業費は対象経費支出予定額である。

3 構成3町の60事業所について、令和3年から令和7年までの
  定員・利用者数・従事者数を確かめた（確認事項No.23・No.78・No.96）。

4 介護老人保健施設の定員に誤りがあり正誤表が示されている。
  正の値（3町計240人）は当方の収録値と一致する。

5 システムに登録されている日常生活圏域は6圏域である
  （東川町1・東神楽町1・美瑛町4）。

6 推計名1006は確定版により作成されたものであり、
  なお手当てを要する箇所が8件（優先A 4件）ある。

出力
  output/第10期計画_見える化の利用マニュアル等の受領点検.xlsx

自己点検で1件でも不適合があると終了コード1で終わる。
"""

import ast
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
import data_mieru_man as MM                                  # noqa: E402
import data_mieru_kekka3 as K3                               # noqa: E402
import data_jigyosho_bunseki as JB                           # noqa: E402

if isinstance(sys.stdout, io.TextIOWrapper):
    sys.stdout.reconfigure(encoding="utf-8")

OUT = os.path.join(RP.OUTPUT,
                   "第10期計画_見える化の利用マニュアル等の受領点検.xlsx")

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


JURYO = datetime.date(2026, 10, 23)

SH = {"about": "00", "uke": "01", "kaiketsu": "02", "kikan": "03",
      "chiiki": "04", "jigyo": "05", "teiin": "06", "jinzai": "07",
      "roken": "08", "kekka": "09", "koiki": "10", "torikumi": "11",
      "chk": "12"}


def shn(key):
    return SH[key] + "シート"


def wareki(d):
    return "令和%d年%d月%d日" % (d.year - 2018, d.month, d.day)


# ============================================================ 読み取り
def literal(fn, name):
    """ソースから literal の代入を読む（固定値を書き写さないため）。"""
    src = io.open(os.path.join(RP.ROOT, fn), encoding="utf-8").read()
    for node in ast.parse(src).body:
        if isinstance(node, ast.Assign) and any(
                getattr(t, "id", None) == name for t in node.targets):
            return ast.literal_eval(node.value)
    raise RuntimeError("%s に %s が見つからない" % (fn, name))


CHECK = literal("build_process_control.py", "CHECK")
LACK = literal("build_process_control.py", "LACK")
BY_C = {x[0]: x for x in CHECK}
BY_L = {x[0]: x for x in LACK}

# 本件により解決した確認事項と、備考を更新した確認事項
KAIKETSU = [
    (191, "令和8年度の実績見込みに用いる月の起点と終期",
     "操作編②により、令和8年5月報告（令和8年3月サービス分）以降の"
     "累計から実績見込みを算定することを確かめた。"
     "受領した画面では5月・6月・7月の3か月が設定されている"),
    (192, "入力する地域支援事業費の額の作り方",
     "操作編②により、総事業費から寄付金その他の収入額を差し引いた額の"
     "うち交付金対象経費として支出する見込みの額"
     "（対象経費支出予定額）であることを確かめた"),
    (148, "操作手引・認証後画面の写しのご提供",
     "操作編①から④及びシステム管理編の5編を受領した。"
     "認証後画面の写しは令和8年9月30日に受領済みであり、本件で全て解決した"),
    (78, "グループホームびえいの郷の定員減の時期",
     "事業所分析データにより、令和3年から令和7年まで18人で"
     "変わっていないことを確かめた"),
]

KOSHIN = [
    (23, "事業所の定員・稼働率データ",
     "事業所分析データにより、構成3町の60事業所について"
     "令和3年から令和7年までの定員・1月あたり延べ利用者数・"
     "利用者数・従事者数を確かめた"),
    (96, "認知症対応型共同生活介護の共同生活住居数と、くるみの郷の定員",
     "事業所分析データには5事業所（定員計81人）が掲げられており、"
     "くるみの郷は現れない。北海道の名簿による99人との差は1事業所分である"),
    (107, "施設・居住系の実績の減少を将来へ延長するか",
     "利用者数が5年続けて前年を下回る事業所は0件であり、"
     "施設・居住系サービスの定員は特定施設の2事業所を除き変わっていない"),
    (116, "介護人材の確保状況を供給の制約として見込量に反映するか",
     "介護職員の常勤換算数が5か年で2割以上減った事業所が11件あり、"
     "うち3件が訪問介護である"),
    (133, "介護老人保健施設の令和8年度の減少をどう見るか",
     "正誤表により、現状分析機能の定員のデータに誤りがあったことを"
     "確かめた。正の値（3町計240人）は当方の収録値と一致する"),
    (88, "日常生活圏域ごとの必要利用定員総数",
     "システムに登録されている日常生活圏域は6圏域である"
     "（東川町1・東神楽町1・美瑛町4）。"
     "また都道府県への提出時に「必要入所（利用）定員、入居定員総数調査票」を"
     "登録することとされている（資料提供依頼No.31）"),
    (12, "日常生活圏域別のデータ",
     "正誤表の「定員（日常生活圏域）」により、"
     "美瑛町が旭・北西地区、美馬牛地区、朗根内地区、"
     "市街地・周辺地区の4圏域に分かれていることを確かめた"),
    (156, "所得段階別第1号被保険者数の将来推計",
     "弾力化の段階が令和8年9月30日に40段階へ拡張された"),
    (137, "1人1月あたり給付費の実績値をどの年度にするか",
     "推計名1006は施設・居住系・在宅とも令和8年度を選んでいる。"
     "令和8年度は実績見込みを算出せず、令和8年5月報告から"
     "設定した月までの累計のみから算定される"),
]


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


# ============================================================ 00
ws = sheet(SH["about"] + "_この資料について",
           "見える化システムの利用マニュアル及び事業所分析データの受領点検",
           "%sに受領した資料を読み、当方の算定・計画素案・確認事項の台帳と"
           "突き合わせたものです。"
           "原本は保管せず、集計値と定めのみを確認のために用いています。"
           % wareki(JURYO),
           [26, 30, 30, 18, 10, 10, 10, 10])
r = 4
r = lead(ws, r, "1　本資料の位置づけ")
r = note(ws, r,
         "本資料は、受領した資料により何が確かめられ、"
         "何が解決し、何がなお残るかを示すものです。\n"
         "・確認事項No.191及びNo.192は、いずれも原典により解決しました"
         "（%s・%s）。\n"
         "・事業所分析データにより、区域内の事業所ごとの定員と従事者の"
         "5か年の推移が得られました（%s）。"
         "これは従前、事業者へ直接お尋ねしなければ得られないものとして"
         "整理していた事項です。\n"
         "・受領した推計（推計名1006）は確定版により作成されたものであり、"
         "なお手当てを要する箇所が%d件あります（%s）。"
         % (shn("kikan"), shn("chiiki"), shn("jigyo"),
            len(K3.NOKORU), shn("kekka")), height=118)
r += 1
r = lead(ws, r, "2　構成")
r = header(ws, r, ["シート", "内容", "", "", "", "", "", ""])
for k, t in [("uke", "受領したもの"),
             ("kaiketsu", "本件により解決した確認事項と備考を更新した確認事項"),
             ("kikan", "令和8年度の実績見込みに用いる月（確認事項No.191）"),
             ("chiiki", "地域支援事業費の額の作り方（確認事項No.192）"),
             ("jigyo", "事業所分析データ（構成3町の事業所）"),
             ("teiin", "施設・居住系及び通所系の定員の5か年の推移"),
             ("jinzai", "介護職員の常勤換算数の5か年の推移"),
             ("roken", "介護老人保健施設の定員の誤データと日常生活圏域"),
             ("kekka", "推計名1006と当方の算定との対照"),
             ("koiki", "広域連合向けの機能と都道府県への提出"),
             ("torikumi", "都道府県への提出時に入力する取組の項目"),
             ("chk", "自己点検")]:
    r = body(ws, r, [shn(k), t, "", "", "", "", "", ""], height=20)

# ============================================================ 01
ws = sheet(SH["uke"] + "_受領したもの", "受領したもの",
           "原本は保管しません。"
           "集計値と定めのみを確認のために用いています。",
           [14, 46, 12, 10, 34, 10, 10, 10])
r = 4
r = header(ws, r, ["区分", "名称", "版", "頁数", "内容", "", "", ""])
for tag, name, han, pg, naiyo in MM.GENTEN:
    r = body(ws, r, [tag, name, han or "―", ("%d頁" % pg) if pg else "―",
                     naiyo, "", "", ""], height=34,
             align={3: "center", 4: "center"})
r = body(ws, r, ["事業所分析", "事業所別従事者数・事業所別利用者数・"
                 "事業所別回転率・施設別定員充足率（各47都道府県）",
                 "―", "―",
                 "構成3町の%d事業所の令和3年から令和7年までの値"
                 % len(JB.JIGYO), "", "", ""], height=34,
         align={3: "center", 4: "center"})
r = body(ws, r, ["推計結果", "推計結果概要（総括表。推計名 %s）" % K3.SUIKEI_MEI,
                 "―", "―", "6シート（推計値サマリほか）", "", "", ""],
         height=22, align={3: "center", 4: "center"})
r = body(ws, r, ["画面の写し", "介護保険事業状況報告の設定（推計名 %s）"
                 % MM.GAMEN_HOUKOKU["推計名"], "―", "―",
                 "令和6年度から令和8年度までの年報・月報の設定", "", "", ""],
         height=22, align={3: "center", 4: "center"})
r += 1
r = lead(ws, r, "資料提供依頼の更新")
r = header(ws, r, ["No.", "名称", "状態", "", "", "", "", ""])
for no in (25, 26, 31):
    x = BY_L[no]
    r = body(ws, r, [no, x[2], x[8], "", "", "", "", ""], height=22,
             align={1: "center", 3: "center"},
             fills={3: (OK_G if x[8] == "受領済" else IN_Y)})
r = note(ws, r,
         "資料提供依頼No.31（第10期必要入所（利用）定員、入居定員総数調査票）は"
         "本件により新たに起票したものです。"
         "都道府県へ推計データを提出する際に登録することとされており、"
         "厚生労働省より別途配布されています。")

# ============================================================ 02
ws = sheet(SH["kaiketsu"] + "_解決した確認事項",
           "本件により解決した確認事項と、備考を更新した確認事項",
           "確認事項は業務工程管理表 03_確認事項一覧から読んでいます。",
           [8, 34, 56, 12, 10, 10, 10, 10])
r = 4
r = lead(ws, r, "1　解決した確認事項")
r = header(ws, r, ["No.", "件名", "確かめたこと", "状態", "", "", "", ""])
for no, mei, naiyo in KAIKETSU:
    r = body(ws, r, [no, mei, naiyo, BY_C[no][7], "", "", "", ""],
             height=46, align={1: "center", 4: "center"}, fills={4: OK_G})
r += 1
r = lead(ws, r, "2　備考を更新した確認事項")
r = header(ws, r, ["No.", "件名", "確かめたこと", "状態", "", "", "", ""])
for no, mei, naiyo in KOSHIN:
    r = body(ws, r, [no, mei, naiyo, BY_C[no][7], "", "", "", ""],
             height=46, align={1: "center", 4: "center"}, fills={4: MID_B})
r = note(ws, r,
         "状態が「受託者案で確定」であるものは、令和8年10月19日のご了承により"
         "受託者の案により作成を進めることとされた事項です。"
         "本件で得られた事実により、その案の根拠を改めたものを含みます。")

# ============================================================ 03
ws = sheet(SH["kikan"] + "_実績見込みの期間",
           "令和8年度の実績見込みに用いる月（確認事項No.191）",
           "利用マニュアル【システム操作編② 将来推計】第11.0版により"
           "確かめたものです。",
           [14, 12, 34, 34, 12, 10, 10, 10])
r = 4
r = lead(ws, r, "1　システムが用いる実績値の期間")
r = header(ws, r, ["年度", "出所", "報告月", "サービス提供月", "", "", "", ""])
for nendo, (shutsu, hokoku, service) in MM.JISSEKI_KIKAN.items():
    r = body(ws, r, [nendo, shutsu, hokoku or "―", service or "―",
                     "", "", "", ""], height=34, align={1: "center",
                                                        2: "center"})
r = note(ws, r, MM.R8_MIKOMI + "\n" + MM.R8_TANKA)
r += 1
r = lead(ws, r, "2　受領した画面の設定")
r = header(ws, r, ["項目", "設定", "", "", "", "", "", ""])
for k, v in MM.GAMEN_HOUKOKU.items():
    if k == "令和8年度の月数":
        continue
    r = body(ws, r, [k, v, "", "", "", "", "", ""], height=22)
r = note(ws, r, MM.R8_SUISHO +
         "\n受領した画面は%dか月であり、推奨される月数に届いていません。"
         "月報が公表されている月まで設定を広げることをお勧めします。"
         % MM.GAMEN_HOUKOKU["令和8年度の月数"])
r += 1
r = lead(ws, r, "3　当方の算定との関係")
r = note(ws, r,
         "当方は令和8年度を令和8年4月から7月までの4か月（サービス提供月）"
         "として整理しており、システムの整理"
         "（令和8年3月から5月までの3か月）とは起点が2か月異なります。\n"
         "本計画の算定は令和7年度（12か月の完結年度）を基準年度としており、"
         "令和8年度を数量の基準に用いていないため、"
         "見込量及び保険料は動きません。\n"
         "計画素案 第6章第2節 注1-1 に両者の期間の違いを注記しました。",
         height=92)
r += 1
r = lead(ws, r, "4　推計に用いる値の算式（原典による）")
r = header(ws, r, ["項目", "算式", "", "", "", "", "", ""])
for k, v in MM.SUIKEI_SHIKI.items():
    r = body(ws, r, [k, v, "", "", "", "", "", ""], height=22)
r = note(ws, r, MM.YOSHIEN_RANKAI)

# ============================================================ 04
ws = sheet(SH["chiiki"] + "_地域支援事業費",
           "地域支援事業費の額の作り方（確認事項No.192）",
           "利用マニュアル【システム操作編② 将来推計】第11.0版により"
           "確かめたものです。",
           [20, 56, 34, 12, 10, 10, 10, 10])
r = 4
r = lead(ws, r, "1　原典の定め")
r = note(ws, r, MM.CHIIKI_HI)
r += 1
r = lead(ws, r, "2　当方の額との違い")
r = header(ws, r, ["区分", "額の作り方", "本計画の扱い", "", "", "", "", ""])
r = body(ws, r, ["システムが想定する額",
                 "総事業費から寄付金その他の収入額を差し引いた額のうち、"
                 "交付金対象経費として支出する見込みの額",
                 "事業別の積上げによる算定に置き直す際に"
                 "この範囲に合わせて整理する", "", "", "", ""], height=46)
r = body(ws, r, ["本計画の額",
                 "令和6年度の介護保険特別会計決算の款4の額から"
                 "保険者機能強化推進事業費及び保険者努力支援事業費を"
                 "除いたもの",
                 "寄付金その他の収入額を差し引いていない点で範囲が異なる",
                 "", "", "", "", ""], height=46)
r = note(ws, r,
         "受領した推計（推計名%s）は地域支援事業費が全ての欄で0であり、"
         "システム上はなお未入力です。\n"
         "計画素案 第6章第3節2に額の範囲の違いを注記しました。"
         % K3.SUIKEI_MEI)

# ============================================================ 05
ws = sheet(SH["jigyo"] + "_事業所分析データ",
           "事業所分析データ（構成3町の事業所）",
           "受領物は都道府県単位の全47件です。"
           "本資料には構成3町の行のみを掲げます。",
           [12, 30, 20, 10, 12, 12, 12, 12])
r = 4
r = lead(ws, r, "1　出どころと対象")
r = note(ws, r,
         "従事者数及び定員は介護サービス情報公表システム（各年9月時点）、"
         "1月あたりサービス利用者数は介護保険総合データベース"
         "（各年9月サービス分）によります。\n"
         "対象は最新年度の9月サービス分の算定実績のある事業所に限られ、"
         "介護予防給付のみの事業所は対象外です。"
         "したがって区域内のすべての事業所を覆うものではありません。\n"
         "定員充足率及び回転率の欄は原本では数式であり"
         "計算済みの値を持たないため、当方で算式のとおり求めています"
         "（%s）。" % MM.GENTEN[0][0], height=92)
r += 1
r = lead(ws, r, "2　群別の件数")
r = header(ws, r, ["群", "事業所数", "東川町", "美瑛町", "東神楽町",
                   "定員の欄", "", ""])
for g in ("訪問系", "通所系", "短期入所系", "多機能系", "施設居住系"):
    rr = JB.rows(gun=g)
    r = body(ws, r, [g, "%d" % len(rr)]
             + ["%d" % sum(1 for _, v in rr if v["町"] == m)
                for m in ("東川町", "美瑛町", "東神楽町")]
             + ["あり" if g in ("通所系", "施設居住系") else "なし", "", ""],
             height=20, align={i: "center" for i in range(2, 7)})
r = body(ws, r, ["計", "%d" % len(JB.JIGYO)]
         + ["%d" % sum(1 for v in JB.JIGYO.values() if v["町"] == m)
            for m in ("東川町", "美瑛町", "東神楽町")] + ["―", "", ""],
         height=20, bold=True, align={i: "center" for i in range(2, 7)})
r += 1
r = lead(ws, r, "3　原本の数式")
r = header(ws, r, ["項目", "算式", "", "", "", "", "", ""])
for k, v in JB.SHIKI.items():
    r = body(ws, r, [k, v, "", "", "", "", "", ""], height=20)

# ============================================================ 06
ws = sheet(SH["teiin"] + "_定員の推移",
           "施設・居住系及び通所系の定員の5か年の推移",
           "定員は介護サービス情報公表システムの各年9月時点の値です。",
           [22, 26, 10, 9, 9, 9, 9, 9])
r = 4
r = lead(ws, r, "1　サービス別の定員（令和7年）")
r = header(ws, r, ["サービス", "事業所数", "定員"] + list(MM.GENTEN[0][:0])
           + ["", "", "", ""])
for c in ("51", "54", "52", "33", "32", "15", "16", "78"):
    r = body(ws, r, [JB.svc_name(c), "%d" % len(JB.rows(svc=c)),
                     "%d人" % JB.teiin(c), "", "", "", "", ""],
             height=20, align={2: "center", 3: "right"})
r += 1
r = lead(ws, r, "2　5か年で定員が動いた事業所")
r = header(ws, r, ["町", "事業所", "サービス"] + list(JB.NENDO) + [""])
_UGOKI = [(k, v) for k, v in JB.JIGYO.items()
          if v.get("定員") and len({x for x in v["定員"] if x is not None}) > 1]
for k, v in _UGOKI:
    r = body(ws, r, [v["町"], v["名"], JB.svc_name(k[2])]
             + [("%d" % x if x is not None else "―") for x in v["定員"]] + [""],
             height=22, align={i: "center" for i in range(4, 9)})
r = note(ws, r,
         "グループホームびえいの郷の定員は令和3年から令和7年まで18人であり、"
         "変わっていません（確認事項No.78）。\n"
         "認知症対応型共同生活介護は本データに5事業所・定員計%d人が"
         "掲げられており、北海道の名簿による99人とは1事業所分の差があります"
         "（確認事項No.96）。" % JB.teiin("32"), height=76)
r += 1
r = lead(ws, r, "3　5年続けて利用者数が前年を下回る事業所")
_RENZOKU = [v for v in JB.JIGYO.values() if JB.renzoku_gensho(v)]
r = note(ws, r,
         "%d件です。"
         "施設・居住系サービスの定員も、特定施設入居者生活介護の2事業所を除き"
         "5か年を通じて変わっていません。"
         "施設・居住系の実績の減少を将来へ延長しないという本計画の扱い"
         "（確認事項No.107）と向きが合います。" % len(_RENZOKU))

# ============================================================ 07
ws = sheet(SH["jinzai"] + "_従事者数の推移",
           "介護職員の常勤換算数の5か年の推移",
           "介護サービス情報公表システムの各年9月時点の値です。"
           "事業所が公表した時点の値であり、記入のない年があるものを含みます。",
           [12, 30, 24, 9, 9, 9, 9, 10])
r = 4
r = lead(ws, r, "1　3町計の推移")
r = header(ws, r, ["区分"] + list(JB.NENDO) + ["", ""])
_KEI = [sum(v["介護職員常勤換算"][i] or 0 for v in JB.JIGYO.values())
        for i in range(5)]
r = body(ws, r, ["介護職員常勤換算（3町計）"]
         + ["%.0f人" % x for x in _KEI] + ["", ""],
         height=20, align={i: "right" for i in range(2, 7)})
for m in ("東川町", "美瑛町", "東神楽町"):
    r = body(ws, r, ["　" + m]
             + ["%.0f人" % sum(v["介護職員常勤換算"][i] or 0
                               for v in JB.JIGYO.values() if v["町"] == m)
                for i in range(5)] + ["", ""],
             height=20, align={i: "right" for i in range(2, 7)})
r = note(ws, r,
         "3町計は令和3年の%.0f人から令和7年の%.0f人へ増えています。"
         "ただし公表した事業所の数と記入の状況により年ごとに振れるため、"
         "区域の介護職員の実人数の推移としては読めません。"
         % (_KEI[0], _KEI[-1]))
r += 1
r = lead(ws, r, "2　5か年で常勤換算数が2割以上減った事業所")
_GENIN = sorted(
    [(k, v, JB.gensho(v, "介護職員常勤換算")) for k, v in JB.JIGYO.items()
     if (JB.gensho(v, "介護職員常勤換算") or 0) <= -0.2],
    key=lambda x: x[2])
r = header(ws, r, ["町", "事業所", "サービス"]
           + ["令和3年", "令和7年", "変化の割合", "", ""])
for k, v, g in _GENIN:
    x = v["介護職員常勤換算"]
    r = body(ws, r, [v["町"], v["名"], JB.svc_name(k[2]),
                     "%.0f" % x[0], "%.0f" % x[-1], "▲%.1f％" % (-g * 100),
                     "", ""], height=22,
             align={4: "right", 5: "right", 6: "right"})
r = note(ws, r,
         "%d事業所のうち%d事業所が訪問介護です。"
         "訪問介護は区域内の3事業所とも減っています。\n"
         "本計画は供給の制約を見込量に反映していません"
         "（確認事項No.116）。"
         "見込量は需要の側から算定したものであり、"
         "これを担う人材が確保できるかは別に確かめる必要があります。"
         % (len(_GENIN),
            sum(1 for k, _v, _g in _GENIN if JB.svc_name(k[2]) == "訪問介護")),
         height=76)

# ============================================================ 08
ws = sheet(SH["roken"] + "_老健の定員と圏域",
           "介護老人保健施設の定員の誤データと日常生活圏域",
           "利用マニュアル【システム操作編①】に同梱されている"
           "正誤表により確かめたものです。",
           [18, 18, 18, 40, 10, 10, 10, 10])
r = 4
r = lead(ws, r, "1　正誤表の概要")
r = note(ws, r,
         MM.ROKEN_SEIGO["概要"] +
         "\n対象機能：" + MM.ROKEN_SEIGO["対象機能"] +
         "\n対象指標：" + "／".join(MM.ROKEN_SEIGO["対象指標"]), height=76)
r += 1
r = header(ws, r, ["区分", "誤", "正", "備考", "", "", "", ""])
for k in ("全国", "北海道", "東川町", "美瑛町", "東神楽町"):
    a, b = MM.ROKEN_SEIGO[k]
    r = body(ws, r, [k, "{:,}人".format(a), "{:,}人".format(b),
                     ("当方の収録値と一致する"
                      if k in ("東川町", "美瑛町", "東神楽町") else "―"),
                     "", "", "", ""], height=20,
             align={2: "right", 3: "right"},
             fills=({3: OK_G} if k in ("東川町", "美瑛町", "東神楽町") else None))
_SUM = sum(MM.ROKEN_SEIGO[m][1] for m in ("東川町", "美瑛町", "東神楽町"))
r = body(ws, r, ["3町計",
                 "{:,}人".format(sum(MM.ROKEN_SEIGO[m][0]
                                     for m in ("東川町", "美瑛町", "東神楽町"))),
                 "{:,}人".format(_SUM),
                 "当方の収録値（3施設240人）と一致する", "", "", "", ""],
         height=20, bold=True, align={2: "right", 3: "right"},
         fills={3: OK_G})
r = note(ws, r,
         "対象は現状分析機能であり、将来推計機能には及びません。\n"
         "本計画の定員は北海道の名簿及び介護サービス情報公表システムに"
         "よっており、誤りの影響を受けていません。")
r += 1
r = lead(ws, r, "2　システムに登録されている日常生活圏域")
r = header(ws, r, ["町", "圏域", "", "", "", "", "", ""])
for m, ken in MM.SEIKATSU_KEN:
    r = body(ws, r, [m, ken, "", "", "", "", "", ""], height=20)
r = note(ws, r,
         "正誤表の「定員（日常生活圏域）」により確かめたものです。"
         "構成3町で%d圏域（東川町1・東神楽町1・美瑛町%d）が"
         "登録されています。\n"
         "計画素案 第6章第4節4は基盤整備の単位を3町（3圏域）とする案を"
         "掲げており、システムに登録されている圏域とは数が異なります"
         "［要協議］。日常生活圏域ごとの必要利用定員総数は"
         "介護保険法第117条第2項第1号の法定記載事項です"
         "（確認事項No.88・No.12）。"
         % (len(MM.SEIKATSU_KEN),
            sum(1 for m2, _k in MM.SEIKATSU_KEN if m2 == "美瑛町")),
         height=92)

# ============================================================ 09
ws = sheet(SH["kekka"] + "_推計1006との対照",
           "推計名%s と当方の算定との対照" % K3.SUIKEI_MEI,
           "確定版（2026年8月1日提供）の将来推計機能により"
           "作成された推計データです。",
           [10, 30, 34, 30, 14, 10, 10, 10])
r = 4
r = lead(ws, r, "1　確定版と暫定版")
r = note(ws, r,
         "確定版の提供：%s／暫定版の提供：%s\n%s\n%s"
         % (K3.__doc__ and MM.KAKUTEIBAN["確定版の提供"],
            MM.KAKUTEIBAN["暫定版の提供"], MM.KAKUTEIBAN["暫定版の扱い"],
            MM.KAKUTEIBAN["足元値"]), height=108)
r += 1
r = lead(ws, r, "2　設定の内容と当方の推奨との対照")
r = header(ws, r, ["判定", "項目", "推計%s の設定" % K3.SUIKEI_MEI,
                   "当方の見方", "", "", "", ""])
for k, v in K3.SETTEI.items():
    han, riyu = K3.SETTEI_HANTEI.get(k, ("―", "―"))
    r = body(ws, r, [han, k, v, riyu, "", "", "", ""], height=34,
             align={1: "center"},
             fills={1: (OK_G if han == "○" else
                        NG_O if han == "×" else GRAY)})
r = note(ws, r,
         "○は当方の推奨と一致するもの、×は異なるもの、"
         "―は判断を要しないものです。")
r += 1
r = lead(ws, r, "3　主な値")
r = header(ws, r, ["項目", "令和9年度", "令和10年度", "令和11年度",
                   "令和12年度", "", "", ""])
for mei, vals in [("第1号被保険者数（人）", K3.HIHO["第1号"][3:7]),
                  ("認定者数（第1号・人）", K3.NINTEI_1GO[3:7]),
                  ("総給付費（千円）", K3.KYUFU[3:7])]:
    r = body(ws, r, [mei] + ["{:,.0f}".format(x) for x in vals]
             + ["", "", ""], height=20,
             align={i: "right" for i in range(2, 6)})
r += 1
r = header(ws, r, ["保険料基準額（月額）"] + list(K3.HOKENRYO)[:5] + ["", ""])
r = body(ws, r, ["推計%s" % K3.SUIKEI_MEI]
         + ["{:,.0f}円".format(K3.HOKENRYO[k]) for k in list(K3.HOKENRYO)[:5]]
         + ["", ""], height=20, align={i: "right" for i in range(2, 7)})
r = note(ws, r,
         "第9期の保険料基準額%.2f円は当方の再現と一致します。\n"
         "第1号被保険者負担割合は令和12年度以降%s と上がることが"
         "システムの係数に織り込まれています。"
         "本計画の中長期推計は23％で固定しているため、"
         "この点でシステムの値と開きます。"
         % (K3.HOKENRYO["第9期"],
            "・".join("%d％" % round(v * 100)
                      for k, v in list(K3.FUTAN_WARIAI.items())[1:])),
         height=76)
r += 1
r = lead(ws, r, "4　なお手当てを要する箇所")
r = header(ws, r, ["優先", "箇所", "内容", "月額への効き", "", "", "", ""])
for yu, mei, naiyo, kiki in K3.NOKORU:
    r = body(ws, r, [yu, mei, naiyo, kiki, "", "", "", ""], height=40,
             align={1: "center", 4: "center"},
             fills={1: (NG_O if yu == "A" else IN_Y if yu == "B" else GRAY)})
r = note(ws, r,
         "5_保険料推計のシートが #DIV/0! となっているのは、"
         "所得段階別の第1号被保険者数が入力されておらず"
         "補正後被保険者数が0であることによります。"
         "1_推計値サマリの保険料基準額は所得段階別の登録前の値です。\n"
         "所得段階別の人数は、令和8年9月30日にお示しした"
         "16段階×8年度の値をそのまま用いることができます。", height=92)

# ============================================================ 10
ws = sheet(SH["koiki"] + "_広域連合向けの機能",
           "広域連合向けの機能と都道府県への提出",
           "利用マニュアル【システム操作編② 将来推計】第11.0版 附録②"
           "及び3-10により確かめたものです。",
           [22, 60, 12, 12, 10, 10, 10, 10])
r = 4
r = lead(ws, r, "1　広域連合（組合）向け将来推計機能")
r = header(ws, r, ["項目", "定め", "", "", "", "", "", ""])
for k, v in MM.KOIKI.items():
    r = body(ws, r, [k, v, "", "", "", "", "", ""], height=46)
r = note(ws, r,
         "保険者単位の将来推計機能が実績値として用いる"
         "「介護保険事業状況報告」は構成市町村単位では存在しません。"
         "本計画がサービス見込量・給付費・保険料を保険者（広域連合）単位で"
         "算定し、町別には算定していないことと整合します。")
r += 1
r = lead(ws, r, "2　都道府県への提出")
r = header(ws, r, ["項目", "内容", "必須", "", "", "", "", ""])
for mei, naiyo, hissu in MM.TEISHUTSU:
    r = body(ws, r, [mei, naiyo, hissu, "", "", "", "", ""], height=40,
             align={3: "center"},
             fills={3: (IN_Y if hissu == "必須" else GRAY)})
r = note(ws, r, MM.TEISHUTSU_GO +
         "\n「必要入所（利用）定員、入居定員総数調査票」は"
         "厚生労働省より別途配布されているものであり、"
         "当方は受領していません（資料提供依頼No.31）。")

# ============================================================ 11
ws = sheet(SH["torikumi"] + "_提出時に入力する取組",
           "都道府県への提出時に入力する「第10期計画の策定における取組」",
           "利用マニュアル【システム操作編② 将来推計】第11.0版 附録③"
           "によります。入力はチェックボックス又はラジオボタンによります。",
           [30, 6, 40, 44, 10, 10, 10, 10])
r = 4
for gun in ("（1）中長期を含むサービス見込量の推計",
            "（2）サービス提供体制の確保",
            "（3）医療・介護連携の推進"):
    r = lead(ws, r, gun)
    r = header(ws, r, ["群", "記号", "問", "選択肢", "", "", "", ""])
    for g, kigo, toi, sel in MM.torikumi_of(gun):
        r = body(ws, r, [g, kigo, toi, "／".join(sel), "", "", "", ""],
                 height=60, align={2: "center"})
    r += 1
r = note(ws, r,
         "（1）ア及びイの年度の選択肢は2030年度・2040年度・2050年度・"
         "その他の年度です。本計画は中長期的な推計として"
         "令和17年度（2035年度）及び令和22年度（2040年度）を掲げており、"
         "令和12年度（2030年度）及び令和32年度（2050年度）は"
         "算定しているものの計画本文には掲げていません。"
         "計画本文に掲げる年度を選択肢に合わせるかはご判断を要します"
         "（確認事項No.193）。\n"
         "（2）及び（3）は現に実施された検討・分析・協議の状況を問うもので"
         "あり、当方が代わりに判断することができません。"
         "入力そのものは発注者が行うものとして整理しています。", height=124)

# ============================================================ 12 自己点検
ws = sheet(SH["chk"] + "_自己点検", "自己点検",
           "1件でも不適合があると本資料の作成を止めます。",
           [6, 40, 28, 20, 8, 10, 10, 10], freeze="A4")

chk(1, "受領した原典が7件であること", "len(GENTEN)",
    "%d件" % len(MM.GENTEN), len(MM.GENTEN) == 7)
chk(2, "操作編②が276頁であること", "GENTEN",
    "%s頁" % MM.genten_of("操作編②")[3], MM.genten_of("操作編②")[3] == 276)
chk(3, "確認事項No.191が決着していること", "台帳の状態",
    BY_C[191][7], BY_C[191][7] == "解決")
chk(4, "確認事項No.192が決着していること", "台帳の状態",
    BY_C[192][7], BY_C[192][7] == "解決")
chk(5, "令和8年度の起点が令和8年5月報告であること", "JISSEKI_KIKAN",
    MM.JISSEKI_KIKAN["令和8年度"][1],
    "令和8年5月報告" in MM.JISSEKI_KIKAN["令和8年度"][1])
chk(6, "地域支援事業費が対象経費支出予定額であること", "CHIIKI_HI",
    "対象経費支出予定額", "対象経費支出予定額" in MM.CHIIKI_HI)
chk(7, "構成3町の事業所が60件であること", "len(JIGYO)",
    "%d件" % len(JB.JIGYO), len(JB.JIGYO) == 60)
chk(8, "介護老人保健施設の定員が収録値と一致すること", "teiin('52')",
    "%d人" % JB.teiin("52"), JB.teiin("52") == 240)
chk(9, "地域密着型介護老人福祉施設の定員が収録値と一致すること",
    "teiin('54')", "%d人" % JB.teiin("54"), JB.teiin("54") == 62)
chk(10, "正誤表の正の値が3町計240人であること", "ROKEN_SEIGO",
    "%d人" % _SUM, _SUM == 240)
chk(11, "日常生活圏域が6圏域であること", "len(SEIKATSU_KEN)",
    "%d圏域" % len(MM.SEIKATSU_KEN), len(MM.SEIKATSU_KEN) == 6)
chk(12, "5年続けて利用者数が減る事業所が0件であること", "renzoku_gensho",
    "%d件" % len(_RENZOKU), len(_RENZOKU) == 0)
chk(13, "常勤換算数が2割以上減った事業所を掲げていること", "gensho",
    "%d件" % len(_GENIN), len(_GENIN) > 0)
chk(14, "推計1006の第9期保険料基準額が当方の再現と一致すること",
    "HOKENRYO['第9期']", "%.2f円" % K3.HOKENRYO["第9期"],
    abs(K3.HOKENRYO["第9期"] - 6427.89) < 0.01)
chk(15, "なお手当てを要する箇所に優先Aが4件あること", "NOKORU",
    "%d件" % sum(1 for x in K3.NOKORU if x[0] == "A"),
    sum(1 for x in K3.NOKORU if x[0] == "A") == 4)
chk(16, "附録③の問が3群8問であること", "TORIKUMI",
    "%d群%d問" % (len({t[0] for t in MM.TORIKUMI}), len(MM.TORIKUMI)),
    len({t[0] for t in MM.TORIKUMI}) == 3 and len(MM.TORIKUMI) == 8)
chk(17, "資料提供依頼No.25・No.26が受領済であること", "台帳の状態",
    "／".join(BY_L[n][8] for n in (25, 26)),
    all(BY_L[n][8] == "受領済" for n in (25, 26)))
chk(18, "資料提供依頼No.31が台帳にあること", "台帳",
    BY_L[31][2] if 31 in BY_L else "なし", 31 in BY_L)
chk(19, "確認事項No.193が台帳にあり未決であること", "台帳",
    BY_C[193][7] if 193 in BY_C else "なし",
    193 in BY_C and BY_C[193][7] == "未回答")
chk(20, "解決した確認事項が全て決着していること", "台帳の状態",
    "／".join(BY_C[n][7] for n, _m, _s in KAIKETSU),
    all(BY_C[n][7] in ("解決", "受託者案で確定", "完了", "了承済")
        for n, _m, _s in KAIKETSU))

# 本資料に書き出した値の走査
_TXT = "\n".join(
    str(c.value) for s in wb.worksheets for row in s.iter_rows()
    for c in row if c.value is not None)
_NAIBU = [w for w in ("runpy", ".py", "固定値", "実物から", "再実行",
                      "リポジトリ", "書き写し", "終了コード")
          if w in _TXT]
chk(21, "受託者の内部の仕組みの語が本資料に出ていないこと", "語で探す",
    "／".join(_NAIBU) or "当たり0件", not _NAIBU)
_KINSHI = [w for w in ("に由来する", "と整合する", "1件も",
                       "有意差がないため", "全国トップ級") if w in _TXT]
chk(22, "禁止表現が本資料に出ていないこと", "語で探す",
    "／".join(_KINSHI) or "当たり0件", not _KINSHI)
PI = [re.compile(r"0\d{1,4}[-(]\d{1,4}[-)]\d{3,4}"),
      re.compile(r"[\w.+-]+@[\w-]+\.[\w.]+")]
_pi = [p.pattern for p in PI if p.search(_TXT)]
chk(23, "電話番号・メールアドレスの形が本資料に出ていないこと", "値の形で探す",
    "%d件" % len(_pi), not _pi)
chk(24, "アカウントの識別子が本資料に出ていないこと", "語で探す",
    "当たり0件" if "W0183" not in _TXT else "当たりあり", "W0183" not in _TXT)
_mk = re.findall(r"\*\*", _TXT)
chk(25, "強調の指定が書き出した値に残っていないこと", "** の有無",
    "%d件" % len(_mk), not _mk)

r = 4
r = header(ws, r, ["No.", "確かめたこと", "式・見るところ", "結果", "判定",
                   "", "", ""])
for no, naiyo, shiki, kekka, han in CHECKS:
    r = body(ws, r, [no, naiyo, shiki, kekka, han, "", "", ""],
             height=30, align={1: "center", 5: "center"},
             fills={5: (OK_G if han == "適合" else NG_O)})
_ng = [c for c in CHECKS if c[4] != "適合"]
r = body(ws, r, ["", "自己点検の結果", "",
                 "適合%d件・不適合%d件" % (len(CHECKS) - len(_ng), len(_ng)),
                 "適合" if not _ng else "不適合", "", "", ""],
         height=22, bold=True, align={5: "center"},
         fills={5: (OK_G if not _ng else NG_O)})

wb.save(OUT)
print("書き出しました:", OUT)
for s in wb.sheetnames:
    print("  -", s, wb[s].max_row, "rows")
print("受領 原典%d件＋事業所分析データ4項目＋推計結果＋画面の写し"
      % len(MM.GENTEN))
print("解決した確認事項 %d件／備考を更新した確認事項 %d件"
      % (len(KAIKETSU), len(KOSHIN)))
print("構成3町の事業所 %d件（定員の欄があるもの %d件）"
      % (len(JB.JIGYO), sum(1 for v in JB.JIGYO.values() if "定員" in v)))
print("自己点検 %d件：適合%d件・不適合%d件"
      % (len(CHECKS), len(CHECKS) - len(_ng), len(_ng)))
if _ng:
    for c in _ng:
        print("  不適合 No.%s %s → %s" % (c[0], c[1], c[3]))
    sys.exit(1)
print("すべての点検に適合しました。")
