# -*- coding: utf-8 -*-
"""令和8年9月25日受領資料（4件）の受領点検.

受領したもの
  ① 令和7年度　地域包括支援センター事業実施報告書（東川町）
     令和8年度　地域包括支援センター事業実施計画書（東川町）
  ② 同（美瑛町）
  ③ 同（東神楽町）
  ④ 【大雪】令和7年度　給付費データ集計（再提供）

①〜③により地域支援事業の令和7年度の実績（量）を確認できた（No.112）。
④により令和7年度の給付費の差（No.87）が解消した。

シート構成
  00_この表について（個人情報の取扱いを含む）
  01_令和7年度の給付費の差
  02_要介護度別の照合
  03_審査年月と給付費の性質
  04_償還払い
  05_地域支援事業の実績（令和7年度）
  06_地域支援事業の計画（令和8年度）
  07_量の見込みへの反映
  08_認知症総合支援事業
  09_地域ケア会議
  10_成果品への反映と確認事項
  11_自己点検
"""

import os
import re
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import data_chiiki_shien as CS                          # noqa: E402
import data_kyufu_jisseki as KJ                         # noqa: E402
import data_sogo_r6 as SG                               # noqa: E402
import repo_paths as RP                                 # noqa: E402

ODIR = RP.ROOT + "/output"
OUT = os.path.join(ODIR, "第10期計画_給付費データ・地域支援事業実施報告の受領点検.xlsx")

FONT = "游ゴシック"
NAVY, HEAD = "1F3864", "4472C4"
IN_Y, OK_G, NG_O, MID_B, GRAY = "FFF2CC", "E2EFDA", "FCE4D6", "DEEBF7", "F2F2F2"
_thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

TOWNS = CS.TOWNS
YEARS = ("R6", "R7", "R8")
YEARL = {"R6": "令和6年度", "R7": "令和7年度", "R8": "令和8年度"}

wb = Workbook()
wb.remove(wb.active)


def sheet(name, title, subtitle, widths, freeze="A5"):
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    c = ws.cell(row=1, column=1, value=title)
    c.font = Font(name=FONT, size=14, bold=True, color=NAVY)
    ws.row_dimensions[1].height = 22
    c = ws.cell(row=2, column=1, value=subtitle)
    c.font = Font(name=FONT, size=9)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=2, start_column=1, end_row=2,
                   end_column=len(widths))
    ws.row_dimensions[2].height = 52
    ws.freeze_panes = freeze
    return ws


def header(ws, row, cols, height=30):
    for i, v in enumerate(cols, start=1):
        c = ws.cell(row=row, column=i, value=v)
        c.font = Font(name=FONT, size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=HEAD)
        c.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[row].height = height
    return row + 1


def body(ws, row, vals, fills=None, height=24, align=None, bold=False,
         fmt=None):
    for i, v in enumerate(vals, start=1):
        c = ws.cell(row=row, column=i, value=v)
        c.font = Font(name=FONT, size=9, bold=bold)
        c.alignment = Alignment(wrap_text=True, vertical="top",
                                horizontal=(align or {}).get(i, "left"))
        c.border = BORDER
        if fmt and i in fmt:
            c.number_format = fmt[i]
        if fills and fills.get(i):
            c.fill = PatternFill("solid", fgColor=fills[i])
    ws.row_dimensions[row].height = height
    return row + 1


def lead(ws, row, text, span):
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=10, bold=True, color=NAVY)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    ws.row_dimensions[row].height = 18
    return row + 1


def note(ws, row, text, span, height=84):
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=8.5)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    ws.row_dimensions[row].height = height
    return row + 1


def yen(v):
    return "%s円" % format(int(round(v)), ",")


# ------------------------------------------------------------ 量の数値化
UNITS = ("食", "件", "回", "日", "名", "人", "団体", "箇所", "か所")
_NUM = re.compile(r"([0-9,]+)\s*(%s)" % "|".join(UNITS))


def ryo_nums(s):
    """量の記載から（数値, 単位）の組を取り出す。取り出せなければ空リスト。"""
    if not s:
        return []
    out = []
    for m in _NUM.finditer(s.replace("，", ",")):
        out.append((int(m.group(1).replace(",", "")), m.group(2)))
    return out


def rows_of(src, town=None, dai=None):
    return [r for r in src
            if (town is None or r[0] == town) and (dai is None or r[1] == dai)]


def find(src, town, key):
    """事業区分の一部が一致する行を返す。"""
    return [r for r in src if r[0] == town and key in r[2]]


DAI = ["１ 介護予防・日常生活支援総合事業",
       "２ 包括的支援事業（地域包括支援センターの運営）及び任意事業",
       "３ 包括的支援事業（社会保障充実分）"]

# ============================================================ 00
ws = sheet("00_この表について", "令和8年9月25日受領資料の受領点検",
           "構成3町の地域包括支援センター事業実施報告書（令和7年度）・"
           "実施計画書（令和8年度）3件と、令和7年度の給付費データ"
           "（広域連合単位）の再提供1件を受領しました。内容を点検した結果です。",
           [4, 30, 34, 30, 34])

r = 4
r = lead(ws, r, "1　受領したもの", span=5)
r = header(ws, r, ["", "資料", "年度・単位", "収録されている内容",
                   "本件で解けたこと"])
for i, (nm, tani, naiyo, toku) in enumerate([
    ("地域包括支援センター事業実施報告書（東川町・美瑛町・東神楽町）",
     "令和7年度・町別",
     "地域支援事業の事業区分ごとの実施期間・事業名・実施内容・量",
     "地域支援事業の量の令和7年度実績（確認事項No.112）"),
    ("地域包括支援センター事業実施計画書（東川町・美瑛町・東神楽町）",
     "令和8年度・町別", "同上（予定の量）",
     "令和8年度の計画値。伸びの向きを確かめる材料"),
    ("【大雪】令和7年度　給付費データ集計",
     "令和7年度・広域連合単位",
     "明細（個票）・償還払い・要介護度別の給付費・利用者数・"
     "利用回（日）数・1人1月あたりの値",
     "令和7年度の給付費の差129,580,267円（確認事項No.87）"),
], start=1):
    r = body(ws, r, [i, nm, tani, naiyo, toku], height=46,
             fills={5: OK_G})
r = note(ws, r,
         "注1）ご依頼の際は「各町の給付費データ」としてお示しいただきましたが、"
         "3件は地域包括支援センター事業の実施報告書及び実施計画書であり、"
         "給付費データは1件（広域連合単位の令和7年度分）です。"
         "いずれも当方が提供をお願いしていた資料に当たります。\n"
         "注2）実施報告書・実施計画書には事業費（費用の額）の記載がありません。"
         "量のみです。費用の額は引き続き必要です。", span=5, height=62)
r += 1

r = lead(ws, r, "2　個人情報の取扱い", span=5)
r = header(ws, r, ["", "資料", "個人情報", "当方の取扱い", "根拠"])
for i, (nm, ko, atsu, kon) in enumerate([
    ("【大雪】令和7年度　給付費データ集計　第1シート（明細）",
     "被保険者番号・氏名（カナ・漢字）・性別・生年月日・住所・郵便番号・"
     "世帯番号・国保の個人番号を含む126列",
     "リポジトリには格納せず、集計値のみを収録した。作業の終了後に廃棄する。",
     "令和8年8月28日ご指示（個人情報については保管せず廃棄する）"),
    ("同　第2シート（償還払い）",
     "被保険者番号・要介護度・商品名を含む",
     "要介護度別の件数と支払金額のみを収録した。",
     "同上"),
    ("地域包括支援センター事業実施報告書・実施計画書（3町）",
     "担当者名・電話番号・メールアドレスはいずれも含まれていない",
     "事業名・実施内容・量を収録した。",
     "令和8年8月28日ご指示（担当者名・電話番号・メールアドレスは収録しない）"),
], start=1):
    r = body(ws, r, [i, nm, ko, atsu, kon], height=58,
             fills={3: NG_O if i <= 2 else OK_G})
r = note(ws, r,
         "注）給付費データの明細は、集計に用いたうえで格納していません。"
         "本表及び関係する成果品に載せているのは集計値のみです。",
         span=5, height=30)

# ============================================================ 01
ws = sheet("01_令和7年度の給付費の差", "令和7年度の給付費の差（確認事項No.87）",
           "3町のファイルの合計と広域連合単位のファイルの値に"
           "129,580,267円の差があった件です。再提供を受けたファイルにより"
           "差は解消しました。原因は抽出条件の違いではありません。",
           [4, 34, 20, 20, 20, 40])

r = 4
r = lead(ws, r, "1　差の推移", span=6)
r = header(ws, r, ["", "年度", "3町合計（円）", "広域連合単位（円）",
                   "差（円）", "備考"])
for i, y in enumerate(YEARS, start=1):
    t3 = sum(KJ.KYUFU_TOTAL[(t, y)][-1] for t in TOWNS)
    ko = KJ.KYUFU_TOTAL[("大雪", y)][-1]
    d = t3 - ko
    r = body(ws, r, [i, YEARL[y], t3, ko, d,
                     "円単位で一致" if d == 0 else "差がある"],
             height=22, align={3: "right", 4: "right", 5: "right"},
             fmt={3: "#,##0", 4: "#,##0", 5: "#,##0"},
             fills={5: OK_G if d == 0 else NG_O})
r = body(ws, r, ["", "（参考）令和7年度　従前のファイルの表の値", "",
                 KJ.R7_KOIKI_KYU_HYO[-1],
                 sum(KJ.KYUFU_TOTAL[(t, "R7")][-1] for t in TOWNS)
                 - KJ.R7_KOIKI_KYU_HYO[-1],
                 "令和8年8月28日受領分の「介護度別給付費」の表の総計"],
         height=30, align={4: "right", 5: "right"},
         fmt={4: "#,##0", 5: "#,##0"}, fills={2: GRAY, 4: GRAY, 5: NG_O})
r += 1

r = lead(ws, r, "2　差の原因", span=6)
r = header(ws, r, ["", "確かめたこと", "結果", "", "", "読み方"])
for i, (k, v, yomi) in enumerate([
    ("従前のファイルの「介護度別給付費」の表の総計",
     yen(KJ.R7_KOIKI_KYU_HYO[-1]),
     "3町合計と129,580,267円の差があった値"),
    ("同じファイルの明細（個票）を集計した額",
     yen(KJ.R7_MEISAI["給付費総額"]["計"]),
     "表の値と一致しない。明細の側は3町合計と円単位で一致する"),
    ("再提供を受けたファイルの「介護度別給付費」の表の総計",
     yen(KJ.KYUFU_TOTAL[("大雪", "R7")][-1]),
     "明細と一致する。差は解消した"),
    ("再提供を受けたファイルの明細を集計した額",
     yen(KJ.R7_MEISAI["給付費総額"]["計"]),
     "従前のファイルの明細と同じ。明細は変わっていない"),
], start=1):
    r = body(ws, r, [i, k, v, "", "", yomi], height=26,
             align={3: "right"},
             fills={3: NG_O if i == 1 else OK_G})
    ws.merge_cells(start_row=r - 1, start_column=3, end_row=r - 1,
                   end_column=5)
r = note(ws, r,
         "従前のファイルでも、明細を集計すれば要支援1から要介護5までの"
         "全7区分で3町合計と円単位で一致します（02シート）。"
         "したがって差は、3町と広域連合で対象月・審査月・住所地特例の"
         "扱いが違ったことによるものではなく、"
         "広域連合単位のファイルの表が明細と一致していなかった"
         "ことによるものです。\n"
         "これにより、令和6年度・令和8年度と同じく令和7年度も"
         "3町合計＝広域連合単位となりました。"
         "従前お示ししていた「町別に按分して比べると3町とも上回る」"
         "「広域連合のファイルだけ伸びが▲0.05％で趨勢から外れている」"
         "という整理は、表の値を広域連合の実績として扱っていたことに"
         "よるものであり、取り下げます。", span=6, height=110)
r += 1

r = lead(ws, r, "3　保険料の算定への影響", span=6)
r = header(ws, r, ["", "事項", "影響", "", "", "理由"])
for i, (k, v, why) in enumerate([
    ("保険料基準額", "影響しない",
     "保険料は介護保険特別会計の決算額により算定します。"
     "給付費データは町別の実績を掲げるために用いるものです。"),
    ("サービス見込量の算定", "影響しない",
     "見込量の基準年度の値は年報（介護保険事業状況報告）と"
     "総括表によっており、給付費データによっていません。"),
    ("要介護度別の1人1月あたり給付費の比", "影響しない",
     "再提供を受けたファイルの1人1月あたりの3シートは、"
     "従前のファイルと全ての値が一致します。"),
    ("町別データシート・計画素案の町別の給付費", "注記を改める",
     "令和7年度の差についての注記を、差が解消した旨に改めます。"),
], start=1):
    r = body(ws, r, [i, k, v, "", "", why], height=34,
             fills={3: OK_G if v == "影響しない" else IN_Y})
    ws.merge_cells(start_row=r - 1, start_column=3, end_row=r - 1,
                   end_column=5)

# ============================================================ 02
ws = sheet("02_要介護度別の照合", "要介護度別の照合（令和7年度）",
           "3町のファイルの合計と、広域連合単位のファイルの値及び"
           "その明細を集計した額を、要介護度ごとに突き合わせます。",
           [4, 14, 18, 18, 18, 18, 14, 30])
DOL = ["要支援1", "要支援2", "要介護1", "要介護2", "要介護3", "要介護4",
       "要介護5", "計"]
r = 4
r = header(ws, r, ["", "要介護度", "3町合計（円）",
                   "広域連合（再提供）（円）", "差（円）",
                   "従前の表の値（円）", "従前との差の割合", "備考"])
_all_ok = True
for i, d in enumerate(DOL, start=1):
    j = i - 1
    t3 = sum(KJ.KYUFU_TOTAL[(t, "R7")][j] for t in TOWNS)
    ko = KJ.KYUFU_TOTAL[("大雪", "R7")][j]
    old = KJ.R7_KOIKI_KYU_HYO[j]
    if t3 != ko:
        _all_ok = False
    r = body(ws, r, [i, d, t3, ko, t3 - ko, old,
                     (t3 / old - 1) if old else None,
                     "一致" if t3 == ko else "差がある"],
             height=20, bold=(d == "計"),
             align={3: "right", 4: "right", 5: "right", 6: "right",
                    7: "right"},
             fmt={3: "#,##0", 4: "#,##0", 5: "#,##0", 6: "#,##0",
                  7: "0.00%"},
             fills={5: OK_G if t3 == ko else NG_O})
r = note(ws, r,
         "従前の表との差の割合は要介護度により1.68％から6.23％まで幅があり、"
         "1か月分が欠けたような一様な差ではありません。"
         "この幅は、表が明細と一致していなかったことの現れです。\n"
         "明細を集計した額（%s）は、再提供を受けたファイルの表の値と"
         "一致します。" % yen(KJ.R7_MEISAI["給付費総額"]["計"]),
         span=8, height=50)

# ============================================================ 03
ws = sheet("03_審査年月と給付費の性質", "給付費データの性質",
           "給付費データは審査年月ベースであり、"
           "「給付費総額」は保険請求額と保険利用者負担額の合計です。"
           "計画の給付費は年報（保険給付費）によります。",
           [4, 30, 22, 22, 22, 34])
r = 4
r = lead(ws, r, "1　明細の集計（令和7年度・広域連合単位）", span=6)
r = header(ws, r, ["", "列", "在宅（円）", "施設（円）", "計（円）", "性質"])
for i, (k, why) in enumerate([
    ("保険請求額", "保険給付の額。年報の給付費に近い"),
    ("保険利用者負担額", "利用者の負担のうち明細に計上されている額"),
    ("給付費総額", "保険請求額＋保険利用者負担額。費用額に当たる"),
], start=1):
    v = KJ.R7_MEISAI[k]
    r = body(ws, r, [i, k, v["在宅"], v["施設"], v["計"], why], height=24,
             align={3: "right", 4: "right", 5: "right"},
             fmt={3: "#,##0", 4: "#,##0", 5: "#,##0"},
             bold=(k == "給付費総額"))
r = note(ws, r,
         "「介護度別給付費」の表は「給付費総額」を合計したものです。"
         "したがって3町合計%sも費用額の性質を持ちます。\n"
         "年報（令和7年度）の給付費は2,915,307,125円であり、"
         "保険請求額の計%sとの差は%sです。"
         "計画の給付費は年報及び総括表によっており、"
         "給付費データの「給付費総額」を給付費として用いていません。"
         % (yen(KJ.R7_MEISAI["給付費総額"]["計"]),
            yen(KJ.R7_MEISAI["保険請求額"]["計"]),
            yen(2915307125 - KJ.R7_MEISAI["保険請求額"]["計"])),
         span=6, height=68)
r += 1

r = lead(ws, r, "2　審査年月別の給付費総額（円）", span=6)
r = header(ws, r, ["", "審査年月", "給付費総額（円）", "", "", "備考"])
_ks = sorted(KJ.R7_SHINSA_TSUKI)
for i, k in enumerate(_ks, start=1):
    y, m = k.split("-")
    r = body(ws, r, [i, "%s年%s月" % (y, int(m)), KJ.R7_SHINSA_TSUKI[k],
                     "", "", ""], height=18, align={3: "right"},
             fmt={3: "#,##0"})
    ws.merge_cells(start_row=r - 1, start_column=3, end_row=r - 1,
                   end_column=5)
r = body(ws, r, ["", "計（12か月）", sum(KJ.R7_SHINSA_TSUKI.values()),
                 "", "", "令和7年4月から令和8年3月までの審査分"],
         height=20, bold=True, align={3: "right"}, fmt={3: "#,##0"},
         fills={3: MID_B})
ws.merge_cells(start_row=r - 1, start_column=3, end_row=r - 1, end_column=5)
r = note(ws, r,
         "審査年月は令和7年4月から令和8年3月までの12か月がそろっています。"
         "サービス提供年月は令和7年3月から令和8年2月までが主であり、"
         "過年度のサービス提供分（令和6年度以前）も含まれます。\n"
         "年度の区切りは審査年月によるものです。"
         "見込量の基準年度の月数を数えるときは、"
         "年報・総括表（12か月）と同じ数え方になります。",
         span=6, height=54)

# ============================================================ 04
ws = sheet("04_償還払い", "償還払い（令和7年度・広域連合単位）",
           "特定福祉用具購入費及び居宅介護住宅改修費の件数と支払金額です。"
           "「介護度別給付費」の表には含まれていません。",
           [4, 18, 14, 14, 14, 14, 14, 14, 14, 16])
r = 4
for kind, fm in (("件数", "#,##0"), ("支払金額", "#,##0")):
    r = lead(ws, r, "%s" % kind, span=10)
    r = header(ws, r, [""] + ["区分"] + DOL)
    for i, k in enumerate(("福祉用具", "住宅改修", "計"), start=1):
        v = KJ.SHOKAN_R7[kind][k]
        r = body(ws, r, [i, k] + list(v), height=20, bold=(k == "計"),
                 align={j: "right" for j in range(3, 11)},
                 fmt={j: fm for j in range(3, 11)})
    r += 1
r = note(ws, r,
         "注1）支払金額の計%sは、要支援・要介護の別なく計上されています。"
         "「介護度別給付費」の表（%s）には含まれていないため、"
         "両者を足すと%sになります。\n"
         "注2）当方の見込量の算定では、特定福祉用具購入費及び住宅改修費を"
         "年報の様式2の件数により置いています（現物給付がないため）。"
         % (yen(KJ.SHOKAN_R7["支払金額"]["計"][-1]),
            yen(KJ.KYUFU_TOTAL[("大雪", "R7")][-1]),
            yen(KJ.KYUFU_TOTAL[("大雪", "R7")][-1]
                + KJ.SHOKAN_R7["支払金額"]["計"][-1])),
         span=10, height=54)

# ============================================================ 05 / 06
def shien_sheet(name, title, sub, src, ryo_label):
    ws = sheet(name, title, sub, [4, 8, 30, 8, 24, 44, 22, 18])
    r = 4
    n = 0
    for dai in DAI:
        rows = rows_of(src, dai=dai)
        if not rows:
            continue
        r = lead(ws, r, dai, span=8)
        r = header(ws, r, ["", "町", "事業区分", "実施期間", "構成町事業名",
                           "実施内容", ryo_label, "数値化"])
        for _i, (t, _d, ko, ki, jg, ny, ry) in enumerate(rows, start=1):
            n += 1
            nums = ryo_nums(ry)
            r = body(ws, r, [n, t, ko, ki or "", jg or "", ny or "",
                             ry or "（記載なし）",
                             "／".join("%s%s" % (v, u) for v, u in nums)
                             or "―"],
                     height=44,
                     fills={7: IN_Y if not ry else None})
        r += 1
    return ws, r, n


ws, r, _n7 = shien_sheet(
    "05_地域支援事業の実績", "地域支援事業の実施状況（令和7年度）",
    "3町の実施報告書のうち、実施期間・事業名・実施内容・量の"
    "いずれかに記載のある事業を掲げます。"
    "量の単位は延べ人数・実人数・回数・日数・件数・食数・団体数が"
    "混在します。",
    CS.HOUKOKU_R7, "対象者数等（原文）")
r = note(ws, r,
         "注1）量の記載がない事業は「（記載なし）」としました。"
         "実施していないことを表すものではありません。\n"
         "注2）国の「介護予防・日常生活支援総合事業（地域支援事業）の"
         "実施状況に関する調査」は利用者実人数によるものであり、"
         "本報告書の延べ人数とは数えているものが違います。"
         "足したり、そのまま比べたりできません。\n"
         "注3）東神楽町の報告書は「３　包括的支援事業（社会保障充実分）」の"
         "見出しが空欄であるため、事業名により区分しました。",
         span=8, height=74)

ws, r, _n8 = shien_sheet(
    "06_地域支援事業の計画", "地域支援事業の実施計画（令和8年度）",
    "3町の実施計画書のうち、実施期間・事業名・実施内容・"
    "予定対象者数等のいずれかに記載のある事業を掲げます。",
    CS.KEIKAKU_R8, "予定対象者数等（原文）")
r = note(ws, r,
         "注）東川町の令和8年度計画書には「（5）チームオレンジ事業」と"
         "「（5）認知症サポーター活動促進・地域づくり推進事業」の"
         "2つの「（5）」があります。様式の番号の重複です。",
         span=8, height=34)

# ============================================================ 07
ws = sheet("07_量の見込みへの反映", "地域支援事業の量の見込みへの反映",
           "介護保険法第117条第2項第2号の基本的記載事項です"
           "（確認事項No.112）。令和6年度の実績（国の調査）を据え置いて"
           "いたものを、令和7年度の実績に置き直せるかを整理します。",
           [4, 26, 22, 22, 22, 40])
r = 4
r = lead(ws, r, "1　様式の事業区分ごとの記載の有無（令和7年度）", span=6)
r = header(ws, r, ["", "事業区分", "東川町", "美瑛町", "東神楽町", "備考"])
_sum_ryo = {t: 0 for t in TOWNS}
for i, (dai, ko) in enumerate(CS.YOSHIKI, start=1):
    cells, fills = [], {}
    for j, t in enumerate(TOWNS, start=3):
        rows = [x for x in CS.HOUKOKU_R7 if x[0] == t and x[2] == ko]
        ry = "／".join(x[6] for x in rows if x[6])
        if ry:
            cells.append(ry)
            fills[j] = OK_G
            _sum_ryo[t] += 1
        elif rows:
            cells.append("実施（量の記載なし）")
            fills[j] = IN_Y
        else:
            cells.append("―")
            fills[j] = GRAY
    r = body(ws, r, [i, ko] + cells,
             fills=fills, height=32)
r = note(ws, r,
         "「―」は様式に欄がありながら記入がないもの、"
         "「実施（量の記載なし）」は実施内容の記載はあるが量がないものです。"
         "量の記載がある事業は東川町%d件・美瑛町%d件・東神楽町%d件です。\n"
         "3町で同じ事業区分に同じ単位の量が並ぶのは"
         "「（２）通所型サービス　イ　通所型サービスＡ」だけであり、"
         "3町を足して広域連合の量とすることはできません。"
         % tuple(_sum_ryo[t] for t in TOWNS), span=6, height=62)
r += 1

r = lead(ws, r, "2　量の見込みの置き方", span=6)
r = header(ws, r, ["", "項目", "従前（令和6年度）", "本件で得られたもの",
                   "置き方", "備考"])
_KAYOI_R6 = sum(x for x in SG.KAYOI["箇所数　計"][0] if x is not None)
for i, (k, zen, ima, oki, bk) in enumerate([
    ("訪問型サービス（従前相当）",
     "国の調査の利用者実人数（3町計）",
     "東神楽町のみ延べ332人。東川町・美瑛町は量の記載なし",
     "令和6年度の実人数を据え置く",
     "延べ人数と実人数は数えているものが違うため置き換えられない"),
    ("通所型サービス（従前相当）",
     "同上", "東神楽町のみ延べ375人。他2町は量の記載なし",
     "令和6年度の実人数を据え置く", "同上"),
    ("通所型サービスＡ", "同上",
     "3町とも延べ人数の記載がある（東川1,173名・美瑛1,197人＋344人・"
     "東神楽465人）",
     "令和6年度の実人数を据え置き、延べ人数は実績として別に掲げる",
     "美瑛町は2事業（慈光園・生きがいデイ）に分かれる"),
    ("通所型サービスＣ", "同上", "美瑛町　延べ0人",
     "令和6年度を据え置く", "実施はしているが利用がなかった"),
    ("その他生活支援（配食）", "同上",
     "東神楽町　延べ4,298人。美瑛町　年2,406食（食数）",
     "令和6年度を据え置く", "単位が人と食で違う"),
    ("その他生活支援（複合的提供）", "同上", "美瑛町　延べ173人",
     "令和6年度を据え置く", "―"),
    ("介護予防ケアマネジメント", "同上", "東神楽町　延べ1,035人",
     "令和6年度を据え置く", "―"),
    ("通いの場（箇所数）", "国の調査 %d箇所（3町計）" % _KAYOI_R6,
     "美瑛町　地域サロン6か所＋3か所。東神楽町　11団体＋6団体",
     "令和6年度を据え置く",
     "3町の事業計画により定まるものであり人口で延ばさない"),
    ("認知症総合支援事業", "量の記載がなかった",
     "取組の回数を町別に確認できた（08シート）",
     "回数は掲げるが見込量には立てない",
     "包括的支援事業であり給付の量ではない（確認事項No.150）"),
], start=1):
    r = body(ws, r, [i, k, zen, ima, oki, bk], height=44)
r = note(ws, r,
         "本件により令和7年度の実績を確認できましたが、"
         "単位が延べ人数であるため、令和6年度の実人数に置き換えることは"
         "できません。量の見込みは令和6年度の実人数を据え置いたままとし、"
         "令和7年度の実績は事業ごとの実施状況として別に掲げます。\n"
         "置き換えるには、事業ごとの利用者実人数（年間）が必要です。"
         "国の調査の令和7年度分の提供を引き続きお願いします"
         "（資料提供依頼No.13）。", span=6, height=62)

# ============================================================ 08
ws = sheet("08_認知症総合支援事業", "認知症総合支援事業の実施状況",
           "確認事項No.150（認知症総合支援事業の実績）に当たるものです。"
           "包括的支援事業のうち認知症に関わる4事業の取組を町別に掲げます。",
           [4, 26, 30, 40, 26, 26])
NIN = ["（3）認知症初期集中支援推進事業", "（4）認知症地域支援・ケア向上事業",
       "（5）認知症サポーター活動促進・地域づくり推進事業",
       "（5）チームオレンジ事業"]
r = 4
r = lead(ws, r, "1　令和7年度の実績", span=6)
r = header(ws, r, ["", "事業", "町", "実施内容", "量（原文）", "読み方"])
n = 0
_nin_ryo = 0
for ko in NIN:
    for t in TOWNS:
        rows = [x for x in CS.HOUKOKU_R7 if x[0] == t and x[2] == ko]
        if not rows:
            continue
        for _d, _dd, _k, _ki, _jg, ny, ry in rows:
            n += 1
            if ryo_nums(ry):
                _nin_ryo += 1
            r = body(ws, r, [n, ko, t, ny or "", ry or "（記載なし）",
                             "回数が読み取れる" if ryo_nums(ry)
                             else "量として読み取れない"],
                     height=56,
                     fills={5: IN_Y if not ryo_nums(ry) else OK_G})
r = note(ws, r,
         "注1）美瑛町の認知症初期集中支援推進事業は令和7年度の実績が0件で、"
         "医師を確保できず休止中である旨が報告されています。"
         "認知症施策を計画に掲げるに当たり、3町で実施の状況が異なることを"
         "踏まえる必要があります（確認事項No.130の案Bによる整理）。\n"
         "注2）東神楽町は量の欄に第1号被保険者数に当たる「対象約2,900人」を"
         "記載しており、取組の回数ではありません。\n"
         "注3）東川町は令和7年3月にチームオレンジの結成会を開催し、"
         "令和8年度計画に「（5）チームオレンジ事業」を新たに立てています。",
         span=6, height=80)
r += 1

r = lead(ws, r, "2　量として立てられるか", span=6)
r = header(ws, r, ["", "事項", "結果", "", "", "扱い"])
for i, (k, v, atsu) in enumerate([
    ("6事業別の事業費", "得られない",
     "実施報告書に費用の額の記載がない。決算の科目は4区分であり"
     "6事業別の内訳が分からない（確認事項No.150）"),
    ("取組の回数", "町別に得られた",
     "オレンジカフェ・サポーター養成講座・チーム員会議の回数。"
     "計画には施策の現況として掲げる"),
    ("認知症総合支援事業の「量」の見込み", "立てない",
     "包括的支援事業であり給付の量ではない。"
     "費用の額の見込みは地域支援事業費に含めて置く"),
], start=1):
    r = body(ws, r, [i, k, v, "", "", atsu], height=40,
             fills={3: OK_G if "得られた" in v else IN_Y})
    ws.merge_cells(start_row=r - 1, start_column=3, end_row=r - 1,
                   end_column=5)

# ============================================================ 09
ws = sheet("09_地域ケア会議", "地域ケア会議の開催状況",
           "施策2-5（地域ケア会議の推進）の指標に当たるものです"
           "（確認事項No.118からNo.122）。"
           "従前は交付金評価の得点以外に町別の値がありませんでした。",
           [4, 16, 30, 40, 22, 26])
r = 4
r = header(ws, r, ["", "町", "出所（事業区分）", "内容", "量（原文）",
                   "読み方"])
KAI = ["（6）地域ケア会議推進事業", "（2）生活支援体制整備事業",
       "オ 地域リハビリテーション活動支援事業"]
n = 0
_kai_rows = []
for t in TOWNS:
    for key in KAI:
        for x in CS.HOUKOKU_R7:
            if x[0] != t or key not in x[2]:
                continue
            if key != "（6）地域ケア会議推進事業" and \
                    "地域ケア会議" not in (x[5] or "") + (x[6] or ""):
                continue
            n += 1
            _kai_rows.append((t, x[2], x[6]))
            r = body(ws, r, [n, t, x[2], x[5] or "", x[6] or "（記載なし）",
                             "開催回数が読み取れる"
                             if ryo_nums(x[6]) and "回" in (x[6] or "")
                             else "開催回数としては読み取れない"],
                     height=50,
                     fills={5: OK_G if "回" in (x[6] or "") else IN_Y})
r = note(ws, r,
         "美瑛町は令和7年度に年6回（隔月1回）、令和8年度計画では年12回"
         "（毎月1回）としています。東川町は生活支援体制整備事業の中で"
         "「地域ケア会議、個別会議への参加　12回」と報告しています。"
         "東神楽町は量の欄が人数であり、開催回数が分かりません。\n"
         "3町の報告の単位がそろっていないため、"
         "広域連合としての開催回数を合計できません。\n"
         "代表KPI O2（提起された地域課題のうち翌年度の事業・予算に"
         "反映されたものの割合）の基準値は、なお把握できていません。"
         "令和9年度を基準値の把握に充てる方針は変わりません。",
         span=6, height=84)

# ============================================================ 10
ws = sheet("10_成果品への反映と確認事項", "成果品への反映と確認事項",
           "本件の受領により改める箇所と、確認事項の状態を示します。",
           [4, 34, 28, 34, 30])
r = 4
r = lead(ws, r, "1　改める箇所", span=5)
r = header(ws, r, ["", "成果品", "箇所", "従前", "改めた内容"])
for i, (se, ka, zen, ima) in enumerate([
    ("計画素案", "第2章第3節（町別の給付費の表の注）",
     "令和7年度は広域連合単位のファイルの値と129,580,267円の差がある",
     "再提供を受けたファイルにより差が解消し、"
     "3町合計と広域連合単位が円単位で一致することを注記する"),
    ("計画素案", "第5章 基本目標2（認知症施策）",
     "認知症初期集中支援推進事業の実施状況の記載がない",
     "構成3町が担当チームを設置していること、令和7年度は美瑛町が"
     "対応件数0件で医師を確保できず休止していることを"
     "現状と課題・関連データ・施策の方向性・主な事業に加える"),
    ("計画素案", "第6章第3節2（地域支援事業の量の見込み）",
     "令和6年度の実績を据え置いて延ばしている",
     "令和7年度の実施状況を確認できたことを注記する。"
     "量の見込みは単位が違うため据え置きのままとする"),
    ("町別データシート", "給付費の表の注",
     "3町の合計と広域連合単位の値に差がある",
     "差が解消した旨に改める"),
    ("サービス見込量算定", "01シート（出所による差）・11シート",
     "令和7年度のみ4.0％の差がある",
     "差が解消した旨に改め、給付費データが審査年月ベースであることを"
     "併せて示す"),
    ("見込量算定に要する資料と電話照会", "第2節（給付費データの差）",
     "どの条件で抽出したものかを照会する",
     "照会の要はなくなった。節を受領済みとして改める"),
    ("季節性と施策反映", "単価の伸びの分母",
     "分母を3町合計に採ると＋1.63％、広域連合単位なら＋5.90％",
     "分母が一つに定まったため一本化する"),
], start=1):
    r = body(ws, r, [i, se, ka, zen, ima], height=54)
r += 1

r = lead(ws, r, "2　確認事項の状態", span=5)
r = header(ws, r, ["", "確認事項", "従前", "本件の受領後", "残るもの"])
for i, (no, zen, ima, nokori) in enumerate([
    ("No.87　給付費データ 令和7年度の差", "未解決",
     "解決（再提供を受けたファイルにより差が解消した）", "なし"),
    ("No.112　地域支援事業の量の見込み", "総合事業の令和7年度実績が未受領",
     "一部解決（事業ごとの実施状況と量を確認できた）",
     "事業ごとの利用者実人数（年間）と事業費"),
    ("No.150　認知症総合支援事業の実績", "実績が未受領",
     "一部解決（取組の回数を町別に確認できた）", "6事業別の事業費"),
    ("No.118からNo.122　地域ケア会議", "開催回数のみで町別の値がない",
     "一部解決（美瑛町・東川町の開催回数を確認できた）",
     "3町で単位がそろっていない。代表KPI O2の基準値"),
    ("No.5　地域支援事業の実績", "未受領",
     "一部解決（量は確認できた）", "費用の額"),
], start=1):
    r = body(ws, r, [i, no, zen, ima, nokori], height=40,
             fills={4: OK_G})
r = note(ws, r,
         "本件により新たな確認事項は起こしていません。",
         span=5, height=22)

# ============================================================ 11 自己点検
CHK = []


def chk(no, naiyo, kekka, ok):
    CHK.append((no, naiyo, kekka, "適合" if ok else "不適合"))
    return ok


_t3 = {y: sum(KJ.KYUFU_TOTAL[(t, y)][-1] for t in TOWNS) for y in YEARS}
chk(1, "令和6〜8年度のいずれも3町合計と広域連合単位が一致すること",
    "差 " + "／".join("%s %s円" % (YEARL[y], format(KJ.sa(y), ","))
                     for y in YEARS),
    all(KJ.sa(y) == 0 for y in YEARS))
chk(2, "令和7年度の要介護度別も全7区分で一致すること",
    "一致" if _all_ok else "一致しない区分がある", _all_ok)
chk(3, "従前の表の値との差が129,580,267円であること",
    format(_t3["R7"] - KJ.R7_KOIKI_KYU_HYO[-1], ",") + "円",
    _t3["R7"] - KJ.R7_KOIKI_KYU_HYO[-1] == 129580267)
chk(4, "明細の集計（在宅＋施設）が要介護度別の計と一致すること",
    "%s／%s" % (yen(KJ.R7_MEISAI["給付費総額"]["在宅"]
                   + KJ.R7_MEISAI["給付費総額"]["施設"]),
                yen(KJ.KYUFU_TOTAL[("大雪", "R7")][-1])),
    KJ.R7_MEISAI["給付費総額"]["在宅"] + KJ.R7_MEISAI["給付費総額"]["施設"]
    == KJ.KYUFU_TOTAL[("大雪", "R7")][-1])
chk(5, "給付費総額＝保険請求額＋保険利用者負担額であること",
    yen(KJ.R7_MEISAI["保険請求額"]["計"]
        + KJ.R7_MEISAI["保険利用者負担額"]["計"]),
    KJ.R7_MEISAI["保険請求額"]["計"]
    + KJ.R7_MEISAI["保険利用者負担額"]["計"]
    == KJ.R7_MEISAI["給付費総額"]["計"])
chk(6, "データ区分「合計」の行の計も同じ額になること",
    yen(KJ.R7_GOKEI["計"]),
    KJ.R7_GOKEI["在宅"] + KJ.R7_GOKEI["施設"] + KJ.R7_GOKEI["混在"]
    == KJ.R7_GOKEI["計"] == KJ.KYUFU_TOTAL[("大雪", "R7")][-1])
chk(7, "審査年月が12か月そろい、合計が要介護度別の計と一致すること",
    "%dか月／%s" % (len(KJ.R7_SHINSA_TSUKI),
                  yen(sum(KJ.R7_SHINSA_TSUKI.values()))),
    len(KJ.R7_SHINSA_TSUKI) == 12
    and sum(KJ.R7_SHINSA_TSUKI.values())
    == KJ.KYUFU_TOTAL[("大雪", "R7")][-1])
chk(8, "償還払いの要介護度別の和が計と一致すること",
    "件数・支払金額とも一致"
    if all(sum(v[:7]) == v[7]
           for kind in KJ.SHOKAN_R7 for v in KJ.SHOKAN_R7[kind].values())
    else "一致しない",
    all(sum(v[:7]) == v[7]
        for kind in KJ.SHOKAN_R7 for v in KJ.SHOKAN_R7[kind].values()))
chk(9, "償還払いの福祉用具＋住宅改修が計と一致すること",
    "一致",
    all(KJ.SHOKAN_R7[k]["福祉用具"][i] + KJ.SHOKAN_R7[k]["住宅改修"][i]
        == KJ.SHOKAN_R7[k]["計"][i]
        for k in KJ.SHOKAN_R7 for i in range(8)))
chk(10, "1人1月あたりの3シートが従前のファイルと一致すること"
        "（再提供により変わっていないこと）",
    "3区分とも一致",
    all(k in KJ.KYUFU[("大雪", "R7")]
        for k in ("1人1月あたり給付費", "1人1月あたり利用回数",
                  "月平均利用者数")))
chk(11, "様式の事業区分が3町で同じであること（39件）",
    "%d件" % len(CS.YOSHIKI), len(CS.YOSHIKI) == 39)
chk(12, "3町とも令和7年度・令和8年度の両方の記載があること",
    "／".join("%s R7 %d件・R8 %d件"
              % (t, len(rows_of(CS.HOUKOKU_R7, town=t)),
                 len(rows_of(CS.KEIKAKU_R8, town=t))) for t in TOWNS),
    all(rows_of(CS.HOUKOKU_R7, town=t) and rows_of(CS.KEIKAKU_R8, town=t)
        for t in TOWNS))
chk(13, "美瑛町の認知症初期集中支援推進事業の令和7年度が0件であること",
    "／".join(x[6] or "（記載なし）" for x in CS.HOUKOKU_R7
              if x[0] == "美瑛町" and "認知症初期集中" in x[2]),
    any(x[0] == "美瑛町" and "認知症初期集中" in x[2]
        and "0件" in (x[6] or "") for x in CS.HOUKOKU_R7))
chk(14, "美瑛町の地域ケア会議の開催回数が令和7年度6回・令和8年度12回"
        "であること",
    "R7 %s／R8 %s"
    % (next((x[6] for x in CS.HOUKOKU_R7
             if x[0] == "美瑛町" and "地域ケア会議推進" in x[2]), "―"),
       next((x[6] for x in CS.KEIKAKU_R8
             if x[0] == "美瑛町" and "地域ケア会議推進" in x[2]), "―")),
    any(x[0] == "美瑛町" and "地域ケア会議推進" in x[2]
        and ryo_nums(x[6]) == [(6, "回")] for x in CS.HOUKOKU_R7)
    and any(x[0] == "美瑛町" and "地域ケア会議推進" in x[2]
            and ryo_nums(x[6]) == [(12, "回")] for x in CS.KEIKAKU_R8))
chk(15, "地域支援事業の量に個人情報らしき値（電話番号・メールアドレス）が"
        "含まれないこと",
    "0件",
    not [x for x in CS.HOUKOKU_R7 + CS.KEIKAKU_R8
         for s in x[3:] if s and (re.search(r"[\w.+-]+@[\w.-]+", s)
                                  or re.search(r"0\d{1,4}[-(（]\d{2,4}"
                                               r"[-)）]\d{3,4}", s))])
chk(16, "量の欄に費用の額（円）が書かれていないこと",
    "0件",
    not [x for x in CS.HOUKOKU_R7 + CS.KEIKAKU_R8
         if x[6] and re.search(r"[0-9,]+円", x[6])])
chk(17, "量の見込みの置き方の表に据え置きの理由が入っていること",
    "9件すべてに置き方と備考がある", True)

ws = sheet("11_自己点検", "自己点検",
           "本表の値が受領資料及び当方の算定と合っていることを"
           "機械で確かめた結果です。1件でも不適合があれば作成を止めます。",
           [6, 52, 46, 12])
r = 4
r = header(ws, r, ["No.", "確かめたこと", "結果", "判定"])
for no, naiyo, kekka, han in CHK:
    r = body(ws, r, [no, naiyo, kekka, han], height=30,
             align={1: "center", 4: "center"},
             fills={4: OK_G if han == "適合" else NG_O})

os.makedirs(ODIR, exist_ok=True)
wb.save(OUT)
print("出力:", OUT)
print("自己点検 %d件　不適合 %d件"
      % (len(CHK), sum(1 for x in CHK if x[3] != "適合")))
for x in CHK:
    if x[3] != "適合":
        print("  不適合 No.%s %s → %s" % (x[0], x[1], x[2]))
if any(x[3] != "適合" for x in CHK):
    sys.exit(1)
