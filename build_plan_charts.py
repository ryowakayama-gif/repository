# -*- coding: utf-8 -*-
"""大雪地区広域連合 第10期介護保険事業計画 図表集（白黒レイアウト）生成スクリプト.

数値の出所は 00_凡例・出典 シートに明記。第9期計画（令和6年3月）及び
素案第9稿に掲載された見える化データのみを使用し、推測値は使用していない。
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.marker import Marker
from openpyxl.chart.data_source import StrRef
from openpyxl.chart.series import SeriesLabel
from openpyxl.drawing.line import LineProperties

FONT = "Carlito"
GRAY = ["000000", "404040", "737373", "A6A6A6", "D9D9D9", "F2F2F2", "FFFFFF"]
DASH = ["solid", "dash", "sysDot", "dashDot", "lgDash"]
MARK = ["circle", "square", "triangle", "diamond", "x"]
IN_Y = "FFF2CC"

thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
wb = Workbook()


def sheet(name, title, subtitle, widths):
    ws = wb.create_sheet(name)
    ws["A1"] = title
    ws["A1"].font = Font(name=FONT, size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="000000")
    ws["A2"] = subtitle
    ws["A2"].font = Font(name=FONT, size=9)
    ws["A2"].fill = PatternFill("solid", fgColor="F2F2F2")
    n = max(len(widths), 8)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n)
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 16
    return ws


def table(ws, row, head, rows, numfmt=None, headfill="404040"):
    for i, h in enumerate(head, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = Font(name=FONT, size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=headfill)
        c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
        c.border = BORDER
    r = row + 1
    for data in rows:
        for i, v in enumerate(data, start=1):
            c = ws.cell(row=r, column=i, value=v)
            c.font = Font(name=FONT, size=9)
            c.border = BORDER
            if i == 1:
                c.alignment = Alignment(horizontal="left", vertical="center")
            else:
                c.alignment = Alignment(horizontal="right", vertical="center")
                if numfmt:
                    c.number_format = numfmt
        r += 1
    return r - 1


def note(ws, row, text):
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=8, italic=True)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=13)
    ws.row_dimensions[row].height = 26


def mono_bar(ws, title, y_title, cats_ref, data_ref, anchor, stacked=False,
             width=22, height=11, gap=60, overlap=None):
    ch = BarChart()
    ch.type = "col"
    ch.style = None
    if stacked:
        ch.grouping = "stacked"
        ch.overlap = 100
    elif overlap is not None:
        ch.overlap = overlap
    ch.title = title
    ch.y_axis.title = y_title
    ch.add_data(data_ref, titles_from_data=True)
    ch.set_categories(cats_ref)
    ch.gapWidth = gap
    ch.width, ch.height = width, height
    for i, s in enumerate(ch.series):
        gp = GraphicalProperties(solidFill=GRAY[i % 5])
        gp.line = LineProperties(solidFill="000000", w=6350)
        s.graphicalProperties = gp
    _axis_mono(ch)
    ws.add_chart(ch, anchor)
    return ch


def mono_hbar(ws, title, cats_ref, data_ref, anchor, width=20, height=10):
    ch = BarChart()
    ch.type = "bar"
    ch.grouping = "percentStacked"
    ch.overlap = 100
    ch.title = title
    ch.add_data(data_ref, titles_from_data=True)
    ch.set_categories(cats_ref)
    ch.gapWidth = 60
    ch.width, ch.height = width, height
    for i, s in enumerate(ch.series):
        gp = GraphicalProperties(solidFill=GRAY[i % 7])
        gp.line = LineProperties(solidFill="000000", w=6350)
        s.graphicalProperties = gp
    _axis_mono(ch)
    ws.add_chart(ch, anchor)
    return ch


def mono_line(ws, title, y_title, cats_ref, data_ref, anchor,
              width=22, height=11, min_=None, max_=None):
    ch = LineChart()
    ch.style = None
    ch.title = title
    ch.y_axis.title = y_title
    ch.add_data(data_ref, titles_from_data=True)
    ch.set_categories(cats_ref)
    ch.width, ch.height = width, height
    if min_ is not None:
        ch.y_axis.scaling.min = min_
    if max_ is not None:
        ch.y_axis.scaling.max = max_
    for i, s in enumerate(ch.series):
        gp = GraphicalProperties()
        gp.line = LineProperties(solidFill="000000", w=22225,
                                 prstDash=DASH[i % len(DASH)])
        s.graphicalProperties = gp
        s.marker = Marker(symbol=MARK[i % len(MARK)], size=6)
        s.marker.graphicalProperties = GraphicalProperties(solidFill=GRAY[i % 5])
        s.marker.graphicalProperties.line = LineProperties(solidFill="000000", w=6350)
        s.smooth = False
    _axis_mono(ch)
    ws.add_chart(ch, anchor)
    return ch


def _axis_mono(ch):
    for ax in (ch.x_axis, ch.y_axis):
        ax.majorGridlines = None
        ax.delete = False
        gp = GraphicalProperties()
        gp.line = LineProperties(solidFill="000000", w=6350)
        ax.graphicalProperties = gp
    ch.legend.position = "b"



def _named_series(src_ws, row_from, row_to):
    """08_ニーズ調査データ の1行から、系列名=C列（地域名）、値=E～J列（全体・年齢階級）の系列を作る。

    D列（サンプル数）を値に含めないため add_data ではなく Series を直接構成する。
    """
    ref = Reference(src_ws, min_col=5, max_col=10, min_row=row_from, max_row=row_to)
    s = Series(ref, title_from_data=False)
    s.tx = SeriesLabel(strRef=StrRef(f"'{src_ws.title}'!$C${row_from}"))
    return s


YEARS = ["平成24年", "平成25年", "平成26年", "平成27年", "平成28年", "平成29年",
         "平成30年", "令和元年", "令和2年", "令和3年", "令和4年", "令和5年"]

# ============================================================ 00 凡例・出典
ws = sheet("00_凡例・出典", "第10期介護保険事業計画 図表集（白黒レイアウト）",
           "令和8（2026）年7月28日作成／第9期計画の図表構成に準拠し、確認済みの数値のみで作図",
           [6, 26, 34, 44, 30])
r = 4
ws.cell(row=r, column=1, value="1　本図表集の考え方").font = Font(name=FONT, size=11, bold=True)
r += 1
note(ws, r, "第9期介護保険事業計画（令和6年3月）の図表構成をそのまま踏襲し、白黒印刷（本文1色）を前提としたグレースケール・線種・マーカーで作図しています。"
            "数値は第9期計画本体及び第10期計画素案第9稿に掲載された見える化データのみを使用し、推測値・補間値は使用していません。"
            "令和6・7年度の実績は未受領のため、該当セルは淡黄色の入力欄としています。")
r += 2
ws.cell(row=r, column=1, value="2　シート構成と出典").font = Font(name=FONT, size=11, bold=True)
r += 1
end = table(ws, r, ["No.", "シート", "図表", "出典・基準時点", "更新時の差替え"], [
    (1, "01_人口推移", "人口の推移（広域連合）", "第9期計画 第2章第1節1／住民基本台帳 各年10月1日", "R6～R8実績を追加"),
    (2, "02_高齢化率推移", "高齢化率の推移（広域連合・3町）", "第9期計画 第2章第1節1／住民基本台帳 各年10月1日", "R6～R8実績を追加"),
    (3, "03_高齢者世帯", "高齢者を含む世帯／構成割合", "第9期計画 第2章第1節2／国勢調査（H22・H27・R2）", "令和7年国勢調査の公表後に更新"),
    (4, "04_認定者数推移", "認定者数の推移（要介護度別）", "第9期計画 第2章第1節3／介護保険事業状況報告 各年9月分", "R6～R8の9月分を追加"),
    (5, "05_認定者割合", "認定者割合の比較（全国・道・3町）", "第9期計画 第2章第1節3／介護保険事業状況報告 令和5年9月分", "最新年9月分で更新"),
    (6, "06_出現率推移", "出現率の推移／上川管内比較", "第9期計画 第2章第1節3／介護保険事業状況報告 各年9月分", "R6～R8を追加。管内比較は最新年で更新"),
    (7, "07_給付費推移", "サービス区分別給付費／地域支援事業費／中長期推計", "第9期計画 第2章第2節（見える化 R6.1.18参照）／素案第9稿 表39", "R6～R8実績の受領後に更新"),
    (8, "08_ニーズ調査データ", "介護予防・日常生活圏域ニーズ調査 14指標の集計値", "第9期計画 第2章第3節／令和4年11月調査（回収4,626票・63.9％）", "第10期ニーズ調査の集計受領後に全面差替え"),
    (9, "09_ニーズ調査グラフ", "同上の年齢階級別・町別グラフ", "同上", "同上"),
])
r = end + 2
ws.cell(row=r, column=1, value="3　作図上の凡例（白黒）").font = Font(name=FONT, size=11, bold=True)
r += 1
end = table(ws, r, ["No.", "要素", "表現", "備考", ""], [
    (1, "積上げ棒・横棒", "黒→濃灰→中灰→淡灰→白の順に塗り分け、輪郭は黒", "5系列までは塗りのみで判別可能"),
    (2, "折れ線", "すべて黒線。実線／破線／点線／一点鎖線で系列を区別", "マーカーは○△□◇で重複を回避"),
    (3, "目盛線", "非表示。軸線のみ黒で表示", "1色印刷時の視認性を優先"),
    (4, "凡例", "グラフ下部に配置", "第9期計画の体裁を踏襲"),
    (5, "入力欄", "淡黄色。R6・R7実績など未受領の値", "受領後に上書きするとグラフに反映"),
])
r = end + 2
ws.cell(row=r, column=1, value="4　作図にあたり確認が必要な事項").font = Font(name=FONT, size=11, bold=True)
r += 1
end = table(ws, r, ["No.", "事項", "内容", "対応", ""], [
    (1, "人口の年齢区分", "第9期計画のグラフは0～39歳／40～64歳／65～74歳／75歳以上の4区分だが、"
        "計画本体のPDFから0～39歳・40～64歳の確定値を復元できなかった", "本図表集は0～64歳（総人口－65歳以上）／65～74歳／75歳以上の3区分で作図。"
        "4区分にする場合は住民基本台帳の元データが必要", "広域連合へ照会"),
    (2, "リスク点数の単位", "第9期計画 第2章第3節⑨は「当広域連合全体で15.8％」と記載しているが、"
        "実際は要支援・要介護リスク評価尺度による平均点（72,099点÷4,560人＝15.8点）", "本図表集では「点」として作図。第10期では単位表記を修正", "第10期計画で修正"),
    (3, "認定者数と出現率の基準月", "第9期計画は各年9月分、素案第9稿の直近値（認定者1,984人・21.8％）はR8年3月末",
        "基準月が異なるため同一系列に接続していない。R6～R8は9月分を受領して追加", "広域連合へ照会"),
    (4, "総給付費の定義", "第9期計画 第2章第2節の総給付費（R5 2,794,391千円）と、"
        "素案第9稿の保険給付費R5決算（2,940.4百万円）は定義が異なる", "別系列として併記。計画掲載時はいずれかに統一", "第10期計画で整理"),
    (5, "ニーズ調査の基準", "掲載値は令和4年11月実施の第9期調査。第10期調査は実施済だが集計未受領",
        "第10期調査の集計受領後に全面差替え。設問・判定ロジックの一致を確認", "広域連合・3町へ照会"),
])

# ============================================================ 01 人口推移
ws = sheet("01_人口推移", "図1　人口の推移（大雪地区広域連合）",
           "資料：住民基本台帳 各年10月1日現在（第9期計画 第2章第1節1）／令和6～8年は淡黄色欄に入力すると連動",
           [16] + [10] * 15)
POP_TOTAL = [28635, 28668, 28800, 28950, 28904, 28929, 28827, 28624, 28288, 28247, 28138, 27887]
POP_65_74 = [3744, 3862, 4032, 4119, 4183, 4166, 4210, 4215, 4269, 4233, 4165, 3986]
POP_75 = [4398, 4491, 4578, 4667, 4768, 4865, 4933, 5006, 5014, 5039, 5109, 5205]
POP_65 = [8142, 8353, 8610, 8786, 8951, 9031, 9143, 9221, 9283, 9272, 9274, 9191]

HD = ["区分"] + YEARS + ["令和6年", "令和7年", "令和8年"]
rows = [
    ["0～64歳"] + [None] * 15,
    ["65～74歳"] + POP_65_74 + [None] * 3,
    ["75歳以上"] + POP_75 + [None] * 3,
    ["65歳以上（再掲）"] + POP_65 + [None] * 3,
    ["総人口"] + POP_TOTAL + [None] * 3,
]
end = table(ws, 4, HD, rows, numfmt="#,##0")
for c in range(2, 17):
    col = get_column_letter(c)
    ws.cell(row=5, column=c).value = f"={col}9-{col}6-{col}7"
    ws.cell(row=5, column=c).number_format = "#,##0"
for r_ in (6, 7, 8, 9):
    for c in range(14, 17):
        ws.cell(row=r_, column=c).fill = PatternFill("solid", fgColor=IN_Y)
note(ws, end + 1, "注1）0～64歳は「総人口－65歳以上」で算出。第9期計画のグラフは0～39歳／40～64歳に分けているが、"
                  "計画本体から確定値を復元できなかったため3区分としている（00_凡例・出典 4-1）。"
                  "注2）65歳以上は65～74歳と75歳以上の合計。令和6～8年（淡黄色欄）は住民基本台帳の実績受領後に入力する。")
cats = Reference(ws, min_col=2, max_col=16, min_row=4)
data = Reference(ws, min_col=1, max_col=16, min_row=5, max_row=7)
mono_bar(ws, "人口の推移（大雪地区広域連合）", "人口（人）", cats, data, "A14", stacked=True, width=26, height=12)

# ============================================================ 02 高齢化率
ws = sheet("02_高齢化率推移", "図2　高齢化率の推移（広域連合・構成3町）",
           "資料：住民基本台帳 各年10月1日現在（第9期計画 第2章第1節1）／単位：％",
           [16] + [10] * 15)
RATE = {
    "広域連合": [28.4, 29.1, 29.9, 30.3, 31.0, 31.2, 31.7, 32.2, 32.8, 32.8, 33.0, 33.0],
    "東川町": [28.7, 29.7, 31.1, 31.5, 32.1, 32.0, 32.2, 32.1, 32.8, 32.3, 31.8, 31.4],
    "美瑛町": [33.6, 34.4, 35.2, 35.9, 36.5, 36.7, 37.3, 37.9, 38.4, 38.6, 38.7, 38.8],
    "東神楽町": [22.5, 23.0, 23.4, 23.8, 24.5, 25.2, 25.8, 26.7, 27.4, 27.7, 28.4, 28.7],
}
end = table(ws, 4, HD, [[k] + v + [None] * 3 for k, v in RATE.items()], numfmt="0.0")
for r_ in range(5, 9):
    for c in range(14, 17):
        ws.cell(row=r_, column=c).fill = PatternFill("solid", fgColor=IN_Y)
note(ws, end + 1, "注）東川町の高齢化率は令和2年の32.8％をピークに緩やかな下降傾向、美瑛町・東神楽町は上昇傾向で推移している"
                  "（第9期計画 第2章第1節1）。令和6～8年（淡黄色欄）は実績受領後に入力する。")
cats = Reference(ws, min_col=2, max_col=16, min_row=4)
data = Reference(ws, min_col=1, max_col=16, min_row=5, max_row=8)
mono_line(ws, "高齢化率の推移", "高齢化率（％）", cats, data, "A13", width=26, height=12, min_=20, max_=42)

# ============================================================ 03 世帯
ws = sheet("03_高齢者世帯", "図3　高齢者を含む世帯の状況",
           "資料：国勢調査（平成22年・平成27年・令和2年）（第9期計画 第2章第1節2）／単位：世帯・％",
           [22, 12, 12, 12, 12, 12, 12, 12, 12, 12])
r = 4
ws.cell(row=r, column=1, value="（1）世帯数").font = Font(name=FONT, size=10, bold=True)
r += 1
SETAI = [
    ["一般世帯　広域連合", 10536, 11050, 11426], ["　東川町", 2965, 3132, 3391],
    ["　美瑛町", 4289, 4274, 4205], ["　東神楽町", 3282, 3644, 3830],
    ["高齢者を含む世帯　広域連合", 4885, 5318, 5622], ["　東川町", 1316, 1506, 1620],
    ["　美瑛町", 2330, 2341, 2353], ["　東神楽町", 1239, 1471, 1649],
    ["高齢者夫婦世帯　広域連合", 1461, 1697, 1950], ["　東川町", 387, 504, 596],
    ["　美瑛町", 667, 739, 743], ["　東神楽町", 407, 454, 611],
    ["高齢者単身世帯　広域連合", 1076, 1331, 1638], ["　東川町", 295, 363, 490],
    ["　美瑛町", 531, 653, 714], ["　東神楽町", 250, 315, 434],
]
end = table(ws, r, ["区分", "平成22年", "平成27年", "令和2年"], SETAI, numfmt="#,##0")
r = end + 2
ws.cell(row=r, column=1, value="（2）高齢者を含む世帯の構成割合").font = Font(name=FONT, size=10, bold=True)
r += 1
KOSEI = [
    ["広域連合（平成22年）", 29.9, 22.0, 48.1], ["広域連合（平成27年）", 31.9, 25.0, 43.1],
    ["広域連合（令和2年）", 34.7, 29.1, 36.2], ["東川町（令和2年）", 36.8, 30.2, 33.0],
    ["美瑛町（令和2年）", 31.6, 30.3, 38.1], ["東神楽町（令和2年）", 37.1, 26.3, 36.6],
]
end2 = table(ws, r, ["区分", "高齢者夫婦世帯", "高齢者単身世帯", "その他の高齢者を含む世帯"], KOSEI, numfmt="0.0")
note(ws, end2 + 1, "注）構成割合は「高齢者を含む世帯」に占める割合。令和2年では高齢者単身世帯の割合は美瑛町が30.3％と最も高く、"
                   "高齢者夫婦世帯の割合は東神楽町が37.1％と最も高い（第9期計画 第2章第1節2）。次回は令和7年国勢調査の公表後に更新する。")
cats = Reference(ws, min_col=1, min_row=r + 1, max_row=r + 6)
data = Reference(ws, min_col=2, max_col=4, min_row=r, max_row=r + 6)
mono_hbar(ws, "高齢者を含む世帯の構成割合", cats, data, f"F{r}", width=20, height=11)
SETAI_HEAD = 5           # 世帯数テーブルの見出し行
FUFU_ROW = SETAI_HEAD + 9    # 高齢者夫婦世帯 広域連合
TANSHIN_ROW = SETAI_HEAD + 13  # 高齢者単身世帯 広域連合
assert ws.cell(row=FUFU_ROW, column=1).value.startswith("高齢者夫婦世帯")
assert ws.cell(row=TANSHIN_ROW, column=1).value.startswith("高齢者単身世帯")
cats2 = Reference(ws, min_col=2, max_col=4, min_row=SETAI_HEAD)
data2 = Reference(ws, min_col=1, max_col=4, min_row=FUFU_ROW, max_row=FUFU_ROW)
d3 = Reference(ws, min_col=1, max_col=4, min_row=TANSHIN_ROW, max_row=TANSHIN_ROW)
ch = BarChart()
ch.type = "col"
ch.title = "高齢者夫婦世帯・高齢者単身世帯の推移（広域連合）"
ch.y_axis.title = "世帯数（世帯）"
ch.add_data(data2, titles_from_data=True, from_rows=True)
ch.add_data(d3, titles_from_data=True, from_rows=True)
ch.set_categories(cats2)
ch.gapWidth = 80
ch.width, ch.height = 14, 9
for i, s in enumerate(ch.series):
    gp = GraphicalProperties(solidFill=GRAY[i * 2])
    gp.line = LineProperties(solidFill="000000", w=6350)
    s.graphicalProperties = gp
_axis_mono(ch)
ws.add_chart(ch, f"F{r + 24}")

# ============================================================ 04 認定者数
ws = sheet("04_認定者数推移", "図4　要介護（要支援）認定者数の推移（大雪地区広域連合）",
           "資料：介護保険事業状況報告 各年9月分（第9期計画 第2章第1節3）／単位：人",
           [16] + [10] * 15)
NINTEI = {
    "要支援1": [169, 221, 264, 325, 297, 274, 289, 275, 271, 254, 239, 248],
    "要支援2": [243, 270, 282, 288, 263, 268, 256, 295, 302, 300, 313, 285],
    "要介護1": [336, 354, 396, 410, 415, 420, 447, 438, 424, 482, 470, 492],
    "要介護2": [283, 274, 301, 315, 306, 323, 320, 315, 323, 320, 350, 329],
    "要介護3": [222, 223, 196, 198, 220, 250, 228, 235, 237, 224, 229, 247],
    "要介護4": [175, 191, 190, 198, 211, 195, 182, 209, 210, 232, 213, 222],
    "要介護5": [181, 198, 188, 170, 177, 170, 179, 161, 154, 158, 158, 153],
}
rows = [[k] + v + [None] * 3 for k, v in NINTEI.items()] + [["合計"] + [None] * 15]
end = table(ws, 4, HD, rows, numfmt="#,##0")
TOT_ROW = end
for c in range(2, 17):
    col = get_column_letter(c)
    ws.cell(row=TOT_ROW, column=c).value = f"=SUM({col}5:{col}11)"
    ws.cell(row=TOT_ROW, column=c).number_format = "#,##0"
    ws.cell(row=TOT_ROW, column=c).font = Font(name=FONT, size=9, bold=True)
for r_ in range(5, 12):
    for c in range(14, 17):
        ws.cell(row=r_, column=c).fill = PatternFill("solid", fgColor=IN_Y)
note(ws, end + 1, "注）第1号・第2号被保険者の合計。令和5年の合計1,976人は第9期計画の記載と一致する。"
                  "要支援1～要介護4は増加傾向、要介護5は減少傾向で推移している（第9期計画 第2章第1節3）。"
                  "令和6～8年（淡黄色欄）は各年9月分の実績受領後に入力する。")
cats = Reference(ws, min_col=2, max_col=16, min_row=4)
data = Reference(ws, min_col=1, max_col=16, min_row=5, max_row=11)
mono_bar(ws, "認定者数の推移（要介護度別）", "認定者数（人）", cats, data, "A16", stacked=True, width=26, height=12)

# ============================================================ 05 認定者割合
ws = sheet("05_認定者割合", "図5　認定者割合の比較（令和5年9月）",
           "資料：介護保険事業状況報告 令和5年9月分（第9期計画 第2章第1節3）／単位：％",
           [16, 11, 11, 11, 11, 11, 11, 11])
WARIAI = [
    ["全国", 14.2, 13.9, 20.7, 16.7, 13.2, 12.7, 8.5],
    ["北海道", 18.4, 14.6, 23.1, 15.7, 10.6, 10.4, 7.2],
    ["広域連合", 12.6, 14.4, 24.9, 16.6, 12.5, 11.2, 7.7],
    ["東川町", 9.9, 14.8, 23.4, 18.1, 11.7, 13.1, 9.1],
    ["美瑛町", 14.2, 14.7, 23.6, 16.3, 14.3, 10.4, 6.4],
    ["東神楽町", 12.5, 13.6, 28.7, 15.7, 10.2, 10.7, 8.6],
]
end = table(ws, 4, ["区分", "要支援1", "要支援2", "要介護1", "要介護2", "要介護3", "要介護4", "要介護5"],
            WARIAI, numfmt="0.0")
note(ws, end + 1, "注）東川町は「要支援1」の割合が9.9％と全国・北海道・他町より低く、美瑛町は「要支援1」14.2％と「要介護3」14.3％が他町より高い。"
                  "東神楽町は「要介護1」の割合が28.7％と全国・北海道・他町より高い（第9期計画 第2章第1節3）。")
cats = Reference(ws, min_col=1, min_row=5, max_row=10)
data = Reference(ws, min_col=2, max_col=8, min_row=4, max_row=10)
mono_hbar(ws, "認定者割合の比較（令和5年9月）", cats, data, "A14", width=22, height=11)

# ============================================================ 06 出現率
ws = sheet("06_出現率推移", "図6　要支援・要介護認定者の出現率",
           "資料：介護保険事業状況報告 各年9月分（第9期計画 第2章第1節3）／出現率＝第1号認定者数÷第1号被保険者数・単位：％",
           [18] + [10] * 15)
r = 4
ws.cell(row=r, column=1, value="（1）出現率の推移（全国・北海道との比較）").font = Font(name=FONT, size=10, bold=True)
r += 1
SHUTSU = {
    "広域連合": [19.3, 20.2, 20.6, 21.2, 20.7, 20.7, 20.5, 20.7, 20.5, 21.0, 21.1, 21.2],
    "北海道": [18.5, 18.9, 19.2, 19.4, 19.5, 19.5, 19.7, 20.0, 20.2, 20.4, 20.6, 20.8],
    "全国": [17.5, 17.8, 17.9, 18.0, 18.0, 18.1, 18.3, 18.5, 18.6, 18.8, 19.1, 19.3],
}
end = table(ws, r, HD, [[k] + v + [None] * 3 for k, v in SHUTSU.items()], numfmt="0.0")
for r_ in range(r + 1, r + 4):
    for c in range(14, 17):
        ws.cell(row=r_, column=c).fill = PatternFill("solid", fgColor=IN_Y)
cats = Reference(ws, min_col=2, max_col=16, min_row=r)
data = Reference(ws, min_col=1, max_col=16, min_row=r + 1, max_row=r + 3)
mono_line(ws, "出現率の推移", "出現率（％）", cats, data, "A11", width=26, height=11, min_=15, max_=23)

r = end + 2
ws.cell(row=r, column=1, value="（2）町別出現率の推移").font = Font(name=FONT, size=10, bold=True)
r += 1
SHUTSU_T = {
    "東川町": [18.7, 19.8, 19.2, 19.1, 17.5, 17.8, 18.4, 18.6, 19.4, 19.6, 19.9, 20.5],
    "美瑛町": [19.3, 20.6, 21.8, 23.0, 22.9, 23.3, 22.8, 23.2, 22.7, 23.4, 23.5, 24.0],
    "東神楽町": [19.8, 19.9, 20.3, 20.8, 20.7, 19.9, 19.3, 19.3, 18.5, 19.1, 18.9, 18.2],
}
end2 = table(ws, r, HD, [[k] + v + [None] * 3 for k, v in SHUTSU_T.items()], numfmt="0.0")
for r_ in range(r + 1, r + 4):
    for c in range(14, 17):
        ws.cell(row=r_, column=c).fill = PatternFill("solid", fgColor=IN_Y)
cats = Reference(ws, min_col=2, max_col=16, min_row=r)
data = Reference(ws, min_col=1, max_col=16, min_row=r + 1, max_row=r + 3)
mono_line(ws, "町別出現率の推移", "出現率（％）", cats, data, f"A{end2 + 3}", width=26, height=11, min_=15, max_=26)

r = end2 + 26
ws.cell(row=r, column=1, value="（3）出現率の比較（令和5年 上川総合振興局管内）").font = Font(name=FONT, size=10, bold=True)
r += 1
KANNAI = [
    ("全国", 19.3), ("北海道", 20.8), ("上川総合振興局管内", 21.2), ("当麻町", 23.8), ("旭川市", 21.8),
    ("中川町", 21.5), ("比布町", 21.4), ("大雪地区広域連合", 21.2), ("和寒町", 21.2), ("上川町", 20.6),
    ("名寄市", 20.5), ("幌加内町", 20.2), ("愛別町", 20.2), ("富良野市", 19.8), ("中富良野町", 19.4),
    ("南富良野町", 19.3), ("美深町", 19.2), ("鷹栖町", 18.7), ("士別市", 18.7), ("剣淵町", 17.7),
    ("下川町", 17.3), ("占冠村", 17.2), ("上富良野町", 16.1), ("音威子府村", 10.8),
]
end3 = table(ws, r, ["区分", "出現率"], [list(x) for x in KANNAI], numfmt="0.0")
note(ws, end3 + 1, "注）当広域連合の出現率は全国・北海道を上回って推移し、平成25年以降は20％を超えている。"
                   "上川総合振興局管内の保険者では高い方から5番目（第9期計画 第2章第1節3）。")
cats = Reference(ws, min_col=1, min_row=r + 1, max_row=end3)
data = Reference(ws, min_col=2, min_row=r, max_row=end3)
ch = BarChart()
ch.type = "col"
ch.title = "出現率の比較（令和5年 上川総合振興局管内）"
ch.y_axis.title = "出現率（％）"
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ch.gapWidth = 40
ch.width, ch.height = 26, 11
gp = GraphicalProperties(solidFill="A6A6A6")
gp.line = LineProperties(solidFill="000000", w=6350)
ch.series[0].graphicalProperties = gp
_axis_mono(ch)
ch.legend = None
ws.add_chart(ch, f"D{r}")

# ============================================================ 07 給付費
ws = sheet("07_給付費推移", "図7　介護給付費等の推移",
           "資料：第9期計画 第2章第2節（見える化システム 令和6年1月18日参照）／素案第9稿 表39（自然体推計）",
           [26, 14, 14, 14, 14, 14, 14, 14, 14])
r = 4
ws.cell(row=r, column=1, value="（1）サービス区分別給付費（実績・単位：千円）").font = Font(name=FONT, size=10, bold=True)
r += 1
KYUFU = [
    ["居宅サービス", 836496, 872021, 896194, None, None, None],
    ["地域密着型サービス", 765808, 718778, 743018, None, None, None],
    ["施設サービス", 973845, 924985, 941503, None, None, None],
    ["居宅介護支援", 111147, 115959, 115772, None, None, None],
    ["介護予防サービス", 68592, 64733, 66451, None, None, None],
    ["地域密着型介護予防サービス", 25160, 22391, 19802, None, None, None],
    ["介護予防支援", 12766, 12144, 11653, None, None, None],
    ["合計", None, None, None, None, None, None],
]
end = table(ws, r, ["区分", "令和3年度", "令和4年度", "令和5年度", "令和6年度", "令和7年度", "令和8年度"],
            KYUFU, numfmt="#,##0")
for c in range(2, 8):
    col = get_column_letter(c)
    ws.cell(row=end, column=c).value = f"=SUM({col}{r+1}:{col}{end-1})"
    ws.cell(row=end, column=c).number_format = "#,##0"
    ws.cell(row=end, column=c).font = Font(name=FONT, size=9, bold=True)
for r_ in range(r + 1, end):
    for c in range(5, 8):
        ws.cell(row=r_, column=c).fill = PatternFill("solid", fgColor=IN_Y)
cats = Reference(ws, min_col=2, max_col=7, min_row=r)
data = Reference(ws, min_col=1, max_col=7, min_row=r + 1, max_row=end - 1)
mono_bar(ws, "サービス区分別給付費の推移", "給付費（千円）", cats, data, "J4", stacked=True, width=20, height=11)

r2 = end + 2
ws.cell(row=r2, column=1, value="（2）総給付費（在宅・居住系・施設別／実績・単位：千円）").font = Font(name=FONT, size=10, bold=True)
r2 += 1
SOU = [
    ["在宅サービス", 1191892, 1218034, 1250603, None, None, None],
    ["居住系サービス", 436392, 399288, 414922, None, None, None],
    ["施設サービス", 1165530, 1113689, 1128866, None, None, None],
    ["総給付費", None, None, None, None, None, None],
]
end2 = table(ws, r2, ["区分", "令和3年度", "令和4年度", "令和5年度", "令和6年度", "令和7年度", "令和8年度"],
             SOU, numfmt="#,##0")
for c in range(2, 8):
    col = get_column_letter(c)
    ws.cell(row=end2, column=c).value = f"=SUM({col}{r2+1}:{col}{end2-1})"
    ws.cell(row=end2, column=c).number_format = "#,##0"
    ws.cell(row=end2, column=c).font = Font(name=FONT, size=9, bold=True)
for r_ in range(r2 + 1, end2):
    for c in range(5, 8):
        ws.cell(row=r_, column=c).fill = PatternFill("solid", fgColor=IN_Y)

r3 = end2 + 2
ws.cell(row=r3, column=1, value="（3）地域支援事業費（実績・単位：円）").font = Font(name=FONT, size=10, bold=True)
r3 += 1
CHIIKI = [
    ["介護予防・日常生活支援総合事業費", 98828508, 100020350, 106455000, None, None, None],
    ["包括的支援事業（センター運営）及び任意事業費", 67884572, 71189693, 65532000, None, None, None],
    ["包括的支援事業（社会保障充実分）", 12746280, 15365016, 18387000, None, None, None],
    ["地域支援事業費 合計", None, None, None, None, None, None],
]
end3 = table(ws, r3, ["区分", "令和3年度", "令和4年度", "令和5年度", "令和6年度", "令和7年度", "令和8年度"],
             CHIIKI, numfmt="#,##0")
for c in range(2, 8):
    col = get_column_letter(c)
    ws.cell(row=end3, column=c).value = f"=SUM({col}{r3+1}:{col}{end3-1})"
    ws.cell(row=end3, column=c).number_format = "#,##0"
    ws.cell(row=end3, column=c).font = Font(name=FONT, size=9, bold=True)
for r_ in range(r3 + 1, end3):
    for c in range(5, 8):
        ws.cell(row=r_, column=c).fill = PatternFill("solid", fgColor=IN_Y)

r4 = end3 + 2
ws.cell(row=r4, column=1, value="（4）給付費の中長期見通し（自然体推計・参考／単位：百万円）").font = Font(name=FONT, size=10, bold=True)
r4 += 1
SHIZEN = [
    ["居宅サービス", 1537.4, 1559.6, 1589.9, 1652.4, 1738.6],
    ["居住系サービス", 380.6, 391.9, 391.9, 423.7, 441.5],
    ["施設サービス", 1138.7, 1138.7, 1138.7, 1265.5, 1312.4],
    ["合計", None, None, None, None, None],
]
end4 = table(ws, r4, ["区分", "令和9年度", "令和10年度", "令和11年度", "令和17年度", "令和22年度"],
             SHIZEN, numfmt="#,##0.0")
for c in range(2, 7):
    col = get_column_letter(c)
    ws.cell(row=end4, column=c).value = f"=SUM({col}{r4+1}:{col}{end4-1})"
    ws.cell(row=end4, column=c).number_format = "#,##0.0"
    ws.cell(row=end4, column=c).font = Font(name=FONT, size=9, bold=True)
cats = Reference(ws, min_col=2, max_col=6, min_row=r4)
data = Reference(ws, min_col=1, max_col=6, min_row=r4 + 1, max_row=end4 - 1)
mono_bar(ws, "給付費の中長期見通し（自然体推計）", "給付費（百万円）", cats, data, f"J{r2}", stacked=True, width=20, height=11)

note(ws, end4 + 2,
     "注1）（1）～（3）は第9期計画 第2章第2節の掲載値（見える化システム 令和6年1月18日参照）。令和6～8年度（淡黄色欄）は実績受領後に入力する。"
     "注2）（2）の総給付費は（1）の介護給付費と介護予防給付費の合計に相当し、令和5年度は2,794,391千円（在宅・居住系・施設の3区分で再集計した値）。"
     "素案第9稿に記載の保険給付費 令和5年度決算2,940.4百万円とは定義が異なるため、同一系列に接続していない（00_凡例・出典 4-4）。"
     "注3）（1）の合計（令和5年度2,794,393千円）と（2）の総給付費（同2,794,391千円）には2千円の差がある。"
     "これは第9期計画の介護予防給付費の合計欄（令和4年度99,267千円・令和5年度97,905千円）に内訳との±1千円の端数差があるためで、"
     "本図表集では内訳を原典どおり掲載し合計を数式で算出している。計画掲載時はいずれかに統一する。"
     "注4）（4）は素案第9稿 表39の自然体推計値で、報酬改定・供給制約・政策効果は未反映の参考値。採用値ではない。")

# ============================================================ 08 ニーズ調査データ
ws = sheet("08_ニーズ調査データ", "図8　介護予防・日常生活圏域ニーズ調査 結果概要",
           "資料：第9期計画 第2章第3節／令和4年11月7日～28日実施・65歳以上7,239人対象・回収4,626票（63.9％）／単位：％（⑨のみ点）",
           [6, 26, 12, 11, 11, 11, 11, 11, 11, 11])
NEEDS = [
    ("①", "フレイルあり（基本チェックリスト8項目以上）", "％",
     [("広域連合", 4554, 18.5, 9.9, 13.1, 20.5, 26.0, 36.7),
      ("東川町", 1367, 18.6, 9.4, 13.3, 19.3, 27.2, 35.2),
      ("美瑛町", 1676, 18.9, 10.3, 12.9, 23.2, 21.9, 37.3),
      ("東神楽町", 1511, 18.0, 10.0, 13.0, 18.8, 30.6, 37.6)]),
    ("②", "運動機能低下者", "％",
     [("広域連合", 4485, 9.7, 4.1, 6.9, 10.6, 13.2, 23.3),
      ("東川町", 1339, 9.6, 3.9, 5.7, 11.0, 11.7, 25.2),
      ("美瑛町", 1655, 10.5, 4.8, 7.7, 11.4, 12.1, 24.1),
      ("東神楽町", 1491, 9.1, 3.7, 7.3, 9.2, 16.4, 20.2)]),
    ("③", "1年間転倒あり", "％",
     [("広域連合", 4477, 32.9, 27.0, 29.7, 33.6, 37.0, 47.1),
      ("東川町", 1339, 32.8, 28.0, 25.7, 34.5, 41.1, 44.7),
      ("美瑛町", 1650, 35.0, 27.9, 35.4, 35.2, 33.6, 50.5),
      ("東神楽町", 1488, 30.6, 25.7, 27.5, 30.7, 37.7, 45.0)]),
    ("④", "物忘れが多い者", "％",
     [("広域連合", 4440, 43.4, 35.4, 41.4, 46.1, 49.4, 52.3),
      ("東川町", 1322, 44.5, 33.1, 42.1, 46.5, 51.8, 56.5),
      ("美瑛町", 1633, 44.8, 36.6, 43.3, 49.7, 48.8, 49.2),
      ("東神楽町", 1485, 40.8, 35.7, 38.7, 41.4, 47.6, 52.3)]),
    ("⑤", "閉じこもり者", "％",
     [("広域連合", 4472, 6.7, 3.2, 4.3, 6.5, 10.5, 15.2),
      ("東川町", 1333, 7.6, 3.1, 5.3, 7.9, 10.4, 17.4),
      ("美瑛町", 1652, 6.5, 3.5, 3.2, 5.6, 11.5, 13.9),
      ("東神楽町", 1487, 6.1, 3.0, 4.7, 6.1, 9.4, 14.7)]),
    ("⑥", "うつ", "％",
     [("広域連合", 4444, 28.9, 20.7, 25.4, 29.7, 37.8, 42.1),
      ("東川町", 1327, 28.9, 22.8, 24.6, 30.0, 36.2, 38.5),
      ("美瑛町", 1646, 30.8, 20.2, 26.1, 33.0, 38.5, 46.5),
      ("東神楽町", 1471, 26.9, 19.8, 25.5, 25.6, 38.5, 39.4)]),
    ("⑦", "口腔機能低下者", "％",
     [("広域連合", 4463, 23.4, 18.0, 21.5, 25.6, 26.6, 32.0),
      ("東川町", 1330, 23.8, 20.6, 20.9, 25.5, 26.2, 30.1),
      ("美瑛町", 1645, 25.3, 18.1, 24.2, 26.7, 28.7, 35.2),
      ("東神楽町", 1488, 21.0, 16.4, 19.4, 24.3, 23.9, 29.5)]),
    ("⑧", "低栄養の傾向", "％",
     [("広域連合", 4351, 6.0, 5.1, 5.9, 5.6, 7.3, 7.2),
      ("東川町", 1298, 6.0, 3.9, 8.1, 3.9, 7.8, 6.1),
      ("美瑛町", 1595, 6.0, 5.7, 4.8, 6.4, 7.7, 6.4),
      ("東神楽町", 1458, 6.0, 5.5, 5.0, 6.6, 6.3, 9.6)]),
    ("⑨", "要支援・要介護リスク点数の平均点", "点",
     [("広域連合", 4560, 15.8, 5.1, 12.1, 18.8, 24.6, 30.0),
      ("東川町", 1367, 16.4, 5.3, 12.2, 18.6, 24.8, 30.1),
      ("美瑛町", 1679, 16.3, 4.9, 12.2, 19.2, 24.5, 30.0),
      ("東神楽町", 1514, 14.7, 5.2, 11.9, 18.6, 24.5, 30.1)]),
    ("⑩", "認知機能低下者（基本チェックリスト）", "％",
     [("広域連合", 4496, 34.5, 28.5, 31.8, 33.2, 42.6, 45.6),
      ("東川町", 1346, 33.7, 24.9, 30.6, 34.0, 43.1, 43.1),
      ("美瑛町", 1656, 34.9, 27.4, 33.8, 31.9, 43.8, 43.6),
      ("東神楽町", 1494, 34.9, 31.7, 30.9, 34.0, 40.5, 51.2)]),
    ("⑪", "IADL（自立度）低下者（1項目以上）", "％",
     [("広域連合", 4500, 9.8, 4.6, 7.1, 9.3, 13.8, 23.5),
      ("東川町", 1346, 10.9, 6.2, 7.5, 9.7, 16.4, 22.9),
      ("美瑛町", 1658, 9.4, 4.5, 5.9, 9.4, 12.0, 23.5),
      ("東神楽町", 1496, 9.2, 3.7, 8.2, 8.8, 13.5, 24.0)]),
    ("⑫", "通いの場参加者（月1回以上）", "％",
     [("広域連合", 4324, 7.3, 3.6, 5.0, 8.0, 13.5, 12.4),
      ("東川町", 1277, 8.1, 3.9, 4.8, 8.9, 17.0, 11.6),
      ("美瑛町", 1593, 7.3, 2.7, 4.3, 8.1, 13.0, 14.2),
      ("東神楽町", 1454, 6.7, 4.3, 5.8, 6.9, 10.7, 10.9)]),
    ("⑬", "ボランティア参加者（月1回以上）", "％",
     [("広域連合", 4301, 9.0, 7.0, 10.8, 7.6, 10.1, 9.7),
      ("東川町", 1271, 10.8, 8.5, 11.1, 8.7, 15.9, 11.2),
      ("美瑛町", 1581, 9.2, 7.2, 12.1, 8.2, 7.6, 10.4),
      ("東神楽町", 1449, 7.2, 5.9, 9.1, 5.9, 7.8, 7.1)]),
    ("⑭", "社会的ネットワーク：友人知人と会う頻度が高い（月1回以上）", "％",
     [("広域連合", 4433, 64.3, 59.0, 66.2, 68.8, 66.4, 59.0),
      ("東川町", 1330, 66.8, 60.4, 67.0, 70.1, 70.9, 64.7),
      ("美瑛町", 1627, 66.6, 61.2, 68.1, 72.1, 70.0, 58.6),
      ("東神楽町", 1476, 59.6, 56.1, 63.5, 63.5, 56.6, 53.5)]),
]
HEAD8 = ["No.", "指標", "地域", "サンプル数", "全体", "65～69歳", "70～74歳", "75～79歳", "80～84歳", "85歳以上"]
rows8 = []
index = {}
r = 5
for no, name, unit, blocks in NEEDS:
    index[no] = (name, unit, r)
    for j, b in enumerate(blocks):
        rows8.append([no if j == 0 else "", name if j == 0 else "", b[0], b[1], b[2], b[3], b[4], b[5], b[6], b[7]])
    r += 4
end = table(ws, 4, HEAD8, rows8, numfmt="0.0")
for rr in range(5, end + 1):
    ws.cell(row=rr, column=4).number_format = "#,##0"
note(ws, end + 1,
     "注1）第9期計画 第2章第3節の掲載値。町の対応はサンプル数と回収票数（東川町1,367票・美瑛町1,679票・東神楽町1,514票）の照合により確定した。"
     "注2）⑨は第9期計画に「15.8％」と記載されているが、実際は要支援・要介護リスク評価尺度による平均点（72,099点÷4,560人＝15.8点）であり、"
     "本表では単位を「点」としている（00_凡例・出典 4-2）。"
     "注3）第10期ニーズ調査は実施済みで集計未受領。設問・判定ロジックの一致を確認したうえで全面差替えとなる。")

# ============================================================ 09 ニーズ調査グラフ
wsg = sheet("09_ニーズ調査グラフ", "図8　介護予防・日常生活圏域ニーズ調査 結果概要（グラフ）",
            "08_ニーズ調査データ を参照。年齢階級別（広域連合）と町別（全体）を各指標について作図",
            [14] * 20)
AGE_CATS = Reference(ws, min_col=5, max_col=10, min_row=4)
anchor_r = 4
for k, (no, name, unit, blocks) in enumerate(NEEDS):
    base = 5 + k * 4
    ytitle = "割合（％）" if unit == "％" else "平均点（点）"
    # 年齢階級別（広域連合）
    ch = BarChart()
    ch.type = "col"
    ch.title = f"{no} {name}（年齢階級別・広域連合）"
    ch.y_axis.title = ytitle
    ch.append(_named_series(ws, base, base))
    ch.set_categories(Reference(ws, min_col=5, max_col=10, min_row=4))
    ch.gapWidth = 60
    ch.width, ch.height = 13, 8
    gp = GraphicalProperties(solidFill="737373")
    gp.line = LineProperties(solidFill="000000", w=6350)
    ch.series[0].graphicalProperties = gp
    _axis_mono(ch)
    ch.legend = None
    wsg.add_chart(ch, f"A{anchor_r}")
    # 町別（全体・年齢階級別）
    ch2 = BarChart()
    ch2.type = "col"
    ch2.title = f"{no} {name}（町別）"
    ch2.y_axis.title = ytitle
    for rr in range(base, base + 4):
        ch2.append(_named_series(ws, rr, rr))
    ch2.set_categories(Reference(ws, min_col=5, max_col=10, min_row=4))
    ch2.gapWidth = 60
    ch2.width, ch2.height = 16, 8
    for i, s in enumerate(ch2.series):
        g = GraphicalProperties(solidFill=GRAY[i % 5])
        g.line = LineProperties(solidFill="000000", w=6350)
        s.graphicalProperties = g
    _axis_mono(ch2)
    wsg.add_chart(ch2, f"J{anchor_r}")
    anchor_r += 17

del wb["Sheet"]
wb.save("/home/user/repository/output/第10期計画_図表集_白黒.xlsx")
print("saved:", len(wb.sheetnames), "sheets")
