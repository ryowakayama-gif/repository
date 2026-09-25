"""小野町 見える化システム 地域支援事業の見込み量推計 ― 画面の形式に合わせた入力値。

令和8年9月25日のご指示
  「地域支援事業まで来ましたが、添付ファイルのような形式のため
    再度確認をお願いします。」

**画面キャプチャ（推計名「将来推計0925-2」）で分かったこと**
  画面名：訪問型・通所型サービス利用者数／事業費の実績値及び計画値の入力
  ・ラジオボタンで「年齢階級別に登録する」／「サービスごとの総数で登録する」
    を選ぶ。**「サービスごとの総数で登録する」が選ばれている（これでよい）。**
  ・右上のボタンで「利用者数を入力する」⇄「事業費を入力する」を切り替える。
    (1) 利用者数 は単位「人／月」、(2) 事業費 は単位「円」。
  ・列は R6・R7・R8（青＝実績値）と R9・R10・R11（オレンジ＝計画値）。
    **実績値の3年度も手入力である。**現在はすべて0。
  ・行は4つ。訪問介護相当サービス利用者／訪問型サービスA利用者／
    通所介護相当サービス利用者／通所型サービスA利用者。

**あわせて気づいた点**
  画面右上が「現在時点」になっている。「将来推計1」「将来推計0925」では
  「公表時点」であった。**版が固定され再現できるため「公表時点」を推奨する。**

出力
  11_見える化出力依頼/小野町_見える化_地域支援事業の画面入力_YYYYMMDD.docx
  11_見える化出力依頼/小野町_見える化_地域支援事業の画面入力_YYYYMMDD.xlsx
"""

import pathlib
import warnings

warnings.filterwarnings("ignore")

import openpyxl
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from build_ono_kaisu import chiiki_plan, chiiki_ryo, sogo_jisseki

ROOT = pathlib.Path(__file__).parent
OUT = ROOT / "小野町_引継ぎ_整理済" / "11_見える化出力依頼"
ASOF = "20260925"
ASOF_JP = "令和8年9月25日"
JP_MIN = "游明朝"
JP_GO = "游ゴシック"

HEAD = PatternFill("solid", fgColor="1F3864")
IN = PatternFill("solid", fgColor="FFF2CC")     # 当方で用意できる
MACHI = PatternFill("solid", fgColor="FCE4E4")  # 町の資料が要る
OK = PatternFill("solid", fgColor="E2EFDA")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

CHIIKI_NEN = 63916236    # 地域支援事業費の年額（年報 様式4 の令和6年度決算）
SOGO_NEN = 34581047      # うち介護予防・日常生活支援総合事業費
HOKATSU_NEN = 29335189   # 包括的支援事業（センター運営）及び任意事業費

# 画面ごとの注意
GAMEN = [
    ("①", "訪問型・通所型サービス利用者数／事業費",
     "**「サービスごとの総数で登録する」のまま。**"
     "右上のボタンで「利用者数」と「事業費」を切り替えて、両方入れます。"
     "**R6〜R8（青）も手入力です**"),
    ("②", "その他の訪問型・通所型サービス（B・C・D ほか）",
     "**国保連月報では令和6年度以降0件です。**"
     "町で実施していなければ0のままで構いません。"
     "**実施している場合は決算額が要ります**"),
    ("③", "介護予防ケアマネジメント",
     "当方で用意できます。第2章の表のとおり"),
    ("④", "審査支払手数料・高額介護予防サービス相当事業等",
     "**2つ合わせて令和6年度530,663円**（年報 様式4 の決算額と国保連月報の差）。"
     "内訳が分かるまでは審査支払手数料に全額を寄せてください"),
    ("⑤", "一般介護予防事業（5事業）",
     "**5つ合わせて令和6年度5,304,726円**（年報 様式4）。"
     "内訳が分かるまでは介護予防普及啓発事業に全額を寄せてください"),
    ("⑥", "包括的支援事業（地域包括支援センターの運営）及び任意事業",
     "**包括的支援事業（社会保障充実分）6事業と合わせて"
     "令和6年度29,335,189円**（年報 様式4）。"
     "内訳が分かるまでは地域包括支援センターの運営に全額を寄せてください"),
    ("⑦", "包括的支援事業（社会保障充実分）",
     "**⑥に寄せるため0で構いません。**"
     "内訳が届いたら割り付け直します"),
    ("⑧", "特定地域居宅サービス等事業・介護情報利活用事業",
     "**0です。**小野町では実施していません"),
]

CHUI = [
    ("1", "**「サービスごとの総数で登録する」のままにしてください。**",
     "「年齢階級別に登録する」を選ぶと、前期高齢者・後期高齢者別の入力になり、"
     "当方が用意できる形になりません。**国保連月報は年齢階級別に分かれていません**"),
    ("2", "**R6〜R8（青い列）も手入力です。**",
     "実績値も自動では入りません。"
     "**令和8年度は国保連月報が3か月分しかありません。**"
     "町の当初予算があればそちらを、なければ3か月の年換算を入れてください"),
    ("3", "**内訳が分からない欄は、合計額を1つの欄に寄せてください。**",
     "**保険料に効くのは「総合事業費」と"
     "「包括的支援事業・任意事業費」の2つの合計だけです。**"
     "総合事業費には調整交付金がかかり、包括的支援事業・任意事業費には"
     "かからないため、この2つの境を越えなければ保険料は変わりません"),
    ("4", "**画面右上が「現在時点」になっています。**",
     "「将来推計1」「将来推計0925」では「公表時点」でした。"
     "**「公表時点」のほうが版が固定され、あとから同じ結果を再現できます。**"
     "令和6・7年度は完結年度なので値は変わらないはずですが、"
     "念のため「実績及び推計方法の設定」でご確認ください"),
    ("5", "**画面上部の保険料額（2,968円）はまだ意味を持ちません。**",
     "所得段階別第1号被保険者数も加算4項目も未入力のためです。"
     "**すべて入れ終えてから見てください**"),
]


# ---------------------------------------------------------------- 入力値

def gamen1():
    """画面①の入力値。(区分, 行名, 単位, R6, R7, R8, R9, R10, R11)。"""
    sg = sogo_jisseki()
    cp = {ev: mik for _k, ev, _j, mik, _n in chiiki_plan()}
    ry = {r[1]: r for r in chiiki_ryo()}
    r6, r7, r8 = (sg[y]["事業"] for y in ("令和6年度", "令和7年度", "令和8年度"))
    m8 = sg["令和8年度"]["月数"]
    out = []
    for ev, lab in (("訪問介護相当サービス", "訪問介護相当サービス利用者"),
                    ("通所介護相当サービス", "通所介護相当サービス利用者")):
        rr = ry[ev]
        out.append(("利用者数", lab, "人／月",
                    r6[ev][0] / 12, r7[ev][0] / 12, r8[ev][0] / m8,
                    rr[5], rr[6], rr[7]))
    for ev, lab in (("訪問型サービスA", "訪問型サービスA利用者"),
                    ("通所型サービスA", "通所型サービスA利用者")):
        out.append(("利用者数", lab, "人／月", 0, 0, 0, 0, 0, 0))
    for ev in ("訪問介護相当サービス", "通所介護相当サービス"):
        out.append(("事業費", ev, "円",
                    r6[ev][1], r7[ev][1], r8[ev][1] / m8 * 12,
                    cp[ev], cp[ev], cp[ev]))
    for ev in ("訪問型サービスA", "通所型サービスA"):
        out.append(("事業費", ev, "円", 0, 0, 0, 0, 0, 0))
    # 画面の並びにそろえる（訪問介護相当→訪問A→通所介護相当→通所A）
    order = ["訪問介護相当サービス利用者", "訪問型サービスA利用者",
             "通所介護相当サービス利用者", "通所型サービスA利用者",
             "訪問介護相当サービス", "訪問型サービスA",
             "通所介護相当サービス", "通所型サービスA"]
    return sorted(out, key=lambda x: (x[0] != "利用者数", order.index(x[1])))


def gamen3():
    """画面③ 介護予防ケアマネジメント。"""
    sg = sogo_jisseki()
    cp = {ev: mik for _k, ev, _j, mik, _n in chiiki_plan()}
    ry = {r[1]: r for r in chiiki_ryo()}
    ev = "介護予防ケアマネジメント"
    r6, r7, r8 = (sg[y]["事業"][ev] for y in ("令和6年度", "令和7年度", "令和8年度"))
    m8 = sg["令和8年度"]["月数"]
    rr = ry[ev]
    return [
        ("利用者数", ev, "人／月", r6[0] / 12, r7[0] / 12, r8[0] / m8,
         rr[5] / 12, rr[6] / 12, rr[7] / 12),
        ("事業費", ev, "円", r6[1], r7[1], r8[1] / m8 * 12,
         cp[ev], cp[ev], cp[ev]),
    ]


def matome():
    """画面④〜⑧の合計額（内訳は町の資料待ち）。"""
    cp = {ev: mik for _k, ev, _j, mik, _n in chiiki_plan()}
    return [
        ("④", "審査支払手数料・高額介護予防サービス相当事業等",
         cp["その他（審査支払手数料等）"], "総合事業",
         "**審査支払手数料に全額を寄せる。**"
         "年報 様式4 の決算額と国保連月報の差による"),
        ("⑤", "一般介護予防事業（5事業）", cp["一般介護予防事業"], "総合事業",
         "**介護予防普及啓発事業に全額を寄せる。**"
         "令和5年度993,702円から5.3倍に増えており、"
         "**何の事業かとあわせて内訳の確認をお願いしています**"),
        ("⑥", "包括的支援事業（センター運営）及び任意事業",
         cp["包括的支援事業・任意事業"], "包括的支援事業・任意事業",
         "**地域包括支援センターの運営に全額を寄せる。**"
         "社会保障充実分6事業と合わせた額で、"
         "**8欄への割り付けには決算書が要ります**"),
        ("⑦", "包括的支援事業（社会保障充実分）", 0, "包括的支援事業・任意事業",
         "**⑥に寄せるため0**"),
        ("⑧", "特定地域居宅サービス等事業・介護情報利活用事業", 0, "―",
         "小野町では実施していない"),
    ]


def selfcheck(G1, G3, MT):
    bad = []
    sogo = (sum(x[6] for x in G1 if x[0] == "事業費")
            + sum(x[6] for x in G3 if x[0] == "事業費")
            + MT[0][2] + MT[1][2])
    if round(sogo) != SOGO_NEN:
        bad.append(f"総合事業費 {sogo:,.0f}≠{SOGO_NEN:,}")
    if MT[2][2] != HOKATSU_NEN:
        bad.append(f"包括的支援事業・任意事業費 {MT[2][2]:,}≠{HOKATSU_NEN:,}")
    if round(sogo) + MT[2][2] != CHIIKI_NEN:
        bad.append(f"合計 {sogo + MT[2][2]:,.0f}≠{CHIIKI_NEN:,}")
    return bad


# ---------------------------------------------------------------- 体裁

def new_doc():
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = JP_MIN
    st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), JP_MIN)
    return doc


def head(doc, text, level):
    h = doc.add_heading(text, level=level)
    for r in h.runs:
        r.font.name = JP_GO
        r._element.rPr.rFonts.set(qn("w:eastAsia"), JP_GO)


def _emph(p, text):
    for i, part in enumerate(str(text).split("**")):
        if not part:
            continue
        r = p.add_run(part)
        r.font.name = JP_MIN
        r._element.rPr.rFonts.set(qn("w:eastAsia"), JP_MIN)
        r.bold = bool(i % 2)


def body(doc, *paras):
    for t in paras:
        _emph(doc.add_paragraph(), t)


def note(doc, text):
    p = doc.add_paragraph()
    _emph(p, text)
    for r in p.runs:
        r.font.size = Pt(9)


def table(doc, header_row, rows, right_from=99):
    t = doc.add_table(rows=1, cols=len(header_row))
    t.style = "Table Grid"
    for i, h in enumerate(header_row):
        c = t.rows[0].cells[i]
        c.text = ""
        p = c.paragraphs[0]
        r = p.add_run(h)
        r.font.name = JP_GO
        r._element.rPr.rFonts.set(qn("w:eastAsia"), JP_GO)
        r.bold = True
        r.font.size = Pt(9)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            p = cells[i].paragraphs[0]
            _emph(p, v)
            for r in p.runs:
                r.font.size = Pt(9)
            if i >= right_from:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    doc.add_paragraph()


def style_head(ws, row=1):
    for c in ws[row]:
        if c.value is None:
            continue
        c.fill = HEAD
        c.font = Font(bold=True, color="FFFFFF", size=9)
        c.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True)
        c.border = BORDER


def body_style(ws, first=2, wrap=()):
    for r in ws.iter_rows(min_row=first):
        for c in r:
            c.font = Font(size=9)
            c.border = BORDER
            c.alignment = Alignment(vertical="top", wrap_text=c.column in wrap)


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w


def _n(v, unit):
    if unit == "人／月":
        return f"{v:,.1f}"
    return f"{v:,.0f}"


# ---------------------------------------------------------------- Word

def build_docx(G1, G3, MT):
    doc = new_doc()
    head(doc, "地域支援事業の見込み量推計 ― 画面の入力値", 0)
    body(doc, f"小野町　第10期介護保険事業計画　／　{ASOF_JP}")

    head(doc, "0　結論", 1)
    table(doc, ["#", "内容"], [
        ("1", "**画面の形式は想定どおりです。**"
              "「サービスごとの総数で登録する」のままにしてください。"
              "右上のボタンで「利用者数」と「事業費」を切り替えて、両方入れます"),
        ("2", "**R6〜R8（青い列）も手入力です。**"
              "実績値も自動では入りません。第2章の表のとおりお入れください"),
        ("3", "**令和8年度は国保連月報が3か月分しかありません。**"
              "町の当初予算があればそちらを、なければ3か月の年換算を"
              "入れてください（第2章の表はこの年換算です）"),
        ("4", "**内訳が分からない欄は、合計額を1つの欄に寄せてください。**"
              "保険料に効くのは「総合事業費」と「包括的支援事業・任意事業費」の"
              "2つの合計だけです（第3章）"),
        ("5", "**画面右上が「現在時点」になっています。**"
              "「将来推計1」「将来推計0925」では「公表時点」でした。"
              "**版が固定される「公表時点」を推奨します**（第4章）"),
    ])

    head(doc, "1　画面ごとの進め方", 1)
    table(doc, ["順", "画面", "すること"],
          [(a, b, c) for a, b, c in GAMEN])

    # ---- 2 画面①
    head(doc, "2　画面①　訪問型・通所型サービス利用者数／事業費", 1)
    head(doc, "(1)　利用者数（単位：人／月）", 2)
    table(doc, ["行", "R6", "R7", "R8", "R9", "R10", "R11"],
          [(lab,) + tuple(
              (f"**{_n(v, unit)}**" if i >= 3 else _n(v, unit))
              for i, v in enumerate([r6, r7, r8, r9, r10, r11]))
           for kind, lab, unit, r6, r7, r8, r9, r10, r11 in G1
           if kind == "利用者数"], right_from=1)
    note(doc, "※ 小数が入らない場合は四捨五入した整数（23／0／51／0 ほか）で"
              "構いません。**保険料には効きません。**"
              "※ R6・R7は国保連月報の年間件数÷12、R8は3か月の件数÷3。"
              "R9〜R11は令和6・7年度の第1号被保険者1人あたりの件数率の平均を"
              "第10期将来推計用推計人口に乗じたものです。")

    head(doc, "(2)　事業費（単位：円）", 2)
    table(doc, ["行", "R6", "R7", "R8", "R9", "R10", "R11"],
          [(lab,) + tuple(
              (f"**{_n(v, unit)}**" if i >= 3 else _n(v, unit))
              for i, v in enumerate([r6, r7, r8, r9, r10, r11]))
           for kind, lab, unit, r6, r7, r8, r9, r10, r11 in G1
           if kind == "事業費"], right_from=1)
    note(doc, "※ R6・R7は国保連月報の総合事業コード（訪問介護相当＝Ａ２、"
              "通所介護相当＝Ａ６）の年間合計。R8は3か月の合計×4。"
              "**R9〜R11は令和6年度の実績を据え置いたものです。**"
              "※ 訪問型サービスA・通所型サービスAは月報では令和6年度以降0件です。"
              "**町で実施していないかご確認ください。**")

    # ---- 3 画面③以降
    head(doc, "3　画面③以降", 1)
    head(doc, "(1)　介護予防ケアマネジメント", 2)
    table(doc, ["項目", "R6", "R7", "R8", "R9", "R10", "R11"],
          [(f"{kind}（{unit}）",) + tuple(
              (f"**{_n(v, unit)}**" if i >= 3 else _n(v, unit))
              for i, v in enumerate([r6, r7, r8, r9, r10, r11]))
           for kind, _lab, unit, r6, r7, r8, r9, r10, r11 in G3], right_from=1)

    head(doc, "(2)　内訳が分からない欄（合計額で仮置き）", 2)
    body(doc,
         "**保険料に効くのは「総合事業費」と「包括的支援事業・任意事業費」の"
         "2つの合計だけです。**"
         "総合事業費には調整交付金がかかり、包括的支援事業・任意事業費には"
         "かかりません。**この2つの境を越えなければ、"
         "どの欄に入れても保険料は変わりません。**")
    table(doc, ["画面", "項目", "令和6年度決算（年額）", "区分", "入れ方"],
          [(a, b, f"**{c:,}円**", d, e) for a, b, c, d, e in MT], right_from=2)
    note(doc, f"※ 令和7年度は年報 様式4 が未入力のため決算額が取れません。"
              f"令和6年度と同額で構いません。"
              f"※ 令和8年度は町の当初予算でお願いします。"
              f"※ R9〜R11は令和6年度決算の据え置きです。")

    head(doc, "(3)　合計の検算", 2)
    sogo = (sum(x[6] for x in G1 if x[0] == "事業費")
            + sum(x[6] for x in G3 if x[0] == "事業費")
            + MT[0][2] + MT[1][2])
    table(doc, ["区分", "年額", "3年計", "内容"], [
        ("介護予防・日常生活支援総合事業費", f"**{sogo:,.0f}円**",
         f"{sogo * 3:,.0f}円",
         "訪問介護相当＋通所介護相当＋介護予防ケアマネジメント＋"
         "審査支払手数料等＋一般介護予防事業。**調整交付金がかかる**"),
        ("包括的支援事業・任意事業費", f"**{MT[2][2]:,}円**",
         f"{MT[2][2] * 3:,}円", "**調整交付金はかからない**"),
        ("**合計**", f"**{CHIIKI_NEN:,}円**", f"**{CHIIKI_NEN * 3:,}円**",
         "年報 様式4 の令和6年度決算"),
    ], right_from=1)

    # ---- 4 気づいた点
    head(doc, "4　あわせて気づいた点", 1)
    table(doc, ["#", "点", "内容"], [(a, b, c) for a, b, c in CHUI])

    head(doc, "(1)　確かめていただきたいこと", 2)
    table(doc, ["#", "内容"], [
        ("1", "**「実績及び推計方法の設定」で、"
              "（3）2）・（4）3）・（4）4）の3つの伸びが"
              "「0（ゼロ）」になっているかをご確認ください。**"
              "新しい推計でも既定で別の値が入っている可能性があります"),
        ("2", "**令和8年度のチェックが外れているかをご確認ください**"),
        ("3", "**「公表時点データを使用する」に戻すかどうかをご判断ください。**"
              "令和6・7年度は完結年度なので値は変わらないはずですが、"
              "版が固定されるほうが再現性があります"),
        ("4", "**訪問型サービスA・通所型サービスA・配食・見守りを"
              "総合事業で実施しているかをご教示ください。**"
              "国保連月報では令和6年度以降0件です"),
    ])

    p = OUT / f"小野町_見える化_地域支援事業の画面入力_{ASOF}.docx"
    doc.save(p)
    return p


# ---------------------------------------------------------------- Excel

def build_xlsx(G1, G3, MT):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet("00_画面①")
    ws.append(["区分", "行", "単位", "R6", "R7", "R8", "R9", "R10", "R11"])
    style_head(ws)
    for kind, lab, unit, *v in G1:
        ws.append([kind, lab, unit] + [round(x, 1) if unit == "人／月"
                                       else round(x) for x in v])
        for c in range(4, 10):
            ws.cell(ws.max_row, c).number_format = \
                "#,##0.0" if unit == "人／月" else "#,##0"
            ws.cell(ws.max_row, c).fill = IN
    body_style(ws)
    widths(ws, [10, 28, 10] + [14] * 6)
    ws.freeze_panes = "D2"
    ws.append([])
    for t in [
        "※ 画面名：訪問型・通所型サービス利用者数／事業費の実績値及び計画値の入力。",
        "※ 「サービスごとの総数で登録する」のまま。右上のボタンで"
        "「利用者数を入力する」⇄「事業費を入力する」を切り替える。",
        "※ R6〜R8（青い列）も手入力。R6・R7は国保連月報の年間合計、"
        "R8は3か月の年換算（町の当初予算があればそちらを優先）。",
        "※ 利用者数に小数が入らない場合は四捨五入した整数でよい。保険料には効かない。",
        "※ 訪問型サービスA・通所型サービスAは月報では令和6年度以降0件。"
        "町で実施していないかを確認する。",
    ]:
        ws.append([t])

    ws = wb.create_sheet("01_画面③以降")
    ws.append(["画面", "項目", "単位", "R6", "R7", "R8", "R9", "R10", "R11",
               "区分", "入れ方"])
    style_head(ws)
    for kind, lab, unit, *v in G3:
        ws.append(["③", f"{lab}（{kind}）", unit]
                  + [round(x, 1) if unit == "人／月" else round(x) for x in v]
                  + ["総合事業", "当方で用意できる"])
        for c in range(4, 10):
            ws.cell(ws.max_row, c).number_format = \
                "#,##0.0" if unit == "人／月" else "#,##0"
            ws.cell(ws.max_row, c).fill = IN
    for a, b, c, d, e in MT:
        ws.append([a, b, "円", c, c, ("町の当初予算" if c else 0), c, c, c, d,
                   e.replace("**", "")])
        for col in (4, 5, 7, 8, 9):
            ws.cell(ws.max_row, col).number_format = "#,##0"
            ws.cell(ws.max_row, col).fill = MACHI
    body_style(ws, wrap=(2, 11))
    widths(ws, [6, 40, 8] + [14] * 6 + [22, 54])
    ws.freeze_panes = "D2"
    ws.append([])
    ws.append(["※ 令和7年度は年報 様式4 が未入力のため令和6年度と同額で置く。"])
    ws.append(["※ R9〜R11は令和6年度決算の据え置き。"])

    ws = wb.create_sheet("02_合計の検算")
    sogo = (sum(x[6] for x in G1 if x[0] == "事業費")
            + sum(x[6] for x in G3 if x[0] == "事業費")
            + MT[0][2] + MT[1][2])
    ws.append(["区分", "項目", "年額", "3年計"])
    style_head(ws)
    KOMOKU = [
        ("総合事業", "訪問介護相当サービス",
         next(x[6] for x in G1 if x[0] == "事業費" and "訪問介護相当" in x[1])),
        ("総合事業", "通所介護相当サービス",
         next(x[6] for x in G1 if x[0] == "事業費" and "通所介護相当" in x[1])),
        ("総合事業", "介護予防ケアマネジメント",
         next(x[6] for x in G3 if x[0] == "事業費")),
        ("総合事業", "審査支払手数料等", MT[0][2]),
        ("総合事業", "一般介護予防事業", MT[1][2]),
    ]
    for ku, nm, v in KOMOKU:
        ws.append([ku, nm, round(v), round(v) * 3])
    ws.append(["総合事業", "小計", round(sogo), round(sogo) * 3])
    for c in (3, 4):
        ws.cell(ws.max_row, c).fill = OK
    ws.append(["包括的支援事業・任意事業", "包括的支援事業・任意事業",
               MT[2][2], MT[2][2] * 3])
    ws.append(["合計", "", CHIIKI_NEN, CHIIKI_NEN * 3])
    for c in (3, 4):
        ws.cell(ws.max_row, c).fill = OK
    for r in range(2, ws.max_row + 1):
        for c in (3, 4):
            ws.cell(r, c).number_format = "#,##0"
    body_style(ws)
    widths(ws, [26, 30, 16, 16])
    ws.freeze_panes = "A2"
    ws.append([])
    ws.append(["※ 出典：介護保険事業状況報告（年報）様式4 の令和6年度決算と、"
               "国民健康保険団体連合会の月報（総合事業コード）。"])
    ws.append(["※ 保険料に効くのは総合事業費と包括的支援事業・任意事業費の"
               "2つの合計だけ。総合事業費には調整交付金がかかる。"])

    ws = wb.create_sheet("03_注意")
    ws.append(["#", "点", "内容"])
    style_head(ws)
    for a, b, c in CHUI:
        ws.append([a, b.replace("**", ""), c.replace("**", "")])
    body_style(ws, wrap=(2, 3))
    widths(ws, [4, 44, 84])
    ws.freeze_panes = "A2"

    p = OUT / f"小野町_見える化_地域支援事業の画面入力_{ASOF}.xlsx"
    wb.save(p)
    return p


# ---------------------------------------------------------------- main

def main():
    OUT.mkdir(parents=True, exist_ok=True)
    G1, G3, MT = gamen1(), gamen3(), matome()
    bad = selfcheck(G1, G3, MT)
    p1 = build_docx(G1, G3, MT)
    p2 = build_xlsx(G1, G3, MT)
    print("出力:", p1)
    print("     ", p2)
    sogo = (sum(x[6] for x in G1 if x[0] == "事業費")
            + sum(x[6] for x in G3 if x[0] == "事業費")
            + MT[0][2] + MT[1][2])
    print(f"  画面① {len(G1)}行／画面③ {len(G3)}行／合計で仮置き {len(MT)}件")
    print(f"  総合事業費 年{sogo:,.0f}円／包括的支援事業・任意事業費 "
          f"年{MT[2][2]:,}円／合計 年{CHIIKI_NEN:,}円")
    print(f"  自己点検 {'OK' if not bad else 'NG: ' + ' / '.join(bad)}")


if __name__ == "__main__":
    main()
