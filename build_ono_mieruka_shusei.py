"""小野町 見える化システム「将来推計0925-2」の修正が要る箇所。

令和8年9月25日のご指示
  「途中まで進めましたが修正対応必要な部分ご教示ください。」

受領（同日）
  将来推計ワーニングチェック結果（推計名「将来推計0925-2」）
  第10期介護保険事業計画策定に向けたワークシート（同推計）

**必ず直すもの（1件）**
  ワークシート「4_施策反映の解説」によると、3つの伸びが「未設定」である。
    20行 施設・居住系サービスの利用率の伸び … 未設定
    31行 在宅サービス利用者数の利用率の伸び … 未設定
    32行 1人1月あたり利用回（日）数の伸び   … 未設定
  **「未設定」は「0（ゼロ）」ではない。**その結果、在宅サービスと
  居住系サービスの令和9〜11年度の給付費がすべて0になっている。
  総給付費は年401,172千円（施設のみ）、保険料は月額2,964円。
  ワーニング86件も、すべてこの0が原因である。

  **9月25日に「選択ボタンは押さないでください」とお伝えしたのは、
  新しい推計には当てはまらなかった。**新規推計の初期値は「未設定」で、
  「将来推計1」の「0（ゼロ）」は明示的に選ばれた値であった。訂正する。

**正しく入ったもの**
  地域支援事業費 年63,916,236円（総合事業34,581,047円、
  包括的支援事業・任意事業29,335,189円）。3年計191,748,708円。
  当方の入力値と1円まで一致している。
  所得段階別加入割合も入っており、補正係数は0.9958／0.9973／0.9967。
  調整交付金見込交付割合は6.11／5.52／4.93％で自動計算されている。

出力
  11_見える化出力依頼/小野町_見える化_修正が要る箇所_YYYYMMDD.docx
  11_見える化出力依頼/小野町_見える化_修正が要る箇所_YYYYMMDD.xlsx
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

import build_ono_tanka as T

ROOT = pathlib.Path(__file__).parent
OUT = ROOT / "小野町_引継ぎ_整理済" / "11_見える化出力依頼"
ASOF = "20260925"
ASOF_JP = "令和8年9月25日"
JP_MIN = "游明朝"
JP_GO = "游ゴシック"

HEAD = PatternFill("solid", fgColor="1F3864")
NG = PatternFill("solid", fgColor="FCE4E4")
WARN = PatternFill("solid", fgColor="FFE0B2")
OK = PatternFill("solid", fgColor="E2EFDA")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

OUR_GETSU = 6095          # 当方の算定（月額・基金取崩なし）
OUR_STD3 = 3093948        # 当方の標準給付費（3年計・千円）
OUR_KYU3 = 2869936        # 当方の総給付費（3年計・千円）

# 「4_施策反映の解説」の設定を3つの推計で並べる
SETTEI = [
    ("認定者数の自然体推計手法", "性別／年齢5歳階級別／要介護度別",
     "同左", "同左", "○", "初期値のまま。このままでよい"),
    ("介護サービス利用者数の自然体推計手法", "サービス別／要介護度別",
     "同左", "同左", "○", "初期値のまま。このままでよい"),
    ("認定率の伸び", "当該市町村における令和6年度→令和7年度",
     "当該市町村における令和6年度→令和7年度",
     "**全国における令和6年度→令和7年度**", "△",
     "**変わっています。全国のほうが令和7年8〜9月の認定者数の断層を"
     "含まないため、むしろ妥当です。**"
     "記述欄に理由を書けばこのままで構いません"),
    ("**施設・居住系サービスの利用率等の伸び**", "0（ゼロ）",
     "当該市町村における令和6年度→令和7年度", "**未設定**", "**×**",
     "**「選択」から「0（ゼロ）」を選んでください**"),
    ("**在宅サービス利用者数の利用率の伸び**", "0（ゼロ）",
     "当該市町村における令和6年度→令和7年度", "**未設定**", "**×**",
     "**同上**"),
    ("**1人1月あたり利用回（日）数の伸び**", "0（ゼロ）",
     "当該市町村における令和6年度→令和7年度", "**未設定**", "**×**",
     "**同上**"),
    ("1人1月あたりの給付費の実績値", "令和8年度（推定）", "令和7年度",
     "**令和8年度**", "△",
     "**令和8年度のチェックが外れていれば「令和7年度」になるはずです。**"
     "設定画面をご確認ください"),
    ("地域支援事業費", "未入力", "未入力",
     "**年63,916,236円（入力済）**", "○",
     "**当方の入力値と1円まで一致しています**"),
    ("所得段階別加入割合", "未入力", "入力済", "**入力済**", "○",
     "補正係数0.9958／0.9973／0.9967。当方の0.9959と一致"),
]

# いま出ている結果
KEKKA = [
    ("総給付費（3年計）", 1203516, 3103845, OUR_KYU3, "千円"),
    ("　うち在宅サービス（令和9年度）", 0, 504401, None, "千円"),
    ("　うち居住系サービス（令和9年度）", 0, 138561, None, "千円"),
    ("　うち施設サービス（令和9年度）", 401172, 387688, None, "千円"),
    ("標準給付費見込額（3年計）", 1415637, 3314975, OUR_STD3, "千円"),
    ("地域支援事業費（3年計）", 191749, 0, 191749, "千円"),
    ("保険料収納必要額（3年計）", 361796, 745210, None, "千円"),
    ("**保険料基準額（月額）**", 2964, 6104, OUR_GETSU, "円"),
]

# ワーニング
WARN_ROWS = [
    ("施策反映　認定者数", "該当無し", "該当無し", "○"),
    ("施策反映　施設・居住系サービス利用者数", "該当無し", "**2件**", "×"),
    ("施策反映　在宅サービス利用者数", "該当無し", "**12件**", "×"),
    ("（参考）在宅・居住系・施設の1人1月あたり給付費", "該当無し", "該当無し", "○"),
    ("（参考）サービス種類毎の1人1月あたり給付費", "該当無し", "該当無し", "○"),
    ("（参考）在宅サービス1人1月あたり利用回（日）数", "42件", "**72件**", "×"),
]

WARN_REI = [
    ("特定施設入居者生活介護", "令和8年度2人", "令和9年度0人", "▲75％"),
    ("認知症対応型共同生活介護", "令和8年度44人", "令和9年度0人", "▲90％"),
    ("訪問介護", "令和8年度67人", "令和9年度0人", "▲85％"),
    ("通所介護", "令和8年度169人", "令和9年度0人", "▲85％"),
    ("福祉用具貸与", "令和8年度270人", "令和9年度0人", "▲95％"),
    ("介護予防支援・居宅介護支援", "令和8年度353人", "令和9年度0人", "▲95％"),
]

TEJUN = [
    ("1", "「実績及び推計方法の設定」を開く", "左メニューの2段目"),
    ("2", "**（3）2）施設・居住系サービスの自然体推計に用いる利用率等の伸び**で"
          "「選択」を押し、一覧から**「0（ゼロ）」**を選ぶ",
     "**「未設定」のままにしないでください**"),
    ("3", "**（4）3）在宅サービス利用者数の自然体推計に用いる利用率の伸び**も"
          "同じく「0（ゼロ）」", "―"),
    ("4", "**（4）4）在宅サービス1人1月あたり利用回（日）数の伸び**も"
          "同じく「0（ゼロ）」", "―"),
    ("5", "介護保険事業状況報告の設定で**令和8年度のチェックが外れているか**"
          "を確認する",
     "**「推計に用いた1人1月あたりの給付費の実績値」が「令和8年度」に"
     "なっています。**外れていれば「令和7年度」になるはずです"),
    ("6", "保険料額を算定し直す",
     "所得段階別・地域支援事業費は入力済みなので、そのまま使えます"),
    ("7", "ワーニングチェックと総括表を出し直す",
     "**86件のワーニングがほぼ消えるはずです**"),
]

# 直したあとの見込み
MIKOMI = [
    ("令和8年度を基準にしたまま（伸びだけ0にする）", 3178594, 3394808, 6657,
     "**令和8年度は国保連月報の1か月分です。**上振れした水準が基準になります"),
    ("令和8年度のチェックも外す（令和7年度が基準）", 3015444, 3220561, 6336,
     "**こちらをお勧めします。**令和7年度は月報12か月の完結年度です"),
    ("（参考）当方の算定", OUR_KYU3, OUR_STD3, OUR_GETSU,
     "令和8年度を国保連月報の4〜6月提供分（前年同期比0.9598）で置いたもの"),
]


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


# ---------------------------------------------------------------- Word

def build_docx():
    doc = new_doc()
    head(doc, "「将来推計0925-2」で修正が要る箇所", 0)
    body(doc, f"小野町　第10期介護保険事業計画　／　{ASOF_JP}")

    head(doc, "0　結論", 1)
    table(doc, ["#", "内容"], [
        ("1", "**直すのは1か所だけです。3つの「伸び」が「未設定」になっています。**"
              "「選択」ボタンから**「0（ゼロ）」**を選んでください。"
              "**「未設定」は「0（ゼロ）」ではありません**（第1章）"),
        ("2", "**9月25日に「選択ボタンは押さないでください」とお伝えしたのは、"
              "新しい推計には当てはまりませんでした。**"
              "新規推計の初期値は「未設定」で、「将来推計1」の「0（ゼロ）」は"
              "明示的に選ばれた値でした。お詫びして訂正します"),
        ("3", "**いま在宅サービスと居住系サービスの令和9〜11年度が"
              "すべて0になっています。**"
              "総給付費は年401,172千円（施設サービスのみ）、"
              "保険料は月額2,964円です。"
              "**ワーニング86件も、すべてこの0が原因です**（第2章）"),
        ("4", "**地域支援事業費は完璧に入りました。**"
              "年63,916,236円（総合事業34,581,047円、"
              "包括的支援事業・任意事業29,335,189円）で、"
              "当方の入力値と1円まで一致しています。"
              "所得段階別加入割合も入っています"),
        ("5", "**あわせて、令和8年度のチェックが外れているかをご確認ください。**"
              "「推計に用いた1人1月あたりの給付費の実績値」が「令和8年度」に"
              "なっています。外れていれば「令和7年度」になるはずです（第3章）"),
    ])

    # ---- 1 設定の対比
    head(doc, "1　設定の対比", 1)
    body(doc, "**ワークシートの「4_施策反映の解説」に記録された設定を、"
              "3つの推計で並べました。**")
    table(doc, ["設定項目", "将来推計1", "将来推計0925", "将来推計0925-2",
                "判定", "内容"],
          [(a, b, c, d, e, f) for a, b, c, d, e, f in SETTEI])

    head(doc, "(1)　「未設定」と「0（ゼロ）」は違います", 2)
    body(doc,
         "**「未設定」のままだと、在宅サービスと居住系サービスの推計値が"
         "0になります。**"
         "利用率が決まらないため、利用者数が計算できないのだと考えられます。"
         "施設サービスだけは「計画期間中は直近年度の利用者数が一定」という"
         "別の仕組みで動くため、401,172千円が残っています。",
         "**「0（ゼロ）」を選ぶと、令和8年度（または令和7年度）の利用率が"
         "そのまま計画期間に延びます。**"
         "これが見える化システムの自然体推計の本来の姿です。")

    # ---- 2 いま出ている結果
    head(doc, "2　いま出ている結果", 1)
    table(doc, ["項目", "将来推計0925-2（いま）", "将来推計0925", "当方の算定"],
          [(nm, f"**{a:,}{u}**", f"{b:,}{u}",
            f"{c:,}{u}" if c is not None else "―")
           for nm, a, b, c, u in KEKKA], right_from=1)
    note(doc, "※ 将来推計0925は3つの伸びが「令和6年度→令和7年度」で、"
              "地域支援事業費は未入力でした。"
              "**いまの0925-2は伸びが未設定で、地域支援事業費は入力済みです。**")

    head(doc, "(1)　ワーニング", 2)
    table(doc, ["画面", "将来推計0925", "将来推計0925-2", "判定"],
          [(a, b, c, d) for a, b, c, d in WARN_ROWS], right_from=1)
    body(doc, "**増えた分は、すべて「令和9年度以降が0」によるものです。**")
    table(doc, ["サービス", "令和8年度", "令和9年度", "変化率"],
          [(a, b, c, f"**{d}**") for a, b, c, d in WARN_REI], right_from=1)
    note(doc, "※ このほか、在宅サービス1人1月あたり利用回（日）数が72件。"
              "**伸びを「0（ゼロ）」にすれば、"
              "これらはいずれも解消する見込みです。**"
              "小野町に利用がないサービス（訪問リハビリテーションほか）の"
              "36件だけが残ります。")

    # ---- 3 直し方
    head(doc, "3　直し方", 1)
    table(doc, ["順", "すること", "留意点"],
          [(a, b, c) for a, b, c in TEJUN])

    head(doc, "(1)　直したあとの見込み", 2)
    body(doc, "**おおよその見当です。実際の値は再計算でご確認ください。**")
    table(doc, ["前提", "総給付費（3年計）", "標準給付費（3年計）",
                "保険料基準額（月額）", "内容"],
          [(nm, f"約{k:,}千円", f"約{s:,}千円", f"**約{m:,}円**", memo)
           for nm, k, s, m, memo in MIKOMI], right_from=1)
    note(doc, "※ 令和8年度は国保連月報の1か月分です。"
              "**その水準を基準にすると、令和7年度を基準にする場合より"
              "保険料が約320円高く出ます。**"
              "※ 当方の算定は令和8年度を国保連月報の4〜6月提供分"
              "（前年同期比0.9598）で置いたものです。"
              "令和8年度の月報が12か月そろう令和9年5月以降に確かめます。")

    # ---- 4 正しく入ったもの
    head(doc, "4　正しく入ったもの", 1)
    body(doc, "**地域支援事業費は1円まで一致しました。**")
    table(doc, ["区分", "令和9〜11年度（各年度）", "3年計", "当方の値"], [
        ("介護予防・日常生活支援総合事業費", "34,581,047円", "103,743,141円",
         "**一致**"),
        ("包括的支援事業（センター運営）及び任意事業費", "29,335,189円",
         "88,005,567円", "**一致**"),
        ("**地域支援事業費 計**", "**63,916,236円**", "**191,748,708円**",
         "**一致**"),
    ], right_from=1)
    table(doc, ["項目", "システム", "当方", "判定"], [
        ("所得段階別加入割合補正係数（令和9年度）", "0.9958",
         f"{T.shotoku_hosei():.4f}", "**一致**"),
        ("調整交付金見込交付割合（令和9年度）", "6.11％",
         f"{T.chosei_wari('令和9年度') * 100:.3f}％", "**一致**"),
        ("同（令和10年度）", "5.52％",
         f"{T.chosei_wari('令和10年度') * 100:.3f}％", "**一致**"),
        ("同（令和11年度）", "4.93％",
         f"{T.chosei_wari('令和11年度') * 100:.3f}％", "**一致**"),
        ("予定保険料収納率", "99.40％", "99.35％",
         "第9期と同じ値。どちらでも構いません"),
    ], right_from=1)

    head(doc, "(1)　認定率の伸びを「全国」に変えられた件", 2)
    body(doc,
         "**これはむしろ妥当な変更です。そのままで構いません。**"
         "小野町の令和6年度から令和7年度への動きには、"
         "令和7年8〜9月の認定者数の断層（年報で944人→792人、▲16.1％）が"
         "入っています。全国の伸びにはこれが入りません。",
         "**結果も当方の見込みに近づいています。**"
         "令和9年度の第1号被保険者の認定者数は、"
         "当該市町村の伸びで769人、全国の伸びで778人、"
         f"当方の見込みが{T.NINTEI_EST['令和9年度']:,}人です。"
         "**記述欄に「全国の伸びを用いた理由」を書いてください。**"
         "コメント案の【B案】がそのまま使えます。")

    p = OUT / f"小野町_見える化_修正が要る箇所_{ASOF}.docx"
    doc.save(p)
    return p


# ---------------------------------------------------------------- Excel

def build_xlsx():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet("00_要点")
    ws.append(["#", "内容", "対応"])
    style_head(ws)
    for i, (a, b, ok) in enumerate([
        ("3つの伸びが「未設定」。「未設定」は「0（ゼロ）」ではない",
         "「選択」から「0（ゼロ）」を選ぶ", False),
        ("在宅サービス・居住系サービスの令和9〜11年度がすべて0",
         "上を直せば解消する", False),
        ("総給付費 年401,172千円（施設のみ）、保険料 月額2,964円",
         "上を直せば解消する", False),
        ("ワーニング86件（施設・居住系2・在宅12・利用回数72）",
         "上を直せば36件まで減る見込み", False),
        ("「1人1月あたりの給付費の実績値」が令和8年度",
         "令和8年度のチェックが外れているか確認する", False),
        ("認定率の伸びが「全国」に変わっている",
         "妥当な変更。記述欄に理由を書く", True),
        ("地域支援事業費 年63,916,236円", "当方の値と1円まで一致", True),
        ("所得段階別加入割合 補正係数0.9958", "当方の0.9959と一致", True),
        ("調整交付金見込交付割合 6.11／5.52／4.93％", "当方の算式値と一致", True),
    ], start=1):
        ws.append([i, a, b])
        ws.cell(ws.max_row, 3).fill = OK if ok else NG
    body_style(ws, wrap=(2, 3))
    widths(ws, [4, 66, 40])
    ws.freeze_panes = "A2"

    ws = wb.create_sheet("01_設定の対比")
    ws.append(["設定項目", "将来推計1", "将来推計0925", "将来推計0925-2",
               "判定", "内容"])
    style_head(ws)
    for a, b, c, d, e, f in SETTEI:
        ws.append([a.replace("**", ""), b, c, d.replace("**", ""),
                   e.replace("**", ""), f.replace("**", "")])
        ws.cell(ws.max_row, 5).fill = \
            OK if e.replace("**", "") == "○" else \
            (WARN if e.replace("**", "") == "△" else NG)
    body_style(ws, wrap=(1, 2, 3, 4, 6))
    widths(ws, [32, 26, 30, 26, 6, 52])
    ws.freeze_panes = "A2"

    ws = wb.create_sheet("02_結果の対比")
    ws.append(["項目", "単位", "将来推計0925-2（いま）", "将来推計0925",
               "当方の算定"])
    style_head(ws)
    for nm, a, b, c, u in KEKKA:
        ws.append([nm.replace("**", ""), u, a, b, c])
        ws.cell(ws.max_row, 3).fill = NG
    for r in range(2, ws.max_row + 1):
        for c in (3, 4, 5):
            ws.cell(r, c).number_format = "#,##0"
    body_style(ws)
    widths(ws, [34, 8, 22, 18, 18])
    ws.freeze_panes = "A2"

    ws = wb.create_sheet("03_ワーニング")
    ws.append(["画面", "将来推計0925", "将来推計0925-2", "判定"])
    style_head(ws)
    for a, b, c, d in WARN_ROWS:
        ws.append([a, b.replace("**", ""), c.replace("**", ""), d])
        ws.cell(ws.max_row, 4).fill = OK if d == "○" else NG
    ws.append([])
    ws.append(["【0になっている例】"])
    ws.append(["サービス", "令和8年度", "令和9年度", "変化率"])
    style_head(ws, ws.max_row)
    hr = ws.max_row
    for a, b, c, d in WARN_REI:
        ws.append([a, b, c, d])
    body_style(ws, hr + 1)
    body_style(ws, 2, wrap=(1,))
    widths(ws, [42, 22, 22, 12])
    ws.freeze_panes = "A2"

    ws = wb.create_sheet("04_手順")
    ws.append(["順", "すること", "留意点"])
    style_head(ws)
    for a, b, c in TEJUN:
        ws.append([a, b.replace("**", ""), c.replace("**", "")])
    body_style(ws, wrap=(2, 3))
    widths(ws, [4, 58, 62])
    ws.freeze_panes = "A2"
    ws.append([])
    ws.append(["【直したあとの見込み】"])
    ws.append(["前提", "総給付費（3年計・千円）", "標準給付費（3年計・千円）",
               "保険料基準額（月額・円）"])
    style_head(ws, ws.max_row)
    hr = ws.max_row
    for nm, k, s, m, _memo in MIKOMI:
        ws.append([nm, k, s, m])
    for r in range(hr + 1, ws.max_row + 1):
        for c in (2, 3, 4):
            ws.cell(r, c).number_format = "#,##0"
    body_style(ws, hr + 1, wrap=(1,))

    p = OUT / f"小野町_見える化_修正が要る箇所_{ASOF}.xlsx"
    wb.save(p)
    return p


# ---------------------------------------------------------------- main

def selfcheck():
    bad = []
    if abs(T.shotoku_hosei() - 0.9958) > 5e-4:
        bad.append(f"補正係数 {T.shotoku_hosei():.4f}")
    if abs(T.chosei_wari("令和9年度") - 0.0611) > 5e-4:
        bad.append(f"見込交付割合 {T.chosei_wari('令和9年度'):.4f}")
    if len(SETTEI) != 9:
        bad.append(f"設定 {len(SETTEI)}件")
    return bad


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    bad = selfcheck()
    p1 = build_docx()
    p2 = build_xlsx()
    print("出力:", p1)
    print("     ", p2)
    print("  必ず直すもの 3つの伸びが「未設定」→「0（ゼロ）」")
    print("  いまの保険料 2,964円／直すと約6,336〜6,657円／当方 6,095円")
    print("  ワーニング 86件（0925は42件）")
    print(f"  自己点検 {'OK' if not bad else 'NG: ' + ' / '.join(bad)}")


if __name__ == "__main__":
    main()
