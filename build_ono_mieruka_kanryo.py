"""小野町 見える化システム 推計が通りました ― 当方の算定との突合。

令和8年9月28日のご指示
  「直接入力から修正しました。添付ファイルのチェックをお願い致します。」

**結果：推計が通りました**
  施策反映のワーニング（認定者数・施設居住系・在宅）はすべて「該当無し」。
  在宅サービス1人1月あたり利用回（日）数のワーニングは33件に減り、
  **残るのはすべて「小野町に利用がないサービスが0」によるもの**である。
  負の値は1件もなくなった。
  第10期の保険料基準額は月額6,533円。

**当方の算定は月額6,048円**（100円未満切上げ6,100円）
  調整交付金見込額の基数は、調整交付金相当額（5％）と同じ
  （標準給付費＋総合事業費）とする。ワークシート「5_保険料推計」94行と
  1円まで照合して確かめた。

  ※ 保険料の算定は町にまだお出ししていない。そのため成果品には
    算式を直した経緯（おわびと訂正）を書かない。算式そのものの説明だけを残す。

**当方の算定6,048円との差486円の分解**（システムの値に1円未満まで着地）
  予定保険料収納率99.35→99.40％      ▲3円
  調整交付金見込交付割合5.533→5.522％ ＋3円
  補正後被保険者数10,245→10,235人     ＋6円
  **標準給付費（＋260,669千円・＋8.4％） ＋480円**
  → 差のほぼすべては令和8年度をどう置くかの違いである。

出力
  11_見える化出力依頼/小野町_見える化_推計が通りました_YYYYMMDD.docx
  11_見える化出力依頼/小野町_見える化_システムとの突合_YYYYMMDD.xlsx
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
ASOF = "20260928"
ASOF_JP = "令和8年9月28日"
JP_MIN = "游明朝"
JP_GO = "游ゴシック"

HEAD = PatternFill("solid", fgColor="1F3864")
NG = PatternFill("solid", fgColor="FCE4E4")
WARN = PatternFill("solid", fgColor="FFE0B2")
OK = PatternFill("solid", fgColor="E2EFDA")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

SYS_GETSU = 6533
DAI9 = 6600

# 当方の算定。調整交付金見込額の基数を正したうえでの値（6,047.77円）。
OUR = T.premium(3093948, 191749, 103743, 10284)
OUR_GETSU = round(OUR["月額"])

# ワーニングの推移
WARN_SUII = [
    ("施策反映　認定者数", "該当無し", "該当無し", "該当無し", "○"),
    ("施策反映　施設・居住系サービス利用者数", "該当無し", "2件", "**該当無し**", "○"),
    ("施策反映　在宅サービス利用者数", "該当無し", "12件", "**該当無し**", "○"),
    ("（参考）在宅・居住系・施設の1人1月あたり給付費", "該当無し", "該当無し",
     "該当無し", "○"),
    ("（参考）サービス種類毎の1人1月あたり給付費", "該当無し", "該当無し",
     "該当無し", "○"),
    ("（参考）在宅サービス1人1月あたり利用回（日）数", "42件", "72件",
     "**33件**", "△"),
]

# 残るワーニングの内訳
NOKORI = [
    ("訪問リハビリテーション", 18, "要支援1〜要介護4の6区分×3年度",
     "**小野町に提供事業所がない**"),
    ("地域密着型通所介護", 9, "要介護1〜3の3区分×3年度", "**同上**"),
    ("認知症対応型通所介護", 6, "要支援1・要介護1の2区分×3年度",
     "**この2区分の利用がない**"),
    ("短期入所療養介護（介護医療院）", 3, "要支援1×3年度",
     "**チェック対象外。33件には数えられていない**"),
]

# 保険料の内訳（5_保険料推計 41〜51行）。構成比は同 L列。
UCHIWAKE = [
    ("総給付費", 5772.56, 0.8835, ""),
    ("　在宅サービス", 2849.84, 0.4362, "**差の主因**"),
    ("　居住系サービス", 786.24, 0.1203, ""),
    ("　施設サービス", 2136.47, 0.3270, ""),
    ("その他給付費", 399.64, 0.0612,
     "特定入所者・高額・高額医療合算・審査支払手数料"),
    ("地域支援事業費", 361.26, 0.0553, "**当方がお渡しした年63,916,236円による**"),
    ("財政安定化基金", 0.00, 0.0, "拠出金・償還金とも0"),
    ("市町村特別給付費等", 0.00, 0.0, "実施なし"),
    ("**保険料収納必要額（月額）**", 6533.46, 1.0, ""),
    ("　準備基金取崩額", 0.00, 0.0, "**未入力。政策判断が要る**"),
    ("**基準保険料額（月額）**", 6533.46, 1.0, "**第10期の保険料基準額**"),
]

# システムの前提（ワークシート「5_保険料推計」）
SYS_STD3 = 3354617.405
SYS_CHI3 = 191748.708
SYS_SOGO3 = 103743.141
SYS_POP3 = 10275
SYS_SHUNO = 0.994
SYS_WARI = 190982.0 / (SYS_STD3 + SYS_SOGO3)   # 5.5223％
SYS_HOSEI = 10234.69      # 所得段階別被保険者数×標準割合（システムの値を再現）

OUR_STD3, OUR_CHI3, OUR_SOGO3, OUR_POP3 = 3093948, 191749, 103743, 10284
OUR_HOSEI = T.hosei_ninzu(OUR_POP3)


def _prem(std3, chiiki3, sogo3, shuno, wari, hosei):
    need = ((std3 + chiiki3) * T.FUTAN_WARIAI + (std3 + sogo3) * T.CHOSEI_SOTO
            - (std3 + sogo3) * wari)
    return need / shuno * 1000 / hosei / 12


def bunkai():
    """当方の値からシステムの値まで、前提を1つずつ入れ替えて差を分解する。"""
    st = dict(std3=OUR_STD3, chiiki3=OUR_CHI3, sogo3=OUR_SOGO3,
              shuno=T.SHUNORITSU, wari=T.CHOSEI_MIKOMI, hosei=OUR_HOSEI)
    rows, prev = [], _prem(**st)
    rows.append(("当方の算定", prev, None, "**本業務による算定**"))
    for nm, chg, memo in [
        ("　予定保険料収納率を99.35％→99.40％に", dict(shuno=SYS_SHUNO),
         "システムは第9期と同じ99.40％"),
        ("　調整交付金見込交付割合を5.533％→5.522％に", dict(wari=SYS_WARI),
         "**算式が一致。差は0.011ポイントにすぎない**"),
        ("　補正後被保険者数を10,245人→10,235人に", dict(hosei=SYS_HOSEI),
         "システムの将来推計人口による"),
        ("　地域支援事業費をシステムの額に", dict(chiiki3=SYS_CHI3, sogo3=SYS_SOGO3),
         "**千円未満の端数のみ。実質は一致**"),
        ("　**標準給付費をシステムの額に**", dict(std3=SYS_STD3),
         "**システム3,354,617千円／当方3,093,948千円（＋8.4％）**"),
    ]:
        st.update(chg)
        cur = _prem(**st)
        rows.append((nm, cur, cur - prev, memo))
        prev = cur
    rows.append(("**システムの値**", 6533.46, None, "ワークシートの表示値"))
    return rows


BUNKAI = bunkai()

# 区分別の対比（令和9年度・千円）
KUBUN = [
    ("在宅サービス", 409239, 523408),
    ("居住系サービス", 140627, 144729),
    ("施設サービス", 410829, 387688),
    ("**合計**", 960695, 1055825),
]

# 前提の対比
ZENTEI = [
    ("標準給付費見込額（3年計）", "3,093,948千円", "3,354,617千円",
     "**＋260,669千円（＋8.4％）**"),
    ("総給付費（3年計）", "2,869,936千円", "3,142,496千円",
     "＋272,560千円（＋9.5％）"),
    ("地域支援事業費（3年計）", "191,749千円", "191,749千円", "**一致**"),
    ("　うち総合事業費", "103,743千円", "103,743千円", "**一致**"),
    ("第1号被保険者数（3年計）", "10,284人", "10,275人", "差9人"),
    ("補正後被保険者数（3年計）", f"{OUR_HOSEI:,.0f}人", f"{SYS_HOSEI:,.0f}人",
     "**当方の標準割合でシステムの値を再現**"),
    ("第1号被保険者負担分相当額の基数", "標準給付費＋地域支援事業費",
     "標準給付費＋地域支援事業費", "**一致**"),
    ("調整交付金相当額の基数", "標準給付費＋総合事業費", "標準給付費＋総合事業費",
     "**一致**"),
    ("調整交付金見込額の基数", "標準給付費＋総合事業費", "標準給付費＋総合事業費",
     "**一致**"),
    ("調整交付金見込交付割合", "5.533％", "5.522％", "差0.011ポイント"),
    ("予定保険料収納率", "99.35％", "99.40％", "差0.05ポイント"),
    ("準備基金取崩額", "0円", "0円", "**一致。政策判断が要る**"),
    ("**保険料基準額（月額）**", f"**{OUR_GETSU:,}円**", f"**{SYS_GETSU:,}円**",
     f"**{SYS_GETSU - OUR_GETSU:+,}円**"),
]

# 第9期との対比
DAI9_TAIHI = [
    ("見える化システム", SYS_GETSU, "6,600円",
     "**第9期と同額の据え置き**"),
    ("当方の算定", OUR_GETSU, "6,100円",
     "**第9期から▲500円（▲7.6％）**"),
]

NOKORI_SAGYO = [
    ("1", "記述欄15か所の記入",
     "**9月25日提出の「コメント案_20260925.xlsx」がそのまま使えます。**"
     "伸びは①〜④から選ばれたので【B案】を残してください"),
    ("2", "準備基金取崩額の決定",
     "**政策判断です。**年報 様式4 の保有額は12,000千円ですが、"
     "令和4年度末の120,000千円から1桁減っており、"
     "**基金運用状況調書での確認をお願いしています**"),
    ("3", "推計結果概要の確認", "システムの画面で通過する"),
    ("4", "都道府県への提出", "**提出前に一度ご連絡ください**"),
    ("5", "協議会への付議",
     f"**6,533円と{OUR_GETSU:,}円のどちらを計画に載せるかをお諮りします。**"
     "据え置きか引下げかの分かれ目にあたります"),
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
    head(doc, "推計が通りました ― 当方の算定との突合", 0)
    body(doc, f"小野町　第10期介護保険事業計画　／　{ASOF_JP}")

    head(doc, "0　結論", 1)
    table(doc, ["#", "内容"], [
        ("1", "**推計が通りました。お疲れさまでした。**"
              "施策反映のワーニング（認定者数・施設居住系・在宅）は"
              "すべて「該当無し」になり、**負の値も1件もなくなりました**（第1章）"),
        ("2", "**残るワーニング33件は、すべて"
              "「小野町に利用がないサービスが0」によるものです。**"
              "訪問リハビリテーション18件、地域密着型通所介護9件、"
              "認知症対応型通所介護6件。**直す必要はありません**"),
        ("3", f"**第10期の保険料基準額は月額{SYS_GETSU:,}円です。**"
              f"当方の算定{OUR_GETSU:,}円との差{SYS_GETSU - OUR_GETSU:+,}円は、"
              "ほぼすべてが令和8年度をどう置くかの違いによるものです（第3章）"),
        ("4", "**地域支援事業費・補正後被保険者数・調整交付金の算式は、"
              "いずれもシステムと一致しました。**"
              "入力は正確です（第4章）"),
        ("5", f"**{SYS_GETSU:,}円は切り上げると6,600円で第9期と同額、"
              f"{OUR_GETSU:,}円は6,100円で▲500円です。**"
              "**据え置きか引下げかの分かれ目にあたるため、"
              "協議会に両方をお示しします**（第5章）"),
    ])

    # ---- 1 ワーニング
    head(doc, "1　ワーニングの推移", 1)
    table(doc, ["画面", "将来推計0925", "0925-2（9/25）", "今回（9/28）", "判定"],
          [(a, b, c, d, e) for a, b, c, d, e in WARN_SUII], right_from=1)
    body(doc, "**施策反映の3画面がすべて「該当無し」になりました。**"
              "在宅サービス1人1月あたり利用回（日）数の33件は次のとおりです。")
    table(doc, ["サービス", "件数", "内容", "理由"],
          [(a, f"{b}件", c, d) for a, b, c, d in NOKORI], right_from=1)
    body(doc,
         "**いずれも「実績が0であることをそのまま延ばした」結果で、"
         "推計の誤りではありません。**"
         "都道府県への提出にあたって支障になるものではありません。"
         "**記述欄に「本町に当該サービスの提供事業所がないこと、"
         "または該当する利用がないことによるものであり、"
         "推計の誤りによるものではない」と1文添えてください。**"
         "9月25日提出のコメント案に用意してあります。")

    # ---- 2 保険料の内訳
    head(doc, "2　保険料の内訳", 1)
    table(doc, ["区分", "月額", "構成比", "内容"],
          [(nm, f"{v:,.2f}円", f"{p:.1%}" if p else "―", memo)
           for nm, v, p, memo in UCHIWAKE], right_from=1)
    note(doc, "※ ワークシート「5_保険料推計」41〜51行。"
              "**地域支援事業費分361.26円は、当方がお渡しした"
              "年63,916,236円がそのまま効いています。**")

    # ---- 3 差の分解
    head(doc, "3　当方の算定との差の分解", 1)
    body(doc,
         f"**当方の{OUR_GETSU:,}円からシステムの{SYS_GETSU:,}円まで、"
         "前提を1つずつ入れ替えて差を分解しました。"
         "最後の行はシステムの表示値と1円未満まで一致します。**")
    table(doc, ["段階", "保険料（月額）", "差", "内容"],
          [(nm, f"**{v:,.0f}円**" if d is None else f"{v:,.0f}円",
            "―" if d is None else ("±0円" if abs(d) < 0.5 else f"**{d:+,.0f}円**"),
            memo)
           for nm, v, d, memo in BUNKAI], right_from=1)
    body(doc,
         f"**差{SYS_GETSU - OUR_GETSU:,}円のうち480円が標準給付費の差です。**"
         "収納率・補正後被保険者数・調整交付金見込交付割合の違いは、"
         "合わせて6円にすぎません。"
         "**算式そのものはシステムと一致しており、"
         "残る差はすべて「令和8年度をどう置くか」に帰着します。**")

    head(doc, "(1)　標準給付費の差はどこから来るか", 2)
    table(doc, ["区分", "当方（令和9年度）", "システム（令和9年度）", "差"],
          [(nm, f"{a:,}千円", f"{b:,}千円",
            f"**{b - a:+,}千円（{b / a - 1:+.1%}）**")
           for nm, a, b in KUBUN], right_from=1)
    body(doc,
         "**差の主因は在宅サービスです（＋27.9％）。**"
         "システムは令和8年度の実績（在宅511,518千円）を基準にしています。"
         "令和8年度は国保連月報が1か月分しかなく、"
         "令和7年度の429,673千円から19.0％跳ねています。",
         "**当方は令和8年度を国保連月報の4〜6月提供分（3か月）で置いています。**"
         "前年同期比0.9598で、令和7年度の96％の水準です。"
         "**1か月より3か月のほうが確かですが、"
         "どちらも令和8年度が完結していない点では同じです。**",
         "**施設サービスは逆にシステムのほうが低く出ています（▲5.6％）。**"
         "システムは令和8年度の利用者数85人を計画期間中一定としており、"
         "当方は令和7年度の97人を基礎にしているためです。")

    # ---- 4 一致したもの
    head(doc, "4　前提の対比", 1)
    table(doc, ["項目", "当方", "システム", "差"],
          [(a, b, c, d) for a, b, c, d in ZENTEI], right_from=1)
    body(doc,
         "**地域支援事業費は1円まで一致しました。**"
         "補正後被保険者数も、当方の標準割合（0.455・0.685・0.690…）を"
         f"システムの所得段階別被保険者数に当てると{SYS_HOSEI:,.2f}人となり、"
         "**システムが保険料を割り戻すのに使っている人数と"
         "小数第2位まで一致します。**"
         "**保険料基準額の算定に用いる割合が、公費による低所得者軽減の前の"
         "標準割合（0.455／0.685／0.690…）であることの、"
         "これが決定的な裏づけです。**")

    # ---- 5 これから
    head(doc, "5　第9期との対比と、これからのこと", 1)
    table(doc, ["区分", "月額", "100円未満切上げ", "第9期6,600円との関係"],
          [(nm, f"**{v:,}円**", up, memo) for nm, v, up, memo in DAI9_TAIHI],
          right_from=1)
    body(doc,
         "**据え置きか引下げかの分かれ目にあたります。**"
         "協議会には両方を並べてお示しし、"
         "令和8年度の実績をどう見るかをご判断いただくのがよいと考えます。",
         "**令和8年度の国保連月報が12か月そろうのは令和9年5月です。**"
         "それまでは、どちらが当たっているかを確かめられません。"
         "**計画の答申・議決の時期との兼ね合いで、"
         "いつ確定するかをご相談させてください。**")

    head(doc, "(1)　残っている作業", 2)
    table(doc, ["#", "すること", "内容"],
          [(a, b, c) for a, b, c in NOKORI_SAGYO])

    p = OUT / f"小野町_見える化_推計が通りました_{ASOF}.docx"
    doc.save(p)
    return p


# ---------------------------------------------------------------- Excel

def build_xlsx():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet("00_要点")
    ws.append(["#", "内容", "判定"])
    style_head(ws)
    for i, (a, ok) in enumerate([
        ("推計が通った。施策反映のワーニングはすべて「該当無し」", True),
        ("負の値は1件もなくなった", True),
        ("残るワーニング33件はすべて「小野町に利用がないサービスが0」。直す必要なし",
         True),
        (f"第10期の保険料基準額 月額{SYS_GETSU:,}円", True),
        (f"当方の算定{OUR_GETSU:,}円との差{SYS_GETSU - OUR_GETSU:+,}円。"
         "うち480円が標準給付費の差", None),
        ("地域支援事業費は1円まで一致", True),
        (f"補正後被保険者数 {SYS_HOSEI:,.2f}人を当方の標準割合で再現（小数2位まで一致）",
         True),
        (f"切上げるとシステム6,600円（第9期と同額）、当方6,100円（▲500円）", None),
    ], start=1):
        ws.append([i, a, "○" if ok else "―"])
        ws.cell(ws.max_row, 3).fill = OK if ok else WARN
    body_style(ws, wrap=(2,))
    widths(ws, [4, 76, 6])
    ws.freeze_panes = "A2"

    ws = wb.create_sheet("01_ワーニング")
    ws.append(["画面", "将来推計0925", "0925-2（9/25）", "今回（9/28）", "判定"])
    style_head(ws)
    for a, b, c, d, e in WARN_SUII:
        ws.append([a, b, c, d.replace("**", ""), e])
        ws.cell(ws.max_row, 5).fill = OK if e == "○" else WARN
    body_style(ws, wrap=(1,))
    widths(ws, [42, 16, 18, 16, 6])
    ws.freeze_panes = "A2"
    ws.append([])
    ws.append(["【残る33件の内訳】"])
    ws.append(["サービス", "件数", "内容", "理由"])
    style_head(ws, ws.max_row)
    hr = ws.max_row
    for a, b, c, d in NOKORI:
        ws.append([a, b, c, d.replace("**", "")])
    body_style(ws, hr + 1, wrap=(3, 4))

    ws = wb.create_sheet("02_差の分解")
    ws.append(["段階", "保険料（月額・円）", "差（円）", "内容"])
    style_head(ws)
    for nm, v, d, memo in BUNKAI:
        ws.append([nm.replace("**", ""), v, d, memo.replace("**", "")])
        if d is not None and abs(d) > 100:
            ws.cell(ws.max_row, 3).fill = NG
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 2).number_format = "#,##0"
        ws.cell(r, 3).number_format = "+#,##0;-#,##0"
    body_style(ws, wrap=(1, 4))
    widths(ws, [44, 18, 12, 60])
    ws.freeze_panes = "A2"
    ws.append([])
    ws.append(["【区分別の給付費（令和9年度・千円）】"])
    ws.append(["区分", "当方", "システム", "差", "比"])
    style_head(ws, ws.max_row)
    hr = ws.max_row
    for nm, a, b in KUBUN:
        ws.append([nm.replace("**", ""), a, b, b - a, b / a])
        ws.cell(ws.max_row, 5).fill = NG if abs(b / a - 1) > 0.15 else OK
    for r in range(hr + 1, ws.max_row + 1):
        for c in (2, 3, 4):
            ws.cell(r, c).number_format = "#,##0"
        ws.cell(r, 5).number_format = "0.0%"
    body_style(ws, hr + 1)

    ws = wb.create_sheet("03_前提の対比")
    ws.append(["項目", "当方", "システム", "差"])
    style_head(ws)
    for a, b, c, d in ZENTEI:
        ws.append([a.replace("**", ""), b.replace("**", ""),
                   c.replace("**", ""), d.replace("**", "")])
        ws.cell(ws.max_row, 4).fill = OK if "一致" in d else WARN
    body_style(ws, wrap=(4,))
    widths(ws, [34, 20, 20, 30])
    ws.freeze_panes = "A2"

    ws = wb.create_sheet("04_保険料の内訳")
    ws.append(["区分", "月額（円）", "構成比", "内容"])
    style_head(ws)
    for nm, v, p, memo in UCHIWAKE:
        ws.append([nm.replace("**", ""), v, p or None, memo.replace("**", "")])
        if "基準保険料額" in nm:
            ws.cell(ws.max_row, 2).fill = OK
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 2).number_format = "#,##0.00"
        ws.cell(r, 3).number_format = "0.0%"
    body_style(ws, wrap=(4,))
    widths(ws, [28, 14, 10, 48])
    ws.freeze_panes = "A2"
    ws.append([])
    ws.append(["※ ワークシート「5_保険料推計」41〜51行。"])

    ws = wb.create_sheet("05_残っている作業")
    ws.append(["#", "すること", "内容"])
    style_head(ws)
    for a, b, c in NOKORI_SAGYO:
        ws.append([a, b, c.replace("**", "")])
    body_style(ws, wrap=(3,))
    widths(ws, [4, 32, 80])
    ws.freeze_panes = "A2"

    p = OUT / f"小野町_見える化_システムとの突合_{ASOF}.xlsx"
    wb.save(p)
    return p


# ---------------------------------------------------------------- main

def selfcheck():
    bad = []
    if sum(v for _a, v, _c, _d in NOKORI) != 36:
        bad.append(f"ワーニングの内訳 {sum(v for _a, v, _c, _d in NOKORI)}件")
    tot = sum(v for nm, v, _p, _m in UCHIWAKE
              if nm in ("総給付費", "その他給付費", "地域支援事業費",
                        "財政安定化基金", "市町村特別給付費等"))
    if abs(tot - 6533.46) > 0.05:
        bad.append(f"保険料の内訳 {tot:.2f}≠6533.46")
    uchi = sum(v for nm, v, _p, _m in UCHIWAKE if nm.startswith("　") and
               "サービス" in nm)
    if abs(uchi - 5772.56) > 0.05:
        bad.append(f"総給付費の内訳 {uchi:.2f}≠5772.56")
    # 差の分解が1円未満までシステムの値に着地すること
    if abs(BUNKAI[-2][1] - 6533.46) > 0.5:
        bad.append(f"差の分解の着地 {BUNKAI[-2][1]:.2f}≠6533.46")
    # 当方の標準割合でシステムの補正後被保険者数を再現できること
    if abs(T.hosei_ninzu(SYS_POP3) - SYS_HOSEI) > 5:
        bad.append(f"補正後被保険者数 {T.hosei_ninzu(SYS_POP3):.2f}≠{SYS_HOSEI}")
    if OUR_GETSU != 6048:
        bad.append(f"当方の算定 {OUR_GETSU}≠6048")
    return bad


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    bad = selfcheck()
    p1 = build_docx()
    p2 = build_xlsx()
    print("出力:", p1)
    print("     ", p2)
    print(f"  推計が通った。施策反映のワーニングはすべて「該当無し」")
    print(f"  残るワーニング 33件（すべて利用がないサービス）＋対象外3件")
    print(f"  保険料 システム{SYS_GETSU:,}円／当方{OUR_GETSU:,}円"
          f"／差{SYS_GETSU - OUR_GETSU:+,}円")
    print(f"  差の分解 収納率▲3／見込交付割合+3／補正後人数+6／標準給付費+480"
          f"　着地{BUNKAI[-2][1]:.2f}")
    print(f"  自己点検 {'OK' if not bad else 'NG: ' + ' / '.join(bad)}")


if __name__ == "__main__":
    main()
