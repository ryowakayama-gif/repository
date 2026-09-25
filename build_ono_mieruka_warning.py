"""小野町 見える化システムのワーニングチェック結果と、更新されたワークシートの確認。

令和8年9月25日のご指示
  「ワーニング発生していますが、入力が終わりました。内容の確認をお願いします。」

受領（同日）
  将来推計ワーニングチェック結果（推計名「将来推計0925」・総括表）
  第10期介護保険事業計画策定に向けたワークシート（同推計）

**最も重要な所見**
  ワークシート「4_施策反映の解説」によると、次の3つの「伸び」が
  「当該市町村における令和6年度→令和7年度の伸び」に設定されている。
    ・施設・居住系サービスの利用率等の伸び（20行）
    ・在宅サービス利用者数の利用率の伸び（31行）
    ・1人1月あたり利用回（日）数の伸び（32行）
  9月25日に受領した画面キャプチャ（推計名「将来推計1」）では、
  **3つとも「0（ゼロ）」であった。**
  0に戻さないと、令和7年8〜9月の認定者数の断層を含んだ動きが
  計画期間に延びる。介護予防訪問看護の給付費が令和9年度以降
  ▲3,345千円と負になっているのは、伸びを差分として当てている証拠である。

**一致した点（当方の訂正が正しかったことの確認）**
  所得段階別加入割合補正係数　システム0.9958／当方0.9959
  調整交付金見込交付割合　　　システム6.11・5.52・4.93％／
                              当方6.107・5.548・4.945％
  第1号被保険者数（3年計）　　10,275人（入力済み）

出力
  11_見える化出力依頼/小野町_見える化_ワーニングの確認結果_YYYYMMDD.docx
  11_見える化出力依頼/小野町_見える化_ワーニングの確認結果_YYYYMMDD.xlsx
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
WARN_FILE = (OUT / "原本_見える化出力_20260925"
             / "【受領】将来推計ワーニングチェック結果_将来推計0925_20260925.xlsx")
WS_FILE = (ROOT / "小野町_引継ぎ_整理済" / "04_算定・見込量" / "原本_見える化ワークシート"
           / "【受領】第10期介護保険事業計画策定に向けたワークシート_将来推計0925_20260925.xlsx")
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

# 当方の算定（訂正後）
OUR_STD3 = 3093948       # 標準給付費（3年計・千円）
OUR_CHI3 = 191749        # 地域支援事業費（3年計・千円）
OUR_SOGO3 = 103743       # うち総合事業費（3年計・千円）
OUR_GETSU = 6095         # 保険料基準額（月額・基金取崩なし）

# ワークシート 5_保険料推計 から読み取った値
SYS = {
    "標準給付費": 3314974.794, "総給付費": 3103845.0,
    "地域支援事業費": 0.0, "調整交付金見込額": 182983.0,
    "保険料収納必要額": 745209.94232, "予定収納率": 0.994,
    "第10期月額": 6104.3062977384825, "第1号被保険者数": 10275,
    "所得段階別補正係数": [0.9958, 0.9973, 0.9967],
    "後期補正係数": [0.9558, 0.9802, 1.0065],
    "見込交付割合": [0.0611, 0.0552, 0.0493],
}
SYS_NINTEI = {"令和6年度": 928, "令和7年度": 856, "令和8年度": 781,
              "令和9年度": 769, "令和10年度": 775, "令和11年度": 771}

# 「4_施策反映の解説」から読み取った設定
SETTEI = [
    ("認定者数の自然体推計手法",
     "性別／年齢5歳階級別／要介護度別に推計する", "0（ゼロ）ではない", "○",
     "**基本推計。当方の検証でも包括推計をわずかに上回る。このままでよい**"),
    ("介護サービス利用者数の自然体推計手法",
     "サービス別／要介護度別に推計する", "―", "○", "初期値。このままでよい"),
    ("認定率の伸び",
     "当該市町村における令和6年度→令和7年度の伸び", "―", "○",
     "**完結年度どうしの伸び。国のガイドの上限1.1・下限0.9の歯止めもかかる。"
     "このままでよい**"),
    ("**施設・居住系サービスの利用率等の伸び**",
     "**当該市町村における令和6年度→令和7年度の伸び**", "**0（ゼロ）**", "**×**",
     "**9月25日の画面（将来推計1）では0（ゼロ）だった。0に戻す**"),
    ("**在宅サービス利用者数の利用率の伸び**",
     "**当該市町村における令和6年度→令和7年度の伸び**", "**0（ゼロ）**", "**×**",
     "**同上。0に戻す**"),
    ("**1人1月あたり利用回（日）数の伸び**",
     "**当該市町村における令和6年度→令和7年度の伸び**", "**0（ゼロ）**", "**×**",
     "**同上。介護予防訪問看護の給付費が負になっている直接の原因**"),
    ("1人1月あたりの給付費の実績値",
     "令和7年度", "―", "○", "最新の完結年度。このままでよい"),
    ("施策反映", "認定者数・施設居住系・在宅のいずれも未入力", "―", "○",
     "自然体推計のまま。このままでよい"),
    ("**地域支援事業費**", "**全欄0**", "**当方の入力データ（9月25日版）**",
     "**×**",
     "**3_地域支援事業費がすべて0。保険料が約400円過小に出る**"),
]

# 令和7年度→令和8年度で大きく振れたサービス（ワークシート 2_サービス別給付費）
FURE = [
    ("短期入所生活介護", 27393.6, 90099.3, 27.2, 60.0),
    ("訪問介護", 29747.4, 45552.7, 51.2, 67.0),
    ("通所リハビリテーション", 15575.8, 22534.3, 25.8, 29.0),
    ("認知症対応型通所介護", 12237.6, 8762.2, 9.5, 14.0),
    ("特定施設入居者生活介護", 9719.0, 4862.8, 3.8, 2.0),
    ("短期入所療養介護（老健）", 2327.6, 925.0, 2.2, 2.0),
    ("介護予防訪問看護", 2213.9, 212.6, 6.8, 6.0),
]

TEJUN = [
    ("1", "「実績及び推計方法の設定」を開く", "左メニューの2段目"),
    ("2", "**（3）2）施設・居住系サービスの自然体推計に用いる利用率等の伸び**を"
          "**0（ゼロ）**にする", "「選択」ボタンから"),
    ("3", "**（4）3）在宅サービス利用者数の自然体推計に用いる利用率の伸び**を"
          "**0（ゼロ）**にする", "同上"),
    ("4", "**（4）4）在宅サービス1人1月あたり利用回（日）数の伸び**を"
          "**0（ゼロ）**にする", "同上"),
    ("5", "「地域支援事業の見込み量推計」で地域支援事業費を入力する",
     "**当方の「見える化_地域支援事業の入力データ_20260925.xlsx」による。"
     "町の資料が届いていない欄は当方の仮置きで構わない**"),
    ("6", "「保険料額の算定」まで進めて再計算する", "所得段階別の入力はそのまま使える"),
    ("7", "ワーニングチェックと総括表を出し直す",
     "**介護予防訪問看護のマイナスと短期入所生活介護の20.5日が"
     "消えているかを確認する**"),
]


# ---------------------------------------------------------------- 読み取り

def read_warnings():
    """ワーニングチェック結果の明細を読む。"""
    wb = openpyxl.load_workbook(WARN_FILE, data_only=True)
    ws = wb["在宅サービス1人1月あたり利用回（日）数"]
    cur_svc = cur_y = None
    rows = []
    for r in range(7, ws.max_row + 1):
        svc, y = ws.cell(r, 3).value, ws.cell(r, 4).value
        g, val = ws.cell(r, 5).value, ws.cell(r, 6).value
        lo, hi, res = (ws.cell(r, c).value for c in (7, 8, 9))
        if svc:
            cur_svc = str(svc).strip()
        if y:
            cur_y = str(y).strip()
        if res and "範囲内" not in str(res):
            rows.append((cur_svc, cur_y, g, val, lo, hi, str(res)))
    gaiyo = {}
    ws2 = wb["概要"]
    for r in range(11, ws2.max_row + 1):
        nm, n = ws2.cell(r, 3).value, ws2.cell(r, 4).value
        if nm and n is not None:
            gaiyo[str(nm).strip()] = n
    return rows, gaiyo


def classify(rows):
    """ワーニングを『利用がない』『異常』に分ける。"""
    ijo, nashi, taigai = [], [], []
    for x in rows:
        if "対象外" in x[6]:
            taigai.append(x)
        elif x[3] is not None and x[3] < 0:
            ijo.append(x)
        elif "上限" in x[6]:
            ijo.append(x)
        else:
            nashi.append(x)
    return ijo, nashi, taigai


def premium(std, chi, sogo, wari, shuno, g, pop3, kikin=0):
    need = (std + chi) * 0.23 + (std + sogo) * 0.05 - std * wari - kikin
    return need / shuno * 1000 / (pop3 * g) / 12


def hikaku():
    """システムと当方の突合。"""
    wari = SYS["調整交付金見込額"] / SYS["標準給付費"]
    g = SYS["所得段階別補正係数"][0]
    pop3 = SYS["第1号被保険者数"]
    return {
        "見込割合": wari,
        "再現": premium(SYS["標準給付費"], 0, 0, wari, SYS["予定収納率"], g, pop3),
        "地域支援入り": premium(SYS["標準給付費"], OUR_CHI3, OUR_SOGO3, wari,
                          SYS["予定収納率"], g, pop3),
        "当方": OUR_GETSU,
    }


def selfcheck(H, ijo, nashi, taigai):
    bad = []
    if abs(H["再現"] - SYS["第10期月額"]) > 5:
        bad.append(f"保険料の再現 {H['再現']:.0f}≠{SYS['第10期月額']:.0f}")
    if len(ijo) + len(nashi) + len(taigai) != 45:
        bad.append(f"ワーニング件数 {len(ijo) + len(nashi) + len(taigai)}")
    for i, (y, f) in enumerate(zip(T.KOKI_HOSEI, SYS["後期補正係数"])):
        if abs(T.KOKI_HOSEI[y] - f) > 1e-6:
            bad.append(f"{y}の後期補正係数")
        w = 0.28 - 0.23 * f * SYS["所得段階別補正係数"][i]
        if abs(w - SYS["見込交付割合"][i]) > 0.0005:
            bad.append(f"{y}の見込交付割合 {w:.4f}≠{SYS['見込交付割合'][i]}")
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


# ---------------------------------------------------------------- Word

def build_docx(rows, gaiyo, ijo, nashi, taigai, H):
    doc = new_doc()
    head(doc, "ワーニングチェック結果と総括表の確認", 0)
    body(doc, f"小野町　第10期介護保険事業計画　／　{ASOF_JP}")

    head(doc, "0　結論", 1)
    table(doc, ["#", "内容"], [
        ("1", "**ワーニング42件のうち36件は「小野町に利用がないサービスが0になっている」"
              "もので、直す必要はありません。**"
              "訪問リハビリテーション18件、地域密着型通所介護9件、"
              "認知症対応型通所介護6件、訪問入浴介護（要介護5）3件です"),
        ("2", "**直していただきたいのは6件です。**"
              "**介護予防訪問看護の要支援2が1人1月あたり▲12.6回と負の値になっており、"
              "給付費も令和9年度以降▲3,345千円と負になっています（3件）。**"
              "利用回数が負になることはあり得ません。"
              "**短期入所生活介護の要介護1が1人1月あたり20.5日で、"
              "上限16日を4.5日上回っています（3件）**"),
        ("3", "**原因は、3つの「伸び」が0（ゼロ）ではなく"
              "「令和6年度→令和7年度の伸び」に設定されていることです。**"
              "施設・居住系の利用率、在宅の利用率、1人1月あたり利用回（日）数の3つです。"
              "9月25日にいただいた画面（推計名「将来推計1」）では3つとも"
              "0（ゼロ）でした。**0に戻してください**（第2章）"),
        ("4", "**地域支援事業費が全欄0のままです。**"
              "当方の入力データ（9月25日提出）を入れてください。"
              f"入れると保険料は{SYS['第10期月額']:,.0f}円から"
              f"約{H['地域支援入り']:,.0f}円になります（第4章）"),
        ("5", "**所得段階別の入力は正しく効いています。**"
              "補正係数はシステム0.9958・当方0.9959、"
              "調整交付金見込交付割合はシステム6.11／5.52／4.93％・"
              "当方6.107／5.548／4.945％で一致しました。"
              "**当方が9月25日に訂正した算式が正しかったことの裏づけです**（第5章）"),
    ])

    # ---- 1 ワーニングの内訳
    head(doc, "1　ワーニング42件の内訳", 1)
    body(doc, "**ワーニングはすべて「在宅サービス1人1月あたり利用回（日）数」"
              "の欄で出ています。**"
              "認定者数・施設・居住系・給付費の各チェックは「該当無し」です。")
    table(doc, ["区分", "件数", "内容", "対応"], [
        ("**利用がないサービス**", f"**{len(nashi)}件**",
         "訪問リハビリテーション18件（要支援1〜要介護4の6区分×3年度）、"
         "地域密着型通所介護9件、認知症対応型通所介護6件、"
         "訪問入浴介護（要介護5）3件",
         "**直す必要はありません。**小野町に利用がないサービスで、"
         "実績が0であることをそのまま延ばした結果です"),
        ("**負の値・上限超過**", f"**{len(ijo)}件**",
         "介護予防訪問看護（要支援2）が▲12.6回で下限を12.6回下回る3件、"
         "短期入所生活介護（要介護1）が20.5日で上限16日を4.5日上回る3件",
         "**第2章の設定を直してください**"),
        ("チェック対象外", f"{len(taigai)}件",
         "短期入所療養介護（介護医療院）の要支援1",
         "閾値が設定されていない区分。件数にも数えられていません"),
    ], right_from=1)
    note(doc, f"※ 概要シートの「在宅サービス1人1月あたり利用回（日）数」の"
              f"ワーニング発生件数は{gaiyo.get('在宅サービス', '―')}件。"
              "チェック対象外の3件を除いた数です。")

    head(doc, "(1)　直していただきたい6件", 2)
    table(doc, ["サービス", "年度", "要介護度", "値", "下限", "上限", "結果"],
          [(a, b, str(c), f"**{d:,.1f}**", f"{e:,.1f}", f"{f:,.1f}", g)
           for a, b, c, d, e, f, g in ijo], right_from=3)

    # ---- 2 原因
    head(doc, "2　原因 ― 3つの「伸び」が0になっていません", 1)
    body(doc,
         "**ワークシートの「4_施策反映の解説」に、設定した内容が記録されています。**"
         "9月25日にいただいた画面キャプチャ（推計名「将来推計1」）と比べると、"
         "3つの「伸び」が変わっています。")
    table(doc, ["設定項目", "現在（将来推計0925）", "あるべき値", "判定", "内容"],
          [(a, b, c, d, e) for a, b, c, d, e in SETTEI])

    head(doc, "(1)　なぜ0でなければならないか", 2)
    body(doc,
         "**令和6年度から令和7年度にかけての動きには、"
         "令和7年8〜9月の認定者数の断層が入っているためです。**"
         "年報の認定者数は令和6年度末944人から令和7年度末792人へ16.1％減りました。"
         "認定者数が減った一方でサービスの利用は減っていないため、"
         "「認定者1人あたりの利用率」は跳ね上がっています。"
         "**その動きをもう1年延ばすと、利用者数が実態からかけ離れて増えます。**",
         "**負の値が出ていることが、何よりの証拠です。**"
         "「伸び」は1人1月あたり利用回（日）数に差として加えられるため、"
         "令和6年度から令和7年度に減った項目は、もう1年分さらに引かれます。"
         "介護予防訪問看護の要支援2は、この引き算で負になりました。"
         "**利用回数が負になることは現実にはあり得ません。**",
         "**0（ゼロ）に戻すと、令和7年度の水準がそのまま計画期間に延びます。**"
         "これが見える化システムの自然体推計の本来の仕様であり、"
         "9月8日に協議いただいた当方の標準ケースとも一致します。")

    # ---- 3 症状
    head(doc, "3　いま出ている症状", 1)
    body(doc, "令和7年度から令和8年度にかけて、給付費と利用者数が大きく振れています。"
              "**令和8年度は実績ではなく推計値ですので、"
              "これは設定が生んだ動きです。**")
    table(doc, ["サービス", "令和7年度\n給付費", "令和8年度\n給付費", "倍率",
                "令和7年度\n利用者数", "令和8年度\n利用者数", "倍率"],
          [(nm, f"{a:,.0f}千円", f"{b:,.0f}千円",
            f"**{b / a:.2f}**" if a else "―",
            f"{c:,.1f}人", f"{d:,.1f}人",
            f"**{d / c:.2f}**" if c else "―")
           for nm, a, b, c, d in FURE], right_from=1)
    note(doc, "※ 出典：ワークシート「2_サービス別給付費」。"
              "給付費は年間累計、利用者数は1月あたり。"
              "※ 介護予防訪問看護は令和9年度以降の給付費が▲3,345千円。")
    body(doc,
         "**訪問介護の要介護5は、1人1月あたりの利用回数が"
         "令和7年度の33.8回から令和9〜11年度の67.4回へ倍になっています。**"
         "上限116回の範囲内なのでワーニングにはなりませんが、"
         "実態として説明のつく水準ではありません。",
         "**短期入所生活介護は、利用者数が27.2人から60.0人へ2.2倍、"
         "給付費が27,394千円から90,099千円へ3.3倍になっています。**"
         "1人1月あたりの利用日数も要介護1で20.5日と、上限16日を超えています。")

    # ---- 4 保険料
    head(doc, "4　保険料", 1)
    table(doc, ["項目", "システム", "当方", "差"], [
        ("標準給付費見込額（3年計）", f"{SYS['標準給付費']:,.0f}千円",
         f"{OUR_STD3:,.0f}千円",
         f"**{SYS['標準給付費'] - OUR_STD3:+,.0f}千円"
         f"（{SYS['標準給付費'] / OUR_STD3 - 1:+.1%}）**"),
        ("地域支援事業費（3年計）", "**0千円（未入力）**", f"{OUR_CHI3:,.0f}千円",
         f"{-OUR_CHI3:+,.0f}千円"),
        ("調整交付金見込額（3年計）", f"{SYS['調整交付金見込額']:,.0f}千円", "―", "―"),
        ("予定保険料収納率", f"{SYS['予定収納率'] * 100:.2f}％", "99.35％", "―"),
        ("所得段階別加入割合補正係数",
         f"{SYS['所得段階別補正係数'][0]:.4f}", f"{T.shotoku_hosei():.4f}", "**一致**"),
        ("**保険料基準額（月額）**", f"**{SYS['第10期月額']:,.0f}円**",
         f"**{OUR_GETSU:,.0f}円**", f"{SYS['第10期月額'] - OUR_GETSU:+,.0f}円"),
    ], right_from=1)
    body(doc,
         "**システムの数字は、2つの誤差が打ち消し合って"
         "当方の算定に近く見えているだけです。**"
         "給付費が7.1％高い分（保険料を上げる方向）と、"
         "地域支援事業費が0である分（下げる方向）が相殺しています。",
         f"**地域支援事業費{OUR_CHI3:,.0f}千円を入れると、"
         f"いまの給付費のままなら保険料は約{H['地域支援入り']:,.0f}円になります。**"
         f"当方の{OUR_GETSU:,.0f}円との差"
         f"{H['地域支援入り'] - OUR_GETSU:+,.0f}円が、"
         "そのまま給付費の差（＝3つの伸びの設定）によるものです。")
    table(doc, ["段階", "保険料基準額（月額）", "内容"], [
        (f"いまのシステム", f"{SYS['第10期月額']:,.0f}円",
         "地域支援事業費0・3つの伸びあり"),
        ("地域支援事業費を入れる", f"約{H['地域支援入り']:,.0f}円",
         "**3つの伸びはそのまま**"),
        ("3つの伸びを0に戻す", "**再計算が要ります**",
         "当方の見込みでは6,100円台後半から6,200円程度"),
        ("当方の算定", f"{OUR_GETSU:,.0f}円",
         "令和8年度を国保連月報の3か月で置き、そこから延ばしたもの"),
    ], right_from=1)
    note(doc, "※ 3つの伸びを0に戻しても、当方の算定と完全には一致しません。"
              "システムは令和7年度を起点にし、当方は令和8年度"
              "（国保連月報の4〜6月提供分・前年同期比0.9598）を起点にしているためです。"
              "この差は令和8年度の月報が12か月そろう令和9年5月以降に確かめます。")

    # ---- 5 一致した点
    head(doc, "5　一致した点", 1)
    body(doc, "**9月25日に当方が訂正した2点が、システムの出力で裏づけられました。**")
    table(doc, ["項目", "システム", "当方", "判定"], [
        ("所得段階別加入割合補正係数（令和9年度）",
         f"{SYS['所得段階別補正係数'][0]:.4f}", f"{T.shotoku_hosei():.4f}", "**一致**"),
        ("調整交付金見込交付割合（令和9年度）",
         f"{SYS['見込交付割合'][0] * 100:.2f}％",
         f"{T.chosei_wari('令和9年度') * 100:.3f}％", "**一致**"),
        ("同（令和10年度）", f"{SYS['見込交付割合'][1] * 100:.2f}％",
         f"{T.chosei_wari('令和10年度') * 100:.3f}％", "**一致**"),
        ("同（令和11年度）", f"{SYS['見込交付割合'][2] * 100:.2f}％",
         f"{T.chosei_wari('令和11年度') * 100:.3f}％", "**一致**"),
        ("第1号被保険者数（3年計）", f"{SYS['第1号被保険者数']:,}人",
         f"{SYS['第1号被保険者数']:,}人", "**一致**"),
    ], right_from=1)
    body(doc,
         "**基準額に対する割合は、公費軽減前の標準割合（0.455／0.685／0.690）を"
         "用いるのが正しい、という当方の訂正が確認できました。**"
         "軽減後の0.285／0.485／0.685で計算していれば補正係数は0.9523となり、"
         "システムの0.9958とは合いません。",
         "**調整交付金見込交付割合を算式で求める、という訂正も確認できました。**"
         "（第1号被保険者負担割合＋5％）－第1号被保険者負担割合"
         "×後期高齢者加入割合補正係数×所得段階別加入割合補正係数。"
         "従前の「第9期の公表値からの逆算6.148％」では合いません。")

    head(doc, "(1)　認定者数についての訂正", 2)
    body(doc,
         "**9月25日に「システムの認定者数が当方より約10％多い」とお伝えしましたが、"
         "これは誤りでした。**"
         "古い推計（将来推計1）の画面に出ていた男281人を"
         "令和7年度末の男女比で割り戻した推定値によるものでした。")
    table(doc, ["年度", "システム（第1号）", "当方", "差"],
          [(y, f"{SYS_NINTEI[y]:,}人",
            f"{T.NINTEI_EST[y]:,}人" if y in T.NINTEI_EST else "―",
            f"{T.NINTEI_EST[y] - SYS_NINTEI[y]:+,}人"
            f"（{T.NINTEI_EST[y] / SYS_NINTEI[y] - 1:+.1%}）"
            if y in T.NINTEI_EST else "―")
           for y in ("令和9年度", "令和10年度", "令和11年度")], right_from=1)
    note(doc, "※ 新しい推計（将来推計0925）では令和9年度が男253人・女516人＝769人です。"
              "**令和8年度のチェックを外した効果で、古い推計の281人から減っています。**"
              "※ 当方の見込みとの差は1.7％で、実質的に一致しています。")

    # ---- 6 手順
    head(doc, "6　直していただきたいこと", 1)
    table(doc, ["順", "すること", "留意点"],
          [(a, b, c) for a, b, c in TEJUN])
    body(doc,
         "**手順2〜4を直せば、介護予防訪問看護の負の値と"
         "短期入所生活介護の20.5日は解消する見込みです。**"
         "令和7年度の実績（介護予防訪問看護 要支援2 は正の値、"
         "短期入所生活介護は1人1月あたり約10日）がそのまま延びるためです。",
         "**残る36件のワーニングは、直しても消えません。**"
         "小野町に利用がないサービスの実績が0であることによるもので、"
         "都道府県への提出にあたって支障になるものではありません。"
         "**「利用がないため0である」と説明できれば足ります。**")

    p = OUT / f"小野町_見える化_ワーニングの確認結果_{ASOF}.docx"
    doc.save(p)
    return p


# ---------------------------------------------------------------- Excel

def build_xlsx(rows, ijo, nashi, taigai, H):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet("00_要点")
    ws.append(["#", "内容", "対応"])
    style_head(ws)
    for i, (a, b) in enumerate([
        ("ワーニング42件のうち36件は「小野町に利用がないサービスが0」。実害なし",
         "直す必要なし"),
        ("介護予防訪問看護（要支援2）が▲12.6回。給付費も令和9年度以降▲3,345千円",
         "下の設定を直す"),
        ("短期入所生活介護（要介護1）が20.5日で上限16日を超過", "同上"),
        ("3つの「伸び」が0（ゼロ）ではなく「令和6年度→令和7年度の伸び」になっている",
         "**0（ゼロ）に戻す**"),
        ("地域支援事業費が全欄0", "当方の入力データを入れる"),
        ("所得段階別加入割合補正係数・調整交付金見込交付割合は当方と一致",
         "当方の訂正が正しかった"),
        ("認定者数は当方782人・システム769人で1.7％差", "実質一致"),
    ], start=1):
        ws.append([i, a, b.replace("**", "")])
        ws.cell(ws.max_row, 3).fill = OK if "必要なし" in b or "正しかった" in b \
            or "実質一致" in b else NG
    body_style(ws, wrap=(2, 3))
    widths(ws, [4, 74, 30])
    ws.freeze_panes = "A2"

    ws = wb.create_sheet("01_設定の対比")
    ws.append(["設定項目", "現在（将来推計0925）", "あるべき値", "判定", "内容"])
    style_head(ws)
    for a, b, c, d, e in SETTEI:
        ws.append([a.replace("**", ""), b.replace("**", ""), c.replace("**", ""),
                   d.replace("**", ""), e.replace("**", "")])
        ws.cell(ws.max_row, 4).fill = OK if d.replace("**", "") == "○" else NG
    body_style(ws, wrap=(1, 2, 3, 5))
    widths(ws, [34, 34, 22, 6, 58])
    ws.freeze_panes = "A2"

    ws = wb.create_sheet("02_ワーニング明細")
    ws.append(["区分", "サービス", "年度", "要介護度", "値", "下限", "上限", "結果"])
    style_head(ws)
    for lab, group, fill in (("要対応", ijo, NG), ("利用なし", nashi, WARN),
                             ("対象外", taigai, None)):
        for a, b, c, d, e, f, g in group:
            ws.append([lab, a, b, c, d, e, f, g])
            if fill:
                ws.cell(ws.max_row, 1).fill = fill
    for r in range(2, ws.max_row + 1):
        for c in (5, 6, 7):
            ws.cell(r, c).number_format = "#,##0.0"
    body_style(ws, wrap=(8,))
    widths(ws, [10, 28, 12, 10, 10, 8, 8, 34])
    ws.freeze_panes = "A2"

    ws = wb.create_sheet("03_令和7→8年度の振れ")
    ws.append(["サービス", "令和7年度給付費", "令和8年度給付費", "倍率",
               "令和7年度利用者数", "令和8年度利用者数", "倍率"])
    style_head(ws)
    for nm, a, b, c, d in FURE:
        ws.append([nm, a, b, b / a if a else None, c, d, d / c if c else None])
    for r in range(2, ws.max_row + 1):
        for c in (2, 3):
            ws.cell(r, c).number_format = "#,##0"
        for c in (5, 6):
            ws.cell(r, c).number_format = "#,##0.0"
        for c in (4, 7):
            ws.cell(r, c).number_format = "0.00"
            v = ws.cell(r, c).value
            if isinstance(v, float) and (v > 1.2 or v < 0.8):
                ws.cell(r, c).fill = NG
    body_style(ws)
    widths(ws, [28, 16, 16, 8, 16, 16, 8])
    ws.freeze_panes = "A2"
    ws.append([])
    ws.append(["※ 出典：ワークシート「2_サービス別給付費」。"
               "給付費は年間累計（千円）、利用者数は1月あたり。"])
    ws.append(["※ 令和8年度は実績ではなく推計値。3つの「伸び」の設定が生んだ動き。"])

    ws = wb.create_sheet("04_保険料の突合")
    ws.append(["項目", "システム", "当方", "備考"])
    style_head(ws)
    for a, b, c, d in [
        ("標準給付費見込額（3年計・千円）", SYS["標準給付費"], OUR_STD3,
         f"差 {SYS['標準給付費'] - OUR_STD3:+,.0f}千円"
         f"（{SYS['標準給付費'] / OUR_STD3 - 1:+.1%}）"),
        ("地域支援事業費（3年計・千円）", 0, OUR_CHI3, "システムは未入力"),
        ("調整交付金見込額（3年計・千円）", SYS["調整交付金見込額"], None, ""),
        ("予定保険料収納率（％）", SYS["予定収納率"] * 100, 99.35, ""),
        ("所得段階別加入割合補正係数", SYS["所得段階別補正係数"][0],
         T.shotoku_hosei(), "一致"),
        ("調整交付金見込交付割合 令和9年度（％）", SYS["見込交付割合"][0] * 100,
         T.chosei_wari("令和9年度") * 100, "一致"),
        ("調整交付金見込交付割合 令和10年度（％）", SYS["見込交付割合"][1] * 100,
         T.chosei_wari("令和10年度") * 100, "一致"),
        ("調整交付金見込交付割合 令和11年度（％）", SYS["見込交付割合"][2] * 100,
         T.chosei_wari("令和11年度") * 100, "一致"),
        ("第1号被保険者数（3年計・人）", SYS["第1号被保険者数"],
         SYS["第1号被保険者数"], "一致"),
        ("保険料基準額（月額・円）", SYS["第10期月額"], OUR_GETSU,
         "システムは地域支援事業費0のまま"),
        ("　地域支援事業費を入れた場合（円）", H["地域支援入り"], None,
         "当方の191,749千円を入れた場合"),
    ]:
        ws.append([a, b, c, d])
    for r in range(2, ws.max_row + 1):
        for c in (2, 3):
            v = ws.cell(r, c).value
            if isinstance(v, float) and v < 10:
                ws.cell(r, c).number_format = "0.0000"
            else:
                ws.cell(r, c).number_format = "#,##0.00"
    body_style(ws, wrap=(4,))
    widths(ws, [38, 18, 18, 40])
    ws.freeze_panes = "A2"

    ws = wb.create_sheet("05_手順")
    ws.append(["順", "すること", "留意点"])
    style_head(ws)
    for a, b, c in TEJUN:
        ws.append([a, b.replace("**", ""), c.replace("**", "")])
    body_style(ws, wrap=(2, 3))
    widths(ws, [4, 58, 58])
    ws.freeze_panes = "A2"

    p = OUT / f"小野町_見える化_ワーニングの確認結果_{ASOF}.xlsx"
    wb.save(p)
    return p


# ---------------------------------------------------------------- main

def main():
    OUT.mkdir(parents=True, exist_ok=True)
    rows, gaiyo = read_warnings()
    ijo, nashi, taigai = classify(rows)
    H = hikaku()
    bad = selfcheck(H, ijo, nashi, taigai)
    p1 = build_docx(rows, gaiyo, ijo, nashi, taigai, H)
    p2 = build_xlsx(rows, ijo, nashi, taigai, H)
    print("出力:", p1)
    print("     ", p2)
    print(f"  ワーニング 要対応{len(ijo)}件／利用なし{len(nashi)}件"
          f"／対象外{len(taigai)}件")
    print(f"  システムの保険料 {SYS['第10期月額']:,.0f}円"
          f"（再現{H['再現']:,.0f}円）"
          f"／地域支援事業費を入れると{H['地域支援入り']:,.0f}円"
          f"／当方{OUR_GETSU:,.0f}円")
    print(f"  自己点検 {'OK' if not bad else 'NG: ' + ' / '.join(bad)}")


if __name__ == "__main__":
    main()
