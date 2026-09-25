"""小野町 未確認事項を見える化システムでどう解消するかを整理する。

令和8年9月25日のご指示
  「未確認事項の解消方法として見える化システムから必要な内容を
    再度ご教示ください。」

未確認事項を3つに分ける。

  A 見える化システムから取れる（依頼済みだが未取得）
      指標IDを特定し、出力を依頼すれば解消する。
  B 見える化システムでは取れない
      指標そのものがない、又は年度が古い。町の資料を要する。
  C 見える化システム上で自前で確かめられる
      受領していない資料の代わりに、システムを操作して判定する。
      推計手法の妥当性（厚生労働省の保険者別配布シートの代替）と、
      第9期策定時の設定がこれにあたる。

出力
  11_見える化出力依頼/小野町_未確認事項の解消方法_見える化システム_YYYYMMDD.docx
  11_見える化出力依頼/小野町_見える化システム追加出力依頼_YYYYMMDD.xlsx
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

ROOT = pathlib.Path(__file__).parent
BASE = ROOT / "小野町_引継ぎ_整理済"
OUT = BASE / "11_見える化出力依頼"
MASTER = OUT / "見える化_指標リスト_目的別.xlsx"
REQ = OUT / "小野町_見える化システム出力依頼リスト_20260804.xlsx"
ASOF = "20260925"
ASOF_JP = "令和8年9月25日"
JP_MIN = "游明朝"
JP_GO = "游ゴシック"

HEAD = PatternFill("solid", fgColor="1F3864")
NG = PatternFill("solid", fgColor="FCE4E4")
WARN = PatternFill("solid", fgColor="FFE0B2")
OK = PatternFill("solid", fgColor="E2EFDA")
CALC = PatternFill("solid", fgColor="EAF1FB")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# 令和7年度の総給付費（年報 様式2）。推計手法の判定の突合先。
R7_KYUFU = 1005144


# ---------------------------------------------------------------- 読み取り

def master():
    """指標リスト（目的別）を {指標ID: (大分類, 中分類, 指標名)} で返す。"""
    wb = openpyxl.load_workbook(MASTER, data_only=True)
    ws = wb["指標リスト"]
    rows = list(ws.iter_rows(values_only=True))
    h = next(i for i, r in enumerate(rows) if r and r[0] == "目的")
    out, cur = {}, ["", "", "", "", ""]
    for r in rows[h + 1:]:
        if not r or all(x is None for x in r):
            continue
        v = [("" if x is None else str(x).strip()) for x in r[:7]]
        for j in range(5):
            if v[j]:
                cur[j] = v[j]
        if v[6]:
            out[v[6]] = (cur[2], cur[3], v[5])
    return out


def jokyo():
    """既存の依頼リストから、依頼済み・取得済みの指標IDを返す。"""
    wb = openpyxl.load_workbook(REQ, data_only=True)
    r = wb["01_依頼指標一覧"]
    g = wb["02_取得済み指標"]
    req = {str(r.cell(i, 3).value) for i in range(2, r.max_row + 1)
           if r.cell(i, 3).value}
    got = {str(g.cell(i, 1).value) for i in range(2, g.max_row + 1)
           if g.cell(i, 1).value}
    return req, got


# 未確認事項ごとに、解消に要する指標のかたまりを定める。
# (区分, 未確認事項, 指標IDの範囲を決める関数, 何が分かるか)
def _rng(pref, lo, hi):
    return lambda M: [k for k in M
                      if k.startswith(pref) and k[len(pref):].isdigit()
                      and lo <= int(k[len(pref):]) <= hi]


BLOCK_A = [
    ("S8", "小野町の交付金の得点（令和6年度）", _rng("W", 126, 137),
     "**総合得点・交付金別・目標別・指標群別の得点がそのまま取れる。**"
     "W128とW130は（ⅰ）体制・取組と（ⅱ）活動の別であり、"
     "計画に書けば取れる点と実績の積上げを要する点を分けられる"),
    ("S8", "成果指標群（目標Ⅳ）の状況", _rng("W", 138, 145),
     "長期的な要介護度の変化（要介護1・2／3〜5）、短期平均要介護度の伸び、"
     "健康寿命延伸の状況。**第4章の成果指標に置くかの判断材料になる**"),
    ("S8", "活動指標群の状況", _rng("W", 146, 155),
     "認知症サポーター数・ステップアップ講座修了者数、"
     "介護人材の定着・資質向上に関する研修の実施状況ほか。"
     "**成果指標の現状値にそのまま使える**"),
    ("S8", "個別の評価項目の該当状況", _rng("W", 1, 125),
     "**市町村評価指標の項目ごとの該当状況。**"
     "全国の多くが得点していて小野町が0点の項目を洗い出せる。"
     "ただし令和4・5年度の体系によるものが66件含まれる"),
    ("C7", "地域包括支援センターの体制と地域ケア会議", _rng("F", 15, 25),
     "設置数・人員体制（3職種別）・地域ケア会議の開催回数。"
     "**第5章の進行管理と、交付金の評価項目の両方に効く**"),
    ("C7", "生活支援体制整備事業の状況", _rng("F", 26, 27),
     "生活支援コーディネーターと協議体。"
     "包括的支援事業（社会保障充実分）の実施状況"),
    ("C8", "認知症関連の研修の受講状況", _rng("J", 1, 15),
     "認知症サポート医・かかりつけ医・病院勤務者・歯科医師・薬剤師・"
     "看護職員の対応力向上研修、認知症介護実践者等養成事業。"
     "**第4章③の現状の記述に使える**"),
]

BLOCK_B = [
    ("包括的支援事業・任意事業費29,335,189円の8欄内訳",
     "**D48-c は歳出の「地域支援事業」の総額のみで、内訳はない。**"
     "見える化に包括的支援事業の費用の内訳を示す指標はない",
     "町の令和6年度決算書（歳出内訳）。依頼票で照会済み"),
    ("同費に交付金を財源とする事業が含まれるか",
     "同上。費用の財源別の区分は見える化にない",
     "**町の決算書。含まれていれば第1号被保険者負担分の算定基礎から除く。"
     "含まれたままなら保険料が過大に出る**"),
    ("一般介護予防事業費5,304,726円の5事業別内訳",
     "同上",
     "町の事業実績。F28〜F40は件数の指数であって費用ではない"),
    ("通いの場の令和3年度以降の実績",
     "**見える化の通いの場（F1〜F14）は令和2年度が最新である。**"
     "週1回以上は平成27年度の3か所・56人を最後に0が続き、"
     "月1回以上は令和1・2年度が40〜42人（1.2％）",
     "**町の事業実績。令和7年度調査では週1回以上の参加が5.5％"
     "（高齢者人口換算で約190人）あり、見える化と大きく食い違う。**"
     "量の見込みを置く前に実態の確認が要る"),
    ("チームオレンジの整備状況",
     "**指標リスト992件にチームオレンジの指標はない**",
     "町の事業実績。交付金の評価項目としては W 分野に現れる"),
    ("令和7年8〜9月の認定者数の計上方法",
     "見える化の認定者数は各年9月末の値であり、"
     "断層そのものは見えるが原因は分からない",
     "町の認定データ。依頼票で照会済み"),
]

# C 見える化システム上で自前で確かめる手順
STEP_SUIKEI = [
    ("1", "推計パターンを2つ新規に作る",
     "**既存のパターンを上書きしない。**"
     "名称は「手法検証_基本推計」「手法検証_包括推計」など区別のつくものにする"),
    ("2", "介護保険事業状況報告の設定で、令和7年度と令和8年度のチェックを外す",
     "**これにより令和6年度が最新の完結年度＝基準年度になる。**"
     "令和8年度は1か月分であり、外さないと基準年度が汚れる"),
    ("3", "一方を「性別／年齢5歳階級別／要介護度別に推計する（基本推計）」にする",
     "推計方法の設定画面"),
    ("4", "他方を包括推計にする", "同上"),
    ("5", "施策反映の3画面は編集せず通過する",
     "**編集するとセルがピンクになり「自然体推計に全て戻す」が押せる状態になる。**"
     "その状態で総括表を出すと推計が固定され、比較にならない"),
    ("6", "それぞれ将来推計総括表を出力する", "令和7年度の推計値が得られる"),
    ("7", "令和7年度の推計値を年報 様式2 の実績と突き合わせる",
     f"**実績は総給付費{R7_KYUFU:,}千円。**"
     "推計値÷実績値が1に近いほうが小野町で当たる手法である"),
]

STEP_DAI9 = [
    ("1", "将来推計の推計パターンの一覧を開く",
     "**第9期（令和5年度策定）のパターンが残っていれば、これだけで解消する**"),
    ("2", "第9期のパターンの「推計方法の設定」画面を開く",
     "基準年度・推計手法（基本推計／包括推計）が分かる"),
    ("3", "「介護保険事業状況報告の設定」画面を開く",
     "**どの年度にチェックが入っていたかが分かる。**"
     "第9期の基準年度が何か月分であったかはここで決まる"),
    ("4", "認定率・利用率の伸びの窓の設定を確認する",
     "どの年度からどの年度への伸びを採っていたか"),
    ("5", "残っていない場合は町の起案文書による",
     "**それも無ければ、第9期の見込みが令和5年度実績より7.3％高いこと"
     "（見込量算定の再検証 第2章）を推定の根拠とするにとどまる**"),
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


# ---------------------------------------------------------------- Word

def build_docx(M, req, got, A):
    n_a = sum(len(ids) for _c, _n, ids, _m in A)
    mi = sum(1 for _c, _n, ids, _m in A for k in ids if k not in got)
    doc = new_doc()
    head(doc, "未確認事項の解消方法 ― 見える化システムから取得すべき内容", 0)
    body(doc, f"小野町　第10期介護保険事業計画　／　{ASOF_JP}")

    head(doc, "0　結論", 1)
    table(doc, ["区分", "内容"], [
        ("**A 見える化から取れる**",
         f"**{n_a}件の指標で7つの未確認事項が解消する。**"
         f"うち{mi}件は令和8年8月4日の依頼リストに載せたまま未取得である。"
         "**とくにW126〜W137は小野町の交付金の得点そのもので、"
         "町の得点票を待たずに分析に入れる**"),
        ("**B 見える化では取れない**",
         "**6件。**費用の内訳（包括的支援事業・任意事業費の8欄、"
         "一般介護予防事業費の5事業）は見える化に指標がなく、"
         "通いの場は令和2年度が最新である。いずれも町の資料を要する"),
        ("**C 見える化で自前で確かめる**",
         "**2件。**推計手法（基本推計／包括推計）の妥当性は、"
         "厚生労働省の保険者別配布シートがなくても"
         "システムを操作して判定できる。"
         "第9期策定時の設定も、推計パターンが残っていれば読み取れる"),
    ])

    head(doc, "1　A　見える化システムから取れるもの", 1)
    body(doc,
         "**下表の指標を出力していただければ解消します。**"
         "いずれも令和8年8月4日の出力依頼リストに含まれていますが、"
         "現時点で取得できていません。")
    def _rangetxt(ids):
        v = sorted(int(k[1:]) for k in ids)
        pre = ids[0][0]
        if len(v) == 1:
            return f"**{pre}{v[0]}**"
        # 連番が欠けている場合は「のうち○件」と書き、範囲だけを示さない
        tobi = "" if len(v) == v[-1] - v[0] + 1 else f"　のうち{len(v)}件"
        return f"**{pre}{v[0]}〜{pre}{v[-1]}**{tobi}"

    table(doc, ["未確認事項", "指標ID", "件数", "何が分かるか"],
          [(nm, _rangetxt(ids), f"{len(ids)}", memo)
           for _c, nm, ids, memo in A], right_from=2)
    note(doc, "※ 指標IDは見える化システムの指標リスト（目的別・992指標）による。"
              "※ W1〜W125には令和4年度・令和5年度の体系による項目が66件含まれる。"
              "令和6年度の体系（8目標・800点）と対応しないため、"
              "**得点の比較にはW126〜W137を用いる。**")

    head(doc, "2　とくに優先していただきたいもの ― W126〜W137", 1)
    body(doc,
         "**これは小野町の交付金の得点そのものです。**"
         "令和8年9月25日に受領した公表資料は都道府県分のみで、"
         "小野町自身の得点分析ができていませんでしたが、"
         "**この12指標で分析に入れます。**")
    table(doc, ["指標ID", "指標名"],
          [(k, M[k][2]) for k in sorted(
              (k for k in M if k.startswith("W") and k[1:].isdigit()
               and 126 <= int(k[1:]) <= 137), key=lambda x: int(x[1:]))])
    body(doc,
         "**W128とW130（指標群別の得点）がとくに重要です。**"
         "各目標は（ⅰ）体制・取組と（ⅱ）活動に分かれ、"
         "（ⅰ）は計画に書けば翌年度の取組として得点になるのに対し、"
         "（ⅱ）は実績の積上げを要します。"
         "**この2つを分けられれば、新たな事業も予算も要さず"
         "計画への記載だけで取れる項目を切り出せます。**")

    head(doc, "3　B　見える化システムでは取れないもの", 1)
    table(doc, ["未確認事項", "見える化に無い理由", "解消の方法"],
          [(nm, why, how) for nm, why, how in BLOCK_B])

    head(doc, "4　C　見える化システム上で自前で確かめるもの", 1)
    head(doc, "4-1　推計手法（基本推計／包括推計）はどちらが当たるか", 2)
    body(doc,
         "**厚生労働省の保険者別配布シート（標準の推計方法を当てはめた誤差）は"
         "未受領ですが、見える化システムを操作すれば小野町で自前に判定できます。**"
         "完結年度を1つ手前に戻して推計し、実績と突き合わせる方法です。")
    table(doc, ["順", "すること", "留意点"],
          [(no, what, memo) for no, what, memo in STEP_SUIKEI])
    body(doc,
         "**推計値÷実績値が1に近いほうが小野町で当たる手法です。**"
         "他町村の案件では、この判定により基本推計が包括推計を"
         "大きく上回ることを確認して手法を確定しています。"
         "**小野町の総括表の現在の設定が基本推計であるかどうかも、"
         "あわせてご確認ください。**")
    note(doc, "※ 配布シートが受領できれば、全国1,500超の保険者との比較ができるため"
              "そちらが優る。本手順はその代替であり、"
              "小野町の1保険者の中での比較にとどまる。")

    head(doc, "4-2　第9期策定時の設定", 2)
    body(doc,
         "**第9期の見込みは、令和5年度の実績より7.3％高いところにありました**"
         "（見込量算定の再検証 第2章）。"
         "実績／見込みが令和6年度98.6％・令和7年度91.1％となった原因が"
         "基準年度の置き方にあるかを確定するには、当時の設定が要ります。")
    table(doc, ["順", "すること", "留意点"],
          [(no, what, memo) for no, what, memo in STEP_DAI9])

    head(doc, "5　あわせてご確認いただきたい画面", 1)
    table(doc, ["画面", "見るところ", "なぜ"], [
        ("将来推計 ＞ 推計方法の設定",
         "**推計手法（基本推計／包括推計）**",
         "第10期の総括表がどちらで出ているかが未確認"),
        ("将来推計 ＞ 介護保険事業状況報告の設定",
         "**令和8年度にチェックが入っているか**",
         "**入っていれば1か月分が年度として扱われる。**"
         "当方の判定（利用者数の非ゼロ26件がすべて整数）の裏づけになる"),
        ("将来推計 ＞ 施策反映（3画面）",
         "**「自然体推計に全て戻す」が押せる状態になっていないか。"
         "セルがピンクになっていないか**",
         "**押せる状態のまま総括表を出すと、伸びの窓を変えても"
         "見込量が動かない。**他町村の案件で実際に起きている"),
        ("将来推計 ＞ 推計パターンの一覧",
         "第9期のパターンが残っているか",
         "4-2のとおり"),
    ])

    head(doc, "6　解消後にできるようになること", 1)
    table(doc, ["解消する事項", "できるようになること"], [
        ("交付金の得点（W126〜W137）",
         "**全国の多くの市町村が得点していて小野町が0点の項目を洗い出し、"
         "取組の水準が低いのか事業に着手していないのかを分けられる。**"
         "協議会資料（08）の分析を都道府県分から小野町分に差し替える"),
        ("活動指標群（W146〜W155）",
         "**第4章の成果指標42件のうち、現状値が未受領の指標に値が入る。**"
         "認知症サポーター数・ステップアップ講座修了者数がこれにあたる"),
        ("推計手法の判定（4-1）",
         "**見込量算定の再検証で唯一検証できていない項目が埋まる。**"
         "手引きの14段階のうち、推計手法の選択の妥当性が確認済みになる"),
        ("第9期の設定（4-2）",
         "**第9期の乖離の原因が推定から確定に変わる。**"
         "第10期で同じ置き方を繰り返さないことの根拠が固まる"),
        ("包括的支援事業・任意事業費の内訳（B）",
         "**交付金を財源とする事業を算定基礎から除ける。"
         "含まれたままなら保険料が過大に出る**"),
    ])

    OUT.mkdir(parents=True, exist_ok=True)
    p = OUT / f"小野町_未確認事項の解消方法_見える化システム_{ASOF}.docx"
    doc.save(p)
    print("出力:", p)


# ---------------------------------------------------------------- Excel

def style_head(ws, row):
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(row, c)
        cell.fill = HEAD
        cell.font = Font(bold=True, color="FFFFFF", size=9, name=JP_GO)
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
        cell.border = BORDER


def xbody(ws, first, wrap=()):
    for r in range(first + 1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            cell.font = Font(size=9, name=JP_MIN)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=c in wrap)


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w


def notes(ws, lines):
    ws.append([])
    for t in lines:
        ws.append([t])
        ws.cell(ws.max_row, 1).font = Font(size=9, italic=True, name=JP_MIN)


def build_xlsx(M, req, got, A):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet("00_要点")
    ws.append([f"小野町 第10期　見える化システム 追加出力依頼　（{ASOF_JP}）"])
    ws.cell(1, 1).font = Font(bold=True, size=13, name=JP_GO)
    ws.append(["未確認事項を、見える化システムで解消できるもの（A）・"
               "できないもの（B）・システム上で自前に確かめるもの（C）に分けた。"])
    ws.append([])
    ws.append(["区分", "未確認事項", "指標ID", "件数", "依頼済", "取得済",
               "何が分かるか"])
    hr = ws.max_row
    for c, nm, ids, memo in A:
        ids = sorted(ids, key=lambda x: int(x[1:]))
        ws.append([c, nm,
                   f"{ids[0]}〜{ids[-1]}" if len(ids) > 1 else ids[0],
                   len(ids),
                   sum(1 for k in ids if k in req),
                   sum(1 for k in ids if k in got),
                   memo.replace("**", "")])
    style_head(ws, hr)
    for r in range(hr + 1, ws.max_row + 1):
        ws.cell(r, 6).fill = NG if ws.cell(r, 6).value == 0 else OK
    xbody(ws, hr, wrap=(2, 7))
    widths(ws, [8, 34, 14, 8, 8, 8, 60])
    notes(ws, [
        "**※ 取得済0件の行は、依頼リスト（令和8年8月4日）に載せたまま"
        "出力されていないものである。**",
        "**※ W126〜W137は小野町の交付金の得点そのもの。**"
        "町の得点票を待たずに分析に入れる。",
    ])

    ws = wb.create_sheet("01_指標の明細")
    ws.append(["区分", "未確認事項", "指標ID", "大分類", "中分類", "指標名",
               "依頼済", "取得済"])
    hr = ws.max_row
    for c, nm, ids, _memo in A:
        for k in sorted(ids, key=lambda x: int(x[1:])):
            d = M.get(k, ("", "", ""))
            ws.append([c, nm, k, d[0], d[1], d[2],
                       "○" if k in req else "", "○" if k in got else ""])
    style_head(ws, hr)
    for r in range(hr + 1, ws.max_row + 1):
        ws.cell(r, 8).fill = OK if ws.cell(r, 8).value == "○" else NG
    xbody(ws, hr, wrap=(2, 5, 6))
    widths(ws, [8, 30, 10, 26, 26, 54, 8, 8])
    ws.freeze_panes = "A2"

    ws = wb.create_sheet("02_見える化では取れない")
    ws.append(["未確認事項", "見える化に無い理由", "解消の方法"])
    hr = ws.max_row
    for nm, why, how in BLOCK_B:
        ws.append([nm, why.replace("**", ""), how.replace("**", "")])
    style_head(ws, hr)
    for r in range(hr + 1, ws.max_row + 1):
        ws.cell(r, 1).fill = WARN
    xbody(ws, hr, wrap=(1, 2, 3))
    widths(ws, [34, 52, 58])

    ws = wb.create_sheet("03_推計手法の自前判定")
    ws.append(["見える化システム上で、基本推計と包括推計のどちらが当たるかを判定する"])
    ws.cell(1, 1).font = Font(bold=True, size=12, name=JP_GO)
    ws.append([])
    ws.append(["順", "すること", "留意点"])
    hr = ws.max_row
    for no, what, memo in STEP_SUIKEI:
        ws.append([no, what, memo.replace("**", "")])
    style_head(ws, hr)
    xbody(ws, hr, wrap=(2, 3))
    widths(ws, [6, 46, 62])
    notes(ws, [
        f"**※ 突合先は年報 様式2 の令和7年度 総給付費{R7_KYUFU:,}千円。**"
        "推計値÷実績値が1に近いほうが小野町で当たる手法である。",
        "※ 厚生労働省の保険者別配布シートが受領できればそちらが優る"
        "（全国1,500超の保険者との比較ができる）。本手順はその代替である。",
    ])

    ws = wb.create_sheet("04_第9期の設定の確認")
    ws.append(["第9期策定時の見える化システムの設定を読み取る"])
    ws.cell(1, 1).font = Font(bold=True, size=12, name=JP_GO)
    ws.append([])
    ws.append(["順", "すること", "留意点"])
    hr = ws.max_row
    for no, what, memo in STEP_DAI9:
        ws.append([no, what, memo.replace("**", "")])
    style_head(ws, hr)
    xbody(ws, hr, wrap=(2, 3))
    widths(ws, [6, 46, 62])
    notes(ws, [
        "**※ 第9期の見込みは3か年の振れ幅が0.42％しかなく自然体推計そのもので、"
        "その水準は令和5年度の実績より7.3％高い。**"
        "実績／見込みは令和6年度98.6％・令和7年度91.1％（見込量算定の再検証 第2章）。",
    ])

    p = OUT / f"小野町_見える化システム追加出力依頼_{ASOF}.xlsx"
    wb.save(p)
    print("出力:", p)


def main():
    M = master()
    req, got = jokyo()
    A = [(c, nm, fn(M), memo) for c, nm, fn, memo in BLOCK_A]
    A = [(c, nm, ids, memo) for c, nm, ids, memo in A if ids]
    OUT.mkdir(parents=True, exist_ok=True)
    build_docx(M, req, got, A)
    build_xlsx(M, req, got, A)
    n = sum(len(ids) for _c, _n, ids, _m in A)
    g = sum(1 for _c, _n, ids, _m in A for k in ids if k in got)
    print(f"  A（見える化から取れる）{n}件のうち取得済み{g}件／"
          f"B（取れない）{len(BLOCK_B)}件／C（自前で確かめる）2件")


if __name__ == "__main__":
    main()
