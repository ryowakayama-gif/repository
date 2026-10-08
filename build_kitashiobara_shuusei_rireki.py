"""計画素案　元ファイルとの差分と修正履歴（マーカー強調つき）ジェネレータ

ご指示（令和8年10月8日）
  「計画素案の修正が終わった段階で元ファイルとの差分について修正履歴を
   ご教示ください。特に国・県の指針と不整合があった箇所、前回計画との差分で
   KPIや目標設定、課題、見込量について修正した箇所についてはマーカーにて
   強調表示して下さい」

元ファイル
  source/北塩原村_骨子案_原本_20260731.docx
  他メンバーが作成した令和8年7月31日版。当方は一切書き換えていない。
  計画素案は、この原本を入力として build_kitashiobara_kosshi_rev.py が
  作り直すため、原本と素案を機械で突き合わせれば差分が出る。

マーカーの色
  黄　国・県の指針と不整合があった箇所（分類A・B）
  緑　KPI・目標設定・課題・見込量を修正した箇所（分類C・D・E）
  なし　記載事項の新設・現況データの更新・体裁の是正（分類F・H・G）

このブックの立て方
 1. 修正履歴は素案の生成器（build_kitashiobara_kosshi_rev.py）の
    changes.append から機械で読み出す。手で書き写さないので取り違えがない。
 2. 分類は語による規則で決める。規則に当たらない項目が出たら自己点検で落ちる。
    生成器に修正を足して分類を決め忘れることがない。
 3. 章節は素案の見出しと突き合わせる。存在しない箇所を指す履歴が残らない。

出力: output/北塩原村_計画素案_修正履歴.docx
"""

import hashlib
import os
import re

import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from docx.table import Table
from docx.text.paragraph import Paragraph

REPO_ROOT = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = f"{REPO_ROOT}/output"
OUT_FILE = f"{OUT_DIR}/北塩原村_計画素案_修正履歴.docx"
GENPON = f"{REPO_ROOT}/source/北塩原村_骨子案_原本_20260731.docx"
SUAN = f"{OUT_DIR}/北塩原村_計画素案.docx"
BUILDER = f"{REPO_ROOT}/build_kitashiobara_kosshi_rev.py"

FONT = "BIZ UDPゴシック"
HEADER_FILL = "1F3864"
SUB_FILL = "2E75B6"
NOTE_FILL = "FFF2CC"
BORDER = "AAAAAA"
ASOF_JP = "令和8年10月8日"

# 分類と、マーカーの色（None はマーカーなし）
CATS = [
    ("A", "国の指針との不整合の是正", "yellow",
     "告示（令和8年3月31日 こども家庭庁・厚生労働省告示第4号）及びその改正に"
     "照らして、原本の記載が指針に合っていなかった箇所、"
     "指針が求める事項が欠けていた箇所"),
    ("B", "県の指針・県計画との整合", "yellow",
     "第7期福島県障がい福祉計画・第3期障がい児福祉計画及び"
     "第5次福島県障がい者計画に照らして、原本に記載がなかった箇所"),
    ("C", "KPI・目標設定の修正", "green",
     "成果目標の数値・倍率・指標の設定にかかる修正"),
    ("D", "課題・評価の修正", "green",
     "現行計画の評価、前回調査の数値、今期調査の結果にかかる修正"),
    ("E", "見込量・算定の修正", "green",
     "サービス見込量の値・推計方法・活動指標・確保方策にかかる修正"),
    ("F", "記載事項の新設", None,
     "原本になかった章節を、業務委託仕様書又は基本指針の記載事項として"
     "新たに設けたもの"),
    ("H", "現況データの更新", None,
     "人口・手帳所持者数等の現況データを最新の実績に更新したもの"),
    ("G", "体裁・表記・参照の是正", None,
     "表題・表記・年度表記・節番号の参照ずれ等"),
]
CAT_NAME = {c: name for c, name, _col, _d in CATS}
CAT_COLOR = {c: col for c, _name, col, _d in CATS}

# 分類を決める語の規則。上から順に当て、最初に当たったものを採る。
RULES = [
    # 国の指針との不整合の是正。指針を名指しするもの、廃止区分、基準時点の訂正
    ("A", (r"告示", r"基本指針", r"別表", r"廃止区分", r"国基準", r"基準時点",
           r"仕様書.*欠落")),
    # 県の計画との整合。県を名指しするもの
    ("B", (r"福島県", r"県計画", r"県障がい者計画", r"会津", r"県の成果目標",
           r"県が定める区域")),
    # 現況データの更新
    ("H", (r"人口データ", r"手帳所持者数", r"将来推計", r"就学状況")),
    # 体裁・表記・参照の是正。E より前に置く。「R9〜R11見込量」のような
    # 引用された旧列名で見込量の修正と取り違えないようにする
    ("G", (r"表記", r"奥付", r"表題", r"列名", r"参照ずれ", r"目次",
           r"年度表記")),
    # 見込量・算定。C より前に置く。「成果目標による補正」のような
    # 算定方法の説明を目標設定の修正と取り違えないようにする
    ("E", (r"見込量", r"推計", r"算定", r"請求件数", r"活動指標", r"確保方策")),
    ("C", (r"目標値", r"倍", r"成果目標", r"セルフプラン率", r"ＫＰＩ", r"KPI")),
    ("D", (r"評価", r"課題", r"回収率", r"アンケート", r"推進状況", r"調査")),
    ("F", (r".",)),
]

# 分類ごとの件数。生成器に修正を足して分類が動いたら自己点検で気づく。
EXPECT = {"A": 13, "B": 8, "C": 3, "D": 5, "E": 6, "F": 9, "H": 2, "G": 8}

# f文字列の差し込みを読める形に直す（生成器の実行時の値）。
# 値は素案の現物と突き合わせる。
PLACEHOLDER = {
    "{len(SERVICE_TAIKEI)}": "30",
    "{len(YOUGO)}": "35",
    "{replaced}": "27",
    "{dropped}": "2",
    "{nm}": "就労定着支援",
    "{hit}": "8",
    "{sum(hits.values())}": "2",
}


# ============================================================
# 書式ヘルパー
# ============================================================
def _el(tag, **attrs):
    e = OxmlElement(tag if ":" in tag else f"w:{tag}")
    for k, v in attrs.items():
        e.set(qn(f"w:{k}"), str(v))
    return e


def _font(run, size=10.5, bold=False, color=None, highlight=None):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    if highlight:
        run._element.rPr.append(_el("highlight", val=highlight))
    return run


def new_doc(title, subtitle):
    doc = docx.Document()
    sec = doc.sections[0]
    sec.top_margin = Cm(2.0)
    sec.bottom_margin = Cm(1.8)
    sec.left_margin = Cm(1.8)
    sec.right_margin = Cm(1.8)
    st = doc.styles["Normal"]
    st.font.name = FONT
    st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _font(p.add_run(subtitle), size=9)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _font(p.add_run(title), size=15, bold=True)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(12)
    return doc


def h2(doc, text):
    p = doc.add_paragraph()
    _font(p.add_run(f"■　{text}"), size=12, bold=True, color=HEADER_FILL)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    return p


def h3(doc, text):
    p = doc.add_paragraph()
    _font(p.add_run(f"　{text}"), size=11, bold=True, color=SUB_FILL)
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    return p


def para(doc, text, indent=True, size=10.5):
    p = doc.add_paragraph()
    _font(p.add_run(text), size=size)
    if indent:
        p.paragraph_format.first_line_indent = Pt(size)
    p.paragraph_format.space_after = Pt(4)
    return p


def table(doc, rows, widths=None, size=9, highlights=None):
    """表を作る。highlights は行ごとのマーカー色（1行目は表頭のため None）。"""
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblpr = t._tbl.tblPr
    for old in tblpr.findall(qn("w:tblBorders")):
        tblpr.remove(old)
    b = _el("tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b.append(_el(edge, val="single", sz=4, space=0, color=BORDER))
    tblpr.append(b)

    for ri, row in enumerate(rows):
        hl = None if ri == 0 or not highlights else highlights[ri - 1]
        for ci, val in enumerate(row):
            cell = t.cell(ri, ci)
            if widths:
                tcpr = cell._tc.get_or_add_tcPr()
                for old in tcpr.findall(qn("w:tcW")):
                    tcpr.remove(old)
                tcpr.append(_el("tcW", w=widths[ci], type="dxa"))
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            lines = str(val).split("\n")
            for li, ln in enumerate(lines):
                if li:
                    _font(p.add_run(), size=size).add_break()
                _font(p.add_run(ln), size=size, bold=(ri == 0),
                      color="FFFFFF" if ri == 0 else None,
                      highlight=hl)
            if ri == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                cell._tc.get_or_add_tcPr().append(
                    _el("shd", val="clear", color="auto", fill=HEADER_FILL))
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def note(doc, text):
    t = doc.add_table(rows=1, cols=1)
    tblpr = t._tbl.tblPr
    for old in tblpr.findall(qn("w:tblBorders")):
        tblpr.remove(old)
    b = _el("tblBorders")
    for edge in ("top", "left", "bottom", "right"):
        b.append(_el(edge, val="single", sz=4, space=0, color="C9A03B"))
    tblpr.append(b)
    cell = t.cell(0, 0)
    cell._tc.get_or_add_tcPr().append(
        _el("shd", val="clear", color="auto", fill=NOTE_FILL))
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    _font(p.add_run(text), size=9.5)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


# ============================================================
# 現物の読み出し
# ============================================================
def doc_profile(path):
    """見出し・段落数・表数・バイトのSHA-256・内容の署名を返す。

    docx は ZIP であるため、同じ内容でも作り直すとバイト列が変わる。
    作り直すたびに値が動く欄を本書に載せると、本書自身が再現しなくなる。
    そのため計画素案には内容の署名（段落数・表数・本文テキストの
    SHA-256 先頭12桁。check_reproducibility.py と同じ作り方）を用い、
    バイトのSHA-256は作り直さない元ファイルにだけ用いる。
    """
    with open(path, "rb") as fh:
        digest = hashlib.sha256(fh.read()).hexdigest()
    d = docx.Document(path)
    heads = [(p.style.name, p.text.strip()) for p in d.paragraphs
             if p.style.name.startswith("Heading") and p.text.strip()]
    texts = [p.text for p in d.paragraphs]
    for t in d.tables:
        for row in t.rows:
            texts += [c.text for c in row.cells]
    sig = hashlib.sha256("\n".join(texts).encode("utf-8")).hexdigest()[:12]
    return {"sha256": digest, "sig": sig, "heads": heads,
            "paras": len(d.paragraphs), "tables": len(d.tables)}


def read_log():
    """生成器の changes.append から修正履歴を読み出す。"""
    with open(BUILDER, encoding="utf-8") as fh:
        src = fh.read()
    calls = re.findall(r"changes\.append\(\s*((?:f?\"[^\"]*\"\s*)+)\)", src)
    out = []
    for c in calls:
        txt = "".join(re.findall(r'f?"([^"]*)"', c))
        for ph, val in PLACEHOLDER.items():
            txt = txt.replace(ph, val)
        out.append(txt)
    if not out:
        raise SystemExit("生成器から修正履歴を読み出せません")
    return out


def classify(text):
    for cat, pats in RULES:
        for pat in pats:
            if re.search(pat, text):
                return cat
    return "F"


def split_entry(text):
    """「章節：内容」に分ける。先頭の管理番号（M-2・E-1 等）は分けて返す。"""
    num = ""
    m = re.match(r"^((?:[A-Z]-\d+(?:・[A-Z]-\d+)*|C-\d+)\s+)(.*)$", text)
    if m:
        num, text = m.group(1).strip(), m.group(2)
    if "：" in text:
        где, naka = text.split("：", 1)
    else:
        где, naka = "―", text
    return num, где.strip(), naka.strip()


# ============================================================
# 原本の記載を直した箇所（不整合の是正）
# 元ファイルの記載, 直した記載, 根拠, 分類
# ============================================================
SEIGO = [
    ("第4章（1）　施設入所者数の基準時点を「令和5年度末」としていた",
     "「令和4年度末（令和5年3月31日時点）」に訂正",
     "告示 第二の一は令和4年度末時点の施設入所者数を基準とする。"
     "第7期県計画も令和4年度末を基準点としている",
     "A"),
    ("第5章　サービスの内容及びサービス区分表に「医療型児童発達支援」を"
     "掲げていた",
     "削除し、≪サービスの内容≫の番号②〜⑦を①〜⑥に繰り上げ",
     "令和4年法律第66号による児童福祉法の改正により、"
     "令和6年4月1日から児童発達支援に一元化された",
     "A"),
    ("第5章3〜7　就労定着支援の前計画欄に「10人日分」と量を載せていた",
     "量の記載を削除（単位欄は人／月のみ）",
     "基本指針 別表第一 三は就労定着支援について利用者数の見込みのみを"
     "求めており、量を定めていない。単位欄と値が食い違っていた",
     "A"),
    ("第4章　≪国の基本指針≫欄の数値が第7期のものだった（8か所）",
     "令和8年3月31日告示第4号の本文と照合し、就労移行支援1.14倍・"
     "就労継続支援Ａ型1.52倍・Ｂ型1.67倍、精神病床1年以上長期入院患者数、"
     "重症心身障害児支援事業所の確保、拠点の全日常生活圏域対象化等に差し替え",
     "告示 第二の一〜八",
     "A"),
    ("第5章1　見込量の算定方法を示していなかった",
     "告示 第三の二の2㈠が「設定するものとする」とした2要件"
     "（継続入所者の控除・保育所等における障害児の受入れ体制）を含め、"
     "算定方法の根拠を表で明示",
     "告示 第三の二の2㈠。本村は別表第五 一の項⑴（全部過疎でないこと）を"
     "満たさないため同表の算定方法は適用されない",
     "A"),
    ("第5章　活動指標の節がなかった",
     "「10　活動指標」を新設（既存の10〜11を11〜12へ繰り下げ）",
     "業務委託仕様書4-Ⅱ(3)⑤及び告示 別表第一。骨子案に欠落していた項目",
     "A"),
    ("第5章　人材確保・生産性向上・経営基盤の節がなかった",
     "「13　障がい福祉人材の確保・定着、生産性向上及び経営基盤の確立」を新設"
     "（意思決定支援を含む）",
     "令和8年7月10日の基本指針の改正により別表第二に六の項として新設された"
     "記載事項",
     "A"),
    ("第5章　関係機関との連携の節がなく、就労関係機関の記載が0か所だった",
     "「14　関係機関との連携」を新設",
     "基本指針 別表第二 五㈠㈡",
     "A"),
    ("第4章（2）　長期入院患者の基盤整備量の項目がなかった",
     "項目を置き、県の算定値を勘案して定めることを明記（県の算定値待ち）",
     "基本指針 別表第二 三㈠⑤の「定めなければならない事項」",
     "A"),
    ("第1章2　県の計画は名称だけを掲げ、計画期間も関係も書いていなかった",
     "「■　福島県の計画との整合」「■　県が定める区域と本計画でいう圏域」"
     "「■　県障がい者計画の基本目標と本計画の施策体系」"
     "「■　県の成果目標との関係」の4節を新設",
     "告示 第三の一の5・7。第7期県計画 第1の2⑵・第1の4・第1の5、"
     "第5次県障がい者計画 第1章第2〜第5・第3章第1",
     "B"),
    ("第1章5(1)　前回調査の回収率を誤った分母で記載していた",
     "現行計画本編及び調査結果報告書に基づき訂正（回収率57.9％の分母は164人）",
     "前回計画 本編及び第7期計画策定に係るアンケート調査結果報告書",
     "D"),
    ("第5章冒頭・第5章7　注記が差し替え前の列名（R6〜R8実績・"
     "R9〜R11見込量）を指したままだった",
     "現在の表頭（令和9〜11年度）に是正し、年度表記を令和表記に統一",
     "―（表と注記の食い違いの是正）",
     "G"),
    ("第5章9　地域生活支援事業の表頭が「R8見込量」等の略記だった",
     "令和表記に統一（他の見込量表と同じ表記）",
     "―（表記の統一）",
     "G"),
]


# ============================================================
# 本文
# ============================================================
def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    for path in (GENPON, SUAN, BUILDER):
        if not os.path.exists(path):
            raise SystemExit(f"見つかりません: {path}")

    a = doc_profile(GENPON)
    b = doc_profile(SUAN)
    log = read_log()
    rows = [split_entry(x) + (classify(x),) for x in log]

    sa = {t for _s, t in a["heads"]}
    sb = {t for _s, t in b["heads"]}
    added = [(s, t) for s, t in b["heads"] if t not in sa]
    gone = [(s, t) for s, t in a["heads"] if t not in sb]

    doc = new_doc("北塩原村　障がい福祉計画・障がい児福祉計画　計画素案\n"
                  "元ファイルとの差分と修正履歴", ASOF_JP)

    h2(doc, "1　本書について")
    para(doc,
         "計画素案は、他メンバーが作成した骨子案の原本（令和8年7月31日版）を"
         "入力として、当方の生成器が作り直したものです。"
         "本書は、その原本と現在の計画素案を機械で突き合わせた差分と、"
         "生成器に記録された修正履歴をまとめたものです。")
    table(doc, [
        ["区分", "ファイル", "段落", "表", "見出し", "照合値"],
        ["元ファイル",
         "source/北塩原村_骨子案_原本_20260731.docx\n"
         "（他メンバー作成。当方は書き換えていない）",
         str(a["paras"]), str(a["tables"]), str(len(a["heads"])),
         f"SHA-256\n{a['sha256'][:16]}"],
        ["計画素案", "output/北塩原村_計画素案.docx",
         str(b["paras"]), str(b["tables"]), str(len(b["heads"])),
         f"内容の署名\n{b['sig']}"],
    ], [1500, 4400, 900, 700, 900, 2400])

    h3(doc, "マーカーの意味")
    para(doc, "ご指示に従い、次の分類にマーカーを付けています。", indent=False)
    hl = []
    cat_rows = [["分類", "内容", "マーカー", "件数"]]
    for cat, name, color, desc in CATS:
        n = sum(1 for r in rows if r[3] == cat)
        cat_rows.append([cat, f"{name}\n{desc}",
                         {"yellow": "黄", "green": "緑"}.get(color, "―"),
                         f"{n}件"])
        hl.append(color)
    table(doc, cat_rows, [700, 7200, 900, 900], highlights=hl)

    h2(doc, "2　分量の差分")
    para(doc,
         f"見出しは{len(a['heads'])}件から{len(b['heads'])}件へ、"
         f"表は{a['tables']}件から{b['tables']}件へ増えています。"
         f"新たに設けた章節は{len(added)}件、"
         f"見出しの文字が変わった（主に節番号の繰り下げによる）ものが"
         f"{len(gone)}件です。")

    h2(doc, "3　修正履歴")
    para(doc,
         f"生成器に記録された{len(rows)}件です。"
         "黄のマーカーが国・県の指針との不整合の是正、"
         "緑のマーカーがKPI・目標設定・課題・見込量の修正です。")
    rireki = [["#", "分類", "箇所", "修正内容", "管理番号"]]
    hl = []
    for i, (num, where, naka, cat) in enumerate(rows, 1):
        rireki.append([str(i), cat, where, naka, num or "―"])
        hl.append(CAT_COLOR[cat])
    table(doc, rireki, [500, 600, 1800, 5800, 1000], highlights=hl)

    h2(doc, "4　元ファイルの記載を直した箇所")
    para(doc,
         "修正履歴のうち、原本の記載そのものが指針等に合っていなかったもの、"
         "又は表と注記が食い違っていたものです。"
         "追加・新設ではなく「直した」ものに絞っています。")
    seigo_rows = [["元ファイルの記載", "直した内容", "根拠", "分類"]]
    hl = []
    for moto, naoshi, konkyo, cat in SEIGO:
        seigo_rows.append([moto, naoshi, konkyo, cat])
        hl.append(CAT_COLOR[cat])
    table(doc, seigo_rows, [2800, 2800, 3300, 600], highlights=hl)

    h2(doc, "5　新たに設けた章節")
    para(doc, f"{len(added)}件です。根拠は修正履歴の該当行をご覧ください。")
    table(doc, [["階層", "見出し"]]
          + [[s.replace("Heading ", "第") + "階層", t] for s, t in added],
          [1200, 8400])

    h2(doc, "6　見出しの文字が変わったもの")
    para(doc,
         f"{len(gone)}件です。節の新設に伴う番号の繰り下げが大半で、"
         "内容がなくなったものではありません。"
         "第3章は「第4次障がい者計画の推進状況（参考）」から"
         "「計画の基本理念・基本目標及び施策体系」へ再構成しています。")
    table(doc, [["元ファイルの見出し", "計画素案での扱い"]]
          + [[t, _gone_reason(t, sb)] for _s, t in gone], [3600, 6000])

    h2(doc, "7　残る確認事項")
    para(doc,
         "本書の修正のうち、村又は県の回答を待っている事項です。"
         "いずれも計画素案に【要確認：村】【要確認：県】として置いています。")
    table(doc, [
        ["番号", "先", "内容", "関係する修正"],
        ["M-10", "村", "継続入所者（旧指定施設等に引き続き入所している"
                      "18歳以上の方）に該当する方がいるか",
         "第5章1（告示が「ものとする」とした要件）"],
        ["M-11", "村", "村内の保育所・認定こども園・放課後児童クラブにおける"
                      "障がい児の受入人数と加配の状況",
         "第5章1・第5章7（同上）"],
        ["M-12", "村・県", "本村の精神病床1年以上長期入院者の実数と"
                        "年齢区分・入院先",
         "第4章（2）（基盤整備量）"],
        ["Q-5", "県", "第8期県計画の策定時期と、別表第四の三の項による"
                     "令和11年度末の基盤整備量の算定値",
         "第4章（2）（定めなければならない事項）"],
        ["Q-6", "県", "会津障がい保健福祉圏域で圏域確保が認められる"
                     "機能の範囲",
         "第1章2・第4章（4）・第5章8・第5章11"],
    ], [800, 900, 4600, 3300])

    note(doc,
         "本書は build_kitashiobara_shuusei_rireki.py が、"
         "原本・計画素案・素案の生成器の3つを読んで作っています。"
         "修正履歴は生成器の記録から機械で読み出しており、"
         "書き写しによる取り違えがありません。"
         "分類は語による規則で決めているため、"
         "生成器に修正を足して分類を決め忘れると自己点検で落ちます。")

    doc.save(OUT_FILE)
    print(f"作成: {OUT_FILE}")
    print(f"  元ファイル 段落{a['paras']}・表{a['tables']}・"
          f"見出し{len(a['heads'])}")
    print(f"  計画素案　 段落{b['paras']}・表{b['tables']}・"
          f"見出し{len(b['heads'])}")
    print(f"  修正履歴 {len(rows)}件　"
          + "／".join(f"{c}{sum(1 for r in rows if r[3] == c)}"
                      for c, _n, _col, _d in CATS))
    marked = sum(1 for r in rows if CAT_COLOR[r[3]])
    print(f"  マーカー対象 {marked}件（黄"
          f"{sum(1 for r in rows if CAT_COLOR[r[3]] == 'yellow')}・緑"
          f"{sum(1 for r in rows if CAT_COLOR[r[3]] == 'green')}）")
    print(f"  新設した章節 {len(added)}件／見出しが変わったもの {len(gone)}件")
    print(f"  元ファイルの記載を直した箇所 {len(SEIGO)}件")
    return a, b, rows, added, gone


def _gone_reason(text, sb):
    """消えた見出しの扱いを、素案の見出し集合から判定する。"""
    m = re.match(r"^([０-９0-9]+)　(.+)$", text)
    if m and any(t.endswith("　" + m.group(2)) for t in sb):
        return "節番号の繰り下げ（内容は残っている）"
    if text.startswith("第3章"):
        return "第3章として再構成（計画の基本理念・基本目標及び施策体系）"
    if text in ("基本理念", "基本目標", "基本施策と令和11年度目標値（第４次計画より）"):
        return "第3章1・2・5として独立した節に格上げ"
    return "要確認"


# ============================================================
# 自己点検
# ============================================================
def run_checks(a, b, rows, added, gone):
    ng = []

    # 1 すべての履歴が分類されていること（規則に当たらないものがない）
    for num, where, naka, cat in rows:
        if cat not in CAT_NAME:
            ng.append(f"分類が決まらない履歴: {where}／{naka[:40]}")

    # 2 分類ごとの件数が想定どおりであること（分類が動いたら気づく）
    got = {c: sum(1 for r in rows if r[3] == c) for c, _n, _col, _d in CATS}
    if got != EXPECT:
        ng.append(f"分類ごとの件数が変わっています: {got}（想定 {EXPECT}）。"
                  "RULES と EXPECT を見直してください")

    # 3 履歴が指す章が素案に実在すること
    heads = {t for _s, t in b["heads"]}
    flat = "".join(heads)
    for _num, where, _naka, _cat in rows:
        for sho in re.findall(r"第(\d+)章", where):
            if not any(t.startswith(f"第{sho}章") for t in heads):
                ng.append(f"履歴が実在しない章を指している: 第{sho}章（{where}）")

    # 4 新設した章節が素案の見出しに実在し、原本にないこと
    sa = {t for _s, t in a["heads"]}
    for _s, t in added:
        if t not in heads or t in sa:
            ng.append(f"新設の判定が合わない: {t}")

    # 5 見出しが変わったものの扱いが残らず判定できていること
    for _s, t in gone:
        if _gone_reason(t, heads) == "要確認":
            ng.append(f"消えた見出しの扱いが判定できない: {t}")

    # 6 f文字列の差し込みが残らず解けていること
    for _num, _where, naka, _cat in rows:
        if "{" in naka or "}" in naka:
            ng.append(f"差し込みが解けていない履歴: {naka[:50]}")

    # 7 差し込みに当てた値が素案の現物と合うこと
    d = docx.Document(SUAN)
    body = [p.text for p in d.paragraphs]
    for t in d.tables:
        for row in t.rows:
            body += [c.text for c in row.cells]
    text = "\n".join(body)
    for probe, need in (("用語解説", f"用語解説を新設（{PLACEHOLDER['{len(YOUGO)}']}語）"),):
        if probe not in text:
            ng.append(f"素案に{probe}が見つからない")
    yougo = next((t for t in d.tables
                  if len(t.columns) == 2
                  and t.rows[0].cells[0].text.strip() == "用語"), None)
    if yougo is None:
        ng.append("素案に用語解説の表が見つからない")
    elif len(yougo.rows) - 1 != int(PLACEHOLDER["{len(YOUGO)}"]):
        ng.append(f"用語解説の語数が合わない: 素案{len(yougo.rows) - 1}／"
                  f"履歴{PLACEHOLDER['{len(YOUGO)}']}")

    # 8 元ファイルの記載を直した箇所の分類が正しく、根拠があること
    for moto, naoshi, konkyo, cat in SEIGO:
        if cat not in CAT_NAME:
            ng.append(f"是正一覧の分類が不正: {cat}")
        if not konkyo.strip():
            ng.append(f"是正一覧に根拠がない: {moto[:30]}")

    # 9 元ファイルが当方の修正で書き換わっていないこと
    if a["sha256"] != GENPON_SHA:
        ng.append(f"元ファイルが変わっています: {a['sha256'][:16]}"
                  f"（記録 {GENPON_SHA[:16]}）")

    if ng:
        print("自己点検 不合格:")
        for x in ng:
            print("  -", x)
        return False
    print(f"自己点検: 合格　履歴{len(rows)}件すべて分類済み／"
          f"分類ごとの件数が想定どおり／履歴が指す章は素案に実在／"
          f"新設{len(added)}件・見出しの変化{len(gone)}件を判定／"
          f"差し込み0件／元ファイルのSHA-256が記録と一致")
    return True


# 元ファイル（他メンバー作成）のSHA-256。当方が書き換えていないことの証し。
GENPON_SHA = "577d7831b7a7b116e203a09ede1a44bcec4591dec30b9628222ae0db30ca00b3"


if __name__ == "__main__":
    import sys
    if not run_checks(*main()):
        sys.exit(1)
