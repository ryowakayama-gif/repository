# -*- coding: utf-8 -*-
"""
北塩原村 第8期障がい福祉計画・第4期障がい児福祉計画
素案 概要版ジェネレータ

出力: output/北塩原村_計画素案_概要版.docx（A4縦。頁数は flow() が決める。令和8年10月24日時点で8ページ）

現行計画の概要版（source/現行計画/第4次障がい者計画等_概要版.pdf、A4・4ページ）
の構成を踏襲する。ただし次の点が異なる。

  - 今回策定するのは障がい福祉計画（第8期）と障がい児福祉計画（第4期）の
    2計画である。障がい者計画（第4次）は令和11年度までを計画期間としており、
    本計画の策定にあわせて中間見直しを行う（正本 第1章2・第3章）。
    その基本理念・基本目標・施策体系は本計画と共有する。
    （令和8年10月24日まで「改定対象外」と書いていた。村の計画策定方針
    （中間見直しを行う）と正本に合わせて改めた。他ＡＩの横断レビュー RT-05）
  - 現行概要版の「4 SDGｓとの関連について」は、今回は扱わない方針のため
    「国の基本指針の見直しのポイント」に差し替える。
  - 成果目標は第8期で8項目に再編されたため、項目ごとの囲みではなく
    一覧表にまとめる。

4ページから6ページに増やした理由（令和8年10月10日）
  他メンバー版への移植を終えた正本には、現行計画になかった節がある。
  概要版が4ページのままでは、委員会とパブリックコメントに出す資料が
  計画の中身と食い違う。次の3つは、読み手が自分に関わるところを探す
  ための入口であり、落とせない。
    - 必要な見込量を確保するための方策（正本 第5章2（1）〜（15））
      医療へのアクセス・災害時・冬季の地域生活を含む
    - 年齢到達に伴うサービスの移行（正本 第5章5）
      18歳と65歳で何が変わり何が続くかを図で示す
    - 他計画との連携（正本 第7章1〜4）
      高齢者福祉計画・介護保険事業計画、こども・子育て計画を含む

6ページから7ページに増やした理由（令和8年10月14日）
  村の方針資料は「新規の追加施策」が一覧で見えることを求めている。
  令和8年10月14日に体系の折衷案（案Ｃ）が採用され、計画本文では
  第3章の章題を「第4次北塩原村障がい者計画の推進状況と本計画で
  新たに加える施策」へ改め、同章の末尾に一覧（表３-３）を置いた。
  概要版も同じ一覧を載せる（節8）。
  一覧は正本と同じ出所（build_kitashiobara_taikei_hanei.py の SHINKI）
  から引いており、概要版だけが別の一覧を持つことはない。
  あわせて節の番号を1つずつ繰り下げ（従前の8から15が9から16になる）、
  成果目標の8項目を2ページに分けず1ページに収めた。

正本から読む（令和8年10月24日）
  他ＡＩの横断レビュー（RT-05〜RT-08）で、概要版の見込量・成果目標・
  第4次計画の扱いが正本と食い違っていることが指摘された。
  見込量は第1次概算（build_kitashiobara_mikomiryo）を、照合は
  旧い計画素案（北塩原村_計画素案.docx）を見ていたため、
  正本が直っても概要版が置いていかれた。
  見込量（令和11年度）と実績（令和７年度）は正本 第5章の表から、
  成果目標は正本 第4章2の表の行と照合して書く。
  第8期の給付費の見込み（3か年計）は正本に無く、根拠の見込量が
  正本と違うため載せない。

工程の書き方（令和8年10月14日）
  従前は「令和8年度中 パブリックコメントの実施」のように時期を丸めて
  書いていた。概要版は住民に配るものであり、日付が分からなければ
  意見を出せない。工程が案Ａで決まったため、
  kitashiobara_koutei.py から日付を引いて書く。

出所
  数値は output/北塩原村_計画素案.docx 及び kitashiobara_common.py と同じ
  出典から取り、村に確認中のものには【要更新】を付す。
  令和8年10月に加えた3節は正本（output/北塩原村_計画素案_正本_移植後.docx）
  にしかないため、その現物と突き合わせる。したがって本書は
  build_kitashiobara_honpon.py の後に実行する。
  手帳所持者数の令和11年度の推計は
  build_kitashiobara_review3_taiou.py の shuusoku() から引く。
  工程は kitashiobara_koutei.py から引く。
  新たに加える施策は build_kitashiobara_taikei_hanei.py の SHINKI から
  引く（いずれも数字と一覧を書き写さない）。
"""

import os

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor

from kitashiobara_common import OUT_DIR
from kitashiobara_gaiyou_taisai import (
    ACCENT, BAND, BODY_PT, BORDER_COLOR, BOX_FILL, CHARS_PER_LINE,
    DARK, FONT, FONT_B, LINE_PT, LINES_PER_PAGE, NOTE_FILL, OBI,
    PAGE_W, SMALL_PT, TBL_PT,
)
from kitashiobara_koutei import (
    GAIYOU_HYO, PABU_KAISHI, PABU_NISSU, PABU_SHURYO, wareki,
)
# 新たに加える施策は正本 第3章4（表３-３）と同じ出所から引く。
# 概要版だけが別の一覧を持つと、計画と概要版で何が新しいかが食い違う。
from build_kitashiobara_taikei_hanei import SHINKI
# 手帳所持者数の令和11年度の推計。概要版で数字を書き写さず、
# 推計した生成器から引く（単一出所）。
from build_kitashiobara_review3_taiou import shuusoku  # noqa: F401
# アンケートの割合（回収40件）。節9の注記で用いる。
from kitashiobara_anke import pc as _anke_pc

OUT_FILE = f"{OUT_DIR}/北塩原村_計画素案_概要版.docx"

# 体裁は前回計画の概要版（令和6年3月）を実測して決めている。
# kitashiobara_gaiyou_taisai.py が生成のたびにＰＤＦを測り直し、
# 書体・本文の大きさ・基調色が前回と合っていることを確かめる。
# 行数と文字数も紙面と文字の大きさから計算しており、手で書かない。
#   書体     ＵＤデジタル教科書体（前回と同じ）
#   本文     10ポイント（前回は9〜10ポイントで10ポイントが主）
#   基調色   緑（前回の見出しの文字の色の実測から作る）


# ============================================================
# 書式ヘルパー
# ============================================================
def _el(tag, **attrs):
    e = OxmlElement(tag if ":" in tag else f"w:{tag}")
    for k, v in attrs.items():
        e.set(qn(f"w:{k}"), str(v))
    return e


def _cell_shade(cell, fill):
    cell._tc.get_or_add_tcPr().append(_el("shd", val="clear", color="auto", fill=fill))


def _cell_margins(cell, top=30, left=80, bottom=30, right=80):
    mar = _el("tcMar")
    for tag, v in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        mar.append(_el(tag, w=v, type="dxa"))
    cell._tc.get_or_add_tcPr().append(mar)


def _run(para, text, size=BODY_PT, bold=False, color=None):
    """文字を置く。

    太字は書体を替える。ＵＤデジタル教科書体は太さごとに別の書体
    （NK-R と NK-B）として登録されているため、レギュラーに太字を
    かけるとＷｏｒｄが字形を合成し、前回の概要版と見え方が変わる。
    """
    run = para.add_run(text)
    fn = FONT_B if bold else FONT
    run.font.name = fn
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = _el("rFonts")
        rpr.insert(0, rf)
    rf.set(qn("w:eastAsia"), fn)
    rf.set(qn("w:ascii"), fn)
    rf.set(qn("w:hAnsi"), fn)
    return run


def para(doc, text="", size=BODY_PT, bold=False, align=None, space_after=2,
         indent=0, color=None, line=LINE_PT):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(space_after)
    pf.line_spacing = Pt(line)
    if indent:
        pf.left_indent = Pt(indent)
        pf.first_line_indent = Pt(-indent)
    if align is not None:
        p.alignment = align
    if text:
        _run(p, text, size=size, bold=bold, color=color)
    return p


def band(doc, number, title):
    """章見出しの帯（濃紺地に白抜き）。"""
    t = doc.add_table(rows=1, cols=1)
    t.autofit = False
    cell = t.cell(0, 0)
    cell.width = Mm(180)
    _cell_shade(cell, DARK)
    _cell_margins(cell, top=50, bottom=50, left=110)
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = Pt(14)
    _run(p, f"{number}　{title}" if number else title,
         size=11.5, bold=True, color="FFFFFF")
    _tbl_width(t)
    _no_borders(t)
    para(doc, "", size=4, space_after=0, line=5)
    return t


def sub(doc, text, fill=BAND):
    """小見出しの帯（淡色地）。"""
    t = doc.add_table(rows=1, cols=1)
    t.autofit = False
    cell = t.cell(0, 0)
    cell.width = Mm(180)
    _cell_shade(cell, fill)
    _cell_margins(cell, top=20, bottom=20, left=100)
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = Pt(LINE_PT)
    _run(p, text, size=SMALL_PT, bold=True, color=DARK)
    _tbl_width(t)
    _no_borders(t)
    para(doc, "", size=3, space_after=0, line=4)
    return t


def bullets(doc, lines, size=SMALL_PT):
    for line in lines:
        para(doc, f"・{line}", size=size, indent=9)


def _tbl_width(t):
    t._tbl.tblPr.append(_el("tblW", w=PAGE_W, type="dxa"))
    t._tbl.tblPr.append(_el("tblLayout", type="fixed"))


def _no_borders(t):
    borders = _el("tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        borders.append(_el(edge, val="nil"))
    t._tbl.tblPr.append(borders)


def _borders(t, color=BORDER_COLOR):
    borders = _el("tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        borders.append(_el(edge, val="single", sz=4, space=0, color=color))
    t._tbl.tblPr.append(borders)


def table(doc, rows, widths, header=True, size=TBL_PT, aligns=None,
          fill_first=None):
    """本文の表。widths は twips で合計 PAGE_W とする。"""
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.autofit = False
    for r, values in enumerate(rows):
        for c, value in enumerate(values):
            cell = t.cell(r, c)
            cell.width = Mm(widths[c] * 180 / PAGE_W)
            _cell_margins(cell)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = Pt(TBL_PT * 1.35)
            if aligns and aligns[c] == "center":
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            is_head = header and r == 0
            if is_head:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _run(p, str(value), size=size, bold=is_head,
                 color="FFFFFF" if is_head else None)
            if is_head:
                _cell_shade(cell, DARK)
            elif fill_first and c == 0:
                _cell_shade(cell, fill_first)
    _tbl_width(t)
    _borders(t)
    # 列幅を明示する
    grid = t._tbl.find(qn("w:tblGrid"))
    for i, gc in enumerate(grid.findall(qn("w:gridCol"))):
        gc.set(qn("w:w"), str(widths[i]))
    para(doc, "", size=4, space_after=0, line=5)
    return t


def box(doc, lines, fill=BOX_FILL, border=BORDER_COLOR, size=SMALL_PT,
        bold_first=False):
    """囲み（強調・注記）。"""
    t = doc.add_table(rows=1, cols=1)
    t.autofit = False
    cell = t.cell(0, 0)
    cell.width = Mm(180)
    _cell_shade(cell, fill)
    _cell_margins(cell, top=60, bottom=60, left=140, right=140)
    first = cell.paragraphs[0]
    for i, line in enumerate(lines):
        p = first if i == 0 else cell.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = Pt(LINE_PT)
        _run(p, line, size=size, bold=(bold_first and i == 0))
    _tbl_width(t)
    _borders(t, border)
    para(doc, "", size=4, space_after=0, line=5)
    return t


def note(doc, text):
    return box(doc, [text], fill=NOTE_FILL, border="C9A03B", size=8)


def _make_break(doc):
    """改ページの段落を作って本文から切り離して返す。"""
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = Pt(1)
    _run(p, "", size=1).add_break(WD_BREAK.PAGE)
    el = p._p
    el.getparent().remove(el)
    return el


def page_break(doc):
    el = _make_break(doc)
    doc.element.body.append(el)


def flow(doc, blocks):
    """ブロックを順に置き、1頁に収まらなくなる手前で改ページを入れる。

    改ページの位置を手で置いていたときは、文章を1行足すたびに
    どの頁が溢れるかを人が確かめ直すことになっていた。
    前回計画の概要版に合わせて本文を9ポイントから10ポイントに
    上げたところ2頁が溢れたため、頁の割り方を計算に任せる形に改めた。
    各ブロックを置いた後に行数を数え、その頁が目安を超えたときは
    そのブロックの手前へ改ページを入れて次の頁に送る。
    """
    body = doc.element.body
    oki = []
    for name, fn in blocks:
        mae = len(list(body.iterchildren()))
        fn(doc)
        els = list(body.iterchildren())[mae:]
        if not els:
            continue
        pages = estimate_lines(doc)
        if mae > 0 and pages[-1] > LINES_PER_PAGE:
            els[0].addprevious(_make_break(doc))
            pages = estimate_lines(doc)
        oki.append((name, len(pages), pages[-1]))
    return estimate_lines(doc), oki


def figure(doc, name, title, src, width_mm=168):
    """図を貼り、下に表題と出所を置く（正本と同じ作法）。"""
    path = f"{OUT_DIR}/図表/{name}"
    if not os.path.exists(path):
        raise SystemExit(
            f"図が見つかりません: {path}\n"
            "build_kitashiobara_graph.py を先に実行してください。")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(1)
    p.add_run().add_picture(path, width=Mm(width_mm))
    cap = para(doc, title, size=8.5, bold=True,
               align=WD_ALIGN_PARAGRAPH.CENTER, space_after=1)
    para(doc, src, size=7.5, space_after=4)
    return cap


# ============================================================
# 本文データ
# ============================================================
HYOKI_TEXT = (
    "本計画では「障害」などの「害」の字の表記について、字に対する印象に配慮するとともに、"
    "障がい者の人権をより尊重する観点から、国の法令等に基づく法律用語や施設名等の固有名称を除き、"
    "可能な限り「害」の字をひらがなで表記しています。このため、本計画では「がい」と「害」の字が"
    "混在する表現となっています。"
)

ICHIZUKE = [
    ["区分", "障害者計画", "障害福祉計画", "障害児福祉計画",
     "成年後見制度利用促進計画"],
    ["根拠法", "障害者基本法\n第11条第３項", "障害者総合支援法\n第88条第１項",
     "児童福祉法\n第33条の20第１項", "成年後見制度の利用の促進に関する法律\n第14条第１項"],
    ["内容", "障がい者施策の基本的方向性について定める計画",
     "障害福祉サービス等の見込量とその確保策を定める計画",
     "障害児通所支援等の提供体制と確保策を定める計画",
     "成年後見制度の利用の促進に関する施策の基本的な計画"],
    ["今回の取扱い", "中間見直し\n（第４次計画）", "第８期として策定",
     "第４期として策定", "本計画の第６章として策定"],
]

KIKAN = [
    ["計画", "計画期間", "今回の取扱い"],
    ["第４次北塩原村障がい者計画", "令和６年度〜令和11年度（６年間）",
     "本計画の策定にあわせて中間見直し。基本理念・基本目標・施策体系を本計画と共有"],
    ["第８期北塩原村障がい福祉計画", "令和９年度〜令和11年度（３年間）", "今回策定"],
    ["第４期北塩原村障がい児福祉計画", "令和９年度〜令和11年度（３年間）", "今回策定"],
    ["成年後見制度の利用促進に関する計画", "令和９年度〜令和13年度（５年間）",
     "今回策定（本計画の第６章）"],
]

SHISHIN_POINTS = [
    "成果目標が７項目から８項目になりました。「障がい福祉人材の確保・定着、"
    "当事者視点に立ったケアの充実のための生産性向上」が新たに加わっています。",
    "就労選択支援（令和７年10月開始）の積極的な利用促進が求められ、"
    "就労選択支援に係る成果目標が新設されました。",
    "障がい児支援では、児童発達支援センター等の「４つの中核機能」の確保と、"
    "インクルージョン（地域社会への参加・包容）推進のための協議の場の設置が求められています。",
    "のぞまないセルフプラン（自己作成のサービス等利用計画）の解消が、"
    "相談支援体制の成果目標に位置づけられました。",
    "障害福祉サービス等情報公表制度の公表率に関する成果目標が新設されました。",
    "人口減少地域におけるサービスの維持・確保、高次脳機能障害のある方への支援、"
    "災害時におけるサービス提供の確保、スポーツ・健康増進活動による社会参加の促進などが、"
    "新たな視点として示されています。",
]

# 平成31年〜令和5年は現行計画（第4次障がい者計画）、令和8年は正本 第2章2 による。
# 以前は令和11年の中位推計（合計145人）を併記していたが、
# その推計は令和5年までの減少を延ばしたものであり、
# 令和8年の実績181人（増加）と前提が合わないため外した。
# 令和8年10月13日に策定委員会資料から療育手帳・精神障害者保健福祉手帳の
# 令和6年・令和7年の値を受領し、6点の回帰で推計をやり直した。
# 令和11年度の推計は build_kitashiobara_review3_taiou.py の
# shuusoku() が計算する。概要版で数字を書き写さない。
TECHOU = [
    ["区分", "平成31年", "令和３年", "令和５年", "令和８年",
     "令和５年→\n令和８年"],
    ["身体障害者手帳", "126人", "128人", "114人", "127人", "＋13人"],
    ["療育手帳", "13人", "14人", "12人", "16人", "＋４人"],
    ["精神障害者保健福祉手帳", "24人", "27人", "30人", "38人", "＋８人"],
    ["合計", "163人", "169人", "156人", "181人", "＋25人"],
]

RINEN = "「障がいのあるなしに関わらず、お互いの人格や個性を尊重し、多様な価値観を認め合い、誰もが自分らしく輝くむら」"

MOKUHYOU = [
    ("①　障がいへの理解を深め、交流を育むむら",
     "障がいの有無に関わらず、個人の人格や個性を尊重し合うことを目指し、"
     "障がいを理由とする差別の解消や障がいのある人に関する正しい理解が必要であるため、"
     "広報等を活用した啓発活動を推進します。"),
    ("②　共に支え合い、誰もが安心して暮らせるむら",
     "すべての人が共に協力し合い、誰もが住み慣れた地域で安心して暮らすことを目指し、"
     "障がいにつながる疾病を予防・軽減するため、保健・医療との連携や"
     "災害などの緊急時における安全・安心の確保に取り組みます。"),
    ("③　みんなが輝き、自立した生活を送れるむら",
     "生きがいを持って活動できる社会を目指し、障がいのある人も様々な可能性の中から"
     "自分らしい生き方を選択できるように、ライフステージに沿った切れ目のない支援を進めます。"),
]

# 施策体系（節9）。
#   「本計画で新たに加える施策」の列は正本 表３-３ と同じ出所（SHINKI）の
#   基本施策の対応から、「目標値」の列は正本 表３-１ の現物から読む。
#   以前は手で書いており、医療的ケア児等への支援を②に、虐待の防止を①に
#   置くなど、正本 表３-３ と対応が食い違っていた（横断レビュー RT 指摘14）。
KIHON = ["①　啓発・広報", "②　保健・医療", "③　福祉", "④　教育・育成",
         "⑤　雇用・就業", "⑥　生活環境", "⑦　スポーツ・文化"]
SHIHYOU_RYAKU = {
    "差別や偏見を感じている障がいがある人の割合": "差別や偏見を感じている人の割合",
    "特定健康診査の受診率": "特定健康診査の受診率",
    "現在利用している福祉サービスに満足している人の割合":
        "福祉サービスに満足している人の割合",
    "医療的ケア児の受入れ体制の整備": "医療的ケア児の受入れ体制の整備",
    "障がい者就労施設等からの物品調達件数": "障がい者就労施設等からの物品調達件数",
    "障がい者に配慮した公共施設トイレの設置率": "公共施設トイレの設置率",
    "個別避難計画の策定件数": "個別避難計画",
    "外出目的が趣味・スポーツ・グループ活動である人の割合":
        "外出の目的が趣味・スポーツ等である人の割合",
}
# 扱いを決める必要がある指標（節9の注記で理由を書く）
YOUKAKUNIN = {"差別や偏見を感じている障がいがある人の割合",
              "現在利用している福祉サービスに満足している人の割合",
              "個別避難計画の策定件数",
              "外出目的が趣味・スポーツ・グループ活動である人の割合"}


def shisaku_rows():
    """節9の表。正本 表３-１ の行と SHINKI の基本施策から組む。"""
    hyo = None
    for t in _honpon_doc().tables:
        rows = [[c.text.strip() for c in r.cells] for r in t.rows]
        if rows[0][:2] == ["基本施策", "指標項目"]:
            hyo = rows[1:]
    if hyo is None:
        raise LookupError("正本に表３-１（基本施策の目標値）が無い")
    out = [["基本施策", "本計画で新たに加える施策", "目標値（現状→令和11年度）"]]
    for k in KIHON:
        kk = k.replace("　", "")
        shinki = [r[0] for r in SHINKI if r[3] == kk]
        moku = []
        for r in hyo:
            if r[0] != kk:
                continue
            m = f"{SHIHYOU_RYAKU[r[1]]}\n{r[2]}→{r[3]}"
            if r[1] in YOUKAKUNIN:
                m += "【要確認】"
            moku.append(m)
        out.append([k, "、".join(shinki) if shinki else
                    "（第４次計画の取組みを継続）", "\n".join(moku)])
    return out


# 成果目標（節10）。値は正本 第4章2 の表の行と照合する（SEIKA_SHOUGOU）。
SEIKA = [
    ["成果目標", "現状（基準）", "令和11年度の目標値"],
    ["（１）福祉施設の入所者の\n　　　地域生活への移行",
     "施設入所者数　４人\n（令和７年度末）",
     "地域生活移行者数　０人\n施設入所者の削減見込み　０人\n"
     "（入所者お一人ずつの意向を年１回以上確かめます）"],
    ["（２）精神障害にも対応した\n　　　地域包括ケアシステムの構築",
     "精神病床に１年以上入院している方　５人\n（第７期計画時点）",
     "退院者数　１人\n心のサポーター養成研修　年１回・５人【新規】\n"
     "Ｋ６等によるこころの状態の把握【新規】\n"
     "長期入院患者の地域生活への移行に伴う基盤整備量【県の算定値待ち】"],
    ["（３）福祉施設から\n　　　一般就労への移行等",
     "一般就労移行者数　０人（令和６年度）\n"
     "就労定着支援事業の利用者数　０人（令和６年度末）",
     "一般就労移行者数　１人\n就労定着支援事業の利用者数　２人\n"
     "就労選択支援の利用者数　２人【新規】"],
    ["（４）障がい児支援の\n　　　提供体制の整備等",
     "児童発達支援センター　０カ所\n保育所等訪問支援事業所　０カ所\n"
     "医療的ケア児等支援の協議の場　検討中\n"
     "医療的ケア児等コーディネーター　１人",
     "４つの中核機能を近隣市町村との連携により確保\n"
     "インクルージョン推進の協議の場　１カ所（会津北部圏域で設置）【新規】\n"
     "医療的ケア児等支援の協議の場　１カ所（近隣町村と共同設置）\n"
     "医療的ケア児等コーディネーター　１人（継続）\n"
     "伴走的な相談支援体制の構築【新規】"],
    ["（５）地域生活支援の充実",
     "会津北部地域生活支援拠点（４町村共同）　１カ所\nコーディネーター　１人\n"
     "運用状況の検証・検討　年１回",
     "拠点の運用状況の検証・検討　年１回以上\n"
     "強度行動障害を有する方の支援体制を圏域の協議の場で整備"],
    ["（６）相談支援体制の\n　　　充実・強化等",
     "基幹相談支援センター　０カ所\n（国の一覧では令和７年４月１日時点で"
     "会津障がい保健福祉圏域13市町村のうち８市町村が設置済み・３カ所）\n"
     "協議会専門部会　１カ所",
     "基幹相談支援センター　１カ所（会津北部４町村で広域設置）\n"
     "協議会専門部会　２カ所\nのぞまないセルフプラン　０件【新規】"],
    ["（７）障がい福祉人材の確保・定着\n　　　及び生産性向上【新規】",
     "―（新たに加わった項目）",
     "県の相談支援専門員研修等への村内関係者の参加　年２人\n"
     "県のワンストップ窓口・生産性向上支援策を事業所へ周知"],
    ["（８）障がい福祉サービス等の\n　　　質の向上に係る体制の構築",
     "県研修への村職員の参加　２人（令和８年度）",
     "県研修への参加　年２人\n審査結果の分析・共有　年１回\n"
     "村内の対象事業所の情報公表　公表率・更新率100％【新規】"],
]

# 成果目標の値が正本 第4章2 の表の行と同じであることの照合
# (正本の表の行の1列目に含まれる語, 2列目の書き出し, 概要版の SEIKA に出る語)
SEIKA_SHOUGOU = [
    ("令和７年度末時点の施設入所者数", "４人", "施設入所者数　４人"),
    ("地域生活移行者数", "０人", "地域生活移行者数　０人"),
    ("削減見込者数", "０人", "削減見込み　０人"),
    ("長期入院者の退院者数", "１人", "退院者数　１人"),
    ("心のサポーター養成研修の実施回数", "１回", "年１回"),
    ("心のサポーター養成研修の受講者数", "５人", "５人【新規】"),
    ("令和６年度の一般就労への移行者数", "０人", "一般就労移行者数　０人"),
    ("一般就労移行者数", "１人", "一般就労移行者数　１人"),
    ("令和６年度末の就労定着支援事業の利用者数", "０人",
     "就労定着支援事業の利用者数　０人"),
    ("就労定着支援事業の利用者数", "２人", "就労定着支援事業の利用者数　２人"),
    ("就労選択支援の利用者数", "２人", "就労選択支援の利用者数　２人"),
    ("児童発達支援センターの設置数", "０カ所", "児童発達支援センター　０カ所"),
    ("保育所等訪問支援事業所の整備数", "０カ所", "保育所等訪問支援事業所　０カ所"),
    ("インクルージョン推進のための協議の場", "未設置", "インクルージョン推進の協議の場"),
    ("医療的ケア児等支援のための協議の場", "検討中", "協議の場　検討中"),
    ("医療的ケア児等支援のための協議の場", "検討中", "１カ所（近隣町村と共同設置）"),
    ("医療的ケア児等コーディネーターの配置", "１人", "コーディネーター　１人"),
    ("令和８年度末時点の設置数", "１カ所", "１カ所"),
    ("令和８年度末時点のコーディネーターの配置人数", "１人", "コーディネーター　１人"),
    ("令和８年度の運用状況の検証・検討回数", "１回", "年１回"),
    ("検証・検討回数", "１回以上", "年１回以上"),
    ("令和８年度末時点の基幹相談支援センターの設置数", "０カ所",
     "基幹相談支援センター　０カ所"),
    ("基幹相談支援センターの設置数", "１カ所", "基幹相談支援センター　１カ所"),
    ("令和８年度末時点の協議会専門部会の設置数", "１カ所", "協議会専門部会　１カ所"),
    ("協議会専門部会の設置数", "２カ所", "協議会専門部会　２カ所"),
    ("のぞまないセルフプランの件数", "０件", "のぞまないセルフプラン　０件"),
    ("県が実施する相談支援専門員研修等への村内関係者の参加人数", "２人", "年２人"),
    ("令和８年度の県研修への村職員の参加人数", "２人", "２人（令和８年度）"),
    ("県研修への村職員の参加人数", "２人", "県研修への参加　年２人"),
    ("審査支払等システムの審査結果の分析・共有", "１回", "年１回"),
    ("情報公表制度の公表率・更新率", "100％", "100％"),
]

PDCA = [
    ["段階", "内容"],
    ["Plan（計画）", "成果目標及びサービスの見込量（活動指標）を定めます。"],
    ["Do（実行）", "計画に基づきサービスの提供体制を確保し、事業を実施します。"],
    ["Check（評価）",
     "少なくとも年１回、実績を把握し、成果目標及び活動指標の達成状況を分析・評価します。"],
    ["Action（改善）",
     "評価の結果を踏まえ、必要に応じて計画の変更やサービス提供体制の確保に係る対策を講じます。"],
]

# ---------------------------------------------------------------------------
# 必要な見込量を確保するための方策（正本 第5章2（1）〜（15）を束ねたもの）
#   15項目をそのまま並べると概要版に入らないため、
#   読み手が自分に関わるところを探せる束ね方にする。
#   どの項を束ねたかは verify() で正本の現物と照合する。
# ---------------------------------------------------------------------------
HOUSAKU = [
    ("情報提供と相談",
     "広報誌・ホームページ・窓口を通じた情報提供と、自立支援協議会を"
     "中心とした関係機関のネットワークによる総合的な相談支援を進めます。"),
    ("広域連携によるサービスの確保",
     "多くのサービスを近隣市町村に依存しているため、会津北部圏域の町村及び"
     "県と連携して確保し、村外の事業所に受入余力を年１回照会します。"
     "供給が不足する場合は県の事業者指定に意見を申し出ます。"),
    ("一般就労への移行等の推進",
     "就労選択支援の活用等により本人に合った働き方を選択できるよう支援し、"
     "ハローワーク等と一体となって一般就労につなげます。"),
    ("障がい児支援の提供体制の確保",
     "乳幼児健診等による早期発見から、療育、就学、卒業後の進路まで、"
     "保健・医療・福祉・保育・教育が連携した切れ目のない支援体制を"
     "確保します。"),
    ("人材の確保と事業所の経営基盤",
     "県の研修や支援策を活用し、人材の確保・定着とサービスの質の向上を"
     "図ります。事業所が存続しなければ見込量は確保できないため、"
     "経営基盤を供給の制約として把握します。"),
    ("住まいの確保",
     "グループホーム及び自立生活援助に加え、住宅セーフティネット制度との"
     "連携により、一般住宅での生活の選択肢を広げます。"),
    ("災害時・冬季の地域生活の継続",
     "事業所の業務継続計画や村外事業所との連絡手段の確保など、災害時に"
     "サービスを止めない備えを進めます。冬季は除雪、通院・通所の確保、"
     "通行規制による中断への備えを整理します。"),
    ("虐待の予防から再発の防止まで",
     "通報を受けてからの対応だけでなく、虐待に至る前の家族の介護負担の"
     "把握から再発の防止までを一続きの流れとして整理します。"),
    ("医療へのアクセス",
     "村内で受けられる医療と圏域の市町に出向いて受ける医療とを分けて示し、"
     "休日・夜間及び急病のときにどこへかかるかをあらかじめ分かる形に"
     "しておきます。"),
]

# ---------------------------------------------------------------------------
# 他計画との連携（正本 第7章1〜4）
# ---------------------------------------------------------------------------
RENKEI = [
    ["村の計画", "計画期間", "連携して進めること"],
    ["第３次健康21・\n北塩原村グッドヘルスプラン", "令和６年度〜\n令和17年度",
     "こころの健康づくりとゲートキーパーの養成、母子保健事業を通じた"
     "医療的ケア児等への支援、障がいのある方に配慮した健診の受診環境、"
     "公民館等を通じた社会参加の機会"],
    ["北塩原村第五次総合振興計画", "平成29年度〜\n令和８年度",
     "村の最上位計画。施策19「穏やかな暮らしの支援」を本計画が"
     "具体化します。防災、地域活動、人口減少への対応で連携します。"
     "令和８年度で計画期間が終了するため、次期計画との整合を図ります"],
    ["第10期北塩原村高齢者福祉計画・\n介護保険事業計画",
     "令和９年度〜\n令和11年度\n（本計画と同一）",
     "65歳及び40歳の到達による制度間の移行、介護保険の適用除外、"
     "共生型サービス。両計画が同じ方を重ねて数えることのないよう、"
     "対象となる方を一人ずつ確かめて両計画に同時に反映します"],
    ["北塩原村こども・子育て計画",
     "令和７年度〜\n令和11年度\n（令和９年度に\n中間見直し）",
     "保育所等における障がい児の受入れと障がい児通所支援は"
     "入れ替わるものではなく併せて使うものです。"
     "令和９年度の中間見直しの場で、両計画の見込量の整合を確かめます"],
]

IKOU_LEAD = (
    "一定の年齢に達すると対象となる制度が変わります。本計画では18歳"
    "（学校卒業）、22歳（障害児入所施設からの移行）、65歳（介護保険）の"
    "引継ぎを整理しており、下の図は18歳と65歳について示したものです。"
)

IKOU_NOTE = (
    "65歳に達しても、一律に介護保険へ移るわけではありません。"
    "介護保険に相当するサービスがない共同生活援助、就労系サービス、"
    "同行援護、行動援護等は、65歳到達後も障がい福祉サービスとして"
    "利用できます。本村では、65歳到達の前に相談支援専門員と"
    "地域包括支援センターが連携して本人の意向を確認し、"
    "必要な場合は両制度を併せて利用できるよう調整します。"
    "また、相談、権利擁護（成年後見制度）、緊急時・災害時の備え、"
    "冬季の除雪と通院・通所の確保は、どの時期にも共通して続く支援です。"
)

# 工程は kitashiobara_koutei.py から引く（単一出所）。
# 以前は「令和８年度中」のように丸めて書いていたが、
# 概要版は住民に配るものであり、
# パブリックコメントの日付が分からなければ意見を出せない。
SCHEDULE = GAIYOU_HYO

# 新たに加える施策（正本 表３-３ と同じ出所）。
# 概要版では紙面の都合で「施策」と「本計画での位置」の2列にする。
SHINKI_HYO = [["本計画で新たに加える施策", "本計画での位置"]] + [
    [shisaku, ichi] for shisaku, _naka, ichi, _kihon, _go in SHINKI
]

# 概要版に載せるサービス。計画期間に値が動くもの（生活介護・就労継続支援Ｂ型・
# 共同生活援助・児童発達支援・放課後等デイサービス）と、新規に立つもの
# （就労選択支援・短期入所・保育所等訪問支援）を落とさない。
# 名称は正本 第5章の表のもの（改行を除いた形）に合わせる。
MIKOMI_PICK = [
    "居宅介護", "生活介護", "自立訓練（生活訓練）", "就労選択支援",
    "就労移行支援", "就労継続支援Ａ型", "就労継続支援Ｂ型", "就労定着支援",
    "短期入所（福祉型）", "共同生活援助", "施設入所支援", "計画相談支援",
    "児童発達支援", "放課後等デイサービス", "保育所等訪問支援",
    "障がい児相談支援",
]

_DOC = []


def _honpon_doc():
    if not _DOC:
        path = f"{OUT_DIR}/北塩原村_計画素案_正本_移植後.docx"
        if not os.path.exists(path):
            raise SystemExit(f"正本が見つかりません: {path}\n"
                             "build_kitashiobara_honpon.py を先に実行してください。")
        _DOC.append(docx.Document(path))
    return _DOC[0]


def _suuji(v):
    """「5人分」「0.3（1）人分」→ 先頭の数（「―」は None）。"""
    import re
    m = re.match(r"^([0-9.]+|―)", v)
    if not m or m.group(1) == "―":
        return None
    x = float(m.group(1))
    return int(x) if x == int(x) else x


def honpon_hyou(nendo):
    """正本 第5章の表から {サービス名: [人／月の値 ×3年度]} を読む。

    nendo="見込" は令和９〜11年度の見込量の表（セルが「5人分」）、
    nendo="実績" は令和６〜８年度の実績の表（セルが「3（4）人分」）。
    第4章の表（セルが「1人」で「分」が付かない）は拾わない。
    """
    head = (["令和９年度", "令和10年度", "令和11年度"] if nendo == "見込"
            else ["令和６年度", "令和７年度", "令和８年度"])
    out = {}
    for t in _honpon_doc().tables:
        rows = [[c.text.strip() for c in r.cells] for r in t.rows]
        if rows[0][:2] != ["サービス種別", "単位"] or rows[0][2:5] != head:
            continue
        for r in rows[1:]:
            if r[1] != "人／月" or not r[2].endswith("分"):
                continue
            out.setdefault(r[0].replace("\n", ""), [_suuji(v) for v in r[2:5]])
    return out


def mikomi_rows():
    """主なサービスの令和７年度の実績と令和11年度の見込量を2列組にする。

    いずれも正本 第5章の表の現物から読む（第1次概算は読まない）。
    """
    jis, mik = honpon_hyou("実績"), honpon_hyou("見込")
    missing = [n for n in MIKOMI_PICK if n not in mik]
    if missing:
        raise LookupError(f"正本の見込量の表に無いサービスです: {missing}")
    pairs = []
    for n in MIKOMI_PICK:
        r7 = jis.get(n, [None, None, None])[1]
        moto = "―" if r7 is None else f"{r7}人"
        pairs.append((n, f"{moto}→{mik[n][2]}人"))
    half = (len(pairs) + 1) // 2
    left, right = pairs[:half], pairs[half:]
    rows = [["サービス", "令和７年度→令和11年度", "サービス",
             "令和７年度→令和11年度"]]
    for i in range(half):
        a = left[i]
        b = right[i] if i < len(right) else ("", "")
        rows.append([a[0], a[1], b[0], b[1]])
    return rows


# ============================================================
# ページ数の見積り（4ページに収まっているかの確認用）
# ============================================================
def estimate_lines(doc):
    """改ページで区切って、各ページのおよその行数を数える。

    図は段落の中に入るため、その高さ（EMU）を行に換算して足す。
    足さないと、図を貼ったページだけ見積りが実際より軽く出る。
    """
    pages = [0]
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            text = "".join(t.text or "" for t in child.iter(qn("w:t")))
            if child.findall(".//" + qn("w:br") + "[@" + qn("w:type") + "='page']"):
                pages.append(0)
                continue
            zu = child.findall(".//" + qn("wp:extent"))
            if zu:
                # 1行＝LINE_PT（行送りを固定しているため）。
                # 1ポイント＝12,700EMU
                pages[-1] += (sum(int(x.get("cy")) for x in zu)
                              / (LINE_PT * 12700))
                continue
            pages[-1] += max(1, -(-len(text) // CHARS_PER_LINE)) if text else 0.4
        elif child.tag == qn("w:tbl"):
            for tr in child.findall(qn("w:tr")):
                longest = 1
                for tc in tr.findall(qn("w:tc")):
                    text = "".join(t.text or "" for t in tc.iter(qn("w:t")))
                    hard = text.count("\n") + 1
                    width = int(tc.find(qn("w:tcPr")).find(qn("w:tcW")).get(qn("w:w"))) \
                        if tc.find(qn("w:tcPr")) is not None and \
                        tc.find(qn("w:tcPr")).find(qn("w:tcW")) is not None else PAGE_W
                    chars = max(4, int(CHARS_PER_LINE * width / PAGE_W))
                    longest = max(longest, max(hard, -(-len(text) // chars)))
                pages[-1] += longest * 0.92
    return pages


# ============================================================
# 各ページ
# ============================================================
def blk_hyoushi(doc):
    # 表紙帯
    t = doc.add_table(rows=1, cols=1)
    t.autofit = False
    cell = t.cell(0, 0)
    cell.width = Mm(180)
    _cell_shade(cell, DARK)
    _cell_margins(cell, top=120, bottom=120, left=160)
    lines = [
        ("◆　第８期北塩原村障がい福祉計画", 13),
        ("◆　第４期北塩原村障がい児福祉計画", 13),
        ("（令和９年度〜令和11年度）　素案　概要版", 11),
    ]
    for i, (text, size) in enumerate(lines):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = Pt(size + 6)
        _run(p, text, size=size, bold=True, color="FFFFFF")
    _tbl_width(t)
    _no_borders(t)
    para(doc, "", size=4, space_after=0, line=5)

    p = para(doc, align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=4)
    _run(p, "作成年月　令和８年10月　／　発行　北塩原村　／　"
            "編集　北塩原村保健福祉課　福祉係　（TEL）0241-23-3113", size=8.5)

def blk_shushi(doc):
    band(doc, "１", "計画策定の背景・趣旨")
    para(doc,
         "本村では、令和６年３月に「第４次北塩原村障がい者計画・第７期北塩原村障がい福祉計画・"
         "第３期北塩原村障がい児福祉計画」を策定し、障がいのある人もない人も互いに人格と個性を"
         "尊重し合う共生社会の実現を目指してきました。")
    para(doc,
         "このうち「第７期北塩原村障がい福祉計画」及び「第３期北塩原村障がい児福祉計画」は"
         "令和８年度で計画期間が終了することから、令和８年３月31日に告示された国の基本指針、"
         "これまでの本村の取組み及び障がいのある方のニーズを踏まえ、令和９年度から令和11年度までの"
         "３年間について、障がい福祉サービス等の見込量とその確保のための方策を定めるものです。")
    para(doc,
         "なお、「第４次北塩原村障がい者計画」は令和11年度までを計画期間としており、"
         "本計画の策定にあわせて中間見直しを行います。基本施策ごとの目標値の"
         "達成状況を確認し、目標値を改める場合は計画本文の第３章に掲げます。"
         "その基本理念・基本目標及び施策体系は、本計画においても共有します。", space_after=4)

    sub(doc, "「障害」と「障がい」の表記について")
    para(doc, HYOKI_TEXT, space_after=4)


def blk_ichizuke(doc):
    band(doc, "２", "計画の位置づけ")
    para(doc, "本計画は、以下の法律に基づいて策定する法定計画です。")
    table(doc, ICHIZUKE, [1150, 2200, 2200, 2200, 2456],
          aligns=["center", None, None, None, None], fill_first=BAND, size=8)

def blk_kikan(doc):
    band(doc, "３", "計画の期間")
    table(doc, KIKAN, [3000, 2700, 4506], fill_first=BAND)


def blk_shishin(doc):
    band(doc, "４", "国の基本指針の見直しのポイント")
    para(doc,
         "第８期障害福祉計画及び第４期障害児福祉計画に係る国の基本指針は令和８年３月31日に"
         "告示され、主な見直し事項として次の点が示されています。")
    bullets(doc, SHISHIN_POINTS, size=SMALL_PT)


def blk_genjou(doc):
    band(doc, "５", "本村の障がい者の現状")
    para(doc,
         "本村の人口（住民基本台帳）は令和８年４月１日現在2,316人で、"
         "令和４年（2,535人）から219人減少しています。"
         "高齢化率は43.1％（令和８年度）で、令和11年度には総人口が2,100人台、"
         "高齢化率は45％を超えることが見込まれます。")

    sub(doc, "障がい者手帳の所持者数の推移")
    table(doc, TECHOU, [2206, 1400, 1400, 1400, 1400, 1400],
          aligns=[None, "center", "center", "center", "center", "center"])
    box(doc, [
        "障がい者手帳をお持ちの方は181人（令和８年４月現在）で、人口の約7.8％、"
        "約13人に１人の割合です。",
        "　身体障がいのある人　127人（5.5％）／知的障がいのある人　16人（0.7％）／"
        "精神障がいのある人　38人（1.6％）",
        "第７期計画の策定時（令和５年４月156人）と比べて25人増えており、"
        "特に精神障害者保健福祉手帳をお持ちの方が増加しています。",
    ])
    #   令和11年度の見通しは正本が単一の値を示していない（村の台帳の
    #   確認を待つ）。概要版だけが数を示すと計画と食い違う（RT 指摘9）。
    note(doc, "【確認中】令和11年度の手帳所持者数の見通しは、手帳の種類ごとに、"
              "過去の推移を延ばす方法と、人口に対する所持率を将来人口に乗じる"
              "方法の二つを併用し、単一の数値ではなく幅で示します。"
              "令和６年から令和８年までの手帳所持者数を村の台帳で確認したうえで"
              "確定し、見込量と突き合わせます。")

def blk_rinen(doc):
    band(doc, "６", "計画の基本理念")
    box(doc, [RINEN], fill=NOTE_FILL, border=ACCENT, size=10.5, bold_first=True)
    para(doc,
         "障がいの有無に関係なく、地域に暮らす人々がお互いに支え合いながら共に生きる"
         "「地域共生社会」の理念を基本としています。本村は、特性の異なる多様な地域が"
         "手を取り合い、各々の個性を磨きながら生活しています。これを踏まえ、人格や個性、"
         "価値観を認め合いながら、誰もが自分らしく生活できる地域社会の実現を基本理念とします。",
         space_after=4)

def blk_mokuhyou(doc):
    band(doc, "７", "基本目標")
    para(doc, "基本理念を実現するため、次の３つの目標を設定し、各種施策を推進します。")
    for title, text in MOKUHYOU:
        sub(doc, title)
        para(doc, text, space_after=4)


def blk_shinki(doc):
    """第８期計画で新たに加える施策の一覧。

    村の方針資料は「新規の追加施策」が一覧で見えることを求めている。
    計画本文では第３章４（表３-３）に置き、第１章４（４）で指している。
    概要版も同じ一覧を載せる。出所は正本と同じ（SHINKI）であり、
    概要版だけが別の一覧を持つことがないようにしている。
    """
    band(doc, "８", "本計画で新たに加える施策")
    para(doc,
         "本計画は第７期計画・第３期計画を引き継ぐものですが、"
         "国の基本指針の見直し、本計画の期間中に動き出す制度、"
         "及びアンケート調査で把握した課題に応えるため、"
         f"次の{len(SHINKI)}件を新たに加えます。"
         "詳しい内容は計画本文の第３章４（表３-３）及び"
         "各章をご覧ください。", size=SMALL_PT, space_after=3)
    table(doc, SHINKI_HYO, [5100, 5106], size=8)
    box(doc, [
        "このうち本村にとって特に重いもの",
        "・冬季の移動と地域生活の継続（除雪、通院・通所の確保）",
        "・災害時にサービスを止めないための備え"
        "（避難行動要支援者名簿、個別避難計画、"
        "医療的ケアに必要な電源の確保）",
        "・相談支援の体制"
        "（基幹相談支援センターの広域での設置、"
        "のぞまないセルフプランの解消）",
        "・障がいのある児童とその家族への支援"
        "（医療的ケア児等への支援、家族の休息）",
        "・18歳・22歳・65歳の到達時の引継ぎ",
    ], fill=NOTE_FILL, border=ACCENT, size=SMALL_PT, bold_first=True)

def blk_taikei(doc):
    band(doc, "９", "継承する施策体系と第８期計画で加わる視点")
    para(doc,
         "基本理念及び基本目標のもとに、第４次北塩原村障がい者計画が定める７つの基本施策を"
         "継承します。第８期計画で新たに加わる視点は、次のとおり各施策に接続します。",
         size=SMALL_PT)
    table(doc, shisaku_rows(), [1500, 5706, 3000], size=8)
    note(doc, "【要確認】目標値は第４次北塩原村障がい者計画が令和11年度を目標年度として"
              "定めたものであり、本計画の策定にあわせて中間見直しを行います。"
              "【要確認】を付した指標は次のとおりです。"
              "①差別や偏見は前回が３択・今回が頻度の設問で比較できないため、"
              f"今回の値（{_anke_pc('差別を感じた')}）は参考値とします。"
              "③満足している人の割合は前回がサービス利用者のみ（n＝25）の２択"
              "であったため、前回と同じくサービスを利用している回答者を分母とした"
              f"{_anke_pc('満足（利用者）')}を参考値とします（無回答を除く"
              f"回答者全員を分母とすると{_anke_pc('満足（無回答を除く全員）')}）。"
              "⑥個別避難計画は基準値（令和５年度）が１件で、現在の件数は"
              "村の資料どうしで１件と実質０件に分かれており、確認しています。"
              "⑦外出の目的は本調査に設問がなく測れないため、指標は取り下げず、"
              "測り方を村の実績（生涯学習・スポーツ・文化の事業への参加の実績等）"
              "に置き換えます。")


def blk_seika(doc):
    band(doc, "10", "令和11年度に向けた成果目標")
    para(doc,
         "障がいのある人の自立支援の観点から、地域生活への移行や就労支援等の課題に対応するため、"
         "令和11年度を目標年度として、国の基本指針に即し、次の８項目について成果目標を設定します。"
         "第７期計画までは７項目でしたが、第８期計画では（７）が新設され８項目になりました。")
    table(doc, SEIKA, [2400, 3400, 4406], size=8)
    note(doc, "【確認中】就労定着支援の目標値（２人）を据え置くか、"
              "施設入所者お一人ずつの意向の確認の状況などを村に確認しています。"
              "目標値は北塩原村障がい者自立支援協議会での審議を経て確定します。")


def blk_mikomi(doc):
    band(doc, "11", "障がい福祉サービス等の見込量")
    bullets(doc, [
        "見込量は、実際に使われた量をそのまま先に延ばしたものではありません。"
        "本村は対象となる方が少なく一人の増減が大きく表れるため、"
        "全国の伸び率や過去の実績の延長によらず、"
        "令和８年度末に支援を必要とする方を起点に一人ずつ積み上げます。",
        "必要な量と、事業所の定員・空き・人材・送迎・冬季の対応等を当てはめた"
        "確保できる量とを分けて把握し、その差（未充足の量）を毎年度の点検評価で"
        "確かめます。",
        "介護保険に相当するサービスがない共同生活援助等は、65歳到達後も利用できます。",
    ], size=SMALL_PT)
    para(doc, "＜主なサービスの利用者数（令和７年度の実績→令和11年度の見込量）＞",
         size=SMALL_PT, bold=True)
    table(doc, mikomi_rows(), [3053, 2050, 3053, 2050], size=8,
          aligns=[None, "center", None, "center"])
    note(doc, "【確認中】人数は１か月当たりです。令和７年度の実績は村の資料による"
              "もので、令和９〜11年度の見込量は実績との離れを村に確認しています"
              "（見込量を置き直すかは村が判断します）。"
              "「―」は令和７年度に実績のないサービスです"
              "（就労選択支援は令和７年10月に始まった制度です）。")


def blk_housaku(doc):
    """12 方策。見込量の直後に置く。"""
    band(doc, "12", "必要な見込量を確保するための方策")
    para(doc,
         "見込量を定めるだけではサービスは確保できません。"
         "村内の事業所が少ないという本村の事情を踏まえ、"
         "次の方策を進めます。", size=SMALL_PT, space_after=3)
    for midashi, text in HOUSAKU:
        p = para(doc, space_after=2, size=SMALL_PT)
        _run(p, f"■　{midashi}　", size=SMALL_PT, bold=True, color=DARK)
        _run(p, text, size=SMALL_PT)
    para(doc, "", size=4, space_after=0, line=5)


def blk_ikou(doc):
    band(doc, "13", "年齢到達に伴うサービスの移行")
    para(doc, IKOU_LEAD, size=SMALL_PT, space_after=3)
    figure(doc, "11_ライフコース_概要版.png",
           "図　ライフコースと制度の移行",
           "資料：北塩原村（本計画第５章５及び第７章により作成）",
           width_mm=158)
    para(doc, IKOU_NOTE, size=SMALL_PT, space_after=4)

def blk_renkei(doc):
    band(doc, "14", "他計画との連携")
    para(doc,
         "本計画は、村の関連計画と整合を図り、連携して施策を推進します。",
         size=SMALL_PT, space_after=3)
    table(doc, RENKEI, [2800, 1900, 5506], size=8)


def blk_koken(doc):
    band(doc, "15", "成年後見制度の利用促進")
    para(doc,
         "本村では、認知症、知的障がい、その他の精神上の障がいがあることにより、"
         "日常生活を支える様々な行為や買い物、財産管理が難しい事例が見られます。"
         "高齢化の進展を踏まえると、成年後見制度の必要性は今後も高まっていくことが見込まれます。",
         size=SMALL_PT)
    para(doc,
         "令和４年７月より中核機関「会津権利擁護・成年後見センター」を会津圏域11市町村で設置し、"
         "制度の周知啓発、広報活動、相談対応を行っています。令和６年度からは市民後見人養成事業が"
         "追加されました。本計画期間においても同センターとの連携を継続します。"
         "なお、令和８年６月24日に民法等の一部を改正する法律"
         "（令和８年法律第45号）が公布され、後見及び保佐の制度を廃止して"
         "補助の制度に基本的に一元化されます。"
         "施行の日は公布の日から２年６月を超えない範囲内において"
         "政令で定める日とされており、本計画の期間中に施行されます。"
         "施行の内容に応じて、中核機関を通じて村の運用を見直します。",
         size=SMALL_PT, space_after=4)

def blk_suishin(doc):
    band(doc, "16", "計画の推進体制")
    para(doc,
         "計画に定める事項については、PDCAサイクルにより進行管理を行います。"
         "北塩原村障がい者自立支援協議会において、成果目標及び活動指標の実績を"
         "少なくとも年１回把握して分析・評価し、その結果を村の広報及び"
         "ホームページで公表します。確保できなかった量とその理由も公表します。",
         size=SMALL_PT)
    table(doc, PDCA, [1800, 8406], size=8, fill_first=BAND)

    band(doc, "", "今後の予定")
    table(doc, SCHEDULE, [2400, 7806], size=8, fill_first=BAND)

    box(doc, [
        "この概要版についてのご意見をお寄せください",
        "本計画の素案は、北塩原村障がい者自立支援協議会での審議を経て、"
        f"{wareki(PABU_KAISHI)}から{wareki(PABU_SHURYO)}まで"
        f"（{PABU_NISSU}日間）パブリックコメント（村民意見公募）を"
        "実施する予定です。"
        "素案の全文及びこの概要版は、村ホームページ及び村役場窓口で"
        "ご覧いただけます。",
        "【お問合せ・ご意見の提出先】北塩原村役場　保健福祉課　福祉係"
        "（TEL）0241-23-3113",
    ], fill=NOTE_FILL, border=ACCENT, size=SMALL_PT, bold_first=True)


# ============================================================
def setup_section(doc):
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.top_margin = Mm(15)
    section.bottom_margin = Mm(15)
    section.left_margin = Mm(15)
    section.right_margin = Mm(15)
    # 既定スタイルを本文書体に合わせる
    style = doc.styles["Normal"]
    style.font.name = FONT
    style.font.size = Pt(BODY_PT)
    rpr = style.element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = _el("rFonts")
        rpr.insert(0, rf)
    rf.set(qn("w:eastAsia"), FONT)
    rf.set(qn("w:ascii"), FONT)
    rf.set(qn("w:hAnsi"), FONT)


# ブロックの並び。改ページは flow() が計算して入れる。
# 手で置いていた改ページを外したのは、文章を1行足すたびに
# どの頁が溢れるかを人が確かめ直すことになっていたためである。
BLOCKS = [
    ("表紙", blk_hyoushi),
    ("1 計画策定の背景・趣旨", blk_shushi),
    ("2 計画の位置づけ", blk_ichizuke),
    ("3 計画の期間", blk_kikan),
    ("4 国の基本指針の見直しのポイント", blk_shishin),
    ("5 本村の障がい者の現状", blk_genjou),
    ("6 計画の基本理念", blk_rinen),
    ("7 基本目標", blk_mokuhyou),
    ("8 本計画で新たに加える施策", blk_shinki),
    ("9 継承する施策体系", blk_taikei),
    ("10 令和11年度に向けた成果目標", blk_seika),
    ("11 障がい福祉サービス等の見込量", blk_mikomi),
    ("12 必要な見込量を確保するための方策", blk_housaku),
    ("13 年齢到達に伴うサービスの移行", blk_ikou),
    ("14 他計画との連携", blk_renkei),
    ("15 成年後見制度の利用促進", blk_koken),
    ("16 計画の推進体制", blk_suishin),
]


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    doc = docx.Document()
    setup_section(doc)
    pages, oki = flow(doc, BLOCKS)
    doc.save(OUT_FILE)

    print("  ブロックの置き場所:")
    for name, pg, nokori in oki:
        print(f"    p.{pg}  {name}（置いた後 {nokori:.0f}行）")

    print(f"作成: {OUT_FILE}")
    print(f"  ページ数: {len(pages)}")
    over = []
    for i, lines in enumerate(pages, 1):
        mark = "" if lines <= LINES_PER_PAGE else "　← 収まりません"
        if mark:
            over.append(i)
        print(f"  p.{i}  推定 {lines:.0f} 行 / 目安 {LINES_PER_PAGE} 行{mark}")
    return over, len(pages)


def honpon_text():
    """移植後の正本の全文（本文と表）を返す。

    令和8年10月に加えた3節（方策・年齢到達・他計画との連携）は
    正本にしかない。思い込みで要約すると、計画に書いていないことを
    概要版が書くことになるため、必ず現物と突き合わせる。
    """
    path = f"{OUT_DIR}/北塩原村_計画素案_正本_移植後.docx"
    if not os.path.exists(path):
        raise SystemExit(
            f"正本が見つかりません: {path}\n"
            "build_kitashiobara_honpon.py を先に実行してください。")
    d = docx.Document(path)
    out = [p.text for p in d.paragraphs]
    for t in d.tables:
        for row in t.rows:
            out += [c.text for c in row.cells]
    return "\n".join(out)


def verify(over, npages):
    """概要版の数値が計画素案と合っていることを確かめる。

    概要版は素案の要約であり、数値が食い違うと村の内部で判断が割れる。
    見込量は素案と同じ入口（build_kitashiobara_mikomiryo.py）から読んでいるが、
    成果目標とKPIは概要版が独自に持つため、ここで素案の現物と突き合わせる。
    """
    ng = [f"p.{i} が1ページに収まりません" for i in over]
    gaiyou = docx.Document(OUT_FILE)
    body = "\n".join([p.text for p in gaiyou.paragraphs]
                     + [c.text for t in gaiyou.tables for r in t.rows
                        for c in r.cells])
    import re as _re
    hon_doc = _honpon_doc()

    # 1 見込量（令和11年度）と実績（令和７年度）が正本 第5章の表の値であること
    #   以前は第1次概算と旧い計画素案を照合していたため、正本の見込量と
    #   ６サービスで食い違っていても通っていた（横断レビュー RT-05・RT-06）。
    jis, mik = honpon_hyou("実績"), honpon_hyou("見込")
    for n in MIKOMI_PICK:
        r7 = jis.get(n, [None, None, None])[1]
        need = f"{'―' if r7 is None else f'{r7}人'}→{mik[n][2]}人"
        if need not in body:
            ng.append(f"概要版の{n}が正本の表と合わない（{need}）")
    for furui in ("第1次概算", "第１次概算", "給付費は約", "億"):
        if furui in body:
            ng.append(f"概要版に取り下げた記述が残っている: {furui}")

    # 2 成果目標が正本 第4章2 の表の行と同じであること
    gyou = []
    for t in hon_doc.tables:
        for r in t.rows:
            cs = [c.text.replace("\n", "").strip() for c in r.cells]
            if len(cs) >= 2:
                gyou.append(cs)
    seika = "\n".join(str(x) for r in SEIKA for x in r)
    for label, atai, go in SEIKA_SHOUGOU:
        if not any(label in g[0] and g[1].startswith(atai) for g in gyou):
            ng.append(f"正本 第4章2 の表に無い値: {label}＝{atai}")
        if go not in seika:
            ng.append(f"概要版の成果目標に無い: {go}（{label}）")
    for furui in ("令和４年度末", "セルフプラン率", "０人（令和４年度末）",
                  "年１〜２人", "（令和４年度）"):
        if furui in seika:
            ng.append(f"概要版の成果目標に直す前の値が残っている: {furui}")

    # 3 第４次障がい者計画は中間見直し（正本 第1章2・第3章）
    if "改定対象外" in body or "対象外\n（第４次" in body:
        ng.append("概要版に「改定対象外」が残っている（正本は中間見直し）")
    if "中間見直し" not in body:
        ng.append("概要版に第４次計画の中間見直しの記述が無い")

    # 4 成年後見制度の利用促進に関する計画（令和９〜13年度。正本 第6章）
    if "令和９年度〜令和13年度" not in body:
        ng.append("概要版の計画の期間に成年後見の計画（令和９〜13年度）が無い")

    # 5 アンケートの割合が回収40件の値であること
    import kitashiobara_anke as _A
    for nm in _A.M:
        kyu, shin = _A.pc(nm, kyu=True), _A.pc(nm)
        if kyu != shin and kyu in body and nm != "外出_冬季_裏磐梯":
            ng.append(f"概要版に38件の割合が残っている: {nm} {kyu}")
    if _re.search(r"有効回答38件|n＝38|n=38", body):
        ng.append("概要版に38件の表記が残っている")

    # 6 推進体制の評価の主体が正本（自立支援協議会）と同じであること
    if "自立支援協議会において" not in body:
        ng.append("概要版の推進体制の評価の主体が正本と違う")
    if "評価は保健福祉課が行い" in body:
        ng.append("概要版に評価の主体を保健福祉課とする記述が残っている")

    # 7 令和8年10月に加えた3節が、正本の現物に裏づけられていること
    honpon = honpon_text()

    #   7a 方策は正本 第5章2 の（1）〜（15）を束ねたもの。
    #      束ねた先の項が正本にあることを見出しで確かめる。
    for midashi in ("（１）住民への情報提供等", "（２）総合的な相談支援",
                    "（３）一般就労への移行等の推進",
                    "（４）広域連携によるサービス提供体制の確保",
                    "（５）障がい児支援の提供体制の確保",
                    "（６）サービスの質の向上と人材の確保",
                    "（７）事業所の受入余力の把握", "（８）住まいの確保",
                    "（９）供給が不足する場合の県への意見申出",
                    "（10）中山間・人口減少地域の制度の活用",
                    "（11）災害時にサービスを止めないための備え",
                    "（12）虐待の予防から再発の防止まで",
                    "（13）経営基盤を供給の制約として見る",
                    "（14）冬季の地域生活の継続",
                    "（15）医療へのアクセス"):
        if midashi not in honpon:
            ng.append(f"正本 第5章2 に項が見つからない: {midashi}。"
                      "HOUSAKU の束ね方を見直してください")

    #   7b 他計画との連携は正本 第7章の4節に対応すること
    for midashi in ("１　第３次健康21・北塩原村グッドヘルスプランとの連携",
                    "２　北塩原村第五次総合振興計画との連携",
                    "３　北塩原村高齢者福祉計画・介護保険事業計画との連携",
                    "４　北塩原村こども・子育て計画との連携"):
        if midashi not in honpon:
            ng.append(f"正本 第7章に節が見つからない: {midashi}。"
                      "RENKEI を見直してください")
    if len(RENKEI) - 1 != 4:
        ng.append(f"RENKEI の行数が第7章の節数と合わない: {len(RENKEI) - 1}")

    #   7c 年齢到達の記述が正本の記述と食い違っていないこと
    for need in ("５　年齢到達に伴うサービスの移行",
                 "（１）学校卒業に伴うサービスの移行",
                 "（２）65歳到達時のサービスの区分"):
        if need not in honpon:
            ng.append(f"正本 第5章5 に見つからない: {need}")

    #   7d 概要版が書いた事実が正本にもあること（言い切りの裏づけ）
    for need in ("介護保険に相当するサービスがないもの",
                 "地域包括支援センター",
                 "住宅セーフティネット",
                 "休日・夜間",
                 "特定地域サービス",
                 "中間見直し"):
        if need not in honpon:
            ng.append(f"概要版が書いた事実が正本にない: {need}")

    #   7e 概要版に貼った図が、図の生成器が作ったものであること
    if not os.path.exists(f"{OUT_DIR}/図表/11_ライフコース_概要版.png"):
        ng.append("概要版のライフコース図が無い")

    #   7f 人口と手帳所持者数が正本の現物と一致すること。
    #      概要版だけが古い数値のまま残ると、委員会で二つの数が並ぶ。
    for need in ("令和８年４月１日現在2,316人", "令和４年（2,535人）",
                 "身体障害者手帳が127人", "療育手帳が16人",
                 "精神障害者保健福祉手帳が38人", "全体で181人",
                 "令和５年４月現在156人", "43.1％"):
        if need not in honpon:
            ng.append(f"正本に見つからない数値: {need}。"
                      "概要版の第5節を見直してください")
    for mae, ato in (("114人", "127人"), ("12人", "16人"),
                     ("30人", "38人"), ("156人", "181人")):
        if [r for r in TECHOU[1:] if r[3] == mae and r[4] != ato]:
            ng.append(f"TECHOU の令和8年の値が正本と合わない: {mae}→{ato}")
    goukei = [sum(int(TECHOU[i][j].rstrip("人")) for i in (1, 2, 3))
              for j in range(1, 5)]
    if goukei != [int(TECHOU[4][j].rstrip("人")) for j in range(1, 5)]:
        ng.append(f"TECHOU の合計が内訳と合わない: {goukei}")

    #   8 工程が単一出所（kitashiobara_koutei.py）と一致すること。
    #     概要版は住民に配るものであり、パブリックコメントの日付が
    #     合っていなければ意見を出せない。
    if SCHEDULE is not GAIYOU_HYO:
        ng.append("概要版の工程表が kitashiobara_koutei.py のものでない")
    for need in (wareki(PABU_KAISHI), wareki(PABU_SHURYO),
                 f"{PABU_NISSU}日間"):
        if need not in body:
            ng.append(f"概要版にパブリックコメントの{need}がない")
    for furui in ("２月８日", "２月22日", "令和８年度中"):
        if furui in body:
            ng.append(f"概要版に従前の工程の記述が残っている: {furui}")

    #   9 新たに加える施策が正本（表３-３）と同じ出所であること。
    #     概要版だけが別の一覧を持つと、何が新しいかが食い違う。
    if len(SHINKI_HYO) != len(SHINKI) + 1:
        ng.append(f"新たに加える施策の行数が合わない: {len(SHINKI_HYO)}")
    for shisaku, _naka, ichi, _kihon, _go in SHINKI:
        if shisaku not in body:
            ng.append(f"概要版に新たに加える施策がない: {shisaku}")
        if shisaku not in honpon:
            ng.append(f"正本に新たに加える施策がない: {shisaku}")
    for need in ("表３-３", "第３章４"):
        if need not in body:
            ng.append(f"概要版が正本の{need}を指していない")
    if "第３章　第４次北塩原村障がい者計画の推進状況と" \
            "本計画で新たに加える施策" not in honpon:
        ng.append("正本の第3章の章題が案Ｃのものになっていない")

    #   10 手帳所持者数の見通しは、正本と同じく単一の値を示さないこと
    #      （正本 第2章2。村の台帳の確認を待つ。横断レビュー RT 指摘9）
    if "単一の数値ではなく幅" not in body or "単一の数値ではなく幅" not in honpon:
        ng.append("手帳所持者数の見通しの書き方が正本と違う")
    if _re.search(r"[0-9]+人から[0-9]+人と見込まれます", body):
        ng.append("概要版が令和11年度の手帳所持者数を数で示している（正本は未確定）")
    if "145人" in body:
        ng.append("概要版に取り下げた推計（145人）が残っている")

    if ng:
        print("自己点検 不合格:")
        for x in ng:
            print("  -", x)
        return False
    print(f"自己点検: 合格　見込量{len(MIKOMI_PICK)}件（令和７年度の実績と令和11年度）"
          f"・成果目標{len(SEIKA_SHOUGOU)}点を正本の表と照合／"
          "第４次計画は中間見直し／アンケートは回収40件／推進体制の主体／"
          f"方策{len(HOUSAKU)}件・他計画{len(RENKEI) - 1}件・"
          f"年齢到達3点・言い切り6点・人口と手帳8点を正本の現物と照合／"
          f"新たに加える施策{len(SHINKI)}件を正本と同じ出所から／"
          f"工程を kitashiobara_koutei.py と照合"
          f"（パブリックコメント {wareki(PABU_KAISHI)}〜"
          f"{wareki(PABU_SHURYO)}）／"
          "手帳の見通しは幅で示し数を出さない／"
          f"{npages}ページに収まる")
    return True


if __name__ == "__main__":
    import sys
    if not verify(*main()):
        sys.exit(1)
