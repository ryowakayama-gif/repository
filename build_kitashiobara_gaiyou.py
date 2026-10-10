# -*- coding: utf-8 -*-
"""
北塩原村 第8期障がい福祉計画・第4期障がい児福祉計画
素案 概要版ジェネレータ

出力: output/北塩原村_計画素案_概要版.docx（A4縦・6ページ）

現行計画の概要版（source/現行計画/第4次障がい者計画等_概要版.pdf、A4・4ページ）
の構成を踏襲する。ただし次の点が異なる。

  - 今回の改定対象は障がい福祉計画（第8期）と障がい児福祉計画（第4期）の
    2計画である。障がい者計画（第4次）は令和11年度までを計画期間としており
    改定対象外のため、その基本理念・基本目標・施策体系は「継承する内容」
    として整理する。
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

出所
  数値は output/北塩原村_計画素案.docx 及び kitashiobara_common.py と同じ
  出典から取り、村に確認中のものには【要更新】を付す。
  令和8年10月に加えた3節は正本（output/北塩原村_計画素案_正本_移植後.docx）
  にしかないため、その現物と突き合わせる。したがって本書は
  build_kitashiobara_honpon.py の後に実行する。
"""

import os

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor

from kitashiobara_common import OUT_DIR, SERVICES_ADULT, SERVICES_CHILD

OUT_FILE = f"{OUT_DIR}/北塩原村_計画素案_概要版.docx"

FONT = "BIZ UDPゴシック"
DARK = "1F3864"        # 章見出しの帯（計画素案の表頭と同色）
ACCENT = "FF7F0E"      # 障がい分野の識別色（build_excel.py の規約）
BAND = "DDEBF7"        # 小見出しの帯
NOTE_FILL = "FFF2CC"   # 【要更新】ボックス
BOX_FILL = "F2F2F2"
BORDER_COLOR = "AAAAAA"

PAGE_W = 10206         # A4（210mm）－左右余白15mm ＝ 180mm（twips）

# 1ページに収まる行数の目安（A4縦・上下余白15mm・9pt・行送り固定12pt）
LINES_PER_PAGE = 57
CHARS_PER_LINE = 54    # 9pt 全角で1行に入るおよその文字数


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


def _run(para, text, size=9, bold=False, color=None):
    run = para.add_run(text)
    run.font.name = FONT
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = _el("rFonts")
        rpr.insert(0, rf)
    rf.set(qn("w:eastAsia"), FONT)
    return run


def para(doc, text="", size=9, bold=False, align=None, space_after=2,
         indent=0, color=None, line=12):
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
    p.paragraph_format.line_spacing = Pt(12)
    _run(p, text, size=9.5, bold=True, color="1F3864")
    _tbl_width(t)
    _no_borders(t)
    para(doc, "", size=3, space_after=0, line=4)
    return t


def bullets(doc, lines, size=9):
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


def table(doc, rows, widths, header=True, size=8.5, aligns=None, fill_first=None):
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
            p.paragraph_format.line_spacing = Pt(11)
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


def box(doc, lines, fill=BOX_FILL, border=BORDER_COLOR, size=9, bold_first=False):
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
        p.paragraph_format.line_spacing = Pt(12)
        _run(p, line, size=size, bold=(bold_first and i == 0))
    _tbl_width(t)
    _borders(t, border)
    para(doc, "", size=4, space_after=0, line=5)
    return t


def note(doc, text):
    return box(doc, [text], fill=NOTE_FILL, border="C9A03B", size=8)


def page_break(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = Pt(1)
    _run(p, "", size=1).add_break(WD_BREAK.PAGE)


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
    ["区分", "障害者計画", "障害福祉計画", "障害児福祉計画"],
    ["根拠法", "障害者基本法\n第11条第３項", "障害者総合支援法\n第88条第１項",
     "児童福祉法\n第33条の20第１項"],
    ["内容", "障がい者施策の基本的方向性について定める計画",
     "障害福祉サービス等の見込量とその確保策を定める計画",
     "障害児通所支援等の提供体制と確保策を定める計画"],
    ["今回の改定", "対象外\n（第４次計画を継続）", "第８期として策定", "第４期として策定"],
]

KIKAN = [
    ["計画", "計画期間", "今回の取扱い"],
    ["第４次北塩原村障がい者計画", "令和６年度〜令和11年度（６年間）",
     "改定対象外。基本理念・基本目標・施策体系を本計画に継承"],
    ["第８期北塩原村障がい福祉計画", "令和９年度〜令和11年度（３年間）", "今回策定"],
    ["第４期北塩原村障がい児福祉計画", "令和９年度〜令和11年度（３年間）", "今回策定"],
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
# 令和6年・令和7年の値は手元にないため列を置かない。
# 以前は令和11年の中位推計（合計145人）を併記していたが、
# その推計は令和5年までの減少を延ばしたものであり、
# 令和8年の実績181人（増加）と前提が合わないため外した。
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

SHISAKU = [
    ["基本施策", "第８期計画で加わる視点", "目標値（現状→令和11年度）"],
    ["①　啓発・広報",
     "合理的配慮の提供義務化の周知、意思疎通支援従事者の養成・派遣体制の整備、"
     "障がい者等に対する虐待の防止",
     "差別や偏見を感じている人の割合\n25.3％→15％【要確認】"],
    ["②　保健・医療",
     "精神障害にも対応した地域包括ケアシステムの構築、医療的ケア児等への支援、"
     "高次脳機能障害のある方への支援",
     "特定健康診査の受診率\n51.3％→60％"],
    ["③　福祉",
     "地域生活支援拠点等の年１回以上の検証、強度行動障害を有する方の支援ニーズの把握、"
     "のぞまないセルフプランの解消、人口減少地域におけるサービスの維持・確保",
     "福祉サービスに満足している人の割合\n52％→60％【要確認】"],
    ["④　教育・育成",
     "児童発達支援センターの４つの中核機能の確保、インクルージョン推進のための"
     "協議の場の設置、こども家庭センターとの連携",
     "医療的ケア児の受入れ体制の整備\n無→有"],
    ["⑤　雇用・就業",
     "就労選択支援の利用促進、法定雇用率2.7％への引上げを踏まえた事業主啓発",
     "障がい者就労施設等からの物品調達件数\n０件→３件"],
    ["⑥　生活環境",
     "災害時における障害福祉サービス提供の確保、個別避難計画の作成促進、"
     "冬季・積雪時の避難支援体制の整備",
     "公共施設トイレの設置率\n56.8％→100％\n個別避難計画\n１件→10件【要確認】"],
    ["⑦　スポーツ・文化",
     "スポーツ・健康増進活動による社会参加等の促進、"
     "地域共生社会の実現に向けた取組の一層の推進",
     "外出の目的が趣味・スポーツ等である人の割合\n28.1％→40％【要確認】"],
]

SEIKA = [
    ["成果目標", "現状（基準）", "令和11年度の目標値"],
    ["（１）福祉施設の入所者の\n　　　地域生活への移行",
     "施設入所者数　４人\n（令和４年度末。令和７年度末の実数は確認中）",
     "地域生活移行者数　０人\n施設入所者削減見込み　０人\n（入所者は重度の方が中心のため、個別の移行"
     "可能性を自立支援協議会で検討します）"],
    ["（２）精神障害にも対応した\n　　　地域包括ケアシステムの構築",
     "精神病床に１年以上入院している方　５人\n（第７期計画時点）",
     "退院者数　１人\n心のサポーター養成講座の実施【新規】\n"
     "Ｋ６等によるこころの状態を把握する機会の提供【新規】\n"
     "長期入院患者の地域生活への移行に伴う基盤整備量【県の算定値待ち】"],
    ["（３）福祉施設から\n　　　一般就労への移行等",
     "一般就労移行者数・就労定着支援利用者数\n（令和６年度実績を確認中）",
     "一般就労移行者数　１人\n就労定着支援事業利用者数　１人\n就労選択支援利用者数　年１〜２人【新規】"],
    ["（４）障がい児支援の\n　　　提供体制の整備等",
     "児童発達支援センター　０か所\n保育所等訪問支援事業所　０か所\n医療的ケア児　１人",
     "近隣市町村との連携により４つの中核機能へのアクセスを確保（０か所）\n"
     "医療的ケア児等コーディネーター　１人\nインクルージョン推進の協議の場を会津北部圏域で設置【新規】\n"
     "医療的ケア児等に関する協議の場を圏域で設置【新規】"],
    ["（５）地域生活支援の充実",
     "会津北部地域生活支援拠点（４町村共同）　１か所\nコーディネーター　１人\n"
     "強度行動障害のある方　０人（令和４年度末）",
     "拠点の運用状況の検証　年１回以上\n強度行動障害を有する方の支援ニーズを圏域で把握【新規】"],
    ["（６）相談支援体制の\n　　　充実・強化等",
     "基幹相談支援センター　未設置\n（県の第７期計画は県全体59市町村・"
     "会津圏域13市町村での設置を目標とするが、圏域の設置は２か所）\n"
     "協議会専門部会　０か所\nセルフプラン率　０％",
     "基幹相談支援センター　１か所（４町村広域）\n協議会専門部会　２か所\n"
     "のぞまないセルフプラン　０件【新規】"],
    ["（７）障がい福祉人材の確保・定着\n　　　及び生産性向上【新規】",
     "県が実施する研修への村職員の参加　１人\n（令和４年度）",
     "県研修への参加　２人以上\n県のワンストップ窓口等の支援策を村内事業所へ周知"],
    ["（８）障がい福祉サービス等の\n　　　質の向上に係る体制の構築",
     "県が実施する研修への村職員の参加　１人",
     "県研修への参加　２人\n村内対象事業所の情報公表制度　公表率100％【新規】"],
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
    "障がい福祉サービスは、一定の年齢に達することにより、対象となる制度が"
    "変わります。本計画では、18歳と65歳の二つの節目について、"
    "何が変わり何が続くのかを整理しました。"
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

SCHEDULE = [
    ["時期", "内容"],
    ["令和８年８月〜９月", "アンケート調査の実施（実施済み）"],
    ["令和８年11月", "北塩原村障がい者自立支援協議会（計画素案の審議）"],
    ["令和８年度中", "パブリックコメントの実施"],
    ["令和８年度中", "庁議（計画案の決定）"],
    ["令和９年３月", "計画の確定・公表（予定）"],
]

# 概要版に載せるサービス。計画期間に値が動くもの（生活介護・就労継続支援Ｂ型・
# 共同生活援助・児童発達支援・放課後等デイサービス）と、新規に立つもの
# （就労選択支援・短期入所・保育所等訪問支援）を落とさない。
# 名称は第1次概算（計画素案の見込量表）のものに合わせる。
MIKOMI_PICK = [
    "居宅介護", "生活介護", "自立訓練（生活訓練）", "就労選択支援",
    "就労移行支援", "就労継続支援Ａ型", "就労継続支援Ｂ型", "就労定着支援",
    "短期入所（福祉型）", "共同生活援助", "施設入所支援", "計画相談支援",
    "児童発達支援", "放課後等デイサービス", "保育所等訪問支援",
    "障がい児相談支援",
]

# kitashiobara_common.py の前計画マスタは就労継続支援を半角のA・Bで持つ。
# 第1次概算は全角のＡ・Ｂで持つため、前計画の値を引くときだけ読み替える。
ZEN_ALIAS = {"就労継続支援Ａ型": "就労継続支援A型",
             "就労継続支援Ｂ型": "就労継続支援B型"}


def load_mikomi():
    """計画素案と同じ算定結果（第1次概算）を読む。

    build_kitashiobara_mikomiryo.py の build_rows を呼ぶ。計画素案
    （build_kitashiobara_kosshi_rev.py の load_mikomiryo）と同じ入口であり、
    素案と概要版で値が二重管理にならない。
    """
    import os
    import runpy
    from kitashiobara_common import JIDO_KYUFU, KAIGO_KYUFU
    mod = runpy.run_path(
        os.path.join(os.path.dirname(os.path.abspath(__file__)),
                     "build_kitashiobara_mikomiryo.py"))
    rows = mod["build_rows"](mod["ADULT_PLAN"], KAIGO_KYUFU)
    rows += mod["build_rows"](mod["CHILD_PLAN"], JIDO_KYUFU)
    return {r["name"]: r for r in rows}


def load_kyufu():
    """給付費の見込み（3か年計と各年度）を見込量算定ブックから読む。

    概要版に数字を書き写さず、算定の正本から読む。素案 第5章12 の財源構成と
    同じ 06_給付費の見込み シートを入口にしている。
    """
    import os
    from openpyxl import load_workbook
    path = f"{OUT_DIR}/北塩原村_サービス見込量算定.xlsx"
    if not os.path.exists(path):
        raise SystemExit(
            f"見込量算定ブックが見つかりません: {path}\n"
            "build_kitashiobara_service_estimate.py を先に実行してください。")
    wb = load_workbook(path, data_only=True)
    ws = wb["06_給付費の見込み"]
    for row in ws.iter_rows(values_only=True):
        if row and str(row[0]).strip() == "合計":
            per, total = int(row[1]), int(row[4])
            return per, total
    raise LookupError("06_給付費の見込み シートに合計の行がありません")


def _oku_man(yen):
    """円を「1億7,286万円」の形にする（万円未満は四捨五入）。"""
    man = round(yen / 10000)
    oku, man = divmod(man, 10000)
    if oku:
        return f"{oku}億{man:,}万円"
    return f"{man:,}万円"


def mikomi_rows():
    """主なサービスの前計画（令和8年度）と見込み（令和11年度）を2列組にする。"""
    zen = {name: users for name, _u, _q, users, *_rest in SERVICES_ADULT}
    zen.update({name: users for name, _u, _q, users, *_rest in SERVICES_CHILD})
    mikomi = load_mikomi()
    missing = [n for n in MIKOMI_PICK if n not in mikomi]
    if missing:
        raise LookupError(f"第1次概算に無いサービスです: {missing}")

    pairs = []
    for n in MIKOMI_PICK:
        now = zen.get(ZEN_ALIAS.get(n, n))
        if now is None:
            raise LookupError(f"前計画マスタに無いサービスです: {n}")
        to = mikomi[n]["users"]
        moto = "―" if now == 0 else f"{now}人"
        pairs.append((n, f"{moto}→{to}人"))
    half = (len(pairs) + 1) // 2
    left, right = pairs[:half], pairs[half:]
    rows = [["サービス", "令和８年度→令和11年度", "サービス",
             "令和８年度→令和11年度"]]
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
                # 1行＝12pt＝152,400EMU（行送りを固定しているため）
                pages[-1] += sum(int(x.get("cy")) for x in zu) / 152400
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
def page1(doc):
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
         "今回の改定の対象ではありません。その基本理念・基本目標及び施策体系は、"
         "本計画においても共有する理念・目標・体系として継承します。", space_after=4)

    sub(doc, "「障害」と「障がい」の表記について")
    para(doc, HYOKI_TEXT, space_after=4)

    band(doc, "２", "計画の位置づけ")
    para(doc, "本計画は、以下の法律に基づいて策定する法定計画です。")
    table(doc, ICHIZUKE, [1250, 2985, 2985, 2986], aligns=["center", None, None, None],
          fill_first=BAND)

    band(doc, "３", "計画の期間")
    table(doc, KIKAN, [3000, 2700, 4506], fill_first=BAND)

    band(doc, "４", "国の基本指針の見直しのポイント")
    para(doc,
         "第８期障害福祉計画及び第４期障害児福祉計画に係る国の基本指針は令和８年３月31日に"
         "告示され、主な見直し事項として次の点が示されています。")
    bullets(doc, SHISHIN_POINTS, size=8.5)


def page2(doc):
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
    note(doc, "【要更新】令和11年度の手帳所持者数の見通しは、"
              "令和５年までの実績（156人）をもとに減少（145人程度）と置いていましたが、"
              "令和８年４月の実績が181人と増加に転じたため、前提が成り立ちません。"
              "令和６年度・令和７年度の手帳所持者数を村から受領し、"
              "第２次算定で推計をやり直します。")

    band(doc, "６", "計画の基本理念")
    box(doc, [RINEN], fill=NOTE_FILL, border=ACCENT, size=10.5, bold_first=True)
    para(doc,
         "障がいの有無に関係なく、地域に暮らす人々がお互いに支え合いながら共に生きる"
         "「地域共生社会」の理念を基本としています。本村は、特性の異なる多様な地域が"
         "手を取り合い、各々の個性を磨きながら生活しています。これを踏まえ、人格や個性、"
         "価値観を認め合いながら、誰もが自分らしく生活できる地域社会の実現を基本理念とします。",
         space_after=4)

    band(doc, "７", "基本目標")
    para(doc, "基本理念を実現するため、次の３つの目標を設定し、各種施策を推進します。")
    for title, text in MOKUHYOU:
        sub(doc, title)
        para(doc, text, space_after=4)


def page3(doc):
    band(doc, "８", "継承する施策体系と第８期計画で加わる視点")
    para(doc,
         "基本理念及び基本目標のもとに、第４次北塩原村障がい者計画が定める７つの基本施策を"
         "継承します。第８期計画で新たに加わる視点は、次のとおり各施策に接続します。")
    table(doc, SHISAKU, [1500, 5706, 3000], size=8)
    note(doc, "【要確認】目標値は第４次北塩原村障がい者計画が令和11年度を目標年度として"
              "定めたものです。本計画の策定に係るアンケート調査の結果と突き合わせたところ、"
              "【要確認】を付した４指標は扱いを決める必要があります。"
              "①差別や偏見は前回が３択・今回が頻度の設問で比較できません。"
              "③満足している人の割合は前回がサービス利用者のみ（n＝25）で、"
              "今回の有効回答38件を分母とする値（34.2％）とは母集団が異なります。"
              "⑦外出の目的は本調査に設問がなく測定できません。"
              "⑥個別避難計画は現状値１件に対し令和８年８月の村の確認が０件で食い違います。")

    band(doc, "９", "令和11年度に向けた成果目標")
    para(doc,
         "障がいのある人の自立支援の観点から、地域生活への移行や就労支援等の課題に対応するため、"
         "令和11年度を目標年度として、国の基本指針に即し、次の８項目について成果目標を設定します。"
         "第７期計画までは７項目でしたが、第８期計画では（７）が新設され８項目になりました。")
    table(doc, SEIKA[:7], [2400, 3400, 4406], size=8)


def page4(doc):
    sub(doc, "９　令和11年度に向けた成果目標（続き）", fill=BAND)
    table(doc, [SEIKA[0]] + SEIKA[7:], [2400, 3400, 4406], size=8)
    note(doc, "【要更新】現状（基準）の数値は、村に令和７年度末時点の実績を確認しているところです。"
              "目標値は北塩原村障がい者自立支援協議会での審議を経て確定します。")

    band(doc, "10", "障がい福祉サービス等の見込量")
    kyufu_per, kyufu_total = load_kyufu()
    bullets(doc, [
        "給付実績が完結した直近の年度である令和７年度を基準年度とし、その水準の維持を"
        "基本に、成果目標のあるサービスは目標に合わせて補正しました。"
        "国が示す推計方法（変化率の平均、人口当たり利用率）は検証に用います。"
        "１人単位の積上げは村の個別データ受領後の第２次算定で行います。",
        "本村は全域が過疎地域であるため、地域差の是正に関する算定方法"
        "（基本指針 別表第五）の適用対象外です。",
        "国の基本指針が「設定するものとする」とした要件により、生活介護・"
        "就労継続支援Ｂ型・施設入所支援は継続入所者を除き、障がい児通所支援は"
        "保育所等における障がい児の受入れ体制を踏まえて設定します。",
        "介護保険に相当するサービスがない共同生活援助等は、65歳到達後も利用できます。",
        f"第８期の３か年の給付費は約{_oku_man(kyufu_total)}"
        f"（各年度約{_oku_man(kyufu_per)}）の見込みです。",
    ], size=8.5)
    para(doc, "＜主なサービスの利用者数（令和８年度の前計画→令和11年度の見込み）＞",
         size=8.5, bold=True)
    table(doc, mikomi_rows(), [3053, 2050, 3053, 2050], size=8,
          aligns=[None, "center", None, "center"])
    note(doc, "【第1次概算】令和９〜11年度の値は令和７年度の給付実績に基づく概算です。"
              "村の個別データを受領したうえで第２次算定を行い確定します。"
              "「―」は現行計画に計上がないサービスです。")


def page4b(doc):
    """p.4 の後半（11 方策）。見込量の直後に置く。"""
    band(doc, "11", "必要な見込量を確保するための方策")
    para(doc,
         "見込量を定めるだけではサービスは確保できません。"
         "村内の事業所が少ないという本村の事情を踏まえ、"
         "次の方策を進めます。", size=8.5, space_after=3)
    for midashi, text in HOUSAKU:
        p = para(doc, space_after=2, size=8.5)
        _run(p, f"■　{midashi}　", size=8.5, bold=True, color=DARK)
        _run(p, text, size=8.5)
    para(doc, "", size=4, space_after=0, line=5)


def page5(doc):
    band(doc, "12", "年齢到達に伴うサービスの移行")
    para(doc, IKOU_LEAD, size=8.5, space_after=3)
    figure(doc, "11_ライフコース_概要版.png",
           "図　ライフコースと制度の移行",
           "資料：北塩原村（本計画第５章５及び第７章により作成）",
           width_mm=158)
    para(doc, IKOU_NOTE, size=8.5, space_after=4)

    band(doc, "13", "他計画との連携")
    para(doc,
         "本計画は、村の関連計画と整合を図り、連携して施策を推進します。",
         size=8.5, space_after=3)
    table(doc, RENKEI, [2800, 1900, 5506], size=8)


def page6(doc):
    band(doc, "14", "成年後見制度の利用促進")
    para(doc,
         "本村では、認知症、知的障がい、その他の精神上の障がいがあることにより、"
         "日常生活を支える様々な行為や買い物、財産管理が難しい事例が見られます。"
         "高齢化の進展を踏まえると、成年後見制度の必要性は今後も高まっていくことが見込まれます。",
         size=8.5)
    para(doc,
         "令和４年７月より中核機関「会津権利擁護・成年後見センター」を会津圏域11市町村で設置し、"
         "制度の周知啓発、広報活動、相談対応を行っています。令和６年度からは市民後見人養成事業が"
         "追加されました。本計画期間においても同センターとの連携を継続します。"
         "なお、成年後見制度については民法等の改正が予定されており、"
         "施行の時期と内容に応じて、中核機関を通じて村の運用を見直します。",
         size=8.5, space_after=4)

    band(doc, "15", "計画の推進体制")
    para(doc,
         "計画に定める事項については、PDCAサイクルにより進行管理を行います。"
         "評価は保健福祉課が行い、北塩原村障がい者自立支援協議会に報告して意見を聴いたうえで、"
         "結果を村ホームページ等で公表します。", size=8.5)
    table(doc, PDCA, [1800, 8406], size=8, fill_first=BAND)

    band(doc, "", "今後の予定")
    table(doc, SCHEDULE, [2400, 7806], size=8, fill_first=BAND)

    box(doc, [
        "この概要版についてのご意見をお寄せください",
        "本計画の素案は、北塩原村障がい者自立支援協議会での審議を経て、"
        "パブリックコメント（村民意見公募）を実施します。"
        "素案の全文及びこの概要版は、村ホームページ及び村役場窓口で"
        "ご覧いただけます。",
        "【お問合せ・ご意見の提出先】北塩原村役場　保健福祉課　福祉係"
        "（TEL）0241-23-3113",
    ], fill=NOTE_FILL, border=ACCENT, size=8.5, bold_first=True)


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
    style.font.size = Pt(9)
    rpr = style.element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = _el("rFonts")
        rpr.insert(0, rf)
    rf.set(qn("w:eastAsia"), FONT)
    rf.set(qn("w:ascii"), FONT)
    rf.set(qn("w:hAnsi"), FONT)


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    doc = docx.Document()
    setup_section(doc)

    page1(doc)
    page_break(doc)
    page2(doc)
    page_break(doc)
    page3(doc)
    page_break(doc)
    page4(doc)
    page4b(doc)
    page_break(doc)
    page5(doc)
    page_break(doc)
    page6(doc)

    doc.save(OUT_FILE)

    pages = estimate_lines(doc)
    print(f"作成: {OUT_FILE}")
    print(f"  ページ数: {len(pages)}")
    over = []
    for i, lines in enumerate(pages, 1):
        mark = "" if lines <= LINES_PER_PAGE else "　← 収まりません"
        if mark:
            over.append(i)
        print(f"  p.{i}  推定 {lines:.0f} 行 / 目安 {LINES_PER_PAGE} 行{mark}")
    return over, len(pages)


def suan_text():
    """計画素案の全文（本文と表）を返す。概要版の数値の照合に用いる。"""
    path = f"{OUT_DIR}/北塩原村_計画素案.docx"
    if not os.path.exists(path):
        raise SystemExit(
            f"計画素案が見つかりません: {path}\n"
            "build_kitashiobara_kosshi_rev.py を先に実行してください。")
    d = docx.Document(path)
    out = [p.text for p in d.paragraphs]
    for t in d.tables:
        for row in t.rows:
            out += [c.text for c in row.cells]
    return "\n".join(out)


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
    suan = suan_text()
    mikomi = load_mikomi()

    # 1 見込量の令和11年度の値が素案の表に現れていること
    #   素案は「7人分」「12人分」のように書くため、その形で照合する。
    for n in MIKOMI_PICK:
        users = mikomi[n]["users"]
        need = "0" if users == 0 else f"{users}人分"
        if need not in suan:
            ng.append(f"素案に{n}の見込量（{need}）が見つからない")

    # 2 概要版の成果目標の数値が素案に現れていること
    for need in ("【県の算定値待ち】", "のぞまないセルフプラン", "１か所",
                 "医療的ケア児等コーディネーター"):
        if need not in suan:
            ng.append(f"素案に成果目標の記載（{need}）が見つからない")

    # 3 KPIの現状値・目標値が素案 第3章5の表と一致すること
    for genjo, mokuhyo in (("25.3％", "15％"), ("51.3％", "60％"),
                           ("52％", "60％"), ("56.8％", "100％"),
                           ("28.1％", "40％")):
        for v in (genjo, mokuhyo):
            if v not in suan:
                ng.append(f"素案にKPIの値（{v}）が見つからない")

    # 4 【要確認】を付した指標が、素案で比較不可・要確認とされていること
    for need in ("比較不可", "条件付き可", "測定不能", "要確認。令和８年８月の村の確認では０件"):
        if need not in suan:
            ng.append(f"素案にKPIの判定（{need}）が見つからない。"
                      "概要版の【要確認】を見直してください")

    # 5 給付費の見込みが算定の正本と一致すること
    #   素案 第5章12 は令和7年度の実績を載せるため、照合先は算定ブックにする。
    per, total = load_kyufu()
    if total != per * 3:
        ng.append(f"給付費の3か年計が各年度の3倍でない: {total}／{per}")
    gaiyou = docx.Document(OUT_FILE)
    body = "\n".join([p.text for p in gaiyou.paragraphs]
                     + [c.text for t in gaiyou.tables for r in t.rows
                        for c in r.cells])
    for need in (_oku_man(total), _oku_man(per)):
        if need not in body:
            ng.append(f"概要版に給付費の見込み（{need}）が見つからない")

    # 6 廃止区分が概要版の見込量に入っていないこと
    if "医療型児童発達支援" in MIKOMI_PICK:
        ng.append("概要版の見込量に廃止区分が入っている")

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

    if ng:
        print("自己点検 不合格:")
        for x in ng:
            print("  -", x)
        return False
    print(f"自己点検: 合格　見込量{len(MIKOMI_PICK)}件・成果目標4点・KPI10点を"
          f"素案の現物と照合／給付費{_oku_man(total)}を算定ブックと照合／"
          f"方策{len(HOUSAKU)}件・他計画{len(RENKEI) - 1}件・"
          f"年齢到達3点・言い切り6点・人口と手帳8点を正本の現物と照合／"
          f"{npages}ページに収まる")
    return True


if __name__ == "__main__":
    import sys
    if not verify(*main()):
        sys.exit(1)
