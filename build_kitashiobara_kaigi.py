# -*- coding: utf-8 -*-
"""
北塩原村 第8期障がい福祉計画・第4期障がい児福祉計画
会議資料ジェネレータ（自立支援協議会・庁議）

出力: output/会議資料/*.docx

業務委託仕様書 4-Ⅱ(7)「計画策定委員会の支援（会議資料の作成、会議への出席、
会議記録の作成）」に対応する。説明の主体は村であり、受託者は資料を作成する。

本スクリプトが生成する資料
  協議会資料1  計画策定の概要
  協議会資料2  国の基本指針の改正点（第7期→第8期）
  協議会資料3  前回計画の評価（第7期・第3期の最終評価、第4次の中間評価）
  協議会資料8  提供体制の課題（圏域のサービス資源）
  庁議資料2    関連計画との整合
  庁議資料4    庁内各課への依頼事項

いずれも村資料の受領を待たずに作成できるもの。数値は原典（source配下）から
読み出し、本スクリプトには直接書かない。

協議会資料3の中身は build_kitashiobara_zenkai_hyouka.py に置いている。
評価の文章を二重に持つと、どちらかが古くなる。
本スクリプトは読み込んで資料の形に組むだけである。
まだ分からない値は空欄にせず「確認中」と示したまま委員会に出す。
分かっていないことを分かっているように見せない。

構成案は docs/北塩原村_協議会等資料_構成案.md による。
"""

import os
import re

import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from openpyxl import load_workbook

import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_kitashiobara_zenkai_hyouka import (  # noqa: E402
    GAIBU, HANTEI_TEJUN, KANGAE, KPI4, KPI7, KUBUN, MIKOMI,
    SHISAKU, SHISAKU_KUBUN, SOUKATSU,
)
from build_kitashiobara_mura_daicho import DK_HYOU  # noqa: E402
from kitashiobara_koutei import (  # noqa: E402
    PABU_KAISHI, PABU_NISSU, PABU_SHURYO, _KOUTEI, taigai, wareki,
)

# 正本（計画素案）の現物。資料1の章立て・庁議資料2の高齢化率・
# 庁議資料4の記載箇所は、ここから読む（令和8年10月24日 横断レビュー）。
HONPON = f"{os.path.dirname(os.path.abspath(__file__))}/output/北塩原村_計画素案_正本_移植後.docx"

REPO_ROOT = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = f"{REPO_ROOT}/output/会議資料"
SHIGEN_DIR = f"{REPO_ROOT}/source/障害福祉資源"

FONT = "BIZ UDPゴシック"
HEADER_FILL = "1F3864"
SUB_FILL = "2E75B6"
NOTE_FILL = "FFF2CC"
BORDER = "AAAAAA"

VILLAGE = "北塩原村"
KEN_HOKUBU = ["北塩原村", "猪苗代町", "磐梯町", "湯川村"]
NEIGHBOR = ["会津若松市", "喜多方市"]
HOUMON = ["居宅介護", "重度訪問介護", "同行援護", "行動援護", "重度障害者等包括支援"]
SHIGEN_ORDER = HOUMON + [
    "生活介護", "自立訓練（機能訓練）", "自立訓練（生活訓練）", "宿泊型自立訓練",
    "就労移行支援", "就労継続支援（A型）", "就労継続支援（B型）", "就労定着支援",
    "療養介護", "短期入所", "自立生活援助", "共同生活援助", "施設入所支援",
]


# ============================================================
# 書式ヘルパー
# ============================================================
def _el(tag, **attrs):
    e = OxmlElement(tag if ":" in tag else f"w:{tag}")
    for k, v in attrs.items():
        e.set(qn(f"w:{k}"), str(v))
    return e


def _font(run, size=10.5, bold=False, color=None):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = _el("rFonts")
        rpr.append(rf)
    for a in ("ascii", "hAnsi", "eastAsia", "cs"):
        rf.set(qn(f"w:{a}"), FONT)


def new_doc(title, subtitle):
    doc = docx.Document()
    sec = doc.sections[0]
    sec.top_margin = Cm(2.0)
    sec.bottom_margin = Cm(1.8)
    sec.left_margin = Cm(2.0)
    sec.right_margin = Cm(2.0)
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
    pf = p.paragraph_format
    pf.space_before = Pt(6)
    pf.space_after = Pt(12)
    return doc


def h2(doc, text):
    p = doc.add_paragraph()
    r = p.add_run(f"■　{text}")
    _font(r, size=12, bold=True, color=HEADER_FILL)
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


def bullet(doc, text, size=10.5):
    p = doc.add_paragraph()
    _font(p.add_run(f"・{text}"), size=size)
    p.paragraph_format.left_indent = Pt(12)
    p.paragraph_format.space_after = Pt(2)
    return p


def table(doc, rows, widths=None, header=True, size=9.5):
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
            r = p.add_run(lines[0])
            _font(r, size=size, bold=(header and ri == 0),
                  color="FFFFFF" if (header and ri == 0) else None)
            for ln in lines[1:]:
                r = p.add_run()
                r.add_break()
                _font(r, size=size)
                r2 = p.add_run(ln)
                _font(r2, size=size, bold=(header and ri == 0),
                      color="FFFFFF" if (header and ri == 0) else None)
            if header and ri == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                sh = _el("shd", val="clear", color="auto", fill=HEADER_FILL)
                cell._tc.get_or_add_tcPr().append(sh)
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
# 原典の読み出し
# ============================================================
def load_shigen():
    """ReMHRAD 障害福祉資源から市町村別の事業所数（2020→2025）を読む。"""
    data = {}
    for name in sorted(os.listdir(SHIGEN_DIR)):
        if not name.endswith(".xlsx"):
            continue
        wb = load_workbook(os.path.join(SHIGEN_DIR, name),
                           data_only=True, read_only=True)
        title = wb.sheetnames[0]
        m = re.search(r"（(.+?)）の状況", title)
        svc = (m.group(1) if m else title).replace("(", "（").replace(")", "）")
        vals = {}
        for r in wb[title].iter_rows(min_row=2, values_only=True):
            if not r or not r[0]:
                continue
            try:
                vals[str(r[0])] = (int(r[2]), int(r[4]))
            except (TypeError, ValueError):
                continue
        wb.close()
        data[svc] = vals
    return data


# ============================================================
# 正本（計画素案）と単一出所から読む
#   令和8年10月24日の他ＡＩの横断レビューで、協議会資料・庁議資料が
#   正本と食い違ったまま残っていることが指摘された（RT-05ほか）。
#   本スクリプトには自己点検が無く、正本が直っても資料が置いていかれた。
#   章立て・工程・実績・高齢化率は、ここから読む。
# ============================================================
def _honpon():
    if not os.path.exists(HONPON):
        raise SystemExit(f"正本がありません（先に honpon を実行）: {HONPON}")
    return docx.Document(HONPON)


def honpon_midashi():
    """[(章の見出し, [節の見出し, ...]), ...]"""
    out = []
    for p in _honpon().paragraphs:
        if p.style.name == "Heading 1":
            out.append((p.text.strip(), []))
        elif p.style.name == "Heading 2" and out:
            out[-1][1].append(p.text.strip())
    return out


def honpon_koureika():
    """正本 表2-1（人口の見通し）の令和11年度の高齢化率。"""
    for t in _honpon().tables:
        rows = [[c.text.strip() for c in r.cells] for r in t.rows]
        if rows and rows[0] and rows[0][0] == "区分" and \
                any(r[0] == "高齢化率" for r in rows):
            j = next(i for i, h in enumerate(rows[0]) if h.startswith("令和11年度"))
            return next(r[j] for r in rows if r[0] == "高齢化率")
    raise LookupError("正本に人口の見通しの表（高齢化率）が無い")


def _dk(svc, tani):
    hits = [r for r in DK_HYOU if r.get("svc") == svc and r.get("tani") == tani]
    if len(hits) != 1:
        raise LookupError(f"実績の表の行が一意でない: {svc} {tani}")
    return [re.match(r"^([0-9.]+|―)（", v).group(1) for v in hits[0]["ato"]]


# 策定の工程に載せる行。日付は kitashiobara_koutei の工程（案Ａ）から読み、
# 仮置きの日は月又は旬に戻す（taigai）。
KOUTEI_NOSERU = ["2026-11-18", "2026-12-16", "2027-01-15", "2027-02-19",
                 "2027-02-24"]


def koutei_rows():
    kou = {hi: k for hi, k, *_r in _KOUTEI}
    rows = [
        ("令和８年８月31日～９月25日", "アンケート調査の実施（障がい者調査・障がい児調査）"),
        ("令和８年10月", "アンケート調査結果の集計・分析、実績と財源の整理"),
        ("令和８年10月～11月", "計画素案の作成"),
    ]
    for hi in KOUTEI_NOSERU[:3]:
        rows.append((taigai(hi), kou[hi]))
    rows.append((f"{wareki(PABU_KAISHI)}～{wareki(PABU_SHURYO)}",
                 f"パブリックコメント（{PABU_NISSU}日間）"))
    for hi in KOUTEI_NOSERU[3:]:
        rows.append((taigai(hi), kou[hi]))
    #   最後の行は工程の一か所（kitashiobara_koutei.GAIYOU_HYO）から読む。
    #   「計画の確定・公表」と書いていたが、計画の決定は２月の庁議である
    #   （令和8年10月24日 レッドチームの点検）
    from kitashiobara_koutei import GAIYOU_HYO
    rows.append(tuple(GAIYOU_HYO[-1]))
    return rows


# ============================================================
# 協議会資料1　計画策定の概要
# ============================================================
def shiryo_gaiyou():
    doc = new_doc("第８期北塩原村障がい福祉計画・第４期北塩原村障がい児福祉計画の策定について",
                  "北塩原村障がい者自立支援協議会　資料１")

    h2(doc, "１　計画の趣旨と法的根拠")
    table(doc, [
        ["計画", "根拠法", "対象期間", "改定"],
        ["第８期北塩原村障がい福祉計画", "障害者総合支援法第88条第１項",
         "令和９年度〜令和11年度", "今回改定"],
        ["第４期北塩原村障がい児福祉計画", "児童福祉法第33条の20第１項",
         "令和９年度〜令和11年度", "今回改定"],
        ["第４次北塩原村障がい者計画", "障害者基本法第11条第３項",
         "令和６年度〜令和11年度", "中間見直し"],
    ], [3600, 3000, 2400, 1400])
    para(doc,
         "障がい福祉計画・障がい児福祉計画は、国の基本指針に即して３年を一期として"
         "作成するものです。第４次障がい者計画は令和11年度までを計画期間としており、"
         "本計画の策定にあわせて中間見直しを行います。"
         "基本施策ごとの目標値の達成状況を確認し、目標値を改める場合は"
         "本計画の第３章に改定後の値を掲げます。"
         "その基本理念・基本目標は本計画にも共有されます。")

    h2(doc, "２　国の基本指針")
    para(doc,
         "「障害福祉サービス等及び障害児通所支援等の円滑な実施を確保するための"
         "基本的な指針」は、令和８年３月31日にこども家庭庁・厚生労働省告示第４号により"
         "改正されました。市町村はこの基本指針に即して計画を作成します。"
         "主な改正点は資料２のとおりです。")

    h2(doc, "３　計画の策定体制")
    table(doc, [
        ["体制", "役割"],
        ["北塩原村障がい者自立支援協議会", "計画内容の審議"],
        ["庁議（村長・副村長・教育長・各課長）", "計画の立案、調整、進行管理"],
        ["アンケート調査", "生活状況、今後の要望・課題の把握"],
        ["パブリックコメント", "村民からの意見募集"],
    ], [4200, 6200])

    h2(doc, "４　策定の工程")
    rows = [["時期", "内容"]] + [list(r) for r in koutei_rows()]
    table(doc, rows, [2800, 7600])

    h2(doc, "５　計画の構成（素案）")
    rows = [["章", "主な内容"]]
    for sho, setsu in honpon_midashi():
        rows.append([sho, "、".join(re.sub(r"^[０-９0-9]+　", "", x)
                                    for x in setsu)])
    table(doc, rows, [4600, 5800], size=9)
    note(doc,
         "第５章の「年齢到達に伴うサービスの移行」と"
         "「障がい福祉サービス等に係る費用の推移と財源構成」は、"
         "現行計画にはなかった項目です。"
         "第８期の基本指針が見込量の算定の考え方を明確にしたことを受けて新設します。")
    return doc, "01_協議会資料1_計画策定の概要.docx"


# ============================================================
# 協議会資料2　国の基本指針の改正点
# ============================================================
SEIKA_HIKAKU = [
    ["成果目標の項目", "第７期（令和６〜８年度）", "第８期（令和９〜11年度）"],
    ["福祉施設入所者の地域生活への移行",
     "地域移行者数：令和４年度末施設入所者数の６％以上\n施設入所者数：５％以上削減",
     "地域移行者数：令和７年度末施設入所者数の６％以上\n施設入所者数：５％以上削減\n"
     "（率は据置き、基準時点が更新）"],
    ["精神障害にも対応した\n地域包括ケアシステムの構築",
     "平均生活日数：325.3日以上\n長期在院者数（65歳未満・65歳以上）\n退院率（３か月・６か月・１年）",
     "平均生活日数：319.3日以上\n長期在院者数（65歳未満・65歳以上・75歳以上・40歳以上の認知症）\n"
     "再入院率：90日10.3％・180日17.4％・365日25.7％以下【新規】\n"
     "心のサポーター：令和15年度末までに全国100万人【新規】\nＫ６による住民の心の状態の把握【新規】"],
    ["福祉施設から一般就労への移行等",
     "一般就労移行者数：令和３年度実績の1.28倍以上\nうち就労移行支援 1.31倍\n"
     "うちＡ型 概ね1.29倍　うちＢ型 概ね1.28倍\n就労定着支援利用者数：1.41倍以上",
     "一般就労移行者数：令和６年度実績の1.31倍以上\nうち就労移行支援 1.14倍\n"
     "うちＡ型 おおむね1.52倍　うちＢ型 おおむね1.67倍\n就労定着支援利用者数：1.47倍以上\n"
     "就労選択支援：圏域ごとに１事業所以上、令和11年度の利用者を全国で82,000人以上【新規】"],
    ["障がい児支援の提供体制の整備等",
     "児童発達支援センターの設置\n地域社会への参加・包容の推進",
     "児童発達支援センター等の「４つの中核機能」の確保\n"
     "インクルージョン推進のための協議の場の設置【新規】\n"
     "主に重症心身障がい児を支援する事業所を各市町村に１か所以上"
     "（市町村単独で困難な場合は圏域での確保も可）\n"
     "障がい児及びその家族への伴走的な相談支援体制の確保【新規】"],
    ["地域生活支援の充実",
     "地域生活支援拠点等の整備、年１回以上の運用状況の検証・検討",
     "拠点等の整備に加え、当該市町村の全ての日常生活圏域を支援の対象とする\n"
     "拠点コーディネーター及び機能を担う事業所等の担当者の配置\n"
     "強度行動障害の状態にある者の支援ニーズの把握と体制整備"],
    ["相談支援体制の充実・強化等",
     "基幹相談支援センターの設置等",
     "基幹相談支援センター・地域生活支援拠点等・協議会を整備し連携させる\n"
     "令和11年度末までに「のぞまないセルフプラン」の件数をゼロとする【新規】"],
    ["障がい福祉人材の確保・定着、\n生産性向上",
     "（項目なし）",
     "【新規項目】都道府県におけるワンストップ窓口の設置、協議会の設置、"
     "相談支援専門員研修等の実施"],
    ["障がい福祉サービス等の質の確保",
     "サービスの質の向上のための体制構築",
     "同左に加え、情報公表制度における公表率・更新率を100％とする"
     "（各都道府県・政令市・中核市の管内事業所）【新規】\n"
     "本計画では、村内の対象事業所について100％を目指して働きかける"
     "（第４章２（８））"],
]

MINAOSHI = [
    ("①入所等から地域生活への移行", "データを活用した地域移行者数の把握、居室の個室化等の推進", ""),
    ("②精神障害にも対応した地域包括ケア", "市町村における相談・援助体制の整備", ""),
    ("③一般就労への移行等", "就労選択支援の積極的な利用促進、成果目標の新設", ""),
    ("④障がい児支援の提供体制の整備等", "４つの中核機能の確保、インクルージョン推進の協議の場の設置", ""),
    ("⑤相談支援体制の充実強化", "基幹相談支援センターの設置推進、協議会への当事者参画", ""),
    ("⑥人材確保・定着、生産性向上", "介護テクノロジーの導入促進、人材確保支援体制の整備", "新規分野"),
    ("⑦サービスの質の確保等", "情報公表制度の公表率等に関する成果目標の新設", ""),
    ("⑧きめ細かい地域ニーズを踏まえた支援体制", "意思疎通支援従事者の養成・派遣体制、ＩＣＴ機器の利用支援", ""),
    ("⑨高次脳機能障害者に対する支援", "高次脳機能障害者支援法の成立を踏まえた相談支援体制の充実", "新規"),
    ("⑩人口減少地域におけるサービスの維持・確保",
     "中山間・人口減少地域におけるサービス提供体制の維持・確保の重要性", "本村に直結"),
    ("⑪地域共生社会の実現に向けた取組", "一層の推進", ""),
    ("⑫住宅セーフティネット制度との連携", "住宅担当部局・居住支援協議会等との連携", ""),
    ("⑬地域差の是正・指定の在り方等",
     "地域差の是正に向けたサービス見込量の算出方法、重度障がい者の個別利用者数の見込み設定",
     "本村に直結"),
    ("⑭障害者等に対する虐待の防止等", "調査の徹底と体制整備の強化", ""),
    ("⑮スポーツ・健康増進活動による社会参加等の促進", "―", ""),
    ("⑯災害時における障害福祉サービス提供の確保", "防災部局・職能団体等との連携、耐災害性強化", ""),
]


def shiryo_shishin():
    doc = new_doc("国の基本指針の改正点（第７期→第８期）",
                  "北塩原村障がい者自立支援協議会　資料２")

    para(doc,
         "「障害福祉サービス等及び障害児通所支援等の円滑な実施を確保するための基本的な指針」は、"
         "令和８年３月31日にこども家庭庁・厚生労働省告示第４号により改正されました。"
         "市町村はこの基本指針に即して計画を作成します。本資料は第７期からの変更点を"
         "告示の本文及び新旧対照表と照合して整理したものです。")

    h2(doc, "１　成果目標の変更点")
    para(doc, "成果目標は第７期の７項目から第８期の８項目に増えました。")
    table(doc, SEIKA_HIKAKU, [2600, 3700, 4100], size=8.5)

    h2(doc, "２　本村に特に関係する変更点")
    h3(doc, "（１）就労継続支援Ｂ型の目標倍率が大幅に引き上げられました")
    b6, b7, _b8 = _dk("就労継続支援\nＢ型", "人／月")
    a6, a7, _a8 = _dk("就労継続支援\nＡ型", "人／月")
    para(doc,
         "第７期は「概ね1.28倍」でしたが、第８期は「おおむね1.67倍」となりました。"
         f"本村の就労系サービスの利用者は就労継続支援Ｂ型（令和６年度 月{b6}人、"
         f"令和７年度 月{b7}人）が中心であり"
         + ("（就労継続支援Ａ型は令和７年度まで利用がありません）"
            if (a6, a7) == ("0", "0") else
            f"（就労継続支援Ａ型は令和６年度 月{a6}人、令和７年度 月{a7}人）")
         + "、この引上げの影響を最も受けます。"
         "なお、一般就労移行者数の全体目標（1.31倍）とは別に、"
         "事業種別ごとの目標を定めることが求められています。")

    h3(doc, "（２）人口減少地域におけるサービスの維持・確保が明記されました")
    para(doc,
         "見直し事項⑩として、中山間・人口減少地域におけるサービス提供体制の維持・確保の"
         "重要性が記載されました。本村は過疎地域の持続的発展の支援に関する特別措置法第２条により"
         "全域が過疎地域に該当し、福島県過疎・中山間地域振興条例の中山間地域（全域）でもあるため、"
         "この視点に正面から該当します。基本指針は、共生型サービス、基準該当障害福祉サービス、"
         "多機能型、従たる事業所等の現行制度の活用を挙げています（詳細は資料８）。")

    h3(doc, "（３）サービス見込量の算定方法に関する規定が新設されました")
    para(doc,
         "見直し事項⑬「地域差の是正」に伴い、別表第五が新設されました。"
         "全部過疎市町村ではなく、かつサービス利用者割合が上位25％の範囲内である市町村は、"
         "令和９〜11年度の見込量を「令和８年度の見込みと全国の伸び率」を用いて定めることが"
         "基本となります。本村は全域が過疎地域に該当するため、この算定方法の適用対象外であり、"
         "利用実績・意向・心身の状況を勘案して見込量を設定する原則によります。")

    h3(doc, "（４）重度障がい者の個別の利用者数の把握が求められます")
    para(doc,
         "生活介護・短期入所・共同生活援助について、強度行動障害の状態にある者、"
         "高次脳機能障害を有する者、医療的ケアを必要とする者等の重度障がい者について、"
         "個別に利用者数の見込みを設定するよう努めることとされました。")

    h2(doc, "３　基本指針の見直し事項（16項目）")
    rows = [["項目", "主な内容", "本村との関係"]]
    for a, b, c in MINAOSHI:
        rows.append([a, b, c])
    table(doc, rows, [3400, 5400, 1600], size=8.5)

    note(doc,
         "本資料の数値は、令和８年３月31日告示第４号の本文及び新旧対照表と１件ずつ照合しています。"
         "なお、国が公表した改正後概要（３頁）は就労定着支援の基準を「令和６年度末実績」と"
         "記載していますが、告示本文は「令和六年度の実績」です。"
         "本村の基準値（令和６年度の利用者数）は、いずれの読み方でも０人であり、"
         "計画素案 第４章２（３）の値は変わりません。")
    note(doc,
         "このほか、令和８年７月10日（社会福祉法等の改正を受けたもの）及び"
         "８月19日（医療的ケア児支援法の改正を受けたもの）に、"
         "基本指針の改正案が国の審議会に示されています。"
         "計画素案で引いている改正案の内容は、"
         "告示の確定後に照合して反映します（Ｓ-35）。")
    return doc, "02_協議会資料2_国の基本指針の改正点.docx"


# ============================================================
# 協議会資料3　前回計画の評価
#   中身は build_kitashiobara_zenkai_hyouka.py にある。
#   ここでは資料の形に組むだけである。
# ============================================================
def shiryo_hyouka():
    doc = new_doc("前回計画の評価",
                  "北塩原村障がい者自立支援協議会　資料３")

    para(doc,
         "本資料は、第８期障がい福祉計画・第４期障がい児福祉計画の"
         "策定にあたり、前回計画の達成状況を整理したものです。"
         "評価の対象は二つあります。"
         "第７期障がい福祉計画・第３期障がい児福祉計画（令和６〜８年度）は"
         "計画期間の最終年度の途中であるため、令和８年度末の見込みによる"
         "評価（暫定）です。令和８年度の値は年度途中までの実績によります。"
         "第４次北塩原村障がい者計画（令和６〜11年度）は"
         "期間の途中であるため中間評価です。")
    for bun in SOUKATSU[1][1].split("\n"):
        para(doc, bun)

    h2(doc, "１　評価の考え方")
    para(doc,
         "評価は成績表ではなく、第８期計画の根拠とするために行います。"
         "次の五つを前提としています。")
    rows = [["前提", "内容"]]
    for a, b in KANGAE:
        rows.append([a, b])
    table(doc, rows, [3000, 7400], size=9)

    h2(doc, "２　成果目標・ＫＰＩの達成状況（第７期・令和８年度末見込み）")
    para(doc,
         "第７期計画の成果目標及び指標14項目について、"
         "計画値と実績を突き合わせたものです。"
         "「達成」としたものも、数値の達成と課題の解決は別であるため、"
         "第８期での扱いを併せて示します。")
    rows = [["成果目標・指標", "項目", "計画値", "実績\n（見込み）", "判定",
             "第８期での評価の仕方", "第８期での扱い"]]
    for rec in KPI7:
        rows.append(list(rec))
    table(doc, rows, [1700, 1500, 900, 900, 800, 2300, 2300], size=8)

    h2(doc, "３　サービス見込量の評価（第７期・令和８年度末見込み）")
    para(doc,
         "見込量は、計画値と実績が一致したかどうかだけでは評価しません。"
         "実績が０であることをもってニーズがないとは評価しません。"
         "次の六つの区分に分けて整理しています。"
         "実績は計画素案 第５章の実績の表"
         "（村の協議会資料２－３と給付の台帳による）の値です。")
    rows = [["区分", "意味"]]
    for a, b in KUBUN:
        rows.append([a, b])
    table(doc, rows, [2000, 8400], size=9)

    para(doc, "区分は次の手順で決めています。")
    for i, t in enumerate(HANTEI_TEJUN, 1):
        bullet(doc, f"{'①②③④⑤⑥'[i - 1]}　{t}", size=9.5)
    rows = [["サービス", "区分", "実績の状況", "第８期での扱い"]]
    for rec in MIKOMI:
        rows.append(list(rec))
    table(doc, rows, [1800, 1300, 4600, 2700], size=8)

    h2(doc, "４　施策の評価（第７期・令和８年度末見込み）")
    para(doc,
         "施策については、本村で実施しているかどうかではなく、"
         "住民の必要が充たされているかどうかで評価します。"
         "村内に事業所がないサービスは、"
         "圏域で対応できていれば課題として扱いません。")
    rows = [["区分", "意味"]]
    for a, b in SHISAKU_KUBUN:
        rows.append([a, b])
    table(doc, rows, [2600, 7800], size=9)

    rows = [["分野", "区分", "評価", "第８期での扱い", "主な論点"]]
    for rec in SHISAKU:
        rows.append(list(rec))
    table(doc, rows, [1500, 1900, 3000, 1200, 2800], size=8.5)

    h2(doc, "５　第４次障がい者計画の指標（中間評価）")
    para(doc,
         "第４次障がい者計画の計画期間は令和６年度から令和11年度までで、"
         "いまは期間の途中です。"
         "このため「未達成」とは評価せず、"
         "基準値から目標値までの進み具合として見ます。"
         "現在値が確認中の指標があります。"
         "村への照会（Ｍ-42・Ｍ-95）の回答をいただいたうえで"
         "確定します。"
         "外出目的の指標（⑦）は、今回の調査票に該当する設問がないため"
         "測ることができず、測り方を村の実績に置き換えます。")
    rows = [["基本施策", "指標項目", "基準値\n（令和５年度）", "現在値",
             "目標\n（令和11年度）", "中間評価の読み方"]]
    for rec in KPI4:
        rows.append(list(rec))
    table(doc, rows, [1500, 2600, 1100, 1500, 1100, 2600], size=8)

    h2(doc, "６　外部の値との比較")
    para(doc,
         "本村の値だけを見ても、高いのか低いのかが分かりません。"
         "全国又は県の値と並べて示します。")
    rows = [["項目", "区分", "値", "備考", "出所"]]
    for rec in GAIBU:
        rows.append(list(rec))
    table(doc, rows, [2400, 1100, 1200, 2900, 2800], size=8)

    h2(doc, "７　第８期計画への反映")
    for bun in SOUKATSU[2][1].split("\n"):
        para(doc, bun)

    note(doc,
         "本資料のうち現在値が「確認中」となっている欄は、"
         "村への照会の回答をいただいたうえで確定します。"
         "分かっていない値を空欄にすると、"
         "値がないのか０なのかが分からなくなるため、"
         "「確認中」と示したままお出ししています。"
         "また、実績が０であるサービスについて"
         "「ニーズなし」と判定したものはありません。"
         "利用がない理由を、制度を知らない・供給がない・"
         "送迎がない・家族が担っている・他の制度で足りている・"
         "本当に必要がない、に分けて確かめたうえで"
         "第８期の見込量に反映します。")
    return doc, "03_協議会資料3_前回計画の評価.docx"


def tanki_bun():
    """短期入所の実績（計画素案 表５-７）。"""
    n6, n7, n8 = _dk("短期入所\n（福祉型）", "人／月")
    d6, d7, d8 = _dk("短期入所\n（福祉型）", "人日／月")
    return (f"第７期計画では見込んでいなかったが、令和６年度は月{n6}人・{d6}人日、"
            f"令和７年度は月{n7}人・{d7}人日の利用があった"
            "（計画素案 表５-７）。圏域の受入可能事業所を把握する")


# ============================================================
# 協議会資料8　提供体制の課題
# ============================================================
def shiryo_teikyou(shigen):
    doc = new_doc("提供体制の課題（会津北部圏域のサービス資源）",
                  "北塩原村障がい者自立支援協議会　資料８")

    para(doc,
         "国の基本指針は、訪問系サービス及び指定通所支援について"
         "「各市町村において事業を実施する事業所を最低一箇所確保できるよう努める」"
         "こととしています（第三の二の２㈡）。本資料は本村及び会津北部圏域の"
         "事業所数の状況を整理したものです。")

    h2(doc, "１　会津北部圏域（４町村）と近隣市の事業所数（令和７年）")
    munis = KEN_HOKUBU + NEIGHBOR
    rows = [["サービス"] + [m for m in munis] + ["圏域４町村計"]]
    for svc in SHIGEN_ORDER:
        vals = shigen.get(svc, {})
        cnt = [vals.get(m, (None, None))[1] for m in munis]
        total = sum(x for m, x in zip(munis, cnt) if m in KEN_HOKUBU and x)
        rows.append([svc.replace("（A型）", "（Ａ型）").replace("（B型）", "（Ｂ型）")]
                    + [str(x if x is not None else "―") for x in cnt] + [str(total)])
    table(doc, rows, [2600] + [1150] * 6 + [1400], size=8.5)
    para(doc,
         "出典：ＲｅＭＨＲＡＤ 障害福祉資源（原典は障害福祉サービス等情報公表システム）。"
         "事業所数は所在地の市町村で計上されるため、本村の住民が圏域外の事業所を"
         "利用している分はこの表には現れません。"
         "会津若松市・喜多方市は会津北部圏域に含まれない近隣市であり、"
         "圏域４町村計には含めていません。", indent=False, size=9)

    h2(doc, "２　この表から読み取れること")
    village = {s: shigen.get(s, {}).get(VILLAGE, (0, 0)) for s in SHIGEN_ORDER}
    has = [s for s in SHIGEN_ORDER if village[s][1] > 0]
    h3(doc, "（１）本村に事業所があるサービス")
    para(doc,
         "令和７年時点で本村に所在する事業所は "
         + "、".join(f"{s}{village[s][1]}か所" for s in has)
         + " です。いずれも令和２年時点は０か所であり、その後に指定を受けたものと考えられます。")

    h3(doc, "（２）圏域の訪問系サービスは本村の事業所のみとなっています")
    hk = {s: {m: shigen.get(s, {}).get(m, (0, 0)) for m in KEN_HOKUBU} for s in HOUMON}
    rows = [["サービス"] + [f"{m}\n令和２年→令和７年" for m in KEN_HOKUBU]]
    for s in HOUMON:
        rows.append([s] + [f"{hk[s][m][0]} → {hk[s][m][1]}" for m in KEN_HOKUBU])
    table(doc, rows, [2800] + [1900] * 4, size=9)
    para(doc,
         "猪苗代町の居宅介護・重度訪問介護は令和２年の各１か所から令和７年には０か所となり、"
         "会津北部圏域４町村で訪問系サービスを提供する事業所は本村の１か所のみとなっています。"
         "この１か所が撤退した場合、圏域全体が会津若松市・喜多方市の事業所への依存となります。")

    h2(doc, "３　基本指針が示す確保方策")
    para(doc,
         "基本指針は、小規模町村等において事業所を確保できない場合の工夫として"
         "次を挙げています（第三の二の２㈡、第一の一の３）。")
    table(doc, [
        ["方策", "内容", "本村での検討事項"],
        ["介護保険事業所への\n指定取得の働きかけ",
         "介護保険の訪問介護事業所や居宅介護支援事業所に対して、"
         "障害者総合支援法に基づく居宅介護事業所としての指定を取るよう促す",
         "村内及び圏域の介護保険事業所に指定取得の意向を確認する"],
        ["共生型サービスの周知",
         "平成30年に創設された共生型サービスの指定制度を事業者へ周知する",
         "圏域４町村と協調して制度周知を行い、指定の意向がある事業所を把握する"],
        ["基準該当障害福祉サービス\nの活用",
         "指定基準を完全には満たさない事業所であっても、市町村が認めた場合に"
         "特例的に給付の対象とすることができる",
         "圏域内で確保が困難なサービスについて活用の可否を県と協議する"],
        ["多機能型・従たる事業所",
         "複数のサービスを一体的に実施する多機能型、主たる事業所と一体的に運営する"
         "従たる事業所により、小規模でも成り立つ運営形態を確保する",
         "圏域事業者に本村への従たる事業所の設置可能性を照会する"],
        ["短期入所事業所の確保",
         "介護者が病気等になったときに対応できる短期入所の充実。"
         "医療型短期入所をはじめとした事業所の確保に努める",
         tanki_bun()],
    ], [2400, 4400, 3600], size=9)

    h2(doc, "４　協議会にお諮りしたい事項")
    for t in [
        "圏域４町村の訪問系サービスの提供体制について、どのように確保していくか。",
        "村内及び圏域の介護保険事業所への共生型サービス・居宅介護指定の働きかけを、"
        "圏域共同で行うことについて。",
        "本村の事業所（上記１）の受入可能人数・送迎範囲について、"
        "実態を把握する方法について。",
    ]:
        bullet(doc, t)

    note(doc,
         "【要確認】本表の出典である障害福祉サービス等情報公表システムに公表されていない"
         "事業所は計上されません。第８期基本指針が公表率・更新率100％を成果目標に"
         "新設したのは、この網羅性の問題への対処です（第二の八）。"
         "村の指定・委託の記録と突き合わせて確認する必要があります。")
    return doc, "04_協議会資料8_提供体制の課題.docx"


# ============================================================
# 庁議資料2　関連計画との整合
# ============================================================
def shiryo_kanren():
    doc = new_doc("関連計画との整合", "庁議　資料２")

    para(doc,
         "第８期障がい福祉計画・第４期障がい児福祉計画は、村の他の計画と整合を図りながら"
         "策定します。国の基本指針も、地域福祉計画、介護保険事業計画、"
         "こども・子育て支援事業計画、賃貸住宅供給促進計画等との調和を求めています。")

    h2(doc, "１　関連計画の一覧")
    table(doc, [
        ["計画", "所管", "本計画との関係"],
        ["第４次北塩原村障がい者計画\n（令和６〜11年度）", "保健福祉課",
         "基本理念・基本目標・施策体系を共有する。"
         "本計画の策定にあわせて中間見直しを行う"],
        ["北塩原村総合振興計画", "企画・政策部門",
         "村の最上位計画。人口の将来見通しの根拠"],
        ["地域福祉計画", "保健福祉課",
         "策定の有無と予定を村に確認中（Ｍ-135）。"
         "策定する場合は重層的支援体制整備事業実施計画との連携"
         "（基本指針 第一の一の４）"],
        ["第３次健康21・北塩原村グッドヘルスプラン", "保健福祉課",
         "ゲートキーパー養成講座と心のサポーター養成の連携、"
         "Ｋ６を活用した心の状態の把握、特定健診受診率"],
        ["北塩原村こども・子育て計画", "こども関係部局",
         "こども家庭センターとの相談導線、保育所等における障がい児の受入体制、"
         "インクルージョン推進の協議の場"],
        ["北塩原村介護保険事業計画・高齢者福祉計画", "保健福祉課",
         "65歳到達に伴うサービス移行の個別判断、共生型サービスの指定、"
         "介護保険事業所への居宅介護指定の働きかけ"],
        ["北塩原村地域防災計画・避難行動要支援者名簿", "総務課（防災）",
         "個別避難計画の策定、災害時における障害福祉サービス提供の確保"],
        ["福島県第８期障がい福祉計画", "福島県障がい福祉課",
         "精神病床の長期入院患者に係る基盤整備量、県の成果目標との整合"],
    ], [3400, 2400, 4600], size=9)

    h2(doc, "２　特に調整が必要な事項")
    h3(doc, "（１）こども・子育て計画とのインクルージョン推進")
    para(doc,
         "第８期基本指針は、インクルージョン推進のための協議の場の設置を"
         "各市町村の成果目標として新設しました。保育・子育て支援・教育等の関係機関が"
         "連携を図る場であり、こども家庭センターの機能と重なります。"
         "障がい福祉部門とこども部門のどちらが設置主体となるかの整理が必要です。"
         "計画素案では、既存の自立支援協議会や会津北部地域生活支援拠点の"
         "協議の場を活用して設置することとしています（第４章２（４））。")

    h3(doc, "（２）介護保険事業計画との関係")
    para(doc,
         f"本村の令和11年度の高齢化率は、計画素案（住民基本台帳の実績による"
         f"見通し）では{honpon_koureika()}、"
         "村のこども・子育て支援事業計画の推計では48.1％であり、"
         "いずれを用いるかを確認しています（Ｍ-142）。"
         "65歳到達に伴うサービス移行は一律ではなく個別判断であり、"
         "介護保険部門と障がい福祉部門の連携が必要です。"
         "また、圏域の訪問系サービス事業所が本村の１か所のみという状況に対し、"
         "介護保険の訪問介護事業所への共生型サービス・居宅介護指定の働きかけが"
         "基本指針で挙げられており、これは介護保険部門との共同作業になります。")

    h3(doc, "（３）健康21・グッドヘルスプランとの連携")
    para(doc,
         "第８期基本指針は「心のサポーター」の養成とＫ６を活用した住民の心の状態の把握を"
         "新設しました。いずれも健康増進部門の既存事業（ゲートキーパー養成講座、健診）と"
         "一体で実施することが現実的です。")
    return doc, "05_庁議資料2_関連計画との整合.docx"


# ============================================================
# 庁議資料4　庁内各課への依頼事項
# ============================================================
IRAI = [
    ("保健福祉課\n（障がい福祉）",
     "・年齢別・サービス別の匿名利用者一覧（利用者番号、生年月、障害支援区分、"
     "利用中のサービス、月利用日数・時間。氏名は不要）\n"
     "・サービス別の支給決定者数（請求件数ではなく実人数）\n"
     "・令和６・７年度の給付実績、令和８年度の見込\n"
     "・令和７年度末の施設入所者数\n"
     "・令和６年度の一般就労移行実績・就労定着支援利用実績（事業種別ごと）\n"
     "・村内３事業所（居宅介護・重度訪問介護・生活介護）の事業所名、指定年月、"
     "共生型サービスの指定の有無、受入可能人数\n"
     "・地域生活支援事業の事業別事業費、国県補助額、村負担額",
     "第４章・第５章の全般"),
    ("保健福祉課\n（保健・精神）",
     "・令和７年度末時点の１年以上長期入院者の実数と年齢区分"
     "（65歳未満／65歳以上／75歳以上／40歳以上の認知症の別）、入院先\n"
     "・ゲートキーパー養成講座の実施回数・受講者数\n"
     "・Ｋ６を活用した心の状態の把握の実施状況",
     "第４章２（２）"),
    ("住民課",
     "・年齢階級別の障がい者手帳所持者数（身体は０〜17／18〜39／40〜64／65〜74／75歳以上、"
     "療育・精神は18歳未満／18〜64／65歳以上）\n"
     "・令和６〜８年度の手帳所持者数（身体・療育・精神）\n"
     "・年齢３区分別人口の最新値",
     "第２章"),
    ("教育委員会",
     "・特別支援学校・特別支援学級の在籍者（学年、卒業予定年度、卒業後の進路の見通し）\n"
     "・令和６・７年度に卒業した方の現在の利用状況\n"
     "・医療的ケア児の受入体制の整備状況",
     "第２章６・第５章５（１）"),
    ("こども関係部局",
     "・こども家庭センターの設置状況と相談導線\n"
     "・保育所・認定こども園・幼稚園における障がい児の受入体制\n"
     "・障害児通所支援の利用児童の学年と就学予定年度\n"
     "・令和６・７年度に児童発達支援が減り放課後等デイサービスが増えた理由"
     "（受給者番号ベースで同一児童の就学移行かどうか）",
     "第４章２（４）・第５章１（５）・第７章４"),
    ("総務課（防災）",
     "・個別避難計画の策定件数\n・避難行動要支援者名簿における障がい者の登載状況",
     "第３章（第４次障がい者計画の指標）"),
    ("総務課（人事・契約）",
     "・障がい者就労施設等からの物品等の調達件数・調達額",
     "第３章（同上）"),
    ("財政担当",
     "・障害者自立支援給付費負担金、障害児施設給付費負担金、"
     "地域生活支援事業費補助金の令和６・７年度交付額\n"
     "・村の一般財源負担額、国庫負担基準額の超過の有無\n"
     "・村単独事業（重度心身障害者医療費助成等）の対象者数と事業費",
     "第５章７（費用の推移と財源構成）"),
]


def shiryo_irai():
    doc = new_doc("計画策定に係る庁内各課への依頼事項", "庁議　資料４")

    para(doc,
         "第８期障がい福祉計画・第４期障がい児福祉計画の策定にあたり、"
         "各課にご提供をお願いする資料です。"
         "本計画では、サービス見込量を過去実績の伸び率ではなく利用者を個別に積み上げる方法で"
         "算定します。本村は利用者が数人のサービスが多く"
         "（例：施設入所支援４人、居宅介護は令和７年度 月"
         f"{_dk('居宅介護', '人／月')[1]}人）、"
         "利用者１人の増減が２割を超える変化となり、"
         "過去実績の伸び率では捉えられないためです。"
         "このため、個々の利用者の状況が分かる資料が必要となります。")

    h2(doc, "依頼事項")
    rows = [["課", "依頼事項", "計画の記載箇所"]]
    for a, b, c in IRAI:
        rows.append([a, b, c])
    table(doc, rows, [1900, 6900, 1600], size=8.5)

    h2(doc, "個人情報の取扱い")
    for t in [
        "利用者一覧は氏名を不要とし、利用者番号による匿名の形式でご提供ください。",
        "受渡しの方法（手交・専用回線等）と保管・廃棄の手順は、"
        "契約約款及び村の個人情報保護条例に従って別途取り決めます。",
        "集計・公表にあたっては、対象者が少数となる項目について"
        "個人が特定されないよう秘匿基準（ｎ＜５）を設けます。",
    ]:
        bullet(doc, t)

    note(doc,
         "上記のうち、年齢別・サービス別の匿名利用者一覧とサービス別の支給決定者数は、"
         "サービス見込量の算定に不可欠なもので、これがないと第５章の見込量が確定しません。"
         "令和８年８月末を目途にご提供をお願いしていたものです。"
         "このうち給付の台帳は令和８年10月22日にご提供いただき、"
         "計画素案 第５章の実績の表に反映しました。"
         "年齢到達の予定（18歳・65歳）の一覧は、村への照会の第２便"
         "（Ｍ-15・Ｍ-20）でお願いしています。")
    return doc, "06_庁議資料4_庁内各課への依頼事項.docx"


# ============================================================
# 自己点検（令和8年10月24日に新設）
#   本スクリプトには点検が無く、協議会資料・庁議資料が正本と食い違った
#   まま残っていた（他ＡＩの横断レビュー RT-05・RT-07 ほか）。
# ============================================================
def _doc_text(doc):
    out = [p.text for p in doc.paragraphs]
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                out.append(c.text)
    return "\n".join(out)


def shousetsu_nai(text, midashi):
    """「第５章７」「第４章２（４）」の形の参照のうち、正本に無い章節を返す。"""
    zen = "０１２３４５６７８９"
    hankaku = str.maketrans(zen, "0123456789")
    nai = []
    for m in re.finditer(r"第([０-９0-9]+)章([０-９0-9]+)", text):
        sho, setsu = int(m.group(1).translate(hankaku)), int(m.group(2).translate(hankaku))
        if sho > len(midashi):
            nai.append(m.group(0))
            continue
        bangou = [int(re.match(r"^([０-９0-9]+)", x).group(1).translate(hankaku))
                  for x in midashi[sho - 1][1] if re.match(r"^[０-９0-9]+", x)]
        if setsu not in bangou:
            nai.append(m.group(0))
    return nai


def verify(docs):
    import kitashiobara_anke as A
    from build_kitashiobara_kakunin_kubun import KUBUN as KK
    ng = []
    tx = {name: _doc_text(doc) for doc, name in docs}
    zen = "\n".join(tx.values())
    midashi = honpon_midashi()
    # ① 第４次障がい者計画の扱い（正本は中間見直し）
    if "改定対象外" in zen:
        ng.append("「改定対象外」が残っている（正本は中間見直し）")
    for name in ("01_協議会資料1_計画策定の概要.docx",
                 "05_庁議資料2_関連計画との整合.docx"):
        if "中間見直し" not in tx[name]:
            ng.append(f"{name} に中間見直しの記述が無い")
    # ② 資料1の章立てが正本の章と一致
    sho = [r for r in tx["01_協議会資料1_計画策定の概要.docx"].split("\n")
           if re.match(r"^第[１-８]章　", r)]
    if sho != [h for h, _ in midashi]:
        ng.append(f"資料1の章立てが正本と違う: {sho[:3]}…")
    # ③ 工程（パブリックコメントは案Ａ。仮置きの日を出さない）
    if wareki(PABU_KAISHI) not in tx["01_協議会資料1_計画策定の概要.docx"]:
        ng.append("資料1にパブリックコメントの開始日が無い")
    for g in ("２月８日", "２月22日", "11月18日", "12月16日", "１月15日",
              "２月24日"):
        if g in zen:
            ng.append(f"従前の工程又は仮置きの日が出ている: {g}")
    # ④ 評価は令和８年度末見込み（暫定）。「最終評価」と書かない
    if "最終評価" in tx["03_協議会資料3_前回計画の評価.docx"]:
        ng.append("資料3に「最終評価」が残っている")
    # ⑤ 社内の呼び名（正本）を委員会資料に出さない
    if "正本" in zen:
        ng.append("委員会資料に社内の呼び名「正本」が出ている")
    # ⑥ 章節の参照が正本に実在する
    nai = shousetsu_nai(zen, midashi)
    if nai:
        ng.append(f"正本に無い章節を参照している: {sorted(set(nai))}")
    # ⑦ 照会番号が確認事項の区分に実在する
    aru = {r[0] for r in KK}
    #   「ＳＳ-1」（指標の追記の番号）を「Ｓ-1」と読まないよう前を見る
    for m in sorted(set(re.findall(r"(?<![Ａ-Ｚ])[ＭＳ]-\d+", zen))):
        if m not in aru:
            ng.append(f"確認事項に無い照会番号: {m}")
    # ⑧ アンケートの割合が38件の値で残っていない
    for nm in A.M:
        kyu, shin = A.pc(nm, kyu=True), A.pc(nm)
        if kyu != shin and kyu in zen and nm != "外出_冬季_裏磐梯":
            ng.append(f"38件の割合が残っている: {nm} {kyu}")
    # ⑨ 実績・高齢化率が正本の値である
    if honpon_koureika() not in tx["05_庁議資料2_関連計画との整合.docx"]:
        ng.append("庁議資料2の高齢化率が正本の表の値でない")
    if "Ａ型（１人）" in zen or "令和６年度２件・令和７年度３件" in zen:
        ng.append("直す前の実績（Ａ型１人・短期入所の件数）が残っている")
    # ⑩ 否定試験73：正本に無い章節を参照すると⑥が止めること
    if not shousetsu_nai("第５章12", midashi):
        ng.append("否定試験73: 正本に無い章節（第５章12）を見逃した")
    if ng:
        print("自己点検 不合格:")
        for x in ng:
            print("   -", x)
        raise SystemExit(1)
    print(f"  自己点検: 第４次計画は中間見直し／章立て{len(midashi)}章が正本と一致／"
          "工程は案Ａ（仮置きの日なし）／暫定評価／社内の呼び名なし／"
          "章節の参照と照会番号が実在／38件の割合なし／"
          "実績・高齢化率は正本の値／否定試験73")


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    shigen = load_shigen()
    missing = [s for s in SHIGEN_ORDER if s not in shigen]
    if missing:
        raise SystemExit(f"障害福祉資源に無いサービスがあります: {missing}")

    docs = [
        shiryo_gaiyou(),
        shiryo_shishin(),
        shiryo_hyouka(),
        shiryo_teikyou(shigen),
        shiryo_kanren(),
        shiryo_irai(),
    ]
    verify(docs)
    for doc, name in docs:
        path = os.path.join(OUT_DIR, name)
        doc.save(path)
        d = docx.Document(path)
        n_p = len(d.paragraphs)
        n_t = len(d.tables)
        print(f"  作成: {name}　段落{n_p}／表{n_t}")


if __name__ == "__main__":
    main()
