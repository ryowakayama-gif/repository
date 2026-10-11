# -*- coding: utf-8 -*-
"""
北塩原村　正本（他メンバー版）への移植

出力: output/北塩原村_計画素案_正本_移植後.docx

経緯
  令和8年10月9日に、他メンバー版の計画素案を正本として進めることが
  決まった。申し送り（北塩原村_申し送り_他メンバー版への移植_20261009.docx）に
  まとめた移植20件と、論点メモからの修正のうちデータ待ちでないもの、
  村資料点検（北塩原村_村資料点検_20261010.xlsx）の修正案のうち
  村の方針に明記があって回答を待たずに書ける5件を、
  正本に実際に入れたものが本書の出力である。
  移植17〜20（県計画との整合・年齢到達・関係機関との連携・費用と財源）は
  網羅性点検の残りのうち、村・県の回答を待たずに入れられるものである。
  第5章に足す3節は、既存の第5章1〜4を番号で参照している箇所があるため、
  間に挿し込まず末尾に足す。

作り
  ・入れる文章は申し送りの生成器（build_kitashiobara_ishoku_okurijo.py）と
    村資料点検の生成器（build_kitashiobara_murashiryo.py の MURA_EDITS）から
    読み込む。指示書と出力が食い違わないようにするため、文章はそちらを
    唯一の出所とする。
  ・挿入位置は申し送りが示した「直前の段落の末尾」を正本の現物から探す。
    見つからなければ止める。
  ・書式は正本に揃える（本文幅9,412dxa、データ表の罫線4F8A6E・
    表頭2D7A57・1列目EAF4EE、注記は罫線3D86C6・地色EEF5FB）。
    申し送りの列幅は本文幅15,200を前提にしているため、9,412へ按分する。
  ・正本の既存の記述は消さない。すべて追加である。

入れていないもの
  村・県・他メンバーの回答を待つ論点（人口推計の元データ、手帳所持者数の
  実数、グループホームの設置、排泄管理支援用具の増減、成年後見の類型別、
  社会福祉法改正への対応、防災の取組の状況ほか）は入れていない。
  一覧は出力の末尾ではなく、北塩原村_論点整理_20261009.xlsx による。
"""

import copy
import os
import re
import sys

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from docx.table import Table
from docx.text.paragraph import Paragraph

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_kitashiobara_ishoku_okurijo import ISHOKU  # noqa: E402
from build_kitashiobara_murashiryo import MURA_EDITS  # noqa: E402
from build_kitashiobara_juten12 import JUTEN_EDITS  # noqa: E402
from build_kitashiobara_hp_shisaku import HP_EDITS  # noqa: E402
from build_kitashiobara_hp_3bunya import HP3_EDITS  # noqa: E402
from build_kitashiobara_redteam_kyukyu import RT_EDITS  # noqa: E402
from build_kitashiobara_mece_houshu import ME_EDITS, KAKIKAE  # noqa: E402
from build_kitashiobara_mikomi_redteam import SA_EDITS  # noqa: E402
from build_kitashiobara_anke_mikomi import AN_EDITS  # noqa: E402
from build_kitashiobara_kaigo_seigo import (  # noqa: E402
    KG_EDITS, KG_KAKIKAE,
)
from build_kitashiobara_kouiki_redteam import (  # noqa: E402
    KO_EDITS, KO_KAKIKAE,
)
from build_kitashiobara_jouhou_kouhyou import KJ_EDITS  # noqa: E402
from build_kitashiobara_ishi_kettei import IK_EDITS  # noqa: E402
from build_kitashiobara_yosan_mece import YS_EDITS  # noqa: E402
from build_kitashiobara_kyoseigata import KY_EDITS  # noqa: E402
from build_kitashiobara_kojino import KN_EDITS  # noqa: E402
from build_kitashiobara_chiikishien import CS_EDITS  # noqa: E402
from build_kitashiobara_graph import GRAPHS, OUT_DIR as ZU_DIR  # noqa: E402
from build_kitashiobara_zuhyo_bangou import ZU, HYO  # noqa: E402

REPO_ROOT = os.path.dirname(os.path.abspath(__file__))
SRC = (f"{REPO_ROOT}/source/他メンバー素案_20261008/"
       "北塩原村_第8期障がい福祉計画_第4期障がい児福祉計画_素案_他メンバー版.docx")
OUT_FILE = f"{REPO_ROOT}/output/北塩原村_計画素案_正本_移植後.docx"

# 正本の書式
BODY_W = 9412          # 本文幅（dxa）
SRC_W = 15200          # 申し送りの列幅が前提にしている幅
BORDER = "4F8A6E"      # データ表の罫線
HEAD_FILL = "2D7A57"   # データ表の表頭
COL1_FILL = "EAF4EE"   # データ表の1列目
NOTE_BORDER = "3D86C6"
NOTE_FILL = "EEF5FB"
INDENT = Pt(11)        # 本文の字下げ（正本と同じ）

changes = []


# ===========================================================================
# 書式ヘルパー（正本に揃える）
# ===========================================================================
def _el(tag, **attrs):
    e = OxmlElement(tag if ":" in tag else f"w:{tag}")
    for k, v in attrs.items():
        e.set(qn(f"w:{k}"), str(v))
    return e


def _borders(color):
    b = _el("tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b.append(_el(side, val="single", sz="8", space="0", color=color))
    return b


def _shade(cell, fill):
    cell._tc.get_or_add_tcPr().append(
        _el("shd", val="clear", color="auto", fill=fill))


def _detach(doc, element):
    """doc.add_* で作った要素を本文末から外し、挿入に使えるようにする。"""
    doc.element.body.remove(element)
    return element


def make_para(doc, text, bold=False, indent=True, color=None, size=None):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.font.bold = bold
    if size:
        r.font.size = Pt(size)
    if color:
        r.font.color.rgb = RGBColor.from_string(color)
    if indent:
        p.paragraph_format.first_line_indent = INDENT
    return _detach(doc, p._p)


def make_sub(doc, text):
    """小見出し。正本の組み方に合わせてスタイルを選ぶ。

      「３　…」のような節の見出し  → Heading 2（目次に出る）
      「（７）…」のような項の見出し → 小見出し スタイル
      「≪…≫」                    → Normal の太字
    """
    t = text.strip()
    if re.match(r"^[０-９0-9]+[　 ]", t):
        p = doc.add_paragraph(style="Heading 2")
        p.add_run(t)
        return _detach(doc, p._p)
    if t.startswith("（"):
        p = doc.add_paragraph(style="小見出し")
        p.add_run(t)
        return _detach(doc, p._p)
    return make_para(doc, t, bold=True, indent=False)


def make_table(doc, rows, widths):
    ncol = len(rows[0])
    w = [max(600, round(x * BODY_W / SRC_W)) for x in widths]
    w[-1] += BODY_W - sum(w)          # 合計を本文幅に合わせる
    t = doc.add_table(rows=len(rows), cols=ncol)
    tblpr = t._tbl.tblPr
    for old in tblpr.findall(qn("w:tblBorders")):
        tblpr.remove(old)
    tblpr.append(_borders(BORDER))
    for old in tblpr.findall(qn("w:tblW")):
        tblpr.remove(old)
    tblpr.append(_el("tblW", w=BODY_W, type="dxa"))
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = t.cell(ri, ci)
            tcpr = cell._tc.get_or_add_tcPr()
            for old in tcpr.findall(qn("w:tcW")):
                tcpr.remove(old)
            tcpr.append(_el("tcW", w=w[ci], type="dxa"))
            para = cell.paragraphs[0]
            para.paragraph_format.space_after = Pt(0)
            run = para.add_run(str(val))
            run.font.size = Pt(9.5)
            if ri == 0:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                _shade(cell, HEAD_FILL)
            else:
                if ci == 0:
                    _shade(cell, COL1_FILL)
                elif ncol >= 4:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return _detach(doc, t._tbl)


def make_note(doc, text):
    t = doc.add_table(rows=1, cols=1)
    tblpr = t._tbl.tblPr
    for old in tblpr.findall(qn("w:tblBorders")):
        tblpr.remove(old)
    tblpr.append(_borders(NOTE_BORDER))
    for old in tblpr.findall(qn("w:tblW")):
        tblpr.remove(old)
    tblpr.append(_el("tblW", w=BODY_W, type="dxa"))
    cell = t.cell(0, 0)
    tcpr = cell._tc.get_or_add_tcPr()
    tcpr.append(_el("tcW", w=BODY_W, type="dxa"))
    _shade(cell, NOTE_FILL)
    para = cell.paragraphs[0]
    para.paragraph_format.space_after = Pt(0)
    run = para.add_run(text)
    run.font.size = Pt(10)
    return _detach(doc, t._tbl)


def make_empty(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    return _detach(doc, p._p)


# ===========================================================================
# 位置の特定
# ===========================================================================
def _norm(s):
    return re.sub(r"[\s　]+", "", s)


def find_para(doc, text):
    """本文の段落のうち、文字列が一致するものを1つだけ見つける。"""
    key = _norm(text)
    hits = [p for p in doc.paragraphs if _norm(p.text) == key]
    if not hits:
        hits = [p for p in doc.paragraphs if key and key in _norm(p.text)]
    if len(hits) != 1:
        raise LookupError(f"挿入位置が一意に決まりません（{len(hits)}件）: "
                          f"{text[:44]}")
    return hits[0]._p


def insert_after(anchor, elements):
    cur = anchor
    for e in elements:
        cur.addnext(e)
        cur = e
    return cur


def find_note_cell(doc, prefix):
    """1x1の注記ボックスのうち、本文が prefix で始まるものを返す。"""
    key = _norm(prefix)
    hits = []
    for t in doc.tables:
        if len(t.rows) == 1 and len(t.columns) == 1:
            if _norm(t.cell(0, 0).text).startswith(key):
                hits.append(t.cell(0, 0))
    if len(hits) != 1:
        raise LookupError(f"注記ボックスが一意に決まりません（{len(hits)}件）: "
                          f"{prefix[:30]}")
    return hits[0]


# ===========================================================================
# 移植の割り当て
#   既定は ISHOKU の ichi を挿入位置とし、ブロックをその直後に入れる。
#   位置が前の移植と同じものは、前の移植の後ろに続けて入れる。
# ===========================================================================
# 位置が特殊なもの
#   移植6・7・8後半　第5章2（6）の末尾に続けて入れる
#   移植8　　　　　　2か所に分かれる（前半は第4章2（7）、後半は第5章2）
#   移植13　　　　　 ≪国の基本指針≫欄への追記を含む
#   移植3・13・論点R-12　第4章2（3）に続けて入れる
#   移植18・19・20　第5章の末尾に3節を続けて入れる

# 移植8の前半の挿入位置（第4章2（7）の【目標設定の考え方】の末尾）
ANCHOR_8A = ("村単独での人材確保策には限りがあることから、"
             "福島県が設置するワンストップ窓口や各種研修を積極的に活用し、"
             "村職員及び村内関係者の受講を促進します。")

# 移植6・7・8後半が続く位置（第5章2（6）の末尾）
ANCHOR_5_2 = ("福島県が実施する研修や人材確保・生産性向上に関する支援策を活用し、"
              "サービス提供に携わる人材の確保・定着とサービスの質の向上を図ります。"
              "また、障害者虐待の防止や、災害時におけるサービス提供の継続についても、"
              "事業所や関係機関と連携して取り組みます。")

# 移植18・19・20（第5章に足す3節）が続く位置（第5章4（2）の末尾）
#   既存の第5章1〜4を番号で参照している箇所があるため、
#   間に挿し込まず第5章の末尾に足す。
ANCHOR_5_4 = ("現在実施中のサービスについては、質の向上の促進に努めます。"
              "理解促進研修・啓発事業、自発的活動支援事業、移動支援事業など、"
              "これまでの実績がない事業については、利用者ニーズの把握や"
              "近隣市町村との連携によるサービス提供事業者の確保に努めます。")


# ===========================================================================
# 論点メモからの修正（データ待ちでないもの）
#   番号, 挿入位置, 入れるもの
# ===========================================================================
MEMO_EDITS = [
    ("R-12", "第４章２（３）　優先調達の進め方",
     "なお、国等による障害者就労施設等からの物品等の調達の推進等に関する"
     "法律（障害者優先調達推進法）第９条により、本村は毎年度、"
     "障害者就労施設等からの物品等の調達の方針を作成し、公表します。"
     "あわせて、当該年度の調達の実績も公表し、"
     "第３章の目標値（障がい者就労施設等からの物品調達件数 ３件）の"
     "達成状況を確認します。",
     "CHAIN13"),
    ("R-15", "第４章２（６）　相談の窓口の二層の役割分担",
     "本村の相談の窓口は二層で考えます。"
     "一層目は、どなたでも気軽に相談できる窓口であり、"
     "村保健福祉課及び相談支援事業所（地域生活支援センターいなわしろ）が"
     "これに当たります。"
     "二層目は、専門的な助言や困難な事例への対応を担う窓口であり、"
     "会津北部４町村で広域設置を目指す基幹相談支援センターがこれに当たります。"
     "一層目で受けた相談のうち、権利擁護、虐待、"
     "複数の制度にまたがる調整を要するものを二層目につなぐ流れを整えます。",
     "本村では現在、該当する事例はありませんが、今後も相談支援事業所との"
     "連携により相談支援専門員を確保し、令和11年度末まで０件を維持します。"),
    ("R-19", "第４章２（４）　インクルーシブ教育との接続",
     "これらの取組みは、障がいのある子どもとない子どもが"
     "可能な限り共に学ぶインクルーシブ教育の考え方と同じ方向にあります。"
     "インクルージョン推進のための協議の場では、"
     "保育所・幼稚園・学校における受入れの状況と必要な配慮を共有し、"
     "第３章の基本施策④教育・育成と一体で進めます。",
     "CHAIN4"),
    ("R-20", "第７章１　母子保健との連携の具体化",
     "また、健康21プランの乳児家庭全戸訪問事業（こんにちは赤ちゃん事業）や"
     "乳幼児健康診査は、障がいの早期発見と早期療育につながる入口です。"
     "国民健康保険のデータヘルス計画による生活習慣病の重症化予防とあわせて、"
     "保健部門と障がい福祉部門が情報を共有して取り組みます。",
     "健康21プランの母子保健事業（妊婦全戸訪問、産後ケア事業等）と、"
     "本計画の医療的ケア児支援体制の整備を、保健師を通じて連携して進めます。"),
    ("R-21", "第５章３　障害者週間に合わせた啓発",
     "理解促進研修・啓発事業は、これまで実績がありません。"
     "実施の方法としては、毎年12月３日から９日までの障害者週間"
     "（障害者基本法第９条）に合わせて、"
     "村の施設における展示、広報誌の特集、"
     "福祉ボランティアの活動の紹介を行うことを想定しています。",
     "理解促進研修・啓発事業については、第４次北塩原村障がい者計画の"
     "目標である「差別や偏見を感じている障がいがある人の割合」の減少に向け、"
     "各年度１件の実施を見込みます。"),
]

# 用語解説に加える語（R-7）
YOUGO_ADD = [
    ("療育手帳", "知的障がいのある方に交付される手帳。"
     "身体障害者手帳（身体障害者福祉法第15条）や"
     "精神障害者保健福祉手帳（精神保健福祉法第45条）と異なり、"
     "法律上の根拠はなく、国の通知に基づく制度として"
     "都道府県・指定都市が交付する。"
     "福島県は重度をＡ判定、中度・軽度をＢ判定としている。"),
    ("難病等", "治療方法が確立していない疾病その他の特殊の疾病であって"
     "政令で定めるもの。障害者総合支援法第４条第１項により、"
     "手帳を所持していなくても障がい福祉サービスの対象となる。"
     "対象疾病は令和７年４月から376疾病。"),
    ("基盤整備量", "精神病床に１年以上入院している方が地域生活へ移行した"
     "場合に、地域で受け止めるために必要となる障がい福祉サービス等の量。"
     "国の基本指針 別表第二 三㈠⑤により、"
     "市町村障害福祉計画で定めなければならない事項とされている。"),
]


# ===========================================================================
# 移植の実行
# ===========================================================================
def build_blocks(doc, rec, skip_first_h3=False, h3_override=None):
    """ISHOKU の nakami を正本の書式の要素に組み立てる。"""
    out = []
    for i, blk in enumerate(rec["nakami"]):
        kind = blk[0]
        if kind == "h3":
            text = blk[1]
            if h3_override is not None:
                text = h3_override
            out.append(make_sub(doc, text))
        elif kind == "p":
            out.append(make_para(doc, blk[1]))
        elif kind == "tbl":
            out.append(make_table(doc, blk[1], blk[2]))
        elif kind == "note":
            out.append(make_note(doc, blk[1]))
    out.append(make_empty(doc))
    return out


def apply_ishoku(doc):
    byno = {r["no"]: r for r in ISHOKU}
    tail = {}      # 挿入位置ごとの「最後に入れた要素」

    def put(anchor_key, anchor_text, elements):
        if anchor_key in tail:
            cur = tail[anchor_key]
        else:
            cur = find_para(doc, anchor_text)
        tail[anchor_key] = insert_after(cur, elements)

    for no in [r["no"] for r in ISHOKU]:
        rec = byno[no]

        if no in ("18", "19", "20"):
            # 第5章に足す3節。末尾に続けて入れる（節の番号をずらさない）
            put("5-4", ANCHOR_5_4, build_blocks(doc, rec))
        elif no == "７":
            put("5-2", ANCHOR_5_2, build_blocks(doc, rec))
        elif no == "８":
            # 前半（意思決定支援）は第4章2（7）、後半（経営基盤）は第5章2の続き
            a = rec["nakami"]
            # 2つめの見出し（第5章に入れる分）で前後に切る。
            # 位置を数で書くと、文章を足したときに切る場所がずれる。
            midashi = [i for i, b in enumerate(a) if b[0] == "h3"]
            if len(midashi) != 2:
                raise LookupError(f"移植８の見出しが2つでない: {midashi}")
            front = a[midashi[0] + 1:midashi[1]]   # p … note
            back = a[midashi[1] + 1:]              # p, tbl, note
            el = []
            for b in front:
                el.append(make_para(doc, b[1]) if b[0] == "p"
                          else make_note(doc, b[1]))
            el.append(make_empty(doc))
            put("4-2-7", ANCHOR_8A, el)
            # 見出しは ISHOKU の指示書きを外して作る（書き写さない）
            h = a[midashi[1]][1].split("】", 1)[1]
            el2 = [make_sub(doc, h)]
            for b in back:
                if b[0] == "p":
                    el2.append(make_para(doc, b[1]))
                elif b[0] == "tbl":
                    el2.append(make_table(doc, b[1], b[2]))
                elif b[0] == "note":
                    el2.append(make_note(doc, b[1]))
            el2.append(make_empty(doc))
            put("5-2", ANCHOR_5_2, el2)
        elif no == "13":
            # ①≪国の基本指針≫欄への追記 ②本文・表・注記
            cell = find_note_cell(doc, "≪国の基本指針≫\n① 一般就労への移行者数")
            add = rec["nakami"][0][1].split("】", 1)[1]
            p = cell.add_paragraph()
            r = p.add_run(add)
            r.font.size = Pt(10)
            el = []
            for b in rec["nakami"][1:]:
                if b[0] == "p":
                    el.append(make_para(doc, b[1].split("】", 1)[1]
                                        if b[1].startswith("【") else b[1]))
                elif b[0] == "tbl":
                    el.append(make_table(doc, b[1], b[2]))
                elif b[0] == "note":
                    el.append(make_note(doc, b[1]))
            el.append(make_empty(doc))
            put("4-2-3", rec["ichi"], el)
        elif no == "６":
            put("5-2", ANCHOR_5_2, build_blocks(doc, rec))
        elif no == "３":
            put("4-2-3", rec["ichi"], build_blocks(doc, rec))
        else:
            put(f"a{no}", rec["ichi"], build_blocks(doc, rec))
        changes.append(f"移植{no}：{rec['saki']}")
    return tail


def find_table(doc, header):
    """表頭が一致する表を1つだけ見つける。"""
    hits = []
    for t in doc.tables:
        hdr = [c.text.strip() for c in t.rows[0].cells]
        if hdr[:len(header)] == list(header):
            hits.append(t)
    if len(hits) != 1:
        raise LookupError(f"表が一意に決まりません（{len(hits)}件）: {header}")
    return hits[0]


def _grid(tbl):
    """tblGrid の列幅（dxa）。"""
    g = tbl.find(qn("w:tblGrid"))
    if g is None:
        return []
    return [int(c.get(qn("w:w"))) for c in g.findall(qn("w:gridCol"))]


def _set_cell_text(tc, text):
    """セルの文字列を書き替える（書式は元のまま使う）。"""
    ps = tc.findall(qn("w:p"))
    for extra in ps[1:]:
        tc.remove(extra)
    p = ps[0]
    runs = p.findall(qn("w:r"))
    keep = runs[0] if runs else None
    for r in runs[1:]:
        p.remove(r)
    if keep is None:
        keep = OxmlElement("w:r")
        p.append(keep)
    for t in keep.findall(qn("w:t")):
        keep.remove(t)
    for br in keep.findall(qn("w:br")):
        keep.remove(br)
    first = True
    for line in text.split("\n"):
        if not first:
            keep.append(OxmlElement("w:br"))
        t = OxmlElement("w:t")
        t.text = line
        t.set(qn("xml:space"), "preserve")
        keep.append(t)
        first = False


def _set_span(tc, n, grid, start):
    """セルの gridSpan と幅を合わせる。"""
    tcpr = tc.get_or_add_tcPr()
    for old in tcpr.findall(qn("w:gridSpan")):
        tcpr.remove(old)
    if n > 1:
        tcpr.append(_el("gridSpan", val=n))
    if grid:
        w = sum(grid[start:start + n])
        for old in tcpr.findall(qn("w:tcW")):
            tcpr.remove(old)
        tcpr.append(_el("tcW", w=w, type="dxa"))


def add_rows(tbl, src_index, rows, spans=None):
    """既存の行を写して行を足す（書式・罫線・幅をそのまま引き継ぐ）。

    spans を渡すと、写した行の gridSpan を組み替える。
    """
    grid = _grid(tbl)
    src = tbl.findall(qn("w:tr"))[src_index]
    for values in rows:
        tr = copy.deepcopy(src)
        # 縦の結合は写さない（新しい行は独立させる。vMerge は呼び手が付ける）
        tcs = tr.findall(qn("w:tc"))
        if spans:
            if len(tcs) != len(spans):
                raise LookupError(
                    f"写した行のセル数が spans と合いません（{len(tcs)}≠{len(spans)}）")
            pos = 0
            for tc, n in zip(tcs, spans):
                _set_span(tc, n, grid, pos)
                pos += n
        if len(tcs) != len(values):
            raise LookupError(
                f"写した行のセル数が値の数と合いません（{len(tcs)}≠{len(values)}）")
        for tc, v in zip(tcs, values):
            _set_cell_text(tc, str(v))
        tbl.append(tr)
    return tbl


def _vmerge(tbl, rows_from_end, col=0):
    """末尾から数えた複数行の指定列を縦に結合する。"""
    trs = tbl.findall(qn("w:tr"))[-rows_from_end:]
    for i, tr in enumerate(trs):
        tc = tr.findall(qn("w:tc"))[col]
        tcpr = tc.get_or_add_tcPr()
        for old in tcpr.findall(qn("w:vMerge")):
            tcpr.remove(old)
        vm = _el("vMerge")
        if i == 0:
            vm.set(qn("w:val"), "restart")
        tcpr.append(vm)


def find_note_tbl(doc, prefix):
    """1x1の注記ボックスの表そのものを返す。"""
    cell = find_note_cell(doc, prefix)
    return cell._tc.getparent().getparent()


def next_tbl(element):
    """ある段落の次に現れる表を返す（直後に続けて入れるため）。"""
    cur = element
    while True:
        cur = cur.getnext()
        if cur is None:
            raise LookupError("後ろに表が見つかりません")
        if cur.tag.split("}")[-1] == "tbl":
            return cur


def apply_juten(doc, records=None, label="重点施策"):
    """再レビューからの追記を入れる（JUTEN_EDITS と HP_EDITS に使う）。"""
    for rec in (JUTEN_EDITS if records is None else records):
        el = []
        for blk in rec["nakami"]:
            if blk[0] in ("h2", "h3"):
                el.append(make_sub(doc, blk[1]))
            elif blk[0] == "p":
                el.append(make_para(doc, blk[1]))
            elif blk[0] == "tbl":
                el.append(make_table(doc, blk[1], blk[2]))
            elif blk[0] == "note":
                el.append(make_note(doc, blk[1]))
            else:
                raise LookupError(f"ブロックの型が未定義: {rec['no']} {blk[0]}")
        el.append(make_empty(doc))
        if rec["kata"] == "注記後":
            anchor = find_note_tbl(doc, rec["ichi"])
        elif rec["kata"] == "表の後":
            anchor = next_tbl(find_para(doc, rec["ichi"]))
        elif rec["kata"] == "本文":
            anchor = find_para(doc, rec["ichi"])
        else:
            raise LookupError(f"挿入位置の型が未定義: {rec['no']} "
                              f"{rec['kata']}")
        insert_after(anchor, el)
        changes.append(f"{label}{rec['no']}：{rec['saki']}")


def make_image(doc, path, width_cm):
    """図（画像）を本文幅に収めて入れる。"""
    from docx.shared import Cm
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    p.add_run().add_picture(path, width=Cm(width_cm))
    return _detach(doc, p._p)


def make_caption(doc, text, size=9.5, center=True):
    """図の表題・出所の行。"""
    p = doc.add_paragraph()
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(text)
    r.font.size = Pt(size)
    return _detach(doc, p._p)


def apply_graphs(doc):
    """当方が作った図を、対応する表又は段落の直後に入れる。

    kata が「表の後」（既定）なら、ichi の段落の次に現れる表の直後に、
    「本文」ならその段落の直後に入れる。
    """
    from PIL import Image
    for g in GRAPHS:
        path = f"{ZU_DIR}/{g['file']}"
        if not os.path.exists(path):
            raise LookupError(f"図が見つかりません: {path}")
        im = Image.open(path)
        cm = min(im.width / 200 * 2.54, 16.6)      # 200dpiで作っている
        el = [
            make_image(doc, path, cm),
            make_caption(doc, g["title"]),
            make_caption(doc, g["src"], size=9, center=False),
            make_empty(doc),
        ]
        kata = g.get("kata", "表の後")
        para = find_para(doc, g["ichi"])
        if kata == "表の後":
            # 表の直後に（各年４月１日現在）のような添え書きがあるときは、
            # その後に入れる（添え書きは表に属するため）
            anchor = next_tbl(para)
            while True:
                tsugi = anchor.getnext()
                if tsugi is None or tsugi.tag.split("}")[-1] != "p":
                    break
                if not _is_kakko(Paragraph(tsugi, doc)):
                    break
                anchor = tsugi
            insert_after(anchor, el)
        elif kata == "本文":
            insert_after(para, el)
        else:
            raise LookupError(f"図の挿入位置の型が未定義: {g['no']} {kata}")
        changes.append(f"図{g['no']}：{g['title']}")


# ===========================================================================
# 図表番号・表題・図表目次（規約は図表一覧06シート）
#   番号は章ごとの通し番号。表題は表の上、図の表題は図の下。
#   注記ボックス（1×1）には番号を振らない。
# ===========================================================================
SUJI = "０１２３４５６７８９"


def _zenkaku(n):
    """1桁は全角、2桁以上は半角（表記の作法）。"""
    return SUJI[n] if n < 10 else str(n)


def _bangou(kigou, no):
    """「表5-34」のような番号を「表５-34」の形に直す。"""
    m = re.match(r"^(図|表)(\d+)-(\d+)$", no)
    if not m:
        raise LookupError(f"番号の形が違う: {no}")
    return f"{m.group(1)}{_zenkaku(int(m.group(2)))}-{_zenkaku(int(m.group(3)))}"


def make_midashi(doc, text, center=False, size=10, bold=True):
    """図表の表題。"""
    p = doc.add_paragraph()
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(1)
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.font.bold = bold
    return _detach(doc, p._p)


def _is_kakko(para):
    """（単位：人）（令和８年４月１日現在）のような添え書きか。

    「（２）財源構成（令和７年度・法定負担割合による試算）」のような
    項の見出しも「（」で始まり「）」で終わるため、
    最初の「（」に対応する「）」が末尾にあることまで確かめる。
    本文スタイル以外（小見出し等）は添え書きではない。
    """
    if hasattr(para, "style") and para.style.name != "Normal":
        return False
    t = (para.text if hasattr(para, "text") else str(para)).strip()
    if len(t) < 2 or t[0] != "（" or t[-1] != "）":
        return False
    fukasa = 0
    for i, ch in enumerate(t):
        if ch == "（":
            fukasa += 1
        elif ch == "）":
            fukasa -= 1
            if fukasa == 0:
                return i == len(t) - 1
    return False


def apply_bangou(doc):
    """図と表に番号と表題を付け、出所を添える。

    表題は表の上に置くが、表のすぐ上に（単位：…）のような添え書きが
    あるときは、その上に置く（添え書きは表に属するため）。
    """
    tag = lambda e: e.tag.split("}")[-1]          # noqa: E731
    zu_i = hyo_i = 0
    sho = 0
    tsuketa = []
    for e in list(doc.element.body.iterchildren()):
        if tag(e) == "p":
            para = Paragraph(e, doc)
            t = para.text.strip()
            m = re.match(r"^第([０-９])章", t)
            if para.style.name == "Heading 1" and m:
                sho = SUJI.index(m.group(1))
            ookii = any(int(x.get("cx")) > 19 * 360000
                        for x in e.findall(".//" + qn("wp:extent")))
            if not e.findall(".//" + qn("a:blip")) or not sho or ookii:
                continue
            # --- 図 ---
            z = ZU[zu_i]
            if z["sho"] != sho:
                raise LookupError(f"図の章が合わない: {z['no']} ← 第{sho}章")
            zu_i += 1
            midashi = f"{_bangou('図', z['no'])}　{z['midashi']}"
            tsugi = e.getnext()
            tsugi_t = (Paragraph(tsugi, doc).text.strip()
                       if tsugi is not None and tag(tsugi) == "p" else "")
            if _norm(tsugi_t) == _norm(z["midashi"]):
                # 当方の図。既にある表題の行に番号を足す
                for r in Paragraph(tsugi, doc).runs:
                    r.text = ""
                Paragraph(tsugi, doc).runs[0].text = midashi
            else:
                insert_after(e, [make_midashi(doc, midashi, center=True,
                                              size=9.5)])
                if z["moto"]:
                    if tsugi is not None and tag(tsugi) == "p" \
                            and _is_kakko(Paragraph(tsugi, doc)):
                        # （令和８年４月１日現在）のような添え書きを
                        # 出所の行に置き換える（日付が二重にならないよう）
                        pp = Paragraph(tsugi, doc)
                        for r in pp.runs[1:]:
                            r.text = ""
                        pp.runs[0].text = z["moto"]
                    elif not tsugi_t.startswith("資料："):
                        insert_after(e.getnext(),
                                     [make_caption(doc, z["moto"], size=9,
                                                   center=False)])
            tsuketa.append(midashi)
        elif tag(e) == "tbl" and sho:
            tbl = Table(e, doc)
            if len(tbl.rows) == 1 and len(tbl.columns) == 1:
                continue
            h = HYO[hyo_i]
            if h["sho"] != sho:
                raise LookupError(f"表の章が合わない: {h['no']} ← 第{sho}章")
            atama = [c.text.strip() for c in tbl.rows[0].cells]
            uniq = []
            for x in atama:
                if x not in uniq:
                    uniq.append(x)
            if [_norm(x) for x in uniq[:len(h["atama"])]] != \
                    [_norm(x) for x in h["atama"]]:
                raise LookupError(f"表頭が合わない: {h['no']} ← {uniq[:2]}")
            hyo_i += 1
            midashi = f"{_bangou('表', h['no'])}　{h['midashi']}"
            # 表のすぐ上の添え書きは表に属する。その上に表題を置く
            ue = e
            while True:
                mae = ue.getprevious()
                if mae is None or tag(mae) != "p":
                    break
                if not _is_kakko(Paragraph(mae, doc)):
                    break
                ue = mae
            ue.addprevious(make_midashi(doc, midashi))
            if h["moto"]:
                shita = e
                while True:
                    tsugi = shita.getnext()
                    if tsugi is None or tag(tsugi) != "p":
                        break
                    if not _is_kakko(Paragraph(tsugi, doc)):
                        break
                    shita = tsugi          # （単位：…）は表に属する
                tsugi = shita.getnext()
                tsugi_t = (Paragraph(tsugi, doc).text.strip()
                           if tsugi is not None and tag(tsugi) == "p" else "")
                if not tsugi_t.startswith("資料："):
                    insert_after(shita, [make_caption(doc, h["moto"], size=9,
                                                      center=False)])
            tsuketa.append(midashi)
    if zu_i != len(ZU) or hyo_i != len(HYO):
        raise LookupError(f"図表の数が合わない: 図{zu_i}/{len(ZU)}・"
                          f"表{hyo_i}/{len(HYO)}")
    changes.append(f"図表番号：図{zu_i}点・表{hyo_i}点に番号と表題を付与")
    return tsuketa


def apply_zuhyo_mokuji(doc):
    """目次の後に図表目次を置く。"""
    saigo = None
    for p in doc.paragraphs:
        if p.style.name.startswith("toc "):
            saigo = p._p
    if saigo is None:
        raise LookupError("目次が見つかりません")
    el = [make_midashi(doc, "図表目次", center=True, size=12)]
    for kigou, items in (("図", ZU), ("表", HYO)):
        el.append(make_midashi(doc, f"【{kigou}】", size=10))
        for it in items:
            no = _bangou(kigou, it["no"])
            el.append(make_para(doc, f"{no}　{it['midashi']}",
                                indent=False, size=9))
    el.append(make_empty(doc))
    insert_after(saigo, el)
    changes.append(f"図表目次：図{len(ZU)}点・表{len(HYO)}点を目次の後に掲載")


def apply_kakikae(doc):
    """他メンバー版の本文を書き換える（事実として誤っているものに限る）。

    挿入ではなく置換であるため、直す前の文が本文にちょうど1つある
    ことを確かめてから書き換える。書式は先頭のランのものを残す。
    """
    for rec in list(KAKIKAE) + list(KG_KAKIKAE) + list(KO_KAKIKAE):
        atari = [p for p in doc.paragraphs
                 if _norm(rec["mae"]) in _norm(p.text)]
        if len(atari) != 1:
            raise LookupError(f"書き換え先が一意でない（{len(atari)}件）: "
                              f"{rec['no']}")
        para = atari[0]
        if _norm(para.text) != _norm(rec["mae"]):
            raise LookupError(f"書き換え先の段落に余分な字がある: {rec['no']}")
        for r in para.runs[1:]:
            r.text = ""
        para.runs[0].text = rec["ato"]
        changes.append(f"書き換え{rec['no']}：{rec['saki']}")


def apply_mura(doc):
    """村資料点検の修正案のうち、最優先の追記5件を入れる。"""
    for rec in MURA_EDITS:
        kata = rec["kata"]
        if kata == "計画の表":
            tbl = find_table(doc, ("計画名", "項目", "内容"))._tbl
            add_rows(tbl, -2, rec["gyou"])      # 直前の2行の書式を写す
            _vmerge(tbl, 2, col=0)
        elif kata == "年度表":
            el = [make_para(doc, t) for t in rec["honbun"]]
            insert_after(find_para(doc, rec["ichi"]), el)
            tbl = find_table(doc, ("年度", "R6\n(2024)"))._tbl
            # 1列目＋R6〜R8（3列）＋R9〜R13（5列）＋R14（1列）＝10列
            add_rows(tbl, -1, rec["gyou"], spans=[1, 3, 5, 1])
        elif kata == "本文":
            el = [make_para(doc, t) for t in rec["honbun"]]
            insert_after(find_para(doc, rec["ichi"]), el)
        elif kata == "項":
            el = []
            for midashi, honbun in rec["ko"]:
                el.append(make_sub(doc, midashi))
                el.append(make_para(doc, honbun))
            insert_after(find_para(doc, rec["ichi"]), el)
        else:
            raise LookupError(f"型が未定義: {rec['no']}")
        changes.append(f"村資料{rec['no']}：{rec['saki']}")


def apply_memo(doc, tail):
    """論点メモからの修正のうち、データ待ちでないものを入れる。"""
    for no, midashi, honbun, anchor in MEMO_EDITS:
        el = [make_para(doc, honbun), make_empty(doc)]
        if anchor == "CHAIN13":
            insert_after(tail["4-2-3"], el)
            tail["4-2-3"] = el[-1]
        elif anchor == "CHAIN4":
            insert_after(tail["a４"], el)
            tail["a４"] = el[-1]
        else:
            insert_after(find_para(doc, anchor), el)
        changes.append(f"論点{no}：{midashi}")

    # 用語解説に3語を加える（R-7ほか）
    yougo = None
    for t in doc.tables:
        if len(t.columns) == 2 and t.rows[0].cells[0].text.strip() == "用語":
            yougo = t
            break
    if yougo is None:
        raise LookupError("用語解説の表が見つかりません")
    for word, desc in YOUGO_ADD:
        row = yougo.add_row()
        for ci, val in enumerate((word, desc)):
            cell = row.cells[ci]
            para = cell.paragraphs[0]
            para.paragraph_format.space_after = Pt(0)
            run = para.add_run(val)
            run.font.size = Pt(9.5)
            if ci == 0:
                _shade(cell, COL1_FILL)
    changes.append(f"第８章５：用語解説に{len(YOUGO_ADD)}語を追加"
                   "（療育手帳・難病等・基盤整備量）")
    return yougo


# ===========================================================================
# 自己点検
# ===========================================================================
def count_images(doc):
    return len(doc.element.body.findall(".//" + qn("a:blip")))


def verify(doc, src_dims):
    ng = []
    parts = []
    for ch in doc.element.body.iterchildren():
        tag = ch.tag.split("}")[-1]
        if tag == "p":
            parts.append(Paragraph(ch, doc).text)
        elif tag == "tbl":
            for row in Table(ch, doc).rows:
                seen = set()
                for c in row.cells:
                    if id(c._tc) in seen:
                        continue
                    seen.add(id(c._tc))
                    parts.append(c.text)
    out = _norm("\n".join(parts))

    # ① 申し送りの文章がすべて入っていること
    miss = 0
    for rec in ISHOKU:
        for blk in rec["nakami"]:
            if blk[0] in ("h3", "p", "note"):
                t = blk[1]
                # 「【第４章２（７）の…に入れる】」のような指示書きだけの
                # 見出しは出力に入れないため、照合の対象から外す
                if blk[0] == "h3" and t.startswith("【第") and \
                        t.endswith("に入れる】"):
                    continue
                if t.startswith("【") and "】" in t:
                    t = t.split("】", 1)[1]
                if _norm(t) not in out:
                    miss += 1
                    if miss <= 6:
                        ng.append(f"移植{rec['no']}の文章が出力にない: {t[:38]}")
            elif blk[0] == "tbl":
                for row in blk[1]:
                    for v in row:
                        if v and _norm(str(v)) not in out:
                            miss += 1
                            if miss <= 6:
                                ng.append(
                                    f"移植{rec['no']}の表のセルが出力にない: {v[:28]}")
    if miss > 6:
        ng.append(f"…ほか{miss - 6}件")

    # ①b 移植8後半の見出し（指示書きを外したもの）が入っていること
    h8 = [b[1] for b in next(r for r in ISHOKU if r["no"] == "８")["nakami"]
          if b[0] == "h3"]
    for h in h8:
        if "】" in h and not h.endswith("に入れる】"):
            if _norm(h.split("】", 1)[1]) not in out:
                ng.append(f"移植8の見出しが出力にない: {h[:40]}")

    # ② 論点メモからの修正が入っていること
    for no, _m, honbun, _a in MEMO_EDITS:
        if _norm(honbun) not in out:
            ng.append(f"論点{no}の文章が出力にない")
    for word, desc in YOUGO_ADD:
        if _norm(desc) not in out:
            ng.append(f"用語解説に入っていない: {word}")

    # ②b 村資料点検の修正案（最優先の追記5件）が入っていること
    for rec in MURA_EDITS:
        for t in rec.get("honbun", []):
            if _norm(t) not in out:
                ng.append(f"村資料{rec['no']}の文章が出力にない: {t[:38]}")
        for midashi, honbun in rec.get("ko", []):
            for t in (midashi, honbun):
                if _norm(t) not in out:
                    ng.append(f"村資料{rec['no']}の文章が出力にない: {t[:38]}")
        for row in rec.get("gyou", []):
            for v in row:
                if v and _norm(v) not in out:
                    ng.append(f"村資料{rec['no']}の表のセルが出力にない: {v[:28]}")

    # ②c 再レビューからの追記が入っていること
    for rec in (list(JUTEN_EDITS) + list(HP_EDITS) + list(HP3_EDITS)
                + list(RT_EDITS) + list(ME_EDITS) + list(SA_EDITS)
                + list(AN_EDITS) + list(KG_EDITS)
                + list(KO_EDITS) + list(KJ_EDITS)
                + list(IK_EDITS) + list(YS_EDITS)
                + list(KY_EDITS) + list(KN_EDITS)
                + list(CS_EDITS)):
        for blk in rec["nakami"]:
            if blk[0] in ("h2", "h3", "p", "note"):
                if _norm(blk[1]) not in out:
                    ng.append(f"重点施策{rec['no']}の文章が出力にない: "
                              f"{blk[1][:38]}")
            elif blk[0] == "tbl":
                for row in blk[1]:
                    for v in row:
                        if v and _norm(str(v)) not in out:
                            ng.append(f"重点施策{rec['no']}の表のセルが"
                                      f"出力にない: {v[:28]}")

    # ②d 当方が作った図とその表題・出所が入っていること
    n_img = count_images(doc)
    want = src_dims[2] + len(GRAPHS)
    if n_img != want:
        ng.append(f"図の数が合わない（画像{n_img}点。"
                  f"正本{src_dims[2]}点＋当方{len(GRAPHS)}点＝{want}点）")
    for g in GRAPHS:
        if _norm(g["title"]) not in out:
            ng.append(f"図の表題が出力にない: {g['title']}")
        if _norm(g["src"]) not in out:
            ng.append(f"図の出所が出力にない: {g['src']}")

    # ③ 正本の既存の記述が消えていないこと
    keep = [
        "障がいのあるなしに関わらず、お互いの人格や個性を尊重し",
        "令和８年４月１日現在2,316人",
        "身体障害者手帳が127人",
        "有効回答率", "22.1", "23.0",
        "会津北部地域生活支援拠点",
        "のぞまないセルフプラン",
        "居宅介護", "生活介護", "就労継続支援Ｂ型", "共同生活援助",
        "施設入所支援", "計画相談支援", "放課後等",
        "成年後見制度の利用促進",
        "北塩原村第五次総合振興計画",
        "北塩原村障がい者自立支援協議会委員名簿",
    ]
    for w in keep:
        if _norm(w) not in out:
            ng.append(f"正本の記述が消えている: {w}")

    # ④ 見込量表が残っていること（8列の実績・見込量の表）
    n_mikomi = 0
    for t in doc.tables:
        hdr = [c.text.strip() for c in t.rows[0].cells]
        if hdr[:2] == ["サービス種別", "単位"] and len(t.rows) > 2:
            n_mikomi += 1
    if n_mikomi < 10:
        ng.append(f"見込量表が{n_mikomi}表しかない（正本は10表）")

    # ⑤ 表記の作法（正本に揃える）
    added = []
    for rec in ISHOKU:
        for blk in rec["nakami"]:
            if blk[0] in ("h3", "p", "note"):
                added.append(blk[1])
            elif blk[0] == "tbl":
                added += [str(v) for row in blk[1] for v in row]
    for _n, _m, honbun, _a in MEMO_EDITS:
        added.append(honbun)
    added += [d for _w, d in YOUGO_ADD]
    joined = "\n".join(added)
    if "〜" in joined:
        ng.append("波ダッシュ（〜）が混じっている")
    if "か所" in joined or "ヶ所" in joined:
        ng.append("「カ所」に揃っていない")
    if re.search(r"令和[0-9]年", joined):
        ng.append("令和の1桁が半角になっている")
    if "%" in joined:
        ng.append("半角の％が混じっている")

    # ⑤b 加えた見出しが正本と同じスタイルで組まれていること
    #     見出しは ISHOKU から機械で拾う（書き写すと取り違える）
    want_h2, want_sub = [], []
    for rec in ISHOKU:
        for blk in rec["nakami"]:
            if blk[0] != "h3":
                continue
            t = blk[1]
            if t.startswith("【") and t.endswith("に入れる】"):
                continue
            if "】" in t:
                t = t.split("】", 1)[1]
            if t.startswith("（"):
                want_sub.append(t)
            elif re.match(r"^[０-９0-9]+[　 ]", t):
                want_h2.append(t)
    for rec in MURA_EDITS:
        for midashi, _honbun in rec.get("ko", []):
            if midashi.startswith("（"):
                want_sub.append(midashi)
    for rec in (list(JUTEN_EDITS) + list(HP_EDITS) + list(HP3_EDITS)
                + list(RT_EDITS) + list(ME_EDITS) + list(SA_EDITS)
                + list(AN_EDITS) + list(KG_EDITS)
                + list(KO_EDITS) + list(KJ_EDITS)
                + list(IK_EDITS) + list(YS_EDITS)
                + list(KY_EDITS) + list(KN_EDITS)
                + list(CS_EDITS)):
        for blk in rec["nakami"]:
            if blk[0] in ("h2", "h3") and blk[1].startswith("（"):
                want_sub.append(blk[1])
            elif blk[0] in ("h2", "h3") and re.match(
                    r"^[０-９0-9]+[　 ]", blk[1]):
                want_h2.append(blk[1])
    styles = {p.text.strip(): p.style.name for p in doc.paragraphs
              if p.text.strip()}
    for t in want_h2:
        if styles.get(t) != "Heading 2":
            ng.append(f"節の見出しが Heading 2 でない: {t}（{styles.get(t)}）")
    for t in want_sub:
        if styles.get(t) != "小見出し":
            ng.append(f"項の見出しが 小見出し でない: {t}（{styles.get(t)}）")

    # ⑤c 村資料点検で足した表の行が実在すること
    keikaku = find_table(doc, ("計画名", "項目", "内容"))
    if len(keikaku.rows) != 9:
        ng.append(f"計画の表が9行でない（{len(keikaku.rows)}行）")
    nendo = find_table(doc, ("年度", "R6\n(2024)"))
    if len(nendo.rows) != 5:
        ng.append(f"年度表が5行でない（{len(nendo.rows)}行）")
    else:
        grid = sum(int(tc.tcPr.find(qn("w:gridSpan")).get(qn("w:val")))
                   if tc.tcPr is not None
                   and tc.tcPr.find(qn("w:gridSpan")) is not None else 1
                   for tc in nendo.rows[4]._tr.tc_lst)
        if grid != 10:
            ng.append(f"年度表に足した行の列数が10でない（{grid}）")

    # ⑥ 増えた量が妥当であること（既存を消していない）
    sp, st = src_dims[0], src_dims[1]
    if len(doc.paragraphs) <= sp:
        ng.append(f"段落が増えていない（{sp}→{len(doc.paragraphs)}）")
    if len(doc.tables) <= st:
        ng.append(f"表が増えていない（{st}→{len(doc.tables)}）")

    # ⑦ 他団体の名を出さない
    for bad in ("小野町", "金ヶ崎", "阿蘇", "札幌"):
        if bad in "\n".join(parts):
            ng.append(f"他団体名が含まれている: {bad}")

    # ⑧ 照会番号の表記が揺れていないこと。
    #    記号は全角（Ｍ・Ｓ）、数字は半角に揃える。
    #    半角のMとSが混ざると、照会文との突き合わせで取りこぼす。
    zenbun = "\n".join(parts)
    han = re.findall(r"(?<![A-Za-zＡ-Ｚａ-ｚ])[MS]-\d+", zenbun)
    if han:
        ng.append(f"照会番号の記号が半角になっている: {sorted(set(han))}")
    zen_suji = re.findall(r"[ＭＳ]-[０-９]+", zenbun)
    if zen_suji:
        ng.append(f"照会番号の数字が全角になっている: "
                  f"{sorted(set(zen_suji))}")

    if ng:
        print("自己点検 不合格:")
        for e in ng:
            print("   -", e)
        raise SystemExit(1)
    print(f"  自己点検: 移植{len(ISHOKU)}件の文章と表が出力に実在／"
          f"論点{len(MEMO_EDITS)}件・用語{len(YOUGO_ADD)}語・"
          f"村資料{len(MURA_EDITS)}件・"
          f"重点施策{len(JUTEN_EDITS)}件・"
          f"村HP{len(HP_EDITS)}件・村HP3分野{len(HP3_EDITS)}件・"
          f"RedTeam{len(RT_EDITS)}件・MECE{len(ME_EDITS)}件・"
          f"見込量{len(SA_EDITS)}件・アンケート{len(AN_EDITS)}件・"
          f"介護整合{len(KG_EDITS)}件・"
          f"広域{len(KO_EDITS)}件・情報公表{len(KJ_EDITS)}件・"
          f"意思決定{len(IK_EDITS)}件・国予算{len(YS_EDITS)}件・"
          f"共生型{len(KY_EDITS)}件・"
          f"高次脳{len(KN_EDITS)}件・"
          f"地域支援{len(CS_EDITS)}件・"
          f"書き換え{len(KAKIKAE) + len(KG_KAKIKAE) + len(KO_KAKIKAE)}件・"
          f"図{len(GRAPHS)}点が実在／"
          f"正本の記述{len(keep)}点が残存／見込量表{n_mikomi}表／"
          "表記の作法4点すべて合")


def main():
    if not os.path.exists(SRC):
        raise SystemExit(f"正本が見つかりません: {SRC}")
    os.makedirs(f"{REPO_ROOT}/output", exist_ok=True)
    doc = docx.Document(SRC)
    src_dims = (len(doc.paragraphs), len(doc.tables), count_images(doc))

    tail = apply_ishoku(doc)
    apply_memo(doc, tail)
    apply_mura(doc)
    apply_juten(doc)
    apply_juten(doc, HP_EDITS, label="村HP")
    apply_juten(doc, HP3_EDITS, label="村HP3分野")
    apply_juten(doc, RT_EDITS, label="RedTeam")
    apply_juten(doc, ME_EDITS, label="MECE")
    apply_juten(doc, SA_EDITS, label="見込量")
    apply_juten(doc, AN_EDITS, label="アンケート")
    apply_juten(doc, KG_EDITS, label="介護整合")
    apply_juten(doc, KO_EDITS, label="広域")
    apply_juten(doc, KJ_EDITS, label="情報公表")
    apply_juten(doc, IK_EDITS, label="意思決定")
    apply_juten(doc, YS_EDITS, label="国予算")
    apply_juten(doc, KY_EDITS, label="共生型")
    apply_juten(doc, KN_EDITS, label="高次脳")
    apply_juten(doc, CS_EDITS, label="地域支援")
    apply_kakikae(doc)
    apply_graphs(doc)
    apply_bangou(doc)
    apply_zuhyo_mokuji(doc)
    verify(doc, src_dims)
    doc.save(OUT_FILE)

    print(f"作成: {OUT_FILE}")
    print(f"  正本: 段落{src_dims[0]}・表{src_dims[1]}")
    print(f"  移植後: 段落{len(doc.paragraphs)}・表{len(doc.tables)}"
          f"（＋{len(doc.paragraphs) - src_dims[0]}段落・"
          f"＋{len(doc.tables) - src_dims[1]}表）")
    print("  入れたもの:")
    for c in changes:
        print(f"    - {c}")


if __name__ == "__main__":
    main()
