"""障がい計画　他メンバー作成の計画素案へ、当方の内容を変更履歴付きで反映する。

何をするか
  『他素案への追加候補の対応表』（build_ono_shogai_tsuika_taiou.py）の
  優先「高」9件（№11・13・14・16・19・22・23・24・25）を、
  他メンバー作成の計画素案（Word）に**変更履歴（挿入・削除）として**書き込む。
  他メンバーが Word の「変更履歴」で1件ずつ承諾・却下できるようにする。

やり方
  docx を展開し word/document.xml を lxml で直接編集する。
  挿入は <w:ins>、削除は <w:del>（文字は <w:delText>）で囲む。
  行の挿入・削除は <w:trPr> に <w:ins>/<w:del> を置く。
  他素案の書式（BIZ UDPゴシック・表の罫線と見出しの塗り）は、
  その場所の既存の段落・表から写して使う。

値の出所
  №13 の第7期の目標値と №16 の第7期の計画値は、**書き写さない。**
  前者は対応表の ZENKAI_MOKUHYO（原本の頁で照合済み）の値が
  前回計画の該当頁にあることを、後者は ono_shogai_zenkai が前回計画のPDFから
  読んだ値を用いる。№19 の文言は ono_setten（高齢者計画と同じもの）から作る。

入出力
  入力  03_計画素案/他メンバー素案/他メンバー素案_受領_20261020.docx
  出力  03_計画素案/他メンバー素案/
        ★小野町_障がい者計画等（素案）_当方反映_変更履歴付き_20261020.docx
"""

import copy
import pathlib
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
import zipfile

from lxml import etree

import build_ono_shogai_tsuika_taiou as T
import ono_setten as SET
import ono_shogai_zenkai as Z

ROOT = pathlib.Path(__file__).parent / "小野町_引継ぎ_整理済" / "03_計画素案" / "他メンバー素案"
SRC = ROOT / "他メンバー素案_受領_20261020.docx"
OUT = ROOT / "★小野町_障がい者計画等（素案）_当方反映_変更履歴付き_20261020.docx"

AUTHOR = "Ryo Wakayama"
DATE = "2026-10-20T09:00:00Z"
TBD = "【要確認】"

WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % WNS
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
NS = {"w": WNS}

_ID = [90000]
LOG = []          # (№, 何をしたか)


def _nid():
    _ID[0] += 1
    return str(_ID[0])


def _mark(tag):
    e = etree.Element(W + tag)
    e.set(W + "id", _nid())
    e.set(W + "author", AUTHOR)
    e.set(W + "date", DATE)
    return e


def txt(e):
    return "".join(e.itertext())


# ------------------------------------------------------------------ 文字
_ZEN = str.maketrans("0123456789", "０１２３４５６７８９")


def zen(s):
    """本文の1桁の数字を全角にする（他素案の本文の書き方。2桁以上は半角のまま）。"""
    return re.sub(r"(?<![0-9.])([0-9])(?![0-9.%])",
                  lambda m: m.group(1).translate(_ZEN), s)


def segs(s):
    """「**太字**」の印で区切り、(文字, 太字か) の並びにする。"""
    out = []
    for i, part in enumerate(re.split(r"\*\*", s)):
        if part:
            out.append((part, i % 2 == 1))
    return out


# ------------------------------------------------------------------ 部品
def _rpr_of(p):
    r = p.find(".//w:r/w:rPr", NS)
    if r is None:
        r = etree.Element(W + "rPr")
    r = copy.deepcopy(r)
    for t in ("b", "bCs", "color", "highlight"):
        for x in r.findall("w:" + t, NS):
            r.remove(x)
    return r


def _ins_run(text, rpr, bold=False):
    ins = _mark("ins")
    r = etree.SubElement(ins, W + "r")
    rp = copy.deepcopy(rpr)
    if bold:
        rp.insert(1 if rp.find("w:rFonts", NS) is not None else 0, etree.Element(W + "b"))
    r.append(rp)
    t = etree.SubElement(r, W + "t")
    t.text = text
    t.set(XML_SPACE, "preserve")
    return ins


def _mark_para_inserted(p):
    ppr = p.find("w:pPr", NS)
    if ppr is None:
        ppr = etree.Element(W + "pPr")
        p.insert(0, ppr)
    rpr = ppr.find("w:rPr", NS)
    if rpr is None:
        rpr = etree.SubElement(ppr, W + "rPr")
    rpr.insert(0, _mark("ins"))


def new_para(template, text, bold_all=False, keep_digits=False):
    """template の段落書式で、挿入扱いの段落を作る。"""
    p = etree.Element(W + "p")
    ppr = template.find("w:pPr", NS)
    if ppr is not None:
        ppr = copy.deepcopy(ppr)
        for x in ppr.findall("w:numPr", NS):
            ppr.remove(x)
        for x in ppr.findall("w:rPr", NS):
            for y in x.findall("w:ins", NS) + x.findall("w:del", NS):
                x.remove(y)
        p.append(ppr)
    _mark_para_inserted(p)
    rpr = _rpr_of(template)
    s = text if keep_digits else zen(text)
    for part, b in segs(s):
        p.append(_ins_run(part, rpr, bold=b or bold_all))
    return p


def insert_after(anchor, *els):
    for e in reversed(els):
        anchor.addnext(e)
    return els[-1]


def _merge_runs(p):
    """書式が同じで文字だけの run が隣り合っていれば、1つにまとめる。

    Word は校正の印などで1つの語句を複数の run に分けて持つ。
    置き換える語句が run をまたぐと見つからないため、まとめてから探す。
    """
    def key(r):
        rp = r.find("w:rPr", NS)
        return etree.tostring(rp) if rp is not None else b""
    for parent in {r.getparent() for r in p.iter(W + "r")}:
        prev = None
        for r in list(parent):
            ok = r.tag == W + "r" and all(c.tag in (W + "rPr", W + "t") for c in r)
            if ok and prev is not None and key(prev) == key(r):
                pt = prev.find("w:t", NS)
                rt = r.find("w:t", NS)
                if pt is not None and rt is not None:
                    pt.text = (pt.text or "") + (rt.text or "")
                    pt.set(XML_SPACE, "preserve")
                    parent.remove(r)
                    continue
            prev = r if ok else None


def replace_in_para(p, old, new):
    """段落の中の old を削除し new を挿入する（old は1つの run に収まること）。"""
    if not any(old in (t.text or "") for t in p.iter(W + "t")):
        _merge_runs(p)
    for r in p.findall(".//w:r", NS):
        t = r.find("w:t", NS)
        if t is None or t.text is None or old not in t.text:
            continue
        before, after = t.text.split(old, 1)
        rpr = r.find("w:rPr", NS)
        parent = r.getparent()
        idx = parent.index(r)
        pieces = []
        if before:
            r1 = copy.deepcopy(r)
            r1.find("w:t", NS).text = before
            r1.find("w:t", NS).set(XML_SPACE, "preserve")
            pieces.append(r1)
        d = _mark("del")
        rd = etree.SubElement(d, W + "r")
        if rpr is not None:
            rd.append(copy.deepcopy(rpr))
        dt = etree.SubElement(rd, W + "delText")
        dt.text = old
        dt.set(XML_SPACE, "preserve")
        pieces.append(d)
        # 【要確認】は黄色の網掛けで書かれている。挿入する値には網掛けを写さない
        rpr_i = copy.deepcopy(rpr) if rpr is not None else etree.Element(W + "rPr")
        for x in rpr_i.findall("w:highlight", NS):
            rpr_i.remove(x)
        for part, b in segs(new):
            pieces.append(_ins_run(part, rpr_i, b))
        if after:
            r2 = copy.deepcopy(r)
            r2.find("w:t", NS).text = after
            r2.find("w:t", NS).set(XML_SPACE, "preserve")
            pieces.append(r2)
        parent.remove(r)
        for k, x in enumerate(pieces):
            parent.insert(idx + k, x)
        return True
    raise ValueError("段落に「%s」が1つの run として見当たらない：%s" % (old, txt(p)[:40]))


def delete_row(tr):
    trpr = tr.find("w:trPr", NS)
    if trpr is None:
        trpr = etree.Element(W + "trPr")
        tr.insert(0, trpr)
    trpr.append(_mark("del"))
    for p in tr.findall(".//w:p", NS):
        ppr = p.find("w:pPr", NS)
        if ppr is None:
            ppr = etree.Element(W + "pPr")
            p.insert(0, ppr)
        rpr = ppr.find("w:rPr", NS)
        if rpr is None:
            rpr = etree.SubElement(ppr, W + "rPr")
        rpr.insert(0, _mark("del"))
        for r in p.findall("w:r", NS):
            d = _mark("del")
            r.addprevious(d)
            d.append(r)
            for t in r.findall("w:t", NS):
                t.tag = W + "delText"
            for x in r.findall("w:lastRenderedPageBreak", NS):
                r.remove(x)


def inserted_row_like(tr, texts, bold_first=False):
    """tr の書式を写した、挿入扱いの行を作る。"""
    nr = copy.deepcopy(tr)
    trpr = nr.find("w:trPr", NS)
    if trpr is None:
        trpr = etree.Element(W + "trPr")
        nr.insert(0, trpr)
    for x in trpr.findall("w:tblHeader", NS):
        trpr.remove(x)
    trpr.append(_mark("ins"))
    tcs = nr.findall("w:tc", NS)
    if len(tcs) != len(texts):
        raise ValueError("列の数が合わない（%d列に%d個）" % (len(tcs), len(texts)))
    for k, (tc, s) in enumerate(zip(tcs, texts)):
        ps = tc.findall("w:p", NS)
        p0 = ps[0]
        rpr = _rpr_of(p0)
        for p in ps[1:]:
            tc.remove(p)
        for r in list(p0):
            if r.tag != W + "pPr":
                p0.remove(r)
        _mark_para_inserted(p0)
        for part, b in segs(s):
            p0.append(_ins_run(part, rpr, b or (bold_first and k == 0)))
    return nr


def new_table(like_tbl, widths, header, rows):
    """like_tbl の罫線・見出しの塗り・余白を写して、挿入扱いの表を作る。"""
    tb = etree.Element(W + "tbl")
    tb.append(copy.deepcopy(like_tbl.find("w:tblPr", NS)))
    grid = etree.SubElement(tb, W + "tblGrid")
    for w_ in widths:
        g = etree.SubElement(grid, W + "gridCol")
        g.set(W + "w", str(w_))
    trs = like_tbl.findall("w:tr", NS)
    hdr_tr, body_tr = trs[0], trs[1]
    hdr_tc = hdr_tr.findall("w:tc", NS)[0]
    body_tc = body_tr.findall("w:tc", NS)[0]

    def make_row(src_tr, src_tc, cells, header):
        tr = etree.SubElement(tb, W + "tr")
        trpr = copy.deepcopy(src_tr.find("w:trPr", NS)) \
            if src_tr.find("w:trPr", NS) is not None else etree.Element(W + "trPr")
        if not header:
            for x in trpr.findall("w:tblHeader", NS):
                trpr.remove(x)
        for x in trpr.findall("w:ins", NS) + trpr.findall("w:del", NS):
            trpr.remove(x)
        trpr.append(_mark("ins"))
        tr.append(trpr)
        for w_, s in zip(widths, cells):
            tc = etree.SubElement(tr, W + "tc")
            tcpr = copy.deepcopy(src_tc.find("w:tcPr", NS))
            tcpr.find("w:tcW", NS).set(W + "w", str(w_))
            tc.append(tcpr)
            sp = src_tc.find("w:p", NS)
            p = etree.SubElement(tc, W + "p")
            ppr = copy.deepcopy(sp.find("w:pPr", NS))
            for x in ppr.iter(W + "ins", W + "del"):
                x.getparent().remove(x)
            if not header:
                for x in ppr.findall("w:jc", NS):
                    ppr.remove(x)
            p.append(ppr)
            _mark_para_inserted(p)
            rpr = _rpr_of(sp)
            for part, b in segs(zen(s)):
                p.append(_ins_run(part, rpr, b or header))
        return tr

    make_row(hdr_tr, hdr_tc, header, True)
    for row in rows:
        make_row(body_tr, body_tc, row, False)
    return tb


# ------------------------------------------------------------------ 探す
class Doc:
    def __init__(self, path):
        self.tree = etree.parse(str(path))
        self.body = self.tree.getroot().find("w:body", NS)

    def paras(self, start):
        return [p for p in self.body.iter(W + "p") if txt(p).startswith(start)]

    def para(self, start):
        hit = self.paras(start)
        if len(hit) != 1:
            raise ValueError("「%s」で始まる段落が%d個（1個であること）" % (start, len(hit)))
        return hit[0]

    def tables(self, start):
        return [t for t in self.body.iter(W + "tbl") if txt(t).startswith(start)]

    def table(self, start):
        hit = self.tables(start)
        if len(hit) != 1:
            raise ValueError("「%s」で始まる表が%d個（1個であること）" % (start, len(hit)))
        return hit[0]

    @staticmethod
    def next_table(el):
        e = el.getnext()
        while e is not None and e.tag != W + "tbl":
            e = e.getnext()
        if e is None:
            raise ValueError("後ろに表がない")
        return e


# ------------------------------------------------------------------ №13・14
# 他素案の表「第7期計画の成果目標の達成状況」の7行に入れる値。
# 数値は対応表の ZENKAI_MOKUHYO と同じ原本の頁にあることを genpon_hyou() で確かめる。
HYOU13 = [
    "移行31人（119.2％）・削減22人（84.6％）。基準は令和4年度末の入所者26人"
    "（移行が入所者を上回り、評価の基準にならない。本章1）",
    "町の成果目標は設定していない（県が設定）",
    "拠点等1か所・コーディネーター1人・検証年3回、"
    "強度行動障がいの支援ニーズの把握 有・支援体制の整備 有",
    "一般就労移行1人（うち就労移行支援1人・Ａ型1人・Ｂ型1人）、"
    "移行5割以上の就労移行支援事業所1か所、就労定着支援の利用者1人、"
    "定着率7割以上の事業所1か所",
    "児童発達支援センター1か所、保育所等訪問支援1か所、"
    "重症心身障がい児の児童発達支援1か所・放課後等デイサービス1か所、"
    "協議の場1か所・コーディネーター1人",
    "基幹相談支援センター 有、指導・助言158件、人材育成の支援104件、"
    "連携強化102回、事例検証57回、主任相談支援専門員1人、"
    "協議会の事例検討1回、専門部会1部会",
    "研修参加1人、分析・共有の体制 有・共有年1回",
]


def genpon_hyou():
    """HYOU13 の数値（○人・○件など）が、前回計画の該当頁にあるかを確かめる。"""
    ng = list(T.genpon_ng())
    for (g, _k, _v, pages, _w), s in zip(T.ZENKAI_MOKUHYO, HYOU13):
        if not pages:
            continue
        body = "".join(T._page(p) for p in pages)
        for tok in re.findall(r"\d+(?:\.\d+)?(?:人|か所|件|回|％|部会)",
                              unicodedata.normalize("NFKC", s)):
            tok = tok.replace("％", "%")
            if tok == "26人" and g == "施設入所者の地域生活移行":
                pass
            if tok not in body:
                ng.append("№13 %s：「%s」が原本 p.%s にない" % (g, tok, pages))
    return ng


def do13(d):
    tb = d.table("成果目標項目第７期の目標")
    trs = tb.findall("w:tr", NS)[1:]
    if len(trs) != len(HYOU13):
        raise ValueError("達成状況の表の行数が%dで、7ではない" % len(trs))
    for tr, s in zip(trs, HYOU13):
        tc = tr.findall("w:tc", NS)[2]
        replace_in_para(tc.find("w:p", NS), "【要確認】", s)
    p = d.para("※第７期の目標は「おのまち障がい者計画")
    replace_in_para(p, "により記載します。", "第３部第５章によります。")
    replace_in_para(p, "【要確認：第７期の目標値】", "")
    LOG.append(("13", "第3部第5章冒頭の達成状況の表、7行の【要確認】を前回計画の値に置換。表の注の【要確認】を削除"))


def do14(d):
    a = d.para("※地域生活移行者数とは、入所施設の入所者が施設を退所し")
    insert_after(a, new_para(a,
        "※第７期計画は、令和４年度末の入所者26人に対し、地域生活移行者数31人（119.2％）、"
        "削減見込者数22人（84.6％）を目標としていました。移行者数が基準の入所者数を上回っており、"
        "目標として成り立たない値でした。また、22人を削減すると入所者は４人となる一方、"
        "同じ計画の施設入所支援の見込量は15人であり、目標と見込量が整合していませんでした。"
        "このため、第７期の達成状況は「未達成」とはせず、令和７年度末の入所者19人を"
        "令和４年度末の26人と比べて評価します（７人減、26.9％減）。"
        "第８期の目標は、令和７年度末の入所者19人を基準に、国の基本指針に即して設定しています。"))
    LOG.append(("14", "第3部第5章1の表の注の後に、前回目標が成り立たない値であったことの注を挿入"))


# ------------------------------------------------------------------ №11
def do11(d):
    a = d.para("障害者総合支援法に基づく協議会である「小野町自立支援協議会」において")
    insert_after(a, new_para(a,
        "また、社会福祉法等の一部を改正する法律（令和8年法律第51号）により、令和9年4月1日から、"
        "協議会は「支援協議会」に名称が改められます（障害者総合支援法第89条の3）。"
        "同条第5項により、支援協議会は、介護保険法第115条の48第1項の支援会議、"
        "社会福祉法第106条の3の2第1項の支援会議、"
        "住宅確保要配慮者に対する賃貸住宅の供給の促進に関する法律第81条第1項の支援協議会"
        "（居住支援協議会）、生活困窮者自立支援法第9条第1項の支援会議などが置かれているときは、"
        "これらと相互に連携を図るよう努めることとされます。"
        "本町に置かれている会議体を確かめたうえで、65歳到達時の移行や共生型サービスなど"
        "高齢者分野にまたがる事柄を中心に、連携の場と頻度を定めます（第９章４）。" + TBD))
    LOG.append(("11", "第3部第3章4の段落の後に、支援協議会への名称変更と5つの会議体との相互連携を挿入"))


# ------------------------------------------------------------------ №22・24
def do22_24(d):
    a = d.para("そのうえで、必要な利用者数を次のとおり組み立て")
    tb = d.next_table(a)
    p22 = new_para(a,
        "なお、生活介護、就労継続支援Ｂ型及び施設入所支援の見込量は、国の基本指針により、"
        "継続入所者（平成24年の制度改正の前から障がい児施設に入所し、その施設が"
        "障害者支援施設の指定を受けた後も引き続き入所している18歳以上の方）の数を除いて"
        "設定するものとされています。本町の継続入所者の有無を確かめ、"
        "該当する場合はその数を除いて設定します。【要確認：継続入所者の有無】")
    p24 = new_para(a,
        "また、国の基本指針は、別表第五の一の項の要件（全部過疎市町村ではないこと、かつ、"
        "人口に占めるサービスの利用者の割合が上位25％の範囲内であること）に該当する市町村について、"
        "生活介護、就労継続支援Ａ型・Ｂ型、共同生活援助、児童発達支援及び放課後等デイサービスの"
        "見込量の算定方法を定めています（同表の二の項）。全部過疎市町村は対象外です。"
        "本町は過疎地域の持続的発展の支援に関する特別措置法第2条第1項の"
        "過疎地域市町村（全部過疎）であるため、別表第五の適用はありません。")
    insert_after(tb, p22, p24)
    LOG.append(("22", "第3部第6章冒頭の表の後に、継続入所者の数を除いて設定する旨を挿入"))
    LOG.append(("24", "同じ位置に、別表第五（地域差の是正）の適用判定を挿入"))


# ------------------------------------------------------------------ №23
def do23(d):
    hits = [p for p in d.paras("第７期の利用実績は、児童発達支援は月５～７人")]
    if len(hits) != 1:
        raise ValueError("第7章1の方策の段落が%d個" % len(hits))
    a = hits[0]
    insert_after(a, new_para(a,
        "見込量は、国の基本指針により、保育所や認定こども園等における障がい児の受入れの体制の"
        "整備状況を踏まえて設定するものとされています。保育所等での受入れが広がれば"
        "児童発達支援等の量は減り、受入れが難しくなれば増えるため、町の保育担当部局と"
        "受入れの状況を毎年度確かめ、見込量に反映します。"
        "【要確認：保育所等における障がい児の受入れの状況】"))
    LOG.append(("23", "第3部第7章1の方策の段落の後に、保育所等の受入れ体制を踏まえる旨を挿入"))


# ------------------------------------------------------------------ №25
def do25(d):
    tb = d.table("項目数値備考令和７年度の公表率・更新率")
    hit = [tr for tr in tb.findall("w:tr", NS) if txt(tr).startswith("【目標値】令和11年度の公表率")]
    if len(hit) != 1:
        raise ValueError("公表率の目標値の行が%d行" % len(hit))
    tc0 = hit[0].findall("w:tc", NS)[0]
    replace_in_para(tc0.find("w:p", NS), "【目標値】令和11年度の公表率・更新率",
                    "（参考）県の目標（管内事業所の公表率・更新率）")
    note_t = d.para("※第７期の目標は「おのまち障がい者計画")
    insert_after(tb, new_para(note_t,
        "※障がい福祉サービス等情報公表制度の公表率・更新率100％は、国の基本指針において、"
        "事業所の指定権者である都道府県・政令市・中核市の目標とされています。"
        "本町は指定権者でないため成果目標とはせず、県の目標を参考として掲げ、"
        "町内事業所への周知・働きかけを行います。"))
    LOG.append(("25", "第3部第5章8の表の【目標値】を（参考）県の目標に置換し、理由の注を挿入"))


# ------------------------------------------------------------------ №16（17を含む）
# 他素案の行の名前 → (前回計画のサービス名, 量の単位 or None)
KEIKAKU = {
    "居宅介護（時間／人）": ("居宅介護", "時間"),
    "重度訪問介護": ("重度訪問介護", "時間"),
    "同行援護": ("同行援護", "時間"),
    "行動援護": ("行動援護", "時間"),
    "重度障がい者等包括支援": ("重度障がい者等包括支援", "時間"),
    "生活介護": ("生活介護", "人日"),
    "自立訓練（機能訓練）": ("自立訓練（機能訓練）", "人日"),
    "自立訓練（生活訓練）": ("自立訓練（生活訓練）", "人日"),
    "就労移行支援": ("就労移行支援", "人日"),
    "就労継続支援（Ａ型）": ("就労継続支援（A型）", "人日"),
    "就労継続支援（Ｂ型）": ("就労継続支援（B型）", "人日"),
    "就労定着支援": ("就労定着支援", None),
    "療養介護": ("療養介護", None),
    "短期入所（福祉型）": ("短期入所（福祉型）", "人日"),
    "短期入所（医療型）": ("短期入所（医療型）", "人日"),
    "自立生活援助": ("自立生活援助", None),
    "共同生活援助": ("共同生活援助", None),
    "施設入所支援": ("施設入所支援", None),
    "計画相談支援": ("計画相談支援", None),
    "地域移行支援": ("地域移行支援", None),
    "地域定着支援": ("地域定着支援", None),
    "児童発達支援": ("児童発達支援", "人日"),
    "放課後等デイサービス": ("放課後等デイサービス", "人日"),
    "保育所等訪問支援": ("保育所等訪問支援", "人日"),
    "居宅訪問型児童発達支援": ("居宅訪問型児童発達支援", "人日"),
    "障がい児相談支援": ("障がい児相談支援", None),
}
SHINSETSU = {"就労選択支援"}      # 第7期には無かったサービス


def keikakuchi(nm):
    svc, tani = KEIKAKU[nm]
    d = Z.SERVICE[svc]
    hito = d["人"]["7期"]
    if tani is None:
        return ["%d人" % h for h in hito]
    ryo = d[tani]["7期"]
    return ["%d%s／%d人" % (q, tani, h) for q, h in zip(ryo, hito)]


def do16(d):
    Z.load()
    els = list(d.body)
    sho8 = [i for i, e in enumerate(els)
            if e.tag == W + "p" and txt(e).startswith("第８章") and "地域生活支援事業" in txt(e)]
    if not sho8:
        raise ValueError("第8章の見出しが見当たらない")
    end = sho8[0]
    n_row = 0
    n_tbl = 0
    jidou_tb = None
    for i, e in enumerate(els[:end]):
        if not (e.tag == W + "p" and txt(e).startswith("■第７期の実績")):
            continue
        tb = Doc.next_table(e)
        n_tbl += 1
        for tr in tb.findall("w:tr", NS)[1:]:
            nm = txt(tr.findall("w:tc", NS)[0]).strip()
            if nm in SHINSETSU:
                cells = ["　第７期の計画値", "―", "―", "―"]
            elif nm in KEIKAKU:
                cells = ["　第７期の計画値"] + keikakuchi(nm)
            else:
                continue
            tr.addnext(inserted_row_like(tr, cells))
            n_row += 1
        if txt(tb).startswith("第７期令和６年度令和７年度令和８年度児童発達支援"):
            jidou_tb = tb
    # 章の冒頭に、計画値を並べたことの説明
    a = d.para("本章では、障がい福祉サービスの種類ごとに、第７期の実績")
    insert_after(a, new_para(a,
        "各サービスの第７期の実績の表には、第７期計画（令和６年３月）で定めた見込量を"
        "「第７期の計画値」として併せて示します。計画値と実績の差は、"
        "年次の点検評価（第１部第２章２）の材料とします。"))
    # 児童発達支援・放課後等デイサービスの計画値の注（対応表№17）
    if jidou_tb is None:
        raise ValueError("障がい児通所支援の実績の表が見当たらない")
    note_t = d.para("※第７期の目標は「おのまち障がい者計画")
    insert_after(jidou_tb, new_para(note_t,
        "※児童発達支援・放課後等デイサービスの第７期の計画値は、第７期計画の表のとおり記載しています。"
        "人数と量（人日）が入れ替わっているとみられ（第６期の実績は児童発達支援６人・24～28人日、"
        "放課後等デイサービス31人・93人日）、計画値と実績の比較には用いません。"))
    LOG.append(("16", "第3部第6章・第7章の実績の表%d点に「第７期の計画値」の行を%d行挿入。章の冒頭に説明、"
                "障がい児通所支援の表に注（№17）" % (n_tbl, n_row)))
    return n_tbl, n_row


# ------------------------------------------------------------------ №19
def do19(d):
    a = d.para("小野町高齢者保健福祉計画・介護保険事業計画（第10期）は、本計画と同じ")
    replace_in_para(a, "両計画が同じ人を重ねて数えることのないよう、次の事項を確かめたうえで"
                       "見込量を定めます。",
                    "両計画の接点を次のとおり定めます。本節は、高齢者保健福祉計画・"
                    "介護保険事業計画に同じ内容を掲げています。")
    old = d.next_table(a)
    paras = [new_para(a, s) for s in SET.LEAD]
    t1 = new_table(old, [450, 1800, 3500, 3250],
                   ["#", "場面", "制度の内容", "本町での扱い"],
                   [[n, k, s, t] for n, k, s, t in SET.SETTEN])
    p_j = new_para(a, SET.JINKO)
    note_t = d.para("※第７期の目標は「おのまち障がい者計画")
    p_n = new_para(note_t, SET.JOBUN)
    p_k = new_para(a, "**上の判定を行うために、次のことを町に確かめます。**")
    kiku = ["5（保険料の分母）", "2・4・7・9（両計画の見込量）",
            "1（両計画の見込量）", "6（事業所の数え方）", "8（引継ぎ）",
            "10（判定を行う場）", "10（判定を行う場）"]
    if len(kiku) != len(SET.KAKUNIN):
        raise ValueError("確かめること%d件・効くもの%d件" % (len(SET.KAKUNIN), len(kiku)))
    t2 = new_table(old, [450, 6300, 2250], ["#", "確かめること", "効くもの"],
                   [[str(i + 1), x, y] for i, (x, y) in enumerate(zip(SET.KAKUNIN, kiku))])
    p_e = new_para(a, "いずれも" + TBD + "です。"
                   "**2の一覧が得られれば、2・4・7・9がまとめて片付きます。**")
    # 旧表の行は、新しい表を写し取った後で削除扱いにする（先に消すと、
    # 削除の印まで新しい表に写ってしまう）
    for tr in old.findall("w:tr", NS):
        delete_row(tr)
    insert_after(old, *paras, t1, p_j, p_n, p_k, t2, p_e)
    LOG.append(("19", "第3部第9章4の表（5事項）を削除し、高齢者計画と同じ接点の節"
                "（10場面・確かめること7件）を挿入。導入の文を改めた"))


# ------------------------------------------------------------------ 本体
def build():
    if not SRC.exists():
        raise FileNotFoundError(SRC)
    ng = genpon_hyou()
    if ng:
        raise ValueError("前回計画の原本と合わない：" + "／".join(ng))
    work = pathlib.Path(tempfile.mkdtemp())
    with zipfile.ZipFile(SRC) as z:
        z.extractall(work)
    for p in work.rglob("*"):
        if p.is_symlink():
            p.unlink()
    d = Doc(work / "word" / "document.xml")
    do13(d)
    do14(d)
    do11(d)
    do22_24(d)
    do23(d)
    do25(d)
    n_tbl, n_row = do16(d)
    do19(d)
    d.tree.write(str(work / "word" / "document.xml"), xml_declaration=True,
                 encoding="UTF-8", standalone=True)
    if OUT.exists():
        OUT.unlink()
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        names = [n for n in zipfile.ZipFile(SRC).namelist()]
        for n in names:
            z.write(work / n, n)
    shutil.rmtree(work)
    return n_tbl, n_row


# ------------------------------------------------------------------ 自己点検
def jiko():
    import docx
    ng = []
    d = docx.Document(str(OUT))
    root = d.element.body
    ins = root.findall(".//" + W + "ins")
    dele = root.findall(".//" + W + "del")
    authors = {e.get(W + "author") for e in ins + dele}
    if authors != {AUTHOR}:
        ng.append("変更履歴の作成者が %s" % authors)
    # 1 9件すべて反映
    if sorted({x[0] for x in LOG}, key=int) != ["11", "13", "14", "16", "19", "22", "23", "24", "25"]:
        ng.append("反映した件が9件そろわない：%s" % [x[0] for x in LOG])
    # 2 変更を承諾した姿の本文に、入れるべき語がある／消すべき語がない
    acc = accepted_text()
    for w in ["支援協議会」に名称", "第89条の３", "目標として成り立たない値", "継続入所者",
              "別表第五", "保育所や認定こども園等", "（参考）県の目標", "第７期の計画値",
              "同じ方の同じ利用が両方の見込量に入る", "計画相談支援から居宅介護支援への引継ぎ"]:
        if w not in acc:
            ng.append("承諾後の本文に「%s」がない" % w)
    for w in ["【要確認：第７期の目標値】", "同じ人を重ねて数えることのないよう",
              "【目標値】令和11年度の公表率"]:
        if w in acc:
            ng.append("承諾後の本文に「%s」が残っている" % w)
    # 3 接点の節が高齢者計画と同じ文言（全角・半角の違いを除く）
    for i, s in enumerate(SET.LEAD):
        if unicodedata.normalize("NFKC", s.replace("**", "")) not in \
                unicodedata.normalize("NFKC", acc):
            ng.append("接点の前書き%dが承諾後の本文にない" % (i + 1))
    # 4 他団体の名称が出ない
    for w in ("北塩原", "福島町", "北海道", "大雪"):
        if w in acc:
            ng.append("他団体の名称が出ている：%s" % w)
    return ng


def accepted_text():
    """変更を承諾した姿の本文（<w:del> を除き <w:ins> を残す）。"""
    with zipfile.ZipFile(OUT) as z:
        root = etree.fromstring(z.read("word/document.xml"))
    for e in list(root.iter(W + "del")):
        if e.getparent().tag not in (W + "rPr", W + "trPr"):
            e.getparent().remove(e)
    return "".join(t.text or "" for t in root.iter(W + "t"))


if __name__ == "__main__":
    n_tbl, n_row = build()
    print("■ 他素案への反映（変更履歴付き）")
    print("   %s" % OUT.relative_to(pathlib.Path(__file__).parent))
    for no, s in LOG:
        print("   №%-3s %s" % (no, s))
    ng = jiko()
    print()
    print("■ 自己点検 4項目")
    if ng:
        for x in ng:
            print("   **%s**" % x)
        sys.exit(1)
    print("   すべて適合")
