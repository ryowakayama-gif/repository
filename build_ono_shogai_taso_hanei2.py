"""障がい計画　他メンバー作成の計画素案へ、図表の点検（令和8年10月23日）の結果を変更履歴付きで反映する（第2弾）。

何をするか（「仕様書の網羅性確認と図表の点検」の「今すぐ進められる作業」1〜5）
  1・2　第2部の調査の出典を令和8年度調査に差し替える。
        令和5年度調査（前回調査）の表は削除し、令和8年度調査の図（ono_shogai_zu）と本文に置き換える。
        調査のなかった章（第4章・第7章1・第8章・第10章）にも図を置く（前回計画と同じく各章の「現状と課題」に置く）。
  3　　図表の番号・題名・資料を高齢者計画の規則にそろえる。
        【図N】【表N】の通し番号、すべての図表に固有の題名、実績・見込量の表に「資料：」。
  4　　「障害」「障がい」の書き分けを前回計画の書き方にそろえる（ono_shogai_hyoki）。変更履歴で示す。
        数字の全角・半角は半角に、「%」は「％」にそろえる。**件数が数千に及ぶため変更履歴にせず直接直す**
        （文字の形だけの変更。自己点検で全角数字が残っていないことを確かめる）。
  5　　当方素案（第5版）にしかない図表を加える。
        手帳所持者数の見通し、障がい別の状況、前回調査との比較、サービス区分ごとの算定方法、1人あたり日数。

入出力
  入力  他メンバー素案_受領_20261020.docx に、第1弾（build_ono_shogai_taso_hanei.py）を当てたもの
  出力  03_計画素案/他メンバー素案/★小野町_障がい者計画等（素案）_当方反映_変更履歴付き_20261023.docx
        （第1弾の変更履歴も含む。第1弾の単独の版は本版に含まれるため残さない）

値の出所
  調査の割合は、すべて集計表（ono_shogai_zu.load）から計算する。書き写さない。
  前回調査の割合は、前回調査の報告書（令和5年度）の本文にあることを自己点検で確かめる。
他団体の資料
  他団体の案件の計画素案は図表の数と構成の対比にだけ用いた。名称・数値・文言は入れない（自己点検）。
"""

import copy
import pathlib
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import unicodedata
import zipfile

from lxml import etree

import build_ono_shogai_taso_hanei as H
import ono_setten as SET
import ono_shogai_hyoki as HY
import ono_shogai_zu as Z

HERE = pathlib.Path(__file__).parent
ROOT = HERE / "小野町_引継ぎ_整理済"
DIR = ROOT / "03_計画素案" / "他メンバー素案"
OUT = DIR / "★小野町_障がい者計画等（素案）_当方反映_変更履歴付き_20261023.docx"
TOHO = (ROOT / "03_計画素案" /
        "おのまち障がい者計画_第4期障がい児福祉計画_第8期障がい福祉計画_素案_第5版_20261023.docx")
ZENKAI_CHOSA = (ROOT / "23_障がいアンケート分析" / "原本_前回調査_受領_20261007" /
                "前回_障がい福祉アンケート調査結果報告書_令和5年度.pdf")
FIG_NISSU = HERE / "output" / "ono_fig" / "f62_shogai_nissu.png"

DATE = "2026-10-23T09:00:00Z"
W = H.W
NS = H.NS
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"
TBD = "【要確認】"
SHIRYO_R8 = "資料：令和8年度 障がい福祉に関するアンケート調査"
LOG = []           # (作業, 内容)
ZENKAI_ATAI = []   # 本文に引いた前回調査の割合（自己点検で前回調査の報告書と照合）
INSERTED_CAP = set()   # 当方が挿入した題名の段落（番号の付け方を変える）
FIG_PARAS = []         # 当方が挿入した図の段落


def tx(e):
    return "".join(t.text or "" for t in e.iter(W + "t"))


# ------------------------------------------------------------------ 調査の値
SHA, JI = "集計_障がい者", "集計_障がい児"


def P(sheet, q, *ss, koumoku=None):
    d = Z.load(sheet)[q]
    c = sum(cc for k, s, cc in d["rows"] if s in ss and (koumoku is None or k == koumoku))
    if not any(s in ss for _k, s, _c in d["rows"]):
        raise KeyError((sheet, q, ss))
    return "%.1f％" % (c / d["n"] * 100)


def PK(sheet, q, *ss, koumoku=None):
    """割合と件数（障がい児調査は回答が25件のため件数を添える）。"""
    d = Z.load(sheet)[q]
    c = sum(cc for k, s, cc in d["rows"] if s in ss and (koumoku is None or k == koumoku))
    return "%.1f％（%d件）" % (c / d["n"] * 100, c)


def PB(q, b, *ss):
    d = Z.load_betsu()[q]
    c = sum(cc[b] for _k, s, cc in d["rows"] if s in ss)
    return "%.1f％" % (c / d["n"][b] * 100)


def PG(q, s, kiso):
    d = Z.load(SHA)[q]
    c = sum(cc for _k, ss, cc in d["rows"] if ss == s)
    return "%.1f％" % (c / kiso * 100)


def ZK(s):
    """前回調査の割合を本文に引く（自己点検で前回調査の報告書と照合する）。"""
    ZENKAI_ATAI.append(s)
    return s


# ------------------------------------------------------------------ 部品
def mark(tag):
    e = etree.Element(W + tag)
    e.set(W + "id", H._nid())
    e.set(W + "author", H.AUTHOR)
    e.set(W + "date", DATE)
    return e


def in_ins(el):
    a = el.getparent()
    while a is not None:
        if a.tag == W + "ins":
            return True
        a = a.getparent()
    return False


def in_del(el):
    a = el.getparent()
    while a is not None:
        if a.tag == W + "del" and a.getparent().tag not in (W + "rPr", W + "trPr"):
            return True
        a = a.getparent()
    return False


def del_para(p):
    """段落の文字と段落記号を、削除の変更履歴にする。"""
    ppr = p.find("w:pPr", NS)
    if ppr is None:
        ppr = etree.Element(W + "pPr")
        p.insert(0, ppr)
    rpr = ppr.find("w:rPr", NS)
    if rpr is None:
        rpr = etree.SubElement(ppr, W + "rPr")
    rpr.insert(0, mark("del"))
    for r in list(p.iter(W + "r")):
        if in_del(r):
            continue
        if in_ins(r):
            raise ValueError("挿入済みの文字を削除しようとした：%s" % tx(p)[:30])
        d = mark("del")
        r.addprevious(d)
        d.append(r)
        for t in r.findall("w:t", NS):
            t.tag = W + "delText"


def del_table(tb):
    for tr in tb.findall("w:tr", NS):
        H.delete_row(tr)


def split_run(r, a, b, new):
    """run の文字の a〜b を削除し new を挿入する（変更履歴）。"""
    t = r.find("w:t", NS)
    s = t.text
    parent = r.getparent()
    idx = parent.index(r)
    rpr = r.find("w:rPr", NS)
    pieces = []
    if s[:a]:
        r1 = copy.deepcopy(r)
        r1.find("w:t", NS).text = s[:a]
        r1.find("w:t", NS).set(H.XML_SPACE, "preserve")
        pieces.append(r1)
    d = mark("del")
    rd = etree.SubElement(d, W + "r")
    if rpr is not None:
        rd.append(copy.deepcopy(rpr))
    dt = etree.SubElement(rd, W + "delText")
    dt.text = s[a:b]
    dt.set(H.XML_SPACE, "preserve")
    pieces.append(d)
    ins = mark("ins")
    ri = etree.SubElement(ins, W + "r")
    if rpr is not None:
        rp = copy.deepcopy(rpr)
        for x in rp.findall("w:highlight", NS):
            rp.remove(x)
        ri.append(rp)
    ti = etree.SubElement(ri, W + "t")
    ti.text = new
    ti.set(H.XML_SPACE, "preserve")
    pieces.append(ins)
    if s[b:]:
        r2 = copy.deepcopy(r)
        r2.find("w:t", NS).text = s[b:]
        r2.find("w:t", NS).set(H.XML_SPACE, "preserve")
        pieces.append(r2)
    parent.remove(r)
    for k, x in enumerate(pieces):
        parent.insert(idx + k, x)
    return pieces[-1] if s[b:] else None


def rep(p, old, new):
    """段落の中の old を new に（変更履歴）。old は段落内で1か所であること。"""
    full = tx(p)
    if full.count(old) != 1:
        raise ValueError("「%s」が段落に%d個（1個であること）：%s" % (old[:30], full.count(old), full[:40]))
    H.replace_in_para(p, old, new)


def rep_all(p, new):
    """段落の文字をすべて削除し new を挿入する（表のセルの書き換えに使う）。"""
    tmpl = None
    for r in p.iter(W + "r"):
        if r.find("w:t", NS) is not None:
            tmpl = r
            break
    rpr = H._rpr_of(p)
    for r in list(p.iter(W + "r")):
        if in_del(r) or r.find("w:t", NS) is None:
            continue
        d = mark("del")
        r.addprevious(d)
        d.append(r)
        for t in r.findall("w:t", NS):
            t.tag = W + "delText"
    for part, b in H.segs(new):
        p.append(H._ins_run(part, rpr, b))
    return tmpl


def para(template, text, center=False):
    p = H.new_para(template, HY.hyoki(text), keep_digits=True)
    if center:
        ppr = p.find("w:pPr", NS)
        for x in ppr.findall("w:ind", NS) + ppr.findall("w:jc", NS):
            ppr.remove(x)
        jc = etree.SubElement(ppr, W + "jc")
        jc.set(W + "val", "center")
    return p


def cap(template, title):
    p = para(template, "■" + title)
    INSERTED_CAP.add(p)
    return p


# ------------------------------------------------------------------ 画像
class Media:
    def __init__(self, work):
        self.work = work
        self.rels_p = work / "word" / "_rels" / "document.xml.rels"
        self.rels = etree.parse(str(self.rels_p))
        self.n = 0
        ct = work / "[Content_Types].xml"
        t = etree.parse(str(ct))
        exts = {e.get("Extension") for e in t.getroot()}
        if "png" not in exts:
            e = etree.SubElement(t.getroot(), "{http://schemas.openxmlformats.org/package/2006/content-types}Default")
            e.set("Extension", "png")
            e.set("ContentType", "image/png")
            t.write(str(ct), xml_declaration=True, encoding="UTF-8", standalone=True)

    def add(self, png):
        self.n += 1
        name = "zu_r8_%02d.png" % self.n
        shutil.copy(png, self.work / "word" / "media" / name)
        rid = "rIdZu%02d" % self.n
        e = etree.SubElement(self.rels.getroot(),
                             "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship")
        e.set("Id", rid)
        e.set("Type", R_NS + "/image")
        e.set("Target", "media/" + name)
        return rid, name

    def save(self):
        self.rels.write(str(self.rels_p), xml_declaration=True, encoding="UTF-8", standalone=True)


def png_size(p):
    with open(p, "rb") as f:
        f.read(16)
        w, h = struct.unpack(">II", f.read(8))
    return w, h


def img_para(template, media, png, width_cm=15.0):
    rid, name = media.add(png)
    w, h = png_size(png)
    cx = int(width_cm * 360000)
    cy = int(cx * h / w)
    p = para(template, "", center=True)
    for r in list(p):
        if r.tag == W + "ins":
            p.remove(r)
    ins = mark("ins")
    r = etree.SubElement(ins, W + "r")
    dr = etree.SubElement(r, W + "drawing")
    xml = f'''<wp:inline xmlns:wp="{WP}" xmlns:a="{A_NS}" xmlns:pic="{PIC}" xmlns:r="{R_NS}" distT="0" distB="0" distL="0" distR="0">
<wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="{8000 + media.n}" name="{name}"/>
<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
<a:graphic><a:graphicData uri="{PIC}"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="{name}"/><pic:cNvPicPr/></pic:nvPicPr>
<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
</pic:pic></a:graphicData></a:graphic></wp:inline>'''
    dr.append(etree.fromstring(xml))
    p.append(ins)
    FIG_PARAS.append(p)
    return p


# ------------------------------------------------------------------ 文書
class D(H.Doc):
    def eq(self, text):
        hit = [p for p in self.body.iter(W + "p") if tx(p).strip() == text]
        if len(hit) != 1:
            raise ValueError("「%s」の段落が%d個（1個であること）" % (text[:30], len(hit)))
        return hit[0]

    def start(self, s):
        hit = [p for p in self.body.iter(W + "p") if tx(p).startswith(s) and not in_del(p)]
        if len(hit) != 1:
            raise ValueError("「%s」で始まる段落が%d個（1個であること）" % (s[:30], len(hit)))
        return hit[0]

    def cont(self, s):
        hit = [p for p in self.body.iter(W + "p") if s in tx(p)]
        if len(hit) != 1:
            raise ValueError("「%s」を含む段落が%d個（1個であること）" % (s[:30], len(hit)))
        return hit[0]


def nxt_until(el, stop):
    """el の後ろから、stop(e) が真になる要素までを並べる（stop の要素を含む）。"""
    out = []
    e = el.getnext()
    while e is not None:
        out.append(e)
        if stop(e):
            return out
        e = e.getnext()
    raise ValueError("終わりの要素が見つからない")


# ------------------------------------------------------------------ 作業1・2　調査の出典を令和8年度に
OLD_SRC = "資料：小野町障がい福祉に関するアンケート調査（令和５年10月実施）"


def block(d, media, cap_start, zus, tmpl):
    """前回調査の題名〜資料の段落までを削除し、令和8年度調査の図に置き換える。"""
    c = d.start(cap_start)
    seq = [c] + nxt_until(c, lambda e: e.tag == W + "p" and tx(e).strip() == OLD_SRC)
    n_t = 0
    for e in seq:
        if e.tag == W + "tbl":
            del_table(e)
            n_t += 1
        elif e.tag == W + "p" and tx(e).strip():
            del_para(e)
    anchor = seq[-1]
    els = []
    for k in zus:
        sho, dai, png = ZU[k]
        els += [cap(tmpl["cap"], dai), img_para(tmpl["img"], media, png)]
    els.append(para(tmpl["src"], SHIRYO_R8))
    H.insert_after(anchor, *els)
    LOG.append(("1・2", "%s：前回調査の表%d件を削除し、令和8年度調査の図%d点に置換" % (cap_start[1:20], n_t, len(zus))))


def add_fig(d, media, heading, text, zus, tmpl):
    """節の「現状と課題」の末尾（〈 施策の方向 〉の前）に、本文と図を加える。"""
    h = d.eq(heading)
    e = h.getnext()
    while not (e.tag == W + "p" and tx(e).strip() == "〈 施策の方向 〉"):
        e = e.getnext()
    els = [para(tmpl["body"], text)]
    for k in zus:
        sho, dai, png = ZU[k]
        els += [cap(tmpl["cap"], dai), img_para(tmpl["img"], media, png)]
    els.append(para(tmpl["src"], SHIRYO_R8))
    for x in els:
        e.addprevious(x)
    LOG.append(("1・2", "%s：令和8年度調査の図%d点と本文を追加" % (heading[:20], len(zus))))


def sakuin(d, media, tmpl):
    n37 = P(SHA, "問37", "名前も内容も知らない")
    j34 = P(JI, "問34", "名前も内容も知らない")
    kiso = Z.gaito(SHA, "問35", ("ある", "少しある"))
    # 第1部第3章 策定体制・調査の概要
    p = d.cont("前計画の策定にあわせて令和５年10月に実施した")
    rep(p, "前計画の策定にあわせて令和５年10月に実施した「障がい福祉に関するアンケート調査」の結果を基礎資料として活用するとともに",
        "令和8年8月から9月にかけて実施した「障がい福祉に関するアンケート調査」の結果を基礎資料として活用し"
        "（前計画の策定にあわせて令和5年10月に実施した調査と比べています）、")
    des = __import__("ono_shogai_chosa").report_design()
    tb = d.tables("項目内容調査の種類")[0]
    for tr in tb.findall("w:tr", NS):
        tcs = tr.findall("w:tc", NS)
        k = tx(tcs[0]).strip()
        newv = {
            "調査対象": "障がい者調査：%s　障がい児調査：%s" % des["調査対象"],
            "抽出方法": "町が対象者を抽出" + TBD,
            "調査方法": des["調査方法"][0],
            "調査時期": des["調査期間"][0],
            "回収結果": "障がい者：配布数%s／有効回答数%s（有効回答率%.1f％）　障がい児：配布数%s／有効回答数%s（有効回答率%.1f％）" % (
                unicodedata.normalize("NFKC", des["配布数"][0]), des["有効回答数"][0],
                float(des["有効回答率"][0].rstrip("%")),
                unicodedata.normalize("NFKC", des["配布数"][1]), des["有効回答数"][1],
                float(des["有効回答率"][1].rstrip("%"))),
        }.get(k)
        if newv:
            rep_all(tcs[1].find("w:p", NS), newv)
    LOG.append(("1・2", "第1部第1章4 策定体制：調査の時期・対象・回収結果を令和8年度調査に差し替え"))

    # 第1部第3章3 障がい児の状況
    p = d.cont("前回のアンケート調査では、障がい児の63.6％が")
    rep(p, "前回のアンケート調査では、障がい児の63.6％が放課後等デイサービスの利用意向を示しています。",
        "令和8年度の調査では、障がいのある子どもの放課後等デイサービスの利用意向は%s で、前回調査（%s）から大きく下がっています。"
        "今後の利用意向は「就労継続支援」%s、「児童発達支援」%sが多くなっています（回答が25件のため、幅をもって見ます）。" % (
            PK(JI, "問28-2", "利用したい", koumoku="放課後等デイ"), ZK("63.6％"),
            PK(JI, "問28-2", "利用したい", koumoku="就労継続"), PK(JI, "問28-2", "利用したい", koumoku="児童発達支援")))

    # 第1部第3章5 課題と対応
    p = d.cont("令和５年10月に実施したアンケート調査と、令和８年８月に実施したヒアリング")
    rep(p, "令和５年10月に実施したアンケート調査と", "令和8年8月から9月に実施したアンケート調査と")
    tb = d.tables("分野アンケート調査・ヒアリングの結果")[0]
    KADAI = {
        "権利擁護": "成年後見制度を「名前も内容も知らない」が障がい者%s、障がい児%s（前回調査 %s・%s）。"
                "ヒアリングでも制度の周知不足が指摘された" % (n37, j34, ZK("38.2％"), ZK("36.4％")),
        "差別の解消": "差別や嫌な思いをする（した）ことが「ある」「少しある」が障がい者%s（精神障がい%s）、障がい児%s" % (
            P(SHA, "問35", "ある", "少しある"), PB("問35", "精神障がい", "ある", "少しある"), P(JI, "問32", "ある", "少しある")),
        "移動・外出": "外出時に困ることは「外出にお金がかかる」%s、「困ったときにどうすればいいのか心配」%s、"
                  "「公共交通機関が少ない」%s（障がい者）" % (
                      P(SHA, "問22", "外出にお金がかかる"), P(SHA, "問22", "困ったときにどうすればいいのか心配"),
                      P(SHA, "問22", "公共交通機関が少ない（公共交通機関が通っていない）")),
        "防災": "災害時に一人で避難「できない」が障がい者%s、障がい児%s。近所に助けてくれる人が「いない」が障がい者%s、障がい児%s" % (
            P(SHA, "問41", "できない"), P(JI, "問38", "できない"), P(SHA, "問42", "いない"), P(JI, "問39", "いない")),
        "サービスの利用意向": "障がい者は相談支援%s、居宅介護%s、生活介護%s、短期入所%s。障がい児は就労継続支援%s、児童発達支援%s" % (
            P(SHA, "問31-2", "利用したい", koumoku="相談支援"), P(SHA, "問31-2", "利用したい", koumoku="居宅介護"),
            P(SHA, "問31-2", "利用したい", koumoku="生活介護"), P(SHA, "問31-2", "利用したい", koumoku="短期入所"),
            PK(JI, "問28-2", "利用したい", koumoku="就労継続"), PK(JI, "問28-2", "利用したい", koumoku="児童発達支援")),
        "就労": "就労支援として必要なことは「職場の上司や同僚に障がいの理解があること」%s、「通勤手段の確保」%s（障がい者）" % (
            P(SHA, "問27", "職場の上司や同僚に障がいの理解があること"), P(SHA, "問27", "通勤手段の確保")),
    }
    hit = 0
    for tr in tb.findall("w:tr", NS):
        tcs = tr.findall("w:tc", NS)
        k = tx(tcs[0]).strip()
        if k in KADAI:
            rep_all(tcs[1].find("w:p", NS), KADAI[k])
            hit += 1
    if hit != len(KADAI):
        raise ValueError("課題の表の行が%d件しか見つからない" % hit)
    p = d.start("資料：小野町障がい福祉に関するアンケート調査（令和５年10月実施）、令和８年８月のヒアリング")
    rep(p, "小野町障がい福祉に関するアンケート調査（令和５年10月実施）",
        "令和8年度 障がい福祉に関するアンケート調査（前回調査は令和5年10月実施）")
    LOG.append(("1・2", "第1部第3章5 課題と対応：表の6行を令和8年度調査の値に差し替え"))

    # 策定経過
    tb = d.tables("年月日内容")[0]
    for tr in tb.findall("w:tr", NS):
        tcs = tr.findall("w:tc", NS)
        if tx(tcs[0]).strip() == "令和５年10月":
            rep_all(tcs[0].find("w:p", NS), "令和8年8月14日～9月25日")
            rep_all(tcs[1].find("w:p", NS), "障がい福祉に関するアンケート調査の実施（障がい者・障がい児）")
    LOG.append(("1・2", "資料編3 策定経過：アンケート調査の時期を令和8年度に差し替え"))

    # 第2部 各章
    p = d.cont("前回のアンケート調査では、成年後見制度の認知度について")
    rep(p, "前回のアンケート調査では、成年後見制度の認知度について、「名前も内容も知らない」が障がい者・障がい児でともに約４割となっており、",
        "令和8年度の調査では、成年後見制度を「名前も内容も知らない」が障がいのある方の%s、障がいのある子どもの%sとなっており"
        "（前回調査はともに約4割）、" % (n37, j34))
    block(d, media, "■成年後見制度の認知度（令和５年度アンケート調査結果）", ["後見"], TM)

    p = d.cont("前回のアンケート調査によると、障がいがあることで差別や嫌な思いを")
    rep(p, "前回のアンケート調査によると、障がいがあることで差別や嫌な思いをする（した）ことがあるかでは、"
           "「ある」と「少しある」の合計が障がい者の約３割、障がい児の約７割となっています。"
           "さらに、どのような場所で差別や嫌な思いをしたかでは、障がい者は「学校・仕事場」、「外出先」、"
           "「病院などの医療機関」、「住んでいる地域」が比較的多く挙げられており、",
        "令和8年度の調査によると、障がいがあることで差別や嫌な思いをする（した）ことが「ある」「少しある」は、"
        "障がいのある方の%s、障がいのある子どもの%sです（前回調査は約3割・約7割）。"
        "障がい別では精神障がいのある方が%s、知的障がいのある方が%sと高くなっています。"
        "差別や嫌な思いをした場所は、経験のある方（%d人）のうち「外出先」%s、「学校・仕事場」%s、「住んでいる地域」%sが多く挙げられており、" % (
            P(SHA, "問35", "ある", "少しある"), P(JI, "問32", "ある", "少しある"),
            PB("問35", "精神障がい", "ある", "少しある"), PB("問35", "知的障がい", "ある", "少しある"),
            kiso, PG("問36-1", "外出先", kiso), PG("問36-1", "学校・仕事場", kiso), PG("問36-1", "住んでいる地域", kiso)))
    H.insert_after(p, para(TM["body"],
                           "また、障害者差別解消法を「知らない」は%s、障害者虐待防止法を「知らない」は%sで、法の周知が課題です。" % (
                               P(SHA, "問40", "知らない"), P(SHA, "問39", "知らない"))))
    block(d, media, "■障がいがあることで差別や嫌な思いをする（した）ことがあるか", ["差別", "差別場所", "法認知"], TM)

    p = d.cont("一方で、前回のアンケート調査では、制度を「名前も内容も知らない」")
    rep(p, "一方で、前回のアンケート調査では、制度を「名前も内容も知らない」が障がい者38.2％、障がい児36.4％であり（本章１）、",
        "一方で、令和8年度の調査では、制度を「名前も内容も知らない」が障がいのある方の%s、障がいのある子どもの%sであり（本章1）、" % (n37, j34))

    p = d.cont("前回のアンケート調査において、外出する時に困ることは")
    rep(p, "前回のアンケート調査において、外出する時に困ることは、障がい者では「困ったときにどうすればいいのか心配」、"
           "「道路や駅に階段や段差が多い」、「公共交通機関が少ない（通っていない）」がいずれも２割を超えて多く、"
           "次いで「外出にお金がかかる」、「列車やバスの乗り降りが困難」となっています。",
        "令和8年度の調査では、外出を「めったにしない」「まったくしない」障がいのある方が%sです。"
        "外出するときに困ることは「外出にお金がかかる」%s、「困ったときにどうすればいいのか心配」%s、"
        "「公共交通機関が少ない」%sが多く、前回調査で多かった「道路や駅に階段や段差が多い」は%sでした。" % (
            P(SHA, "問19", "めったに外出しない", "まったく外出しない"), P(SHA, "問22", "外出にお金がかかる"),
            P(SHA, "問22", "困ったときにどうすればいいのか心配"),
            P(SHA, "問22", "公共交通機関が少ない（公共交通機関が通っていない）"),
            P(SHA, "問22", "道路や駅に階段や段差が多い")))
    block(d, media, "■１週間の外出頻度", ["外出頻度", "外出困難"], TM)

    p = d.cont("前回のアンケート調査によると、災害時に一人で避難できるかは")
    rep(p, tx(p).strip(),
        "令和8年度の調査によると、災害時に一人で避難「できない」は障がいのある方の%s、障がいのある子どもの%sで、"
        "「わからない」を合わせると障がいのある方の%sが一人での避難に不安を抱えています。"
        "近所に助けてくれる人が「いない」は障がいのある方の%s、障がいのある子どもの%sです。"
        "災害時に困ることは「投薬や治療が受けられない」%sが最も多くなっています。" % (
            P(SHA, "問41", "できない"), P(JI, "問38", "できない"), P(SHA, "問41", "できない", "わからない"),
            P(SHA, "問42", "いない"), P(JI, "問39", "いない"), P(SHA, "問43", "投薬や治療が受けられない")))
    block(d, media, "■火事や地震等の災害時に一人で避難できるか", ["避難", "近所", "災害"], TM)

    p = d.cont("前回のアンケート調査では、現在受けている医療ケアは")
    rep(p, "前回のアンケート調査では、現在受けている医療ケアは、障がい者では「服薬」が最も多く、「その他」を除くと「透析」、「カテーテル留置」と続いています。",
        "令和8年度の調査では、現在受けている医療ケアは、障がいのある方では「服薬」%sが最も多く、"
        "「その他」を除くと「透析」%s、「ストマ」%sと続いています。" % (
            P(SHA, "問16", "服薬"), P(SHA, "問16", "透析"), P(SHA, "問16", "ストマ（人工肛門・人工膀胱）")))
    block(d, media, "■現在受けている医療ケア（上位項目・複数回答）", ["医療"], TM)

    p = d.cont("前回のアンケート調査によると、障がい福祉サービスの今後の利用意向は")
    sv = lambda k: P(SHA, "問31-2", "利用したい", koumoku=k)
    ima = lambda k: P(SHA, "問31-1", "利用している", koumoku=k)
    rep(p, tx(p).strip(),
        "令和8年度の調査によると、障がい福祉サービスの今後の利用意向は、障がいのある方では「相談支援」%s、"
        "「居宅介護」%s、「生活介護」%s、「短期入所」%sが高く、いずれも現在の利用（%s・%s・%s・%s）を大きく上回ります。"
        "障がいのある子どもでは「就労継続支援」%s、「児童発達支援」%sが多くなっています。" % (
            sv("相談支援"), sv("居宅介護"), sv("生活介護"), sv("短期入所"),
            ima("相談支援"), ima("居宅介護"), ima("生活介護"), ima("短期入所"),
            PK(JI, "問28-2", "利用したい", koumoku="就労継続"), PK(JI, "問28-2", "利用したい", koumoku="児童発達支援")))
    H.insert_after(p, para(TM["body"],
                           "地域で生活するために必要な支援は「経済的な負担の軽減」%s、「必要な在宅サービスが適切に利用できること」%s、"
                           "「在宅で医療ケアなどが適切に得られること」%sが上位です。" % (
                               P(SHA, "問18", "経済的な負担の軽減"), P(SHA, "問18", "必要な在宅サービスが適切に利用できること"),
                               P(SHA, "問18", "在宅で医療ケアなどが適切に得られること"))))
    block(d, media, "■今後の障がい福祉サービスの利用意向", ["利用意向", "地域生活"], TM)

    p = d.cont("前回のアンケート調査によると、障がい者の就労支援として必要なことは")
    rep(p, "前回のアンケート調査によると、障がい者の就労支援として必要なことは、「職場の上司や同僚に障がいの理解があること」が最も多く、"
           "次いで「通勤手段の確保」、「短時間勤務や勤務日数等の配慮」が続いています。",
        "令和8年度の調査によると、平日の日中を「自宅で過ごしている」障がいのある方は%sで、収入を得る仕事をしている方は%sです。"
        "障がい者の就労支援として必要なことは、「職場の上司や同僚に障がいの理解があること」%sが最も多く、"
        "次いで「通勤手段の確保」%s、「短時間勤務や勤務日数等の配慮」%sが続いています。" % (
            P(SHA, "問23", "自宅で過ごしている"), P(SHA, "問23", "会社勤めや、自営業、家業などで収入を得て仕事している"),
            P(SHA, "問27", "職場の上司や同僚に障がいの理解があること"), P(SHA, "問27", "通勤手段の確保"),
            P(SHA, "問27", "短時間勤務や勤務日数等の配慮")))
    block(d, media, "■障がい者の就労支援として必要なこと（上位項目・複数回答）", ["日中", "就労"], TM)

    # 調査のなかった章
    add_fig(d, media, "第４章　情報アクセシビリティの向上及び意思疎通支援の充実",
            "令和8年度の調査では、障がいや福祉サービスの情報の入手先は「本や新聞、雑誌の記事、テレビやラジオのニュース」%s、"
            "「家族や親せき、友人・知人」%s、「かかりつけの医師や看護師」%s、「行政機関の広報誌」%sの順で、"
            "「どこで情報を入手したらよいかわからない」方も%sいます。" % (
                P(SHA, "問34", "本や新聞、雑誌の記事、テレビやラジオのニュース"), P(SHA, "問34", "家族や親せき、友人・知人"),
                P(SHA, "問34", "かかりつけの医師や看護師"), P(SHA, "問34", "行政機関の広報誌"),
                P(SHA, "問34", "どこで情報を入手したらよいかわからない")), ["情報"], TM)
    add_fig(d, media, "１　意思決定支援及び相談支援体制の充実",
            "令和8年度の調査では、悩みや困ったことの相談相手は「家族や親せき」が%sと突出し、「かかりつけの医師や看護師」%s、"
            "「友人・知人」%sが続きます。「相談支援事業所などの民間の相談窓口」「行政機関の相談窓口」はいずれも%sにとどまり、"
            "「相談する相手がいない」方も%sいます。" % (
                P(SHA, "問33", "家族や親せき"), P(SHA, "問33", "かかりつけの医師や看護師"), P(SHA, "問33", "友人・知人"),
                P(SHA, "問33", "行政機関の相談窓口"), P(SHA, "問33", "相談する相手がいない")), ["相談"], TM)
    add_fig(d, media, "１　インクルーシブ教育システムの推進",
            "令和8年度の調査（障がいのある子ども、回答25件）では、平日の日中は「一般の小中学校、高校に通っている」が%sで最も多く、"
            "「一般の小中学校の支援学級に通っている」%s、「特別支援学校（小中高等部）に通っている」%sです。"
            "地域で生活するために必要な支援は「経済的な負担の軽減」%s、「必要な在宅サービスが適切に利用できること」%s、"
            "「障がい者に適した住居の確保」%sが上位です。" % (
                PK(JI, "問17", "一般の小中学校、高校に通っている"), PK(JI, "問17", "一般の小中学校の支援学級に通っている"),
                PK(JI, "問17", "特別支援学校（小中高等部）に通っている"), PK(JI, "問24", "経済的な負担の軽減"),
                PK(JI, "問24", "必要な在宅サービスが適切に利用できること"), PK(JI, "問24", "障がい者に適した住居の確保")),
            ["児日中", "児支援"], TM)
    add_fig(d, media, "１　文化芸術活動・スポーツ活動等の振興",
            "令和8年度の調査では、外出の目的は「医療機関へ受診」%s、「買い物に行く」%sが多く、"
            "「趣味やスポーツをする」は%s、「グループ活動に参加する」は%sにとどまります。" % (
                P(SHA, "問21", "医療機関へ受診"), P(SHA, "問21", "買い物に行く"),
                P(SHA, "問21", "趣味やスポーツをする"), P(SHA, "問21", "グループ活動に参加する")), ["外出目的"], TM)


# ------------------------------------------------------------------ 作業5　当方の図表を加える
def toho_tables():
    """当方素案（第5版）の表を題名で読む。{題名: rows}"""
    root = etree.fromstring(zipfile.ZipFile(TOHO).read("word/document.xml"))
    kids = list(root.find(W + "body"))
    out = {}
    for i, e in enumerate(kids):
        m = re.match(r"【表\d+】(.+)", tx(e)) if e.tag == W + "p" else None
        if m and i + 1 < len(kids) and kids[i + 1].tag == W + "tbl":
            out[m.group(1).strip()] = [[tx(c) for c in r.findall(W + "tc")] for r in kids[i + 1].findall(W + "tr")]
    return out


def tw(cm):
    return [int(x * 567) for x in cm]


def add_toho(d, media):
    T = toho_tables()
    like = d.tables("分野アンケート調査・ヒアリングの結果")[0]
    conv = lambda s: HY.hyoki(s.replace("【町データ待ち】", TBD))

    # a 手帳所持者数の見通し
    a = d.start("資料：健康福祉課、小野町地域福祉計画（各年１月１日現在）")
    rows = T["障がい者手帳所持者数の見通し（所持率による推計の幅）"]
    if rows[0][0] != "手帳の種類" or "令和11年" not in rows[0]:
        raise ValueError("当方素案の手帳の見通しの表の形が違う")
    H.insert_after(a,
                   para(TM["body"], "このほか、手帳の所持率（総人口に占める所持者の割合）を用いて、計画期間の所持者数を2つの方法で推計すると、次のとおりです。"
                                    "Aは令和8年の所持率を据え置く方法、Bは令和元～8年の所持率の傾向を延ばす方法で、表は小さい方～大きい方の幅で示します。"
                                    "上の推計（小野町地域福祉計画）とあわせて見て、見込量の前提とします。"),
                   cap(TM["cap"], "障がい者手帳所持者数の見通し（所持率による推計の幅）"),
                   H.new_table(like, tw([4.0, 2.2, 2.2, 2.2, 2.2, 2.6]), [conv(x) for x in rows[0]],
                               [[conv(x) for x in r] for r in rows[1:]]),
                   para(TM["src"], "資料：町の見込量・目標値設定シート（令和8年9月）の手帳所持者数と総人口から推計。町の見込は同シートの令和11年の値"))
    LOG.append(("5", "第1部第3章2：手帳所持者数の見通し（所持率による推計の幅）の表を追加"))

    # b 障がい別の状況・前回調査との比較（第1部第3章5の末尾）
    a = d.start("資料：令和8年度 障がい福祉に関するアンケート調査（前回調査は令和5年10月実施）")
    r15 = T["主な項目の障がい別の状況（令和8年度調査）"]
    r22 = T["前回調査（令和5年度）と今期調査（令和8年度）の比較"]
    if r15[0][:2] != ["障がい", "項目"] or r22[0][0] != "調査項目":
        raise ValueError("当方素案の障がい別・前回比較の表の形が違う")
    hin = [[conv(x) for x in r[:3]] for r in r22[1:]]
    H.insert_after(a,
                   para(TM["body"], "障がい別にみると、次の項目で障がいのある方全体との差が大きくなっています。"),
                   cap(TM["cap"], "主な項目の障がい別の状況（令和8年度調査）"),
                   H.new_table(like, tw([2.6, 8.4, 2.4, 2.4]), [conv(x) for x in r15[0]],
                               [[conv(x) for x in r] for r in r15[1:]]),
                   para(TM["src"], SHIRYO_R8 + "（障がい別は所持している手帳による）"),
                   para(TM["body"], "前回調査（令和5年度）と比べると、主な項目は次のとおりです。"),
                   cap(TM["cap"], "前回調査（令和5年度）と今期調査（令和8年度）の比較"),
                   H.new_table(like, tw([5.0, 5.4, 5.4]), [conv(x) for x in r22[0][:3]], hin),
                   para(TM["src"], "資料：障がい福祉アンケート調査結果報告書（令和5年度）、" + SHIRYO_R8[3:]
                        + "。※は前回と設問・基数が異なるため、比べる際に注意を要するもの"))
    LOG.append(("5", "第1部第3章5：主な項目の障がい別の状況・前回調査との比較の表を追加"))

    # c 算定方法・d 1人あたり日数（第3部第6章 見込量の定め方の後）
    a = d.tables("加減項目拠り所")[0]
    r36, r40 = T["サービスの区分ごとの見込量の算定方法"], T["1人あたりの利用日数　前回計画の見込みと実績"]
    if r36[0][:3] != ["区分", "サービス", "算定方式"] or r40[0][0] != "サービス":
        raise ValueError("当方素案の算定方法・1人あたり日数の表の形が違う")
    H.insert_after(a,
                   para(TM["body"], "サービスの区分ごとの算定の方法は、次のとおりです。"),
                   cap(TM["cap"], "サービスの区分ごとの見込量の算定方法"),
                   H.new_table(like, tw([2.4, 4.6, 4.2, 4.6]), [conv(x) for x in r36[0]],
                               [[conv(x) for x in r] for r in r36[1:]]),
                   para(TM["body"], "前回計画（第6期）の見込量と実績を比べると、人数はおおむね見込みどおりである一方、"
                                    "1人あたりの利用日数が見込みを大きく下回っていました。"
                                    "このため第8期は、量（人日）を「見込人数×1人あたり日数」とし、1人あたり日数は直近の実績の値を出発点とします。"
                                    "日数を引き上げる場合は、通えない理由（送迎・費用・事業所の有無）への手当てを第2部の施策に書きます。"),
                   cap(TM["cap"], "1人あたりの利用日数　前回計画の見込みと実績"),
                   H.new_table(like, tw([3.6, 2.6, 2.6, 7.0]), [conv(x) for x in r40[0]],
                               [[conv(x) for x in r] for r in r40[1:]]),
                   para(TM["src"], "資料：前回計画（令和6年3月）の第6期の見込量・実績から算出（人日÷人。令和3～5年度の幅）"),
                   cap(TM["cap"], "1人あたりの利用日数　見込みと実績（第6期）"),
                   img_para(TM["img"], media, FIG_NISSU),
                   para(TM["src"], "資料：前回計画（令和6年3月）の第6期の見込量・実績から算出（令和3～5年度の平均）。右端は実績÷見込み"))
    LOG.append(("5", "第3部第6章：算定方法の表、1人あたり利用日数の表と図を追加"))


# ------------------------------------------------------------------ 作業3　番号・題名・資料
SAKU = "施策の方向"
# 題名のない表のうち、直前が本文で、題名を決めて付けるもの（1行目の1列目 → 題名）
TEHAI = {
    "年齢階層": "年齢階層別人口",
    "機能": "地域生活支援拠点等の5つの機能",
    "活動指標": "発達障がい者等に対する支援の活動指標",
    "段階": "見込量を定めるために把握する段階",
    "加減": "必要な利用者数の組み立て",
    "区分現在の人数": "重度の障がいのある人の人数と主に必要となるサービス",
    "サービス町内": "田村地方の事業所の状況とサービスごとの確保の方針",
    "照会する事項": "事業所の受入余力について照会する事項",
    "事項両計画での扱い": "介護保険と障がい福祉の両計画での扱い",
    "#場面": "介護保険と障がい福祉の接点（場面と本町での扱い）",
    "#確かめること": "両計画の接点について町に確かめること",
    "部門": "小野町自立支援協議会委員名簿",
    "年月日": "計画の策定経過",
    "年主な動向": "近年の障がい者施策の動向",
    "種別": "田村地方の障がい福祉サービス事業所一覧",
    "行": "用語集",
    "期間区分": "前計画と本計画の期間",
    "分野関係機関": "関係機関と連携の内容",
    "年齢": "年齢到達に伴うサービスの移行",
    "事項": "地域生活支援事業の実施に必要な事項",
}
GENERIC = re.compile(r"^■(第[６７８６-８\d]期の(実績|見込量|目標値|計画値)|第[７7]期の実績と|見込量)")
NUM = re.compile(r"^(（[0-9０-９]+）|\([0-9]+\)|[①-⑳]|[0-9０-９]+　|[0-9０-９]+\s)")


def is_del_tbl(tb):
    trs = tb.findall("w:tr", NS)
    return trs and all(tr.find("w:trPr/w:del", NS) is not None for tr in trs)


def is_del_p(p):
    return p.find("w:pPr/w:rPr/w:del", NS) is not None and not tx(p).strip()


def has_drawing(p):
    return p.find(".//" + W + "drawing") is not None


def has_chart(p):
    return p.find(".//{%s}chart" % CHART) is not None


def clean(s):
    s = s.strip().lstrip("■")
    s = re.sub(r"^第[0-9０-９]+章\s*", "", s)
    s = NUM.sub("", s).strip()
    s = re.sub(r"【[^】]*】", "", s).strip()
    return s


def bango(d, tmpl):
    """図・表に通し番号と固有の題名を付ける。"""
    kids = list(d.body)
    heading = sub = ""
    nz = nt = 0
    added_cap = added_src = 0
    skip_draw = set()
    i = 0
    while i < len(kids):
        e = kids[i]
        if e.tag == W + "p":
            t = tx(e).strip()
            st = e.find("w:pPr/w:pStyle", NS)
            sv = st.get(W + "val") if st is not None else ""
            if sv in ("1", "Heading1", "11") and t:
                heading = t
                sub = ""
            elif sv in ("2", "Heading2", "21") or re.match(r"^[0-9０-９]+　", t) and len(t) < 40 and not has_drawing(e):
                heading = t
                sub = ""
            elif re.match(r"^(（[0-9０-９]+）|\([0-9]+\)|[①-⑳])", t) and len(t) < 60:
                sub = t
            elif (t and len(t) < 45 and not t.endswith("。") and not has_drawing(e)
                  and not re.match(r"^(■|【|資料|※|〈|・|出所)", t)):
                sub = t     # サービス名だけの行（「居宅介護」など）
            # 図
            if has_drawing(e) and e not in skip_draw and not in_del(e):
                capp = None
                if t.startswith("■"):
                    capp = e
                else:
                    j = i - 1
                    while j >= 0 and kids[j].tag == W + "p" and (not tx(kids[j]).strip() or is_del_p(kids[j])):
                        j -= 1
                    if kids[j].tag == W + "p" and tx(kids[j]).strip().startswith("■"):
                        capp = kids[j]
                if capp is not None or has_chart(e) or e in FIG_PARAS:
                    nz += 1
                    if capp is None:
                        title = "人口構成（人口ピラミッド）" if "人口構成" in t else clean(sub or heading)
                        np_ = para(tmpl["cap"], "【図%d】%s" % (nz, title))
                        e.addprevious(np_)
                        added_cap += 1
                    else:
                        set_num(capp, "図", nz, sub or heading)
                    # 同じ題名に続く図（2枚で1つの図）は数えない
                    k = i + 1
                    while k < len(kids) and kids[k].tag == W + "p" and has_drawing(kids[k]) and not tx(kids[k]).strip():
                        skip_draw.add(kids[k])
                        k += 1
        elif e.tag == W + "tbl" and not is_del_tbl(e):
            first = [tx(c).strip() for c in e.findall("w:tr", NS)[0].findall("w:tc", NS)]
            if first and first[0] == "田村市":    # 事業所一覧の凡例（1行）。一覧と一体
                i += 1
                continue
            nt += 1
            j = i - 1
            while j >= 0 and kids[j].tag == W + "p" and (not tx(kids[j]).strip() or is_del_p(kids[j])):
                j -= 1
            prev = kids[j] if kids[j].tag == W + "p" else None
            pt = tx(prev).strip() if prev is not None and not has_drawing(prev) else ""
            if pt.startswith("■"):
                set_num(prev, "表", nt, sub or heading)
            else:
                key1 = first[0]
                key2 = first[0] + (first[1] if len(first) > 1 else "")
                if key1.startswith("①") or pt == "〈 施策の方向 〉":
                    title = "%sの%s" % (clean(heading), SAKU)
                elif key2 in TEHAI:
                    title = TEHAI[key2]
                elif key1 in TEHAI and not (pt and NUM.match(pt)):
                    title = TEHAI[key1]
                elif pt and (NUM.match(pt) or len(pt) < 40):
                    base = clean(pt)
                    title = base + ("の内容" if key1 in ("給付の種類", "内容", "サービス名") else
                                    "の目標" if key1 == "項目" else "")
                else:   # 直前が本文：見出し（小見出し）から題名を作る。題名の一覧は自己点検の出力で確かめる
                    base = clean(sub or heading)
                    title = base + ("の内容" if key1 in ("給付の種類", "内容", "サービス名") else
                                    "の目標" if key1 == "項目" else "")
                    if not base:
                        raise ValueError("表の題名を決められない：%s／直前「%s」" % (first[:2], pt[:30]))
                np_ = para(tmpl["cap"], "【表%d】%s" % (nt, title))
                e.addprevious(np_)
                added_cap += 1
                pt = title
            # 資料：実績・見込量の表で資料がないもの
            k = i + 1
            while k < len(kids) and kids[k].tag == W + "p" and not tx(kids[k]).strip():
                k += 1
            nx = tx(kids[k]).strip() if k < len(kids) and kids[k].tag == W + "p" else ""
            if "資料" not in nx[:12]:
                if first and first[0] in ("第７期", "第7期"):
                    e.addnext(para(tmpl["src"], "資料：健康福祉課（障がい福祉サービス等の給付実績。令和8年度は見込み）、"
                                                "第7期の計画値は前回計画（令和6年3月）"))
                    added_src += 1
                elif first and first[0] in ("第８期", "第8期"):
                    e.addnext(para(tmpl["src"], "資料：本計画による推計（本章冒頭の見込量の定め方による）"))
                    added_src += 1
        i += 1
    LOG.append(("3", "図%d点・表%d点に通し番号。題名のない図表%d点に題名を追加、実績・見込量の表%d点に資料を追加"
                % (nz, nt, added_cap, added_src)))
    return nz, nt


def set_num(p, kind, n, sub):
    """題名の段落の「■」を【図N】【表N】に。見込量・実績の題名には、サービス名を添えて固有の題名にする。"""
    t = tx(p).strip()
    head = "【%s%d】" % (kind, n)
    if GENERIC.match(t) and sub:
        head += clean(sub) + "の"
    if p in INSERTED_CAP:
        for tt in p.iter(W + "t"):
            if tt.text and "■" in tt.text:
                tt.text = tt.text.replace("■", head, 1)
                return
    # 図の段落に題名が同居しているもの・既存の題名：変更履歴で置き換える
    for r in p.iter(W + "r"):
        tt = r.find("w:t", NS)
        if tt is not None and tt.text and "■" in tt.text and not in_del(r):
            if in_ins(r):
                tt.text = tt.text.replace("■", head, 1)
            else:
                a = tt.text.index("■")
                split_run(r, a, a + 1, head)
            return
    raise ValueError("題名の「■」が見当たらない：%s" % t[:30])


# ------------------------------------------------------------------ 作業4　書き分け・半角
def hyoki_all(d):
    n = 0
    for p in d.body.iter(W + "p"):
        if in_del(p):
            continue
        H._merge_runs(p)
        changed = True
        while changed:
            changed = False
            for r in list(p.iter(W + "r")):
                tt = r.find("w:t", NS)
                if tt is None or not tt.text or in_del(r):
                    continue
                if in_ins(r):
                    s2 = HY.hyoki(tt.text)
                    if s2 != tt.text:
                        tt.text = s2
                        n += 1
                    continue
                for rx, b in HY.RULES:
                    m = rx.search(tt.text)
                    if m:
                        split_run(r, m.start(), m.end(), b)
                        n += 1
                        changed = True
                        break
                if changed:
                    break
    LOG.append(("4", "「障害」「障がい」の書き分けを前回計画の規則に：%d か所（変更履歴）" % n))


_HAN = str.maketrans("０１２３４５６７８９", "0123456789")


def hankaku(d, styles=None):
    n = 0
    roots = [d.body] + ([styles] if styles is not None else [])
    for root in roots:
        for t in root.iter(W + "t"):
            if not t.text:
                continue
            s2 = t.text.translate(_HAN).replace("%", "％")
            if s2 != t.text:
                n += sum(1 for a, b in zip(t.text, s2) if a != b)
                t.text = s2
    LOG.append(("4", "全角数字を半角に、「%%」を「％」に：%d 文字（文字の形だけの変更のため変更履歴にしない）" % n))
    return n


# ------------------------------------------------------------------ 組み立て
ZU = {}
TM = {}


def build():
    global ZU
    ZU = Z.build_all()
    work = pathlib.Path(tempfile.mkdtemp())
    # 第1弾を一時の場所に作り、その上に重ねる
    tmp1 = work / "dai1.docx"
    H.OUT = tmp1
    H.build()
    ng1 = H.jiko()
    if ng1:
        raise ValueError("第1弾の自己点検：" + "／".join(ng1))
    H._ID[0] = 95000
    wd = work / "x"
    with zipfile.ZipFile(tmp1) as z:
        z.extractall(wd)
        names = z.namelist()
    for p in wd.rglob("*"):
        if p.is_symlink():
            p.unlink()
    d = D(wd / "word" / "document.xml")
    ids = [int(e.get(W + "id")) for e in d.body.iter() if (e.get(W + "id") or "").isdigit()]
    H._ID[0] = max(ids) + 100
    media = Media(wd)
    TM["cap"] = copy.deepcopy(d.start("■成年後見制度の認知度（令和５年度アンケート調査結果）"))
    TM["src"] = copy.deepcopy(d.start("資料：健康福祉課、小野町地域福祉計画（各年１月１日現在）"))
    TM["body"] = copy.deepcopy(d.cont("家庭、障がい者福祉施設、職場において虐待を見つけた人には通報の義務がある"))
    TM["img"] = TM["body"]

    sakuin(d, media, TM)
    add_toho(d, media)
    hyoki_all(d)
    bango(d, TM)
    st_p = wd / "word" / "styles.xml"
    hankaku(d)
    media.save()
    d.tree.write(str(wd / "word" / "document.xml"), xml_declaration=True, encoding="UTF-8", standalone=True)
    if OUT.exists():
        OUT.unlink()
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        all_names = list(names) + ["word/media/" + p.name for p in (wd / "word" / "media").glob("zu_r8_*.png")]
        for n in all_names:
            z.write(wd / n, n)
    shutil.rmtree(work)


# ------------------------------------------------------------------ 自己点検
def accepted(path):
    with zipfile.ZipFile(path) as z:
        root = etree.fromstring(z.read("word/document.xml"))
    for e in list(root.iter(W + "del")):
        if e.getparent().tag not in (W + "rPr", W + "trPr"):
            e.getparent().remove(e)
    for tr in list(root.iter(W + "tr")):
        if tr.find("w:trPr/w:del", NS) is not None:
            tr.getparent().remove(tr)
    return root


def jiko():
    ng = []
    root = accepted(OUT)
    paras = ["".join(t.text or "" for t in p.iter(W + "t")) for p in root.iter(W + "p")]
    acc = "\n".join(paras)
    # 1 前回調査を土台にした記述が残っていない
    for w in ("前回のアンケート調査", "令和５年度アンケート調査結果", "（令和５年10月実施）", "（令和5年10月実施）\n"):
        if w in acc:
            ng.append("承諾後の本文に「%s」が残っている" % w.strip())
    if acc.count("令和8年度 障がい福祉に関するアンケート調査") < 14:
        ng.append("令和8年度調査の資料の行が少ない（%d）" % acc.count("令和8年度 障がい福祉に関するアンケート調査"))
    # 2 図の数
    with zipfile.ZipFile(OUT) as z:
        n_media = len([n for n in z.namelist() if n.startswith("word/media/zu_r8_")])
        doc = z.read("word/document.xml").decode("utf-8")
        rels = z.read("word/_rels/document.xml.rels").decode("utf-8")
    want = sum(len(v) for v in ZU_HAICHI.values()) + 1
    if n_media != want:
        ng.append("挿入した図が%d点（%d点のはず）" % (n_media, want))
    for k in range(1, n_media + 1):
        if 'r:embed="rIdZu%02d"' % k not in doc or 'Id="rIdZu%02d"' % k not in rels:
            ng.append("図%dの参照が切れている" % k)
    # 3 書き分け・半角
    bad = HY.ihan(acc)
    if bad:
        ng.append("書き分けの規則に合わない語が%d件：%s" % (len(bad), bad[:3]))
    zen = re.findall(r"[０-９]", acc)
    if zen:
        ng.append("全角数字が%d文字残っている" % len(zen))
    if "%" in acc:
        ng.append("半角の「%」が残っている")
    # 4 図表の番号が1から連続し、題名が付いている
    for kind in ("図", "表"):
        ns = [int(m.group(1)) for m in re.finditer(r"【%s(\d+)】" % kind, acc)]
        if ns != list(range(1, len(ns) + 1)):
            ng.append("%sの番号が連続しない：%s" % (kind, ns[:12]))
        if any(p.strip().startswith("【%s" % kind) and len(re.sub(r"【.\d+】", "", p).strip()) < 3 for p in paras):
            ng.append("%sの題名が空のものがある" % kind)
    n_tbl = len([tb for tb in root.iter(W + "tbl") if tb.find("w:tr", NS) is not None])   # 削除した表（行なし）を除く
    n_hyo = len(re.findall(r"【表\d+】", acc))
    if n_tbl - 1 != n_hyo:     # 事業所一覧の凡例（1行の表）を除く
        ng.append("表%d個（凡例を除く）に対し【表N】が%d個" % (n_tbl - 1, n_hyo))
    if any(p.strip().startswith("■") and ("第７期" in p or "第8期" in p or "第７期" in p) for p in paras):
        ng.append("番号のない題名（■）が残っている")
    # 5 前回調査の割合が前回調査の報告書にある
    zt = subprocess.run(["pdftotext", "-layout", str(ZENKAI_CHOSA), "-"], capture_output=True, text=True).stdout
    zt = unicodedata.normalize("NFKC", re.sub(r"\s+", "", zt))
    for v in ZENKAI_ATAI:
        if unicodedata.normalize("NFKC", v).replace("％", "%") not in zt:
            ng.append("前回調査の割合 %s が前回調査の報告書にない" % v)
    # 6 調査の図の割合が報告書と一致
    ng += Z.zu_check()
    # 7 両計画の接点の文言（書き分けを除いて同じ）
    a2 = unicodedata.normalize("NFKC", HY.kyotsu(acc))
    for i, s in enumerate(SET.LEAD):
        if unicodedata.normalize("NFKC", HY.kyotsu(s.replace("**", ""))) not in a2:
            ng.append("接点の前書き%dが承諾後の本文にない" % (i + 1))
    # 8 他団体・禁止語
    for w in ("北塩原", "福島町", "北海道", "大雪", "川崎町", "金ヶ崎", "裏磐梯", "会津", "訂正", "お詫び", "おわび"):
        if w in acc:
            ng.append("「%s」が出ている" % w)
    # 9 変更履歴の作成者
    r2 = etree.fromstring(doc.encode("utf-8"))
    au = {e.get(W + "author") for e in r2.iter(W + "ins", W + "del")}
    if au != {H.AUTHOR}:
        ng.append("変更履歴の作成者が %s" % au)
    return ng


# 図の置き場所（自己点検で数を確かめる）
ZU_HAICHI = {
    "第2章1": ["後見"], "第2章2": ["差別", "差別場所", "法認知"], "第3章2": ["外出頻度", "外出困難"],
    "第4章": ["情報"], "第5章1": ["避難", "近所", "災害"], "第6章2": ["医療"], "第7章1": ["相談"],
    "第7章2": ["利用意向", "地域生活"], "第8章1": ["児日中", "児支援"], "第9章1": ["日中", "就労"], "第10章1": ["外出目的"],
}


if __name__ == "__main__":
    build()
    print("■ 他素案への反映 第2弾（変更履歴付き）")
    print("   %s" % OUT.relative_to(HERE))
    for no, s in LOG:
        print("   [%s] %s" % (no, s))
    ng = jiko()
    print("\n■ 自己点検 9項目")
    for x in ng:
        print("   NG", x)
    print("   すべて適合" if not ng else "   **%d件**" % len(ng))
    sys.exit(1 if ng else 0)
