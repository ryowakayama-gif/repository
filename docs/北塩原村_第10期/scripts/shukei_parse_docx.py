# -*- coding: utf-8 -*-
"""調査票のdocxから設問と選択肢を抽出してコードブックの素案を作る。
   ルビ（w:rt）は読み仮名なので読み飛ばす。
   使い方: python3 shukei_parse_docx.py <docx> [--json out.json]
"""
import sys, zipfile, re, json
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def text_of(el):
    """ルビの読み（w:rt）を除いたテキストを返す"""
    out = []
    def rec(node):
        if node.tag == W + "rt":       # ルビの読み。本文ではないので飛ばす
            return
        if node.tag == W + "t":
            out.append(node.text or "")
        elif node.tag == W + "tab":
            out.append("\t")
        elif node.tag == W + "br":
            out.append("\n")
        for c in node:
            rec(c)
    rec(el)
    return "".join(out)


def lines_of(path):
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    body = root.find(W + "body")
    lines = []

    def walk(el):
        for child in el:
            if child.tag == W + "p":
                t = re.sub(r"\s+", " ", text_of(child)).strip()
                if t:
                    lines.append(t)
            elif child.tag == W + "tbl":
                for tr in child.findall(W + "tr"):
                    for tc in tr.findall(W + "tc"):
                        for p in tc.findall(W + "p"):
                            t = re.sub(r"\s+", " ", text_of(p)).strip()
                            if t:
                                lines.append(t)
                        for tbl in tc.findall(W + "tbl"):
                            walk(tc)
            elif child.tag in (W + "sdt", W + "sdtContent", W + "txbxContent"):
                walk(child)
    walk(body)
    return lines


# 設問見出し：問1 / 問１ / （1） / (1) / ① / A票 問6-1 など
RE_MON = re.compile(r"^問\s*([０-９0-9]+)(?:\s*[-－ー]\s*([０-９0-9]+))?")
RE_EDA = re.compile(r"^[（(]\s*([０-９0-9]+)\s*[)）]")
RE_MARU = re.compile(r"^([①-⑳])")
# 選択肢：１．〜 / 1. / 10． を1行から複数拾う
RE_OPT = re.compile(r"(?:^|\s)([０-９0-9]{1,2})[．.]\s*([^０-９]*?)(?=(?:\s[０-９0-9]{1,2}[．.])|$)")

Z2H = str.maketrans("０１２３４５６７８９", "0123456789")


def balance(lab):
    """選択肢ラベルの括弧を閉じる（正規表現の切り出しで末尾が欠けるため）"""
    for op, cl in (("（", "）"), ("(", ")")):
        if lab.count(op) > lab.count(cl):
            lab += cl * (lab.count(op) - lab.count(cl))
    return lab


def parse(path):
    items, cur = [], None
    mon = eda = ""
    for ln in lines_of(path):
        m_mon, m_eda, m_maru = RE_MON.match(ln), RE_EDA.match(ln), RE_MARU.match(ln)
        head = None
        if m_mon:
            head = "問" + m_mon.group(1).translate(Z2H)
            if m_mon.group(2):
                head += "-" + m_mon.group(2).translate(Z2H)
        elif m_eda:
            head = "(" + m_eda.group(1).translate(Z2H) + ")"
        elif m_maru:
            head = m_maru.group(1)
        if head:
            body = ln[len(head):] if not m_maru else ln[1:]
            if m_mon:
                mon, eda = head, ""
                key = head
            elif m_eda:
                eda = head
                key = mon + head
            else:
                key = mon + eda + head
            cur = {"番号": key, "見出し": head, "設問文": body.strip("　 |"), "選択肢": []}
            items.append(cur)
            rest = body
        else:
            rest = ln
        if cur is None:
            continue
        for num, lab in RE_OPT.findall(rest):
            lab = balance(lab.strip("　 |"))
            if lab:
                cur["選択肢"].append((int(num.translate(Z2H)), lab))
    # 選択肢の重複を除き番号順に
    for it in items:
        seen, opts = set(), []
        for n, l in it["選択肢"]:
            if n in seen:
                continue
            seen.add(n)
            opts.append([n, l])
        it["選択肢"] = sorted(opts)
    return items


if __name__ == "__main__":
    items = parse(sys.argv[1])
    if "--json" in sys.argv:
        out = sys.argv[sys.argv.index("--json") + 1]
        json.dump(items, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        print(f"{len(items)}件 -> {out}")
    else:
        for it in items:
            print(f"■ {it['番号']}　{it['設問文'][:50]}")
            for n, l in it["選択肢"]:
                print(f"    {n}. {l[:40]}")
