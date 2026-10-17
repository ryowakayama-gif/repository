# -*- coding: utf-8 -*-
"""計画素案 Ver.2.4 → Ver.2.5（書体と文字の大きさを大雪広域の協議用素案に合わせる）

大雪地区広域連合の協議用素案（令和8年8月）の styles.xml を調べたところ、
次のようになっていた。

  docDefaults      テーマフォント・11.0pt・行送り1.15（after=200 line=276 auto）
  Normal（既定）   游ゴシック 10.5pt　※既定の段落スタイルが定義されている
  Heading1         游ゴシック 14.0pt 太字 色365F91
  Heading2         游ゴシック 13.0pt 太字 色4F81BD
  Heading3         游ゴシック 10.5pt 太字 色4F81BD
  Caption          9.0pt 太字
  本文の run       10.0pt（629 run）
  表内の run       8.5pt（6,167 run）

一方、川崎町の素案には次の3つの問題があった。

  ① 既定の段落スタイル（Normal）が定義されていない
     Title・Heading1〜6・Strong・ListParagraph はいずれも
     <w:basedOn w:val="Normal"/> を持つが、Normal 自体が styles.xml にない。
     Word は既定の段落スタイルがない文書を開くと自前の「標準」を当てるため、
     本文が明朝体で表示されることがある。これが「見づらい」の主因と考えられる。
  ② 同じ用途の要素に複数の大きさが混在している
     ▌小見出し 11.0pt が152か所・10.5pt が7か所（Ver.2.2以降に加えた分）
     ⚠注記 10.5pt が31か所・9.5pt が17か所・11.0pt が7か所
  ③ 本文10.5pt・表9.0pt であり、大雪広域（本文10.0pt・表8.5pt）と揃っていない

本スクリプトは次を行う。

  1. 既定の段落スタイル Normal（游ゴシック 10.5pt）を定義し、
     すべてのスタイルに游ゴシックを明示する
  2. 文字の大きさを用途ごとに揃え、大雪広域の水準に合わせる
       章扉        15.0 → 14.0pt（大雪 Heading1）
       節見出し    12.5 → 13.0pt（大雪 Heading2）
       ▌小見出し  11.0／10.5 → 10.5pt（大雪 Heading3）
       本文        10.5 → 10.0pt（大雪の本文）
       ●要点      10.0 → 10.0pt（据置き）
       ⚠注記      10.5／9.5／11.0 → 10.0pt（本文に統一）
       図キャプション 9.0 → 9.0pt（大雪 Caption）
       出典         8.5 → 8.5pt（据置き）
       表内         9.0 → 8.5pt（大雪の表）
       表内（体系図）11.0 → 10.5pt
       目次の表     10.0 → 10.0pt（据置き）
  3. 書体の指定がない run に游ゴシックを補う
  4. ヘッダーの版表記（「計画書素案v1.0」のまま）を改める

⚠ 表紙・ご挨拶・目次の見出しの大きさは、体裁上の意図があるため変更しない。
"""
import copy
import re
import shutil
import sys
import zipfile

import docx
from docx.oxml.ns import qn

SRC = "01_第10期_最新版成果品/川崎町_計画書素案_v2.4_制度改正反映版.docx"
DST = "01_第10期_最新版成果品/川崎町_計画書素案_v2.5_書体統一版.docx"

GOTHIC = "游ゴシック"

# ══════════════════════════ 既定の段落スタイル（大雪広域の Normal と同じ）
NORMAL_STYLE = (
    '<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    ' w:type="paragraph" w:default="1" w:styleId="Normal">'
    '<w:name w:val="Normal"/><w:qFormat/>'
    '<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr>'
    f'<w:rPr><w:rFonts w:ascii="{GOTHIC}" w:hAnsi="{GOTHIC}"'
    f' w:eastAsia="{GOTHIC}" w:cs="{GOTHIC}"/>'
    '<w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:style>')

# ══════════════════════════ 大雪広域に合わせる見出しの大きさ（半ポイント単位）
STYLE_SZ = {
    "Heading1": 28,   # 14.0pt
    "Heading2": 26,   # 13.0pt
    "Heading3": 21,   # 10.5pt
    "Heading4": 21,
    "Heading5": 21,
    "Heading6": 21,
}
STYLE_COLOR = {
    "Heading1": "365F91",
    "Heading2": "4F81BD",
    "Heading3": "4F81BD",
}

# ══════════════════════════ 用途ごとの大きさ（ポイント）
SZ_SHO = 14.0        # 章扉
SZ_SETSU = 13.0      # 節見出し（1-1 等）
SZ_SUB = 10.5        # ▌小見出し
SZ_HONBUN = 10.0     # 本文・⚠注記・●要点
SZ_CAP = 9.0         # 図のキャプション
SZ_SHUTTEN = 8.5     # 出典・資料
SZ_TABLE = 8.5       # 表内
SZ_TABLE_SUB = 10.5  # 表内の体系図等

RE_SHO = re.compile(r"^第\s*\d+\s*章$")
RE_SETSU = re.compile(r"^\d+-\d+　")
RE_CAP = re.compile(r"^図\d+-\d+　")


def set_sz(run_el, pt):
    """run の w:sz / w:szCs を pt に揃える。"""
    rpr = run_el.find(qn("w:rPr"))
    if rpr is None:
        rpr = run_el.makeelement(qn("w:rPr"), {})
        run_el.insert(0, rpr)
    val = str(int(round(pt * 2)))
    for tag in ("w:sz", "w:szCs"):
        e = rpr.find(qn(tag))
        if e is None:
            e = rpr.makeelement(qn(tag), {qn("w:val"): val})
            rpr.append(e)
        else:
            e.set(qn("w:val"), val)


def set_font(run_el):
    """run に游ゴシックを明示する。"""
    rpr = run_el.find(qn("w:rPr"))
    if rpr is None:
        rpr = run_el.makeelement(qn("w:rPr"), {})
        run_el.insert(0, rpr)
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = rpr.makeelement(qn("w:rFonts"), {})
        rpr.insert(0, rf)
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(a), GOTHIC)


def role_of(text):
    """段落の文言から用途を判じる。"""
    t = text.strip()
    if not t:
        return None
    if RE_SHO.match(t):
        return "章"
    if RE_SETSU.match(t):
        return "節"
    if t.startswith("▌") or t.startswith("　▌"):
        return "小見出し"
    if RE_CAP.match(t):
        return "図"
    if t.startswith("出典：") or t.startswith("出典:") or \
            t.startswith("資料：") or t.startswith("資料:"):
        return "出典"
    if t.startswith("●") or t.startswith("⚠") or t.startswith("※"):
        return "本文"
    return "本文"


ROLE_SZ = {"章": SZ_SHO, "節": SZ_SETSU, "小見出し": SZ_SUB,
           "図": SZ_CAP, "出典": SZ_SHUTTEN, "本文": SZ_HONBUN}

HEADER_NEW = "川崎町高齢者保健福祉計画・第10期介護保険事業計画　計画書素案 Ver.2.5"


def fix_styles(path):
    """styles.xml に Normal を加え、全スタイルに游ゴシックを明示する。"""
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, \
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/styles.xml":
                s = data.decode("utf-8")
                if 'w:styleId="Normal"' not in s:
                    s = s.replace("</w:docDefaults>",
                                  "</w:docDefaults>" + NORMAL_STYLE, 1)
                # 既定の行送り（大雪広域と同じ 1.15 行・段落後 200）
                s = s.replace("<w:pPrDefault/>",
                              '<w:pPrDefault><w:pPr>'
                              '<w:spacing w:after="120" w:line="276"'
                              ' w:lineRule="auto"/>'
                              '</w:pPr></w:pPrDefault>', 1)
                # 各スタイルに游ゴシックと大きさを明示する
                def repl(m):
                    body = m.group(0)
                    sid = re.search(r'w:styleId="([^"]+)"', body).group(1)
                    if sid == "Normal":
                        return body
                    rpr = re.search(r"<w:rPr>(.*?)</w:rPr>", body, re.S)
                    fonts = (f'<w:rFonts w:ascii="{GOTHIC}" w:hAnsi="{GOTHIC}"'
                             f' w:eastAsia="{GOTHIC}" w:cs="{GOTHIC}"/>')
                    inner = rpr.group(1) if rpr else ""
                    if "<w:rFonts" not in inner:
                        inner = fonts + inner
                    if sid in STYLE_SZ:
                        v = str(STYLE_SZ[sid])
                        inner = re.sub(r'<w:sz w:val="\d+"/>',
                                       f'<w:sz w:val="{v}"/>', inner)
                        inner = re.sub(r'<w:szCs w:val="\d+"/>',
                                       f'<w:szCs w:val="{v}"/>', inner)
                        if "<w:sz " not in inner:
                            inner += f'<w:sz w:val="{v}"/><w:szCs w:val="{v}"/>'
                        if "<w:b/>" not in inner:
                            inner = "<w:b/><w:bCs/>" + inner
                    if sid in STYLE_COLOR:
                        c = STYLE_COLOR[sid]
                        if "<w:color " in inner:
                            inner = re.sub(r'<w:color w:val="[^"]*"/>',
                                           f'<w:color w:val="{c}"/>', inner)
                        else:
                            inner += f'<w:color w:val="{c}"/>'
                    if rpr:
                        return body[:rpr.start()] + f"<w:rPr>{inner}</w:rPr>" \
                            + body[rpr.end():]
                    return body.replace("</w:style>",
                                        f"<w:rPr>{inner}</w:rPr></w:style>")
                s = re.sub(r"<w:style [^>]*>.*?</w:style>", repl, s, flags=re.S)
                data = s.encode("utf-8")
            elif item.filename == "word/header1.xml":
                s = data.decode("utf-8")
                s = re.sub(r"計画書素案\s*v?V?er?\.?[0-9.]+", "", s)
                # 本文の <w:t> を書き換える
                s = re.sub(r"(<w:t[^>]*>)[^<]*(</w:t>)",
                           lambda m: m.group(1) + HEADER_NEW + m.group(2), s,
                           count=1)
                data = s.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, path)


# ══════════════════════════ 版表記
VER = {
    4: "計画書素案 Ver.2.5",
    5: "（書体統一版）",
    21: ("⚠ 頁番号は、本素案Ver.2.5の組版によるものです。"
         "図表の差替え・加筆により頁の割付けが動いた場合は、"
         "目次の頁番号も併せて更新します。"),
}
VER_SWAP = [("本素案Ver.2.4", "本素案Ver.2.5")]


def main():
    shutil.copy(SRC, DST)
    doc = docx.Document(DST)
    body = doc.element.body
    sys.path.insert(0, "07_ソーススクリプト")
    from fix_soan_v111 import set_el
    ps = [el for el in body.iterchildren() if el.tag == qn("w:p")]
    for i, txt in VER.items():
        set_el(ps[i - 1], txt)
    for el in body.iterchildren():
        if el.tag != qn("w:p"):
            continue
        t = "".join(n.text or "" for n in el.iter(qn("w:t")))
        new = t
        for a, b in VER_SWAP:
            new = new.replace(a, b)
        if new != t:
            set_el(el, new)

    # 目次の表の位置を求める（これより前は表紙・ご挨拶・目次であり、
    # 体裁上の意図があるため大きさを変えない）
    toc_idx = None
    kids = list(body.iterchildren())
    for i, el in enumerate(kids):
        if el.tag == qn("w:tbl"):
            t = docx.table.Table(el, doc)
            txt = "".join(c.text for r in t.rows[:3] for c in r.cells)
            if "ご挨拶" in txt or "計画策定の背景と目的" in txt:
                toc_idx = i
                break
    if toc_idx is None:
        raise SystemExit("目次の表が見つからない")
    toc_tbl = kids[toc_idx]

    # ══════════ Ver.2.2以降に加えた本文の太字を落とす
    #   fix_soan_v220〜v240 は本文の雛形として ps[301-1] を deepcopy していたが、
    #   これは「●要点」の段落であり太字である。
    #   そのため1-5・1-6・6-1・9-1・9-3・10-4に加えた本文が太字になっていた。
    #   ●要点・章扉のサブタイトル・▌小見出し・章・節は太字のまま残す。
    n_bold = 0
    prev_sho = False
    for i, el in enumerate(kids):
        if el.tag != qn("w:p"):
            continue
        t = "".join(n.text or "" for n in el.iter(qn("w:t"))).strip()
        is_sho = bool(RE_SHO.match(t))
        if i > toc_idx and t and not is_sho and not prev_sho:
            keep = (t.startswith("●") or t.startswith("▌")
                    or t.startswith("　▌") or RE_SETSU.match(t))
            if not keep:
                for r in el.iter(qn("w:r")):
                    rpr = r.find(qn("w:rPr"))
                    if rpr is None:
                        continue
                    for tag in ("w:b", "w:bCs"):
                        e = rpr.find(qn(tag))
                        if e is not None:
                            rpr.remove(e)
                            n_bold += 1
        prev_sho = is_sho if t else prev_sho

    # ══════════ Markdown の強調記号が本文に残っていたものを落とす
    n_ast = 0
    for el in body.iterchildren():
        if el.tag != qn("w:p"):
            continue
        t = "".join(n.text or "" for n in el.iter(qn("w:t")))
        if "**" in t:
            set_el(el, t.replace("**", ""))
            n_ast += 1

    n_p = n_t = n_font = 0
    changed = {}
    for i, el in enumerate(kids):
        if el.tag == qn("w:p"):
            for r in el.iter(qn("w:r")):
                set_font(r)
                n_font += 1
            if i <= toc_idx:
                continue
            txt = "".join(n.text or "" for n in el.iter(qn("w:t")))
            role = role_of(txt)
            if role is None:
                continue
            for r in el.iter(qn("w:r")):
                set_sz(r, ROLE_SZ[role])
            changed[role] = changed.get(role, 0) + 1
            n_p += 1
        elif el.tag == qn("w:tbl"):
            is_toc = el is toc_tbl
            for r in el.iter(qn("w:r")):
                set_font(r)
                n_font += 1
                if is_toc:
                    continue
                rpr = r.find(qn("w:rPr"))
                cur = None
                if rpr is not None:
                    e = rpr.find(qn("w:sz"))
                    if e is not None:
                        cur = int(e.get(qn("w:val"))) / 2
                set_sz(r, SZ_TABLE_SUB if cur and cur >= 11.0 else SZ_TABLE)
                n_t += 1

    doc.save(DST)
    fix_styles(DST)

    print("保存：", DST)
    print(f"  書体を游ゴシックに明示した run {n_font:,}")
    print(f"  段落の大きさを揃えた {n_p:,}（{changed}）")
    print(f"  表内の run の大きさを揃えた {n_t:,}")
    print(f"  本文の太字を落とした run {n_bold:,}")
    print(f"  本文に残っていた強調記号（**）を落とした段落 {n_ast}")

    # ══════════ 検算
    d2 = docx.Document(DST)

    def sz(r):
        rpr = r.find(qn("w:rPr"))
        if rpr is None:
            return None
        e = rpr.find(qn("w:sz"))
        return int(e.get(qn("w:val"))) / 2 if e is not None else None

    import collections
    by_role = collections.defaultdict(set)
    for p in d2.paragraphs:
        t = p.text.strip()
        if not t:
            continue
        role = role_of(t)
        for r in p._p.iter(qn("w:r")):
            by_role[role].add(sz(r))
    print("  用途ごとの大きさ（1つに揃っていること）：")
    for k in ("章", "節", "小見出し", "本文", "図", "出典"):
        v = sorted(x for x in by_role.get(k, set()) if x)
        mark = "○" if len(v) <= 1 else "×"
        print(f"    {mark} {k}：{v}")
    tb = collections.Counter()
    for t in d2.tables:
        for r in t._tbl.iter(qn("w:r")):
            tb[sz(r)] += 1
    print("  表内：", dict(tb.most_common(5)))
    nofont = sum(1 for r in d2.element.body.iter(qn("w:r"))
                 if r.find(qn("w:rPr")) is None
                 or r.find(qn("w:rPr")).find(qn("w:rFonts")) is None)
    print(f"  書体の指定がない run：{nofont}")
    with zipfile.ZipFile(DST) as z:
        st = z.read("word/styles.xml").decode("utf-8")
        hd = z.read("word/header1.xml").decode("utf-8")
    print("  既定の段落スタイル Normal：",
          "あり" if 'w:styleId="Normal"' in st else "なし")
    print("  游ゴシックの指定がないスタイル：",
          [m for m in re.findall(r'w:styleId="([^"]+)"', st)
           if f'w:styleId="{m}"' in st and
           GOTHIC not in re.search(
               r'<w:style [^>]*w:styleId="%s"[^>]*>.*?</w:style>' % re.escape(m),
               st, re.S).group(0)])
    print("  ヘッダー：", re.sub(r"<[^>]+>", "", hd).strip())
    txt = "\n".join(p.text for p in d2.paragraphs)
    print("  旧版の表記 Ver.2.4 の残り：", txt.count("Ver.2.4"))


if __name__ == "__main__":
    main()
