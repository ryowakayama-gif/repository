"""成果品の紙面（頁の上での見え方）を機械で点検する。

`check_ono_taisai.py` は docx の中身（用紙・余白・書体・列幅・塗り・図表番号）を
見るもので、**頁の切れ目に何が来るかは見ていない。** 本スクリプトはそこを見る。

  ・枠からのはみ出し　　列幅の取り違えで字が右の余白に出ていないか
  ・見出しの取り残し　　見出しが頁の最後に来て、中身が次の頁にないか
  ・図とその題の離れ　　【図N】と図・出典が別の頁に分かれていないか
  ・小さすぎる字　　　　8pt より小さい字がないか

**表の孤立行は検出していない。** 続きの頁に最後の1行だけが残る欠陥である。
当方の環境では確実に見分けられなかった。試したのは次の2通りで、いずれも誤った。

  ① 行の文字列で頁を突き合わせる
     同じ様式の表を並べた文書では、見出し行や値の行が表どうしで同じ文字列に
     なる。成果指標の仕様の文書は「区分／内容」の同じ表が48個並び、
     **別の表の行に当たって3件を誤検出した。**
  ② 頁の最初の中身の行が表の最後の行かどうかを見る
     語を行にまとめる際の丸めが粗く、1行の中の語が分かれることがある。
     長い表では LibreOffice が見出し行を続きの頁の頭で繰り返すため、
     「頁の最初の行」が見出し行になる場合もある。

**代わりに、起きないようにしてある。** `ono_style.tbl` が、見出し行を含めて
12行までの表について、最後から2番目の行に `keepNext` を掛け、最後の行と
一緒に次の頁へ送る。これより長い表は、見出し行が続きの頁の頭で繰り返される
ため、続きが1行でも読める。**検出ではなく予防で担保している。**

やり方は2段である。

  1. `soffice` で docx を PDF に変換し、`pdftotext -bbox-layout` で
     1語ごとの位置を得る。はみ出しと頁数はここから読む。
  2. docx の段落・表の行の文字列が PDF の何頁に載ったかを、
     空白を除いた文字列の一致で突き合わせる。
     孤立行・見出しの取り残し・図と題の離れはここから読む。

**この確認でできないこと。** 游ゴシックがこの環境になく IPAゴシックに
置き換えて組まれるため、行の折り返しの位置と頁数は町の Word と異なり得る。
一方、**表の割れ・図の潰れ・見出しの取り残し・枠からのはみ出しは
書体によらず見える**ため、本点検はこれらについては有効である。

**頁数について分かっていないこと（令和8年10月14日）。**
`ono_style` に Word の見出しスタイルを当てた前後で、計画素案の頁数が
LibreOffice の組版で **62頁から70頁に増えた**。両者の docx の差は
`<w:pStyle w:val="HeadingN"/>` が76か所増えただけで、本文・表・図は
1文字も変わっていない（タグ単位で照合して確認した）。
増える原因は次のいずれでもないことを、1つずつ外して組み直して確かめた。

  スタイルの字の大きさ／行間／段落の前後の空き／書体／keepLines／
  ぶら下げの設定／outlineLvl／basedOn／言語の指定／スタイル名

**LibreOffice が、docx のスタイル定義より自分の組み込みの見出しの体裁を
優先しているためと見ているが、特定できていない。** 見出しの大きさ・色・
太字・段落の前後の空きは**段落と run に直に書いてあり**、Word はそちらを
優先するため、Word の紙面は従前と変わらないと考えられる。
ただし**この環境では確かめられない。** 町の Word で開いた頁数を
いただければ、どちらであるかが決まる（確認事項）。
本点検が見る「頁の切れ目に何が来るか」は、どちらであっても有効である。

  python3 check_ono_shimen.py              TARGETS をすべて見る
  python3 check_ono_shimen.py <docx> ...   指定した docx だけを見る
  python3 check_ono_shimen.py --jiko       自己試験（検査が現に検出することを確かめる）
"""

import html
import pathlib
import re
import shutil
import subprocess
import sys
import tempfile
import warnings

warnings.filterwarnings("ignore")

from docx import Document

ROOT = pathlib.Path(__file__).parent / "小野町_引継ぎ_整理済"

# 頁数の枠があるもの・表の多いものを見る。
# （`check_ono_taisai.TARGETS` のうち、紙面の割れが起きうる分量のもの）
TARGETS = [
    "03_計画素案/小野町高齢者保健福祉計画_第10期介護保険事業計画_"
    "素案_第12版_20261020.docx",
    "09_骨子案/小野町高齢者保健福祉計画_第10期介護保険事業計画_"
    "骨子案_第4版_20261006.docx",
    "09_骨子案/おのまち障がい者計画_第4期障がい児福祉計画_"
    "第8期障がい福祉計画_骨子案_第3版_20261015.docx",
    "08_協議会資料/小野町高齢者福祉サービス推進協議会資料_第2回_"
    "骨子案と見込量・保険料_20261006.docx",
    "04_算定・見込量/小野町_第10期_サービス見込量推計結果報告書_20261006.docx",
    "10_納品・印刷/小野町_第10期計画_印刷の発注仕様書_20261006.docx",
    "04_算定・見込量/小野町_令和8年法改正の反映と3区分の仕分け_20261015.docx",
    "03_計画素案/小野町_第10期_成果指標の仕様_20261014.docx",
    "16_町データ依頼/小野町_必要資料の優先順位_20261015.docx",
    "04_算定・見込量/小野町_決算・交付金の受領と反映方針_20261015.docx",
    "16_町データ依頼/小野町_高齢者_必要資料の再確認_20261015.docx",
    "04_算定・見込量/小野町_計画の整合のエビデンスの点検_20261015.docx",
    "04_算定・見込量/小野町_県高齢者居住安定確保計画の反映_20261015.docx",
    "04_算定・見込量/小野町_県保健医療計画の反映_20261015.docx",
    "04_算定・見込量/小野町_障がい計画_県計画との整合の点検_20261015.docx",
    "16_町データ依頼/小野町_確認事項の一覧_20261015.docx",
    "16_町データ依頼/小野町_町データ待ちの切り分け_20261019.docx",
    "03_計画素案/小野町_法定記載事項チェック_20261015.docx",
    "03_計画素案/おのまち障がい者計画_第4期障がい児福祉計画_"
    "第8期障がい福祉計画_素案_第1版_20261015.docx",
    "03_計画素案/小野町_障がい計画_他素案への追加候補の対応表_20261020.docx",
    "03_計画素案/小野町_障がい計画_論点漏れの点検_20261020.docx",
    "12_進捗管理/小野町_第10期_策定チェックリストの確認結果_20261020.docx",
    "03_計画素案/小野町_審議会論点の整理_第271回介護給付費分科会・第160回障害者部会_20261020.docx",
]

A4W_PT = 595.276          # A4の幅（pt）＝21.0cm
MARGIN_CM = 1.9           # 左右の余白
PT_PER_CM = 72 / 2.54
RIGHT_EDGE = A4W_PT - MARGIN_CM * PT_PER_CM     # 本文の右端（pt）
OVER_TOL = 2.0            # はみ出しの許容（pt）。罫線の太さと丸めの分
MIN_PT = 8.0              # これより小さい字は読めない
SHORT_TBL = 7             # 「短い表」とみなす行数（見出し行を含む）

# 行末でぶら下がってよい文字（日本語の組み方として正しい）
BURASAGE = "。、，．）」』】〕〉》”’!?！？:;：；"


def _norm(s):
    """突き合わせ用に、空白と制御文字を落とす。"""
    return re.sub(r"\s+", "", str(s))


def to_pdf(src, workdir):
    """docx を PDF に変換する。"""
    dst = pathlib.Path(workdir) / (pathlib.Path(src).stem + ".pdf")
    subprocess.run(
        ["soffice", "--headless",
         "-env:UserInstallation=file://%s/lo" % workdir,
         "--convert-to", "pdf", "--outdir", str(workdir), str(src)],
        check=True, capture_output=True, timeout=600,
        env={"HOME": str(workdir), "PATH": "/usr/bin:/bin"})
    if not dst.exists():
        raise RuntimeError("PDFに変換できなかった")
    return dst


def pdf_pages_text(pdf):
    """PDFの各頁の文字列（空白を除いたもの）を返す。"""
    txt = subprocess.run(["pdftotext", "-layout", str(pdf), "-"],
                         check=True, capture_output=True, timeout=300
                         ).stdout.decode("utf-8", "ignore")
    return [_norm(p) for p in txt.split("\f")]


def pdf_word_boxes(pdf):
    """1語ごとの (頁, x左, x右, y上, y下, 文字列) を返す。"""
    xh = subprocess.run(["pdftotext", "-bbox-layout", str(pdf), "-"],
                        check=True, capture_output=True, timeout=300
                        ).stdout.decode("utf-8", "ignore")
    out, page = [], 0
    for m in re.finditer(
            r'<page width="([\d.]+)"|'
            r'<word xMin="([\d.]+)" yMin="([\d.]+)" '
            r'xMax="([\d.]+)" yMax="([\d.]+)">(.*?)</word>', xh):
        if m.group(1):
            page += 1
            continue
        out.append((page, float(m.group(2)), float(m.group(4)),
                    float(m.group(3)), float(m.group(5)),
                    html.unescape(m.group(6))))
    return out


def find_page(needle, pages, cur, minlen=4, span=None):
    """文字列が載った頁を、cur から前へ探す。見つからなければ None。

    minlen … これより短い文字列は、取り違えるため探さない。
             表の行は「行5」のように短いことがあるため、
             呼ぶ側が cur と span で範囲を狭めたうえで下げてよい。
    span   … cur から何頁先まで見るか。表のように連続している
             ものを探すときに、遠くの同じ文字列に当たらないようにする。
    """
    needle = _norm(needle)
    if len(needle) < minlen:
        return None
    key = needle[:40]
    last = len(pages) if span is None else min(len(pages), cur + span + 1)
    for i in range(cur, last):
        if key in pages[i]:
            return i
    return None


def table_start_page(rows, pages, cur):
    """表が始まる頁を求める。

    見出し行は「区分」「値」のように短く、単独では探せない。
    先頭2行をつないで探し、それで見つからなければ1行ずつ試す。
    """
    for cand, ml, sp in (("".join(rows[:2]), 3, None),
                         (rows[0], 2, 3),
                         (rows[1] if len(rows) > 1 else "", 2, 3)):
        pg = find_page(cand, pages, cur, minlen=ml, span=sp)
        if pg is not None:
            return pg
    return None




# ---------------------------------------------------------------- 各点検
def check_hamidashi(boxes):
    """枠からのはみ出し。行末のぶら下げ組は除く。"""
    bad = []
    for pg, x0, x1, _y0, _y1, w in boxes:
        if x1 <= RIGHT_EDGE + OVER_TOL:
            continue
        # 末尾のぶら下げてよい文字を落として測り直す
        core = w.rstrip(BURASAGE)
        if not core:
            continue
        shrink = (x1 - x0) * (len(w) - len(core)) / max(len(w), 1)
        if x1 - shrink > RIGHT_EDGE + OVER_TOL:
            bad.append("p%d 枠から %.1fpt はみ出し「%s」"
                       % (pg, x1 - shrink - RIGHT_EDGE, w[:20]))
    return bad


def check_chiisai(doc):
    """小さすぎる字。docx の run から直に読む。"""
    bad = []
    seen = set()
    for q in doc.paragraphs:
        for r in q.runs:
            if r.font.size and r.font.size.pt < MIN_PT:
                k = (r.font.size.pt, r.text[:12])
                if k not in seen:
                    seen.add(k)
                    bad.append("%.1fpt の字「%s」" % (r.font.size.pt, r.text[:20]))
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                for q in c.paragraphs:
                    for r in q.runs:
                        if r.font.size and r.font.size.pt < MIN_PT:
                            k = (r.font.size.pt, r.text[:12])
                            if k not in seen:
                                seen.add(k)
                                bad.append("%.1fpt の字（表）「%s」"
                                           % (r.font.size.pt, r.text[:20]))
    return bad


def _body_items(doc):
    """本文の段落と表を、現れる順に (種類, もの) で返す。"""
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    from docx.oxml.ns import qn as _qn
    out = []
    for ch in doc.element.body.iterchildren():
        if ch.tag == _qn("w:p"):
            out.append(("p", Paragraph(ch, doc)))
        elif ch.tag == _qn("w:tbl"):
            out.append(("t", Table(ch, doc)))
    return out


def body_start(pages):
    """本文が始まる頁を返す。

    目次の行は見出しと同じ文字列を持つため、頁を頭から探すと
    **本文の見出しが目次の行に当たってしまう**（自己試験で実際に
    誤検出した）。目次の頁を飛ばしてから本文を探す。
    目次の頁は、点リーダー（………）が何度も出ることで見分ける。
    """
    last = -1
    for i, t in enumerate(pages[:14]):      # 目次は前付にしかない
        if t.count("…") + t.count("....") >= 5:
            last = i
    return last + 1


def outline_pages(pdf):
    """PDFのしおりから (見出しの文字列, 頁) を順に返す。

    見出しの頁を**文字列の突き合わせで探さない**ための要。
    目次の行は見出しと同じ文字列を持つため突き合わせでは取り違える。
    しおりは本文の見出しだけを指すので、取り違えが起きない。
    （しおりが出るのは、見出しにWordの見出しスタイルを当てたことによる。）
    """
    try:
        from pypdf import PdfReader
    except ImportError:
        return None
    r = PdfReader(str(pdf))
    out = []

    def walk(node):
        for x in node:
            if isinstance(x, list):
                walk(x)
            else:
                try:
                    out.append((_norm(x.title), r.get_destination_page_number(x)))
                except Exception:
                    pass
    try:
        walk(r.outline)
    except Exception:
        return None
    return out


def page_lines(boxes):
    """語の箱を行にまとめ、頁ごとに (y上, y下, 行の文字列) を返す。

    同じ行に属する語は yMin がほぼ揃う。2pt の幅で丸めてまとめる。
    """
    tmp = {}
    for pg, x0, _x1, y0, y1, w in boxes:
        tmp.setdefault((pg, round(y0 / 2.0)), []).append((x0, y0, y1, w))
    out = {}
    for (pg, _k), ws in tmp.items():
        ws.sort()
        out.setdefault(pg, []).append(
            (min(w[1] for w in ws), max(w[2] for w in ws),
             _norm("".join(w[3] for w in ws))))
    for pg in out:
        out[pg].sort()
    return out


def heading_orphans(pdf, boxes, page_h=841.89, foot=60.0):
    """見出しの取り残し ― 見出しだけが頁の最後に残っていないか。

    しおりで見出しの頁を定め、その頁の語の位置から
    「見出しより下に本文が無い」ことで判定する。
    フッタ（頁番号）は下から foot pt を除いて数えない。
    """
    marks = outline_pages(pdf)
    if not marks:
        return []
    lines = page_lines(boxes)
    bad = []
    for title, pg0 in marks:
        rows = lines.get(pg0 + 1, [])           # boxes の頁は1起点
        if not rows or not title:
            continue
        key = title[:12]
        # 見出しの行を探す。**語ではなく行で照合する。**
        # 語で照合すると、題の一部に当たる短い語が頁の下の方にあるだけで
        # 「見出しは頁の最後」と誤って判定する（自己試験で実際に起きた）。
        hit = [y1 for _y0, y1, t in rows if key and key in t]
        if not hit:
            continue
        ybottom = min(hit)                      # 最初に現れた行を見出しとみなす
        below = [t for y0, _y1, t in rows
                 if y0 > ybottom + 1 and y0 < page_h - foot and t]
        if not below:
            bad.append("p%d 見出しだけが頁の最後に残っている「%s」"
                       % (pg0 + 1, title[:28]))
    return bad


def check_koritsu_midashi(doc, pages, lines):
    """表の孤立行・図とその題の離れ。"""
    bad = []
    cur = body_start(pages)
    items = _body_items(doc)
    # 各要素が載った頁を先に求める
    loc = []
    for kind, obj in items:
        if kind == "p":
            pg = find_page(obj.text, pages, cur)
        else:
            rows = ["".join(c.text for c in r.cells) for r in obj.rows]
            pg = table_start_page(rows, pages, cur) if rows else None
        if pg is not None:
            cur = pg
        loc.append(pg)

    for i, ((kind, obj), pg) in enumerate(zip(items, loc)):
        if pg is None:
            continue
        # （表の孤立行は検出しない。理由は冒頭の説明のとおり）
        # （見出しの取り残しは heading_orphans が PDF のしおりで見る）
        # --- 図とその題の離れ
        if kind == "p" and re.match(r"【図\d+】", obj.text):
            j = i + 1
            while j < len(items) and items[j][0] == "p" \
                    and not _norm(items[j][1].text) \
                    and "graphicData" not in items[j][1]._p.xml:
                j += 1
            if j < len(items) and loc[j] is not None and loc[j] > pg:
                bad.append("p%d 図の題と図が別の頁になっている「%s」"
                           % (pg + 1, obj.text[:28]))
    return bad


# ---------------------------------------------------------------- 実行
def check(path, workdir):
    doc = Document(path)
    pdf = to_pdf(path, workdir)
    pages = pdf_pages_text(pdf)
    boxes = pdf_word_boxes(pdf)
    bad = []
    bad += check_hamidashi(boxes)
    bad += check_chiisai(doc)
    bad += heading_orphans(pdf, boxes)
    bad += check_koritsu_midashi(doc, pages, page_lines(boxes))
    return bad, len(pages) - (1 if pages and not pages[-1] else 0)


def jiko_shiken():
    """自己試験。壊したものを検査が現に見つけることを確かめる。"""
    import ono_style as S
    ok = []
    with tempfile.TemporaryDirectory() as wd:
        # (1) 枠からのはみ出し ― 列幅の合計を本文幅より広くする
        r = S.Report("自己試験1")
        r.h1("見出し")
        r.tbl([["区分", "内容"], ["あ", "い" * 60]], widths=[3.0, 20.0])
        p = r.save(pathlib.Path(wd) / "s1.docx")
        b, _ = check(p, wd)
        ok.append(("枠からのはみ出し",
                   any("はみ出し" in x for x in b), b[:1]))

        # (2) 小さすぎる字
        r = S.Report("自己試験2")
        r.h1("見出し")
        r.p("小さい字", sz=6)
        p = r.save(pathlib.Path(wd) / "s2.docx")
        b, _ = check(p, wd)
        ok.append(("小さすぎる字",
                   any("の字" in x for x in b), b[:1]))

        # (3) 見出しの取り残し ― 見出しの直後に改頁を入れる。
        #     見出しには keep_with_next が掛かっているため、ふつうに
        #     組んでいる限りこの欠陥は起きない。起きるのは、1頁大の図版の
        #     ように改頁を自分で入れたときである（他案件で実際に起きた形）。
        r = S.Report("自己試験3")
        for i in range(5):
            r.p("本文の行 %d" % i)
        r.h2("頁の最後に来る見出し")
        r.pagebreak()
        r.p("中身は次の頁にある")
        p = r.save(pathlib.Path(wd) / "s3.docx")
        b, _ = check(p, wd)
        ok.append(("見出しの取り残し",
                   any("見出しだけ" in x for x in b), b[:1]))


    print("■ 自己試験（壊したものを検査が見つけるか）\n")
    rc = 0
    for name, hit, sample in ok:
        print("  %s  %s" % ("検出した" if hit else "**見落とした**", name))
        if sample:
            print("        %s" % sample[0])
        if not hit:
            rc = 1
    print("\n  %d件中 %d件を検出" % (len(ok), sum(1 for _, h, _ in ok if h)))
    return rc


def main(argv):
    if "--jiko" in argv:
        return jiko_shiken()
    rels = argv[1:] or TARGETS
    rc, n = 0, 0
    with tempfile.TemporaryDirectory() as wd:
        for rel in rels:
            p = pathlib.Path(rel)
            if not p.exists():
                p = ROOT / rel
            if not p.exists():
                print("** 見つからない %s" % rel)
                rc = 1
                continue
            n += 1
            try:
                bad, pg = check(p, wd)
            except Exception as e:
                print("NG  %s　%s" % (p.name, e))
                rc = 1
                continue
            head = "%s　%d頁" % (p.name, pg)
            if bad:
                rc = 1
                print("NG  %s" % head)
                for b in bad[:15]:
                    print("      %s" % b)
                if len(bad) > 15:
                    print("      ほか%d件" % (len(bad) - 15))
            else:
                print("OK  %s" % head)
    print("\n紙面の点検 %d件　%s" % (n, "欠陥なし" if rc == 0 else "要修正あり"))
    print("※ 游ゴシックがこの環境になくIPAゴシックで組まれている。"
          "行の折り返しと頁数は町のWordと異なり得る。")
    return rc


if __name__ == "__main__":
    if not shutil.which("soffice") or not shutil.which("pdftotext"):
        sys.exit("soffice と pdftotext が要る："
                 "apt-get install libreoffice-writer poppler-utils")
    sys.exit(main(sys.argv))
