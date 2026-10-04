# -*- coding: utf-8 -*-
"""送付区分の成果品に、受託者の内部の仕組み・作業経過の語がないかを走査する.

CLAUDE.md §1（令和8年9月24日 ご指示）
  「発注者へ提供する文章に、受託者の内部の仕組み・作業経過を書かない」

§4 の「ご指示は同じ区分の成果品すべてに当てる。1冊だけ直して終わりにしない」
に対する手立てである。成果品ごとの自己点検とは別に、
`data_dispatch.DISPATCH` の区分が「送付」である成果品を横断して走査する。

走査する範囲
  docx  本文・表（入れ子を含む）・ヘッダー・フッター
  xlsx  全シートの全セルの文字列

判定
  不適合   内部の仕組み・作業経過の語に当たったもの
  要判断   成果品の名に含まれるため許容している語に当たったもの
  未実施   成果品の実体が無く走査できなかったもの

不適合が1件でもあれば終了コード1で終わる。

使い方
  python3 tools/sofu_word_check.py              送付区分の全件を走査する
  python3 tools/sofu_word_check.py <名の一部>    名が当たるものだけを走査する
"""

import os
import re
import sys

import openpyxl
import docx
from docx.oxml.ns import qn

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import repo_paths as RP                                       # noqa: E402
import data_dispatch as DD                                    # noqa: E402

# ------------------------------------------------------------ 走査する語
# 正規表現で探すもの（スクリプト名・モジュール名・ファイル名）
RE_NG = [
    (r"[A-Za-z0-9_]+\.py", "スクリプト名・モジュール名"),
    # 「個票がなくても再現できます」は当方の都合である（§1）。
    # 算定式の検証としての「再現できる」は差し支えないため、
    # 個票・個人情報が近くにあるものだけを不適合とする。
    (r"個票[^。]{0,40}再現でき|再現でき[^。]{0,40}個票", "個票の再現の可否"),
    (r"\brunpy\b", "読み込みの仕組み"),
    (r"\bast\b", "読み取りの仕組み"),
    (r"終了コード\s*[01]", "実行の仕組み"),
]
# 語で探すもの
WORD_NG = [
    ("スクリプト", "作り方の語"),
    ("モジュール", "作り方の語"),
    ("リポジトリ", "当方の作業環境の語"),
    ("固定値", "作り方の語"),
    ("実物から", "作り方の語"),
    ("直書き", "作り方の語"),
    ("書き写", "作り方の語"),
    ("再実行", "作業経過の語"),
    ("内部管理用", "送付区分と食い違う語"),
    ("読んで判定", "判定方法の説明"),
    ("章節ごとに", "判定方法の説明"),
    ("記述が薄", "内部の課題整理"),
    ("節が埋ま", "内部の課題整理"),
    ("1節が埋", "内部の課題整理"),
]
# 成果品の名に含まれるため許容する語（§1「成果品の名までは書いてよい」）
# 「内部管理用の一覧」は発注者の側の資料の作り方を問うものであり、
# 受託者の内部の仕組みではない。
ALLOW = ["トークスクリプト", "内部管理用の一覧"]


def _strip_allow(t):
    for a in ALLOW:
        t = t.replace(a, "　" * len(a))
    return t


def doc_texts(path):
    d = docx.Document(path)
    out = []
    for el in d.element.body.iter(qn("w:p")):
        out.append(("本文", "".join(t.text or "" for t in el.iter(qn("w:t")))))
    for i, s in enumerate(d.sections, start=1):
        for nm, part in (("ヘッダー", s.header), ("フッター", s.footer)):
            for p in part.paragraphs:
                out.append(("%s（節%d）" % (nm, i), p.text))
    return out


def xlsx_texts(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = []
    for ws in wb.worksheets:
        for r in ws.iter_rows(values_only=True):
            for v in r:
                if isinstance(v, str) and v.strip():
                    out.append((ws.title, v))
    wb.close()
    return out


def odt_texts(path):
    """odt は zip の content.xml からタグを落として読む。"""
    import zipfile
    with zipfile.ZipFile(path) as z:
        x = z.read("content.xml").decode("utf-8", "replace")
    x = re.sub(r"<text:p[^>]*>", "\n", x)
    return [("本文", t) for t in re.sub(r"<[^>]+>", "", x).split("\n") if t.strip()]


def scan(path):
    """(場所, 語, 理由, 前後) の一覧を返す。"""
    if path.endswith(".docx"):
        ts = doc_texts(path)
    elif path.endswith(".xlsx"):
        ts = xlsx_texts(path)
    elif path.endswith(".odt"):
        ts = odt_texts(path)
    else:
        return None
    hits = []
    for where, raw in ts:
        t = _strip_allow(raw)
        for pat, why in RE_NG:
            for m in re.finditer(pat, t):
                hits.append((where, m.group(0), why,
                             raw[max(0, m.start() - 30):m.end() + 30]))
        for w, why in WORD_NG:
            i = t.find(w)
            if i >= 0:
                hits.append((where, w, why,
                             raw[max(0, i - 30):i + len(w) + 30]))
    return hits


def main(argv):
    q = argv[1] if len(argv) > 1 else ""
    ng, mi, ok = [], [], 0
    for name in sorted(DD.DISPATCH):
        if DD.DISPATCH[name][0] != "送付":
            continue
        if q and q not in name:
            continue
        p = os.path.join(RP.OUTPUT, name)
        if not os.path.exists(p):
            mi.append((name, "実体が無い"))
            continue
        try:
            hits = scan(p)
        except Exception as e:                       # noqa: BLE001
            mi.append((name, "読めない（%s）" % e))
            continue
        if hits is None:
            # PDF は同名の xlsx から変換したものであるため、
            # その xlsx を走査していれば足りる。
            x = os.path.join(RP.OUTPUT, os.path.splitext(name)[0] + ".xlsx")
            if name.endswith(".pdf") and os.path.exists(x):
                ok += 1
            else:
                mi.append((name, "走査の対象としていない形式"))
            continue
        if hits:
            ng.append((name, hits))
        else:
            ok += 1

    print("■ 送付区分の成果品の走査")
    print("  適合 %d件／不適合 %d件／未実施 %d件" % (ok, len(ng), len(mi)))
    for name, hits in ng:
        print("")
        print("【不適合】%s（%d件）" % (name, len(hits)))
        seen = set()
        for where, w, why, ctx in hits:
            k = (where, w)
            if k in seen:
                continue
            seen.add(k)
            print("  %s｜%s（%s）" % (where, w, why))
            print("    …%s…" % ctx.replace("\n", " "))
    for name, why in mi:
        print("")
        print("【未実施】%s　%s" % (name, why))
    if mi:
        print("")
        print("※ 未実施は「適合」ではない。実体を作成してから走査する。")
    return 1 if ng else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
