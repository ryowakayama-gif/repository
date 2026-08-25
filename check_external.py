# -*- coding: utf-8 -*-
"""成果品を外部送付の観点で機械的に点検する.

発注者へ送付する資料に残っていると不適切な表現を検出する。
  ① 内部品質管理用語（レッドチーム、的中 ほか）
  ② 稿番号（第○稿）
  ③ 記述ルールで禁止した5表現
  ④ 自己評価的な文言
  ⑤ 記号（★）
  ⑥ 文書プロパティ（本文検索では見つからない題名・作成者・作成ソフト）
  ⑦ 内部保管資料への参照（送付資料だけで参照が閉じているか）

使い方
  python3 check_external.py            一覧を表示する
  python3 check_external.py -v         該当箇所の文面も表示する
"""

import glob
import os
import re
import sys
from collections import Counter, defaultdict

ODIR = "/home/user/repository/output"

GROUPS = [
    # 「撤回」「誤読」は、記述を改めた事実の説明として必要なため対象外とする。
    ("内部品質管理", [
        "レッドチーム", "レッド チーム", "的中",
        "自らの整理を否定", "内部批判",
    ]),
    ("稿番号", [r"第\d+稿", r"第\d+稿→第\d+稿"]),
    ("禁止表現", [
        "全国トップ級", "に由来する", r"と整合する(?!か)", "1件も",
        "有意差がないため",
    ]),
    ("自己評価", [
        "所定の到達点", "本報告の中心", "最大の特徴", "最も大きな特徴",
        "明らかになった", "完璧", "万全",
    ]),
    # ★は、評価尺度（★★★）としての使用は差し支えない。
    # 文の頭に付けて強調する用法のみを対象とする。
    ("記号", [r"★(?=[^★\s])"]),
]

# 検出しても差し支えない文脈（用語を禁止・訂正するために引用している等）
ALLOW = [
    "等の表現を用いない", "は用いない", "を用いない（",
    "記述ルール", "禁止", "不可", "と判定していた", "当初は",
    "改めた", "改めます", "に置き換え", "としない",
    "修正指示書", "内部保管", "用いない表現", "代わりに用いる",
    "訂正", "という所見", "が成り立たない", "従来", "旧", "と書かない",
]

# 用語そのものを定義・列挙するためのファイル（点検の対象外）
EXEMPT_FILES = ["第10期計画_成果品の送付区分表.xlsx"]


def is_ok(text, hit_pos):
    """前後の文脈から、禁止語を「使っている」のではなく「扱っている」場合を除く。"""
    ctx = text[max(0, hit_pos - 60):hit_pos + 80]
    return any(a in ctx for a in ALLOW)


def scan_docx(path):
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    d = Document(path)
    out = []
    for ch in d.element.body.iterchildren():
        if ch.tag.endswith("}p"):
            out.append(Paragraph(ch, d).text)
        elif ch.tag.endswith("}tbl"):
            for r in Table(ch, d).rows:
                for c in r.cells:
                    out.append(c.text)
    return [("", t) for t in out if t]


def scan_xlsx(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = []
    for ws in wb:
        for row in ws.iter_rows(values_only=True):
            for v in row:
                if isinstance(v, str) and v:
                    out.append((ws.title, v))
    return out


def scan_odt(path):
    import zipfile
    from xml.etree import ElementTree as ET
    x = ET.fromstring(zipfile.ZipFile(path).read("content.xml"))
    return [("", "".join(x.itertext()))]


def scan(path):
    e = os.path.splitext(path)[1].lower()
    if e == ".docx":
        return scan_docx(path)
    if e == ".xlsx":
        return scan_xlsx(path)
    if e == ".odt":
        return scan_odt(path)
    return []


def run(verbose=False, only_sofu=True):
    from data_dispatch import dispatch_of
    files = sorted(glob.glob(os.path.join(ODIR, "*.xlsx")) +
                   glob.glob(os.path.join(ODIR, "*.docx")) +
                   glob.glob(os.path.join(ODIR, "*.odt")))
    skipped = [f for f in files
               if dispatch_of(os.path.basename(f))[0] in ("内部保管", "対象外")]
    if only_sofu:
        files = [f for f in files if f not in skipped]
        print("送付区分「内部保管」「対象外」の %d 件は対象から除いた"
              "（-a で含める）" % len(skipped))
    total = Counter()
    perfile = defaultdict(Counter)
    detail = defaultdict(list)
    for p in files:
        fn = os.path.basename(p)
        if fn in EXEMPT_FILES:
            continue
        try:
            cells = scan(p)
        except Exception as ex:                       # noqa: BLE001
            print("  読めない:", fn, ex)
            continue
        for sh, text in cells:
            for gname, pats in GROUPS:
                for pat in pats:
                    for m in re.finditer(pat, text):
                        if is_ok(text, m.start()):
                            continue
                        total[gname] += 1
                        perfile[fn][gname] += 1
                        if len(detail[(fn, gname)]) < 3:
                            i = m.start()
                            detail[(fn, gname)].append(
                                "[%s] …%s…" % (sh or "本文",
                                               text[max(0, i - 45):i + 55]
                                               .replace("\n", " ")))
    print("=" * 78)
    print("外部送付の観点による点検　対象 %d ファイル" % len(files))
    print("=" * 78)
    for fn in sorted(perfile, key=lambda x: -sum(perfile[x].values())):
        c = perfile[fn]
        print("\n%-56s 計%3d" % (fn[:56], sum(c.values())))
        for g, n in c.most_common():
            print("    %-12s %3d" % (g, n))
            if verbose:
                for d in detail[(fn, g)]:
                    print("        " + d)
    # ------------------------------------------------ 文書プロパティ
    from docmeta import scan_meta
    meta_ng = []
    for p in files:
        h = scan_meta(p)
        if h:
            meta_ng.append((os.path.basename(p), h))

    # ------------------------------------------------ 内部保管資料への参照
    naibu = [os.path.splitext(f)[0]
             .replace("第10期計画_", "").replace("第10期計画素案_", "")
             for f, (k, _w) in dispatch_of.__globals__["DISPATCH"].items()
             if k == "内部保管"]
    ref_ng = []
    for p in files:
        fn = os.path.basename(p)
        if fn in EXEMPT_FILES:
            continue
        try:
            cells = scan(p)
        except Exception:                         # noqa: BLE001
            continue
        got = sorted({k for _sh, t in cells for k in naibu if k and k in t})
        if got:
            ref_ng.append((fn, got))

    print("\n" + "-" * 78)
    print("文書プロパティの点検: %s"
          % ("該当なし" if not meta_ng
             else "／".join("%s（%s）" % (f, "・".join(h)) for f, h in meta_ng)))
    print("内部保管資料への参照: %s"
          % ("該当なし" if not ref_ng
             else "／".join("%s→%s" % (f, "・".join(g)) for f, g in ref_ng)))
    print("-" * 78)
    print("区分別の合計:",
          "／".join("%s %d" % (k, v) for k, v in total.most_common()))
    print("該当ファイル数: %d / %d" % (len(perfile), len(files)))
    return perfile


if __name__ == "__main__":
    run(verbose="-v" in sys.argv, only_sofu="-a" not in sys.argv)
