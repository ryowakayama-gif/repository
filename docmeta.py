# -*- coding: utf-8 -*-
"""文書のプロパティ（メタデータ）を送付できる内容に整える.

本文を置換しても、DOCX の docProps/core.xml、ODT の meta.xml には
原本の題名・作成者・作成ソフト等が残る。これらは本文検索では見つからない。

  ・他団体名（原本の流用元）
  ・受託者内部の版の表記（第○稿）
  ・個人名（担当者名は収録しない運用による）
  ・作成に用いたソフトウェアの名称

使い方
  from docmeta import clean_docx, clean_odt, scan_meta
  clean_docx(path, title="…", subject="…")
  clean_odt(path, title="…", subject="…")
"""

import os
import re
import shutil
import zipfile

ORG = "大雪地区広域連合"
VENDOR = "ビズアップ公共コンサルティング株式会社"

# プロパティに残っていてはならない語
# （作成に用いたライブラリ名は業務上の秘匿情報ではないため対象としない）
NG_WORDS = ["小野町", "RED TEAM", "REDTEAM", "Red Team", "レッドチーム",
            "OpenAI", "ChatGPT", "Claude"]
NG_RE = re.compile(r"第\d+稿")


def _tag(xml, tag, value):
    """<ns:tag>…</ns:tag> の中身を差し替える。無ければ何もしない。"""
    pat = re.compile(r"(<[\w]+:%s(?:\s[^>]*)?>)(.*?)(</[\w]+:%s>)"
                     % (tag, tag), re.S)
    return pat.sub(lambda m: m.group(1) + value + m.group(3), xml)


def _drop(xml, tag):
    """<ns:tag>…</ns:tag> を丸ごと取り除く（空要素も含む）。"""
    xml = re.sub(r"<[\w]+:%s(?:\s[^>]*)?>.*?</[\w]+:%s>" % (tag, tag), "",
                 xml, flags=re.S)
    return re.sub(r"<[\w]+:%s(?:\s[^>]*)?/>" % tag, "", xml)


def _rewrite(path, edits):
    """ZIP 内の指定エントリを書き換えて保存し直す。他の内容は変えない。"""
    with zipfile.ZipFile(path) as z:
        order = z.namelist()
        data = {n: z.read(n) for n in order}
    for name, fn in edits.items():
        if name in data:
            data[name] = fn(data[name].decode("utf-8")).encode("utf-8")
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n in order:
            if n == "mimetype":                  # ODF は無圧縮で先頭に置く
                z.writestr(zipfile.ZipInfo(n), data[n],
                           compress_type=zipfile.ZIP_STORED)
            else:
                z.writestr(n, data[n])
    shutil.move(tmp, path)


def clean_docx(path, title, subject="", creator=None):
    """DOCX の文書プロパティを送付できる内容に置き換える。"""
    creator = creator or VENDOR

    def core(x):
        x = _tag(x, "title", title)
        x = _tag(x, "subject", subject)
        x = _tag(x, "creator", creator)
        x = _tag(x, "lastModifiedBy", creator)
        x = _tag(x, "description", "")
        x = _tag(x, "keywords", "")
        x = _tag(x, "category", "")
        return _drop(x, "lastPrinted")

    def app(x):
        x = _tag(x, "Company", ORG)
        x = _tag(x, "Manager", "")
        x = _tag(x, "Application", "")
        # TitlesOfParts（目次の見出しと原本の題名の一覧）にも稿番号が残る
        def lpstr(m):
            t = m.group(2)
            t = re.sub(r"第\d+稿・", "", t)
            t = NG_RE.sub("令和8年8月", t)
            for w in NG_WORDS:
                t = t.replace(w, "")
            return m.group(1) + t + m.group(3)
        return re.sub(r"(<vt:lpstr>)([^<]*)(</vt:lpstr>)", lpstr, x)

    _rewrite(path, {"docProps/core.xml": core, "docProps/app.xml": app})
    return path


def clean_xlsx(path, title, subject="", creator=None):
    """XLSX の文書プロパティを送付できる内容に置き換える。"""
    return clean_docx(path, title, subject, creator)


def clean_odt(path, title, subject="", creator=None):
    """ODT の meta.xml を送付できる内容に置き換える。"""
    creator = creator or VENDOR

    def meta(x):
        x = _tag(x, "title", title)
        x = _tag(x, "subject", subject)
        x = _tag(x, "description", "")
        x = _tag(x, "keyword", "")
        x = _tag(x, "creator", creator)
        x = _tag(x, "initial-creator", creator)
        x = _drop(x, "user-defined")
        # <meta:generator>…</meta:generator>
        return _tag(x, "generator", "")

    _rewrite(path, {"meta.xml": meta})
    return path


def scan_meta(path):
    """プロパティに残る不適切な語を返す。送付前の点検に用いる。"""
    try:
        z = zipfile.ZipFile(path)
    except Exception:                             # noqa: BLE001
        return []
    names = z.namelist()
    blob = ""
    for n in ("docProps/core.xml", "docProps/app.xml", "meta.xml"):
        if n in names:
            blob += z.read(n).decode("utf-8", "replace")
    hits = [w for w in NG_WORDS if w in blob]
    hits += sorted(set(NG_RE.findall(blob)))
    return hits


if __name__ == "__main__":
    import glob
    bad = 0
    for p in sorted(glob.glob("/home/user/repository/output/*.docx") +
                    glob.glob("/home/user/repository/output/*.xlsx") +
                    glob.glob("/home/user/repository/output/*.odt")):
        h = scan_meta(p)
        if h:
            bad += 1
            print("%-56s %s" % (os.path.basename(p)[:56], "／".join(h)))
    print("プロパティに残存: %d件" % bad)
