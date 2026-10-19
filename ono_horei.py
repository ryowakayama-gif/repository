"""法令の条文を e-Gov 法令API（v2）から取る。

なぜ要るか
  条文番号は記憶で書くと項・号を取り違える。現にこれまで2件の取り違えが
  あった（市町村計画に医療計画との整合性の確保の義務があるとしたもの、
  障がい2法の調和を第5項としたもの）。**引くたびに現物で確かめる。**

v1 ではなく v2 を使う理由
  v1（`/api/1/lawdata/<法令番号>`）は**現行条文しか返さない。**
  v2 は `asof=YYYY-MM-DD` を付けると、**その日に効力をもつ版**を返す。
  第10期の計画期間は令和9年4月1日に始まり、その日に

    ・介護保険法（第117条が14項→15項。第5項【新設】勘案義務ほか）
    ・社会福祉法（第106条の6が丸ごと差し替え。支援会議は第106条の3の2へ）
    ・老人福祉法・障害者総合支援法・児童福祉法

  の改正が施行される（社会福祉法等の一部を改正する法律＝令和8年法律
  第51号、公布 令和8年6月25日）。**計画期間に効力をもつ条文は改正後の
  ものであり、v2 でなければ原典で確かめられない。**

使い方
    import ono_horei as H
    root = H.get("介護保険法")              # 現行（今日）
    root = H.get("介護保険法", H.TENKI)     # 第10期の初日
    body = H.find(root, "117", "5", None)   # 第117条第5項の本文

取れないとき
  None を返す。**取れないことと、取って違っていたことを分ける。**
  ネットワークが遮断された環境では前者になる。
"""

import pathlib
import re
import subprocess
import urllib.parse
import xml.etree.ElementTree as ET

CACHE = pathlib.Path("/tmp/claude-0/-home-user-repository/"
                     "d7dca1d7-b918-515d-8b6d-9309b5e7ef32/scratchpad/egov2")
API = "https://laws.e-gov.go.jp/api/2/law_data/"

# 第10期介護保険事業計画・第8期障がい福祉計画の初日。
# この日に効力をもつ条文が、計画期間を通じて適用される条文である。
TENKI = "2027-04-01"

# 社会福祉法等の一部を改正する法律（令和8年法律第51号。公布 令和8年6月25日）。
# 本件に関わる規定の多くが令和9年4月1日に施行される。
KAISEI = "令和8年法律第51号（社会福祉法等の一部を改正する法律）"

LAWS = {
    "介護保険法": "平成九年法律第百二十三号",
    "障害者総合支援法": "平成十七年法律第百二十三号",
    "児童福祉法": "昭和二十二年法律第百六十四号",
    "高齢者住まい法": "平成十三年法律第二十六号",
    "住宅セーフティネット法": "平成十九年法律第百十二号",
    "手話施策推進法": "令和七年法律第七十八号",
    "老人福祉法": "昭和三十八年法律第百三十三号",
    "認知症基本法": "令和五年法律第六十五号",
    "社会福祉法": "昭和二十六年法律第四十五号",
    "障害者基本法": "昭和四十五年法律第八十四号",
    "医療介護総合確保促進法": "平成元年法律第六十四号",
    "生活困窮者自立支援法": "平成二十五年法律第百五号",
    "子ども・子育て支援法": "平成二十四年法律第六十五号",
}

_MEM = {}


def _path(nm, asof):
    key = "%s_%s" % (re.sub(r"[^\w]", "_", nm), asof or "now")
    return CACHE / ("%s.xml" % key)


def get(nm, asof=None):
    """法令の XML の根を返す。取れなければ None。

    asof は "YYYY-MM-DD"。省くと現行。
    一度取ったものはファイルに残し、二度目からは取りに行かない。
    """
    k = (nm, asof)
    if k in _MEM:
        return _MEM[k]
    CACHE.mkdir(parents=True, exist_ok=True)
    p = _path(nm, asof)
    if not (p.exists() and p.stat().st_size > 2000):
        if nm not in LAWS:
            return None
        url = API + urllib.parse.quote(LAWS[nm]) + "?response_format=xml"
        if asof:
            url += "&asof=" + asof
        r = subprocess.run(["curl", "-sS", "--max-time", "180", url,
                            "-o", str(p)], capture_output=True)
        if r.returncode != 0 or not p.exists() or p.stat().st_size < 2000:
            _MEM[k] = None
            return None
    try:
        root = ET.parse(p).getroot()
    except Exception:
        _MEM[k] = None
        return None
    # 404 などは小さな XML／JSON が返る。条がなければ取れていない。
    if not list(root.iter("Article")):
        _MEM[k] = None
        return None
    _MEM[k] = root
    return root


def honbun(root):
    """本則（MainProvision）だけを返す。

    **附則にも同じ条番号の条がある。** 本則に限らないと、附則の条を
    拾って「第117条は1項しかない」といった取り違えが起きる。
    """
    mp = list(root.iter("MainProvision"))
    return mp[0] if mp else root


def arts(root):
    """本則の条を {条番号: 要素} で返す。先に現れたものを採る。"""
    d = {}
    for a in honbun(root).iter("Article"):
        d.setdefault(a.get("Num"), a)
    return d


def tx(el):
    return re.sub(r"\s+", "", "".join(el.itertext()))


def find(root, jo, ko, go=None):
    """条・項・号の本文を返す。見つからなければ None。

    条番号の「の」は "_"（第33条の20 → "33_20"、第4条の2 → "4_2"）。
    """
    a = arts(root).get(jo)
    if a is None:
        return None
    for ph in a.findall("Paragraph"):
        if ph.get("Num") != ko:
            continue
        if go is None:
            return tx(ph)
        for it in ph.findall("Item"):
            if it.get("Num") == go:
                return tx(it)
        return None
    return None


def kozu(root, jo):
    """条の組立て（項の数と、項ごとの号の数）を返す。"""
    a = arts(root).get(jo)
    if a is None:
        return None
    return [(p.get("Num"), len(p.findall("Item")))
            for p in a.findall("Paragraph")]


def midashi(root, jo):
    a = arts(root).get(jo)
    if a is None:
        return None
    c = a.find("ArticleCaption")
    return tx(c) if c is not None else ""


def ban(root):
    """取った版の情報（改正法と施行日）を返す。"""
    ri = root.find("revision_info")
    if ri is None:
        return {}
    out = {}
    for k in ("law_revision_id", "amendment_enforcement_date",
              "amendment_law_num", "amendment_law_title"):
        e = ri.find(k)
        out[k] = e.text if e is not None else None
    return out


if __name__ == "__main__":
    for nm in LAWS:
        a = get(nm)
        b = get(nm, TENKI)
        def _s(r):
            if r is None:
                return "取れず"
            v = ban(r)
            return "%s（%s）" % (v.get("amendment_enforcement_date") or "?",
                               v.get("amendment_law_num") or "?")
        print("%-12s 現行 %-28s 第10期初日 %s" % (nm, _s(a), _s(b)))
