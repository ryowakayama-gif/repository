"""令和8年度交付金 該当状況調査票集計表（市町村分）から小野町の値を取り出す。

**順位の母数について（重要）**
  本表は**市町村分のみ**で、広域連合分は別表であり受領していない。
  同表の令和7年度の順位欄は1位から1,741位まであるのに対し、
  本表の行は1,536しかなく、上位100位のうち83しか埋まらない（4位が欠ける）。
  **つまり国の公表順位の母数は広域連合を含む全保険者であり、
  本表だけでは再現できない。**
  小野町の令和7年度は、本表の中で1,339位、公表順位は1,513位である。
  したがって本モジュールが返す「全国順位」は
  **市町村分1,536保険者の中での順位**であって、公表順位ではない。
  「全国平均」「全国該当率」も同じく市町村分の中での値である。
  福島県は59市町村すべてが本表にあり（広域連合がない）、
  県内順位は母数59で完全である。

令和8年9月28日に受領した
  001732614_3.xlsx（全国集計（市町村）・1,574保険者×1,563列）
  001732616.pdf（評価指標（市町村分）・42頁）
による。**これにより、小野町の得点を明細まで追えるようになった。**

**これまで用いていた地域包括ケア「見える化」システムの値（W126〜W155）は
令和6年度交付金（令和5年度の取組に対する評価）のものであり、
本表は令和8年度交付金（令和7年度の取組に対する評価）である。**
2年の間に小野町の得点は大きく変わっているため、本表を正とする。

県内比較は、個別の団体名と個別の得点を出さず、
分布（平均・中央値・四分位・最大・最小）と小野町の順位のみを用いる。
"""

import json
import pathlib
import statistics as st
import warnings

warnings.filterwarnings("ignore")

import openpyxl

ROOT = pathlib.Path(__file__).parent
SRC = (ROOT / "小野町_引継ぎ_整理済" / "21_給付適正化・交付金" / "原本_受領_20260928"
       / "【受領】令和8年度交付金_該当状況調査票集計表_市町村分_20260928.xlsx")
CACHE = ROOT / "output" / "ono_kofukin_cache.json"

HOKENSHA = "075226"        # 小野町の保険者番号
KEN = "福島県"

# 合計・順位の列（採点列から除く）
TOTAL_COLS = {12, 13, 54, 55, 56, 76, 77, 78, 105, 106, 107, 140, 141,
              235, 236, 237, 264, 265, 266, 298, 299, 300, 333, 334, 335}

# 見出しの読み方（列番号 → 表示名）
SUMMARY = [
    (335, "推進・支援合計", 800), (141, "保険者機能強化推進交付金 計", 400),
    (334, "介護保険保険者努力支援交付金 計", 400),
    (56, "推進Ⅰ　持続可能な地域のあるべき姿をかたちにする", 100),
    (78, "推進Ⅱ　公正・公平な給付を行う体制を構築する（介護給付適正化）", 100),
    (107, "推進Ⅲ　介護人材の確保その他のサービス提供基盤の整備", 100),
    (140, "推進Ⅳ　高齢者の自立した日常生活（成果指標群）", 100),
    (237, "支援Ⅰ　介護予防／日常生活支援を推進する", 100),
    (266, "支援Ⅱ　認知症総合支援を推進する", 100),
    (300, "支援Ⅲ　在宅医療・在宅介護連携の体制を構築する", 100),
    (333, "支援Ⅳ　（推進Ⅳと同一の成果指標群）", 100),
]
# 指標群ごとの配点は評価指標（市町村分）の各群の見出しに明記されている。
# 明細の列の最大値を足し上げると成果指標群のように配点を超えることがある
# （ア・イが択一で「より上位となった方で得点」するため）ので、配点は
# 評価指標に書かれた数を用いる。
GUNBETSU = [
    (54, "推進Ⅰ　体制・取組", 64), (55, "推進Ⅰ　活動", 36),
    (76, "推進Ⅱ　体制・取組", 68), (77, "推進Ⅱ　活動", 32),
    (105, "推進Ⅲ　体制・取組", 64), (106, "推進Ⅲ　活動", 36),
    (235, "支援Ⅰ　体制・取組", 52), (236, "支援Ⅰ　活動", 48),
    (264, "支援Ⅱ　体制・取組", 64), (265, "支援Ⅱ　活動", 36),
    (298, "支援Ⅲ　体制・取組", 68), (299, "支援Ⅲ　活動", 32),
]
# 目標Ⅳ（成果指標群）は2つの交付金で同一の指標であり、同じ得点が両方に計上される。
# 支援側（301〜333）は推進側（108〜140）の重複であるため、明細の集計では除く。
DUP_RANGE = range(301, 334)


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _extract():
    wb = openpyxl.load_workbook(SRC, data_only=True, read_only=True)
    ws = wb["全国集計（市町村）"]
    head, rows = [], []
    for i, r in enumerate(ws.iter_rows(values_only=True), 1):
        if i <= 8:
            head.append(list(r))
            continue
        if r[3] is None:
            continue
        rows.append(list(r))
    wb.close()

    ono = next(r for r in rows if r[3] == HOKENSHA)
    ken = [r for r in rows if r[6] == KEN]

    # 見出しは結合セルのため、値を右へ引き継いで組み立てる。
    # シートの行2＝交付金／行3＝目標／行4＝指標群／行5＝指標番号／
    # 行6＝指標名／行7・8＝枝番（head[0] がシートの行1）。
    carry, label = {}, {}
    for c in range(14, 336):
        for r in (1, 2, 3, 4, 5):
            v = head[r][c] if c < len(head[r]) else None
            if v is not None:
                carry[r] = str(v).replace("\n", "")
                for rr in range(r + 1, 6):
                    carry.pop(rr, None)
        sub = " ".join(str(head[r][c]) for r in (6, 7)
                       if c < len(head[r]) and head[r][c] is not None)
        label[c] = {"交付金": carry.get(1, ""), "目標": carry.get(2, ""),
                    "指標群": carry.get(3, ""), "番号": carry.get(4, ""),
                    "指標": carry.get(5, ""), "枝番": sub.strip()}

    def stats(c, haiten=None):
        """1列の分布と小野町の位置。順位は同点を同じ順位とする（競技順位）。"""
        a = [x for x in (_num(r[c]) for r in rows) if x is not None]
        k = [x for x in (_num(r[c]) for r in ken) if x is not None]
        o = _num(ono[c])
        out = {"小野町": o, "全国平均": st.mean(a), "全国最大": max(a),
               "全国件数": len(a), "県平均": st.mean(k),
               "県中央": st.median(k), "県最大": max(k), "県最小": min(k),
               "県件数": len(k)}
        if haiten is not None:
            out["全国満点数"] = sum(1 for x in a if x >= haiten)
            out["県満点数"] = sum(1 for x in k if x >= haiten)
        if o is not None:
            out["全国順位"] = sorted(a, reverse=True).index(o) + 1
            out["県内順位"] = sorted(k, reverse=True).index(o) + 1
            out["全国同点数"] = a.count(o)
            out["県同点数"] = k.count(o)
            q = sorted(k)
            out["県第3四分位"] = q[int(len(q) * 0.75)]
            out["県第1四分位"] = q[int(len(q) * 0.25)]
        return out

    d = {
        "基本": {"人口": _num(ono[8]), "第1号被保険者数": _num(ono[9]),
                "過疎地域": ono[11], "前年度得点": _num(ono[12]),
                "前年度順位": _num(ono[13]),
                "前年度全国平均": st.mean([x for x in
                                      (_num(r[12]) for r in rows)
                                      if x is not None]),
                # 本表は市町村分のみで、広域連合分は別表（未受領）。
                # 令和7年度の順位欄は1〜1,741まであり、本表の1,536行では
                # 上位100位のうち83しか埋まらない。つまり公表順位の母数は
                # 広域連合を含む全保険者であり、本表だけでは再現できない。
                "前年度順位母数": max(x for x in (_num(r[13]) for r in rows)
                                if x is not None),
                "前年度表内順位": (
                    sorted((x for x in (_num(r[12]) for r in rows)
                            if x is not None), reverse=True)
                    .index(_num(ono[12])) + 1)},
        "合計": {nm: dict(stats(c, mx), 満点=mx) for c, nm, mx in SUMMARY},
        "指標群": {nm: dict(stats(c, hp), 配点=hp) for c, nm, hp in GUNBETSU},
        # 得点の分布。団体名は持たず、値だけを並べる。
        # 「もしあと何点とれたら何位か」を後から計算できるようにしておく。
        "県分布": sorted(x for x in (_num(r[335]) for r in ken)
                       if x is not None),
        "全国分布": sorted(x for x in (_num(r[335]) for r in rows)
                        if x is not None),
        "分布": {nm: {"全国": sorted(x for x in (_num(r[c]) for r in rows)
                                  if x is not None),
                    "県": sorted(x for x in (_num(r[c]) for r in ken)
                                if x is not None)}
               for c, nm, _mx in SUMMARY},
        "明細": [],
    }
    for c in range(14, 336):
        if c in TOTAL_COLS:
            continue
        a = [x for x in (_num(r[c]) for r in rows) if x is not None]
        o = _num(ono[c])
        if o is None or not a:
            continue
        d["明細"].append({
            "列": c, **label[c], "小野町": o, "満点": max(a),
            "全国平均": st.mean(a),
            "全国該当率": sum(1 for v in a if v > 0) / len(a),
            "重複": c in DUP_RANGE,
        })
    return d


def load(refresh=False):
    """抽出結果を返す。2度目からはキャッシュを読む（元表は2MB・1,563列）。"""
    if not refresh and CACHE.exists():
        return json.loads(CACHE.read_text(encoding="utf-8"))
    d = _extract()
    CACHE.parent.mkdir(parents=True, exist_ok=True)
    CACHE.write_text(json.dumps(d, ensure_ascii=False), encoding="utf-8")
    return d


D = load()


def meisai(**cond):
    """明細を絞り込む。重複（支援側の目標Ⅳ）は既定で除く。"""
    dup = cond.pop("重複", False)
    out = [m for m in D["明細"] if dup or not m["重複"]]
    for k, v in cond.items():
        out = [m for m in out if v in str(m.get(k, ""))]
    return out


# 指標ごとの配点。評価指標（市町村分）の「配点」欄に書かれている数を写した。
# 枝番の列の最大値を足し上げた数とは、次の2つの指標群で食い違う。
#   目標Ⅳ（成果指標群）　　ア・イが択一で「より上位となった方で得点」するため
#                       列の最大値の和は各20点の2倍になる
#   支援Ⅰ（ⅱ）の指標9　　「多様なサービス・活動の実施状況」は3つの割合の
#                       いずれかで最大4点
HAITEN = {
    ("保険者機能強化推進交付金", "目標Ⅰ", "（ⅰ）"): {"1": 16, "2": 16,
                                          "3": 16, "4": 16},
    ("保険者機能強化推進交付金", "目標Ⅰ", "（ⅱ）"): {"1": 12, "2": 12, "3": 12},
    ("保険者機能強化推進交付金", "目標Ⅱ", "（ⅰ）"): {"1": 32, "2": 36},
    ("保険者機能強化推進交付金", "目標Ⅱ", "（ⅱ）"): {"1": 16, "2": 16},
    ("保険者機能強化推進交付金", "目標Ⅲ", "（ⅰ）"): {"1": 30, "2": 34},
    ("保険者機能強化推進交付金", "目標Ⅲ", "（ⅱ）"): {"1": 12, "2": 12, "3": 12},
    ("保険者機能強化推進交付金", "目標Ⅳ", "成果"): {"1": 20, "2": 20, "3": 20,
                                        "4": 20, "5": 20},
    ("介護保険保険者努力支援交付金", "目標Ⅰ", "（ⅰ）"): {"1": 6, "2": 9, "3": 7,
                                           "4": 7, "5": 7, "6": 9,
                                           "7": 7},
    ("介護保険保険者努力支援交付金", "目標Ⅰ", "（ⅱ）"): {"1": 4, "2": 12, "3": 4,
                                           "4": 8, "5": 4, "6": 4,
                                           "7": 4, "8": 4, "9": 4},
    ("介護保険保険者努力支援交付金", "目標Ⅱ", "（ⅰ）"): {"1": 25, "2": 19, "3": 20},
    ("介護保険保険者努力支援交付金", "目標Ⅱ", "活動"): {"1": 12, "2": 12, "3": 12},
    ("介護保険保険者努力支援交付金", "目標Ⅲ", "（ⅰ）"): {"1": 26, "2": 21, "3": 21},
    ("介護保険保険者努力支援交付金", "目標Ⅲ", "（ⅱ）"): {"1": 16, "2": 16},
}


def _haiten(m):
    gun = "成果" if "成果" in m["指標群"] else (
        "（ⅰ）" if "（ⅰ）" in m["指標群"] else (
            "（ⅱ）" if "（ⅱ）" in m["指標群"] else "活動"))
    return HAITEN.get((m["交付金"], m["目標"][:3], gun), {}).get(m["番号"])


def shihyo():
    """指標（枝番をまとめた単位）ごとの得点。評価指標の配点を満点とする。"""
    out = {}
    for m in meisai():
        k = (m["交付金"], m["目標"], m["指標群"], m["番号"], m["指標"])
        a = out.setdefault(k, {"交付金": m["交付金"], "目標": m["目標"],
                               "指標群": m["指標群"], "番号": m["番号"],
                               "指標": m["指標"], "小野町": 0.0,
                               "全国平均": 0.0, "配点": _haiten(m)})
        a["小野町"] += m["小野町"]
        a["全国平均"] += m["全国平均"]
    return list(out.values())


def zero_items(rate=0.40):
    """0点で、全国の多くが得点している項目。失点の大きい順。"""
    z = [m for m in meisai() if m["小野町"] == 0 and m["全国該当率"] >= rate]
    return sorted(z, key=lambda m: (-m["満点"], -m["全国該当率"]))


def full_items():
    """満点の項目。小野町の強み。"""
    return [m for m in meisai() if m["小野町"] >= m["満点"] > 0]


def rank_if(nm, value, ken=False):
    """その得点なら何位か。他の保険者の得点が変わらないとした場合の目安。"""
    a = D["分布"][nm]["県" if ken else "全国"]
    return sum(1 for x in a if x > value) + 1, len(a)


if __name__ == "__main__":
    d = load(refresh=True)
    b, g = d["基本"], d["合計"]["推進・支援合計"]
    print(f"小野町　人口{b['人口']:,.0f}人／第1号{b['第1号被保険者数']:,.0f}人"
          f"／過疎{b['過疎地域']}")
    print(f"  令和8年度交付金 {g['小野町']:,.0f}点／800点"
          f"　全国平均{g['全国平均']:.1f}　県平均{g['県平均']:.1f}")
    print(f"  市町村分{g['全国順位']:,}位／{g['全国件数']:,}　"
          f"県内{g['県内順位']}位／{g['県件数']}")
    print(f"  ※ 公表順位の母数は広域連合を含む{b['前年度順位母数']:,.0f}保険者。"
          f"令和7年度は表内{b['前年度表内順位']:,}位に対し公表"
          f"{b['前年度順位']:,.0f}位で、{b['前年度順位'] - b['前年度表内順位']:,.0f}位の差")
    print(f"  令和7年度交付金 {b['前年度得点']:,.0f}点"
          f"（市町村分の平均{b['前年度全国平均']:.1f}）"
          f"　→　{g['小野町'] - b['前年度得点']:+,.0f}点")
    z = zero_items()
    print(f"  0点かつ全国該当率40％以上 {len(z)}項目／失点{sum(m['満点'] for m in z):.0f}点")
    print(f"  満点の項目 {len(full_items())}件／明細 {len(meisai())}件")
