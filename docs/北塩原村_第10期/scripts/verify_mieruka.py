# -*- coding: utf-8 -*-
"""地域包括ケア「見える化」システムの収録データと、計画素案・図の数値との突合.

見える化から出力した実物（data/mieruka_raw/*.xlsx）を読み、
本村の数値が素案の表・図の正本と一致しているかを機械で確かめる。

**本村の数値のみを突合の対象とする。** 県内の他自治体の数値は参照値であり、
成果品には団体名も数値も持ち込まない（素案には掲げない）。
福島県平均・全国平均は公表値であり、比較の相手として用いる。

  python3 scripts/verify_mieruka.py          突合の結果を表示する
  python3 scripts/verify_mieruka.py --csv    突合の表を data/ に書き出す

1件でも不適合があると終了コード1で終わる。
"""
import io
import os
import re
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from openpyxl import load_workbook

import data_zuhyo as DZ
import soan_content as SC

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW = os.path.join(BASE, "data", "mieruka_raw")
RESULTS = []


def chk(no, name, ok, detail=""):
    RESULTS.append((no, name, ok, detail))


def book(pat):
    for f in sorted(os.listdir(RAW)):
        if f.startswith(pat):
            return load_workbook(os.path.join(RAW, f), data_only=True)
    raise FileNotFoundError(pat)


def periods(wb, sheet="表"):
    """見える化の期の名（年・年度）を返す。値の並びと同じ添字で対応する"""
    ws = wb[sheet]
    for r in ws.iter_rows(min_row=1, max_row=6, values_only=True):
        head = [v for v in r[4:]]
        if sum(1 for v in head if v not in (None, "")) >= 5:
            return [_pkey(v) for v in head]
    return []


def _pkey(v):
    """期の名を突き合わせるための正規化.

    見える化と成果品では期の書き方が違う。
      見える化  2010 ／ 平成30年度 ／ 令和元年度 ／ 令和2年<改行>3月末
      成果品    2010年 ／ H30年度 ／ R元年度 ／ 令和2年3月末
    同じ形（2010・H30・R1・R2）に寄せてから突き合わせる。
    """
    if v is None:
        return ""
    t = re.sub(r"[\s\u3000]", "", str(v))
    t = re.sub(r"[（(].*?[）)]", "", t)          # (R8/2月サービス提供分まで) を落とす
    t = t.replace("平成", "H").replace("令和", "R").replace("昭和", "S")
    t = t.replace("元年", "1年")
    t = t.replace("年度", "年").replace("時点", "")
    # 月は落とさない。P2には令和8年3月末と令和8年5月末の2列があり、
    # 落とすと同じ期になって別の時点の値を突き合わせてしまう
    t = re.sub(r"(\d+)月末?", r"/\1", t)
    # 和暦を西暦に寄せる。素案の表は和暦（平成22年）、見える化のP1は西暦（2010）
    # であり、揃えないと1行も突き合わないまま「照合した」ことになってしまう
    GENGO = {"H": 1988, "R": 2018, "S": 1925}
    m = re.match(r"^([HRS])(\d+)年(.*)$", t)
    if m:
        t = str(GENGO[m.group(1)] + int(m.group(2))) + m.group(3)
    t = t.replace("年", "")
    return t


def rows_of(wb, sheet="表"):
    """(見出しの連なり, 値の並び) を返す。見出しはA〜D列の連結"""
    ws = wb[sheet]
    out = []
    for r in ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True):
        # A〜D列（添字0〜3）が見出しと単位、E列（添字4）から値
        lab = "／".join(str(v).strip() for v in r[:4] if v not in (None, ""))
        vals = [v for v in r[4:]]
        out.append((lab, vals))
    return out


def pick(rs, *words, exclude=()):
    """見出しに words をすべて含み、exclude を含まない最初の行の値"""
    for lab, vals in rs:
        if all(w in lab for w in words) and not any(e in lab for e in exclude):
            return [v for v in vals if v is not None]
    return None


def near(a, b, tol=0.051):
    if a is None or b is None:
        return False
    return abs(float(a) - float(b)) <= tol


def cmp_series(no, name, got, want, tol=0.051, gp=None, wp=None):
    """見える化の値（got）と成果品の値（want）を突き合わせる。

    **期の名で合わせる。** 見える化と成果品では収める期間の始まりが違うことが
    あり（P3は平成30年度から、図2-9は平成29年度から）、位置で重ねると
    実際には一致しているものをずれとして拾ってしまう。
    期の名が渡されないときだけ先頭から重ねる。
    """
    if got is None:
        chk(no, name, False, "見える化の側に該当の行がない")
        return
    if gp and wp:
        gm = {k: v for k, v in zip(gp, got) if k}
        wm = {_pkey(k): v for k, v in zip(wp, want)}
        keys = [k for k in wm if k in gm]
        if not keys:
            chk(no, name, False,
                f"期の名が突き合わない（見える化 {gp[:2]} / 成果品 {wp[:2]}）")
            return
        bad = [k for k in keys if not near(gm[k], wm[k], tol)]
        if bad:
            d = "／".join(f"{k} 見える化{gm[k]}≠成果品{wm[k]}" for k in bad[:3])
            chk(no, name, False, d)
        else:
            nog = [k for k in wm if k not in gm]
            extra = f"（成果品にあり見える化にない期 {len(nog)}）" if nog else ""
            chk(no, name, True, f"{len(keys)}点すべて一致{extra}")
        return
    n = min(len(got), len(want))
    bad = [i for i in range(n) if not near(got[i], want[i], tol)]
    if bad:
        d = "／".join(f"{i + 1}番目 見える化{got[i]}≠成果品{want[i]}" for i in bad[:3])
        chk(no, name, False, d)
    else:
        chk(no, name, True, f"{n}点すべて一致")


def soan_table(sec_no, head0=None, head_has=None):
    for ch in SC.CH:
        for sec in ch["sections"]:
            if sec["no"] != sec_no:
                continue
            for b in sec["blocks"]:
                if b["t"] not in ("table", "kpi"):
                    continue
                if head0 and b["head"][0] != head0:
                    continue
                if head_has and not all(h in b["head"] for h in head_has):
                    continue
                return b
    return None


def num(v):
    t = re.sub(r"[,\s　人円点％%か所年度月世帯台位]", "", str(v))
    t = t.replace("▲", "-").replace("＋", "").replace("△", "-")
    try:
        return float(t)
    except ValueError:
        return None


def main():
    Z = DZ.load(use_book=False)

    # ── P1 人口の推移 ────────────────────────────────
    _b = book("P1_"); rs = rows_of(_b); gp1 = periods(_b)
    C3 = Z["fig2-3_将来推計人口"]["cats"]
    C4 = Z["fig2-4_前期後期高齢者"]["cats"]
    C2 = Z["fig2-2_高齢化率推移"]["cats"]
    cmp_series(1, "P1 総人口が図2-3と一致すること",
               pick(rs, "人口", "（人)", exclude=("15歳", "40歳", "65歳", "75歳",
                                                  "生産年齢", "高齢者")),
               [v for n, v in Z["fig2-3_将来推計人口"]["series"] if n == "総人口"][0],
               gp=gp1, wp=C3)
    for lab, ser in [("15歳未満", "15歳未満"), ("生産年齢人口", "15〜64歳"),
                     ("高齢者人口", "65歳以上")]:
        cmp_series(1, f"P1 {lab}が図2-3と一致すること", pick(rs, lab),
                   [v for n, v in Z["fig2-3_将来推計人口"]["series"] if n == ser][0],
                   gp=gp1, wp=C3)
    cmp_series(2, "P1 前期高齢者が図2-4と一致すること", pick(rs, "65歳～75歳未満"),
               [v for n, v in Z["fig2-4_前期後期高齢者"]["series"] if n == "65〜74歳"][0],
               gp=gp1, wp=C4)
    cmp_series(2, "P1 後期高齢者が図2-4と一致すること", pick(rs, "75歳以上"),
               [v for n, v in Z["fig2-4_前期後期高齢者"]["series"] if n == "75歳以上"][0],
               gp=gp1, wp=C4)
    for lab, ser in [("高齢化率／（%）", "北塩原村"), ("高齢化率（福島県）", "福島県"),
                     ("高齢化率（全国）", "全国")]:
        cmp_series(3, f"P1 {lab}が図2-2と一致すること",
                   pick(rs, *lab.split("／")),
                   [v for n, v in Z["fig2-2_高齢化率推移"]["series"] if n == ser][0],
                   gp=gp1, wp=C2)

    # ── P2 認定者数・認定率 ──────────────────────────
    _b = book("P2_"); rs = rows_of(_b); gp2 = periods(_b)
    C6 = Z["fig2-6_要介護度別認定者数"]["cats"]
    C5 = Z["fig2-5_認定率推移"]["cats"]
    for ser in ["要支援1", "要支援2", "要介護1", "要介護2", "要介護3", "要介護4", "要介護5"]:
        key = ser.replace("1", "１").replace("2", "２").replace("3", "３") \
                 .replace("4", "４").replace("5", "５")
        got = pick(rs, f"認定者数（{key}）", exclude=("グラフ",))
        want = [v for n, v in Z["fig2-6_要介護度別認定者数"]["series"] if n == ser][0]
        cmp_series(4, f"P2 {ser}の認定者数が図2-6と一致すること", got, want,
                   gp=gp2, wp=C6)
    for lab, ser in [("認定率／（%）", "北塩原村"), ("認定率（福島県）", "福島県"),
                     ("認定率（全国）", "全国")]:
        got = pick(rs, *lab.split("／"))
        cmp_series(5, f"P2 {lab}が図2-5と一致すること", got,
                   [v for n, v in Z["fig2-5_認定率推移"]["series"] if n == ser][0],
                   gp=gp2, wp=C5)
    # 認定者数の合計（素案5-3・2-3が用いる系列）
    nin = pick(rs, "認定者数", "（人）", exclude=("要支援", "要介護", "経過的", "グラフ"))
    chk(6, "P2 令和8年3月末の認定者数が素案の214人と一致すること",
        near(nin[7], 214), f"見える化{nin[7]}人")
    chk(7, "P2 令和8年5月末の認定者数が211人であること（素案の211は推計ではない）",
        near(nin[8], 211), f"見える化{nin[8]}人（5月末時点）")

    # ── P3 費用額 ──────────────────────────────────
    _b = book("P3_"); rs = rows_of(_b); gp3 = periods(_b)
    C9 = Z["fig2-9_1人あたり費用額"]["cats"]
    for lab, ser in [("第1号被保険者1人1月あたり費用額／（円）", "北塩原村"),
                     ("第1号被保険者1人1月あたり費用額（福島県）", "福島県"),
                     ("第1号被保険者1人1月あたり費用額（全国）", "全国")]:
        got = pick(rs, *lab.split("／"))
        want = [v for n, v in Z["fig2-9_1人あたり費用額"]["series"] if n == ser][0]
        cmp_series(8, f"P3 {lab}が図2-9と一致すること", got, want, tol=0.6,
                   gp=gp3, wp=C9)

    # ── P4 保険料額 ────────────────────────────────
    _b = book("P4_"); rs = rows_of(_b); gp4 = periods(_b)
    C17 = Z["fig2-17_保険料と必要額"]["cats"]
    cmp_series(9, "P4 必要保険料額が図2-17と一致すること",
               pick(rs, "必要保険料額（合計）"),
               [v for n, v in Z["fig2-17_保険料と必要額"]["series"]
                if n == "必要保険料額"][0], gp=gp4, wp=C17)
    cmp_series(9, "P4 保険料基準額が図2-17と一致すること",
               pick(rs, "保険料基準額", "（円）",
                    exclude=("福島県", "全国")),
               [v for n, v in Z["fig2-17_保険料と必要額"]["series"]
                if n == "保険料基準額"][0], gp=gp4, wp=C17)
    for lab, ser in [("保険料基準額", "北塩原村"), ("保険料基準額（福島県）", "福島県平均"),
                     ("保険料基準額（全国）", "全国平均")]:
        got = pick(rs, lab, "（円）",
                   exclude=() if "（" in lab else ("福島県", "全国"))
        want = [v for n, v in Z["fig2-16_保険料比較"]["series"] if n == ser][0]
        # 図2-16は期ごと（第7・8・9期）。P4は年度ごとなので期の初年度を採る
        cmp_series(10, f"P4 {lab}が図2-16（期別）と一致すること",
                   [got[0], got[3], got[6]] if got else None, want)

    # ── 素案の表との突合（図を経ない直接の照合）──────────────────
    rs1 = rows_of(book("P1_"))
    t21 = soan_table("2-1", head0="年")
    if t21 and "高齢化率" in " ".join(t21["head"]):
        idx = {h: i for i, h in enumerate(t21["head"])}
        got = {}
        for col, lab in [("総人口", "人口／（人)"), ("15歳未満", "15歳未満"),
                         ("生産年齢人口", "生産年齢人口"), ("65〜74歳", "65歳～75歳未満"),
                         ("75歳以上", "75歳以上")]:
            if col not in idx:
                continue
            gv = pick(rs1, *lab.split("／"))
            gm = {k: v for k, v in zip(periods(book("P1_")), gv)}
            bad, n = [], 0
            for r in t21["rows"]:
                k = _pkey(r[0])
                if k not in gm:
                    continue
                n += 1
                if not near(gm[k], num(r[idx[col]])):
                    bad.append(f"{r[0]} 見える化{gm[k]}≠素案{r[idx[col]]}")
            chk(11, f"素案2-1の{col}が P1 と一致すること", not bad and n > 0,
                "／".join(bad[:2]) if bad
                else (f"{n}行を照合" if n else "1行も突き合っていない（期の名を確かめる）"))

    rs2 = rows_of(book("P2_"))
    gp2b = periods(book("P2_"))
    t23 = soan_table("2-3", head0="時点", head_has=["北塩原村", "福島県", "全国"])
    if t23:
        for col, lab in [("北塩原村", "認定率／（%）"), ("福島県", "認定率（福島県）"),
                         ("全国", "認定率（全国）")]:
            i = t23["head"].index(col)
            gm = {k: v for k, v in zip(gp2b, pick(rs2, *lab.split("／")))}
            n = sum(1 for r in t23["rows"] if _pkey(r[0]) in gm)
            bad = [f"{r[0]} 見える化{gm[_pkey(r[0])]}≠素案{r[i]}"
                   for r in t23["rows"]
                   if _pkey(r[0]) in gm and not near(gm[_pkey(r[0])], num(r[i]))]
            chk(12, f"素案2-3の認定率（{col}）が P2 と一致すること", not bad and n > 0,
                "／".join(bad[:2]) if bad
                else (f"{n}行を照合" if n else "1行も突き合っていない（期の名を確かめる）"))

    t23b = soan_table("2-3", head0="時点", head_has=["要支援1", "要介護5"])
    if t23b:
        for col in ["要支援1", "要支援2", "要介護1", "要介護2", "要介護3",
                    "要介護4", "要介護5"]:
            if col not in t23b["head"]:
                continue
            i = t23b["head"].index(col)
            key = col.translate(str.maketrans("12345", "１２３４５"))
            gm = {k: v for k, v in
                  zip(gp2b, pick(rs2, f"認定者数（{key}）", exclude=("グラフ",)))}
            n = sum(1 for r in t23b["rows"] if _pkey(r[0]) in gm)
            bad = [f"{r[0]} 見える化{gm[_pkey(r[0])]}≠素案{r[i]}"
                   for r in t23b["rows"]
                   if _pkey(r[0]) in gm and not near(gm[_pkey(r[0])], num(r[i]))]
            chk(13, f"素案2-3の認定者数（{col}）が P2 と一致すること", not bad and n > 0,
                "／".join(bad[:2]) if bad
                else (f"{n}行を照合" if n else "1行も突き合っていない（期の名を確かめる）"))

    # ── 順位（順位シート）と素案の記載の突合 ─────────────────────
    #    順位は「順位」シートにあり、表のシートには入っていない。
    #    素案は費用額を県内44番目・全国1,153番目と書いていたが、実物は
    #    45番目・1,152番目であった。1つずれるため、必ず実物から読む。
    #
    #    **母数（59保険者など）だけで突き合わせない。** 認定率も費用額も
    #    必要保険料額も母数が同じ59保険者であり、母数で結ぶと別の指標の
    #    順位を拾ってしまう。指標を表す語で結ぶ。
    import re as _re
    txt = "\n".join(
        str(b2.get("v", "")) if b2["t"] in ("p", "h3", "note")
        else " ".join(str(x) for r in b2.get("rows", []) for x in r)
        for ch in SC.CH for sec in ch["sections"] for b2 in sec["blocks"])

    def rank_pairs(pat):
        """順位シートから (見出し, 時点, [(順位, 母数)]) を返す"""
        wb = book(pat)
        if "順位" not in wb.sheetnames:
            return None
        vals = [str(v).strip() for r in wb["順位"].iter_rows(values_only=True)
                for v in r if v not in (None, "")]
        head = vals[0] if vals else ""
        out, tp = [], ""
        for i2, v in enumerate(vals):
            if v.startswith("（") and ("時点" in v or "年" in v):
                tp = v
            m = _re.match(r"^([\d,]+)番目$", v)
            if m and i2 + 1 < len(vals):
                m2 = _re.match(r"^([\d,]+)保険者$", vals[i2 + 1])
                if m2:
                    out.append((m.group(1), m2.group(1), tp))
        return head, out

    # 素案が現に順位を書いている指標だけを見る。
    # **文の中で探さない。** 1つの文に認定率と費用額の順位が並ぶことがあり、
    # 文単位では取り違える。順位シートから「N保険者中M番目」の形を組み立て、
    # その文字列が素案にあること、かつ誤った順位がないことを見る。
    # (点検番号, ファイル, 名前, 用いる組の位置)
    RANKS = [(14, "P2_", "認定率", 0), (15, "P3_", "費用額", 0),
             (16, "P1_", "高齢化率", 1)]
    for no, pat, kw, which in RANKS:
        got = rank_pairs(pat)
        if not got:
            chk(no, f"{kw}の順位が見える化と一致すること", False, "順位シートがない")
            continue
        _head, pairs = got
        use = pairs[which * 2: which * 2 + 2]
        if len(use) < 2:
            chk(no, f"{kw}の順位が見える化と一致すること", False,
                f"順位シートに{which + 1}組目がない（{len(pairs)}組）")
            continue
        bad, n = [], 0
        for rank, tot, _tp in use:
            want = f"{tot}保険者中{rank}番目"
            if want in txt:
                n += 1
                continue
            # 同じ母数で書かれている順位をすべて挙げる。
            # 1つの母数を複数の指標が使うため、1件だけ出すと取り違える
            ms = sorted(set(_re.findall(rf"{_re.escape(tot)}保険者中([\d,]+)番目", txt)))
            if ms:
                bad.append(f"{tot}保険者中：見える化{rank}番目／"
                           f"素案にあるのは{'・'.join(ms)}番目")
            else:
                bad.append(f"{tot}保険者中{rank}番目の記載が素案にない")
        if not bad and n == 0:
            chk(no, f"{kw}の順位が見える化と一致すること", True, "素案に順位の記載なし")
        else:
            chk(no, f"{kw}の順位が見える化と一致すること", not bad,
                "／".join(sorted(set(bad))[:2]) if bad
                else f"{n}対を照合（{use[0][2] or '時点の記載なし'}）")

    # ── 見える化に収録されていない期（突合できない値）─────────────
    UNCOV = []
    for fig, bk in [("fig2-2_高齢化率推移", "P1_"), ("fig2-3_将来推計人口", "P1_"),
                    ("fig2-4_前期後期高齢者", "P1_"), ("fig2-5_認定率推移", "P2_"),
                    ("fig2-6_要介護度別認定者数", "P2_"),
                    ("fig2-9_1人あたり費用額", "P3_"),
                    ("fig2-17_保険料と必要額", "P4_")]:
        gp = set(periods(book(bk)))
        miss = [c for c in Z[fig]["cats"] if _pkey(c) not in gp]
        if miss:
            UNCOV.append((fig, miss))

    # ── 新しく得られた値 ───────────────────────────
    rs3 = rows_of(book("P3_"))
    rs4 = rows_of(book("P4_"))
    rs2 = rows_of(book("P2_"))
    NEW = [
        ("認定者数（令和8年5月末）",
         pick(rs2, "認定者数", "（人）",
              exclude=("要支援", "要介護", "経過的", "グラフ"))[8], "人",
         "図2-6・素案2-3は令和8年3月末まで。5月末の値は本計画では用いない"),
        ("1人1月あたり費用額（令和8年度）",
         pick(rs3, "第1号被保険者1人1月あたり費用額／（円）".split("／")[0], "（円）",
              exclude=("福島県", "全国"))[8], "円",
         "令和8年度はR8/3月サービス提供分までの1か月分のみ。年度の値として扱えない"),
        ("必要保険料額（令和8年度）",
         pick(rs4, "必要保険料額（合計）")[8], "円",
         "同上。1か月分のみのため年度の値として扱えない"),
    ]

    # ── 出力 ───────────────────────────────────
    w = max(len(n) for _, n, _, _ in RESULTS)
    print("■ 見える化（地域分析 P1〜P4）と成果品の突合")
    ng = 0
    for no, name, ok, detail in RESULTS:
        mark = "適合" if ok else "不適合"
        print(f"  {no:>3}  {name:<{w}}  {mark}" + (f"  {detail}" if detail else ""))
        if not ok:
            ng += 1
    print(f"\n  {len(RESULTS)}件のうち適合{len(RESULTS)-ng}件／不適合{ng}件")
    print("\n■ 見える化に新しく収録された値（本計画では用いていない）")
    for nm, v, u, why in NEW:
        print(f"  {nm:<34} {v}{u}　{why}")
    print("\n■ 見える化に収録されておらず、突合できない期")
    if not UNCOV:
        print("  なし")
    for fig, miss in UNCOV:
        print(f"  {fig:<30} {'・'.join(miss)}")
    print("\n  ※ 県内の他自治体の数値は参照値であり、本突合の対象としていない。")
    print("     成果品には団体名も数値も掲げない。")
    return 1 if ng else 0


if __name__ == "__main__":
    sys.exit(main())
