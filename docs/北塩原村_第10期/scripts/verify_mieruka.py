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


import os as _os


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
    # 元号の直後の「元」は1年。D47-aは「R元」のように年を付けずに呼ぶ
    t = re.sub(r"^([HRS])元", r"\g<1>1", t)
    m = re.match(r"^([HRS])(\d+)年(.*)$", t)
    if m:
        t = str(GENGO[m.group(1)] + int(m.group(2))) + m.group(3)
    else:
        # 「H26」のように年を付けない呼び方（D47系列の歳入の年度）
        m = re.match(r"^([HRS])(\d+)$", t)
        if m:
            t = str(GENGO[m.group(1)] + int(m.group(2)))
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


# 図ごとの突合の状態。(状態, 相手または理由)
#   "済"      … 見える化の系列と値で突き合わせている
#   "できない" … 見える化に相手の系列がない。出所が別であることを記す
COVER = {
    "fig2-1_人口構成比較": ("できない",
        "国勢調査の年齢階級別。見える化は15歳未満・15〜64歳・65歳以上の3区分までで、"
        "15〜39歳・40〜64歳の分けがない（総人口・15〜64歳・65歳以上は点検1で突合）"),
    "fig2-2_高齢化率推移": ("済", "P1・A2系列（点検1）"),
    "fig2-3_将来推計人口": ("済", "P1系列（点検1）"),
    "fig2-4_前期後期高齢者": ("済", "P1・A3系列（点検2）"),
    "fig2-5_認定率推移": ("済", "P2系列（点検3）"),
    "fig2-6_要介護度別認定者数": ("済", "B3-a系列（点検25）"),
    "fig2-7_要支援1": ("済", "図2-6と同じB3-a系列の要支援1（点検25に含む）"),
    "fig2-8_新規認定要支援1": ("済", "B9系列（点検29）"),
    "fig2-9_1人あたり費用額": ("済", "P3系列（点検5）"),
    "fig2-10_費用額内訳": ("済",
        "D13系列の給付費から給付率で突合（点検41）。費用額は自己負担を含むため"
        "給付費と厳密には一致せず、給付率の帯で見ている"),
    "fig2-11_在宅施設別給付月額": ("済", "D6-a・D6-b系列（点検38・39）"),
    "fig2-12_1人あたり定員": ("済", "D29・D30系列（点検30・31）"),
    "fig2-13_地域支援事業費": ("済", "D48-c系列の決算の内訳（点検37）"),
    "fig2-14_サービス区分別変化": ("済", "D13系列の27サービスの和（点検40）"),
    "fig2-15_財政指数": ("済", "D48-b・D48-c・D47-a系列の決算（点検34〜36）"),
    "fig2-16_保険料比較": ("済", "P4系列（点検6）"),
    "fig2-17_保険料と必要額": ("済", "P4系列（点検7）"),
    "fig2-18_交付金指標群": ("済", "W系列（点検21）"),
    "fig2-19_交付金得点推移": ("済", "W系列（点検18・19）"),
    "fig2-20_通いの場": ("できない",
        "村から受領した令和6年度の実数。見える化のF1・F4系列は平成30年度で止まっている"),
    "fig2-21_調整済み認定率": ("済", "B5-a・B6-a・B6-b系列（点検23〜28）"),
}


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
        # **両側を正規化する。** 片側だけを寄せると、見える化と成果品が同じ
        # 書き方をしている系列（令和2年3月末など）で1件も突き合わない。
        # _pkey は冪等であり、すでに寄せた名を渡しても結果は変わらない。
        gm = {_pkey(k): v for k, v in zip(gp, got) if k}
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

    # ── 交付金：見える化の年度ラベルと素案の年度ラベルの対応 ──────────
    #    見える化は取組の年度で、厚生労働省の該当状況調査票集計表は評価・交付の
    #    年度で呼ぶ。同じ評価が1年ずれた名で載るため、突き合わせるときに
    #    取り違えやすい。実物で対応を確かめ、ずれたら気づくようにする。
    import csv as _csv
    BATCH = os.path.join(BASE, "data", "mieruka_batch.csv")
    if os.path.exists(BATCH):
        w126 = {}
        with io.open(BATCH, encoding="utf-8") as f:
            for r in _csv.DictReader(f):
                if r["file"].startswith("W126") and "北塩原" in r["region"]:
                    w126[r["indicator"]] = float(r["value"])
        # 素案2-8の得点の表
        t28 = soan_table("2-8", head0="")
        got = {}
        for ch in SC.CH:
            for sec in ch["sections"]:
                if sec["no"] != "2-8":
                    continue
                for b2 in sec["blocks"]:
                    if b2["t"] != "table" or "令和6年度" not in b2["head"]:
                        continue
                    i6 = b2["head"].index("令和6年度")
                    for row in b2["rows"]:
                        got[str(row[0])] = row[i6]
        PAIR = [("推進合計", "保険者機能強化推進交付金（400点）"),
                ("支援合計", "介護保険保険者努力支援交付金（400点）"),
                ("推進・支援合計", "合計（800点）")]
        bad18 = []
        n18 = 0
        for wk, sk in PAIR:
            if wk not in w126 or sk not in got:
                continue
            n18 += 1
            if not near(w126[wk], num(got[sk])):
                bad18.append(f"{sk}：見える化(令和5年度){w126[wk]:.0f}≠素案(令和6年度){got[sk]}")
        chk(18, "交付金：見える化の令和5年度が素案の令和6年度と一致すること",
            not bad18 and n18 > 0,
            "／".join(bad18[:2]) if bad18
            else (f"{n18}件一致（見える化は取組の年度、素案は評価の年度で呼ぶ）"
                  if n18 else "突き合わせる値がない"))

        # 目標別の内訳も確かめる（W127＝推進・W129＝支援）
        mokuhyo = {}
        with io.open(BATCH, encoding="utf-8") as f:
            for r in _csv.DictReader(f):
                if r["file"][:4] in ("W127", "W129") and "北塩原" in r["region"]:
                    mokuhyo[(r["file"][:4], r["indicator"])] = float(r["value"])
        s127 = sum(v for (fk, ind), v in mokuhyo.items()
                   if fk == "W127" and ind.startswith("目標") and "合計" in ind
                   and "目標別" not in ind)
        s129 = sum(v for (fk, ind), v in mokuhyo.items()
                   if fk == "W129" and ind.startswith("目標") and "合計" in ind
                   and "目標別" not in ind)
        ok19 = near(s127, w126.get("推進合計", -1)) and \
            near(s129, w126.get("支援合計", -1))
        chk(19, "交付金：目標別の内訳が推進・支援の計と合うこと", ok19,
            f"推進{s127:.0f}／支援{s129:.0f}"
            + ("" if ok19 else f"（計は推進{w126.get('推進合計')}／支援{w126.get('支援合計')}）"))
    else:
        chk(18, "交付金：見える化の令和5年度が素案の令和6年度と一致すること", False,
            "mieruka_batch.csv がない（parse_mieruka_batch.py を実行してください）")

    # ── 一括受領分から加えた記載の突合 ──────────────────────────
    if os.path.exists(BATCH):
        BR = []
        with io.open(BATCH, encoding="utf-8") as f:
            BR = [r for r in _csv.DictReader(f) if "北塩原" in r["region"]]

        def bval(ind, per=None, fpre=None):
            for r in BR:
                if r["indicator"] != ind:
                    continue
                if per is not None and r["period"] != per:
                    continue
                if fpre and not r["file"].startswith(fpre):
                    continue
                return float(r["value"])
            return None

        # 20 村内に所在する事業所の数（素案2-5）
        t20 = soan_table("2-5", head0="サービス")
        bad20, n20 = [], 0
        if t20 is None:
            bad20.append("2-5に村内の事業所数の表がない")
        else:
            YY = [h for h in t20["head"][1:]]
            for row in t20["rows"]:
                for i, y in enumerate(YY):
                    g = bval("サービス提供事業所数（%s）" % row[0], y)
                    if g is None:
                        continue
                    n20 += 1
                    if not near(g, num(row[1 + i])):
                        bad20.append("%s %s：見える化%.0f≠素案%s" % (row[0], y, g, row[1 + i]))
        chk(20, "素案2-5の村内の事業所数が見える化と一致すること",
            not bad20 and n20 > 0,
            "／".join(bad20[:2]) if bad20
            else (f"{n20}点を照合" if n20 else "突き合わせる値がない"))

        # 21 指標群別の得点（素案2-8）
        t21b = soan_table("2-8", head0="指標群")
        bad21, n21 = [], 0
        if t21b is None:
            bad21.append("2-8に指標群別の表がない")
        else:
            MAP = {"取組・体制指標群": "取組・体制指標群合計",
                   "活動指標群": "活動指標群合計",
                   "成果指標群": "成果指標群合計",
                   "計": "指標群別合計"}
            for row in t21b["rows"]:
                ind = MAP.get(row[0])
                if not ind:
                    continue
                for i, fpre in enumerate(("W128", "W130")):
                    g = bval(ind, fpre=fpre)
                    if g is None:
                        continue
                    n21 += 1
                    if not near(g, num(row[1 + i])):
                        bad21.append("%s %s：見える化%.0f≠素案%s"
                                     % (row[0], "推進" if i == 0 else "支援", g, row[1 + i]))
                # 計の列
                a1, a2_ = bval(ind, fpre="W128"), bval(ind, fpre="W130")
                if a1 is not None and a2_ is not None and len(row) > 3:
                    n21 += 1
                    if not near(a1 + a2_, num(row[3])):
                        bad21.append("%s 計：見える化%.0f≠素案%s" % (row[0], a1 + a2_, row[3]))
        chk(21, "素案2-8の指標群別の得点が見える化と一致すること",
            not bad21 and n21 > 0,
            "／".join(bad21[:2]) if bad21
            else (f"{n21}点を照合" if n21 else "突き合わせる値がない"))

    # ── 22 認定の状況が令和8年3月末で揃っていること ──────────────
    #    見える化には3月末と5月末の2時点がある。村のご判断により本計画は
    #    3月末による。5月末の値（認定者211人・認定率20.7%）が本文に紛れ込むと、
    #    同じ指標に2つの値が並ぶ。順位だけは5月末のものしか公表がないため例外。
    rs22 = rows_of(book("P2_"))
    gp22 = periods(book("P2_"))
    nin22 = pick(rs22, "認定者数", "（人）",
                 exclude=("要支援", "要介護", "経過的", "グラフ"))
    ninr22 = pick(rs22, "認定率", "（%）", exclude=("福島県", "全国"))
    m22 = {k: v for k, v in zip(gp22, nin22)}
    r22 = {k: v for k, v in zip(gp22, ninr22)}
    # 期のキーは _pkey が決める（和暦は西暦に寄せる）。決め打ちしない
    K3, K5 = _pkey("令和8年3月末"), _pkey("令和8年5月末")
    v3, v5 = m22.get(K3), m22.get(K5)
    p3, p5 = r22.get(K3), r22.get(K5)
    bad22 = []
    if v3 is None or v5 is None:
        bad22.append("見える化に3月末・5月末の両方がない")
    else:
        # 3月末の値が本文にあること
        if f"{v3:.0f}人" not in txt:
            bad22.append(f"3月末の認定者数{v3:.0f}人が本文にない")
        if f"{p3}%" not in txt:
            bad22.append(f"3月末の認定率{p3}%が本文にない")
        # 5月末の値は、時点を断った注記の中にだけあってよい。
        # ただし同じ数字が別の指標の値として出ることがある（令和9年度の認定率20.7%）。
        # 文または表のます目の単位で見て、第10期の年度を名指ししているものは別の指標
        # として扱う。値だけを数えると、関わりのない一致で鳴ってしまう。
        import re as _re22
        MIKOMI22 = ("令和9年度", "令和10年度", "令和11年度", "令和12年度",
                    "令和17年度", "令和22年度")

        def _tani22():
            """素案を、文と表のます目に分けて（その単位が注記かどうかと一緒に）返す"""
            for ch in SC.CH:
                for sec in ch["sections"]:
                    for b2 in sec["blocks"]:
                        if b2["t"] in ("p", "note", "h3"):
                            for bun in _re22.split(r"(?<=。)", str(b2.get("v", ""))):
                                yield bun, b2["t"] == "note", None
                        elif b2["t"] == "bullets":
                            for x in b2.get("v", []):
                                yield str(x), False, None
                        elif b2["t"] in ("table", "kpi"):
                            for r in b2.get("rows", []):
                                for x in r:
                                    # ます目は行の見出しと一緒に見る（年度は行の見出しにある）
                                    yield str(x), False, str(r[0])
        for val, lab in ((f"{v5:.0f}人", "認定者数"), (f"{p5}%", "認定率")):
            soto = 0
            for bun, is_note, gyomi in _tani22():
                if val not in bun:
                    continue
                if is_note:
                    continue
                mawari = bun + (gyomi or "")
                if any(y in mawari for y in MIKOMI22):
                    continue        # 第10期以降の見込みの値。5月末の実績ではない
                soto += 1
            if soto:
                bad22.append(f"5月末の{lab}{val}が注記の外にある（{soto}件）")
    chk(22, "認定の状況が令和8年3月末で揃っていること", not bad22,
        "／".join(bad22[:2]) if bad22
        else f"3月末（{v3:.0f}人・{p3}%）を本文に用い、5月末（{v5:.0f}人・{p5}%）は注記のみ")

    # ── 23　性・年齢調整済み認定率が見える化 B5-a・B6-a・B6-b と一致すること ──
    #    「注目する地域のみ」と「他地域と比較」で同じ本村の値が異なるため、
    #    どちらを用いているかを明示したうえで突き合わせる。
    try:
        import collections as _col23
        BAT = os.path.join(BASE, "data", "mieruka_batch.csv")
        d23 = _col23.defaultdict(dict)
        with open(BAT, encoding="utf-8") as f:
            for r in _csv.DictReader(f):
                if r["region"] != "北塩原村" or "注目" not in r["indicator"]:
                    continue
                per = r["period"].replace("時点", "")
                if r["file"].startswith("B5-a_"):
                    d23[r["indicator"]][per] = float(r["value"])
                elif r["file"].startswith("B6-a_"):
                    d23["重度"][per] = float(r["value"])
                elif r["file"].startswith("B6-b_"):
                    d23["軽度"][per] = float(r["value"])
        PER23 = DZ.cats("fig2-21_調整済み認定率")
        tot = [d23["【注目する地域のみ】合計調整済み認定率"].get(p) for p in PER23]
        ni2 = [round(sum(d23["【注目する地域のみ】調整済み認定率（要介護%d）" % g].get(p, 0)
                         for g in (2, 3, 4, 5)), 2) for p in PER23]
        cmp_series(23, "調整済み認定率（合計）が見える化 B5-a と一致すること",
                   tot, DZ.vals("fig2-21_調整済み認定率", "合計"), gp=PER23, wp=PER23)
        cmp_series(24, "調整済み認定率（軽度）が見える化 B6-b と一致すること",
                   [d23["軽度"].get(p) for p in PER23],
                   DZ.vals("fig2-21_調整済み認定率", "軽度（要支援1〜要介護2）"),
                   gp=PER23, wp=PER23)
        cmp_series(25, "調整済み認定率（重度）が見える化 B6-a と一致すること",
                   [d23["重度"].get(p) for p in PER23],
                   DZ.vals("fig2-21_調整済み認定率", "重度（要介護3以上）"),
                   gp=PER23, wp=PER23)
        cmp_series(26, "調整済み要介護2以上が B5-a の要介護2〜5の和と一致すること",
                   ni2, DZ.vals("fig2-21_調整済み認定率", "要介護2以上"),
                   gp=PER23, wp=PER23)
        # 国の評価指標の8.65%（令和5年度）と、令和6年3月末の値が一致すること
        v65 = dict(zip(PER23, DZ.vals("fig2-21_調整済み認定率", "要介護2以上")))["令和6年3月末"]
        chk(27, "調整済み要介護2以上が国の評価指標の8.65%（令和5年度）と一致すること",
            abs(v65 - 8.65) <= 0.051,
            f"令和6年3月末 {v65}% と 8.65%（端数は要介護度ごとの小数第1位による）"
            if abs(v65 - 8.65) <= 0.051 else f"令和6年3月末 {v65}% ≠ 8.65%")
        # 軽度＋重度が合計に一致すること（区分の取り方の確かめ）
        bad23 = []
        for i, p in enumerate(PER23):
            g = (d23["軽度"].get(p, 0) + d23["重度"].get(p, 0))
            if abs(g - tot[i]) > 0.11:
                bad23.append(f"{p} 軽度＋重度{g:.1f}≠合計{tot[i]}")
        chk(28, "調整済み認定率の軽度＋重度が合計と一致すること", not bad23,
            "／".join(bad23[:2]) if bad23
            else f"{len(PER23)}点すべて一致（軽度＝要支援1〜要介護2、重度＝要介護3以上）")
    except Exception as e:
        chk(23, "調整済み認定率が見える化と一致すること", False, f"照合できない（{e}）")

    # ── 29〜31　受領データで突き合わせられる図を増やす ──────────────
    #    正本の系列と見える化の系列を値で突き合わせて相手を探し、
    #    一致が確かめられたものを点検として固定した。
    try:
        import collections as _c29
        BAT = os.path.join(BASE, "data", "mieruka_batch.csv")
        G = _c29.defaultdict(dict)
        with open(BAT, encoding="utf-8") as f:
            for r in _csv.DictReader(f):
                if r["region"] != "北塩原村":
                    continue
                try:
                    G[(r["file"].split("_")[0], r["indicator"])][_pkey(r["period"])] = \
                        float(r["value"])
                except ValueError:
                    pass

        def pick29(bk, ind, periods):
            for (b, i), d in G.items():
                if b == bk and i == ind:
                    return [d.get(_pkey(p)) for p in periods]
            return None

        # 29 新規認定者に占める要支援1の割合（B9 の実数から割合を作って突き合わせる）
        P29 = DZ.cats("fig2-8_新規認定要支援1")
        ys1 = pick29("B9", "新規要支援・要介護認定者（要支援１）", P29)
        cmp_series(29, "新規認定の要支援1の割合が見える化 B9 と一致すること",
                   ys1, DZ.vals("fig2-8_新規認定要支援1", "要支援1の割合"),
                   gp=P29, wp=P29)
        # 30・31 要支援・要介護者1人あたり定員
        P30 = DZ.cats("fig2-12_1人あたり定員")
        cmp_series(30, "1人あたり定員（通所介護）が見える化 D30 と一致すること",
                   pick29("D30", "要支援・要介護者1人あたり定員（通所介護）", P30),
                   DZ.vals("fig2-12_1人あたり定員", "通所系（通所介護）"),
                   tol=0.0006, gp=P30, wp=P30)
        cmp_series(31, "1人あたり定員（認知症GH）が見える化 D29 と一致すること",
                   pick29("D29", "要支援・要介護者1人あたり定員（認知症対応型共同生活介護）",
                          P30),
                   DZ.vals("fig2-12_1人あたり定員", "居住系（認知症GH）"),
                   tol=0.0006, gp=P30, wp=P30)
        # 33 B9の「割合」が「実数÷合計」であることを確かめる。
        #    B9には「新規…者数（要支援１）」（人）と「新規…者（要支援１）」（%）の
        #    2つがあり、名がよく似ている。取り違えていないことを押さえておく。
        n1 = pick29("B9", "新規要支援・要介護認定者数（要支援１）", P29)
        nt = pick29("B9", "新規要支援・要介護認定者数合計", P29)
        bad33 = []
        for i, p33 in enumerate(P29):
            if n1[i] is None or not nt[i] or ys1[i] is None:
                bad33.append(f"{p33} の実数または合計がない")
                continue
            if abs(n1[i] / nt[i] * 100 - ys1[i]) > 0.051:
                bad33.append(f"{p33} 実数{n1[i]:.0f}÷合計{nt[i]:.0f}"
                             f"＝{n1[i] / nt[i] * 100:.1f}%≠割合{ys1[i]}%")
        chk(33, "B9の割合が実数÷合計であること（指標の取り違えの確認）", not bad33,
            "／".join(bad33[:2]) if bad33
            else f"{len(P29)}点すべて一致（割合の系列は「者（要支援１）」、"
                 f"実数の系列は「者数（要支援１）」）")
    except Exception as e:
        chk(29, "受領データで突き合わせられる図", False, f"照合できない（{e}）")

    # ── 34〜37　村の決算による図（図2-16・図2-18）を見える化のD47・D48と突き合わせる ──
    #    決算の系列は「平成27年3月末」のように**締めの日**で期を呼ぶ。
    #    平成27年3月末に締めるのは平成26年度の決算であり、成果品の「H26年度」にあたる。
    #    1年ずらさないと1件も突き合わない（交付金の年度の呼び方と同じ種類のずれ）。
    #    歳入（保険料）D47-a だけは「H26」と年度そのもので呼ぶためずらさない。
    try:
        import collections as _c34
        import re as _re34
        BAT34 = os.path.join(BASE, "data", "mieruka_batch.csv")

        def _nendo(period):
            """決算の「N年3月末」を年度に直す。それ以外はそのまま。"""
            k = _pkey(period)
            m = _re34.match(r"^(\d{4})/3$", k)
            return str(int(m.group(1)) - 1) if m else k

        D = _c34.defaultdict(dict)
        with open(BAT34, encoding="utf-8") as f:
            for r in _csv.DictReader(f):
                if r["region"] != "北塩原村":
                    continue
                bk = r["file"].split("_")[0]
                if not (bk.startswith("D47") or bk.startswith("D48")):
                    continue
                try:
                    D[(bk, r["indicator"])][_nendo(r["period"])] = float(r["value"])
                except ValueError:
                    pass

        def _pick34(bk, ind, periods, warizan=1000.0):
            d = D.get((bk, ind), {})
            return [None if d.get(p) is None else d[p] / warizan for p in periods]

        P34 = [_pkey(c) for c in DZ.cats("fig2-15_財政指数")]
        for no, nm, bk, ind in ((34, "保険給付費", "D48-b", "合計"),
                                (35, "地域支援事業費", "D48-c", "合計"),
                                (36, "保険料収入", "D47-a", "保険料")):
            cmp_series(no, f"財政指数の{nm}が見える化 {bk} の決算と一致すること",
                       _pick34(bk, ind, P34), DZ.vals("fig2-15_財政指数", nm),
                       tol=0.51, gp=P34, wp=P34)

        # 37 図2-16 地域支援事業費の内訳（百万円）
        P37 = [_pkey(c) for c in DZ.cats("fig2-13_地域支援事業費")]
        bad37 = []
        for nm, ind in (("一般介護予防事業費", "一般介護予防事業費"),
                        ("介護予防・生活支援サービス事業費", "介護予防・生活支援サービス事業費"),
                        ("包括的支援事業費", "包括的支援事業･任意事業")):
            got = _pick34("D48-c", ind, P37, 1000000.0)
            want = DZ.vals("fig2-13_地域支援事業費", nm)
            for p, g, w in zip(P37, got, want):
                if g is None:
                    bad37.append(f"{nm} {p} が見える化にない")
                elif abs(g - w) > 0.011:
                    bad37.append(f"{nm} {p} 見える化{g:.2f}≠正本{w}")
        # 内訳の和＋その他が合計に一致すること
        gokei = _pick34("D48-c", "合計", P37, 1000000.0)
        for i, p in enumerate(P37):
            uchi = sum(DZ.vals("fig2-13_地域支援事業費", k)[i] for k in
                       ("一般介護予防事業費", "介護予防・生活支援サービス事業費",
                        "包括的支援事業費", "その他"))
            if gokei[i] is not None and abs(uchi - gokei[i]) > 0.021:
                bad37.append(f"{p} 内訳の和{uchi:.2f}≠合計{gokei[i]:.2f}")
        chk(37, "地域支援事業費の内訳が見える化 D48-c と一致すること", not bad37,
            "／".join(bad37[:3]) if bad37
            else f"3系列×{len(P37)}点と、内訳の和が合計に一致")
        # 38・39 在宅／施設・居住系別の1人1月あたり給付月額（D6-a・D6-b）
        D6 = _c34.defaultdict(dict)
        with open(BAT34, encoding="utf-8") as f:
            for r in _csv.DictReader(f):
                if r["region"] != "北塩原村" or not r["file"].startswith("D6-"):
                    continue
                try:
                    D6[(r["file"].split("_")[0], r["indicator"])][_pkey(r["period"])] = \
                        float(r["value"])
                except ValueError:
                    pass
        P38 = [_pkey(c) for c in DZ.cats("fig2-11_在宅施設別給付月額")]

        def _p38(bk, kw):
            for (b, i), d in D6.items():
                if b == bk and kw in i:
                    return [d.get(p) for p in P38]
            return None
        cmp_series(38, "在宅サービスの給付月額が見える化 D6-a と一致すること",
                   _p38("D6-a", "在宅サービス"),
                   DZ.vals("fig2-11_在宅施設別給付月額", "在宅サービス"),
                   tol=0.51, gp=P38, wp=P38)
        cmp_series(39, "施設・居住系の給付月額が見える化 D6-b と一致すること",
                   _p38("D6-b", "施設および居住系"),
                   DZ.vals("fig2-11_在宅施設別給付月額", "施設・居住系サービス"),
                   tol=0.51, gp=P38, wp=P38)
    except Exception as e:
        chk(34, "村の決算による図が見える化 D47・D48 と一致すること", False,
            f"照合できない（{e}）")

    # ── 40・41　サービス区分別の図をD13（サービス別の給付月額）から組み直す ──
    #    D13は27サービスそれぞれの第1号被保険者1人あたり給付月額を持つ。
    #    図2-14の7区分はその和であり、1円のずれもなく再現できる。
    #    図2-10（費用額）は**給付費ではなく費用額**であり、自己負担を含むため
    #    給付費より約1割大きい。厳密には一致しないので、給付率の帯で見る。
    try:
        import collections as _c40
        import re as _re40
        BAT40 = os.path.join(BASE, "data", "mieruka_batch.csv")
        D13 = _c40.defaultdict(dict)
        B1N = {}

        def _nd40(p40):
            k = _pkey(p40)
            m = _re40.match(r"^(\d{4})/3$", k)
            return str(int(m.group(1)) - 1) if m else k

        with open(BAT40, encoding="utf-8") as f:
            for r in _csv.DictReader(f):
                if r["region"] != "北塩原村":
                    continue
                try:
                    v40 = float(r["value"])
                except ValueError:
                    continue
                if r["file"].startswith("D13-"):
                    svc = r["indicator"].replace(
                        "第１号被保険者１人あたり給付月額（", "").rstrip("）")
                    D13[svc][_pkey(r["period"])] = v40
                elif r["file"].startswith("B1_") and r["indicator"] == "第１号被保険者数":
                    B1N[_nd40(r["period"])] = v40

        HOUMON = ["訪問介護", "訪問入浴介護", "訪問看護", "訪問リハビリテーション",
                  "居宅療養管理指導", "定期巡回・随時対応型訪問介護看護",
                  "夜間対応型訪問介護"]
        TSUSHO = ["通所介護", "通所リハビリテーション", "認知症対応型通所介護",
                  "地域密着型通所介護"]
        TANKI = ["短期入所生活介護", "短期入所療養介護"]
        YOGU = ["福祉用具貸与", "特定福祉用具販売", "住宅改修"]
        SHIEN40 = ["介護予防支援・居宅介護支援"]
        KYOJU = ["特定施設入居者生活介護", "認知症対応型共同生活介護",
                 "地域密着型特定施設入居者生活介護", "小規模多機能型居宅介護",
                 "看護小規模多機能型居宅介護"]
        SHISETSU = ["介護老人福祉施設", "介護老人保健施設", "介護療養型医療施設",
                    "地域密着型介護老人福祉施設入所者生活介護", "介護医療院"]
        KU40 = [("訪問系", HOUMON), ("通所系", TSUSHO), ("短期入所", TANKI),
                ("福祉用具・住宅改修", YOGU), ("居宅介護支援", SHIEN40),
                ("居住系", KYOJU), ("施設系", SHISETSU)]

        def _wa(svcs, k40):
            return sum(D13[s].get(k40, 0) for s in svcs)

        # 40 図2-14（7区分の給付月額）
        #    この図は **系列が年度・区分が横軸** である。向きを取り違えると
        #    系列名で引けずに止まる（実際に一度取り違えた）。
        bad40 = []
        n40 = 0
        KUNAME = {k: v for k, v in KU40}
        for nendo40 in [x[0] for x in
                        next(z["series"] for z in DZ.ZU
                             if z["name"] == "fig2-14_サービス区分別変化")]:
            k40 = _pkey(nendo40)
            want40 = dict(zip(DZ.cats("fig2-14_サービス区分別変化"),
                              DZ.vals("fig2-14_サービス区分別変化", nendo40)))
            for ku, svcs in KU40:
                n40 += 1
                g = _wa(svcs, k40)
                w = want40.get(ku)
                if w is None:
                    bad40.append(f"{nendo40} に区分{ku}がない")
                elif abs(g - w) > 0.51:
                    bad40.append(f"{nendo40} {ku} 見える化の和{g:.0f}≠正本{w:.0f}")
        chk(40, "サービス区分別変化の7区分がD13の27サービスの和と一致すること", not bad40,
            "／".join(bad40[:3]) if bad40
            else f"2期×7区分（{n40}点）すべて一致（27サービスを7区分に足し上げ）")

        # 41 図2-10（費用額）は給付費に自己負担を加えたもの。給付率の帯で見る
        GRP41 = [("在宅サービス", HOUMON + TSUSHO + TANKI + YOGU + SHIEN40),
                 ("居住系サービス", KYOJU), ("施設サービス", SHISETSU)]
        LO, HI = 0.86, 0.94          # 給付率の取り得る幅
        bad41, kyufu = [], []
        for nm41, svcs in GRP41:
            want = DZ.vals("fig2-10_費用額内訳", nm41)
            for c41, w41 in zip(DZ.cats("fig2-10_費用額内訳"), want):
                k40 = _pkey(c41)
                n = B1N.get(k40)
                if n is None:
                    bad41.append(f"{c41} の第1号被保険者数が見える化にない")
                    continue
                g41 = _wa(svcs, k40) * n * 12 / 1e6
                if not w41:
                    continue
                ritsu = g41 / w41          # 給付費 ÷ 費用額 ＝ 給付率
                kyufu.append(ritsu)
                if not (LO <= ritsu <= HI):
                    bad41.append(f"{nm41} {c41} の給付率が{ritsu:.1%}（{LO:.0%}〜{HI:.0%}の外）")
        chk(41, "費用額内訳がD13の給付費と給付率で整合すること", not bad41,
            "／".join(bad41[:3]) if bad41
            else f"{len(kyufu)}点の給付率が{min(kyufu):.1%}〜{max(kyufu):.1%}"
                 f"（費用額は給付費に自己負担を加えたものであり、厳密には一致しない）")
    except Exception as e:
        chk(40, "サービス区分別の図が見える化 D13 と整合すること", False,
            f"照合できない（{e}）")

    # ── 32　正本のすべての図について、突合の状態が定められていること ──────
    #    突合していない図を黙って残さないための点検。
    #    突き合わせられない図は、なぜできないかを COVER に書く。
    try:
        zu32 = sorted(d["name"] for d in DZ.ZU)
        bad32 = [n for n in zu32 if n not in COVER]
        over32 = [n for n in COVER if n not in zu32]
        nasi = [n for n in zu32 if COVER.get(n, ("", ""))[0] == "できない"]
        chk(32, "正本のすべての図に突合の状態が定められていること",
            not bad32 and not over32,
            f"状態なし {bad32}／余り {over32}" if (bad32 or over32)
            else f"{len(zu32)}図のうち突合済み{len(zu32) - len(nasi)}図／"
                 f"見える化に相手がない{len(nasi)}図")
    except Exception as e:
        chk(32, "正本のすべての図に突合の状態が定められていること", False,
            f"照合できない（{e}）")

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

    # ── 72　居住系の受給者数の算出が、見える化の居住系受給者数と合うこと ────
    #    サービス種類別の受給者数の系列はないが、
    #      受給者数 ＝ 第1号被保険者1人あたり給付月額 × 第1号被保険者数
    #                 ÷ 受給者1人あたり給付月額
    #    により算出できる。この算出が正しいなら、居住系の各サービスの和は
    #    見える化の居住系受給者数（D1）と一致するはずである。
    #    算出をやめて按分に戻すと、ここが合わなくなる。
    try:
        import csv as _csv72, io as _io72, re as _re72, collections as _c72, statistics as _st72
        import paths as _P72
        gen = list(_csv72.DictReader(_io72.open(
            _os.path.join(_P72.DATA, "第10期_給付費のみ把握サービス.csv"),
            encoding="utf-8-sig")))
        san = {}
        for lab in ("令和5年度", "令和6年度", "令和7年度"):
            san[lab] = sum(float(r[lab + "実績(人/月・参考)"]) for r in gen if r["区分"] == "居住系")
        # 見える化 D1 の居住系受給者数（月次）を年度の平均にする
        bat = list(_csv72.DictReader(_io72.open(
            _os.path.join(_P72.DATA, "mieruka_batch.csv"), encoding="utf-8-sig")))
        tsuki = _c72.defaultdict(list)
        for r in bat:
            if r["region"] != "北塩原村" or not r["file"].startswith("D1_"):
                continue
            if r["indicator"] != "居住系受給者数":
                continue
            m = _re72.match(r"令和([0-9元]+)年([0-9]+)月", r["period"])
            if not m:
                continue
            y = int(m.group(1).replace("元", "1"))
            mo = int(m.group(2))
            nd = y if mo >= 4 else y - 1
            try:
                tsuki["令和%d年度" % nd].append(float(r["value"]))
            except (TypeError, ValueError):
                pass
        bad72, n72 = [], 0
        for lab in ("令和5年度", "令和6年度", "令和7年度"):
            if len(tsuki.get(lab, [])) != 12:
                bad72.append("%s の月次が12か月そろっていない" % lab)
                continue
            mie = _st72.mean(tsuki[lab])
            n72 += 1
            # 0.5人までの差を認める（給付費が0のサービスには残差を割り振れない）
            if abs(san[lab] - mie) > 0.5:
                bad72.append("%s：算出の和%.1f人 と 見える化%.1f人 の差が%.1f人"
                             % (lab, san[lab], mie, abs(san[lab] - mie)))
        # 和だけでは按分との差が小さく出るため、サービス単位でも算出と突き合わせる
        tidy = list(_csv72.DictReader(_io72.open(
            _os.path.join(_P72.DATA, "mieruka_tidy.csv"), encoding="utf-8-sig")))
        ix72 = {}
        for r in tidy:
            if r["region"] == "北塩原村":
                ix72.setdefault((r["code"], r["indicator"], r["year"]), r["value"])

        def _f72(code, ind, y):
            v = ix72.get((code, ind, y))
            try:
                return float(v)
            except (TypeError, ValueError):
                return None
        YR72 = {"2023": "令和5年度", "2024": "令和6年度", "2025": "令和7年度"}
        SV72 = [("特定施設入居者生活介護", "D13-q", "D17-k"),
                ("認知症対応型共同生活介護", "D13-w", "D17-q")]
        n_sv = 0
        for nm, c13, c17 in SV72:
            row = [r for r in gen if r["サービス種類"] == nm]
            if not row:
                bad72.append("%s が給付費のみ把握サービスの表にない" % nm)
                continue
            for y, lab in YR72.items():
                a13 = _f72(c13, "第１号被保険者１人あたり給付月額（%s）" % nm, y)
                a17 = _f72(c17, "受給者1人あたり給付月額（%s）" % nm, y)
                ni = _f72("D2", "第1号被保険者数", y)
                if not (a13 and a17 and ni):
                    continue
                n_sv += 1
                mach = round(a13 * ni / a17, 1)
                ari = float(row[0][lab + "実績(人/月・参考)"])
                if abs(mach - ari) > 0.15:
                    bad72.append("%s %s：表の%.1f人 と 算出の%.1f人 が違う"
                                 % (nm, lab, ari, mach))
        chk(72, "居住系の受給者数の算出が見える化と合うこと", not bad72,
            "・".join(bad72[:3]) if bad72
            else "サービス別%d件が算出と一致し、3か年とも和が見える化の居住系受給者数と"
                 "0.5人以内で一致する" % n_sv)
    except Exception as e:
        chk(72, "居住系の受給者数の算出が見える化と合うこと", False, "照合できない（%s）" % e)

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
