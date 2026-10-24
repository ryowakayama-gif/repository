# -*- coding: utf-8 -*-
"""資料編（付表1・付表2）の全数照合。

前身の verify_shiryo_appendix.py との違いは「取りこぼしを黙って飛ばさない」こと。
表の中のデータセルを1つ残らず

    一致 ／ 不一致 ／ 未照合（理由つき）

のいずれかに分類し、最後に「総セル数 ＝ 一致＋不一致＋未照合」が成立することを
突き合わせて表示する。提出用の照合記録として、走査漏れがないことを機械が示す。

使い方
  python3 verify_shiryo_full.py [出力txt]
"""
import re
import sys
import unicodedata
from decimal import Decimal, ROUND_HALF_UP

import openpyxl
from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

BASE = "/home/user/repository/kawasaki_project"
XL = f"{BASE}/04_調査・入力・分析/R8調査データ/川崎町_在宅介護実態調査_集計データ_20260807.xlsx"
DOC = f"{BASE}/09_元資料/R8調査データ/R8.9.9受領版/川崎町_資料編_R8.9.9版.docx"

MARK = ("〇", "○", "◯")


def pct1(k, n):
    """百分率を小数第1位で四捨五入（Python の round は偶数丸めのため使わない）"""
    if not n:
        return None
    return float((Decimal(k * 100) / Decimal(n)).quantize(
        Decimal("0.1"), rounding=ROUND_HALF_UP))


def norm(s):
    s = unicodedata.normalize("NFKC", str(s))
    s = re.sub(r"[\s　]+", "", s)
    s = re.sub(r"[（）()［］\[\]「」・,、。／/？?]", "", s)
    return s.replace("～", "-").replace("〜", "-").replace("~", "-")


# ---------------------------------------------------------------- 個票の読み込み
wb = openpyxl.load_workbook(XL, data_only=True)
ws = wb["データ"]
ROWS = [r for r in range(5, ws.max_row + 1)
        if ws.cell(r, 1).value is not None
        and str(ws.cell(r, 1).value).strip() != "計"]

hdr = {}
for c in range(1, ws.max_column + 1):
    g = ws.cell(1, c).value
    if g:
        hdr.setdefault(str(g).strip(), []).append(c)

SEX, AGE, CARE = hdr["性別"][0], hdr["年齢階級"][0], hdr["要介護状態区分"][0]


def qblock(sc):
    """先頭列 sc から、行4が 'n' に当たるまでを1設問のブロックとして切り出す"""
    cols, ncol = [], None
    for c in range(sc, ws.max_column + 1):
        k = ws.cell(4, c).value
        if k is None:
            break
        if str(k).strip() == "n":
            ncol = c
            break
        cols.append(c)
    return cols, ncol


QB = {}
for name, cs in hdr.items():
    if not re.match(r"^[AB]問", name):
        continue
    c = cs[0]
    while c <= ws.max_column:
        sub = str(ws.cell(2, c).value or "").strip()
        cols, ncol = qblock(c)
        if not cols:
            break
        QB.setdefault(name, []).append((sub, cols, ncol))
        c = (ncol or cols[-1]) + 1
        v = str(ws.cell(1, c).value or "").strip()
        if v and v != name:
            break

SVC = {
    "訪問介護（ホームヘルプ）": "訪問介護", "訪問入浴介護": "訪問入浴",
    "訪問看護": "訪問看護", "訪問リハビリテーション": "訪問リハ",
    "通所介護（デイサービス）": "通所介護", "通所リハビリテーション（デイケア）": "通所リハ",
    "夜間対応型訪問介護": "夜間対応型訪問", "定期巡回・随時対応型訪問介護看護": "定期巡回",
    "小規模多機能型居宅介護": "小多機", "看護小規模多機能型居宅介護": "看多機",
    "短期入所生活介護・療養介護（ショートステイ）": "ショートステイ",
    "居宅療養管理指導": "居宅療養管理",
}


def band(v):
    """資料編が使う3区分（本来の区分名は誤り。第10章の指摘を参照）"""
    v = str(v or "")
    if v in ("65歳未満", "65～69歳", "70～74歳"):
        return norm("65歳未満～74歳")
    if v in ("75～79歳", "80～84歳"):
        return norm("75～84歳")
    if v in ("85～89歳", "90歳以上"):
        return norm("85～94歳")
    return None


def sel(level):
    """区分ラベルから該当する個票の行番号を返す"""
    L = norm(level)
    if L in ("-", ""):
        return ROWS
    for c in (SEX, AGE, CARE):
        g = [r for r in ROWS if norm(ws.cell(r, c).value or "") == L]
        if g:
            return g
    if "／" in level or "/" in level:
        s, b = re.split("[／/]", level, 1)
        return [r for r in ROWS
                if norm(ws.cell(r, SEX).value or "") == norm(s)
                and band(ws.cell(r, AGE).value) == norm(b)]
    return None


def resolve_block(grp, rest):
    """キャプションの設問名から個票の設問ブロックを1つに決める"""
    if grp == "A問8":
        bl = [x for x in QB.get(grp, []) if x[0] == SVC.get(rest)]
    else:
        bl = QB.get(grp, [])
    return bl[0] if len(bl) == 1 else None


def map_columns(heads, l2c):
    """表の列見出しを個票の選択肢列に対応づける。
    完全一致→（見出しが「…」で切れている場合）切れる前までの前方一致 の順に試す。
    決められない列は None を返し、呼び出し側が「未照合」として計上する。"""
    out, how = [], []
    for h in heads:
        hn = norm(h)
        if hn in l2c:
            out.append(l2c[hn])
            how.append("完全一致")
            continue
        # 「…」「...」で切れている見出しは、切れる前までの前方一致で解決する
        stem = re.sub(r"[…\.]+$", "", hn)
        hit = [c for k, c in l2c.items() if k.startswith(stem)] if stem else []
        if len(hit) == 1:
            out.append(hit[0])
            how.append("前方一致")
        else:
            out.append(None)
            how.append(f"不能（候補{len(hit)}）")
    return out, how


# ---------------------------------------------------------------- 資料編の読み込み
doc = Document(DOC)


def blocks(d):
    for ch in d.element.body.iterchildren():
        if ch.tag == qn("w:p"):
            yield Paragraph(ch, d)
        elif ch.tag == qn("w:tbl"):
            yield Table(ch, d)


def ctext(c):
    return "\n".join(p.text.strip() for p in c.paragraphs if p.text.strip())


L = []          # 出力行
ok = bad = unv = na = 0
bad_list, unv_list, note, na_list = [], [], [], []
small = []      # n<10 かつ件数1〜2 のセル
hundred = []    # n<10 で 100.0% と表示しているセル
tbl_ledger = []
h1_truth = {}   # 付表1の正解値（付表2の全体行との突合に使う）
p2_total = {}   # 付表2の全体行の値
cross_rows = []  # 層合計 vs 全体
n_axis = 0       # 突き合わせた属性軸の数

cap = ""
for b in blocks(doc):
    if isinstance(b, Paragraph):
        if b.text.strip():
            cap = b.text.strip()
        continue

    grid = [[ctext(c) for c in r.cells] for r in b.rows]
    tag = cap.split("　")[0] if "　" in cap else cap[:8]

    # ------------------------------------------------ 付表1（単純集計）
    m1 = re.match(r"^付表1-\d+\s+(A問\d+|B問\d+)\s+(.*)$", cap)
    if m1:
        grp = m1.group(1)
        rest = re.sub(r"（複数回答）|（[^）]*選択可）", "", m1.group(2)).strip()
        blk = resolve_block(grp, rest)
        cells_total = sum(max(0, len(r) - 1) for r in grid[1:])
        if blk is None:
            unv += cells_total
            unv_list.append(f"{tag}：設問ブロックを特定できず（{cells_total}セル）")
            tbl_ledger.append((tag, cells_total, 0, 0, 0, cells_total))
            continue
        _, cols, ncol = blk
        l2c = {norm(ws.cell(3, c).value): c for c in cols}
        n_true = sum(1 for x in ROWS if ws.cell(x, ncol).value not in (None, "", 0))
        t_ok = t_bad = t_unv = t_na = 0
        for r in grid[1:]:
            lab = norm(r[0])
            if lab in ("有効回答数n", "有効回答数", "合計", "計", "n", "総数"):
                for cell in r[1:]:
                    mm = re.search(r"([0-9,]+)", cell)
                    if not mm:
                        # 有効回答数の行の割合欄は「－」。割合が定義されないための
                        # 意図的な空欄であり、照合の対象にならない
                        if norm(cell) in ("-", "", "—", "―"):
                            t_na += 1
                            na_list.append(f"{tag}「{r[0][:14]}」割合欄が「－」（意図的）")
                        else:
                            t_unv += 1
                            unv_list.append(f"{tag}「{r[0][:14]}」数値なし：[{cell[:20]}]")
                        continue
                    v = int(mm.group(1).replace(",", ""))
                    if v != n_true:
                        t_bad += 1
                        bad_list.append(f"{tag}「{r[0][:14]}」有効回答数 {v}→正解{n_true}")
                    else:
                        t_ok += 1
                continue
            c = l2c.get(lab)
            if c is None:
                stem = re.sub(r"[…\.]+$", "", lab)
                hit = [x for k, x in l2c.items() if k.startswith(stem)] if stem else []
                c = hit[0] if len(hit) == 1 else None
            if c is None:
                t_unv += max(0, len(r) - 1)
                unv_list.append(f"{tag}「{r[0][:20]}」選択肢を個票に対応づけられず")
                continue
            kt = sum(1 for x in ROWS if ws.cell(x, c).value in MARK)
            pt = pct1(kt, n_true)
            h1_truth[(grp, rest, lab)] = (kt, pt, n_true)
            for cell in r[1:]:
                s = cell.replace(" ", "").replace("件", "")
                mp = re.search(r"([0-9]+\.[0-9]+)\s*[%％]", s)
                mk = re.match(r"^([0-9,]+)\s*$", s)
                if mp:                                   # 割合セル
                    pr = float(mp.group(1))
                    if pt is None or abs(pr - pt) > 0.001:
                        t_bad += 1
                        bad_list.append(f"{tag}「{r[0][:16]}」割合 {pr}%→正解{pt}%")
                    else:
                        t_ok += 1
                elif mk:                                 # 件数セル
                    kr = int(mk.group(1).replace(",", ""))
                    if kr != kt:
                        t_bad += 1
                        bad_list.append(f"{tag}「{r[0][:16]}」件数 {kr}→正解{kt}")
                    else:
                        t_ok += 1
                else:
                    t_unv += 1
                    unv_list.append(f"{tag}「{r[0][:16]}」数値として読めず：[{cell[:20]}]")
        ok += t_ok
        bad += t_bad
        unv += t_unv
        na += t_na
        tbl_ledger.append((tag, cells_total, t_ok, t_bad, t_na, t_unv))
        continue

    # ------------------------------------------------ 付表2（属性別クロス）
    m2 = re.match(r"^付表2-\d+\s+(A問\d+|B問\d+)\s+(.*)$", cap)
    if not m2:
        note.append(f"付表の書式に合わないキャプション：{cap[:40]}")
        continue
    grp = m2.group(1)
    rest = re.sub(r"（\d/\d）|［.*$", "", m2.group(2)).strip()
    rest = re.sub(r"（複数選択可）|（[^）]*選択可）", "", rest).strip()
    blk = resolve_block(grp, rest)
    heads = grid[0][3:]
    cells_total = sum(1 + max(0, len(r) - 3) for r in grid[1:])
    if blk is None:
        unv += cells_total
        unv_list.append(f"{tag}：設問ブロックを特定できず（{cells_total}セル）")
        tbl_ledger.append((tag, cells_total, 0, 0, 0, cells_total))
        continue
    _, cols, ncol = blk
    l2c = {norm(ws.cell(3, c).value): c for c in cols}
    colmap, how = map_columns(heads, l2c)
    for h, w in zip(heads, how):
        if w != "完全一致":
            note.append(f"{tag} 列見出し「{h[:26]}」→ {w}で解決")

    t_ok = t_bad = t_unv = t_na = 0
    lay_n = {}
    axis_n = {}     # 属性軸 → その軸に属する層の n の合計
    axis = ""
    for r in grid[1:]:
        if r[0].strip():
            axis = r[0].strip()
        lev = (r[1] or "－").strip()
        g = sel(lev)
        if g is None:
            t_unv += 1 + max(0, len(r) - 3)
            unv_list.append(f"{tag}／区分「{lev}」を個票に対応づけられず")
            continue
        ntr = sum(1 for x in g if ws.cell(x, ncol).value not in (None, "", 0))
        # n セル
        mm = re.search(r"([0-9,]+)", r[2])
        if not mm:
            t_unv += 1
            unv_list.append(f"{tag}／{lev}：n が数値として読めず [{r[2][:16]}]")
        else:
            nr = int(mm.group(1).replace(",", ""))
            if nr != ntr:
                t_bad += 1
                bad_list.append(f"{tag}／{lev}：n {nr}→正解{ntr}")
            else:
                t_ok += 1
            lay_n[lev] = nr
            if norm(lev) != "-":
                axis_n[axis] = axis_n.get(axis, 0) + nr
        # 値セル
        for i, cell in enumerate(r[3:]):
            if i >= len(colmap) or colmap[i] is None:
                t_unv += 1
                unv_list.append(f"{tag}／{lev}／列{i+1}：列を対応づけられず")
                continue
            s = cell.replace(" ", "").replace("件", "")
            mm = re.match(r"^([0-9,]+)\s*[\(（]([0-9.]+)[%％][\)）]?", s)
            if not mm:
                t_unv += 1
                unv_list.append(f"{tag}／{lev}／{heads[i][:14]}：読めず [{cell[:20]}]")
                continue
            kr, pr = int(mm.group(1).replace(",", "")), float(mm.group(2))
            kt = sum(1 for x in g if ws.cell(x, colmap[i]).value in MARK)
            pt = pct1(kt, ntr)
            if ntr == 0:
                # 母数0の層。件数は照合できるが割合は定義できない（0÷0）
                if kr != kt:
                    t_bad += 1
                    bad_list.append(f"{tag}／{lev}／{heads[i][:14]}：件数{kr}→{kt}")
                else:
                    t_na += 1
                    na_list.append(
                        f"{tag}／{lev}(n=0)／{heads[i][:14]}：割合 {pr}% と表示（0÷0）")
                continue
            e = []
            if kr != kt:
                e.append(f"件数{kr}→{kt}")
            if pt is None or abs(pr - pt) > 0.001:
                e.append(f"割合{pr}%→{pt}%")
            if e:
                t_bad += 1
                bad_list.append(f"{tag}／{lev}／{heads[i][:14]}：" + " ".join(e))
            else:
                t_ok += 1
            if lev in ("全　体", "全体", "－") or norm(lev) == "-":
                p2_total[(grp, rest, norm(heads[i]))] = (kt, pt, ntr)
            if 0 < kt <= 2 and ntr < 10:
                small.append(f"{tag}／{lev}(n={ntr})／{heads[i][:14]}={kt}件")
            if ntr < 10 and pt == 100.0:
                hundred.append(f"{tag}／{lev}(n={ntr})／{heads[i][:14]}")
    ok += t_ok
    bad += t_bad
    unv += t_unv
    na += t_na
    tbl_ledger.append((tag, cells_total, t_ok, t_bad, t_na, t_unv))

    # 層合計 vs 全体（内部整合）。属性軸ごとに、その軸の層の n の合計を全体と比べる
    tot = next((v for k, v in lay_n.items() if norm(k) == "-"), None)
    for ax, s in axis_n.items():
        n_axis += 1
        if tot is not None and s != tot:
            cross_rows.append(f"{tag}／{ax}：層合計{s} ≠ 全体{tot}（差{tot - s}）")


# ---------------------------------------------------------------- 出力
L.append("川崎町 資料編（R8.9.9版）　全数照合の記録")
L.append(f"個票：{XL.split('/')[-1]}（{len(ROWS)}件）")
L.append(f"照合先：{DOC.split('/')[-1]}")
L.append("")
L.append("■ 走査したセルの内訳（表ごと）")
L.append(f"{'付表':<10}{'総セル':>7}{'一致':>7}{'不一致':>7}{'対象外':>7}{'未照合':>7}")
for t, c, o, b_, a_, u in tbl_ledger:
    L.append(f"{t:<10}{c:>7}{o:>7}{b_:>7}{a_:>7}{u:>7}")
tt = sum(x[1] for x in tbl_ledger)
L.append(f"{'合計':<10}{tt:>7}{ok:>7}{bad:>7}{na:>7}{unv:>7}")
L.append("")
L.append(f"■ 照合の完全性　総セル {tt} ＝ 一致 {ok} ＋ 不一致 {bad} "
         f"＋ 対象外 {na} ＋ 未照合 {unv}"
         f"　→ {'成立' if tt == ok + bad + na + unv else '★不成立（集計に漏れ）'}")
L.append("")

L.append(f"■ 不一致（個票と値が異なるもの） {bad} 件")
for x in bad_list:
    L.append("  ✗ " + x)
if not bad_list:
    L.append("  （なし）")
L.append("")

L.append(f"■ 照合の対象外 {na} 件（値の誤りではないが、表記の手当てを要するもの）")
kinds = {}
for x in na_list:
    k = "母数0の層に割合を表示（0÷0）" if "n=0" in x else "有効回答数の行の割合欄が「－」"
    kinds.setdefault(k, []).append(x)
for k, v in kinds.items():
    L.append(f"  ・{k}：{len(v)} 件")
    for x in v[:6]:
        L.append("      " + x)
    if len(v) > 6:
        L.append(f"      …ほか {len(v)-6} 件")
L.append("")

L.append(f"■ 未照合 {unv} 件")
for x in unv_list[:60]:
    L.append("  ? " + x)
if not unv_list:
    L.append("  （なし。上記を除く全セルを個票と突き合わせた）")
elif len(unv_list) > 60:
    L.append(f"  …ほか {len(unv_list)-60} 件")
L.append("")

L.append(f"■ 列見出しの解決方法に注意を要したもの {len(note)} 件")
for x in note:
    L.append("  ・" + x)
if not note:
    L.append("  （すべて完全一致で解決）")
L.append("")

L.append(f"■ 内部整合：付表2 の「層の n の合計」と「全体の n」")
L.append(f"  突き合わせた属性軸 {n_axis} 組／合計が全体と一致しない軸 {len(cross_rows)} 組")
for x in cross_rows[:20]:
    L.append("  △ " + x)
if len(cross_rows) > 20:
    L.append(f"  …ほか {len(cross_rows)-20} 組")
L.append("")

L.append("■ 付表1 と 付表2「全　体」行の突合")
n_cmp = n_dif = 0
for (grp, rest, lab), (kt, pt, n) in h1_truth.items():
    v = p2_total.get((grp, rest, lab))
    if v is None:
        continue
    n_cmp += 1
    if v[0] != kt:
        n_dif += 1
        L.append(f"  ✗ {grp} {rest} 「{lab[:20]}」付表1={kt} / 付表2全体={v[0]}")
L.append(f"  照合 {n_cmp} 項目／相違 {n_dif} 件")
L.append("")

L.append(f"■ 個人特定に関わるセル")
L.append(f"  n＜10 かつ件数1〜2：{len(small)} 箇所")
L.append(f"  n＜10 で 100.0% 表示：{len(hundred)} 箇所")

out = sys.argv[1] if len(sys.argv) > 1 else "chk_shiryo_全数照合.txt"
with open(out, "w", encoding="utf-8") as f:
    f.write("\n".join(L) + "\n")
print("\n".join(L[:200]))
print(f"\n→ {out}")
