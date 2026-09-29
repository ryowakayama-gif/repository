# -*- coding: utf-8 -*-
"""エクセルの成果品について、図（グラフ・画像）と表の関係を機械で確かめる。

使い方
    python3 tools/xlsx_zuhyo_check.py                 output の xlsx をすべて
    python3 tools/xlsx_zuhyo_check.py <xlsx> [<xlsx>] 指定したものだけ
    python3 tools/xlsx_zuhyo_check.py --csv <path>    結果をCSVにも書き出す

確かめること
    L1 重なり      グラフ・画像が置かれた範囲に、値の入ったセルがないか
    L2 はみ出し    グラフがシートの右端・下端を大きく越えていないか
    L3 参照範囲    系列の参照が見出しの行・列を巻き込んでいないか
    L4 合計の巻込み 系列の参照に「合計」「計」の行・列が入っていないか
    L5 長さ        カテゴリの数と各系列の値の数が一致するか
    L6 中身        参照先に数値でないセル・空のセルが混じっていないか
    L7 内訳と合計  積上げグラフの内訳の和が、同じ表の合計の欄と一致するか

判定
    不適合  直さなければ送付・公表できない（L1・L3・L4・L5）
    要判断  文脈により適否が分かれる（L2・L6・L7）
    未実施  読めなかったもの（理由を残す。適合に丸めない）

不適合が1件でもあると終了コード1で終わる。

列の幅・行の高さからの位置の計算は目安である（フォントにより実際の描画は動く）。
**重なりは「無い」と出たことを根拠にせず、紙面でも確かめる。**
"""

import csv
import os
import re
import sys

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, range_boundaries

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

EMU_PX = 9525           # 1画素（96dpi）あたりのEMU
EMU_PT = 12700          # 1ポイントあたりのEMU
DEF_COL = 8.43          # 既定の列の幅（文字数）
DEF_ROW = 18.0          # 既定の行の高さ（ポイント）

FUTEKI = "不適合"
YOHAN = "要判断"
MIJISSHI = "未実施"

RESULTS = []            # (ファイル, シート, 点検, 判定, 内容)


def add(f, s, ten, han, naiyo):
    RESULTS.append((os.path.basename(f), s, ten, han, naiyo))


# ------------------------------------------------------------ 位置の計算
def col_px(ws, i):
    """0起点の列番号 i の幅（画素）。"""
    d = ws.column_dimensions.get(get_column_letter(i + 1))
    w = d.width if (d is not None and d.width) else (
        ws.sheet_format.defaultColWidth or DEF_COL)
    return int(round(w * 7)) + 5


def row_pt(ws, i):
    """0起点の行番号 i の高さ（ポイント）。"""
    d = ws.row_dimensions.get(i + 1)
    h = d.height if (d is not None and d.height) else (
        ws.sheet_format.defaultRowHeight or DEF_ROW)
    return h or DEF_ROW


def spread(ws, anchor):
    """図が占めるセルの範囲を (列の始め, 行の始め, 列の終り, 行の終り) で返す。

    いずれも0起点・終りを含む。二点指定の錨は指定をそのまま使い、
    一点指定の錨は大きさ（EMU）から列の幅・行の高さを足して求める。
    """
    fr = getattr(anchor, "_from", None)
    if fr is None:
        return None
    c0, r0 = fr.col, fr.row
    to = getattr(anchor, "to", None)
    if to is not None:
        return c0, r0, to.col, to.row
    ext = getattr(anchor, "ext", None)
    if ext is None:
        return None
    c, used = c0, 0
    while used < ext.cx and c < c0 + 200:
        used += col_px(ws, c) * EMU_PX
        c += 1
    r, used = r0, 0
    while used < ext.cy and r < r0 + 2000:
        used += row_pt(ws, r) * EMU_PT
        r += 1
    return c0, r0, max(c - 1, c0), max(r - 1, r0)


def zu_of(ws):
    """シートの図（グラフ・画像）を (種別, 表題, 錨) で返す。"""
    out = []
    for ch in getattr(ws, "_charts", []):
        try:
            t = ch.title.tx.rich.p[0].r[0].t
        except Exception:
            t = None
        out.append(("グラフ", t or "（表題なし）", ch.anchor, ch))
    for im in getattr(ws, "_images", []):
        out.append(("画像", getattr(im, "path", None) or "（画像）",
                    im.anchor, None))
    return out


# ------------------------------------------------------------ 参照範囲を読む
REF = re.compile(r"^'?([^'!]+)'?!(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)$")


def cells_of(wb, ref):
    """参照の文字列から (シート, 範囲, セルの一覧) を返す。"""
    m = REF.match(ref or "")
    if not m:
        return None
    sn, rng = m.group(1), m.group(2).replace("$", "")
    if sn not in wb.sheetnames:
        return None
    ws = wb[sn]
    c = ws[rng]
    if not isinstance(c, tuple):
        c = ((c,),)
    elif c and not isinstance(c[0], tuple):
        c = tuple((x,) for x in c)
    return ws, rng, [x for row in c for x in row]


def series_of(wb, ch):
    """グラフの系列を (名, 名の参照, 値の参照, 値のセル, カテゴリのセル) で返す。"""
    out = []
    for ser in ch.series:
        nm, nmref = None, None
        if ser.tx is not None and ser.tx.strRef is not None:
            nmref = ser.tx.strRef.f
            got = cells_of(wb, nmref)
            nm = got[2][0].value if got and got[2] else None
        vref = (ser.val.numRef.f
                if (ser.val is not None and ser.val.numRef is not None) else None)
        vc = cells_of(wb, vref) if vref else None
        cc = None
        if ser.cat is not None:
            cr = ser.cat.numRef or ser.cat.strRef
            if cr:
                cc = cells_of(wb, cr.f)
        out.append((nm, nmref, vref, vc, cc))
    return out


GOKEI = re.compile(r"^\s*(合計|計|総計|小計|総合計|合\s*計)\s*$")


def is_formula(v):
    return isinstance(v, str) and v.startswith("=")


def is_head(v):
    """見出しらしいセルか。式のセルは値が保存されていないだけなので除く。"""
    if v is None or isinstance(v, (int, float)) or is_formula(v):
        return False
    return True


# ------------------------------------------------------------ 点検
def check_book(path):
    try:
        wb = load_workbook(path)
    except Exception as e:
        add(path, "―", "読み込み", MIJISSHI, "開けなかった：%s" % str(e)[:60])
        return
    for sn in wb.sheetnames:
        ws = wb[sn]
        zus = zu_of(ws)
        # ---- L1 重なり／L2 はみ出し
        used_c = (ws.max_column or 1) - 1
        used_r = (ws.max_row or 1) - 1
        for shu, mei, anc, ch in zus:
            sp = spread(ws, anc)
            if sp is None:
                add(path, sn, "L1 重なり", MIJISSHI,
                    "%s「%s」の錨を読めなかった" % (shu, mei))
                continue
            c0, r0, c1, r1 = sp
            atari = []
            for r in range(r0, min(r1, used_r) + 1):
                for c in range(c0, min(c1, used_c) + 1):
                    v = ws.cell(row=r + 1, column=c + 1).value
                    if v not in (None, ""):
                        atari.append("%s%d" % (get_column_letter(c + 1), r + 1))
            if atari:
                add(path, sn, "L1 重なり", FUTEKI,
                    "%s「%s」（%s%d〜%s%d）が値のあるセル %d件と重なる（%s ほか）"
                    % (shu, mei, get_column_letter(c0 + 1), r0 + 1,
                       get_column_letter(c1 + 1), r1 + 1, len(atari),
                       "・".join(atari[:5])))
            if c1 > used_c + 40 or r1 > used_r + 200:
                add(path, sn, "L2 はみ出し", YOHAN,
                    "%s「%s」が表の範囲を大きく越える（右端 %s／下端 %d行）"
                    % (shu, mei, get_column_letter(c1 + 1), r1 + 1))
        # ---- L3〜L7 グラフの参照
        for shu, mei, anc, ch in zus:
            if ch is None:
                continue
            sers = series_of(wb, ch)
            if not sers:
                add(path, sn, "L3 参照範囲", MIJISSHI,
                    "グラフ「%s」の系列を読めなかった" % mei)
                continue
            ncat = None
            gk_done = False
            for nm, nmref, vref, vc, cc in sers:
                lab = nm or nmref or "（名なし）"
                if vc is None:
                    add(path, sn, "L3 参照範囲", MIJISSHI,
                        "グラフ「%s」系列［%s］の値の参照を読めなかった（%s）"
                        % (mei, lab, vref))
                    continue
                vws, vrng, vcells = vc
                vals = [x.value for x in vcells]
                # L3 見出しの巻込み
                atama = [x for x in vcells[:1] if is_head(x.value)]
                if atama:
                    add(path, sn, "L3 参照範囲", FUTEKI,
                        "グラフ「%s」系列［%s］の値の参照 %s が見出しのセルから始まる"
                        % (mei, lab, vrng))
                # L4 合計の巻込み
                if not gk_done and cc is not None:
                    gk = [x.coordinate for x in cc[2]
                          if isinstance(x.value, str) and GOKEI.match(x.value)]
                    if gk:
                        gk_done = True
                        add(path, sn, "L4 合計の巻込み", YOHAN,
                            "グラフ「%s」のカテゴリに合計の行・列がある（%s）。"
                            "内訳と並べて示す意図であれば差し支えないが、"
                            "内訳の合計として二重に数えていないかを見る"
                            % (mei, "・".join(gk)))
                # L6 中身
                kara = sum(1 for v in vals if v is None)
                shiki = sum(1 for v in vals if is_formula(v))
                moji = [str(v)[:12] for v in vals
                        if isinstance(v, str) and not is_formula(v)]
                if moji:
                    add(path, sn, "L6 中身", YOHAN,
                        "グラフ「%s」系列［%s］の参照に数値でないセルがある（%s）"
                        % (mei, lab, "・".join(moji[:4])))
                if shiki:
                    add(path, sn, "L6 中身", MIJISSHI,
                        "グラフ「%s」系列［%s］の参照は式のセル %d件であり、"
                        "値が保存されていないため中身を確かめられない"
                        % (mei, lab, shiki))
                if kara and kara == len(vals):
                    add(path, sn, "L6 中身", YOHAN,
                        "グラフ「%s」系列［%s］の参照がすべて空である" % (mei, lab))
                # L5 長さ
                n = len(vals)
                if cc is not None:
                    nc = len(cc[2])
                    if nc != n:
                        add(path, sn, "L5 長さ", FUTEKI,
                            "グラフ「%s」系列［%s］はカテゴリ %d件に対し値 %d件"
                            % (mei, lab, nc, n))
                if ncat is None:
                    ncat = n
                elif ncat != n:
                    add(path, sn, "L5 長さ", FUTEKI,
                        "グラフ「%s」の系列ごとに値の数が違う（%d件と %d件）"
                        % (mei, ncat, n))
        # ---- L7 内訳と合計
        for shu, mei, anc, ch in zus:
            if ch is None or getattr(ch, "grouping", None) not in ("stacked",
                                                                   "percentStacked"):
                continue
            sers = series_of(wb, ch)
            cols = []
            for nm, nmref, vref, vc, cc in sers:
                if vc is None:
                    cols = []
                    break
                cols.append([x.value for x in vc[2]])
            if not cols or len({len(c) for c in cols}) != 1:
                continue
            wa = [sum(c[i] for c in cols if isinstance(c[i], (int, float)))
                  for i in range(len(cols[0]))]
            # 同じ表の中に合計の行・列があれば突き合わせる
            got = cells_of(wb, sers[0][2])
            if got is None:
                continue
            vws, vrng, _ = got
            mn_c, mn_r, mx_c, mx_r = range_boundaries(vrng)
            hit = None
            for r in range(1, (vws.max_row or 1) + 1):
                v = vws.cell(row=r, column=mn_c - 1 if mn_c > 1 else 1).value
                if isinstance(v, str) and GOKEI.match(v) and r not in range(mn_r, mx_r + 1):
                    hit = r
                    break
            if hit is None:
                continue
            gvals = [vws.cell(row=hit, column=c).value
                     for c in range(mn_c, mx_c + 1)]
            if len(gvals) != len(wa):
                continue
            zure = [i for i, (a, b) in enumerate(zip(wa, gvals))
                    if isinstance(b, (int, float)) and abs(a - b) > 0.005]
            if zure:
                add(path, sn, "L7 内訳と合計", YOHAN,
                    "グラフ「%s」の内訳の和が %d行目の合計と合わない（%d箇所）"
                    % (mei, hit, len(zure)))


def main(argv):
    csv_out = None
    if "--csv" in argv:
        i = argv.index("--csv")
        csv_out = argv[i + 1]
        argv = argv[:i] + argv[i + 2:]
    files = argv or sorted(
        os.path.join("output", f) for f in os.listdir("output")
        if f.endswith(".xlsx") and not f.startswith("~$"))
    if not files:
        print("点検する xlsx がない（未実施）")
        return 2
    for f in files:
        check_book(f)
    # 同じ内容の重複を落とす（系列ごとに同じことを言うものがある）
    seen, uniq = set(), []
    for x in RESULTS:
        if x in seen:
            continue
        seen.add(x)
        uniq.append(x)
    RESULTS[:] = uniq
    n_f = sum(1 for x in RESULTS if x[3] == FUTEKI)
    n_y = sum(1 for x in RESULTS if x[3] == YOHAN)
    n_m = sum(1 for x in RESULTS if x[3] == MIJISSHI)
    print("点検したブック %d件" % len(files))
    for han in (FUTEKI, YOHAN, MIJISSHI):
        rows = [x for x in RESULTS if x[3] == han]
        if not rows:
            continue
        print("\n==== %s %d件" % (han, len(rows)))
        for f, s, ten, _, naiyo in rows[:60]:
            print("  [%s] %s／%s　%s" % (ten, f, s, naiyo))
        if len(rows) > 60:
            print("  … ほか %d件" % (len(rows) - 60))
    print("\n不適合 %d件／要判断 %d件／未実施 %d件" % (n_f, n_y, n_m))
    print("※ 位置の計算は列の幅・行の高さからの目安である。"
          "重なりが「無い」ことは紙面でも確かめる。")
    if csv_out:
        with open(csv_out, "w", encoding="utf-8-sig", newline="") as fp:
            w = csv.writer(fp)
            w.writerow(["ファイル", "シート", "点検", "判定", "内容"])
            w.writerows(RESULTS)
        print("書き出し", csv_out)
    return 1 if n_f else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
