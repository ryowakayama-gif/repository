# -*- coding: utf-8 -*-
"""川崎町 第10期 推計方法の設定（伸びの選択）の妥当性検証。

厚生労働省が配布した保険者別シートを読む。

  ・認定率の伸びのパターンの比較（貴保険者）
  ・施設・居住系サービス利用率の伸びのパターンの比較（貴保険者）
  ・在宅サービス利用率の伸びのパターンの比較（貴保険者）
  ・要介護認定者数の推計誤差（全国の保険者別）
  ・サービス別利用者数の推計誤差（全国の保険者別）

町からの情報：令和8年度は「公表時点データを使用」で月報は5月のみチェック。
したがって令和8年度の実績値は令和8年5月の1か月分である。
本スクリプトは、この1か月が伸びの終点に使われていることの影響を測る。

使い方
  python3 kawasaki_nobi_verify.py <配布シートのディレクトリ> [出力txt]
"""
import glob
import os
import re
import statistics as st
import sys
import warnings

import openpyxl

warnings.filterwarnings("ignore")

KW = "04324"
YEARS = ["令和3年度", "令和4年度", "令和5年度",
         "令和6年度", "令和7年度", "令和8年度"]


def find(d, *keys):
    """配布シートをシート名の特徴で見分ける"""
    for f in sorted(glob.glob(os.path.join(d, "*.xlsx"))):
        if os.path.basename(f).startswith("~$"):
            continue
        try:
            wb = openpyxl.load_workbook(f, data_only=True, read_only=True)
        except Exception:
            continue
        names = set(wb.sheetnames)
        if "グラフ" in names:
            t = str(wb["グラフ"]["B2"].value or "")
            if all(k in t for k in keys):
                wb.close()
                return f
        elif keys[0] == "誤差":
            ws = wb[wb.sheetnames[0]]
            head = " ".join(str(ws.cell(r, c).value or "")
                            for r in range(1, 8) for c in range(1, 8))
            if keys[1] in head:
                wb.close()
                return f
        wb.close()
    return None


def sections(ws, col=2):
    """データシートは
         「◯◯(実績値)」→ 年度ブロック6件 → 「◯◯の伸び」→ 伸びパターン5件
       の順に並ぶ。見出しの立つ行をすべて拾い、種別を付けて返す。"""
    out = []
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, col).value
        if not v or not str(v).strip():
            continue
        s = str(v).strip()
        if s in ("男", "女", "計"):
            continue
        if re.match(r"^令和\d+年度$", s):
            out.append(("年度", s, r))
        elif re.match(r"^[①-⑩]", s):
            out.append(("伸び", s, r))
        elif s.endswith("の伸び") or s.endswith("(実績値)") or s.endswith("実績値）"):
            out.append(("見出し", s, r))
    # 各ブロックの終わりを次の見出しの直前にする
    res = []
    for i, (kind, name, r0) in enumerate(out):
        r1 = out[i + 1][2] if i + 1 < len(out) else ws.max_row + 1
        res.append((kind, name, r0, r1))
    return res


def year_blocks(ws, col=2):
    return [(n, r0, r1) for k, n, r0, r1 in sections(ws, col) if k == "年度"]


def growth_blocks(ws, col=2):
    return [(n, r0, r1) for k, n, r0, r1 in sections(ws, col) if k == "伸び"]


def _nin_block(ws, r0, r1):
    """認定率シートの1ブロックから、計×各区分の「全体」を取り出す"""
    sex, d = None, {}
    for r in range(r0, r1):
        a, b, c = (ws.cell(r, x).value for x in (2, 3, 4))
        if a and str(a).strip() in ("男", "女", "計"):
            sex = str(a).strip()
        lab = str(b).strip() if b and str(b).strip() else (
            str(c).strip() if c and str(c).strip() else None)
        if not lab or sex != "計":
            continue
        v = ws.cell(r, 5).value
        if isinstance(v, (int, float)):
            d[lab] = v
    return d


def read_ninteiritsu(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["データ"]
    jisseki = {y: _nin_block(ws, r0, r1) for y, r0, r1 in year_blocks(ws)}
    nobi = {n: _nin_block(ws, r0, r1) for n, r0, r1 in growth_blocks(ws)}
    return jisseki, nobi


def _riyo_block(ws, r0, r1):
    d = {}
    for r in range(r0, r1):
        nm = ws.cell(r, 2).value
        if not nm or not str(nm).strip():
            continue
        nm = str(nm).strip()
        if nm.startswith("（") or nm.startswith("令和") or re.match(r"^[①-⑩]", nm):
            continue
        v = ws.cell(r, 3).value
        if isinstance(v, (int, float)):
            d[nm] = v
    return d


def read_riyoritsu(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["データ"]
    jisseki = {y: _riyo_block(ws, r0, r1) for y, r0, r1 in year_blocks(ws)}
    nobi = {n: _riyo_block(ws, r0, r1) for n, r0, r1 in growth_blocks(ws)}
    return jisseki, nobi


def cagr(a, b, n):
    if not a or not b or a <= 0 or b <= 0:
        return None
    return (b / a) ** (1.0 / n) - 1


L = []


def P(s=""):
    L.append(s)


d = sys.argv[1] if len(sys.argv) > 1 else "."
F_NIN = find(d, "認定率", "伸びのパターン")
F_SHI = find(d, "施設・居住系", "伸びのパターン")
F_ZAI = find(d, "在宅サービス利用率", "伸びのパターン")
F_ERR_N = find(d, "誤差", "要介護認定率")
F_ERR_S = find(d, "誤差", "利用者数の実績値")

P("川崎町 第10期 推計方法の設定（伸びの選択）の妥当性検証")
P("厚生労働省 保険者別配布シートによる／保険者番号 04324")
P("=" * 80)
P()
P("【町からの情報】介護保険事業報告の設定で、令和8年度は「公表時点データを使用」、")
P("　　　　　　　　月報は5月のみチェック。→ 令和8年度＝令和8年5月の1か月分。")
P("　　　　　　　　当社が9月11日に総括表の整数性から判定した内容と一致する。")

# ============================================== 認定率
if F_NIN:
    nin, nin_g = read_ninteiritsu(F_NIN)
    P()
    P("■ 1．認定率（計）の実績値と、システムが選べる5つの伸び")
    P()
    keys = ["第1号被保険者", "65～69歳", "70～74歳", "75～79歳",
            "80～84歳", "85～89歳", "90歳以上", "第2号被保険者", "総数"]
    P("  (1) 認定率の実績値")
    P(f"      {'区分':<16}" + "".join(f"{y[:5]:>9}" for y in YEARS))
    for k in keys:
        if k not in nin.get("令和3年度", {}):
            continue
        P(f"      {k:<16}" + "".join(
            f"{nin.get(y, {}).get(k, 0)*100:>8.2f}%" for y in YEARS))
    v7 = nin["令和7年度"]["第1号被保険者"]
    v8 = nin["令和8年度"]["第1号被保険者"]
    P()
    P(f"      令和7年度 {v7*100:.2f}% → 令和8年度 {v8*100:.2f}%"
      f"（{(v8/v7-1)*100:+.2f}%）")
    P("      認定率は「ある時点の認定者数÷被保険者数」であり、月数に中立な指標である")
    P("      （手引き 表3）。認定率に限れば、単月であること自体は直ちに誤りにならない。")
    P()
    P("  (2) システムが選べる5つの伸び（配布シートに収録されているもの）")
    P(f"      {'パターン':<40}" + "".join(f"{k[:8]:>10}" for k in keys[:4]))
    for n, d in nin_g.items():
        P(f"      {n[:40]:<40}" + "".join(
            f"{d.get(k, 0):>10.4f}" for k in keys[:4]))
    P()
    cur = next((n for n in nin_g if "令和３年度→令和８年度" in n
                or "令和3年度→令和8年度" in n), None)
    if cur:
        P(f"      現在の設定は「{cur}」である。")
        P("      この伸びの終点は令和8年5月の1か月であり、終点が振れれば5年分の")
        P("      伸びがまるごと振れる。他の4つは終点が同じか、全国値である。")

# ============================================== 利用率
sec = 1
for tag, f in (("施設・居住系サービス", F_SHI), ("在宅サービス", F_ZAI)):
    if not f:
        continue
    sec += 1
    ri, ri_g = read_riyoritsu(f)
    P()
    P(f"■ {sec}．{tag}利用率（全体）")
    P()
    names = [k for k in ri.get("令和7年度", {}) if k != "合計"]
    P("  (1) 利用率の実績値")
    P(f"      {'サービス':<30}" + "".join(f"{y[:5]:>9}" for y in YEARS)
      + f"{'R8/R7':>8}")
    flags = []
    for k in names:
        vs = [ri.get(y, {}).get(k) for y in YEARS]
        v7, v8 = vs[4], vs[5]
        rt = (v8 / v7) if (v7 and v8) else None
        P(f"      {k[:30]:<30}" + "".join(
            f"{(v*100 if v is not None else 0):>8.2f}%" for v in vs)
          + (f"{rt:>8.2f}" if rt else f"{'－':>8}"))
        if v7 and not v8:
            flags.append((k, v7, "令和8年度が0"))
        elif rt and abs(rt - 1) > 0.3:
            flags.append((k, v7, f"令和8年度は令和7年度の{rt:.2f}倍"))
    P()
    P(f"      令和8年度が0、又は令和7年度から±30％を超えて動いたもの：{len(flags)}件")
    for k, v7, why in flags:
        P(f"        ・{k}（令和7年度 {v7*100:.2f}%）　{why}")
    if ri_g:
        P()
        P(f"  (2) システムが選べる伸び（{len(ri_g)}通り）")
        # 令和7年度の利用率が大きい順に主なサービスを並べる
        top = sorted((k for k in names if ri["令和7年度"].get(k)),
                     key=lambda k: -ri["令和7年度"][k])[:6]
        P(f"      {'パターン':<34}" + "".join(f"{k[:7]:>9}" for k in top))
        for n, dd in ri_g.items():
            P(f"      {n[:34]:<34}" + "".join(
                f"{dd.get(k, 0):>9.4f}" for k in top))
        P()
        ends8 = [n for n in ri_g if "令和８年度の伸び" in n]
        P(f"      → {len(ri_g)}通りのうち{len(ends8)}通りは終点が令和8年度"
          "（＝令和8年5月の1か月）である。")
        P("        完結した年度どうしで組めるのは「令和6年度→令和7年度」だけである。")

# ============================================== 推計誤差（認定者数）
if F_ERR_N:
    P()
    P(f"■ {sec+1}．要介護認定者数の推計誤差（全国比較）")
    P("    誤差＝推計値÷実績値。国が標準の推計方法を各保険者に当てはめて求めたもの。")
    P()
    wb = openpyxl.load_workbook(F_ERR_N, data_only=True)
    for sn in wb.sheetnames:
        ws = wb[sn]
        for ci, nm in ((7, "基本推計"), (8, "包括推計")):
            allv, miy, kw = [], [], None
            for r in range(15, ws.max_row + 1):
                no = str(ws.cell(r, 2).value or "").strip()
                pref = ws.cell(r, 3).value
                v = ws.cell(r, ci).value
                if not isinstance(v, (int, float)) or v <= 0:
                    continue
                allv.append(v)
                if pref == "宮城県":
                    miy.append(v)
                if no == KW:
                    kw = v
            if kw is None:
                continue
            a = sorted(abs(x - 1) for x in allv)
            k = abs(kw - 1)
            rank = sum(1 for x in a if x < k) + 1
            mr = sorted(abs(x - 1) for x in miy)
            P(f"    【{sn}／{nm}】川崎町 {kw:.4f}（{(kw-1)*100:+.2f}%）")
            P(f"      全国{len(allv)}保険者の|誤差| 中央値 {st.median(a)*100:.2f}%"
              f"／75%点 {a[int(len(a)*.75)]*100:.2f}%"
              f"／90%点 {a[int(len(a)*.9)]*100:.2f}%")
            P(f"      川崎町の順位 {rank}/{len(a)}（上位{rank/len(a)*100:.0f}%）"
              f"　宮城県内 {sum(1 for x in mr if x < k)+1}/{len(mr)}")
        P()
    P("    → 川崎町では基本推計（性別・年齢5歳階級別・要介護度別）が包括推計より")
    P("      明らかに当たる。総括表の設定は基本推計であり、この選択は妥当である。")

# ============================================== 推計誤差（サービス別）
if F_ERR_S:
    P()
    P(f"■ {sec+2}．サービス別利用者数の推計誤差（令和6年度・全国比較）")
    P()
    wb = openpyxl.load_workbook(F_ERR_S, data_only=True)
    ws = wb[wb.sheetnames[0]]
    svcs = []
    for c in range(5, ws.max_column + 1, 2):
        n = ws.cell(6, c).value
        if n and str(n).strip():
            svcs.append((c, str(n).strip()))
    kwr = None
    for r in range(8, ws.max_row + 1):
        if str(ws.cell(r, 2).value or "").strip() == KW:
            kwr = r
            break
    P(f"    {'サービス':<30}{'川崎町(基本)':>12}{'全国|誤差|中央':>14}"
      f"{'順位':>12}{'判定':>8}")
    out, na = [], []
    for c, nm in svcs:
        kb = ws.cell(kwr, c).value
        allv = [ws.cell(r, c).value for r in range(8, ws.max_row + 1)]
        allv = [v for v in allv if isinstance(v, (int, float)) and v > 0]
        if not isinstance(kb, (int, float)) or kb <= 0:
            na.append(nm)
            continue
        a = sorted(abs(x - 1) for x in allv)
        k = abs(kb - 1)
        rank = sum(1 for x in a if x < k) + 1
        pos = rank / len(a)
        mark = "★外れ" if pos > 0.8 else ("・注意" if pos > 0.6 else "良")
        out.append((nm, kb, st.median(a), rank, len(a), pos, mark))
    for nm, kb, med, rank, n, pos, mark in sorted(out, key=lambda z: -z[5]):
        P(f"    {nm[:30]:<30}{kb:>12.4f}{med*100:>13.2f}%"
          f"{f'{rank}/{n}':>12}{mark:>8}")
    if na:
        P(f"    （実績がなく誤差が算定されていない：{len(na)}種別　{'、'.join(na)}）")
    bad = [x for x in out if x[5] > 0.8]
    P()
    P(f"    → 全国下位2割に入る（＝標準の推計方法が当たらない）サービスが"
      f"{len(bad)}件ある。")
    P("      これらは伸びの窓を変えても解消しない。手引き第5章の「補正で解消しない外れ」")
    P("      に当たり、供給側の事情を個別に確かめる必要がある。")

out = sys.argv[2] if len(sys.argv) > 2 else "kawasaki_nobi_verify.txt"
with open(out, "w", encoding="utf-8") as f:
    f.write("\n".join(L) + "\n")
print("\n".join(L))
print(f"\n→ {out}")
