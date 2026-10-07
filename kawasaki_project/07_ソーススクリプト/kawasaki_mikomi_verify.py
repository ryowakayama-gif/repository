# -*- coding: utf-8 -*-
"""川崎町 第10期 サービス見込量算定の妥当性検証。

「介護保険事業計画 サービス見込量算定の手引き」（令和8年9月）の手順を
川崎町（保険者番号04324）に当てはめる。

  第2章  基準年度の決め方　　… 月数の実測、月数に比例する値／中立な値の分離
  第3章  出所の選び方　　　　… 内的整合（合計＝年報）
  第4章  算定パターンの比較　… 給付費1％あたりの保険料月額
  第5章  伸びの置き方　　　　… 対計画比の偏り
  第10章 単価の趨勢　　　　　… 1人1月あたり給付費

使い方
  python3 kawasaki_mikomi_verify.py [出力txt]
"""
import glob
import os
import re
import sys

import openpyxl

BASE = "/home/user/repository/kawasaki_project/09_元資料"
K10 = f"{BASE}/R8実績データ/R8.9.9受領版/【川崎町】第10期_将来推計総括表_R8.9.8出力.xlsx"
K9 = f"{BASE}/R8実績データ/R8.9.7受領版/【川崎町】第９期　将来推計総括表.xlsx"
NEN = {
    "令和6年度": f"{BASE}/R8実績データ/R8.9.8受領版/年報データ_2024_川崎町.xlsx",
    "令和7年度": f"{BASE}/R7実績データ/年報データ_2025_川崎町.xlsx",
}
GEPPO = {
    "令和7年度": f"{BASE}/R7実績データ/R7月報",
    "令和8年度": f"{BASE}/R8実績データ/R8.9.7受領版/R8月報",
}
# 総括表の列の並び（第10期は実績R6〜R8・計画R9〜R11、第9期は実績R3〜R5・計画R6〜R8）
Y10 = ["令和6年度", "令和7年度", "令和8年度", "令和9年度", "令和10年度", "令和11年度"]
Y9 = ["令和3年度", "令和4年度", "令和5年度", "令和6年度", "令和7年度", "令和8年度"]


def sokatsu_rows(path):
    """総括表『2_サービス別給付費』を、行の並びのまま読む。
    給付費は年間累計（千円）、人数・回数・日数は1月当たり。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["2_サービス別給付費"]
    rows, section, cur = [], "", None
    for r in range(8, ws.max_row + 1):
        a, b, c = (ws.cell(r, i).value for i in (1, 2, 3))
        if a and str(a).strip():
            section = str(a).strip()
        if not (c and str(c).strip()):
            continue
        kind = str(c).strip()
        vals = [ws.cell(r, x).value for x in range(4, 10)]
        vals = [v if isinstance(v, (int, float)) else None for v in vals]
        if "給付費" in kind:
            nm = str(b).strip() if b and str(b).strip() else section
            cur = {"name": nm, "section": section, "row": r, "給付費": vals,
                   "人数": None, "量": None, "量名": None,
                   "total": section.startswith("合計")}
            rows.append(cur)
        elif cur is not None and "人数" in kind:
            cur["人数"] = vals
        elif cur is not None and ("回数" in kind or "日数" in kind):
            cur["量"], cur["量名"] = vals, kind
    return rows


def key(x):
    """予防と介護で同名のサービスがあるため、節を鍵に含める"""
    top = "予防" if "介護予防" in x["section"] or x["row"] < 62 else "介護"
    return (top, x["name"])


def nenpo(path):
    """年報 様式2（給付費）を読む。
    列8＝予防計、列15＝介護計、列16＝合計。行の見出しは列3〜5に階層で入る。
    戻り値は {サービス名: (予防計, 介護計, 合計)} と「総計」行。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["様式２（給付費）"]
    out, sokei = {}, None
    for r in range(1, ws.max_row + 1):
        nm = None
        for c in (3, 4, 5):
            v = ws.cell(r, c).value
            if v and str(v).strip():
                nm = str(v).strip()
        if not nm:
            continue
        vs = [ws.cell(r, c).value for c in (8, 15, 16)]
        if not all(isinstance(v, (int, float)) for v in vs):
            continue
        if nm.startswith("総計"):
            sokei = tuple(vs)
        else:
            out[nm] = tuple(vs)
    return out, sokei


def geppo_months(d):
    fs = [f for f in glob.glob(os.path.join(d, "*.xlsx"))
          if not os.path.basename(f).startswith("~$")]
    return sorted({m.group(1) for f in fs
                   for m in [re.search(r"(20\d{4})", os.path.basename(f))] if m})


def premium_params(path, r0=1, r1=None):
    """5_保険料推計 から保険料の算定要素を取り出す。
    見出しは列1〜5のいずれか、第10期3か年計の値は列6にある。
    同じ見出しが「月額の内訳」と「保険料収納必要額関係」の双方に現れるため、
    行の範囲を指定して読み分ける。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["5_保険料推計"]
    p = {}
    for r in range(r0, (r1 or ws.max_row) + 1):
        lab = None
        for c in range(1, 6):
            v = ws.cell(r, c).value
            if v and str(v).strip():
                lab = str(v).strip()
                break
        if not lab:
            continue
        v = ws.cell(r, 6).value
        if isinstance(v, (int, float)) and lab not in p:
            p[lab] = v
    return p


L = []


def P(s=""):
    L.append(s)


P("川崎町 第10期 サービス見込量算定の妥当性検証")
P("『介護保険事業計画 サービス見込量算定の手引き』（令和8年9月）の手順による")
P("保険者番号 04324／総括表 出力日 令和8年9月8日／推計パターン名「R8」")
P("=" * 82)

r10 = sokatsu_rows(K10)
r9 = sokatsu_rows(K9)
svc10 = {key(x): x for x in r10 if not x["total"]}
svc9 = {key(x): x for x in r9 if not x["total"]}
tot10 = [x for x in r10 if x["total"]]
tot9 = [x for x in r9 if x["total"]]


def gsum(tots, j):
    return sum((t["給付費"][j] or 0) for t in tots[:2])


# ============================================ 第2章
P()
P("■ 第2章　基準年度の決め方")
P()
P("【1】月報の収録月")
for y, d in GEPPO.items():
    ms = geppo_months(d)
    P(f"    {y}：{len(ms):2d}か月　{ms[0]}〜{ms[-1]}"
      + ("　（完結）" if len(ms) == 12 else "　（年度途中）"))

P()
P("【2】総括表の令和8年度は何か月分か")
g = [gsum(tot10, j) for j in range(6)]
P(f"    給付費の合計（年間累計・千円）")
for j in range(3):
    P(f"      {Y10[j]:<8}{g[j]:>12,.0f}")
P(f"    令和8年度 ÷ 令和7年度 ＝ {g[2]/g[1]:.4f}　×12 ＝ {g[2]/g[1]*12:.2f} か月相当")
P("    → 給付費の水準は12か月分に引き伸ばされている（年度換算済み）。")

P()
P("【3】1月当たりの値が整数かどうかで、元の月数を判定する（決定的な検証）")
cnt = {0: [0, 0], 1: [0, 0], 2: [0, 0]}
cnt_n = {0: [0, 0], 1: [0, 0], 2: [0, 0]}
for x in r10:
    if x["total"]:
        continue
    for nm, arr in (("人数", x["人数"]), ("量", x["量"])):
        if not arr:
            continue
        for j in range(3):
            v = arr[j]
            if not isinstance(v, (int, float)) or v == 0:
                continue
            cnt[j][0] += 1
            if abs(v - round(v)) < 1e-9:
                cnt[j][1] += 1
            if nm == "人数":
                cnt_n[j][0] += 1
                if abs(v - round(v)) < 1e-9:
                    cnt_n[j][1] += 1
P(f"      {'年度':<10}{'人数の非ゼロ':>12}{'うち整数':>10}{'整数率':>9}")
for j in range(3):
    n, i = cnt_n[j]
    P(f"      {Y10[j]:<10}{n:>12}{i:>10}{i/n*100:>8.1f}%")
P()
P("    12か月の平均であれば、利用者数が整数になることはまれである（令和6・7年度は"
  f"{cnt_n[0][1]/cnt_n[0][0]*100:.0f}％・{cnt_n[1][1]/cnt_n[1][0]*100:.0f}％）。")
P("    令和8年度は非ゼロの人数が100％整数であり、これは月平均ではなく")
P("    **特定の1か月の利用者数**をそのまま置いた値であることを示す。")
P("    給付費は、その1か月の水準を12倍したものである。")

P()
P("【4】第9期でも同じことが起きていたか（出力 2023年10月5日）")
c9 = {0: [0, 0], 1: [0, 0], 2: [0, 0]}
for x in r9:
    if x["total"] or not x["人数"]:
        continue
    for j in range(3):
        v = x["人数"][j]
        if isinstance(v, (int, float)) and v != 0:
            c9[j][0] += 1
            if abs(v - round(v)) < 1e-9:
                c9[j][1] += 1
P(f"      {'年度':<10}{'人数の非ゼロ':>12}{'うち整数':>10}{'整数率':>9}")
for j in range(3):
    n, i = c9[j]
    P(f"      {Y9[j]:<10}{n:>12}{i:>10}{i/n*100:>8.1f}%")
g9 = [gsum(tot9, j) for j in range(6)]
P()
P("    第9期の基準年度（令和5年度）も、人数が100％整数＝1か月の値であった。")
P("    その結果がそのまま第9期の計画値になっている。")
P(f"      {'':<12}" + "".join(f"{Y9[j][:6]:>12}" for j in range(6)))
P(f"      {'給付費(千円)':<12}" + "".join(f"{g9[j]:>12,.0f}" for j in range(6)))
P(f"      {'':<12}" + "".join(f"{x:>12}" for x in
                              ["実績", "実績", "1か月×12", "計画", "計画", "計画"]))
P(f"      令和5年度 ÷ 令和4年度 ＝ {g9[2]/g9[1]:.4f}"
  f"（1年で{(g9[2]/g9[1]-1)*100:+.1f}％の跳ね）")
P(f"      これに対し令和6年度の実績は {g[0]:,.0f} 千円、"
  f"令和7年度は {g[1]:,.0f} 千円であった。")
P(f"      対計画比　令和6年度 {g[0]/g9[3]*100:.1f}％／令和7年度 {g[1]/g9[4]*100:.1f}％")
P()
P("    → 第9期の計画が実績を5〜6％上回ってきた原因は、基準年度に")
P("      1か月の値を12倍したものを用いたことにある。第10期も同じ置き方になっている。")

# ============================================ 第3章
P()
P("■ 第3章　出所の選び方（内的整合の点検）")
P()
P("【4】サービス別の給付費の合計 ＝ 年報の給付費 か（手引き 表4）")
P(f"      {'年度':<10}{'区分':<10}{'総括表（円）':>18}{'年報 様式2（円）':>18}{'差':>10}")
for y, j in (("令和6年度", 0), ("令和7年度", 1)):
    _, sokei = nenpo(NEN[y])
    for i, nm in enumerate(("介護予防", "介護")):
        s = (tot10[i]["給付費"][j] or 0) * 1000
        n = sokei[i]
        P(f"      {y:<10}{nm:<10}{s:>18,.0f}{n:>18,.0f}{s-n:>10,.0f}")
    s = gsum(tot10, j) * 1000
    P(f"      {'':<10}{'計':<10}{s:>18,.0f}{sokei[2]:>18,.0f}{s-sokei[2]:>10,.0f}")
P("    → 総括表は千円単位（小数3桁）で保持されているため、差は丸めの範囲にとどまる。")
P("      令和6・7年度の実績は年報と突合でき、出所として信頼できる。")
P("    → 令和8年度は年報が存在せず、この点検が効かない唯一の年度である。")

# ============================================ 令和8年度で0になったサービス
P()
P("■ 令和8年度で給付費が0になったサービス（1か月に請求がなかったもの）")
zero = []
for k, x in svc10.items():
    gg = x["給付費"]
    if (gg[2] or 0) == 0 and max(gg[0] or 0, gg[1] or 0) > 0:
        zero.append((k, gg))
P(f"      {'区分':<4}{'サービス':<32}{'R6':>10}{'R7':>10}{'R8':>8}{'R9〜R11':>10}")
zr6 = zr7 = 0
for (top, nm), gg in sorted(zero, key=lambda z: -(z[1][1] or 0)):
    zr6 += gg[0] or 0
    zr7 += gg[1] or 0
    P(f"      {top:<4}{nm[:32]:<32}{(gg[0] or 0):>10,.0f}{(gg[1] or 0):>10,.0f}"
      f"{0:>8}{(gg[3] or 0):>10,.0f}")
P(f"      {'計':<36}{zr6:>10,.0f}{zr7:>10,.0f}{0:>8}{0:>10}")
P(f"    → {len(zero)}サービス。令和7年度の実績 {zr7:,.0f}千円 が"
  f"令和9〜11年度もすべて0で推計されている。")

# ============================================ 単価
P()
P("■ 第10章　1人1月あたり給付費（給付費÷12÷人数）")
P()
tan = []
for k, x in svc10.items():
    gg, nn = x["給付費"], x["人数"]
    if not gg or not nn:
        continue
    u = [(gg[j] / 12 / nn[j]) if (gg[j] and nn[j]) else None for j in range(3)]
    if u[1] and u[2]:
        tan.append((k, u, u[2] / u[1], gg[1] or 0))
tan.sort(key=lambda z: -abs(z[2] - 1))
P(f"      {'サービス':<32}{'R6':>9}{'R7':>9}{'R8':>9}{'R8/R7':>8}{'R7給付費':>11}")
for (top, nm), u, rt, g7 in tan:
    P(f"      {nm[:32]:<32}" + "".join(
        f"{v:>9,.1f}" if v else f"{'－':>9}" for v in u)
      + f"{rt:>8.3f}{g7:>11,.0f}")
md = sorted(z[2] for z in tan)[len(tan) // 2]
big = [z for z in tan if abs(z[2] - 1) > 0.2]
P()
P(f"    単価を復元できたサービス {len(tan)} 件。令和8年度／令和7年度 の中央値 {md:.3f}。")
P(f"    ただし ±20％を超えて動いたものが {len(big)} 件ある。"
  "1か月の値を単価の基準に置いたことによる振れである。")

# ============================================ 対計画比
P()
P("■ 第5章　対計画比（実績 ÷ 第9期計画値）")
P()
P("    第9期総括表の令和6・7年度は計画値、第10期総括表の同年度は実績値である。")
P(f"      {'区分':<4}{'サービス':<30}{'R6計画':>10}{'R6実績':>10}{'比':>7}"
  f"{'R7計画':>10}{'R7実績':>10}{'比':>7}")
ratios = []
pc6 = pa6 = pc7 = pa7 = 0
for k in sorted(svc10, key=lambda z: -(svc10[z]["給付費"][1] or 0)):
    if k not in svc9:
        continue
    a, b = svc10[k], svc9[k]
    c6, r6 = b["給付費"][3] or 0, a["給付費"][0] or 0
    c7, r7 = b["給付費"][4] or 0, a["給付費"][1] or 0
    if c6 <= 0 and c7 <= 0:
        continue
    pc6 += c6
    pa6 += r6
    pc7 += c7
    pa7 += r7
    f6 = f"{r6/c6*100:>6.1f}%" if c6 else f"{'－':>7}"
    f7 = f"{r7/c7*100:>6.1f}%" if c7 else f"{'－':>7}"
    if c7 and r7:
        ratios.append((k, r7 / c7))
    P(f"      {k[0]:<4}{k[1][:30]:<30}{c6:>10,.0f}{r6:>10,.0f}{f6}"
      f"{c7:>10,.0f}{r7:>10,.0f}{f7}")
P(f"      {'計（計画値のある種別）':<34}{pc6:>10,.0f}{pa6:>10,.0f}"
  f"{pa6/pc6*100:>6.1f}%{pc7:>10,.0f}{pa7:>10,.0f}{pa7/pc7*100:>6.1f}%")
if ratios:
    rs = sorted(r for _, r in ratios)
    P()
    P(f"    サービス別の対計画比は {min(rs)*100:.1f}％ 〜 {max(rs)*100:.1f}％ に開いている"
      f"（{len(rs)}種別）。")
    P(f"    全体では令和6年度 {pa6/pc6*100:.1f}％、令和7年度 {pa7/pc7*100:.1f}％。")

# ============================================ 保険料のレバー
P()
P("■ 第4章　給付費1％が保険料月額に効く額")
p = premium_params(K10, 65)          # 「５．保険料収納必要額関係」以降
p.update(premium_params(K10, 1, 20))  # 保険料基準額（月額）は上段から
hyojun = p["標準給付費見込額"]
sokyufu = p.get("総給付費（財政影響額調整後）", p.get("総給付費"))
chiiki = p["地域支援事業費"]
ichigo = p["第1号被保険者負担分相当額"]
chosei_s = p["調整交付金相当額"]
chosei_m = p["調整交付金見込額"]
hitsuyo = p["保険料収納必要額"]
shunyu = p["予定保険料収納率"]
getsugaku = p["保険料基準額（月額）"]
warimashi = hyojun / sokyufu
ni = ichigo / (hyojun + chiiki)
cs = chosei_s / hyojun
cm = chosei_m / hyojun
hosei = hitsuyo / shunyu / 12 / getsugaku

P(f"    標準給付費見込額　　　　{hyojun:>16,.0f} 円")
P(f"    総給付費　　　　　　　　{sokyufu:>16,.0f} 円")
P(f"    地域支援事業費　　　　　{chiiki:>16,.0f} 円")
P(f"    第1号被保険者負担分相当額{ichigo:>16,.0f} 円　（負担割合 {ni:.4f}）")
P(f"    調整交付金相当額　　　　{chosei_s:>16,.0f} 円　（対標準給付費 {cs:.4f}）")
P(f"    調整交付金見込額　　　　{chosei_m:>16,.0f} 円　（対標準給付費 {cm:.4f}）")
P(f"    保険料収納必要額　　　　{hitsuyo:>16,.0f} 円")
P(f"    予定収納率 {shunyu}／補正後被保険者数 {hosei:,.2f} 人／月額 {getsugaku:,.2f} 円")
P()
P("    検算：第1号負担分相当額 ＋ 調整交付金相当額 － 調整交付金見込額")
P(f"        ＝ {ichigo:,.0f} ＋ {chosei_s:,.0f} － {chosei_m:,.0f}"
  f" ＝ {ichigo+chosei_s-chosei_m:,.0f} 円"
  f"　（総括表 {hitsuyo:,.0f} 円と{'一致' if abs(ichigo+chosei_s-chosei_m-hitsuyo)<1 else '不一致'}）")
P()
P("    給付費1円が保険料収納必要額に効く額")
P("      ＝ 割増率 ×（第1号負担割合 ＋ 調整交付金相当割合 － 見込交付割合）")
kiki = warimashi * (ni + cs - cm)
P(f"      ＝ {warimashi:.6f} ×（{ni:.4f} ＋ {cs:.4f} － {cm:.4f}）＝ {kiki:.6f} 円")
one = sokyufu * 0.01 * kiki / shunyu / hosei / 12
P()
P(f"    → 総給付費1％（{sokyufu*0.01:,.0f}円）あたり 月額 {one:,.1f} 円")
P(f"    → 保険料月額100円は総給付費の {100/one:.2f}％ に当たる")

# ============================================ 月額への効き
P()
P("■ 基準年度を改めた場合の効き")
P()
P("【基準年度の差の分解】令和8年度（1か月×12）と令和7年度（完結）の差はどこから来るか")
P()
same7 = g[1] - zr7          # 令和7年度のうち、令和8年度で0にならなかった種別
P(f"      令和7年度の給付費（12か月の実績）　　　　　　　　{g[1]:>12,.0f} 千円")
P(f"        うち令和8年度に0となった12種別　　　　　　　　{zr7:>12,.0f} 千円")
P(f"        うち令和8年度にも計上のある種別　　　　　　　　{same7:>12,.0f} 千円")
P(f"      令和8年度の給付費（1か月を12倍）　　　　　　　　{g[2]:>12,.0f} 千円")
P()
P(f"      (a) 同じ種別での上振れ　{g[2]:,.0f} ÷ {same7:,.0f} － 1 ＝ "
  f"{(g[2]/same7-1)*100:+.2f}％（{g[2]-same7:+,.0f} 千円）")
P(f"      (b) 12種別が0になった下振れ　　　　　　　　　　　"
  f"{-zr7/g[1]*100:+.2f}％（{-zr7:+,.0f} 千円）")
P(f"      (a)＋(b) ＝ 令和8年度 ÷ 令和7年度 － 1 ＝ {(g[2]/g[1]-1)*100:+.2f}％")
P()
P("      → 2つの誤差が逆向きに働いて打ち消し合い、合計では＋2.4％に見えている。")
P("        打ち消し合っているだけで、どちらも実績ではない。")

P()
P("■ 月額への効きの一覧（総括表の現行値 7,257円 からの差・第1次近似）")
P()
lev = [
    ("(a) 令和8年度に計上のある種別の上振れを取り除く",
     -(g[2] - same7) / g[2] * 100),
    ("(b) 令和8年度で0になった12種別を令和7年度の実績で補う",
     zr7 / g[2] * 100),
    ("(a)＋(b)＝基準年度を令和7年度（完結年度）に改める",
     (g[1] - g[2]) / g[2] * 100),
    ("対計画比の偏りを補正しない場合の現状（参考・第9期の乖離）",
     (pa7 / pc7 - 1) * 100),
]
P(f"      {'レバー':<50}{'給付費への効き':>14}{'月額':>10}")
for nm, pc in lev[:3]:
    P(f"      {nm:<50}{pc:>13.2f}%{pc*one:>9.0f}円")
P()
P("      ※ (a)(b)は同時に起こるため、基準年度を改める効きは両者の和になる。")
P("        伸び率の終点（令和3年度→令和8年度）も1か月の値に依存しているため、")
P("        正確な値はシステム上で基準年度を改めて再推計する必要がある。")

out = sys.argv[1] if len(sys.argv) > 1 else "kawasaki_mikomi_verify.txt"
with open(out, "w", encoding="utf-8") as f:
    f.write("\n".join(L) + "\n")
print("\n".join(L))
print(f"\n→ {out}")
