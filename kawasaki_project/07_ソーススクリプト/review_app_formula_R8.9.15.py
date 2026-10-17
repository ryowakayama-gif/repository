# -*- coding: utf-8 -*-
"""他メンバーのアプリの計算式を川崎町の実データに当てて、当社算定との差を測る。

アプリの式（HTMLの calcService より）
  intensity = Σ(量×月数) ÷ Σ(利用者×月数)         1人1月あたりの量
  rate      = Σ(利用者×月数) ÷ Σ(認定者×月数)     受給率（サービス単位・要介護度別ではない）
  users     = 将来認定者数 × rate × uptake
  volume    = users × intensity × intensityFactor
  unitCost  = Σ費用 ÷ Σ(量×月数)                   1回（日）あたり単価
  cost      = volume × 12 × unitCost               ← 量がないサービスは算定不能
"""
import importlib.util, io, contextlib, warnings
warnings.filterwarnings("ignore")
spec = importlib.util.spec_from_file_location(
    "m", "/home/user/repository/kawasaki_project/07_ソーススクリプト/project_mikomi_R8.9.15.py")
m = importlib.util.module_from_spec(spec)
with contextlib.redirect_stdout(io.StringIO()):
    spec.loader.exec_module(m)

ROWS, NIN = m.ROWS, m.NIN
YS = ["令和9年度", "令和10年度", "令和11年度"]

# ── 1. アプリの式で「給付費を自動算定できない」サービスの割合
has_q = {r["name"]: (r["R7量"] is not None) for r in ROWS}
tot = sum(r["R7給付費"] for r in ROWS)
na = sum(r["R7給付費"] for r in ROWS if r["R7量"] is None)
print("■ アプリの式（cost = volume×12×unitCost）で自動算定できるか（令和7年度実績・千円）")
print(f"  量（回・日）がある　→ 算定可　　{tot - na:>12,.0f}　{(tot - na) / tot:>7.1%}")
print(f"  量がない　　　　　→ 算定不可　{na:>12,.0f}　{na / tot:>7.1%}")
print("  ※算定不可の主なもの")
for r in sorted([x for x in ROWS if x["R7量"] is None],
                key=lambda x: -x["R7給付費"])[:10]:
    print(f"    {r['name']:<34}{r['R7給付費']:>10,.0f}")

# ── 2. 要介護度別 vs サービス単位（受給率の粒度）の差
#   サービス単位：rate = R7利用者計 ÷ R7認定者計（対象区分）、将来 = 将来認定者計 × rate
S_IDX, K_IDX = (0, 1), (2, 3, 4, 5, 6)
print("\n■ 受給率の粒度による差（総給付費・千円）")
print(f"{'年度':<10}{'要介護度別（当社）':>18}{'サービス単位（アプリ）':>22}{'差':>14}{'差率':>9}")
for y in YS:
    fine = sum(r[y + "給付費"] for r in ROWS)
    coarse = 0.0
    for r in ROWS:
        idx = S_IDX if r["区分"] == "予防" else K_IDX
        base_rec = sum(NIN["令和8年度"][i] for i in idx)
        fut_rec = sum(NIN[y][i] for i in idx)
        coarse += r["R7給付費"] * fut_rec / base_rec
    print(f"{y:<10}{fine:>18,.0f}{coarse:>22,.0f}{coarse - fine:>14,.0f}"
          f"{(coarse - fine) / fine:>9.2%}")

print("\n■ サービス別に見た差（令和9年度・千円・差の大きい順）")
print(f"{'サービス':<34}{'要介護度別':>12}{'サービス単位':>14}{'差':>10}")
diffs = []
for r in ROWS:
    idx = S_IDX if r["区分"] == "予防" else K_IDX
    base_rec = sum(NIN["令和8年度"][i] for i in idx)
    fut_rec = sum(NIN["令和9年度"][i] for i in idx)
    c = r["R7給付費"] * fut_rec / base_rec
    diffs.append((r["name"], r["令和9年度給付費"], c, c - r["令和9年度給付費"]))
for nm, f, c, d in sorted(diffs, key=lambda x: -abs(x[3]))[:12]:
    print(f"{nm:<34}{f:>12,.0f}{c:>14,.0f}{d:>10,.0f}")
