# -*- coding: utf-8 -*-
"""基準年度（利用率の分母）の取り方による差を測る。

  当社の公表値 ： 令和7年度実績 × 認定者数(年) ÷ 認定者数(令和8年度)
                  → 見える化システムの構造（令和8年度が基準年度）に合わせたもの
  アプリのM03  ： 令和7年度実績 × 認定者数(年) ÷ 認定者数(令和7年度)
                  → 利用率の分子（サービス実績）と分母（認定者数）を同じ年度でそろえたもの
"""
import importlib.util, io, contextlib, warnings
warnings.filterwarnings("ignore")
spec = importlib.util.spec_from_file_location(
    "m", "/home/user/repository/kawasaki_project/07_ソーススクリプト/project_mikomi_R8.9.15.py")
m = importlib.util.module_from_spec(spec)
with contextlib.redirect_stdout(io.StringIO()):
    spec.loader.exec_module(m)
ROWS, NIN, SV, kyu, proj = m.ROWS, m.NIN, m.SV, m.kyu, m.proj
YS = ["令和9年度", "令和10年度", "令和11年度"]
D7 = m.D7

for base in ("令和8年度", "令和7年度"):
    print(f"■ 分母＝{base} の伸び率")
    for y in YS:
        r = [NIN[y][i] / NIN[base][i] for i in range(7)]
        print(f"  {y:<10}" + "".join(f"{v:>9.3f}" for v in r))
print()

G7 = {}
for nm, krow, nyear, qyear, unit, kubun in SV:
    G7[nm] = kyu(krow)
S_IDX, K_IDX = (0, 1), (2, 3, 4, 5, 6)


def total(base, y):
    t = 0.0
    for nm, krow, nyear, qyear, unit, kubun in SV:
        g = G7[nm]
        rt = [NIN[y][i] / NIN[base][i] for i in range(7)]
        t += sum(g[i] * rt[i] for i in range(7))
    return t / 1000.0


print(f"{'年度':<10}{'分母R8（公表値）':>18}{'分母R7（利用率整合）':>22}{'差':>12}{'差率':>9}")
for y in YS:
    a, b = total("令和8年度", y), total("令和7年度", y)
    print(f"{y:<10}{a:>18,.0f}{b:>22,.0f}{b - a:>12,.0f}{(b - a) / a:>9.2%}")
a3 = sum(total("令和8年度", y) for y in YS)
b3 = sum(total("令和7年度", y) for y in YS)
print(f"{'3か年計':<10}{a3:>18,.0f}{b3:>22,.0f}{b3 - a3:>12,.0f}{(b3 - a3) / a3:>9.2%}")
print(f"\n保険料への影響の目安　{(b3 - a3) * 1000 * 0.2211 / (0.96 * 12 * 9611 * 0.964047):+,.1f} 円/月")

print("\n■ サービス別の差（令和9年度・千円・差の大きい順）")
print(f"{'サービス':<36}{'分母R8':>12}{'分母R7':>12}{'差':>10}")
d = []
for nm, krow, nyear, qyear, unit, kubun in SV:
    g = G7[nm]
    for idx in (K_IDX, S_IDX):
        a = sum(g[i] * NIN["令和9年度"][i] / NIN["令和8年度"][i] for i in idx) / 1000
        b = sum(g[i] * NIN["令和9年度"][i] / NIN["令和7年度"][i] for i in idx) / 1000
        if a == 0 and b == 0:
            continue
        lab = nm if idx == K_IDX else "（予防）" + nm
        d.append((lab, a, b, b - a))
for lab, a, b, x in sorted(d, key=lambda z: -abs(z[3]))[:12]:
    print(f"{lab:<36}{a:>12,.0f}{b:>12,.0f}{x:>10,.0f}")
