# -*- coding: utf-8 -*-
"""所得段階別第1号被保険者数（13段階・令和7年度実績）を反映した保険料の再算定。

出所：介護保険事業状況報告（年報・令和7年度／様式1 所得段階別・年度末現在被保険者数）
"""
import importlib.util, io, contextlib, warnings
warnings.filterwarnings("ignore")
spec = importlib.util.spec_from_file_location(
    "m", "/home/user/repository/kawasaki_project/07_ソーススクリプト/project_mikomi_R8.9.15.py")
M = importlib.util.module_from_spec(spec)
with contextlib.redirect_stdout(io.StringIO()):
    spec.loader.exec_module(M)

YEARS, HIHO3 = M.YEARS, M.hihoken3
SHUNO3 = sum(M.FIN[y]["保険料収納必要額"] for y in YEARS)
RITSU = [0.455, 0.685, 0.690, 0.90, 1.00, 1.20, 1.30, 1.50, 1.70,
         1.90, 2.10, 2.30, 2.40]
R6 = [412, 308, 304, 383, 699, 375, 419, 218, 67, 25, 13, 7, 31]
R7 = [383, 298, 305, 343, 679, 410, 427, 232, 85, 26, 17, 7, 28]
FUND = 152_500_000
H_NOW = M.HOSEI          # 見える化システムに登録されている補正係数（0.964047）


def hosei(w):
    s = sum(w)
    return sum(n * r for n, r in zip(w, RITSU)) / s


def premium(h, torikuzushi=0):
    return SHUNO3 / 0.96 / 12 / (HIHO3 * h) - torikuzushi / (12 * HIHO3 * h)


CASES = [("現在の登録値（第10〜13段階が0人）", H_NOW),
         ("令和6年度の実績分布（13段階）", hosei(R6)),
         ("令和7年度の実績分布（13段階）", hosei(R7))]
print("■ 所得段階別加入割合による補正係数と保険料")
print(f"{'分布':<34}{'補正係数':>10}{'補正後被保険者数':>16}"
      f"{'A 取崩なし':>13}{'B 50%取崩':>12}{'C 全額取崩':>12}")
for nm, h in CASES:
    print(f"{nm:<34}{h:>10.6f}{HIHO3 * h:>16,.1f}"
          f"{premium(h):>13,.0f}{premium(h, FUND * .5):>12,.0f}"
          f"{premium(h, FUND):>12,.0f}")

h0, h7 = H_NOW, hosei(R7)
print(f"\n現在の登録値 → 令和7年度の実績分布に改めた場合の差"
      f"　{premium(h7) - premium(h0):+,.1f} 円／月")
print(f"第9期（6,500円）との差　A {premium(h7) - 6500:+,.0f}円／"
      f"B {premium(h7, FUND * .5) - 6500:+,.0f}円／"
      f"C {premium(h7, FUND) - 6500:+,.0f}円")
print(f"第9期の取崩前（7,206.92円）との差　A {premium(h7) - 7206.92:+,.0f}円")
print("\n■ 令和7年度の13段階分布（年報・様式1所得段階別）")
print(f"{'段階':<8}{'割合':>7}{'人数':>7}{'構成比':>9}{'加重':>9}")
for i, (n, r) in enumerate(zip(R7, RITSU), 1):
    print(f"第{i}段階{'':<3}{r:>7.3f}{n:>7,}{n / sum(R7) * 100:>8.1f}%{n * r:>9.1f}")
print(f"{'計':<8}{'':>7}{sum(R7):>7,}{100.0:>8.1f}%{sum(n * r for n, r in zip(R7, RITSU)):>9.1f}")
