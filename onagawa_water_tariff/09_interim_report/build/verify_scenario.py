# -*- coding: utf-8 -*-
"""試算条件切替計算表の数式を評価して検算する。

LibreOffice がこの環境で動かないため、formulas で数式を評価している。
    pip install formulas
    python3 verify_scenario.py
"""
import formulas, os, sys
fn = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'onagawa_scenario_calc.xlsx')
xl = formulas.ExcelModel().loads(fn).finish()
sol = xl.calculate()
def get(sheet, cell):
    for k,v in sol.items():
        if k.upper().endswith(f"]{sheet.upper()}'!{cell}"):
            try: return float(v.value[0,0])
            except Exception:
                try: return v.value[0,0]
                except Exception: return v
    return None
CHECKS = [
 ('02_総括原価','C14','総括原価 合計', 1175810),
 ('03_指標と改定率','C6','給水原価', 272.4),
 ('03_指標と改定率','C7','供給単価 税抜', 136.9),
 ('03_指標と改定率','C8','総括原価回収割合', 0.503),
 ('03_指標と改定率','C9','年間不足額', 116993),
 ('03_指標と改定率','C10','必要改定率', 1.990),
 ('03_指標と改定率','C11','料金不足額 定義③', -29488),
 ('03_指標と改定率','C13','赤字半減の改定率', 1.1248),
 ('03_指標と改定率','C14','適用改定率', 1.990),
 ('04_料金への影響','F8','③6〜10㎥ 現行', 1265),
 ('04_料金への影響','H8','③6〜10㎥ パターン1', 2517),
 ('04_料金への影響','J8','③6〜10㎥ 3-A', 1465),
 ('04_料金への影響','F14','年間収入 現行（百万円）', 118.2),
 ('04_料金への影響','H14','年間収入 パターン1', 235.2),
 ('04_料金への影響','J14','年間収入 3-A', 235.2),
 ('04_料金への影響','F15','回収割合 現行', 0.503),
 ('04_料金への影響','H15','回収割合 パターン1', 1.000),
 ('04_料金への影響','J15','回収割合 3-A', 1.000),
 ('04_料金への影響','F19','標準家庭 現行', 2497),
 ('04_料金への影響','H19','標準家庭 パターン1', 4969),
 ('04_料金への影響','J19','標準家庭 3-A', 4166),
 ('05_口径別料金表','D7','3-A 基本料金 20mm', 1396),
 ('06_ケース比較','C5','採用条件 総括原価', 1175810),
 ('06_ケース比較','F5','採用条件 必要改定率', 1.990),
 ('06_ケース比較','C6','繰入金控除なし 総括原価', 1448807),
 ('06_ケース比較','F6','繰入金控除なし 必要改定率', 2.452),
 ('06_ケース比較','C8','差引方式に戻す 総括原価', 1134492),
 ('06_ケース比較','C9','機械電気20年 総括原価', 1247275),
]
ng=0
for sh, cell, name, want in CHECKS:
    got = get(sh, cell)
    ok = isinstance(got,(int,float)) and abs(got-want) < max(abs(want)*0.002, 0.6)
    if not ok: ng+=1
    g = format(got,',.3f') if isinstance(got,(int,float)) else str(got)
    print(f'{"OK " if ok else "NG "} {name:28s} {sh}!{cell:5s} 期待 {want:>12,.3f}  実測 {g:>14}')
print('\n不一致', ng, '件')
sys.exit(1 if ng else 0)
