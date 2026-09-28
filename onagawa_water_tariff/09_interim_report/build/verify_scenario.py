# -*- coding: utf-8 -*-
"""試算条件切替計算表の数式を評価して検算する。

LibreOffice がこの環境で動かないため、formulas で数式を評価している。
    pip install formulas
    python3 verify_scenario.py
"""
import os, sys, shutil, tempfile
import formulas, openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, '..', 'onagawa_scenario_calc.xlsx')


def solve(changes=None):
    path = SRC
    if changes:
        tmp = os.path.join(tempfile.gettempdir(), 'scenario_case.xlsx')
        shutil.copy(SRC, tmp)
        wb = openpyxl.load_workbook(tmp)
        for cell, v in changes.items():
            wb['01_設定'][cell] = v
        wb.save(tmp)
        path = tmp
    return formulas.ExcelModel().loads(os.path.abspath(path)).finish().calculate()


def get(sol, sheet, cell):
    for k, v in sol.items():
        if k.upper().endswith(f"]{sheet.upper()}'!{cell}"):
            try:
                return float(v.value[0, 0])
            except Exception:
                return v.value[0, 0]
    return None


BASE = [
 ('02_総括原価', 'C14', '総括原価 合計', 1175810),
 ('03_指標と改定率', 'C6', '給水原価', 272.43),
 ('03_指標と改定率', 'C7', '供給単価 税抜', 136.90),
 ('03_指標と改定率', 'C8', '総括原価回収割合', 0.503),
 ('03_指標と改定率', 'C9', '年間不足額', 116993),
 ('03_指標と改定率', 'C10', '必要改定率', 1.990),
 ('03_指標と改定率', 'C11', '料金不足額 定義③', -29488),
 ('03_指標と改定率', 'C13', '赤字半減の改定率', 1.125),
 ('03_指標と改定率', 'C14', '適用改定率', 1.990),
 ('03_指標と改定率', 'C15', '3-A水準係数（自動）', 0.976223),
 ('04_料金への影響', 'F9', '③6〜10㎥ 現行', 1265),
 ('04_料金への影響', 'H9', '③6〜10㎥ パターン1', 2517),
 ('04_料金への影響', 'J9', '③6〜10㎥ 3-A', 1465),
 ('04_料金への影響', 'F16', '年間収入 現行（百万円）', 118.169),
 ('04_料金への影響', 'H16', '年間収入 パターン1', 235.162),
 ('04_料金への影響', 'J16', '年間収入 3-A', 235.162),
 ('04_料金への影響', 'F17', '回収割合 現行', 0.503),
 ('04_料金への影響', 'H17', '回収割合 パターン1', 1.000),
 ('04_料金への影響', 'J17', '回収割合 3-A', 1.000),
 ('04_料金への影響', 'F23', '含意される母集団補正率', 1.15145),
 ('04_料金への影響', 'F27', '標準家庭 現行', 2497),
 ('04_料金への影響', 'H27', '標準家庭 パターン1', 4969),
 ('04_料金への影響', 'J27', '標準家庭 3-A', 4166),
 ('05_口径別料金表', 'D8', '3-A 基本料金 20mm', 1396),
 ('06_ケース比較', 'C6', '採用条件 総括原価', 1175810),
 ('06_ケース比較', 'F6', '採用条件 必要改定率', 1.990),
 ('06_ケース比較', 'C7', '繰入金控除なし 総括原価', 1448807),
 ('06_ケース比較', 'F7', '繰入金控除なし 必要改定率', 2.452),
 ('06_ケース比較', 'C9', '差引方式に戻す 総括原価', 1134492),
 ('06_ケース比較', 'C10', '機械電気20年 総括原価', 1247275),
]

# 条件を変えたときに全体が追随するかの確認
CASES = [
 ('給水収益を650,845千円に', {'C26': 650845},
  [('03_指標と改定率', 'C10', '必要改定率', 1.807),
   ('04_料金への影響', 'F16', '現行年間収入', 130.169),
   ('04_料金への影響', 'H17', 'パターン1の回収割合', 1.000),
   ('04_料金への影響', 'J17', '3-Aの回収割合', 1.000)]),
 ('他会計補助金を控除しない', {'C20': '控除しない'},
  [('03_指標と改定率', 'C10', '必要改定率', 2.452),
   ('03_指標と改定率', 'C15', '3-A水準係数', 1.2028),
   ('04_料金への影響', 'J17', '3-Aの回収割合', 1.000)]),
 ('3-Aを手入力（0.976194）に固定', {'C32': '手入力'},
  [('03_指標と改定率', 'C15', '3-A水準係数', 0.976194),
   ('04_料金への影響', 'J17', '3-Aの回収割合', 1.000)]),
 ('手入力のまま繰入金を控除しない', {'C32': '手入力', 'C20': '控除しない'},
  [('04_料金への影響', 'J17', '3-Aの回収割合', 0.812)]),
 ('1円未満切捨て', {'C28': '1円未満切捨て'},
  [('04_料金への影響', 'J9', '③6〜10㎥ 3-A', 1464)]),
]


def check(sol, rows, head):
    ng = 0
    print(f'\n■ {head}')
    for sh, cell, name, want in rows:
        got = get(sol, sh, cell)
        ok = isinstance(got, (int, float)) and abs(got - want) < max(abs(want) * 0.003, 0.6)
        if not ok:
            ng += 1
        g = format(got, ',.4f') if isinstance(got, (int, float)) else str(got)
        print(f'  {"OK " if ok else "NG "} {name:24s} {sh}!{cell:5s} 期待 {want:>12,.4f}  実測 {g:>14}')
    return ng


total = check(solve(), BASE, '既定条件')
for label, chg, rows in CASES:
    total += check(solve(chg), rows, label)
print(f'\n不一致 {total} 件')
sys.exit(1 if total else 0)
