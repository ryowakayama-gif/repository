# -*- coding: utf-8 -*-
"""旧版Excelの明らかな誤りを直し、中間報告書と一致しない旨の見出しを入れる。

対象は「計算式そのものの誤り」に限る。前提条件（長期前受金の推計値など）は
当時の条件として残す。中間報告書の数値は `09_interim_report/onagawa_adopted_calc.xlsx`
で再現する。
"""
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RED, YELLOW = 'C00000', 'FFF2CC'

BANNER = ('【この版は中間報告書（改訂版）の前提と一致しません】'
          '　計算根拠には使用しないでください。'
          '　本文の数値は 09_interim_report/onagawa_adopted_calc.xlsx（採用計算表）で再現します。')


def snapshot(path):
    wb = openpyxl.load_workbook(path)
    return {sh: (wb[sh].max_row, wb[sh].max_column,
                 sorted(str(m) for m in wb[sh].merged_cells.ranges)) for sh in wb.sheetnames}


def add_banner(ws, row, span=8):
    ws.cell(row=row, column=1, value=BANNER).font = Font(name='Arial', size=10, bold=True, color=RED)
    for i in range(1, span + 1):
        ws.cell(row=row, column=i).fill = PatternFill('solid', fgColor=YELLOW)
    ws.cell(row=row, column=1).alignment = Alignment(vertical='center')
    ws.row_dimensions[row].height = 20


def first_free_row(ws, limit=8):
    """先頭付近の空行を探す。見つからなければ None。"""
    for r in range(2, limit + 1):
        if all(ws.cell(row=r, column=c).value in (None, '') for c in range(1, 12)):
            return r
    return None


def main():
    changes = []

    # --- 02_total_cost_simulation：5年合計が累計の合計になっている誤りを直す ---
    p = os.path.join(ROOT, '02_deliverables', '02_total_cost_simulation.xlsx')
    before = snapshot(p)
    wb = openpyxl.load_workbook(p)
    ws = wb['総括原価シミュレーション']
    for col in 'DEFG':
        old = ws[f'{col}15'].value
        new = f'=SUM({col}8:{col}14)'
        if old != new:
            ws[f'{col}15'] = new
            changes.append(f'02_total_cost_simulation 総括原価シミュレーション!{col}15  {old} → {new}')
    inp = wb['前提条件（入力）']
    if inp['B7'].value != 0.03:
        changes.append(f'02_total_cost_simulation 前提条件（入力）!B7  '
                       f'{inp["B7"].value} → 0.03（仕様書の正式ケース）')
        inp['B7'] = 0.03
    for sh in ('前提条件（入力）', '総括原価シミュレーション'):
        r = first_free_row(wb[sh])
        if r:
            add_banner(wb[sh], r)
            changes.append(f'02_total_cost_simulation {sh}!A{r}  注意書きを追加')
    wb.save(p)
    after = snapshot(p)
    assert before.keys() == after.keys(), 'シート構成が変わった'
    for sh in before:
        assert before[sh][2] == after[sh][2], f'{sh} の結合セルが変わった'

    # --- 注意書きだけ入れるもの ---
    for rel, sheets in [
            ('02_deliverables/01_rate_reform_simulation_MAIN.xlsx', ['01_前提条件', '02_総括原価']),
            ('02_deliverables/03_tariff_period_cost_calc.xlsx', None),
            ('06_capital_plan_update/09_assumption_switch_cost_calc.xlsx', ['01_設定', '06_主要ケース']),
    ]:
        p = os.path.join(ROOT, rel)
        before = snapshot(p)
        wb = openpyxl.load_workbook(p)
        for sh in (sheets or wb.sheetnames):
            if sh not in wb.sheetnames:
                continue
            r = first_free_row(wb[sh])
            if r:
                add_banner(wb[sh], r)
                changes.append(f'{os.path.basename(rel)} {sh}!A{r}  注意書きを追加')
        wb.save(p)
        after = snapshot(p)
        assert before.keys() == after.keys(), f'{rel} のシート構成が変わった'
        for sh in before:
            assert before[sh][2] == after[sh][2], f'{rel} {sh} の結合セルが変わった'

    for c in changes:
        print(' ', c)
    print(f'\n{len(changes)}件を更新しました。')


if __name__ == '__main__':
    main()
