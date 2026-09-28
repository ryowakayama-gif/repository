# -*- coding: utf-8 -*-
"""受け渡し用zipを作成する。

zip内のファイル名はすべて半角英数字にする（日本語名のままだとWindowsで文字化けするため）。
対応表 00_filename_map.md を同梱する。
"""
import os, sys, zipfile, datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT  = os.path.join(ROOT, '09_interim_report', 'onagawa_interim_report_v2.zip')

# (zip内のパス, 元のパス, 日本語名または説明)
ITEMS = [
 # 中間報告書一式
 ('01_report/interim_report.docx',
  '09_interim_report/女川町水道料金改定_中間報告書.docx',
  '女川町水道料金改定_中間報告書.docx'),
 ('01_report/review_response_table.docx',
  '09_interim_report/女川町水道料金改定_中間報告書_指摘事項対応表.docx',
  '女川町水道料金改定_中間報告書_指摘事項対応表.docx'),
 ('01_report/schedule_to_final_report.docx',
  '09_interim_report/女川町水道料金改定_最終報告に向けた工程表.docx',
  '女川町水道料金改定_最終報告に向けた工程表.docx'),
 ('01_report/README.md', '09_interim_report/README.md', '中間報告書の説明'),

 # WBS・進捗
 ('01_report/wbs.xlsx', '10_wbs/onagawa_wbs_R8.xlsx', 'WBS（業務分解構成）'),
 ('01_report/wbs_and_progress.docx',
  '10_wbs/女川町上下水道経営指標評価_WBSと進捗状況.docx',
  '女川町上下水道経営指標評価_WBSと進捗状況.docx'),
 ('01_report/wbs_README.md', '10_wbs/README.md', 'WBSの説明'),

 # 算定に用いた計算表
 ('02_calculation/assumption_switch_cost_calc.xlsx',
  '06_capital_plan_update/09_assumption_switch_cost_calc.xlsx',
  '総括原価 前提条件切替計算表'),
 ('02_calculation/construction_cost_survey_R8toR17.xlsx',
  '06_capital_plan_update/construction_cost_survey_R8toR17.xlsx',
  '建設改良費調査票（女川町回答）'),
 ('02_calculation/tariff_model.py', '09_interim_report/build/tariff_model.py',
  '料金パターンの算定に用いた作業用ファイル'),
 ('02_calculation/figures.json', '09_interim_report/build/figures.json',
  '算定結果'),

 # 料金改定シミュレーション成果品
 ('03_simulation/01_rate_reform_simulation_MAIN.xlsx',
  '02_deliverables/01_rate_reform_simulation_MAIN.xlsx', '料金改定シミュレーション本体'),
 ('03_simulation/02_total_cost_simulation.xlsx',
  '02_deliverables/02_total_cost_simulation.xlsx', '総括原価シミュレーション'),
 ('03_simulation/03_tariff_period_cost_calc.xlsx',
  '02_deliverables/03_tariff_period_cost_calc.xlsx', '料金算定期間原価計算'),
 ('03_simulation/04_tariff_1m3_unit_by_diameter.xlsx',
  '02_deliverables/04_tariff_1m3_unit_by_diameter.xlsx', '口径別1㎥単価'),
 ('03_simulation/05_volume_zone_by_wateramount.xlsx',
  '02_deliverables/05_volume_zone_by_wateramount.xlsx', '水量ゾーン別集計'),
 ('03_simulation/06_volume_zone_analysis_R1toR7.xlsx',
  '02_deliverables/06_volume_zone_analysis_R1toR7.xlsx', '水量ゾーン分析 R1〜R7'),
 ('03_simulation/07_zone_analysis_R1toR6.xlsx',
  '02_deliverables/07_zone_analysis_R1toR6.xlsx', 'ゾーン分析 R1〜R6'),
 ('03_simulation/08_billing_summary_R1toR7_v2.xlsx',
  '02_deliverables/08_billing_summary_R1toR7_v2.xlsx', '調定集計 R1〜R7'),

 # 参考資料（県内統一化の検討）
 ('04_reference/onagawa_tariff_comparison.docx',
  '04_prefecture_unification/01_final_deliverables/女川町水道料金_比較検討資料.docx',
  '女川町水道料金_比較検討資料.docx'),
 ('04_reference/prefecture_unification_report.docx',
  '04_prefecture_unification/01_final_deliverables/宮城県水道料金体系統一化_検討報告書.docx',
  '宮城県水道料金体系統一化_検討報告書.docx'),
 ('04_reference/prefecture_unification_materials.docx',
  '04_prefecture_unification/01_final_deliverables/宮城県水道料金体系統一化_検討資料.docx',
  '宮城県水道料金体系統一化_検討資料.docx'),
 ('04_reference/tariff_detail_confirmed.docx',
  '04_prefecture_unification/01_final_deliverables/料金詳細確認_確定版.docx',
  '料金詳細確認_確定版.docx'),
 ('04_reference/fishery_impact_quantified.docx',
  '04_prefecture_unification/01_final_deliverables/水産加工業影響_定量化資料.docx',
  '水産加工業影響_定量化資料.docx'),
 ('04_reference/coastal_municipality_comparison.docx',
  '04_prefecture_unification/01_final_deliverables/県内臨海団体比較_影響試算.docx',
  '県内臨海団体比較_影響試算.docx'),

 # 根拠資料（受領資料）
 ('05_source/mgmt_strategy_simulation_sheet.xlsx',
  '01_source_evidence/mgmt_strategy_simulation_sheet.xlsx',
  '経営戦略 財政シミュレーション（受領資料）'),
 ('05_source/mgmt_strategy_report_draft.docx',
  '01_source_evidence/mgmt_strategy_report_draft.docx',
  '経営戦略（案）本文（受領資料）'),
 ('05_source/municipal_bond_ledger.xlsx',
  '01_source_evidence/municipal_bond_ledger.xlsx', '企業債台帳（受領資料）'),
 ('05_source/R7_billing_summary.xlsx',
  '01_source_evidence/R7_billing_summary.xlsx', '令和7年度 調定集計（受領資料）'),
 ('05_source/R6_billing_summary.xlsx',
  '01_source_evidence/R6_billing_summary.xlsx', '令和6年度 調定集計（受領資料）'),
 ('05_source/MHLW_tariff_guideline.pdf',
  '01_source_evidence/MHLW_tariff_guideline_shiryo2-2.pdf',
  '資産維持費に関する資料（受領資料）'),

 # 確認記録
 ('06_check/file_check_report.md', '05_review/file_check_report.md', 'ファイル確認結果'),
 ('06_check/word_document_review.md', '05_review/word_document_review.md', 'Word文書の確認結果'),
 ('06_check/working_rules.md', 'CLAUDE.md', '作業上の取り決め'),
]

HEAD = """# ファイル名対応表

このzipのファイル名は、Windowsでの文字化けを避けるためすべて半角英数字にしています。
下表で元のファイル名・内容を確認してください。

作成日：{date}
作成：ビズアップ公共コンサルティング株式会社

| zip内のファイル | 元のファイル名・内容 |
|---|---|
"""

FOOT = """

## フォルダの構成

| フォルダ | 内容 |
|---|---|
| `01_report/` | 中間報告書・指摘事項対応表・最終報告に向けた工程表 |
| `02_calculation/` | 総括原価と料金パターンの算定に用いた計算表 |
| `03_simulation/` | 料金改定シミュレーションの成果品（Excel 8点） |
| `04_reference/` | 県内料金体系統一化の検討資料（参考） |
| `05_source/` | 算定の根拠とした受領資料 |
| `06_check/` | 確認の記録と作業上の取り決め |

## 主要な数値（中間報告書・改訂版）

| 指標 | 値 |
|---|---:|
| 総括原価（令和8〜12年度 5年計） | 1,175,810千円 |
| 給水原価 | 272.4円/㎥ |
| 供給単価（令和8〜12年度・税抜） | 136.9円/㎥ |
| 総括原価回収割合 | 50.3% |
| 年間不足額 | 116,993千円/年 |
| 必要改定率 | 1.99倍 |
| 　（基準内繰入金を控除しない場合） | 2.45倍 |

1.99〜2.45倍は繰入金の控除条件のみを変えた比較値であり、必要改定率の上限・下限ではありません。
前提条件と限界は中間報告書 第5章を参照してください。
"""

def main():
    rows = []
    missing = []
    with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
        for arc, rel, label in ITEMS:
            src = os.path.join(ROOT, rel)
            if not os.path.exists(src):
                missing.append(rel); continue
            z.write(src, arc)
            rows.append(f'| `{arc}` | {label} |')
        body = HEAD.format(date=datetime.date.today().strftime('%Y年%-m月%-d日')) \
               + '\n'.join(rows) + FOOT
        z.writestr('00_filename_map.md', body.encode('utf-8'))
    for m in missing:
        print('見つかりません:', m)
    size = os.path.getsize(OUT)
    with zipfile.ZipFile(OUT) as z:
        n = len(z.namelist())
        non_ascii = [x for x in z.namelist() if not x.isascii()]
    print(f'作成: {OUT}')
    print(f'  {n}ファイル / {size/1024/1024:.1f}MB')
    print(f'  半角英数字以外のファイル名: {non_ascii if non_ascii else "なし"}')

if __name__ == '__main__':
    main()
