# -*- coding: utf-8 -*-
"""受け渡し用zipを作成する。

zip内のファイル名はすべて半角英数字にする（日本語名のままだとWindowsで文字化けするため）。
対応表 00_filename_map.md を同梱する。
"""
import os, sys, zipfile, datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT  = os.path.join(ROOT, '09_interim_report', 'onagawa_interim_report_v3.zip')

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
 ('02_calculation/adopted_calc.xlsx', '09_interim_report/onagawa_adopted_calc.xlsx',
  '★採用計算表。中間報告書の数値はこの1冊から再現できる（07_本文照合表で対応を確認）'),
 ('02_calculation/scenario_calc.xlsx', '09_interim_report/onagawa_scenario_calc.xlsx',
  '★試算条件切替計算表。01_設定の黄色いセルを変えると改定率まで再計算される'),
 ('02_calculation/construction_cost_survey_R8toR17.xlsx',
  '06_capital_plan_update/construction_cost_survey_R8toR17.xlsx',
  '建設改良費調査票（女川町回答）'),
 ('02_calculation/tariff_model.py', '09_interim_report/build/tariff_model.py',
  '採用計算表を作成する作業用ファイル'),
 ('02_calculation/figures.json', '09_interim_report/build/figures.json',
  '算定結果'),

 # 調定データの集計（前提の影響を受けない実績の整理）
 ('03_billing_analysis/04_tariff_1m3_unit_by_diameter.xlsx',
  '02_deliverables/04_tariff_1m3_unit_by_diameter.xlsx', '口径別1㎥単価'),
 ('03_billing_analysis/05_volume_zone_by_wateramount.xlsx',
  '02_deliverables/05_volume_zone_by_wateramount.xlsx', '水量ゾーン別集計'),
 ('03_billing_analysis/06_volume_zone_analysis_R1toR7.xlsx',
  '02_deliverables/06_volume_zone_analysis_R1toR7.xlsx', '水量ゾーン分析 R1〜R7'),
 ('03_billing_analysis/07_zone_analysis_R1toR6.xlsx',
  '02_deliverables/07_zone_analysis_R1toR6.xlsx', 'ゾーン分析 R1〜R6'),
 ('03_billing_analysis/08_billing_summary_R1toR7_v2.xlsx',
  '02_deliverables/08_billing_summary_R1toR7_v2.xlsx', '調定集計 R1〜R7'),

 # 旧版（使用不可・履歴として同梱）
 ('09_old_do_not_use/01_rate_reform_simulation_MAIN.xlsx',
  '02_deliverables/01_rate_reform_simulation_MAIN.xlsx',
  '★旧版。総括原価2,552,199千円ほか旧前提のまま。使用不可'),
 ('09_old_do_not_use/02_total_cost_simulation.xlsx',
  '02_deliverables/02_total_cost_simulation.xlsx',
  '★旧版。5年合計が累計の合計となる誤式、資産維持率1％のまま。使用不可'),
 ('09_old_do_not_use/03_tariff_period_cost_calc.xlsx',
  '02_deliverables/03_tariff_period_cost_calc.xlsx', '★旧版。使用不可'),
 ('09_old_do_not_use/assumption_switch_cost_calc.xlsx',
  '06_capital_plan_update/09_assumption_switch_cost_calc.xlsx',
  '★前提条件切替計算表。主要ケースは旧値のまま。前提の比較にのみ使用'),

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

## はじめにお読みください

中間報告書の数値を再現できるのは **`02_calculation/adopted_calc.xlsx`（採用計算表）** だけです。
同表の `07_本文照合表` に、報告書の各数値と計算表の対応を載せています。

前提を変えて改定率を確かめたいときは **`02_calculation/scenario_calc.xlsx`（試算条件切替計算表）**
をお使いください。`01_設定` の黄色いセルを変えると、総括原価から必要改定率、
利用者の月額までが自動で計算し直されます。Excelで開いてご利用ください。

`09_old_do_not_use/` のファイルは、中間報告書より前の前提で作成した旧版です。
数値が報告書と一致しないため、**計算根拠として使用しないでください**。
経緯を追えるように同梱していますが、町への説明や条例化の根拠には用いません。

| zip内のファイル | 元のファイル名・内容 |
|---|---|
"""

FOOT = """

## フォルダの構成

| フォルダ | 内容 | 使い方 |
|---|---|---|
| `01_report/` | 中間報告書・指摘事項対応表・工程表・WBS | 本体 |
| `02_calculation/` | 採用計算表・試算条件切替計算表・建設改良費調査票 | **数値の根拠はここ** |
| `03_billing_analysis/` | 調定データの集計（前提の影響を受けない実績の整理） | 参考 |
| `04_reference/` | 県内料金体系統一化の検討資料 | 参考 |
| `05_source/` | 算定の根拠とした受領資料 | 原票 |
| `06_check/` | 確認の記録と作業上の取り決め | 経緯 |
| `09_old_do_not_use/` | 旧版のExcel | **使用不可** |

## 主要な数値（中間報告書・改訂版）

| 指標 | 値 |
|---|---:|
| 総括原価（令和8〜12年度 5年計・税抜） | 1,175,810千円 |
| 給水原価 | 272.4円/㎥ |
| 供給単価（令和8〜12年度・税抜） | 136.9円/㎥ |
| 総括原価回収割合 | 50.3% |
| 年間不足額 | 116,993千円/年 |
| 必要改定率（試算値） | 1.99倍 |
| 　（基準内繰入金を控除しない場合） | 2.45倍 |

## 数値を引用するときの注意

- **1.99倍は確定した改定水準ではありません。** 他会計補助金272,997千円の全額を
  基準内繰入金として控除できる場合の試算値です。控除の根拠は町の資料で未確認で、
  控除しない場合は2.45倍になります。
- **口径別3-Aの「回収割合100％」は検証値ではありません。** 総括原価と一致するよう
  料金水準を逆算した結果であり、実際の調定データで確かめた値ではありません。
- **産業用特例（3-B／D案）は未設計です。** 適用の境界で料金が約6割下がる逆転があり、
  町全体の減収額も算定していません。住民・議会説明や条例化の対象としないでください。
- **新料金の適用開始は令和9年4月以降**です。令和8年度の料金は現行のままで、
  報告書の増収額は通年ベースの試算値です。
- 料金表と利用者への影響は税込、収支・総括原価・供給単価は税抜です。
  令和7年度実績の供給単価は税込128.2円/㎥・税抜116.6円/㎥で、
  将来推計136.9円/㎥と比較できるのは税抜の116.6円/㎥です。

前提条件と限界は中間報告書 第5章にまとめています。
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
