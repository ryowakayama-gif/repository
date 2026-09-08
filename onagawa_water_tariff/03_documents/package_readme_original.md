# 女川町上水道 料金改定検討 一式

作成日：2026年9月8日
案件：令和8年度 上下水道経営指標評価書作成等業務委託（女川町）
担当：ビズアップ公共コンサルティング株式会社

ダウンロード時の文字化けを避けるため、フォルダ内のファイル名はすべて半角英数字にしています。元の日本語ファイル名は下表を参照してください。

---

## 01_source_evidence（作業元エビデンス）

分析の根拠となった元データです。成果品の数値はすべてこのフォルダのファイルから導出しています。

| ファイル名 | 元ファイル名 / 内容 |
|---|---|
| MHLW_tariff_guideline_shiryo2-2.pdf | 0000134952.pdf（厚生労働省「水道料金の適正化について」資料2-2） |
| R1_billing_summary.xlsx | R1料金調定集計表_20251209174132.xlsx |
| R2_billing_summary.xlsx | R2料金調定集計表_20251209174116.xlsx |
| R3_billing_summary.xlsx | R3料金調定集計表_20251209174051.xlsx |
| R4_billing_summary.xlsx | R4料金調定集計表_20251209174029.xlsx |
| R5_billing_summary.xlsx | R5料金調定集計表_20251209174005.xlsx |
| R6_billing_summary.xlsx | R6水道_料金調定集計表_20251209173820.xlsx |
| R7_billing_summary.xlsx | R7水道_料金調定集計表_.xlsx |
| monthly_billing_detail_R1toR7.xlsx | 検索一覧表_上水道.xlsx（R1〜R7月次個票。全使用者の約44%サンプル） |
| mgmt_strategy_simulation_sheet.xlsx | 女川町_水道_経営戦略試算シート20251209版.xlsx（財政シミュのマスター） |
| municipal_bond_ledger.xlsx | 企業債_AII_布設替入り.xlsx |
| mgmt_strategy_report_draft.docx | ＿報告書_女川町_水道事業_経営戦略_案_.docx |
| sewer_fee_analysis_memo.docx | 下水道使用料分析_引継ぎメモ.docx |

---

## 02_deliverables（成果品）

| ファイル名 | 内容 | シート数 |
|---|---|---|
| 01_rate_reform_simulation_MAIN.xlsx | **料金改定シミュレーション本体**。仕様書パターン1〜3に対応 | 8 |
| 02_total_cost_simulation.xlsx | 総括原価シミュレーション。資産維持率1〜3%を可変で再計算 | 3 |
| 03_tariff_period_cost_calc.xlsx | 料金算定対象期間の設定と原価算定 | 2 |
| 04_tariff_1m3_unit_by_diameter.xlsx | 口径別1㎥刻み調定分析 | 9 |
| 05_volume_zone_by_wateramount.xlsx | 水量区分別ボリュームゾーン分析 | 4 |
| 06_volume_zone_analysis_R1toR7.xlsx | 年度別ゾーン推移（R1〜R7）と経営戦略整合確認 | 11 |
| 07_zone_analysis_R1toR6.xlsx | 過年度ゾーン分析（R1〜R6推計＋R7実績） | 8 |
| 08_billing_summary_R1toR7_v2.xlsx | 調定集計整理と経営戦略整合確認メモ | 11 |

### 01_rate_reform_simulation_MAIN.xlsx のシート構成

| シート | 仕様書対応 |
|---|---|
| 01_前提条件 | **B14セル（資産維持率）が全シート連動のマスターキー。既定値3%** |
| 02_総括原価 | 資産維持費3%算入。給水原価の二重表現（572円／242円）を整理 |
| 03_パターン1_現行体系維持 | 仕様書パターン1 |
| 04_パターン2_赤字半減案 | 仕様書パターン2 |
| 05_パターン3_口径別料金表案 | 仕様書パターン3（3-A標準逓増型／3-B産業配慮型） |
| 06_水産加工業影響 | 19社モデル×4改定案の影響試算 |
| 07_改定案比較表 | A〜E案の横断比較 |
| 08_住民議会用要約 | 住民・議会説明用の切り抜き資料 |

---

## 03_documents（ドキュメント）

| ファイル名 | 内容 |
|---|---|
| spec_alignment_summary.md | 業務仕様書との整合確認まとめ（onagawa_rate_revision_summary.md） |
| handoff_memo.md | 引き継ぎメモ。数値前提、確定値と推計値の区別、残作業を記載 |

---

## 数値の取扱いについて

成果品の数値には確定値と推計値が混在しています。詳細は handoff_memo.md の第4章を参照してください。

**確定値**：現行料金体系、R7実績（有収水量1,024,576㎥、料金収入131,428,825円）、財政シミュの費目、県内4団体の料金

**推計値（要更新）**：控除項目（長期前受金戻入益・他会計補助金等）、対象資産の期首・期末残高、水産加工業19社モデルの使用水量

R6・R7決算確定値が揃い次第、01_rate_reform_simulation_MAIN.xlsx の02_総括原価シートと08_住民議会用要約シートの数値更新が必要です。
