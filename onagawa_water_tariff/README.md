# 女川町 上水道 料金改定検討

**業務名**：令和8年度 上下水道経営指標評価書作成等業務委託（女川町）
**委託元**：女川町 ／ **受託**：ビズアップ公共コンサルティング株式会社
**中間報告期限**：令和8年9月末

claude.ai のチャットで進めていた検討作業を、Claude Code に移行したものです。

---

## まず読むもの

| 順 | ファイル | 内容 |
|:--:|---|---|
| 1 | [`06_capital_plan_update/README.md`](06_capital_plan_update/README.md) | **建設改良費の改定と必要改定幅への影響。最新の論点** |
| 2 | [`05_review/file_check_report.md`](05_review/file_check_report.md) | Excel成果品の確認レポート。既知の不整合6件と残作業 |
| 3 | [`05_review/word_document_review.md`](05_review/word_document_review.md) | Word文書の確認レポート。表現ルール違反1件ほか |
| 4 | [`03_documents/handoff_memo.md`](03_documents/handoff_memo.md) | 引継ぎメモ。数値前提、確定値と推計値の区別 |
| 5 | [`03_documents/spec_alignment_summary.md`](03_documents/spec_alignment_summary.md) | 業務仕様書との整合確認 |

---

## ディレクトリ構成

```
onagawa_water_tariff/
├── 01_source_evidence/          作業元エビデンス（触らない・上書きしない）
├── 02_deliverables/             成果品Excel 8ブック / 全65シート
├── 03_documents/                引継ぎメモ・仕様書整合確認
├── 04_prefecture_unification/   宮城県水道料金体系統一化 検討（参考資料の位置づけ）
├── 05_review/                   確認・レビュー記録
└── 06_capital_plan_update/      建設改良費調査票（女川町回答）と影響再計算
```

### 01_source_evidence（作業元エビデンス）

| ファイル | 元ファイル名 / 内容 |
|---|---|
| `MHLW_tariff_guideline_shiryo2-2.pdf` | 厚生労働省「水道料金の適正化について」資料2-2 |
| `R1〜R7_billing_summary.xlsx` | 各年度の料金調定集計表 |
| `monthly_billing_detail_R1toR7.xlsx` | 検索一覧表_上水道（R1〜R7月次個票。全使用者の約44%サンプル） |
| `mgmt_strategy_simulation_sheet.xlsx` | 女川町_水道_経営戦略試算シート20251209版（**財政シミュのマスター**） |
| `municipal_bond_ledger.xlsx` | 企業債_AII_布設替入り |
| `mgmt_strategy_report_draft.docx` | 報告書_女川町_水道事業_経営戦略（案） |
| `sewer_fee_analysis_memo.docx` | 下水道使用料分析_引継ぎメモ |

### 02_deliverables（成果品）

| ファイル | 内容 | シート数 |
|---|---|:--:|
| `01_rate_reform_simulation_MAIN.xlsx` | **料金改定シミュレーション本体**。仕様書パターン1〜3に対応 | 8 |
| `02_total_cost_simulation.xlsx` | 総括原価・資産維持率感度分析 ※要修正（レポート#1） | 3 |
| `03_tariff_period_cost_calc.xlsx` | 料金算定対象期間の設定と原価算定 | 2 |
| `04_tariff_1m3_unit_by_diameter.xlsx` | 口径別1㎥刻み調定分析 | 9 |
| `05_volume_zone_by_wateramount.xlsx` | 水量区分別ボリュームゾーン分析 | 4 |
| `06_volume_zone_analysis_R1toR7.xlsx` | 年度別ゾーン推移（R1〜R7）と経営戦略整合確認 | 11 |
| `07_zone_analysis_R1toR6.xlsx` | 過年度ゾーン分析（R1〜R6推計＋R7実績） | 8 |
| `08_billing_summary_R1toR7_v2.xlsx` | 調定集計整理と経営戦略整合確認メモ | 11 |

#### `01_rate_reform_simulation_MAIN.xlsx` のシート構成

| シート | 仕様書対応 |
|---|---|
| `01_前提条件` | **B14セル（資産維持率）が全シート連動のマスターキー。既定値3%** |
| `02_総括原価` | 資産維持費3%算入。給水原価の二重表現（572円／242円）を整理 |
| `03_パターン1_現行体系維持` | 仕様書パターン1 |
| `04_パターン2_赤字半減案` | 仕様書パターン2 |
| `05_パターン3_口径別料金表案` | 仕様書パターン3（3-A標準逓増型／3-B産業配慮型） |
| `06_水産加工業影響` | 19社モデル×4改定案の影響試算 |
| `07_改定案比較表` | A〜E案の横断比較 |
| `08_住民議会用要約` | 住民・議会説明用の切り抜き資料 |

### 04_prefecture_unification（県内統一化検討）

宮城県内臨海4団体（女川町・気仙沼市・石巻広域・塩竈市）の料金体系比較。
仕様書上は**参考資料**の位置づけで、主資料は女川町単独の改定シミュレーション（02_deliverables）です。

検証スクリプトは再実行可能です。

```bash
cd 04_prefecture_unification/03_verification_scripts
python3 verify_all_rates.py      # 各団体の公式早見表との逆算検証（全件一致）
python3 impact_simulation.py     # 政策シナリオ別の影響試算
```

### 06_capital_plan_update（建設改良費の改定）

女川町から回答のあった建設改良費調査票（R8〜R17）と、それが総括原価・必要改定幅に
与える影響の再計算です。**成果品Excelの数値はまだこの改定を織り込んでいません。**

| ファイル | 内容 |
|---|---|
| `construction_cost_survey_R8toR17.xlsx` | 建設改良費調査票（女川町回答・全4事業） |
| `recalc_capital_impact.py` | 新旧計画の比較と総括原価への影響の再計算 |
| `recalc_output.txt` | 上記の実行結果 |
| `README.md` | 結論・感応度分析・調査票の不備・確認事項 |

```bash
cd 06_capital_plan_update && python3 recalc_capital_impact.py
```

---

## 旧チャット環境からのパス対応

`handoff_memo.md` 第8章に記載のパスは旧環境のものです。

| 旧環境 | 本リポジトリ |
|---|---|
| `/mnt/user-data/uploads/` | `01_source_evidence/` |
| `/mnt/user-data/outputs/` | `02_deliverables/` |

| 旧ファイル名 | 本リポジトリ |
|---|---|
| `onagawa_rate_reform_v1.xlsx` | `02_deliverables/01_rate_reform_simulation_MAIN.xlsx` |
| `onagawa_cost_sim_v1.xlsx` | `02_deliverables/02_total_cost_simulation.xlsx` |
| `onagawa_tariff_cost_calc.xlsx` | `02_deliverables/03_tariff_period_cost_calc.xlsx` |
| `onagawa_water_tariff_1m3_unit.xlsx` | `02_deliverables/04_tariff_1m3_unit_by_diameter.xlsx` |
| `onagawa_vol_zone_by_wateramt.xlsx` | `02_deliverables/05_volume_zone_by_wateramount.xlsx` |
| `onagawa_volume_zone_analysis_v2.xlsx` | `02_deliverables/06_volume_zone_analysis_R1toR7.xlsx` |
| `onagawa_water_R1toR6_zone_analysis.xlsx` | `02_deliverables/07_zone_analysis_R1toR6.xlsx` |
| `水道料金調定集計_整理済み_R1-R7_v2.xlsx` | `02_deliverables/08_billing_summary_R1toR7_v2.xlsx` |
| `onagawa_rate_revision_summary.md` | `03_documents/spec_alignment_summary.md` |

---

## 数値の取扱い

成果品には確定値と推計値が混在しています。詳細は `handoff_memo.md` 第4章。

**確定値**（出典と突合済み）
- 現行料金体系（女川町公式サイト）
- R7実績：有収水量 1,024,576㎥ ／ 料金収入 131,428,825円 ／ 供給単価 128.2円/㎥
- 財政シミュの費目（維持管理費・減価償却費・企業債利息）
- 県内4団体の料金（各団体の条例別表・公式早見表から逆算検証済み）

**推計値（要更新）**
- 控除項目（長期前受金戻入益・他会計補助金等）
- 対象資産の期首・期末残高
- 水産加工業19社モデルの使用水量

R6・R7決算確定値が揃い次第、`01_..._MAIN.xlsx` の `02_総括原価` と `08_住民議会用要約` の更新が必要です。
