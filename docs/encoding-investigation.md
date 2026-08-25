# 文字コード調査報告

| 項目 | 内容 |
| --- | --- |
| 調査日 | 2026-08-25 |
| 対象 | 福祉計画コラム部品ライブラリ（`build_excel.py` / `build_component_images.py` / `output/`） |
| 依頼内容 | 「デフォルトの文字エンコードを知りたい。行政系で弾かれるのでUTF-8へ修正したい」 |
| ブランチ | `claude/default-encoding-utf8-5772it` |

---

## 1. 結論

**このリポジトリの文字エンコードは調査時点ですでに全レイヤーがUTF-8であり、UTF-8化のための修正対象は存在しなかった。**

行政系システムで弾かれていた原因は文字コードではなく、次の2点である可能性が高い。

1. **出力ファイル名が日本語（非ASCII）** — 全11ファイルが該当
2. **CSV提出時の文字コード** — 行政系はUTF-8ではなく Shift_JIS(CP932) 指定が多い

この2点に対応する切り替え機構を実装した（第5章）。

---

## 2. 調査方法と結果

### 2.1 ソースコード

各ファイル1行目のエンコード宣言と、実ファイルのバイト列の両方を確認した。

```bash
head -1 build_excel.py build_component_images.py   # => # -*- coding: utf-8 -*-
file build_excel.py build_component_images.py
```

```
build_excel.py:            Python script, Unicode text, UTF-8 text executable
build_component_images.py: Python script, Unicode text, UTF-8 text executable
```

宣言・実体ともUTF-8。**問題なし。**

### 2.2 Python実行時

```
Python 3.11.15

sys.getdefaultencoding()             = utf-8
sys.getfilesystemencoding()          = utf-8
locale.getpreferredencoding(False)   = utf-8
sys.stdout.encoding                  = utf-8
sys.flags.utf8_mode                  = 1
```

すべてUTF-8。**問題なし。**

### 2.3 出力 `.xlsx`

`.xlsx` はZIP書庫であり、中身のXMLを取り出してバイト列を直接判定した。

```python
z = zipfile.ZipFile('output/01_共通_基本コラム部品.xlsx')
b = z.read('xl/worksheets/sheet1.xml')
b.decode('utf-8')   # => 成功
b.decode('cp932')   # => UnicodeDecodeError
```

CP932としては解釈できずUTF-8としてのみ解釈できる、すなわち**UTF-8で確定**。
そもそもOOXML仕様上 `.xlsx` の内部XMLはUTF-8固定であり、openpyxlに選択の余地はない。**問題なし。**

### 2.4 出力 `.png`

バイナリ形式。文字コードの概念が存在しない。**対象外。**

### 2.5 テキスト書き出しの有無

```bash
grep -n "open(\|csv\.\|json\.dump\|\.write(" build_excel.py build_component_images.py
# => 該当なし（1行目の coding 宣言のみ）
```

調査時点の両スクリプトには**テキストファイルの書き出しが1箇所も存在しなかった**。
`open()` の `encoding` 引数を指定する対象自体がなく、この意味でも「UTF-8へ直す」箇所は無かった。

### 2.6 まとめ

| レイヤー | 調査時点のエンコード | 判定 |
| --- | --- | --- |
| Pythonソース（宣言・実体） | UTF-8 | 修正不要 |
| Python実行時（default / filesystem / locale / stdout） | UTF-8 | 修正不要 |
| 出力 `.xlsx` 内部XML | UTF-8（仕様上固定） | 修正不要 |
| 出力 `.png` | バイナリ | 対象外 |
| テキスト出力処理 | 存在しない | 対象外 |

---

## 3. 弾かれる原因の分析

文字コードが原因でないとすると、行政系システムのアップロードで弾かれる要因は次が候補になる。

### 3.1 ファイル名が日本語（最有力）

調査時点の出力11ファイルすべてが非ASCII文字を含んでいた。

```
output/00_全計画マスター管理表.xlsx
output/01_共通_基本コラム部品.xlsx
output/images_basic/BC-01_ポイント.png   ...ほか
```

行政系のアップロード画面は「ファイル名は半角英数字のみ」を要求することが多い。
また日本語ファイル名はブラウザ・サーバ間でのURLエンコードやZIP格納時の
文字化けの原因にもなりやすく、**最も疑わしい要因**と判断した。

### 3.2 CSVの文字コード指定（UTF-8とは限らない）

**依頼の前提と実態が食い違う可能性がある点。** 行政系システムのCSV取り込みは
UTF-8ではなく **Shift_JIS(CP932) を指定するケースが多い**。
「UTF-8へ直す」ことがかえって弾かれる原因になり得るため、
既定をCP932とし、UTF-8系も選べる形で実装した。

### 3.3 ファイル形式の制限

`.xlsx` を受け付けず `.xls` / `.csv` / `.pdf` のみ、といった制限。
提出先の仕様確認が必要（第7章）。

---

## 4. CP932変換可否の事前検証

CSVをShift_JISで出力する前提として、既存データが実際にCP932へ変換できるかを
全項目（`COMPONENTS` / `LAYOUTS` / `DESIGN_RULES` / `RECOMMEND_SIZE` / `MERGE_FIELDS`）
について1文字ずつ検査した。

```
=== CP932でエンコード不可の文字 ===
  なし
```

差込データ本体は**全文字がCP932で表現可能**であることを確認した。
ただし将来の追記に備え、変換不能文字を代替文字へ置き換える処理を実装している（第5.3節）。

---

## 5. 実施した対応

### 5.1 `naming.py` の新設

出力ファイル名と文字エンコードの規則を1箇所へ集約した。
`build_excel.py` と `build_component_images.py` の双方がこれを参照するため、
両者の命名が食い違わない。

| 定義 | 役割 |
| --- | --- |
| `BOOKS` | ブック名の 日本語 / 半角英数 対応表 |
| `PART_SLUGS` | 部品IDごとの半角英数スラッグ |
| `CP932_SUBSTITUTIONS` | CP932非対応文字の代替表 |
| `write_csv()` | 文字コード・改行コードを制御したCSV書き出し |

### 5.2 半角英数ファイル名モード

`--ascii` オプション、または環境変数 `ASCII_FILENAMES=1` で切り替える。

| 日本語名（既定） | 半角英数名（`--ascii`） |
| --- | --- |
| `00_全計画マスター管理表.xlsx` | `00_master_all_plans.xlsx` |
| `01_共通_基本コラム部品.xlsx` | `01_common_basic_columns.xlsx` |
| `02_高齢者介護保険事業計画.xlsx` | `02_senior_care_plan.xlsx` |
| `03_障がい福祉計画.xlsx` | `03_disability_welfare_plan.xlsx` |
| `04_こども計画.xlsx` | `04_child_plan.xlsx` |
| `BC-01_ポイント.png` | `BC-01_point.png` |
| `BC-05_データの見方.png` | `BC-05_data_guide.png` |

部品を追加した場合は `naming.py` の `BOOKS` / `PART_SLUGS` も更新すること。
`PART_SLUGS` 未登録のIDは部品IDから機械的にスラッグを生成するため、
**未登録でも非ASCIIファイル名にはならない**。

### 5.3 提出用CSV出力

差込データを `output/csv/` へCSVとして出力する処理を追加した。

- **文字コード**: 既定 `cp932`（Shift_JIS）。環境変数 `CSV_ENCODING` で変更可
- **改行コード**: CRLF固定（Excel・行政系システムの慣例）
- **変換不能文字**: `CP932_SUBSTITUTIONS` により代替文字へ置換し、置換内容を標準出力へ警告

| `CSV_ENCODING` | 用途 |
| --- | --- |
| `cp932`（既定） | Shift_JIS指定の行政系システム。最も多い |
| `utf-8-sig` | UTF-8指定かつExcelで開く場合（BOM付き） |
| `utf-8` | BOMなしUTF-8 |

置換の実例（`〜` は入力時点で全角チルダのため置換対象外）:

```
! cp932 で表現できない文字を置換しました: '—'→'―'、'✓'→'*'
```

### 5.4 オプション不一致の検知

2つのスクリプトは同じオプションで実行する必要がある。
不一致の場合、`build_component_images.py` が貼り付け先ブックを見つけられないため、
理由を表示して**終了コード1で停止**するようにした。

```
エラー: 貼り付け先のブックがありません -> /home/user/repository/output/01_共通_基本コラム部品.xlsx
       先に同じオプションで build_excel.py を実行してください。
       （現在のモード … ファイル名: 日本語 / CSV文字コード: cp932）
```

### 5.5 その他

- 差込データの列定義（`MERGE_HEADERS`）と行生成（`merge_data_row()`）を
  ExcelシートとCSVで共用するよう整理し、二重管理を解消
- `CLAUDE.md`（方針・運用手順）と `.gitignore`（`__pycache__/`）を追加

---

## 6. 検証結果

両モードでスクリプトを実際に実行し、以下を確認した。

| 検証項目 | 結果 |
| --- | --- |
| 日本語名モードでの生成 | 従来の11ファイルを同名で再現（＋CSV5件） |
| 半角英数モードでの生成 | 16ファイル生成、**非ASCIIファイル名 0件** |
| CSVがShift_JISであること | 全10ファイルで「CP932デコード成功 / UTF-8デコード失敗」= Shift_JIS確定 |
| CSVの改行コード | CRLF 7件 / LF単独 0件 |
| `CSV_ENCODING=utf-8-sig` | BOM(`EF BB BF`)付与を確認 |
| xlsxの内容保持 | シート構成・埋込画像6枚を両モードで確認 |
| オプション不一致時 | 終了コード1で停止することを確認 |
| CP932非対応文字の置換 | `✓`→`*`、`—`→`―` の置換と警告表示を確認 |

生成物は日本語名・半角英数名の両方をリポジトリに含めている（合計32ファイル / 884KB）。

---

## 7. 補足と残課題

### 7.1 実行環境による差異（注意点）

本調査での `sys.flags.utf8_mode = 1` は、コンテナで `LANG` が未設定のため
PythonのUTF-8モードが自動で有効になった結果である。
**Windows日本語環境で実行した場合、`locale.getpreferredencoding()` は `cp932` を返す。**

現状の実装では、テキスト出力は `naming.write_csv()` に集約されており
文字コードを明示指定しているため、**この差異による影響は受けない**。
ただし今後 `open()` でテキストを書き出す処理を追加する場合は、
必ず `encoding=` を明示すること。

### 7.2 未確認事項

本調査は手元のリポジトリ内で完結しており、**提出先システムの実際の仕様は未確認**である。
原因の特定には次の情報が必要。

- 提出先の行政系システム名
- アップロード時に表示された実際のエラーメッセージ
- 仕様上要求されるファイル形式・文字コード・ファイル名の制約

なお提出先仕様の調査は、通常のweb検索ではなく
**Claude in Chrome でページを開いて確認する**方針とした（`CLAUDE.md` に記載）。

---

## 8. 運用手順

```bash
# 通常運用（日本語ファイル名）
python3 build_excel.py
python3 build_component_images.py

# 行政系へ提出する場合（半角英数ファイル名 + Shift_JIS CSV）
python3 build_excel.py --ascii
python3 build_component_images.py --ascii

# 提出先がUTF-8指定の場合
CSV_ENCODING=utf-8-sig python3 build_excel.py --ascii
```

2つのスクリプトは**必ず同じオプション**で実行する。

依存: `openpyxl`, `Pillow`, IPAゴシック（`/usr/share/fonts/opentype/ipafont-gothic/`）
