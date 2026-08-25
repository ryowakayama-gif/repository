# 福祉計画コラム部品ライブラリ

Word差し込み印刷用のコラム部品（管理表Excel・見本画像）を生成するリポジトリ。

## 構成

| ファイル | 役割 |
| --- | --- |
| `naming.py` | 出力ファイル名と文字エンコードの共通設定。他2スクリプトが参照する |
| `build_excel.py` | 管理表ブック（マスター＋計画別4冊）と提出用CSVを生成 |
| `build_component_images.py` | 基本コラム部品6種の見本画像を生成し、Excelへ貼り付け |
| `output/` | 生成物（xlsx / csv / png） |
| `docs/encoding-investigation.md` | 文字コード調査報告（方針の根拠と検証結果） |

## 実行

```bash
python3 build_excel.py              # 日本語ファイル名（既定）
python3 build_component_images.py

python3 build_excel.py --ascii      # 半角英数ファイル名
python3 build_component_images.py --ascii
```

2つのスクリプトは **必ず同じオプション** で実行する。ブック名が一致しないと
`build_component_images.py` が貼り付け先を見つけられず、終了コード1で停止する。

依存: `openpyxl`, `Pillow`, IPAゴシック（`/usr/share/fonts/opentype/ipafont-gothic/`）

## 文字コードとファイル名の方針

行政系システムへの提出を前提に、次の2点を切り替えられるようにしている。

**ファイル名**
`--ascii` または `ASCII_FILENAMES=1` で、出力ファイル名を半角英数字にする。
行政系のアップロードは「ファイル名は半角英数字のみ」を要求することが多く、
日本語ファイル名が弾かれる原因になりやすい。対応表は `naming.py` の
`BOOKS` と `PART_SLUGS` にあり、部品を追加したらここも更新する。

**CSVの文字コード**
`CSV_ENCODING` で指定する。既定は `cp932`。

| 値 | 用途 |
| --- | --- |
| `cp932`（既定） | Shift_JIS指定の行政系システム。最も多い |
| `utf-8-sig` | UTF-8指定かつExcelで開く場合（BOM付き） |
| `utf-8` | BOMなしUTF-8 |

改行はCRLF固定。CP932で表現できない文字（`✓` `—` など）は
`naming.CP932_SUBSTITUTIONS` で代替文字へ置換し、置換内容を標準出力へ警告する。

**Python・Excel側は変更不要**
ソースは全てUTF-8（各ファイル1行目に `# -*- coding: utf-8 -*-`）。
`.xlsx` の中身はOOXML仕様でUTF-8固定、`.png` はバイナリ。
したがって「UTF-8へ直す」対象は存在せず、提出時に問題になるのは
上記のファイル名とCSVの文字コードの2点に限られる。

## 作業メモ

- **Web検索は Claude in Chrome で開いて行う。** 特に提出先の行政系システムの
  仕様（許可されるファイル名・文字コード・拡張子）を調べるときは、
  通常のweb検索ではなく Claude in Chrome でページを開いて確認すること。
