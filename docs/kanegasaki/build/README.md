# 成果品のビルド元

金ケ崎町　高齢者福祉計画・第10期介護保険事業計画・認知症施策推進計画の
**成果品の本文を組むPython**を置く。**ビルド結果（docx・png）は置かない。**

## 置くもの

| ファイル | 成果品 |
|---|---|
| `ch_soan.py` | 計画素案_第10期.docx |
| `ch_iinkai2.py` | 第2回策定委員会資料.docx |
| `ch_bp2.py` | 別表2_施策体系表.docx |
| `ch_kofu2.py` | 交付金の体系と本町の位置.docx（別冊1） |
| `ch_ninchi_i.py` | 認知症施策推進計画の位置づけと体系.docx（別冊2） |
| `ch_shokai.py` | 確認事項照会_第10期.docx |
| `ch_shiryo.py` | 確認事項の区分.docx |
| `ch_sabun.py` | 資料提供依頼一覧との差分.docx |
| `ch_kikin3.py` | 介護給付費準備基金の取崩額_3案の試算.docx |
| `mk_toc.py` | 計画素案の目次 |
| `chk_*.py` | 照合用（表の合計行・文書間の突合せ・図のデータ・記載事項） |

体裁のヘルパー `fmt.py` と図 `figs.py`・`figs_soan.py` は `../style/` にある。

## 手順

```bash
# 1　スクラップパッドに複写する
cp docs/kanegasaki/style/*.py docs/kanegasaki/build/*.py "$SCRATCHPAD"/
cd "$SCRATCHPAD" && mkdir -p out/fig

# 2　図を先に生成する（matplotlib と IPAGothic が要る）
python3 figs.py && python3 figs_soan.py

# 3　本文をビルドする
{ cat fmt.py; echo; cat ch_soan.py; } > mkf_soan.py && python3 mkf_soan.py

# 4　目次を作る
python3 mk_toc.py

# 5　提示前に点検する（指摘0件にする）
python3 .claude/skills/docx-design/check_docx.py out/*.docx --figs figs.py figs_soan.py
```

## 直したら必ず戻す

**スクラップパッドは作業環境が作り直されると空になる。**
`ch_*.py` を直したら、**同じセッションのうちにこのディレクトリへ戻してコミットする。**

令和8年10月4日、作業環境の作り直しにより、スクラップパッドにのみ置いていた
`ch_*.py` と成果品8点を失った。追跡下にあった `style/` の体裁ファイルと図は、
再出力で戻った。このディレクトリはその再発を防ぐために設けたものである。

## 計画素案は2本ある

| | ビルド元 | 成果品 | 用途 |
|---|---|---|---|
| **協議用** | `ch_soan.py` | `計画素案_第10期.docx` | 町との協議。分析の過程・未確定箇所の表示・附・限界を含む |
| **公表用** | `ch_kohyo.py` | `計画素案_第10期_公表用.docx` | 策定委員会・パブリックコメント・計画書。第1回策定委員会の目次案のとおり |

**公表用は協議用から内容を採って書き直したものである。数値は協議用を正とする。**
`python3 -I chk_kohyo.py` が、公表用の数値が協議用に現れるものだけであることを毎回確かめる。

**目次のページ番号の対応表は文書ごとに分ける。**

```
python3 mk_toc.py "out/計画素案_第10期.docx"        mkf_soan.py
python3 mk_toc.py "out/計画素案_第10期_公表用.docx" mkf_kohyo.py "out/_toc_pages_kohyo.json"
```

同じ対応表を使うと、あとからビルドしたほうが先のページ番号を壊す。

- **`chk_taisai.py`**　docx の体裁の実測値を一覧にする（用紙・余白・Normal・本文と表の書体／サイズ／太字／色の分布・ヘッダ・フッタ・ページ番号）。
  前回計画書との対照表の「本計画」欄を機械で埋めるために使う。`python3 -I chk_taisai.py <docx...>`

- **`chk_taisai_pdf.py`**　ＰＤＦの体裁の実測値を一覧にする（用紙・文字の占める範囲・フォント／サイズ／色の分布）。
  前回計画書（第9期・ＰＤＦ）の体裁を測り、対照表の右欄を埋めるために使う。`python3 -I chk_taisai_pdf.py <pdf>`（要 PyMuPDF）

## 公表用の計画書のビルド

**公表用（`ch_kohyo.py`）だけは体裁が違う。**前回（第9期）計画書の実測値に合わせるため、体裁プロファイル `../style/style_kohyo.py` を先に連結する。

```
{ cat style_kohyo.py; cat fmt.py; echo; cat ch_kohyo.py; } > mkf_kohyo.py && python3 mkf_kohyo.py
```

**他の文書は従前どおり `{ cat fmt.py; echo; cat ch_XXX.py; }` でよい**（体裁は `fmt.py` の既定値）。

- **`ch_bp5.py`**　別表5「所得段階別第1号被保険者数と基準額に対する割合」の本文。`output/所得段階別第1号被保険者数_登録用.csv` と標準段階区分から機械で組んだ（3か年計14,252人・補正後13,920.56人が計画素案 第8章5⑦と一致する）。

- **`ch_kikin3.py`**　介護給付費準備基金の取崩額3案（0円・7,200万円・9,000万円）の試算。算式は第9期計画書 p77。`{ cat fmt.py; echo; cat ch_kikin3.py; } > mkf_kikin3.py && python3 mkf_kikin3.py`

- **`get_jigyosho.py`**　町内介護サービス提供事業所一覧を、介護サービス情報公表システムのオープンデータ（厚生労働省がＣＳＶで公表。全国一括・サービス種類別35ファイル）から組む。`python3 -I get_jigyosho.py out/町内介護サービス提供事業所一覧.csv [時点のスタンプ]`。要ネットワーク `www.mhlw.go.jp`。

- **`ch_sabun.py`**　令和8年10月20日にご提供いただいた「計画素案の完成に向けた資料提供・確認依頼一覧」（Excel・6シート）と、当方の照会文書・計画素案の未確定箇所との突合。記録は `../資料提供依頼一覧との差分.md`。
