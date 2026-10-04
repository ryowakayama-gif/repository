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
