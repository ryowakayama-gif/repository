# 機械で確かめる手順（川崎町 第10期）

実行はすべて `kawasaki_project/` から行う。スクリプトは相対パスで
成果品を読むため、`07_ソーススクリプト/` に降りて実行すると素案が見つからない。

```bash
cd /home/user/repository/kawasaki_project
```

---

## 0　まず今の状態を記録する（読み取りだけ）

```bash
git status --short
git log --oneline -5
ls -t 01_第10期_最新版成果品/*.docx | head -5     # 素案の最新版
```

素案の版は `fix_soan_vNNN*.py` を**前の版の docx に当てて次の版を作る**
積み上げである。最初から作り直すものではない。
点検するだけなら**作り直さない**。

```bash
S=01_第10期_最新版成果品/川崎町_計画書素案_v2.6_図表整理版.docx
python3 - <<'EOF'
import docx, sys, os, glob
p = sorted(glob.glob("01_第10期_最新版成果品/川崎町_計画書素案_v*.docx"))[-1]
d = docx.Document(p)
print(os.path.basename(p), "段落", len(d.paragraphs), "表", len(d.tables))
EOF
```

---

## 1　足りないものだけ入れる

コンテナを作り直すたびに消えるものがある。

```bash
python3 -c "import docx, openpyxl, pypdf; print('docx/openpyxl/pypdf ok')"
python3 -c "import matplotlib; print('matplotlib ok')" || pip install matplotlib
which soffice || echo "LibreOffice がない（PDF 化・頁数の点検は未実施）"
```

日本語のフォントは環境により名が違う。**決め打ちにしない。**

```bash
fc-list :lang=ja file | head -5
# 本コンテナにあるもの
#   /usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf
#   /usr/share/fonts/truetype/fonts-japanese-gothic.ttf
# NotoSansCJK は無い。游ゴシックも無い
```

**⚠ コンテナに游ゴシックが入っていないため、LibreOffice の PDF 化では
別の書体に置き換わる。** PDF は**頁数・頁の割付けを数えるためだけ**に用い、
見え方の確認には用いない。見え方は町のパソコンで確かめていただく。

---

## 2　第1層　機械で確かめる

```bash
S=01_第10期_最新版成果品/川崎町_計画書素案_v2.6_図表整理版.docx

python3 07_ソーススクリプト/check_soan_selfcheck_R8.9.28.py "$S"   # 検査1〜6
python3 07_ソーススクリプト/check_hokenryo_R8.9.30.py       "$S"   # 検算1〜9・39項目
python3 07_ソーススクリプト/check_font_R8.9.30.py           "$S"   # 検査1〜8・13項目
python3 07_ソーススクリプト/build_zuhyo_daicho.py                   # 図表 自己点検8件
python3 07_ソーススクリプト/check_2kai_selfcheck_R8.9.28.py         # 委員会資料 27件
```

**終了コード0でなければそこで止まる。**

```bash
for f in check_soan_selfcheck_R8.9.28 check_hokenryo_R8.9.30 check_font_R8.9.30; do
  python3 "07_ソーススクリプト/$f.py" "$S" >/dev/null 2>&1 \
    && echo "○ $f" || echo "× $f（終了コード1）"
done
```

### docx の検証

場所は環境による。**存在を確かめてから呼び、無ければ「未実施」と記す。**

```bash
V=${DOCX_VALIDATE:-/mnt/skills/public/docx/scripts/office/validate.py}
[ -f "$V" ] && python3 "$V" "$S" || echo "docx の検証は未実施（validate.py が見つからない）"
```

### 目次の頁番号

見出しを増減したとき、本文・表・図が増減したとき、
**図の大きさを変えたときも**作り直す。

```bash
python3 07_ソーススクリプト/fill_toc_pages_v200.py "$S"
```

出力に次の3つが揃っていることを確かめる。

- `未特定 0件`
- `目次自身の頁を拾っている行はありません。`
- `頁番号は単調増加です。`

**「未特定0件」だけでは足りない。** Ver.2.3 までは全行に数が入っていながら
ご挨拶・第1章〜第4章の24行がすべて目次自身の頁（3）を指していた。

---

## 3　第2層　数値の出所を辿る

### 図の数値

正本は `07_ソーススクリプト/data_zuhyo.py` の1か所。
直したら必ず次の順で実行する。

```bash
python3 07_ソーススクリプト/build_plan_figures.py    # 08_図表 の png を作り直す
python3 07_ソーススクリプト/build_zuhyo_daicho.py    # 台帳と自己点検
python3 07_ソーススクリプト/fix_soan_v260_figures.py # 素案へ差し替え（次の版を作る）
python3 07_ソーススクリプト/fill_toc_pages_v200.py "$S_NEW"
```

### 素案に現にある章節を出す

反映先を書くときは、**素案に現にある**章節から採る。

```bash
python3 - <<'EOF'
import docx, re, sys
p = sys.argv[1] if len(sys.argv) > 1 else \
    "01_第10期_最新版成果品/川崎町_計画書素案_v2.6_図表整理版.docx"
for q in docx.Document(p).paragraphs:
    t = q.text.strip()
    if re.match(r"^第\s*\d+\s*章", t) or re.match(r"^\d+-\d+[　 ]", t):
        print(t[:60])
EOF
```

### 表の中身を見る

```bash
python3 - <<'EOF'
import docx
d = docx.Document("01_第10期_最新版成果品/川崎町_計画書素案_v2.6_図表整理版.docx")
for i, t in enumerate(d.tables, 1):
    head = " / ".join(c.text.strip()[:12] for c in t.rows[0].cells)
    print(i, len(t.rows), "行", len(t.columns), "列", head[:90])
EOF
```

### ⚠ 副作用のあるモジュール

`project_mikomi_R8.9.15.py` は **import しただけで
`川崎町_第10期_計画見込量_R8.9.15.xlsx` を書き換える。**
読んだあとは必ず戻す。

```bash
git status --short -- "05_試算・管理シート"
git checkout -- "05_試算・管理シート/川崎町_第10期_計画見込量_R8.9.15.xlsx"
```

---

## 4　第3層　守るべき制約の走査

本文・表（入れ子を含む）・ヘッダー・フッター・テキストボックスを走査する。
**走査できなかった要素は件数で示す。**

```bash
python3 - <<'EOF'
import docx, re, sys
from docx.oxml.ns import qn

P = sys.argv[1] if len(sys.argv) > 1 else \
    "01_第10期_最新版成果品/川崎町_計画書素案_v2.6_図表整理版.docx"
d = docx.Document(P)

def texts():
    """本文・表・ヘッダー・フッターの段落の文字を、場所とともに返す。"""
    for i, p in enumerate(d.paragraphs, 1):
        yield ("本文 P%d" % i, p.text)
    for ti, t in enumerate(d.tables, 1):
        for ri, r in enumerate(t.rows, 1):
            for ci, c in enumerate(r.cells, 1):
                yield ("表%d[%d,%d]" % (ti, ri, ci), c.text)
    for si, s in enumerate(d.sections, 1):
        for nm, part in (("ヘッダー", s.header), ("フッター", s.footer)):
            for p in part.paragraphs:
                yield ("%s%d" % (nm, si), p.text)

NG = {
    "受託者の語": ["受託者", "当社", "ビズアップ", "確認事項No",
                   "スクリプト", "再実行", "判定しています"],
    "他団体名":   ["大雪", "東川町", "東神楽町", "上川町"],
    "強調記号":   ["**", "__"],
}
YOHANDAN = {
    "禁止表現": ["に由来する", "と整合する", "1件も", "有意差がないため",
                 "全国トップ級"],
}
PAT = {
    "電話番号":       re.compile(r"0\d{1,4}-\d{1,4}-\d{3,4}"),
    "メールアドレス": re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+"),
    "［］以外の括弧": re.compile(r"[\[【〔].{0,20}?[\]】〕]"),
}
n = 0
for where, t in texts():
    n += 1
    for kind, ws in NG.items():
        for w in ws:
            if w in t:
                print("不適合 %-8s %-14s %s" % (kind, where, t.strip()[:60]))
    for kind, ws in YOHANDAN.items():
        for w in ws:
            if w in t:
                print("要判断 %-8s %-14s %s" % (kind, where, t.strip()[:60]))
    for kind, pat in PAT.items():
        m = pat.search(t)
        if m:
            print("要判断 %-8s %-14s %s" % (kind, where, m.group(0)[:40]))
print("走査した要素 %d" % n)
print("⚠ 図の中の文字・画像・氏名・住所・自由記述は機械では拾えない（未実施）")
EOF
```

**当たったものは候補である。** 文脈を見て判定を分ける
（「大雪」は地名として本文に出ることはないが、対比の成果品では正当）。

---

## 5　成果品を送付する前に

```bash
# ① すべての点検を通す
# ② 目次の頁番号を作り直す
# ③ 頁数を数える
python3 - <<'EOF'
import os, subprocess, tempfile, pypdf
S = "01_第10期_最新版成果品/川崎町_計画書素案_v2.6_図表整理版.docx"
with tempfile.TemporaryDirectory() as wd:
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf",
                    "--outdir", wd, os.path.abspath(S)],
                   check=True, env=dict(os.environ, HOME=wd),
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    f = [x for x in os.listdir(wd) if x.endswith(".pdf")][0]
    print("総頁数", len(pypdf.PdfReader(os.path.join(wd, f)).pages))
EOF
# ④ 別添整理表・業務工程管理表を更新する
# ⑤ commit して push する
git add -A && git commit && git push -u origin claude/kawasaki-town-handoff-skn48x
```

**PR は明示の依頼があるときだけ作る。**
