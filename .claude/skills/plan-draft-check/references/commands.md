# 機械で確かめる手順

## 0　環境を整える

セッションが替わるとパッケージが入っていない。

```bash
pip install python-docx openpyxl matplotlib pymupdf numpy scipy pillow
apt-get update && apt-get install -y libreoffice-writer libreoffice-calc
```

`apt-get update` を省くと 404 で失敗する。
docx→PDF には writer、xlsx→PDF には calc が要る（core だけでは足りない）。

作業ツリーが別の内容になっていることもある。

```bash
git fetch origin claude/daisetsu-care-plan-docs-pvjqid
git reset --hard origin/claude/daisetsu-care-plan-docs-pvjqid
```

## 1　第1層　機械で確かめる

```bash
python3 build_plan_draft.py          # 協議用素案
python3 build_plan_public.py         # 公表版（協議用素案を読んで組み替える）
python3 build_soan_kisai_taihi.py    # 網羅性の対比・レッドチーム
python3 build_soan_kadai.py          # 未確定箇所と確認事項
python3 build_process_control.py     # 業務工程管理表（確認事項の台帳）
```

**いずれも終了コード0でなければそこで止まる。**

docx の検証（作ったら必ず行う）。

```bash
for f in "output/第10期介護保険事業計画_協議用素案_令和8年8月.docx" \
         "output/第10期介護保険事業計画_公表版.docx"; do
  python3 /mnt/skills/public/docx/scripts/office/validate.py "$f"
done
```

素案の規模を実物から数える。

```bash
python3 -c "import repo_paths as RP; print(RP.draft_label())"
```

見出しを増減したら目次のページ番号を作り直す（ページ番号が変わらなくなるまで繰り返す）。

```bash
python3 build_toc_pages.py
python3 build_public_toc.py
```

紙面の点検（体裁を変えたら実行する）。

```bash
python3 tools/check_pages.py
```

再出力しただけの成果品を戻す。**本文の文字だけで比べない**（zip の中身で比べる）。

```bash
python3 tools/restore_unchanged.py
```

## 2　第2層　数値の出所を辿る

固定値の書き写しを探す。算定の値は runpy で読んでいるか。

```bash
grep -n "runpy" build_plan_draft.py | head
```

自己点検を持つスクリプトと件数を数える。

```bash
python3 - <<'PY'
import ast, glob
n = s = 0
for f in glob.glob("build_*.py"):
    try:
        t = ast.parse(open(f, encoding="utf-8").read())
    except Exception:
        continue
    c = sum(1 for x in ast.walk(t)
            if isinstance(x, ast.Call) and getattr(x.func, "id", "") == "chk")
    if c:
        n += 1
        s += c
print("自己点検を持つ %d本・計 %d件" % (n, s))
PY
```

確認事項の台帳に欠番・重複がないかを確かめる。

```bash
python3 - <<'PY'
import ast
src = open("build_process_control.py", encoding="utf-8").read()
for n in ast.parse(src).body:
    if isinstance(n, ast.Assign) and getattr(n.targets[0], "id", None) == "CHECK":
        C = ast.literal_eval(n.value)
no = sorted(x[0] for x in C)
print("件数 %d／最大 %d" % (len(no), max(no)))
print("欠番", [i for i in range(1, max(no) + 1) if i not in set(no)])
print("重複", [i for i in set(no) if no.count(i) > 1])
PY
```

## 3　第3層　記述の当否と守るべき制約

素案（協議用・公表版）の本文と表をまとめて走査する。

```bash
python3 - <<'PY'
import re
from docx import Document
import repo_paths as RP

# 「協議用」は協議用素案の表紙にある語であり、公表版でだけ見る。
FILES = [(RP.DRAFT, False),
         (RP.ROOT + "/output/第10期介護保険事業計画_公表版.docx", True)]
NAIBU = ("受託者", "本素案", "修正指示書", "確認事項No",
         "固定値", "実物から", "再実行", "章節ごとに", "判定しています",
         "runpy", ".py", "書き写して")
NAIBU_PUB = NAIBU + ("協議用",)
KIN = ("に由来する", "1件も", "全国トップ級", "有意差がないため")
TADAN = ("北塩原村", "浜田地区", "浜田圏域", "川崎町", "金ヶ崎町",
         "空知中部", "日高中部", "後志広域", "京極町")
OK_CH = re.compile(r"[ぁ-んァ-ヴ一-龥々ー０-９0-9A-Za-z"
                   r"、。・（）「」〜％　 \n\t（）［］【】〔〕：／\-—…‐"
                   r"＋±×÷＝＜＞％　°Ⅰ-Ⅻ①-⑳㎡㎞　]")

for f, pub in FILES:
    d = Document(f)
    t = "\n".join(p.text for p in d.paragraphs)
    for tb in d.tables:
        for r in tb.rows:
            for c in r.cells:
                t += "\n" + c.text
    print("====", f.split("/")[-1])
    print(" 内部の語 ", [w for w in (NAIBU_PUB if pub else NAIBU) if w in t]
          or "なし")
    print(" 禁止表現 ", [w for w in KIN if w in t] or "なし")
    print(" 他団体名 ", [w for w in TADAN if w in t] or "なし")
    print(" 電話番号 ", len(re.findall(r"0\d{1,4}-\d{1,4}-\d{3,4}", t)))
    print(" メール   ", len(re.findall(r"[\w.+-]+@[\w.-]+\.\w+", t)))
    print(" 未確定   ", len(re.findall(r"［(要協議|要確認|要内訳)］", t)))
    print(" 見出しの受託者", [p.text for p in d.paragraphs
                              if p.text.strip().startswith("【")
                              and "受託者" in p.text] or "なし")
PY
```

**個人情報は値の形（正規表現）で探す。** 語で探すと偽陽性が出る。

素案の章節の一覧（反映先を書く前に突き合わせる）。

```bash
python3 - <<'PY'
import re
from docx import Document
from docx.text.paragraph import Paragraph
import repo_paths as RP

d = Document(RP.DRAFT)
SHO = re.compile(r"^(第\d+章|資料\d+|資料編)")
SETSU = re.compile(r"^(第\d+節|基本目標\d)")
sho = ""
for ch in d.element.body.iterchildren():
    if ch.tag.split("}")[-1] != "p":
        continue
    t = Paragraph(ch, d).text.strip()
    if "\t" in t:
        continue
    if SHO.match(t):
        sho = t.split()[0]
        print(sho)
    elif SETSU.match(t):
        print("   ", sho, t)
PY
```

## 4　成果品を送付する前に

```bash
python3 build_deliverable_index.py    # 登録漏れ・実体なしを確かめる
python3 tools/restore_unchanged.py    # 再出力しただけのものを戻す
git status --short
```

新しい成果品を加えたら `data_dispatch.py` の `DISPATCH` にも登録する。
区分が「送付」のものは、**内部の仕組みの語の走査を自己点検に置く。**
