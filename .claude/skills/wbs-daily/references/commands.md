# 機械で確かめる手順

下のコード片はいずれもそのまま実行できる（リポジトリの直下で動かす）。

---

## 0　環境を整える（セッションが替わったとき）

```bash
python3 -c "import openpyxl, docx" 2>/dev/null || \
  pip install -q python-docx openpyxl matplotlib pymupdf numpy scipy pillow
```

docx を PDF にするときは `apt-get update && apt-get install -y libreoffice-writer`、
xlsx のときは `libreoffice-calc`（`apt-get update` を省くと404で失敗する）。

**開発ブランチの作業ツリーが別の内容になっていることがある。**
始めに**読み取りだけで**状態を記録する。

```bash
git -C . rev-parse --short HEAD
git -C . status --short
git -C . log --oneline -5
```

`git reset --hard` は通常の手順に置かない。
戻すのは、`git status` が空であることを確かめたうえで、必要なときだけ。

---

## 1　その日の作業を拾う

```bash
D=$(date +%F)
git log --since="$D 00:00" --until="$D 23:59" --format='%h %s'
git log --since="$D 00:00" --until="$D 23:59" --name-only --format='--- %s' | head -80
```

受領資料は作業領域にある（セッションが替わると消える）。

```bash
ls -la /root/.claude/uploads/*/ 2>/dev/null | head -20
```

---

## 2　WBSを更新する

`data_progress.py` を次の順で直す。

1. 進捗率を動かすなら、**更新前の値を `KIROKU`・`KIROKU_UGOKI` へ移す**
2. `PROGRESS` の値と動きの欄を書き換える
3. `KIJUNBI` を当日に進める
4. `SHINCHOKU_RIYU_R8_MM_DD` を新設する

移し忘れていないかを確かめる。

```bash
python3 - <<'PY'
import data_progress as DP
print("基準日", DP.KIJUNBI, "／全体", DP.overall_pct(), "％")
print("記録の最新", sorted(DP.KIROKU)[-1])
assert DP.KIJUNBI not in DP.KIROKU, "現時点の値が記録に紛れている"
prev = sorted(DP.KIROKU)[-1]
for no, nm, v, st, d in DP.PROGRESS:
    p = DP.KIROKU[prev].get(no)
    if v is None or p is None:
        continue
    ugoki = round((v - p) * 100)
    print("%-5s %-28s %4.0f%% (記録 %3.0f%% / 差 %+d pt) 欄=%s"
          % (no, nm, v * 100, p * 100, ugoki, d))
PY
```

**差（pt）と動きの欄が食い違っていたら、どちらかが誤っている。**

```bash
python3 build_process_control.py
```

---

## 3　日次の表を作る

```bash
python3 build_nikkan.py              # 基準日は実行した日
python3 build_nikkan.py 2026-10-01   # 基準日を指定する
echo "終了コード $?"
```

終了コード1で終わったときは06シートを見る。

```bash
python3 - <<'PY'
import openpyxl, repo_paths as RP, os
wb = openpyxl.load_workbook(
    os.path.join(RP.OUTPUT, "第10期計画_日次の状況と翌日の作業順位.xlsx"))
for row in wb["06_自己点検"].iter_rows(min_row=4, values_only=True):
    if row and row[4] == "不適合":
        print(row[0], row[1], "｜", row[3])
PY
```

**よく出る不適合**

| 点検 | 意味 | 直し方 |
|---|---|---|
| 8 | 進捗の記録の基準日が当日でない | 先にWBSを更新する |
| 1 | 確認事項に欠番がある | 別の成果品で起票したものを台帳へ登録する |
| 12・13 | 禁止表現・個人情報の形が出た | 本表の文言を直す（台帳の文言のこともある） |

---

## 4　台帳を読む（確認事項・資料提供依頼）

```bash
python3 - <<'PY'
import ast, io, collections
src = io.open("build_process_control.py", encoding="utf-8").read()
g = {}
for n in ast.parse(src).body:
    if isinstance(n, ast.Assign):
        for t in n.targets:
            if getattr(t, "id", None) in ("CHECK", "LACK"):
                g[t.id] = ast.literal_eval(n.value)
C, L = g["CHECK"], g["LACK"]
KANRYO = ("完了", "了承済", "了承済（保管せず廃棄）", "代替により解消", "解決")
M = [x for x in C if x[7] not in KANRYO]
print("確認事項 %d件（未決 %d件）" % (len(C), len(M)))
print("状態 ", collections.Counter(x[7] for x in C).most_common())
print("期限 ", collections.Counter(x[8] for x in M).most_common(8))
no = sorted(x[0] for x in C)
print("欠番 ", [n for n in range(1, max(no) + 1) if n not in set(no)] or "なし")
print("資料提供依頼 %d件（未受領 %d件）"
      % (len(L), sum(1 for x in L if x[8] not in ("受領済", "完了", "解消"))))
PY
```

**同じ問いが既に台帳にないかを語で探す**（言い換えで二重に起票したことがある）。

```bash
python3 - <<'PY'
import ast, io, sys
WORD = "認知症"          # ← 探す語
src = io.open("build_process_control.py", encoding="utf-8").read()
for n in ast.parse(src).body:
    if isinstance(n, ast.Assign) and any(getattr(t, "id", None) == "CHECK"
                                         for t in n.targets):
        C = ast.literal_eval(n.value)
for x in C:
    if WORD in str(x[3]) + str(x[4]):
        print("No.%-4d [%s] %s" % (x[0], x[7], x[3]))
PY
```

---

## 5　作業継続可能なものを確かめる

自己点検を備えたスクリプトを数え、全て再実行する。

```bash
python3 - <<'PY'
import ast, glob, io, os, subprocess
# 受領資料をセッション固有の場所から読むため再実行できないもの
NG = {"build_kickoff_fix.py", "build_minutes_fix.py",
      "build_survey_crosstab.py"}
tgt = []
for f in sorted(glob.glob("build_*.py")):
    if os.path.basename(f) in NG:
        continue
    src = io.open(f, encoding="utf-8").read()
    n = sum(1 for node in ast.walk(ast.parse(src))
            if isinstance(node, ast.Call)
            and getattr(node.func, "id", "") == "chk")
    if n:
        tgt.append((f, n))
print("自己点検を備えたスクリプト %d本・計%d件"
      % (len(tgt), sum(n for _f, n in tgt)))
bad = []
for f, _n in tgt:
    rc = subprocess.run(["python3", f], capture_output=True).returncode
    if rc:
        bad.append(f)
print("不適合で終わったもの:", bad or "なし")
PY
```

**実行に時間がかかる。** 毎日は回さず、締めの前か週に1回でよい。

---

## 6　締める

```bash
python3 tools/xlsx_zuhyo_check.py output/第10期計画_日次の状況と翌日の作業順位.xlsx
python3 tools/restore_unchanged.py
git -C . add -A
git -C . commit -m "日次 WBSの更新と翌日の作業順位（<基準日>）"
git -C . push -u origin claude/daisetsu-care-plan-docs-pvjqid
```

押せないときは2秒・4秒・8秒・16秒と間を置いて4回まで繰り返す。
**プルリクエストは明示の依頼がない限り作らない。**

---

## 7　件数を述べている成果品を作り直す

台帳を増減したら必要になる。

```bash
python3 tools/suchi_impact.py --map 173      # 古い件数がどこに残っているか
python3 build_check_irai.py
python3 build_kofukin_torimatome.py
python3 build_kanri_tanaoroshi.py
```
