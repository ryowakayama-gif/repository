# 機械で確かめる手順

実行できなかった手順は**「未実施」として記録し、適合に丸めない。**

## 0　まず今の状態を記録する（読み取りだけ）

点検は**対象の版を変えずに**始める。何を点検したのかが後から分かるようにする。

```bash
git rev-parse --short HEAD
git status --short          # 未コミットの作業があるかを見る
ls -l --time-style=long-iso output/*.docx output/*.xlsx | head -20
```

**未コミットの作業があるときは、それが点検の対象である。**
勝手に戻さない。戻すかどうかは作業している人が決める。

## 1　足りないものだけ入れる

まず実行し、`ModuleNotFoundError` が出たものだけ入れる。

```bash
python3 -c "import docx, openpyxl, matplotlib, fitz, numpy, scipy, PIL" \
  || pip install python-docx openpyxl matplotlib pymupdf numpy scipy pillow
```

docx→PDF（紙面の点検・目次のページ番号）には LibreOffice の Writer、
xlsx→PDF には Calc が要る。**無い環境では紙面の点検を「未実施」とする。**

```bash
command -v soffice >/dev/null \
  || echo "LibreOffice が無い。紙面の点検・目次のページ番号は未実施"
# 入れられる環境なら（core だけでは足りない。update を先に行う）
# apt-get update && apt-get install -y libreoffice-writer libreoffice-calc
```

## 1-2　作業ツリーが壊れているときだけ（例外の手順）

**通常の手順ではない。** 未コミットの作業が無いことを `git status` で
確かめたうえで、必要なときに限り行う。破棄したくない変更があるときは
先に `git stash` する。

```bash
git status --short            # 空であることを確かめる
git fetch origin claude/daisetsu-care-plan-docs-pvjqid
git switch claude/daisetsu-care-plan-docs-pvjqid
# それでも直らないときに限り（追跡中の未コミット変更は失われる）
# git reset --hard origin/claude/daisetsu-care-plan-docs-pvjqid
```

## 2　第1層　機械で確かめる

既にある成果品を点検するだけなら、**作り直さずに**次を実行する。

```bash
python3 build_soan_kisai_taihi.py    # 網羅性の対比・レッドチーム
python3 build_soan_kadai.py          # 未確定箇所と確認事項
python3 build_zuhyo_daicho.py        # 図の数値と図表集との突合
python3 build_process_control.py     # 業務工程管理表（確認事項の台帳）
```

素案そのものを直したときは、併せて作り直す。

```bash
python3 build_plan_figures.py        # 本文に差し込む図（data_zuhyo.py から）
python3 build_plan_draft.py          # 協議用素案
python3 build_plan_public.py         # 公表版（協議用素案を読んで組み替える）
```

**いずれも終了コード0でなければそこで止まる。**

docx の検証（作ったら必ず行う）。**場所は環境による。**

```bash
V=${DOCX_VALIDATE:-/mnt/skills/public/docx/scripts/office/validate.py}
if [ -f "$V" ]; then
  for f in "output/第10期介護保険事業計画_協議用素案_令和8年8月.docx" \
           "output/第10期介護保険事業計画_公表版.docx"; do
    python3 "$V" "$f"
  done
else
  echo "docx の検証は未実施（validate.py が見つからない）"
fi
```

素案の規模を実物から数える。

```bash
python3 -c "import repo_paths as RP; print(RP.draft_label())"
```

見出しを増減したとき、**及び本文・表・図が増減してページが動いたとき**は
目次のページ番号を作り直す（ページ番号が変わらなくなるまで繰り返す。
繰り返しても収まらないときは、何が動いているかを見てから進める）。

```bash
python3 build_toc_pages.py
python3 build_public_toc.py
```

紙面の点検。**体裁を変えたときだけでなく、本文・表・図を増減したときも行う。**

```bash
D="output/第10期介護保険事業計画_協議用素案_令和8年8月.docx"
soffice --headless --convert-to pdf --outdir /tmp "$D" >/dev/null
python3 tools/check_pages.py "/tmp/$(basename "${D%.docx}").pdf" "$D"
```

再出力しただけの成果品を戻す。**本文の文字だけで比べない**（zip の中身で比べる）。
画像（PNG）は作成日時が変わるだけのことがあるため、画素で比べて戻す。

```bash
python3 tools/restore_unchanged.py
```

## 3　第2層　数値の出所を辿る

固定値の書き写しを探す。算定の値は runpy で読んでいるか。
**`runpy` の字があることは、出所が正しいことの証しではない。**
出力の値から入力のセル・算式・単位・期間までたどれるかを見る。

```bash
grep -n "runpy" build_plan_draft.py | head
```

自己点検を持つスクリプトと件数を数える。
**これは `chk(` の呼出しの数であり、実際に走った検査の数ではない。**
構文解析に失敗したファイルは件数として別に示す。

```bash
python3 - <<'PY'
import ast, glob
n = s = 0
ng = []
for f in sorted(glob.glob("build_*.py")):
    try:
        t = ast.parse(open(f, encoding="utf-8").read())
    except Exception as e:
        ng.append((f, str(e)[:40]))
        continue
    c = sum(1 for x in ast.walk(t)
            if isinstance(x, ast.Call) and getattr(x.func, "id", "") == "chk")
    if c:
        n += 1
        s += c
print("自己点検を持つ %d本・計 %d件（chk の呼出しの数）" % (n, s))
print("構文解析できなかったもの", ng or "なし")
PY
```

確認事項の台帳に欠番・重複がないかを確かめる。
**番号が連続していても登録漏れは起きる**ため、成果品の側で挙げた論点と
併せて突き合わせる。

```bash
python3 - <<'PY'
import ast, sys
src = open("build_process_control.py", encoding="utf-8").read()
C = None
for n in ast.parse(src).body:
    if isinstance(n, ast.Assign) and getattr(n.targets[0], "id", None) == "CHECK":
        C = ast.literal_eval(n.value)
if C is None:
    print("CHECK を読めなかった（未実施）")
    sys.exit(2)
no = sorted(x[0] for x in C)
ketsu = [i for i in range(1, max(no) + 1) if i not in set(no)]
juf = [i for i in set(no) if no.count(i) > 1]
print("件数 %d／最大 %d" % (len(no), max(no)))
print("欠番", ketsu or "なし")
print("重複", juf or "なし")
sys.exit(1 if (ketsu or juf) else 0)
PY
```

## 4　第3層　守るべき制約の走査

**本文・入れ子を含む表・ヘッダー・フッター・テキストボックス**を走査する。
検出があれば終了コード1で終わる。

```bash
python3 - <<'PY'
import re, sys
from docx import Document
from docx.oxml.ns import qn
import repo_paths as RP


def texts(path):
    """文書の文字を、どこにあるかと対にして取り出す。"""
    d = Document(path)
    out = []

    def walk(el, where):
        # iter() は入れ子の表・テキストボックスの中の段落も文書順に返す
        for p in el.iter(qn("w:p")):
            t = "".join(n.text or "" for n in p.iter(qn("w:t")))
            if t.strip():
                out.append((where, t))

    walk(d.element.body, "本文")
    for i, s in enumerate(d.sections, start=1):
        for nm, part in (("ヘッダー", s.header), ("フッター", s.footer),
                         ("最初のページのヘッダー", s.first_page_header),
                         ("最初のページのフッター", s.first_page_footer)):
            try:
                walk(part._element, "第%d節の%s" % (i, nm))
            except Exception:
                pass
    return out


def norm(s):
    """表記のゆれをそろえる。走査の前に必ず通す。"""
    return re.sub(r"[ \t]+", " ", s.replace("%", "％").replace("　", " "))


NAIBU = ("受託者", "本素案", "修正指示書", "確認事項No", "固定値", "実物から",
         "再実行", "章節ごとに", "判定しています", "runpy", ".py", "書き写して")
NAIBU_PUB = NAIBU + ("協議用",)
KIN = ("に由来する", "と整合する", "1件も", "全国トップ級", "有意差がないため")
TEL = re.compile(r"0\d{1,4}[-－ー―]\d{1,4}[-－ー―]\d{3,4}|0[789]0\d{8}")
MAIL = re.compile(r"[\w.+-]+@[\w.-]+\.\w+")
PH = re.compile(r"［(要協議|要確認|要内訳)］")
PH_YURE = re.compile(r"[\[【〔]\s*(要協議|要確認|要内訳)\s*[\]】〕]")
# 許容する文字。当たらないものは別の言語の文字が混じった疑いがある（要判断）
OK_CH = re.compile(r"[ぁ-んァ-ヴー一-龥々〆ヶ０-９0-9A-Za-zＡ-Ｚａ-ｚ"
                   r"、。・（）「」『』【】〜％　 \n\t()［］〔〕：；／＼\-—…‐"
                   r"＋±×÷＝＜＞°Ⅰ-Ⅻ①-⑳㎡㎞㎏№→←↑↓※〇○●◇◆□■△▲★☆"
                   r"，．！？＆＃＄＠＊｜～－―≒⑴-⑿²³δ:=_"
                   r"~\"'`,.!?&#$@*|/\\]")

# 他団体名は収録している一覧から作る（書き写すと増減に追随しない）
JI = ("大雪地区広域連合", "東川町", "美瑛町", "東神楽町")
try:
    import data_kofukin_rengo as KR
    TADAN = sorted({r[1] for r in KR.RENGO} | {r[1] for r in KR.KAISAN}
                   | {m[0] for r in KR.RENGO for m in r[3]}
                   | {m[0] for r in KR.KAISAN for m in r[4]})
except Exception:
    TADAN = []
# 比較材料として読んだ団体は収録していないため、別に挙げる
HIKAKU = {"北塩原村", "浜田地区広域行政組合", "浜田圏域", "川崎町", "金ヶ崎町"}
TADAN = sorted((set(TADAN) | HIKAKU) - set(JI))
print("他団体名の一覧 %d件（当連合・構成3町を除く）" % len(TADAN))

FILES = [("協議用素案", RP.DRAFT, NAIBU),
         ("公表版", RP.ROOT + "/output/第10期介護保険事業計画_公表版.docx", NAIBU_PUB)]

FUTEKI = YOHANDAN = 0
for name, path, naibu in FILES:
    ts = texts(path)
    joined = "\n".join(norm(t) for _, t in ts)
    print("====", name, "　走査した段落 %d" % len(ts))
    hit = [(w, where) for w in naibu for where, t in ts if w in norm(t)]
    print("  内部の語［不適合］", sorted(set(w for w, _ in hit)) or "なし")
    for w, where in hit[:5]:
        print("      ［%s］%s" % (w, where))
    FUTEKI += len(set(w for w, _ in hit))
    tad = [w for w in TADAN if w in joined]
    print("  他団体名［不適合］", tad or "なし")
    FUTEKI += len(tad)
    yure = PH_YURE.findall(joined)
    print("  未確定箇所 %d件（正しい表記）／ほかの括弧［不適合］ %d件"
          % (len(PH.findall(joined)), len(yure)))
    FUTEKI += len(yure)
    kin = sorted({w for w in KIN if w in joined})
    print("  禁止表現［要判断］", kin or "なし")
    YOHANDAN += len(kin)
    tel, mail = TEL.findall(joined), MAIL.findall(joined)
    print("  電話番号［要判断］ %d件／メール［要判断］ %d件" % (len(tel), len(mail)))
    YOHANDAN += len(tel) + len(mail)
    gai = sorted(set(OK_CH.sub("", joined)))
    print("  許容する文字の外［要判断］", "".join(gai) if gai else "なし")
    YOHANDAN += len(gai)
    mido = [t for _, t in ts if t.strip().startswith(("【", "［")) and "受託者" in t]
    print("  見出しの受託者［不適合］", mido or "なし")
    FUTEKI += len(mido)
print("不適合 %d件／要判断 %d件" % (FUTEKI, YOHANDAN))
print("※ 氏名・住所・自由記述・画像の中の文字は機械では拾えない。目で確かめる。")
sys.exit(1 if FUTEKI else 0)
PY
```

**受入れの確認**（直したあとに、わざと違反を入れて落ちることを確かめる）

| 入れるもの | 期待する結果 |
|---|---|
| ヘッダーに「受託者」 | 位置付きで検出される |
| 入れ子の表の中に「受託者」 | 検出される |
| 本文に「[要確認]」 | ほかの括弧として検出される |
| 本文に「と整合する」 | 要判断として検出される |
| 資料6から1項目を落とす | 原典の期待リストとの差で検出される |

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
    if "\t" in t:          # 目次の行はページ番号をタブで連結している
        continue
    if SHO.match(t):
        sho = t.split()[0]
        print(sho)
    elif SETSU.match(t):
        print("   ", sho, t)
PY
```

## 5　成果品を送付する前に

```bash
python3 build_deliverable_index.py    # 登録漏れ・実体なしを確かめる
python3 tools/restore_unchanged.py    # 再出力しただけのものを戻す
git status --short
```

新しい成果品を加えたら `data_dispatch.py` の `DISPATCH` にも登録する。
区分が「送付」のものは、**内部の仕組みの語の走査を自己点検に置く。**

公表版を確定として出すのは、SKILL.md の「どこまで進めてよいかの区分」の
**公表版として確定してよい**の条件をすべて満たしたときに限る。
