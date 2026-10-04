"""小野町 計画策定業務　成果品がリポジトリだけで組み直せるかを点検する。

**なぜこの点検が要るか。**

金ケ崎町の案件で、成果品の本文（ch_*.py）とビルド結果をスクラップパッドに
置き、リポジトリには記録用のMDと体裁ファイルだけを残す運用にしていた。
スクラップパッドはセッションごとの一時領域であり、環境が作り直された
時点で消える。その結果、計画素案の文言1か所を直すこともできなくなった。

同じことが起きる条件は1つだけである。
**成果品を組み直すのに要るものが、リポジトリの外にあること。**

本点検は、追跡しているスクリプトが参照するファイルの場所を調べ、
リポジトリの外を向いているものを洗い出す。見落としやすいのは次の3つ。

 1. セッションのアップロード先（/root/.claude/uploads/<セッションID>/…）
    受領した原本を直に読んでいると、次のセッションでは読めない。
 2. スクラップパッド（/tmp/claude-*/…/scratchpad/…）
    作業用の置き場であり、残らない。
 3. 作業ディレクトリにあるが git に入れていないファイル
    手元では動くが、clone した人の手元では動かない。

一時的な作業先として /tmp に書き出すこと自体は差し支えない。
消えても作り直せるためである。**読み取り元が外にあることだけが問題になる。**

実行:
  python3 check_ono_jiritsu.py
"""

import ast
import pathlib
import re
import subprocess
import sys

ROOT = pathlib.Path(__file__).parent.resolve()

# 読み取り元が外を向いていたら成果品を組み直せなくなる場所。
SOTO = [
    (re.compile(r"/root/\.claude/uploads/"), "セッションのアップロード先"),
    (re.compile(r"/tmp/claude-[^/]*/.*scratchpad"), "スクラップパッド"),
    (re.compile(r"/home/[^/]+/(?!repository)"), "リポジトリ外のホーム"),
    (re.compile(r"/root/(?!\.claude/uploads)"), "リポジトリ外の /root"),
]

# 書き出し先としてなら差し支えないもの（消えても作り直せる）。
KAKIDASHI_OK = re.compile(r"^/tmp/(?!claude-)")

# 読み取りに使う関数・メソッドの名前。これらの引数に出てくるパスを見る。
YOMI = {"load_workbook", "open", "read_text", "read_bytes", "Document",
        "read_csv", "imread", "glob", "rglob", "iterdir"}


def paths_in(src, fname):
    """ソースから、文字列として書かれているパスらしきものを拾う。

    f文字列・連結・pathlib の / 演算子をまたぐため、構文木ではなく
    文字列リテラルを総当たりする。取りこぼすより拾いすぎるほうが安全である。
    """
    try:
        tree = ast.parse(src)
    except SyntaxError as e:
        return [("構文エラー", str(e), 0)]
    out = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Constant) and isinstance(node.value, str):
            v = node.value
            if "/" in v and len(v) > 3:
                out.append((v, fname, getattr(node, "lineno", 0)))
    return out


def git(*args):
    """git を呼ぶ。日本語のファイル名が \\346\\234\\254… とエスケープされると
    追跡の有無を取り違えるため、core.quotepath=false を明示する。
    リポジトリ側の設定に結果が左右されないようにするためである。
    """
    r = subprocess.run(["git", "-C", str(ROOT), "-c", "core.quotepath=false",
                        *args], capture_output=True, text=True)
    return r.stdout


def tracked_files():
    return set(git("ls-files").splitlines())


def main():
    bad, warn = [], []
    tracked = tracked_files()
    pys = sorted(p for p in tracked if p.endswith(".py"))

    # (1) 追跡スクリプトが、リポジトリ外を読んでいないか
    me = pathlib.Path(__file__).name
    for rel in pys:
        f = ROOT / rel
        if not f.exists():
            continue
        if f.name == me:
            continue        # 本点検の正規表現そのものを拾ってしまうため
        src = f.read_text(encoding="utf-8")
        for v, _fn, ln in paths_in(src, rel):
            for pat, naze in SOTO:
                if pat.search(v):
                    if KAKIDASHI_OK.match(v):
                        continue
                    bad.append((rel, ln, naze, v[:90]))

    # (2) 作業ディレクトリにあって追跡していないファイル（成果品・原本・スクリプト）
    for line in git("status", "--porcelain", "--untracked-files=all").splitlines():
        if not line.startswith("??"):
            continue
        rel = line[3:].strip().strip('"')
        if rel.endswith((".py", ".docx", ".xlsx", ".xls", ".pdf", ".csv", ".json")):
            warn.append(("追跡していない", rel))

    # (3) 成果品フォルダの中身がすべて追跡されているか
    seika = ROOT / "小野町_引継ぎ_整理済"
    if seika.exists():
        for f in seika.rglob("*"):
            if f.is_dir():
                continue
            rel = str(f.relative_to(ROOT))
            if rel not in tracked:
                warn.append(("成果品フォルダで未追跡", rel))

    print(f"点検対象　追跡スクリプト {len(pys)}本")
    if bad:
        print(f"\n■ リポジトリ外を読んでいる箇所 {len(bad)}件")
        print("  **環境が作り直されると、この成果品は組み直せなくなります。**")
        for rel, ln, naze, v in bad:
            print(f"  NG {rel}:{ln}　{naze}")
            print(f"       {v}")
    else:
        print("\n■ リポジトリ外を読んでいる箇所　なし")

    if warn:
        print(f"\n■ 追跡していないファイル {len(warn)}件")
        for naze, rel in warn[:30]:
            print(f"  ! {naze}　{rel}")
        if len(warn) > 30:
            print(f"  … ほか {len(warn) - 30}件")
    else:
        print("■ 追跡していないファイル　なし")

    ok = not bad and not warn
    print("\n" + ("自立性の点検　適合（リポジトリだけで組み直せます）"
                  if ok else "自立性の点検　要修正"))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
