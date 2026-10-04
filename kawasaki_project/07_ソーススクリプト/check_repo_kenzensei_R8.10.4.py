# -*- coding: utf-8 -*-
"""リポジトリの健全性の点検（令和8年10月4日）

金ヶ崎町のブランチで、**計画素案の本文を作るスクリプト（ch_*.py）と
ビルド結果をスクラップパッドにのみ置く運用であったため、
環境が作り直された時点で消え、文言1か所の直しもできなくなった**事象が
起きた。同じことが川崎町で起きないかを機械で点検する。

スクラップパッドは**セッションごとに作り直される一時の場所**であり、
そこにしかないものは失われる。成果品を再び作れる状態を保つには、
次がすべてリポジトリに追跡されている必要がある。

  検査1　成果品（docx・xlsx・png）が追跡されているか
  検査2　作成のスクリプトが追跡されているか
  検査3　スクリプトが、リポジトリの外のパスを入力にしていないか
  検査4　素案の版の連なりが切れていないか
         （fix_soan_vNNN.py の入力の docx がリポジトリにあるか）
  検査5　図の数値の正本（data_zuhyo.py）と画像が揃っているか
  検査6　.gitignore が成果品を除外していないか
  検査7　未追跡・未コミットの変更が残っていないか
  検査8　アップロード領域の受領資料が、リポジトリに退避されているか
         （中身の MD5 で突合する。ファイル名は変えてよい）

  python3 07_ソーススクリプト/check_repo_kenzensei_R8.10.4.py

1件でも不適合があれば終了コード1で終わる。
"""
import hashlib
import os
import re
import subprocess
import sys

ROOT = "kawasaki_project"
UPLOAD = "/root/.claude/uploads"
# 成果品を置く場所
SEIKA_DIRS = ["01_第10期_最新版成果品", "03_委員会・説明資料",
              "05_試算・管理シート", "08_図表"]
SCRIPT_DIR = "07_ソーススクリプト"
# リポジトリの外（セッションごとに消える）を指すパス
SOTO = re.compile(r"(/root/\.claude|/tmp/claude-|scratchpad|/home/claude)")


def git(*args):
    return subprocess.run(["git"] + list(args), capture_output=True,
                          text=True, cwd="..").stdout


def tracked():
    """追跡しているファイルの一覧。

    ⚠ git は既定で非ASCIIのパスを \343\201… の形に escape する
    （core.quotepath）。そのままでは日本語のファイル名が一致しないため、
    -z（NUL区切り・escape しない）で受け取る。
    """
    out = subprocess.run(["git", "ls-files", "-z"], capture_output=True,
                         text=True, cwd="..").stdout
    return set(x for x in out.split("\0") if x)


def zip_naka(path, repo_md5):
    """書庫の中のファイルがすべてリポジトリに退避されているか。"""
    import zipfile
    try:
        z = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, OSError):
        return False
    names = [n for n in z.namelist() if not n.endswith("/")]
    if not names:
        return False
    for n in names:
        if hashlib.md5(z.read(n)).hexdigest() not in repo_md5:
            return False
    return True


def main():
    os.chdir(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
    tr = tracked()
    ng, note = [], []

    # ══════════════════════════ 検査1　成果品
    EXT = (".docx", ".xlsx", ".png", ".pdf", ".md")
    seika, untracked = 0, []
    for d in SEIKA_DIRS:
        for dirpath, _, files in os.walk(d):
            for f in files:
                if not f.endswith(EXT) or f.startswith("~$"):
                    continue
                p = os.path.join(dirpath, f)
                seika += 1
                if f"{ROOT}/{p}" not in tr:
                    untracked.append(p)
    if untracked:
        ng.append(f"検査1　追跡されていない成果品 {len(untracked)}件："
                  + "／".join(untracked[:5]))

    # ══════════════════════════ 検査2　スクリプト
    scripts, s_untracked = 0, []
    for f in sorted(os.listdir(SCRIPT_DIR)):
        if not f.endswith(".py"):
            continue
        scripts += 1
        if f"{ROOT}/{SCRIPT_DIR}/{f}" not in tr:
            s_untracked.append(f)
    if s_untracked:
        ng.append(f"検査2　追跡されていないスクリプト {len(s_untracked)}件："
                  + "／".join(s_untracked[:5]))

    # ══════════════════════════ 検査3　外のパスを入力にしていないか
    soto = []
    for f in sorted(os.listdir(SCRIPT_DIR)):
        if not f.endswith(".py"):
            continue
        p = os.path.join(SCRIPT_DIR, f)
        with open(p, encoding="utf-8", errors="ignore") as fh:
            for i, line in enumerate(fh, 1):
                if line.lstrip().startswith("#"):
                    continue
                m = SOTO.search(line)
                if m:
                    soto.append((f, i, line.strip()[:76]))
    # 入力として使っているもの（SRC_DIR 等）を取り出す
    soto_in = [x for x in soto
               if re.search(r"(SRC|IN|INPUT|UPLOAD)", x[2], re.I)]
    soto_out = [x for x in soto if x not in soto_in]
    # ⚠ 外のパスを**代替**として置くことは差し支えない。
    #   「リポジトリの中の相対パスを第一の入力とし、存在を確かめてから
    #   外へ落ちる」形になっていれば適合とする。その形であることを、
    #   ①リポジトリ相対のパスの定数 ②os.path.exists による確認
    #   の双方が同じファイルにあることで見る。
    soto_hard = []
    for f, i, line in soto_in:
        src = open(os.path.join(SCRIPT_DIR, f), encoding="utf-8",
                   errors="ignore").read()
        has_repo = re.search(r'=\s*\(?\s*"0\d_[^"]+"', src)
        has_guard = "os.path.exists" in src
        if not (has_repo and has_guard):
            soto_hard.append((f, i, line))
    soto_soft = [x for x in soto_in if x not in soto_hard]
    if soto_hard:
        ng.append(f"検査3　リポジトリの外だけを入力にしているスクリプト "
                  f"{len(soto_hard)}件")
        for f, i, line in soto_hard:
            ng.append(f"　　{f}:{i}  {line}")
    if soto_soft:
        note.append(f"検査3　外のパスを代替の入力として置いている箇所 "
                    f"{len(soto_soft)}件"
                    "（リポジトリの中を第一の入力にしたうえでの代替）")
    if soto_out:
        note.append(f"検査3　外のパスを一時の置き場として使っている箇所 "
                    f"{len(soto_out)}件（出力先・作業用。入力ではない）")

    # ══════════════════════════ 検査4　素案の版の連なり
    chain, broken = [], []
    for f in sorted(os.listdir(SCRIPT_DIR)):
        if not re.match(r"fix_soan_v\d+", f):
            continue
        p = os.path.join(SCRIPT_DIR, f)
        src = dst = None
        with open(p, encoding="utf-8", errors="ignore") as fh:
            for line in fh:
                m = re.match(r'\s*SRC\s*=\s*"([^"]+\.docx)"', line)
                if m:
                    src = m.group(1)
                m = re.match(r'\s*DST\s*=\s*"([^"]+\.docx)"', line)
                if m:
                    dst = m.group(1)
        if src:
            chain.append((f, src, dst))
            if not os.path.exists(src):
                broken.append(f"{f} の入力 {os.path.basename(src)} がない")
            elif f"{ROOT}/{src}" not in tr:
                broken.append(f"{f} の入力 {os.path.basename(src)} が"
                              "追跡されていない")
    if broken:
        ng.append("検査4　素案の版の連なりが切れている：" + "／".join(broken))

    # ══════════════════════════ 検査5　図の正本と画像
    sys.path.insert(0, SCRIPT_DIR)
    try:
        from data_zuhyo import ZU
        miss = [d["png"] for d in ZU
                if not os.path.exists(os.path.join("08_図表", d["png"]))]
        miss_tr = [d["png"] for d in ZU
                   if f"{ROOT}/08_図表/{d['png']}" not in tr]
        if miss:
            ng.append("検査5　data_zuhyo.py の図の画像がない："
                      + "／".join(miss))
        if miss_tr:
            ng.append("検査5　図の画像が追跡されていない："
                      + "／".join(miss_tr))
        n_zu = len(ZU)
    except Exception as e:
        ng.append(f"検査5　data_zuhyo.py を読めない：{e}")
        n_zu = 0

    # ══════════════════════════ 検査6　.gitignore
    ign = []
    gi = "../.gitignore"
    if os.path.exists(gi):
        with open(gi, encoding="utf-8", errors="ignore") as fh:
            for line in fh:
                s = line.strip()
                if not s or s.startswith("#"):
                    continue
                if re.search(r"\.(docx|xlsx|png|py)$|^\*|output|成果", s):
                    ign.append(s)
    if ign:
        note.append(".gitignore に成果品に関わりうる行がある："
                    + "／".join(ign[:6]))

    # ══════════════════════════ 検査7　未コミットの変更
    st = git("status", "--porcelain").splitlines()
    mine = [x for x in st if ROOT in x or x.startswith("?? .claude")]
    if mine:
        note.append(f"未コミットの変更 {len(mine)}件"
                    "（作業中であれば差し支えない）")

    # ══════════════════════════ 検査8　受領資料の退避
    #
    # ⚠ アップロード領域（/root/.claude/uploads）は**セッションごとに
    #   作り直される。** 受領した資料をそこに置いたままにすると、
    #   次のセッションでは原本がなく、成果品を作り直せない。
    #   ファイル名は内容の分かる形に改めてよいので、**中身の MD5** で突合する。
    up_total = up_miss = 0
    up_note = ""
    if os.path.isdir(UPLOAD):
        repo_md5 = {}
        for dirpath, dirs, files in os.walk("."):
            if "/.git" in dirpath:
                continue
            for f in files:
                p = os.path.join(dirpath, f)
                try:
                    with open(p, "rb") as fh:
                        repo_md5.setdefault(
                            hashlib.md5(fh.read()).hexdigest(), p)
                except OSError:
                    pass
        miss = []
        for dirpath, dirs, files in os.walk(UPLOAD):
            for f in files:
                p = os.path.join(dirpath, f)
                up_total += 1
                try:
                    with open(p, "rb") as fh:
                        k = hashlib.md5(fh.read()).hexdigest()
                except OSError:
                    continue
                if k in repo_md5:
                    continue
                # 書庫（zip）は、中身が退避されていれば退避済みとみなす
                if f.lower().endswith(".zip") and zip_naka(p, repo_md5):
                    continue
                miss.append(f)
        up_miss = len(miss)
        if miss:
            ng.append(f"検査8　リポジトリに退避されていない受領資料 "
                      f"{len(miss)}件：" + "／".join(miss[:6]))
    else:
        up_note = "検査8　未実施（アップロード領域がない。" \
                  "セッションが作り直された後は確かめられない）"
        note.append(up_note)

    # ══════════════════════════ 出力
    print("リポジトリの健全性の点検")
    print(f"  成果品 {seika}件／スクリプト {scripts}件"
          f"／素案の版の連なり {len(chain)}段／図 {n_zu}点")
    print(f"  追跡されている成果品 {seika - len(untracked)}／{seika}")
    print(f"  追跡されているスクリプト {scripts - len(s_untracked)}／{scripts}")
    for m in note:
        print("   △", m)
    if ng:
        for m in ng:
            print("   ×", m)
        print("  ⚠ スクラップパッドはセッションごとに作り直されます。"
              "そこにしかないものは失われます。")
        sys.exit(1)
    print("   ○ 成果品・スクリプト・図の正本・版の連なりは"
          "すべてリポジトリに追跡されている")
    print("   ○ スクラップパッドにしかない入力はない")
    if up_total and not up_miss:
        print(f"   ○ 受領資料 {up_total}件はすべてリポジトリに"
              "退避されている（MD5で突合）")


if __name__ == "__main__":
    main()
