# -*- coding: utf-8 -*-
"""納品する電子媒体の中身を組み立てる（仕様書5③／WBS Ⅳ-91）

nohin_data.N に「収録する」と定めたファイルだけを、媒体での置き場所ごとに
並べる。**収録しないと定めたファイルが1件でも入ったら止まる。**

  python3 scripts/build_nohin.py            何が入るかを一覧で示すだけ（組み立てない）
  python3 scripts/build_nohin.py --write    実際に組み立てる

組み立て先は output/_納品媒体/ である。版管理には入れない（.gitignore）。
中身は元のファイルの写しであり、ここを直しても元には戻らない。

媒体の構成
  01_計画書／02_調査報告書／03_委員会資料／04_図表データ／05_策定の記録／
  06_算定の根拠／00_媒体の構成.txt
"""
import sys
sys.dont_write_bytecode = True
import hashlib
import os
import shutil

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from nohin_data import N, MIKAN, MOUSHIOKURI, FOLDER_SETSUMEI

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.abspath(__file__)))))
REPO = os.path.join(ROOT, "repository") if os.path.isdir(
    os.path.join(ROOT, "repository")) else ROOT
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEST = os.path.join(REPO, "output", "_納品媒体")

# 媒体での置き場所。nohin_data の3列目がこのいずれかであること
FOLDERS = ["01_計画書", "02_調査報告書", "03_委員会資料", "04_図表データ",
           "05_策定の記録", "06_算定の根拠"]


def src_path(rel):
    """nohin_data のパス（output/… または data/…）を実体に直す"""
    pre, _, rest = rel.partition("/")
    if pre == "output":
        return os.path.join(REPO, "output", rest)
    if pre == "data":
        return os.path.join(BASE, "data", rest)
    return os.path.join(REPO, rel)


def plan():
    """(置き場所, 元のパス, 媒体でのパス, 大きさ) の並びと、問題の一覧を返す"""
    rows, bad = [], []
    noroku = {x[0] for x in N if x[1] == "収録しない"}
    for rel, kubun, loc, _why in N:
        if kubun == "収録しない":
            continue
        if loc not in FOLDERS:
            bad.append(f"{rel}：置き場所「{loc}」が媒体の構成にない")
            continue
        p = src_path(rel)
        if not os.path.exists(p):
            bad.append(f"{rel}：実体がない")
            continue
        nm = os.path.basename(rel.rstrip("/"))
        size = (sum(os.path.getsize(os.path.join(dp, f))
                    for dp, _, fs in os.walk(p) for f in fs)
                if os.path.isdir(p) else os.path.getsize(p))
        rows.append((loc, rel, os.path.join(loc, nm), size))
    # 収録しないと定めたものが紛れていないこと
    for loc, rel, _dst, _s in rows:
        if rel in noroku:
            bad.append(f"{rel}：収録しないと定めたものが入っている")
    return rows, bad


def _oriakeru(t, n):
    """長い文を n 文字ずつに折る（等幅で読まれることを前提にしている）"""
    return [t[i:i + n] for i in range(0, len(t), n)] or [""]


def kousei_txt(rows):
    """媒体に入れる「00_媒体の構成.txt」の中身"""
    out = ["第10期北塩原村高齢者福祉計画・第10期北塩原村介護保険事業計画",
           "電子媒体の構成", "", "【この媒体に入っているもの】", ""]
    for loc in FOLDERS:
        sel = [r for r in rows if r[0] == loc]
        if not sel:
            continue
        out.append(f"■ {loc}")
        setsu = FOLDER_SETSUMEI.get(loc)
        if setsu:
            out.append(f"    {setsu[0]}")
            for line in _oriakeru(setsu[1], 62):
                out.append(f"    {line}")
            out.append("")
        for _l, rel, dst, size in sorted(sel, key=lambda x: x[2]):
            out.append(f"      {os.path.basename(dst)}（{size // 1024:,} KB）")
        out.append("")
    out += ["", "【この媒体に入れていないもの】", ""]
    for m in MOUSHIOKURI:
        out.append("・" + m)
        out.append("")
    if MIKAN:
        out += ["", "【納品の時点で作るもの】", ""]
        for nm, _k, loc, why in MIKAN:
            out.append(f"・{nm}（{loc}）　{why}")
    return "\n".join(out) + "\n"


def main():
    rows, bad = plan()
    print("■ 納品する電子媒体の構成")
    for loc in FOLDERS:
        sel = [r for r in rows if r[0] == loc]
        tot = sum(r[3] for r in sel)
        print(f"\n  {loc}　{len(sel)}件　{tot // 1024:,} KB")
        for _l, rel, dst, size in sorted(sel, key=lambda x: x[2]):
            print(f"    {os.path.basename(dst):<52} {size // 1024:>7,} KB"
                  f"　←　{rel}")
    print(f"\n  合計 {len(rows)}件　{sum(r[3] for r in rows) // 1024 // 1024:,} MB")
    n_nai = sum(1 for x in N if x[1] == "収録しない")
    print(f"  収録しないと定めたもの {n_nai}件は入れていない")
    if bad:
        print("\n■ 組み立てられない")
        for b in bad:
            print("  ✕", b)
        return 1
    if "--write" not in sys.argv:
        print("\n  （一覧のみ。組み立てるには --write を付ける）")
        return 0
    if os.path.isdir(DEST):
        shutil.rmtree(DEST)
    for loc in FOLDERS:
        os.makedirs(os.path.join(DEST, loc), exist_ok=True)
    for _l, rel, dst, _s in rows:
        p, q = src_path(rel), os.path.join(DEST, dst)
        if os.path.isdir(p):
            shutil.copytree(p, q)
        else:
            shutil.copy2(p, q)
    with open(os.path.join(DEST, "00_媒体の構成.txt"), "w", encoding="utf-8") as f:
        f.write(kousei_txt(rows))
    # 組み立てたものに、収録しないと定めたファイルが紛れていないことを確かめる
    noroku_nm = {os.path.basename(x[0].rstrip("/")) for x in N
                 if x[1] == "収録しない"}
    magire = [os.path.join(dp, fn) for dp, _, fs in os.walk(DEST) for fn in fs
              if fn in noroku_nm]
    if magire:
        print("\n■ 収録しないと定めたものが入った（組み立てを取り消す）")
        for m in magire:
            print("  ✕", m)
        shutil.rmtree(DEST)
        return 1
    n_file = sum(len(fs) for _, _, fs in os.walk(DEST))
    print(f"\n  組み立てた: {DEST}（{n_file}ファイル）")
    print("  00_媒体の構成.txt に、入れていないものと申し送る事項を書いている")
    return 0


if __name__ == "__main__":
    sys.exit(main())
