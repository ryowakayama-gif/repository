"""小野町 引継ぎ資料の全ファイルマニフェストを生成する。

受領した `ono_deliverables_20260730_full.zip` に同梱の
`00_引継ぎ/manifest.csv` は旧18ファイル構成（00_source〜04_process_reference）
のままで、現物57ファイルのうち39ファイルが記載されていない。

本スクリプトは、展開済みの `小野町_引継ぎ_整理済` 配下の全ファイルについて
sha256・サイズ・更新日を採取し、受領時点の完全なマニフェストを出力する。
以後、資料を追加・差替えした際に再実行して差分を確認する。
"""

import csv
import hashlib
import pathlib

ROOT = pathlib.Path(__file__).parent / "小野町_引継ぎ_整理済"
OUT = ROOT / "00_引継ぎ" / "manifest_全ファイル_20260731.csv"

# 旧manifest.csvのarcname接頭辞と、現物フォルダの対応。
# 旧manifestに記載のある18ファイルを特定するために使う。
LEGACY_DIR_MAP = {
    "00_source": "01_仕様書・公告",
    "01_kickoff": "02_キックオフ・業務計画",
    "02_plan_drafts": "03_計画素案",
    "03_calculation_logic": "04_算定・見込量",
    "04_process_reference": "05_参考過程",
}


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def load_legacy():
    """旧manifest.csvを読み、現物パス -> (sha256, size) の辞書にする。"""
    legacy = {}
    src = ROOT / "00_引継ぎ" / "manifest.csv"
    if not src.exists():
        return legacy
    with open(src, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            old_dir, name = row["arcname"].split("/", 1)
            new_dir = LEGACY_DIR_MAP.get(old_dir)
            if new_dir is None:
                continue
            legacy[f"{new_dir}/{name}"] = (row["sha256"], int(row["size_bytes"]))
    return legacy


def main():
    legacy = load_legacy()
    rows = []
    for path in sorted(p for p in ROOT.rglob("*") if p.is_file()):
        rel = path.relative_to(ROOT).as_posix()
        digest = sha256(path)
        size = path.stat().st_size
        if rel in legacy:
            old_digest, old_size = legacy[rel]
            state = "旧manifest一致" if (digest, size) == (old_digest, old_size) else "旧manifest不一致"
        else:
            state = "旧manifest未記載"
        rows.append(
            {
                "path": rel,
                "category": rel.split("/", 1)[0],
                "size_bytes": size,
                "sha256": digest,
                "legacy_manifest": state,
            }
        )

    with open(OUT, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(
            f, fieldnames=["path", "category", "size_bytes", "sha256", "legacy_manifest"]
        )
        writer.writeheader()
        writer.writerows(rows)

    counts = {}
    for r in rows:
        counts[r["legacy_manifest"]] = counts.get(r["legacy_manifest"], 0) + 1
    print(f"出力: {OUT}")
    print(f"総ファイル数: {len(rows)}")
    for state, n in sorted(counts.items()):
        print(f"  {state}: {n}")


if __name__ == "__main__":
    main()
