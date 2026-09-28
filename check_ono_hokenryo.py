"""成果品に古い保険料額が残っていないかを機械的に点検する。

令和8年9月28日、保険料算式の誤り（調整交付金見込額の基数に総合事業費を
含めていなかった）を訂正した。当方の第10期の保険料基準額は
月額6,095円から6,048円に変わる（100円未満切上げの6,100円は不変）。

本スクリプトは `小野町_引継ぎ_整理済` 配下の docx・xlsx を走査し、
訂正前の額が本文に残っていないかを確かめる。
受領した原本（【受領】・原本_ で始まるもの）は町・システムの出力なので除く。

  python3 check_ono_hokenryo.py
"""

import pathlib
import re
import sys
import warnings
import zipfile

warnings.filterwarnings("ignore")

ROOT = pathlib.Path(__file__).parent / "小野町_引継ぎ_整理済"

# 訂正前の額。これらが成果品に残っていたら要修正。
# 6,300円は単価の趨勢ケース（標準ケースとは別の置き方）の切上げ額としても
# 現れるため、ここには入れない。標準ケースの切上げ額は POSITIVE で確かめる。
STALE = {
    "6,095円": "当方の第10期保険料基準額（訂正前）。正しくは6,048円",
    "6,096円": "同上",
    "6,210円": "さらに前の版の額",
    "6,211円": "所得段階別の割合を訂正する前の額",
    "6,449円": "従前の方式によるケース（訂正前）。正しくは6,402円",
    "6,442円": "第9期の実数による検算（訂正前）。正しくは6,865円",
}

# 標準ケースの額を載せているはずの成果品。載っていなければ作り直し漏れ。
POSITIVE = {
    "小野町高齢者保健福祉計画_第10期介護保険事業計画_素案_第5版": ("6,048円", "6,100円"),
    "小野町_第10期_サービス見込量の算定結果": ("6,048円",),
    "小野町_見える化_所得段階別被保険者数の入力": ("6,048円",),
}

# 訂正の経緯を説明する文書では、訂正前の額に言及してよい。
ALLOW = ("推計が通りました", "システムとの突合", "業務進捗", "進捗管理")


def is_gen(p):
    """当方が生成した成果品か（受領した原本でないか）。"""
    if any(part.startswith(("【受領】", "原本_")) for part in p.parts):
        return False
    return p.suffix in (".docx", ".xlsx")


def text_of(p):
    """docx・xlsx の共有文字列と本文を雑に取り出す。"""
    out = []
    try:
        with zipfile.ZipFile(p) as z:
            for n in z.namelist():
                if not n.endswith(".xml"):
                    continue
                if not (n.startswith("word/") or n.startswith("xl/")):
                    continue
                try:
                    out.append(z.read(n).decode("utf-8", "ignore"))
                except Exception:
                    pass
    except Exception:
        return ""
    # XMLタグをまたいで分割された文字列も拾えるよう、タグを落としてから繋ぐ
    return re.sub(r"<[^>]+>", "", "\n".join(out))


def main():
    hits, missing, seen = [], [], set()
    n = 0
    for p in sorted(ROOT.rglob("*")):
        if not p.is_file() or not is_gen(p):
            continue
        n += 1
        t = None
        for key, musts in POSITIVE.items():
            if key in p.name:
                seen.add(key)
                t = text_of(p) if t is None else t
                for must in musts:
                    if must not in t:
                        missing.append((p.relative_to(ROOT), must))
        if any(a in p.name for a in ALLOW):
            continue
        t = text_of(p) if t is None else t
        for bad, why in STALE.items():
            if bad in t:
                hits.append((p.relative_to(ROOT), bad, why))
    for key in POSITIVE:
        if key not in seen:
            missing.append((key, "そもそも成果品が見つからない"))
    print(f"走査 {n}件（うち経緯説明文書は古い額の点検から除外）")
    if not hits and not missing:
        print("古い保険料額は残っていない。標準ケースの額も載っている。")
        return 0
    if hits:
        print(f"**古い額が {len(hits)}件 残っている**")
        for rel, bad, why in hits:
            print(f"  {rel}\n      {bad}  … {why}")
    if missing:
        print(f"**載っているべき額が {len(missing)}件 欠けている**")
        for rel, must in missing:
            print(f"  {rel}\n      {must} が見当たらない")
    return 1


if __name__ == "__main__":
    sys.exit(main())
