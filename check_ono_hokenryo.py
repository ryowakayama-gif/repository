"""成果品に古い保険料額が残っていないかを機械的に点検する。

令和8年9月28日、保険料算式の誤り（調整交付金見込額の基数に総合事業費を
含めていなかった）を訂正した。当方の第10期の保険料基準額は
月額6,095円から6,048円に変わる（100円未満切上げの6,100円は不変）。

本スクリプトは `小野町_引継ぎ_整理済` 配下の docx・xlsx を走査し、
訂正前の額が本文に残っていないかを確かめる。
受領した原本（【受領】・原本_ で始まるもの）は町・システムの出力なので除く。

  python3 check_ono_hokenryo.py
"""

import html
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
    "6,449円": "従前の方式によるケース（旧版）。正しくは6,402円",
    "6,442円": "第9期の実数による検算（旧版）。正しくは6,865円",
    # 保険料の算定は町にまだお出ししていない。算式を直した経緯を成果品に書かない。
    "おわびと訂正": "保険料算式の経緯。町にお出ししていない額の訂正は書かない",
    "算式の訂正": "同上",
    "訂正前": "同上",
    # 令和8年10月21日：町に出す成果品には、訂正・おわびの語をいっさい出さない（ご指示による）。
    "訂正": "訂正の語。直した経緯は書かず、今の内容だけを書く",
    "お詫び": "同上",
    "おわび": "同上",
}

# 標準ケースの額を載せているはずの成果品。載っていなければ作り直し漏れ。
POSITIVE = {
    "小野町高齢者保健福祉計画_第10期介護保険事業計画_素案_第16版": ("6,053円", "6,100円"),
    "小野町_第10期_サービス見込量の算定結果": ("6,048円",),
    "小野町_見える化_所得段階別被保険者数の入力": ("6,048円",),
    "小野町高齢者福祉サービス推進協議会資料_第9期計画の評価と令和7年度調査結果_20260929":
        ("6,048円", "6,533円"),
}

# 読み手に見えてはいけない強調マーカー。`**` を太字に変換せず本文に出していないか。
# 令和8年9月28日に素案（708か所）と協議会資料（196か所）で見つかり、直した。
NO_MARKER = (
    "小野町_見える化ワークシートの確認結果",
    "小野町_第10期_見える化ワークシートとの突合",
    "小野町高齢者保健福祉計画_第10期介護保険事業計画_素案_第7版",
    "小野町高齢者福祉サービス推進協議会資料_第9期計画の評価と令和7年度調査結果_20260929",
)

# 保険料の算定は町にまだお出ししていない。そのため成果品には、算式を直した
# 経緯（おわびと訂正）も、訂正前の額も書かない。除外する文書はない。
ALLOW = ()

# 古い額と同じ数字が、別の算定で正しく出る場合。直前（150字以内）にこの語があれば古い額とみなさない。
# 令和8年10月22日：認知症対応型共同生活介護1ユニット整備のケースは、標準ケース6,048円に
# 整備分を足して6,095円になる（ono_seibi.cases）。
BETSU = {
    "6,095円": ("認知症対応型共同生活介護を1ユニット",),
}


def stale_in(t, bad):
    """古い額が、BETSU の文脈以外で出ているか。"""
    for m in re.finditer(re.escape(bad), t):
        mae = t[max(0, m.start() - 150):m.start()]
        if not any(w in mae for w in BETSU.get(bad, ())):
            return True
    return False


def is_gen(p):
    """当方が生成した成果品か（受領した原本でないか）。"""
    if any(part.startswith(("【受領】", "原本_")) for part in p.parts):
        return False
    # 【社内】のフォルダは町に出さない資料（旧版・経過資料、社内の作業様式）なので除く
    if any("【社内】" in part for part in p.parts):
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
    # XMLタグをまたいで分割された文字列も拾えるよう、タグを落としてから繋ぐ。
    # xlsx は日本語を文字参照（&#20870; など）で持つことがあるので実体に戻す。
    # これを戻さないと「6,442円」のような円つきの語を取りこぼす。
    return html.unescape(re.sub(r"<[^>]+>", "", "\n".join(out)))


def main():
    hits, missing, seen, marks = [], [], set(), []
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
        if any(k in p.name for k in NO_MARKER):
            t = text_of(p) if t is None else t
            if "**" in t:
                marks.append((p.relative_to(ROOT), t.count("**")))
        if any(a in p.name for a in ALLOW):
            continue
        t = text_of(p) if t is None else t
        for bad, why in STALE.items():
            if stale_in(t, bad):
                hits.append((p.relative_to(ROOT), bad, why))
    for key in POSITIVE:
        if key not in seen:
            missing.append((key, "そもそも成果品が見つからない"))
    print(f"走査 {n}件")
    if not hits and not missing and not marks:
        print("古い保険料額は残っていない。標準ケースの額も載っている。")
        print("強調マーカー（**）が本文に出ている成果品もない。")
        return 0
    if marks:
        print(f"**強調マーカーが本文に出ている成果品が {len(marks)}件**")
        for rel, c in marks:
            print(f"  {rel}\n      ** が {c}か所。太字に変換されていない")
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
