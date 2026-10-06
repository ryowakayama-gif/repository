# -*- coding: utf-8 -*-
"""確認事項の影響度の採点（日次の表と事前打合せの資料の双方が読む）.

令和8年10月1日のご指示により日次の表（`build_nikkan.py`）に置いた採点を、
令和8年10月6日に本モジュールへ切り出した。
**同じ採点を2か所に書かない**（CLAUDE.md §4）。

採点の規則は `RULE`・`TOME_MAX`・`BAND` の3つに全て書く。
読む側はここを引くため、規則を変えれば読む側の表も変わる。

台帳・据え置き・既定値はいずれも実物又はソースから読む。
**固定値を書かない**（CLAUDE.md §4）。

  CHECK   確認事項（業務工程管理表。正の台帳）
  LACK    資料提供依頼
  KITEI   決着しない場合の当方の扱い（既定値。`data_kitei.all_kitei()`）
  KUBUN   据え置きの出どころ（Z／A／B／C）
  SUEOKI  据え置き（見込量算定）
  HANEI   既定値のとおり計画素案へ反映し終えた確認事項

使い方

    import kakunin_score as KS
    sc = KS.make_score(KS.yyyymm(基準日))     # 期限の採点に基準日を要する
    pt, uchi = sc(x)
"""

import ast
import io
import os
import re
import runpy
import sys

import repo_paths as RP

sys.path.insert(0, RP.ROOT)
import data_kitei as DK                                      # noqa: E402


# ============================================================ 読み取り
def literal(fn, name):
    """ソースから literal の代入を読む（固定値を書き写さないため）。"""
    src = io.open(os.path.join(RP.ROOT, fn), encoding="utf-8").read()
    for node in ast.parse(src).body:
        if isinstance(node, ast.Assign) and any(
                getattr(t, "id", None) == name for t in node.targets):
            return ast.literal_eval(node.value)
    raise RuntimeError("%s に %s が見つからない" % (fn, name))


def run_script(fn):
    """スクリプトを読み込んで名前空間を得る（標準出力は捨てる）。"""
    d = open(os.devnull, "w")
    o, e = sys.stdout, sys.stderr
    sys.stdout = sys.stderr = d
    try:
        return runpy.run_path(os.path.join(RP.ROOT, fn))
    finally:
        sys.stdout, sys.stderr = o, e
        d.close()


def wareki(d):
    """西暦の date を和暦の文字列にする（令和のみ）。"""
    return "令和%d年%d月%d日" % (d.year - 2018, d.month, d.day)


def yyyymm(d):
    """date を令和の通算月にする。"""
    return (d.year - 2018) * 12 + d.month


# 確認事項（正の台帳）。(0)No (1)業務内容 (2)工程 (3)表題 (4)内容
#                       (5)止めている成果物 (6)確認先 (7)状態 (8)期限 (9)回答
CHECK = literal("build_process_control.py", "CHECK")
# 資料提供依頼。(0)No (1)領域 (2)資料 (3)内容 (4)確定する主張 (5)入手先
#               (6)希望時期 (7)優先度 (8)状態 (9)備考
LACK = literal("build_process_control.py", "LACK")
# 決着しない場合の当方の扱い（既定値）。1か所から引く。
KITEI = DK.all_kitei()
# 据え置きの出どころ（A＝受託者で確定できる／B＝発注者・3町／C＝国）
KUBUN = literal("build_mikomi_juryo_nashi.py", "KUBUN")
# 既定値のとおり計画素案へ反映し終えた確認事項（ご決定はなお待っているもの）。
HANEI = {x[0] for x in literal("build_ikenkokankai.py", "SUSUMETA")}

SANTEI = run_script("build_mikomiryo_santei.py")
# 据え置き。(0)区分 (1)項目 (2)理由 (3)当方の扱い (4)関係する確認事項 (5)効き
SUEOKI = SANTEI["SUEOKI"]

KANRYO = ("完了", "了承済", "了承済（保管せず廃棄）", "代替により解消", "解決")
MACHI = [x for x in CHECK if x[7] not in KANRYO]
LACK_MACHI = [x for x in LACK if x[8] not in ("受領済", "完了", "解消")]


# ============================================================ 期限を読む
def kigen_m(s):
    """「R8.10」のような期限を令和の通算月にする。読めなければ None。"""
    m = re.search(r"R(\d+)\s*[.．]\s*(\d+)", str(s or ""))
    return (int(m.group(1)) * 12 + int(m.group(2))) if m else None


# ============================================================ 月額への効き
# 据え置きの「効き」の欄から、月額を何円動かし得るかを読む。
# 確認事項No. は据え置きの「関係する確認事項」の欄から拾う。
_YEN = re.compile(r"([0-9,]+)\s*円")
_NO = re.compile(r"No\.\s*([0-9]+)")
# 「月額」に続く範囲だけを月額とみる。
# 効きの欄には月額でない額も現れる（「差5,854,697円で月額約▲7円」のように、
# 同じ文に総額と月額が並ぶ）。文中の円を無差別に拾うと総額を月額と読む。
TSUKI_WIN = 24          # 「月額」の後ろ何文字までを月額の記載とみるか
GETSU_MAX = 2000        # 1件の月額の効きとしてあり得る上限（点検に用いる）


def max_yen(text):
    """「月額」に続いて現れる円のうち最大のものを返す。無ければ None。"""
    t = text or ""
    v = []
    for m in re.finditer("月額", t):
        seg = t[m.end():m.end() + TSUKI_WIN]
        v += [int(x.replace(",", "")) for x in _YEN.findall(seg)]
    return max(v) if v else None


def no_of(text):
    """「No.112・No.150」から確認事項No.を拾う。"""
    return [int(n) for n in _NO.findall(str(text or ""))]


GETSUGAKU = {}          # 確認事項No. -> (円, 据え置きの項目)
for _s in SUEOKI:
    _y = max_yen(_s[5])
    if _y is None:
        continue
    for _no in no_of(_s[4]):
        if _y > GETSUGAKU.get(_no, (0, ""))[0]:
            GETSUGAKU[_no] = (_y, _s[1])


# ============================================================ 影響度の採点
# 期限は当区域ではほとんどが超過又は当月であり、それだけでは順位が付かない。
# 順位を分けるのは**放置したときに何が動くか**であるため、
# 月額への効きと「止めている成果物」を重く、期限を軽くしている。
RULE = [
    ("期限", "回答期限が基準日の月より前（超過している）", 2),
    ("期限", "回答期限が基準日と同じ月", 1),
    ("期限", "回答期限が翌月", 0),
    ("止めている成果物", "概算（第1次・第2次）を止めている", 3),
    ("止めている成果物", "保険料に関わるものを止めている", 2),
    ("止めている成果物", "計画素案を止めている", 1),
    ("月額への効き", "据え置きに紐づき月額100円以上を動かし得る", 4),
    ("月額への効き", "同 50円以上100円未満", 2),
    ("月額への効き", "同 50円未満（効きの記載がある）", 1),
    ("法定記載事項", "介護保険法第117条の記載事項に直結する", 2),
    ("既定値", "決着しない場合の当方の扱いを置けていない", 3),
    ("当方で動かせるか", "国の告示・公布を待つもの（催促できない）", -2),
]
TOME_MAX = 4            # 「止めている成果物」の加点の上限
BAND = [("至急", 9, 99, "翌日に着手する。相手がある場合は翌日の午前に出す"),
        ("優先", 6, 8, "今週のうちに着手する"),
        ("通常", -99, 5, "期限の月に入ってから着手する")]

# 国の告示・公布を待つ据え置き（区分C）に紐づく確認事項No.
KUNI_NO = set()
for _i, _s in enumerate(SUEOKI, start=1):
    if KUBUN.get(_i) == "C":
        KUNI_NO.update(no_of(_s[4]))

_HOTEI = ("法第117条", "法定記載事項", "第117条")


def oki(text):
    """算定の表の言い回しを、送付する資料の文に直す。

    算定の側の表で用いている言い回し（自己点検の番号、シートの番号、
    「本表」）は写した先では意味を持たないため落とす。
    確定した日を述べる文も、注で述べるためはじめに置かない。
    **写すときに落とす処理を1か所に置く**（CLAUDE.md §4）。
    """
    v = str(text).replace("**", "")
    if v.startswith("令和8年10月1日") and "。" in v:
        v = v.split("。", 1)[1]
    v = re.sub(r"（点検[0-9０-９・点検]*）", "", v)
    v = re.sub(r"[0-9０-９]+シートに", "サービス見込量の算定に", v)
    v = v.replace("本表は", "見込量は")
    return v.strip()


def make_score(now_m):
    """基準日（令和の通算月）を与えて採点の関数を作る。"""

    def score(x):
        """確認事項1件の影響度を採点し、(合計, [内訳]) を返す。"""
        uchi = []
        km = kigen_m(x[8])
        if km is not None:
            if km < now_m:
                uchi.append(("期限", "超過（%s）" % x[8], 2))
            elif km == now_m:
                uchi.append(("期限", "当月（%s）" % x[8], 1))
        tome, t = 0, str(x[5])
        if "概算" in t:
            tome += 3
            uchi.append(("止めている成果物", "概算", 3))
        if "保険料" in t:
            tome += 2
            uchi.append(("止めている成果物", "保険料", 2))
        if "計画素案" in t:
            tome += 1
            uchi.append(("止めている成果物", "計画素案", 1))
        if tome > TOME_MAX:          # 上限で頭打ちにする
            uchi = [u for u in uchi if u[0] != "止めている成果物"]
            uchi.append(("止めている成果物", "概算・保険料・計画素案（上限）",
                         TOME_MAX))
            tome = TOME_MAX
        g = GETSUGAKU.get(x[0])
        if g:
            p = 4 if g[0] >= 100 else (2 if g[0] >= 50 else 1)
            uchi.append(("月額への効き", "最大%s円（%s）"
                         % ("{:,}".format(g[0]), g[1]), p))
        if any(w in str(x[4]) for w in _HOTEI):
            uchi.append(("法定記載事項", "法第117条に直結", 2))
        if x[0] not in KITEI and "決着しない場合" not in str(x[9]):
            uchi.append(("既定値", "置けていない", 3))
        if x[0] in KUNI_NO:
            uchi.append(("当方で動かせるか", "国の告示・公布待ち", -2))
        return sum(u[2] for u in uchi), uchi

    return score


def band_of(pt):
    """点から区分（至急・優先・通常）を返す。"""
    for nm, lo, hi, _d in BAND:
        if lo <= pt <= hi:
            return nm
    return BAND[-1][0]


def sorted_check(now_m, rows=None):
    """(点, 内訳, 行) を点の高い順に並べて返す。"""
    sc = make_score(now_m)
    out = []
    for x in (rows if rows is not None else MACHI):
        pt, uchi = sc(x)
        out.append((pt, uchi, x))
    out.sort(key=lambda s: (-s[0], kigen_m(s[2][8]) or 9999, s[2][0]))
    return out
