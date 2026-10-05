# -*- coding: utf-8 -*-
"""仕様書（spec_data.S）と WBS（wbs_data.W）と確認事項（wbs_kakunin.K）の整合を機械的に点検する。

   不適合があれば内容を出力して終了コード 1 を返す。

   点検1 WBS の ID が一意であること
   点検2 仕様書の要求事項が参照する WBS ID が実在すること
   点検3 すべての要求事項に対応作業があること（＝仕様書からの漏れがない）
   点検4 仕様書に根拠を持つ WBS 作業が、いずれかの要求事項から参照されていること（＝宙に浮いた作業がない）
   点検5 WBS の「仕様書該当」列と、参照元の要求事項の条項が一致すること
   点検6 仕様書外の作業（手引き・確認）がすべて SHIEN で支援先を示されていること
   点検7 確認事項の「関連WBS No.」が実在する ID を指していること
   点検8 WBS の表示順が 大分類→中分類 の順に整っていること
   点検9 未解決の優先度S・Aのすべてに影響度が与えられていること（計画策定側は IMPACT、調査工程側は機械）
   点検10 影響度が LEVEL に定義された値であること、かつ IMPACT に調査工程側の項目が混ざっていないこと
   点検11 影響度の判定に挙げた「止まる対象」のWBS番号が実在すること
   点検12 照会の束がすべてのWBSを覆い、束の間で重複がないこと
   点検13 翌営業日の作業（READY）が実在する未完了のWBSを指し、影響度1（停止）の項目で止まっていないこと
   点検14 未完了のWBSが READY か HOLD_REASON のいずれかで扱われていること
   点検15 進捗の実績日の整合（完了に実績完了日があること、着手に実績開始日があること、
          未着手に実績日がないこと、実績完了日が実績開始日より後であること）
   点検16 納品物の収録範囲が、実際に作られているすべてのファイルを覆っていること
          （新しいファイルを置いたら区分を書くまで通らない）
   点検17 納品する電子媒体が組み立てられること（収録すると定めたものが実在し、
          置き場所が媒体の構成にあり、収録しないものが紛れないこと）
   点検18 成果品を組み直すのに要るものがすべて版管理にあり、置き場所をじか書き
          していないこと（作業環境が作り直されても組み直せること）
"""
import os, re, sys
sys.dont_write_bytecode = True   # 古いバイトコードで誤った結果が出ることを防ぐ
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from spec_data import S, SHIEN, GAIBU
from wbs_data import W
from wbs_kakunin import K, SOLVED
from wbs_progress import P
from wbs_pending import LEVEL, IMPACT, BUNDLE, READY, HOLD_REASON
from nohin_data import N as NOHIN, MIKAN, MOUSHIOKURI

NG = []
def chk(no, name, ok, detail=""):
    print(f"  {'○' if ok else '✕'} 点検{no} {name}" + (f"　… {detail}" if detail else ""))
    if not ok:
        NG.append(f"点検{no} {name}：{detail}")

ids = [w[0] for w in W]
IDX = {w[0]: w for w in W}

# 1
dup = [i for i in set(ids) if ids.count(i) > 1]
chk(1, "WBS IDの一意性", not dup, f"重複 {dup}" if dup else f"{len(ids)}件すべて一意")

# 2
miss = sorted({i for _, _, _, refs, _ in S for i in refs if i not in IDX})
chk(2, "要求事項が参照するWBS IDの実在", not miss, f"不明 {miss}" if miss else f"参照{sum(len(r[3]) for r in S)}件")

# 3
nolink = [(c, t) for c, t, tan, refs, _ in S if not refs and tan != "村"]
chk(3, "仕様書の要求事項の充足", not nolink,
    "／".join(f"{c} {t}" for c, t in nolink) if nolink else f"要求事項{len(S)}件すべてに対応作業あり")

# 4
referred = {i for _, _, _, refs, _ in S for i in refs}
orphan = [(w[0], w[3]) for w in W if w[4].startswith("仕様書") or w[4].startswith("通知")
          if w[0] not in referred]
chk(4, "仕様書に根拠をもつ作業の被参照", not orphan,
    "／".join(f"{a} {b}" for a, b in orphan) if orphan else "")

# 5 条項の一致（「仕様書該当」列の条項から、その作業を参照している要求事項があるか）
bad5 = []
for w in W:
    if not (w[4].startswith("仕様書") or w[4].startswith("通知")):
        continue
    src_clauses = [re.sub(r"-[a-z]$", "", c) for c, _, _, refs, _ in S if w[0] in refs]
    if not any(c == w[4] or c.startswith(w[4]) for c in src_clauses):
        bad5.append(f"{w[0]}（列={w[4]} / 参照元={'・'.join(sorted(set(src_clauses))) or 'なし'}）")
chk(5, "「仕様書該当」列と参照元条項の一致", not bad5, "／".join(bad5) if bad5 else "")

# 6
gaibu = [w[0] for w in W if w[4] in GAIBU]
nosh = [i for i in gaibu if i not in SHIEN]
chk(6, "仕様書外の作業の位置づけの明示", not nosh,
    f"支援先未記載 {nosh}" if nosh else f"仕様書外{len(gaibu)}件すべてに支援先を明示")

# 7
refs_k, bad7 = 0, set()
for row in list(K) + list(SOLVED):
    for r in re.split(r"[,、]\s*", row[3]):
        r = r.strip()
        if not r or r == "―":
            continue
        refs_k += 1
        if r not in IDX:
            bad7.add(r)
chk(7, "確認事項の関連WBS No.の実在", not bad7,
    f"不明 {sorted(bad7)}" if bad7 else f"参照{refs_k}件すべて実在")

# 8
MAJ = ["０ 契約・立上げ", "Ⅰ ニーズ調査", "Ⅱ 計画策定", "Ⅲ 打合せ等", "Ⅳ 成果品・納品", "Ⅴ 管理"]
seq, seen, bad8 = [], set(), []
for w in W:
    key = (w[1], w[2])
    if not seq or seq[-1] != key:
        if key in seen:
            bad8.append(f"{w[1]}／{w[2]}")
        seq.append(key)
        seen.add(key)
if [m for m in (k[0] for k in seq)] != sorted({k[0] for k in seq}, key=MAJ.index) * 0 + \
        [m for m in dict.fromkeys(k[0] for k in seq)] or True:
    pass
order_ok = [MAJ.index(m) for m in dict.fromkeys(k[0] for k in seq)] == \
           sorted([MAJ.index(m) for m in dict.fromkeys(k[0] for k in seq)])
chk(8, "WBSの表示順", not bad8 and order_ok,
    f"分断 {bad8}" if bad8 else ("大分類の順序が不正" if not order_ok else f"{len(seq)}区分が整列"))

# ══════════ ペンディング整理の点検（点検9〜14）══════════
def _ids(k):
    return [x.strip() for x in re.split(r"[,、]\s*", k[3])
            if x.strip() and x.strip() != "―"]

PROG = {}
for w in W:
    e = P.get(w[3])
    PROG[w[0]] = (e[1] if e else 0.0, e[0] if e else "未着手", w[3])

SA = [(i, k) for i, k in enumerate(K, 1) if k[5] in ("S", "A")]
# IMPACT は確認事項の**見出し**で引く（並びの位置で引くと足し引きでずれる）
MIDASHI = {i: k[1] for i, k in SA}
SURVEY = {k[1] for i, k in SA if _ids(k) and all(x.startswith("Ⅰ") for x in _ids(k))}
PLAN = {k[1] for i, k in SA} - SURVEY

# 9
miss9 = sorted(PLAN - set(IMPACT))
amari9 = sorted(set(IMPACT) - PLAN)       # 見出しが変わると残る
chk(9, "未解決の優先度S・Aへの影響度の付与", not miss9 and not amari9,
    ((f"影響度なし {[m[:22] for m in miss9[:3]]}" if miss9 else "")
     + (f"／見出しが合わない {[m[:22] for m in amari9[:3]]}" if amari9 else ""))
    if (miss9 or amari9)
    else f"計画策定側{len(PLAN)}件に判定／調査工程側{len(SURVEY)}件は機械で影響度5")

# 10
bad10 = [m[:22] for m, v in IMPACT.items() if v[0] not in LEVEL]
bad10 += [m[:22] + "（調査工程）" for m in IMPACT if m in SURVEY]
chk(10, "影響度の値と対象の正しさ", not bad10, "／".join(bad10) if bad10
    else f"影響度0〜4の{len(IMPACT)}件すべてが定義済みの値")

# 11
bad11 = set()
for nm11, v in IMPACT.items():
    for m in re.findall(r"[０-９Ⅰ-Ⅴ]+-\d+", v[1]):
        if m not in IDX:
            bad11.add(f"{nm11[:18]}:{m}")
chk(11, "「止まる対象」のWBS番号の実在", not bad11,
    f"不明 {sorted(bad11)}" if bad11 else "影響度の判定54件の参照はすべて実在")

# 12
bw, dup = [], []
seen12 = {}
for no, nm, wids, _ in BUNDLE:
    for x in wids:
        if x not in IDX:
            bw.append(f"{no}:{x}")
        if x in seen12:
            dup.append(f"{x}（{seen12[x]}と{no}）")
        seen12[x] = no
uncovered = [w[0] for w in W if w[0] not in seen12]
chk(12, "照会の束の網羅と非重複", not bw and not dup and not uncovered,
    "／".join(bw + dup + [f"束に属さない未完了 {uncovered}" if uncovered else ""]).strip("／")
    if (bw or dup or uncovered) else f"{len(BUNDLE)}束でWBS {len(W)}件をすべて覆い重複なし")

# 13
bad13 = []
STOP = {x for _m, v in IMPACT.items() if v[0] == 1
        for x in re.findall(r"[０-９Ⅰ-Ⅴ]+-\d+", v[1])}
for no, wid, doing, judge, hrs in READY:
    if wid not in IDX:
        bad13.append(f"順位{no}:{wid} が実在しない")
        continue
    if PROG[wid][0] >= 1.0:
        bad13.append(f"順位{no}:{wid} は完了済み")
    if wid in STOP:
        bad13.append(f"順位{no}:{wid} は影響度1（停止）の項目で止まっている")
nos = [r[0] for r in READY]
if nos != list(range(1, len(nos) + 1)):
    bad13.append(f"順位が1からの連番でない {nos}")
chk(13, "翌営業日の作業の妥当性", not bad13, "／".join(bad13) if bad13
    else f"{len(READY)}件すべてが未完了・ペンディングなし")

# 14
ready_ids = {r[1] for r in READY}
bad14 = [w[0] for w in W
         if PROG[w[0]][0] < 1.0 and PROG[w[0]][1] != "対象外"
         and not w[0].startswith("Ⅰ")
         and w[0] not in ready_ids and w[0] not in HOLD_REASON]
over14 = [x for x in HOLD_REASON if x not in IDX or PROG.get(x, (0,))[0] >= 1.0
          or x in ready_ids]
chk(14, "未完了のWBSの扱いの明示", not bad14 and not over14,
    "／".join([f"扱い未記載 {bad14}" if bad14 else "", f"不要な理由 {over14}" if over14 else ""]).strip("／")
    if (bad14 or over14) else f"READY {len(ready_ids)}件＋理由 {len(HOLD_REASON)}件で網羅")


# 15
bad15 = []
for nm_, e in P.items():
    st_, pr_, s0, s1, _note = e
    if st_ == "完了" and not s1:
        bad15.append(f"{nm_}：完了に実績完了日がない")
    if st_ == "着手" and not s0:
        bad15.append(f"{nm_}：着手に実績開始日がない")
    if st_ == "未着手" and (s0 or s1):
        bad15.append(f"{nm_}：未着手に実績日がある")
    if s1 and not s0:
        bad15.append(f"{nm_}：実績完了日だけがある")
    if s0 and s1 and s1 < s0:
        bad15.append(f"{nm_}：実績完了日が実績開始日より前")
    if st_ == "完了" and pr_ != 1.00:
        bad15.append(f"{nm_}：完了なのに進捗が{pr_}")
chk(15, "進捗の実績日の整合", not bad15,
    "／".join(bad15) if bad15 else f"進捗反映{len(P)}件すべて整合")


# 16
import os as _os16
# scripts/ → 北塩原村_第10期/ → docs/ → repository/
BASE16 = _os16.path.dirname(_os16.path.dirname(_os16.path.abspath(__file__)))
ROOT16 = _os16.path.dirname(_os16.path.dirname(BASE16))
GEN16 = {"output": _os16.path.join(ROOT16, "output"),
         "data": _os16.path.join(BASE16, "data")}
aru16, bad16 = set(), []
for pre, d in GEN16.items():
    if not _os16.path.isdir(d):
        bad16.append(f"{d} がない")
        continue
    for nm in _os16.listdir(d):
        # 先頭が . や _ のものは組み立ての置き場であり、元のファイルではない
        # （output/_納品媒体 は build_nohin.py が写しを並べる場所）
        if nm.startswith(".") or nm.startswith("_"):
            continue
        aru16.add(f"{pre}/{nm}")
kiji16 = {x[0] for x in NOHIN}
mibunrui = sorted(aru16 - kiji16)
amari16 = sorted(kiji16 - aru16)
if mibunrui:
    bad16.append("区分が書かれていない " + "・".join(mibunrui[:4])
                 + (f" ほか{len(mibunrui) - 4}件" if len(mibunrui) > 4 else ""))
if amari16:
    bad16.append("実在しないのに区分がある " + "・".join(amari16[:4]))
KUBUN16 = {"成果品", "関連", "収録しない"}
warui = [x[0] for x in NOHIN if x[1] not in KUBUN16]
if warui:
    bad16.append("区分の名が違う " + "・".join(warui[:3]))
naiyou = [x[0] for x in NOHIN if not x[3].strip()]
if naiyou:
    bad16.append("理由が書かれていない " + "・".join(naiyou[:3]))
n_syu = sum(1 for x in NOHIN if x[1] == "成果品")
n_kan = sum(1 for x in NOHIN if x[1] == "関連")
n_nai = sum(1 for x in NOHIN if x[1] == "収録しない")
chk(16, "納品物の収録範囲の網羅", not bad16,
    "／".join(bad16[:2]) if bad16
    else f"実在{len(aru16)}件すべてに区分あり（成果品{n_syu}／関連{n_kan}／"
         f"収録しない{n_nai}）＋納品時に作るもの{len(MIKAN)}件")


# 17
try:
    import build_nohin as _BN
    rows17, bad17 = _BN.plan()
    n17 = {}
    for loc, _r, _d, _s in rows17:
        n17[loc] = n17.get(loc, 0) + 1
    # 媒体の各フォルダに、何のためのものかの説明があること
    from nohin_data import FOLDER_SETSUMEI as _FS17
    for loc17 in _BN.FOLDERS:
        if any(r[0] == loc17 for r in rows17) and loc17 not in _FS17:
            bad17.append(f"{loc17} に説明がない（村が使い道を判断できない）")
    chk(17, "納品する電子媒体が組み立てられること", not bad17,
        "／".join(bad17[:3]) if bad17
        else "／".join(f"{k} {v}件" for k, v in sorted(n17.items()))
             + f"（計{len(rows17)}件）")
except Exception as e:
    chk(17, "納品する電子媒体が組み立てられること", False, f"照合できない（{e}）")


# 18
#   金ケ崎町の案件では、本文が版管理の外（スクラッチパッド）にあり、
#   作業環境が作り直された時点で消えた。同じことが起きないかを機械で見る。
#     (a) 本文・体裁・データの各モジュールが版管理にあること
#     (b) 置き場所をじか書きしていないこと（別の場所に復元しても組み直せる）
#     (c) 外部のライブラリの版が宣言されていること
import subprocess as _sp18
try:
    bad18 = []
    SCR18 = _os16.path.dirname(_os16.path.abspath(__file__))
    tracked = set(_sp18.check_output(
        ["git", "-C", BASE16, "ls-files"], text=True).split("\n"))
    # (a) 組立てに要るモジュールが追跡されていること
    IRU = ["soan_content.py", "shiryo_content.py", "shiryo3_content.py",
           "figures_map.py", "data_zuhyo.py", "wbs_data.py", "wbs_progress.py",
           "wbs_kakunin.py", "wbs_pending.py", "spec_data.py", "nohin_data.py",
           "shukei_data.py", "shihyo_dict.py", "paths.py", "docxlib.js",
           "build_soan.py", "build_soan_docx.js", "build_shiryo_docx.js",
           "build_figures.py", "build_wbs.py", "build_zuhyo_daicho.py"]
    for nm in IRU:
        if ("scripts/" + nm) not in tracked:
            bad18.append(f"{nm} が版管理にない")
    # (b) 置き場所のじか書き。
    #     この点検自身が探す語を素のまま持つと自分を拾ってしまうため、組み立てる。
    JIKA = "/home/user/" + "repository"
    TMPLIB = "/tmp/node_" + "modules"
    import glob as _g18
    jika = []
    for f18 in sorted(_g18.glob(_os16.path.join(SCR18, "*.py"))
                      + _g18.glob(_os16.path.join(SCR18, "*.js"))):
        nm = _os16.path.basename(f18)
        if nm in ("paths.py", "docxlib.js"):      # ここだけが置き場所を知ってよい
            continue
        with open(f18, encoding="utf-8") as fh:
            src = fh.read()
        if JIKA in src:
            jika.append(nm)
        if TMPLIB in src:
            jika.append(nm + "（/tmp のライブラリ）")
    if jika:
        bad18.append("置き場所のじか書き " + "・".join(sorted(set(jika))[:4]))
    # (c) 外部のライブラリの版の宣言
    pkg = _os16.path.join(BASE16, "package.json")
    if not _os16.path.exists(pkg):
        bad18.append("package.json がなく docx の版が宣言されていない")
    elif not _sp18.check_output(
            ["git", "-C", BASE16, "ls-files", pkg], text=True).strip():
        bad18.append("package.json が版管理にない")
    chk(18, "成果品を組み直せること（版管理と置き場所）", not bad18,
        "／".join(bad18[:3]) if bad18
        else f"組立てに要る{len(IRU)}件が版管理にあり、置き場所のじか書きなし、"
             f"docx の版は package.json で固定")
except Exception as e:
    chk(18, "成果品を組み直せること（版管理と置き場所）", False, f"照合できない（{e}）")


# ── 19　成果品に書き方の記号が残っていないこと ───────────────────
#    備考や説明の文は手元の記録から持ってくるため、マークダウンの強調（**…**）が
#    そのまま表に出る。00_媒体の構成.txt で現に起きた。
#    村が開く成果品（xlsx・txt）を走査して、残っていないことを確かめる。
try:
    import glob as _gl19
    import io as _io19
    import paths as _P19
    import openpyxl as _xl19
    OUT19 = _P19.OUT
    bad19, n19 = [], 0
    for f in sorted(_gl19.glob(_os16.path.join(OUT19, "*.xlsx"))):
        if _os16.path.basename(f).startswith("~$"):
            continue
        wb = _xl19.load_workbook(f, read_only=True, data_only=True)
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                for v in row:
                    if isinstance(v, str):
                        n19 += 1
                        if "**" in v:
                            bad19.append("%s %s：%s"
                                         % (_os16.path.basename(f), ws.title, v[:30]))
        wb.close()
    for f in sorted(_gl19.glob(_os16.path.join(OUT19, "_納品媒体", "**", "*.txt"),
                               recursive=True)):
        t = _io19.open(f, encoding="utf-8").read()
        n19 += 1
        if "**" in t:
            bad19.append("%s：強調記号" % _os16.path.basename(f))
    # docx も見る。本文モジュールに ** を書けば村の目に触れる。
    # 連が割れても拾えるよう、段落ごとに w:t をつないでから見る。
    import re as _re19
    import zipfile as _zip19
    for f in sorted(_gl19.glob(_os16.path.join(OUT19, "*.docx"))):
        if _os16.path.basename(f).startswith("~$"):
            continue
        z = _zip19.ZipFile(f)
        for nm in z.namelist():
            if not (nm.startswith("word/") and nm.endswith(".xml")):
                continue
            x = z.read(nm).decode("utf-8", "replace")
            for para in _re19.split(r"</w:p>", x):
                t = "".join(_re19.findall(r"<w:t[^>]*>([^<]*)</w:t>", para))
                if not t:
                    continue
                n19 += 1
                if "**" in t:
                    bad19.append("%s：%s" % (_os16.path.basename(f), t[:30]))
        z.close()
    chk(19, "成果品に書き方の記号が残っていないこと", not bad19,
        "／".join(sorted(set(bad19))[:3]) if bad19
        else "xlsx・txt・docx の文字列%d件に マークダウンの強調（**）は残っていない" % n19)
except Exception as e:
    chk(19, "成果品に書き方の記号が残っていないこと", False, "照合できない（%s）" % e)

# ── 20　照会票が組み立てられること ─────────────────────────
#    村に出す照会票は、確認事項から導いている。導き方に穴があると、
#    尋ね忘れ（どの票にも出ない）か二度尋ね（2つの票に出る）が起きる。
#    ・未解決の確認事項は、束1〜6のどれか1つにだけ現れること
#    ・ただし束7（調査工程・他メンバー担当）と影響度0（社内）は出さない
#    ・影響度1（停止）の件は第1部に置かれていること
try:
    import irai_bundle as _IB20
    bad20 = []
    doko = {}
    for nm, _, _ in _IB20.bundles():
        for lv, title, _, rows in _IB20.items(nm):
            for r in rows:
                doko.setdefault(r[5], []).append((nm, lv))   # 鍵は元の見出し
            if lv == 1 and not title.startswith("第1部"):
                bad20.append("%s：影響度1が%sに置かれている" % (nm, title.split("　")[0]))
    nido = [k for k, v in doko.items() if len(v) > 1]
    if nido:
        bad20.append("2つ以上の票に出る %d件（%s）" % (len(nido), nido[0][:24]))

    # 数の突合：未解決 ＝ 票に出る分 ＋ 束7 ＋ 影響度0
    machi20 = [k for k in K if k[6] in _IB20.MACHI]
    n_b7 = sum(1 for k in machi20 if _IB20.assign(k) == "束7")
    n_lv0 = sum(1 for k in machi20
                if _IB20.assign(k) != "束7"
                and (IMPACT.get(k[1]) or (None,))[0] == 0)
    n_hyo = len(doko)
    if n_hyo + n_b7 + n_lv0 != len(machi20):
        bad20.append("票%d＋束7の%d＋社内の%d＝%d が未解決の%d件に合わない"
                     % (n_hyo, n_b7, n_lv0, n_hyo + n_b7 + n_lv0, len(machi20)))
    nashi = [k[1] for k in machi20
             if _IB20.assign(k) != "束7"
             and (IMPACT.get(k[1]) or (None,))[0] != 0
             and k[1] not in doko]
    if nashi:
        bad20.append("どの票にも出ない %d件（%s）" % (len(nashi), nashi[0][:24]))

    kara = [r[5] for nm, _, _ in _IB20.bundles()
            for lv, _, _, rows in _IB20.items(nm) for r in rows
            if r[0] == "―" or len(r[0]) < 8]
    if kara:
        bad20.append("尋ねたいことが空になる %d件（%s）" % (len(kara), kara[0][:24]))

    chk(20, "照会票が組み立てられること", not bad20,
        "／".join(bad20[:3]) if bad20
        else "未解決%d件のうち、票に%d件（束1〜6）・束7が%d件・社内が%d件。"
             "二度尋ねなし、尋ね忘れなし、影響度1は第1部"
             % (len(machi20), n_hyo, n_b7, n_lv0))
except Exception as e:
    chk(20, "照会票が組み立てられること", False, "照合できない（%s）" % e)

print()
if NG:
    print(f"不適合 {len(NG)}件")
    for n in NG:
        print("  -", n)
    sys.exit(1)
print(f"適合：仕様書の要求事項 {len(S)}件／WBS {len(W)}件／確認事項の参照 {refs_k}件／影響度の判定 {len(IMPACT)}件／照会の束 {len(BUNDLE)}束／翌営業日の作業 {len(READY)}件"
      f"／納品物の区分 {len(NOHIN)}件")
