# -*- coding: utf-8 -*-
"""公表用素案を組む（WBS Ⅱ-119・Ⅱ-120）

【何をするか】
  kouhyou_map.py の振り分け（村のご承認 令和8年10月20日）にしたがって、
  協議用素案の節を第9期計画の目次体系に並べ替え、番号を付け替える。

    ① 第9期の目次（K9）＋追加（TSUIKA）の順に項目を並べる
    ② 1つの項目に複数の節を当てるときは、節の題を小見出し（h3）にして束ねる
       もとの小見出し（h3）は太字の段落に降ろす
    ③ 本文の相互参照（「5-4(1)」「施策4-6」「資10」）を新しい番号に付け替える
    ④ 協議用のみに置く節への参照は落とすか、手当てが要るものとして挙げる

【番号の付け方】
  章の項　　　「1-1」「2-4」「5-2」の形（章-項）
  第4章の施策　「施策1-1」の形（第9期と同じ）
  第4章の法定の項目　「第4章5」「第4章6」の形（施策と紛れないように）
  資料編　　　「資料編○1」の形

  python3 scripts/kouhyou_build.py           組み立ててJSONを書く
  python3 scripts/kouhyou_build.py --refs    付け替え表と、手当てが要る参照を示す
"""
import collections
import json
import os
import re
import sys
sys.dont_write_bytecode = True
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import paths as P
import soan_content as SC
import kouhyou_map as KM
from figures_map import FIGS

# ══════════════════════════════════════════════════════════════
# 第4章の施策の並び（第9期の番号 → 協議用素案の施策）。MAP から起こす
def _new_no(ch, no):
    """章と番号から、公表用の参照の形を作る"""
    if ch == "資料編":
        return "資料編" + no
    if ch == "第4章":
        return ("施策" + no) if re.match(r"^\d+-\d+$", no) else ("第4章" + no)
    return ch.replace("第", "").replace("章", "") + "-" + no


def shisaku_map():
    """協議用素案の施策 → 公表用の参照の形。
       第5章3に移す施策3-5のように、第4章の外へ移るものも拾う。"""
    out = {}
    for (ch, no), (ss, _m) in KM.MAP.items():
        if no.startswith("基本目標"):
            continue
        for s in ss:
            if s.startswith("施策"):
                out[s] = _new_no(ch, no)
    for ch, no, _nm, ss, _e, _k in KM.TSUIKA:
        for s in ss:
            if s.startswith("施策"):
                out[s] = _new_no(ch, no)
    return out


def setsu_map():
    """協議用素案の節（施策を除く） → 公表用の参照の形"""
    out = {}
    for (ch, no), (ss, _m) in KM.MAP.items():
        if no.startswith("基本目標"):
            continue
        for s in ss:
            if not s.startswith("施策"):
                out[s] = _new_no(ch, no)
    for ch, no, _nm, ss, _e, _k in KM.TSUIKA:
        for s in ss:
            if not s.startswith("施策"):
                out[s] = _new_no(ch, no)
    return out


SHISAKU = shisaku_map()
SETSU = setsu_map()
# 協議用のみに置く節。参照は落とすか手当てが要る
KYOGI = set(KM.KYOGI)
# 施策の題（「3-5（介護給付の適正化）」のように題を添えて書かれた参照を
# 節ではなく施策と見分けるために使う）
SHISAKU_NA = {}
for _ch in SC.CH:
    for _s in _ch["sections"]:
        if _s["no"].startswith("施策"):
            SHISAKU_NA[_s["no"][2:]] = _s["title"].split("　")[0].split("【")[0]


def ref_table():
    """付け替え表。(種類, 旧, 新) の一覧"""
    rows = []
    for k in sorted(SETSU, key=lambda x: (len(x), x)):
        rows.append(("節", k, SETSU[k]))
    for k in sorted(SHISAKU, key=lambda x: (len(x), x)):
        rows.append(("施策", k, SHISAKU[k] or "―"))
    for k in sorted(KYOGI):
        rows.append(("協議用のみ", k, "―"))
    return rows


# ══════════════════════════════════════════════════════════════
# 相互参照の付け替え
# 枝（例 5-5(2)）まで拾う。協議用5-5を公表用で2つの節に分けたため、
# 枝ごとに行き先が違う。枝を拾わないと「5-2(2)」のような存在しない参照になる。
TOK = re.compile(r"(施策\d+-\d+|資\d+|(?<![0-9０-９\-－第])\d-\d{1,2}(?:\(\d\))?(?![0-9]))")
# 丸ごと落とせる括弧（中が協議用のみの節だけでできているもの）
KAKKO = re.compile(r"（((?:5-12|2-8|2-9|2-11|3-5|資[1-4]|資7|資8)(?:・[^（）]{0,8})*)）")

# 協議用のみに置く節のうち、中身が公表用の別の節に織り込まれるもの。
# 参照はその節へ振り替える（kouhyou_map.KYOGI に書いた理由のとおり）。
KYOGI_REF = {
 "2-8": "5-4",    # 保険者機能の評価 → 交付金の活用
 "2-9": "6-1",    # 第9期計画の進捗と評価 → PDCA
 "2-11": "3-3",   # 本村の課題 → 計画の基本目標
 "3-3": "3-3",    # 第10期の重点 → 計画の基本目標
 "3-5": "3-4",    # 施策体系の対応表 → 計画の体系及びSDGs
}

# 5-12（現時点で据え置いた項目）は公表用に対応する節がない。
# 参照している文は、5-12を指さない言い方に書き替える。
# （協議用素案はそのままにする。公表用を組むときだけ当てる）
REWRITE = [
 ("・5-12）", "）"),
 ("（5-12(1)）", ""),

 ("5-12(2)の国の政令・告示等が示され", "国の政令・告示等が示され"),
 ("該当する項目は5-12に掲げています。", ""),
 ("5-12の確定の条件に従って再算定します", "確定の条件に従って再算定します"),
 ("該当する項目は5-12に一覧化しています。", ""),
 ("5-12に一覧化した31項目のうち", "確定を待っている31項目のうち"),
 ("5-12(1)の据え置いた項目と重なります", "確定を待っている項目と重なります"),
 # 題を添えずに書いた施策への参照。節の番号と紛れるため、施策と分かる形にする
 ("（2-8・4-4）", "（2-8・施策4-4）"),
 ("32点／100点（4-4）", "32点／100点（施策4-4）"),
]

# 基本目標5（協議用の4-5）に置いた施策。公表用では第6章2に移る
SHISAKU_SOTO = {"施策5-1": "6-2", "施策5-2": "6-2", "施策5-3": "6-2", "施策5-4": "6-2"}

# ══════════════════════════════════════════════════════════════
# 公表用の実測（令和8年10月20日・組んで PDF に変換して数えた本文の頁数）
#   素案（編集注記を含む）と、計画書（注記を落としたもの）を分けて持つ。
#   村にお示しするのは計画書の頁数である。前付4頁を加えて合計になる。
#
#   kouhyou_map.kouhyou_pages() の推定は本文108頁だったが、実測は117頁だった。
#   推定は節の高さを足しただけで、束ねるときに足す小見出し・基本目標の見出し・
#   章の変わり目の改頁を見ていなかった。**推定は当てにしない。**
#
#   測り直すには scripts/jissoku_kouhyou.py を走らせる（組んで変換して数える）。
#   令和8年10月21日に測り直した（doc110）。例規集の全分野と突き合わせて
#   施策1-2・施策4-6・4-6(2)・6-3・6-4を直した段階では5通りとも前と同じだったが、
#   さらに施策4-6へキャリアアップ補助金の別表（補助対象5事業と限度額）を加えると
#   計画書は4通りとも1頁増えた（素案127頁は変わらず）。
#   同じ加筆で協議用の計画書は7通りとも動いていない。改頁の位置の動きは
#   組み方ごとに違うため、**一方の実測から他方を推すことはできない。**
JISSOKU_KOUHYOU = {
    "":            127,   # 素案（注記を含む）は本文127頁＝131頁
}
JISSOKU_KEIKAKU = {
    "":            118,   # 案を当てない計画書　＝122頁（許容115頁を7頁超える）
    "ABCDE":       110,   # ＝114頁
    "ABCDEHJ":     107,   # ＝111頁
    "ABCDEGHJKL":  104,   # ＝108頁（目安100頁に対し+8頁・許容に7頁の余裕）
}
MAE_PG = 4

NOKORI = []          # 手当てが要る参照 (節, 旧の語, 前後)


def _is_shisaku_gloss(t, end):
    """「3-5（介護給付の適正化）」のように題を添えた参照は施策を指す"""
    m = re.match(r"（([^（）]{2,20})）", t[end:])
    if not m:
        return False
    g = m.group(1)
    # 「村外施設の利用支援」と「村外施設サービスの利用支援」のように
    # 題を縮めて書いてある参照も拾う。4文字以上が続けて一致すれば同じものと見る
    for v in SHISAKU_NA.values():
        if g in v or v in g:
            return True
        for i in range(len(g) - 3):
            if g[i:i + 4] in v:
                return True
    return False


def fix_refs(t, sec):
    """1つの文字列の相互参照を付け替える"""
    if not t:
        return t
    # ⓪ 5-12 を指す文を書き替える
    for a, b in REWRITE:
        t = t.replace(a, b)
    # ① 中が協議用のみの節だけの括弧は丸ごと落とす
    def _k(m):
        inner = m.group(1)
        toks = re.split(r"[・，,]", inner)
        if all(x.strip() in KYOGI for x in toks if x.strip()):
            NOKORI.append((sec, m.group(0), "括弧ごと落とした"))
            return ""
        return m.group(0)
    t = KAKKO.sub(_k, t)

    # ② 残りの語を付け替える
    def _r(m):
        tok = m.group(1)
        if tok.startswith("施策"):
            return SHISAKU.get(tok, SHISAKU_SOTO.get(tok, tok))
        # 題を添えた参照は施策を指す
        if not tok.startswith("資") and _is_shisaku_gloss(t, m.end()):
            return SHISAKU.get("施策" + tok, tok)
        if tok in SETSU:                      # 枝つきはここで当たる
            return SETSU[tok]
        me = EDA.match(tok)
        if me:                                # 枝つきで行き先が無いもの
            moto = me.group(1)
            if moto in KYOGI:
                if moto in KYOGI_REF:
                    return KYOGI_REF[moto] + "(%s)" % me.group(2)
                a = max(0, m.start() - 18)
                NOKORI.append((sec, tok, t[a:m.end() + 16].replace("\n", " ")))
                return tok
            # 節まるごとが1つの節に移るなら、枝はそのまま付けて移す
            return SETSU.get(moto, moto) + "(%s)" % me.group(2)
        if tok in KYOGI:
            if tok in KYOGI_REF:
                return KYOGI_REF[tok]
            a = max(0, m.start() - 18)
            NOKORI.append((sec, tok, t[a:m.end() + 16].replace("\n", " ")))
            return tok
        if tok not in SETSU:
            # 行き先を知らない節番号は、素通しすると公表用の別の節を指す。
            # 協議用5-10が公表用5-5になるため、裸の「5-5」を素通しすると
            # 第5章5（低所得者への対応）を指してしまう。黙って通さない。
            a = max(0, m.start() - 18)
            NOKORI.append((sec, tok, t[a:m.end() + 16].replace("\n", " ")))
            return tok
        return SETSU[tok]
    return TOK.sub(_r, t)


def walk_fix(o, sec):
    if isinstance(o, str):
        return fix_refs(o, sec)
    if isinstance(o, dict):
        return {k: walk_fix(v, sec) for k, v in o.items()}
    if isinstance(o, list):
        return [walk_fix(v, sec) for v in o]
    if isinstance(o, tuple):
        return tuple(walk_fix(v, sec) for v in o)
    return o


# ══════════════════════════════════════════════════════════════
def _sec_index():
    """節の索引。削減案を当てた状態で呼べるよう、その都度数え直す"""
    return {s["no"]: (ch["no"], s) for ch in SC.CH for s in ch["sections"]}


# 節の枝を指す書き方（例 "5-5(2)"）。節まるごとではなく小見出しの範囲を取る
EDA = re.compile(r"^(\d-\d{1,2})\((\d)\)$")


def _src(SEC, s0):
    """s0 が指す中身を (章, 題, ブロック) で返す。

    "5-5" のように節番号だけなら節まるごと。
    "5-5(2)" のように枝を添えたら、その小見出し「(2) …」から
    次の「(…)」の小見出しまでを取る。題は小見出しから「(2) 」を外したもの。
    枝の中に入れ子の小見出し（5-5(1)の「量の見込みの立て方」）があっても、
    「(…)」で始まらないものは枝の切れ目としない。

    協議用5-5は量の見込みと事業費を1つの節に持つ。第9期は事業費を
    第5章2（給付費等の今後の見込み）に置いていたため、公表用では
    (1) を第4章4-3へ、(2) を第5章2へ分ける（村のご判断による）。
    """
    m = EDA.match(s0)
    if not m:
        ch, sec = SEC[s0]
        return ch, sec["title"], sec["blocks"]
    no, eda = m.group(1), m.group(2)
    if no not in SEC:
        raise SystemExit("節%sがない（%sを解けない）" % (no, s0))
    ch, sec = SEC[no]
    atama = "(%s)" % eda
    st = None
    for i, b in enumerate(sec["blocks"]):
        if b.get("t") != "h3":
            continue
        v = str(b.get("v", ""))
        if st is None:
            if v.startswith(atama):
                st = i
        elif re.match(r"^\(\d\)", v):
            return ch, v0_title(sec["blocks"][st]), sec["blocks"][st + 1:i]
    if st is None:
        raise SystemExit("節%sに小見出し「%s」がない" % (no, atama))
    return ch, v0_title(sec["blocks"][st]), sec["blocks"][st + 1:]


def _eda_moto(s0):
    """枝つきの書き方から節番号だけを取る（図の対応表の鍵は節番号で持つ）"""
    m = EDA.match(s0)
    return m.group(1) if m else s0


def v0_title(b):
    """小見出し「(2) 地域支援事業費の見込み」→「地域支援事業費の見込み」"""
    return re.sub(r"^\(\d\)\s*", "", str(b.get("v", "")))

# 第1章3「計画の対象」は協議用素案にない。第9期の記述により新設する
SHINSETSU = {
 ("第1章", "3"): [
   {"t": "p", "v": "本計画の対象は、村内にお住まいの65歳以上の方と、"
                   "その方を介護するご家族です。あわせて、要介護・要支援の認定を"
                   "受けている40歳以上65歳未満の方（第2号被保険者）と、"
                   "高齢者の暮らしを支える地域の皆様、事業者、関係機関を対象とします。"},
   {"t": "note", "v": "本節は第9期計画 第1章3（計画の対象）の記述によります。"
                      "協議用素案に対応する節がないため新設しました。"
                      "対象の書き方を村にご確認いただきます。"},
 ],
 ("第2章", "3"): [
   {"t": "p", "v": "村が独自に実施している高齢者福祉サービスの実施状況は、"
                   "第4章5（高齢者福祉サービスの量の目標）の表に実績として示しています。"
                   "同じ表を二度掲げないため、本節では再掲しません。"},
 ],
}


def h3_to_bold(b):
    """束ねるときに、もとの小見出しを太字の段落に降ろす"""
    if b["t"] == "h3":
        return {"t": "p", "v": "■ " + b["v"], "bold": True}
    return b


def build():
    """公表用の章の木を作る"""
    SEC = _sec_index()
    chapters = []
    figs_new = {}
    for ch, nm, items in KM.K9:
        secs = []
        # 第4章は基本目標の見出しを挟む
        cur_moku = None
        rows = [(no, inm, KM.MAP[(ch, no)][0], False) for no, inm, _pg in items]
        rows += [(t[1], t[2], t[3], True) for t in KM.TSUIKA if t[0] == ch]
        for no, inm, ss, tsuika in rows:
            if ch == "第4章" and re.match(r"^\d+-\d+$", no):
                m = no.split("-")[0]
                if m != cur_moku:
                    cur_moku = m
                    mk = KM.MAP.get(("第4章", "基本目標" + m))
                    blocks = []
                    if mk:
                        for s0 in mk[0]:
                            blocks += [walk_fix(b, s0) for b in SEC[s0][1]["blocks"]]
                    secs.append({"no": "基本目標" + m,
                                 "title": KM.K9_MOKUHYO[m], "blocks": blocks})
            blocks = []
            if (ch, no) in SHINSETSU:
                blocks = list(SHINSETSU[(ch, no)])
            elif len(ss) == 1:
                sch, _sti, sbl = _src(SEC, ss[0])
                blocks = [walk_fix(b, ss[0]) for b in sbl]
                for e in FIGS:
                    if e[0] == "%s|%s" % (sch, _eda_moto(ss[0])):
                        figs_new.setdefault("%s|%s" % (ch, no), []).append(e)
            else:
                for s0 in ss:
                    sch, sti, sbl = _src(SEC, s0)
                    ti = sti.split("　")[0].split("【")[0]
                    blocks.append({"t": "h3", "v": ti})
                    blocks += [walk_fix(h3_to_bold(b), s0) for b in sbl]
                    for e in FIGS:
                        if e[0] == "%s|%s" % (sch, _eda_moto(s0)):
                            figs_new.setdefault("%s|%s" % (ch, no), []).append(e)
            secs.append({"no": no, "title": inm, "blocks": blocks,
                         "tsuika": tsuika})
        chapters.append({"no": ch, "title": nm, "sections": secs})
    return chapters, figs_new


def to_json(out_path=None, an=""):
    """an に削減案の記号（例 "ABCDEGHJKL"）を渡すと、その案を当てた形で組む"""
    if an:
        import estimate_sakugen as ES
        tg, yg, fg = ES.an_targets(an)
        with ES.nuki(tg, yg, fg):
            chapters, figs_new = build()
    else:
        chapters, figs_new = build()
    figs = {}
    figs_inline = {}
    for k, es in figs_new.items():
        for _sec, fn, cap, src, inline in es:
            rec = {"file": fn, "caption": cap, "source": src}
            if inline:
                figs_inline[fn] = rec
            else:
                figs.setdefault(k, []).append(rec)
    for _sec, fn, cap, src, inline in FIGS:
        if inline:
            figs_inline.setdefault(fn, {"file": fn, "caption": cap, "source": src})
    out = {"title": SC.TITLE, "title2": SC.TITLE2, "subtitle": SC.SUBTITLE,
           "draft": "素　案（公表用）", "issuer": SC.ISSUER, "date": SC.DATE,
           "chapters": chapters, "figures": figs, "figures_inline": figs_inline}
    fp = out_path or P.build("kouhyou_soan%s.json" % ("_" + an if an else ""))
    with open(fp, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    return fp, chapters


def main():
    if "--refs" in sys.argv:
        print("■ 相互参照の付け替え表\n")
        print("  %-8s %-7s %s" % ("種類", "協議用", "公表用"))
        for t, a, b in ref_table():
            print("  %-8s %-7s %s" % (t, a, b))
        build()
        print("\n■ 手当てが要る参照　%d件\n" % len([x for x in NOKORI if x[2] != "括弧ごと落とした"]))
        c = collections.Counter(x[1] for x in NOKORI if x[2] != "括弧ごと落とした")
        for k, v in c.most_common():
            ex = next(x[2] for x in NOKORI if x[1] == k and x[2] != "括弧ごと落とした")
            print("  %-5s %3d件　例：%s" % (k, v, ex[:70]))
        print("\n  括弧ごと落とした参照 %d件"
              % len([x for x in NOKORI if x[2] == "括弧ごと落とした"]))
        return 0
    an = ""
    if "--an" in sys.argv:
        i = sys.argv.index("--an")
        an = sys.argv[i + 1] if i + 1 < len(sys.argv) else ""
    fp, chapters = to_json(an=an)
    n = sum(len(c["sections"]) for c in chapters)
    if an:
        print("■ 削減案 %s を当てた\n" % an)
    print("■ 公表用素案を組んだ\n")
    for c in chapters:
        print("  %-6s %-24s %2d項目" % (c["no"], c["title"][:24], len(c["sections"])))
    print("\n  計 %d項目　書き出した： %s" % (n, fp))
    print("\n■ 紙面（実測。組んで PDF に変換して数えた）\n")
    print("  %-14s %6s %6s %6s %8s" % ("行う案", "本文", "前付", "合計", "許容115頁"))
    print("  " + "─" * 50)
    print("  %-14s %6d %6d %6d %8s"
          % ("素案（注記あり）", JISSOKU_KOUHYOU[""], MAE_PG,
             JISSOKU_KOUHYOU[""] + MAE_PG, "―"))
    for k, v in JISSOKU_KEIKAKU.items():
        g = v + MAE_PG
        print("  %-14s %6d %6d %6d %+8d"
              % (("計画書（なし）" if not k else "　＋" + k), v, MAE_PG, g, g - 115))
    print("\n  ※ 村にお示しするのは計画書の頁数である。"
          "案を1つも行わないと許容115頁を6頁超える。")
    print("  ※ 案F・M・Nは、対象の節（2-8・資8・資7）が協議用のみのため公表用では効かない。")
    nok = len([x for x in NOKORI if x[2] != "括弧ごと落とした"])
    print("  相互参照：括弧ごと落とした %d件／手当てが要る %d件（--refs で一覧）"
          % (len(NOKORI) - nok, nok))
    return 0


if __name__ == "__main__":
    sys.exit(main())
