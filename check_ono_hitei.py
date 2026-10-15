"""否定試験 ― 点検そのものが働いているかを確かめる。

点検が通っていることと、点検が働いていることは別である。
点検スクリプトが4本（体裁・紙面・一次資料・接点）と、生成器の自己点検が
十数本に増えたため、**わざと誤った値を入れて止まるかどうか**を確かめる。

なぜ要るか
  点検は、通っているあいだは何も言わない。条件の書き間違いで常に通る
  状態になっていても、表示は「すべて適合」のままである。
  **「すべて適合」という表示そのものを信用するために、
  誤りを入れたら止まることを示す。**

やり方
  1. すべて**その場（メモリ上）で差し替える**。リポジトリのファイルは
     書き換えない。終わったら必ず元に戻す（finally）。
  2. 1件につき**1つだけ**誤りを入れる。複数入れると、どの規則が
     働いたのか分からない。
  3. 「止まるはずなのに止まらなかった」ものだけを並べる。
     止まったものは、その規則が働いている証拠である。

出力
  画面のみ（成果品ではない）。11_品質管理に記録を残すのは別途。
"""

import contextlib
import copy
import pathlib
import sys
import warnings

warnings.filterwarnings("ignore")

ROOT = pathlib.Path(__file__).parent
sys.path.insert(0, str(ROOT))

SCRATCH = pathlib.Path("/tmp/claude-0/-home-user-repository/"
                       "d7dca1d7-b918-515d-8b6d-9309b5e7ef32/scratchpad/hitei")

RESULT = []     # (対象, 入れた誤り, 止まったか, 捕まえた規則)


def shiken(taisho, ayamari, fn):
    """fn() が「止まった理由」を返せば合格。空なら不合格。"""
    try:
        riyu = fn()
    except AssertionError as e:                 # assert で止まるのも合格
        riyu = "assert: %s" % (e or "（条件不成立）")
    except Exception as e:                      # 例外で落ちるのは不合格
        RESULT.append((taisho, ayamari, "例外", "%s: %s"
                       % (type(e).__name__, str(e)[:60])))
        return
    RESULT.append((taisho, ayamari, "止まった" if riyu else "**素通り**",
                   str(riyu)[:90] if riyu else "―"))


@contextlib.contextmanager
def sashikae(obj, name, value):
    """属性をその場で差し替え、必ず元に戻す。"""
    old = getattr(obj, name)
    setattr(obj, name, value)
    try:
        yield
    finally:
        setattr(obj, name, old)


# ---------------------------------------------------------------- 1 一次資料
def t_ichiji():
    """地域支援事業費の値を1円ずらすと、一次資料との突合が止まるか。"""
    import build_ono_tanka as T
    import check_ono_ichiji as I

    hon = copy.deepcopy(T.CHIIKI_JISSEKI)
    nokori = list(I.RESULT)
    try:
        T.CHIIKI_JISSEKI["令和6年度"]["一般介護予防事業"] += 1
        I.RESULT.clear()
        I.check_keiri(I.nenpo_books())
        ng = [x for x in I.RESULT if x[4] == "NG"]
        return "NG %d件（%s）" % (len(ng), ng[0][1]) if ng else ""
    finally:
        T.CHIIKI_JISSEKI.clear()
        T.CHIIKI_JISSEKI.update(hon)
        I.RESULT.clear()
        I.RESULT.extend(nokori)


def t_kofukin():
    """令和7年度の値をずらすと、交付金実績報告との突合が止まるか。"""
    import build_ono_tanka as T
    import check_ono_ichiji as I

    hon = copy.deepcopy(T.CHIIKI_JISSEKI)
    nokori = list(I.RESULT)
    try:
        T.CHIIKI_JISSEKI["令和7年度"]["包括的支援事業・任意事業"] -= 1000
        I.RESULT.clear()
        I.check_chiiki_r7()
        ng = [x for x in I.RESULT if x[4] == "NG"]
        return "NG %d件（%s）" % (len(ng), ng[0][1]) if ng else ""
    finally:
        T.CHIIKI_JISSEKI.clear()
        T.CHIIKI_JISSEKI.update(hon)
        I.RESULT.clear()
        I.RESULT.extend(nokori)


# ---------------------------------------------------------------- 2 体裁
def _kowasu(kind):
    """わざと体裁を崩した docx を作って、その道を返す。"""
    import ono_style as S
    SCRATCH.mkdir(parents=True, exist_ok=True)
    r = S.Report("否定試験")
    r.h1("見出し")
    r.p("本文です。")
    r.tbl([["あ", "い"], ["1", "2"]], widths=[8.6, 8.6], caption="表")
    p = SCRATCH / ("kowasu_%s.docx" % kind)
    r.save(p)
    from docx import Document
    from docx.shared import Cm, Pt
    d = Document(str(p))
    if kind == "font":
        d.styles["Normal"].font.name = "ＭＳ 明朝"
    elif kind == "size":
        d.styles["Normal"].font.size = Pt(9)
    elif kind == "margin":
        d.sections[0].left_margin = Cm(3.0)
    elif kind == "width":
        # 列幅は gridCol と各セルの tcW の両方に書かれている。
        # 点検は行のセル幅を見るため、セル側を変える。
        for row in d.tables[0].rows:
            for c in row.cells:
                c.width = Cm(3.0)
    d.save(str(p))
    return p


def t_taisai(kind, nazenara):
    def f():
        import check_ono_taisai as C
        bad = C.check(_kowasu(kind))[0]
        return "；".join(bad[:2]) if bad else ""
    return f


# ---------------------------------------------------------------- 3 接点
def t_setten():
    """接点の文言を1字変えると、両計画の照合が止まるか。"""
    import ono_setten as T

    hon = list(T.SETTEN)
    try:
        n, k, s, t = T.SETTEN[0]
        T.SETTEN[0] = (n, k + "（否定試験）", s, t)
        with open("/dev/null", "w") as dev:       # 画面に出さない
            old, sys.stdout = sys.stdout, dev
            try:
                rc = T.check()
            finally:
                sys.stdout = old
        return "check() が 1（NGあり）を返した" if rc else ""
    finally:
        T.SETTEN[:] = hon


# ---------------------------------------------------------------- 4 法定記載
def t_hotei():
    """素案から語が1つ欠けたことにすると、✕と判定するか。"""
    import build_ono_hotei_check as H

    mark, _hit = H.judge("老人福祉事業の量の目標を定めます。", ["老人福祉事業の量の目標"])
    nashi, _h2 = H.judge("この文には目標がありません。", ["老人福祉事業の量の目標"])
    if mark != "○":
        return "そろっているのに○にならない（%s）" % mark
    return "" if nashi == "○" else "欠けを %s と判定した" % nashi


# ---------------------------------------------------------------- 5 自己点検
def t_kakunin_bango():
    """確認事項の通し番号を重複させると、自己点検が止まるか。"""
    import build_ono_kakunin_ichiran as K

    hon = list(K.KAKUNIN)
    try:
        x = list(K.KAKUNIN[5])
        x[0] = K.KAKUNIN[4][0]        # 1つ前と同じ番号にする
        K.KAKUNIN[5] = tuple(x)
        ng = K.jiko()
        return "；".join(ng[:2]) if ng else ""
    finally:
        K.KAKUNIN[:] = hon


def t_kakunin_atesaki():
    """解決済み（Ａ）なのに宛先を入れると、自己点検が止まるか。"""
    import build_ono_kakunin_ichiran as K

    hon = list(K.KAKUNIN)
    try:
        x = list(K.KAKUNIN[0])
        x[3] = "町"                    # Ａは宛先が「―」でなければならない
        K.KAKUNIN[0] = tuple(x)
        ng = K.jiko()
        return "；".join(ng[:2]) if ng else ""
    finally:
        K.KAKUNIN[:] = hon


def t_shogai_setsu():
    """障がい素案の節と施策の数をずらすと、自己点検が止まるか。"""
    import build_ono_soan_shogai as G

    hon = G.SHISAKU["第2章"]
    try:
        G.SHISAKU["第2章"] = (hon[0], hon[1][:-1])     # 1節ぶん減らす
        ng = G.jiko(None)
        return "；".join(ng[:2]) if ng else ""
    finally:
        G.SHISAKU["第2章"] = hon


def t_shogai_hotei():
    """障がい素案の法定記載事項から新設の号を落とすと、止まるか。"""
    import build_ono_soan_shogai as G

    hon = list(G.HOTEI)
    try:
        G.HOTEI[:] = [x for x in G.HOTEI if "第88条第3項第4号" not in x[0]]
        ng = G.jiko(None)
        return "；".join(ng[:2]) if ng else ""
    finally:
        G.HOTEI[:] = hon


def t_houkaisei():
    """法改正の仕分けで、事項を「〜関係」に丸めると止まるか。"""
    import build_ono_houkaisei as H

    hon = list(H.K51)
    try:
        x = list(H.K51[0])
        x[0] = "介護保険法関係"          # 条・項・号で指していない
        H.K51[0] = tuple(x)
        import ono_style as S
        ng = H.jiko(S.Report("否定試験"))
        return "；".join(ng[:2]) if ng else ""
    finally:
        H.K51[:] = hon


# ================================================================
SHIKEN = [
    ("一次資料との突合（年報 様式4）", "令和6年度の一般介護予防事業費を1円増やす",
     t_ichiji),
    ("一次資料との突合（交付金実績報告）",
     "令和7年度の包括的支援事業・任意事業を1,000円減らす", t_kofukin),
    ("体裁の点検", "本文の書体をＭＳ明朝にする", t_taisai("font", "書体")),
    ("体裁の点検", "本文のサイズを9ptにする", t_taisai("size", "サイズ")),
    ("体裁の点検", "左余白を3.0cmにする", t_taisai("margin", "余白")),
    ("体裁の点検", "表の列幅の合計を6.0cmにする", t_taisai("width", "列幅")),
    ("両計画の接点の照合", "接点1の見出しに5字加える", t_setten),
    ("法定記載事項チェック", "判定に用いる語が本文にない状態を作る", t_hotei),
    ("確認事項の一覧　自己点検", "通し番号を1つ重複させる", t_kakunin_bango),
    ("確認事項の一覧　自己点検", "解決済み（Ａ）に宛先を入れる", t_kakunin_atesaki),
    ("障がい素案　自己点検", "第2章の施策を1節ぶん減らす", t_shogai_setsu),
    ("障がい素案　自己点検", "法定記載事項から第88条第3項第4号を落とす",
     t_shogai_hotei),
    ("法改正の仕分け　自己点検", "事項を「介護保険法関係」に丸める", t_houkaisei),
]


def main():
    for taisho, ayamari, fn in SHIKEN:
        shiken(taisho, ayamari, fn)

    print("■ 否定試験　わざと誤りを入れて、点検が止まるかを確かめる")
    print()
    last = None
    for taisho, ayamari, kekka, riyu in RESULT:
        if taisho != last:
            print("── %s" % taisho)
            last = taisho
        print("   %-8s %s" % (kekka, ayamari))
        if kekka == "止まった":
            print("            └ %s" % riyu)
        elif kekka != "止まった":
            print("            └ **%s**" % riyu)
    print()
    sudoori = [x for x in RESULT if x[2] != "止まった"]
    print("   試験 %d件　止まった %d件／**素通り・例外 %d件**"
          % (len(RESULT), len(RESULT) - len(sudoori), len(sudoori)))
    if not sudoori:
        print("   **すべての規則が働いています。**")
    return sudoori


if __name__ == "__main__":
    sys.exit(1 if main() else 0)
