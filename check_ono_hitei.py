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


def t_shogai_techo():
    """手帳の内訳を1人ずらすと、計と合わなくなって止まるか。"""
    import build_ono_soan_shogai as G

    hon = [(nm, list(v)) for nm, v in G.TECHO_SUII]
    try:
        nm, v = G.TECHO_SUII[0]
        v2 = list(v)
        v2[0] += 1                     # 令和元年の身体を1人増やす
        G.TECHO_SUII[0] = (nm, v2)
        ng = G.jiko(None)
        return "；".join(ng[:2]) if ng else ""
    finally:
        G.TECHO_SUII[:] = [(nm, list(v)) for nm, v in hon]



def t_zenkai_pdf():
    """前回計画の読み取りを1つずらすと、原本との照合が止まるか。"""
    import ono_shogai_zenkai as Z

    Z.load()
    hon = list(Z.SERVICE["生活介護"]["人日"]["6期"])
    try:
        Z.SERVICE["生活介護"]["人日"]["6期"][1] += 1
        ng = Z.jiko()
        return "；".join(ng[:2]) if ng else ""
    finally:
        Z.SERVICE["生活介護"]["人日"]["6期"][:] = hon



def t_zenkai_shisaku():
    """前回計画の施策の章を1つ消すと、9章そろっているかの点検が止まるか。"""
    import ono_shogai_zenkai as Z

    Z.load()
    hon = list(Z.SHISAKU)
    try:
        Z.SHISAKU[:] = [x for x in Z.SHISAKU if not x[1].startswith("第10章")]
        ng = Z.jiko()
        return "；".join(ng[:2]) if ng else ""
    finally:
        Z.SHISAKU[:] = hon



def t_kiriwake():
    """切り分けから1件落とすと、確認事項の自己点検が止まるか。"""
    import build_ono_kakunin_ichiran as K
    import build_ono_machidata_kiriwake as M

    hon = list(M.KIRIWAKE)
    try:
        M.KIRIWAKE[:] = [x for x in M.KIRIWAKE if x[0] != 32]
        ng = K.jiko()
        return "；".join(ng[:2]) if ng else ""
    finally:
        M.KIRIWAKE[:] = hon


def t_kiriwake_yokei():
    """確認事項に無い№を切り分けに入れると、止まるか。"""
    import build_ono_machidata_kiriwake as M

    hon = list(M.KIRIWAKE)
    try:
        M.KIRIWAKE.append((999, "Ⅳ", "ないもの", "ないもの"))
        ng = M.jiko()
        return "；".join(ng[:2]) if ng else ""
    finally:
        M.KIRIWAKE[:] = hon



def t_jobun():
    """条文番号をわざと1つずらすと、原典との照合が止まるか。"""
    import check_ono_jobun as J

    hon = list(J.JOBUN)
    nokori = list(J.RESULT)
    try:
        # 調和の規定を第7項から第5項に戻す（かつての誤り）
        J.JOBUN[:] = [("障害者総合支援法", J.GEN, "88", "5", None,
                       ["市町村地域福祉計画", "調和"], "否定試験")]
        J.RESULT.clear()
        ng = J.main()
        return "NG %d件（%s）" % (len(ng), ng[0][5]) if ng else ""
    finally:
        J.JOBUN[:] = hon
        J.RESULT.clear()
        J.RESULT.extend(nokori)


def t_jobun_jiten():
    """**時点を取り違えると止まるか。**

    改正後にしかない条項（介護保険法第117条第5項の勘案義務）を
    「現行」として引くと、見当たらないと判定されなければならない。
    時点を見ないつくりに戻すと、この試験が素通りする。
    """
    import check_ono_jobun as J

    hon = list(J.JOBUN)
    nokori = list(J.RESULT)
    try:
        J.JOBUN[:] = [("介護保険法", J.GEN, "117", "5", "1",
                       ["有料老人ホーム及び登録住宅のそれぞれの入居定員総数"],
                       "否定試験（改正後の条項を現行として引く）")]
        J.RESULT.clear()
        ng = J.main()
        return "NG %d件（%s）" % (len(ng), ng[0][5]) if ng else ""
    finally:
        J.JOBUN[:] = hon
        J.RESULT.clear()
        J.RESULT.extend(nokori)


def t_joritsu():
    """**保険料の標準乗率を1段階だけ動かすと、政令との照合が止まるか。**

    13段階の乗率は介護保険法施行令第38条第1項が漢数字で定めている。
    従前は見える化システムの画面で確かめていただけで、
    **政令の現物とは突き合わせていなかった。**
    """
    import build_ono_tanka as T
    import check_ono_jobun as J

    hon = list(T.SHOTOKU_R7)
    try:
        nm, hyojun, fuka, n = T.SHOTOKU_R7[5]      # 第6段階 1.200
        T.SHOTOKU_R7[5] = (nm, 1.250, fuka, n)
        ng = [x for x in J.joritsu() if x[0] == "**NG**"]
        return "NG %d件（%s）" % (len(ng), ng[0][5]) if ng else ""
    finally:
        T.SHOTOKU_R7[:] = hon


def t_setten_kakunin_kazu():
    """**接点の「確かめること」を1件足したとき、黙って落ちないか。**

    以前は `zip(KAKUNIN, 効くもの)` で表を組んでいたため、
    KAKUNIN に1件足すと**短い方で止まり、7件目が出力されなかった。**
    表は6行のまま、どの点検も通ってしまう（接点の照合だけが
    「確認事項7がない」で捕まえた）。いまは数が合わなければ止まる。
    """
    import ono_setten as SET

    hon = list(SET.KAKUNIN)
    try:
        SET.KAKUNIN.append("否定試験のために足した確かめること")
        try:
            SET.render(_NullDoc(), _null_tbl, _null_body, _null_note, "―")
        except ValueError as e:
            return "数が合わないとして止まった（%s）" % str(e)[:40]
        return ""
    finally:
        SET.KAKUNIN[:] = hon


class _NullDoc:
    """render に渡すだけの受け皿。何も書かない。"""

    def p(self, *a, **k):
        pass

    def spacer(self, *a, **k):
        pass

    def h2(self, *a, **k):
        pass

    def h3(self, *a, **k):
        pass

    def src(self, *a, **k):
        pass


def _null_tbl(*a, **k):
    pass


def _null_body(*a, **k):
    pass


def _null_note(*a, **k):
    pass


def t_tsuika_genpon():
    """**他素案へ入れる前回計画の目標値を1つ変えると、原本照合が止まるか。**

    対応表の第7期の目標値は書き写さず、前回計画の該当頁にあるかを毎回見る。
    """
    import build_ono_shogai_tsuika_taiou as T

    hon = list(T.ZENKAI_MOKUHYO)
    try:
        g, k, v, pg, _w = hon[5]
        T.ZENKAI_MOKUHYO[5] = (g, k, v, pg, ["159件"])   # 原本は158件
        ng = T.genpon_ng()
        return "NG %d件（%s）" % (len(ng), ng[0]) if ng else ""
    finally:
        T.ZENKAI_MOKUHYO[:] = hon


def t_ronten_genten():
    """**論点漏れの点検の根拠の文言を1字変えると、原典照合が止まるか。**"""
    import build_ono_shogai_ronten_more as R

    hon = list(R.RONTEN)
    try:
        x = list(R.RONTEN[0])
        x[4] = ["の数を除いて設定するものとしない"]     # 原典は「ものとする」
        R.RONTEN[0] = tuple(x)
        ng = R.genten_ng()
        return "NG %d件（%s）" % (len(ng), ng[0]) if ng else ""
    finally:
        R.RONTEN[:] = hon


def t_mijotei_shirushi():
    """**未確定の印でない【　】を本文に書くと止まるか。**

    以前は「数えないもの」の一覧で除いていたため、本文に
    【令和9年4月1日新設】のような強調を書くと、**未確定箇所の数に
    紛れ込んだ**（166箇所が170箇所と出た）。
    いまは5種類だけを数え、知らない【　】があれば例外を投げる。
    """
    import build_ono_soan_elderly as E

    class _P:
        def __init__(self, t):
            self.text = t

    class _Doc:
        def __init__(self, texts):
            self.paragraphs = [_P(t) for t in texts]
            self.tables = []

    # 定義した印だけ → 数えられる
    ok = E.count_markers(_Doc([E.TBD, E.TGT, "【図3】", "【現状と課題】"]))
    if sum(ok.values()) != 2:
        return ""
    # 知らない印 → 例外
    try:
        E.count_markers(_Doc(["本文です。【令和9年4月1日新設】の規定です。"]))
    except ValueError:
        return "定義していない【　】で例外を投げた"
    return ""


def t_kakunin_no():
    """**本文から番号を引く `_no` が、当たらないときに黙らないか。**

    以前、本文に番号を直書きしていたため、通し番号を振り直したとき
    本文だけが古い番号のまま残っていた（基金残高を№25と書いていたが
    実際は№27）。`_no` は当たるものが1件でなければ例外を投げる。
    """
    import build_ono_kakunin_ichiran as K

    for kotoba in ("この文字列はどの事項にも当たらない", "の"):
        try:
            K._no(kotoba)
        except ValueError:
            continue
        return ""   # 例外を投げなかった＝素通り
    # 正しく引けることも併せて確かめる
    if not K._no("介護給付費準備基金の残高").startswith("№"):
        return ""
    return "当たらない語・多く当たる語のいずれでも例外を投げた"



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
    ("障がい素案　自己点検", "手帳の内訳（令和元年の身体）を1人増やす",
     t_shogai_techo),
    ("前回計画の読み取り", "生活介護の実績を1人日増やす", t_zenkai_pdf),
    ("前回計画の読み取り", "第2部の施策から第10章を落とす",
     t_zenkai_shisaku),
    ("条文番号の原典確認", "調和の規定を第7項から第5項に戻す", t_jobun),
    ("条文番号の原典確認（時点）",
     "改正後にしかない条項を「現行」として引く", t_jobun_jiten),
    ("確認事項の一覧　本文の番号",
     "当たらない語・多く当たる語で番号を引く", t_kakunin_no),
    ("未確定箇所の数え方",
     "未確定の印でない【　】を本文に書く", t_mijotei_shirushi),
    ("保険料の標準乗率と政令の突合",
     "第6段階の標準割合を1.200から1.250にする", t_joritsu),
    ("他素案への追加候補　前回計画の目標値",
     "相談支援の指導・助言の目標を158件から159件にする", t_tsuika_genpon),
    ("論点漏れの点検　原典照合",
     "継続入所者の控除の文言を「ものとしない」に変える", t_ronten_genten),
    ("両計画の接点　確かめることの数",
     "「確かめること」を1件足す（表の組み方が黙って落とさないか）",
     t_setten_kakunin_kazu),
    ("町データ待ちの切り分け", "切り分けから№32を落とす", t_kiriwake),
    ("町データ待ちの切り分け", "確認事項に無い№を切り分けに入れる",
     t_kiriwake_yokei),
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
