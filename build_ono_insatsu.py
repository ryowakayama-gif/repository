"""小野町 第10期計画　印刷の仕様書と見積依頼の様式。

仕様書4の成果品のうち、（５）計画書と（６）概要版は印刷が納品物に含まれる。

  ※（５）（６）のデータ（ＣＤ―Ｒ）１部
    計画書印刷５０部＜仕様＞ Ａ４版、４色刷り ８０頁程度
  ※（７）のデータ（ＣＤ－Ｒ）１部、
    概要版印刷５０部＜仕様＞ Ａ４版、４色刷り ８頁

本ファイルが作るもの:
  1. 印刷の発注仕様書（docx）　印刷業者に示す条件。判型・色数・用紙・綴じ・
     校正の回数・納品場所・納期・個人情報と著作権の取扱いまで。
  2. 見積依頼の様式（xlsx）　業者が記入する内訳の様式と、
     集まった見積を同じ土俵で並べる比較表。
  3. 印刷工程の逆算（docx・xlsx の両方に入る）

設計方針:
 1. **仕様は仕様書の条文から引く。**判型・色数・頁数・部数を手で書かない。
    SPEC から組み立て、自己点検で条文と照合する。
 2. **日付は build_ono_kotei2 の逆算から引く。**
    工程変更協議資料と印刷の仕様書で校了日が食い違うことを防ぐ。
    印刷日数のつまみ（18日・25日）もあちらの CASE_E と同じものを使う。
 3. **見積の様式は「比べられる形」で出す。**
    業者ごとに内訳の切り方が違うと、合計額しか比べられない。
    工程を固定した様式に記入してもらう。
 4. 部数の増減単価を必ず書いてもらう。
    パブリックコメントの結果で頁数や部数が動くことがあるため。

出力:
  10_納品・印刷/小野町_第10期計画_印刷の発注仕様書_YYYYMMDD.docx
  10_納品・印刷/小野町_第10期計画_見積依頼書と内訳の様式_YYYYMMDD.xlsx
"""

import datetime
import pathlib
import re

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

import ono_style as S
import build_ono_kotei2 as K
from build_ono_gyomu_shinchoku import (BORDER, HEAD, SUB, body_style,
                                       head_row, widths)

ROOT = pathlib.Path(__file__).parent
OUT = ROOT / "小野町_引継ぎ_整理済" / "10_納品・印刷"
ASOF = "20261006"
ASOF_JP = "令和8年10月6日"

# 仕様書4 成果品（原典の表記のまま）。ここを唯一の出所にする。
SPEC = {
    "計画書": {
        "名称": "小野町高齢者保健福祉計画・第10期介護保険事業計画",
        "成果品": "（５）",
        "判型": "A4版", "色数": "4色刷り", "頁数": "80頁程度", "部数": 50,
        "条文": "計画書印刷５０部＜仕様＞ Ａ４版、４色刷り ８０頁程度",
    },
    "概要版": {
        "名称": "小野町高齢者保健福祉計画・第10期介護保険事業計画概要版",
        "成果品": "（６）",
        "判型": "A4版", "色数": "4色刷り", "頁数": "8頁", "部数": 50,
        "条文": "概要版印刷５０部＜仕様＞ Ａ４版、４色刷り ８頁",
    },
}

# 仕様書が定めていない条件。町と決めるもの。
# 「業者に任せる」と書くと見積が比べられなくなるため、当方の推奨を置く。
MIKETTEI = [
    ("用紙（本文）", "マットコート 73kg（四六判）",
     "4色刷りで写真と図表が入る。上質紙は図表のベタが沈む",
     "**町のご判断**"),
    ("用紙（表紙）", "コート紙 180kg（四六判）＋PP加工（マット）",
     "50部を窓口・図書館に置く。手に取られる前提で表紙を厚くする",
     "**町のご判断**"),
    ("綴じ（計画書）", "無線綴じ（くるみ製本）",
     "80頁は中綴じの限界（おおむね64頁）を超える",
     "仕様書の頁数から決まる"),
    ("綴じ（概要版）", "中綴じ（2つ折り・2か所針金）",
     "8頁。中綴じが最も安い",
     "仕様書の頁数から決まる"),
    ("色校正", "本紙校正 1回（計画書・概要版とも）",
     "4色刷りの色の出かたは画面と紙で変わる。簡易校正（DDCP）だと"
     "紙の風合いまでは見えない",
     "**町のご判断**（簡易校正にすると費用と日数が減る）"),
    ("校正の回数（文字）", "初校・再校の2回",
     "町の確認を各回に挟む。3回にすると校了が1週間後ろにずれる",
     "**町のご判断**"),
    ("納品場所", "小野町役場 保健福祉課",
     "分納（概要版を先に納める等）の要否は町のご指示による",
     "**町のご確認**"),
    ("奥付", "発行者・発行年月・問合せ先を記載",
     "発行年月は令和9年3月とする",
     "**町のご確認**"),
]

# 見積に記入してもらう工程。ここを固定することで業者間で比べられる。
MITSUMORI = [
    ("1", "企画・デザイン", "表紙デザイン、本文のフォーマット設計", "一式", 1),
    ("2", "組版・レイアウト（計画書）",
     "原稿（Word）からの流し込み、図表の配置", "頁", 80),
    ("3", "組版・レイアウト（概要版）",
     "計画書からの抽出・図表の作り直しを含む", "頁", 8),
    ("4", "図版の作成・加工", "本文中の図表の仕上げ（当方が元データを提供）",
     "点", 30),
    ("5", "校正（初校・再校）", "2回。赤字の反映を含む", "一式", 1),
    ("6", "色校正（本紙）", "計画書・概要版とも1回", "一式", 1),
    ("7", "刷版（計画書）", "4色×台数", "式", 1),
    ("8", "印刷・製本（計画書）", "A4・4色刷り・80頁・無線綴じ", "部", 50),
    ("9", "刷版（概要版）", "4色", "式", 1),
    ("10", "印刷・製本（概要版）", "A4・4色刷り・8頁・中綴じ", "部", 50),
    ("11", "簡易製本（成果品①〜④）",
     "介護保険給付費等分析結果報告書、サービス見込量推計結果報告書、"
     "計画骨子案、計画素案　各3部", "部", 12),
    ("12", "CD-Rの作成・ラベル印刷",
     "成果品①〜④で1部、（５）（６）で1部、（７）で1部", "枚", 3),
    ("13", "梱包・納品", "小野町役場 保健福祉課への搬入", "一式", 1),
]

# 見積のときに必ず書いてもらう条件。合計額だけでは比べられないもの。
JOKEN = [
    ("見積の有効期限", "提出日から60日以上",
     "発注が12月になるため、11月の見積が失効しないようにする"),
    ("部数の増減単価", "計画書・概要版とも 10部単位の増減単価",
     "パブリックコメントの結果で配布先が増えることがある"),
    ("頁数の増減単価", "計画書の8頁（1台）単位の増減単価",
     "「80頁程度」であり、素案の分量から前後する"),
    ("校正の追加単価", "文字校正を1回増やす場合の単価",
     "町の確認が3回になった場合に備える"),
    ("色校正の種別", "本紙校正と簡易校正（DDCP）の両方の単価",
     "費用と日数の比較のため"),
    ("最短の所要日数", "校了（下版）から納品までの営業日数",
     "**当方の逆算は18日と25日の2案で引いている。**"
     "これより短くできる業者があれば、協議会の日程に余裕が出る"),
    ("データの取扱い", "入稿データ・版下データの帰属と、業務終了後の消去",
     "町の著作物になるため"),
    ("再委託", "再委託の有無と、ある場合の委託先・範囲",
     "個人情報は含まないが、未公表の計画案を扱う"),
]


def jp(d):
    return K.jp(d) if hasattr(K, "jp") else d.isoformat()


def _pick(chain, key):
    for nm, d, _days in chain:
        if key in nm:
            return d
    raise SystemExit(f"逆算の鎖に「{key}」がありません")


def kotei():
    """印刷に関わる日付を build_ono_kotei2 の逆算から取り出す。

    ここで日付を書き直すと、工程変更協議資料と食い違う。

    逆算の各行の日付は**その工程を終える期限**である。
    したがって「★ 校了・下版」の日が印刷の着手（入稿）の日であり、
    「印刷・製本」の行の日付は刷り上がりの期限になる。
    """
    out = []
    for mark, kaigi_n, pdays in K.CASE_E:
        g = K.gyakusan(K.NOUKI_E, K.chain_e(kaigi_n, pdays))
        out.append({
            "案": mark, "協議会": kaigi_n, "印刷日数": pdays,
            "校了": _pick(g, "校了"),          # ＝印刷の着手（入稿）
            "印刷完了": _pick(g, "印刷・製本"),
            "納品": _pick(g, "納品・検収"),
            "原稿確定": _pick(g, "計画書の原稿確定"),
            "協議会2": _pick(g, "協議会 第2回"),
        })
    return out


MITSU_CH = [("発注（契約）", 3),
            ("見積の比較・業者の決定", 4),
            ("見積の提出期限", 14),
            ("★ 見積依頼の発送", 0)]


def _biz(d):
    """土日と年末年始（12月29日〜1月3日）に落ちた期限を手前に寄せる。

    工程変更協議資料の _biz は土日しか見ていない。印刷の発注は年末に
    かかるため、役場も印刷業者も閉まっている期間を避ける必要がある。
    """
    d = K._biz(d)
    while (d.month, d.day) >= (12, 29) or (d.month, d.day) <= (1, 3):
        d -= datetime.timedelta(days=1)
        d = K._biz(d)
    return d


def hacchu(genko_kakutei):
    """最遅の逆算。印刷業者の作業が始まる日から、見積依頼の期限を引く。

    **印刷業者の作業はレイアウト・組版から始まる。**校了ではない。
    逆算の鎖で「計画書の原稿確定」の次が「レイアウト・組版」であり、
    その時点で業者が決まっていなければならない。
    """
    ch = [("印刷業者の作業の開始（レイアウト・組版）", 0)] + MITSU_CH
    d = genko_kakutei
    out = []
    for nm, days in ch:
        out.append((nm, _biz(d), days))
        d -= datetime.timedelta(days=days)
    return list(reversed(out))


# 協議会 第2回で骨子案・見込量・保険料がそろう。ここで計画書の分量の見当が
# つくため、見積を取れるようになる。推奨はこの日を起点に順に進める。
SUISHO_MACHI = 7            # 協議会の後、見積依頼を出すまでの日数


def suisho(kaigi2):
    """推奨の日程。協議会 第2回を起点に、順に（逆算ではなく）進める。"""
    out, d = [], kaigi2 + datetime.timedelta(days=SUISHO_MACHI)
    out.append(("★ 見積依頼の発送", _biz(d), SUISHO_MACHI))
    for nm, days in reversed(MITSU_CH[:-1]):
        d += datetime.timedelta(days=days)
        out.append((nm, _biz(d), days))
    return out


# ---------------------------------------------------------------- 自己点検

def selfcheck(kt, hc, sc):
    ok, ng = [], []

    def chk(no, name, kitai, kekka):
        (ok if kitai == kekka else ng).append(
            (no, name, str(kitai), str(kekka)))

    # 1 仕様が条文と合っているか（判型・色数・頁数・部数を手で書いていないか）
    for k, v in SPEC.items():
        jo = v["条文"].translate(str.maketrans("０１２３４５６７８９",
                                               "0123456789"))
        hit = (v["判型"].replace("版", "") in jo.replace("Ａ", "A")
               and v["色数"].replace("色刷り", "") in jo.replace("４", "4")
               and str(v["部数"]) in jo
               and re.sub(r"[^0-9]", "", v["頁数"]) in jo)
        chk(f"1-{k}", f"{k}の仕様が仕様書の条文と合う", True, hit)

    # 2 見積の数量が仕様の部数・頁数と合う（別々に書いて食い違わないように）
    q = {r[1]: r[4] for r in MITSUMORI}
    chk("2a", "計画書の部数", SPEC["計画書"]["部数"], q["印刷・製本（計画書）"])
    chk("2b", "概要版の部数", SPEC["概要版"]["部数"], q["印刷・製本（概要版）"])
    chk("2c", "計画書の頁数",
        int(re.sub(r"[^0-9]", "", SPEC["計画書"]["頁数"])),
        q["組版・レイアウト（計画書）"])
    chk("2d", "概要版の頁数",
        int(re.sub(r"[^0-9]", "", SPEC["概要版"]["頁数"])),
        q["組版・レイアウト（概要版）"])

    # 3 逆算が単調（納期に近い工程ほど日付が後）
    for r in kt:
        chk(f"3-{r['案']}", f"案{r['案']}の逆算が順に進む", True,
            r["校了"] < r["印刷完了"] < r["納品"])

    # 4 納期を超えていない
    nk = datetime.date.fromisoformat(K.NOUKI_E)
    for r in kt:
        chk(f"4-{r['案']}", f"案{r['案']}の納品が納期以前", True, r["納品"] <= nk)

    # 5 見積依頼が基準日以降（すでに過ぎていたら逆算の意味がない）
    base = datetime.date.fromisoformat("2026-10-06")
    for mark, g in hc.items():
        chk(f"5-{mark}", f"案{mark}の見積依頼が基準日以降", True, g[0][1] >= base)

    # 6 印刷日数を削ると校了（入稿）が後ろになる（つまみが効いている）
    a25 = [r for r in kt if r["印刷日数"] == 25 and r["協議会"] == 4][0]
    a18 = [r for r in kt if r["印刷日数"] == 18 and r["協議会"] == 4][0]
    chk("6", "印刷日数を25日→18日にすると校了が後ろになる", True,
        a18["校了"] > a25["校了"])

    # 7 刷り上がりの期限は印刷日数によらず同じ（納品・検収から引くため）
    chk("8", "刷り上がりの期限は印刷日数によらない",
        1, len({r["印刷完了"] for r in kt}))

    # 7 見積の工程に番号の飛びがない
    nos = [int(r[0]) for r in MITSUMORI]
    chk("7", "見積の工程の番号", list(range(1, len(MITSUMORI) + 1)), nos)

    # 9 推奨の発注が最遅より前（前でなければ推奨の意味がない）
    for mark in hc:
        chk(f"9-{mark}", f"案{mark}の推奨の発注が最遅より前", True,
            sc[mark][-1][1] < hc[mark][-2][1])

    # 10 推奨の発注が1月末まで（＝2〜3月の印刷の集中より前に業者を押さえる）
    for mark in sc:
        d = sc[mark][-1][1]
        chk(f"10-{mark}", f"案{mark}の推奨の発注が令和9年1月末まで", True,
            d <= datetime.date(2027, 1, 31))

    # 11 日付が年末年始（12月29日〜1月3日）に落ちていない
    allday = [d for g in list(sc.values()) + list(hc.values())
              for _n, d, _y in g]
    chk("11", "年末年始に落ちた日付", 0,
        sum(1 for d in allday
            if (d.month, d.day) >= (12, 29) or (d.month, d.day) <= (1, 3)))

    # 12 土日に落ちた日付がない
    chk("12", "土日に落ちた日付", 0,
        sum(1 for d in allday if d.weekday() >= 5))

    return ok, ng


# ---------------------------------------------------------------- docx

def build_docx(kt, hc, sc, ok, ng):
    r = S.Report("小野町高齢者保健福祉計画・第10期介護保険事業計画　印刷の発注仕様書")
    r.cover(["第10期計画　印刷の発注仕様書（案）"],
            ["小野町高齢者保健福祉計画・第10期介護保険事業計画",
             "同　概要版",
             f"{ASOF_JP}"])
    r.toc_here()
    r.body_here()

    r.h1("1　本書の位置づけ")
    r.p("委託仕様書4の成果品のうち、**（５）計画書と（６）概要版は印刷が"
        "納品物に含まれます。**本書は、印刷業者に示す条件を一枚にまとめたものです。")
    r.p("委託仕様書が定めているのは**判型・色数・頁数・部数の4つだけ**であり、"
        "用紙・綴じ・校正の回数・納品場所は定めがありません。"
        "これらは見積額に大きく効くため、**業者に見積を依頼する前に町と決める"
        "必要があります。**3章に当方の推奨を置きました。")
    r.src("※ 本書は案であり、町のご確認をいただいたうえで業者に示します。")

    r.h1("2　委託仕様書が定めている条件")
    r.tbl([["区分", "成果品", "判型", "色数", "頁数", "部数"]] +
          [[k, v["成果品"], v["判型"], v["色数"], v["頁数"], f"{v['部数']}部"]
           for k, v in SPEC.items()],
          widths=[2.6, 1.8, 2.2, 2.4, 3.0, 5.2], right=(5,),
          caption="委託仕様書4が定めている印刷の条件")
    r.src("資料：小野町高齢者保健福祉計画・第10期介護保険事業計画策定業務委託"
          "仕様書 4 成果品。")
    r.p("このほか、**成果品①〜④は冊子（簡易製本）各3部**、"
        "**データ（CD-R）は①〜④で1部、（５）（６）で1部**が納品物です。")
    r.src("※ 仕様書には「（７）のデータ（ＣＤ－Ｒ）１部」との記載がありますが、"
          "成果品の列挙は（６）までで（７）がありません。"
          "概要版（６）に係る記載と解していますが、町にご確認いただく事項です"
          "（依頼票C-6）。本書の見積様式ではCD-Rを3枚で見ています。")

    r.h1("3　委託仕様書に定めがなく、町と決める条件")
    r.p("下表の「決め方」に**町のご判断**とあるものが、"
        "**見積を依頼する前にお決めいただきたい項目**です。")
    r.tbl([["項目", "当方の推奨", "推奨の理由", "決め方"]] +
          [list(x) for x in MIKETTEI],
          widths=[3.0, 4.2, 6.4, 3.6],
          caption="委託仕様書に定めのない条件と当方の推奨")
    r.p("**いずれも「業者の標準による」としないでください。**"
        "用紙も綴じも業者ごとに標準が違い、"
        "**同じ条件で見積を比べられなくなります。**")

    r.h1("4　見積依頼で必ず書いていただく条件")
    r.p("合計額だけでは比べられません。下表は**見積書に必ず記載いただく条件**です。")
    r.tbl([["項目", "記載いただく内容", "なぜ必要か"]] +
          [list(x) for x in JOKEN],
          widths=[3.2, 6.0, 8.0],
          caption="見積書に記載を求める条件")

    r.h1("5　印刷工程の逆算")
    r.p(f"動かせない期限は**納期（{jp(datetime.date.fromisoformat(K.NOUKI_E))}・"
        "委託仕様書3）**です。ここから逆算すると、"
        "**校了（下版）の日と印刷の着手日が決まります。**")
    r.tbl([["案", "協議会", "印刷日数", "校了・下版（＝入稿）",
            "刷り上がりの期限", "納品・検収"]] +
          [[f"案{r_['案']}", f"{r_['協議会']}回", f"{r_['印刷日数']}日",
            jp(r_["校了"]), jp(r_["印刷完了"]), jp(r_["納品"])] for r_ in kt],
          widths=[1.8, 2.0, 2.2, 3.8, 3.8, 3.6],
          caption="納期から逆算した印刷の日程（4案）")
    r.src("資料：小野町_工程変更協議資料_第2版。"
          "本表の日付は同資料と同じ計算（build_ono_kotei2.gyakusan）で引いて"
          "おり、両者がずれることはありません。")

    r.h2("（1）業者をいつまでに押さえるか")
    r.p("**印刷業者の作業は校了からではなく、レイアウト・組版から始まります。**"
        "上の鎖でいえば「計画書の原稿確定」の次です。"
        "その時点で業者が決まっていなければなりません。")
    late = min(g[0][1] for g in hc.values())
    rec = max(g[-1][1] for g in sc.values())
    r.tbl([["案", "印刷業者の作業の開始", "見積依頼の発送（最遅）",
            "発注（最遅）", "見積依頼の発送（推奨）", "発注（推奨）"]] +
          [[f"案{m}",
            jp([x for x in kt if x["案"] == m][0]["原稿確定"]),
            jp(hc[m][0][1]), jp(hc[m][-2][1]),
            jp(sc[m][0][1]), jp(sc[m][-1][1])] for m in hc],
          widths=[1.6, 3.2, 3.4, 2.8, 3.4, 2.8],
          caption="見積依頼と発注の期限（最遅）と推奨")

    r.h3("最遅")
    r.p("**逆算だけで見れば、発注は令和9年2月中旬まで待てます。**"
        f"どの案でも、見積依頼は{jp(late)}までに出せば間に合う計算です。")
    for mark, g in hc.items():
        case = [x for x in kt if x["案"] == mark][0]
        r.h3(f"案{mark}（協議会{case['協議会']}回・印刷{case['印刷日数']}日）の最遅")
        r.tbl([["工程", "期限", "所要日数"]] +
              [[nm, jp(d), f"{dy}日" if dy else "―"] for nm, d, dy in g],
              widths=[8.0, 5.2, 4.0], right=(2,))

    r.h3("推奨")
    r.p(f"**それでも、遅くとも{jp(rec)}までに発注を終えることをお勧めします。**"
        "逆算が成り立つのは、見積を依頼した業者がすぐに応じ、"
        "日程どおりに作業できる場合に限られます。"
        "2月から3月は自治体の計画書・報告書の印刷が集中する時期であり、"
        "**2月に発注先を探し始めると、色校正の差し戻しに応じる余裕が"
        "業者側に残っていないことがあります。**"
        "これは当方の実務上の見立てであり、逆算から出る数字ではありません。")
    r.p("起点は**協議会 第2回**です。ここで骨子案・見込量・保険料がそろい、"
        "**計画書の分量（頁数）の見当がつく**ため、見積を取れるようになります。")
    for mark, g in sc.items():
        case = [x for x in kt if x["案"] == mark][0]
        r.h3(f"案{mark}（協議会 第2回 {jp(case['協議会2'])}）の推奨")
        r.tbl([["工程", "日程", "前の工程からの日数"]] +
              [[nm, jp(d), f"{dy}日" if dy else "―"] for nm, d, dy in g],
              widths=[8.0, 5.2, 4.0], right=(2,))
    r.p("**推奨のとおり進めれば、発注は協議会の回数が多い案Aで"
        f"{jp(sc['A'][-1][1])}、最も遅い案でも{jp(rec)}に終わります。**"
        "最遅との差が、業者が見つからなかった場合や、"
        "色校正で刷り直しが生じた場合の余裕になります。"
        "**協議会を3回に減らす案（B・D）では協議会 第2回が12月下旬となり、"
        "発注が1月にずれます。**協議会の回数は、"
        "印刷の発注時期にも効くということです。")
    r.src("※ 業務進捗管理表およびペンディング整理では、本作業の理由を"
          "「12月までに発注しないと令和9年2月下旬の校了に間に合わない」と"
          "記載していました。**逆算で確かめたところ、これは誤りです。**"
          "校了に間に合わせるだけなら発注は2月中旬まで待てます。"
          "また、推奨の日程でも12月中に発注を終えられるのは案Aだけで、"
          "案B・C・Dは1月になります。早く発注すべき理由は校了の期限ではなく、"
          "年度末の印刷の集中です。記載を改めます。")

    r.h2("（2）日程が詰まった場合に削れるもの")
    r.tbl([["削るもの", "縮む日数", "失うもの"],
           ["色校正を本紙校正から簡易校正（DDCP）に", "3〜5日",
            "**紙に刷った色を見られない。**写真の肌色と図表のベタが"
            "画面と変わることがある"],
           ["文字校正を初校・再校の2回から初校のみに", "5日",
            "**再校が効かない。**初校の赤字が正しく反映されたかを"
            "紙で確かめられない"],
           ["印刷日数を25日から18日に", "7日",
            "**色校正の差し戻しが1回も効かない。**"
            "刷り直しが生じると納期に間に合わない"],
           ["概要版を計画書と分けて先に発注", "―",
            "刷版と用紙の手配が2回になり、**費用が増える**"]],
          widths=[5.0, 2.6, 9.6],
          caption="日程が詰まった場合に削れるものと、失うもの")
    r.p("**当方としては、色校正（本紙・1回）と文字校正（初校・再校）は"
        "残すことをお勧めします。**50部とはいえ3年間使われる計画書であり、"
        "刷り直しが利きません。")

    r.h1("6　見積依頼の様式")
    r.p("同名の xlsx を併せてお出しします。シートは3枚です。")
    r.tbl([["シート", "記入する人", "内容"],
           ["01_見積依頼書", "―（当方が作成）",
            "宛先・件名・提出期限・条件。業者にそのまま送れる"],
           ["02_内訳の様式", "**印刷業者**",
            f"下表の{len(MITSUMORI)}工程の単価と金額。"
            "業者ごとに内訳の切り方が違うと比べられないため、工程を固定する"],
           ["03_比較表", "―（当方が集計）",
            "集まった見積を同じ工程で横に並べ、合計と工程別の差を出す"]],
          widths=[3.4, 3.4, 10.4],
          caption="見積依頼の様式（xlsx）のシート構成")
    r.tbl([["No", "工程", "内容", "単位", "数量"]] +
          [[a, b, c, d, f"{e:,}"] for a, b, c, d, e in MITSUMORI],
          widths=[1.2, 4.2, 7.4, 1.8, 2.6], right=(4,),
          caption="見積の内訳の様式（02_内訳の様式）")
    r.src("※ 数量は現時点の見込みです。計画書の頁数は「80頁程度」であり、"
          "素案の分量から前後します。増減単価を併せてご記入いただきます。")

    r.h1("7　町にご確認いただきたいこと")
    for i, t in enumerate([
        "3章の「**町のご判断**」5項目（本文用紙・表紙用紙・色校正の種別・"
        "文字校正の回数・納品場所）。",
        "印刷業者の選定方法（指名競争・見積合わせ・随意契約のいずれか）と、"
        "**町で業者を指定されるか、当方で複数社から見積を取るか。**",
        "見積を依頼する業者の数（当方としては3社をお勧めします）。",
        "成果品①〜④の簡易製本（各3部）を、印刷業者に併せて発注するか、"
        "当方で行うか。",
        "「（７）のデータ（ＣＤ－Ｒ）１部」の解釈（2章の注記）。",
        "奥付の発行者名・問合せ先の表記。",
    ], start=1):
        r.p(f"{i}　{t}")

    r.h1("8　自己点検")
    r.p("本書の数値と日付が、委託仕様書および工程変更協議資料と"
        "食い違っていないことを機械で確かめています。")
    r.tbl([["No", "点検項目", "期待", "結果"]] +
          [[a, b, c, d] for a, b, c, d in ok + ng],
          widths=[1.6, 7.2, 4.2, 4.2],
          caption="自己点検の結果")
    r.p(f"**{len(ok)}項目すべて適合**" if not ng
        else f"**要修正 {len(ng)}項目**")

    OUT.mkdir(parents=True, exist_ok=True)
    r.build_toc()
    path = OUT / f"小野町_第10期計画_印刷の発注仕様書_{ASOF}.docx"
    r.save(path)
    return path


# ---------------------------------------------------------------- xlsx

def _title(ws, row, text):
    ws.cell(row=row, column=1, value=text).font = Font(bold=True, size=11,
                                                       color="1F3864")


def build_xlsx(kt, hc, sc, ok, ng):
    wb = openpyxl.Workbook()

    # ---- 01 見積依頼書
    ws = wb.active
    ws.title = "01_見積依頼書"
    widths(ws, [3, 22, 60, 26])
    rows = [
        ("", "", "", f"{ASOF_JP}"),
        ("", "", "", "（受託者名）"),
        ("", "", "", ""),
        ("", "件名", "小野町高齢者保健福祉計画・第10期介護保険事業計画"
                     "及び同概要版の印刷製本に係るお見積のお願い", ""),
        ("", "", "", ""),
        ("", "1　印刷物",
         "（1）小野町高齢者保健福祉計画・第10期介護保険事業計画"
         f"　{SPEC['計画書']['判型']}・{SPEC['計画書']['色数']}・"
         f"{SPEC['計画書']['頁数']}・{SPEC['計画書']['部数']}部", ""),
        ("", "",
         "（2）同　概要版"
         f"　{SPEC['概要版']['判型']}・{SPEC['概要版']['色数']}・"
         f"{SPEC['概要版']['頁数']}・{SPEC['概要版']['部数']}部", ""),
        ("", "", "（3）報告書等の簡易製本　4種　各3部", ""),
        ("", "", "（4）CD-Rの作成・ラベル印刷　3枚", ""),
        ("", "", "", ""),
        ("", "2　仕様", "別紙「印刷の発注仕様書」のとおり", ""),
        ("", "", "", ""),
        ("", "3　記入方法",
         "シート「02_内訳の様式」の網かけ欄にご記入ください。"
         "工程を固定しているのは、各社の見積を同じ区分で比べるためです。"
         "該当がない工程は「0」とし、別途必要な工程があれば"
         "「14」以降の行に追記してください。", ""),
        ("", "", "", ""),
        ("", "4　提出期限", "", ""),
        ("", "5　有効期限", "提出日から60日以上", ""),
        ("", "6　提出先", "", ""),
        ("", "", "", ""),
        ("", "7　併せて記載いただく条件", "", ""),
    ]
    for i, r_ in enumerate(rows, start=1):
        for j, v in enumerate(r_, start=1):
            c = ws.cell(row=i, column=j, value=v)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            if j == 2 and v:
                c.font = Font(bold=True, size=10)
    base = len(rows) + 1
    ws.cell(row=base, column=2, value="項目").font = Font(bold=True, size=9,
                                                          color="FFFFFF")
    ws.cell(row=base, column=3, value="記載いただく内容").font = Font(
        bold=True, size=9, color="FFFFFF")
    ws.cell(row=base, column=4, value="ご記入欄").font = Font(
        bold=True, size=9, color="FFFFFF")
    for col in (2, 3, 4):
        ws.cell(row=base, column=col).fill = HEAD
        ws.cell(row=base, column=col).border = BORDER
        ws.cell(row=base, column=col).alignment = Alignment(
            wrap_text=True, vertical="center", horizontal="center")
    for i, (a, b, _c) in enumerate(JOKEN, start=base + 1):
        ws.cell(row=i, column=2, value=a)
        ws.cell(row=i, column=3, value=b)
        ws.cell(row=i, column=4, value="")
        for col in (2, 3, 4):
            cc = ws.cell(row=i, column=col)
            cc.border = BORDER
            cc.font = Font(size=9)
            cc.alignment = Alignment(wrap_text=True, vertical="top")
            if col == 4:
                cc.fill = PatternFill("solid", fgColor="FFF2CC")
    first = min(g[0][1] for g in hc.values())
    ws.cell(row=15, column=4, value=f"{jp(first)}（厳守）").font = Font(
        bold=True, size=10, color="C00000")

    # ---- 02 内訳の様式
    ws = wb.create_sheet("02_内訳の様式")
    widths(ws, [5, 24, 44, 7, 8, 12, 14, 24])
    ws.append(["No", "工程", "内容", "単位", "数量", "単価（円）",
               "金額（円）", "備考"])
    head_row(ws)
    for a, b, c, d, e in MITSUMORI:
        ws.append([a, b, c, d, e, None, None, None])
    n = len(MITSUMORI)
    for i in range(2, n + 2):
        ws.cell(row=i, column=7,
                value=f"=IF(F{i}=\"\",\"\",E{i}*F{i})")
        for col in (6, 7, 8):
            ws.cell(row=i, column=col).fill = PatternFill(
                "solid", fgColor="FFF2CC")
    # 追記用の空行
    for k in range(3):
        ws.append([str(n + 1 + k), None, None, None, None, None, None, None])
        for col in (2, 3, 4, 5, 6, 7, 8):
            ws.cell(row=n + 2 + k, column=col).fill = PatternFill(
                "solid", fgColor="FFF2CC")
    last = n + 4
    ws.append(["", "小計", "", "", "", "", f"=SUM(G2:G{last})", ""])
    ws.append(["", "消費税（10％）", "", "", "", "",
               f"=ROUND(G{last + 1}*0.1,0)", ""])
    ws.append(["", "合計", "", "", "", "",
               f"=G{last + 1}+G{last + 2}", ""])
    body_style(ws, wrap_cols=(1, 2, 7))
    for i in range(last + 1, last + 4):
        for col in range(1, 9):
            ws.cell(row=i, column=col).fill = SUB
            ws.cell(row=i, column=col).font = Font(bold=True, size=9)
            ws.cell(row=i, column=col).border = BORDER
    for i in range(2, last + 4):
        ws.cell(row=i, column=6).number_format = "#,##0"
        ws.cell(row=i, column=7).number_format = "#,##0"
        ws.cell(row=i, column=5).number_format = "#,##0"
    ws.freeze_panes = "A2"

    r0 = last + 6
    _title(ws, r0, "■ 増減単価（必ずご記入ください）")
    ws.cell(row=r0 + 1, column=2, value="項目").font = Font(bold=True, size=9)
    ws.cell(row=r0 + 1, column=6, value="単価（円）").font = Font(bold=True,
                                                                  size=9)
    for i, t in enumerate([
        "計画書　10部増の場合の単価", "計画書　10部減の場合の減額",
        "概要版　10部増の場合の単価", "概要版　10部減の場合の減額",
        "計画書　8頁（1台）増の場合の単価",
        "文字校正を1回追加する場合の単価",
        "色校正を簡易校正（DDCP）とした場合の減額",
    ], start=r0 + 2):
        ws.cell(row=i, column=2, value=t).font = Font(size=9)
        ws.cell(row=i, column=2).border = BORDER
        c = ws.cell(row=i, column=6)
        c.fill = PatternFill("solid", fgColor="FFF2CC")
        c.border = BORDER
        c.number_format = "#,##0"

    # ---- 03 比較表
    ws = wb.create_sheet("03_比較表")
    widths(ws, [5, 24, 7, 8, 13, 13, 13, 16])
    ws.append(["No", "工程", "単位", "数量", "A社（円）", "B社（円）",
               "C社（円）", "最低額との差（最高額）"])
    head_row(ws)
    for a, b, _c, d, e in MITSUMORI:
        ws.append([a, b, d, e, None, None, None, None])
    for i in range(2, n + 2):
        ws.cell(row=i, column=8,
                value=f"=IF(COUNT(E{i}:G{i})=0,\"\",MAX(E{i}:G{i})"
                      f"-MIN(E{i}:G{i}))")
        for col in (5, 6, 7):
            ws.cell(row=i, column=col).fill = PatternFill(
                "solid", fgColor="E2EFDA")
    ws.append(["", "合計（税抜）", "", "",
               f"=SUM(E2:E{n + 1})", f"=SUM(F2:F{n + 1})",
               f"=SUM(G2:G{n + 1})", ""])
    ws.append(["", "消費税（10％）", "", "",
               f"=ROUND(E{n + 2}*0.1,0)", f"=ROUND(F{n + 2}*0.1,0)",
               f"=ROUND(G{n + 2}*0.1,0)", ""])
    ws.append(["", "合計（税込）", "", "",
               f"=E{n + 2}+E{n + 3}", f"=F{n + 2}+F{n + 3}",
               f"=G{n + 2}+G{n + 3}", ""])
    body_style(ws, wrap_cols=(1,))
    for i in range(n + 2, n + 5):
        for col in range(1, 9):
            ws.cell(row=i, column=col).fill = SUB
            ws.cell(row=i, column=col).font = Font(bold=True, size=9)
            ws.cell(row=i, column=col).border = BORDER
    for i in range(2, n + 5):
        for col in (5, 6, 7, 8):
            ws.cell(row=i, column=col).number_format = "#,##0"
    ws.freeze_panes = "A2"
    r0 = n + 6
    _title(ws, r0, "■ 金額以外の比較（価格だけで決めない）")
    for i, t in enumerate([
        "校了から納品までの所要日数（短いほど協議会の日程に余裕が出る）",
        "色校正の種別（本紙／簡易）と回数",
        "増減単価（部数・頁数）",
        "過去の自治体計画書の実績",
        "再委託の有無",
        "入稿データの形式（Word・InDesign・PDF）と、版下データの引渡しの可否",
    ], start=r0 + 1):
        ws.cell(row=i, column=2, value=t).font = Font(size=9)

    # ---- 04 工程（逆算）
    ws = wb.create_sheet("04_工程の逆算")
    widths(ws, [6, 10, 10, 16, 18, 16, 16])
    ws.append(["案", "協議会", "印刷日数", "見積依頼の発送",
               "校了・下版（＝入稿）", "刷り上がりの期限", "納品・検収"])
    head_row(ws)
    for r_ in kt:
        g = hc[r_["案"]]
        ws.append([f"案{r_['案']}", f"{r_['協議会']}回", f"{r_['印刷日数']}日",
                   jp(g[0][1]), jp(r_["校了"]), jp(r_["印刷完了"]),
                   jp(r_["納品"])])
    body_style(ws)
    r0 = len(kt) + 3
    _title(ws, r0, "■ 業者を押さえるまでの逆算（案ごと）")
    row = r0 + 1
    ws.cell(row=row, column=1, value="案").font = Font(bold=True, size=9)
    ws.cell(row=row, column=2, value="工程").font = Font(bold=True, size=9)
    ws.cell(row=row, column=4, value="期限").font = Font(bold=True, size=9)
    ws.cell(row=row, column=5, value="所要日数").font = Font(bold=True, size=9)
    row += 1
    for mark, g in hc.items():
        for nm, d, dy in g:
            ws.cell(row=row, column=1, value=f"案{mark}").font = Font(size=9)
            ws.cell(row=row, column=2, value=nm).font = Font(size=9)
            ws.cell(row=row, column=4, value=jp(d)).font = Font(size=9)
            ws.cell(row=row, column=5,
                    value=(f"{dy}日" if dy else "―")).font = Font(size=9)
            row += 1

    # ---- 05 自己点検
    ws = wb.create_sheet("05_自己点検")
    widths(ws, [8, 48, 24, 24, 10])
    ws.append(["No", "点検項目", "期待", "結果", "判定"])
    head_row(ws)
    for a, b, c, d in ok:
        ws.append([a, b, c, d, "OK"])
    for a, b, c, d in ng:
        ws.append([a, b, c, d, "要修正"])
    body_style(ws, wrap_cols=(1, 2, 3))

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / f"小野町_第10期計画_見積依頼書と内訳の様式_{ASOF}.xlsx"
    wb.save(path)
    return path


def main():
    kt = kotei()
    hc = {r["案"]: hacchu(r["原稿確定"]) for r in kt}
    sc = {r['案']: suisho(r['協議会2']) for r in kt}
    ok, ng = selfcheck(kt, hc, sc)

    dx = build_docx(kt, hc, sc, ok, ng)
    xl = build_xlsx(kt, hc, sc, ok, ng)
    print("出力:", dx)
    print("     ", xl)
    for a, b, c, d in ok:
        print(f"  OK {a} {b}　期待 {c}／結果 {d}")
    for a, b, c, d in ng:
        print(f"  !! {a} {b}　期待 {c}／結果 {d}")
    if ng:
        raise SystemExit(f"\n自己点検 要修正 {len(ng)}項目")
    print(f"\n自己点検 {len(ok)}項目 すべて適合")
    first = min(g[0][1] for g in hc.values())
    print(f"見積依頼の発送期限 {jp(first)}")


if __name__ == "__main__":
    main()
