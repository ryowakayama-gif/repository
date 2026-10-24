# -*- coding: utf-8 -*-
"""概要版を最新の協議用素案から組み直す（令和8年10月23日）

  python3 07_ソーススクリプト/build_gaiyo_v230.py

ご指示「概要版を最新の協議用素案から組み直して下さい」（令和8年10月23日）による。

従前の概要版（build_gaiyo_R8.10.4.py／R8.10.5.py・計画素案_概要版_R8.10.5.docx）は
素案 Ver.2.20 を読み、本文の数値をスクリプトに書いていたため、次が素案と合わなくなっていた。
  ・4つの重点の名称が素案 4-3 と違う（本文・図・囲みの3か所でそれぞれ別の名称）
  ・素案にない数値（全国・県の高齢化率を本文に書き、県内順位は令和8年6月末の値と並べていた）
  ・素案で［要確認］のもの（認知症サポーター累計642名・チームオレンジの整備年度）を言い切っていた
  ・ヘッダーに版（Ver.2.16）、書体は游ゴシック
本スクリプトは**数値・名称をすべて協議用素案の docx から読む**（概要版のために数値を作らない。
CLAUDE.md §3）。住民向けの言い換えの文はここに置くが、数値はいずれも素案の文・表・図から引く。

入力　01_第10期_最新版成果品/川崎町_計画書素案_v2.30_協議用.docx（内容の正本）
      01_第10期_最新版成果品/川崎町_計画書案_公表用_v2.30.docx（書式の雛形：書体・ヘッダー・頁番号）
      07_ソーススクリプト/data_zuhyo.py（図の数値の正本）
出力　01_第10期_最新版成果品/川崎町_計画書案_概要版_v2.30.docx
      08_図表/概要版/g1〜g8.png（1色刷り）

仕様書8　A4版／コート紙／1色刷り／8頁程度／3,500部
⚠ 素案で【町確認】【委員会協議】［要確認］［要協議］のものは載せない（住民の方に配るもの）。
⚠ 保険料は第3回策定委員会で決まる。確定前である旨を明記する。
⚠ 紙面に版を出さない。版はファイル名で管理する（公表用と同じ）。
終了コード　0＝自己点検 全件適合／1＝1件でも不適合
"""
import collections
import copy
import importlib.util
import os
import re
import shutil
import sys
import zipfile

import docx
from docx.oxml.ns import qn
from docx.shared import Emu

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from fix_soan_v111 import clone_table, set_el            # noqa: E402

BAN = "2.33"
SOAN = f"01_第10期_最新版成果品/川崎町_計画書素案_v{BAN}_協議用.docx"
HINAGATA = f"01_第10期_最新版成果品/川崎町_計画書案_公表用_v{BAN}.docx"
OUT = f"01_第10期_最新版成果品/川崎町_計画書案_概要版_v{BAN}.docx"


def load(name, fname):
    spec = importlib.util.spec_from_file_location(name, os.path.join(HERE, fname))
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


# 図（g1〜g7）と画像の後始末は従前の関数を使う（数値は data_zuhyo.py から読む）。
# g8（柱と重点）は名称を素案から読むため本スクリプトで作る。
G4 = load("gaiyo4", "build_gaiyo_R8.10.4.py")
FIGDIR = G4.FIGDIR
EMU_IN = G4.EMU_IN
MIKAKUTEI = ("【町確認】", "【委員会協議】", "［要確認］", "［要協議］", "［要内訳］", "確認中")


# ══════════════════════════ 素案から読む
class Soan:
    def __init__(self, path):
        self.d = docx.Document(path)
        self.P = [p.text for p in self.d.paragraphs]
        self.T = [[[c.text for c in r.cells] for r in t.rows] for t in self.d.tables]
        self.zen = "\n".join(self.P + [c for t in self.T for r in t for c in r])

    def hiku(self, pat, n=1):
        """正規表現に当たる箇所の値の組。当たる値が1通りでなければ止める。"""
        hit = {m.groups() for s in self.P for m in re.finditer(pat, s)}
        if len(hit) != 1:
            raise SystemExit(f"素案から1通りに読めない（{len(hit)}通り）：{pat}")
        g = hit.pop()
        return g if n > 1 else g[0]

    def hyo(self, *atama):
        """先頭の行の先頭のセルが atama のいずれかで始まる表。"""
        h = [t for t in self.T if any(t[0][0].strip().startswith(a) for a in atama)]
        if not h:
            raise SystemExit(f"素案の表が見つからない：{atama}")
        return h


def yomu(S):
    V = {}
    V["koreika"], V["jinko"], V["kourei"] = S.hiku(
        r"令和8年6月末時点で([\d.]+)%（総人口([\d,]+)人・うち65歳以上([\d,]+)人）", 3)
    V["koreika7"], V["juni"] = S.hiku(
        r"令和7年3月31日時点では([\d.]+)%であり、宮城県内35市町村中(\d+)位", 2)
    t = [t for t in S.hyo("区分") if "令和7年度末" in t[0][2]][0]
    g = [r for r in t if r[0].startswith("（再掲）後期高齢者")][0]
    V["kouki3"], V["kouki7"], V["kouki_r"] = g[1], g[2], g[4].replace("+", "＋")
    V["jukyu"] = S.hiku(r"サービス受給者は延べ([\d.]+)人")
    V["kyufu"], V["shisetsu"] = S.hiku(r"給付費合計は約([\d.]+)億円。施設サービスが([\d.]+)％", 2)
    V["jutoku"] = S.hiku(r"住所地特例被保険者は(\d+)人")
    V["keikaku"], V["jisseki"], V["tassei"], V["kyotaku"], V["shisetsu_u"] = S.hiku(
        r"令和7年度の給付費は、計画([\d,]+)千円に対し実績([\d,]+)千円（達成率([\d.]+)％）.*?"
        r"居宅サービスが計画を([\d.]+)％下回る一方、施設サービスは([\d.]+)％上回って", 5)
    # 3-4 第10期に向けた課題
    t = S.hyo("No.")
    t = [x for x in t if x[0][1] == "課題"][0]
    V["kadai"] = [r[1] for r in t[1:]]
    # 2-7 5年間の推移
    t = [t for t in S.T if len(t) > 1 and t[1][0].startswith("介護老人福祉施設（定員30人）の稼働率")][0]
    gyo = {r[0][:6]: r for r in t}
    V["mitsu"] = [r for r in t if r[0].startswith("地域密着型")][0]
    V["tanki"] = [r for r in t if r[0].startswith("短期入所")][0]
    V["kaigo"] = [r for r in t if r[0].startswith("施設・居住系の介護職員")][0]
    V["aki"] = S.hiku(r"増やした定員のうち約(\d+)人分が埋まっていない")
    V["jigyosho_su"] = S.hiku(r"これらを含む(\d+)事業所が町内にあります")
    # 3-4 の重点課題（5本）（令和8年10月24日のレビュー A07：11の課題との橋渡しを概要版にも）
    t = [t for t in S.T if t[0][:2] == ["重点課題", "根拠となる事実"]][0]
    V["juten_kadai"] = [[r[0], r[1]] for r in t[1:]]
    # 4-1 基本理念
    i = S.P.index("　▌ 基本理念（案）")
    V["rinen"] = S.P[i + 1].strip()
    i = S.P.index("　▌ サブタイトル（第10期で新たに付加）")
    V["sub"] = S.P[i + 1].strip()
    # 4-2 柱（名称と事業数）
    t = [t for t in S.T if t[0][:2] == ["柱", "内容"]][0]
    V["hashira"] = []
    for r in t[1:]:
        m = re.match(r"柱(\d)．(.+)", r[0].replace("\n", "").replace("　", ""))
        na = re.sub(r"【.*?】", "", m.group(2)).strip()
        jig = re.search(r"（(?:(\d+)節)?(\d+)項(\d+)事業）", r[1])
        V["hashira"].append((int(m.group(1)), na, int(jig.group(3))))
    V["kosei"] = S.hiku(r"(\d+)つの柱(\d+)節(\d+)項(\d+)事業により構成", 4)
    # 新規・拡充（節の見出しから柱へ）
    SHO = {"5-1": 1, "5-2": 2, "5-3": 3, "5-4": 4, "5-5": 6, "6-4": 5}
    shin = collections.Counter()
    kaku = collections.Counter()
    sec = None
    for s in S.P:
        m = re.match(r"^(\d+-\d+)　", s.strip())
        if m:
            sec = m.group(1)
        m = re.search(r"新たに立てる事業・広げる事業（新規(\d+)件・拡充(\d+)件）", s)
        if m:
            # 第8章の冒頭（8-1 の前）の表は直前の節番号 7-6 のもとにある。柱7（第8章）に当てる
            h = SHO.get(sec, 7 if sec and sec[:2] in ("7-", "8-") else None)
            if h is None:
                raise SystemExit(f"新規・拡充の節を柱に当てられない：{sec}")
            shin[h] += int(m.group(1))
            kaku[h] += int(m.group(2))
    V["shin"], V["kaku"] = shin, kaku
    # 新規の事業名（素案の「新たに立てる事業」の表の「新規」の行）
    V["shinki_mei"] = [re.sub(r"^（[^）]*）", "", r[1]).strip()
                       for t in S.T if t[0][:2] == ["区分", "事業"] for r in t[1:]
                       if r[0].strip() == "新規"]
    # 4-3 重点
    t = [t for t in S.T if t[0][:2] == ["重点", "テーマ"]][0]
    V["juten"] = [(r[0], r[1].replace("\n", "").replace(" ", "")) for r in t[1:]]
    # 第6章
    V["ninchi_kosei"] = S.hiku(r"6-4　認知症施策の体系（(\d+節\d+項\d+事業)）")
    V["ninchi_su"], V["ninchi_r"], V["mci_su"], V["mci_r"] = S.hiku(
        r"認知症の方は([\d,]+)人（([\d.]+)％）、軽度認知障害（ＭＣＩ）の方は([\d,]+)人（([\d.]+)％）", 4)
    t = [t for t in S.hyo("手") if t[0][0].replace("　", "") == "手続"][0]
    g = [r for r in t if r[0].startswith("認知症の人及びその家族等からの意見の聴取")][0]
    V["choshu"] = re.search(r"(令和\d+年\d+月から\d+月)に実施", g[3]).group(1)
    # 1-4 調査
    V["chosa"] = S.hiku(r"①②は(令和\d+年\d+月から\d+月)にかけて郵送により実施し、①は(\d+)件、②は(\d+)件", 3)
    # 1-2 ほかの計画
    t = [t for t in S.T if t[0][:3] == ["区分", "計画名", "整合の方針"]][0]
    V["jouikei"] = [r[1] for r in t if r[0] == "上位計画"][0]
    V["kenkei"] = [r[1] for r in t if r[0] == "県の計画"][0]
    # 9-3 保険料
    V["hok"] = S.hiku(r"パターンA ([\d,]+)円、パターンB ([\d,]+)円、パターンC ([\d,]+)円", 3)
    V["hok9"] = S.hiku(r"第9期の([\d,]+)円との差は")
    V["dankai"] = S.hiku(r"令和6年度から適用している(\d+)段階の乗率")
    # 10-4 KPI（未確定の印のない行だけ）
    t = [t for t in S.T if t[0][:3] == ["指標", "区分", "現状値"]][0]
    V["kpi"] = [[r[0], r[2], r[3]] for r in t[1:]
                if not any(w in r[2] + r[3] for w in MIKAKUTEI)]
    # 2-1 日常生活圏域
    V["chiku"] = S.hiku(r"日常生活圏域は(\d)圏域とし、町内を(\d)地区", 2)
    return V


# 住民向けの言い換え（素案の新規事業名の一部 → 説明）。数値は持たない。
# ⚠ 鍵は素案の「新規」の事業名に1件だけ当たること（自己点検）。
SHINKI_SETSUMEI = [
    ("家族介護教室", "家族介護教室",
     "介護をしているご家族に、介護の仕方や体の使い方をお伝えする教室です。"),
    ("ピアサポート", "ご家族の相談・交流の場",
     "同じように介護をしている方どうしが話し合える場をつくります。"),
    ("レスパイト", "短期入所（レスパイト）の利用のお手伝い",
     "介護をしているご家族が休めるよう、短期入所の利用をお手伝いします。"),
    ("仕事と介護の両立", "仕事と介護の両立のお手伝い",
     "介護のために仕事を辞めずにすむよう、相談に応じます。"),
    ("ヤングケアラー", "ヤングケアラー・ダブルケアへの対応",
     "若い世代の介護、子育てと介護が重なるご家庭に気づき、必要な支援につなぎます。"),
    ("本人ミーティング", "認知症の本人ミーティング",
     "認知症の本人が集まって語り合い、そのお話を町の取組に活かします。"),
    ("認知症の日", "認知症の日（9月21日）の行事",
     "認知症について知っていただく行事を、認知症の日と認知症月間（9月）に行います。"),
    ("乖離の把握と公表", "計画と実際の差をたしかめて公表すること",
     "計画した量と実際の利用の差を毎年たしかめ、結果を公表します。"),
]

SODAN = [
    ("介護保険・高齢者福祉のご相談", "川崎町保健福祉課"),
    ("介護・福祉・権利を守ることの総合相談", "川崎町地域包括支援センター"),
    ("もの忘れが気になるとき", "川崎町地域包括支援センター（もの忘れ相談）"),
    ("生活支援・福祉移送サービスのご案内", "川崎町社会福祉協議会"),
]

KOTOBA = [
    ("要介護認定", "介護が必要な状態かどうか、どの程度かを定める手続です。"
                   "全国共通の基準により、要支援1・2と要介護1〜5に分かれます。"),
    ("地域包括支援センター", "高齢者の総合相談の窓口です。介護・福祉・権利を守ることについて"
                             "相談に応じます。"),
    ("住所地特例", "施設に入るために他の市町村へ住所を移した方について、"
                   "移る前の市町村が引き続き保険者となる仕組みです。"),
    ("介護給付費準備基金", "保険料の余りをためておき、次の3年間の保険料の上がり方を"
                           "やわらげるために使うお金です。"),
    ("認知症サポーター", "認知症について正しく理解し、認知症の人とご家族を見守る応援者です。"),
    ("チームオレンジ", "認知症サポーターが、認知症の人やご家族の身近な支援につながるよう"
                       "結成するチームです。"),
    ("通いの場", "住民が主体となって運営する、体操や交流などの集まりです。"),
]


def g_hashira(V):
    """7つの柱と4つの重点（1色刷り）。名称は素案 4-2・4-3 による。"""
    import matplotlib.pyplot as plt
    from matplotlib.patches import FancyBboxPatch
    import textwrap
    fig, ax = plt.subplots(figsize=(6.4, 3.9), dpi=150)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 6.6)
    ax.axis("off")
    for i, (no, na, _) in enumerate(V["hashira"]):
        c, r = i % 4, i // 4
        x = 0.25 + c * 2.45
        y = 4.15 - r * 1.45
        dark = (no == 5)
        ax.add_patch(FancyBboxPatch((x, y), 2.25, 1.2,
                                    boxstyle="round,pad=0.03,rounding_size=0.10",
                                    facecolor=G4.K1 if dark else G4.K3, edgecolor=G4.K0,
                                    linewidth=1.3))
        ax.text(x + 1.125, y + 0.6, f"柱{no}\n" + "\n".join(textwrap.wrap(na, 9)),
                ha="center", va="center", fontproperties=G4.JP_XS,
                color="white" if dark else G4.K0, linespacing=1.3)
    ax.text(5.0, 6.15, f"{len(V['hashira'])}つの柱で進めます", ha="center",
            fontproperties=G4.JP_T, color=G4.K0)
    ax.text(5.0, 2.05, f"とくに力を入れる{len(V['juten'])}つのこと", ha="center",
            fontproperties=G4.JP_M, color=G4.K0)
    for i, (j, te) in enumerate(V["juten"]):
        ax.text(0.3 + (i % 2) * 5.0, 1.45 - (i // 2) * 0.6, f"■ {j}　{te}",
                ha="left", fontproperties=G4.JP_XS, color=G4.K0)
    return G4.save(fig, "g8_hashira.png")


def honbun(V):
    """概要版の並び。（種別, 内容[, 列幅]）。種別：表題／副題／見出し／本文／囲み／図／T"""
    h9 = V["hok9"]
    a, b, c = V["hok"]
    m, tn, kg = V["mitsu"], V["tanki"], V["kaigo"]
    sa = lambda x: int(x.replace(",", "")) - int(h9.replace(",", ""))   # noqa: E731
    fugo = lambda v: ("＋" if v > 0 else "▲") + f"{abs(v):,}"           # noqa: E731
    hashira_hyo = [["7つの柱", "事業の数", "うち新しく始める・広げるもの"]]
    for no, na, n in V["hashira"]:
        hashira_hyo.append([f"柱{no}　{na}", f"{n}", f"{V['shin'][no] + V['kaku'][no]}"])
    hashira_hyo.append(["合計", f"{sum(n for _, _, n in V['hashira'])}",
                        f"{sum(V['shin'].values()) + sum(V['kaku'].values())}"])
    shin_hyo = [["新しく始めること", "どういうものか"]]
    for kagi, na, setsu in SHINKI_SETSUMEI:
        shin_hyo.append([na, setsu])
    P = [
        ("表題", "川崎町高齢者保健福祉計画"),
        ("表題", "第10期介護保険事業計画（案）"),
        ("表題", "概　要　版"),
        ("本文", "令和9年度から令和11年度までの3年間の、高齢者福祉と介護保険の進め方を"
                 "定める計画の案です。65歳以上の方の介護保険料と、町が行う高齢者福祉の"
                 "3年間を定めます。"),
        ("囲み", "この冊子は計画の案の段階のものです。保険料の額などは、これからの策定委員会の"
                 "協議を経て決まります。"),
        ("T", [["この計画のこと", "内容"],
               ["計画の名前", "川崎町高齢者保健福祉計画・第10期介護保険事業計画"],
               ["計画の期間", "令和9年度から令和11年度までの3年間"],
               ["中身", "高齢者福祉の進め方（高齢者保健福祉計画）と、介護保険の3年間の見通し・"
                        "保険料（介護保険事業計画）を1つにまとめたものです。第10期からは、"
                        "認知症の取組（認知症施策推進計画）もこの中に位置づけます。"],
               ["ほかの計画との関係", f"町の最上位の計画である「{V['jouikei']}」のもとで、"
                                      "地域福祉計画・障害者計画などと足並みをそろえます。"
                                      f"宮城県の「{V['kenkei'].split('（')[0]}」とも"
                                      "整合を図ります。"],
               ["つくりかた", f"高齢者の方へのアンケート（{V['chosa'][0]}・{V['chosa'][1]}件と"
                              f"{V['chosa'][2]}件のご回答）、策定委員会でのご審議、認知症の本人と"
                              "ご家族からの意見の聴取、みなさまからのご意見の募集を経てつくります。"],
               ["計画の見直し", "3年ごとに見直します。毎年度、進み具合をたしかめて公表します。"]],
         [30, 70]),
        ("見出し", "川崎町の高齢者のいま"),
        ("本文", f"川崎町の高齢化率は令和8年6月末で{V['koreika']}％です（総人口{V['jinko']}人の"
                 f"うち65歳以上が{V['kourei']}人）。県内の順位をくらべられる令和7年3月31日では"
                 f"{V['koreika7']}％で、宮城県内35市町村のなかで{V['juni']}番目に高い町です。"),
        ("図", "g2_koreika.png"),
        ("図", "g1_nenrei.png"),
        ("本文", f"65歳以上の方の数はほぼ横ばいですが、75歳以上の方は4年で{V['kouki3']}から"
                 f"{V['kouki7']}へ{V['kouki_r']}増えています。介護が必要になる方が増えるのは、"
                 "この年代からです。"),
        ("見出し", "介護サービスの利用とお金"),
        ("図", "g3_jukyusha.png"),
        ("本文", f"令和7年度は、ひと月あたり延べ{V['jukyu']}人の方が介護サービスを利用しました。"),
        ("図", "g4_kyufuhi.png"),
        ("本文", f"介護にかかったお金は年間約{V['kyufu']}億円です。そのうち約半分"
                 f"（{V['shisetsu']}％）は特別養護老人ホームなどの施設サービスです。"
                 f"町外の施設を利用している方も多く、令和8年3月末で{V['jutoku']}人の方が"
                 "住所地特例の対象となっています。"),
        ("見出し", "第9期（令和6〜8年度）をふりかえって"),
        ("図", "g5_taihi.png"),
        ("本文", f"令和7年度は、計画の{V['keikaku']}千円に対し、実際には{V['jisseki']}千円でした"
                 f"（{V['tassei']}％）。全体ではほぼ計画どおりですが、自宅で受けるサービスが"
                 f"計画を{V['kyotaku']}％下回る一方、施設のサービスは{V['shisetsu_u']}％"
                 "上回りました。"),
        ("本文", f"町内{V['jigyosho_su']}事業所の5年間の動きをみると、"
                 f"短期入所（ショートステイ）の利用はひと月{tn[1]}から{tn[3]}へほぼ半分になり、"
                 f"施設・居住系の介護職員数（常勤換算）は{kg[1]}から{kg[3]}に減っています。令和5年に定員を"
                 f"増やした地域密着型の特別養護老人ホームは、増やした定員のうち約{V['aki']}人分が"
                 "あいています。いま受けられているサービスを3年後も受けられるようにすることが、"
                 "この計画の大きな課題です。"),
        ("本文", f"第10期に向けて、{len(V['kadai'])}の課題を整理しました。これらを"
                 f"{len(V['juten_kadai'])}つの重点課題に集約して取り組みます。"),
        ("T", [["", "第10期に向けた課題"]]
         + [["①②③④⑤⑥⑦⑧⑨⑩⑪⑫"[i], k] for i, k in enumerate(V["kadai"])], [10, 90]),
        ("T", [["重点課題", "根拠となる事実"]] + V["juten_kadai"], [32, 68]),
        ("見出し", "第10期で大切にすること"),
        ("本文", f"基本理念「{V['rinen']}」を第9期から引き継ぎ、第10期では"
                 f"「{V['sub'].strip('〜 ')}」を副題に加えます。"),
        ("図", "g8_hashira.png"),
        ("本文", f"{V['kosei'][0]}つの柱に沿って{V['kosei'][3]}の事業を進め、"
                 f"{len(V['juten_kadai'])}つの重点課題のうち、とくに力を入れる{len(V['juten'])}つの"
                 "ことを重点として取り組みます。"),
        ("T", hashira_hyo, [56, 18, 26]),
        # 令和8年10月24日 再レビュー B12：柱の立て方（計画の章の組み立て）は委員会のご承認の前
        ("囲み", f"{V['kosei'][0]}つの柱の立て方（計画の章の組み立て）は、策定委員会でご審議いただく案であり、"
                 "まだ確定していません。"),
        ("見出し", "第9期から変わること"),
        ("本文", f"第10期では、新しく始める事業が{len(V['shinki_mei'])}件あります。"
                 "おもなものをお示しします。"),
        ("T", shin_hyo, [36, 64]),
        ("見出し", "認知症とともに生きるまちへ"),
        ("本文", "令和6年に認知症基本法が施行されました。この法律は、認知症になってからも"
                 "希望をもって自分らしく暮らすことができる社会をめざしています。"
                 "川崎町は第10期から、認知症の取組をこの計画の中に位置づけ"
                 f"（{V['ninchi_kosei']}）、とくに力を入れることの1つめとします。"),
        ("本文", f"町の65歳以上の方のうち、認知症の方は{V['ninchi_su']}人（{V['ninchi_r']}％）、"
                 f"軽度認知障害（ＭＣＩ）の方は{V['mci_su']}人（{V['mci_r']}％）と見込まれます。"),
        ("図", "g6_kpi.png"),
        ("本文", f"認知症の本人とご家族からお話をうかがったうえで計画をつくります"
                 f"（{V['choshu']}に行います）。認知症カフェ「喫茶みかん」、認知症サポーターの"
                 "養成、チームオレンジの活動、認知症初期集中支援チームによる訪問を続け、"
                 "第10期からは、認知症の本人が集まって語り合う「本人ミーティング」を"
                 "新しく始めます。"),
        ("囲み", "もの忘れが気になったら　川崎町地域包括支援センターにご相談ください。"),
        ("見出し", "介護保険料"),
        ("本文", "介護保険料は、3年間に必要な介護サービスのお金から、国・県・町が出す分と、"
                 "40歳から64歳の方が出す分を差し引いて、65歳以上の方で分け合って決めます。"
                 "これまでにためてきた基金（準備基金）を取りくずすと、保険料を抑えることが"
                 "できます。"),
        ("図", "g7_hokenryo.png"),
        ("本文", f"第9期の月額{h9}円に対して、3つの案があります。基金を取りくずさないと"
                 f"{a}円（{fugo(sa(a))}円）、半分を取りくずすと{b}円（{fugo(sa(b))}円）、"
                 f"全額を取りくずすと{c}円（{fugo(sa(c))}円）です。基金を使えばいまの保険料は"
                 "下がりますが、次の第11期に使える分がなくなります。"),
        ("囲み", "保険料の額は、第3回の策定委員会で決めます。上の金額は試算であり、"
                 f"まだ確定していません。保険料は、所得に応じて{V['dankai']}の段階に分かれます。"),
        ("見出し", "計画の進め方とご相談先"),
        ("本文", "計画は、毎年度その進み具合をたしかめ、結果を町のホームページで公表します。"
                 "計画に定めた目安がどこまで達成できたかを、策定委員会にご報告します。"
                 "おもな目安は次のとおりです。"),
        ("T", [["何をみるか", "いまのところ", "3年後の目安"]] + V["kpi"], [50, 22, 28]),
        ("T", [["こんなとき", "ご相談先"]] + [list(x) for x in SODAN], [46, 54]),
        ("本文", "この計画の案については、ご意見を募る手続（パブリックコメント）を行います。"
                 "時期と方法は、町の広報紙とホームページでお知らせします。"
                 "みなさまのご意見をお寄せください。"),
        ("見出し", "この冊子に出てくることば"),
    ]
    P += [("本文", f"{k}　{v}") for k, v in KOTOBA]
    P.insert(-len(KOTOBA), ("本文", f"日常生活圏域　サービスを考えるうえでの区域です。川崎町は"
                                    f"町全体を{V['chiku'][0]}つの区域とし、{V['chiku'][1]}つの"
                                    "地区ごとに状況を把握します。"))
    P.append(("本文", "［発行］川崎町保健福祉課"))
    # 見出しに番号
    out, i = [], 0
    for item in P:
        if item[0] == "見出し":
            i += 1
            out.append(("見出し", f"{'０１２３４５６７８９'[i]}　{item[1]}"))
        else:
            out.append(item)
    return out


PPR_ATO = ("w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
           "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
           "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
           "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
           "w:textDirection", "w:textAlignment", "w:textboxTightWrap",
           "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange")


def kakomi(el):
    """段落に罫線の枠と淡い地色を付ける（1色刷り）。pPr の子の順を守る。"""
    from docx.oxml import OxmlElement
    pPr = el.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        el.insert(0, pPr)
    for t in ("w:pBdr", "w:shd"):
        x = pPr.find(qn(t))
        if x is not None:
            pPr.remove(x)
    bdr = OxmlElement("w:pBdr")
    for side in ("top", "left", "bottom", "right"):
        e = OxmlElement(f"w:{side}")
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), "6")
        e.set(qn("w:space"), "4")
        e.set(qn("w:color"), "595959")
        bdr.append(e)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), "EDEDED")
    ato = [c for c in pPr if c.tag in {qn(t) for t in PPR_ATO}]
    if ato:
        ato[0].addprevious(bdr)
    else:
        pPr.append(bdr)
    bdr.addnext(shd)
    return el


def soroe(tbl):
    """表の文字揃え。見出しの行は中央、数だけのセルは右、ほかは左
    （中央揃えは表中の文言が見えにくい。令和8年10月のご指示）。"""
    from docx.oxml import OxmlElement
    JUN = ["cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit",
           "trHeight", "tblHeader", "tblCellSpacing", "jc", "hidden"]
    for ri, tr in enumerate(tbl.findall(qn("w:tr"))):
        # 行を頁で分けない。見出しの行は頁ごとに繰り返す
        trPr = tr.find(qn("w:trPr"))
        if trPr is None:
            trPr = OxmlElement("w:trPr")
            tc0 = tr.find(qn("w:tc"))
            tc0.addprevious(trPr)
        for t in (["cantSplit", "tblHeader"] if ri == 0 else ["cantSplit"]):
            if trPr.find(qn("w:" + t)) is None:
                trPr.append(OxmlElement("w:" + t))
        for c in sorted(list(trPr), key=lambda c: JUN.index(c.tag.split("}")[1])
                        if c.tag.split("}")[1] in JUN else 99):
            trPr.append(c)
        for tc in tr.findall(qn("w:tc")):
            s = "".join(t.text or "" for t in tc.iter(qn("w:t"))).strip()
            val = "center" if ri == 0 else ("right" if re.fullmatch(r"[\d,]+", s) else "left")
            for p in tc.findall(qn("w:p")):
                pPr = p.find(qn("w:pPr"))
                if pPr is None:
                    pPr = OxmlElement("w:pPr")
                    p.insert(0, pPr)
                jc = pPr.find(qn("w:jc"))
                if jc is None:
                    jc = OxmlElement("w:jc")
                    nochi = [c for c in pPr if c.tag in (qn("w:textDirection"), qn("w:textAlignment"),
                                                          qn("w:outlineLvl"), qn("w:rPr"))]
                    if nochi:
                        nochi[0].addprevious(jc)
                    else:
                        pPr.append(jc)
                jc.set(qn("w:val"), val)
    return tbl


def tsugi(el):
    from docx.oxml import OxmlElement
    pPr = el.find(qn("w:pPr"))
    if pPr.find(qn("w:keepNext")) is None:
        e = OxmlElement("w:keepNext")
        st = pPr.find(qn("w:pStyle"))
        if st is not None:
            st.addnext(e)
        else:
            pPr.insert(0, e)


def kaipage(el):
    from docx.oxml import OxmlElement
    pPr = el.find(qn("w:pPr"))
    if pPr.find(qn("w:pageBreakBefore")) is None:
        e = OxmlElement("w:pageBreakBefore")
        mae = [c for c in pPr if c.tag in (qn("w:pStyle"), qn("w:keepNext"), qn("w:keepLines"))]
        if mae:
            mae[-1].addnext(e)
        else:
            pPr.insert(0, e)


def main():
    for p in (SOAN, HINAGATA):
        if not os.path.exists(p):
            raise SystemExit("入力がない：" + p)
    S = Soan(SOAN)
    V = yomu(S)
    figs = [fn() for k, fn in G4.FIGS if k != "g8"] + [g_hashira(V)]
    P = honbun(V)

    shutil.copy(HINAGATA, OUT)
    doc = docx.Document(OUT)
    body = doc.element.body
    tpl = {}
    for el in body.iterchildren():
        if el.tag == qn("w:p"):
            s = "".join(n.text or "" for n in el.iter(qn("w:t"))).strip()
            if el.findall(".//" + qn("w:drawing")):
                tpl.setdefault("図", copy.deepcopy(el))
                continue
            if "表題" not in tpl and s.startswith("第 1 章"):
                tpl["表題"] = copy.deepcopy(el)
            if "見出し" not in tpl and s.startswith("▌"):
                tpl["見出し"] = copy.deepcopy(el)
            if "本文" not in tpl and s.startswith("令和7年度から令和8年度への"):
                tpl["本文"] = copy.deepcopy(el)
        elif el.tag == qn("w:tbl"):
            n = len(el.findall(qn("w:tr"))[0].findall(qn("w:tc")))
            if n in (2, 3) and f"T{n}" not in tpl and len(el.findall(qn("w:tr"))) > 2:
                tpl[f"T{n}"] = copy.deepcopy(el)
    for k in ("表題", "見出し", "本文", "図", "T2", "T3"):
        if k not in tpl:
            raise SystemExit("雛形が見つからない：" + k)
    tpl["囲み"] = kakomi(copy.deepcopy(tpl["本文"]))

    sectPr = body.find(qn("w:sectPr"))
    for el in list(body.iterchildren()):
        if el is not sectPr:
            body.remove(el)

    def add(el):
        sectPr.addprevious(el)

    n_fig = n_tbl = 0
    midashi1 = True
    for item in P:
        kind, v = item[0], item[1]
        if kind == "図":
            doc.add_picture(os.path.join(FIGDIR, v), width=Emu(int(4.8 * EMU_IN)))  # 重点課題の表を加えた（R8.10.24）ため 5.2→4.8 で8頁に収める
            p = doc.paragraphs[-1]._p
            p.getparent().remove(p)
            pPr = tpl["図"].find(qn("w:pPr"))
            if pPr is not None:
                old = p.find(qn("w:pPr"))
                if old is not None:
                    p.remove(old)
                p.insert(0, copy.deepcopy(pPr))
            add(p)
            n_fig += 1
        elif kind == "T":
            add(soroe(clone_table(tpl[f"T{len(v[0])}"], v, item[2])))
            e = copy.deepcopy(tpl["本文"])
            set_el(e, "")
            add(e)
            n_tbl += 1
        else:
            el = copy.deepcopy(tpl[kind])
            set_el(el, v)
            if kind == "見出し":
                tsugi(el)   # 見出しだけが頁の下に残らないようにする
            if kind == "見出し" and midashi1:
                # 1頁目（表紙と「この計画のこと」）の後で頁を改める
                kaipage(el)
                midashi1 = False
            add(el)
    doc.save(OUT)
    G4.strip_unused_media(OUT)

    # ══════════════════════════ 自己点検
    d2 = docx.Document(OUT)
    text = "\n".join(q.text for q in d2.paragraphs)
    tbl = "\n".join(c.text for t in d2.tables for r in t.rows for c in r.cells)
    zen = text + "\n" + tbl
    hdr = "\n".join(p.text for s in d2.sections for p in s.header.paragraphs)
    ng = []

    def chk(ok, mi):
        print(("   ○ " if ok else "   × ") + mi)
        if not ok:
            ng.append(mi)

    print("概要版を最新の協議用素案から組み直す")
    print(f"  内容の正本 {SOAN}")
    print("  ── 検査1　図・表")
    z = zipfile.ZipFile(OUT)
    n_media = sum(1 for n in z.namelist() if n.startswith("word/media"))
    chk(n_fig == 8 and n_media == 8, f"図 {n_fig}点・画像 {n_media}点（8点）")
    chk(len(d2.tables) == n_tbl, f"表 {len(d2.tables)}件")
    print("  ── 検査2　未確定のもの・受託者の語・他団体名・版を載せていない")
    for w in MIKAKUTEI + ("⚠", "確認事項No", "受託者", "当社", "ビズアップ", "スクリプト",
                          "**", "大雪", "北塩原", "金ヶ崎", "金ケ崎", "川崎市", "Ver.", "協議用"):
        chk(w not in zen + hdr, f"「{w}」0件")
    chk(not re.search(r"v\d+\.\d+", zen + hdr), "版の番号（v2.xx）0件")
    # 令和8年10月24日のレビュー A10・A11 の再発防止
    chk(not re.search(r"[1-9]\d+つ", zen), "「11つ」のような2桁の数＋「つ」0件")
    chk(not re.search(r"\d+の事業所", zen), "「8の事業所」の形 0件")
    chk("（常勤換算）" in zen, "介護職員数に（常勤換算）を明記")
    chk(all(k[0] in zen for k in V["juten_kadai"]), f"重点課題{len(V['juten_kadai'])}本の表（素案 3-4 と同じ名称）")
    print("  ── 検査3　数値はすべて素案にある")
    sujı = set(re.findall(r"\d[\d,]*(?:\.\d+)?(?=[％%人円件億千])", zen))
    zu = " ".join(str(x) for d in G4.ZU for r in d["rows"] for x in r)
    nai = sorted(x for x in sujı if x not in S.zen and x not in zu
                 and not (x.isdigit() and int(x) < 100))
    # 差（＋322円 等）は素案の本文にもある
    chk(not nai, f"素案・図の数値にない数 {len(nai)}件　" + "／".join(nai[:8]))
    print("  ── 検査4　名称は素案と同じ")
    for no, na, _ in V["hashira"]:
        chk(f"柱{no}　{na}" in zen, f"柱{no}「{na}」")
    for j, te in V["juten"]:
        chk(te in S.zen.replace("\n", "").replace(" ", ""), f"{j}「{te}」が素案 4-3 にある")
    chk(V["rinen"] in zen, f"基本理念「{V['rinen']}」")
    chk(sum(n for _, _, n in V["hashira"]) == int(V["kosei"][3]),
        f"柱ごとの事業数の和 {sum(n for _, _, n in V['hashira'])} ＝ 素案の {V['kosei'][3]}事業")
    print("  ── 検査5　新しく始めることの鍵は素案の新規事業に1件ずつ当たる")
    for kagi, na, _ in SHINKI_SETSUMEI:
        h = [x for x in V["shinki_mei"] if kagi in x]
        chk(len(h) == 1, f"「{kagi}」→ {h[0] if h else '当たりなし'}")
    print("  ── 検査6　ご相談先は素案にある名称")
    for _, na in SODAN:
        chk(na.split("（")[0] in S.zen, f"「{na.split('（')[0]}」")
    print("  ── 検査7　見出しの番号が順")
    mi = [q.text.strip()[0] for q in d2.paragraphs if re.match(r"^[０-９]　", q.text.strip())]
    chk(mi == list("１２３４５６７８９")[:len(mi)], f"見出し {len(mi)}件")
    print("  ── 検査8　保険料が確定前である旨")
    chk("まだ確定していません" in zen, "確定前である旨の断り")
    chk("柱の立て方（計画の章の組み立て）は、策定委員会でご審議いただく案" in zen,
        "柱の立て方が委員会のご承認の前である旨（再レビュー B12）")
    n_char = len(text.replace("\n", "")) + len(tbl.replace("\n", ""))
    print(f"  {OUT}　段落 {len(d2.paragraphs)}／図 {n_fig}／表 {len(d2.tables)}／字数 {n_char:,}字")
    if ng:
        print(f"   × 不適合 {len(ng)}件")
        sys.exit(1)
    print("   ○ 検査1〜8　不適合0件")


if __name__ == "__main__":
    main()
