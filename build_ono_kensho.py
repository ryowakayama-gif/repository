"""小野町 第10期 見込量算定の再検証（他町村の検証手順を当てはめる）。

令和8年9月25日のご指示
  「大雪広域及び川崎町のブランチ確認の上、見込量算定について再度検証を
    お願い致します。」

他町村の案件で用いている検証の軸を小野町のデータに当てはめた。
**他団体の数値は本書に載せない。方法のみを用いる。**

  ① 総給付費1％が保険料の月額で何円にあたるかを先に出し、論点の重みを測る
  ② 基準年度が何か月分かを実測し、第9期でも同じことが起きていなかったかを見る
  ③ 見える化の令和8年度を「同じ種別での上振れ」と「0になった種別の下振れ」に分ける
  ④ 見える化システムの自然体推計の仕様どおりに組み直し、当方の置き方との差を測る
  ⑤ 国のガイドが課す上限1.1・下限0.9に適合しているかを確かめる

出力
  04_算定・見込量/小野町_第10期_見込量算定の再検証_YYYYMMDD.docx
  04_算定・見込量/小野町_第10期_見込量算定の再検証_YYYYMMDD.xlsx
"""

import pathlib
import warnings

warnings.filterwarnings("ignore")

import openpyxl
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

import build_ono_tanka as T
import ono_mieruka_ws as W
import ono_shizentai as SZ
from build_ono_kaisu import chiiki_plan

ROOT = pathlib.Path(__file__).parent
OUT = ROOT / "小野町_引継ぎ_整理済" / "04_算定・見込量"
ASOF = "20260925"
ASOF_JP = "令和8年9月25日"
JP_MIN = "游明朝"
JP_GO = "游ゴシック"

HEAD = PatternFill("solid", fgColor="1F3864")
NG = PatternFill("solid", fgColor="FCE4E4")
WARN = PatternFill("solid", fgColor="FFE0B2")
OK = PatternFill("solid", fgColor="E2EFDA")
CALC = PatternFill("solid", fgColor="EAF1FB")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# 第9期計画の公表値（標準給付費見込額・千円）。計画書 令和6年3月。
DAI9_MIKOMI = {"令和6年度": 1185894, "令和7年度": 1189712, "令和8年度": 1190880}
YS = ["令和3年度", "令和4年度", "令和5年度", "令和6年度", "令和7年度"]


# ---------------------------------------------------------------- 算定

def kijun():
    """総給付費1％が保険料の月額で何円にあたるか。論点の重みを測る基準。"""
    A = SZ.plan_rows()["集計"]
    cp = chiiki_plan()
    chi3 = sum(v[3] for v in cp) * 3 / 1000
    sogo3 = sum(v[3] for v in cp if v[0] == "総合事業") * 3 / 1000
    pop3 = sum(T.POP_DAI10[y]["1号"] for y in T.PLAN_YEARS)
    std3, kyu3 = sum(A["標準給付費"]), sum(A["計"])
    base = T.premium(std3, chi3, sogo3, pop3)["月額"]
    r = std3 / kyu3
    up = T.premium(std3 + kyu3 * 0.01 * r, chi3, sogo3, pop3)["月額"]
    return {"標準給付費": std3, "総給付費": kyu3, "割増率": r,
            "月額": base, "1％": up - base,
            "円換算": lambda amt: (T.premium(std3 + amt * r, chi3, sogo3,
                                          pop3)["月額"] - base)}


def jisseki():
    """年報の総給付費と標準給付費（審査支払手数料を除く・千円）。

    **加算3項目は比で置かず、年報 様式3 の実額を用いる。**
    第9期の対比（04_算定・見込量／10_給付費等分析）と同じ組み方にそろえるため。
    審査支払手数料は年報に計上欄がないので、実績側には含めない。
    """
    act, _s, gen = T.extract()
    v, std = {}, {}
    for y in YS:
        g = gen[y]
        v[y] = g["総給付費"] / 1000.0
        std[y] = (g["総給付費"] + g["特定入所"] + g["高額"]
                  + g["高額合算"]) / 1000.0
    return v, std, T.kasan_ratios(gen)


def dai9():
    """第9期の見込みが、どの年度の実績の水準にあるか。

    第9期の見込みは3か年ほぼ同値であり、自然体推計をそのまま採った形である。
    その水準が実績のどこにあたるかを見れば、基準年度が膨らんでいたかが分かる。
    """
    v, std, _rat = jisseki()
    furê = max(DAI9_MIKOMI.values()) / min(DAI9_MIKOMI.values()) - 1
    rows = [(y, v[y], std[y], DAI9_MIKOMI["令和6年度"] / std[y])
            for y in ("令和4年度", "令和5年度")]
    taihi = [(y, v[y], std[y], DAI9_MIKOMI[y], std[y] / DAI9_MIKOMI[y])
             for y in ("令和6年度", "令和7年度")]
    return {"割増率": std["令和5年度"] / v["令和5年度"], "振れ幅": furê,
            "水準": rows, "対比": taihi}


def bunkai():
    """見える化の令和8年度を、同じ種別の上振れと0になった種別の下振れに分ける。"""
    mie, _tot, _ku, _hoken = W.read()
    zero, keijo7, keijo8, zero7 = [], 0.0, 0.0, 0.0
    for kind in ("介護", "予防"):
        for name, d in mie[kind].items():
            k = d.get("給付費") or {}
            v7, v8 = k.get("令和7年度"), k.get("令和8年度")
            if v7 is None or v8 is None or v7 <= 0:
                continue
            if (v8 or 0) == 0:
                zero.append((kind, name, v7))
                zero7 += v7
            else:
                keijo7 += v7
                keijo8 += v8
    r7 = keijo7 + zero7
    g = SZ.r8_options()[0]["月報の実績"][0]
    return {"令和7年度": r7, "計上あり_R7": keijo7, "計上あり_R8": keijo8,
            "0落ち_R7": zero7, "上振れ": keijo8 / keijo7 - 1,
            "下振れ": -zero7 / r7, "比": keijo8 / r7,
            "月報比": g, "月報との差": keijo8 / r7 / g - 1,
            "0落ち明細": sorted(zero, key=lambda x: -x[2])}


def shiyo_sa():
    """見える化の仕様（標準）と、第1号被保険者数で一律に延ばした場合との差。

    再検証の結果、標準の算定を見える化の仕様に改めた。
    従前の置き方は ichiritsu() に残してある。
    """
    m = SZ.mieruka_spec()
    A = SZ.plan_rows()["集計"]
    i = SZ.ichiritsu()
    SZ.selfcheck()          # 仕様どおりに組めているかの検算
    return {"仕様": m, "標準": A["計"], "従前": i["年度別"],
            "差": m["3年計"] - i["3年計"],
            "率": m["3年計"] / i["3年計"] - 1}


def jogen():
    """国のガイドの上限1.1・下限0.9への適合。

    標準の算定（見える化の仕様）では、居住系・在宅の利用率は恒等的に変化0となる。
    施設だけは利用者数一定であるため、認定者数が減る分だけ利用率が上がる。
    """
    n8 = T.NINTEI_EST["令和8年度"] + T.NINTEI_2GO
    shisetsu = [n8 / (T.NINTEI_EST[y] + T.NINTEI_2GO) for y in T.PLAN_YEARS]
    return [
        ("年齢階層別認定率", [1.0] * 3, "変化を0として据え置いている"),
        ("サービス利用率（居住系・在宅）", [1.0] * 3,
         "認定者数・在宅対象者で延ばしているため恒等的に変化0"),
        ("サービス利用率（施設）", shisetsu,
         "**利用者数を一定としているため、認定者数が減る分だけ利用率が上がる。**"
         "見える化システムの仕様そのものによる"),
        ("1人1月あたり回（日）数", [1.0] * 3,
         "人数と量に同じ係数を当てているため恒等的に変化0"),
        ("1人1月あたり給付費", [1.0] * 3, "同上"),
    ]


# ---------------------------------------------------------------- 体裁

def new_doc():
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = JP_MIN
    st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), JP_MIN)
    return doc


def head(doc, text, level):
    h = doc.add_heading(text, level=level)
    for r in h.runs:
        r.font.name = JP_GO
        r._element.rPr.rFonts.set(qn("w:eastAsia"), JP_GO)
    return h


def _emph(p, text):
    """**…** を太字にして段落に流し込む。"""
    for i, part in enumerate(text.split("**")):
        if not part:
            continue
        r = p.add_run(part)
        r.font.name = JP_MIN
        r._element.rPr.rFonts.set(qn("w:eastAsia"), JP_MIN)
        r.bold = bool(i % 2)


def body(doc, *paras):
    for t in paras:
        _emph(doc.add_paragraph(), t)


def note(doc, text):
    p = doc.add_paragraph()
    _emph(p, text)
    for r in p.runs:
        r.font.size = Pt(9)
    return p


def table(doc, header_row, rows, right_from=99):
    t = doc.add_table(rows=1, cols=len(header_row))
    t.style = "Table Grid"
    for i, h in enumerate(header_row):
        c = t.rows[0].cells[i]
        c.text = ""
        p = c.paragraphs[0]
        r = p.add_run(h)
        r.font.name = JP_GO
        r._element.rPr.rFonts.set(qn("w:eastAsia"), JP_GO)
        r.bold = True
        r.font.size = Pt(9)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            p = cells[i].paragraphs[0]
            _emph(p, str(v))
            for r in p.runs:
                r.font.size = Pt(9)
            if i >= right_from:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    doc.add_paragraph()
    return t


# ---------------------------------------------------------------- Word

def build_docx(K, D9, B, S, J):
    yen = K["円換算"]
    doc = new_doc()
    head(doc, "第10期 サービス見込量算定の再検証", 0)
    body(doc, f"小野町　第10期介護保険事業計画　／　{ASOF_JP}")

    head(doc, "0　結論", 1)
    table(doc, ["#", "確かめたこと", "結果"], [
        ("1", "論点の重みの基準",
         f"**総給付費1％＝保険料の月額{K['1％']:.1f}円。**"
         "以下の論点はこの基準で重みを測った"),
        ("2", "第9期でも同じことが起きていなかったか",
         f"**起きていた。**第9期の見込みは3か年の振れ幅が"
         f"{D9['振れ幅'] * 100:.2f}％しかなく自然体推計そのものであり、"
         f"その水準は令和5年度の実績より"
         f"**{(D9['水準'][1][3] - 1) * 100:.1f}％高い。**"
         f"第9期の実績／見込みが令和6年度{D9['対比'][0][4] * 100:.1f}％・"
         f"令和7年度{D9['対比'][1][4] * 100:.1f}％となった理由がこれで説明できる"),
        ("3", "見える化の令和8年度の中身",
         f"**同じ種別での上振れ{B['上振れ'] * 100:+.2f}％と、"
         f"0になった種別の下振れ{B['下振れ'] * 100:+.2f}％**に分かれる。"
         f"差引{(B['比'] - 1) * 100:+.2f}％にしか見えないが、"
         f"**国保連月報の実績{B['月報比']:.4f}と比べると"
         f"{B['月報との差'] * 100:+.1f}％高い。**"
         "第9期と同じ幅であり、そのまま使えば同じ誤りを繰り返す"),
        ("4", "見える化の仕様との差",
         f"**従前は人数・量・給付費のすべてに第1号被保険者数の比を当てていたが、"
         f"見える化は施設を利用者数一定とし、居住系・在宅を認定者数で延ばす。**"
         f"差は3年計{S['差']:+,.0f}千円（{S['率'] * 100:+.2f}％）、"
         f"保険料の月額{yen(S['差']):+.0f}円。"
         "**本検証を受けて、標準の算定を見える化の仕様に改めた。**"
         "100円未満切上げ後の6,300円は変わらない"),
        ("5", "国のガイドの上限1.1・下限0.9",
         f"**{len(J)}項目とも適合。**最も振れるのは施設の利用率で"
         f"{max(abs(x - 1) for v in J for x in v[1]) * 100:.2f}％"
         "（利用者数一定の仕様そのものによるもの）"),
    ])

    head(doc, "1　論点の重みの基準 ― 総給付費1％＝月額58円", 1)
    body(doc,
         "**どの論点が意思決定に効き、どれが誤差の範囲かを分けるために、"
         "先に換算の基準を出します。**"
         "他町村の案件で用いている手順です。")
    table(doc, ["区分", "金額", "保険料への効き"], [
        ("総給付費（3年計）", f"{K['総給付費']:,.0f}千円", "―"),
        ("標準給付費（3年計）", f"{K['標準給付費']:,.0f}千円",
         f"総給付費の{K['割増率']:.4f}倍"),
        ("保険料基準額", f"月額{K['月額']:,.0f}円", "―"),
        ("**総給付費1％**", f"{K['総給付費'] * 0.01:,.0f}千円",
         f"**月額{K['1％']:.1f}円**"),
    ], right_from=1)
    note(doc, "※ 総給付費が動くと特定入所者介護サービス費等・高額介護サービス費等・"
              "高額医療合算介護サービス費等も比例して動くため、"
              "標準給付費は総給付費の変化の1.0781倍で動く。"
              "審査支払手数料は件数に比例するため、この換算では総給付費と同率で置いた。")

    head(doc, "2　第9期でも同じことが起きていた", 1)
    body(doc,
         "**第9期の見込みは、3か年の振れ幅が"
         f"{D9['振れ幅'] * 100:.2f}％しかありません。**"
         "見える化システムの自然体推計は、施設を利用者数一定、"
         "居住系・在宅の利用率と1人1月あたりの回（日）数の変化を0とするため、"
         "初期値のまま推計すれば見込量は3か年ほぼ同値になります。"
         "**第9期はこれをそのまま採ったものです。**")
    table(doc, ["年度", "第9期の標準給付費見込額"],
          [(y, f"{v:,}千円") for y, v in DAI9_MIKOMI.items()], right_from=1)
    body(doc,
         "**問題は、その水準が実績のどこにあったかです。**"
         "年報 様式3 の総給付費に特定入所者介護サービス費等・"
         "高額介護サービス費等・高額医療合算介護サービス費等の実額を加えて"
         "標準給付費を出し"
         f"（令和5年度で総給付費の{D9['割増率']:.4f}倍）、"
         "第9期の令和6年度の見込みと比べます。")
    table(doc, ["年度", "総給付費（年報 様式3）", "標準給付費",
                "第9期 令和6年度見込との比"],
          [(y, f"{v:,.0f}千円", f"{s:,.0f}千円", f"**{r:.4f}**")
           for y, v, s, r in D9["水準"]], right_from=1)
    body(doc,
         f"**第9期の見込みの水準は、令和5年度の実績より"
         f"{(D9['水準'][1][3] - 1) * 100:.1f}％高いところにありました。**"
         "その結果が次の対比です。")
    table(doc, ["年度", "総給付費（実績）", "標準給付費", "第9期の見込み",
                "実績／見込み"],
          [(y, f"{v:,.0f}千円", f"{s:,.0f}千円", f"{m:,}千円",
            f"**{r * 100:.1f}％**") for y, v, s, m, r in D9["対比"]],
          right_from=1)
    body(doc,
         "**令和6年度は実績が前年度から5.8％伸びたため見込みに近づきましたが"
         "（98.6％）、令和7年度に7.4％下がって91.1％まで開きました。**"
         "基準年度の水準が7％高ければ、実績がどう動いてもこの幅の乖離は残ります。",
         "**第10期で同じことを繰り返さないことが、本件の算定の出発点です。**")

    head(doc, "3　見える化の令和8年度は、そのままでは使えない", 1)
    body(doc,
         "**見える化ワークシートの令和8年度は、令和7年度の"
         f"{B['比']:.4f}倍に見えます。**"
         "ところが中身を分けると、上がっている種別と0になった種別が"
         "打ち消し合っているだけです。他町村の案件で用いている分け方です。")
    table(doc, ["区分", "金額", "効き"], [
        ("令和7年度の給付費（見える化）", f"{B['令和7年度']:,.0f}", "―"),
        ("　うち令和8年度にも計上のある種別",
         f"{B['計上あり_R7']:,.0f}", "―"),
        ("　うち令和8年度に0となった種別", f"{B['0落ち_R7']:,.0f}", "―"),
        ("**(a) 同じ種別での上振れ**",
         f"{B['計上あり_R8'] - B['計上あり_R7']:+,.0f}",
         f"**{B['上振れ'] * 100:+.2f}％**"),
        ("**(b) 0となった種別の下振れ**", f"{-B['0落ち_R7']:+,.0f}",
         f"**{B['下振れ'] * 100:+.2f}％**"),
        ("(a)＋(b)＝令和8年度／令和7年度", "―",
         f"{(B['比'] - 1) * 100:+.2f}％"),
    ], right_from=1)
    if B["0落ち明細"]:
        table(doc, ["区分", "0となった種別", "令和7年度の給付費"],
              [(k, n, f"{v:,.0f}") for k, n, v in B["0落ち明細"]],
              right_from=2)
    body(doc,
         f"**国保連月報でみると、令和8年度4〜6月提供分は前年の同じ月の"
         f"{B['月報比'] * 100:.1f}％です。**"
         f"見える化の令和8年度は、これより{B['月報との差'] * 100:.1f}％高く出ています。"
         f"**第9期の基準年度が実績より"
         f"{(D9['水準'][1][3] - 1) * 100:.1f}％高かったのとほぼ同じ幅です。**",
         "**本件が令和8年度を月報の実績で置いているのは、この一点によります。**")

    head(doc, "4　延ばし方を見える化の仕様にそろえた", 1)
    body(doc,
         "**従前は、人数・量・給付費のすべてに第1号被保険者数の比を"
         "当てていました。**"
         "総額を月報で裏づけるのが目的であり、サービスの構成と1人当たりの水準は"
         "令和7年度のまま動かさない置き方です。",
         "**一方、見える化システムの自然体推計は、サービスを3つに分けて"
         "別々に延ばします。**"
         "厚生労働省『自然体推計の計算過程確認シートのガイド』によります。"
         "**本検証を受けて、標準の算定をこちらに改めました。**")
    table(doc, ["区分", "見える化の置き方（標準に採用）", "従前の置き方"], [
        ("施設（介護老人福祉施設・老健・介護医療院・地域密着型特養）",
         "**計画期間中は直近年度の利用者数が一定**",
         "第1号被保険者数の比で逓減"),
        ("居住系（特定施設・認知症対応型共同生活介護・地域密着型特定施設）",
         "認定者数 × 利用率（利用率の変化は0）",
         "同上"),
        ("在宅（上記以外。短期入所を含む）",
         "（認定者数 − 施設 − 居住系）× 利用率（同）", "同上"),
    ])
    body(doc,
         "**認定者数は第1号被保険者数より速く減るため、両者は一致しません。**"
         "第1号被保険者数は令和8年度から令和11年度で0.89％減りますが、"
         "認定者数は1.88％減ります。")
    table(doc, ["年度", "見える化の仕様（標準）", "従前の置き方", "差"],
          [(y, f"{S['仕様']['年度別'][j]:,.0f}千円",
            f"{S['従前'][j]:,.0f}千円",
            f"{S['仕様']['年度別'][j] - S['従前'][j]:+,.0f}千円")
           for j, y in enumerate(T.PLAN_YEARS)]
          + [("**3年計**", f"**{S['仕様']['3年計']:,.0f}千円**",
              f"**{sum(S['従前']):,.0f}千円**",
              f"**{S['差']:+,.0f}千円**")], right_from=1)
    table(doc, ["区分", "見える化の仕様（3年計）"],
          [(k, f"{sum(v):,.0f}千円") for k, v in S["仕様"]["区分"].items()],
          right_from=1)
    body(doc,
         f"**差は{S['率'] * 100:+.2f}％、保険料の月額で{yen(S['差']):+.0f}円です。**"
         f"従前の置き方による{K['月額'] - yen(S['差']):,.0f}円に対し"
         f"**{K['月額']:,.0f}円**となります。"
         "**100円未満を切り上げた6,300円は変わりません。**",
         "**町が見える化システムに令和8年度の実績見込み値を入れ直して"
         "自然体推計を出し直すと、システムはこの値を返します。**"
         "素案・見込量の数値とシステムの出力が突き合うようになりました。")
    note(doc, "※ 区分別に組めていることは、サービス別を通さずに仕様を直に組んだ値と"
              "突き合わせて確かめている（ono_shizentai.selfcheck）。"
              "区分の割り当てを取り違えるとここで止まる。")

    head(doc, "5　国のガイドの上限1.1・下限0.9には適合している", 1)
    body(doc,
         "**国のガイドは、要介護認定率及びサービス利用率の自然体推計に、"
         "基準年度実績×1.1を上限、×0.9を下限とする制約を課しています。**"
         "当方の算定がこれに収まるかを確かめました。")
    table(doc, ["項目", "令和9年度", "令和10年度", "令和11年度", "判定", "内容"],
          [(nm, f"{v[0]:.4f}", f"{v[1]:.4f}", f"{v[2]:.4f}",
            "**適合**" if all(0.9 <= x <= 1.1 for x in v) else "**超過**", memo)
           for nm, v, memo in J], right_from=1)
    note(doc, "※ 基準年度（令和8年度）の値を1.0000としたときの比。")

    head(doc, "6　協議会にお諮りする事項", 1)
    table(doc, ["No.", "事項", "受託者の案"], [
        ("1", "見込量の延ばし方",
         f"**見える化の仕様に合わせた（9月25日）。**"
         f"町がシステムを動かしたときの値と一致する。"
         f"差は月額{yen(S['差']):+.0f}円で、切上げ後の6,300円は変わらない"),
        ("2", "令和8年度を月報の実績で置くこと",
         "**そのままとする。**第9期の乖離が基準年度の水準で説明できる以上、"
         "見える化の令和8年度をそのまま使うことはできない"),
        ("3", "施設を利用者数一定とすることの是非",
         "**妥当と考える。**町内の定員が制約として働いており、"
         "認定者数が減っても施設の利用者数がそれに比例して減るとは限らない"),
    ])

    head(doc, "7　町にご確認いただきたいこと", 1)
    table(doc, ["#", "内容", "優先度"], [
        ("1", "**第9期策定時の見える化の設定**"
              "（基準年度、認定率・利用率の伸びの窓）。"
              "第9期の乖離の原因を確定するために要する", "B"),
        ("2", "**【解消済】厚生労働省の保険者別配布シート**"
              "（標準の推計方法を当てはめた誤差。基本推計と包括推計の別）。"
              "**見える化システムの「推計パターン毎の乖離状況」ボタンから"
              "同じものが出力できた（令和8年9月25日受領）。**"
              "推計手法の妥当性検証（別冊）を参照", "―"),
        ("3", "令和8年度の残りの月の国保連月報（7月提供分以降）", "A"),
        ("4", "**包括的支援事業・任意事業費29,335,189円（年報 様式4 の令和6年度決算）に、"
              "保険者機能強化推進交付金及び介護保険保険者努力支援交付金を財源とする"
              "事業が含まれているか**", "A"),
    ])
    note(doc, "※ 4について。**交付金を財源とする事業は、"
              "第1号被保険者負担分相当額の算定基礎に含めない。**"
              "他町村の案件では、地域支援事業費にこれらを含めていたことが分かり、"
              "算定から除く修正を行っている。"
              "小野町の地域支援事業費は年報 様式4 の令和6年度決算を"
              "そのまま据え置いており、同じことが起きていないかの確認を要する。"
              "**含まれていれば保険料は過大に出る。**"
              "なお、調整交付金の算定基礎に加える総合事業費については、"
              "本件は介護予防・生活支援サービス事業（29,276,321円）と"
              "一般介護予防事業（5,304,726円）の合計で置いており、"
              "国の定義に沿っている。")
    note(doc, "※ 2は令和8年9月25日に解消した。"
              "**小野町では基本推計と包括推計の優劣が年度で逆転しており"
              "（令和6年度は包括、令和7年度は基本）、どちらとも言えない。**"
              "2年の平均では基本推計3.16％・包括推計3.57％で基本推計が"
              "わずかに優り、現行の設定を変える理由はない。"
              "**より重要なのは、標準の推計方法が小野町では全国の中位より下"
              "（1,267位・1,015位／約1,570保険者）であり、"
              "サービス別では誤差10％超が10件あることである。**"
              "当方が実績を基点として総額を月報で裏づける置き方を採っていることの"
              "裏づけになる（推計手法の妥当性検証）。")

    OUT.mkdir(parents=True, exist_ok=True)
    p = OUT / f"小野町_第10期_見込量算定の再検証_{ASOF}.docx"
    doc.save(p)
    print("出力:", p)


# ---------------------------------------------------------------- Excel

def style_head(ws, row):
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(row, c)
        cell.fill = HEAD
        cell.font = Font(bold=True, color="FFFFFF", size=9, name=JP_GO)
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
        cell.border = BORDER


def xbody(ws, first, wrap=()):
    for r in range(first + 1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            cell.font = Font(size=9, name=JP_MIN)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top",
                                       wrap_text=c in wrap)


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w


def notes(ws, lines):
    ws.append([])
    for t in lines:
        ws.append([t])
        ws.cell(ws.max_row, 1).font = Font(size=9, italic=True, name=JP_MIN)


def build_xlsx(K, D9, B, S, J):
    yen = K["円換算"]
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet("00_結論")
    ws.append([f"小野町 第10期 サービス見込量算定の再検証　（{ASOF_JP}）"])
    ws.cell(1, 1).font = Font(bold=True, size=13, name=JP_GO)
    ws.append(["他町村の案件で用いている検証の軸を小野町のデータに当てはめた。"
               "他団体の数値は用いていない。"])
    ws.append([])
    ws.append(["#", "確かめたこと", "結果", "保険料への効き"])
    hr = ws.max_row
    for no, nm, res, eff in [
        (1, "論点の重みの基準", f"総給付費1％＝月額{K['1％']:.1f}円", "―"),
        (2, "第9期の基準年度の水準",
         f"令和5年度の実績より{(D9['水準'][1][3] - 1) * 100:+.1f}％高い",
         f"第9期の乖離（令和7年度{D9['対比'][1][4] * 100:.1f}％）の原因"),
        (3, "見える化の令和8年度",
         f"上振れ{B['上振れ'] * 100:+.2f}％と0落ち{B['下振れ'] * 100:+.2f}％。"
         f"月報より{B['月報との差'] * 100:+.1f}％高い", "使わない判断の根拠"),
        (4, "見える化の仕様との差（標準に採用）",
         f"従前の置き方から3年計{S['差']:+,.0f}千円"
         f"（{S['率'] * 100:+.2f}％）",
         f"月額{yen(S['差']):+.0f}円"),
        (5, "上限1.1・下限0.9",
         f"{len(J)}項目とも適合（最大{max(abs(x - 1) for v in J for x in v[1]) * 100:.2f}％）",
         "―"),
    ]:
        ws.append([no, nm, res, eff])
    style_head(ws, hr)
    xbody(ws, hr, wrap=(2, 3, 4))
    widths(ws, [5, 30, 52, 34])

    ws = wb.create_sheet("01_第9期の検証")
    ws.append(["第9期の見込みは、どの年度の実績の水準にあったか"])
    ws.cell(1, 1).font = Font(bold=True, size=12, name=JP_GO)
    ws.append([])
    ws.append(["年度", "総給付費（年報 様式3・千円）", "標準給付費（千円）",
               "第9期の見込み（千円）", "実績／見込み"])
    hr = ws.max_row
    v, std, _rat = jisseki()
    mul = D9["割増率"]
    for y in YS:
        m = DAI9_MIKOMI.get(y)
        ws.append([y, round(v[y]), round(std[y]), m or None,
                   (std[y] / m) if m else None])
    style_head(ws, hr)
    for r in range(hr + 1, ws.max_row + 1):
        for c in (2, 3, 4):
            ws.cell(r, c).number_format = "#,##0"
        ws.cell(r, 5).number_format = "0.0%"
        x = ws.cell(r, 5).value
        if x is not None:
            ws.cell(r, 5).fill = NG if x < 0.95 else WARN if x < 0.99 else OK
    xbody(ws, hr)
    widths(ws, [12, 26, 24, 22, 14])
    notes(ws, [
        "※ 標準給付費＝総給付費＋特定入所者＋高額＋高額医療合算（いずれも年報 様式3 の実額）。"
        "審査支払手数料は年報に計上欄がないため実績側には含めない。",
        f"**※ 第9期の見込みは3か年の振れ幅が{D9['振れ幅'] * 100:.2f}％しかなく、"
        "自然体推計をそのまま採った形である。**",
        f"**※ その水準は令和5年度の実績より"
        f"{(D9['水準'][1][3] - 1) * 100:.1f}％高い。**"
        "基準年度の水準が高ければ、実績がどう動いてもこの幅の乖離は残る。",
    ])

    ws = wb.create_sheet("02_令和8年度の分解")
    ws.append(["見える化の令和8年度を、上振れと0落ちに分ける"])
    ws.cell(1, 1).font = Font(bold=True, size=12, name=JP_GO)
    ws.append([])
    ws.append(["区分", "金額（千円）", "率"])
    hr = ws.max_row
    for nm, amt, rate in [
        ("令和7年度の給付費（見える化）", B["令和7年度"], None),
        ("　うち令和8年度にも計上のある種別", B["計上あり_R7"], None),
        ("　うち令和8年度に0となった種別", B["0落ち_R7"], None),
        ("(a) 同じ種別での上振れ", B["計上あり_R8"] - B["計上あり_R7"],
         B["上振れ"]),
        ("(b) 0となった種別の下振れ", -B["0落ち_R7"], B["下振れ"]),
        ("令和8年度／令和7年度", None, B["比"] - 1),
        ("国保連月報の実績（令和8年度4〜6月提供分）", None, B["月報比"] - 1),
        ("見える化と月報の差", None, B["月報との差"]),
    ]:
        ws.append([nm, round(amt) if amt is not None else None, rate])
    style_head(ws, hr)
    for r in range(hr + 1, ws.max_row + 1):
        ws.cell(r, 2).number_format = "#,##0"
        ws.cell(r, 3).number_format = "+0.00%;-0.00%"
    ws.cell(ws.max_row, 3).fill = NG
    xbody(ws, hr)
    widths(ws, [40, 16, 12])
    ws.append([])
    ws.append(["令和8年度に0となった種別", "区分", "令和7年度の給付費（千円）"])
    hr2 = ws.max_row
    for k, n, x in B["0落ち明細"]:
        ws.append([n, k, round(x)])
    style_head(ws, hr2)
    for r in range(hr2 + 1, ws.max_row + 1):
        ws.cell(r, 3).number_format = "#,##0"

    ws = wb.create_sheet("03_見える化仕様との差")
    ws.append(["見える化システムの自然体推計の仕様で組み直した場合"])
    ws.cell(1, 1).font = Font(bold=True, size=12, name=JP_GO)
    ws.append([])
    ws.append(["年度", "見える化の仕様＝標準（千円）", "従前の置き方（千円）",
               "差（千円）"])
    hr = ws.max_row
    for j, y in enumerate(T.PLAN_YEARS):
        ws.append([y, round(S["仕様"]["年度別"][j]), round(S["従前"][j]),
                   round(S["仕様"]["年度別"][j] - S["従前"][j])])
    ws.append(["3年計", round(S["仕様"]["3年計"]), round(sum(S["従前"])),
               round(S["差"])])
    style_head(ws, hr)
    for r in range(hr + 1, ws.max_row + 1):
        for c in (2, 3, 4):
            ws.cell(r, c).number_format = "#,##0"
    ws.cell(ws.max_row, 4).fill = WARN
    xbody(ws, hr)
    widths(ws, [14, 24, 22, 16])
    ws.append([])
    ws.append(["区分", "見える化の仕様（3年計・千円）", "延ばし方"])
    hr2 = ws.max_row
    for k, v in S["仕様"]["区分"].items():
        ws.append([k, round(sum(v)),
                   {"施設": "計画期間中は直近年度の利用者数が一定",
                    "居住系": "認定者数 × 利用率（利用率の変化は0）",
                    "在宅": "（認定者数 − 施設 − 居住系）× 利用率（同）"}[k]])
    style_head(ws, hr2)
    for r in range(hr2 + 1, ws.max_row + 1):
        ws.cell(r, 2).number_format = "#,##0"
    xbody(ws, hr2, wrap=(3,))
    notes(ws, [
        f"**※ 差は{S['率'] * 100:+.2f}％、保険料の月額で{yen(S['差']):+.0f}円。**"
        f"従前{K['月額'] - yen(S['差']):,.0f}円に対し{K['月額']:,.0f}円。"
        "100円未満切上げ後の6,300円は変わらない。",
        "※ 従前は人数・量・給付費のすべてに同じ係数（第1号被保険者数の比）を"
        "当てていた。認定者数は第1号被保険者数より速く減るため一致しなかった。",
    ])

    ws = wb.create_sheet("04_上限下限への適合")
    ws.append(["国のガイドの上限1.1・下限0.9への適合"])
    ws.cell(1, 1).font = Font(bold=True, size=12, name=JP_GO)
    ws.append([])
    ws.append(["項目", "令和9年度", "令和10年度", "令和11年度", "判定", "内容"])
    hr = ws.max_row
    for nm, v, memo in J:
        ok = all(0.9 <= x <= 1.1 for x in v)
        ws.append([nm, v[0], v[1], v[2], "適合" if ok else "超過", memo])
    style_head(ws, hr)
    for r in range(hr + 1, ws.max_row + 1):
        for c in (2, 3, 4):
            ws.cell(r, c).number_format = "0.0000"
        ws.cell(r, 5).fill = OK if ws.cell(r, 5).value == "適合" else NG
    xbody(ws, hr, wrap=(6,))
    widths(ws, [26, 12, 12, 12, 8, 52])
    notes(ws, ["※ 基準年度（令和8年度）の値を1.0000としたときの比。",
               "※ 国のガイドは、要介護認定率及びサービス利用率の自然体推計に"
               "基準年度実績×1.1を上限、×0.9を下限とする制約を課している。"])

    p = OUT / f"小野町_第10期_見込量算定の再検証_{ASOF}.xlsx"
    wb.save(p)
    print("出力:", p)


def main():
    K, D9, B, S, J = kijun(), dai9(), bunkai(), shiyo_sa(), jogen()
    OUT.mkdir(parents=True, exist_ok=True)
    build_docx(K, D9, B, S, J)
    build_xlsx(K, D9, B, S, J)
    print(f"  総給付費1％＝月額{K['1％']:.1f}円／"
          f"第9期の基準年度は令和5年度実績の{D9['水準'][1][3]:.4f}倍／"
          f"見える化仕様との差 {S['差']:+,.0f}千円（月額{K['円換算'](S['差']):+.0f}円）")


if __name__ == "__main__":
    main()
