# -*- coding: utf-8 -*-
"""概要版の拡充（6頁 → 8頁程度へ）（令和8年10月5日）

08_作業順位 の順位19。委託仕様書8は概要版を
**A4版／コート紙／1色刷り／8頁程度／3,500部**と定めている。
令和8年10月4日の下書きは6頁（2,458字・図8点）であったため、
**内容を加えて8頁程度にする。**

加えるのは次の4つ。いずれも計画素案と既にある成果品から起こせる。

  ① 7つの柱ごとの主な取組（事業数・新規／拡充の数）
     ← 05_試算・管理シート/川崎町_第10期_施策事業統合体系表_R8.9.25b.xlsx
  ② 第9期から変わること（新しく始めること・やめること）
     ← 同 03_新規拡充一覧
  ③ 主な目安（KPI）　← 計画素案 10-4
  ④ ご相談先の一覧　← 計画素案 第10章

  出力　01_第10期_最新版成果品/川崎町_計画素案_概要版_R8.10.5.docx

⚠ **本文と図は `build_gaiyo_R8.10.4.py` から読む。**
  概要版のために新しい数値を作らない。同スクリプトの `P`（本文の並び）に
  加える形にしており、**本文を2か所に持っていない。**

⚠ 【町確認】【委員会協議】が残っている指標は概要版に載せない。
  住民の方にお配りするものであり、未確定の目標値を示さない。

⚠ 保険料は令和9年1月の策定委員会で決まる。下書きと同じく、
  確定前であることを明記したまま置く。
"""
import copy
import importlib.util
import os
import re
import shutil
import sys
import zipfile

import docx
import openpyxl
from docx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from fix_soan_v111 import clone_table, set_el            # noqa: E402


def load(name, fname):
    spec = importlib.util.spec_from_file_location(name,
                                                  os.path.join(HERE, fname))
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


G4 = load("gaiyo4", "build_gaiyo_R8.10.4.py")

SOAN = G4.SOAN
TAIKEI = "05_試算・管理シート/川崎町_第10期_施策事業統合体系表_R8.9.25b.xlsx"
OUT = "01_第10期_最新版成果品/川崎町_計画素案_概要版_R8.10.5.docx"
FIGDIR = G4.FIGDIR
EMU_IN = G4.EMU_IN

# 柱の番号と、概要版での呼び方（素案の章立てに合わせる）
HASHIRA = {
    "第１章": "柱1　健康づくりと介護予防",
    "第２章": "柱2　住まい・防災・防犯",
    "第３章": "柱3　移動・見守り・社会参加と家族への支援",
    "第４章": "柱4　地域包括ケアを支える事業",
    "第５章": "柱5　認知症とともに生きるまちづくり",
    "第６章": "柱6　介護サービス",
    "第７章": "柱7　計画の進め方と介護人材の確保",
}

# 概要版に載せる目安（KPI）。
# ⚠ 目標値が【町確認】【委員会協議】のものは載せない。
KPI = [
    ["外出を控えている高齢者の割合", "19.1％", "減らす"],
    ["認知症の相談窓口を知っている人の割合", "47.7％", "増やす"],
    ["自立支援型地域ケア会議での個別事例の話し合い", "―", "年12回以上"],
    ["生活支援コーディネーターの人数", "25名", "いまの人数を保つ"],
    ["計画と実績の差を毎年たしかめて公表すること", "していない",
     "毎年度行い公表する"],
    ["事業所からお話をうかがうこと", "していない", "年1回行う"],
]

KIKAN = [
    ["計画の名前",
     "川崎町高齢者保健福祉計画・第10期介護保険事業計画"],
    ["計画の期間", "令和9年度から令和11年度までの3年間"],
    ["中身", "高齢者福祉の進め方（高齢者保健福祉計画）と、"
             "介護保険の3年間の見通し・保険料（介護保険事業計画）を"
             "1つにまとめたものです。"
             "第10期からは、認知症の取組（認知症施策推進計画）も"
             "この中に位置づけます。"],
    ["ほかの計画との関係",
     "町の最上位の計画である「第6次川崎町長期総合計画」のもとで、"
     "地域福祉計画・健康増進計画・障害者計画などと足並みをそろえます。"],
    ["つくりかた",
     "住民のみなさまへのアンケート（令和8年6〜7月）、"
     "策定委員会でのご審議、認知症の本人とご家族からの意見の聴取、"
     "みなさまからのご意見の募集を経てつくります。"],
    ["計画の見直し", "3年ごとに見直します。"
                     "毎年度、進み具合をたしかめて公表します。"],
]

SODAN = [
    ["介護保険・高齢者福祉のご相談", "川崎町 保健福祉課"],
    ["介護・福祉・権利を守ることの総合相談", "川崎町地域包括支援センター"],
    ["もの忘れが気になるとき", "川崎町地域包括支援センター（毎月の相談日）"],
    ["認知症の方とご家族の集まり", "認知症カフェ「喫茶みかん」"],
    ["生活支援・通いの場のご案内", "川崎町社会福祉協議会"],
]


def yomu_taikei():
    """統合体系表から、柱ごとの事業数と新規・拡充の数を読む。"""
    wb = openpyxl.load_workbook(TAIKEI)
    ws = wb["01_統合体系表"]
    import collections
    n = collections.Counter()
    kubun = collections.defaultdict(collections.Counter)
    sho_order = []
    for r in range(2, ws.max_row + 1):
        sho = str(ws.cell(r, 1).value or "").strip()
        if not sho or sho == "章":
            continue
        k = str(ws.cell(r, 5).value or "").strip()
        key = sho[:3]
        if key not in sho_order:
            sho_order.append(key)
        # ⚠ 「再掲」は他の章に既に数えた事業である。
        #   02_集計シートも除いているため、ここでも数えない
        #   （数えると第5章が32件・合計147件になり、集計と食い違う）。
        if k == "再掲":
            continue
        n[key] += 1
        kubun[key][k] += 1
    rows = [["7つの柱", "事業の数", "うち新しく始める・広げるもの"]]
    for key in sho_order:
        name = HASHIRA.get(key)
        if name is None:
            raise SystemExit("柱の名が分からない：" + key)
        shin = kubun[key]["新規"] + kubun[key]["拡充"]
        rows.append([name, f"{n[key]}", f"{shin}"])
    rows.append(["合計", f"{sum(n.values())}",
                 f"{sum(k['新規'] + k['拡充'] for k in kubun.values())}"])
    # 02_集計シートと突き合わせる（数えかたが食い違っていないか）
    shu = {}
    ws2 = wb["02_集計"]
    for r in range(2, ws2.max_row + 1):
        t = str(ws2.cell(r, 1).value or "").strip()
        if t.startswith("第"):
            shu[t[:3]] = ws2.cell(r, 4).value
    chigai = [f"{k}（本表{n[k]}／集計{shu.get(k)}）"
              for k in sho_order if n[k] != shu.get(k)]
    if chigai:
        raise SystemExit("事業の数が 02_集計シートと合わない："
                         + "／".join(chigai))
    return rows


def yomu_shinki():
    """新規・拡充の一覧から、住民の方に関わりの深いものを拾う。"""
    wb = openpyxl.load_workbook(TAIKEI)
    ws = wb["03_新規拡充一覧"]
    head = [str(ws.cell(1, c).value or "") for c in range(1, 9)]
    i_jigyo = head.index("事業") + 1 if "事業" in head else 4
    i_kubun = head.index("区分") + 1 if "区分" in head else 5
    shin = []
    for r in range(2, ws.max_row + 1):
        k = str(ws.cell(r, i_kubun).value or "").strip()
        j = str(ws.cell(r, i_jigyo).value or "").strip()
        if k == "新規" and j:
            shin.append(re.sub(r"^（[^）]*）", "", j))
    return shin


def tsuika(P0, taikei, shinki):
    """下書きの本文に、加える4つを差し込む。

    ⚠ **本文は `build_gaiyo_R8.10.4.py` の P を正本とする。**
      ここでは差し込むだけで、同じ文を書き直さない。
    """
    P = []
    for item in P0:
        P.append(item)
        kind, v = item[0], item[1]
        # ⓪ 冒頭のあとに、計画の期間と位置づけ
        if kind == "本文" and v.startswith("この計画は、65歳以上の方の"):
            P.append(("T2", [["この計画のこと", "内容"]] + KIKAN, [30, 70]))
        # ① 柱の説明のあとに、柱ごとの表
        if kind == "本文" and v.startswith("柱1は健康づくりと介護予防"):
            P.append(("T3", taikei, [52, 20, 28]))
        # ② 「いちばんの課題です。」のあとに、第9期から変わること
        if kind == "本文" and v.endswith("この計画のいちばんの課題です。"):
            P.append(("見出し", "＃　第9期から変わること"))
            P.append(("本文",
                      "第10期では、新しく始めることが"
                      f"{len(shinki)}件あります。おもなものをお示しします。"))
            P.append(("T2", [["新しく始めること", "どういうものか"]]
                      + SHINKI_SETSUMEI, [38, 62]))
            P.append(("本文",
                      "第9期にあって第10期で記載しない事業は2件です"
                      "（高齢者外出タクシー利用助成事業ほか）。"
                      "いずれも利用の実績と他の事業との重なりを見たうえで"
                      "整理したものです。"))
        # ③ 計画の進め方のあとに、主な目安（KPI）
        if kind == "本文" and v.startswith("計画は、毎年度その進み具合を"):
            P.append(("本文", "おもな目安は次のとおりです。"))
            P.append(("T3", [["何をみるか", "いまのところ", "3年後の目安"]]
                      + KPI, [46, 22, 32]))
        # ④ ご相談先の囲みのあとに、相談先の一覧
        if kind == "囲み" and v.startswith("ご相談は"):
            P.append(("T2", [["こんなとき", "ご相談先"]] + SODAN, [44, 56]))
    return bangou(P)


KAN = "０１２３４５６７８９"


def bangou(P):
    """見出しの番号を、出てくる順に振り直す。

    ⚠ 下書きの見出しは「１　…」から「８　…」まで番号を持っている。
      **間に見出しを差し込むと番号が重なる**ため、組み立ての前に
      すべて振り直す（差し込んだものは「＃」と書いておく）。
    """
    out, i = [], 0
    for item in P:
        if item[0] == "見出し":
            i += 1
            na = re.sub(r"^[0-9０-９＃]+[　 ]*", "", item[1])
            out.append(("見出し", f"{KAN[i]}　{na}"))
        else:
            out.append(item)
    return out


# ⚠ 新しく始めることの説明は**当方の書き起こし**である。
#   事業の内容のご確認（確認事項No.164）に含まれる。
SHINKI_SETSUMEI = [
    ["家族介護教室",
     "介護をしているご家族に、介護の仕方や体の使い方をお伝えする教室です。"],
    ["ご家族の相談・交流の場",
     "同じように介護をしている方どうしが話し合える場をつくります。"],
    ["短期入所（レスパイト）の利用のお手伝い",
     "介護をしているご家族が休めるよう、短期入所の利用をお手伝いします。"],
    ["仕事と介護の両立のお手伝い",
     "介護のために仕事を辞めずにすむよう、"
     "相談の窓口と事業所への働きかけを行います。"],
    ["ヤングケアラー・ダブルケアへの対応",
     "若い世代の介護、子育てと介護が重なるご家庭に気づき、"
     "必要な支援につなぎます。"],
    ["認知症の本人ミーティング",
     "認知症の本人が集まって語り合い、"
     "そのお話を町の取組に活かします。"],
    ["認知症の日・認知症月間の行事",
     "毎年9月21日とその前後に、認知症について知っていただく"
     "行事を行います。"],
    ["計画と実績の差をたしかめて公表すること",
     "計画した量と実際の利用の差を毎年たしかめ、結果を公表します。"],
]


def main():
    for p in (SOAN, TAIKEI):
        if not os.path.exists(p):
            raise SystemExit("入力がない：" + p)
    # 図は下書きと同じものを使う（作り直して同じになることを確かめる）
    made = [fn() for _, fn in G4.FIGS]
    taikei = yomu_taikei()
    shinki = yomu_shinki()
    P = tsuika(G4.P, taikei, shinki)

    shutil.copy(SOAN, OUT)
    doc = docx.Document(OUT)
    body = doc.element.body
    tpl = {}
    for el in body.iterchildren():
        if el.tag == qn("w:p"):
            s = "".join(n.text or "" for n in el.iter(qn("w:t"))).strip()
            if el.findall(".//" + qn("w:drawing")):
                if "図" not in tpl:
                    tpl["図"] = copy.deepcopy(el)
                continue
            if "表題" not in tpl and s.startswith("第 1 章"):
                tpl["表題"] = copy.deepcopy(el)
            if "見出し" not in tpl and s.startswith("▌"):
                tpl["見出し"] = copy.deepcopy(el)
            if "本文" not in tpl and s.startswith("令和7年度から令和8年度への"):
                tpl["本文"] = copy.deepcopy(el)
            if "囲み" not in tpl and s.startswith("⚠ 宮城県自身の評価"):
                tpl["囲み"] = copy.deepcopy(el)
        elif el.tag == qn("w:tbl"):
            n = len(el.findall(qn("w:tr"))[0].findall(qn("w:tc")))
            tpl.setdefault(f"T{n}", copy.deepcopy(el))
    for k in ("表題", "見出し", "本文", "囲み", "図", "T2", "T3"):
        if k not in tpl:
            raise SystemExit("雛形が見つからない：" + k)

    sectPr = body.find(qn("w:sectPr"))
    for el in list(body.iterchildren()):
        if el is not sectPr:
            body.remove(el)

    def add(el):
        if sectPr is not None:
            sectPr.addprevious(el)
        else:
            body.append(el)

    n_fig = n_tbl = 0
    for item in P:
        kind, v = item[0], item[1]
        if kind == "図":
            doc.add_picture(os.path.join(FIGDIR, v),
                            width=G4.Emu(int(5.6 * EMU_IN)))
            p = doc.paragraphs[-1]._p
            p.getparent().remove(p)
            pPr = tpl["図"].find(qn("w:pPr"))
            if pPr is not None:
                old = p.find(qn("w:pPr"))
                if old is not None:
                    p.remove(old)
                p.insert(0, copy.deepcopy(pPr))
            tr = tpl["図"].find(qn("w:r"))
            rPr = tr.find(qn("w:rPr")) if tr is not None else None
            if rPr is not None:
                for r in p.findall(qn("w:r")):
                    if r.find(qn("w:rPr")) is None:
                        r.insert(0, copy.deepcopy(rPr))
            add(p)
            n_fig += 1
        elif kind.startswith("T"):
            add(clone_table(tpl[kind], v, item[2]))
            e = copy.deepcopy(tpl["本文"])
            set_el(e, "")
            add(e)
            n_tbl += 1
        else:
            el = copy.deepcopy(tpl[kind])
            set_el(el, v)
            add(el)
    doc.save(OUT)
    G4.strip_unused_media(OUT)

    # ══════════════════════════ 自己点検
    d2 = docx.Document(OUT)
    text = "\n".join(q.text for q in d2.paragraphs)
    tbl_text = "\n".join(c.text for t in d2.tables
                         for r in t.rows for c in r.cells)
    zen = text + "\n" + tbl_text
    ng = []
    z = zipfile.ZipFile(OUT)
    n_media = sum(1 for n in z.namelist() if n.startswith("word/media"))
    n_img = len(re.findall(r"<w:drawing>",
                           z.read("word/document.xml").decode("utf-8")))
    if n_img != len(G4.FIGS):
        ng.append(f"図が{n_img}点（{len(G4.FIGS)}点のはず）")
    if n_media != len(G4.FIGS):
        ng.append(f"使っていない画像が残っている（{n_media}点）")
    if len(d2.tables) != n_tbl:
        ng.append(f"表が{len(d2.tables)}件（{n_tbl}件のはず）")
    for w in ("受託者", "当社", "ビズアップ", "確認事項No",
              "【町確認】", "【委員会協議】", "**", "__"):
        if w in zen:
            ng.append(f"概要版に「{w}」が残っている")
    for w in ("大雪", "東川", "東神楽", "上川", "金ヶ崎", "金ケ崎", "川崎市"):
        if w in zen:
            ng.append(f"他団体の名称が入っている：{w}")
    # 下書きの本文がすべて引き継がれていること
    for item in G4.P:
        if item[0] in ("図",):
            continue
        s = item[1]
        if item[0] == "見出し":
            # 番号は振り直すため、番号を除いた見出し名で見る
            s = re.sub(r"^[0-9０-９＃]+[　 ]*", "", s)
        if s and s.split("\n")[0][:20] not in zen:
            ng.append(f"下書きの本文が落ちている：{item[1][:24]}")
    # 見出しの番号が1から順に重なりなく振られていること
    mi = [q.text.strip() for q in d2.paragraphs
          if re.match(r"^[０-９]　", q.text.strip())]
    want = [KAN[i] for i in range(1, len(mi) + 1)]
    if [m[0] for m in mi] != want:
        ng.append(f"見出しの番号が順でない：{[m[0] for m in mi]}")
    # 加えた4つが入っていること
    for name, w in (("柱ごとの表", "柱7　計画の進め方と介護人材の確保"),
                    ("第9期から変わること", "本人ミーティング"),
                    ("主な目安", "3年後の目安"),
                    ("ご相談先", "川崎町社会福祉協議会")):
        if w not in zen:
            ng.append(f"加えるはずの{name}が入っていない")
    # 事業数が統合体系表と合っていること
    goukei = taikei[-1][1]
    if goukei not in zen:
        ng.append(f"事業の数の合計{goukei}が入っていない")
    # 保険料が確定前である旨
    if "まだ確定していません" not in zen:
        ng.append("保険料が確定前である旨の断りがない")

    n_char = sum(len(q.text) for q in d2.paragraphs) + len(tbl_text)
    print("概要版の拡充（6頁 → 8頁程度へ）")
    print("保存：", OUT)
    print(f"  段落 {len(d2.paragraphs)}／図 {n_img}／表 {len(d2.tables)}"
          f"／字数 {n_char:,}字")
    print(f"  下書き（6頁・2,458字）に、柱ごとの表・第9期から変わること・"
          f"主な目安{len(KPI)}件・ご相談先{len(SODAN)}件を加えた")
    print(f"  統合体系表から読んだ事業の数 {goukei}件"
          f"（うち新規・拡充 {taikei[-1][2]}件）")
    print("  ── 自己点検")
    if ng:
        for m in ng:
            print("   ×", m)
        sys.exit(1)
    print("   ○ 下書きの本文はすべて引き継がれている"
          "（本文を2か所に持っていない）")
    print("   ○ 【町確認】【委員会協議】の残っている目標値は載せていない")
    print("   ○ 受託者を主語とする語・他団体の名称なし")
    print("  ⚠ 頁数は組版によります。Word で開いてご確認ください。"
          "当方の環境（游ゴシックがない）で PDF にすると7頁です。"
          "仕様書8は「8頁程度」と定めており、"
          "冊子は4の倍数の頁で組むため8頁が収まりのよい形です。"
          "Word でご覧になって不足する場合は、"
          "ことばの説明と相談先を広げて調整いたします。")
    print("  ⚠ 入稿は第3回策定委員会（令和9年1月）の後になります。"
          "保険料の額が確定してから差し替えます。")


if __name__ == "__main__":
    main()
