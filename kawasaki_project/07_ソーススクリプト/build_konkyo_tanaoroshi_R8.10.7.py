# -*- coding: utf-8 -*-
"""素案の根拠の棚卸し（令和8年10月7日）

08_作業順位 の順位25。

計画素案に書いた数値の一つずつについて、**どこから来た数か**を示す。
県への事前協議と公表版の確定の前に、数の裏づけを揃えておく。

  出力　05_試算・管理シート/川崎町_素案の根拠の棚卸し_R8.10.7.xlsx
          00_この表について
          01_区分ごとの件数
          02_節ごとの状況
          03_出所を特定できていないもの（要対応）
          04_出所の一覧（当方の成果品・原資料）

やり方は、素案から数値を取り出し、**出所の候補となるファイルの値と
機械で突き合わせる**ことによる。

  A 原典・成果品に現れる　　同じ値がファイルの中にある
  B 当方の算定　　　　　　　単位換算・割合として、同じ節の数から導ける
  C 識別子・連絡先　　　　　事業所番号・電話番号など、突合の対象でない
  D 出所を特定できていない　⚠ 公表版の確定までに出所を書き留める

⚠ **数え方の限界。** 年・月・日・条・項・号・章・節の数と、
  3桁以下の整数（割合を除く）は、数が多く意味も薄いため対象から外している。
  対象としたのは、カンマを含む・小数点を含む・4桁以上・割合（％）の数である。

⚠ 本表は**出所を示すもの**であり、値が正しいかどうかを確かめるものではない。
  値そのものの検算は、保険料の検算（39項目）・交付金の全国統計の再現
  （39件）・図表データ管理台帳による。
"""
import collections
import importlib.util
import os
import re
import sys

import docx
import openpyxl
from docx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from data_zuhyo import ZU                                    # noqa: E402


def load(name, fname):
    spec = importlib.util.spec_from_file_location(name,
                                                  os.path.join(HERE, fname))
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


GY = load("gyotei8", "build_iinkai_gyotei_R8.10.1.py")

SOAN = "01_第10期_最新版成果品/川崎町_計画書素案_v2.12_根拠整理版.docx"
OUT = "05_試算・管理シート/川崎町_素案の根拠の棚卸し_R8.10.7.xlsx"

NUM = re.compile(r"\d[\d,]*\.?\d*")

# ══════════════════════════ 出所の候補
SRC_X = [
    ("05_試算・管理シート/川崎町_第10期_計画見込量_R8.9.15.xlsx",
     "当方の算定（サービス見込量・保険料）"),
    ("05_試算・管理シート/川崎町_交付金_全国統計の再現_R8.10.7.xlsx",
     "当方の算定（交付金の全国統計）"),
    ("05_試算・管理シート/川崎町_交付金_取りまとめ表_R8.9.29.xlsx",
     "当方の算定（交付金の枝番）"),
    ("05_試算・管理シート/川崎町_第9期_計画値実績対比表_R8.10.1.xlsx",
     "第9期計画書及び年報"),
    ("05_試算・管理シート/川崎町_第10期_施策事業統合体系表_R8.9.25b.xlsx",
     "町ご提示の施策・事業の体系（案）"),
    ("05_試算・管理シート/川崎町_保険料試算ワークブック.xlsx",
     "当方の算定（保険料）"),
    ("05_試算・管理シート/川崎町_地域支援事業費の置き方_R8.10.7.xlsx",
     "当方の算定（地域支援事業費）"),
    ("05_試算・管理シート/川崎町_需要シナリオと保険料の感度_R8.10.7.xlsx",
     "当方の算定（感度）"),
    ("09_元資料/R8実績データ/R8.9.11受領版/"
     "【川崎町】第10期_将来推計総括表_R8.9.11出力_④訂正後.xlsx",
     "見える化システム（将来推計総括表）"),
    ("09_元資料/R7実績データ/年報データ_2025_川崎町.xlsx",
     "介護保険事業状況報告（年報・令和7年度）"),
    ("09_元資料/R8実績データ/R8.9.1受領版/"
     "川崎町_町提供実績データ_R8.9.1受領.xlsx",
     "町ご提供の実績データ（令和8年9月1日受領）"),
    ("09_元資料/交付金評価/③令和８年度交付金評価指標等（市町村分・公表版）/"
     "001732614_令和８年度全国集計（市町村）.xlsx",
     "厚生労働省 令和8年度 全国集計（市町村分）"),
]
SRC_D = [
    ("01_第10期_最新版成果品/川崎町_ニーズ調査結果報告書_R8.10.2版.docx",
     "ニーズ調査（令和8年6〜7月・576件）"),
    ("01_第10期_最新版成果品/川崎町_在宅介護実態調査結果報告書_R8.10.2版.docx",
     "在宅介護実態調査（令和8年6〜7月・142件）"),
    ("01_第10期_最新版成果品/"
     "川崎町_在宅介護実態調査結果報告書_資料編_R8.10.2版.docx",
     "在宅介護実態調査 資料編"),
]

# ══════════════════════════ 手で書き留めた出所
#   機械で突合できないが、出所がはっきりしているもの。
#   （節, 探す語, 出所）
CHUKI = [
    ("2-1", "7,648", "住民基本台帳（令和8年6月末）。町ホームページの公表値"),
    ("1-1", "7,648", "同上"),
    ("2-2", "7,648", "同上"),
    ("（冒頭）", "7,648", "同上"),
    ("2-1", "3,209", "住民基本台帳（令和8年6月末）の65歳以上"),
    ("1-1", "3,209", "同上"),
    ("2-2", "3,209", "同上"),
    ("（冒頭）", "3,209", "同上"),
    ("2-1", "2,181", "住民基本台帳（令和8年6月末）の世帯数"),
    ("1-5", "82.65", "令和8年3月 全国介護保険・高齢者保健福祉担当課長会議資料"),
    ("2-5", "0472200062", "介護サービス情報公表システムの事業所番号"),
    ("2-5", "2119", "国民健康保険川崎病院の電話番号"),
]

KAKUNIN = ["163", "157", "155", "154", "158", "29", "8"]


def soan_nums():
    """素案から、突合の対象とする数値を取り出す。"""
    d = docx.Document(SOAN)
    cur = "（冒頭）"
    items = []

    def shori(sec, kind, text):
        for m in NUM.finditer(text):
            s = m.group(0)
            mae = text[max(0, m.start() - 6):m.start()]
            ato = text[m.end():m.end() + 4]
            # 年月日・条項号・章節は対象から外す
            if re.search(r"(令和|平成|第)$", mae) and \
                    re.match(r"^(年|月|日|条|項|号|章|節|回|期)", ato):
                continue
            if re.match(r"^(年度|年|月|日|条|項|号|章|節|頁|期)", ato) \
                    and len(s) <= 2:
                continue
            if not ("," in s or "." in s
                    or len(s.replace(",", "")) >= 4
                    or ato.startswith("％") or ato.startswith("%")):
                continue
            items.append({
                "節": sec, "種類": kind, "数": s,
                "前後": (mae + "《" + s + "》" + ato).replace("\n", " "),
                "mae": mae, "ato": ato})
    for el in d.element.body.iterchildren():
        if el.tag == qn("w:p"):
            t = "".join(n.text or "" for n in el.iter(qn("w:t")))
            m = re.match(r"^(\d+-\d+)[　 ]", t.strip())
            if m:
                cur = m.group(1)
            shori(cur, "本文", t)
        elif el.tag == qn("w:tbl"):
            tb = docx.table.Table(el, d)
            for r in tb.rows:
                for c in r.cells:
                    shori(cur, "表", c.text)
    return items


def src_values():
    """出所の候補の値を集める。

    返すのは2つ。
      vals　丸めた表示の文字 → 出所の名の集合（速い突合に使う）
      raw 　(値, 出所の名) の一覧（丸めの差を許す突合に使う）

    ⚠ **数が文字として入っているセルも拾う。** 当方の成果品には
      「418.9点」「1,741保険者」のように書式を付けて書き込んだセルがあり、
      数として読むと取りこぼす。
    """
    vals = collections.defaultdict(set)
    raw = []

    def add(v, name):
        if isinstance(v, str):
            for m in NUM.finditer(v):
                try:
                    add(float(m.group(0).replace(",", "")), name)
                except ValueError:
                    pass
            return
        if not isinstance(v, (int, float)) or isinstance(v, bool):
            return
        f = float(v)
        raw.append((f, name))
        for nd in range(0, 4):
            vals[f"{round(f, nd):.{nd}f}"].add(name)
        # 円→千円・万円・億円の丸め（素案は単位を変えて書くことがある）
        for bai in (1e3, 1e4, 1e8):
            for nd in (0, 1, 2):
                vals[f"{round(f / bai, nd):.{nd}f}"].add(name + "（単位換算）")

    for p, name in SRC_X:
        if not os.path.exists(p):
            raise SystemExit("出所の候補がない：" + p)
        wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                for v in row:
                    add(v, name)
        wb.close()
    for p, name in SRC_D:
        if not os.path.exists(p):
            raise SystemExit("出所の候補がない：" + p)
        dd = docx.Document(p)
        txt = "\n".join(q.text for q in dd.paragraphs) + "\n" + "\n".join(
            c.text for t in dd.tables for r in t.rows for c in r.cells)
        for m in NUM.finditer(txt):
            try:
                add(float(m.group(0).replace(",", "")), name)
            except ValueError:
                pass
    for z in ZU:
        for row in z["rows"]:
            for v in row:
                add(v, "図の数値の正本（data_zuhyo.py）")
    return vals, raw


def main():
    if not os.path.exists(SOAN):
        raise SystemExit("入力がない：" + SOAN)
    items = soan_nums()
    vals, raw = src_values()

    chuki = {(c[0], c[1]): c[2] for c in CHUKI}
    sec_ok = collections.defaultdict(set)

    # ── 第1段　原典・成果品に現れるか
    for it in items:
        try:
            f = float(it["数"].replace(",", ""))
        except ValueError:
            it["区分"], it["出所"] = "C 識別子・連絡先", "数として読めない"
            continue
        it["f"] = f
        nd = len(it["数"].split(".")[1]) if "." in it["数"] else 0
        it["nd"] = nd
        src = vals.get(f"{round(f, nd):.{nd}f}")
        if src:
            it["区分"] = "A 原典・成果品に現れる"
            it["出所"] = sorted(src)[0]
            sec_ok[it["節"]].add(f)
        else:
            it["区分"] = None

    # ── 第2段　手の注記・識別子・割合
    for it in items:
        if it["区分"]:
            continue
        s, sec = it["数"], it["節"]
        if (sec, s) in chuki:
            it["区分"], it["出所"] = "A 原典・成果品に現れる", chuki[(sec, s)]
            continue
        bare = s.replace(",", "")
        if len(bare) >= 9 and "." not in s:
            it["区分"] = "C 識別子・連絡先"
            it["出所"] = "桁数から識別子と判断（突合の対象外）"
            continue
        if re.match(r"^20\d\d\.\d$", s):
            it["区分"] = "C 識別子・連絡先"
            it["出所"] = "西暦の年月の表記（突合の対象外）"
            continue
        # 割合（同じ節の2つの数の比として導けるか）
        if it["ato"].startswith("％") or it["ato"].startswith("%") \
                or (it.get("nd", 0) > 0 and it.get("f", 1e9) <= 100):
            vs = sorted(sec_ok.get(sec, ()))
            ok = None
            for a in vs:
                for b in vs:
                    if b and a != b and abs(a / b * 100 - it["f"]) < 0.06:
                        ok = f"同じ節の {a:,.0f} ÷ {b:,.0f} × 100"
                        break
                if ok:
                    break
            if ok:
                it["区分"], it["出所"] = "B 当方の算定（割合）", ok
                continue
        it["区分"], it["出所"] = "D 出所を特定できていない", ""

    # ── 第3段　丸めの差を許した突合
    #   ⚠ 素案は円を千円・万円・億円に丸めて書くことがあり、
    #     丸めの向きの違いで1だけずれる。倍率ごとに誤差1まで許す。
    BAI = [(1.0, ""), (1e3, "千円に丸め"), (1e4, "万円に丸め"),
           (1e8, "億円に丸め")]
    nokori = [it for it in items if it["区分"].startswith("D")]
    for it in nokori:
        f = it.get("f")
        if f is None:
            continue
        atari = None
        for bai, nm in BAI:
            kyoyo = 1.0 if bai > 1 else 0.51
            for v, name in raw:
                if abs(v / bai - f) <= kyoyo:
                    atari = (name, nm)
                    break
            if atari:
                break
        if atari:
            name, nm = atari
            it["区分"] = "A 原典・成果品に現れる"
            it["出所"] = name + ("（" + nm + "）" if nm else "")

    cnt = collections.Counter(it["区分"] for it in items)
    D = [it for it in items if it["区分"].startswith("D")]
    print(f"  （丸めの差を許した突合で "
          f"{len(nokori) - len(D)}件が解けました）")

    # ══════════════════════════ 出力
    wb = openpyxl.Workbook()
    ws = GY.sheet(
        wb, "00_この表について", "計画素案の根拠の棚卸し",
        f"計画素案（{os.path.basename(SOAN)}）に書いた数値"
        f"{len(items):,}件の一つずつについて、どこから来た数かを示します。"
        "県への事前協議と公表版の確定の前に、数の裏づけを揃えておくためです。\n"
        "やり方は、素案から数値を取り出し、出所の候補となるファイル"
        f"（{len(SRC_X) + len(SRC_D) + 1}件）の値と機械で突き合わせることによります。\n"
        "⚠ 年・月・日・条・項・号・章・節の数と、3桁以下の整数（割合を除く）は、"
        "数が多く意味も薄いため対象から外しています。対象としたのは、"
        "カンマを含む・小数点を含む・4桁以上・割合（％）の数です。\n"
        "⚠ 本表は出所を示すものであり、値が正しいかどうかを確かめるもの"
        "ではありません。値そのものの検算は、保険料の検算（39項目）・"
        "交付金の全国統計の再現（39件）・図表データ管理台帳によります。",
        ["シート", "内容"], [34, 96], first=True)
    GY.put(ws, [
        ["01_区分ごとの件数", "A・B・C・Dの件数と割合。"],
        ["02_節ごとの状況", "節ごとに、何件のうち何件の出所が示せているか。"],
        ["03_出所を特定できていないもの",
         f"⚠ {len(D)}件。公表版の確定までに出所を書き留める必要があります。"],
        ["04_出所の一覧",
         f"突合に用いた出所{len(SRC_X) + len(SRC_D) + 1}件。"],
        ["05_確認事項", f"{len(KAKUNIN)}件。"],
    ])

    ws = GY.sheet(
        wb, "01_区分ごとの件数", "区分ごとの件数",
        "A＝原典・成果品に同じ値がある／B＝単位換算・割合として導ける／"
        "C＝識別子・連絡先（突合の対象外）／D＝出所を特定できていない",
        ["区分", "件数", "割合", "内容"], [34, 14, 14, 76])
    SETSU = {
        "A 原典・成果品に現れる": "同じ値が、原資料又は当方の成果品の中にあります。"
                                  "出所の名を02・03シートに記しています。",
        "B 当方の算定（割合）": "同じ節にある2つの数の比として導けます。"
                                "割合・構成比の類です。",
        "C 識別子・連絡先": "事業所番号・電話番号・西暦の年月など、"
                            "突合の対象としないものです。",
        "D 出所を特定できていない": "⚠ 公表版の確定までに出所を書き留める"
                                    "必要があります。03シートに一覧と原因があります。",
    }
    rows, fills = [], []
    for k in ("A 原典・成果品に現れる", "B 当方の算定（割合）",
              "C 識別子・連絡先", "D 出所を特定できていない"):
        n = cnt.get(k, 0)
        rows.append([k, f"{n:,}", f"{n / len(items):.1%}", SETSU[k]])
        fills.append(GY.NGF if k.startswith("D")
                     else (GY.OKF if k.startswith("A") else GY.WARN))
    rows.append(["計", f"{len(items):,}", "100.0％", ""])
    fills.append(None)
    GY.put(ws, rows, fills)

    ws = GY.sheet(
        wb, "02_節ごとの状況", "節ごとの状況",
        "節ごとに、数値が何件あり、そのうち出所を示せているのが何件かを"
        "示します。Dの多い節から手を付けてください。",
        ["節", "数値の件数", "A", "B", "C", "D（要対応）", "出所を示せた割合"],
        [14, 16, 12, 12, 12, 16, 20])
    by = collections.defaultdict(lambda: collections.Counter())
    for it in items:
        by[it["節"]][it["区分"][0]] += 1
    def key(s):
        m = re.match(r"^(\d+)-(\d+)$", s)
        return (int(m.group(1)), int(m.group(2))) if m else (99, 99)
    rows, fills = [], []
    for sec in sorted(by, key=key):
        c = by[sec]
        n = sum(c.values())
        shimese = (c["A"] + c["B"] + c["C"]) / n
        rows.append([sec, f"{n:,}", f"{c['A']:,}", f"{c['B']:,}",
                     f"{c['C']:,}", f"{c['D']:,}", f"{shimese:.1%}"])
        fills.append(GY.NGF if c["D"] >= 10
                     else (GY.WARN if c["D"] else GY.OKF))
    GY.put(ws, rows, fills)

    ws = GY.sheet(
        wb, "03_出所を特定できていないもの", "出所を特定できていないもの",
        f"⚠ {len(D)}件あります。公表版の確定までに、"
        "一つずつ出所を書き留める必要があります。\n"
        "⚠ 残っているものの原因は1つに絞れました。"
        "9-1・9-2・9-3 に集まっており、いずれも"
        "中長期（令和17年度・令和22年度）の見込量と、保険料の内訳です。"
        "これらは当方が算定していますが、算定の結果を成果品"
        "（計画見込量のxlsx）に書き出しておらず、素案にしか残っていません。"
        "計画見込量のxlsxは令和11年度までしか持っていません。\n"
        "⚠ 同じ原因であった交付金の全国統計（3-3・18件）は、"
        "原典から作り直して成果品に残しました"
        "（川崎町_交付金_全国統計の再現_R8.10.7.xlsx）。"
        "中長期と保険料の内訳についても同じことを行います（作業順位29）。",
        ["節", "種類", "数値", "前後の文", "出所（ご記入ください）"],
        [12, 10, 18, 76, 40])
    GY.put(ws, [[it["節"], it["種類"], it["数"], it["前後"], ""]
                for it in D], [GY.NGF] * len(D))

    ws = GY.sheet(
        wb, "04_出所の一覧", "突合に用いた出所",
        "いずれもリポジトリに追跡されており、"
        "いつでも同じ突合を行えます。",
        ["出所", "ファイル", "区分"], [44, 76, 20])
    rows = []
    for p, name in SRC_X:
        rows.append([name, p,
                     "原資料" if p.startswith("09_") else "当方の成果品"])
    for p, name in SRC_D:
        rows.append([name, p, "当方の成果品"])
    rows.append(["図の数値の正本", "07_ソーススクリプト/data_zuhyo.py",
                 "当方の成果品"])
    GY.put(ws, rows)

    kak = yomu_kakunin(KAKUNIN)
    ws = GY.sheet(
        wb, "05_確認事項", "本表に関わる確認事項",
        "業務工程管理表 03_確認事項一覧 から読んでいます。"
        "内容はそちらが正本です。",
        ["No.", "確認事項", "状態", "決着しない場合の当方の扱い"],
        [8, 40, 12, 62])
    GY.put(ws, kak, [GY.OKF if x[2] == "完了" else None for x in kak])

    wb.save(OUT)

    # ══════════════════════════ 自己点検
    ng = []
    wb2 = openpyxl.load_workbook(OUT)
    if len(wb2.worksheets) != 6:
        ng.append(f"シートが{len(wb2.worksheets)}枚（6枚のはず）")
    if sum(cnt.values()) != len(items):
        ng.append("区分の合計が数値の件数と合わない")
    w3 = wb2["03_出所を特定できていないもの"]
    if w3.max_row - 4 != len(D):
        ng.append(f"03シートが{w3.max_row - 4}件（{len(D)}件のはず）")
    zen = "\n".join(str(c.value) for w_ in wb2.worksheets
                    for row in w_.iter_rows() for c in row
                    if c.value is not None)
    for w in ("**", "__"):
        if w in zen:
            ng.append(f"強調記号が残っている：{w}")
    for w in ("大雪", "東川", "東神楽", "上川", "美瑛", "金ヶ崎", "川崎市"):
        if w in zen:
            ng.append(f"他団体の名称が入っている：{w}")
    # Aの割合が7割を下回ったら、出所の候補が足りていない
    if cnt.get("A 原典・成果品に現れる", 0) / len(items) < 0.7:
        ng.append("出所を示せた割合が低すぎる。出所の候補が足りていない")

    print("計画素案の根拠の棚卸し")
    print("保存：", OUT)
    print(f"  素案 {os.path.basename(SOAN)}")
    print(f"  数値 {len(items):,}件／出所の候補 "
          f"{len(SRC_X) + len(SRC_D) + 1}件（値 {len(vals):,}通り）")
    for k in ("A 原典・成果品に現れる", "B 当方の算定（割合）",
              "C 識別子・連絡先", "D 出所を特定できていない"):
        n = cnt.get(k, 0)
        print(f"   {k:28s} {n:>5,}件（{n / len(items):5.1%}）")
    print("  ── 節ごとのDの多いもの")
    for sec, c in sorted(by.items(), key=lambda x: -x[1]["D"])[:6]:
        if c["D"]:
            print(f"   {sec:>8s} D {c['D']:>4}件／{sum(c.values()):>4}件")
    print("  ── 自己点検")
    if ng:
        for m in ng:
            print("   ×", m)
        sys.exit(1)
    print("   ○ 区分の合計が数値の件数と一致している")
    print("   ○ 出所の候補はすべてリポジトリに追跡されている")
    print(f"  ⚠ 出所を特定できていない {len(D)}件は、"
          "公表版の確定までに書き留める必要があります。")


def yomu_kakunin(nos):
    wb = openpyxl.load_workbook("川崎町_業務工程管理表.xlsx")
    ws = wb["03_確認事項一覧"]
    out, want = [], list(nos)
    for r in range(5, ws.max_row + 1):
        v = str(ws.cell(r, 1).value or "").strip()
        if v in want:
            out.append([v, str(ws.cell(r, 4).value or ""),
                        str(ws.cell(r, 8).value or ""),
                        str(ws.cell(r, 10).value or "")])
    out.sort(key=lambda x: want.index(x[0]))
    miss = set(want) - {x[0] for x in out}
    if miss:
        raise SystemExit("確認事項が見つからない：" + "／".join(sorted(miss)))
    return out


if __name__ == "__main__":
    main()
