# -*- coding: utf-8 -*-
"""
北塩原村　成果物の横断整合検査

なぜ置くか
  令和8年10月24日の他ＡＩの横断レビュー（RT-01〜RT-17）と、当方が
  同日に行ったレッドチームの点検で、成果物どうしの食い違いが多数見つかった。
    ・アンケートの回収数（38件と40件）が正本・概要版・協議会資料で混在
    ・前回計画の評価（協議会資料3）が、台帳で直す前の実績のまま
    ・概要版の見込量・成果目標が正本と違い、第4次計画を「改定対象外」
    ・照会の期限の曜日の誤り（（金）と書いた日が土曜日）
  いずれも各生成器の自己点検は通っていた。自己点検は「自分の出力と
  自分の入力」を見るだけで、他の成果物と食い違っていても止まらない。
  本書は、成果物の現物どうしを横に並べて照らす点検を一か所に置く。
  全生成器の最後に実行し、一つでも不合格があれば止まる。

点検の一覧（KENSA）
  K1  日付と曜日が合うこと
  K2  アンケートの回収数（障がい者調査 有効回答40件）の値であること
  K3  第4次障がい者計画を「改定対象外」と書いていないこと（中間見直し）
  K4  第7期の評価を「最終評価」と書いていないこと（令和8年度末の見込み）
  K5  委員会・住民に出す資料に社内の呼び名（正本）が出ていないこと
  K6  概要版の見込量・実績が正本 第5章の表と同じであること
  K7  協議会資料3の見込量の評価の実績が正本 第5章の表と同じであること
  K8  パブリックコメントの期間が工程（案Ａ）と同じであること
  K9  参照している照会番号が確認事項の台帳に実在すること
  K10 図表の点数が図表番号一覧と同じであること
  K11 令和11年度の高齢化率が正本の値であること（村の推計を併記する場合は
      照会番号を添える）
  K12 他団体の名前・機微な語（利用先の事業所の種類が分かる語）が無いこと
  K13 「第５章２（６）」等の章節の参照が正本の見出しにあること
  K14 「表５-25」等の図表番号の参照が正本の図表にあること
  K15 台帳の反映前に実績０で、反映後に利用があるサービスを「実績０」と
      書いていないこと（短期入所・意思疎通支援事業・理解促進研修等）

点検の範囲（令和8年10月24日に広げた）
  上の23点（SEIKA）に加えて、output の他の成果物（作業の成果物）にも
  K1・K3・K12・K15 を当てる。作業の成果物にも村・他メンバーに渡したものが
  あり、曜日の誤り・他団体の名前・古い実績が残ると同じ誤りが広がるため。
  K13・K14 は正本を指す参照が多い23点に限る（作業の成果物は、移植前の
  当方の素案（北塩原村_計画素案.docx）の章立てを指すものが多いため）。

許す例外（RE_YURUSU）
  古い値を「直す前」として示す成果物（対応履歴・差し替えの新旧・
  他ＡＩのレビューの確認）は、K2・K3・K8 の対象から外す。
  外す理由を成果物ごとに書いておき、黙って外さない。

否定試験（74）
  各点検に、誤りをわざと入れた文を与え、止まることを確かめる。
  K13〜K15 を足した後も、すべての点検について行う。

出力: output/北塩原村_成果物の横断整合検査.xlsx
"""

import datetime as dt
import hashlib
import os
import re
import sys

import docx
from openpyxl import load_workbook, Workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from kitashiobara_common import (  # noqa: E402
    add_sheet, ensure_out_dir, style_header_row, write_row,
)
import kitashiobara_anke as A  # noqa: E402
from kitashiobara_koutei import PABU_KAISHI, PABU_SHURYO, wareki  # noqa: E402

REPO_ROOT = os.path.dirname(os.path.abspath(__file__))
O = f"{REPO_ROOT}/output"
OUT_FILE = f"{O}/北塩原村_成果物の横断整合検査.xlsx"

# 点検の対象（委員会・村・他メンバーに出す成果物と、その説明の資料）
SEIKA = {
    "正本": f"{O}/北塩原村_計画素案_正本_移植後.docx",
    "概要版": f"{O}/北塩原村_計画素案_概要版.docx",
    "協議会資料1": f"{O}/会議資料/01_協議会資料1_計画策定の概要.docx",
    "協議会資料2": f"{O}/会議資料/02_協議会資料2_国の基本指針の改正点.docx",
    "協議会資料3": f"{O}/会議資料/03_協議会資料3_前回計画の評価.docx",
    "協議会資料8": f"{O}/会議資料/04_協議会資料8_提供体制の課題.docx",
    "庁議資料2": f"{O}/会議資料/05_庁議資料2_関連計画との整合.docx",
    "庁議資料4": f"{O}/会議資料/06_庁議資料4_庁内各課への依頼事項.docx",
    "前回計画の評価_概要版": f"{O}/北塩原村_前回計画の評価_概要版.docx",
    "アンケート結果報告_概要版": f"{O}/北塩原村_アンケート結果報告_概要版.docx",
    "前回計画の評価整理": f"{O}/北塩原村_前回計画の評価整理.xlsx",
    "策定委員会資料_修正指示": f"{O}/北塩原村_策定委員会資料_修正指示.docx",
    "村への照会": f"{O}/北塩原村_村への照会_6通.docx",
    "他メンバーへの確認": f"{O}/北塩原村_他メンバーへの確認_1通.docx",
    "確認事項の区分": f"{O}/北塩原村_確認事項の区分.xlsx",
    "介護保険側への申し送り": f"{O}/北塩原村_介護保険側への申し送り.docx",
    "KPI定義票": f"{O}/北塩原村_KPI定義票_様式案.xlsx",
    "見込量の再算定_参考値": f"{O}/北塩原村_見込量の再算定_参考値.xlsx",
    "村提供台帳の確認": f"{O}/北塩原村_村提供台帳の確認.xlsx",
    "他AIレビューの確認": f"{O}/北塩原村_他AIレビューの確認.xlsx",
    "横断レビューへの対応": f"{O}/北塩原村_横断レビューへの対応.xlsx",
    "アンケート40件への差し替え": f"{O}/北塩原村_アンケート40件への差し替え.xlsx",
    "他メンバー版への対応履歴": f"{O}/北塩原村_他メンバー版への対応履歴.xlsx",
}

# 委員会・住民に直接出すもの（K5 の対象）
TAIGAI = ["概要版", "協議会資料1", "協議会資料2", "協議会資料3", "協議会資料8",
          "庁議資料2", "庁議資料4", "前回計画の評価_概要版",
          "アンケート結果報告_概要版"]

# 古い値を「直す前」として示すため、K2・K3・K8 の対象から外すもの
RE_YURUSU = {
    "他メンバー版への対応履歴": "書き換えの「直す前」の列に他メンバー版の文を示す",
    "アンケート40件への差し替え": "38件と40件の新旧を並べて示す",
    "他AIレビューの確認": "他ＡＩのレビューの指摘（直す前の状態）を引用する",
    "横断レビューへの対応": "他ＡＩの横断レビューの指摘（直す前の状態）を引用する",
}


# K15 だけを外すもの（台帳による書き換えの「直す前」の列を持つ）
RE_YURUSU_K15 = {
    "村提供台帳の確認": "台帳による書き換え（ＤＫ書換）の「直す前」の列に他メンバー版の文を示す",
}

# 作業の成果物のうち、古い記述を「直す前」として引くため K3・K15 から外すもの
RE_YURUSU_SAGYOU = {
    "北塩原村_申し送り_他メンバー版への移植_20261009.docx":
        "令和8年10月9日時点の申し送りの記録。他メンバー版の文を引用している",
    "北塩原村_計画素案.docx":
        "他メンバー版への移植前の当方の素案（令和8年10月9日まで）。"
        "以後の修正は正本で行い、この版は記録として残す",
}

# ===========================================================================
# 1. 成果物を読む
# ===========================================================================
def yomu(path):
    """成果物の文を、段落・表のセル・シートの行ごとに返す。"""
    if path.endswith(".docx"):
        d = docx.Document(path)
        out = [p.text for p in d.paragraphs]
        for t in d.tables:
            for r in t.rows:
                for c in r.cells:
                    out.append(c.text)
        return out
    wb = load_workbook(path, read_only=True, data_only=True)
    out = []
    for ws in wb.worksheets:
        for r in ws.iter_rows(values_only=True):
            out += [str(x) for x in r if x not in (None, "")]
    wb.close()
    return out


def yomu_gyou(path):
    """成果物の文を、段落・表の行・シートの行ごとに返す（行のセルは「 | 」で
    つなぐ）。サービス名と値が別のセルにある表を、行として読むために使う。"""
    if path.endswith(".docx"):
        d = docx.Document(path)
        out = [p.text for p in d.paragraphs]
        for t in d.tables:
            for r in t.rows:
                out.append(" | ".join(c.text for c in r.cells))
        return out
    wb = load_workbook(path, read_only=True, data_only=True)
    out = []
    for ws in wb.worksheets:
        for r in ws.iter_rows(values_only=True):
            v = [str(x) for x in r if x not in (None, "")]
            if v:
                out.append(" | ".join(v))
    wb.close()
    return out


def sagyou_seika():
    """SEIKA の外にある output の成果物（作業の成果物）。本書の出力は除く。"""
    import glob
    seika = {os.path.abspath(p) for p in SEIKA.values()}
    out = {}
    for p in sorted(glob.glob(f"{O}/*.docx") + glob.glob(f"{O}/*.xlsx")
                    + glob.glob(f"{O}/会議資料/*.docx")):
        if os.path.abspath(p) in seika or os.path.abspath(p) == OUT_FILE:
            continue
        out[os.path.relpath(p, O)] = p
    return out


ZEN = str.maketrans("０１２３４５６７８９", "0123456789")
YOUBI = "月火水木金土日"


def _han(s):
    return int(str(s).translate(ZEN))


# ===========================================================================
# 2. 点検
# ===========================================================================
def k1_youbi(text):
    """「令和８年10月23日（金）」の曜日が暦と合うか。年の無い「10月30日（金）」
    は同じ文の前の年を引き継ぎ、無ければ4〜12月は令和8年、1〜3月は令和9年とする。"""
    ng = []
    nen = None
    for m in re.finditer(r"(?:令和([０-９0-9]+)年)?([０-９0-9]{1,2})月([０-９0-9]{1,2})日"
                         r"（([月火水木金土日])）", text):
        if m.group(1):
            nen = _han(m.group(1))
        tsuki, hi = _han(m.group(2)), _han(m.group(3))
        n = nen if nen else (8 if tsuki >= 4 else 9)
        try:
            d = dt.date(2018 + n, tsuki, hi)
        except ValueError:
            ng.append(f"暦に無い日: {m.group(0)}")
            continue
        if YOUBI[d.weekday()] != m.group(4):
            ng.append(f"{m.group(0)} は{YOUBI[d.weekday()]}曜日")
    return ng


def _kyu_wariai():
    """38件の値のうち40件の値と違い、他の値と紛れないもの。"""
    shin = {A.pc(k) for k in A.M}
    out = {}
    for k in A.M:
        kyu = A.pc(k, kyu=True)
        if kyu != A.pc(k) and kyu not in shin and kyu not in out:
            out[kyu] = k
    #   障がい別の値（精神障がい50.0％等）と同じ値は除く。
    #   57.9％は前回調査（令和５年）の回収率と同じ値であるため除く。
    for v in ("50.0％", "16.7％", "66.7％", "83.3％", "100.0％", "0.0％",
              "57.9％"):
        out.pop(v, None)
    return out


KYU = _kyu_wariai()
KYU_HYOKI = ["（n=38）", "n=38", "ｎ＝38", "n＝38", "有効回答38", "回答者数38人",
             "回答者38人", "38人（22.1％）"]
#   調査と関係なく現れうる値は、調査の語が同じ文にあるときだけ数える
CHOUSA_GO = ("アンケート", "回答", "調査", "％）", "割合")


def k2_anke(text):
    ng = []
    for g in KYU_HYOKI:
        if g in text:
            ng.append(f"38件の表記: {g}")
    for v, nm in KYU.items():
        #   「57.9％」の中の「7.9％」を拾わないよう、前が数字でないことを見る
        for m in re.finditer(r"(?<![0-9.])" + re.escape(v), text):
            if any(g in text for g in CHOUSA_GO):
                i = m.start()
                ng.append(f"38件の割合 {v}（{nm}）: …{text[max(0, i - 24):i + 8]}")
                break
    return ng


def k3_dai4ji(text):
    return ["「改定対象外」"] if "改定対象外" in text else []


def k4_saishu(text):
    return ["「最終評価」"] if "最終評価" in text else []


def k5_shanai(text):
    return ["社内の呼び名「正本」"] if "正本" in text else []


def k8_pabu(text):
    ng = []
    #   村の策定委員会資料の直す前の日程を引く文は除く（修正指示・照会で、
    #   どこを直すかを示すために引いている）
    if "策定委員会資料" in text or "【現在の記述】" in text or "【理由】" in text:
        return ng
    for g in ("２月８日", "２月22日", "2月8日", "2月22日"):
        if g in text and "パブリック" in text:
            ng.append(f"従前のパブリックコメントの日: {g}")
    return ng


def k12_kibi(text):
    ng = []
    for g in ("ダルク", "小野町", "金ヶ崎", "金ケ崎", "阿蘇", "札幌", "川崎町",
              "大雪"):
        if g in text:
            ng.append(f"出してはならない語: {g}")
    return ng


# ---------------------------------------------------------------------------
# 数値の照合（正本の表の現物と、概要版・協議会資料3）
# ---------------------------------------------------------------------------
def _honpon_hyou(nendo):
    from build_kitashiobara_gaiyou import honpon_hyou
    return honpon_hyou(nendo)


def k6_gaiyou():
    """概要版の「令和７年度→令和11年度」の表を正本の表と照らす。"""
    ng = []
    jis, mik = _honpon_hyou("実績"), _honpon_hyou("見込")
    d = docx.Document(SEIKA["概要版"])
    n = 0
    for t in d.tables:
        rows = [[c.text.strip() for c in r.cells] for r in t.rows]
        if not rows or "令和７年度→令和11年度" not in rows[0]:
            continue
        for r in rows[1:]:
            for name, val in ((r[0], r[1]), (r[2], r[3])):
                if not name:
                    continue
                n += 1
                m = re.match(r"^(―|[0-9.]+人)→([0-9.]+)人$", val)
                if not m:
                    ng.append(f"形が違う: {name} {val}")
                    continue
                r7 = jis.get(name, [None, None, None])[1]
                want7 = "―" if r7 is None else f"{r7}人"
                if m.group(1) != want7:
                    ng.append(f"{name} の令和７年度 {m.group(1)}（正本 {want7}）")
                if name not in mik or m.group(2) != str(mik[name][2]):
                    ng.append(f"{name} の令和11年度 {m.group(2)}人"
                              f"（正本 {mik.get(name, ['', '', '無し'])[2]}）")
    if n == 0:
        ng.append("概要版に見込量の表が見つからない")
    return ng, n


def k7_hyouka():
    """協議会資料3の見込量の評価の「令和６年度 Ｘ人分（見込量Ｙ人分）」を
    正本の実績の表の「Ｘ（Ｙ）人分」と照らす。"""
    ng = []
    d = docx.Document(SEIKA["協議会資料3"])
    hon = docx.Document(SEIKA["正本"])
    hgyou = []
    for t in hon.tables:
        for r in t.rows:
            hgyou.append([c.text.replace("\n", "") for c in r.cells])
    n = 0
    for t in d.tables:
        rows = [[c.text for c in r.cells] for r in t.rows]
        if not rows or rows[0][:3] != ["サービス", "区分", "実績の状況"]:
            continue
        for r in rows[1:]:
            for m in re.finditer(r"(令和[６７８]年度)(?:は年度途中までの平均で)? ?"
                                 r"([0-9.]+)([^（]*?)（見込量([0-9.]+)", r[2]):
                n += 1
                cell = f"{m.group(2)}（{m.group(4)}）"
                hit = [g for g in hgyou
                       if g and g[0].startswith(r[0][:4]) and
                       any(x.startswith(cell) for x in g)]
                if not hit:
                    ng.append(f"{r[0]} {m.group(1)} {cell} が正本の表に無い")
    if n == 0:
        ng.append("協議会資料3に見込量の評価の表が見つからない")
    return ng, n


def k9_shoukai(text, aru):
    ng = []
    for m in re.finditer(r"(?<![Ａ-Ｚ])[ＭＳ]-\d+", text):
        if m.group(0) not in aru:
            ng.append(f"台帳に無い照会番号: {m.group(0)}")
    return ng


def k10_zuhyou(text, nz, nh):
    """計画素案（正本）の図表の点数を書いた文に限る（報告書の図表の点数は別）。"""
    ng = []
    for m in re.finditer(r"図(\d+)点・表(\d+)点", text):
        mado = text[max(0, m.start() - 40):m.end() + 20]
        if not any(g in mado for g in ("図表番号", "図表目次", "正本")):
            continue
        if (int(m.group(1)), int(m.group(2))) != (nz, nh):
            ng.append(f"図表の点数 {m.group(0)}（図表番号一覧は図{nz}点・表{nh}点）")
    return ng


def k11_koureika(zen):
    """村の推計（48.1％）を書く成果物は、照会（Ｍ-142）を示していること。
    以前は工程の照会（Ｍ-123）を誤って示していた。"""
    ng = []
    zen = zen.replace("48.1%", "48.1％")
    if "48.1％" in zen and "Ｍ-142" not in zen:
        ng.append("村の推計（48.1％）を書きながら照会（Ｍ-142）を示していない")
    if "48.1％" in zen and "Ｍ-123" in zen and "Ｍ-123は工程" not in zen:
        for line in zen.split("\n"):
            if "Ｍ-123" in line and ("人口" in line or "高齢化率" in line):
                ng.append("人口の見通しに工程の照会（Ｍ-123）を示している")
                break
    return ng


# ---------------------------------------------------------------------------
# 参照の照合（章節・図表番号）と、台帳の反映前の実績
# ---------------------------------------------------------------------------
def _nfkc(s):
    import unicodedata
    return unicodedata.normalize("NFKC", s)


def honpon_mokuji():
    """正本の見出しから {(章, 節, 項)} と図表番号の集合を作る。

    章は「第５章」（Heading 1）、節は「２　」（Heading 2）、項は「（６）」
    （Heading 3、又は節の中で「（６）」で始まる段落）。"""
    d = docx.Document(SEIKA["正本"])
    secs, caps = set(), set()
    sho = setsu = None
    for p in d.paragraphs:
        t = _nfkc(p.text.strip())
        st = p.style.name if p.style is not None else ""
        m = re.match(r"(表|図)(\d+)-(\d+)\s", t)
        if m:
            caps.add(f"{m.group(1)}{m.group(2)}-{m.group(3)}")
        if st == "Heading 1":
            m = re.match(r"第(\d+)章", t)
            if m:
                sho, setsu = m.group(1), None
                secs.add((sho, None, None))
            continue
        if st == "Heading 2" and sho:
            m = re.match(r"(\d+)\s", t)
            if m:
                setsu = m.group(1)
                secs.add((sho, setsu, None))
            continue
        m = re.match(r"\((\d+)\)", t)
        if m and sho and setsu:
            secs.add((sho, setsu, m.group(1)))
    return secs, caps


#   正本ではない文書（前回計画・県の計画・報告書・国の告示等）の章節・図表を
#   指す文は、参照の直前にその文書の名がある。その場合は照合しない。
HOKA_BUNSHO = re.compile(
    r"(前回|現行計画|県|第7期|報告書|骨子|第4次|介護保険|高齢者|国の|基本指針|"
    r"旧|他メンバー版|移植前|医療計画|本編|当方の素案|レビュー|指針|告示|法|"
    r"集計表|調査票|資料\d|図表)")


def k13_shousetsu(text, secs):
    ng = []
    t = _nfkc(text)
    for m in re.finditer(r"第(\d+)章(\d+)?(?:\((\d+)\))?", t):
        if m.groups() in secs:
            continue
        if HOKA_BUNSHO.search(t[max(0, m.start() - 20):m.start()]):
            continue
        if "医療計画" in t[max(0, m.start() - 80):m.start()]:
            continue
        ng.append(f"正本に無い章節 {m.group(0)}: …{t[max(0, m.start() - 24):m.end() + 12]}")
    return ng


def k14_zuhyou_ref(text, caps):
    ng = []
    t = _nfkc(text)
    for m in re.finditer(r"(?<![図0-9A-Za-z])(表|図)(\d+)-(\d+)(?!-\d)", t):
        key = f"{m.group(1)}{m.group(2)}-{m.group(3)}"
        if key in caps:
            continue
        if HOKA_BUNSHO.search(t[max(0, m.start() - 30):m.start()]):
            continue
        ng.append(f"正本に無い図表番号 {key}: …{t[max(0, m.start() - 24):m.end() + 12]}")
    return ng


def zero_kara_riyou():
    """台帳の反映前は3年度とも0で、反映後に利用があるサービスの名（括弧の前）。"""
    from build_kitashiobara_mura_daicho import DK_HYOU
    out = set()
    for r in DK_HYOU:
        if all(v.startswith("0（") for v in r["mae"]) and \
                any(not v.startswith(("0（", "―")) for v in r["ato"]):
            #   日常生活用具は種目ごと（「日常生活用具：居宅生活動作補助用具」）
            out.add(r["svc"].replace("\n", "").split("（")[0].split("：")[-1])
    return sorted(out)


ZERO_RIYOU = zero_kara_riyou()


def k15_jisseki0(text):
    """「短期入所の実績は0」等、台帳の反映前の値を書いた文。前回計画の期間
    （令和３〜５年度）の実績を述べる文は除く。"""
    ng = []
    jis0 = r"実績(?:は|が|・見込とも|・見込量とも)?0(?![.0-9])"
    #   表の行で、利用のあるサービスを「未実施」としているもの
    for svc in ZERO_RIYOU:
        for cell_gyou in [_nfkc(text)]:
            sel = [c.strip() for c in cell_gyou.split("|")]
            if any(c.startswith(svc) or c.endswith(svc) for c in sel) and "未実施" in sel:
                ng.append(f"{svc}を未実施としている: {cell_gyou[:70]}")
    for bun in re.split(r"[。\n]", _nfkc(text)):
        if re.search(r"令和[3-5](?:年度|〜)", bun):
            continue
        for svc in ZERO_RIYOU:
            #   サービス名のすぐ後に「実績0」があるもの（「短期入所の実績は0」
            #   「意思疎通支援事業 | 実績0人」）。「短期入所には利用があり、
            #   実績が0であった移動支援」のように別の事柄を述べる文は除く
            if re.search(re.escape(svc) + r"(?:（[^）]*）)?\s*(?:\|\s*)?"
                         r"(?:の|は|が|で)?\s*" + jis0, bun):
                ng.append(f"{svc}を実績0としている: {bun[:70]}")
            #   「理解促進研修・啓発事業、…など、これまでの実績がない事業」
            if re.search(re.escape(svc) + r"[^。|]{0,40}?(?:これまでの実績がない|"
                         r"実績がありません|実績はありません)", bun):
                ng.append(f"{svc}を実績がないとしている: {bun[:70]}")
    #   短期入所の旧い表し方（年間の請求件数を、単位を示さずに実績としたもの）
    for m in re.finditer(r"令和6年度2件・令和7年度3件", _nfkc(text)):
        if "請求" not in _nfkc(text)[max(0, m.start() - 12):m.start()]:
            ng.append(f"短期入所の実績を請求件数のまま書いている: "
                      f"{_nfkc(text)[max(0, m.start() - 20):m.end() + 10]}")
    return ng


# ---------------------------------------------------------------------------
# 委員会・住民に出す前の確認（不合格にはしない。ご判断のための数え上げ）
#   他ＡＩの横断レビュー（第5節 Step 3）が、残っている「要確認」「M-***」等の
#   対外公開の可否を確かめるよう求めたもの。照会番号や確認中の注記を
#   委員会資料に残すかどうかは、村とご相談のうえで決める事項であるため、
#   数と例を示すにとどめる。
# ---------------------------------------------------------------------------
KOUKAI_GO = [
    ("照会番号（Ｍ-）", r"(?<![Ａ-Ｚ])Ｍ-\d+"),
    ("照会番号（Ｓ-）", r"(?<![Ａ-Ｚ])Ｓ-\d+"),
    ("【村へ確認】等の注記", r"【(?:村|県|他メンバー)へ確認"),
    ("【確認中】【要確認】", r"【(?:確認中|要確認)】"),
    ("「確認中」「確認しています」", r"確認中|確認しています"),
]


def koukai_mae(texts):
    out = []
    for name in TAIGAI:
        zen = "\n".join(texts.get(name, []))
        for label, pat in KOUKAI_GO:
            hits = re.findall(pat, zen)
            if hits:
                m = re.search(pat, zen)
                rei = zen[max(0, m.start() - 20):m.end() + 20].replace("\n", " ")
                out.append((name, label, len(hits), rei))
    return out


# ===========================================================================
# 3. 全体の実行
# ===========================================================================
KENSA = [
    ("K1", "日付と曜日が合う", "すべて（作業の成果物を含む）"),
    ("K2", "アンケートは有効回答40件の値", "すべて（直す前を示すものを除く）"),
    ("K3", "第4次計画を「改定対象外」と書かない",
     "すべて（作業の成果物を含む。直す前を示すものを除く）"),
    ("K4", "第7期の評価を「最終評価」と書かない",
     "協議会資料3・前回計画の評価_概要版・前回計画の評価整理"),
    ("K5", "社内の呼び名（正本）を出さない", "委員会・住民に出す資料"),
    ("K6", "概要版の見込量・実績＝正本 第5章の表", "概要版"),
    ("K7", "協議会資料3の実績＝正本 第5章の表", "協議会資料3"),
    ("K8", "パブリックコメントの期間は案Ａ", "すべて（直す前を示すものを除く）"),
    ("K9", "照会番号が確認事項の台帳に実在", "すべて"),
    ("K10", "図表の点数が図表番号一覧と同じ", "すべて"),
    ("K11", "高齢化率の村の推計には照会番号を添える", "すべて"),
    ("K12", "他団体の名前・機微な語を出さない", "すべて（作業の成果物を含む）"),
    ("K13", "章節の参照が正本の見出しにある", "23点"),
    ("K14", "図表番号の参照が正本の図表にある", "23点"),
    ("K15", "台帳の反映前の「実績0」を書かない", "すべて（作業の成果物を含む）"),
]


def jikkou(texts=None, gyou=None, sagyou=None):
    """texts を与えると、成果物の代わりにその文で点検する（否定試験用）。

    texts  : 23点の文（段落・セルごと）
    gyou   : 23点の文（段落・表の行ごと）。K13〜K15 に使う
    sagyou : 作業の成果物の文（行ごと）。K1・K3・K12・K15 を当てる
    """
    from build_kitashiobara_kakunin_kubun import KUBUN as KK
    from build_kitashiobara_zuhyo_bangou import HYO, ZU
    aru = {r[0] for r in KK}
    secs, caps = honpon_mokuji()
    if texts is None:
        texts = {k: yomu(p) for k, p in SEIKA.items()}
        gyou = {k: yomu_gyou(p) for k, p in SEIKA.items()}
        sagyou = {k: yomu_gyou(p) for k, p in sagyou_seika().items()}
    if gyou is None:
        gyou = texts
    kekka = []          # (点検, 成果物, 内容)
    for name, lines in texts.items():
        zen = "\n".join(lines)
        for line in lines:
            for x in k1_youbi(line):
                kekka.append(("K1", name, x))
            for x in k12_kibi(line):
                kekka.append(("K12", name, x))
            for x in k9_shoukai(line, aru):
                kekka.append(("K9", name, x))
            for x in k10_zuhyou(line, len(ZU), len(HYO)):
                kekka.append(("K10", name, x))
            if name not in RE_YURUSU:
                for x in k2_anke(line):
                    kekka.append(("K2", name, x))
                for x in k3_dai4ji(line):
                    kekka.append(("K3", name, x))
                for x in k8_pabu(line):
                    kekka.append(("K8", name, x))
            if name in ("協議会資料3", "前回計画の評価_概要版", "前回計画の評価整理"):
                for x in k4_saishu(line):
                    kekka.append(("K4", name, x))
            if name in TAIGAI:
                for x in k5_shanai(line):
                    kekka.append(("K5", name, x))
        for x in k11_koureika(zen):
            kekka.append(("K11", name, x))
    for name, lines in gyou.items():
        for line in lines:
            for x in k13_shousetsu(line, secs):
                kekka.append(("K13", name, x))
            for x in k14_zuhyou_ref(line, caps):
                kekka.append(("K14", name, x))
            if name not in RE_YURUSU and name not in RE_YURUSU_K15:
                for x in k15_jisseki0(line):
                    kekka.append(("K15", name, x))
    for name, lines in (sagyou or {}).items():
        for line in lines:
            for x in k1_youbi(line):
                kekka.append(("K1", name, x))
            for x in k12_kibi(line):
                kekka.append(("K12", name, x))
            if name not in RE_YURUSU_SAGYOU:
                for x in k3_dai4ji(line):
                    kekka.append(("K3", name, x))
                for x in k15_jisseki0(line):
                    kekka.append(("K15", name, x))
    return kekka


def hitei_shiken():
    """否定試験74：誤りを入れた文を与え、各点検が止めることを確かめる。"""
    warui = {
        "否定試験": [
            "期限 令和８年10月31日（金）",                              # K1
            "有効回答率22.1％（n=38）で、回答者の65.8％が65歳以上",       # K2
            "第４次北塩原村障がい者計画は改定対象外",                     # K3
            "本村の正本では",                                             # K5
            "パブリックコメント 令和９年２月８日～２月22日",              # K8
            "（Ｍ-999）を照会中",                                         # K9
            "図表番号は図26点・表95点です",                               # K10
            "高齢化率は48.1％",                                           # K11
            "利用先はダルク",                                             # K12
            "大雪地区で用いた手順",                                       # K12
            "当方は第５章13に書いています",                               # K13
            "第４章２（９）の目標",                                       # K13
            "表５-99のとおり",                                            # K14
            "短期入所の実績は0だが利用希望が10.0％",                      # K15
            "意思疎通支援事業 | 実績0人",                                 # K15
            "理解促進研修・啓発事業は実績・見込とも0",                     # K15
            "必須 | 意思疎通支援事業 | 法第77条第1項第6号 | 人 | 0 | 未実施",  # K15
            "理解促進研修・啓発事業、自発的活動支援事業など、これまでの実績がない事業",
            "短期入所は令和６年度２件・令和７年度３件の実績",            # K15
        ],
    }
    #   K4・K5 は対象の成果物の名前で効くため、名前を借りて入れる
    warui["協議会資料3"] = ["第７期・最終評価"]
    warui["概要版"] = ["正本 第５章"]
    kekka = jikkou(warui)
    deta = {k for k, _n, _x in kekka}
    hoshii = {"K1", "K2", "K3", "K4", "K5", "K8", "K9", "K10", "K11", "K12",
              "K13", "K14", "K15"}
    #   作業の成果物の側（K1・K3・K12・K15）も止まること
    kekka2 = jikkou({}, {}, {"否定試験（作業）": [
        "期限 令和８年10月31日（金）", "第４次計画は改定対象外",
        "川崎町で用いた手順", "短期入所の実績は0"]})
    deta2 = {k for k, _n, _x in kekka2}
    nuke = sorted(hoshii - deta) + [f"作業の成果物の{k}" for k in
                                    sorted({"K1", "K3", "K12", "K15"} - deta2)]
    #   正しい文では止まらないこと（誤検出の試験）
    yoi = jikkou({"否定試験（正）": [
        "第５章２（６）サービスの質の向上と人材の確保", "表５-25のとおり",
        "第８次福島県医療計画の第10章", "短期入所は令和３〜５年度とも実績0だった",
        "前回計画の第３章５",
        "第7期計画で見込んでいなかった短期入所には毎年度利用があり、"
        "利用の実績が0であった移動支援等"]}, None, {})
    nuke += [f"誤検出: {x}" for x in yoi]
    return nuke, kekka


def main():
    nuke, _ = hitei_shiken()
    if nuke:
        raise SystemExit(f"否定試験74：止まらない点検がある: {nuke}")
    kekka = jikkou()
    SAGYOU = sagyou_seika()
    ng6, n6 = k6_gaiyou()
    ng7, n7 = k7_hyouka()
    kekka += [("K6", "概要版", x) for x in ng6]
    kekka += [("K7", "協議会資料3", x) for x in ng7]

    ensure_out_dir()
    wb = Workbook()
    wb.remove(wb.active)
    ws = add_sheet(wb, "00_結論", "成果物の横断整合検査",
                   f"対象{len(SEIKA)}点（ほかに作業の成果物{len(SAGYOU)}点）・"
                   f"点検{len(KENSA)}種。"
                   f"不合格{len(kekka)}件。否定試験74は全点検で止まることを確認。",
                   [8, 44, 46, 12])
    style_header_row(ws, 5, ["点検", "内容", "対象", "不合格"])
    r = 6
    for i, (k, naiyou, taishou) in enumerate(KENSA):
        cnt = sum(1 for x in kekka if x[0] == k)
        r = write_row(ws, r, [k, naiyou, taishou, cnt], alt=(i % 2 == 1))
    r += 1
    write_row(ws, r, ["", f"K6 で照らした見込量 {n6}件・K7 で照らした実績 {n7}件",
                      "", ""])
    ws = add_sheet(wb, "01_不合格の明細", f"不合格 {len(kekka)}件",
                   "0件であれば合格。", [8, 26, 100])
    style_header_row(ws, 5, ["点検", "成果物", "内容"])
    r = 6
    for i, x in enumerate(kekka):
        r = write_row(ws, r, list(x), alt=(i % 2 == 1))
    ws = add_sheet(wb, "02_対象の成果物", f"対象 {len(SEIKA)}点",
                   "点検した時点のハッシュ（SHA-256の先頭16字）。",
                   [28, 70, 20, 40])
    style_header_row(ws, 5, ["成果物", "ファイル", "ハッシュ", "例外"])
    r = 6
    for i, (k, p) in enumerate(SEIKA.items()):
        h = hashlib.sha256(open(p, "rb").read()).hexdigest()[:16]
        r = write_row(ws, r, [k, os.path.relpath(p, REPO_ROOT), h,
                              RE_YURUSU.get(k, "") or RE_YURUSU_K15.get(k, "")],
                      alt=(i % 2 == 1))
    ws = add_sheet(wb, "03_作業の成果物", f"作業の成果物 {len(SAGYOU)}点",
                   "K1・K3・K12・K15 を当てたもの。点検した時点のハッシュ"
                   "（SHA-256の先頭16字）。", [60, 20, 40])
    style_header_row(ws, 5, ["ファイル", "ハッシュ", "例外"])
    r = 6
    for i, (k, p) in enumerate(SAGYOU.items()):
        h = hashlib.sha256(open(p, "rb").read()).hexdigest()[:16]
        r = write_row(ws, r, [os.path.relpath(p, REPO_ROOT), h,
                              RE_YURUSU_SAGYOU.get(k, "")], alt=(i % 2 == 1))
    km = koukai_mae({k: yomu(SEIKA[k]) for k in TAIGAI})
    ws = add_sheet(wb, "04_委員会に出す前の確認",
                   "委員会・住民に出す資料に残る照会番号・確認中の注記",
                   "不合格にはしない。照会番号や確認中の注記を残すかどうかは、"
                   "村とご相談のうえで決める事項（他ＡＩの横断レビュー 第5節 Step 3）。",
                   [26, 26, 8, 70])
    style_header_row(ws, 5, ["資料", "語", "数", "例"])
    r = 6
    for i, x in enumerate(km):
        r = write_row(ws, r, list(x), alt=(i % 2 == 1))
    wb.save(OUT_FILE)
    print(f"作成: {OUT_FILE}")
    if kekka:
        print(f"自己点検 不合格: {len(kekka)}件")
        for x in kekka[:200]:
            print("   -", x)
        raise SystemExit(1)
    print(f"  自己点検: 対象{len(SEIKA)}点＋作業の成果物{len(SAGYOU)}点・"
          f"点検{len(KENSA)}種すべて合格"
          f"（見込量{n6}件・実績{n7}件を正本の表と照合）／否定試験74")


if __name__ == "__main__":
    main()
