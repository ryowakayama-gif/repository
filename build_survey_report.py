# -*- coding: utf-8 -*-
"""大雪地区広域連合 第10期介護保険事業計画　アンケート調査の集計分析報告書.

業務仕様書「４．業務内容（3）実施済み調査結果の集計・分析」の成果として、
対象4調査の集計・分析の結果を1冊にまとめる。

  ① 在宅生活改善調査　　　　事業所票15件・利用者票99票
  ② 居所変更実態調査　　　　施設等票18件
  ③ 介護人材実態調査　　　　事業所票27件・職員個票317人・職員票26件
  ④ 健康とくらしの調査　　　個票4,729票（別冊「調査クロス集計・分析」24シート）

4調査を横断したクロス集計は、利用者票の所在地区の記入形式が統一されておらず
個票を地区に割り付けられないため、発注者のご意向により行わないこととした
（14シート）。各調査内のクロス集計及び公表データとの突合により、
供給構造の面から4調査の結果を接続する。

シート構成
  00_報告書の構成と調査の概要
  01_調査の実施状況と回収
  02_①在宅生活改善調査の結果
  03_②居所変更実態調査の結果
  04_③介護人材実態調査の結果
  05_④健康とくらしの調査の結果
  06_供給構造①_事業所数と定員
  07_供給構造②_サービスの実施地域
  08_供給構造③_運営法人の集中
  09_供給構造④_従事者数
  10_供給と需要の対照
  11_調査結果の限界と留保
  12_主要所見と計画本文への反映
  13_横断クロス集計を行わないことの整理
  14_公表データによる補完と留保の解消
"""

import collections
import os

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import runpy as _runpy
import sys as _sys

import data_survey2025 as S
import data_survey_kpi as KP
import data_survey_cross as C
import data_mieruka_km as MK
import data_hokkaido_roster as R
import data_hokkaido_shitei as H
import repo_paths as RP


class _Sink(object):
    """標準出力の捨て場。

    読み込む側のスクリプトが sys.stdout.buffer を包み直すことがあるため、
    buffer 属性も備える。close() は無視する。
    """

    closed = False
    encoding = "utf-8"
    errors = "strict"
    newlines = None
    line_buffering = False
    name = "<sink>"
    mode = "w"

    def __init__(self):
        self.buffer = self

    def fileno(self):
        raise OSError("sink")

    def isatty(self):
        return False

    def writelines(self, _lines):
        pass

    def detach(self):
        return self

    def reconfigure(self, **_k):
        pass

    def write(self, *_a, **_k):
        return 0

    def flush(self):
        pass

    def close(self):
        pass

    def readable(self):
        return False

    def writable(self):
        return True

    def seekable(self):
        return False


def _load(name):
    """別のビルドスクリプトの算定結果を読み込む（標準出力は捨てる）。

    見込量・突合の数値を本報告書に固定値で書き写すと、
    算定を改めたときに計画素案と本報告書とで別の答えが並ぶ。
    （令和8年9月17日の点検で、10シートの見込量が
      将来推計 第2段階の値のままであり、
      計画素案 第6章（第1次概算による値）と食い違っていたことが分かった。）
    """
    old = _sys.stdout
    _sys.stdout = _Sink()
    try:
        return _runpy.run_path(RP.ROOT + "/" + name)
    finally:
        _sys.stdout = old


_M = _load("build_mikomiryo_santei.py")      # サービス見込量 第1次概算
_X = _load("build_survey_jisseki_cross.py")  # 調査結果と年報実績の突合

# 第1次概算による施設・居住系の月平均利用者数（保険者単位）
_JIS, _SOG = _M["JISSEKI"], _M["sogaku"]
SHISETSU_LAB = [l for l in _M["SHISETSU_LAB"] if l.startswith("施設サービス")]
KYOJU_LAB = [l for l in _M["SHISETSU_LAB"] if l.startswith("居住系サービス")]
SHI_R7 = sum(sum(_JIS[l]) for l in SHISETSU_LAB)
SHI_R11 = sum(_SOG(l, "2029") for l in SHISETSU_LAB)
KYO_R7 = sum(sum(_JIS[l]) for l in KYOJU_LAB)
KYO_R11 = sum(_SOG(l, "2029") for l in KYOJU_LAB)
NINTEI_R7 = sum(_M["NIN_R7"])

# 定員に対する到達率が最も高いサービス（第1次概算 10シート）
_TEIIN = _M["TEIIN_MAP"]
_TOTATSU = sorted(
    ((_SOG(l, "2029") / _TEIIN[l][0], l) for l in _TEIIN),
    reverse=True)
TOTATSU_TOP = _TOTATSU[0]

# 調査①の代表性の検定（突合クロス集計 02シート）
CHI = _X["CU_CHI"]
CHI_1P = _X["CHI_1"]
CU_DO_OBS, CU_DO_EXP = _X["CU_DO"], _X["CU_EXP"]

# 調査②の在籍者と年報の受給者（突合クロス集計 03シート）
ZAISEKI_VS = [
    (nm, _X["zaiseki"](nm), _X["nenpo_kei"](keys))
    for nm, keys in _X["TAIO"] if keys
]
ROKEN_Z, ROKEN_J = [(z, j) for nm, z, j in ZAISEKI_VS
                    if nm == "介護老人保健施設"][0]

SW = [k for k in H.KOHYO if k["事業所名"] == "さわやか東神楽館"][0]
SW_N = sum(SW["要介護度別入居者数"].values())
SW_KAIGO = SW["介護職員_常勤"] + SW["介護職員_非常勤"]

OUT = RP.ROOT + "/output/第10期計画_アンケート調査の集計分析報告書.xlsx"

FONT = "游ゴシック"
NAVY, HEAD = "1F3864", "4472C4"
IN_Y, OK_G, NG_O, MID_B, GRAY = "FFF2CC", "E2EFDA", "FCE4D6", "DEEBF7", "F2F2F2"
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

TOWNS = ["東川町", "美瑛町", "東神楽町"]

# 事業所票（JIN）の並び順に対応するサービス区分。
# 別冊「実施済み3調査の受領点検と集計」のCATと同一である。
CAT = [
    "GH", "GH", "住宅型有料", "住宅型有料", "GH", "老健", "通所リハ", "GH",
    "サ高住", "通所介護等", "特定施設", "地密特養", "特定施設", "老健",
    "通所リハ", "通所介護等", "通所介護等", "特養", "地密特養", "特養",
    "通所リハ", "老健", "住宅型有料", "GH", "訪問介護", "訪問介護", "訪問介護",
]

# ------------------------------------------------ 計画素案の実物を読む
def _draft_sections():
    """計画素案（docx）を章節ごとのテキストに分ける。

    反映状況を固定値で書くと、素案を改訂したときに実際と合わなくなる。
    （令和8年9月17日の点検で、12シートが第2章第4節・第6章を
      いずれも「着手可」のままとしていたが、
      第2章第4節は令和8年9月15日に、第6章は同月16日に反映済みであった。）
    節を見ずに語の有無だけで判定すると偽陽性が出る。
    （同じ点検で、「除雪」は第1章第7節の除雪率の記述に当たり、
      「同規模保険者」は第3章第4節の記述に当たるため、
      いずれも所見の反映先とは別の箇所であった。）

    戻り値は {"第2章第4節": "その節の本文と表のテキスト", ...}。
    """
    if not os.path.exists(RP.DRAFT):
        return {}
    import re
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    d = Document(RP.DRAFT)
    sec, cur, out, started = {}, None, [], False
    body = d.element.body
    for child in body.iterchildren():
        if child.tag.endswith("}p"):
            t = Paragraph(child, d).text.strip()
            if re.match(r"^第\d+章", t):
                # 目次（章節が連続して並ぶ部分）を読み飛ばす
                started = True
                cur = t.split("\u3000")[0]
                out = []
                sec.setdefault(cur, [])
                chap = cur
                continue
            if started and re.match(r"^第\d+節", t) and "\t" not in t:
                cur = chap + t.split("\u3000")[0].split(" ")[0]
                sec.setdefault(cur, [])
                continue
            if cur:
                sec[cur].append(t)
        elif child.tag.endswith("}tbl") and cur:
            for row in Table(child, d).rows:
                sec[cur].extend(c.text for c in row.cells)
    return {k: "\n".join(v) for k, v in sec.items()}


DRAFT_SEC = _draft_sections()


def in_draft(section, *words):
    """指定した章節（例「第2章第4節」）にすべての語が現れるとき True。"""
    t = DRAFT_SEC.get(section, "")
    return bool(t) and all(w in t for w in words)


# ---------------------------------------------------------------- 集計
# ① 利用者票の分母は設問ごとに異なる。
#   問2（在宅生活の維持の見通し）の有効回答は98票（無回答1票）であり、
#   H06（在宅生活継続困難割合）は72票÷98票＝73.5％である。
#   要介護度の有効回答も98票であるが、これは別の設問の有効回答であって
#   同じ票を指すとは限らない。問3（より適切と思われるサービス）は
#   99票すべてに回答があるため、H12の分母は99票である。
RIYO_N = KP.RIYO_N
IJI_YUKO, IJI_KONNAN, H06 = KP.H06_YUKO, KP.H06_BUNSHI, KP.H06
H12_BUNSHI, H12_BUNBO, H12 = KP.H12_BUNSHI, KP.H12_BUNBO, KP.H12
H12_NAME = KP.H12_NAME

SIS = [j for j in S.JIN if j["区分"] == "施設・通所系"]
HOU = [j for j in S.JIN if j["区分"] != "施設・通所系"]
N_SIS = sum(j["職員"] or 0 for j in SIS)
N_HOU = sum(j["職員"] or 0 for j in HOU)
DUP = 13                       # 華とフラワーの重複（点検事項No.1）
N_ALL = N_SIS + N_HOU - DUP
# 同一法人が施設・通所系と訪問系の両方に同一の職員13人を記入しているため
# （点検事項No.1）、その事業所の採用7人・離職2人も二重に計上される。
# 単純合計は採用90人・離職67人であり、重複分を除くと採用83人・離職65人。
SAIYO_NAMA = sum(j["採用"] or 0 for j in S.JIN)
RISHOKU_NAMA = sum(j["離職"] or 0 for j in S.JIN)
SAIYO = SAIYO_NAMA - 7
RISHOKU = RISHOKU_NAMA - 2
GAIKOKU = sum(j["外国人"] or 0 for j in S.JIN) - 0
N_GAI_JIG = sum(1 for j in S.JIN if (j["外国人"] or 0) > 0)

# サービス区分別の介護職員数（華の重複13人は住宅型有料から除く）
BYCAT = collections.Counter()
for _c, _j in zip(CAT, S.JIN):
    BYCAT[_c] += _j["職員"] or 0
BYCAT["住宅型有料"] -= DUP
NOSER = sum(BYCAT[k] for k in ["GH", "特定施設", "住宅型有料", "サ高住"])

SK = S.SHOKU
SK_N = SK["件数"]
SK_YUKO = SK_N - 1             # 無効1人

# 名簿による定員
CAP_TOKUTEI = sum(y["定員"] for y in R.YU if y["類型"] == "介護付")
CAP_JUTAKU = sum(y["定員"] for y in R.YU if y["類型"] == "住宅型")
CAP_SAKO = sum(y["戸数"] for y in R.SA)
CAP_KEIHI = sum(y["定員"] for y in R.KE)
CAP_TOKUYO = sum(t["定員"] for t in H.TOKUYO if not t["地域密着型"])
CAP_CHITOKU = sum(t["定員"] for t in H.TOKUYO if t["地域密着型"])
CAP_SHORT = sum(t["ショート専用定員"] or 0 for t in H.TOKUYO)
CAP_ROKEN = sum(s["定員"] or 0 for s in S.SHI if s["種別"] == 7)
CAP_GH = sum(s["定員"] or 0 for s in S.SHI if s["種別"] == 4) + 18

# 指定事業所の延べ件数（運営規程の「通常の事業の実施地域」を把握した件数）
SHITEI_NOBE = sum(len(v) for v in H.SHITEI.values())
_SHITEI_NOBE = SHITEI_NOBE

# 実事業所（法人×事業所名でユニーク化）
UNIQ = {}
for _sv, _rows in H.SHITEI.items():
    for _r in _rows:
        UNIQ.setdefault((_r["法人"], _r["事業所名"]),
                        {"町": _r["町"], "区分": set()})["区分"].add(_sv)
BYH = collections.Counter(k[0] for k in UNIQ)
N_JIG = len(UNIQ)
N_HOJIN = len(BYH)
TOP6 = BYH.most_common(6)


def klast(code):
    v = MK.K[code]["値"]
    ks = [k for k in v if v[k] is not None]
    return int(v[ks[-1]])


wb = Workbook()


def sheet(name, title, subtitle, widths, freeze="A5"):
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    c = ws.cell(row=1, column=1, value=title)
    c.font = Font(name=FONT, size=14, bold=True, color=NAVY)
    ws.row_dimensions[1].height = 22
    c = ws.cell(row=2, column=1, value=subtitle)
    c.font = Font(name=FONT, size=9)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(widths))
    ws.row_dimensions[2].height = 52
    ws.freeze_panes = freeze
    return ws


def header(ws, row, cols, height=30):
    for i, v in enumerate(cols, start=1):
        c = ws.cell(row=row, column=i, value=v)
        c.font = Font(name=FONT, size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=HEAD)
        c.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[row].height = height
    return row + 1


def body(ws, row, vals, fills=None, height=22, align=None, bold=False):
    for i, v in enumerate(vals, start=1):
        c = ws.cell(row=row, column=i, value=v)
        c.font = Font(name=FONT, size=9, bold=bold)
        c.alignment = Alignment(wrap_text=True, vertical="top",
                                horizontal=(align or {}).get(i, "left"))
        c.border = BORDER
        if fills and i in fills and fills[i]:
            c.fill = PatternFill("solid", fgColor=fills[i])
    ws.row_dimensions[row].height = height
    return row + 1


def lead(ws, row, text, span):
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=10, bold=True, color=NAVY)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    ws.row_dimensions[row].height = 18
    return row + 1


def note(ws, row, text, span, height=None):
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=8.5)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    ws.row_dimensions[row].height = height or max(
        30, 13 * (len(text) // (span * 11) + 1))
    return row + 1


def dist(ws, r, title, d, labels, span, total=None, note_=""):
    """度数分布を描く。labels はコード→表示名。"""
    r = lead(ws, r, title, span)
    r = header(ws, r, ["区分", "人数・件数", "構成比"] + [""] * (span - 3))
    tot = total or sum(v for k, v in d.items() if k != "-")
    for code, lab in labels.items():
        v = d.get(code, 0)
        r = body(ws, r, [lab, v, "%.1f％" % (v / tot * 100) if tot else "―"]
                 + [""] * (span - 3), height=20,
                 align={2: "right", 3: "right"})
    mu = d.get("-", 0)
    if mu:
        r = body(ws, r, ["無回答・無効", mu, "―"] + [""] * (span - 3),
                 {1: GRAY}, height=20, align={2: "right"})
    r = body(ws, r, ["有効回答", tot, "100.0％"] + [""] * (span - 3),
             {1: GRAY}, height=20, align={2: "right", 3: "right"}, bold=True)
    if note_:
        r = note(ws, r, note_, span)
    return r + 1


# ============================================================ 00 構成
# 14シート（公表データによる補完と留保の解消）の明細。
# 00シートの要約が件数を固定値で書かないよう、ここで定義して数える。
HOKAN_KAISHO = [
    ("1", "訪問系は13事業所のうち3事業所からの回答であり、"
     "区域全体を代表しない（11シートNo.3・No.5）",
     "介護サービス情報公表システムの個別公表画面（13事業所）",
     "訪問介護13事業所すべての訪問介護員等の実人数を把握した。"
     "実人数181人で、調査回答分42人の約4.3倍である。"
     "調査と公表画面の双方がある3事業所のうち2事業所は"
     "実人数が完全に一致しており、"
     "公表画面の値は調査と同じ水準で扱える。",
     "第2章第6節に全数把握の表を追加した",
     "解消"),
    ("2", "訪問介護の担い手の構造が把握できない",
     "同上（事業所の所在する建物の記載）",
     "13事業所のうち8事業所（訪問介護員等128人・70.7％、"
     "利用者225人・65.0％）が住宅型有料老人ホーム、"
     "サービス付き高齢者向け住宅又は介護付有料老人ホームに"
     "併設されていることが確認できた。"
     "併設のない事業所は5事業所（53人）である。",
     "第2章第6節。"
     "住まいの整備方針と訪問介護の供給力が一体であることの根拠",
     "解消"),
    ("3", "訪問系の人材の動きが把握できない",
     "同上（前年度の採用者数・退職者数）",
     "前年度の採用23人・退職28人で、退職が5人上回る。"
     "退職が採用を上回る事業所は13事業所中5事業所である。",
     "第2章第6節、第5章 基本目標4",
     "解消"),
    ("4", "④健康とくらしの調査の世帯構成が"
     "国勢調査の世帯類型と対応しない（11シートNo.1）",
     "国勢調査（令和2年・令和7年速報）、"
     "見える化A5〜A8、在宅生活改善調査",
     "3つの出所を対照し、"
     "水準（美瑛町が高い）と増加の速さ（東神楽町が速い）が"
     "別であることを確認した。"
     "高齢者単身世帯の10年増加率は"
     "東神楽町＋73.6％・東川町＋66.1％・美瑛町＋34.5％である。",
     "第2章第1節の世帯の記述を是正した",
     "解消"),
    ("5", "H12（当初の定義「%s」）のデータ源がない" % KP.H12_KYU_NAME,
     "①在宅生活改善調査（より適切と思われるサービス。本報告書02シート）",
     "区域内に事業所が存在しない4サービスを"
     "「より適切と思われるサービス」として選んだ票は"
     "%d票中%d票（%.1f％）である。"
     "分母はこの設問の有効回答%d票である"
     "（要介護度の有効回答%d票ではない）。"
     "H12の代理指標として定義できる。指標名は"
     "「%s」に改める。"
     % (H12_BUNBO, H12_BUNSHI, H12, H12_BUNBO, IJI_YUKO, H12_NAME),
     "第4章第3節・資料1のH12",
     "解消"),
    ("6", "H07・H08のデータ源がない（在宅介護実態調査を未実施）",
     "④健康とくらしの調査 サブコア1【問17】4）及び"
     "【大雪－問4】1）",
     "主に介護をしている者の割合3.4％（n=2,217）、"
     "世話をしてくれる人がいない者の割合8.7％（n=4,599）。"
     "いずれも代理指標として算定できる。",
     "第4章第3節・資料1のH07・H08",
     "解消"),
]
HOKAN_ZANTEI = [
    ("7", "①利用者票は課題のある利用者を抽出する設計である"
     "（11シートNo.2）",
     "調査の設計そのものによるものであり、留保自体は残る。"
     "母集団が「在宅生活の維持が難しい利用者」であり、"
     "区域内の在宅利用者全体ではない。"
     "ただし年報（令和7年度）の要介護度別の明細を確認できたことにより、"
     "偏りの大きさと向きを測定できるようになった。",
     "測定はできた。留保そのものは解消できない。"
     "第11期に向けて設計の見直しを提案する",
     "「調査対象となった%d人のうち」と明記する。" % RIYO_N +
     "偏りをχ²＝%.1f（自由度6。1％点%.2f）として示せるようになったため、"
     "「課題のある層に偏っている」ではなく"
     "「要介護3が期待の%.1f倍、要支援1が期待の%.2f倍」と書ける。"
     "H12の%.1f％も同じ留保が付く"
     % (CHI, CHI_1P, CU_DO_OBS[4] / CU_DO_EXP[4],
        CU_DO_OBS[0] / CU_DO_EXP[0], H12),
     "測定済み"),
    ("8", "②居所変更実態調査は18施設からの回答である",
     "区域内の住替え全体を母集団としていない。"
     "在宅から在宅への住替えが含まれない。",
     "解消できない",
     "施設側からみた住替えとして記述する",
     "未解消"),
    ("9", "①②③は配布数の記録がなく回収率を算定できない",
     "配布の記録がないため事後に算定できない。"
     "訪問介護13事業所・居宅介護支援13事業所については"
     "公表データで母数を置き換えられる。",
     "解消できない（訪問介護を除く）",
     "母数を置き換えられるもののみ回収率を示す",
     "一部解消"),
    ("10", "令和7年国勢調査の世帯の家族類型が未公表である",
     "人口等基本集計の公表が令和8年11月頃の見込みである。",
     "令和7年国勢調査 人口等基本集計",
     "令和2年国勢調査による。"
     "公表後に更新する旨を注記した",
     "公表待ち"),
    ("11", "介護老人保健施設の介護職員数が"
     "見える化システムと一致しない",
     "「介護職員」の範囲の解釈が確認できていない。"
     "調査77人に対し見える化は令和6年度59人である。",
     "職種の範囲の定義の確認",
     "［要確認］として注記を残す",
     "未解消"),
    ("12", "同一法人が施設・通所系と訪問系の両方に"
     "同一の職員13人を記入している",
     "個票の突合により重複は特定したが、"
     "総数をどちらに計上するかの決定を要する。",
     "総数の取扱いの決定",
     "［要確認］として注記を残す",
     "決定待ち"),
    ("13", "②の在籍者を保険者の受給者と同じものとして扱えない"
     "（11シートNo.15）",
     "数える対象が異なるため一致しない。"
     "年報は施設の所在地を問わず当広域連合の被保険者を数え、"
     "調査②は保険者を問わず区域内の施設の在籍者を数える。",
     "年報（令和7年度）の要介護度別の明細により"
     "区分ごとに受給者数を実数で置けるようになった",
     "10シートで在籍と受給を並べ、"
     "数える対象の違いを注記した。"
     "区域内の定員を必要利用定員総数の上限として扱わない"
     "（確認事項No.88・No.141）",
     "測定済み"),
    ("14", "①の代表性を判定できない（11シートNo.3の一部）",
     "配布数の記録はないままであるが、"
     "回答の構成が母集団と食い違うかどうかは判定できるようになった。",
     "年報（令和7年度）の要介護度別の明細と、"
     "施設・居住系の要介護度別の差引きによる在宅の認定者の分布",
     "χ²検定により代表しないことを示したうえで、"
     "「在宅で困難を抱える方の状況」として引く"
     "（確認事項No.140）",
     "測定済み"),
]
HOKAN_ST = collections.Counter(a[-1] for a in
                               HOKAN_KAISHO + HOKAN_ZANTEI)
HOKAN_SUM = "・".join("%s%d件" % (k, v) for k, v in HOKAN_ST.items())

ws = sheet("00_報告書の構成と調査の概要", "アンケート調査の集計分析報告書",
           "業務仕様書４（3）が定める4調査の集計・分析の結果をまとめる。"
           "①②③は令和8年8月4日に第10期分を受領し、"
           "④は令和8年7月に個票データを受領した。"
           "4調査を横断したクロス集計は、"
           "利用者票の所在地区の記入形式が統一されておらず個票を地区に"
           "割り付けられないため、発注者のご意向により行わないこととし、"
           "各調査内のクロス集計と公表データによる供給構造の分析により"
           "調査相互を接続する（13シート）。"
           "初版作成日：令和8年8月5日／最終更新日：令和8年9月25日。",
           [4, 26, 52, 44, 20], freeze="A5")

r = header(ws, 4, ["No.", "シート", "内容", "主な結果", "備考"])
for i, (nm, cont, res, rem) in enumerate([
    ("01_調査の実施状況と回収", "4調査の実施時期・対象・回収状況を整理する。",
     "①事業所票15件・利用者票99票、②施設等票18件、"
     "③事業所票27件・職員個票317人・職員票26件、④個票4,729票。",
     "回収率は③のみ算定できる"),
    ("02_①在宅生活改善調査の結果", "在宅生活の継続を困難にしている要因と"
     "必要な生活支援を集計する。",
     "在宅生活の維持が困難%d人（有効回答%d票の%.1f％）。"
     "必要な生活支援は外出同行51件・見守り34件・通いの場34件。"
     % (IJI_KONNAN, IJI_YUKO, H06) +
     "より適切なサービスは小規模多機能44件が最多。", "第2章第4節の根拠"),
    ("03_②居所変更実態調査の結果", "施設・居住系の入退所の実態を集計する。",
     "18施設・定員652人・入所者536人（入所欄の合計。記入のない1施設を"
     "要介護度別の内訳58人で補うと594人）。"
     "新規入所335人の入所前は病院・診療所57.0％・自宅27.5％。"
     "退去349人の退去先は病院・診療所47.0％・死亡27.2％。",
     "第2章第6節・第6章第5節の根拠"),
    ("04_③介護人材実態調査の結果", "介護人材の確保・定着の実態を集計する。",
     "介護職員%d人（重複を除く）。介護福祉士67.1％、"
     "勤続1年未満19.9％、"
     "外国人職員は施設・通所系の介護職員%d人の%.1f％。"
     "令和6年度の採用%d人・離職%d人。"
     % (N_ALL, N_SIS, GAIKOKU / N_SIS * 100, SAIYO, RISHOKU),
     "第2章第6節・第5章基本目標4の根拠"),
    ("05_④健康とくらしの調査の結果", "一般高齢者の健康・社会参加・"
     "生活支援ニーズを集計する。別冊24シートの要約である。",
     "フレイル該当19.1％。小学校区12地区別・年齢調整による地域比較を実施。"
     "同規模保険者40との比較で外出と社会参加が最大の課題。",
     "別冊「調査クロス集計・分析」24シート"),
    ("06_供給構造①_事業所数と定員", "北海道の名簿・見える化K系列により"
     "区域内の事業所数と定員を確定する。",
     "指定事業所は延べ124件・実%d事業所。"
     "施設%d人・居住系%d人・介護保険外の住まい%d人（戸）。"
     % (N_JIG, CAP_TOKUYO + CAP_CHITOKU + CAP_ROKEN,
        CAP_GH + CAP_TOKUTEI, CAP_JUTAKU + CAP_SAKO + CAP_KEIHI),
     "調査の未回答を補う"),
    ("07_供給構造②_サービスの実施地域", "運営規程の「通常の事業実施地域」により"
     "町ごとに使えるサービスを整理する。",
     "訪問介護は美瑛町を対象とするのが4事業所と最も薄い。"
     "小規模多機能は東川町・東神楽町を対象とするのが1事業所のみ。",
     "第1章第7節の根拠"),
    ("08_供給構造③_運営法人の集中", "実%d事業所を運営する%d法人の分布を見る。"
     % (N_JIG, N_HOJIN),
     "上位6法人が%d事業所（%.1f％）を運営する。"
     "小規模多機能5事業所はすべて社会福祉法人美瑛慈光会である。"
     % (sum(v for _, v in TOP6), sum(v for _, v in TOP6) / N_JIG * 100),
     "調査結果の偏りの背景"),
    ("09_供給構造④_従事者数", "調査の介護職員数と見える化M2系列を突合する。",
     "見える化に従事者数の系列がないサービスの介護職員116人（33.3％）は"
     "本調査が唯一の把握手段である。", "第2章第6節の根拠"),
    ("10_供給と需要の対照",
     "供給量とサービス見込量の算定 第1次概算を対照する"
     "（計画素案 第6章と同じ値）。",
     "施設・居住系は合計では定員内に収まるが、"
     "地域密着型介護老人福祉施設は令和11年度に定員62人の99.9％となる。"
     "認定者の24.5％（486人）がサービスを利用していない。"
     "在籍と受給は数える対象が異なり、介護老人保健施設だけが"
     "在籍%d人＞受給%.1f人／月である。"
     % (ROKEN_Z, ROKEN_J),
     "第6章第2節・第4節の根拠"),
    ("11_調査結果の限界と留保", "本報告書の数値を計画本文に用いる際の留保を"
     "整理する。",
     "母集団・時点・回収状況の3点に留保がある。"
     "点検事項35件のうち重大6件は取扱いの決定を要する。"
     "調査①が在宅の認定者全体を代表しないことは"
     "χ²＝%.1f（自由度6。1％点%.2f）として測定済みである。" % (CHI, CHI_1P),
     "必読"),
    ("12_主要所見と計画本文への反映", "所見と計画素案の反映先を対応させる。"
     "反映済みかどうかは計画素案の実物を読んで判定する。",
     "所見20件。反映済みのものは計画素案から機械的に判定する。",
     "反映の設計図"),
    ("13_横断クロス集計を行わないことの整理",
     "4調査横断のクロス集計を行わない理由と、"
     "これに代わる接続の方法を記録する。",
     "所在地区の記入形式が9種類に分かれ個票を地区に割り付けられない。"
     "供給構造の分析により調査相互を接続する。", "発注者のご意向による"),
    ("14_公表データによる補完と留保の解消",
     "11シートの留保のうち、その後に取得した公表データ及び"
     "外部統計により解消したもの・しなかったものを整理する。",
     HOKAN_SUM + "。"
     "訪問介護13事業所181人の全数把握により、"
     "訪問系の代表性の留保が解消した。"
     "年報の要介護度別の明細を確認できたことにより、"
     "調査①の偏りと、在籍と受給の違いを測定できるようになった。",
     "更新の記録"),
]):
    r = body(ws, r, [i + 1, nm, cont, res, rem], height=52,
             align={1: "center"})

r += 1
r = lead(ws, r, "【4調査の位置づけ】", 5)
r = header(ws, r, ["調査", "対象", "把握するもの", "計画での用途", "実施時期"])
for nm, tgt, what, use, when in [
    ("① 在宅生活改善調査", "居宅介護支援事業所・小規模多機能・地域包括の"
     "利用者（要介護者本人・家族）",
     "在宅生活の継続を困難にしている要因、必要な生活支援、"
     "より適切と思われるサービス",
     "第2章第4節、第5章基本目標2、代表KPI H06", "令和7年度"),
    ("② 居所変更実態調査", "施設・居住系の事業所",
     "入所前の居場所、退去先、退去理由、待機者、受入可能な医療処置",
     "第2章第6節、第6章第2節・第4節・第5節、代表KPI H11", "令和7年4月1日現在"),
    ("③ 介護人材実態調査", "介護サービス事業所とその職員",
     "介護職員数、資格、雇用形態、勤続年数、採用・離職、外国人職員",
     "第2章第6節、第5章基本目標4、代表KPI H13・H14", "令和7年4月1日現在"),
    ("④ 健康とくらしの調査", "65歳以上の一般高齢者及び総合事業対象者"
     "（要支援・要介護認定者を含まない）",
     "フレイル、社会参加、閉じこもり、うつ、IADL、生活支援ニーズ",
     "第2章第4節、第5章基本目標1、代表KPI H04・H05",
     "令和7年11月17日〜12月8日"),
]:
    r = body(ws, r, [nm, tgt, what, use, when], height=48,
             align={5: "center"})

note(ws, r + 1,
     "注1）本報告書は受託者による集計・分析の結果であり、"
     "計画本文への反映は発注者のご意向を確認したうえで行う。"
     "注2）個票データ（在宅生活改善調査の利用者票99票、"
     "介護人材実態調査の職員個票317人、健康とくらしの調査の個票4,729票）は"
     "個人情報を含むため、成果品には集計値のみを収録する。"
     "注3）本報告書の点検の詳細は別冊"
     "「実施済み3調査の受領点検と集計」（16シート）による。"
     "④の詳細は別冊「調査クロス集計・分析」（24シート）による。"
     "注4）本報告書は初版を令和8年8月5日に作成し、"
     "令和8年9月25日に更新した。"
     "見込量は計画素案 第6章（サービス見込量の算定 第1次概算。"
     "令和8年9月16日）と同じ値であり、"
     "10シートの在籍と受給の対照及び11シートの留保は"
     "年報（令和7年度）の要介護度別の明細により更新している。"
     "点検事項の取扱いの決定により数値が変わる箇所は"
     "11シート及び12シートに示す。", 5)

# ============================================================ 01 回収
ws = sheet("01_調査の実施状況と回収", "調査の実施状況と回収",
           "4調査の対象、配布・回収の状況を整理する。"
           "①②③は配布数の記録がないため回収率を算定できない。"
           "③は前回（令和5年）に回収率77.1％の記録がある。",
           [22, 14, 14, 14, 12, 44], freeze="A5")

r = header(ws, 4, ["調査・票種", "受領件数", "母数（公表データ等）",
                   "回収の状況", "町別", "備考"])
for nm, n, bo, st, tw, memo in [
    ("① 事業所票", "15件", "居宅介護支援13・小規模多機能5・地域包括2",
     "居宅介護支援10・小規模多機能3・地域包括2", "―",
     "北海道の指定事業所一覧では居宅介護支援は13事業所である。"
     "母数を13とすると回収率76.9％となる"),
    ("① 利用者票", "99票", "―", "10事業所から提出", "集計不能",
     "所在地区の記入形式が9種類に分かれ、地区別の集計ができない"
     "（点検事項No.3）"),
    ("② 施設等票", "18件", "介護保険の指定施設19＋介護保険外の住まい12＝31",
     "指定施設14・介護保険外4", "―",
     "13施設が未回答（指定5・指定外8。把握率58.1％）。"
     "うち8施設は介護保険外の住まいである（点検事項No.30・No.32・No.34）"),
    ("③ 事業所票（施設・通所系）", "24件", "―", "―", "―",
     "施設・居住系・通所系の事業所"),
    ("③ 事業所票（訪問系）", "3件", "訪問介護13事業所（見える化K3a・名簿とも）",
     "回収率23.1％", "―",
     "訪問系の母数が確定した。回収率が低く、"
     "訪問系の職員数は区域全体を代表しない（点検事項No.23）"),
    ("③ 職員個票", "317人", "施設・通所系24事業所の介護職員319人",
     "99.4％", "―", "2人分の個票が不足（点検事項No.15〜No.19）"),
    ("③ 職員票（訪問系）", "26件", "訪問系3事業所の介護職員42人",
     "61.9％", "―", "訪問系の職員票は個票と様式が異なる"),
    ("④ 健康とくらしの調査", "4,729票", "65歳以上7,121人",
     "回収4,798票・回収率67.4％（分析対象4,729票）", "3町別に集計可",
     "小学校区12地区別の集計ができる。"
     "要支援・要介護認定者は対象に含まれない"),
]:
    r = body(ws, r, [nm, n, bo, st, tw, memo], height=44,
             align={2: "center", 4: "center", 5: "center"})

note(ws, r + 1,
     "注1）①②③は発注者が実施した調査であり、"
     "受託者は配布・回収に関与していない（仕様書４（3））。"
     "配布先の一覧が確認できないため、回収率は母数を公表データで"
     "置き換えられるものに限り算定した。"
     "注2）③の訪問系は13事業所のうち3事業所からの回答であり、"
     "回収率23.1％である。"
     "訪問系の介護職員42人は区域内の訪問系職員の一部にすぎない。"
     "注3）④は要支援者・要介護者を含まない。"
     "①②が要介護者・施設入所者を対象とするのと母集団が異なるため、"
     "両者の数値を足し合わせたり比率を比較したりしない（11シート）。", 6)

# ============================================================ 02 ①
ws = sheet("02_①在宅生活改善調査の結果", "① 在宅生活改善調査の結果",
           "利用者票99票により、在宅生活の継続を困難にしている要因と"
           "必要な生活支援を集計する。"
           "事業所票15件により、居場所の変更の実績を集計する。",
           [30, 12, 12, 12, 12, 40], freeze="A5")

r = lead(ws, 4, "【在宅生活の維持の見通し（有効回答%d票）】" % IJI_YUKO, 6)
r = header(ws, r, ["区分", "件数", "構成比", "", "", "見方"])
IJI = S.RIYO["生活の維持"]
tot = sum(v for k, v in IJI.items() if k != "-")
for code, lab, view in [
    ("1", "1 現在の状態では在宅生活の維持が困難",
     "回答の大半を占める。事業所が「困難」と判断した利用者を"
     "抽出して回答する設計であるため、"
     "区域内の在宅利用者全体の割合ではない"),
    ("2", "2 当面は在宅生活を維持できる", "―"),
]:
    v = IJI.get(code, 0)
    r = body(ws, r, [lab, v, "%.1f％" % (v / tot * 100), "", "", view],
             height=32, align={2: "right", 3: "right"})
    ws.merge_cells(start_row=r - 1, start_column=4, end_row=r - 1, end_column=5)
r = body(ws, r, ["有効回答", tot, "100.0％", "", "", ""], {1: GRAY},
         height=20, align={2: "right", 3: "right"}, bold=True)
r = body(ws, r, ["無回答", IJI.get("-", 0), "―", "", "",
                 "回収票%d票との差。構成比の分母は有効回答%d票である"
                 % (RIYO_N, tot)], height=20,
         align={2: "right", 3: "right"})

r += 1
r = lead(ws, r, "【必要な生活支援（複数回答・99票）】", 6)
r = header(ws, r, ["生活支援", "件数", "回答者に対する割合", "", "", "見方"])
SUP = S.RIYO["必要な生活支援"]
for k, v in sorted(SUP.items(), key=lambda x: -x[1]):
    view = ""
    if k.startswith("外出同行"):
        view = "最多。移動の確保が在宅生活の継続の鍵である。" \
               "健康とくらしの調査でも自力移動の手段がない層が" \
               "20％台であることと符合する"
    elif k == "見守り、声かけ":
        view = "介護保険の給付では担いにくい。生活支援体制整備事業の対象"
    elif k == "サロンなどの定期的な通いの場":
        view = "通いの場は介護予防だけでなく在宅生活の継続の資源でもある"
    r = body(ws, r, [k, v, "%.1f％" % (v / RIYO_N * 100), "", "", view],
             height=24, align={2: "right", 3: "right"})
    ws.merge_cells(start_row=r - 1, start_column=4, end_row=r - 1, end_column=5)

r += 1
r = lead(ws, r, "【より適切と思われるサービス（複数回答・99票）】", 6)
r = header(ws, r, ["サービス", "件数", "回答者に対する割合", "在宅/施設等",
                   "", "見方"])
SVC = S.RIYO["より適切なサービス"]
ZAITAKU = {"ｼｮｰﾄｽﾃｲ", "訪問介護､訪問入浴", "夜間対応型訪問介護", "訪問看護",
           "訪問ﾘﾊ", "通所介護､通所ﾘﾊ､認知症対応型通所", "定期巡回ｻｰﾋﾞｽ",
           "小規模多機能", "看護小規模多機能", "訪問診療", "居宅療養管理指導"}
for k, v in sorted(SVC.items(), key=lambda x: -x[1]):
    if v == 0:
        continue
    kb = "在宅" if k in ZAITAKU else "施設等"
    view = ""
    if k == "小規模多機能":
        view = "最多。ただし区域内の小規模多機能5事業所はすべて" \
               "社会福祉法人美瑛慈光会が美瑛町で運営しており、" \
               "うち36票が同一法人の2事業所からの提出である" \
               "（点検事項No.2・08シート）"
    elif k == "夜間対応型訪問介護":
        view = "区域内に事業所がない（見える化K3n＝0）"
    elif k == "定期巡回ｻｰﾋﾞｽ":
        view = "区域内に事業所がない（見える化K3m＝0）"
    elif k == "看護小規模多機能":
        view = "区域内に事業所がない（見える化K3q＝0）"
    r = body(ws, r, [k, v, "%.1f％" % (v / RIYO_N * 100), kb, "", view],
             {4: MID_B if kb == "在宅" else IN_Y}, height=30,
             align={2: "right", 3: "right", 4: "center"})

r += 1
r = lead(ws, r, "【区域内に事業所がないサービスへの言及】", 6)
r = header(ws, r, ["サービス", "件数", "見える化\n事業所数", "受給率", "", "見方"])
for nm, key, code in [("夜間対応型訪問介護", "夜間対応型訪問介護", "K3n"),
                      ("定期巡回・随時対応型", "定期巡回ｻｰﾋﾞｽ", "K3m"),
                      ("看護小規模多機能型", "看護小規模多機能", "K3q")]:
    r = body(ws, r, [nm, SVC.get(key, 0), klast(code), "0.0％", "",
                     "事業所がないため利用できない。"
                     "利用者票で必要と回答されていることは、"
                     "第6章第4節（24時間対応サービスの確保方策）の根拠となる"],
             {3: NG_O}, height=32,
             align={2: "right", 3: "right", 4: "right"})
r = body(ws, r, ["計", sum(SVC.get(k, 0) for k in
                           ["夜間対応型訪問介護", "定期巡回ｻｰﾋﾞｽ",
                            "看護小規模多機能"]), "―", "―", "",
                 "延べ26件。上記3サービスの選択件数の合計であり、"
                 "実人数ではない（複数回答のため）"], {1: GRAY},
         height=32, align={2: "right"}, bold=True)

r += 1
r = lead(ws, r, "【代表KPI H12の算定】", 6)
r = header(ws, r, ["項目", "値", "", "", "", "内容"])
for nm, v, naiyo in [
    ("指標名", KP.H12_NAME,
     "当初の定義「%s」の代理指標として新たに定義したもの"
     % KP.H12_KYU_NAME),
    ("対象とする4サービス", "4区分",
     "夜間対応型訪問介護・定期巡回・随時対応型訪問介護看護・"
     "看護小規模多機能型居宅介護・介護医療院。"
     "いずれも区域内に事業所が存在しない"),
    ("分子", "%d票" % KP.H12_BUNSHI,
     "4サービスのいずれかを「より適切と思われるサービス」として"
     "挙げた実人数。**複数を選んでいても1人1票として数える**"
     "（延べ件数ではない）"),
    ("分母", "%d票" % KP.H12_BUNBO,
     "当該設問の有効回答。要介護度の有効回答%d票ではない"
     % KP.DO_YUKO),
    ("基準値", "%.1f％" % KP.H12, "令和7年度"),
    ("母集団", "課題事例の抽出",
     "事業所が在宅生活の継続に課題があると判断した利用者を"
     "抽出して回答する設計であり、"
     "区域内の在宅利用者全体の未充足率ではない（11シートNo.2）"),
]:
    r = body(ws, r, [nm, v, "", "", "", naiyo], height=32,
             align={2: "center"})
    ws.merge_cells(start_row=r - 1, start_column=2, end_row=r - 1, end_column=5)

r += 1
r = lead(ws, r, "【事業所票15件による居場所の変更（171人）】", 6)
r = header(ws, r, ["変更先", "人数", "構成比", "", "", "見方"])
Z3 = collections.Counter()
for *_x, q3 in S.ZAI:
    for k, v in q3.items():
        Z3[k] += v or 0
tz3 = sum(Z3.values())
for k, v in Z3.most_common():
    r = body(ws, r, [k, v, "%.1f％" % (v / tz3 * 100), "", "", ""],
             height=20, align={2: "right", 3: "right"})
r = body(ws, r, ["合計", tz3, "100.0％", "", "", ""], {1: GRAY}, height=20,
         align={2: "right", 3: "right"}, bold=True)

note(ws, r + 1,
     "注1）利用者票は、事業所が在宅生活の継続に課題があると判断した利用者を"
     "抽出して回答する設計である。"
     "したがって「在宅生活の維持が困難%.1f％」は"
     "区域内の在宅利用者全体の割合ではない。"
     "計画本文では「調査対象となった%d人のうち」と母集団を明記する。"
     "注1-2）割合の分母は設問ごとの有効回答による。"
     "在宅生活の維持の見通しは%d票（無回答1票）、"
     "必要な生活支援及びより適切と思われるサービスは%d票である。"
     "要介護度の有効回答も%d票であるが、"
     "これは別の設問の有効回答であって同じ票を指すとは限らない。"
     % (H06, RIYO_N, IJI_YUKO, RIYO_N, IJI_YUKO) +
     "注2）「より適切と思われるサービス」は、"
     "回答者である事業所が自らの提供するサービスを挙げる構造になっている"
     "（点検事項No.2）。"
     "小規模多機能44件のうち31件は同一法人の小規模多機能2事業所からの"
     "提出票によるものである。"
     "計画本文では提出元を除いた集計を併記するか、"
     "件数を示さず「事業所からは小規模多機能・住宅型有料・"
     "グループホームが挙げられた」と質的に記述する。"
     "注3）所在地区の記入形式が統一されていないため、"
     "地区別・圏域別の集計は行わない（点検事項No.3・13シート）。", 6)

# ============================================================ 03 ②
ws = sheet("03_②居所変更実態調査の結果", "② 居所変更実態調査の結果",
           "施設等票18件により、施設・居住系の入退所の実態を集計する。"
           "回答した18施設は区域内の施設・住まい31施設のうちの一部であり"
           "（把握率58.1％）、"
           "定員は北海道の名簿による（06シート）。",
           [26, 12, 12, 12, 12, 42], freeze="A5")

r = lead(ws, 4, "【種別ごとの規模】（調査に回答した18施設）", 6)
r = header(ws, r, ["種別", "施設数", "定員", "入所者", "入所率", "備考"])
CAPD = {}
for s_ in S.SHI:
    g = {1: "住宅型有料・サ高住", 3: "住宅型有料・サ高住", 4: "グループホーム",
         5: "特定施設", 7: "介護老人保健施設", 9: "特養・地域密着型特養",
         10: "特養・地域密着型特養"}.get(s_["種別"], "その他")
    d = CAPD.setdefault(g, {"n": 0, "cap": 0, "res": 0})
    d["n"] += 1
    d["cap"] += s_["定員"] or 0
    d["res"] += s_["入所"] or 0
for g in ["特養・地域密着型特養", "介護老人保健施設", "グループホーム",
          "特定施設", "住宅型有料・サ高住"]:
    d = CAPD[g]
    res = d["res"] + (58 if g == "特養・地域密着型特養" else 0)
    memo = ""
    if g == "特養・地域密着型特養":
        memo = "美瑛慈光園の入所者数が未記入のため要介護度別の内訳58人を加えた"
    elif g == "特定施設":
        memo = "さわやか東神楽館（定員100人）が未回答。" \
               "区域内定員は156人、入居者は公表データを合わせて154人である"
    r = body(ws, r, [g, d["n"], d["cap"], res,
                     "%.1f％" % (res / d["cap"] * 100), memo], height=32,
             align={2: "right", 3: "right", 4: "right", 5: "right"})
tc = sum(d["cap"] for d in CAPD.values())
tr = sum(d["res"] for d in CAPD.values()) + 58
r = body(ws, r, ["計", sum(d["n"] for d in CAPD.values()), tc, tr,
                 "%.1f％" % (tr / tc * 100), ""], {1: GRAY}, height=20,
         align={2: "right", 3: "right", 4: "right", 5: "right"}, bold=True)

# 入所者数の数え方の照合（同じ「入所者」「在籍」の語で3通りの数が現れる）
r += 1
r = lead(ws, r, "【入所者数の数え方の照合】", 6)
r = header(ws, r, ["種別", "入所欄への記入", "要介護度別内訳の合計",
                   "補完後（本報告書の採用値）", "", "採用値の用いどころ"])
_GRP = {1: "住宅型有料・サ高住", 3: "住宅型有料・サ高住", 4: "グループホーム",
        5: "特定施設", 7: "介護老人保健施設", 9: "特養・地域密着型特養",
        10: "特養・地域密着型特養"}
_CNT = {}
for s_ in S.SHI:
    g = _GRP.get(s_["種別"], "その他")
    d = _CNT.setdefault(g, {"nyu": 0, "do": 0})
    d["nyu"] += s_["入所"] or 0
    d["do"] += sum(x or 0 for x in s_["度別"])
_YOTO = {
    "特養・地域密着型特養": "入所率（本シート）は補完後、"
                            "在籍と受給の対照（10シート）は要介護度別内訳",
    "介護老人保健施設": "入所率は入所欄、"
                        "在籍と受給の対照（10シート）は要介護度別内訳。"
                        "5人の差は要介護度別内訳のみに計上された分である",
    "グループホーム": "いずれも同じ",
    "特定施設": "いずれも同じ",
    "住宅型有料・サ高住": "いずれも同じ。介護保険の給付の対象外である",
}
_t = [0, 0, 0]
for g in ["特養・地域密着型特養", "介護老人保健施設", "グループホーム",
          "特定施設", "住宅型有料・サ高住"]:
    d = _CNT[g]
    ho = d["nyu"] + (58 if g == "特養・地域密着型特養" else 0)
    fl = {2: IN_Y} if d["nyu"] != d["do"] else {}
    r = body(ws, r, [g, d["nyu"], d["do"], ho, "", _YOTO[g]], fl, height=32,
             align={2: "right", 3: "right", 4: "right"})
    ws.merge_cells(start_row=r - 1, start_column=5, end_row=r - 1, end_column=5)
    _t[0] += d["nyu"]
    _t[1] += d["do"]
    _t[2] += ho
r = body(ws, r, ["計", _t[0], _t[1], _t[2], "",
                 "入所欄の合計%d人・要介護度別内訳の合計%d人・"
                 "補完後%d人の3通りがある" % tuple(_t)], {1: GRAY},
         height=32, align={2: "right", 3: "right", 4: "right"}, bold=True)
note(ws, r,
     "注）「入所者」「在籍」の語で3通りの数が現れるため、"
     "本表で対照する。"
     "①入所欄への記入は施設が入所者数の欄に記入した人数であり、"
     "特養1施設が未記入である。"
     "②要介護度別内訳の合計は要介護度別の人数の和であり、"
     "未記入の施設も内訳は記入している。"
     "③補完後は①に当該施設の内訳58人を加えたものである。"
     "介護老人保健施設は①%d人・②%d人で%d人の差があり、"
     "要介護度別の内訳のみに計上された分である。"
     "要介護度別に扱う集計（10シートの在籍と受給の対照、"
     "居所変更実態調査の要介護度構成）は②による。"
     % (_CNT["介護老人保健施設"]["nyu"], _CNT["介護老人保健施設"]["do"],
        _CNT["介護老人保健施設"]["do"] - _CNT["介護老人保健施設"]["nyu"]), 6)
r += 1

r += 1
r = lead(ws, r, "【新規入所者の入所前の居場所（335人）】", 6)
r = header(ws, r, ["入所前の居場所", "人数", "構成比", "", "", "見方"])
IN_ = collections.Counter()
for g, d in C.CS["種別×入所前の居場所"].items():
    for k, v in d.items():
        IN_[k] += v
ti = sum(IN_.values())
for k, v in IN_.most_common():
    view = ""
    if "病院" in k:
        view = "最多。医療機関からの退院が施設入所の主要な経路である。" \
               "第6章第5節（医療と介護の連携）の根拠"
    elif k.startswith("自宅"):
        view = "在宅からの直接入所。第2章第4節と接続する"
    r = body(ws, r, [k, v, "%.1f％" % (v / ti * 100), "", "", view],
             height=24, align={2: "right", 3: "right"})
    ws.merge_cells(start_row=r - 1, start_column=4, end_row=r - 1, end_column=5)
r = body(ws, r, ["合計", ti, "100.0％", "", "", ""], {1: GRAY}, height=20,
         align={2: "right", 3: "right"}, bold=True)

r += 1
r = lead(ws, r, "【種別ごとの区域内・区域外の別（新規入所）】", 6)
r = header(ws, r, ["種別", "区域内", "区域外", "区域外の割合", "", "見方"])
for g, d in C.CS["種別×入所前の居場所"].items():
    inn = sum(v for k, v in d.items() if "区域内" in k or "町内" in k)
    out = sum(v for k, v in d.items() if "区域外" in k or "町外" in k)
    if inn + out == 0:
        continue
    view = ""
    if g == "グループホーム" or g == "特定施設":
        view = "区域外からの入所が大半を占める。" \
               "住所地特例により保険者は従前の市町村のままとなる。" \
               "区域内の定員が区域内の被保険者のために使われているとは限らない"
    r = body(ws, r, [g, inn, out, "%.1f％" % (out / (inn + out) * 100), "",
                     view], height=32,
             align={2: "right", 3: "right", 4: "right"})
    ws.merge_cells(start_row=r - 1, start_column=5, end_row=r - 1, end_column=5)

r += 1
r = lead(ws, r, "【未回答の特定施設を公表データで補った場合】", 6)
r = header(ws, r, ["項目", "調査（2施設）\n令和7年4月1日現在",
                   "公表データ\n（さわやか東神楽館）\n令和7年10月6日現在",
                   "計", "北海道の名簿による定員", "見方"])
_TK = [x for x in S.SHI if x["種別"] == 5]
_TK_CAP = sum(x["定員"] or 0 for x in _TK)          # 調査の自己申告 58
_TK_RES = sum(x["入所"] or 0 for x in _TK)          # 57
_SW_CAP, _SW_RES = SW["利用定員"], SW_N             # 100 / 97
r = body(ws, r, ["定員", _TK_CAP, _SW_CAP, _TK_CAP + _SW_CAP, CAP_TOKUTEI,
                 "調査の自己申告の合計は%d人で、"
                 "北海道の名簿による区域内3施設の定員%d人と%+d人ちがう"
                 "（有料老人ホームゆうが調査22人・名簿20人）"
                 % (_TK_CAP + _SW_CAP, CAP_TOKUTEI,
                    _TK_CAP + _SW_CAP - CAP_TOKUTEI)],
         {5: IN_Y}, height=32,
         align={2: "right", 3: "right", 4: "right", 5: "right"})
r = body(ws, r, ["入居者", _TK_RES, _SW_RES, _TK_RES + _SW_RES, "―", ""],
         height=22, align={2: "right", 3: "right", 4: "right", 5: "right"})
r = body(ws, r, ["入居率",
                 "%.1f％" % (_TK_RES / _TK_CAP * 100),
                 "%.1f％" % (_SW_RES / _SW_CAP * 100),
                 "%.1f％（参考）" % ((_TK_RES + _SW_RES)
                                     / (_TK_CAP + _SW_CAP) * 100),
                 "%.1f％（参考）" % ((_TK_RES + _SW_RES) / CAP_TOKUTEI * 100),
                 "施設ごとにみるといずれも高い入居率である。"
                 "計の欄は時点の異なる値を合算した参考値であり、"
                 "区域内の入居率の確定値としては用いない"],
         {4: IN_Y, 5: IN_Y, 6: IN_Y}, height=44,
         align={2: "right", 3: "right", 4: "right", 5: "right"})
note(ws, r,
     "注）調査に回答した2施設は令和7年4月1日現在、"
     "未回答の1施設は介護サービス情報公表システムの"
     "令和7年10月6日現在の公表値である。"
     "同システムの「入居率」の欄は100.0％と表示されているが、"
     "同じ画面の入居者数（男25人・女72人。要介護度別の和も同数）は"
     "%d人であり、定員%d人に対して%.1f％となる。"
     "公表資料の中で食い違っているため、"
     "本報告書では入居者数から計算した%.1f％を用いる（点検事項No.32）。"
     % (_SW_RES, _SW_CAP, _SW_RES / _SW_CAP * 100,
        _SW_RES / _SW_CAP * 100) +
     "時点が異なるため、両者を足した入居率は参考値である。"
     "分母を調査の自己申告の合計%d人とすると%.1f％、"
     "北海道の名簿による区域内定員%d人とすると%.1f％となる。"
     "居住系の定員に余裕があるという見方は、"
     "個々の施設の空きを意味しない。"
     % (_TK_CAP + _SW_CAP,
        (_TK_RES + _SW_RES) / (_TK_CAP + _SW_CAP) * 100,
        CAP_TOKUTEI, (_TK_RES + _SW_RES) / CAP_TOKUTEI * 100), 6)
r += 1

r += 1
r = lead(ws, r, "【さわやか東神楽館の入居者の要介護度（97人）】", 6)
r = header(ws, r, ["要介護度", "人数", "構成比", "", "", "見方"])
_kd = SW["要介護度別入居者数"]
for k, v in _kd.items():
    if v == 0 and k == "自立":
        continue
    view = ""
    if k == "要介護1":
        view = "単独で最多。認定者全体で要介護1に集中していることと同じ方向"
    r = body(ws, r, [k, v, "%.1f％" % (v / SW_N * 100), "", "", view],
             height=20, align={2: "right", 3: "right"})
    ws.merge_cells(start_row=r - 1, start_column=4, end_row=r - 1, end_column=5)
_light = _kd["要支援1"] + _kd["要支援2"] + _kd["要介護1"]
r = body(ws, r, ["要支援1〜要介護1", _light, "%.1f％" % (_light / SW_N * 100),
                 "", "",
                 "【参考】本施設では軽度者が7割を占める。"
                 "特定施設が重度者の受け皿として機能しているとは限らない。"
                 "要支援者が26人入居しており、"
                 "居住系の見込量を要介護度別に置く際に要支援を無視できない。"
                 "1施設の値であり、区域内3施設の構成を示すものではない"],
         {2: IN_Y}, height=40, align={2: "right", 3: "right"}, bold=True)

note(ws, r + 1,
     "注1）本調査に回答したのは18施設であり、"
     "区域内の施設・住まい31施設のうち13施設が未回答である"
     "（把握率58.1％。06シート）。"
     "とくに区域内最大の特定施設（さわやか東神楽館・定員100人）が"
     "回答していないため、入所前の居場所・退去先の集計は"
     "定員56人分の2施設に基づく（点検事項No.32）。"
     "定員・入居者数・要介護度別の内訳・従業者数は"
     "介護サービス情報公表システムの個別画面（記入日 令和7年10月6日）により"
     "補うことができるが、入所前の居場所と退去先は公表されていない。"
     "注2）入所率は調査に回答した施設の定員に対するものである。"
     "区域内の定員は北海道の名簿による（06シート）。"
     "注3）区域内・区域外の別は、調査票の「市内／市外」の選択肢による。"
     "「市内」が町単位か広域連合単位かが定義されていないため、"
     "介護サービス自給率（δ）との接続には確認を要する（点検事項No.25）。", 6)

# ============================================================ 04 ③
ws = sheet("04_③介護人材実態調査の結果", "③ 介護人材実態調査の結果",
           "事業所票27件・職員個票317人・職員票26件により、"
           "介護人材の確保・定着の実態を集計する。"
           "施設・通所系と訪問系で様式が異なるため、区分して示す。",
           [26, 12, 12, 12, 12, 42], freeze="A5")

r = lead(ws, 4, "【介護職員数】", 6)
r = header(ws, r, ["区分", "事業所数", "介護職員", "常勤", "非常勤", "備考"])
for nm, lst in [("施設・通所系", SIS), ("訪問系", HOU)]:
    r = body(ws, r, [nm, len(lst), sum(j["職員"] or 0 for j in lst),
                     sum(j["常勤"] or 0 for j in lst),
                     sum(j["非常勤"] or 0 for j in lst),
                     "訪問系は13事業所のうち3事業所からの回答である"
                     if nm == "訪問系" else ""], height=24,
             align={2: "right", 3: "right", 4: "right", 5: "right"})
r = body(ws, r, ["単純合計", len(S.JIN), N_SIS + N_HOU, "", "", ""],
         height=20, align={2: "right", 3: "right"})
r = body(ws, r, ["重複を除く", len(S.JIN) - 1, N_ALL, "", "",
                 "有料老人ホーム華とヘルパーステーションフラワーが"
                 "同一の職員13人を計上している。"
                 "住宅型有料老人ホームは介護保険の指定サービスではなく"
                 "見える化に現れないため、訪問系13人を実数とする"
                 "（受託者の案・点検事項No.1）"], {3: OK_G}, height=44,
         align={2: "right", 3: "right"}, bold=True)

r += 1
r = dist(ws, r, "【職員個票　資格の取得・研修の修了（317人）】", SK["資格"],
         {"1": "1 介護福祉士", "2": "2 実務者研修", "3": "3 初任者研修",
          "4": "4 いずれも該当しない"}, 6,
         note_="介護福祉士は有効回答316人の67.1％である。"
               "北海道の評価指標及び第10期基本指針案が求める"
               "介護職員の質の指標に対応する。")

r = dist(ws, r, "【職員個票　雇用形態（317人）】", SK["雇用形態"],
         {"1": "1 常勤", "2": "2 非常勤"}, 6)

r = dist(ws, r, "【職員個票　現在の事業所での勤務年数（317人）】",
         SK["勤務年数"], {"1": "1 1年以上", "2": "2 1年未満"}, 6,
         note_="勤続1年未満は有効回答316人の19.9％である。"
               "代表KPI H14は「採用率と離職率の差」であり、"
               "本値は定着の補足指標として併せて見るものである。")

r += 1
r = lead(ws, r, "【雇用形態 × 勤務年数（316人）】", 6)
r = header(ws, r, ["雇用形態", "1年以上", "1年未満", "計", "1年未満の割合",
                   "見方"])
for lab, a, b, view in [
    ("常勤", 206, 47, "常勤の18.6％が勤続1年未満である"),
    ("非常勤", 47, 16, "非常勤の25.4％が勤続1年未満であり、常勤より高い"),
]:
    r = body(ws, r, [lab, a, b, a + b, "%.1f％" % (b / (a + b) * 100), view],
             height=24, align={2: "right", 3: "right", 4: "right",
                               5: "right"})
r = body(ws, r, ["計", 253, 63, 316, "19.9％",
                 "勤続1年未満63人のうち47人（74.6％）は常勤である。"
                 "定着の課題は非常勤に限られない"], {1: GRAY}, height=32,
         align={2: "right", 3: "right", 4: "right", 5: "right"}, bold=True)

r += 1
r = dist(ws, r, "【職員個票　年齢（317人）】", SK["年齢"],
         {"1": "1 20歳未満", "2": "2 20歳代", "3": "3 30歳代",
          "4": "4 40歳代", "5": "5 50歳代", "6": "6 60歳代",
          "7": "7 70歳以上"}, 6,
         note_="60歳以上は42人（13.3％）である。"
               "40歳代が77人（24.4％）と最も多い。")

r += 1
r = lead(ws, r, "【採用と離職（令和6年度・重複を除く）】", 6)
r = header(ws, r, ["区分", "人数", "割合（分母は下欄のとおり）", "", "",
                   "見方"])
for nm, v, view in [
    ("採用者数", SAIYO, "分母は回答全体の介護職員%d人" % N_ALL),
    ("離職者数", RISHOKU, "分母は回答全体の介護職員%d人" % N_ALL),
    ("差（採用－離職）", SAIYO - RISHOKU,
     "北海道の評価指標⑥及び代表KPI H14は「採用率と離職率の差」を用いるが、"
     "各年9月30日現在の在籍者数を要する。"
     "本調査は令和7年4月1日現在であるため、そのままでは算定できない"),
    ("外国人職員", GAIKOKU,
     "%d事業所に配置されている。"
     "いずれも施設・通所系であり訪問系3事業所には配置がない。"
     "施設・通所系の介護職員%d人に対する割合は%.1f％である"
     "（訪問系を含む回答全体%d人を分母とすると%.1f％）"
     % (N_GAI_JIG, N_SIS, GAIKOKU / N_SIS * 100,
        N_ALL, GAIKOKU / N_ALL * 100)),
]:
    # 外国人職員は全員が施設・通所系であるため分母を施設・通所系とする。
    _bunbo = N_SIS if nm == "外国人職員" else N_ALL
    pct = "%.1f％" % (v / _bunbo * 100) if nm != "差（採用－離職）" else "―"
    r = body(ws, r, [nm, v, pct, "", "", view], height=32,
             align={2: "right", 3: "right"})
    ws.merge_cells(start_row=r - 1, start_column=4, end_row=r - 1, end_column=5)

note(ws, r + 1,
     "注1）介護職員数は、事業所票の「介護職員数」の合計である。"
     "職種別（看護職員・生活相談員等）の内訳は調査票にない。"
     "見える化M2系列は職種別の実人員であるため、"
     "突合の際は範囲の違いに留意する（09シート）。"
     "注2）訪問系は13事業所のうち3事業所からの回答であり、"
     "訪問系の介護職員42人は区域内の訪問系職員の一部にすぎない。"
     "訪問系を含む合計値を「区域内の介護職員数」と記述しない。"
     "注3）勤務年数と雇用形態は周辺度数（253人・63人）が一致するが、"
     "個票レベルでは一致しない（上表のとおり）。"
     "同じ数値が並ぶことによる転記の誤りに留意する。"
     "注4）介護老人保健施設3施設の介護職員77人は"
     "見える化M2b（令和6年度59人）と一致しない（点検事項No.31）。"
     "通所リハビリテーションも同様である。"
     "この2サービスは全施設が回答したうえで差が生じている。", 6)

# ============================================================ 05 ④
ws = sheet("05_④健康とくらしの調査の結果", "④ 健康とくらしの調査の結果",
           "個票4,729票の分析結果の要約である。"
           "詳細は別冊「調査クロス集計・分析」（24シート）による。"
           "本調査は一般高齢者及び総合事業対象者を対象とし、"
           "要支援者・要介護者を含まない。",
           [26, 14, 14, 46, 24], freeze="A5")

r = header(ws, 4, ["領域", "主な指標", "値", "所見", "反映先"])
for area, ind, val, fin, dest in [
    ("心身の状態", "フレイル該当割合", "19.1％",
     "代表KPI H04の基準値。町別では美瑛町21.6％・東神楽町17.0％だが、"
     "年齢調整後は差が縮む。町間比較は年齢調整後の値による。",
     "第2章第4節\n第4章第3節"),
    ("心身の状態", "1年間の転倒あり", "35.5％",
     "同規模保険者40（30.0％）を5.5ポイント上回る。"
     "65〜69歳で同規模を9.6ポイント上回り、若い年齢層で顕著である。",
     "第5章 基本目標1"),
    ("心身の状態", "口腔機能低下", "24.7％",
     "同規模保険者40（21.8％）を上回る。", "第5章 基本目標1"),
    ("社会参加", "友人知人と会う頻度が高い", "63.0％",
     "同規模保険者40（71.2％）を8.2ポイント下回る。"
     "外出と社会参加が最大の課題である。",
     "第2章第4節\n第5章 基本目標1・2"),
    ("社会参加", "通いの場への参加", "―",
     "参加者は独居30.3％・生活動作の困りごとあり54.6％と、"
     "支援の必要度が高い層の割合が高い。"
     "通いの場は既にハイリスク層に届いている。",
     "第5章 基本目標1"),
    ("生活支援", "生活動作の困りごと：除雪", "24.3％",
     "困りごとの最上位。介護保険の給付対象外であり、"
     "生活支援体制整備事業及び3町の施策で対応する。",
     "第5章 基本目標2"),
    ("生活支援", "困りごとを解決できず困っている", "0.6％",
     "代表KPI H06（在宅生活継続困難割合）の補完指標である。"
     "H06の基準値は①在宅生活改善調査による" + KP.h06_str() + "。"
     "本調査とは母集団・設問・分母がいずれも異なるため、"
     "同じH06の基準値として置き換えることはできない。",
     "第4章第3節"),
    ("介護・介助", "介護・介助が必要だが受けていない", "7.0％",
     "認定を受けながらサービスを利用していない層とは別の集団である。"
     "合算しない。", "第2章第2節\n第5章 基本目標1・2"),
    ("支え手", "世話をしてくれる人がいない", "8.7％",
     "独居者では32.4％。身寄りのない高齢者の支援体制の対象規模を示す。",
     "第5章 基本目標3"),
    ("意向", "自宅での介護を希望", "49.2％",
     "在宅サービスの見込量の前提となる。"
     "一方、24時間対応の3サービスが区域内に存在しない。",
     "第6章第2節・第4節"),
    ("認知症", "相談窓口を知らない", "64.3％",
     "本人又は家族に症状がある者に限っても47.3％が知らない。"
     "見える化に認知症指標のデータ登録がないため本設問が補完指標となる。",
     "第5章 基本目標2"),
    ("地域資源", "連帯感（ソーシャルキャピタル）", "美瑛町14位・東川町15位",
     "参加74市町村中で上位。健康指標の課題と地域資源の強みを分けて記述する。",
     "第2章第4節\n第5章 基本目標1・2"),
]:
    r = body(ws, r, [area, ind, val, fin, dest], height=44,
             align={3: "center"})

note(ws, r + 1,
     "注1）本シートは別冊「調査クロス集計・分析」（24シート）の要約である。"
     "所見24件の全文及び根拠となる数値は同資料12シートによる。"
     "注2）本調査は要支援者・要介護者を含まない。"
     "①在宅生活改善調査（要介護者）、②居所変更実態調査（施設入所者）とは"
     "母集団が異なるため、割合を直接比較しない。"
     "注3）同規模保険者40との比較には、"
     "比較対象の77.5％が要支援者を調査対象に含むという偏りがある。"
     "「良好」と判定した指標は断定せず、"
     "「課題」と判定した指標は偏りが良好側に働いてもなお下回るため"
     "確度が高い。", 5)

# ============================================================ 06 供給①
ws = sheet("06_供給構造①_事業所数と定員", "供給構造①　事業所数と定員",
           "北海道の名簿及び見える化システムのK系列により、"
           "区域内の事業所数と定員を確定する。"
           "調査に回答していない施設を補い、"
           "計画本文に載せる供給量の基礎とする。",
           [28, 12, 12, 12, 12, 40], freeze="A5")

r = lead(ws, 4, "【介護保険の指定を受ける施設・居住系】", 6)
r = header(ws, r, ["区分", "事業所数", "定員", "調査の把握", "差", "備考"])
for nm, n, cap, sn, scap, memo in [
    ("介護老人福祉施設（特養）", 3, CAP_TOKUYO, 2, 110,
     "東神楽町アゼリアハイツ（50人）が調査に未回答。"
     "定員160人は資料依頼No.7③の「定員減少234人→160人」と一致する"),
    ("地域密着型介護老人福祉施設", 3, CAP_CHITOKU, 2, 42,
     "アゼリアハイツ（ユニット型・20人）が未回答"),
    ("介護老人保健施設", 3, CAP_ROKEN, 3, CAP_ROKEN,
     "見える化K1cと一致し、区域内のすべてを把握している"),
    ("認知症対応型共同生活介護", 7, CAP_GH, 5, 81,
     "名簿は7事業所。公表システムでファミリー18人を確認した。"
     "くるみの郷1事業所の定員が未確定である"),
    ("特定施設入居者生活介護", 3, CAP_TOKUTEI, 2, 58,
     "さわやか東神楽館（100人）が未回答"),
]:
    r = body(ws, r, [nm, n, cap, "%d施設・%d人" % (sn, scap),
                     "%+d人" % (cap - scap), memo], height=36,
             align={2: "right", 3: "right", 4: "center", 5: "center"})
r = body(ws, r, ["計", 19, CAP_TOKUYO + CAP_CHITOKU + CAP_ROKEN + CAP_GH
                 + CAP_TOKUTEI, "14施設・%d人" % (110 + 42 + CAP_ROKEN + 81
                                                 + 58),
                 "%+d人" % (CAP_TOKUYO + CAP_CHITOKU + CAP_GH + CAP_TOKUTEI
                            - (110 + 42 + 81 + 58)),
                 "くるみの郷の定員を除く"], {3: OK_G}, height=22, bold=True,
         align={2: "right", 3: "right", 4: "center", 5: "center"})

r += 1
r = lead(ws, r, "【介護保険の指定を受けない住まい】", 6)
r = header(ws, r, ["区分", "施設数", "定員・戸数", "調査の把握", "差", "備考"])
for nm, n, cap, sn, scap, memo in [
    ("住宅型有料老人ホーム", 9, CAP_JUTAKU, 3, 91,
     "6施設が未回答。介護保険の指定サービスではないため"
     "見える化のK系列・M2系列に現れず、名簿が唯一の把握手段である"),
    ("サービス付き高齢者向け住宅", 2, CAP_SAKO, 1, 30,
     "桜華（東川町・36戸）が未回答。2件とも東川町にある"),
    ("軽費老人ホーム（ケアハウス）", 1, CAP_KEIHI, 0, 0,
     "ケアハウスびえい（美瑛町）。特定施設の指定はない"),
]:
    r = body(ws, r, [nm, n, cap, "%d施設・%d人" % (sn, scap),
                     "%+d" % (cap - scap), memo], height=36,
             align={2: "right", 3: "right", 4: "center", 5: "center"})
r = body(ws, r, ["計", 12, CAP_JUTAKU + CAP_SAKO + CAP_KEIHI,
                 "4施設・121人",
                 "%+d" % (CAP_JUTAKU + CAP_SAKO + CAP_KEIHI - 121),
                 "調査で把握できたのは3分の1にとどまる"], {3: OK_G},
         height=22, bold=True,
         align={2: "right", 3: "right", 4: "center", 5: "center"})

r += 1
r = lead(ws, r, "【在宅サービスの事業所数（名簿と見える化の対照）】", 6)
r = header(ws, r, ["サービス", "名簿\nR8.6.30", "見える化\nR6年度", "差", "",
                   "備考"])
for nm, key, code, memo in [
    ("訪問介護", "訪問介護", "K3a", "介護人材実態調査の母数。回収率23.1％"),
    ("訪問看護", "訪問看護", "K3c", ""),
    ("訪問リハビリテーション", "訪問リハ", "K3d",
     "医療機関のみなし指定は名簿に載らない"),
    ("通所介護", "通所介護", "K3f", ""),
    ("地域密着型通所介護", "地域通所", "K3g", "4事業所のうち3事業所が美瑛町"),
    ("短期入所生活介護", "短期生活", "K3i",
     "ショート専用定員%d人のほか、4施設で空床利用が可能" % CAP_SHORT),
    ("小規模多機能型居宅介護", "小規模居宅", "K3p",
     "5事業所すべてが社会福祉法人美瑛慈光会（08シート）"),
    ("居宅介護支援", "居宅支援", "K3s", "在宅生活改善調査の事業所票の母数"),
]:
    n = len(H.SHITEI.get(key, []))
    kn = klast(code)
    r = body(ws, r, [nm, n, kn, "%+d" % (n - kn) if n != kn else "±0", "",
                     memo], {4: OK_G if n == kn else IN_Y}, height=24,
             align={2: "right", 3: "right", 4: "center"})

r += 1
r = lead(ws, r, "【区域内に事業所がないサービス】", 6)
r = header(ws, r, ["サービス", "事業所数", "受給率", "利用者票の言及", "",
                   "備考"])
for nm, code, key in [("定期巡回・随時対応型訪問介護看護", "K3m", "定期巡回ｻｰﾋﾞｽ"),
                      ("夜間対応型訪問介護", "K3n", "夜間対応型訪問介護"),
                      ("認知症対応型通所介護", "K3o", None),
                      ("看護小規模多機能型居宅介護", "K3q", "看護小規模多機能"),
                      ("訪問入浴介護", "K3b", None),
                      ("介護医療院", "K1e", None),
                      ("地域密着型特定施設入居者生活介護", "K2b", None)]:
    m = SVC.get(key, 0) if key else "―"
    r = body(ws, r, [nm, klast(code), "0.0％", m, "",
                     "受給率0.0％と事業所数0が整合する"], {2: NG_O},
             height=22, align={2: "right", 3: "right", 4: "right"})

note(ws, r + 1,
     "注1）出典は、施設・住まいの定員が北海道の名簿"
     "（届出済有料老人ホーム一覧・特別養護老人ホーム名簿・"
     "住所地特例適用施設一覧・軽費老人ホーム名簿。"
     "いずれも令和8年6月30日から7月1日現在）、"
     "在宅サービスの事業所数が北海道の介護保険事業所一覧"
     "（令和8年6月30日現在）及び見える化システムのK系列（令和6年度）である。"
     "詳細は別冊「住まいと施設の公表名簿との突合」（11シート）による。"
     "注2）介護老人保健施設とグループホームの定員は名簿の対象外であり、"
     "老健は居所変更実態調査、"
     "グループホームは同調査と介護サービス情報公表システムの"
     "個別公表画面による。"
     "注3）名簿と見える化の時点は1年8か月異なる。"
     "認知症対応型共同生活介護（＋2）と居宅介護支援（＋3）の差は"
     "この間の開設によるものと考えられる。", 6)

# ============================================================ 07 供給②
ws = sheet("07_供給構造②_サービスの実施地域",
           "供給構造②　サービスの通常の事業実施地域",
           "北海道の介護保険事業所一覧に収録された運営規程の"
           "「通常の事業実施地域」により、町ごとに使えるサービスを整理する。"
           "実施地域外であっても利用できないわけではないが、"
           "交通費等の実費を負担することがある。",
           [26, 14, 12, 12, 12, 44], freeze="A5")

r = header(ws, 4, ["サービス", "事業所数", "東川町", "美瑛町", "東神楽町",
                   "見方"])
for key, label, memo in [
    ("訪問介護", "訪問介護", "美瑛町を対象とするのが最も少ない。"
     "3事業所（フラワー・桜華・恩送り。いずれも東川町）は実施地域欄が空欄"),
    ("訪問看護", "訪問看護",
     "区域全体を対象とする事業所が3ある。医療的ケアの受け皿は比較的厚い"),
    ("通所介護", "通所介護（一般）", "2事業所とも複数町を対象とする"),
    ("地域通所", "地域密着型通所介護",
     "美瑛町の3事業所のうち2事業所は美瑛町のみを対象とする。"
     "地域密着型は原則として区域内の被保険者のみが利用できる"),
    ("小規模居宅", "小規模多機能型居宅介護",
     "5事業所すべてが美瑛町にあり、"
     "東川町・東神楽町を対象とするのは「七彩」1事業所のみである。"
     "利用者票で小規模多機能が最多に挙げられたことの背景である"),
    ("居宅支援", "居宅介護支援", "計画作成の担い手。各町に事業所がある"),
    ("短期生活", "短期入所生活介護", "施設併設。区域全体を対象とするものが多い"),
]:
    recs = H.SHITEI.get(key, [])
    cov = [sum(1 for x in recs if t in (x["実施地域"] or "")) for t in TOWNS]
    fills = {i + 3: (NG_O if cov[i] <= 1 else (IN_Y if cov[i] <= 3 else OK_G))
             for i in range(3)}
    r = body(ws, r, [label, len(recs)] + cov + [memo], fills, height=36,
             align={2: "right", 3: "center", 4: "center", 5: "center"})

r += 1
r = lead(ws, r, "【町別の実事業所数】", 6)
r = header(ws, r, ["町", "実事業所数", "延べ登録数", "構成比", "", "見方"])
cnt_t = collections.Counter(v["町"] for v in UNIQ.values())
cnt_d = collections.Counter()
for _sv, _rows in H.SHITEI.items():
    for _r in _rows:
        cnt_d[_r["町"]] += 1
for t in TOWNS:
    r = body(ws, r, [t, cnt_t[t], cnt_d[t],
                     "%.1f％" % (cnt_t[t] / N_JIG * 100), "",
                     "美瑛町は人口規模に比して事業所数が多い。"
                     "社会福祉法人美瑛慈光会の事業所が集中している"
                     if t == "美瑛町" else ""], height=24,
             align={2: "right", 3: "right", 4: "right"})
    ws.merge_cells(start_row=r - 1, start_column=5, end_row=r - 1, end_column=5)
r = body(ws, r, ["計", N_JIG, sum(cnt_d.values()), "100.0％", "", ""],
         {1: GRAY}, height=20, bold=True,
         align={2: "right", 3: "right", 4: "right"})

note(ws, r + 1,
     "注1）「通常の事業の実施地域」は運営規程に定めるものであり、"
     "これを超えて利用する場合は交通費等の実費を負担することがある。"
     "実施地域外であっても利用できないわけではない。"
     "注2）本シートは第1章第7節（訪問・通所困難地域）の根拠となる。"
     "健康とくらしの調査の「自力移動の手段がない層の分布」"
     "（別冊 03シート）と併せて用いる。"
     "注3）実施地域欄が空欄の事業所は、所在町を対象としているとみるのが"
     "自然だが、確認を要する。"
     "注4）延べ登録数は介護予防・総合事業の区分を含むため、"
     "同一の事業所が複数回数えられている。"
     "実事業所数は法人名と事業所名によりユニーク化したものである。", 6)

# ============================================================ 08 供給③
ws = sheet("08_供給構造③_運営法人の集中", "供給構造③　運営法人の集中",
           "区域内の実%d事業所を運営する%d法人の分布を見る。"
           "調査結果の偏りが、調査設計によるものか供給構造によるものかを分けて示す。"
           % (N_JIG, N_HOJIN),
           [34, 12, 12, 12, 12, 40], freeze="A5")

r = header(ws, 4, ["運営法人", "事業所数", "構成比", "累積", "主な町", "見方"])
cum = 0
for h, n in BYH.most_common(10):
    cum += n
    towns = collections.Counter(v["町"] for k, v in UNIQ.items() if k[0] == h)
    memo = ""
    if "美瑛慈光会" in h:
        memo = "小規模多機能5事業所すべて、地域密着型特養、特養、老健、" \
               "地域密着型通所介護等を運営する。区域内で最大の事業者である"
    elif "栄友" in h:
        memo = "グループホーム・特定施設・サービス付き高齢者向け住宅・" \
               "訪問介護・訪問看護・通所介護・居宅介護支援を一体で運営する"
    elif "旭川福祉事業会" in h:
        memo = "特養・老健・訪問リハ・地域密着型通所介護等を運営する"
    r = body(ws, r, [h, n, "%.1f％" % (n / N_JIG * 100),
                     "%.1f％" % (cum / N_JIG * 100),
                     "・".join(t for t, _ in towns.most_common()), memo],
             height=32, align={2: "right", 3: "right", 4: "right"})
r = body(ws, r, ["上位10法人", cum, "%.1f％" % (cum / N_JIG * 100), "", "",
                 "%d法人のうち%d法人は1事業所のみを運営する"
                 % (N_HOJIN, sum(1 for v in BYH.values() if v == 1))],
         {1: GRAY}, height=24, bold=True,
         align={2: "right", 3: "right"})

r += 1
r = lead(ws, r, "【小規模多機能型居宅介護の運営法人】", 6)
r = header(ws, r, ["事業所名", "法人", "町", "通常の事業実施地域", "", "見方"])
for rec in H.SHITEI.get("小規模居宅", []):
    cov = [t for t in TOWNS if t in (rec["実施地域"] or "")]
    r = body(ws, r, [rec["事業所名"], rec["法人"], rec["町"],
                     rec["実施地域"] or "（記載なし）", "",
                     "区域全体を対象とする唯一の事業所"
                     if len(cov) == 3 else ""], height=24,
             align={3: "center"})
    ws.merge_cells(start_row=r - 1, start_column=5, end_row=r - 1, end_column=5)

r += 1
r = lead(ws, r, "【利用者票の偏りとの関係】", 6)
r = header(ws, r, ["観察された事実", "", "解釈", "", "", "取扱い"])
for fact, interp, act in [
    ("利用者票99票のうち36票（36.4％）が同一法人の"
     "小規模多機能2事業所からの提出である",
     "区域内の小規模多機能5事業所はすべて社会福祉法人美瑛慈光会が"
     "美瑛町で運営している。"
     "小規模多機能の利用者票が同一法人に集中するのは必然である。",
     "調査設計の問題ではなく供給構造の反映である。"
     "「小規模多機能が最多」という結果を"
     "そのまま需要の強さと読まない"),
    ("「より適切と思われるサービス」で小規模多機能が44件と最多である。"
     "うち31件が提出元の2事業所による",
     "回答者である事業所が自らの提供するサービスを挙げる構造がある。",
     "提出元を除いた集計を併記するか、"
     "件数を示さず質的に記述する（受託者の案）"),
    ("東川町・東神楽町の住民が使える小規模多機能は"
     "「七彩」1事業所のみである",
     "小規模多機能は地域密着型サービスであり、"
     "実施地域外の利用は想定されていない。"
     "東川町・東神楽町では選択肢が実質的に1つである。",
     "第6章第2節の小規模多機能の見込量は、"
     "町別の実施地域を踏まえて設定する"),
]:
    r = body(ws, r, [fact, "", interp, "", "", act], height=52)
    ws.merge_cells(start_row=r - 1, start_column=1, end_row=r - 1, end_column=2)
    ws.merge_cells(start_row=r - 1, start_column=3, end_row=r - 1, end_column=5)

note(ws, r + 1,
     "注1）本シートは、調査結果の偏りが調査設計によるものか"
     "供給構造によるものかを分けて示すためのものである。"
     "偏りの存在は調査の欠陥を意味しない。"
     "注2）法人の集中は、"
     "①事業者の撤退が地域のサービス供給に与える影響が大きいこと、"
     "②事業所間の比較による質の評価が成り立ちにくいことを意味する。"
     "第6章第4節（整備の方針）の留意事項とする。"
     "注3）本シートの解釈は受託者によるものであり、"
     "計画本文に記述する場合は表現をご確認いただく。"
     "特定の法人の評価と受け取られない書き方とする。", 6)

# ============================================================ 09 供給④
ws = sheet("09_供給構造④_従事者数", "供給構造④　従事者数",
           "介護人材実態調査の介護職員数と見える化システムのM2系列"
           "（介護サービス施設・事業所調査による職種別従事者数）を突合する。"
           "見える化に系列がないサービスは、本調査が唯一の把握手段である。",
           [28, 12, 12, 12, 12, 40], freeze="A5")

r = header(ws, 4, ["サービス", "調査の\n介護職員", "見える化\nM2系列", "差",
                   "", "見方"])
for nm, cat, code, memo in [
    ("介護老人福祉施設", "特養", "M2a", "未回答1施設で差が説明できる"),
    ("地域密着型介護老人福祉施設", "地密特養", "M2k",
     "未回答1施設で差が説明できる。1施設当たり11.0人で一致する"),
    ("介護老人保健施設", "老健", "M2b",
     "3施設すべてが回答したうえで差が生じている。"
     "見える化はR4 78人→R5 56人→R6 59人と推移しており、"
     "調査値77人はR4の水準に近い（点検事項No.31）"),
    ("通所リハビリテーション", "通所リハ", "M2g",
     "老健と同様に差が生じている。「介護職員」の範囲の解釈を確認する"),
    ("訪問介護", "訪問介護", "M2e",
     "13事業所のうち3事業所からの回答であり、母集団が異なる"),
    ("認知症対応型共同生活介護", "GH", None,
     "見える化に従事者数の系列がない"),
    ("特定施設入居者生活介護", "特定施設", None,
     "見える化に従事者数の系列がない。"
     "さわやか東神楽館が未回答であり、調査値も一部にとどまる"),
    ("住宅型有料老人ホーム", "住宅型有料", None,
     "介護保険の指定サービスではないため見える化に現れない。"
     "華の13人を重複として除いた後の数値である"),
    ("サービス付き高齢者向け住宅", "サ高住", None,
     "回答した1件の介護職員は0人である"),
]:
    got = BYCAT[cat]
    mv = None
    if code:
        v = MK.M[code]["職種"].get("介護職員", {})
        ks = [k for k in v if v[k] is not None]
        mv = int(v[ks[-1]]) if ks else None
    r = body(ws, r, [nm, got, mv if mv is not None else "系列なし",
                     ("%+d" % (got - mv)) if mv is not None else "―", "",
                     memo], {3: NG_O if mv is None else None}, height=40,
             align={2: "right", 3: "right", 4: "center"})
r = body(ws, r, ["計", sum(BYCAT.values()), "―", "―", "",
                 "華の重複13人を除く"], {1: GRAY}, height=20, bold=True,
         align={2: "right"})

r += 1
r = lead(ws, r, "【見える化で把握できない介護職員】", 6)
r = header(ws, r, ["区分", "介護職員", "介護職員に対する割合", "", "", "見方"])
r = body(ws, r, ["グループホーム・特定施設・住宅型有料・サ高住", NOSER,
                 "%.1f％" % (NOSER / N_ALL * 100), "", "",
                 "見える化システムに従事者数の系列がないサービスである。"
                 "介護人材実態調査が唯一の把握手段であり、"
                 "調査を継続する意義がここにある。"
                 "重複を除く前の基準では129人（361人の35.7％）となる"],
         {2: OK_G}, height=44, align={2: "right", 3: "right"})
ws.merge_cells(start_row=r - 1, start_column=4, end_row=r - 1, end_column=5)
r = body(ws, r, ["見える化で把握できるサービス", N_ALL - NOSER,
                 "%.1f％" % ((N_ALL - NOSER) / N_ALL * 100), "", "",
                 "特養・地域密着型特養・老健・通所リハ・通所介護等・訪問介護"],
         height=24, align={2: "right", 3: "right"})
ws.merge_cells(start_row=r - 1, start_column=4, end_row=r - 1, end_column=5)
r = body(ws, r, ["計", N_ALL, "100.0％", "", "", "重複13人を除く"],
         {1: GRAY}, height=20, bold=True, align={2: "right", 3: "right"})

r += 1
r = lead(ws, r, "【公表システムの個別画面で補える介護職員】", 6)
r = header(ws, r, ["事業所", "介護職員", "常勤／非常勤", "退職者数", "",
                   "見方"])
for nm, tot, kin, hi, ret, view in [
    ("さわやか東神楽館（特定施設）", SW_KAIGO, SW["介護職員_常勤"],
     SW["介護職員_非常勤"],
     SW["介護職員の退職者数_常勤"] + SW["介護職員の退職者数_非常勤"],
     "介護人材実態調査に回答していない。"
     "非常勤が半数を占め、退職者数は介護職員の35.5％にあたる。"
     "記入日は令和7年10月6日であり、調査（令和7年4月1日現在）とは時点が異なる"),
    ("グループホームファミリー", 12, 12, 0, "―",
     "介護人材実態調査には介護付きホームファミリー（特定施設）の"
     "事業所票のみがあり、グループホームの分はない。"
     "記入日は令和7年10月20日である"),
]:
    r = body(ws, r, [nm, tot, "常勤%s・非常勤%s" % (kin, hi), ret, "", view],
             height=40, align={2: "right", 4: "center"})
r = body(ws, r, ["計", SW_KAIGO + 12, "―", "―", "",
                 "調査の348人に加えると379人となるが、"
                 "時点が異なるため単純に合算せず、"
                 "「調査で把握した348人のほか、"
                 "公表データにより2事業所43人を確認した」と記述する"],
         {2: OK_G}, height=40, bold=True, align={2: "right"})

r += 1
r = lead(ws, r, "【人材の確保に関する数値】", 6)
r = header(ws, r, ["項目", "値", "", "", "", "計画での用い方"])
for nm, v, use in [
    ("介護職員（重複を除く）", "%d人" % N_ALL,
     "第2章第6節の基礎数値。訪問系は3事業所分にとどまることを注記する"),
    ("介護福祉士の割合", "67.1％",
     "質の指標。北海道の評価指標及び第10期基本指針案に対応する"),
    ("勤続1年未満の割合", "19.9％",
     "定着の**補足指標**。代表KPI H14は「採用率と離職率の差」であり、"
     "本値はH14そのものではない"),
    ("令和6年度の採用・離職", "採用%d人・離職%d人" % (SAIYO, RISHOKU),
     "採用率と離職率の差を算定するには各年9月30日現在の在籍者数を要する"),
    ("外国人職員", "%d人（%d事業所）" % (GAIKOKU, N_GAI_JIG),
     "施設・通所系の介護職員%d人の%.1f％。"
     "第5章基本目標4の外国人材の受入れの根拠"
     % (N_SIS, GAIKOKU / N_SIS * 100)),
    ("60歳以上の職員", "42人（13.3％）",
     "今後10年の退職見込み。人材確保の時間軸を示す"),
]:
    r = body(ws, r, [nm, v, "", "", "", use], height=28, align={2: "center"})
    ws.merge_cells(start_row=r - 1, start_column=2, end_row=r - 1, end_column=5)

note(ws, r + 1,
     "注1）見える化M2系列は介護サービス施設・事業所調査（各年10月1日現在）"
     "による職種別の実人員である。"
     "本調査は令和7年4月1日現在であり、半年のずれがある。"
     "注2）本調査の「介護職員数」は事業所票の単一の欄によるものであり、"
     "職種別の内訳がない。"
     "見える化M2系列の「介護職員」と範囲が一致するかは確認を要する。"
     "介護老人保健施設の差（点検事項No.31）はこの範囲の問題である"
     "可能性がある。"
     "注3）介護サービス情報公表システムの個別公表画面には"
     "職種別の従業者数・常勤換算・前年度の退職者数が掲載されており、"
     "事業所単位で突合できる。"
     "オープンデータのCSVには従業者数が含まれないため、"
     "個別画面によることとなる"
     "（別冊「実施済み3調査の受領点検と集計」15シート）。", 6)

# ============================================================ 10 対照
ws = sheet("10_供給と需要の対照", "供給と需要の対照",
           "供給量（06シート）とサービス見込量の算定 第1次概算"
           "（令和8年9月16日。計画素案 第6章と同じ値）を対照する。"
           "見込量は、要介護度別の利用率と1人1月あたり給付費を"
           "令和7年度の値で固定した算定による。"
           "実績は年報（令和7年度）の月平均であり、"
           "保険者（広域連合）を単位として当広域連合の被保険者を数えたものである。",
           [26, 14, 14, 14, 14, 40], freeze="A5")

r = header(ws, 4, ["区分", "令和7年度\n実績（人／月）", "令和11年度\n見込み（人／月）",
                   "区域内定員", "定員に対する\n割合", "見方"])
for nm, act, est, cap, memo in [
    ("施設サービス", SHI_R7, SHI_R11, CAP_TOKUYO + CAP_CHITOKU + CAP_ROKEN,
     "合計では定員に収まるが、種別別では逼迫するものがある"
     "（地域密着型介護老人福祉施設は令和11年度に定員62人の99.9％）。"
     "特養の待機者は別に存在する"),
    ("居住系サービス", KYO_R7, KYO_R11, CAP_GH + CAP_TOKUTEI,
     "合計では定員に収まる。"
     "区域内の定員は当広域連合の被保険者にとっての上限ではない。"
     "年報は施設の所在地を問わず当広域連合の被保険者を数えるため、"
     "区域外の施設の利用が含まれる"),
]:
    r = body(ws, r, [nm, round(act, 1), round(est, 1), cap,
                     "%.1f％" % (est / cap * 100), memo],
             height=44, align={2: "right", 3: "right", 4: "right",
                               5: "right"})

r += 1
r = lead(ws, r, "【施設の側から数えた在籍者と、保険者の側から数えた受給者】", 6)
r = header(ws, r, ["区分", "調査②\n在籍（人）", "年報 受給\n（人／月）",
                   "差\n（在籍−受給）", "", "見方"])
for nm, z, j in ZAISEKI_VS:
    if z >= j:
        yomi = ("在籍＞受給。区域内の施設に他の保険者の被保険者が"
                "入居している（住所地特例を含む）")
    else:
        yomi = ("受給＞在籍。当広域連合の被保険者が区域外の施設を"
                "利用している。区域内の定員では説明できない")
    r = body(ws, r, [nm, z, round(j, 1), round(z - j, 1), "", yomi],
             {2: IN_Y if nm == "介護老人保健施設" else None},
             height=32, align={2: "right", 3: "right", 4: "right"})
    ws.merge_cells(start_row=r - 1, start_column=5, end_row=r - 1,
                   end_column=5)

r += 1
r = lead(ws, r, "【認定者の利用状況（令和7年度）】", 6)
r = header(ws, r, ["区分", "推計人数", "認定者に対する割合", "", "", "見方"])
for nm, v, pct, memo in [
    ("在宅サービス", 1016, 51.2, "訪問・通所・短期入所等"),
    ("居住系サービス", 145, 7.3, "グループホーム・特定施設"),
    ("施設サービス", 337, 17.0, "特養・老健・地域密着型特養"),
    ("いずれも利用していない", 486, 24.5,
     "認定者1,984人の4分の1。"
     "健康とくらしの調査の「介護・介助が必要だが受けていない7.0％」とは"
     "別の集団であり、合算しない。"
     "内訳（資料依頼No.8）の受領後に要因を分解する"),
]:
    r = body(ws, r, [nm, v, "%.1f％" % pct, "", "", memo],
             {2: IN_Y if "利用していない" in nm else None}, height=32,
             align={2: "right", 3: "right"})
    ws.merge_cells(start_row=r - 1, start_column=4, end_row=r - 1, end_column=5)

r += 1
r = lead(ws, r, "【調査結果と供給構造の接続】", 6)
r = header(ws, r, ["調査で示されたこと", "", "供給構造の側から見えること", "",
                   "", "計画での扱い"])
for a, b, c in [
    ("利用者票で必要とされた生活支援の最多は外出同行（51件・51.5％）である",
     "訪問系13事業所のうち美瑛町を実施地域とするのは4事業所である。"
     "健康とくらしの調査では自力移動の手段がない層が20％台にある。",
     "移動の確保を第5章基本目標2の重点に位置づける。"
     "介護保険の給付では担えない部分は生活支援体制整備事業による"),
    ("利用者票で夜間対応型・定期巡回・看多機が計26件挙げられた",
     "この3サービスは区域内に事業所がなく、受給率も0.0％である。",
     "第6章第4節（24時間対応サービスの確保方策）の根拠とする。"
     "整備するか、既存の訪問介護・訪問看護の連携で代替するかを決める"),
    ("居所変更実態調査で新規入所の57.0％が病院・診療所からである",
     "老健3施設240人が区域内の中間施設として機能している。"
     "在宅への復帰か施設入所かの分岐点が医療機関にある。",
     "第6章第5節（医療と介護の連携）の根拠とする"),
    ("介護人材実態調査で介護職員の33.3％が"
     "見える化で把握できないサービスに従事している",
     "グループホーム7・特定施設3・住宅型有料9・サ高住2・軽費1の"
     "22施設が該当する。"
     "うち調査に回答したのは一部にとどまる。",
     "第2章第6節では、把握範囲を明記したうえで記述する。"
     "本調査を継続する意義を第10期に位置づける"),
    ("利用者票で小規模多機能が最多（44件）に挙げられた",
     "区域内の小規模多機能5事業所はすべて1法人が美瑛町で運営し、"
     "東川町・東神楽町を実施地域とするのは1事業所のみである。",
     "第6章第2節の見込量は町別の実施地域を踏まえて設定する。"
     "件数をそのまま需要の強さと読まない"),
]:
    r = body(ws, r, [a, "", b, "", "", c], height=52)
    ws.merge_cells(start_row=r - 1, start_column=1, end_row=r - 1, end_column=2)
    ws.merge_cells(start_row=r - 1, start_column=3, end_row=r - 1, end_column=5)

note(ws, r + 1,
     "注1）見込量は別冊「サービス見込量の算定 第1次概算」（18シート）による。"
     "計画素案 第6章に反映済みの値と同じである"
     "（令和8年9月16日。従前は将来推計 第2段階の値を掲げていた）。"
     "前提は確定しておらず、第2次概算（令和8年10月30日）により改まる。"
     "注2）本シートは、4調査を横断したクロス集計に代えて"
     "調査相互を接続するものである（13シート）。"
     "個票を突き合わせたものではなく、集計値と供給構造による対照である。"
     "注3）人数は、認定者数%s人（令和8年3月末）に"
     "令和7年度の合計利用率（見える化D45系列）を乗じて算定した推計人数であり、"
     "月平均の受給者実績（在宅1,013人・居住系%.1f人・施設%.1f人）とは"
     "算定の方法が異なるため一致しない。4区分の合計は%s人となる。"
     "注4）「いずれも利用していない486人」と"
     "健康とくらしの調査の「介護・介助が必要だが受けていない7.0％」は"
     "母集団が異なる（前者は認定者、後者は一般高齢者）。"
     "合算も比率の比較もしない。"
     "注5）在籍と受給は数える対象が異なる。"
     "年報（受給）は施設の所在地を問わず当広域連合の被保険者を数え、"
     "調査②（在籍）は保険者を問わず区域内の施設の在籍者を数える。"
     "介護老人保健施設だけが在籍%d人＞受給%.1f人／月（＋%.1f人）であり、"
     "他の3区分は受給＞在籍である。"
     "したがって区域内の定員は必要利用定員総数の上限ではなく、"
     "到達率が100％を超えることは誤りではない（確認事項No.88・No.141）。"
     % ("{:,}".format(NINTEI_R7), KYO_R7, SHI_R7,
        "{:,}".format(NINTEI_R7), ROKEN_Z, ROKEN_J, ROKEN_Z - ROKEN_J), 6)

# ============================================================ 11 限界
ws = sheet("11_調査結果の限界と留保", "調査結果の限界と留保",
           "本報告書の数値を計画本文に用いる際の留保を整理する。"
           "留保を付さずに用いてよい数値と、"
           "母集団や回収状況を明記して用いる数値を分ける。",
           [4, 26, 50, 44, 16], freeze="A5")

r = header(ws, 4, ["No.", "留保の内容", "具体的に何が言えないか", "対処", "区分"])
for no, nm, cant, act, kb in [
    (1, "母集団が調査ごとに異なる",
     "①は事業所が課題ありと判断した利用者、"
     "②は施設入所者、③は事業所の職員、"
     "④は一般高齢者（要支援・要介護者を含まない）である。"
     "割合を相互に比較したり、人数を足し合わせたりできない。",
     "各数値に母集団を明記する。"
     "「認定者の○％」「回答した99人のうち○％」と書き分ける。", "必須"),
    (2, "①利用者票は課題のある利用者を抽出する設計である",
     "「在宅生活の維持が困難%.1f％」を"
     "区域内の在宅利用者全体の割合と読めない。"
     "年報の要介護度別の明細との突合により、偏りの大きさと向きを測定した。"
     % H06 +
     "要介護度分布は在宅の認定者の分布と食い違う"
     "（χ²＝%.1f、自由度6。1％点%.2f）。"
     "要介護3は期待%.1f人に対し%d人（%.1f倍）、"
     "要支援1は期待%.1f人に対し%d人（%.2f倍）である。"
     % (CHI, CHI_1P, CU_DO_EXP[4], CU_DO_OBS[4],
        CU_DO_OBS[4] / CU_DO_EXP[4], CU_DO_EXP[0], CU_DO_OBS[0],
        CU_DO_OBS[0] / CU_DO_EXP[0]),
     "「調査対象となった%d人のうち」と明記する。" % RIYO_N +
     "率ではなく実数で記述する。"
     "「在宅で困難を抱える方の状況」として引き、"
     "在宅の認定者全体の割合としては引かない"
     "（確認事項No.140。突合クロス集計 02シート）。", "必須"),
    (3, "①②③は配布数の記録がなく回収率を算定できない",
     "回答が区域内の実態を代表しているかを判定できない。",
     "公表データで母数を置き換えられるもの"
     "（訪問介護13事業所、居宅介護支援13事業所）のみ回収率を示す。",
     "必須"),
    (4, "②は13施設が未回答である（把握率58.1％）",
     "施設・住まいの入所者数・入退去の状況を区域全体で示せない。"
     "とくに区域内最大の特定施設が欠けている。",
     "定員は名簿による。入所者数は回答した施設の分であることを明記する。",
     "必須"),
    (5, "③の訪問系は13事業所のうち3事業所からの回答である",
     "訪問系の介護職員数・資格・勤続年数を区域全体で示せない。",
     "訪問系を含む合計を「区域内の介護職員数」と記述しない。"
     "施設・通所系と訪問系を分けて示す。", "必須"),
    (6, "③の介護職員数に重複がある",
     "総数が348人か361人か確定しない。",
     "点検事項No.1の取扱いの決定による。"
     "受託者の案は348人（訪問系13人を実数とする）。", "決定待ち"),
    (7, "③の老健・通所リハの介護職員数が見える化と一致しない",
     "従事者数の推移を見える化で記述できるかどうかが定まらない。",
     "点検事項No.31の取扱いの決定による。"
     "介護サービス情報公表システムの個別画面による確認を提案する。",
     "決定待ち"),
    (8, "①の所在地区の記入形式が統一されていない",
     "地区別・圏域別の集計ができない。"
     "4調査を横断したクロス集計もできない。",
     "地区別の記述を行わない（13シート）。", "決定済"),
    (9, "①の利用者票に回答の偏りがある",
     "「より適切と思われるサービス」の件数を"
     "そのまま需要の強さと読めない。",
     "点検事項No.2の取扱いの決定による。"
     "提出元を除いた集計の併記又は質的な記述を提案する。", "決定待ち"),
    (10, "②の「市内／市外」の定義が示されていない",
     "町単位か広域連合単位かが分からず、"
     "介護サービス自給率（δ）と接続できない。",
     "点検事項No.25。発注者への確認による。", "決定待ち"),
    (11, "④は要支援者・要介護者を含まない",
     "④の値を①②の代替として用いることはできない。",
     "代表KPI H06・H11は補完として位置づける。", "必須"),
    (12, "④の同規模保険者比較に対象者範囲の偏りがある",
     "「良好」と判定した指標を断定できない。",
     "「課題」と判定した指標は偏りが良好側に働いてもなお下回るため"
     "確度が高い。良好側は断定を避ける。", "必須"),
    (13, "供給構造の分析は個票の突合ではない",
     "調査結果と供給構造の対応は関連の指摘であって"
     "因果関係を示すものではない。",
     "10シートの記述は「〜と符合する」「〜の背景にある」にとどめ、"
     "「〜に由来する」と書かない。", "必須"),
    (14, "名簿と調査・見える化の時点が異なる",
     "同一時点の断面として扱えない。",
     "名簿は令和8年6月30日〜7月1日現在、"
     "調査は令和7年4月1日現在、"
     "見える化K・M2系列は令和6年度（10月1日現在）である旨を明記する。",
     "必須"),
    (15, "②の在籍者と年報の受給者は数える対象が異なる",
     "施設の在籍者数と保険者の受給者数を同じものとして比べられない。"
     "年報は施設の所在地を問わず当広域連合の被保険者を数え、"
     "調査②は保険者を問わず区域内の施設の在籍者を数える。"
     "介護老人保健施設だけが在籍%d人＞受給%.1f人／月であり、"
     "他の3区分は受給＞在籍である。"
     % (ROKEN_Z, ROKEN_J),
     "在籍と受給を並べるときは数える対象の違いを注記する（10シート）。"
     "区域内の定員を必要利用定員総数の上限として扱わない。"
     "到達率が100％を超えることは誤りではない"
     "（確認事項No.88・No.141）。", "必須"),
    (16, "①の要介護度分布は在宅の認定者の分布と食い違う",
     "①の割合を在宅の認定者全体の割合として引けない。"
     "χ²＝%.1f（自由度6。1％点%.2f）で、"
     "要介護3が期待の%.1f倍、要支援1が期待の%.2f倍である。"
     % (CHI, CHI_1P, CU_DO_OBS[4] / CU_DO_EXP[4],
        CU_DO_OBS[0] / CU_DO_EXP[0]),
     "留保No.2の測定値である。"
     "代表KPI H12（" + H12_NAME + "）も同じ母集団の指標となるため、"
     "母集団を明記して用いる（確認事項No.140）。", "必須"),
]:
    fl = {"必須": IN_Y, "決定待ち": NG_O, "決定済": OK_G}[kb]
    r = body(ws, r, [no, nm, cant, act, kb], {5: fl}, height=52,
             align={1: "center", 5: "center"})

note(ws, r + 1,
     "注1）「必須」は、計画本文に数値を載せる際に必ず付す留保である。"
     "「決定待ち」は、点検事項の取扱いが決まるまで数値が確定しないものである。"
     "注2）本報告書では、"
     "「〜に由来する」「〜と整合する（根拠として）」「1件も〜ない」"
     "「有意差がないため関係がない」「全国トップ級」の5つの表現を用いない"
     "（受託者の自己点検（記述ルール））。"
     "注3）点検事項35件の全文は別冊"
     "「実施済み3調査の受領点検と集計」02シートによる。", 5)

# ============================================================ 12 所見
ws = sheet("12_主要所見と計画本文への反映", "主要所見と計画本文への反映",
           "本報告書の所見と、計画素案の反映先を対応させる。"
           "点検事項の取扱いの決定を要しないものから着手する。",
           [4, 30, 46, 26, 16, 14], freeze="A5")

r = header(ws, 4, ["No.", "所見", "根拠となる数値", "反映先", "前提", "状態"])
FIND = [
    (1, "医療機関からの入所が施設入所の主要な経路である",
     "新規入所335人の入所前の居場所は病院・診療所57.0％・自宅27.5％。"
     "退去349人の退去先は病院・診療所47.0％・死亡27.2％・自宅12.3％。",
     "第2章第3節\n第6章第5節", "なし", "着手可"),
    (2, "施設・住まいの供給量が確定した",
     "特養160人・地域密着型特養62人・老健240人・"
     "グループホーム99人・特定施設156人。"
     "介護保険外は住宅型有料223人・サ高住66戸・軽費50人。",
     "第2章第3節", "なし", "着手可"),
    (3, "24時間対応の3サービスが区域内に存在しない",
     "定期巡回・夜間対応・看多機の事業所数0、受給率0.0％。"
     "利用者票では計26件が必要と回答されている。",
     "第6章第4節", "なし", "着手可"),
    (4, "訪問・通所困難地域の判定材料がそろった",
     "区域内の指定事業所124件の通常の事業実施地域。"
     "訪問介護13事業所のうち美瑛町を対象とするのは4事業所。",
     "第1章第7節", "なし", "着手可"),
    (5, "運営法人の集中が供給構造の特徴である",
     "実75事業所を28法人が運営し、上位6法人が39事業所（52.0％）。"
     "小規模多機能5事業所はすべて社会福祉法人美瑛慈光会。",
     "第2章第3節\n第6章第4節", "表現の確認", "着手可"),
    (6, "介護職員の33.3％は見える化で把握できないサービスに従事している",
     "グループホーム・特定施設・住宅型有料・サ高住の介護職員116人。",
     "第2章第6節", "なし", "着手可"),
    (7, "外出同行が最も必要とされる生活支援である",
     "利用者票99票のうち外出同行51件（51.5％）、"
     "見守り34件、通いの場34件。",
     "第2章第5節\n第5章 基本目標2", "なし", "着手可"),
    (8, "除雪が生活動作の困りごとの最上位である",
     "健康とくらしの調査の回答者4,128人のうち24.3％（1,002人）。生活動作の困りごとの最上位。介護保険の給付対象外。",
     "第2章第4節\n第5章 基本目標2", "なし", "着手可"),
    (9, "外出と社会参加が同規模保険者との比較で最大の課題である",
     "友人知人と会う頻度が高い者63.0％（同規模71.2％）。"
     "転倒35.5％（同30.0％）、口腔機能低下24.7％（同21.8％）。",
     "第2章第4節\n第5章 基本目標1・2", "なし", "着手可"),
    (10, "介護職員の総数",
     "施設・通所系319人＋訪問系42人。重複13人を除くと348人。",
     "第2章第6節", "点検No.1の決定", "決定待ち"),
    (11, "介護人材の定着の状況",
     "介護福祉士67.1％、勤続1年未満19.9％、"
     "常勤の18.6％・非常勤の25.4％が勤続1年未満。"
     "令和6年度の採用83人・離職65人。",
     "第2章第6節\n第5章 基本目標4", "点検No.1の決定", "決定待ち"),
    (12, "老健・通所リハの従事者数を見える化で記述できるか",
     "老健3施設の回答77人に対し見える化M2b（令和6年度）は59人。",
     "第2章第6節", "点検No.31の決定", "決定待ち"),
    (13, "施設等の入所者数と入所率",
     "回答18施設の定員652人。入所者は入所欄の合計536人（入所率82.2％）と、"
     "入所欄の記入がない1施設を要介護度別の内訳58人で補った594人"
     "（入所率91.1％）の2通りがある。どちらを用いるかの決定を要する。",
     "第2章第3節", "点検No.32の決定", "決定待ち"),
    (14, "区域外からの入所の構造",
     "グループホーム84.6％・特定施設84.2％が区域外からの入所。"
     "特養系は区域内82.2％。",
     "第2章第3節\n第6章第4節", "点検No.25・No.32の決定", "決定待ち"),
    (15, "在宅生活の継続が困難な理由",
     "利用者票%d票のうち、この設問の有効回答%d票において"
     "在宅生活の維持が困難%d人（%.1f％）。"
     % (RIYO_N, IJI_YUKO, IJI_KONNAN, H06),
     "第2章第4節", "母集団の記述の確認", "決定待ち"),
    (16, "より適切と思われるサービス",
     "小規模多機能44件・住宅型有料27件・グループホーム20件・"
     "特別養護老人ホーム20件。",
     "第6章第2節", "点検No.2の決定", "決定待ち"),
    (17, "特別養護老人ホームの待機者",
     "待機者137人。同一法人の2施設で同数の重複を除くと82人。",
     "第6章第4節", "点検No.13の決定", "決定待ち"),
    (19, "【参考・時点混在】特定施設の入居者の要介護度構成",
     "さわやか東神楽館（定員100人・入居者97人・令和7年10月6日現在）の"
     "要介護度は要支援1・2が26人（26.8％）、要介護1が42人（43.3％）で、"
     "要支援1〜要介護1が68人（70.1％）を占める。"
     "要介護4・5は9人（9.3％）にとどまる。"
     "調査に回答した2施設（57人・令和7年4月1日現在）と合算すると"
     "要支援1〜要介護2は110人／154人（71.4％）となるが、"
     "回答2施設のみでは30人／57人（52.6％）であり、"
     "時点と施設構成に強く左右される。"
     "同一時点の3施設のデータへ更新するまで、"
     "整備の必要量の算定には使用しない。",
     "第2章第3節\n第6章第2節・第4節", "同一時点の3施設データ", "決定待ち"),
    (20, "特定施設は施設ごとにみると高い入居率である",
     "調査に回答した2施設は98.3％（令和7年4月1日現在）、"
     "未回答の1施設は公表データで97.0％（令和7年10月6日現在）である。"
     "時点が異なるため、区域内合計の入居率"
     "（定員156人に対し入居者154人・98.7％）は参考値である。",
     "第2章第3節\n第6章第4節", "なし", "着手可"),
    (18, "地区別・圏域別の分析",
     "健康とくらしの調査は小学校区12地区別に集計できるが、"
     "①②③は地区別に集計できない。",
     "第1章第7節\n第2章第4節", "④のみで記述する", "着手可"),
]
# 計画素案の実物を読み、反映済みのものを「反映済」に改める。
# 反映状況を固定値で書くと、素案を改訂したときに実際と合わなくなる。
# 所見No.→（判定する章節, その節に現れるべき語）
HANEI = {
    1: ("第6章第5節", "57.0％", "医療機関"),
    2: ("第2章第3節", "160", "62", "240", "99", "156"),
    3: ("第6章第2節", "定期巡回", "看護小規模多機能"),
    4: ("第1章第7節", "通常の事業の実施地域", "%d件" % _SHITEI_NOBE),
    5: ("第2章第3節", "%d法人" % N_HOJIN, "上位6法人"),
    6: ("第2章第6節", "116人", "33.3％"),
    7: ("第2章第5節", "外出同行"),
    8: ("第2章第4節", "除雪"),
    9: ("第2章第4節", "同規模", "転倒"),
    18: ("第2章第4節", "小学校区"),
    20: ("第6章第4節", "98.3％", "97.0％"),
}
FIND = [(no, nm, ev, dest, pre,
         "反映済" if (no in HANEI and st == "着手可"
                      and in_draft(*HANEI[no])) else st)
        for no, nm, ev, dest, pre, st in FIND]

for no, nm, ev, dest, pre, st in sorted(FIND, key=lambda t: t[0]):
    fl = {"着手可": OK_G, "決定待ち": IN_Y, "反映済": MID_B}[st]
    r = body(ws, r, [no, nm, ev, dest, pre, st], {6: fl}, height=52,
             align={1: "center", 6: "center"})

r += 1
NST = collections.Counter(f[5] for f in FIND)
r = lead(ws, r, "【着手の順序】", 6)
r = header(ws, r, ["区分", "件数", "内容", "", "", "時期"])
for kb, n, cont, when in [
    ("反映済", NST["反映済"],
     "計画素案の本文に反映済みであることを素案の実物で確認したもの。"
     "第2章第4節（高齢者の生活実態）は健康とくらしの調査の結果を入れ、"
     "4段落・表1から30段落・表8となった。"
     "第6章第2節・第4節はサービス見込量の算定 第1次概算の値へ差し替えた。"
     "第1章第7節（通常の事業の実施地域）・第2章第3節（運営法人の集中）・"
     "第6章第4節（区域内の入居状況と受給者数）・"
     "第6章第5節（入所の経路）は、"
     "点検事項の決定を要しないものとして反映した。", "反映済"),
    ("着手可", NST["着手可"],
     "点検事項の取扱いの決定を要しないもの。"
     "決定を要しないものはすべて反映済みである。", "―"),
    ("決定待ち", NST["決定待ち"],
     "点検事項No.1・No.2・No.13・No.25・No.31・No.32の取扱いが"
     "決まってから反映する。"
     "ヒアリングシート項目13・項目14でご意向を確認する。", "R8.9"),
]:
    r = body(ws, r, [kb, n, cont, "", "", when],
             {1: {"反映済": MID_B, "着手可": OK_G}.get(kb, IN_Y)}, height=52,
             align={2: "right", 6: "center"})
    ws.merge_cells(start_row=r - 1, start_column=3, end_row=r - 1, end_column=5)

note(ws, r + 1,
     "注1）所見%d件のうち%d件は反映済みであり、"
     "点検事項の取扱いの決定を要せず着手できるものは%d件である。"
     "注2）反映先は計画素案の章節による。"
     "実際の反映は別管理表で管理し、計画本文には確認事項の注記を書かない。"
     "注3）本シートの所見は受託者によるものであり、"
     "計画本文への反映は発注者のご意向を確認したうえで行う。"
     "注4）「反映済」は、計画素案（%s）の本文と表を章節ごとに読んで"
     "判定している。固定値では書いていない。"
     % (len(FIND), NST["反映済"], NST["着手可"], RP.draft_label()), 6)

# ============================================================ 13 横断
ws = sheet("13_横断クロス集計を行わないことの整理",
           "4調査横断のクロス集計を行わないことの整理",
           "仕様書４（3）は4調査の集計・クロス集計・分析を求めている。"
           "このうち4調査を横断したクロス集計について、"
           "実施できない理由と、これに代わる接続の方法を記録する。"
           "行わないことは発注者のご意向による。",
           [4, 30, 50, 44, 16], freeze="A5")

r = header(ws, 4, ["No.", "項目", "内容", "帰結", "区分"])
for no, nm, cont, res, kb in [
    (1, "横断クロス集計に必要な結合キー",
     "4調査は対象者が異なるため個票を1対1で結合できない。"
     "結合できるのは地区（日常生活圏域）又は事業所である。"
     "④健康とくらしの調査は小学校区12地区別に集計できる。",
     "地区を結合キーとする方法が唯一の選択肢であった。", "前提"),
    (2, "①利用者票の所在地区が集計できない",
     "99票の所在地区の記入形式が9種類に分かれる。"
     "町名46件（美瑛町34・東川町11・「1.美瑛町」1）、"
     "「市街地」15件・「郊外」2件、番号13件、未記入22件、区域外1件。"
     "事業所ごとに記入の仕方が異なり、番号の意味が特定できない。",
     "①の個票を地区に割り付けられない。"
     "①と④を地区で結合できない。", "障害"),
    (3, "②居所変更実態調査に地区の欄がない",
     "所在地の選択肢は「市内／市外」のみである（点検事項No.25）。"
     "「市内」が町単位か広域連合単位かの定義も示されていない。",
     "②を地区に割り付けられない。", "障害"),
    (4, "③介護人材実態調査に所在町の欄がない",
     "事業所票に所在町を記入する欄がない（点検事項No.26）。"
     "北海道の公表データにより事業所名から所在町を特定できるが、"
     "職員個票は事業所単位でしか町に割り付けられない。",
     "③は事業所を経由して町までは割り付けられるが、"
     "地区までは割り付けられない。", "障害"),
    (5, "事業所を結合キーとする方法",
     "①の事業所票と③の事業所票は事業所名で結合できる。"
     "ただし①は居宅介護支援・小規模多機能・地域包括の15件、"
     "③は施設・通所系24件・訪問系3件であり、重なる事業所が少ない。",
     "結合できる事業所が少なく、集計に耐えない。", "障害"),
    (6, "発注者のご意向",
     "令和8年8月5日、4調査を横断したクロス集計は行わず"
     "完了とするご意向をいただいた。",
     "本業務では横断クロス集計を行わない。", "決定"),
    (7, "これに代わる接続の方法",
     "各調査内のクロス集計（別冊「実施済み3調査の受領点検と集計」"
     "11〜13シート、別冊「調査クロス集計・分析」03〜10シート）と、"
     "公表データによる供給構造の分析（本報告書06〜09シート）により、"
     "集計値の水準で調査相互を接続する。",
     "10シートに接続の結果を示した。"
     "個票を突き合わせたものではないため、"
     "関連の指摘にとどめ因果関係を述べない。", "代替"),
    (8, "地区別の記述の扱い",
     "地区別・圏域別の記述は④健康とくらしの調査の"
     "小学校区12地区別集計によることとし、"
     "①②③については行わない。",
     "第1章第7節（圏域の設定）及び第2章第4節（高齢者の状況）は"
     "④のみによる。"
     "①②③は区域全体の記述にとどめる。", "方針"),
    (9, "第11期以降への申し送り",
     "①の調査票に日常生活圏域の選択肢を設ける、"
     "②の「市内／市外」を町名の選択肢に改める、"
     "③の事業所票に所在町の欄を設ける、"
     "の3点により、次期は横断クロス集計が可能になる。",
     "調査票の設計の改善点として記録する。"
     "調査票の設計は発注者が行うものであり、"
     "受託者は提案にとどまる（仕様書４（3））。", "申し送り"),
]:
    fl = {"前提": GRAY, "障害": NG_O, "決定": OK_G, "代替": MID_B,
          "方針": IN_Y, "申し送り": GRAY}[kb]
    r = body(ws, r, [no, nm, cont, res, kb], {5: fl}, height=56,
             align={1: "center", 5: "center"})

note(ws, r + 1,
     "注1）仕様書４（3）は「調査結果の集計・クロス集計・分析」を求めており、"
     "各調査内のクロス集計は実施済みである。"
     "4調査を横断したクロス集計は、"
     "個票を共通の単位に割り付けられないため実施できない。"
     "注2）本シートは、横断クロス集計を行わないことの経緯を"
     "記録として残すためのものである。"
     "業務工程管理表05シートの状態と対応させる。"
     "注3）調査票の設計は発注者が行うものであり"
     "（仕様書は新たな調査の企画・設計・配布・回収・入力及び督促を"
     "行わないと定めている）、"
     "No.9は第11期に向けた提案である。", 5)

# ============================================================ 14
ws = sheet("14_公表データによる補完と留保の解消",
           "公表データによる補完と、留保の解消の状況",
           "本報告書の11シートに掲げた留保のうち、"
           "その後に取得した公表データ及び外部統計により"
           "解消したもの、及び解消しなかったものを整理する。"
           "計画本文に用いる数値の確からしさを更新するためのものである。",
           [4, 24, 34, 34, 24, 12], freeze="A5")

r = lead(ws, 4, "【1　解消した留保】", 6)
r = header(ws, r, ["No.", "留保の内容", "補完に用いたデータ", "解消の内容",
                   "計画本文への反映", "状態"])
for a in HOKAN_KAISHO:
    r = body(ws, r, list(a), {6: OK_G}, height=88, align={1: "center",
                                                          6: "center"})

r = note(ws, r,
         "注1）本シートの「解消」は、"
         "留保が完全になくなったことを意味しない。"
         "No.1〜No.3は公表画面の記入日が事業所により異なり、"
         "1事業所は常勤換算数及び1人当たりの月間サービス提供時間が"
         "0と記入されている（記入漏れとみて集計に用いていない）。"
         "注2）No.5・No.6は代理指標であり、"
         "測るものが当初の定義から変わっている。"
         "指標名の改称を伴う。", 6)

r += 1
r = lead(ws, r, "【2　解消しなかった留保】", 6)
r = header(ws, r, ["No.", "留保の内容", "なぜ解消しないか", "必要なもの",
                   "計画本文での扱い", "状態"])
for a in HOKAN_ZANTEI:
    f = {6: NG_O} if a[5] == "未解消" else (
        {6: OK_G} if a[5] == "測定済み" else
        {6: IN_Y} if a[5] in ("公表待ち", "決定待ち", "一部解消") else {})
    r = body(ws, r, list(a), f, height=76, align={1: "center", 6: "center"})

note(ws, r + 1,
     "注3）留保No.7は本報告書で最も重い留保である。"
     "①在宅生活改善調査から算定するH12（%.1f％）は"
     "第10期の代表KPIに用いる案としているが、"
     "母集団が抽出であることを指標の定義に明記する必要がある。"
     "偏りの大きさと向きは測定できるようになったが、"
     "留保そのものは解消していない。" % H12 +
     "注4）本シートは、公表データの取得により"
     "報告書作成時点から変わった部分を記録するものである。"
     "報告書の本体（02〜13シート）の数値のうち、"
     "10シートの見込量は令和8年9月16日に"
     "サービス見込量の算定 第1次概算の値へ差し替えた"
     "（計画素案 第6章と同じ値にするため）。"
     "そのほかの数値は変更していない。"
     "注5）「測定済み」は、留保そのものは残るが、"
     "その大きさと向きを数値で示せるようになったことを表す。", 6)

del wb["Sheet"]

# ================================================================ 自己点検
# 令和8年9月25日の再レビューによる。
# 分母・時点・同じ指標に2つの値がないことを機械で確かめる。
CHECKS = []


def chk(no, naiyo, kekka, ok):
    CHECKS.append((no, naiyo, kekka, "適合" if ok else "不適合"))


_ALL = [c for _ws in wb for _row in _ws.iter_rows(values_only=True)
        for c in _row if isinstance(c, str)]
_TXT = "".join(_ALL)

chk(1, "在宅生活の維持が困難な者の割合が設問別の有効回答によること",
    "%d人÷%d票＝%.1f％" % (IJI_KONNAN, IJI_YUKO, H06),
    "%.1f％" % H06 in _TXT
    and "%.1f％" % (IJI_KONNAN / RIYO_N * 100) not in _TXT)
chk(2, "H12の分母が当該設問の有効回答であること",
    "%d票÷%d票＝%.1f％" % (H12_BUNSHI, H12_BUNBO, H12),
    "%.1f％" % H12 in _TXT and "20.4％" not in _TXT)
chk(3, "補完と解消の件数が明細と一致すること",
    "%d件＝%s" % (len(HOKAN_KAISHO) + len(HOKAN_ZANTEI), HOKAN_SUM),
    sum(HOKAN_ST.values()) == len(HOKAN_KAISHO) + len(HOKAN_ZANTEI))
chk(4, "外国人職員が全員施設・通所系であること",
    "%d人（施設・通所系%d人・訪問系%d人）"
    % (GAIKOKU,
       sum(j["外国人"] or 0 for j in SIS),
       sum(j["外国人"] or 0 for j in HOU)),
    sum(j["外国人"] or 0 for j in HOU) == 0)
chk(5, "外国人職員の割合の分母が施設・通所系であること",
    "%d人÷%d人＝%.1f％" % (GAIKOKU, N_SIS, GAIKOKU / N_SIS * 100),
    "%.1f％" % (GAIKOKU / N_SIS * 100) in _TXT)
chk(6, "特定施設の施設別の入居率が表に出ていること",
    "調査2施設%.1f％・公表データ%.1f％"
    % (_TK_RES / _TK_CAP * 100, _SW_RES / _SW_CAP * 100),
    "%.1f％" % (_TK_RES / _TK_CAP * 100) in _TXT
    and "%.1f％" % (_SW_RES / _SW_CAP * 100) in _TXT)
chk(7, "特定施設の合計の入居率を参考値としていること",
    "計%d人／名簿の定員%d人" % (_TK_RES + _SW_RES, CAP_TOKUTEI),
    "参考値" in _TXT)
chk(8, "入所者数の3通りが照合できること",
    "入所欄%d人・要介護度別内訳%d人・補完後%d人" % (_t[0], _t[1], _t[2]),
    _t[2] == _t[0] + 58 and _t[1] != _t[0])
chk(9, "H06の基準値を④の値と取り違えていないこと",
    "0.6％は補完指標",
    "補完指標" in _TXT)
chk(10, "利用者票の割合の分母が回収票数であること（複数回答の設問）",
    "外出同行%d件÷%d票＝%.1f％"
    % (SUP["外出同行（通院、買い物など）"], RIYO_N,
       SUP["外出同行（通院、買い物など）"] / RIYO_N * 100),
    "%.1f％" % (SUP["外出同行（通院、買い物など）"] / RIYO_N * 100) in _TXT)

wb.save(OUT)
print("saved:", OUT)
for ws in wb:
    print("  -", ws.title, ws.max_row, "rows")
print("介護職員", N_ALL, "／実事業所", N_JIG, "法人", N_HOJIN,
      "／所見", len(FIND))
_ng = [c for c in CHECKS if c[3] != "適合"]
print("自己点検 %d件：適合%d件・不適合%d件"
      % (len(CHECKS), len(CHECKS) - len(_ng), len(_ng)))
for c in _ng:
    print("  不適合:", c[0], c[1], c[2])
if _ng:
    sys.exit(1)
