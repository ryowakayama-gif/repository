# -*- coding: utf-8 -*-
"""
北塩原村　アンケートの値の一か所（障がい者調査 有効回答40件の集計）

なぜ置くか
  令和8年10月21日に報告書用集計表が改訂され、障がい者調査の有効回答数が
  38件から40件になった（有効回答率 22.1％→23.3％）。正本・概要版・
  協議会資料・照会・KPI定義票は、それぞれの生成器が割合を文字で
  持っていたため、38件の値が残った（令和8年10月24日の他ＡＩの横断
  レビュー RT-01・RT-10）。
  割合はすべてここから読む。生成器に割合を書き写さない。

値の決め方
  ・単一の選択肢 … 集計_障がい者 シートの件数÷回答者数
  ・合計（例：在宅を希望＝「今の自宅で」＋「家族と一緒に」） … 件数の和÷回答者数
    （割合の和ではない。丸めの差を持ち込まないため）
  ・利用者に限った満足度 … 「満足している」÷（回答者－「利用していない」－無回答）
  ・地区別・年齢別 … 4章_クロス集計 の件数÷その区分の回答者数
  ・障がい別（手帳による区分）と障がい児調査は、増えた2件がどの手帳も
    持たない方であるため動かない（build_kitashiobara_anke_kaishuu で確認済み）。
    それでもここで新旧の両方を読み、動かないことを毎回確かめる。
"""

import os
import re

from openpyxl import load_workbook

REPO_ROOT = os.path.dirname(os.path.abspath(__file__))
KYU = f"{REPO_ROOT}/source/アンケート結果/報告書用集計表.xlsx"
SHIN = (f"{REPO_ROOT}/source/アンケート結果/"
        "報告書用集計表_令和8年10月21日版_回収40.xlsx")

_CACHE = {}


def _tally(path):
    """(問, 選択肢) → (件数, 回答者数)"""
    if ("t", path) in _CACHE:
        return _CACHE[("t", path)]
    wb = load_workbook(path, read_only=True, data_only=True)
    out = {}
    for r in wb["集計_障がい者"].iter_rows(values_only=True):
        if not r or not r[0] or not isinstance(r[3], (int, float)):
            continue
        out[(str(r[0]).strip(), str(r[2]).strip())] = (int(r[3]), int(r[5]))
    wb.close()
    _CACHE[("t", path)] = out
    return out


def _cross(path, title_key):
    """4章_クロス集計 の件数の表。title_key（例：'問3×問14'）を含む表題の表。
    → (区分の列名の並び, {区分: 回答者数}, {項目: {区分: 件数}})"""
    k = ("c", path, title_key)
    if k in _CACHE:
        return _CACHE[k]
    wb = load_workbook(path, read_only=True, data_only=True)
    rows = [[x for x in r if x is not None]
            for r in wb["4章_クロス集計"].iter_rows(values_only=True)]
    wb.close()
    rows = [r for r in rows if r]
    for i, r in enumerate(rows):
        if isinstance(r[0], str) and r[0].startswith("表") and title_key in r[0]:
            head = rows[i + 2]
            assert head[0] == "区分", head
            cols = [str(x) for x in head[1:]]
            nrow = rows[i + 3]
            assert str(nrow[0]).startswith("回答者数"), nrow
            n = dict(zip(cols, [int(x) for x in nrow[1:]]))
            items = {}
            for rr in rows[i + 4:]:
                if str(rr[0]).startswith("【割合】"):
                    break
                items[str(rr[0])] = dict(zip(cols, [int(x) for x in rr[1:]]))
            _CACHE[k] = (cols, n, items)
            return _CACHE[k]
    raise LookupError(f"クロス集計の表が無い: {title_key}")


def _gaiyou(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    out = {}
    for r in wb["1-2章_概要・要点"].iter_rows(min_row=1, max_row=12,
                                           values_only=True):
        for i, v in enumerate(r or ()):
            if v is not None and str(v).strip() in ("配布数", "有効回答数"):
                ato = [x for x in r[i + 1:] if x is not None]
                out[str(v).strip()] = (int(ato[0]), int(ato[1]))
                break
    wb.close()
    return out


# ===========================================================================
# 指標の定義（文書で用いる呼び名 → 求め方）
#   ("q", [(問, 選択肢), ...])        … 件数の和 ÷ その問の回答者数
#   ("x", 表題キー, 区分の列名 又は [列名...], 項目)
#   ("manzoku",)                        … 利用者に限った満足度
#   ("manzoku2",)                       … 無回答を除く全員を分母とする満足度
#   ("kaishuu", 障がい者/障がい児/合計)  … 有効回答率
#   ("betsu", 値)                       … 障がい別・障がい児調査（動かない値）
# ===========================================================================
M = {
    "有効回答率_障がい者": ("kaishuu", 0),
    "有効回答率_障がい児": ("kaishuu", 1),
    "有効回答率_合計": ("kaishuu", 2),
    "65歳以上": ("q", [("問2", "65歳以上")]),
    "身体障がい（主たる）": ("q", [("問6", "身体障がい")]),
    "困りごと_特にない": ("q", [("問10", "特にない")]),
    "困りごと_将来の生活": ("q", [("問10", "将来の生活")]),
    "困りごと_将来の生活_65歳未満": ("x", "問2×問10",
                                 ["30～39歳", "40～49歳", "50～64歳"],
                                 "将来の生活"),
    "暮らし_今の自宅": ("q", [("問11", "今の自宅で暮らしたい")]),
    "暮らし_家族と一緒": ("q", [("問11", "家族と一緒に暮らしたい")]),
    "在宅を希望": ("q", [("問11", "今の自宅で暮らしたい"),
                       ("問11", "家族と一緒に暮らしたい")]),
    "外出_冬季": ("q", [("問14", "冬季・積雪時に外出しにくい")]),
    "外出_冬季_裏磐梯": ("x", "問3×問14", "裏磐梯地区", "冬季・積雪時に外出しにくい"),
    "外出_付き添い": ("q", [("問14", "家族等の付き添いが必要")]),
    "移動_自分で運転": ("q", [("問13", "自分で運転")]),
    "移動_家族の送迎": ("q", [("問13", "家族・知人の送迎")]),
    "通院している": ("q", [("問15", "通院している")]),
    "意思決定支援_必要": ("q", [("問21", "必要である（現在すでに支援を受けている、"
                                     "または今後受けたい）"),
                             ("問21", "今は必要ないが、将来的には必要になると思う")]),
    "相談相手_家族": ("q", [("問22", "家族・親族")]),
    "相談相手_相談支援専門員": ("q", [("問22", "相談支援専門員")]),
    "相談相手_村役場": ("q", [("問22", "村役場")]),
    "相談への期待_情報提供": ("q", [("問26", "わかりやすい情報提供")]),
    "相談への期待_サービス利用": ("q", [("問26", "サービス利用の相談")]),
    "相談への期待_将来の暮らし": ("q", [("問26", "将来の暮らしの相談")]),
    "知らない_重度訪問介護": ("q", [("問27-2", "重度訪問介護")]),
    "知らない_同行援護・行動援護": ("q", [("問27-2", "同行援護・行動援護")]),
    "知らない_就労選択支援": ("q", [("問27-2", "就労選択支援")]),
    "知らない_移動支援": ("q", [("問27-2", "移動支援")]),
    "利用したい_施設入所支援": ("q", [("問27-3", "施設入所支援")]),
    "利用したい_グループホーム": ("q", [("問27-3", "グループホーム")]),
    "利用したい_短期入所": ("q", [("問27-3", "短期入所（ショートステイ）")]),
    "拠点_緊急時の受入れ": ("q", [("問29", "緊急時の受入れ")]),
    "拠点_緊急時の相談": ("q", [("問29", "緊急時の相談")]),
    "満足（利用者）": ("manzoku",),
    "満足（無回答を除く全員）": ("manzoku2",),
    "日中_自宅": ("q", [("問31", "自宅で過ごすことが多い")]),
    "働き方_短時間": ("q", [("問32", "短時間・体調に合わせて働きたい")]),
    "働き方_一般企業": ("q", [("問32", "一般企業等で働きたい")]),
    "就労支援_職場の理解": ("q", [("問33", "職場の障がい理解")]),
    "就労支援_体調に合わせた勤務": ("q", [("問33", "体調に合わせた勤務")]),
    "差別を感じた": ("q", [("問34", "週に１回程度"), ("問34", "月に１回程度"),
                          ("問34", "年に１回程度")]),
    "配慮_わかりやすい説明": ("q", [("問36", "わかりやすい説明")]),
    "配慮_手続きの支援": ("q", [("問36", "手続きの支援")]),
    "成年後見_知らない": ("q", [("問37", "知らない")]),
    "成年後見_名前だけ": ("q", [("問37", "名前だけ聞いたことがある")]),
    "成年後見_十分に知らない": ("q", [("問37", "知らない"),
                                   ("問37", "名前だけ聞いたことがある")]),
    "災害不安_避難先での生活": ("q", [("問38", "避難先での生活")]),
    "災害不安_薬・医療機器": ("q", [("問38", "薬・医療機器")]),
    "災害不安_冬季の避難": ("q", [("問38", "冬季・積雪時の避難")]),
    "個別避難計画_知らない": ("q", [("問39", "知らない")]),
    "施策_相談支援体制": ("q", [("問44", "相談支援体制")]),
    "施策_災害時支援": ("q", [("問44", "災害時支援")]),
    "施策_在宅生活支援": ("q", [("問44", "在宅生活支援")]),
    "施策_移動支援": ("q", [("問44", "移動支援")]),
    "施策_事業所の確保": ("q", [("問44", "サービス事業所の確保")]),
}


def _kazu(name, path):
    """(件数, 分母)"""
    d = M[name]
    if d[0] == "q":
        t = _tally(path)
        ken = sum(t[k][0] for k in d[1])
        bo = {t[k][1] for k in d[1]}
        assert len(bo) == 1, (name, bo)
        return ken, bo.pop()
    if d[0] == "x":
        cols, n, items = _cross(path, d[1])
        cs = d[2] if isinstance(d[2], list) else [d[2]]
        return sum(items[d[3]][c] for c in cs), sum(n[c] for c in cs)
    if d[0] == "manzoku":
        t = _tally(path)
        man = t[("問30", "満足している")]
        bo = man[1] - t[("問30", "利用していない")][0] - t[("問30", "無回答")][0]
        return man[0], bo
    if d[0] == "manzoku2":
        #   分母に「利用していない」を含める読み方（第4次計画の指標③の
        #   分母が決まっていないため、両方の読み方を示す）
        t = _tally(path)
        man = t[("問30", "満足している")]
        return man[0], man[1] - t[("問30", "無回答")][0]
    if d[0] == "kaishuu":
        g = _gaiyou(path)
        hai, yuu = g["配布数"], g["有効回答数"]
        if d[1] == 2:
            return yuu[0] + yuu[1], hai[0] + hai[1]
        return yuu[d[1]], hai[d[1]]
    raise ValueError(name)


def kazu(name, kyu=False):
    return _kazu(name, KYU if kyu else SHIN)


def ritsu(name, kyu=False):
    ken, bo = kazu(name, kyu)
    return ken / bo


def pc(name, kyu=False, kigou=True):
    """「35.0％」の形（小数1桁。四捨五入）。"""
    v = ritsu(name, kyu) * 100
    s = f"{_shisha(v):.1f}"
    return s + ("％" if kigou else "")


def _shisha(v):
    """小数第2位を四捨五入（float の銀行丸めを避ける）。"""
    from decimal import Decimal, ROUND_HALF_UP
    return float(Decimal(repr(v)).quantize(Decimal("0.1"), ROUND_HALF_UP))


def n(kyu=False):
    """障がい者調査の有効回答数。"""
    return kazu("有効回答率_障がい者", kyu)[0]


def kakomu(text):
    """文中の「（n=38）」等の古い表記を検出するための正規表現。"""
    return re.findall(r"[nｎ]\s*[=＝]\s*38", text)
