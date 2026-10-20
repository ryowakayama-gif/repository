# -*- coding: utf-8 -*-
"""北海道の施設名簿（定員を持つもの）により確かめた構成3町の施設と定員。

令和8年10月20日、北海道のウェブサイト
「介護保険事業所・老人福祉施設等一覧」から
**令和8年9月1日現在の名簿**を取得して確かめた。

**電話番号及びＦＡＸ番号は収録しない**（CLAUDE.md §1）。
施設名・設置者名（法人名）・定員のみを収める。

**「全事業所一覧」（令和8年8月31日現在）に定員の列はない。**
列は No.／サービス／申請者名称／事業所名称／事業所番号／郵便番号／
所在地／電話番号／ＦＡＸ番号／通常の事業実施地域である。
**認知症対応型共同生活介護の定員（確認事項No.78）はこの名簿では解けない。**
定員を持つ名簿は特別養護老人ホーム・養護老人ホーム・軽費老人ホーム・
届出済有料老人ホーム・未届有料老人ホームである。

確かめたこと

| 事項 | 当方の収録値 | 名簿（令和8年9月1日現在） |
|---|---|---|
| 介護老人福祉施設（広域型）の定員 | 160人 | **160人（一致）** |
| 地域密着型介護老人福祉施設の定員 | 62人 | **62人（一致）** |
| 特定施設入居者生活介護の定員 | 156人 | **156人（一致）** |
| 住宅型有料老人ホーム | 9施設223人 | **9施設223人（一致）** |
| 軽費老人ホーム | 1施設50人 | **1施設50人（一致）** |
| 未届有料老人ホーム | ― | **構成3町に該当する記載はない** |
"""

KAKUNIN_BI = "令和8年10月20日"
MEIBO_BI = "令和8年9月1日現在"
ZEN_BI = "令和8年8月31日現在"   # 全事業所一覧（定員の列はない）

# ---------------------------------------------------------------------------
# 特別養護老人ホーム名簿　（施設名, 町, 定員, ショート専用定員, 空床利用, 地域密着型）
# ---------------------------------------------------------------------------
TOKUYO = [
    ("東神楽町特別養護老人ホームアゼリアハイツ", "東神楽町", 50, 10, True, False),
    ("東神楽町特別養護老人ホームアゼリアハイツ（ユニット型）", "東神楽町", 20, None, True, True),
    ("特別養護老人ホーム東川町羽衣園", "東川町", 50, 5, True, False),
    ("地域密着型介護老人福祉施設 美瑛慈光園", "美瑛町", 24, 4, True, True),
    ("特別養護老人ホーム美瑛慈光園", "美瑛町", 60, None, False, False),
    ("美瑛慈光園サテライト特養 燈", "美瑛町", 18, 2, True, True),
]

# ---------------------------------------------------------------------------
# 軽費老人ホーム名簿　（区分, 施設名, 町, 定員, 特定指定, 地域密着型特定指定）
# ---------------------------------------------------------------------------
KEIHI = [
    ("ケ", "ケアハウスびえい", "美瑛町", 50, False, False),
]

# ---------------------------------------------------------------------------
# 届出済有料老人ホーム一覧　（類型, 施設名, 町, 設置者, 居室数, 定員数, 特定施設の指定）
#   居室数の「18室」のような表記は整数に直して収めた
# ---------------------------------------------------------------------------
YURYO = [
    ("介護付", "高齢者ふれあいハウスファミリー", "東川町", "(有)アスト", 36, 36, True),
    ("介護付", "ゆう", "東川町", "(株)栄友", 20, 20, True),
    ("介護付", "さわやか東神楽館", "東神楽町", "(株)さわやか倶楽部", 100, 100, True),
    ("住宅型", "花時計", "東神楽町", "(医)回生会", 30, 30, False),
    ("住宅型", "ひばりの森", "東神楽町", "(株)オレンジサポート", 31, 43, False),
    ("住宅型", "ひがしかぐらふらわーはうす", "東神楽町", "(特非)まこと", 18, 18, False),
    ("住宅型", "華", "東川町", "(有)フラワー", 30, 31, False),
    ("住宅型", "縁結び", "東川町", "(株)エステートジョイ", 20, 22, False),
    ("住宅型", "びえいの郷", "美瑛町", "シルバーハウス(株)", 30, 30, False),
    ("住宅型", "わくらば", "東川町", "(株)北海道光健康社", 6, 6, False),
    ("住宅型", "わくらばⅡ", "東川町", "(株)北海道光健康社", 18, 18, False),
    ("住宅型", "恩送り", "東川町", "ルラシオン(同)", 24, 25, False),
]

# 未届有料老人ホーム一覧（令和8年9月1日現在）に構成3町の記載はない
MITODOKE = []

# 定員の列を持たない名簿
TEIIN_NASHI = [
    ("全事業所一覧", ZEN_BI,
     "列は No.／サービス／申請者名称／事業所名称／事業所番号／郵便番号／"
     "所在地／電話番号／ＦＡＸ番号／通常の事業実施地域。"
     "定員の列はない。当方の `data_hokkaido_shitei` の基礎である。"),
]


def tokuyo_kouiki():
    """介護老人福祉施設（広域型）の定員の合計。"""
    return sum(t for _, _, t, _, _, m in TOKUYO if not m)


def tokuyo_mitchaku():
    """地域密着型介護老人福祉施設の定員の合計。"""
    return sum(t for _, _, t, _, _, m in TOKUYO if m)


def tokuyo_machi(mitchaku=True):
    """町別の特別養護老人ホームの定員。"""
    d = {}
    for _, machi, t, _s, _k, m in TOKUYO:
        if m == mitchaku:
            d[machi] = d.get(machi, 0) + t
    return d


def short_teiin():
    """ショート専用の定員の合計。"""
    return sum(s for _, _, _, s, _, _ in TOKUYO if s)


def yuryo_sum(ruikei):
    """類型ごとの施設数と定員の合計。"""
    xs = [x for x in YURYO if x[0] == ruikei]
    return len(xs), sum(x[5] for x in xs)


def jutakugata():
    """指定を受けない住まいのうち住宅型有料老人ホーム（施設数, 定員）。"""
    return yuryo_sum("住宅型")


if __name__ == "__main__":
    import sys
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    ng = 0
    print("北海道の施設名簿により確かめた構成3町の施設（%s／名簿は%s）"
          % (KAKUNIN_BI, MEIBO_BI))
    print("特別養護老人ホーム　%d施設" % len(TOKUYO))
    print("  広域型 %d人／地域密着型 %d人／ショート専用 %d人"
          % (tokuyo_kouiki(), tokuyo_mitchaku(), short_teiin()))
    print("  地域密着型の町別　%s" % tokuyo_machi(True))
    print("  広域型の町別　　　%s" % tokuyo_machi(False))
    print("軽費老人ホーム　%d施設 %d人" % (len(KEIHI), sum(x[3] for x in KEIHI)))
    for r in ("介護付", "住宅型"):
        n, t = yuryo_sum(r)
        print("有料老人ホーム（%s）　%d施設 %d人" % (r, n, t))
    print("未届有料老人ホーム　構成3町の記載 %d件" % len(MITODOKE))

    # 点検1　介護老人福祉施設（広域型）の定員が算定の TEIIN_MAP と一致する
    import ast, io as _io
    try:
        src = _io.open("build_mikomiryo_santei.py", encoding="utf-8").read()
        tm = None
        for nd in ast.parse(src).body:
            if isinstance(nd, ast.Assign) and any(
                    getattr(t, "id", None) == "TEIIN_MAP" for t in nd.targets):
                tm = ast.literal_eval(nd.value)
        if tm is None:
            print("不適合　TEIIN_MAP を読めない"); ng += 1
        else:
            pairs = [("施設サービス 介護老人福祉施設", tokuyo_kouiki()),
                     ("施設サービス 地域密着型介護老人福祉施設入所者生活介護",
                      tokuyo_mitchaku()),
                     ("居住系サービス 特定施設入居者生活介護", yuryo_sum("介護付")[1])]
            for k, v in pairs:
                if k not in tm:
                    print("不適合　TEIIN_MAP に無い：%s" % k); ng += 1
                elif tm[k][0] != v:
                    print("不適合　定員が合わない：%s 算定%d／名簿%d"
                          % (k, tm[k][0], v)); ng += 1
    except OSError:
        print("未実施　build_mikomiryo_santei.py を読めない")
    # 点検2　住宅型有料老人ホームは9施設223人である
    if jutakugata() != (9, 223):
        print("不適合　住宅型有料老人ホームが9施設223人でない：%s" % (jutakugata(),)); ng += 1
    # 点検3　軽費老人ホームは1施設50人である
    if (len(KEIHI), sum(x[3] for x in KEIHI)) != (1, 50):
        print("不適合　軽費老人ホームが1施設50人でない"); ng += 1
    # 点検4　地域密着型の町別の和が合計と一致する
    if sum(tokuyo_machi(True).values()) != tokuyo_mitchaku():
        print("不適合　地域密着型の町別の和が合計と合わない"); ng += 1
    # 点検5　介護付の全件が特定施設の指定を受けている
    for x in YURYO:
        if (x[0] == "介護付") != x[6]:
            print("不適合　類型と特定施設の指定が食い違う：%s" % x[1]); ng += 1
    # 点検6　電話番号・メールアドレスを収めていない
    import re as _re
    blob = repr(TOKUYO) + repr(KEIHI) + repr(YURYO) + repr(TEIIN_NASHI)
    for pat, mei in ((r"0\d{1,4}-\d{1,4}-\d{3,4}", "電話番号"),
                     (r"[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}", "メールアドレス")):
        if _re.search(pat, blob):
            print("不適合　%sの形がある" % mei); ng += 1
    # 点検7　町名は構成3町に限る
    for x in TOKUYO + YURYO:
        machi = x[1] if len(x) == 6 and isinstance(x[2], int) else x[2]
        if machi not in ("東川町", "美瑛町", "東神楽町"):
            print("不適合　構成3町以外の町：%s" % machi); ng += 1
    print("不適合 %d件" % ng)
    sys.exit(1 if ng else 0)
