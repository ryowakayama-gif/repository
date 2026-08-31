# -*- coding: utf-8 -*-
"""成果物とエビデンスを分けてZIPにまとめる.

【成果物】仕様書４が求める成果品。納品対象。
  D1 計画本体と管理表　　　計画素案・骨子案・図表集・管理表・会議資料
  D2 第9期の評価と分析　　 評価、調査・地域差・推計
  D3 設計と他団体比較　　　概要版の構成案、修正指示書、他計画の参考資料

【エビデンス】成果物の数値の根拠となるデータ。納品対象ではないが、
              数値の検証と第11期への引継ぎのために添える。
  E1 データ集　　　　　　　統計・事業所・調査の集計値（Excel 4冊）
  E2 図表の画像　　　　　　計画素案及び報告書に掲載した図表（PNG 79点）

成果物のZIPには図表を入れない（エビデンス側に集約する）。
ZIP名及びファイル名は日本語のため、UTF-8フラグを立てて格納する。
"""

import os
import zipfile

ROOT = "/home/user/repository/output"
DEST = os.path.join(ROOT, "zip")
EDIR = os.path.join(ROOT, "evidence")

# ---------------------------------------------------------------- 成果物
SEIKA = [
    ("第10期計画_成果物_1_計画本体と管理表.zip", [
        "第10期計画_送付資料一覧.xlsx",
        "第10期計画_成果品一覧.xlsx",
        "第10期計画_成果品の送付区分表.xlsx",
        "第10期介護保険事業計画_骨子案.docx",
        "第10期介護保険事業計画_協議用素案_令和8年8月.docx",
        "第10期計画_図表集_白黒.xlsx",
        "第10期計画_必要事項の一覧.xlsx",
        "第10期計画_資料提供依頼_第9期の施策事業実績.xlsx",
        "第10期計画_確認依頼書.docx",
        "第10期計画_業務工程管理表.xlsx",
        "第10期計画_業務進捗報告書_令和8年8月分.docx",
        "第10期計画_計画素案の別管理表.xlsx",
        "第10期計画_3町ヒアリング資料.docx",
        "第10期計画_3町別の論点整理.xlsx",
        "第10期計画_キックオフ会議資料_点検反映版.docx",
        "第10期計画_キックオフ会議ヒアリングシート.docx",
        "第10期計画_キックオフ会議資料_令和8年8月.docx",
        "第10期計画_キックオフ会議資料（更新版）の校正結果.xlsx",
        "第10期計画_キックオフ会議_トークスクリプト_60分.docx",
        "第10期計画_キックオフ会議_ファシリテーションの分岐.xlsx",
        "第10期計画_キックオフ資料の点検結果.xlsx",
        "第10期計画_キックオフ会議議事録_令和8年8月6日.odt",
        "第10期計画_キックオフ会議議事録の校正結果.xlsx",
    ], []),
    ("第10期計画_成果物_2_第9期の評価と分析.zip", [
        "第10期計画_妥当性検証報告書.xlsx",
        "第10期計画_第9期施策と調査・KPIの紐付けレビュー.xlsx",
        "第10期計画_第9期計画の評価・検証_中間報告.docx",
        "第10期計画_中間報告の根拠対照表.xlsx",
        "第10期計画_第9期施策別評価表と暫定評価ルール.xlsx",
        "第10期計画_事業所調査の照会票と確定値管理表.xlsx",
        "第10期計画_調査クロス集計・分析.xlsx",
        "第10期計画_認定率の年齢調整分析.xlsx",
        "第10期計画_地域差の分析.xlsx",
        "第10期計画_世帯構成の突合.xlsx",
        "第10期計画_3町の社会資源一覧との突合.xlsx",
        "第10期計画_従業員数の重複計上の整理.xlsx",
        "第10期計画_給付実績データの受領点検.xlsx",
        "第10期計画_令和8年8月28日受領資料の点検結果.xlsx",
        "第10期計画_見える化総括表の受領点検.xlsx",
        "第10期計画_総合事業実施状況調査の受領点検.xlsx",
        "第10期計画_年報月報の受領点検.xlsx",
        "第10期計画_令和6年度決算書の受領点検.xlsx",
        "第10期計画_基金条例の確認.xlsx",
        "第10期計画_人口推計の基礎の検証.xlsx",
        "第10期計画_人口推計の補正_65歳以上75歳以上の突合.xlsx",
        "第10期計画_将来推計_人口と認定者数.xlsx",
        "第10期計画_将来推計_第2段階_サービス見込量.xlsx",
        "第10期計画_将来推計_第3段階_給付費と保険料.xlsx",
        "第10期計画_将来推計_需要3シナリオの感度表.xlsx",
        "第10期計画_実施済み3調査の受領点検と集計.xlsx",
        "第10期計画_要介護認定データの確認.xlsx",
        "第10期計画_サービス受給者数データの確認.xlsx",
        "第10期計画_住まいと施設の公表名簿との突合.xlsx",
        "第10期計画_アンケート調査の集計分析報告書.xlsx",
        "第10期計画_実施済み調査_結果報告書.docx",
    ], []),
    ("第10期計画_成果物_3_設計と他団体比較.zip", [
        "第10期計画_保険料と施策評価の他団体比較.xlsx",
        "第10期計画_保険料の所得段階と低所得者軽減の検証.xlsx",
        "第10期計画_追加調査報告書.xlsx",
        "第10期計画_施策体系新旧対照表.xlsx",
        "第10期計画_代表KPIの振替案と確認事項の精査.xlsx",
        "第10期計画_概要版の構成案.xlsx",
        "第10期計画素案_修正指示書_令和8年7月.xlsx",
        "第10期計画素案_修正指示書_令和8年8月.xlsx",
        "第10期計画_発注者確認事項一覧.xlsx",
        "第10期計画_レビュー対応記録.xlsx",
        "第10期計画_第9期評価の自己点検記録.xlsx",
        "第10期計画_主張の根拠水準の棚卸しと再レビュー.xlsx",
        "00_全計画マスター管理表.xlsx",
        "01_共通_基本コラム部品.xlsx",
        "02_高齢者介護保険事業計画.xlsx",
        "03_障がい福祉計画.xlsx",
        "04_こども計画.xlsx",
    ], []),
]

# ------------------------------------------------ 第9期評価（工程提出用の抜粋）
# 仕様書５の令和8年8月の工程末に提出する「第9期計画の評価・検証の中間報告」と、
# その根拠となる成果品をまとめたもの。成果物ZIPとは別に、提出単位で束ねる。
TEISHUTSU = [
    ("第10期計画_第9期計画の評価・検証_令和8年8月.zip", ROOT, [
        "第10期計画_第9期計画の評価・検証_中間報告.docx",
        "第10期計画_中間報告の根拠対照表.xlsx",
        "第10期計画_妥当性検証報告書.xlsx",
        "第10期計画_第9期施策と調査・KPIの紐付けレビュー.xlsx",
        "第10期計画_第9期施策別評価表と暫定評価ルール.xlsx",
        "第10期計画_施策体系新旧対照表.xlsx",
        "第10期計画_業務進捗報告書_令和8年8月分.docx",
    ], []),
]

# ------------------------------------------------ 計画素案（読むための一式）
# 素案は［要協議］等90箇所を本文に残す運用のため、
# 決定事項一覧・別管理表・図表集と併せて読む必要がある。
SOAN = [
    ("第10期計画_計画素案一式_協議用素案（令和8年8月時点）.zip", ROOT, [
        "第10期介護保険事業計画_協議用素案_令和8年8月.docx",
        "第10期計画_発注者確認事項一覧.xlsx",
        "第10期計画_計画素案の別管理表.xlsx",
        "第10期計画_図表集_白黒.xlsx",
        "第10期計画素案_修正指示書_令和8年8月.xlsx",
        "第10期介護保険事業計画_骨子案.docx",
        "第10期計画_施策体系新旧対照表.xlsx",
        "第10期計画_代表KPIの振替案と確認事項の精査.xlsx",
    ], []),
]

# ------------------------------------------------ 中間報告の検証用エビデンス
# 第三者が中間報告の記述を検証するための一式。
# 報告本体・根拠対照表・根拠となる成果品・原データを1つにまとめる。
# 個票（個人情報）及び見える化δは収録しない（根拠対照表 05シートに明示）。
KENSHO_ROOT = [
    "第10期計画_第9期計画の評価・検証_中間報告.docx",
    "第10期計画_中間報告の根拠対照表.xlsx",
    "第10期計画_妥当性検証報告書.xlsx",
    "第10期計画_第9期施策と調査・KPIの紐付けレビュー.xlsx",
    "第10期計画_第9期施策別評価表と暫定評価ルール.xlsx",
    "第10期計画_施策体系新旧対照表.xlsx",
    "第10期計画_認定率の年齢調整分析.xlsx",
    "第10期計画_地域差の分析.xlsx",
    "第10期計画_調査クロス集計・分析.xlsx",
    "第10期計画_アンケート調査の集計分析報告書.xlsx",
    "第10期計画_サービス受給者数データの確認.xlsx",
    "第10期計画_要介護認定データの確認.xlsx",
    "第10期計画_業務進捗報告書_令和8年8月分.docx",
]
KENSHO_EVID = [
    "第10期計画_エビデンス_0_索引.xlsx",
    "第10期計画_エビデンス_1_統計データ.xlsx",
    "第10期計画_エビデンス_2_事業所と公表情報.xlsx",
    "第10期計画_エビデンス_3_調査の集計値.xlsx",
]

# ---------------------------------------------------------------- エビデンス
EVID = [
    ("第10期計画_エビデンス_1_データ集.zip", EDIR, [
        "第10期計画_エビデンス_0_索引.xlsx",
        "第10期計画_エビデンス_1_統計データ.xlsx",
        "第10期計画_エビデンス_2_事業所と公表情報.xlsx",
        "第10期計画_エビデンス_3_調査の集計値.xlsx",
    ], []),
    ("第10期計画_エビデンス_2_図表の画像.zip", ROOT, [],
     ["figures", "figures_report", "images_basic"]),
]


def build_by_dispatch(missing):
    """送付区分により、送付用と内部保管用のZIPを分けて作る。"""
    from data_dispatch import DISPATCH
    groups = {"送付": [], "条件付き": [], "内部保管": []}
    for fn, (kb, _why) in DISPATCH.items():
        if kb in groups and os.path.exists(os.path.join(ROOT, fn)):
            groups[kb].append(fn)

    n_sofu = 0
    nm = "第10期計画_送付用_令和8年8月.zip"
    with zipfile.ZipFile(os.path.join(DEST, nm), "w",
                         zipfile.ZIP_DEFLATED) as z:
        for sub, kb in [("01_送付可能（留保明記済み）", "送付"),
                        ("02_お諮りする内容を含むもの", "条件付き")]:
            for fn in sorted(groups[kb]):
                add(z, "%s/%s" % (sub, fn), os.path.join(ROOT, fn))
                n_sofu += 1
    print("  %-52s %3d件 %6.0f KB"
          % (nm, n_sofu, os.path.getsize(os.path.join(DEST, nm)) / 1024))

    # 計画本文に掲載する図表の画像を同送する
    fig = os.path.join(ROOT, "figures")
    n_png = 0
    if os.path.isdir(fig):
        with zipfile.ZipFile(os.path.join(DEST, nm), "a",
                             zipfile.ZIP_DEFLATED) as z:
            for f in sorted(os.listdir(fig)):
                add(z, "03_図表の画像/" + f, os.path.join(fig, f))
                n_png += 1
                n_sofu += 1
        print("  %-52s 図表の画像 %d点を同梱" % ("", n_png))

    doc = "/home/user/repository/docs/令和8年8月_送付状.md"
    if os.path.exists(doc):
        with zipfile.ZipFile(os.path.join(DEST, nm), "a",
                             zipfile.ZIP_DEFLATED) as z:
            add(z, "00_送付状_令和8年8月.md", doc)
            n_sofu += 1
        print("  %-52s 送付状を同梱" % "")

    nm2 = "第10期計画_内部保管_作業記録.zip"
    n_naibu = 0
    with zipfile.ZipFile(os.path.join(DEST, nm2), "w",
                         zipfile.ZIP_DEFLATED) as z:
        for fn in sorted(groups["内部保管"]):
            add(z, fn, os.path.join(ROOT, fn))
            n_naibu += 1
    print("  %-52s %3d件 %6.0f KB"
          % (nm2, n_naibu, os.path.getsize(os.path.join(DEST, nm2)) / 1024))
    return n_sofu, n_naibu


def add(z, arc, src):
    info = zipfile.ZipInfo.from_file(src, arc)
    info.flag_bits |= 0x800                      # ファイル名がUTF-8である旨
    info.compress_type = zipfile.ZIP_DEFLATED
    with open(src, "rb") as f:
        z.writestr(info, f.read())


def build(name, base, files, dirs, missing):
    n = 0
    with zipfile.ZipFile(os.path.join(DEST, name), "w",
                         zipfile.ZIP_DEFLATED) as z:
        for fn in files:
            src = os.path.join(base, fn)
            if not os.path.exists(src):
                missing.append(fn)
                continue
            add(z, fn, src)
            n += 1
        for d in dirs:
            p = os.path.join(base, d)
            if not os.path.isdir(p):
                missing.append(d + "/")
                continue
            for fn in sorted(os.listdir(p)):
                add(z, "%s/%s" % (d, fn), os.path.join(p, fn))
                n += 1
    size = os.path.getsize(os.path.join(DEST, name)) / 1024
    print("  %-52s %3d件 %6.0f KB" % (name, n, size))
    return n


if __name__ == "__main__":
    os.makedirs(DEST, exist_ok=True)
    # 旧名のZIPを残さない
    for fn in os.listdir(DEST):
        if fn.startswith("第10期計画_成果品_"):
            os.remove(os.path.join(DEST, fn))

    missing = []
    print("【成果物】")
    ns = sum(build(nm, ROOT, f, d, missing) for nm, f, d in SEIKA)
    print("【エビデンス】")
    ne = sum(build(nm, base, f, d, missing) for nm, base, f, d in EVID)
    print("【計画素案（読むための一式）】")
    nso = sum(build(nm, base, f, d, missing) for nm, base, f, d in SOAN)
    print("【工程提出用（第9期評価）】")
    nt = sum(build(nm, base, f, d, missing) for nm, base, f, d in TEISHUTSU)

    print("【送付区分別】")
    n_so, n_na = build_by_dispatch(missing)

    # 検証用は1つのZIPに2つの階層で格納する
    print("【中間報告の検証用エビデンス】")
    _nm = "第10期計画_中間報告の検証用エビデンス.zip"
    nk = 0
    with zipfile.ZipFile(os.path.join(DEST, _nm), "w",
                         zipfile.ZIP_DEFLATED) as z:
        for fn in KENSHO_ROOT:
            src = os.path.join(ROOT, fn)
            if not os.path.exists(src):
                missing.append(fn)
                continue
            add(z, "01_報告と根拠/" + fn, src)
            nk += 1
        for fn in KENSHO_EVID:
            src = os.path.join(EDIR, fn)
            if not os.path.exists(src):
                missing.append(fn)
                continue
            add(z, "02_原データ/" + fn, src)
            nk += 1
        d = os.path.join(ROOT, "figures")
        for fn in sorted(os.listdir(d)):
            add(z, "03_図表/" + fn, os.path.join(d, fn))
            nk += 1
    _sz = os.path.getsize(os.path.join(DEST, _nm)) / 1024
    print("  %-52s %3d件 %6.0f KB" % (_nm, nk, _sz))
    print()
    print("成果物 %d件／エビデンス %d件／素案一式 %d件／"
          "工程提出用 %d件／検証用 %d件" % (ns, ne, nso, nt, nk))
    print("送付用 %d件／内部保管 %d件" % (n_so, n_na))
    print("欠落: %s" % ("なし" if not missing else "、".join(missing)))

    # 成果品一覧との突合（一覧にあってZIPに入っていないものを検出する）
    import io as _io
    import runpy as _runpy
    import sys as _sys
    _buf, _old = _io.StringIO(), _sys.stdout
    _sys.stdout = _buf
    try:
        _G = _runpy.run_path("build_deliverable_index.py")
    finally:
        _sys.stdout = _old
    idx = {fn for _c, fn, _a, _b, _d in _G["ITEMS"]}
    zipped = set()
    for _nm, _f, _d2 in SEIKA:
        zipped |= set(_f)
    gap = sorted(idx - zipped)
    print("成果品一覧との突合: %s"
          % ("全件がZIPに入っている" if not gap else "ZIP未収録 " + "、".join(gap)))

    # 送付用ZIPに内部保管のものが混ざっていないことを確かめる
    from data_dispatch import dispatch_of
    with zipfile.ZipFile(
            os.path.join(DEST, "第10期計画_送付用_令和8年8月.zip")) as z:
        bad = [a for a in z.namelist()
               if dispatch_of(os.path.basename(a))[0] in ("内部保管", "対象外")]
    print("送付用ZIPの点検: %s"
          % ("内部保管・対象外は含まれていない" if not bad
             else "混入 " + "、".join(bad)))
