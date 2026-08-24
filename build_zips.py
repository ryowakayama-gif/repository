# -*- coding: utf-8 -*-
"""成果物とエビデンスを分けてZIPにまとめる.

【成果物】仕様書４が求める成果品。納品対象。
  D1 計画本体と管理表　　　計画素案・骨子案・図表集・管理表・会議資料
  D2 第9期の評価と分析　　 評価、レッドチームレビュー、調査・地域差・推計
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
        "第10期計画_成果品一覧.xlsx",
        "第10期介護保険事業計画_骨子案.docx",
        "第10期介護保険事業計画_素案_第12稿.docx",
        "第10期計画_図表集_白黒.xlsx",
        "第10期計画_必要事項の一覧.xlsx",
        "第10期計画_資料提供依頼_第9期の施策事業実績.xlsx",
        "第10期計画_確認依頼書.docx",
        "第10期計画_業務工程管理表.xlsx",
        "第10期計画_業務進捗報告書_令和8年8月分.docx",
        "第10期計画_計画素案の別管理表.xlsx",
        "第10期計画_3町ヒアリング資料.docx",
        "第10期計画_キックオフ会議資料_第6稿.docx",
        "第10期計画_キックオフ会議ヒアリングシート.docx",
        "第10期計画_キックオフ会議資料（更新版）_校正反映.docx",
        "第10期計画_キックオフ会議資料（更新版）の校正結果.xlsx",
        "第10期計画_キックオフ会議_トークスクリプト_60分.docx",
        "第10期計画_キックオフ会議_ファシリテーションの分岐.xlsx",
        "第10期計画_キックオフ資料の点検結果.xlsx",
        "第10期計画_キックオフ会議議事録_校正反映.odt",
        "第10期計画_キックオフ会議議事録の校正結果.xlsx",
    ], []),
    ("第10期計画_成果物_2_第9期の評価と分析.zip", [
        "第10期計画_妥当性検証報告書.xlsx",
        "第10期計画_第9期評価のレッドチームレビュー.xlsx",
        "第10期計画_第9期施策と調査・KPIの紐付けレビュー.xlsx",
        "第10期計画_第9期施策別評価表と暫定評価ルール.xlsx",
        "第10期計画_事業所調査の照会票と確定値管理表.xlsx",
        "第10期計画_主張の根拠水準の棚卸しと再レビュー.xlsx",
        "第10期計画_調査クロス集計・分析.xlsx",
        "第10期計画_認定率の年齢調整分析.xlsx",
        "第10期計画_地域差の分析.xlsx",
        "第10期計画_世帯構成の突合.xlsx",
        "第10期計画_3町の社会資源一覧との突合.xlsx",
        "第10期計画_人口推計の基礎の検証.xlsx",
        "第10期計画_人口推計の補正_65歳以上75歳以上の突合.xlsx",
        "第10期計画_将来推計_人口と認定者数.xlsx",
        "第10期計画_将来推計_第2段階_サービス見込量.xlsx",
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
        "第10期計画素案_第9稿→第10稿_修正指示書.xlsx",
        "第10期計画素案_第11稿→第12稿_修正指示書.xlsx",
        "第10期計画_レビュー指摘への対応と決定事項一覧.xlsx",
        "00_全計画マスター管理表.xlsx",
        "01_共通_基本コラム部品.xlsx",
        "02_高齢者介護保険事業計画.xlsx",
        "03_障がい福祉計画.xlsx",
        "04_こども計画.xlsx",
    ], []),
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
    print()
    print("成果物 %d件／エビデンス %d件" % (ns, ne))
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
