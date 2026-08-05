# -*- coding: utf-8 -*-
"""成果品を3つのZIPにまとめる.

  1 計画本体と管理表　　　　計画素案・図表集・管理表・会議資料
  2 第9期の評価と分析　　　 評価、レッドチームレビュー、調査・地域差・推計
  3 分析の残りと設計・参考　他団体比較、設計資料、他計画の参考資料

ZIP名及びファイル名は日本語のため、UTF-8フラグを立てて格納する。
"""

import os
import zipfile

ROOT = "/home/user/repository/output"
DEST = os.path.join(ROOT, "zip")

BUNDLES = [
    ("第10期計画_成果品_1_計画本体と管理表.zip", [
        "第10期計画_成果品一覧.xlsx",
        "第10期介護保険事業計画_素案_第10稿.docx",
        "第10期計画_図表集_白黒.xlsx",
        "第10期計画_必要事項の一覧.xlsx",
        "第10期計画_確認依頼書.docx",
        "第10期計画_業務工程管理表.xlsx",
        "第10期計画_計画素案の別管理表.xlsx",
        "第10期計画_キックオフ会議資料_第6稿.docx",
        "第10期計画_キックオフ会議ヒアリングシート.docx",
        "第10期計画_キックオフ資料の点検結果.xlsx",
    ], ["figures"]),
    ("第10期計画_成果品_2_第9期の評価と分析.zip", [
        "第10期計画_妥当性検証報告書.xlsx",
        "第10期計画_第9期評価のレッドチームレビュー.xlsx",
        "第10期計画_主張の根拠水準の棚卸しと再レビュー.xlsx",
        "第10期計画_調査クロス集計・分析.xlsx",
        "第10期計画_認定率の年齢調整分析.xlsx",
        "第10期計画_地域差の分析.xlsx",
        "第10期計画_人口推計の基礎の検証.xlsx",
        "第10期計画_将来推計_人口と認定者数.xlsx",
        "第10期計画_実施済み3調査の受領点検と集計.xlsx",
        "第10期計画_要介護認定データの確認.xlsx",
        "第10期計画_サービス受給者数データの確認.xlsx",
        "第10期計画_将来推計_第2段階_サービス見込量.xlsx",
        "第10期計画_住まいと施設の公表名簿との突合.xlsx",
    ], []),
    ("第10期計画_成果品_3_分析の残りと設計・参考.zip", [
        "第10期計画_保険料と施策評価の他団体比較.xlsx",
        "第10期計画_保険料の所得段階と低所得者軽減の検証.xlsx",
        "第10期計画_追加調査報告書.xlsx",
        "第10期計画_施策体系新旧対照表.xlsx",
        "第10期計画素案_第9稿→第10稿_修正指示書.xlsx",
        "00_全計画マスター管理表.xlsx",
        "01_共通_基本コラム部品.xlsx",
        "02_高齢者介護保険事業計画.xlsx",
        "03_障がい福祉計画.xlsx",
        "04_こども計画.xlsx",
    ], ["images_basic"]),
]


def add(z, arc, src):
    info = zipfile.ZipInfo.from_file(src, arc)
    info.flag_bits |= 0x800                      # ファイル名がUTF-8である旨
    info.compress_type = zipfile.ZIP_DEFLATED
    with open(src, "rb") as f:
        z.writestr(info, f.read())


os.makedirs(DEST, exist_ok=True)
missing = []
for name, files, dirs in BUNDLES:
    n = 0
    with zipfile.ZipFile(os.path.join(DEST, name), "w",
                         zipfile.ZIP_DEFLATED) as z:
        for fn in files:
            src = os.path.join(ROOT, fn)
            if not os.path.exists(src):
                missing.append(fn)
                continue
            add(z, fn, src)
            n += 1
        for d in dirs:
            base = os.path.join(ROOT, d)
            for fn in sorted(os.listdir(base)):
                add(z, "%s/%s" % (d, fn), os.path.join(base, fn))
                n += 1
    size = os.path.getsize(os.path.join(DEST, name)) / 1024
    print("%s  %d件  %.0f KB" % (name, n, size))

print("欠落: %s" % ("なし" if not missing else "、".join(missing)))
