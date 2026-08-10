# -*- coding: utf-8 -*-
"""発注者ご作成のキックオフ会議議事録の校正.

原本
  99672316-_________10_____________________1.odt
  （大雪地区広域連合 第10期 介護保険事業計画策定支援業務　キックオフ会議
    令和8年8月6日開催）

誤字脱字及び明らかな表記の誤りを反映した版を作成する。
原本の書式を保つため、ODF の content.xml の文字列のみを置換し、
段落・表・スタイルには手を加えない。

判断を要する事項（文体・会議名の呼称・実質の論点）は反映せず、
校正結果の一覧に「提案」「要協議」として掲げる。

出力
  output/第10期計画_キックオフ会議議事録_校正反映.odt
  output/第10期計画_キックオフ会議議事録の校正結果.xlsx
"""

import os
import shutil
import zipfile

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SRC = ("/root/.claude/uploads/54f527c9-842b-534d-b822-e5a6c91f837c/"
       "99672316-_________10_____________________1.odt")
OUT = ("/home/user/repository/output/"
       "第10期計画_キックオフ会議議事録_校正反映.odt")
OUTX = ("/home/user/repository/output/"
        "第10期計画_キックオフ会議議事録の校正結果.xlsx")

# ------------------------------------------------------------ 反映する校正
# (区分, 箇所, 置換前, 置換後, 理由)
FIX = [
    ("誤字", "2. 基礎データ",
     "前回の保険慮", "前回の保険料",
     "「保険慮」は「保険料」の誤りである。"),
    ("誤字", "2. 既存調査／4. 令和8年9月",
     "居宅変更実態調査", "居所変更実態調査",
     "調査の正式名称は「居所変更実態調査」である。2か所。"),
    ("表記", "2. 計画の位置付け／基礎データ、3-2",
     "大雪広域連合", "大雪地区広域連合",
     "正式名称は「大雪地区広域連合」である。3か所。"),
]

# 構造を伴う置換（XMLの断片を直接置き換える）
FIX_XML = [
    ("表記", "3. 議題別整理",
     '<text:span text:style-name="T128">3. 議題別整理</text:span>'
     '<text:span text:style-name="T129">。</text:span>',
     '<text:span text:style-name="T128">3. 議題別整理</text:span>',
     "見出しの末尾に句点が付いている。他の見出しには付いていない。"),
    ("脱字", "4. 令和8年10月　主な成果・判断",
     '<text:span text:style-name="T390">案、骨子・KPI目標、町別論点</text:span>',
     '<text:span text:style-name="T390">案）、骨子・KPI目標、町別論点</text:span>',
     "「推計（案」の括弧が閉じていない。"),
    ("誤記", "3-4. 最終報告・答申　想定内容",
     '<text:p text:style-name="P306">パブコメ反映、/計画原案・保険料方向性'
     '</text:p><text:p text:style-name="P307">第10計画（最終案）作成</text:p>',
     '<text:p text:style-name="P306">パブコメ反映、計画原案・保険料の方向性'
     '</text:p><text:p text:style-name="P307">第10期計画（最終案）作成</text:p>',
     "不要な「/」が混入している。また「第10計画」は「第10期計画」の脱字である。"),
    ("誤記", "2. 既存調査",
     '<text:span text:style-name="T59">令和7年度の介護予防・'
     '日常生活圏域ニーズ調査及び</text:span>'
     '<text:span text:style-name="T60">、実施済みの</text:span>',
     '<text:span text:style-name="T59">令和7年度に実施した</text:span>'
     '<text:span text:style-name="T60"></text:span>',
     "「介護予防・日常生活圏域ニーズ調査」と「健康とくらしの調査"
     "（JAGES調査）」は同一の調査である。"
     "「及び」で並べると別の調査を2件実施したと読めるため、"
     "④に統合する。"),
    ("誤記", "2. 既存調査",
     '<text:span text:style-name="T63">生活改善調査、②居所変更実態調査、'
     '③介護人材実態調査、④健康とくらしの調査（JAGES調査）</text:span>',
     '<text:span text:style-name="T63">生活改善調査、②居所変更実態調査、'
     '③介護人材実態調査及び④介護予防・日常生活圏域ニーズ調査'
     '（健康とくらしの調査、JAGES調査）</text:span>',
     "上に同じ。④に調査名を統合する。"),
    ("表記", "3-4. 最終報告・答申　時期・状態",
     '<text:span text:style-name="T314">19日事務管理者・主監</text:span>',
     '<text:span text:style-name="T314">19日、事務管理者・主監</text:span>',
     "日付と会議名が続けて書かれており、区切りがない。"),
    ("表記", "3-4. 成案完成　時期・状態",
     '<text:p text:style-name="P324">令和9年3月2日全員協議会情報提供</text:p>'
     '<text:p text:style-name="P325">令和9年3月3日第10期介護保険事業計画決定'
     '（条例改正提出準備）</text:p>'
     '<text:p text:style-name="P326">令和9年3月23日広域連合会議</text:p>',
     '<text:p text:style-name="P324">令和9年3月2日　全員協議会情報提供</text:p>'
     '<text:p text:style-name="P325">令和9年3月3日　第10期介護保険事業計画決定'
     '（条例改正提出準備）</text:p>'
     '<text:p text:style-name="P326">令和9年3月23日　広域連合会議</text:p>',
     "日付と会議名が続けて書かれており、区切りがない。3行とも。"),
]

# ------------------------------------------------------------ 反映しない指摘
PROPOSE = [
    ("提案", "5. ③ KPI・役割分担",
     "本文KPI12＋補完4",
     "「代表KPI16項目（うち補完4項目）」",
     "内容は同じである（12＋4＝16）。"
     "ただし「本文KPI12」と「補完4」が別枠と読める余地があるため、"
     "計画骨子案及びKPI定義書の表記に揃えることを提案する。"
     "原本の記載を変えることになるため反映していない。"),
    ("提案", "2. 計画の位置付け",
     "保険事業運営",
     "「介護保険事業の運営」",
     "「保険事業運営」は生命保険等で用いる語であり、"
     "介護保険では通常「介護保険事業の運営」又は「保険者としての運営」"
     "と表す。"),
    ("提案", "2. 計画の位置付け",
     "福祉の中でも上位に位置する計画",
     "「市町村介護保険事業計画として、構成3町の老人福祉計画と"
     "一体的に作成する計画」",
     "介護保険事業計画は、老人福祉計画と一体のものとして作成し、"
     "市町村地域福祉計画等と調和が保たれたものでなければならない"
     "（介護保険法第117条）。"
     "「福祉の中でも上位」とすると、3町の地域福祉計画との関係を"
     "説明できなくなる。"),
    ("提案", "3-3. サービス見込量・保険料",
     "「〜します」（敬体）",
     "「〜する」（常体）",
     "3-1・3-2は常体、3-3のみ敬体である。議事録の文体を統一する。"),
    ("提案", "3-4. 推進協議会・パブリックコメント",
     "「第2回協議会」",
     "「第1回」の記載の追加、又は「委員会」との呼称の統一",
     "第2回の前に第1回の記載がない。"
     "また4. の表では「第1回委員会資料」とあり、"
     "「委員会」と「協議会」の呼称が混在している。"),
    ("提案", "3-4. 10月構成町意見照会",
     "段階名「10月構成町意見照会」／時期「10〜11月頃」",
     "段階名と時期を一致させる",
     "段階名が10月、時期が10〜11月頃で食い違っている。"),
    ("提案", "3-4. パブリックコメント",
     "想定内容欄と時期・状態欄",
     "重複の解消",
     "「令和9年1月に概要版を用いWEBにて2週間実施」が"
     "2つの欄に重複して記載されている。"),
    ("提案", "3-4. 最終報告・答申",
     "「主監会議」",
     "正式名称の確認",
     "「主監会議」の正式名称を確認のうえ表記する。"),
    ("提案", "2. 重点論点",
     "「３町の特異性を平準化させる」",
     "「3町の差を把握したうえで、広域として共通の施策体系に整理する」",
     "「平準化」は3町の地域差を均すとも、様式を揃えるとも読める。"
     "当方は地域差の分析（第10期計画_地域差の分析.xlsx）を実施済みであり、"
     "地域差は把握して施策に反映する対象である。"
     "本文の記述と齟齬が生じないよう、意味を特定する。"),
]

KYOGI = [
    ("要協議", "3-1. 現状分析・将来推計",
     "「人口推計は地方創生計画による人口推計の値を使用」",
     "3町一律には適用できない。"
     "第3期総合戦略の総人口目標は、東神楽町が「令和11年度までに9,500人維持」、"
     "東川町が8,635人、美瑛町は総人口目標を設定していない。"
     "令和7年国勢調査（速報）は東神楽町9,588人・東川町8,726人・美瑛町9,337人で、"
     "東川町は目標を91人上回っている一方、"
     "東神楽町は年▲1.09％のペースで令和8年中に9,500人を割る見込みである。"
     "美瑛町は使用する目標値が存在しない。"
     "議事録の記載のままでは、町ごとに扱いが異なることが読み取れない。"),
    ("要協議", "3-1. 現状分析・将来推計",
     "「社人研の自然推計を元にした見える化システムとの乖離を確認する」",
     "方向は妥当であるが、総人口の乖離だけを見ると判断を誤る。"
     "社人研推計（見える化A1）の令和7年値と国勢調査（速報）の実績の差は、"
     "東川町＋513人・美瑛町＋444人・東神楽町▲408人で町ごとに向きが異なる。"
     "一方、住民基本台帳（令和8年）との突合では、"
     "総人口は住基が622人多いのに、65歳以上は144人、75歳以上は256人少ない。"
     "給付費に効くのは65歳以上・75歳以上であり、"
     "突合はこの2区分で行う必要がある"
     "（第10期計画_世帯構成の突合.xlsx 06シート）。"),
    ("要協議", "3-3. サービス見込量・保険料　手順2",
     "「令和8年実績で再基準化」（8〜9月）",
     "令和8年度は第9期の最終年度であり、実績が確定するのは令和9年3月末である。"
     "8〜9月の時点で用いられるのは令和7年度実績と令和8年度の月報までである。"
     "「令和8年実績」が令和8年度の実績を指すのか、"
     "令和8年（暦年）時点で得られる直近実績を指すのかを特定する。"),
    ("要協議", "2. 保険料",
     "「第10期保険料は現行水準の据え置きを目標とする」",
     "現行（第9期）の基準額は月額6,400円である。"
     "見える化システムの自然体推計による第10期の3年平均は6,238円"
     "（基金取崩0円・予定収納率99.0％）で、現行を162円下回る。"
     "据え置きは現時点で無理のない目標であるが、"
     "第3段階（給付費・保険料）の推計は単価・基金残高・収納率・"
     "所得段階別被保険者数の受領後に確定する。"
     "また「米価の上昇により農業収入が高かったため基金の取崩しがなかった」"
     "という因果は確かめられていない。"
     "給付費が想定を下回った可能性もあるため、議事録では"
     "「基金の取崩しは想定を下回った」という事実の記載にとどめることを提案する。"),
    ("要協議", "5. ③ KPI・役割分担（☑確認）",
     "代表KPI4項目の代理指標が未決である",
     "受領済みの調査には、H12（必要サービス未充足率）の受入困難の設問と、"
     "H16（災害・感染症時のサービス継続率）のBCPの設問が含まれていない。"
     "H07（主介護者の高負担割合）・H08（介護離職懸念）は"
     "在宅介護実態調査によるもので、同調査は実施していない。"
     "この4項目は現在のデータでは算定できない。"
     "「確認」で閉じると、基準値のないKPIが計画に残る。"
     "代理指標への振替え、項目の削除、又は2問の追加照会のいずれかを"
     "決定する必要がある。"),
    ("要協議", "5. ① 調査データの受渡し（☑確認）",
     "事業所実態調査は受領済みである",
     "令和8年8月6日に「現時点で提出いただいている調査結果のすべて」として"
     "受領済みであり、受渡し方法を確認する段階ではない。"
     "受領した回答は事業所票27件・職員個票317人・訪問系職員票26件で、"
     "訪問系は13事業所のうち3事業所からの回答である"
     "（残る10事業所は介護サービス情報公表システムの個別公表画面により"
     "把握済み）。"
     "記載を実態に合わせることを提案する。"),
]


def main():
    with zipfile.ZipFile(SRC) as z:
        names = z.namelist()
        data = {n: z.read(n) for n in names}

    x = data["content.xml"].decode("utf-8")
    done = {}

    for _k, _w, old, new, _r in FIX:
        n = x.count(old)
        if n == 0:
            raise RuntimeError("該当なし: %r" % old)
        x = x.replace(old, new)
        done[old] = n

    for _k, _w, old, new, _r in FIX_XML:
        n = x.count(old)
        if n == 0:
            raise RuntimeError("該当なし（XML）: %r" % old[:60])
        x = x.replace(old, new)
        done[old[:40]] = n

    data["content.xml"] = x.encode("utf-8")

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    if os.path.exists(OUT):
        os.remove(OUT)
    # mimetype は無圧縮で先頭に置く（ODF の要件）
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(zipfile.ZipInfo("mimetype"), data["mimetype"],
                   compress_type=zipfile.ZIP_STORED)
        for n in names:
            if n == "mimetype":
                continue
            z.writestr(n, data[n])

    print("saved:", OUT)
    for k, v in done.items():
        print("  OK %d件  %s" % (v, k.replace("\n", "/")[:60]))

    # ------------------------------------------------------ 校正結果の一覧
    FONT = "游ゴシック"
    NAVY, HEAD = "1F3864", "4472C4"
    OK_G, IN_Y, NG_O = "E2EFDA", "FFF2CC", "FCE4D6"
    thin = Side(style="thin", color="BFBFBF")
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
    wb = Workbook()
    ws = wb.active
    ws.title = "校正結果"
    ws.sheet_view.showGridLines = False
    for i, w in enumerate([6, 9, 26, 26, 26, 56], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    c = ws.cell(row=1, column=1,
                value="キックオフ会議議事録（令和8年8月6日開催）の校正結果")
    c.font = Font(name=FONT, size=14, bold=True, color=NAVY)
    c = ws.cell(row=2, column=1,
                value="誤字脱字及び明らかな表記の誤りは"
                      "「第10期計画_キックオフ会議議事録_校正反映.odt」に"
                      "反映済みである。"
                      "「提案」及び「要協議」は反映していない。")
    c.font = Font(name=FONT, size=9)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=6)
    ws.row_dimensions[2].height = 30
    ws.freeze_panes = "A5"

    r = 4
    for i, v in enumerate(["No.", "区分", "箇所", "現在の記載",
                           "修正後／提案", "理由"], start=1):
        c = ws.cell(row=r, column=i, value=v)
        c.font = Font(name=FONT, size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=HEAD)
        c.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[r].height = 26
    r += 1

    FILL = {"誤字": OK_G, "脱字": OK_G, "誤記": OK_G, "表記": OK_G,
            "提案": IN_Y, "要協議": NG_O}
    no = 0
    rows = ([(k, w, o, n, why) for k, w, o, n, why in FIX]
            + [(k, w, ("XMLの断片（%s）" % w), "反映済み", why)
               for k, w, _o, _n, why in FIX_XML]
            + [(k, w, o, n, why) for k, w, o, n, why in PROPOSE]
            + [(k, w, o, "―", why) for k, w, o, why in KYOGI])
    for k, w, o, n, why in rows:
        no += 1
        for i, v in enumerate([no, k, w, o, n, why], start=1):
            c = ws.cell(row=r, column=i, value=v)
            c.font = Font(name=FONT, size=9)
            c.alignment = Alignment(wrap_text=True, vertical="top",
                                    horizontal="center" if i in (1, 2)
                                    else "left")
            c.border = BORDER
            if i == 2:
                c.fill = PatternFill("solid", fgColor=FILL[k])
        ws.row_dimensions[r].height = max(30, 13 * (len(why) // 34 + 1))
        r += 1

    wb.save(OUTX)
    print("saved:", OUTX)
    print("  校正 %d件（反映 %d件・提案 %d件・要協議 %d件）"
          % (no, len(FIX) + len(FIX_XML), len(PROPOSE), len(KYOGI)))


if __name__ == "__main__":
    main()
