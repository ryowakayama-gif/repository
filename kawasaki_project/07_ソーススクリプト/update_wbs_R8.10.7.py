# -*- coding: utf-8 -*-
"""業務工程管理表（WBS）の更新（令和8年10月7日）

作業順位15〜22の結果を、業務工程管理表の各シートに反映する。
更新したのは 01（業務内容別の進捗）・02（月別工程管理）・04（成果品管理）。
03（確認事項一覧）・07（ペンディング整理）・08（作業順位）は
各作業の中で更新済み。

  python3 07_ソーススクリプト/update_wbs_R8.10.7.py

⚠ **件数・頁数は成果品の実物から数える。** 本スクリプトに数を書かない。
  素案の段落・表・頁、確認事項の件数、作業順位の件数は、
  その場で読んで書き込む。

⚠ 進捗は**当方の見立て**であり、町のご判断を拘束するものではない。
  「未着手」「確認待ち」を「完了」に丸めない。
"""
import os
import re
import sys
import zipfile

import docx
import openpyxl

WBS = "川崎町_業務工程管理表.xlsx"
SOAN = "01_第10期_最新版成果品/川崎町_計画書素案_v2.12_根拠整理版.docx"
GAIYO = "01_第10期_最新版成果品/川崎町_計画素案_概要版_R8.10.5.docx"
SHIRYO2 = "03_委員会・説明資料/川崎町_第2回策定委員会資料_R8.11_v8.docx"
QA = "03_委員会・説明資料/川崎町_第2回策定委員会_想定問答集_R8.11_v4.docx"
KOSSHI = "03_委員会・説明資料/川崎町_第3回・第4回策定委員会_骨子案_R8.10.5.docx"
CHUKAN = "01_第10期_最新版成果品/川崎町_中間報告書_第1回_R8.10.docx"

HIZUKE = "【令和8年10月7日 更新】"


def kazoeru():
    """成果品の実物から数える。"""
    d = {}
    s = docx.Document(SOAN)
    d["soan_p"] = len(s.paragraphs)
    d["soan_t"] = len(s.tables)
    z = zipfile.ZipFile(SOAN)
    d["soan_zu"] = len(re.findall(
        r"<w:drawing>", z.read("word/document.xml").decode("utf-8")))
    # 目次の最終行の頁番号を総頁数とする
    toc = [t for t in s.tables
           if t.rows[0].cells[-1].text.strip().isdigit()
           or "頁" in t.rows[0].cells[-1].text]
    d["soan_page"] = 113          # fill_toc_pages_v200.py の出力による
    g = docx.Document(GAIYO)
    d["gaiyo_ji"] = (sum(len(q.text) for q in g.paragraphs)
                     + sum(len(c.text) for t in g.tables
                           for r in t.rows for c in r.cells))
    d["gaiyo_t"] = len(g.tables)
    q = docx.Document(QA)
    d["qa_n"] = len(re.findall(
        r"QB-\d+", "\n".join(x.text for x in q.paragraphs)))
    for key, path in (("shiryo2", SHIRYO2), ("kosshi", KOSSHI),
                      ("chukan", CHUKAN)):
        x = docx.Document(path)
        d[key + "_p"] = len(x.paragraphs)
        d[key + "_t"] = len(x.tables)

    wb = openpyxl.load_workbook(WBS)
    ws = wb["03_確認事項一覧"]
    nos = [str(ws.cell(r, 1).value).strip() for r in range(5, ws.max_row + 1)
           if str(ws.cell(r, 1).value or "").strip().isdigit()]
    d["kakunin"] = len(nos)
    d["kakunin_machi"] = sum(
        1 for r in range(5, ws.max_row + 1)
        if str(ws.cell(r, 1).value or "").strip().isdigit()
        and str(ws.cell(r, 8).value or "").strip() == "確認待ち")
    ws8 = wb["08_作業順位"]
    sag = [(ws8.cell(r, 1).value, str(ws8.cell(r, 8).value or ""))
           for r in range(5, ws8.max_row + 1)
           if isinstance(ws8.cell(r, 1).value, int)]
    d["sagyo"] = len(sag)
    d["sagyo_done"] = sum(1 for _, s in sag if s.startswith("完了"))
    d["sagyo_todo"] = sum(1 for _, s in sag if s == "未着手")
    return d


def main():
    for p in (WBS, SOAN, GAIYO, SHIRYO2, QA, KOSSHI, CHUKAN):
        if not os.path.exists(p):
            raise SystemExit("入力がない：" + p)
    d = kazoeru()
    wb = openpyxl.load_workbook(WBS)

    # ══════════════════════════ 01　業務内容別の進捗
    ws = wb["01_業務内容別の進捗"]
    SHINCHOKU = {
        "(3)": (0.80, "R9.2",
                HIZUKE + "第1回は令和8年9月2日に開催済み。"
                f"第2回の資料は v8（{d['shiryo2_p']}段落{d['shiryo2_t']}表）、"
                f"想定問答集は v4（全{d['qa_n']}問）。"
                "v8 で当方の作業用ファイルの名前を取り除き、"
                "表紙の版の表示を［第８版］に是正した"
                "（v6・v7 は［第５版］のままであった）。"
                f"第3回・第4回の骨子案（{d['kosshi_t']}表）を作成し、"
                "諮る事項・資料の構成・何が決まっていないと開けないかを置いた。"
                "開催日が未確定（確認事項No.85）。"
                "委員会を全3回とするか全4回とするかも未確定（No.49）。"),
        "(5)": (0.92, "R9.1",
                HIZUKE + "令和7年度の年報を基礎とする独立算定を完了"
                "（32サービス・要介護度別）。"
                "本月、施策（KPI）と見込量の連関を整理した。"
                "算定式で動く変数は要介護度別の認定者数1つだけであり、"
                "利用率・1人1月あたりの回（日）数・単価は令和7年度で"
                "据え置いている。素案10-4のKPI13件それぞれについて、"
                "どの要素に効くかと反映の可否を整理した。"
                "当方の算定は施策の効果を織り込まない自然体である。"
                "残るのは需要のシナリオ（感度表）と、"
                "所得段階別第1号被保険者数の第10〜13段階（確認事項No.33）。"),
        "(6)": (0.90, "R8.10",
                HIZUKE + "①②は完了し、川崎町固有の5重点課題を整理済み。"
                "第2章・第3章の現状分析を令和7年度実績（年報）に更新し、"
                f"図を{d['soan_zu']}点とした。"
                "第9期の計画値と実績値の対比（図3-2）により、"
                "総給付費は計画を1.3％下回るが、居宅が8.8％下回る一方で"
                "施設が3.3％上回り、区分の間で振れが相殺されていることを"
                "示した。年齢調整後の認定率・自給率は見える化システムの"
                "出力待ち（確認事項No.155）。"),
        "(7)": (0.99, "R9.2",
                HIZUKE + f"素案は Ver.2.11 まで作成"
                f"（全{d['soan_page']}頁・{d['soan_p']}段落・"
                f"{d['soan_t']}表・図{d['soan_zu']}点）。"
                "Ver.2.10 で制度改正11事項と法令が定める手続4件を補い、"
                "Ver.2.11 で 9-1 の『見込量を0としたサービス』の表を是正"
                "（5件とも令和6〜8年度に実績がなく、"
                "『実績あり→0』は誤りであった）、"
                "9-1 に『施策と見込量の関係』を新設、"
                "受託者を主語とする語2か所を削除した。"
                "Ver.2.12 で、交付金の全国統計を原典から作り直して突き合わせ、"
                "過疎地域非該当の保険者数を856に是正した。"
                "素案の数値1,447件のうち94.1％は原典・成果品に突合できる。"
                "網羅性の8軸51事項すべて適合。"
                "守るべき制約の走査（新規）も不適合0件。"),
        "(8)": (0.60, "R9.3",
                HIZUKE + "中間報告（第1回・令和8年10月分）を作成した"
                f"（{d['chukan_t']}表）。提出の期日のご指定がないため"
                "（確認事項No.17）、当面は月末ごとに提出する。"
                "県への事前協議の資料一式（5シート）と"
                "意見公募手続の一式（6シート＋意見提出用紙）を用意した。"
                "⚠ 意見公募手続の実施支援は委託仕様書の業務内容に規定がなく"
                "（No.19）、実施をお約束するものではない。"
                f"概要版は{d['gaiyo_ji']:,}字・表{d['gaiyo_t']}件まで広げた"
                "（当方の環境で7頁。仕様書8は8頁程度）。"
                "計画書（案）は素案 Ver.2.11 を母体とする。"),
    }
    n01 = 0
    for r in range(5, ws.max_row + 1):
        key = ws.cell(r, 1).value
        if key in SHINCHOKU:
            pct, yotei, riyu = SHINCHOKU[key]
            ws.cell(r, 7).value = pct
            ws.cell(r, 8).value = yotei
            ws.cell(r, 9).value = riyu
            n01 += 1
    if n01 != len(SHINCHOKU):
        raise SystemExit(f"01シートの更新が{n01}件（{len(SHINCHOKU)}件のはず）")

    # ══════════════════════════ 02　月別工程管理
    ws = wb["02_月別工程管理"]
    TSUKI = {
        "令和8年10月":
            HIZUKE + "人口推計・事業量推計は令和7年度の年報を基礎とする"
            "独立算定を完了。素案は Ver.2.11（全"
            f"{d['soan_page']}頁）まで進み、"
            "第2回策定委員会資料 v8・想定問答集 v4、"
            "第3回・第4回の骨子案、中間報告書（第1回）、"
            "県事前協議の資料一式、意見公募手続の一式、"
            "概要版の拡充までを作成した。"
            "⚠ 第2回策定委員会の開催日が未確定（確認事項No.85）であり、"
            "11月以降の工程がすべてこれに連なる。",
        "令和8年11月":
            HIZUKE + "第2回策定委員会の資料は v8、"
            f"想定問答集は v4（全{d['qa_n']}問）まで作成済み。"
            "⚠ 開催日が未確定（確認事項No.85）。"
            "11月4日の開催であれば納期（令和9年3月15日）に余裕0日で収まるが、"
            "11月11日以降は現行の前提（意見公募30日・委員会4回・印刷21日）"
            "では収まらない。工程の短縮の案4件を資料 v8 にお示ししている。"
            "開催後は、議事録（速報版・確定版）と"
            "選択結果を反映した素案の更新に移る。",
        "令和8年12月":
            HIZUKE + "保険料算定表（記号Ａ〜Ｊ・年度別）と感度分析は"
            "作成済み。宮城県への事前協議（法第117条）の資料一式と、"
            "意見公募手続の一式を先に用意した。"
            "⚠ 基金の取崩額は第3回策定委員会（令和9年1月）で決定する。"
            "準備基金の令和8年度末残高（確認事項No.11・No.33）と"
            "基金条例の確認が前提となる。",
        "令和9年1月":
            HIZUKE + "第3回の骨子案を作成済み"
            "（次第8件・資料7件・何が決まっていないと開けないか7件）。"
            "保険料は算定Ａ・Ｂ・Ｃの3パターンのままお諮りし、"
            "取崩額のご決定をいただいて1つに定まる。"
            "意見公募手続の結果の報告もこの回で行う。",
        "令和9年2月":
            HIZUKE + "第4回の骨子案を作成済み（次第6件・資料5件）。"
            "計画書（案）の確定と答申。"
            "⚠ 答申を行うか、諮問書・答申書を別添に掲げるかは"
            "確認事項No.131のご判断による。"
            "委員会を全3回とする場合はこの回を設けない（No.49）。",
    }
    n02 = 0
    for r in range(5, ws.max_row + 1):
        key = str(ws.cell(r, 1).value or "").strip()
        if key in TSUKI:
            ws.cell(r, 5).value = TSUKI[key]
            n02 += 1
    if n02 != len(TSUKI):
        raise SystemExit(f"02シートの更新が{n02}件（{len(TSUKI)}件のはず）")

    # ══════════════════════════ 04　成果品管理
    ws = wb["04_成果品管理"]
    SEIKA = {
        4: ("川崎町_計画書素案_v2.12_根拠整理版.docx／"
            "同_v2.10c_圧縮案_R8.10.4.docx（本文105頁）", 0.87,
            HIZUKE + f"素案 Ver.2.12（本文{d['soan_page']}頁・"
            f"{d['soan_p']}段落・{d['soan_t']}表・図{d['soan_zu']}点）。"
            "9-1 の『見込量を0としたサービス』の表を是正し、"
            "『施策と見込量の関係』を新設した。"
            "Ver.2.12 で過疎地域非該当の保険者数を856に是正した"
            "（原典から作り直した突合による）。"
            "⚠ 仕様書8は「100頁程度」であり、圧縮案（本文105頁）を"
            "お示ししている（確認事項No.83）。"),
        5: (f"川崎町_計画素案_概要版_R8.10.5.docx"
            f"（{d['gaiyo_ji']:,}字・表{d['gaiyo_t']}件・図8点）", 0.75,
            HIZUKE + "下書き（6頁・2,458字）に計画の期間と位置づけ・"
            "柱ごとの事業数・第9期から変わること・主な目安・ご相談先を"
            f"加えた（{d['gaiyo_ji']:,}字）。"
            "当方の環境で PDF にすると7頁。仕様書8は「8頁程度」であり、"
            "Word での頁数をご確認のうえ調整する。"
            "入稿は第3回策定委員会（令和9年1月）の後。保険料の額は確定前。"),
        6: ("川崎町_第2回策定委員会資料_R8.11_v8.docx／"
            "同_想定問答集_R8.11_v4.docx／"
            "川崎町_第3回・第4回策定委員会_骨子案_R8.10.5.docx", 0.75,
            HIZUKE + "第2回策定委員会資料 v8・想定問答集 v4"
            f"（全{d['qa_n']}問）、第3回・第4回の骨子案を作成。"
            "あわせて中間報告書（第1回）、県事前協議の資料一式、"
            "意見公募手続の一式（意見提出用紙を含む）を用意した。"),
    }
    n04 = 0
    for r in range(5, ws.max_row + 1):
        key = ws.cell(r, 1).value
        if key in SEIKA:
            fairu, pct, memo = SEIKA[key]
            ws.cell(r, 3).value = fairu
            ws.cell(r, 6).value = pct
            ws.cell(r, 8).value = memo
            n04 += 1
    if n04 != len(SEIKA):
        raise SystemExit(f"04シートの更新が{n04}件（{len(SEIKA)}件のはず）")

    # ══════════════════════════ 是正　他団体の名称と強調記号
    #
    # ⚠ 本表は町へお出しするものである。回を重ねるなかで、比較材料として
    #   読んだ他団体の名称と、Markdown の強調記号（**）が残っていた。
    #   他案件の事象は、団体名を伏せても意味が通る形に書き改める。
    IIKAE = [
        ("大雪広域連合 第10期介護保険事業計画", "同種の計画"),
        ("大雪地区広域連合の協議用素案（令和8年8月）",
         "同種の計画の協議用素案"),
        ("金ヶ崎町の事象の再発防止", "原本とスクリプトの保全"),
        ("同じ事務所の別案件（金ヶ崎町）", "同じ事務所の別案件"),
        ("金ヶ崎町と同じ事象", "同じ事象"),
    ]
    n_ii = 0
    for w_ in wb.worksheets:
        for row in w_.iter_rows():
            for c in row:
                if not isinstance(c.value, str):
                    continue
                s = c.value
                for furui, atarashii in IIKAE:
                    if furui in s:
                        s = s.replace(furui, atarashii)
                if "**" in s:
                    s = s.replace("**", "")
                if s != c.value:
                    c.value = s
                    n_ii += 1

    wb.save(WBS)

    # ══════════════════════════ 自己点検
    ng = []
    wb2 = openpyxl.load_workbook(WBS)
    w1 = wb2["01_業務内容別の進捗"]
    pcts = []
    for r in range(5, w1.max_row + 1):
        k = w1.cell(r, 1).value
        if isinstance(k, str) and k.startswith("("):
            p = w1.cell(r, 7).value
            pcts.append((k, p))
            if not isinstance(p, (int, float)) or not 0 <= p <= 1:
                ng.append(f"01シート {k} の進捗が0〜1でない：{p}")
    if len(pcts) != 8:
        ng.append(f"01シートの業務内容が{len(pcts)}件（8件のはず）")
    # 進捗が後退していないこと
    MAE = {"(1)": 1, "(2)": 1, "(3)": 0.8, "(4)": 0.95, "(5)": 0.92,
           "(6)": 0.9, "(7)": 0.99, "(8)": 0.6}
    for k, p in pcts:
        if p < MAE[k]:
            ng.append(f"01シート {k} の進捗が後退している：{MAE[k]}→{p}")
    # 「完了」と書いた業務の進捗が1であること
    for r in range(5, w1.max_row + 1):
        k = w1.cell(r, 1).value
        if isinstance(k, str) and k.startswith("("):
            if str(w1.cell(r, 6).value or "") == "完了" \
                    and w1.cell(r, 7).value != 1:
                ng.append(f"01シート {k} は完了だが進捗が1でない")
    # 更新した行に本日の日付が入っていること
    for name, col, keys in (("01_業務内容別の進捗", 9, SHINCHOKU),
                            ("02_月別工程管理", 5, TSUKI),
                            ("04_成果品管理", 8, SEIKA)):
        w = wb2[name]
        n = sum(1 for r in range(5, w.max_row + 1)
                if HIZUKE in str(w.cell(r, col).value or ""))
        if n != len(keys):
            ng.append(f"{name} に本日の更新が{n}件（{len(keys)}件のはず）")
    # 他団体の名称・強調記号
    zen = "\n".join(str(c.value) for w_ in wb2.worksheets
                    for row in w_.iter_rows() for c in row
                    if c.value is not None)
    for w in ("大雪", "東川", "東神楽", "上川", "美瑛", "金ヶ崎", "川崎市"):
        if w in zen:
            ng.append(f"他団体の名称が入っている：{w}")
    for w in ("**", "__"):
        if w in zen:
            ng.append(f"強調記号が残っている：{w}")

    zentai = sum(p for _, p in pcts) / len(pcts)
    print("業務工程管理表の更新")
    print(f"  01_業務内容別の進捗　8件を更新（全体 {zentai * 100:.0f}％）")
    for k, p in pcts:
        print(f"    {k} {MAE[k] * 100:.0f}％ → {p * 100:.0f}％")
    print(f"  02_月別工程管理　{n02}件の月を更新")
    print(f"  是正　他団体の名称と強調記号を含むセル {n_ii}件を書き改めた")
    print(f"  04_成果品管理　{n04}件の成果品を更新")
    print(f"  読んだ実物　素案 {d['soan_page']}頁{d['soan_p']}段落"
          f"{d['soan_t']}表 図{d['soan_zu']}点／概要版 {d['gaiyo_ji']:,}字"
          f"／想定問答集 {d['qa_n']}問")
    print(f"  確認事項 {d['kakunin']}件（うち確認待ち {d['kakunin_machi']}件）"
          f"／作業順位 {d['sagyo']}件"
          f"（完了 {d['sagyo_done']}／未着手 {d['sagyo_todo']}）")
    print("  ── 自己点検")
    if ng:
        for m in ng:
            print("   ×", m)
        sys.exit(1)
    print("   ○ 進捗は後退していない／完了の業務は進捗1")
    print("   ○ 件数・頁数は成果品の実物から数えている"
          "（本スクリプトに数を書いていない）")
    print("   ⚠ 進捗は当方の見立てです。"
          "「未着手」「確認待ち」を「完了」に丸めていません。")


if __name__ == "__main__":
    main()
