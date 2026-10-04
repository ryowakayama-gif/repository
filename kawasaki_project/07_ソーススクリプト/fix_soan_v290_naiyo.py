# -*- coding: utf-8 -*-
"""計画素案 Ver.2.8 → Ver.2.9（事業一覧表に「内容」の欄を加える）

08_作業順位 の順位6。素案の第5章・第6章の事業一覧表は
「項／事業／区分／所管」の4欄であり、**事業が何をするものかが書いていない。**
名称だけの一覧では計画として読めないため、「内容」の欄を加える。

内容の出所は2つ。**作文はしない。**

  ① 第9期計画書（令和6年3月）の本文　継続事業の多くはここに記述がある。
     事業名の後ろの最初の一文を機械で取り出す（57件）。
  ② 当方が書き起こすもの　①で取れなかった事業のうち、
     介護保険法・国の指針が内容を定めているもの（施設サービス等）と、
     素案の本文に記述があるもの（町独自の事業）に限る（44件）。
     **根拠のないものは［要確認］とし、推測で書かない。**

  python3 07_ソーススクリプト/fix_soan_v290_naiyo.py

⚠ 欄を1つ増やすため、表が縦に伸びて頁数が増える。
  圧縮案（v2.8c）と併せてご判断いただく。
"""
import copy
import os
import re
import shutil
import sys

import docx
import openpyxl
import pypdf
from docx.oxml.ns import qn
from docx.shared import Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

SRC = "01_第10期_最新版成果品/川崎町_計画書素案_v2.8_第9期対比版.docx"
DST = "01_第10期_最新版成果品/川崎町_計画書素案_v2.9_事業内容版.docx"
TAIKEI = "05_試算・管理シート/川崎町_第10期_施策事業統合体系表_R8.9.25b.xlsx"
PDF9 = "09_元資料/川崎町_第9期計画_04324.pdf"
LEDGER = "05_試算・管理シート/川崎町_継続101事業_本文記述_作業台帳_R8.10.2.xlsx"

VER = {4: "計画書素案 Ver.2.9", 5: "（事業内容版）",
       21: ("⚠ 頁番号は、本素案Ver.2.9の組版によるものです。"
            "図表の差替え・加筆により頁の割付けが動いた場合は、"
            "目次の頁番号も併せて更新します。")}
VER_SWAP = [("本素案Ver.2.8", "本素案Ver.2.9"),
            ("計画書素案 Ver.2.8", "計画書素案 Ver.2.9")]

# ══════════════════════════ ② 当方が書き起こすもの
#   介護保険法・国の指針が内容を定めているもの、
#   又は素案の本文に記述があるものに限る。
NAIYO = {
    # ── 第1章 健康づくりの推進
    "（３）住民との協働による健康づくりの推進":
        "健康推進員・食生活改善推進員などの住民組織と協働し、"
        "地区ごとの健康づくりの活動を進めます。",
    "（１）高齢者の感染症予防と発生時の対応":
        "感染症や予防接種に関する正しい知識の普及と情報提供を行い、"
        "医療機関等と連携して予防接種の円滑な実施に取り組みます。",
    "（１）特定健康診査":
        "40歳から74歳の方に特定健康診査を実施し、"
        "生活習慣の改善が必要な方に保健指導を行います。"
        "75歳以上の方には後期高齢者健康診査を実施します。",
    # ── 第3章 地域生活を支援する取り組みの充実
    "（１）外出支援サービス（移送サービス）事業【重点】":
        "社会福祉協議会・NPOによる福祉移送サービスにより、"
        "通院等の外出を支援します。"
        "令和7年3月の高齢者外出タクシー助成の終了に伴う"
        "移動支援3層構造の中核となる事業です。",
    "（２）デマンド型乗合タクシーの運行【重点】":
        "予約に応じて運行する乗合タクシーにより、"
        "路線によらない移動の手段を確保します。",
    "（３）町民バスの運行【重点】":
        "バス会社の運行による町民バスにより、"
        "地区と町の中心部を結ぶ移動の手段を確保します。",
    "（３）ふれあいネットワーク活動の推進":
        "活動員15名・協力員130名（計145名）による見守りの網により、"
        "独居高齢者・高齢者世帯の安否の確認と声かけを行います。",
    "（３）高齢者の健康づくり・スポーツ活動の振興":
        "高齢者を対象としたストレッチなどの健康づくり教室、"
        "スポーツ教室、転倒予防のための運動教室を開催します。",
    "（６）シルバー人材センターの活用":
        "これまでの経験や能力を活かした臨時的・短期的な就業の機会を"
        "確保し、高齢者の社会参加と生きがいづくりを進めます。",
    # ── 第4章 地域支援事業の充実
    "（１）訪問型サービス（訪問介護相当サービス、訪問型サービスＢ・Ｃ）":
        "要支援者等の居宅を訪問し、身体介護・生活援助を行います。"
        "指定事業所による訪問介護相当サービスのほか、"
        "住民主体のＢ、短期集中予防のＣを実施します。",
    "（２）通所型サービス（通所介護相当サービス、通所型サービスＢ）":
        "要支援者等が通う場で、入浴・食事・機能訓練等を行います。"
        "指定事業所による通所介護相当サービスのほか、"
        "住民主体のＢを実施します。",
    "（５）高額介護予防サービス相当事業等":
        "総合事業のサービスの利用者負担が一定の額を超えた場合に、"
        "超えた分を支給します。",
    "（１）介護予防把握事業":
        "基本チェックリスト等により、介護予防の支援を要する方を把握し、"
        "必要な事業につなぎます。",
    "（２）介護予防普及啓発事業（元気まんてん介護予防教室、元気いきいきセミナー、"
    "パドル体操教室、ノルディックウォーキング教室等）":
        "元気まんてん介護予防教室・元気いきいきセミナー・パドル体操教室・"
        "ノルディックウォーキング教室等を開催し、"
        "介護予防の知識と実践の場を広げます。",
    # ── 認知症施策（認知症基本法の基本的施策に対応するもの）
    "（４）少なくとも5年ごとの検討":
        "認知症基本法第12条第6項により、計画に少なくとも5年ごとに"
        "検討を加え、必要があると認めるときは変更します。",
    "（７）移動のための交通手段の確保":
        "認知症の人が移動しやすいよう、移動支援3制度"
        "（福祉移送サービス・デマンド型乗合タクシー・町民バス）の"
        "利用の案内と配慮を進めます。",
    "（８）交通の安全の確保":
        "運転に不安を感じる方の相談に応じ、運転免許の自主返納を支援するとともに、"
        "交通安全の啓発を行います。",
    "（４）企業・学校等における認知症サポーターの養成":
        "町内の企業・学校等に出向いて認知症サポーター養成講座を開き、"
        "地域で支える裾野を広げます。",
    "（１）成年後見制度の利用促進":
        "成年後見制度の周知と相談に応じ、"
        "必要な方が制度を利用できるよう支援します。",
    "（２）日常生活自立支援事業の活用":
        "社会福祉協議会が行う日常生活自立支援事業により、"
        "金銭の管理や福祉サービスの利用の手続を支援します。",
    "（３）意思決定支援研修の実施":
        "認知症の人の意思が尊重されるよう、"
        "関係者を対象とした意思決定支援の研修を実施します（第10期に新設）。",
    "（４）消費者被害の防止の啓発":
        "消費生活相談の窓口と連携し、"
        "高齢者・認知症の人を狙った消費者被害の防止に向けた啓発を行います。",
    "（６）介護予防・地域ささえあいサポート拠点事業":
        "身近な場所に地域の支え合いの拠点を設け、"
        "介護予防と交流の場として運営を支援します。",
    "（８）介護予防ポイント事業":
        "介護予防の活動やボランティアへの参加にポイントを付与し、"
        "参加を促します。交付金 支援 目標Ⅰで県内1位の取組です。",
    "（９）アウトリーチによる参加勧奨":
        "通いの場等に参加していない方へ個別に働きかけ、参加につなげます。"
        "参加意向55.2％に対し参加率15.9％という開きを埋める取組です。",
    "（７）ユニバーサルサポーター制度の推進":
        "14種別約400名の町独自のサポーター制度により、"
        "住民主体の介護予防・生活支援・見守りの活動を進めます。",
    "（４）包括的・継続的ケアマネジメント支援事業":
        "介護支援専門員への助言・指導、地域ケア会議の開催等により、"
        "包括的・継続的なケアマネジメントを支援します。",
    "（４）生活支援・介護予防サービスの推進":
        "不足するサービスの創出、担い手づくり、活動の場の確保などの"
        "資源の開発と、関係者の間の連携により、"
        "生活支援・介護予防のサービスを広げます。",
    "（２）成年後見制度利用支援事業":
        "判断能力が十分でない方の権利を守るため、"
        "成年後見制度の利用に要する費用を助成し、町長による申立てを行います。",
    "（３）介護用品（紙おむつ等）の支給":
        "在宅で要介護高齢者を介護する家族に紙おむつ等の介護用品を支給し、"
        "介護の負担を軽くします。",
    "（４）配食サービス":
        "調理が困難な独居高齢者等に食事を届け、あわせて安否を確認します。",
    # ── 第5章 認知症施策の推進
    "（２）もの忘れ相談の充実":
        "もの忘れが気になる方と家族の相談に応じ、"
        "医療機関や地域包括支援センターにつなぎます（年12回）。",
    "（４）難聴高齢者の早期発見・早期介入":
        "聞こえの低下を早期に見つけ、補聴器の利用や受診につなぎます。"
        "交付金では2年連続で20点満点を得ている取組です。",
    "（３）認知症の早期発見・早期治療の体制づくりの推進":
        "もの忘れ相談・認知症初期集中支援チーム・国保川崎病院との連携により、"
        "早期の発見と対応の体制を整えます。",
    "（１）認知症地域支援推進員の配置":
        "医療・介護・地域の支援を結びつける認知症地域支援推進員を配置し、"
        "認知症の人と家族を支えます。",
    "（３）認知症カフェ（喫茶みかん）の運営支援":
        "認知症の人・家族・地域の方・専門職が集う場として、"
        "認知症カフェ「喫茶みかん」（年47回）の運営を支援します。",
    "（４）認知症の人と家族の相談支援":
        "地域包括支援センターを窓口に、認知症の人と家族の相談に応じ、"
        "必要な支援につなぎます。",
    "（５）認知症高齢者等見守りＱＲコード活用事業":
        "衣類等に貼るＱＲコードにより、"
        "行方が分からなくなった認知症の人の早期の発見を図ります。",
    "（２）認知症キャラバンメイトの養成・活動支援":
        "認知症サポーター養成講座の講師を務めるキャラバンメイトを養成し、"
        "活動を支援します（累計80名）。",
    "（５）認知症の人と家族を支える人材の育成":
        "認知症サポーター・キャラバンメイト・"
        "ステップアップ講座修了者の養成により、支える人材を育てます。",
    # ── 第6章 介護給付・介護予防給付サービス（内容は法令が定める）
    "（３）訪問看護・介護予防訪問看護":
        "看護師等が居宅を訪問し、療養上の世話や診療の補助を行います。",
    "（９）短期入所療養介護（ショートステイ）・介護予防短期入所療養介護":
        "介護老人保健施設等に短期間入所し、"
        "看護・医学的管理のもとで介護や機能訓練を受けます。",
    "（１）定期巡回・随時対応型訪問介護看護":
        "日中・夜間を通じて、定期の訪問と、"
        "呼び出しに応じた随時の訪問・対応を行います。"
        "区域内に事業所はありませんが給付の実績があります。",
    "（５）認知症対応型共同生活介護・介護予防認知症対応型共同生活介護"
    "（グループホーム）":
        "認知症の方が少人数で共同生活を営む住居で、"
        "入浴・排泄・食事等の介護や機能訓練を受けます。",
    "（７）地域密着型介護老人福祉施設入居者生活介護"
    "（小規模特別養護老人ホーム）":
        "定員29人以下の特別養護老人ホームに入所し、"
        "介護や機能訓練を受けます。",
    "（１）介護老人福祉施設":
        "常時介護が必要で居宅での生活が難しい方が入所し、"
        "介護や機能訓練、健康管理を受けます（特別養護老人ホーム）。",
    "（２）介護老人保健施設":
        "病状が安定し在宅復帰を目指す方が入所し、"
        "看護・医学的管理のもとで介護や機能訓練を受けます。",
    "（３）介護医療院":
        "長期の療養が必要な方が入所し、"
        "医療と日常生活上の世話を一体的に受けます。"
        "本町に利用の実績はありません。",
    # ── 第7章 地域マネジメント
    "（１）介護給付の適正化":
        "要介護認定の適正化、ケアプランの点検、縦覧点検・医療情報との突合、"
        "介護給付費通知により、給付の適正化に取り組みます。",
    "（１）事業所への助言・指導":
        "地域密着型サービス事業所の指定・指導監督を行い、"
        "運営の改善に向けた助言・指導を行います。",
    "（１）情報提供・相談体制の充実":
        "広報・ホームページ・パンフレット等により制度とサービスの情報を"
        "提供し、地域包括支援センターを窓口とする相談の体制を充実します。",
    "（１）介護ロボット・ＩＣＴの導入支援":
        "介護現場の生産性向上のため、介護ロボット・ＩＣＴの導入について"
        "情報提供と宮城県の事業への橋渡しを行います。",
}

MARK = re.compile(r"[(（][0-9０-９]{1,2}[)）]")
PART = ("を", "の", "に", "が", "は", "で", "と", "や", "も", "から", "など",
        "、", "。", "）", ")")


def key9(n):
    return re.sub(r"【.*?】", "", re.sub(r"^（[0-9０-９]+）", "", n)).strip()


def condense(seg):
    s = re.sub(r"\s+", "", seg).lstrip("・･")
    for c in ("■", "※"):
        i = s.find(c)
        if i > 0:
            s = s[:i]
    m = re.search(r"^(.{10,110}?。)", s)
    return m.group(1) if m else ""


def from_9ki():
    """第9期計画書の本文から、事業ごとの最初の一文を取り出す。"""
    rd = pypdf.PdfReader(PDF9)
    flat = re.sub(r"[ \t　]+", " ",
                  "\n".join((p.extract_text() or "") for p in rd.pages))
    wb = openpyxl.load_workbook(TAIKEI, data_only=True)
    ws = wb["01_統合体系表"]
    out = {}
    for r in range(2, ws.max_row + 1):
        n = ws.cell(r, 4).value
        if not n or str(ws.cell(r, 5).value or "") != "継続":
            continue
        n = str(n)
        kk = key9(n)
        best = None
        for m in re.finditer(re.escape(kk), flat):
            seg = flat[m.end(): m.end() + 260].replace("\n", "")
            score = len(MARK.findall(seg[:150]))
            if best is None or score < best[0]:
                best = (score, seg)
        if not best or best[0] > 1:
            continue
        s = re.sub(r"\s+", "", best[1]).lstrip("・･")
        if s.startswith(PART):
            continue
        c = condense(best[1])
        if len(c) >= 12:
            out[n] = c
    return out


def set_cell(cell, text, src=None):
    """セルの文字を差し替える。

    `add_column` が作るセルには run がなく、`set_tc` は run のない段落を
    黙って素通りする。隣のセルの段落を複製して書式ごと持ってくる。
    """
    from fix_soan_v111 import set_cell as f
    if not cell.paragraphs or not cell.paragraphs[0].runs:
        if src is None:
            raise SystemExit("run のないセルに雛形がない")
        tc = cell._tc
        for q in tc.findall(qn("w:p")):
            tc.remove(q)
        tc.append(copy.deepcopy(src._tc.findall(qn("w:p"))[0]))
    f(cell, text)


def from_shinki():
    """新規・拡充の事業の内容は、統合体系表の03シート（備考）から採る。"""
    wb = openpyxl.load_workbook(TAIKEI, data_only=True)
    ws = wb["03_新規拡充一覧"]
    out = {}
    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, 3).value
        bikou = ws.cell(r, 5).value
        if not name or not bikou:
            continue
        s = re.sub(r"\s+", "", str(bikou))
        m = re.search(r"^(.{10,110}?。)", s)
        out[str(name)] = m.group(1) if m else s[:110]
    return out


def main():
    auto = from_9ki()
    shin = from_shinki()
    naiyo = dict(auto)
    naiyo.update(shin)          # 新規・拡充
    naiyo.update(NAIYO)          # 当方が書き起こしたものを優先
    shutil.copy(SRC, DST)
    doc = docx.Document(DST)

    from fix_soan_v111 import set_el

    def widths(tbl, total, ratio):
        """列幅を ratio の比率に振り直す。総幅は欄を増やす前のまま保つ。"""
        grid = tbl.find(qn("w:tblGrid"))
        cols = grid.findall(qn("w:gridCol"))
        s2 = sum(ratio)
        new = [int(total * r / s2) for r in ratio]
        new[-1] = total - sum(new[:-1])
        for c, w in zip(cols, new):
            c.set(qn("w:w"), str(w))
        for tr in tbl.findall(qn("w:tr")):
            for tc, w in zip(tr.findall(qn("w:tc")), new):
                tcPr = tc.find(qn("w:tcPr"))
                if tcPr is None:
                    continue
                tcW = tcPr.find(qn("w:tcW"))
                if tcW is not None:
                    tcW.set(qn("w:w"), str(w))
                    tcW.set(qn("w:type"), "dxa")
    body = doc.element.body

    tbls = [t for t in doc.tables
            if [c.text.strip() for c in t.rows[0].cells][:2] == ["項", "事業"]]
    n_row = n_auto = n_man = n_shin = n_none = 0
    nashi = []
    for t in tbls:
        grid0 = t._tbl.find(qn("w:tblGrid"))
        total0 = sum(int(c.get(qn("w:w")))
                     for c in grid0.findall(qn("w:gridCol")))
        t.add_column(Pt(120))
        # 見出し行のセルの書式（地色など）を隣の欄から写す
        srcc = t.rows[0].cells[3]
        dst = t.rows[0].cells[4]._tc
        pr = srcc._tc.find(qn("w:tcPr"))
        if pr is not None:
            old = dst.find(qn("w:tcPr"))
            if old is not None:
                dst.remove(old)
            dst.insert(0, copy.deepcopy(pr))
        set_cell(t.rows[0].cells[4], "内容", src=srcc)
        for row in t.rows[1:]:
            name = row.cells[1].text.strip()
            if not name:
                continue
            n_row += 1
            v = naiyo.get(name)
            if v is None:
                # 表記の揺れを吸収して探す
                k = re.sub(r"\s", "", name)
                v = next((vv for kk, vv in naiyo.items()
                          if re.sub(r"\s", "", kk) == k), None)
            if v is None:
                v = "［要確認］　第9期計画書に記述がなく、町に事業の概要の" \
                    "ご確認をお願いします（確認事項No.8）。"
                n_none += 1
                nashi.append(name)
            elif name in NAIYO:
                n_man += 1
            elif name in shin:
                n_shin += 1
            else:
                n_auto += 1
            set_cell(row.cells[4], v, src=row.cells[3])
        widths(t._tbl, total0, (0.16, 0.25, 0.05, 0.14, 0.40))

    # ══════════════════════════ 版表記
    ps = [el for el in body.iterchildren() if el.tag == qn("w:p")]
    for i, txt in VER.items():
        set_el(ps[i - 1], txt)
    for el in body.iterchildren():
        if el.tag != qn("w:p"):
            continue
        t = "".join(n.text or "" for n in el.iter(qn("w:t")))
        new = t
        for a, b in VER_SWAP:
            new = new.replace(a, b)
        if new != t:
            set_el(el, new)
    doc.save(DST)

    # ヘッダー
    import zipfile
    tmp = DST + ".tmp"
    with zipfile.ZipFile(DST) as zin, \
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/header1.xml":
                data = re.sub(r"計画書素案 ?Ver\.[0-9.]+",
                              "計画書素案 Ver.2.9",
                              data.decode("utf-8")).encode("utf-8")
            info = zipfile.ZipInfo(item.filename, item.date_time)
            info.compress_type = item.compress_type
            info.external_attr = item.external_attr
            zout.writestr(info, data)
    shutil.move(tmp, DST)

    print("事業一覧表に「内容」の欄を加えました")
    print("  保存：", DST)
    print(f"  事業一覧表 {len(tbls)}／事業の行 {n_row}")
    print(f"   第9期計画書から取り出した内容　{n_auto}件")
    print(f"   統合体系表から採った内容（新規・拡充）{n_shin}件")
    print(f"   当方が書き起こした内容　　　　{n_man}件")
    print(f"   ［要確認］のまま残したもの　　{n_none}件")
    for x in nashi[:12]:
        print("      ", x[:46])
    # ══════════════════════════ 自己点検
    d2 = docx.Document(DST)
    t2 = [t for t in d2.tables
          if [c.text.strip() for c in t.rows[0].cells][:2] == ["項", "事業"]]
    bad = [t for t in t2 if len(t.columns) != 5]
    empty = 0
    for t in t2:
        for row in t.rows[1:]:
            if row.cells[1].text.strip() and not row.cells[4].text.strip():
                empty += 1
    print("  ── 自己点検")
    ng = []
    if len(t2) != len(tbls):
        ng.append(f"事業一覧表の数が変わった（{len(tbls)}→{len(t2)}）")
    if bad:
        ng.append(f"5欄になっていない表が{len(bad)}")
    if empty:
        ng.append(f"内容が空の行が{empty}")
    if ng:
        for m in ng:
            print("   ×", m)
        sys.exit(1)
    print(f"   ○ {len(t2)}表すべてが5欄で、{n_row}行すべてに内容がある")
    print("  ⚠ 欄を増やしたため表が縦に伸びます。頁数は目次の作り直しで数えます。")


if __name__ == "__main__":
    main()
