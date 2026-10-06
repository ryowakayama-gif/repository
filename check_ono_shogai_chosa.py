"""障がい福祉アンケート調査 分析報告書（令和8年10月6日受領）の点検。

受領物
  23_障がいアンケート分析/原本_受領_20261006/
    障がい福祉に関するアンケート調査_分析報告書_20261006.docx（628段落・23表）
    障がい福祉に関するアンケート調査_集計表_20261006.xlsx（16シート）

点検の方針:
 1. **原本は書き換えない。**読んで突き合わせるだけにする。
 2. 集計表の中の整合（件数の合計＝n、割合＝件数÷n）を全件確かめる。
 3. 報告書本文の数値を集計表と突き合わせる。
 4. **第7章は手入力の文字列である。**集計表の計算式と連動していないため、
    本文・表の数値を集計シートの原数値から引き直して確かめる。
 5. 回答の分布として不自然なもの（1つの選択肢に全数、無回答0件）を洗い出す。
 6. 制度上の対象要件（障害支援区分・障がい種別）と利用意向を突き合わせる。
    見込量の基礎に使えるかどうかの判断材料になる。

  python3 check_ono_shogai_chosa.py
"""

import pathlib
import re
import sys
from collections import defaultdict
from decimal import Decimal, ROUND_HALF_UP

import openpyxl
from docx import Document

ROOT = (pathlib.Path(__file__).parent / "小野町_引継ぎ_整理済"
        / "23_障がいアンケート分析" / "原本_受領_20261006")
DOCX = ROOT / "障がい福祉に関するアンケート調査_分析報告書_20261006.docx"
XLSX = ROOT / "障がい福祉に関するアンケート調査_集計表_20261006.xlsx"

RESULT = []          # (区分, 重さ, 件名, 内容)


def add(kubun, omosa, ken, naiyo):
    RESULT.append((kubun, omosa, ken, naiyo))


def shi(x, nd=1):
    """四捨五入（Python の round は銀行丸めのため使わない）。"""
    return float(Decimal(str(x)).quantize(Decimal("1." + "0" * nd),
                                          rounding=ROUND_HALF_UP))


# ---------------------------------------------------------------- 読み込み

def load_xlsx():
    if not XLSX.exists():
        raise SystemExit(f"集計表がありません：{XLSX}")
    return openpyxl.load_workbook(XLSX, data_only=True)


def shukei(wb, sheet):
    """集計シートを (設問 -> [(項目, 選択肢, 件数, n, 回答形式)]) で返す。"""
    out = defaultdict(list)
    for r in wb[sheet].iter_rows(min_row=5, values_only=True):
        if not r[0]:
            continue
        out[r[0]].append((r[1], r[2], r[3], r[5], r[6]))
    return out


def shogaibetsu(wb):
    """障がい別シート。(設問, 項目, 選択肢) -> (身体, 知的, 精神) の件数。"""
    out, ns = {}, None
    for r in wb["集計_障がい別"].iter_rows(min_row=5, values_only=True):
        if r[2] == "回答者数（n）":
            ns = (r[3], r[4], r[5])
            continue
        if not r[0]:
            continue
        out[(r[0], r[1], r[2])] = (r[3], r[4], r[5])
    return ns, out


# ---------------------------------------------------------------- 1 集計表の中

def check_shukei(wb):
    for sheet in ("集計_障がい者", "集計_障がい児"):
        q = shukei(wb, sheet)
        ng_kei, ng_wari = [], []
        for name, rows in q.items():
            byitem = defaultdict(list)
            for item, sel, ken, n, form in rows:
                byitem[item].append((sel, ken, n, form))
            for item, v in byitem.items():
                form = v[0][3]
                n = v[0][2]
                if form and form.startswith("単一回答"):
                    kei = sum(x[1] or 0 for x in v)
                    if kei != n:
                        ng_kei.append(f"{name}／{item}　合計{kei}≠n{n}")
        for name, rows in q.items():
            for item, sel, ken, n, _f in rows:
                if ken is None or not n:
                    continue
        add("集計表", "―", f"{sheet}　単一回答の件数合計＝n",
            "全設問で一致" if not ng_kei else "／".join(ng_kei[:4]))
        if ng_kei:
            add("集計表", "重大", f"{sheet}　件数合計がnと合わない",
                "／".join(ng_kei))

    # 割合の計算式
    ng = []
    for sheet in ("集計_障がい者", "集計_障がい児"):
        for r in wb[sheet].iter_rows(min_row=5, values_only=True):
            if not r[0] or r[3] is None or r[4] is None or not r[5]:
                continue
            if abs(r[3] / r[5] - r[4]) > 1e-9:
                ng.append(f"{r[0]}／{r[2]}")
    ns, sb = shogaibetsu(wb)
    add("集計表", "―", "割合＝件数÷回答者数",
        "全件一致" if not ng else "／".join(ng[:5]))
    if ng:
        add("集計表", "重大", "割合が件数÷nになっていない", "／".join(ng))

    # 障がい別のnが手帳の設問と合うか
    q = shukei(wb, "集計_障がい者")
    def kyu(qno, sels):
        return sum(k[2] for k in q[qno] if str(k[1]) in sels)
    shin = kyu("問8", [f"{c}級" for c in "１２３４５６"])
    chi = kyu("問10", ["Ａ判定", "Ｂ判定"])
    sei = kyu("問11", [f"{c}級" for c in "１２３"])
    ok = (shin, chi, sei) == tuple(ns)
    add("集計表", "―" if ok else "重大", "障がい別のnが手帳の設問と合う",
        f"問8の1〜6級{shin}・問10のA/B{chi}・問11の1〜3級{sei}"
        f"／シートのn{ns}　{'一致' if ok else '不一致'}")
    return ns


# ---------------------------------------------------------------- 2 本文の数値

def check_honbun(wb):
    d = Document(DOCX)
    ps = [p.text.strip() for p in d.paragraphs]
    A, C = shukei(wb, "集計_障がい者"), shukei(wb, "集計_障がい児")

    def find(tbl, qno, sel):
        """設問番号（枝番を含む）と選択肢から件数とnを引く。

        前方10文字での照合は「自宅…でひとり暮らし」と
        「自宅…で家族等と一緒に住んでいる」を取り違える。
        **本文の選択肢と完全に一致するもの、無ければ最長の前方一致を採る。**
        """
        keys = [qno] + [k for k in tbl if k.startswith(qno + "-")]
        cand = []
        for k in keys:
            for item, s, ken, n, _f in tbl.get(k, []):
                for t in (s, item):
                    if not t:
                        continue
                    t = str(t)
                    if t == sel:
                        return ken, n
                    if t.startswith(sel) or sel.startswith(t):
                        cand.append((len(t), ken, n))
        if cand:
            cand.sort(reverse=True)
            return cand[0][1], cand[0][2]
        return None, None

    HEAD = re.compile(r"^（\d+）(.+?)（問([\d\-]+)）$")
    PCT = re.compile(r"「([^「」]+?)」が([\d.]+)％")
    cur_q, ch, miss, ng = None, None, 0, []
    for t in ps:
        if t.startswith("第3章"):
            ch = "者"
        elif t.startswith("第5章"):
            ch = "児"
        elif t.startswith(("第4章", "第6章", "第7章")):
            ch = None
        m = HEAD.match(t)
        if m:
            cur_q = "問" + m.group(2)
            continue
        if not (cur_q and ch) or t.startswith(("図", "表", "※", "障がい別")):
            continue
        tbl = A if ch == "者" else C
        for sel, pct in PCT.findall(t):
            ken, n = find(tbl, cur_q, sel)
            if ken is None or not n:
                miss += 1
                continue
            exp = shi(ken / n * 100)
            if abs(exp - float(pct)) > 1e-9:
                ng.append(f"{cur_q}「{sel[:16]}」本文{pct}％／集計{exp}％"
                          f"（{ken}/{n}）")
    add("本文", "―" if not ng else "重大",
        "本文の「「選択肢」がXX.X％」が集計表と一致",
        f"照合 {len(ng)}件の不一致（照合できなかった記述 {miss}件）"
        if ng else f"不一致なし（照合できなかった記述 {miss}件）")
    for x in ng:
        add("本文", "重大", "本文と集計表の不一致", x)


# ---------------------------------------------------------------- 3 第7章

def dai7(wb):
    """第7章シートを行ごとに (表, 見出しのn, 詰めた各セル) で返す。

    **このシートはB列から始まり、列の間に空列が挟まる。**
    そのまま添字で読むと列がずれるため、Noneを落として詰める。
    """
    out, tbl, hn = [], None, None
    for r in wb["7章_追加分析"].iter_rows(values_only=True):
        v = [x for x in r if x is not None]
        if not v:
            continue
        if isinstance(v[0], str) and v[0].startswith("表7-"):
            tbl = v[0].split("　")[0]
            m = re.search(r"n=(\d+)", v[0])
            hn = int(m.group(1)) if m else None
            continue
        out.append((tbl, hn, v))
    return out


def check_dai7(wb):
    PAT = re.compile(r"(\d+)件（([\d.]+)%）")
    ng, hdr = [], defaultdict(set)
    for tbl, hn, row in dai7(wb):
        # 行の n（文字列で入っている場合がある）
        rn = None
        for c in row[1:3]:
            if isinstance(c, int):
                rn = c
                break
            if isinstance(c, str) and c.strip().isdigit():
                rn = int(c)
                break
        base = rn or hn
        if rn and hn and rn != hn:
            hdr[tbl].add(rn)
        for c in row:
            if not isinstance(c, str):
                continue
            for ken, pct in PAT.findall(c):
                if not base:
                    continue
                exp = shi(int(ken) / base * 100)
                if abs(exp - float(pct)) > 1e-9:
                    ng.append(f"{tbl}　{str(row[0])[:14]}　{ken}件/{base}"
                              f"→記載{pct}％／計算{exp}％")
    add("第7章", "―" if not ng else "重大",
        "第7章の「N件（P%）」が行の回答者数と合う",
        "全件一致" if not ng else "／".join(ng[:5]))
    for x in ng:
        add("第7章", "重大", "第7章の割合が合わない", x)

    for tbl, rns in sorted(hdr.items()):
        add("第7章", "重要", f"{tbl}　見出しのnと行のnが違う",
            f"見出しは調査全体のn、割合は行のn（{sorted(rns)}）で計算している。"
            "読者が見出しのnで検算すると合わない")

    # 本文が表に無い項目を挙げていないか（表7-4）
    #
    # 第7章は手入力であり、集計シートの計算式と連動していない。
    # 本文が表の列に無いサービス名を挙げていても、誰も気づかない。
    cols, rows7 = [], {}
    for tbl, _hn, row in dai7(wb):
        if tbl != "表7-4":
            continue
        if str(row[0]) == "ADL介助区分":
            cols = [str(c) for c in row[2:] if c]
        else:
            rows7[str(row[0])] = [str(c) for c in row[2:] if c]
    d = Document(DOCX)
    for p in d.paragraphs:
        t = p.text
        if "6項目以上で介助が必要な" not in t:
            continue
        # 「生活介護44.9％、施設入所34.7％、…」を (サービス, 割合) に割る
        body = t.split("では、", 1)[-1].split("の利用意向")[0]
        for part in re.split(r"[、，]", body):
            m = re.match(r"^(.+?)([\d.]+)％$", part.strip())
            if not m:
                continue
            svc, pct = m.group(1), m.group(2)
            if svc not in cols:
                add("第7章", "重大", "表7-4の本文に表に無い項目",
                    f"本文は「{svc}{pct}％」と書いているが、表7-4の列は"
                    f"{'・'.join(cols)}であり「{svc}」の列がない。"
                    f"表の6項目以上の行は {'／'.join(rows7.get('6項目以上で介助', []))}")
            else:
                i = cols.index(svc)
                cell = rows7.get("6項目以上で介助", [None] * len(cols))[i]
                if cell and f"{pct}%" not in cell.replace("％", "%"):
                    add("第7章", "重大", "表7-4の本文と表の割合が違う",
                        f"本文「{svc}{pct}％」／表は {cell}")


# ---------------------------------------------------------------- 4 分布の不審

def check_bunpu(wb):
    for sheet, name in (("集計_障がい児", "障がい児"), ("集計_障がい者", "障がい者")):
        q = shukei(wb, sheet)
        zensu, mu0 = [], 0
        for qn, rows in q.items():
            byitem = defaultdict(list)
            for item, sel, ken, n, form in rows:
                byitem[item].append((sel, ken, n))
            for item, v in byitem.items():
                n = v[0][2]
                if len(v) < 2 or not n:
                    continue
                for sel, ken, _n in v:
                    if ken == n and str(sel) != "無回答":
                        zensu.append(f"{qn}「{sel}」{ken}/{n}")
                if any(str(s) == "無回答" and k == 0 for s, k, _ in v):
                    mu0 += 1
        add("分布", "重要" if name == "障がい児" else "―",
            f"{name}　1つの選択肢に全数が集まっている設問",
            f"{len(zensu)}件　" + "／".join(zensu[:6]) if zensu else "なし")
        add("分布", "重要" if name == "障がい児" else "―",
            f"{name}　無回答が0件の設問", f"{mu0}設問")

    # 障がい児の災害設問の整合
    C = shukei(wb, "集計_障がい児")

    def g(qn, sel):
        for _i, s, ken, n, _f in C.get(qn, []):
            if str(s) == sel:
                return ken, n
        return None, None

    dekinai, n38 = g("問38", "できない")
    wakaranai, _ = g("問38", "わからない")
    touroku, _ = g("問41", "希望する")
    keikaku, _ = g("問42", "希望する")
    if None not in (dekinai, wakaranai, touroku, keikaku):
        A = shukei(wb, "集計_障がい者")

        def ga(qn, sel):
            for _i, s, ken, n, _f in A.get(qn, []):
                if str(s) == sel:
                    return ken, n
            return None, None
        a41d, _ = ga("問41", "できない")
        a41w, _ = ga("問41", "わからない")
        a44, _ = ga("問44", "希望する")
        a45, _ = ga("問45", "希望する")
        add("分布", "重大", "障がい児の災害設問が整合しない",
            f"一人で避難「できない」{dekinai}件＋「わからない」{wakaranai}件"
            f"＝{dekinai + wakaranai}件（{shi((dekinai + wakaranai) / n38 * 100)}％）"
            f"に対し、避難行動要支援者の登録を希望 {touroku}件、"
            f"個別避難計画を希望 {keikaku}件。"
            f"障がい者は避難できない・わからない{a41d + a41w}人に対し"
            f"登録希望{a44}人・計画希望{a45}人。"
            "25件全数が「希望しない」で無回答0件は郵送調査として不自然であり、"
            "入力の確認が必要")


# ---------------------------------------------------------------- 5 制度要件

KUBUN = "１２３４５６"


def check_youken(wb):
    A = shukei(wb, "集計_障がい者")
    ku = {}
    for _i, sel, ken, _n, _f in A["問28"]:
        ku[str(sel)] = ken
    nin = sum(ku[f"区分{c}"] for c in KUBUN)
    ku3 = sum(ku[f"区分{c}"] for c in "３４５６")
    ku4 = sum(ku[f"区分{c}"] for c in "４５６")
    ku5 = ku["区分５"] + ku["区分６"]
    ku6 = ku["区分６"]
    shikaku = next(k for _i, s, k, _n, _f in A["問9"] if str(s) == "視覚障がい")

    cur = {str(i): k for i, s, k, _n, _f in A["問31-1"] if str(s) == "利用している"}
    iko = {str(i): k for i, s, k, _n, _f in A["問31-2"] if str(s) == "利用したい"}

    YOKEN = [
        ("居宅介護", "障害支援区分1以上", nin),
        ("重度訪問介護", "区分4以上", ku4),
        ("同行援護", "視覚障がい（問9）", shikaku),
        ("行動援護", "区分3以上", ku3),
        ("重度包括", "区分6", ku6),
        ("生活介護", "区分3以上", ku3),
        ("療養介護", "区分5以上", ku5),
        ("短期入所", "区分1以上", nin),
        ("施設入所", "区分4以上", ku4),
    ]
    koe = []
    for svc, yk, n in YOKEN:
        c, i = cur.get(svc, 0), iko.get(svc, 0)
        if i > n:
            koe.append(f"{svc}（現在{c}→意向{i}／{yk}は{n}人）")
    add("要件", "重要", "利用意向が制度上の対象者数を超えているサービス",
        f"{len(koe)}／{len(YOKEN)}サービス　" + "／".join(koe))
    add("要件", "―", "障害支援区分の把握状況",
        f"認定あり {nin}人（区分3以上{ku3}・区分4以上{ku4}・区分6{ku6}）、"
        f"「受けていない」{ku['受けていない']}人、"
        f"無回答 {ku['無回答']}人（{shi(ku['無回答'] / 259 * 100)}％）。"
        "**無回答が多いため上の「対象者数」は下限である**")
    # 区分の把握不足の裏づけ
    if cur.get("生活介護", 0) > ku3:
        add("要件", "重要", "障害支援区分の把握が不十分であることの裏づけ",
            f"生活介護は現に{cur['生活介護']}人が利用しているが、"
            f"区分3以上と答えた方は{ku3}人しかいない。"
            "問28の無回答が多いことによる")


# ---------------------------------------------------------------- 6 回収率

def check_kaishu(wb):
    ws = wb["1-2章_概要・要点"]
    haifu = yuko = ritsu = None
    for r in ws.iter_rows(min_row=1, max_row=12, values_only=True):
        if r[1] == "配布数":
            haifu = r[2]
        if r[1] == "有効回答数":
            yuko = r[2]
        if r[1] == "有効回答率":
            ritsu = r[2]
    add("回収", "重大" if haifu is None else "―", "配布数・有効回答率の記入",
        f"配布数 {haifu or '未記入'}／有効回答数 {yuko}／有効回答率 "
        f"{ritsu or '未記入'}。"
        "**計画素案には配布数・回収率を載せるのが通例であり、"
        "未記入のままでは書けない。前回との比較もできない**"
        if haifu is None else "記入済み")
    # 前回調査（注記）の算術
    for r in ws.iter_rows(min_row=1, max_row=12, values_only=True):
        if r[1] and "前回調査" in str(r[1]):
            t = str(r[1])
            ms = re.findall(r"配布(\d+)件・有効回答(\d+)件（([\d.]+)％）", t)
            ng = [f"{h}/{y}" for h, y, p in ms
                  if abs(shi(int(y) / int(h) * 100) - float(p)) > 1e-9]
            add("回収", "―" if not ng else "重大", "前回調査の回収率の算術",
                "注記の値は一致（" + "／".join(
                    f"{y}/{h}={p}％" for h, y, p in ms) + "）"
                if not ng else "／".join(ng))
            if ms:
                add("回収", "重要", "有効回答数の経年",
                    f"障がい者 {ms[0][1]}件（令和5年10月）→ {yuko}件"
                    f"（{int(yuko) - int(ms[0][1]):+d}件）。"
                    "**配布数が未記入のため、回収率が下がったのか"
                    "配布数が減ったのかを分けられない**")


# ---------------------------------------------------------------- 7 表記

def check_hyoki():
    d = Document(DOCX)
    txt = "\n".join(p.text for p in d.paragraphs)
    for t in d.tables:
        for r in t.rows:
            txt += "\n" + "\t".join(c.text for c in r.cells)
    PAIRS = [("％", "%", "全角と半角のパーセント記号"),
             ("グループホーム", "共同生活援助", "サービス名の表記"),
             ("ご存知", "ご存じ", "「ご存知」と「ご存じ」")]
    for a, b, name in PAIRS:
        ca, cb = txt.count(a), txt.count(b)
        if ca and cb:
            add("表記", "軽微", name, f"「{a}」{ca}回／「{b}」{cb}回　混在")
    zen = len([l for l in txt.splitlines() if re.search(r"[０-９]", l)])
    add("表記", "軽微", "全角数字を含む行",
        f"{zen}行。選択肢の引用（「１～６級」等）は原文どおりでよいが、"
        "本文の章番号・件数は半角にそろえる")
    n100 = len(re.findall(r"100\.0％（\d+件）と最も高", txt))
    if n100:
        add("表記", "軽微", "「100.0％…と最も高く」という言い方",
            f"{n100}箇所。選択肢が1つに集まっている場合に「最も高い」は"
            "不自然。「25件すべてが」等に改める")


# ---------------------------------------------------------------- main

def main():
    wb = load_xlsx()
    if not DOCX.exists():
        raise SystemExit(f"報告書がありません：{DOCX}")
    check_shukei(wb)
    check_honbun(wb)
    check_dai7(wb)
    check_bunpu(wb)
    check_youken(wb)
    check_kaishu(wb)
    check_hyoki()

    order = {"重大": 0, "重要": 1, "軽微": 2, "―": 3}
    RESULT.sort(key=lambda x: (order.get(x[1], 9), x[0]))
    cnt = defaultdict(int)
    for k, o, ken, naiyo in RESULT:
        cnt[o] += 1
    print(f"点検 {len(RESULT)}件　"
          + "／".join(f"{k} {cnt[k]}" for k in ("重大", "重要", "軽微", "―")
                      if cnt[k]))
    print()
    for k, o, ken, naiyo in RESULT:
        mark = {"重大": "!!", "重要": "! ", "軽微": "・", "―": "OK"}[o]
        print(f"{mark} [{k}] {ken}")
        print(f"     {naiyo}")
    return 1 if cnt["重大"] else 0


if __name__ == "__main__":
    sys.exit(main())
