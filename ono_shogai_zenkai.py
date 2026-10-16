"""前回計画（令和6年3月）から、第6期の実績と第7期の見込量を読み取る。

なぜ原本から読むのか
  **数値を書き写すと、書き写した先が原本から離れていく。**
  本モジュールは原本の PDF を毎回読み、表から値を取り出す。
  書き写した定数を持たないため、原本と食い違いようがない。

原本
  13_関連計画/おのまち障がい者基本計画_第7期障がい福祉計画_
    第3期障がい児福祉計画_令和6年3月.pdf（128頁）

表の形（pypdf の抽出テキスト）
  ◆第６期の実績                     （1 か月あたり）
  第６期 ／ 見込量・実績 ／ 令和３年度 令和４年度 令和５年度
  見込量 実 績 見込量 実 績 見込量 実績見込
   生活介護
  728 人日 589 人日 728 人日 551 人日 728 人日 504 人日
  33 人 31 人 33 人 29 人 33 人 28 人

  ◆第７期の見込量                    （1 か月あたり）
  第７期 ／ 見込量 ／ 令和６年度 令和７年度 令和８年度
   生活介護
  728 人日 728 人日 728 人日
  26 人 26 人 26 人

読み方の約束
  ・サービス名は、値の行の直前にある「数字を含まない短い行」とする。
  ・1つのサービスに複数の単位（人日と人）があるときは、単位ごとに持つ。
  ・空欄（値のない欄）は None とする。**0 と区別する。**
    前回計画には「実績なし」を空欄にしている欄が多く、0 と同じではない。
  ・第6期は6つ（令和3〜5年度の 見込量・実績）、
    第7期は3つ（令和6〜8年度の見込量）の値を期待する。
    数が合わない塊は FUTOU に入れ、使わない。

使い方
  import ono_shogai_zenkai as Z
  Z.SERVICE["生活介護"]["人日"]["6期"]   → [728, 589, 728, 551, 728, 504]
  Z.SERVICE["生活介護"]["人"]["7期"]     → [26, 26, 26]
"""

import pathlib
import re
import warnings

warnings.filterwarnings("ignore")

ROOT = pathlib.Path(__file__).parent
PDF = (ROOT / "小野町_引継ぎ_整理済" / "13_関連計画"
       / "おのまち障がい者基本計画_第7期障がい福祉計画_"
         "第3期障がい児福祉計画_令和6年3月.pdf")

UNIT = r"人日|人回|時間|人|日|件|か所|箇所"
VAL = re.compile(r"(?:([0-9,]+)\s*)?(" + UNIT + r")(?=\s|$)")
# 値の行かどうか。単位が2つ以上並んでいれば値の行とみなす。
NG_NAME = re.compile(r"見込量|実\s*績|年度|第[６-９6-9]期|給付の種類|内\s*容|"
                     r"対象者|か月あたり|単位|第\s*\d|部\s|章\s|^-|^\s*$")

# 原本の単位が行まるごと誤っているもの。**理由を書いたものだけ直す。**
#   自立生活援助は人で数えるサービス（指定基準・基本指針とも「人」）。
#   前回計画の第6期の実績の行は「0 人日」が6つ並んでおり、単位が誤っている。
#   第7期の見込量の行は「0 人」で正しい。値はいずれも0のため、
#   単位を直しても数値は動かない。
TANI_TEISEI = {("自立生活援助", "6期"): ("人日", "人")}

# 単位が入れ替わっていると見られるもの。**直さずに印を付けるだけにする。**
#   障がい児の「◆第７期の目標」の表は、児童発達支援を「38 人・4 人日」、
#   放課後等デイサービスを「391 人・23 人日」としている。
#   第6期の実績は児童発達支援が6人・24〜28人日、放課後等デイサービスが
#   31人・93人日であり、人と人日が入れ替わっているとみられる。
#   **本町の手帳所持者は508人であり、放課後等デイサービスを月391人が
#   利用することはありえない。**
#   ただし当方の読み違いである可能性も残るため、値は原本のまま持ち、
#   AYASHII に入れて「確かめるもの」として扱う。
AYASHII = [
    ("児童発達支援", "7期",
     "「38 人・4 人日」。第6期の実績は6人・24〜28人日。"
     "**人と人日が入れ替わっているとみられる**"),
    ("放課後等デイサービス", "7期",
     "「391 人・23 人日」。第6期の実績は31人・93人日。"
     "**本町の手帳所持者は508人であり、月391人の利用はありえない**"),
]

SERVICE = {}        # 名前 → {単位 → {"6期": [...], "7期": [...]}}
FUTOU = []          # 数が合わなかった塊（名前, 期, 単位, 値）
TANI_MIDARE = []    # 原本の単位が1行の中で混ざっているもの
_PAGES = []


def _text():
    import pypdf
    r = pypdf.PdfReader(str(PDF))
    return [(i + 1, p.extract_text() or "") for i, p in enumerate(r.pages)]


def _vals(ln):
    """値の行なら [値 or None, ...] と単位を返す。値の行でなければ None。

    **原本に単位の誤りがある行がある。**
    施設入所支援の第6期の実績は「19 人 16 人 19 人日 15 人 19 人日 14 人」で、
    同じ行に 人 と 人日 が混ざっている（施設入所支援は人で数えるもので、
    人日は誤り）。多数を占める単位を採り、TANI_MIDARE に記録する。
    黙って直さず、記録して見えるようにする。
    """
    got = VAL.findall(ln)
    if len(got) < 2:
        return None, None
    units = [u for _v, u in got]
    u = max(set(units), key=units.count)
    if len(set(units)) != 1:
        TANI_MIDARE.append((ln.strip()[:60], sorted(set(units)), u))
    return [int(v.replace(",", "")) if v else None for v, _u in got], u


def _blocks(txt):
    """（期, 塊の行）に切る。"""
    out, cur, kind = [], [], None
    for ln in txt.split("\n"):
        if "◆第６期の実績" in ln:
            if kind:
                out.append((kind, cur))
            kind, cur = "6期", []
            continue
        if "◆第７期の見込量" in ln or "◆第７期の目標" in ln:
            # 障がい児の章だけ「◆第７期の目標」という見出しになっている。
            # 中身は見込量の表である。
            if kind:
                out.append((kind, cur))
            kind, cur = "7期", []
            continue
        if kind:
            cur.append(ln)
    if kind:
        out.append((kind, cur))
    return out


def _parse():
    SERVICE.clear()
    FUTOU.clear()
    TANI_MIDARE.clear()
    want = {"6期": 6, "7期": 3}
    for pno, txt in _PAGES:
        for kind, body in _blocks(txt):
            # サービス名は2行3行に折り返されている
            # （「就労継続支援」＋「（Ｂ型）」など）。
            # 値の行が来るまでに溜めた行をつないで1つの名にする。
            name, buf, atode = None, [], False
            for ln in body:
                vals, u = _vals(ln)
                if vals is None:
                    s = ln.strip()
                    if s and not NG_NAME.search(s) and 1 <= len(s) <= 24:
                        if atode:          # 値の行のあとに来た名は次のもの
                            buf, atode = [], False
                        buf.append(s)
                    continue
                # 値と同じ行に名が書かれていることがある
                # （「 計画相談支援 7 人 11 人 …」）。その場合は行頭の
                # 文字列を名とし、溜めた行より優先する。
                atama = ln[:VAL.search(ln).start()].strip()
                atama = re.sub(r"\s+", "", atama)
                if atama and not NG_NAME.search(atama) and len(atama) <= 24:
                    name, buf, atode = atama, [], False
                elif buf:
                    name = re.sub(r"\s+", "", "".join(buf))
                atode = True
                if not name:
                    continue
                if len(vals) != want[kind]:
                    FUTOU.append((name, kind, u, vals, pno))
                    continue
                naoshi = TANI_TEISEI.get((name, kind))
                if naoshi and naoshi[0] == u:
                    TANI_MIDARE.append(
                        ("%s %s の単位 %s → %s（原本の誤り）"
                         % (name, kind, naoshi[0], naoshi[1]), [u], naoshi[1]))
                    u = naoshi[1]
                d = SERVICE.setdefault(name, {}).setdefault(u, {})
                if kind in d:                     # 同じ名が2度出たら後を捨てる
                    FUTOU.append((name, kind, u, vals, pno))
                    continue
                d[kind] = vals


def load():
    if not _PAGES:
        _PAGES.extend(_text())
        _parse()
        _shisaku()
    return SERVICE


# ------------------------------------------------- 第2部の施策（9分野）
MARU = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮"
SHISAKU = []        # (頁, 章, 項目名, 取組の数)


def _shisaku():
    """第2部（障がい者計画）の「施策の方向」の項目を取り出す。

    各章は「施策の方向」の見出しのあと、①②③… の項目が並び、
    それぞれに ○ で始まる取組が付く。
    **「施策の方向」より前は見ない。** そこは現状・課題と
    アンケートの選択肢（①居宅介護…）であり、施策ではない。
    """
    SHISAKU.clear()
    sho = ""
    for pno, t in _PAGES:
        if "第 2 部 障がい者計画" not in t:
            continue
        m = re.search(r"第\s*(\d+)\s*章\s+(\S.{0,40})", t)
        if m:
            sho = "第%s章　%s" % (m.group(1), m.group(2).strip())
        lines = [x.strip() for x in t.split("\n")]
        for h, ln in enumerate(lines):
            if "施策の方向" in ln:
                lines = lines[h + 1:]
                break
        else:
            continue
        j = 0
        while j < len(lines):
            cur = lines[j]
            if cur and cur[0] in MARU:
                name, k = cur, j + 1
                while (k < len(lines) and lines[k]
                       and lines[k][0] not in MARU + "○"
                       and len(name) < 34 and k - j < 3):
                    name += lines[k]
                    k += 1
                # **○ は項目名と同じ行に始まることがある。**
                # 「①意思疎通支援事業の充実 ○手話通訳者及び…」
                # 名を切る前に、塊の全体で ○ を数える。
                k2 = k
                while k2 < len(lines) and not (lines[k2]
                                               and lines[k2][0] in MARU):
                    k2 += 1
                n = "".join(lines[j:k2]).count("○")
                name = re.sub(r"\s+", "", name).split("○")[0]
                SHISAKU.append((pno, sho, name, n))
                j = k
                continue
            j += 1


# 前回計画そのものの誤記。**直さず、見つけたものとして残す。**
GOKI = [
    ("第10章のページヘッダ", "文化芸術活動・スポーツ活動等の**進行**",
     "目次と本文の見出しは「**振興**」。ヘッダだけが誤り"),
    ("第10章の節の名", "１ **総合的な就労支援**",
     "**第9章第1節の名がそのまま入っている。**"
     "章は文化芸術活動・スポーツ活動等の振興"),
]


# ---------------------------------------------------------------- 自己点検
def jiko():
    load()
    ng = []
    if not PDF.exists():
        return ["原本が見当たらない：%s" % PDF.name]
    # 1 よく知られた値と一致する（人が原本で読んだもの）
    TASHIKA = [
        ("生活介護", "人日", "6期", [728, 589, 728, 551, 728, 504]),
        ("生活介護", "人", "6期", [33, 31, 33, 29, 33, 28]),
        ("生活介護", "人日", "7期", [728, 728, 728]),
        ("就労継続支援（B型）", "人日", "6期",
         [998, 561, 1078, 595, 1164, 612]),
        ("計画相談支援", "人", "6期", [7, 11, 7, 12, 7, 12]),
    ]
    for nm, u, k, want in TASHIKA:
        got = SERVICE.get(nm, {}).get(u, {}).get(k)
        if got != want:
            ng.append("%s %s %s が %s（原本は %s）" % (nm, u, k, got, want))
    # 2 どのサービスも、値の数が期ごとに決まった数である
    for nm, by in SERVICE.items():
        for u, by2 in by.items():
            for k, v in by2.items():
                n = 6 if k == "6期" else 3
                if len(v) != n:
                    ng.append("%s %s %s の値が%d個" % (nm, u, k, len(v)))
    # 3 単位の乱れは、記録したものだけである（黙って直していない）
    if len(TANI_MIDARE) > 4:
        ng.append("単位の乱れが%d件ある（記録を確かめること）" % len(TANI_MIDARE))
    # 4 入れ替わりの疑いがあるものは、原本のまま持っている（直していない）
    for nm, k, _riyu in AYASHII:
        by = SERVICE.get(nm, {})
        if not any(k in d for d in by.values()):
            ng.append("%s %s が読めていない（AYASHII に挙げたもの）" % (nm, k))
    # 5 第2部の施策が9章ぶん読めている
    n_sho = len({x[1] for x in SHISAKU})
    if n_sho != 9:
        ng.append("第2部の施策の章が%d件（9件のはず）" % n_sho)
    if len(SHISAKU) < 40:
        ng.append("第2部の施策の項目が%d件しかない" % len(SHISAKU))
    # 6 何も読めていない、ということがない
    if len(SERVICE) < 15:
        ng.append("読み取れたサービスが%d件しかない" % len(SERVICE))
    return ng


if __name__ == "__main__":
    load()
    print("■ 前回計画から読み取った実績・見込量")
    print("   原本 %s（%d頁）" % (PDF.name, len(_PAGES)))
    print("   サービス %d件／数が合わなかった塊 %d件"
          % (len(SERVICE), len(FUTOU)))
    print()
    for nm in sorted(SERVICE):
        for u, by in sorted(SERVICE[nm].items()):
            r6 = by.get("6期")
            r7 = by.get("7期")
            print("   %-22s %-3s  6期 %-34s  7期 %s"
                  % (nm, u, r6 if r6 else "―", r7 if r7 else "―"))
    if AYASHII:
        print()
        print("■ 単位が入れ替わっている疑いがあるもの（原本のまま持っている）")
        for nm, k, riyu in AYASHII:
            print("   %-18s %s  %s" % (nm, k, riyu.replace("**", "")))
    if TANI_MIDARE:
        print()
        print("■ 原本の単位の乱れ（多数の単位を採った／理由を書いて直した）")
        for ln, us, u in TANI_MIDARE:
            print("   %-62s %s → %s" % (ln, "・".join(us), u))
    if FUTOU:
        print()
        print("■ 数が合わなかった塊（使っていない）")
        for nm, k, u, v, pno in FUTOU[:24]:
            print("   p%-4d %-20s %-3s %-3s %s" % (pno, nm, k, u, v))
        if len(FUTOU) > 24:
            print("   ほか%d件" % (len(FUTOU) - 24))
    print()
    print()
    print("■ 第2部の施策（9分野）")
    last = None
    for pno, sho, nm, n in SHISAKU:
        if sho != last:
            print("   ── %s" % sho)
            last = sho
        print("      p%-4d %-32s 取組%d" % (pno, nm, n))
    print("   章%d／項目%d／取組%d"
          % (len({x[1] for x in SHISAKU}), len(SHISAKU),
             sum(x[3] for x in SHISAKU)))
    print()
    print("■ 前回計画そのものの誤記（直していない）")
    for a, b, c in GOKI:
        print("   %-18s %s" % (a, b.replace("**", "")))
        print("      %s" % c.replace("**", ""))
    ng = jiko()
    print()
    print("■ 自己点検")
    if ng:
        for x in ng[:20]:
            print("   **%s**" % x)
    else:
        print("   すべて適合")
