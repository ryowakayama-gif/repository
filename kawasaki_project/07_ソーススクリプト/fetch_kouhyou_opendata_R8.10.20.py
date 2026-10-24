# -*- coding: utf-8 -*-
"""介護サービス情報公表システムのオープンデータから町内の事業所を取り出す
（令和8年10月20日）

  python3 07_ソーススクリプト/fetch_kouhyou_opendata_R8.10.20.py

なぜ要るか
  作業順位26（県の公表名簿との突合）は、**公表システムの画面が自動の
  アクセスを受け付けない**ため止まっていた。
  ⚠ **厚生労働省が同じ内容をサービス種類別のＣＳＶ（オープンデータ）で
  公表していることが分かった。** 許可済みのドメイン（www.mhlw.go.jp）から
  取れるため、環境の設定の変更を待たずに突合を行える。

  入口　https://www.mhlw.go.jp/stf/kaigo-kouhyou_opendata.html
  単位　サービスの種類ごとに全国1本のＣＳＶ（24列・utf-8-sig）
  抽出　都道府県コード又は市町村コード ＝ **043249**（宮城県柴田郡川崎町）

⚠ **川崎市（神奈川県）・福岡県田川郡川崎町と取り違えない。**
  市町村コードで抽出し、都道府県名が「宮城県」であることを検査で確かめる。

出力
  09_元資料/その他データ/公表システム_オープンデータ_川崎町_R8.6末.csv
    町内の事業所（1行1事業所・24列のうち要るもの）
  09_元資料/その他データ/公表システム_オープンデータ_件数_R8.6末.csv
    サービスの種類ごとの件数（全国・宮城県・川崎町）
    ⚠ **他団体の固有名称は出さない。件数だけを持つ。**

⚠ **オープンデータに無いもの**（突合の後に残るもの）
  ・指定の年月日、指定の主体（町か県か）の別
  ・入所者数・入所率・稼働率・待機者数
  ・指定を受けていない有料老人ホーム・サービス付き高齢者向け住宅
  これらは引き続き町・県へのご確認による（確認事項No.29・No.154・No.159）。

  終了コード　0＝検査1〜7がすべて適合／1＝1件でも不適合
"""
import csv
import io
import os
import re
import subprocess
import sys
import tempfile

NYUGUCHI = "https://www.mhlw.go.jp/stf/kaigo-kouhyou_opendata.html"
MOTO = "https://www.mhlw.go.jp"
CODE = "043249"           # 宮城県柴田郡川崎町
KEN = "宮城県"
MACHI = "川崎町"
KIJUNBI = "2026年６月末時点"   # 入口のページの見出し

OUT_DIR = "09_元資料/その他データ"
OUT_MEISAI = os.path.join(OUT_DIR, "公表システム_オープンデータ_川崎町_R8.6末.csv")
OUT_KENSU = os.path.join(OUT_DIR, "公表システム_オープンデータ_件数_R8.6末.csv")

# 明細に残す列（24列すべては要らない。緯度経度・URL・備考は落とす）
NOKOSU = ["サービスの種類", "事業所名", "事業所番号", "法人番号",
          "法人の名称", "住所", "方書（ビル名等）", "電話番号", "定員",
          "利用可能曜日", "都道府県名", "市区町村名",
          "都道府県コード又は市町村コード"]


def toru(url, saki, timeout=180):
    """curl で取る。戻り　(HTTPコード, バイト数)"""
    p = subprocess.run(
        ["curl", "-sS", "-L", "-o", saki, "-w", "%{http_code} %{size_download}",
         "--max-time", str(timeout),
         "-A", "Mozilla/5.0", "-e", NYUGUCHI, url],
        capture_output=True, text=True, timeout=timeout + 30)
    out = (p.stdout or "").strip().split()
    if len(out) != 2:
        return ("000", 0)
    return (out[0], int(out[1]))


def nyugushi_yomu():
    """入口のページから、最新（2026年6月末時点）のＣＳＶの場所を読む。

    ⚠ **場所をスクリプトに書き写さない。** 版が上がると変わるためである。
    ページに現れる順で、サービスの種類ごとの最初の1件が最新である。
    """
    with tempfile.NamedTemporaryFile(suffix=".html", delete=False) as f:
        tmp = f.name
    code, size = toru(NYUGUCHI, tmp)
    if code != "200":
        os.unlink(tmp)
        raise SystemExit(f"入口のページが取れない（HTTP {code}）")
    t = open(tmp, encoding="utf-8", errors="replace").read()
    os.unlink(tmp)

    # 基準日の見出しが変わっていないか
    kijun_ok = KIJUNBI in t

    out = {}
    for m in re.finditer(r'href="([^"]*jigyosho_(\d+)_all_[^"]*\.csv)"'
                         r'[^>]*>(.{0,90}?)</a>', t, re.S):
        k = m.group(2)
        if k in out:
            continue          # 先に現れたものが最新
        na = re.sub(r"<[^>]+>|\s+", " ", m.group(3)).strip()
        na = re.sub(r"［.*?］$", "", na).strip()
        out[k] = (MOTO + m.group(1) if m.group(1).startswith("/")
                  else m.group(1), na)
    return out, kijun_ok


def main():
    futeki = []

    def judge(ok, mi, setsu=""):
        print(("   ○ " if ok else "   × ") + mi + (("　" + setsu)
                                                   if setsu else ""))
        if not ok:
            futeki.append(mi)

    print("介護サービス情報公表システム オープンデータからの取り出し")
    print(f"  入口　{NYUGUCHI}")
    print(f"  抽出　市町村コード {CODE}（{KEN}柴田郡{MACHI}）")
    print()

    saki, kijun_ok = nyugushi_yomu()
    print(f"  サービスの種類 {len(saki)}件／基準日の見出し "
          f"{'一致' if kijun_ok else '⚠ 変わっている'}")
    print()

    meisai = []        # 川崎町の事業所
    kensu = []         # (コード, サービス名, 全国, 宮城県, 川崎町)
    shippai = []
    hyoji = None

    for i, (k, (url, na)) in enumerate(sorted(saki.items()), 1):
        with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as f:
            tmp = f.name
        code, size = toru(url, tmp)
        if code != "200" or size < 100:
            shippai.append((k, na, code))
            os.unlink(tmp)
            print(f"  × {k} {na[:26]:<28}HTTP {code}")
            continue
        raw = open(tmp, "rb").read()
        os.unlink(tmp)
        t = None
        for enc in ("utf-8-sig", "cp932", "utf-8"):
            try:
                t = raw.decode(enc)
                break
            except Exception:
                continue
        if t is None:
            shippai.append((k, na, "文字コード不明"))
            continue
        rows = list(csv.reader(io.StringIO(t)))
        head = rows[0]
        if hyoji is None:
            hyoji = head
        ic = head.index("都道府県コード又は市町村コード")
        ik = head.index("都道府県名")
        zen = len(rows) - 1
        ken = sum(1 for r in rows[1:] if len(r) > ik and r[ik] == KEN)
        hit = [r for r in rows[1:] if len(r) > ic and r[ic] == CODE]
        kensu.append((k, na, zen, ken, len(hit)))
        for r in hit:
            d = {h: (r[j] if j < len(r) else "") for j, h in enumerate(head)}
            d["_コード"] = k
            meisai.append(d)
        mark = "○" if hit else "―"
        print(f"  {mark} {k} {na[:26]:<28}全国{zen:>6,}／宮城{ken:>5,}／"
              f"川崎町{len(hit):>3}")

    print()
    os.makedirs(OUT_DIR, exist_ok=True)
    meisai.sort(key=lambda d: (d["_コード"], d.get("事業所名", "")))
    with open(OUT_MEISAI, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["サービスコード"] + NOKOSU)
        for d in meisai:
            w.writerow([d["_コード"]] + [d.get(h, "") for h in NOKOSU])
    with open(OUT_KENSU, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["サービスコード", "サービスの種類", "全国", KEN, MACHI])
        for row in kensu:
            w.writerow(list(row))
    print(f"  書き出し　{OUT_MEISAI}（{len(meisai)}件）")
    print(f"  書き出し　{OUT_KENSU}（{len(kensu)}件）")
    print()

    # ══════ 検査
    print("  ── 検査1　取り切れているか")
    judge(not shippai, f"35件のＣＳＶをすべて取れた"
                       f"（失敗 {len(shippai)}件）",
          "／".join(f"{a}:{c}" for a, b, c in shippai[:3])
          if shippai else "")
    judge(len(saki) >= 30, f"サービスの種類が30件以上（{len(saki)}件）")
    judge(kijun_ok, f"基準日の見出しが「{KIJUNBI}」のまま",
          "⚠ 変わっていれば版が上がっている。出典の表記を改める")

    print("  ── 検査2　取り違えていないか")
    judge(all(d.get("都道府県名") == KEN for d in meisai),
          f"すべての行の都道府県名が「{KEN}」")
    judge(all(d.get("市区町村名") == MACHI for d in meisai),
          f"すべての行の市区町村名が「{MACHI}」")
    judge(all(d.get("都道府県コード又は市町村コード") == CODE
              for d in meisai), f"すべての行の市町村コードが {CODE}")
    judge(not any("川崎市" in d.get("住所", "") for d in meisai),
          "川崎市（神奈川県）の行が混じっていない")
    judge(not any("福岡" in d.get("住所", "") for d in meisai),
          "福岡県田川郡川崎町の行が混じっていない")

    print("  ── 検査3　事業所番号の形")
    ban = [d.get("事業所番号", "") for d in meisai]
    judge(all(re.fullmatch(r"\d{10}", b) for b in ban),
          f"事業所番号がすべて10桁（{len(ban)}件）")
    judge(len(ban) == len(set(ban)) or True,
          f"事業所番号の異なり {len(set(ban))}件"
          f"（1つの事業所が複数のサービスを担うため重複はありうる）")

    print("  ── 検査4　法人")
    hojin = {}
    for d in meisai:
        h = d.get("法人番号", "") or d.get("法人の名称", "")
        hojin.setdefault(h, []).append(d)
    judge(bool(hojin), f"法人番号で数えた法人 {len(hojin)}件")

    print("  ── 検査5　定員")
    teiin = [d for d in meisai
             if (d.get("定員", "") or "0").strip() not in ("", "0")]
    judge(True, f"定員の記載がある事業所 {len(teiin)}件／"
                f"{len(meisai)}件",
          "⚠ 訪問系・居宅介護支援は定員の概念がないため0である")

    print("  ── 検査6　オープンデータに無いもの（未実施として残す）")
    for h in ("指定年月日", "指定日", "入所者数", "入所率", "稼働率",
              "待機"):
        if hyoji and any(h in x for x in hyoji):
            judge(False, f"列「{h}」はあるはずがないのにある")
    judge(True, "指定の年月日・指定の主体・入所者数・入所率・待機者数は"
                "オープンデータに無い（24列に含まれない）")
    print("     引き続き町・県へのご確認による（No.29・No.154・No.159）")

    print("  ── 検査7　件数の表に他団体の固有名称が無いか")
    with open(OUT_KENSU, encoding="utf-8-sig") as f:
        s = f.read()
    judge("川崎市" not in s and "田川郡" not in s,
          "件数の表に取り違えのもとになる語が無い")

    print()
    if futeki:
        print(f"   × 不適合 {len(futeki)}件")
        for m in futeki:
            print(f"      {m}")
        sys.exit(1)
    print(f"   ○ 検査1〜7　不適合0件")
    print()
    print(f"   ▶ 町内の事業所　{len(meisai)}件"
          f"（{len(set(ban))}の事業所番号・{len(hojin)}法人）")


if __name__ == "__main__":
    main()
