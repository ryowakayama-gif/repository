# -*- coding: utf-8 -*-
"""
未確認先の計画ページからPDFを一括取得する

この環境では自治体サイトへの直接アクセスが組織のegressポリシーで遮断されて
いる（全ホストがCONNECTに403）。そのため本スクリプトは、制限のない環境で
実行することを前提にしている。

使い方:
  python3 fetch_plan_pdfs.py                 # 既定の出力先 ./plan_pdfs へ取得
  python3 fetch_plan_pdfs.py --out D:/pdfs   # 出力先を指定
  python3 fetch_plan_pdfs.py --only 函館市 帯広市

取得後:
  python3 batch_extract_plan_periods.py <出力先>
  で計画期間の一覧CSVが出る。

TARGETS は、これまでの調査で確認できた計画掲載ページ。
PDFへの直リンクではなく掲載ページを指しているため、ページ内のPDFリンクを
辿って取得する。リンク構成が変わっている場合はページを開いて手で落とすこと。
"""

import argparse
import os
import re
import sys
import urllib.parse
import urllib.request

# (自治体名, 優先度, 計画名・状況, 計画掲載ページ)
TARGETS = [
    ("函館市", "最優先", "環境基本計画［第3次計画］ 令和2年3月策定・終期未確認（中核市）",
     "https://www.city.hakodate.hokkaido.jp/docs/2020033100148/"),
    ("青森市", "最優先", "環境基本計画の有無から要確認（中核市）。計画一覧から探す",
     "https://www.city.aomori.aomori.jp/shisei/machizukuri/1005734/"),
    ("北海道", "優先", "北海道環境基本計画［第3次計画］ 令和3年3月策定・終期未確認",
     "https://www.pref.hokkaido.lg.jp/ks/ksk/kihonkeikaku.html"),
    ("秋田県", "優先", "第3次秋田県環境基本計画 令和3年3月策定・終期未確認",
     "https://www.pref.akita.lg.jp/pages/genre/13255"),
    ("福島県", "優先", "福島県環境基本計画（第5次）・終期未確認",
     "https://www.pref.fukushima.lg.jp/sec/16005a/5th-kankyoukihonkeikaku.html"),
    ("帯広市", "中", "第三期帯広市環境基本計画 令和2年3月策定・終期未確認",
     "https://www.city.obihiro.hokkaido.jp/kurashi/kankyo/kankobutsu/1003775.html"),
    ("北広島市", "中", "第3次北広島市環境基本計画 令和3年度〜・終期未確認",
     "https://www.city.kitahiroshima.hokkaido.jp/hotnews/detail/00140680.html"),
    ("北見市", "中", "第2次北見市環境基本計画（改定版）・期間未確認",
     "https://www.city.kitami.lg.jp/detail.php?content=4642"),
    ("大崎市", "中", "第2次大崎市環境基本計画 令和2年3月策定・終期未確認",
     "https://www.city.osaki.miyagi.jp/shisei/soshikikarasagasu/shiminkyodousuishimbu/kankyohozenka/3/3329.html"),
    ("鶴岡市", "中", "第2次鶴岡市環境基本計画・期間未確認",
     "https://www.city.tsuruoka.lg.jp/seibi/kankyo/kihonkeikaku/dainizi_kankyoukihon.html"),
    ("小樽市", "中", "環境基本計画の有無・期間とも未確認。計画一覧から探す",
     "https://www.city.otaru.lg.jp/categories/bunya/keikaku/"),
    ("十和田市", "中", "環境基本計画の有無から要確認。各種計画等一覧から探す",
     "https://www.city.towada.lg.jp/shisei/other/2024-1220-1521-66.html"),
    ("上士幌町", "中", "第2次上士幌町環境基本計画 令和6年度策定・終期未確認",
     "https://www.kamishihoro.jp/page/00000416"),
    ("新得町", "中", "環境基本計画・温対・適応計画の一体掲載・期間未確認",
     "https://www.shintoku-town.jp/gyousei/koukai_kouhyou/tyoumin/kankyoukihonkeikaku/"),
    ("池田町", "中", "池田町環境基本計画（第3次改訂）令和8年度改訂・終期未確認",
     "https://www.town.hokkaido-ikeda.lg.jp/kurashi/kankyo/430.html"),
    ("厚岸町", "中", "第2期厚岸町豊かな環境を守り育てる基本計画・期間未確認",
     "https://www.akkeshi-town.jp/gyosei/seisaku/kankyo/kankyo19/"),
    ("鶴居村", "中", "第2次鶴居村環境基本計画・期間未確認",
     "https://www.vill.tsurui.lg.jp/soshikikarasagasu/kikakuzaiseika/muranokeikaku/804.html"),
    ("弟子屈町", "低", "外部DBで計画名のみ確認。公式ページ・期間とも未確認",
     "https://www.town.teshikaga.hokkaido.jp/"),
]

UA = "Mozilla/5.0 (compatible; plan-period-collector/1.0)"
PDF_RE = re.compile(r'href=["\']([^"\']+?\.pdf[^"\']*)["\']', re.IGNORECASE)


def fetch(url, timeout=30):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=timeout) as res:
        return res.read()


def collect(name, page_url, out_dir, limit):
    try:
        html = fetch(page_url).decode("utf-8", errors="replace")
    except Exception as exc:
        return f"ページ取得に失敗: {type(exc).__name__} {exc}"
    links = []
    for href in PDF_RE.findall(html):
        full = urllib.parse.urljoin(page_url, href)
        if full not in links:
            links.append(full)
    if not links:
        return "ページ内にPDFリンクが見つかりません（手動で確認してください）"
    saved = 0
    for i, link in enumerate(links[:limit], 1):
        suffix = "" if len(links[:limit]) == 1 else f"_{i}"
        dst = os.path.join(out_dir, f"{name}_環境基本計画{suffix}.pdf")
        try:
            data = fetch(link, timeout=90)
        except Exception as exc:
            print(f"      ! {link} 取得失敗: {type(exc).__name__}")
            continue
        if not data.startswith(b"%PDF"):
            continue
        with open(dst, "wb") as f:
            f.write(data)
        saved += 1
        print(f"      保存: {os.path.basename(dst)}（{len(data)/1024:.0f}KB）")
    return f"{saved}件保存" if saved else "PDFを保存できませんでした"


def main():
    ap = argparse.ArgumentParser(description="未確認先の計画ページからPDFを一括取得する")
    ap.add_argument("--out", default="plan_pdfs", help="出力フォルダ（既定 plan_pdfs）")
    ap.add_argument("--limit", type=int, default=3, help="1団体あたり取得するPDFの上限（既定3）")
    ap.add_argument("--only", nargs="*", help="対象の自治体名を絞る")
    ap.add_argument("--list", action="store_true", help="対象一覧を表示するだけ")
    args = ap.parse_args()

    targets = TARGETS
    if args.only:
        targets = [t for t in TARGETS if t[0] in args.only]
        if not targets:
            print("指定した自治体が TARGETS にありません。--list で一覧を確認してください。")
            sys.exit(1)

    if args.list:
        print(f"{'自治体':10} {'優先度':6} 状況 / 掲載ページ")
        for name, pri, note, url in targets:
            print(f"{name:10} {pri:6} {note}\n{'':17} {url}")
        return

    os.makedirs(args.out, exist_ok=True)
    print(f"{len(targets)}件を取得します -> {args.out}\n")
    results = []
    for name, pri, note, url in targets:
        print(f"  [{pri}] {name}  {url}")
        msg = collect(name, url, args.out, args.limit)
        print(f"      {msg}")
        results.append((name, msg))
    ng = [n for n, m in results if "保存" not in m or m.startswith("0件")]
    print(f"\n完了。取得できなかった団体: {len(ng)}件" + (f" -> {'、'.join(ng)}" if ng else ""))
    print(f"次: python3 batch_extract_plan_periods.py {args.out}")


if __name__ == "__main__":
    main()
