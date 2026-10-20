# -*- coding: utf-8 -*-
"""環境設定で許可すべきドメインの候補と、いまの疎通の確認（令和8年10月20日）

  python3 10_内部資料_非成果品/check_setsuzoku_zenan_R8.10.20.py

⚠ **これは当方の内部資料であり、町へお出しする成果品ではない。**
  複数の案件（川崎町・北塩原村ほか）のドメインを一覧にするため、
  他団体の固有名称を含む。だから 10_内部資料_非成果品/ に置いている。
  川崎町だけの疎通の確認は 07_ソーススクリプト/check_setsuzoku_R8.10.19.py。

なぜ要るか
  クラウドの作業環境は**外向きの通信が既定で遮断**されており、
  ホストごとに許可を足す必要がある。**許可は環境ごとに別である。**
  現に、北塩原村の環境では福島県・北塩原村のサイトが届かないのに、
  川崎町の環境では宮城県・川崎町のサイトが届く。
  どのホストを足すべきかを、**用途と効きどころつきで**一覧にする。

  終了コード　0＝どの案件でも要るもの（必須）がすべて通る／1＝1つでも通らない

⚠ **「Allow package managers」のチェックは外さないこと。**
  外すと pypi・npm が通らなくなり、openpyxl・python-docx・matplotlib の
  導入ができなくなる。

⚠ **地域包括ケア「見える化」システム（mieruka.mhlw.go.jp）は挙げない。**
  ＩＤ・パスワードはお預かりしているが当方で直接データを触る想定ではなく、
  見える化対応は町・村が進めるためである（CLAUDE.md 3）。
"""
import subprocess
import sys

# （区分, 用途, ホスト, 必須か, 何に効くか）
#   必須＝どの案件（高齢者・介護保険・障がい）でも要るもの
SAKI = [
    # ── 国（どの案件でも要る）
    ("国", "厚生労働省", "www.mhlw.go.jp", True,
     "基本指針・全国課長会議資料・交付金の評価指標・各種オープンデータ。"
     "⚠ **介護サービス情報公表システムの事業所一覧（ＣＳＶ）もここにある**"),
    ("国", "総務省統計局", "www.stat.go.jp", True,
     "国勢調査・人口推計・日本の統計（死亡率ほか）"),
    ("国", "政府統計の総合窓口（e-Stat）", "www.e-stat.go.jp", True,
     "介護保険事業状況報告・人口・国勢調査の統計表"),
    ("国", "総務省", "www.soumu.go.jp", True,
     "住民基本台帳に基づく人口（年齢階級別・市区町村別）・過疎地域の指定"),
    ("国", "法令検索（e-Gov）", "laws.e-gov.go.jp", True,
     "条文の原文。資料編の条文抜粋に要る"),
    ("国", "介護サービス情報公表システム", "www.kaigokensaku.mhlw.go.jp",
     True,
     "⚠ **画面は自動のアクセスを受け付けない。**"
     "事業所の一覧は上のオープンデータ（www.mhlw.go.jp）から取る"),

    # ── 国（案件により要る）
    ("国", "福祉医療機構（ＷＡＭ ＮＥＴ）", "www.wam.go.jp", False,
     "介護事業所情報の補助的な確認・老健事業の報告書"),
    ("国", "こども家庭庁", "www.cfa.go.jp", False,
     "障がい児福祉計画・子ども子育て支援事業計画"),
    ("国", "国土交通省", "www.mlit.go.jp", False,
     "住生活基本計画・高齢者の住まい"),
    ("国", "サービス付き高齢者向け住宅情報提供システム",
     "www.satsuki-jutaku.mlit.go.jp", False,
     "⚠ **www.satsuki-jutaku.jp は 301 でこのホストへ移る。**"
     "移転先を許可しないと届かない。サ高住の登録状況（基本指針 別表 八）"),
    ("国", "国立社会保障・人口問題研究所", "www.ipss.go.jp", False,
     "⚠ **日本の地域別将来推計人口（令和5年推計）の市区町村・5歳階級別。**"
     "認知症の人の数の将来の推計に直に効く"),
    ("国", "文部科学省", "www.mext.go.jp", False, "生涯学習・学校との連携"),
    ("国", "国立精神・神経医療研究センター", "www.ncnp.go.jp", False,
     "障がい（精神）の統計・調査"),

    # ── 自治体（案件ごとに足す）
    ("自治体", "宮城県", "www.pref.miyagi.jp", False,
     "［川崎町］指定介護サービス事業所の一覧・有料老人ホームの届出状況"),
    ("自治体", "川崎町", "www.town.kawasaki.miyagi.jp", False,
     "［川崎町］町の計画・例規・過疎地域持続的発展計画"),
    ("自治体", "福島県", "www.pref.fukushima.lg.jp", False,
     "［北塩原村］県の計画・指定事業所の一覧・特定健診の受診率"),
    ("自治体", "北塩原村", "www.vill.kitashiobara.fukushima.jp", False,
     "［北塩原村］村の計画・例規・広報"),
    ("自治体", "岩手県", "www.pref.iwate.jp", False, "［金ケ崎町］"),
    ("自治体", "金ケ崎町", "www.town.kanegasaki.iwate.jp", False,
     "［金ケ崎町］"),

    # ── その他
    ("その他", "三菱ＵＦＪリサーチ＆コンサルティング", "www.murc.jp", False,
     "老人保健健康増進等事業の報告書（国の委託調査の本体）"),
    ("その他", "九州大学 衛生・公衆衛生学",
     "www.eph.med.kyushu-u.ac.jp", False,
     "認知症の有病率の研究報告書の本体（国の資料は要約）"),
    ("その他", "障害者情報ネットワーク（ノーマネット）",
     "www.dinf.ne.jp", False, "障がい関係の資料"),
]

# 挙げないもの
IWANAI = [
    ("地域包括ケア「見える化」システム", "mieruka.mhlw.go.jp",
     "⚠ **挙げない。** 当方で直接データを触る想定ではない（CLAUDE.md 3）。"
     "見える化対応は町・村が進める"),
    ("e-Stat の API", "api.e-stat.go.jp",
     "アプリケーションＩＤの取得が要る。画面（www.e-stat.go.jp）で足りる"),
]


def tameshi(host, timeout=12):
    """HTTPS の疎通を試す。戻り　(コード, 説明, リダイレクト先)"""
    try:
        p = subprocess.run(
            ["curl", "-sS", "-o", "/dev/null", "-D", "-",
             "-w", "\n%{http_code}", "--max-time", str(timeout),
             f"https://{host}/"],
            capture_output=True, text=True, timeout=timeout + 10)
        out = (p.stdout or "")
        code = out.strip().splitlines()[-1] if out.strip() else ""
        saki = ""
        for ln in out.splitlines():
            if ln.lower().startswith("location:"):
                saki = ln.split(":", 1)[1].strip()
        err = (p.stderr or "").strip()
        if code == "000" or not code.isdigit():
            if "403" in err:
                return ("×", "許可されていない（プロキシが拒む）", "")
            if "resolve" in err:
                return ("×", "⚠ 名前解決できない（ドメインが違うおそれ）", "")
            if "timed out" in err.lower():
                return ("×", "応答なし", "")
            return ("×", err.splitlines()[-1][:48] if err else "通らない", "")
        return (code, "通る" if code.startswith(("2", "3")) else
                "接続はできたが応答がこの状態", saki)
    except subprocess.TimeoutExpired:
        return ("×", "応答なし", "")
    except FileNotFoundError:
        return ("—", "curl がない（未実施）", "")


def main():
    print("環境設定で許可すべきドメインの候補と、いまの疎通（令和8年10月20日）")
    print("  ⚠ **この結果は、いま動いている環境（川崎町）のものである。**")
    print("    許可は環境ごとに別であり、他の案件の環境では違う。")
    print()

    kekka = []
    ima = None
    for kubun, yoto, host, hissu, nani in SAKI:
        if kubun != ima:
            print(f"  ══════ {kubun}")
            ima = kubun
        code, setsu, saki = tameshi(host)
        tsuru = code.isdigit() and code.startswith(("2", "3"))
        kekka.append((kubun, yoto, host, hissu, tsuru, code, setsu, nani))
        mark = "○" if tsuru else "×"
        print(f"   {mark} {yoto[:30]:<30}{'必須' if hissu else '任意':<5}"
              f"{code:<6}{setsu}")
        print(f"       {host}")
        if saki:
            print(f"       → 移る先 {saki}")
        print(f"       {nani}")

    print()
    print("  ══════ 挙げないもの")
    for yoto, host, riyu in IWANAI:
        print(f"   － {yoto}（{host}）")
        print(f"       {riyu}")

    print()
    print("  ── まとめ")
    tsuru = [x for x in kekka if x[4]]
    hissu_ng = [x for x in kekka if x[3] and not x[4]]
    nin_ng = [x for x in kekka if not x[3] and not x[4]]
    print(f"   通る {len(tsuru)}件／通らない {len(kekka) - len(tsuru)}件"
          f"（うち必須 {len(hissu_ng)}件）")
    if nin_ng:
        print()
        print("   △ 足していただくと効くもの（任意）")
        for x in nin_ng:
            print(f"      {x[2]:<42}{x[1]}")
    print()
    print("   ⚠ **「Allow package managers」のチェックは外さないでください。**")
    print("      外すと pypi・npm が通らなくなり、openpyxl・python-docx・")
    print("      matplotlib の導入ができません。")

    if hissu_ng:
        print()
        print("   × **必須の先が通りません。**次を許可してください。")
        for x in hissu_ng:
            print(f"      {x[2]}")
        sys.exit(1)
    print()
    print("   ○ **必須の先はすべて通ります。**")


if __name__ == "__main__":
    main()
