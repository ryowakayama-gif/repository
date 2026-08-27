# -*- coding: utf-8 -*-
"""
新人教育カリキュラムの一覧ページ（Artifact 用 HTML）を生成する。
データは build_curriculum.py の定義をそのまま読み込むため、Excel と内容が乖離しない。
"""
import html
import os

from build_curriculum import (
    LEVELS, LEVEL_TARGET, ROADMAP, SKILLMAP, RUBRIC, CALENDAR, THEMES,
    KOKAIKEI, KOEIKIGYO, KOREISHA, SHOGAI, CHIIKIFUKUSHI, KODOMO, SONOTA, KYOTSU,
)
from curriculum_sales import SALES_LEVELS, SALES_MATRIX, SALES_KIT, SALES_PREP

OUT = "/tmp/claude-0/-home-user-repository/9fe4feb1-e025-51cd-8697-f63e2951be22/scratchpad/curriculum.html"
e = html.escape

THEME_SLUG = {
    "04_公会計": ("kokaikei", "公会計", "財務書類・台帳精緻化・所有外管理資産"),
    "05_公営企業会計": ("koei", "公営企業会計", "会計支援・経営戦略・料金改定・経営分析"),
    "06_高齢者・介護保険": ("korei", "高齢者福祉・介護保険事業計画", "3年1期／保険料算定まで"),
    "07_障害者・障害児": ("shogai", "障害者・障害児計画", "3計画一体／成果目標と見込量"),
    "08_地域福祉": ("chiiki", "地域福祉計画", "上位計画／重層的支援体制"),
    "09_こども・子育て": ("kodomo", "こども計画・子育て支援事業計画", "5年1期／量の見込みと確保方策"),
    "10_その他計画": ("sonota", "その他の行政計画", "食育・健康・農業・男女・財政・公共施設"),
    "11_共通スキル": ("kyotsu", "共通スキル", "調査・執筆・会議・PM・調達・積算"),
}


def phase_key(p):
    return "ⅠⅡⅢⅣⅤⅥ".index(p.strip()[0]) + 1 if p and p.strip()[0] in "ⅠⅡⅢⅣⅤⅥ" else 0


def lv_chip(lv):
    lv = (lv or "").strip()
    end = lv.split("→")[-1].strip()
    n = end[1:] if end.startswith("L") else ""
    return f'<span class="lv lv{e(n)}">{e(lv)}</span>'


# ---------- 各セクション ----------
def sec_levels():
    cards = []
    for code, name, when, define, role, resp, check in LEVELS:
        n = code[1:]
        cards.append(f"""
        <li class="rung rung-{n}">
          <div class="rung-head">
            <span class="rung-code">{e(code)}</span>
            <span class="rung-name">{e(name)}</span>
            <span class="rung-when">{e(when)}</span>
          </div>
          <p class="rung-def">{e(define)}</p>
          <dl class="rung-meta">
            <div><dt>担える役割</dt><dd>{e(role)}</dd></div>
            <div><dt>責任範囲</dt><dd>{e(resp)}</dd></div>
            <div><dt>確認方法</dt><dd>{e(check)}</dd></div>
          </dl>
        </li>""")
    targets = "".join(
        f'<li><span class="tg-when">{e(a)}</span><span class="tg-lv">{e(b)}</span>'
        f'<span class="tg-desc">{e(c)}</span></li>' for a, b, c in LEVEL_TARGET)
    return f"""
    <section id="levels" class="sec">
      <header class="sec-head">
        <p class="eyebrow">01　階層</p>
        <h2>「知っている」ではなく「できた実績」で判定する</h2>
        <p class="lead">全テーマ共通の6段階。判定には必ず具体的な案件名・成果物名を紐づける。同種の作業を2案件以上、単独で完遂して初めてL3とする。</p>
      </header>
      <ol class="ladder">{''.join(cards)}</ol>
      <div class="targets">
        <h3>年次ごとの到達目標</h3>
        <ul class="tg-list">{targets}</ul>
      </div>
    </section>"""


def sec_roadmap():
    years = {1: [], 2: [], 3: []}
    for row in ROADMAP:
        mo, when, phase, focus, learn, goal, ojt, lv, mile = row
        y = 1 if mo <= 12 else (2 if mo <= 24 else 3)
        is_mile = mile.startswith("【")
        goals = [g.strip() for g in goal.replace("①", "\n①").replace("②", "\n②").replace("③", "\n③").split("\n") if g.strip()]
        goal_html = "".join(f"<li>{e(g)}</li>" for g in goals)
        years[y].append(f"""
        <li class="mrow{' is-mile' if is_mile else ''}" data-phase="{phase_key(phase)}">
          <div class="m-when">
            <span class="m-no">{mo}</span>
            <span class="m-mo">{e(when)}</span>
          </div>
          <div class="m-body">
            <div class="m-top">
              <span class="m-phase p{phase_key(phase)}">{e(phase)}</span>
              <span class="m-focus">{e(focus)}</span>
              {lv_chip(lv)}
            </div>
            <p class="m-learn"><span class="k">学ぶ</span>{e(learn)}</p>
            <p class="k k-block">できるようになる</p>
            <ul class="m-goal">{goal_html}</ul>
            <div class="m-foot">
              <p><span class="k">OJT</span>{e(ojt)}</p>
              <p class="m-mile"><span class="k">確認</span>{e(mile)}</p>
            </div>
          </div>
        </li>""")
    panels = ""
    labels = {1: ("1年目", "基礎導入 → 実務補助 → 部分自立"),
              2: ("2年目", "単独遂行 ── 案件を一人で回す"),
              3: ("3年目", "応用・指導 ── 複数案件と後輩育成")}
    for y in (1, 2, 3):
        t, s = labels[y]
        panels += f"""
      <div class="ypanel" id="y{y}" role="tabpanel" aria-labelledby="tab{y}"{'' if y == 1 else ' hidden'}>
        <p class="ypanel-sub">{e(s)}</p>
        <ol class="months">{''.join(years[y])}</ol>
      </div>"""
    tabs = "".join(
        f'<button class="ytab{" is-on" if y == 1 else ""}" id="tab{y}" role="tab" '
        f'aria-selected="{"true" if y == 1 else "false"}" aria-controls="y{y}" data-y="{y}">'
        f'{labels[y][0]}<span>{(y-1)*12+1}–{y*12}ヶ月</span></button>' for y in (1, 2, 3))
    return f"""
    <section id="roadmap" class="sec">
      <header class="sec-head">
        <p class="eyebrow">02　月次ロードマップ</p>
        <h2>入社1ヶ月目から36ヶ月目まで</h2>
        <p class="lead">4月入社を前提に、自治体の年度サイクル（予算要求期＝8〜10月／公告期＝1〜2月／納品期＝12〜3月）へ学習を重ねている。テーマ固有の詳細は下の到達基準と、配布する管理表を参照。</p>
      </header>
      <div class="ytabs" role="tablist" aria-label="育成年次">{tabs}</div>
      {panels}
    </section>"""


THEME_ROWS = {
    "04_公会計": KOKAIKEI, "05_公営企業会計": KOEIKIGYO, "06_高齢者・介護保険": KOREISHA,
    "07_障害者・障害児": SHOGAI, "08_地域福祉": CHIIKIFUKUSHI, "09_こども・子育て": KODOMO,
    "10_その他計画": SONOTA, "11_共通スキル": KYOTSU,
}


def sec_themes():
    smap = {}
    for theme, area, *lv in SKILLMAP:
        smap.setdefault(theme, []).append((area, lv))
    order = list(smap.keys())
    blocks = []
    for i, (sheet, title, sub, color, rows) in enumerate(THEMES):
        slug, short, tag = THEME_SLUG[sheet]
        theme_key = order[i] if i < len(order) else None
        areas = smap.get(theme_key, [])
        n_items = len(rows)
        sales = sum(1 for r in rows if "営業" in str(r[2]))   # r[2] = 区分
        grid = ""
        for area, lv in areas:
            cells = "".join(
                f'<td class="g{j+1}">{e(t)}</td>' for j, t in enumerate(lv))
            grid += f'<tr><th scope="row">{e(area)}</th>{cells}</tr>'
        # 代表的な学習項目（大項目）
        heads, seen = [], set()
        for r in rows:
            h = str(r[3])
            if h not in seen:
                seen.add(h)
                heads.append(h)
        chips = "".join(f"<li>{e(h)}</li>" for h in heads[:9])
        blocks.append(f"""
      <article class="theme" id="{slug}" style="--tc:#{color}">
        <header class="theme-head">
          <span class="theme-dot" aria-hidden="true"></span>
          <div>
            <h3>{e(short)}</h3>
            <p class="theme-tag">{e(tag)}</p>
          </div>
          <p class="theme-count"><b>{n_items}</b><span>学習項目<br>うち営業 {sales}</span></p>
        </header>
        <p class="theme-sub">{e(sub)}</p>
        <ul class="theme-chips">{chips}</ul>
        <div class="tablewrap">
          <table class="grid">
            <caption class="vh">{e(short)}の階層別到達基準</caption>
            <thead><tr><th scope="col">領域</th><th scope="col">L1 理解する</th><th scope="col">L2 補助する</th><th scope="col">L3 自立する</th><th scope="col">L4 応用・指導</th><th scope="col">L5 統括する</th></tr></thead>
            <tbody>{grid}</tbody>
          </table>
        </div>
      </article>""")
    return f"""
    <section id="themes" class="sec">
      <header class="sec-head">
        <p class="eyebrow">03　テーマ別 到達基準</p>
        <h2>8テーマ × 領域 × 階層</h2>
        <p class="lead">主担当テーマは入社3ヶ月目に決定し、2年目に副担当、3年目に第3テーマへ広げる。営業と業務は別々に判定する ── 業務がL3でも営業がL1なら、案件は取れても回せない、あるいはその逆になる。</p>
      </header>
      <div class="themes">{''.join(blocks)}</div>
    </section>"""


KIT_THEMES = ["公会計", "公営企業会計", "高齢者・介護保険", "障害者・障害児",
              "地域福祉", "こども・子育て", "その他計画", "共通"]
KIT_SLUG = {t: f"k{i}" for i, t in enumerate(KIT_THEMES)}


def lv_from(text):
    t = (text or "").strip()
    for n in "12345":
        if t.startswith("L" + n):
            return n
    return "0"


def sec_sales():
    # レベル定義
    rungs = ""
    for code, kata, konkyo, out, aite, seika, ng, when in SALES_LEVELS:
        rungs += f"""
        <li class="srung rung-{code[1:]}">
          <div class="srung-head">
            <span class="rung-code">{e(code)}</span>
            <span class="rung-name">{e(kata)}</span>
            <span class="rung-when">{e(when)}</span>
          </div>
          <dl class="srung-grid">
            <div><dt>根拠にするもの</dt><dd>{e(konkyo)}</dd></div>
            <div><dt>出してよい提案物</dt><dd>{e(out)}</dd></div>
            <div><dt>会う相手</dt><dd>{e(aite)}</dd></div>
            <div><dt>成果の目安</dt><dd>{e(seika)}</dd></div>
            <div class="ng"><dt>やってはいけない</dt><dd>{e(ng)}</dd></div>
          </dl>
        </li>"""

    # 深度マトリクス
    mrows = ""
    for axis, l1, l2, l3, l4, l5, note in SALES_MATRIX:
        cells = "".join(f'<td class="g{i}">{e(t)}</td>' for i, t in enumerate((l1, l2, l3, l4, l5), 1))
        mrows += (f'<tr><th scope="row">{e(axis)}</th>{cells}'
                  f'<td class="cav">{e(note)}</td></tr>')

    # 提案ネタ台帳
    krows = ""
    for theme, cat, name, use, lv, src in SALES_KIT:
        fin = " is-fin" if cat == "財源" else ""
        krows += (f'<tr class="krow{fin}" data-t="{KIT_SLUG.get(theme, "")}">'
                  f'<td class="kt">{e(theme)}</td>'
                  f'<td><span class="cat" data-c="{e(cat)}">{e(cat)}</span></td>'
                  f'<td class="kn">{e(name)}</td>'
                  f'<td>{e(use)}</td>'
                  f'<td class="kl"><span class="lv lv{lv_from(lv)}">{e(lv)}</span></td>'
                  f'<td class="ks">{e(src)}</td></tr>')
    chips = '<button class="kchip is-on" data-f="all">すべて<span>' + str(len(SALES_KIT)) + '</span></button>'
    for t in KIT_THEMES:
        n = sum(1 for r in SALES_KIT if r[0] == t)
        chips += f'<button class="kchip" data-f="{KIT_SLUG[t]}">{e(t)}<span>{n}</span></button>'

    # 訪問前準備
    prep = {}
    for lv, item, todo, done in SALES_PREP:
        prep.setdefault(lv, []).append((item, todo, done))
    pblocks = ""
    for lv in ("L1", "L2", "L3", "L4", "L5"):
        if lv not in prep:
            continue
        lis = "".join(
            f'<li><span class="pi">{e(a)}</span><p class="pt">{e(b)}</p>'
            f'<p class="pd"><span class="k">完了の基準</span>{e(c)}</p></li>'
            for a, b, c in prep[lv])
        pblocks += (f'<div class="pgroup"><span class="lv lv{lv[1]}">{lv}</span>'
                    f'<ul class="plist">{lis}</ul></div>')

    fin_n = sum(1 for r in SALES_KIT if r[1] == "財源")
    return f"""
    <section id="sales" class="sec">
      <header class="sec-head">
        <p class="eyebrow">04　営業提案の深度</p>
        <h2>同じテーマでも、レベルによって話してよい深さが違う</h2>
        <p class="lead">背伸びした提案は、失注より重い事故を招く ── 財源の誤案内、他社を排除する仕様書案、未公表情報の持ち込み。
        提案の材料を9つに分解し、レベルごとに「どこまで使うか」を定めた。</p>
      </header>

      <h3 class="sub-h">提案の型（L1〜L5）</h3>
      <ol class="ladder srungs">{rungs}</ol>

      <div class="warn">
        <p><b>財源に関する原則</b>　補助金・地方債・交付税措置は、対象経費・補助率・充当率・措置率・適用期限が年度ごとに変わる。
        本カリキュラムは「どの制度を、どのレベルで、どう使うか」を整理したものであり、
        数値と期限は必ず当年度の要綱・通知・地方債同意等基準で確認してから顧客に提示すること。
        L3未満は口頭でも数値を出さない。</p>
      </div>

      <h3 class="sub-h">提案素材 × レベル</h3>
      <div class="tablewrap">
        <table class="grid smatrix">
          <thead><tr><th scope="col">提案素材</th><th scope="col">L1 説明する</th><th scope="col">L2 伝える</th>
          <th scope="col">L3 当てはめる</th><th scope="col">L4 組み立てる</th><th scope="col">L5 創る</th>
          <th scope="col">共通の注意</th></tr></thead>
          <tbody>{mrows}</tbody>
        </table>
      </div>

      <h3 class="sub-h">提案ネタ台帳 <span class="sub-n">{len(SALES_KIT)}件（うち財源 {fin_n}件）</span></h3>
      <p class="sub-lead">国の指針、作成マニュアル、他自治体事例、補助金、起債などの財源、関連施策、現行計画の課題 ──
      提案の材料を実名で並べたもの。赤い行が財源で、「計画に位置付けられていることが起債・補助の要件」という順序が、
      計画策定業務の必要性を説明する最も強いロジックになる。</p>
      <div class="kfilter">{chips}</div>
      <div class="tablewrap">
        <table class="grid kit">
          <thead><tr><th scope="col">テーマ</th><th scope="col">区分</th><th scope="col">具体名（正式名称）</th>
          <th scope="col">提案での使い方</th><th scope="col">使える<br>レベル</th><th scope="col">確認先・更新タイミング</th></tr></thead>
          <tbody>{krows}</tbody>
        </table>
      </div>

      <h3 class="sub-h">訪問前 準備チェックリスト</h3>
      <p class="sub-lead">自分のレベルまでの項目をすべて満たしてから訪問する。準備が足りないまま訪問した回数だけ、その顧客での次の機会が遠のく。</p>
      <div class="preps">{pblocks}</div>
    </section>"""


def sec_calendar():
    cells = ""
    peak = {"8月": "営業", "9月": "営業", "10月": "営業", "1月": "営業", "2月": "営業",
            "7月": "業務", "12月": "業務", "3月": "業務"}
    for mo, gov, sales, work, rookie in CALENDAR:
        p = peak.get(mo, "")
        cells += f"""
        <div class="cal-col{' peak-' + ('s' if p == '営業' else 'w') if p else ''}">
          <div class="cal-mo">{e(mo)}{f'<span class="cal-peak">{e(p)}の山</span>' if p else ''}</div>
          <div class="cal-cell"><span class="k">自治体</span>{e(gov)}</div>
          <div class="cal-cell"><span class="k">営業</span>{e(sales)}</div>
          <div class="cal-cell"><span class="k">業務</span>{e(work)}</div>
          <div class="cal-cell rookie"><span class="k">新人の関与</span>{e(rookie)}</div>
        </div>"""
    return f"""
    <section id="calendar" class="sec">
      <header class="sec-head">
        <p class="eyebrow">05　年間サイクル</p>
        <h2>学習計画は自治体の年度に従属する</h2>
        <p class="lead">新人が「今月なにをやるか」は本人の習熟度だけでは決まらない。顧客の年度が決める。ロードマップの各月はこの表と対応している。</p>
      </header>
      <div class="tablewrap cal-wrap">
        <div class="cal">{cells}</div>
      </div>
    </section>"""


def sec_rubric():
    rows = "".join(
        f'<tr><th scope="row">{e(a[0])}</th>' +
        "".join(f'<td class="g{j+1}">{e(t)}</td>' for j, t in enumerate(a[1:])) +
        "</tr>" for a in RUBRIC)
    return f"""
    <section id="rubric" class="sec">
      <header class="sec-head">
        <p class="eyebrow">06　評価ルーブリック</p>
        <h2>テーマ知識に依存しない9つの軸</h2>
        <p class="lead">テーマ別スキルマップと合わせて判定する。自己評価と上長評価に2段階以上の差が出た項目は、必ず面談で認識を合わせる。</p>
      </header>
      <div class="tablewrap">
        <table class="grid rubric">
          <thead><tr><th scope="col">評価軸</th><th scope="col">L1 理解する</th><th scope="col">L2 補助する</th><th scope="col">L3 自立する</th><th scope="col">L4 応用・指導</th><th scope="col">L5 統括する</th></tr></thead>
          <tbody>{rows}</tbody>
        </table>
      </div>
    </section>"""


def sec_ops():
    ops = [("毎月", "上長と本人で30分の面談。当月の学習項目と習得目標の達成状況を確認し、翌月の重点を決める。"),
           ("四半期", "成果物レビューを1件必ず実施し、各項目の「確認方法」欄の方法で到達度を確認する。"),
           ("半期", "個人別チェックシートで自己評価と上長評価を突き合わせ、差異の原因を対話する。"),
           ("年次", "レベル判定を確定し、次年度の主担当・副担当テーマと育成計画を決定する。")]
    o = "".join(f'<li><span class="op-when">{e(a)}</span><p>{e(b)}</p></li>' for a, b in ops)
    files = [("10_新人教育カリキュラム_マスター管理表.xlsx",
              "16シート。階層定義・36ヶ月ロードマップ・テーマ別カリキュラム130項目・スキルマップ・個人別チェックシート・ルーブリック・教材リスト・年間カレンダー。"),
             ("11_個人別_習得度チェックシート.xlsx",
              "配布用の単票。24の評価項目に自己評価／上長評価／判定日／根拠となる実績を記入する。新人1名につき1ファイル。")]
    f = "".join(f'<li><code>{e(a)}</code><p>{e(b)}</p></li>' for a, b in files)
    return f"""
    <section id="ops" class="sec">
      <header class="sec-head">
        <p class="eyebrow">07　運用</p>
        <h2>回し方と配布物</h2>
      </header>
      <div class="ops-grid">
        <div>
          <h3>運用サイクル</h3>
          <ul class="ops">{o}</ul>
        </div>
        <div>
          <h3>配布する管理表</h3>
          <ul class="files">{f}</ul>
          <p class="note">前提：4月入社／育成期間36ヶ月／営業と業務の双方を担う総合職。中途入社は先にレベル判定を行い、該当レベルの月から開始する。制度改正（基本指針・報酬改定・会計基準・各省ガイドライン）が生じた際は、該当テーマの学習内容を年1回見直す。</p>
        </div>
      </div>
    </section>"""


CSS = """
:root{
  --paper:#F1F3F2; --surface:#FBFCFB; --surface-2:#E7ECEB; --ground:#E4E9E8;
  --ink:#17242A; --ink-2:#455A61; --ink-3:#6F838A;
  --rule:#D0D8D7; --rule-2:#BCC7C6;
  --accent:#1E5D6E; --accent-2:#2C7D91; --accent-soft:#DBE9EC;
  --warm:#8E6318; --warm-soft:#EFE3CE;
  --l1:#DEEAEC; --l1fg:#1C3A42; --l2:#B7D2D9; --l2fg:#17323A;
  --l3:#8CB8C3; --l3fg:#122A31; --l4:#4A8598; --l4fg:#F4FAFB; --l5:#1E5D6E; --l5fg:#F1F8F9;
  --shadow:0 1px 2px rgba(23,36,42,.06), 0 8px 24px -18px rgba(23,36,42,.5);
  --maxw:1180px;
}
@media (prefers-color-scheme: dark){
  :root:not([data-theme="light"]){
    --paper:#101719; --surface:#182226; --surface-2:#1F2B30; --ground:#0B1113;
    --ink:#E2EAEA; --ink-2:#A6B9BE; --ink-3:#7A9198;
    --rule:#2A383D; --rule-2:#38494F;
    --accent:#6FB6C7; --accent-2:#8ECBD9; --accent-soft:#1A363E;
    --warm:#D8AA57; --warm-soft:#2E2617;
    --l1:#203237; --l1fg:#BCD3D9; --l2:#2A4750; --l2fg:#CFE3E8;
    --l3:#3A6270; --l3fg:#E0F0F4; --l4:#4E8494; --l4fg:#F0FAFC; --l5:#7FC2D2; --l5fg:#0E1D22;
    --shadow:0 1px 2px rgba(0,0,0,.4), 0 10px 30px -20px rgba(0,0,0,.9);
  }
}
:root[data-theme="dark"]{
  --paper:#101719; --surface:#182226; --surface-2:#1F2B30; --ground:#0B1113;
  --ink:#E2EAEA; --ink-2:#A6B9BE; --ink-3:#7A9198;
  --rule:#2A383D; --rule-2:#38494F;
  --accent:#6FB6C7; --accent-2:#8ECBD9; --accent-soft:#1A363E;
  --warm:#D8AA57; --warm-soft:#2E2617;
  --l1:#203237; --l1fg:#BCD3D9; --l2:#2A4750; --l2fg:#CFE3E8;
  --l3:#3A6270; --l3fg:#E0F0F4; --l4:#4E8494; --l4fg:#F0FAFC; --l5:#7FC2D2; --l5fg:#0E1D22;
  --shadow:0 1px 2px rgba(0,0,0,.4), 0 10px 30px -20px rgba(0,0,0,.9);
}

*,*::before,*::after{box-sizing:border-box}
body{
  margin:0; background:var(--paper); color:var(--ink);
  font-family:"Zen Kaku Gothic New","Hiragino Sans","Noto Sans JP",system-ui,sans-serif;
  font-size:15px; line-height:1.85; letter-spacing:.01em;
  -webkit-font-smoothing:antialiased;
}
h1,h2,h3{font-family:"Zen Old Mincho","Hiragino Mincho ProN",serif; font-weight:600; text-wrap:balance; letter-spacing:.02em; margin:0}
code,.num,.m-no,.theme-count b{font-family:"IBM Plex Mono",ui-monospace,monospace; font-variant-numeric:tabular-nums}
a{color:var(--accent)}
:focus-visible{outline:2px solid var(--accent-2); outline-offset:2px; border-radius:2px}
.vh{position:absolute;width:1px;height:1px;overflow:hidden;clip:rect(0 0 0 0);white-space:nowrap}

.wrap{max-width:var(--maxw); margin:0 auto; padding:0 clamp(18px,4vw,44px); min-width:0}
.sec,.sec-head,.ops-grid>*,.tablewrap{min-width:0}

/* ---------- header ---------- */
.masthead{border-bottom:1px solid var(--rule); background:var(--surface)}
.mast-in{max-width:var(--maxw); margin:0 auto; padding:clamp(34px,6vw,68px) clamp(18px,4vw,44px) clamp(26px,4vw,44px)}
.org{font-size:11.5px; letter-spacing:.18em; color:var(--ink-3); text-transform:none; margin:0 0 18px;
  font-family:"IBM Plex Mono",monospace}
.org b{color:var(--accent); font-weight:500}
h1{font-size:clamp(29px,5vw,48px); line-height:1.3; margin:0 0 18px; max-width:17ch}
.h1-sub{display:block; font-size:.44em; line-height:1.7; letter-spacing:.05em; color:var(--ink-3);
  margin-top:16px; max-width:46ch; font-family:"Zen Kaku Gothic New","Hiragino Sans",sans-serif; font-weight:500}
.mast-lead{max-width:60ch; color:var(--ink-2); margin:0 0 30px; font-size:15.5px}
.facts{display:flex; flex-wrap:wrap; gap:0; border-top:1px solid var(--rule); padding-top:20px; margin:0; list-style:none}
.facts li{padding-right:clamp(20px,4vw,46px); margin-right:clamp(20px,4vw,46px); border-right:1px solid var(--rule)}
.facts li:last-child{border-right:0;margin-right:0;padding-right:0}
.facts b{display:block; font-family:"IBM Plex Mono",monospace; font-size:26px; font-weight:500; color:var(--accent); line-height:1.2}
.facts span{font-size:11.5px; color:var(--ink-3); letter-spacing:.06em}

/* ---------- nav ---------- */
.nav{position:sticky; top:0; z-index:20; background:color-mix(in srgb, var(--paper) 92%, transparent);
  backdrop-filter:blur(8px); border-bottom:1px solid var(--rule)}
.nav-in{max-width:var(--maxw); margin:0 auto; padding:0 clamp(18px,4vw,44px); display:flex; gap:2px; overflow-x:auto}
.nav a{padding:12px 14px; font-size:12.5px; color:var(--ink-2); text-decoration:none; white-space:nowrap;
  border-bottom:2px solid transparent}
.nav a:hover{color:var(--accent); border-bottom-color:var(--accent)}
.nav a .n{font-family:"IBM Plex Mono",monospace; color:var(--ink-3); margin-right:7px; font-size:11px}

/* ---------- sections ---------- */
.sec{padding:clamp(48px,7vw,86px) 0; border-bottom:1px solid var(--rule)}
.sec:last-child{border-bottom:0}
.sec-head{margin-bottom:clamp(26px,4vw,42px)}
.eyebrow{font-family:"IBM Plex Mono",monospace; font-size:11px; letter-spacing:.16em; color:var(--accent);
  margin:0 0 12px}
.sec h2{font-size:clamp(22px,3.2vw,32px); line-height:1.4; margin:0 0 14px; max-width:31ch}
.lead{max-width:64ch; color:var(--ink-2); margin:0; font-size:14.5px}
h3{font-size:17px; margin:0 0 14px}

/* ---------- ladder ---------- */
.ladder{list-style:none; margin:0 0 34px; padding:0; display:grid; gap:2px}
.rung{background:var(--surface); border:1px solid var(--rule); border-left:5px solid var(--rule-2);
  padding:16px 20px; display:grid; gap:8px}
.rung-1{border-left-color:var(--l1)} .rung-2{border-left-color:var(--l2)}
.rung-3{border-left-color:var(--l3)} .rung-4{border-left-color:var(--l4)} .rung-5{border-left-color:var(--l5)}
.rung-head{display:flex; align-items:baseline; gap:14px; flex-wrap:wrap}
.rung-code{font-family:"IBM Plex Mono",monospace; font-size:15px; font-weight:500; color:var(--accent); min-width:26px}
.rung-name{font-family:"Zen Old Mincho",serif; font-size:17px; font-weight:600}
.rung-when{margin-left:auto; font-size:11.5px; color:var(--ink-3); font-family:"IBM Plex Mono",monospace}
.rung-def{margin:0; font-size:14px; color:var(--ink-2); max-width:76ch}
.rung-meta{display:flex; flex-wrap:wrap; gap:8px 26px; margin:2px 0 0}
.rung-meta div{display:flex; gap:8px; align-items:baseline}
.rung-meta dt{font-size:10.5px; color:var(--ink-3); letter-spacing:.08em; margin:0}
.rung-meta dd{margin:0; font-size:12.5px; color:var(--ink-2)}
.targets{background:var(--surface-2); border:1px solid var(--rule); padding:22px 24px}
.tg-list{list-style:none; margin:0; padding:0; display:grid; gap:12px}
.tg-list li{display:grid; grid-template-columns:130px 250px 1fr; gap:16px; align-items:baseline;
  padding-bottom:12px; border-bottom:1px dashed var(--rule-2)}
.tg-list li:last-child{border-bottom:0; padding-bottom:0}
.tg-when{font-family:"IBM Plex Mono",monospace; font-size:12.5px; color:var(--accent)}
.tg-lv{font-size:13px; font-weight:500}
.tg-desc{font-size:13px; color:var(--ink-2)}
@media(max-width:760px){.tg-list li{grid-template-columns:1fr; gap:4px}}

/* ---------- level chips ---------- */
.lv{font-family:"IBM Plex Mono",monospace; font-size:11px; padding:2px 9px; white-space:nowrap;
  border-radius:2px; letter-spacing:.04em}
.lv1{background:var(--l1); color:var(--l1fg)} .lv2{background:var(--l2); color:var(--l2fg)}
.lv3{background:var(--l3); color:var(--l3fg)} .lv4{background:var(--l4); color:var(--l4fg)}
.lv5{background:var(--l5); color:var(--l5fg)}

/* ---------- roadmap ---------- */
.ytabs{display:flex; gap:2px; margin-bottom:0; border-bottom:1px solid var(--rule)}
.ytab{appearance:none; background:none; border:0; border-bottom:2px solid transparent; cursor:pointer;
  padding:11px 20px 13px; font-family:"Zen Old Mincho",serif; font-size:16px; color:var(--ink-3);
  display:flex; align-items:baseline; gap:9px; transition:color .15s}
.ytab span{font-family:"IBM Plex Mono",monospace; font-size:10.5px; letter-spacing:.04em}
.ytab:hover{color:var(--ink)}
.ytab.is-on{color:var(--accent); border-bottom-color:var(--accent)}
.ypanel-sub{font-size:12.5px; color:var(--ink-3); margin:16px 0 18px; letter-spacing:.04em}
.months{list-style:none; margin:0; padding:0; display:grid; gap:2px}
.mrow{display:grid; grid-template-columns:74px 1fr; background:var(--surface); border:1px solid var(--rule)}
.mrow.is-mile{border-color:var(--warm); background:var(--warm-soft)}
.m-when{border-right:1px solid var(--rule); padding:16px 0; display:flex; flex-direction:column;
  align-items:center; gap:2px; background:var(--surface-2)}
.mrow.is-mile .m-when{background:transparent; border-right-color:var(--warm)}
.m-no{font-family:"IBM Plex Mono",monospace; font-size:21px; font-weight:500; color:var(--accent); line-height:1.1}
.mrow.is-mile .m-no{color:var(--warm)}
.m-mo{font-size:11px; color:var(--ink-3)}
.m-body{padding:15px 20px 16px; display:grid; gap:9px}
.m-top{display:flex; align-items:center; gap:11px; flex-wrap:wrap}
.m-phase{font-size:11px; padding:2px 10px; border-radius:2px; letter-spacing:.04em}
.p1{background:var(--l1);color:var(--l1fg)} .p2{background:var(--l2);color:var(--l2fg)}
.p3{background:var(--l3);color:var(--l3fg)} .p4{background:var(--l4);color:var(--l4fg)}
.p5{background:var(--l5);color:var(--l5fg)} .p6{background:var(--warm);color:var(--paper)}
.m-focus{font-family:"Zen Old Mincho",serif; font-size:16px; font-weight:600}
.m-top .lv{margin-left:auto}
.k{display:inline-block; font-family:"IBM Plex Mono",monospace; font-size:10px; letter-spacing:.1em;
  color:var(--ink-3); margin-right:10px; vertical-align:1px}
.m-learn{margin:0; font-size:13.5px; color:var(--ink-2)}
.m-goal{margin:0; padding:0; list-style:none; font-size:13.5px}
.k-block{display:block; margin:2px 0 -5px}
.m-goal li{padding-left:0; margin-top:2px}
.m-foot{display:grid; gap:3px; border-top:1px dashed var(--rule-2); padding-top:9px; margin-top:2px}
.m-foot p{margin:0; font-size:12.5px; color:var(--ink-3)}
.m-mile{color:var(--ink-2)}
.mrow.is-mile .m-mile{color:var(--warm); font-weight:500}
@media(max-width:620px){.mrow{grid-template-columns:1fr}
  .m-when{flex-direction:row; gap:10px; border-right:0; border-bottom:1px solid var(--rule); padding:7px 16px; justify-content:flex-start}
  .m-no{font-size:16px} .m-top .lv{margin-left:0}}

/* ---------- themes ---------- */
.themes{display:grid; gap:2px; min-width:0}
.theme{background:var(--surface); border:1px solid var(--rule); padding:22px clamp(16px,3vw,26px) 24px; min-width:0}
.theme-head{display:flex; align-items:flex-start; gap:14px}
.theme-dot{width:10px; height:10px; background:var(--tc); border-radius:50%; margin-top:9px; flex:none}
.theme-head h3{font-size:19px; margin:0}
.theme-tag{margin:2px 0 0; font-size:11.5px; color:var(--ink-3); letter-spacing:.04em}
.theme-count{margin:0 0 0 auto; text-align:right; display:flex; align-items:baseline; gap:9px}
.theme-count b{font-size:24px; font-weight:500; color:var(--accent)}
.theme-count span{font-size:10.5px; color:var(--ink-3); line-height:1.5; text-align:left}
.theme-sub{margin:12px 0 14px; font-size:13.5px; color:var(--ink-2); max-width:80ch}
.theme-chips{list-style:none; display:flex; flex-wrap:wrap; gap:5px; margin:0 0 18px; padding:0}
.theme-chips li{font-size:11.5px; padding:3px 10px; background:var(--surface-2); border:1px solid var(--rule);
  color:var(--ink-2); border-radius:2px}
.tablewrap{overflow-x:auto; border:1px solid var(--rule); background:var(--surface)}
.grid{border-collapse:collapse; width:100%; min-width:900px; font-size:12.5px}
.grid th,.grid td{border:1px solid var(--rule); padding:9px 12px; text-align:left; vertical-align:top}
.grid thead th{background:var(--surface-2); font-family:"Zen Kaku Gothic New",sans-serif; font-weight:700;
  font-size:11.5px; letter-spacing:.05em; white-space:nowrap; color:var(--ink-2)}
.grid tbody th{background:var(--surface-2); font-weight:700; white-space:nowrap; width:88px; font-size:12.5px}
.grid td{color:var(--ink-2); line-height:1.7}
.grid td.g1{background:color-mix(in srgb,var(--l1) 40%,transparent)}
.grid td.g2{background:color-mix(in srgb,var(--l2) 30%,transparent)}
.grid td.g3{background:color-mix(in srgb,var(--l3) 26%,transparent)}
.grid td.g4{background:color-mix(in srgb,var(--l4) 20%,transparent)}
.grid td.g5{background:color-mix(in srgb,var(--l5) 15%,transparent)}
.rubric tbody th{width:150px; color:var(--ink)}
@media(max-width:640px){.theme-count{margin-left:0; width:100%; justify-content:flex-start}
  .theme-head{flex-wrap:wrap}}

/* ---------- calendar ---------- */
.cal-wrap{background:var(--surface)}
.cal{display:grid; grid-template-columns:repeat(12,minmax(168px,1fr)); min-width:900px}
.cal-col{border-right:1px solid var(--rule); display:grid; grid-template-rows:auto 1fr 1fr 1fr 1fr}
.cal-col:last-child{border-right:0}
.cal-col.peak-s{background:color-mix(in srgb,var(--accent-soft) 60%,transparent)}
.cal-col.peak-w{background:color-mix(in srgb,var(--warm-soft) 60%,transparent)}
.cal-mo{font-family:"IBM Plex Mono",monospace; font-size:13px; padding:9px 12px; border-bottom:1px solid var(--rule);
  background:var(--surface-2); display:flex; align-items:baseline; gap:8px; flex-wrap:wrap}
.cal-peak{font-size:9.5px; letter-spacing:.06em; padding:1px 6px; background:var(--accent); color:var(--paper); border-radius:2px}
.peak-w .cal-peak{background:var(--warm)}
.cal-cell{padding:9px 12px; font-size:11.5px; color:var(--ink-2); border-bottom:1px dashed var(--rule)}
.cal-cell:last-child{border-bottom:0}
.cal-cell .k{display:block; margin-bottom:3px}
.cal-cell.rookie{color:var(--ink); background:color-mix(in srgb,var(--surface-2) 55%,transparent)}

/* ---------- ops ---------- */
.ops-grid{display:grid; grid-template-columns:1fr 1fr; gap:clamp(24px,4vw,52px)}
@media(max-width:820px){.ops-grid{grid-template-columns:1fr}}
.ops,.files{list-style:none; margin:0; padding:0; display:grid; gap:2px}
.ops li{display:grid; grid-template-columns:78px 1fr; gap:14px; background:var(--surface);
  border:1px solid var(--rule); padding:13px 16px; align-items:baseline}
.op-when{font-family:"IBM Plex Mono",monospace; font-size:12.5px; color:var(--accent)}
.ops p,.files p{margin:0; font-size:13px; color:var(--ink-2)}
.files li{background:var(--surface); border:1px solid var(--rule); padding:13px 16px}
.files code{display:block; font-size:12px; color:var(--ink); margin-bottom:5px; word-break:break-all}
.note{margin:18px 0 0; font-size:12px; color:var(--ink-3); border-left:3px solid var(--rule-2); padding-left:14px; line-height:1.9}

footer.foot{padding:34px 0 46px; color:var(--ink-3); font-size:11.5px; letter-spacing:.04em}
@media (prefers-reduced-motion:reduce){*{transition:none!important; animation:none!important; scroll-behavior:auto!important}}
html{scroll-behavior:smooth}

/* ---------- 営業提案セクション ---------- */
.sub-h{font-size:17px; margin:38px 0 12px; padding-bottom:8px; border-bottom:1px solid var(--rule);
  display:flex; align-items:baseline; gap:12px; flex-wrap:wrap}
.sub-n{font-family:"IBM Plex Mono",monospace; font-size:11.5px; color:var(--ink-3); font-weight:400}
.sub-lead{margin:0 0 16px; font-size:13.5px; color:var(--ink-2); max-width:72ch}
.srungs{gap:2px}
.srung{background:var(--surface); border:1px solid var(--rule); border-left:5px solid var(--rule-2); padding:16px 20px}
.srung.rung-1{border-left-color:var(--l1)} .srung.rung-2{border-left-color:var(--l2)}
.srung.rung-3{border-left-color:var(--l3)} .srung.rung-4{border-left-color:var(--l4)} .srung.rung-5{border-left-color:var(--l5)}
.srung-head{display:flex; align-items:baseline; gap:14px; flex-wrap:wrap; margin-bottom:10px}
.srung-grid{display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:7px 30px; margin:0}
.srung-grid>div{display:grid; grid-template-columns:112px 1fr; gap:12px; align-items:baseline;
  padding-bottom:6px; border-bottom:1px dashed var(--rule)}
.srung-grid>div:nth-last-child(-n+1){border-bottom:0}
.srung-grid dt{margin:0; font-size:10.5px; letter-spacing:.06em; color:var(--ink-3)}
.srung-grid dd{margin:0; font-size:13px; color:var(--ink-2)}
.srung-grid .ng{grid-column:1/-1; border-bottom:0}
.srung-grid .ng dt{color:#B03A2E} .srung-grid .ng dd{color:var(--ink)}
:root[data-theme="dark"] .srung-grid .ng dt{color:#E8897C}
@media (prefers-color-scheme: dark){:root:not([data-theme="light"]) .srung-grid .ng dt{color:#E8897C}}
@media(max-width:820px){.srung-grid{grid-template-columns:1fr}}

.warn{margin:26px 0 4px; border:1px solid var(--warm); border-left-width:5px;
  background:var(--warm-soft); padding:15px 20px}
.warn p{margin:0; font-size:13px; color:var(--ink); max-width:88ch}
.warn b{color:var(--warm)}

.smatrix tbody th{width:180px; color:var(--ink)}
.smatrix .cav{background:var(--warm-soft); color:var(--ink-2); min-width:230px}

.kfilter{display:flex; flex-wrap:wrap; gap:4px; margin:0 0 12px}
.kchip{appearance:none; cursor:pointer; font-family:inherit; font-size:12px; padding:5px 12px;
  border:1px solid var(--rule-2); background:var(--surface); color:var(--ink-2); border-radius:2px;
  display:inline-flex; align-items:baseline; gap:7px}
.kchip span{font-family:"IBM Plex Mono",monospace; font-size:10.5px; color:var(--ink-3)}
.kchip:hover{border-color:var(--accent); color:var(--ink)}
.kchip.is-on{background:var(--accent); border-color:var(--accent); color:var(--paper)}
.kchip.is-on span{color:var(--paper); opacity:.75}
.kit{min-width:1080px; font-size:12.5px}
.kit .kt{white-space:nowrap; width:120px; color:var(--ink); font-weight:500}
.kit .kn{width:280px; color:var(--ink)}
.kit .kl{width:78px; text-align:center; white-space:nowrap}
.kit .ks{width:190px; font-size:11.5px; color:var(--ink-3)}
.kit tr.is-fin .kn{color:#B03A2E; font-weight:700}
.kit tr.is-fin .ks{color:#B03A2E}
:root[data-theme="dark"] .kit tr.is-fin .kn,:root[data-theme="dark"] .kit tr.is-fin .ks{color:#EE9C90}
@media (prefers-color-scheme: dark){:root:not([data-theme="light"]) .kit tr.is-fin .kn,
  :root:not([data-theme="light"]) .kit tr.is-fin .ks{color:#EE9C90}}
.cat{display:inline-block; font-size:10.5px; padding:2px 8px; border-radius:2px; white-space:nowrap;
  background:var(--surface-2); color:var(--ink-2); border:1px solid var(--rule-2)}
.cat[data-c="財源"]{background:#B03A2E; color:#fff; border-color:#B03A2E}
.cat[data-c="現行課題"]{background:var(--warm); color:var(--paper); border-color:var(--warm)}
.cat[data-c="指針"],.cat[data-c="マニュアル"],.cat[data-c="ガイドライン"]{background:var(--accent); color:var(--paper); border-color:var(--accent)}
.cat[data-c="注意"]{background:#7C2D22; color:#fff; border-color:#7C2D22}

.preps{display:grid; gap:2px}
.pgroup{background:var(--surface); border:1px solid var(--rule); padding:16px 20px;
  display:grid; grid-template-columns:44px 1fr; gap:18px; align-items:start}
.plist{list-style:none; margin:0; padding:0; display:grid; gap:12px}
.plist li{display:grid; grid-template-columns:170px 1fr; gap:6px 18px}
.pi{font-weight:700; font-size:13px; grid-row:span 2}
.pt{margin:0; font-size:13px; color:var(--ink-2)}
.pd{margin:0; font-size:12px; color:var(--ink-3)}
@media(max-width:760px){.pgroup{grid-template-columns:1fr; gap:10px}
  .plist li{grid-template-columns:1fr} .pi{grid-row:auto}}
"""

JS = """
(function(){
  var tabs=[].slice.call(document.querySelectorAll('.ytab'));
  tabs.forEach(function(t){
    t.addEventListener('click',function(){
      tabs.forEach(function(o){
        var on=o===t;
        o.classList.toggle('is-on',on);
        o.setAttribute('aria-selected',on?'true':'false');
        document.getElementById('y'+o.dataset.y).hidden=!on;
      });
    });
  });

  var chips=[].slice.call(document.querySelectorAll('.kchip')),
      rows=[].slice.call(document.querySelectorAll('.krow'));
  chips.forEach(function(c){
    c.addEventListener('click',function(){
      chips.forEach(function(o){o.classList.toggle('is-on',o===c);});
      var f=c.dataset.f;
      rows.forEach(function(r){ r.hidden = !(f==='all' || r.dataset.t===f); });
    });
  });
})();
"""


def build():
    n_items = sum(len(r) for *_, r in THEMES)
    page = f"""<title>公共コンサル育成ロードマップ</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Zen+Old+Mincho:wght@400;600;700&family=Zen+Kaku+Gothic+New:wght@400;500;700&family=IBM+Plex+Mono:wght@400;500&display=swap">
<style>{CSS}</style>

<header class="masthead">
  <div class="mast-in">
    <p class="org">BIZUP PUBLIC CONSULTING ／ <b>新人教育カリキュラム</b></p>
    <h1>自治体の年度に、育成の年度を重ねる
      <span class="h1-sub">公会計・公営企業会計・行政計画　8テーマの階層別 習得度管理</span></h1>
    <p class="mast-lead">入社してから36ヶ月、各月に何を学び、何ができるようになるべきか。テーマごとに営業と業務を分けて到達基準を定め、レベル判定で管理する。</p>
    <ul class="facts">
      <li><b>36</b><span>ヶ月ロードマップ</span></li>
      <li><b>8</b><span>テーマ</span></li>
      <li><b>{n_items}</b><span>学習項目</span></li>
      <li><b>L0–L5</b><span>階層区分</span></li>
      <li><b>{len(SALES_KIT)}</b><span>営業の提案ネタ</span></li>
      <li><b>33</b><span>評価項目</span></li>
    </ul>
  </div>
</header>

<nav class="nav" aria-label="目次">
  <div class="nav-in">
    <a href="#levels"><span class="n">01</span>階層</a>
    <a href="#roadmap"><span class="n">02</span>月次ロードマップ</a>
    <a href="#themes"><span class="n">03</span>テーマ別到達基準</a>
    <a href="#sales"><span class="n">04</span>営業提案の深度</a>
    <a href="#calendar"><span class="n">05</span>年間サイクル</a>
    <a href="#rubric"><span class="n">06</span>評価ルーブリック</a>
    <a href="#ops"><span class="n">07</span>運用</a>
  </div>
</nav>

<main class="wrap">
  {sec_levels()}
  {sec_roadmap()}
  {sec_themes()}
  {sec_sales()}
  {sec_calendar()}
  {sec_rubric()}
  {sec_ops()}
  <footer class="foot">
    <p>本カリキュラムは4月入社・育成期間36ヶ月・営業／業務の双方を担う総合職を前提としたモデルです。計画の「期」や制度の細目は年度により進行するため、案件着手時に必ず最新の法令・指針を確認してください。</p>
  </footer>
</main>
<script>{JS}</script>
"""
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(page)
    print(f"作成: {OUT}  ({len(page)/1024:.0f} KB)")


if __name__ == "__main__":
    build()
