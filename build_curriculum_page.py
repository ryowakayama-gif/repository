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
from curriculum_talk import HEARING_COMMON, HEARING_THEME, TALK_SCRIPT
from curriculum_csv import (CSV_HEADER, CSV_RULES, REMARK_GUIDE, THEME_TITLES,
                            VISIT_FORM, DATA_STATUS)
from curriculum_case import CASES, PLAN_HOSEI, PLAN_NEXT
from curriculum_train import TRAIN_A, TRAIN_B, TRAIN_SERIES, TRAIN_RULES

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


HT_THEMES = ["公会計", "公営企業会計", "高齢者・介護保険", "障害者・障害児",
             "地域福祉", "こども・子育て", "その他計画"]
HT_SLUG = {t: f"h{i}" for i, t in enumerate(HT_THEMES)}


def sec_talk():
    # --- トークスクリプト（レベルタブ） ---
    by_lv = {}
    for lv, scene, step, script, aim, ng in TALK_SCRIPT:
        by_lv.setdefault(lv, []).append((scene, step, script, aim, ng))
    tpanels = ""
    lv_note = {"L1": "同行が中心。話す量は少なく、記録が仕事。",
               "L2": "単独訪問。国の動きを届け、現状を聞いて帰る。売り込まない。",
               "L3": "課題を示し、財源を添えて、予算要求に載せるところまで。"}
    for i, lv in enumerate(("L1", "L2", "L3"), 1):
        cards = ""
        cur = None
        for scene, step, script, aim, ng in by_lv.get(lv, []):
            if scene != cur:
                cards += f'<li class="scene-h"><span class="lv lv{lv[1]}">{lv}</span>{e(scene)}</li>'
                cur = scene
            speech = script.startswith("「")
            body = (f'<blockquote class="speech">{e(script)}</blockquote>' if speech
                    else f'<p class="stage">{e(script)}</p>')
            cards += f"""
          <li class="tcard">
            <p class="tstep">{e(step)}</p>
            {body}
            <dl class="tmeta">
              <div><dt>意図</dt><dd>{e(aim)}</dd></div>
              <div class="ng"><dt>これはNG</dt><dd>{e(ng)}</dd></div>
            </dl>
          </li>"""
        tpanels += (f'<div class="tpanel" id="t{i}" role="tabpanel" aria-labelledby="ttab{i}"'
                    f'{"" if i == 1 else " hidden"}>'
                    f'<p class="ypanel-sub">{e(lv_note[lv])}</p>'
                    f'<ul class="talks">{cards}</ul></div>')
    ttabs = "".join(
        f'<button class="ytab ttab{" is-on" if i == 1 else ""}" id="ttab{i}" role="tab" '
        f'aria-selected="{"true" if i == 1 else "false"}" aria-controls="t{i}" data-t="{i}">'
        f'{lv}<span>{len(by_lv.get(lv, []))}場面</span></button>'
        for i, lv in enumerate(("L1", "L2", "L3"), 1))

    # --- ヒアリング項目（共通） ---
    hc = ""
    cur = None
    for lv, scene, q, aim, nxt, must in HEARING_COMMON:
        key = (lv, scene)
        if key != cur:
            hc += (f'<li class="scene-h"><span class="lv lv{lv[1]}">{lv}</span>{e(scene)}</li>')
            cur = key
        tag = '<span class="must">必須</span>' if must == "必須" else (
            '<span class="opt">推奨</span>' if must == "推奨" else '<span class="opt">任意</span>')
        qh = (f'<blockquote class="speech q">{e(q)}</blockquote>' if q.startswith("「")
              else f'<p class="stage">{e(q)}</p>')
        hc += f"""
      <li class="hcard">
        <div class="hq">{qh}{tag}</div>
        <dl class="tmeta">
          <div><dt>何が分かるか</dt><dd>{e(aim)}</dd></div>
          <div><dt>次の一手</dt><dd>{e(nxt)}</dd></div>
        </dl>
      </li>"""

    # --- ヒアリング項目（テーマ別） ---
    hrows = ""
    for theme, lv, q, aim in HEARING_THEME:
        hrows += (f'<tr class="hrow" data-h="{HT_SLUG.get(theme, "")}">'
                  f'<td class="kt">{e(theme)}</td>'
                  f'<td class="kl"><span class="lv lv{lv[1]}">{e(lv)}</span></td>'
                  f'<td class="hqt">{e(q)}</td><td>{e(aim)}</td></tr>')
    hchips = ('<button class="kchip hchip is-on" data-g="all">すべて<span>'
              + str(len(HEARING_THEME)) + "</span></button>")
    for t in HT_THEMES:
        n = sum(1 for r in HEARING_THEME if r[0] == t)
        hchips += f'<button class="kchip hchip" data-g="{HT_SLUG[t]}">{e(t)}<span>{n}</span></button>'

    return f"""
    <section id="talk" class="sec">
      <header class="sec-head">
        <p class="eyebrow">05　ヒアリングとトークスクリプト</p>
        <h2>L1〜L3が現場で実際に口に出す言葉</h2>
        <p class="lead">台詞は暗記するものではなく、意図を理解したうえで自分の言葉に直して使うもの。
        ただし「これはNG」に挙げた失敗だけは、そのまま避けてほしい ── 取り返すのに年単位かかる。
        ◯◯・△△は訪問前に必ず埋める。</p>
      </header>

      <h3 class="sub-h">トークスクリプト</h3>
      <div class="ytabs ttabs" role="tablist" aria-label="レベル">{ttabs}</div>
      {tpanels}

      <h3 class="sub-h">ヒアリング項目（共通） <span class="sub-n">{len(HEARING_COMMON)}問</span></h3>
      <p class="sub-lead">場面ごとに、聞く順番のまま並べてある。L1は相手に質問するより「記録する」ことが仕事。</p>
      <ul class="hlist">{hc}</ul>

      <h3 class="sub-h">ヒアリング項目（テーマ別） <span class="sub-n">{len(HEARING_THEME)}問</span></h3>
      <p class="sub-lead">訪問前にテーマで絞り込んで印刷して持参する。答えは相手の言葉のまま記録する ──
      要約した時点で、次の提案に使える情報ではなくなる。</p>
      <div class="kfilter">{hchips}</div>
      <div class="tablewrap">
        <table class="grid hkit">
          <thead><tr><th scope="col">テーマ</th><th scope="col">Lv</th>
          <th scope="col">質問文（このまま聞ける形）</th><th scope="col">聞く目的・使いどころ</th></tr></thead>
          <tbody>{hrows}</tbody>
        </table>
      </div>
    </section>"""


def sec_csv():
    # 入力規則
    rrows = ""
    for no, name, must, fmt, ch, ex, warn in CSV_RULES:
        req = "req" if must.startswith("必須") else "opt"
        rrows += (f'<tr><td class="cno">{no}</td><td class="cname">{e(name)}</td>'
                  f'<td class="creq"><span class="{req}">{e(must)}</span></td>'
                  f'<td>{e(fmt)}</td><td class="cch">{e(ch)}</td>'
                  f'<td class="cex">{e(ex)}</td><td class="cwarn">{e(warn)}</td></tr>')

    # 欠測率（順位付き横棒・単一系列）
    ranked = sorted(DATA_STATUS, key=lambda r: -r[3])
    bars = ""
    for name, ok, ng, pct, note in ranked:
        flag = '<span class="prio">要対応</span>' if pct >= 33 else ""
        bars += f"""
        <div class="brow">
          <div class="blab">{e(name)}{flag}</div>
          <div class="btrack"><div class="bfill" style="width:{pct}%"></div></div>
          <div class="bval">{pct}<span>%</span></div>
          <div class="bnote">{e(note) if note != "―" else ""}</div>
        </div>"""

    # 備考の書き方
    rg = ""
    for kind, ex, note in REMARK_GUIDE:
        cls = ("good" if kind.startswith("◎") else
               "bad" if kind.startswith("×") else
               "warn2" if kind.startswith("△") else "neutral")
        rg += (f'<li class="rg {cls}"><span class="rgk">{e(kind)}</span>'
               f'<p class="rgx">{e(ex)}</p><p class="rgn">{e(note)}</p></li>')

    # テーマ・タイトル
    trows = ""
    prev = None
    for theme, title, cur, sheet in THEME_TITLES:
        first = theme != prev
        prev = theme
        trows += (f'<tr class="{"tsep" if first else ""}">'
                  f'<td class="kt">{e(theme) if first else ""}</td>'
                  f'<td class="kn">{e(title)}</td><td>{e(cur)}</td>'
                  f'<td class="ks">{e(sheet)}</td></tr>')

    csv_items = [(no, item) for blk, no, item, *_ in VISIT_FORM if no]
    int_items = [item for blk, no, item, *_ in VISIT_FORM if not no]
    csv_l = "".join(f'<li><span class="cnum">{no}</span>{e(it)}</li>'
                    for no, it in sorted(csv_items))
    int_l = "".join(f"<li>{e(it)}</li>" for it in int_items)

    return f"""
    <section id="csv" class="sec">
      <header class="sec-head">
        <p class="eyebrow">06　訪問記録とCSV</p>
        <h2>記録は営業活動の資産になって初めて意味がある</h2>
        <p class="lead">最終的にTeams上でCSVとして管理するため、訪問記録の様式はCSVの15列をそのまま含む。
        列構成は変更しない。選択肢は実データ141件で実際に使われている値をそのまま採用した。</p>
      </header>

      <h3 class="sub-h">様式の構成</h3>
      <div class="fsplit">
        <div class="fcol">
          <p class="fhead"><b>CSV出力項目</b>　15列 ── Teamsに上げる</p>
          <ol class="flist csv">{csv_l}</ol>
        </div>
        <div class="fcol">
          <p class="fhead"><b>社内記録</b>　CSVには出さない</p>
          <ul class="flist">{int_l}</ul>
        </div>
      </div>
      <p class="note2">1回の訪問で複数テーマを話した場合、CSVはテーマごとに1行になる。訪問記録も1テーマ1枚とし、
      訪問日・団体名・担当者は各行に同じ値を入れる。金額の枠感や競合の具体名は社内記録にとどめ、備考には書かない。</p>

      <h3 class="sub-h">入力規則</h3>
      <div class="tablewrap">
        <table class="grid crules">
          <thead><tr><th scope="col">No</th><th scope="col">列名</th><th scope="col">必須区分</th>
          <th scope="col">入力形式</th><th scope="col">選択肢</th><th scope="col">記入例</th>
          <th scope="col">よくある誤り・注意</th></tr></thead>
          <tbody>{rrows}</tbody>
        </table>
      </div>

      <figure class="chart">
        <figcaption>
          <h4>列ごとの欠測率</h4>
          <p>現行データ141件。欠測率33%以上を要対応とした。契約予定・タイトル・前回策定の3列が、
          受注確度の判断とテーマ別の集計をそのまま不能にしている。</p>
        </figcaption>
        <div class="bars">{bars}</div>
        <p class="chart-src">出典：訪問記録CSV（141件、空行2行を除く）</p>
      </figure>

      <h3 class="sub-h">備考の書き方</h3>
      <p class="sub-lead">15列のうち、読み手にとって最も価値があるのが備考。
      型は［現況］＋［根拠・時期］＋［次の一手］を1行で。</p>
      <ul class="rglist">{rg}</ul>

      <h3 class="sub-h">テーマ・タイトル標準リスト</h3>
      <p class="sub-lead">タイトルはここから選ぶ。現行データでは
      「公共施設等総合管理計画」と「〜改定業務」、「台帳精緻化（所有外）」「所有外管理資産」
      「固定資産台帳精緻化（所有外管理資産）」が併存しており、そのままでは業務別の集計ができない。</p>
      <div class="tablewrap">
        <table class="grid ttl">
          <thead><tr><th scope="col">CSVのテーマ</th><th scope="col">標準タイトル</th>
          <th scope="col">カリキュラム上のテーマ</th><th scope="col">訪問前に見るシート</th></tr></thead>
          <tbody>{trows}</tbody>
        </table>
      </div>
      <div class="warn">
        <p><b>区分の当てはめで迷いやすいもの</b>　こども計画・子ども子育て支援事業計画・高齢者福祉計画は「福祉計画」
        （現行データでは「その他計画」に入っている）。固定資産台帳精緻化は「公会計」。公営企業は、
        料金改定＝経営改善／戦略の策定・改定＝経営戦略／日常の会計・決算＝会計支援。
        なお「社会生活計画」と「その他計画」の線引きは現状定まっておらず（環境系が両方に存在）、
        上表は案。社内で確定したうえで運用を統一してほしい。</p>
      </div>
    </section>"""


def _mins(t):
    """'0:30–1:30' → 秒数。時間表記でない見出し（帰社後など）は 0 を返す。"""
    t = t.replace("–", "-")
    if "-" not in t or ":" not in t:
        return 0
    a, b = t.split("-")
    def sec(x):
        m, s2 = x.split(":")
        return int(m) * 60 + int(s2)
    try:
        return sec(b) - sec(a)
    except ValueError:
        return 0


def sec_case():
    tabs, panels = "", ""
    for n, c in enumerate(CASES, 1):
        on = n == 1
        hot = c.get("axisgrp") == "A"
        tabs += (f'<button class="ytab ctab{" is-on" if on else ""}{" is-hot" if hot else ""}" '
                 f'id="ctab{n}" role="tab" '
                 f'aria-selected="{"true" if on else "false"}" aria-controls="cs{n}" data-c="{n}">'
                 f'{e(c["label"])}<span>{e(c["theme"])}</span></button>')

        segs = [(_mins(t), t, h) for t, h, *_ in c["steps"]]
        tl = ""
        for i, (d, t, h) in enumerate(segs, 1):
            g = min(i, 5)
            flexed = f"flex:{d}" if d else "flex:0 0 auto"
            tl += (f'<div class="tseg g{g}" style="{flexed}">'
                   f'<span class="tsl">{e(t.split("–")[0])}</span>'
                   f'<span class="tsh">{e(h)}</span></div>')

        cards = "".join(f"""
        <li class="ccard">
          <p class="ctime">{e(t)}</p>
          <div class="cbody">
            <h4>{e(head)}</h4>
            <blockquote class="speech">{e(script)}</blockquote>
            <p class="caim"><span class="k">ねらい</span>{e(aim)}</p>
          </div>
        </li>""" for t, head, script, aim in c["steps"])

        qa = "".join(
            f'<li class="qa"><p class="qq">{e(q)}</p>'
            + (f'<blockquote class="speech q">{e(a)}</blockquote>' if a.startswith("「")
               else f'<p class="stage">{e(a)}</p>')
            + f'<p class="qn"><span class="k">補足</span>{e(nt)}</p></li>'
            for q, a, nt in c["qa"])
        ng = "".join(f"<li>{e(x)}</li>" for x in c["ng"])

        panels += f"""
      <div class="cpanel" id="cs{n}" role="tabpanel" aria-labelledby="ctab{n}"{"" if on else " hidden"}>
        <div class="axis{" is-hot" if c.get("axisgrp") == "A" else ""}">
          <p class="axis-t">{"今期の補正 ── 期限が迫っている" if c.get("axisgrp") == "A" else "この5分で伝えることは1つだけ"}</p>
          <p class="axis-b">{e(c["axis"])}</p>
          <p class="axis-g"><span class="k">対象</span>{e(c["target"])}</p>
        </div>
        <div class="timeline">{tl}</div>
        <ol class="ccards">{cards}</ol>
        <h4 class="sub-h2">想定問答</h4>
        <ul class="qalist">{qa}</ul>
        <h4 class="sub-h2">やってはいけないこと</h4>
        <ul class="nglist">{ng}</ul>
      </div>"""

    def plan(rows, themed=False):
        out = ""
        prev = None
        for row in rows:
            th, k, h, b = row if themed else (None, *row)
            if themed and th != prev:
                out += f'<li class="pgrp">{e(th)}</li>'
                prev = th
            out += (f'<li class="pl"><span class="plk" data-k="{e(k)}">{e(k)}</span>'
                    f'<div><h5>{e(h)}</h5><p>{e(b)}</p></div></li>')
        return out

    n_steps = sum(len(c["steps"]) for c in CASES)
    return f"""
    <section id="case" class="sec">
      <header class="sec-head">
        <p class="eyebrow">07　ケース別の営業</p>
        <h2>断られ方には型がある</h2>
        <p class="lead">訪問記録141件の分布に合わせて5ケース。公会計が44件で最多、
        結果では「他社随契」20件・「不在」15件。うち他社随契の7件が「再訪不要」で打ち切られているが、
        その多くは改定時期がR9〜R14と先で、本来は改定前年度に戻るべき先だった。
        型を持っていないと、その場で引くか、言ってはいけないことを言うかのどちらかになる。</p>
      </header>

      <h3 class="sub-h">5分トーク集 <span class="sub-n">{len(CASES)}ケース／{n_steps}場面</span></h3>
      <div class="ytabs ctabs" role="tablist" aria-label="ケース">{tabs}</div>
      {panels}

      <h3 class="sub-h">営業方針 <span class="sub-n">A 今期の補正／B 来期案件獲得</span></h3>
      <div class="plans">
        <div class="plan" style="--pc:#2CA02C">
          <h4>A　今期の補正に向けた提案営業</h4>
          <p class="psub">高齢者介護（第10期）／障がい3計画（第8期・第4期）</p>
          <p class="pintro">どちらもR8年度が策定年度で、小規模団体では同じ課が両方を持っている。
          <b>狙う枠は12月補正</b>（財政課への要求は11月上旬）、逆算して10月中に見積を出す。
          9月補正はすでに間に合わない。</p>
          <ul class="pllist">{plan(PLAN_HOSEI, True)}</ul>
        </div>
        <div class="plan" style="--pc:#D6336C">
          <h4>B　来期案件獲得のための営業</h4>
          <p class="psub">地域福祉計画／こども計画（R9年度当初予算）</p>
          <p class="pintro">どちらも努力義務で、待っていても発注は生まれない。
          <b>動くのは8〜10月</b>、11〜12月が査定、1〜3月が公告。
          予算要求書が固まってから訪問しても、その年度には載らない。</p>
          <ul class="pllist">{plan(PLAN_NEXT, True)}</ul>
        </div>
      </div>
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
        <p class="eyebrow">08　年間サイクル</p>
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
        <p class="eyebrow">09　評価ルーブリック</p>
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
    def script_block(rows, gid):
        segs = "".join(
            f'<div class="tseg g{min(i,5)}" style="flex:{_mins(t)}">'
            f'<span class="tsl">{e(t.split("–")[0])}</span>'
            f'<span class="tsh">{e(h)}</span></div>'
            for i, (t, h, *_r) in enumerate(rows, 1))
        cards = "".join(f"""
        <li class="ccard">
          <p class="ctime">{e(t)}</p>
          <div class="cbody">
            <h4>{e(head)}</h4>
            <blockquote class="speech">{e(script)}</blockquote>
            <p class="caim"><span class="k">押さえどころ</span>{e(note)}</p>
          </div>
        </li>""" for t, head, script, note in rows)
        return f'<div class="timeline" id="{gid}">{segs}</div><ol class="ccards">{cards}</ol>'

    srows = "".join(
        f'<tr class="{"ax-a" if ax.startswith("A") else "ax-b"}">'
        f'<td class="wax">{e(ax)}</td><td class="wno">第{n}回</td><td class="wth">{e(th)}</td>'
        f'<td class="wmsg">{e(msg)}</td><td class="wmat">{e(mat)}</td>'
        f'<td class="wex">{e(ex)}</td></tr>'
        for ax, n, th, msg, mat, ex in TRAIN_SERIES)

    rules = "".join(
        f'<li class="rule{" is-no" if k == "やらないこと" else ""}">'
        f'<span class="rk">{e(k)}</span><div><h5>{e(h)}</h5><p>{e(w)}</p></div></li>'
        for k, h, w in TRAIN_RULES)

    ops = [("毎月", "上長と本人で30分の面談。当月の学習項目と習得目標の達成状況を確認し、翌月の重点を決める。"),
           ("四半期", "成果物レビューを1件必ず実施し、各項目の「確認方法」欄の方法で到達度を確認する。"),
           ("半期", "個人別チェックシートで自己評価と上長評価を突き合わせ、差異の原因を対話する。"),
           ("年次", "レベル判定を確定し、次年度の主担当・副担当テーマと育成計画を決定する。")]
    o = "".join(f'<li><span class="op-when">{e(a)}</span><p>{e(b)}</p></li>' for a, b in ops)
    files = [("10_新人教育カリキュラム_マスター管理表.xlsx",
              "30シート。階層定義、36ヶ月ロードマップ、テーマ別カリキュラム、営業提案の深度と提案ネタ台帳、"
              "ヒアリング項目とトークスクリプト、訪問記録とCSV、5分トーク集、来期営業方針、社内研修。"),
             ("11_個人別_習得度チェックシート.xlsx",
              "配布用の単票。33の評価項目に自己評価／上長評価／判定日／根拠となる実績を記入する。新人1名につき1ファイル。"),
             ("13_訪問記録_CSVテンプレート.csv",
              "Teams提出用のヘッダ行。BOM付きUTF-8・CRLF・全項目クォート（現行ファイルと同形式）。")]
    f = "".join(f'<li><code>{e(a)}</code><p>{e(b)}</p></li>' for a, b in files)

    return f"""
    <section id="ops" class="sec">
      <header class="sec-head">
        <p class="eyebrow">10　運用と社内展開</p>
        <h2>配って終わりにしないための段取り</h2>
        <p class="lead">研修の中身は、いま動かすもの（今期の補正）と、これから仕込むもの（来期の獲得）に分けた。
        5分は<b>説明3分＋演習2分</b>で組み、毎回「今週やること」を1つだけ指定して締める ──
        行動を指定しない研修は、聞いて終わる。</p>
      </header>

      <h3 class="sub-h">A　補正編 5分台本 <span class="sub-n">高齢者介護／障がい計画・今期の12月補正</span></h3>
      <div class="axis is-hot">
        <p class="axis-t">今期の補正 ── 期限が迫っている</p>
        <p class="axis-b">R8年度は高齢者の第10期と障がいの第8期・第4期が重なる策定年度。
        狙う枠は12月補正だけで、財政課への要求は11月上旬。逆算して10月中に見積を出す。
        そして<b>高齢者と障がいは必ず分けて見積を出す</b> ── 1本にまとめると金額が膨らみ、査定で両方とも落ちる。</p>
        <p class="axis-g"><span class="k">対象</span>営業メンバー全員（8〜9月に実施）</p>
      </div>
      {script_block(TRAIN_A, "trA")}

      <h3 class="sub-h">B　来期編 5分台本 <span class="sub-n">地域福祉計画／こども計画・R9年度当初予算</span></h3>
      <div class="axis">
        <p class="axis-t">この5分で伝えることは1つだけ</p>
        <p class="axis-b">どちらも努力義務で、待っていても発注は生まれない。動くのは8〜10月。
        地域福祉は<b>束ねる</b>、こどもは<b>制度変更を理由にする</b>。
        計画本体より、交付金の使える実態調査を先に取る。</p>
        <p class="axis-g"><span class="k">対象</span>営業メンバー全員（10〜11月に実施）</p>
      </div>
      {script_block(TRAIN_B, "trB")}

      <h3 class="sub-h">週次5分研修シリーズ <span class="sub-n">A編4回 → B編4回・約2か月で一巡</span></h3>
      <div class="tablewrap">
        <table class="grid weekly">
          <thead><tr><th scope="col">軸</th><th scope="col">回</th><th scope="col">テーマ</th>
          <th scope="col">伝えること（1回に1つだけ）</th><th scope="col">使う教材</th>
          <th scope="col">その場の演習（2分）</th></tr></thead>
          <tbody>{srows}</tbody>
        </table>
      </div>

      <h3 class="sub-h">運営のしかた</h3>
      <ul class="rules">{rules}</ul>

      <h3 class="sub-h">育成の運用サイクルと配布物</h3>
      <div class="ops-grid">
        <div>
          <h4 class="mini-h">運用サイクル</h4>
          <ul class="ops">{o}</ul>
        </div>
        <div>
          <h4 class="mini-h">配布する管理表</h4>
          <ul class="files">{f}</ul>
          <p class="note">前提：4月入社／育成期間36ヶ月／営業と業務の双方を担う総合職。中途入社は先にレベル判定を行い、
          該当レベルの月から開始する。制度改正（基本指針・報酬改定・会計基準・各省ガイドライン）が生じた際は、
          該当テーマの学習内容を年1回見直す。</p>
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
.nav a{padding:12px 12px; font-size:12.5px; color:var(--ink-2); text-decoration:none; white-space:nowrap;
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

/* ---------- ヒアリング／トークスクリプト ---------- */
.talks,.hlist{list-style:none; margin:0; padding:0; display:grid; gap:2px}
.scene-h{display:flex; align-items:baseline; gap:12px; font-family:"Zen Old Mincho",serif;
  font-size:15px; font-weight:600; padding:18px 2px 7px; border-bottom:1px solid var(--rule-2)}
.scene-h:first-child{padding-top:4px}
.tcard,.hcard{background:var(--surface); border:1px solid var(--rule); padding:15px 20px 16px; display:grid; gap:10px}
.tstep{margin:0; font-size:11px; letter-spacing:.08em; color:var(--ink-3);
  font-family:"IBM Plex Mono",monospace}
.speech{margin:0; padding:11px 0 11px 18px; border-left:3px solid var(--accent); white-space:pre-line;
  font-size:15px; line-height:1.95; color:var(--ink); font-feature-settings:"palt" 1; max-width:78ch}
.speech.q{font-size:14.5px}
.stage{margin:0; padding:9px 0 9px 18px; border-left:3px dashed var(--rule-2);
  font-size:13.5px; color:var(--ink-2); max-width:78ch}
.tmeta{margin:0; display:grid; gap:5px}
.tmeta>div{display:grid; grid-template-columns:88px 1fr; gap:14px; align-items:baseline}
.tmeta dt{margin:0; font-size:10.5px; letter-spacing:.06em; color:var(--ink-3)}
.tmeta dd{margin:0; font-size:13px; color:var(--ink-2)}
.tmeta .ng dt{color:#B03A2E}
.tmeta .ng dd{color:var(--ink)}
:root[data-theme="dark"] .tmeta .ng dt{color:#E8897C}
@media (prefers-color-scheme: dark){:root:not([data-theme="light"]) .tmeta .ng dt{color:#E8897C}}
.hq{display:flex; align-items:flex-start; gap:14px}
.hq>*:first-child{flex:1; min-width:0}
.must,.opt{flex:none; font-size:10px; letter-spacing:.06em; padding:2px 8px; border-radius:2px; margin-top:12px}
.must{background:#B03A2E; color:#fff}
.opt{background:var(--surface-2); color:var(--ink-3); border:1px solid var(--rule-2)}
.ttabs{margin-bottom:0}
.hkit{min-width:820px}
.hkit .hqt{width:400px; color:var(--ink)}

/* ---------- 訪問記録とCSV ---------- */
.fsplit{display:grid; grid-template-columns:1fr 1fr; gap:2px}
@media(max-width:820px){.fsplit{grid-template-columns:1fr}}
.fcol{background:var(--surface); border:1px solid var(--rule); padding:16px 20px 18px; min-width:0}
.fhead{margin:0 0 12px; font-size:12.5px; color:var(--ink-3); letter-spacing:.04em;
  padding-bottom:9px; border-bottom:1px solid var(--rule)}
.fhead b{color:var(--ink); font-size:14px; margin-right:8px}
.flist{list-style:none; margin:0; padding:0; display:grid; gap:5px; font-size:13.5px; color:var(--ink-2)}
.flist li{display:flex; gap:11px; align-items:baseline}
.flist.csv li{color:var(--ink)}
.cnum{flex:none; width:22px; text-align:right; font-family:"IBM Plex Mono",monospace;
  font-size:11px; color:var(--accent)}
.flist:not(.csv) li::before{content:"—"; color:var(--ink-3); flex:none}
.note2{margin:14px 0 0; font-size:12.5px; color:var(--ink-3); border-left:3px solid var(--rule-2);
  padding-left:14px; line-height:1.9; max-width:84ch}

.crules{min-width:1120px}
.crules .cno{width:34px; text-align:center; font-family:"IBM Plex Mono",monospace; color:var(--ink-3)}
.crules .cname{width:100px; color:var(--ink); font-weight:700; white-space:nowrap}
.crules .creq{width:112px}
.crules .cch{width:200px; font-size:11.5px}
.crules .cex{width:180px; color:var(--ink)}
.crules .cwarn{width:330px}
.req,.opt{font-size:10.5px; padding:2px 8px; border-radius:2px; white-space:nowrap; display:inline-block}
.req{background:#B03A2E; color:#fff}
.opt{background:var(--surface-2); color:var(--ink-2); border:1px solid var(--rule-2)}

/* 欠測率（単一系列・順位付き横棒） */
.chart{margin:26px 0 0; padding:20px 22px 16px; background:var(--surface); border:1px solid var(--rule)}
.chart figcaption{margin:0 0 18px}
.chart h4{margin:0 0 6px; font-family:"Zen Old Mincho",serif; font-size:16px; font-weight:600}
.chart figcaption p{margin:0; font-size:12.5px; color:var(--ink-2); max-width:76ch}
.bars{display:grid; gap:7px}
.brow{display:grid; grid-template-columns:150px minmax(90px,1fr) 52px minmax(0,1.5fr);
  gap:14px; align-items:center}
.blab{font-size:12.5px; color:var(--ink); display:flex; align-items:baseline; gap:8px; justify-content:flex-end}
.prio{font-size:9.5px; letter-spacing:.04em; padding:1px 6px; border-radius:2px;
  background:var(--warm); color:var(--paper); flex:none}
.btrack{height:16px; background:var(--surface-2); border-radius:0 2px 2px 0; position:relative}
.bfill{height:100%; background:var(--accent); border-radius:0 4px 4px 0}
.bval{font-family:"IBM Plex Mono",monospace; font-size:12.5px; font-variant-numeric:tabular-nums;
  text-align:right; color:var(--ink)}
.bval span{font-size:9.5px; color:var(--ink-3); margin-left:1px}
.bnote{font-size:11.5px; color:var(--ink-3); line-height:1.6}
.chart-src{margin:16px 0 0; font-size:10.5px; color:var(--ink-3); letter-spacing:.03em}
@media(max-width:820px){.brow{grid-template-columns:112px 1fr 46px}
  .bnote{grid-column:1/-1; padding-left:126px; margin-top:-2px}}
@media(max-width:520px){.brow{grid-template-columns:100px 1fr 44px}
  .blab{justify-content:flex-start; flex-wrap:wrap; gap:4px}
  .bnote{padding-left:0}}

.rglist{list-style:none; margin:0; padding:0; display:grid; gap:2px}
.rg{background:var(--surface); border:1px solid var(--rule); border-left:5px solid var(--rule-2);
  padding:13px 18px; display:grid; grid-template-columns:88px 1fr; gap:4px 16px; align-items:baseline}
.rg.good{border-left-color:#4C8B54} .rg.bad{border-left-color:#B03A2E}
.rg.warn2{border-left-color:var(--warm)} .rg.neutral{border-left-color:var(--accent)}
.rgk{font-size:11.5px; font-weight:700; color:var(--ink-2)}
.rg.good .rgk{color:#3F7547} .rg.bad .rgk{color:#B03A2E} .rg.warn2 .rgk{color:var(--warm)}
:root[data-theme="dark"] .rg.good .rgk{color:#84C08D}
:root[data-theme="dark"] .rg.bad .rgk{color:#E8897C}
@media (prefers-color-scheme: dark){:root:not([data-theme="light"]) .rg.good .rgk{color:#84C08D}
  :root:not([data-theme="light"]) .rg.bad .rgk{color:#E8897C}}
.rgx{margin:0; font-size:13.5px; color:var(--ink)}
.rgn{margin:0; grid-column:2; font-size:12px; color:var(--ink-3); line-height:1.75}
@media(max-width:620px){.rg{grid-template-columns:1fr} .rgn{grid-column:1}}

.ttl{min-width:760px}
.ttl .tsep td{border-top:2px solid var(--rule-2)}
.ttl .kn{width:300px}

/* ---------- ケース別の営業 ---------- */
.sub-h2{font-size:14px; margin:30px 0 12px; font-family:"Zen Old Mincho",serif; font-weight:600;
  color:var(--ink); display:flex; align-items:baseline; gap:10px}
.sub-h2::after{content:""; flex:1; height:1px; background:var(--rule)}
.timeline{display:flex; gap:2px; margin:22px 0 18px; min-height:56px}
.tseg{padding:9px 12px; display:flex; flex-direction:column; gap:3px; justify-content:center; min-width:0}
.tseg.g1{background:var(--l1); color:var(--l1fg)} .tseg.g2{background:var(--l2); color:var(--l2fg)}
.tseg.g3{background:var(--l3); color:var(--l3fg)} .tseg.g4{background:var(--l4); color:var(--l4fg)}
.tseg.g5{background:var(--l5); color:var(--l5fg)}
.tsl{font-family:"IBM Plex Mono",monospace; font-size:10.5px; opacity:.8}
.tsh{font-size:12px; font-weight:700; line-height:1.4}
@media(max-width:760px){.timeline{flex-direction:column} .tseg{flex:none!important}}

.ccards{list-style:none; margin:0; padding:0; display:grid; gap:2px}
.ccard{background:var(--surface); border:1px solid var(--rule); padding:16px 20px;
  display:grid; grid-template-columns:82px 1fr; gap:18px}
.ctime{margin:0; font-family:"IBM Plex Mono",monospace; font-size:12px; color:var(--accent); font-weight:500}
.cbody{display:grid; gap:9px; min-width:0}
.cbody h4{margin:0; font-family:"Zen Old Mincho",serif; font-size:16px; font-weight:600}
.caim{margin:0; font-size:12.5px; color:var(--ink-2)}
@media(max-width:620px){.ccard{grid-template-columns:1fr; gap:8px}}

.qalist{list-style:none; margin:0; padding:0; display:grid; gap:2px}
.qa{background:var(--surface); border:1px solid var(--rule); border-left:5px solid var(--rule-2);
  padding:14px 20px; display:grid; gap:8px}
.qq{margin:0; font-size:13.5px; font-weight:700; color:var(--ink)}
.qn{margin:0; font-size:12.5px; color:var(--ink-3)}
.nglist{list-style:none; margin:0; padding:0; display:grid; gap:2px}
.nglist li{background:var(--surface); border:1px solid var(--rule); border-left:5px solid #B03A2E;
  padding:11px 18px; font-size:13px; color:var(--ink-2)}

.plans{display:grid; grid-template-columns:1fr 1fr; gap:2px}
@media(max-width:900px){.plans{grid-template-columns:1fr}}
.plan{background:var(--surface); border:1px solid var(--rule); border-top:4px solid var(--pc);
  padding:20px 22px 22px; min-width:0}
.plan h4{margin:0 0 10px; font-family:"Zen Old Mincho",serif; font-size:18px; font-weight:600}
.pintro{margin:0 0 18px; font-size:13px; color:var(--ink-2); padding-bottom:16px;
  border-bottom:1px solid var(--rule)}
.pintro b{color:var(--ink)}
.pllist{list-style:none; margin:0; padding:0; display:grid; gap:14px}
.pl{display:grid; grid-template-columns:78px 1fr; gap:14px; align-items:start}
.plk{font-size:10.5px; padding:3px 0; text-align:center; border-radius:2px; white-space:nowrap;
  background:var(--surface-2); color:var(--ink-2); border:1px solid var(--rule-2)}
.plk[data-k^="狙い目"]{background:#B03A2E; color:#fff; border-color:#B03A2E}
.plk[data-k="財源"]{background:var(--warm); color:var(--paper); border-color:var(--warm)}
.plk[data-k="タイムライン"]{background:var(--accent); color:var(--paper); border-color:var(--accent)}
.pl h5{margin:0 0 3px; font-size:13.5px; font-weight:700; color:var(--ink)}
.pl p{margin:0; font-size:12.5px; color:var(--ink-2); line-height:1.85}
@media(max-width:560px){.pl{grid-template-columns:1fr; gap:4px} .plk{justify-self:start; padding:3px 10px}}

.ctabs{flex-wrap:wrap}
.ctab{padding:10px 16px 12px}
.ctab span{max-width:150px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap}
.axis{background:var(--warm-soft); border:1px solid var(--warm); border-left-width:5px;
  padding:16px 20px; margin:18px 0 0}
.axis-t{margin:0 0 7px; font-size:11px; letter-spacing:.08em; color:var(--warm); font-weight:700}
.axis-b{margin:0; font-size:13.5px; color:var(--ink); max-width:88ch}
.axis-g{margin:11px 0 0; font-size:12.5px; color:var(--ink-2); padding-top:10px;
  border-top:1px dashed var(--rule-2)}

.mini-h{font-size:14px; margin:0 0 14px; font-family:"Zen Old Mincho",serif; font-weight:600}
.weekly{min-width:980px}
.weekly .wno{width:56px; text-align:center; font-family:"IBM Plex Mono",monospace;
  color:var(--accent); font-weight:500; white-space:nowrap}
.weekly .wth{width:130px; color:var(--ink); font-weight:700}
.weekly .wmsg{width:300px; color:var(--ink)}
.weekly .wmat{width:200px; font-size:11.5px; color:var(--ink-3)}
.weekly .wex{width:250px; background:color-mix(in srgb,var(--accent-soft) 55%,transparent)}
.rules{list-style:none; margin:0; padding:0; display:grid; gap:2px}
.rule{background:var(--surface); border:1px solid var(--rule); border-left:5px solid var(--accent);
  padding:13px 18px; display:grid; grid-template-columns:96px 1fr; gap:16px; align-items:baseline}
.rule.is-no{border-left-color:#B03A2E}

.axis.is-hot{background:color-mix(in srgb,#B03A2E 12%,var(--surface)); border-color:#B03A2E}
.axis.is-hot .axis-t{color:#B03A2E}
:root[data-theme="dark"] .axis.is-hot .axis-t{color:#EE9C90}
@media (prefers-color-scheme: dark){:root:not([data-theme="light"]) .axis.is-hot .axis-t{color:#EE9C90}}
.ctab.is-hot span::after{content:"　補正"; color:#B03A2E; font-weight:700}
:root[data-theme="dark"] .ctab.is-hot span::after{color:#EE9C90}
@media (prefers-color-scheme: dark){:root:not([data-theme="light"]) .ctab.is-hot span::after{color:#EE9C90}}
.psub{margin:-6px 0 12px; font-size:12px; color:var(--ink-3); letter-spacing:.03em}
.pgrp{font-family:"IBM Plex Mono",monospace; font-size:10.5px; letter-spacing:.08em; color:var(--ink-3);
  padding:10px 0 4px; border-bottom:1px solid var(--rule); margin-bottom:2px}
.pgrp:first-child{padding-top:0}
.weekly .wax{width:64px; text-align:center; font-size:10.5px; white-space:nowrap; font-weight:700}
.weekly tr.ax-a .wax{background:#2CA02C; color:#fff}
.weekly tr.ax-b .wax{background:#D6336C; color:#fff}

.rk{font-size:11px; letter-spacing:.06em; color:var(--ink-3)}
.rule.is-no .rk{color:#B03A2E}
.rule h5{margin:0 0 3px; font-size:13.5px; font-weight:700; color:var(--ink)}
.rule p{margin:0; font-size:12.5px; color:var(--ink-2); line-height:1.85}
@media(max-width:620px){.rule{grid-template-columns:1fr; gap:3px}}




@media(max-width:620px){.tmeta>div{grid-template-columns:1fr; gap:2px}
  .speech{font-size:14px} .hq{flex-wrap:wrap}}
"""

JS = """
(function(){
  var tabs=[].slice.call(document.querySelectorAll('.ytab[data-y]'));
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

  var chips=[].slice.call(document.querySelectorAll('.kchip[data-f]')),
      rows=[].slice.call(document.querySelectorAll('.krow'));
  chips.forEach(function(c){
    c.addEventListener('click',function(){
      chips.forEach(function(o){o.classList.toggle('is-on',o===c);});
      var f=c.dataset.f;
      rows.forEach(function(r){ r.hidden = !(f==='all' || r.dataset.t===f); });
    });
  });

  var tt=[].slice.call(document.querySelectorAll('.ytab[data-t]'));
  tt.forEach(function(t){
    t.addEventListener('click',function(){
      tt.forEach(function(o){
        var on=o===t;
        o.classList.toggle('is-on',on);
        o.setAttribute('aria-selected',on?'true':'false');
        document.getElementById('t'+o.dataset.t).hidden=!on;
      });
    });
  });
  var hc=[].slice.call(document.querySelectorAll('.kchip[data-g]')),
      hr=[].slice.call(document.querySelectorAll('.hrow'));
  hc.forEach(function(c){
    c.addEventListener('click',function(){
      hc.forEach(function(o){o.classList.toggle('is-on',o===c);});
      var g=c.dataset.g;
      hr.forEach(function(r){ r.hidden = !(g==='all' || r.dataset.h===g); });
    });
  });

  var ct=[].slice.call(document.querySelectorAll('.ytab[data-c]'));
  ct.forEach(function(t){
    t.addEventListener('click',function(){
      ct.forEach(function(o){
        var on=o===t;
        o.classList.toggle('is-on',on);
        o.setAttribute('aria-selected',on?'true':'false');
        document.getElementById('cs'+o.dataset.c).hidden=!on;
      });
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
      <li><b>{len(HEARING_COMMON) + len(HEARING_THEME)}</b><span>ヒアリング項目</span></li>
      <li><b>33</b><span>評価項目</span></li>
    </ul>
  </div>
</header>

<nav class="nav" aria-label="目次">
  <div class="nav-in">
    <a href="#levels"><span class="n">01</span>階層</a>
    <a href="#roadmap"><span class="n">02</span>ロードマップ</a>
    <a href="#themes"><span class="n">03</span>到達基準</a>
    <a href="#sales"><span class="n">04</span>提案の深度</a>
    <a href="#talk"><span class="n">05</span>ヒアリング</a>
    <a href="#csv"><span class="n">06</span>記録・CSV</a>
    <a href="#case"><span class="n">07</span>ケース別</a>
    <a href="#calendar"><span class="n">08</span>年間の動き</a>
    <a href="#rubric"><span class="n">09</span>ルーブリック</a>
    <a href="#ops"><span class="n">10</span>運用・研修</a>
  </div>
</nav>

<main class="wrap">
  {sec_levels()}
  {sec_roadmap()}
  {sec_themes()}
  {sec_sales()}
  {sec_talk()}
  {sec_csv()}
  {sec_case()}
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
