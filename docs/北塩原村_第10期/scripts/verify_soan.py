# -*- coding: utf-8 -*-
"""計画素案の数値の自己点検

【背景】
  素案（scripts/soan_content.py）の数値は、算定スクリプトの出力を書き写したもので
  ある。算定を改めたときに素案が追随しないと、両者が静かにずれる。
  他案件（大雪地区広域連合）では素案の生成が算定を直接読む形に改められたが、
  本村は当面この点検により乖離を検出する。

【使い方】
  python3 scripts/verify_soan.py
  すべて適合すれば終了コード0、1件でも不適合があれば1を返す。
"""
import csv, json, os, re, sys, subprocess
sys.dont_write_bytecode = True   # 古いバイトコードで誤った結果が出ることを防ぐ

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(BASE, 'data')
MD = os.path.join(BASE, '18_計画素案.md')

RESULTS = []


def chk(no, name, ok, detail=''):
    RESULTS.append((no, name, ok, detail))


def num(v):
    try:
        return float(str(v).replace(',', '').replace('円', '').replace('人', '')
                     .replace('千', '').replace('%', '').replace('／月', '')
                     .replace('▲', '-').replace('＋', '').strip())
    except (TypeError, ValueError):
        return None


def has_num(t, v):
    """素案の本文に数値vが（3桁区切りの有無を問わず）現れるか"""
    n = num(v)
    if n is None:
        return False
    i = int(round(n))
    return f'{i:,}' in t or str(i) in t


def load_csv(name):
    path = os.path.join(DATA, name)
    if not os.path.exists(path):
        return None
    with open(path, encoding='utf-8-sig') as f:
        return list(csv.DictReader(f))


def md_text():
    with open(MD, encoding='utf-8') as f:
        return f.read()


def in_md(s, *needles):
    """素案の本文に needles がすべて含まれるか"""
    return all(n in s for n in needles)


def main():
    if not os.path.exists(MD):
        print('18_計画素案.md がありません。build_soan.py を先に実行してください。')
        return 1
    t = md_text()

    # ── 1〜3　人口と認定者数（案Bと素案5-2・5-3の一致）────────────
    cert = load_csv('第10期_要介護度別認定者数推計.csv') or []
    b = {r['年']: r for r in cert if r['案'] == 'B' and r['認定率の基準'] == '令和8年'}
    for i, y in enumerate(['令和9年', '令和10年', '令和11年'], start=1):
        r = b.get(y)
        if not r:
            chk(i, f'5-2 第1号被保険者数（{y}）', False, '算定CSVに行がない')
            continue
        ins, ninntei = r['第1号被保険者計'], r['認定者計']
        ok = has_num(t, ins) and has_num(t, ninntei)
        chk(i, f'5-2・5-3 {y}（第1号{int(num(ins)):,}人・認定者{ninntei}人）', ok)

    # ── 4　区分別受給者数（令和11年度）────────────────────
    kubun = load_csv('第10期_サービス区分別受給者数推計.csv') or []
    tot = {r['区分']: r['令和11年度'] for r in kubun if r['要介護度'] == '計'}
    ok = all(has_num(t, v) for v in tot.values()) and len(tot) == 3
    chk(4, '5-4 区分別受給者数の令和11年度（在宅%s・居住系%s・施設%s）'
        % (tot.get('在宅サービス'), tot.get('居住系サービス'), tot.get('施設サービス')), ok)

    # ── 5　サービス諸費と標準給付費の内部整合 ──────────────────
    svc = {'令和9年度': 263470, '令和10年度': 263505, '令和11年度': 262488}
    hojo = 21517
    ok = True
    detail = []
    for y, v in svc.items():
        std = v + hojo
        if not has_num(t, std):
            ok = False
            detail.append(f'{y}の標準給付費{std:,}が素案にない')
    if not has_num(t, sum(svc.values()) + hojo*3):
        ok = False
        detail.append('3か年計854,014が素案にない')
    chk(5, '5-6 標準給付費＝サービス諸費＋補助給付等（年度別と3か年計）', ok, '／'.join(detail))

    # ── 6　介護給付費＋予防給付費＝サービス諸費 ─────────────────
    split = load_csv('第10期_介護給付費_予防給付費の分離.csv') or []
    ok, detail = True, []
    for r in split:
        y = r['年度']
        if y not in svc:
            continue
        kaigo, yobo, total = num(r['介護給付費(千円)']), num(r['予防給付費(千円)']), num(r['サービス諸費(千円)'])
        if kaigo is None or yobo is None:
            continue
        if abs(kaigo + yobo - total) > 1:
            ok = False
            detail.append(f'{y}の内訳の合計がサービス諸費と合わない')
        if not has_num(t, kaigo) or not has_num(t, yobo):
            ok = False
            detail.append(f'{y}の内訳が素案にない')
    chk(6, '5-6 介護給付費＋予防給付費＝サービス諸費（素案と算定の一致）', ok, '／'.join(detail))

    # ── 7　地域支援事業費 ─────────────────────────
    ch = load_csv('第10期_地域支援事業費見込み.csv') or []
    row = next((r for r in ch if r['区分'] == '合計'), None)
    ok, detail = False, ''
    if row:
        vals = [row.get(k) for k in row if k.startswith('令和')]
        ok = all(has_num(t, v) for v in vals if v)
        three = sum(num(v) or 0 for v in vals)
        if not has_num(t, three):
            ok = False
            detail = f'3か年計{int(three):,}千円が素案にない'
    chk(7, '5-5(2) 地域支援事業費（年度別と3か年計）', ok, detail)

    # ── 8　年度別の保険料収納必要額 ──────────────────────
    yr = load_csv('第10期_保険料の年度別必要額.csv') or []
    ok, detail = bool(yr), []
    for r in yr:
        for key, lab in [('保険料収納必要額(千円)', '必要額'), ('月額(円)', '月額')]:
            v = num(r[key])
            if v is None:
                continue
            if not has_num(t, v):
                ok = False
                detail.append(f"{r['年度']}の{lab}{int(v):,}が素案にない")
    chk(8, '5-7 年度別の保険料収納必要額と月額', ok, '／'.join(detail))

    # ── 9　中長期の保険料 ─────────────────────────
    lt = load_csv('第10期_中長期保険料見通し.csv') or []
    ok, detail = bool(lt), []
    for r in lt:
        if not r['区分'].startswith('保険料基準額'):
            continue
        for k, v in r.items():
            if k == '区分' or not v:
                continue
            if not has_num(t, v):
                ok = False
                detail.append(f'{r["区分"]}の{k}={v}が素案にない')
    chk(9, '5-11 中長期の保険料基準額（αの2案）', ok, '／'.join(detail))

    # ── 10　村内定員の到達率 ───────────────────────
    sup = load_csv('第10期_村内定員の到達率.csv') or []
    ok = bool(sup)
    for r in sup:
        for k, v in r.items():
            if not k.endswith('の到達率(%)') or not v:
                continue
            if f'{float(v):.1f}%' not in t:
                ok = False
    chk(10, '5-4(5) 村内定員に対する到達率', ok)

    # ── 11　利用率の分母の検算が素案に反映されているか ───────────
    chk(11, '5-4 利用率の分母（要介護度別認定者数であること）',
        in_md(t, '要介護度別認定者数', '第1号被保険者数', '0件'))

    # ── 12　単価の趨勢と報酬改定率の二重計上の断り ────────────
    chk(12, '5-4(10)① 趨勢と改定率を重ねない旨の注記', in_md(t, '二重に数える'))

    # ── 13　推計の誤差と6-4の見直しの目安の整合 ─────────────
    chk(13, '6-4 見直しの目安（±12%が推計の誤差11.9%に対応）',
        in_md(t, '11.9%', '±12%超'))

    # ── 14　認知症基本法の条項 ──────────────────────
    chk(14, '4-7 認知症基本法の条項（第14条〜第21条の8項目）',
        in_md(t, '第14条から第21条までの8項目'))

    # ── 15　施策番号と節番号の衝突がないこと ─────────────────
    heads = re.findall(r'^## (\S+?)[　 ]', t, re.M)
    dup = {h for h in heads if heads.count(h) > 1}
    chk(15, '節番号の重複がないこと', not dup, '重複: ' + '・'.join(sorted(dup)) if dup else '')

    # ── 16　据え置き項目の一覧があること ───────────────────
    chk(16, '5-12 据え置き項目と確定の条件の一覧', in_md(t, '現時点で据え置いた項目と確定の条件'))

    # ── 17　3-5 施策体系の対応表が第4章の施策と一致すること ────────
    import soan_content as SC
    ch = {c["no"]: c for c in SC.CH}
    sis = []
    for sec in ch["第4章"]["sections"]:
        if sec["no"].startswith("施策"):
            ti = sec["title"]
            sis.append((f'{sec["no"].replace("施策", "")} {ti.split("　【")[0]}',
                        ti.split("【")[1].rstrip("】")))
        elif sec["no"] == "4-5":
            for b in sec["blocks"]:
                if b["t"] == "table":
                    sis += [(r[0], "新設") for r in b["rows"]]
    t35 = [x for x in ch["第3章"]["sections"] if x["no"] == "3-5"]
    if not t35:
        chk(17, '3-5 施策体系の対応表と第4章の施策の一致', False, '3-5 がない')
    else:
        rows = t35[0]["blocks"][1]["rows"]
        bad = [f'{a}≠{r[0]}' for (a, k), r in zip(sis, rows) if r[0] != a or r[1] != k]
        chk(17, '3-5 施策体系の対応表と第4章の施策の一致',
            len(rows) == len(sis) and not bad,
            f'件数 {len(rows)}／{len(sis)}・' + '・'.join(bad[:3]) if (bad or len(rows) != len(sis))
            else f'施策{len(sis)}本すべて一致')

    # ── 18　2-8 目標別の内訳が交付金の集計表の算定と一致すること ──────
    import csv as _csv
    f18 = '/home/user/repository/docs/北塩原村_第10期/data/交付金_目標別の県内比較_令和8年度.csv'
    if not os.path.exists(f18):
        chk(18, '2-8 目標別の内訳と交付金の算定の一致', False, 'CSVがない（parse_kofukin_shichoson.py を実行）')
    else:
        D = {r['区分']: r for r in _csv.DictReader(open(f18, encoding='utf-8-sig'))}
        t28 = None
        for sec in ch['第2章']['sections']:
            if sec['no'] == '2-8':
                for b in sec['blocks']:
                    if b['t'] == 'table' and b['head'] == ['交付金・目標', '本村', '県平均', '県内順位', '全国平均', '全国順位']:
                        t28 = b
        bad = []
        if t28 is None:
            bad.append('2-8の目標別表がない')
        else:
            for row in t28['rows']:
                kf = ('推進' if row[0].startswith('推進') else '支援') + row[0].split('目標')[1][0]
                d = D.get(kf)
                if not d:
                    bad.append(f'{kf}が算定にない'); continue
                for i, k in ((1, '本村'), (2, '県平均'), (3, '県内順位'), (4, '全国平均'), (5, '全国順位')):
                    got = row[i].replace('点', '').replace('位', '').replace(',', '')
                    exp = str(d[k])
                    if float(got) != float(exp):
                        bad.append(f'{kf} {k} 素案{row[i]}≠算定{d[k]}')
        chk(18, '2-8 目標別の内訳と交付金の算定の一致', not bad,
            '・'.join(bad[:4]) if bad else f'8目標×5項目すべて一致')

    # ── 19　6-3 の取得をめざす評価指標の表が61点で、対応先が実在すること ──
    t63 = None
    for sec in ch['第6章']['sections']:
        if sec['no'] == '6-3':
            for b in sec['blocks']:
                if b['t'] == 'table' and b['head'][0] == '交付金・指標':
                    t63 = b
    if t63 is None:
        chk(19, '6-3 取得をめざす評価指標の表', False, '表がない')
    else:
        rows, last = t63['rows'][:-1], t63['rows'][-1]
        tot = sum(int(r[1].replace('点', '')) for r in rows)
        # 「本計画での対応」に挙げた節・施策がすべて実在すること
        nos = set()
        for c in SC.CH:
            for sec in c['sections']:
                nos.add(sec['no'])
        miss = sorted({x.strip() for r in rows for x in r[4].split('・')} - nos)
        chk(19, '6-3 取得をめざす評価指標の表',
            tot == 61 and last[1] == '61点' and not miss,
            f'合計{tot}点／表記{last[1]}' + (f'・対応先が不明 {miss}' if miss else '')
            if (tot != 61 or last[1] != '61点' or miss) else f'{len(rows)}項目・61点、対応先すべて実在')

    # ── 20　交付金の11項目が第4章・第5章に書かれていること ────────────
    md = t
    key20 = [('4つの場面ごとの目指すべき姿', '施策4-3'), ('難聴高齢者の早期発見・早期介入（普及啓発）', '施策4-4'),
             ('成年後見制度利用支援事業の実施要綱', '施策4-7'),
             ('村外施設入居者の実態把握', '5-8')]
    bad20 = [f'{k}（{w}）' for k, w in key20 if k not in md]
    chk(20, '交付金で取得をめざす記述が本文にあること', not bad20,
        '欠落: ' + '・'.join(bad20) if bad20 else f'{len(key20)}件すべて記載')

    # ── 21　3-4 の施策体系図が登録され、ファイルが実在すること ──────────
    from figures_map import FIGS
    f21 = [x for x in FIGS if x[0] == "第3章|3-4"]
    fp = "/home/user/repository/output/figures/fig3-1_施策体系図.png"
    chk(21, "3-4 施策体系図の登録とファイルの実在",
        len(f21) == 1 and os.path.exists(fp),
        "図表マップに未登録" if len(f21) != 1 else ("PNGがない（build_figure_taikei.py を実行）"
                                              if not os.path.exists(fp) else "図3-1 を登録・実在"))

    # ── 22　第1号被保険者数に同じ時点で2つの値が併存しないこと ──────────
    import shihyo_dict as SD
    bad22 = []
    # 2-9 の乖離表（見える化の年度の値）
    t29 = None
    for sec in ch['第2章']['sections']:
        if sec['no'] == '2-9':
            for b in sec['blocks']:
                if b['t'] == 'table' and b['head'] == ['', '計画値', '実績', '差']:
                    t29 = t29 or b
    if t29 is None:
        bad22.append('2-9の乖離表がない')
    else:
        for r in t29['rows']:
            yr = r[0].replace('年度', '年度')
            exp = SD.HIHOKENSHA_MIERUKA.get(yr)
            got = int(''.join(c for c in r[2] if c.isdigit()))
            if exp is not None and got != exp:
                bad22.append(f'2-9 {yr} 素案{got}≠見える化{exp}')
            if exp is None and yr == '令和8年度' and got != SD.HIHOKENSHA_BAN['令和8年'][2]:
                bad22.append(f"2-9 {yr} 素案{got}≠B案{SD.HIHOKENSHA_BAN['令和8年'][2]}")
    # 2-2 と 5-2 の推計表が B案と一致すること
    for chn, secno in (('第2章', '2-2'), ('第5章', '5-2')):
        for sec in ch[chn]['sections']:
            if sec['no'] != secno:
                continue
            for b in sec['blocks']:
                if b['t'] != 'table' or b['head'][:2] != ['', '前期高齢者']:
                    continue
                for r in b['rows']:
                    key = r[0].replace('（実績）', '').replace('令和7年12月末', '令和8年')
                    if key not in SD.HIHOKENSHA_BAN:
                        continue
                    z, k, tot = SD.HIHOKENSHA_BAN[key]
                    got = int(r[3].replace(',', ''))
                    if got != tot:
                        bad22.append(f'{secno} {r[0]} 素案{got}≠B案{tot}')
    chk(22, '第1号被保険者数の値が出所ごとに一致すること', not bad22,
        '・'.join(bad22[:4]) if bad22 else '2-2・2-9・5-2 の全行が出所の値と一致')

    # ── 23　指標辞書の値が素案に書かれていること ─────────────────
    def num_in(v):
        s1 = f'{v:,}' if isinstance(v, int) and v >= 1000 else str(v)
        return s1 in t or str(v) in t
    bad23 = [k for k, (v, *_rest) in SD.S.items() if not num_in(v)]
    chk(23, '指標辞書の値が素案に記載されていること', not bad23,
        '素案にない: ' + '・'.join(bad23[:4]) if bad23 else f'{len(SD.S)}件すべて記載')

    # ── 24　出所の異なる同じ名前の指標に、時点の断りがあること ─────────
    need24 = [('1,009', '65歳以上人口'), ('1,012', '令和7年12月末'),
              ('1,015', '令和6年度'), ('1,011', '令和7年度')]
    bad24 = [f'{n}（{w}）' for n, w in need24 if n in t and w not in t]
    chk(24, '第1号被保険者数の3つの出所に時点の断りがあること', not bad24,
        '断りがない: ' + '・'.join(bad24) if bad24 else '4件とも時点を明記')

    # ── 25　2-3 の認定者数（各年3月末）が見える化の実データと一致すること ──
    #    他案件（金ケ崎町）で「推計の系列が3通りあった」ことに倣い、
    #    素案に載せた実績は必ず原典の系列から引く。
    MI = os.path.join(DATA, 'mieruka_tidy.csv')
    b3 = {}
    if os.path.exists(MI):
        for r in csv.DictReader(open(MI, encoding='utf-8-sig')):
            if (r['code'] == 'B3-a' and r['indicator'] == '合計認定者数'
                    and r['region'] == '北塩原村'):
                b3[r['period'].replace('時点', '')] = int(float(r['value']))
    t23 = None
    for sec in SC.CH[1]['sections']:
        if sec['no'] == '2-3':
            for b in sec['blocks']:
                if b['t'] == 'table' and b['head'][0] == '時点' and '合計' in b['head']:
                    t23 = b
    bad25 = []
    if not b3 or t23 is None:
        bad25.append('原典または2-3の表がない')
    else:
        for r in t23['rows']:
            k, v = r[0], int(r[-1])
            if k not in b3:
                bad25.append(f'{k}が原典にない')
            elif b3[k] != v:
                bad25.append(f'{k} 素案{v}≠原典{b3[k]}')
    chk(25, '2-3 の認定者数が見える化 B3-a と一致すること', not bad25,
        '・'.join(bad25[:3]) if bad25 else f'{len(t23["rows"]) if t23 else 0}時点すべて一致')

    # ── 26　5-3 の要介護度別認定者数が再計算と一致し、計が四捨五入前の和であること ──
    S53 = {sec['no']: sec for sec in SC.CH[4]['sections']}['5-3']
    tr = [b for b in S53['blocks'] if b['t'] == 'table' and b['head'][0] == '']
    trate = None
    for sec in SC.CH[4]['sections']:
        for b in sec['blocks']:
            if (b['t'] == 'table' and b['head'][:3] == ['区分', '前期高齢者', '後期高齢者']
                    and any(r[0] == '要支援1' for r in b['rows'])):
                trate = b
    import shihyo_dict as SD2
    BAN = SD2.HIHOKENSHA_BAN
    LV = ['要支援1', '要支援2', '要介護1', '要介護2', '要介護3', '要介護4', '要介護5']
    bad26 = []
    if not tr or trate is None:
        bad26.append('5-3の表または認定率の表がない')
    else:
        rate = {r[0]: (num(r[1]) / 100, num(r[2]) / 100) for r in trate['rows'] if r[0] in LV}
        tb = tr[0]
        for yi, y in enumerate(tb['head'][1:], start=1):
            if y not in BAN:
                continue
            zen, kou, _tot = BAN[y]
            raw = {lv: zen * rate[lv][0] + kou * rate[lv][1] for lv in LV}
            row = [r for r in tb['rows'] if r[0] == y]
            if not row:
                continue
            got = row[0]
            for i, lv in enumerate(LV, start=1):
                if int(got[i]) != round(raw[lv]):
                    bad26.append(f'{y} {lv} 素案{got[i]}≠再計算{round(raw[lv])}')
            if int(got[8]) != round(sum(raw.values())):
                bad26.append(f'{y} 計 素案{got[8]}≠四捨五入前の和{round(sum(raw.values()))}')
    ok26 = not bad26 and in_md(t, '四捨五入する前に合計')
    chk(26, '5-3 の認定者数が認定率×人口の再計算と一致すること', ok26,
        '・'.join(bad26[:3]) if bad26 else '令和9〜22年の要介護度別と計が再計算と一致（端数の断りあり）')

    # ── 27　5-4 の（参考）認定者数が各年度末の値で、年率が再計算と一致すること ──
    t27 = None
    for sec in SC.CH[4]['sections']:
        for b in sec['blocks']:
            if b['t'] == 'table' and b['head'][:2] == ['サービス', '令和2年度']:
                t27 = b
    bad27 = []
    if t27 is None or not b3:
        bad27.append('5-4の表または原典がない')
    else:
        row = [r for r in t27['rows'] if '認定者数' in r[0]]
        if not row:
            bad27.append('（参考）認定者数の行がない')
        else:
            a, b_ = int(row[0][1]), int(row[0][2])
            gs = str(row[0][3]).replace('＋', '').replace('%', '')
            g = float(gs)
            # 令和2年度末＝令和3年3月末、令和7年度末＝令和8年3月末
            wa, wb = b3.get('令和3年3月末'), b3.get('令和8年3月末')
            if (a, b_) != (wa, wb):
                bad27.append(f'素案{a}／{b_} ≠ 年度末{wa}／{wb}')
            else:
                r5 = ((b_ / a) ** (1 / 5) - 1) * 100
                if abs(r5 - g) > 0.01:
                    bad27.append(f'年率 素案{g}%≠再計算{r5:.2f}%')
                if f'＋{gs}%' not in t:
                    bad27.append('本文の年率が表と一致しない')
    chk(27, '5-4 の認定者数の伸びが年度末の系列と一致すること', not bad27,
        '・'.join(bad27[:3]) if bad27 else '令和2年度末186人→令和7年度末214人・年率＋2.84%')

    # ── 28　成果品の本文に受託者の内部の仕組みの語がないこと ─────────
    #    他案件（大雪地区広域連合）で発注者に提供する文章から内部の仕組み・
    #    作業経過の語を落とす是正があったことに倣う。
    NAIGO = ['scripts/', '.py', 'verify_', '.csv', 'Python', 'GitHub', 'コミット']
    hit28 = []
    for c in SC.CH:
        for sec in c['sections']:
            for b in sec['blocks']:
                for v in ([str(b.get('v', ''))] +
                          [str(x) for r in b.get('rows', []) for x in r] +
                          [str(x) for x in b.get('head', [])]):
                    for g in NAIGO:
                        if g in v:
                            hit28.append(f'{sec["no"]}:{g}')
    chk(28, '素案の本文に内部の仕組みの語がないこと', not hit28,
        '・'.join(sorted(set(hit28))[:4]) if hit28 else f'{len(NAIGO)}語のいずれも本文にない')

    # ── 29　成果品の地の文にマークダウンの記号が残っていないこと ───────
    #    docxビルダは ** を太字に変換しないため、そのまま紙に出てしまう。
    hit29 = []
    for c in SC.CH:
        for sec in c['sections']:
            for b in sec['blocks']:
                for v in ([str(b.get('v', ''))]
                          + [str(x) for r in b.get('rows', []) for x in r]
                          + [str(x) for x in b.get('head', [])]):
                    for g in ('**', '__', '~~', '](' ):
                        if g in v:
                            hit29.append(f'{sec["no"]}:{g}')
    chk(29, '本文にマークダウンの記号が残っていないこと', not hit29,
        '・'.join(sorted(set(hit29))[:4]) if hit29 else '**・__・~~・リンク記法のいずれもない')

    # ── 30　2-8 の目標別得点が交付金の算定と一致すること ─────────────
    KF = os.path.join(DATA, '交付金_目標別の県内比較_令和8年度.csv')
    t30 = None
    for sec in SC.CH[1]['sections']:
        if sec['no'] == '2-8':
            for b in sec['blocks']:
                if (b['t'] == 'table' and b['head'][0] == '交付金・目標'
                        and '県内順位' in b['head'] and '全国順位' in b['head']):
                    t30 = b
    bad30 = []
    if not os.path.exists(KF) or t30 is None:
        bad30.append('算定のCSVまたは2-8の表がない')
    else:
        D = {r['目標名']: r for r in csv.DictReader(open(KF, encoding='utf-8-sig'))}
        for r in t30['rows']:
            name = r[0].split('\u3000', 1)[-1].strip()
            d = D.get(name)
            if not d:
                bad30.append(f'{name[:12]}が算定にない'); continue
            for i, key in ((1, '本村'), (2, '県平均'), (3, '県内順位'),
                           (4, '全国平均'), (5, '全国順位')):
                a = float(str(r[i]).replace('点', '').replace('位', '').replace(',', ''))
                if abs(a - float(d[key])) > 0.05:
                    bad30.append(f'{name[:8]} {key} 素案{r[i]}≠算定{d[key]}')
    chk(30, '2-8 の目標別得点が交付金の算定と一致すること', not bad30,
        '・'.join(bad30[:3]) if bad30 else '8目標×5項目すべて一致')

    # ── 31　交付金の満点の定義が一貫していること ───────────────────
    need31 = ['800点満点', '成果指向型配分枠', '評価指標による満点']
    miss31 = [x for x in need31 if x not in t]
    bad31 = []
    if '160点' in t and '100点満点で上限' not in t and '満点は100点' not in t:
        bad31.append('目標Ⅳの満点の断りがない')
    chk(31, '交付金の満点の定義が一貫していること', not miss31 and not bad31,
        ('欠落: ' + '・'.join(miss31) if miss31 else '') + '・'.join(bad31)
        if (miss31 or bad31) else '800点は評価指標による満点。成果指向型配分枠は別枠と明記')

    # ── 32　給付適正化の主要3事業の名称が節をまたいで一致すること ─────────
    SAN = ['要介護認定の適正化', 'ケアプラン等の点検', '縦覧点検・医療情報との突合']
    where32 = {}
    for c in SC.CH:
        for sec in c['sections']:
            txt = ''.join([str(b.get('v', ''))
                           + str(b.get('rows', '')) + str(b.get('head', ''))
                           for b in sec['blocks']])
            if '主要3事業' in txt or '主要３事業' in txt:
                where32[sec['no']] = [x for x in SAN if x in txt]
    bad32 = [f'{k}（{len(v)}/3）' for k, v in where32.items() if len(v) not in (0, 3)]
    chk(32, '主要3事業の名称が節をまたいで一致すること', len(where32) >= 3 and not bad32,
        '一部しか一致しない: ' + '・'.join(bad32) if bad32
        else f'主要3事業に触れる{len(where32)}節で名称が一致')

    # ── 33　5-4(12) の需要に対する供給の表が再計算と一致すること ─────────
    import estimate_jinzai as EJ
    t33 = None
    for sec in SC.CH[4]['sections']:
        for b in sec['blocks']:
            if (b['t'] == 'table' and b['head'][0] == ''
                    and b['rows'] and b['rows'][0][0] == '生産年齢人口'):
                t33 = b
    bad33 = []
    if t33 is None:
        bad33.append('5-4(12)の表がない')
    else:
        calc = {r[0]: r for r in EJ.run()}
        ROW = {'生産年齢人口': 1, '認定者数（需要）': 2, '　うち要支援1・2': 3,
               '　うち要介護1以上': 4, '供給の枠（生産年齢人口に比例）': 5,
               '要支援の方に回せる量': 6, '要支援の充足率': 7}
        for row in t33['rows']:
            idx = ROW.get(row[0])
            if idx is None:
                bad33.append(f'{row[0]}が再計算にない'); continue
            for j, y in enumerate(t33['head'][1:], start=1):
                got = str(row[j]).replace('人', '').replace('%', '').replace(',', '')
                exp = calc[y][idx]
                if abs(float(got) - float(exp)) > 0.05:
                    bad33.append(f'{y} {row[0]} 素案{row[j]}≠再計算{exp}')
    chk(33, '5-4(12) 需要に対する供給が再計算と一致すること', not bad33,
        '・'.join(bad33[:3]) if bad33 else '4時点×7行すべて一致（要支援の充足率100.0→18.6%）')

    # ── 34　5-6 の補助給付等の比率と感応度が算定と一致すること ─────────
    KAN = load_csv('第10期_保険料感応度.csv')
    t34 = None
    for sec in SC.CH[4]['sections']:
        for b in sec['blocks']:
            if b['t'] == 'table' and b['head'][:2] == ['前提', '標準給付費への差']:
                t34 = b
    bad34 = []
    if t34 is None or not KAN:
        bad34.append('5-6の表または感応度のCSVがない')
    else:
        D = {r['項目']: r for r in KAN if r.get('項目')}
        for row in t34['rows']:
            key = [k for k in D if row[0][:12] in k or k.startswith(row[0][:10])]
            if not key:
                bad34.append(f'{row[0][:16]}が算定にない'); continue
            d = D[key[0]]
            a = float(str(row[1]).replace('＋', '').replace(',', '').replace('千円', ''))
            if abs(a * 1000 - float(d['給付費等の増減(円)'])) > 1000:
                bad34.append(f'{row[0][:10]} 差 素案{row[1]}≠算定{d["給付費等の増減(円)"]}')
            m = float(str(row[2]).replace('＋', '').replace('円', ''))
            if abs(m - float(d['月額への効き(円)'])) > 0.5:
                bad34.append(f'{row[0][:10]} 月額 素案{row[2]}≠算定{d["月額への効き(円)"]}')
    ok34 = not bad34 and in_md(t, '対標準給付費比', '9.02%', '7.56%')
    chk(34, '5-6 の補助給付等の比率と感応度が算定と一致すること', ok34,
        '・'.join(bad34[:3]) if bad34 else '比率の系列と感応度2件が算定と一致')

    # ── 35　5-4(11) の施策反映が手引きの3区分で、織り込んでいない旨があること ──
    need35 = ['① 認定者数', '② 施設・居住系サービス', '③ 在宅サービス',
              '施策の効果を織り込んでいません']
    miss35 = [x for x in need35 if x not in t]
    chk(35, '5-4(11) 施策反映が手引きの3区分であること', not miss35,
        '欠落: ' + '・'.join(miss35) if miss35 else '3区分と、効果を織り込んでいない旨を明記')

    # ── 36　計画の積算が仕様書の約100頁に収まること ────────────────
    #    仕様書の約100頁は成果品としての計画にかかる。素案の編集注記（⚙）と
    #    5-12（据え置いている項目の一覧）はいずれも計画の確定時に削除するため、
    #    判定はこの2つを除いた確定版の頁数で行う。素案のままの頁数も併せて示す。
    try:
        import estimate_pages as EP
        import math
        import os as _os
        from figures_map import FIGS as _FIGS
        _figs = {}
        for _e in _FIGS:              # inline の図も同じ節に積む
            _figs.setdefault(_e[0], []).append(_e[1])

        def _pages(skip_notes=False, skip_secs=()):
            tot = 0
            for ch in SC.CH:
                h = 300 + 30 * 20                      # 章見出し
                for sec in ch['sections']:
                    if sec['no'] in skip_secs:
                        continue
                    h += 320 + 25 * 20 + 180           # 節見出し
                    for fn in _figs.get(f'{ch["no"]}|{sec["no"]}', []):
                        w, hh = EP.png_size(_os.path.join(EP.FIGDIR, fn))
                        h += (160 + (6.3 * hh / w) * 1440 + 60
                              + 18 * 20 + 40 + 16 * 20 + 200)
                    for b in sec['blocks']:
                        if skip_notes and b['t'] == 'note':
                            continue
                        h += EP.block_h(b)
                tot += math.ceil(h / EP.BODY_H)        # 章の変わり目で改頁
            return tot + sum(n for _, n in EP.FRONT)

        est = _pages()
        final = _pages(skip_notes=True, skip_secs=('5-12',))
        ok36 = 95 <= final <= 105
        chk(36, '計画の積算が仕様書の約100頁に収まること', ok36,
            f'確定版{final}頁（素案のままは{est}頁。編集注記と5-12で{est - final}頁）'
            + ('' if ok36 else '　約100頁から外れています')) 
    except Exception as e:
        chk(36, '計画の積算が仕様書の約100頁に収まること', False, f'照合できない（{e}）')

    # ── 37　交付金の要改善項目の件数・点数が再計算と一致すること ─────────
    #    素案2-8・6-3 の「32項目・122点」「11項目・61点」を、
    #    3か年の項目別データから作り直して突き合わせる。
    try:
        import analyze_kofukin as AK
        rows_k, K_k = AK.load()
        yk = AK.yokaizen(rows_k, K_k)
        n32, p122 = len(yk), sum(x['配点'] for x in yk)
        kisai = [x for x in yk if x['判定'] == '可（記載）']
        p56 = sum(x['配点'] for x in kisai)
        bad37 = []
        if f'{n32}項目・{p122:.0f}点' not in t:
            bad37.append(f'素案に「{n32}項目・{p122:.0f}点」の記載がない')
        # 6-3 の一覧（11項目61点）は 可（記載）56点＋認知症施策推進計画5点
        t63 = None
        for sec in SC.CH[5]['sections']:
            if sec['no'] == '6-3':
                for b in sec['blocks']:
                    if b['t'] == 'table' and b['head'][0] == '交付金・指標':
                        t63 = b
        if t63 is None:
            bad37.append('6-3の一覧がない')
        else:
            body = t63['rows'][:-1]
            tot = sum(float(str(r[1]).replace('点', '')) for r in body)
            if len(body) != 11 or abs(tot - 61) > 0.5:
                bad37.append(f'6-3の一覧が{len(body)}項目{tot:.0f}点')
            if abs(tot - (p56 + 5)) > 0.5:
                bad37.append(f'6-3の61点が 可（記載）{p56:.0f}点＋認知症計画5点と合わない')
        chk(37, '交付金の要改善項目の件数・点数が再計算と一致すること', not bad37,
            '・'.join(bad37[:3]) if bad37
            else f'{n32}項目{p122:.0f}点／6-3は11項目61点（可（記載）{p56:.0f}点＋5点）')
    except Exception as e:
        chk(37, '交付金の要改善項目の件数・点数が再計算と一致すること', False, f'照合できない（{e}）')

    # ── 38　配布データ（A案）の値が素案の中で1通りであること ────────────
    #    同じ指標を2か所に書いていたため、954人と953人が併存していた。
    A = load_csv('第10期_要介護度別認定者数推計.csv')
    ban = {}
    for r in A:
        if r.get('案') == 'A':
            ban[r['年']] = int(r['第1号被保険者計'])
    bad38 = []
    if not ban:
        bad38.append('A案の算定がない')
    else:
        for y in ('令和8年', '令和9年', '令和10年', '令和11年'):
            v = ban.get(y)
            if v is None:
                continue
            # 算定の値の前後1人の値が本文に出てこないこと（同じ指標に2つの値が併存する型）
            for w2 in (v - 1, v + 1):
                if f'{w2}人' in t:
                    bad38.append(f'{y} 算定{v}人に対し本文に{w2}人がある')
        # 令和8年の値と実績との差が1通りであること
        v8 = ban.get('令和8年')
        if v8 and f'{v8}人' in t:
            sa = 1012 - v8
            if f'{sa}人下回' not in t and f'{sa}人多い' not in t and f'{sa}人' not in t:
                bad38.append(f'令和8年の差{sa}人の記載がない')
            if f'{sa - 1}人（' in t or f'{sa + 1}人（' in t:
                bad38.append('差の人数が2通りある')
    chk(38, '配布データ（A案）の値が素案の中で1通りであること', not bad38,
        '・'.join(bad38[:3]) if bad38
        else f'令和8年{ban.get("令和8年")}人・令和9〜11年{ban.get("令和9年")}／{ban.get("令和10年")}／{ban.get("令和11年")}人で統一')

    # ── 39　資6 に基本指針の別表11事項がすべて掲げられていること ────────
    #    網羅性の根拠となる表なので、事項の脱落と「対応」欄の空白を見る。
    try:
        KAN = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十', '十一']
        t6 = None
        for c in SC.CH:
            for sec in c['sections']:
                if sec['no'] == '資6':
                    for b in sec['blocks']:
                        if b['t'] == 'table' and b['head'][0] == '事項' and '対応' in b['head']:
                            t6 = b
        bad39, part = [], []
        if t6 is None:
            bad39.append('資6の対応表がない')
        else:
            got = [r[0] for r in t6['rows']]
            for k in KAN:
                if k not in got:
                    bad39.append(f'事項{k}がない')
            for r in t6['rows']:
                if not str(r[2]).strip():
                    bad39.append(f'事項{r[0]}の該当箇所が空欄')
                if str(r[3]).strip() not in ('対応', '一部', '未対応'):
                    bad39.append(f'事項{r[0]}の対応欄が「{r[3]}」')
            # 「一部」とした事項は理由と代替の表に必ず現れること
            part = [r[0] for r in t6['rows'] if str(r[3]).strip() == '一部']
            t6b = None
            for c in SC.CH:
                for sec in c['sections']:
                    if sec['no'] == '資6':
                        for b in sec['blocks']:
                            if b['t'] == 'table' and '理由' in b['head']:
                                t6b = b
            if t6b is None and part:
                bad39.append('「一部」の事項があるのに理由の表がない')
            elif t6b is not None:
                got2 = [r[0] for r in t6b['rows']]
                for k in part:
                    if k not in got2:
                        bad39.append(f'事項{k}が一部なのに理由が書かれていない')
        chk(39, '資6 の別表11事項がすべて掲げられていること', not bad39,
            '・'.join(bad39[:3]) if bad39
            else f'11事項すべて掲げ、一部対応{len(part)}件に理由と代替を記載')
    except Exception as e:
        chk(39, '資6 の別表11事項がすべて掲げられていること', False, f'照合できない（{e}）')

    # ── 40　図表番号一覧が本文に現れる順と一致し、番号が連番であること ──────
    #    番号は figures_map.py の並び順から振るため、手書きとのずれを機械で見る。
    try:
        from figures_map import FIGS as F40, index_rows as IR40
        t8 = None
        for c in SC.CH:
            for sec in c['sections']:
                if sec['no'] == '資8':
                    for b in sec['blocks']:
                        if b['t'] == 'table' and b['head'][0] == '図番号':
                            t8 = b
        bad40 = []
        if t8 is None:
            bad40.append('資8の図表番号一覧がない')
        else:
            if [list(r) for r in t8['rows']] != IR40():
                bad40.append('一覧が図の対応表と一致しない')
            # 図の登録順が本文に現れる順（章節の並び）と一致すること。
            # 一覧と対応表はどちらも同じ並びから作るため、両者を比べても
            # 並び自体の誤りは出ない。素案の章節の並びと突き合わせる。
            order = {}
            for ci, c in enumerate(SC.CH):
                for si, sec in enumerate(c['sections']):
                    order[f'{c["no"]}|{sec["no"]}'] = (ci, si)
            pos = []
            for _e in F40:
                key = _e[0]
                if key not in order:
                    bad40.append(f'図の登録先 {key} が素案にない')
                else:
                    pos.append((order[key], key))
            for (a1, k1), (a2, k2) in zip(pos, pos[1:]):
                if a2 < a1:
                    bad40.append(f'図の登録順が本文の順と違う（{k1} の後に {k2}）')
                    break
            # 章ごとに1から連番であること（飛び・重複を見る）
            seq = {}
            for fnum, _title, _loc, _src in t8['rows']:
                ch, n = fnum.replace('図', '').split('-')
                seq.setdefault(ch, []).append(int(n))
            for ch, ns in seq.items():
                if ns != list(range(1, len(ns) + 1)):
                    bad40.append(f'図{ch}-の番号が連番でない（{ns}）')
            # 一覧の件数が登録した図の数と一致すること
            if len(t8['rows']) != len(F40):
                bad40.append(f"一覧{len(t8['rows'])}件と図{len(F40)}件が合わない")
        chk(40, '図表番号一覧が本文の順と一致し番号が連番であること', not bad40,
            '・'.join(bad40[:3]) if bad40
            else f'図{len(F40)}点、章ごとに連番で一覧と一致')
    except Exception as e:
        chk(40, '図表番号一覧が本文の順と一致し番号が連番であること', False, f'照合できない（{e}）')

    # ── 41　制度改正の主な事項が 1-8 に掲げられていること ─────────────
    #    C軸の走査語。保険料・給付費に効くものが落ちると計画の前提が狂う。
    try:
        need41 = ['82.65', '7区分', '激変緩和', '補足給付', '令和8年度介護報酬改定',
                  '一定以上所得', '給与所得控除', '事業状況報告', '高額医療合算',
                  '返還義務', '電子資格確認', '協力医療機関', '地域類型',
                  '特定地域サービス', '入居定員']
        s18 = ''
        for c in SC.CH:
            for sec in c['sections']:
                if sec['no'] == '1-8':
                    for b in sec['blocks']:
                        if b['t'] in ('table', 'kpi'):
                            s18 += ' '.join(b['head']) + ' '
                            s18 += ' '.join(str(x) for r in b['rows'] for x in r) + ' '
                        elif b['t'] == 'bullets':
                            s18 += ' '.join(b['v']) + ' '
                        else:
                            s18 += str(b['v']) + ' '
        miss41 = [k for k in need41 if k not in s18]
        chk(41, '制度改正の主な事項が 1-8 に掲げられていること', not miss41,
            '1-8にない：' + '・'.join(miss41) if miss41
            else f'{len(need41)}件すべて記載')
    except Exception as e:
        chk(41, '制度改正の主な事項が 1-8 に掲げられていること', False, f'照合できない（{e}）')

    # ── 42　資6・資7 が指す章節が現に存在すること ──────────────────
    #    存在しない節を指す「対応しています」を防ぐ。
    try:
        import re as _re
        exist = {sec['no'] for c in SC.CH for sec in c['sections']}
        bad42 = []
        for c in SC.CH:
            for sec in c['sections']:
                if sec['no'] not in ('資6', '資7'):
                    continue
                for b in sec['blocks']:
                    if b['t'] != 'table':
                        continue
                    for r in b['rows']:
                        for cell in r:
                            for ref in _re.findall(r'(?:施策)?\d+-\d+(?:\(\d+\))?', str(cell)):
                                base = ref.split('(')[0]
                                cand = [base, '施策' + base]
                                if not any(x in exist for x in cand):
                                    bad42.append(f'{sec["no"]}が指す{ref}がない')
        bad42 = sorted(set(bad42))
        chk(42, '資6・資7 が指す章節が現に存在すること', not bad42,
            '・'.join(bad42[:4]) if bad42 else '参照先すべて実在')
    except Exception as e:
        chk(42, '資6・資7 が指す章節が現に存在すること', False, f'照合できない（{e}）')

    # ── 43　登録した図がすべて本文に1回だけ現れ、画像が実在すること ────────
    #    inline の図は soan_content の {"t":"fig"} で置くため、置き忘れると
    #    図表番号一覧には載るのに本文から消える。番号だけが残る事故を防ぐ。
    try:
        from figures_map import FIGS as F43
        import os as _os2
        FIGDIR = '/home/user/repository/output/figures'
        placed = {}
        for c in SC.CH:
            for sec in c['sections']:
                for b in sec['blocks']:
                    if b['t'] == 'fig':
                        placed[b['v']] = placed.get(b['v'], 0) + 1
        bad43 = []
        reg = set()
        for _e in F43:
            key, fn = _e[0], _e[1]
            inline = bool(_e[4]) if len(_e) > 4 else False
            reg.add(fn)
            if not _os2.path.exists(_os2.path.join(FIGDIR, fn)):
                bad43.append(f'{fn} の画像がない')
            n = placed.get(fn, 0)
            if inline and n != 1:
                bad43.append(f'{fn} はinlineだが本文に{n}回（1回であること）')
            if not inline and n:
                bad43.append(f'{fn} は節の冒頭に出す図だがfigブロックもある')
        for fn in placed:
            if fn not in reg:
                bad43.append(f'{fn} が図の対応表に登録されていない')
        chk(43, '登録した図がすべて本文に1回だけ現れ画像が実在すること', not bad43,
            '・'.join(bad43[:3]) if bad43
            else f'図{len(reg)}点、うち本文中に置くもの{len(placed)}点、画像すべて実在')
    except Exception as e:
        chk(43, '登録した図がすべて本文に1回だけ現れ画像が実在すること', False,
            f'照合できない（{e}）')

    # ── 44　図の数値の正本が1か所であること ──────────────────────
    #    (a) 作図スクリプトに数値がじか書きされていないこと
    #    (b) 画像の中の表題に図表番号を焼き込んでいないこと
    #        （番号は figures_map.py が振るため、焼き込むと採番のやり直しで食い違う）
    try:
        import re as _re4
        import os as _os4
        _sc = _os4.path.join(_os4.path.dirname(_os4.path.abspath(__file__)))
        with open(_os4.path.join(_sc, 'build_figures.py'), encoding='utf-8') as _f4:
            bf = _f4.read()
        bad44 = []
        # (b) 表題に図表番号がないこと
        yaki = _re4.findall(r'(?:set_title\(|suptitle\()"図[0-9]+-[0-9]+', bf)
        if yaki:
            bad44.append(f'画像の表題に図表番号が{len(yaki)}件焼き込まれている')
        # (a) 数値の並びがじか書きされていないこと（3つ以上の数の並び）
        jika = []
        for m in _re4.finditer(r'^\s*([A-Za-z_][A-Za-z_0-9]*)\s*=\s*\[([^\]]*)\]',
                               bf, _re4.M):
            body = m.group(2)
            nums = _re4.findall(r'-?\d+\.?\d*', body)
            if len(nums) >= 3 and not _re4.search(r'[A-Za-z_]{2,}', body):
                jika.append(m.group(1))
        # 見るのは「計測した値」である。次は軸の位置・区分の名・体裁であり、
        # 正本に持つ対象ではないため除く（区分の名は data_zuhyo にもあるが、
        # 作図は目盛りの位置として数で持つ必要がある）
        LAYOUT = {'x', 'lab', 'labels', 'cat', 'cats', 'names', 'grp', 'yrs',
                  'ORDER', 'ys', 'ylab', 'widths', 'pos', 'bottom', 'b',
                  'CHART', 'FIGSIZE', 'idx', 'y'}
        jika = [v for v in jika if v not in LAYOUT]
        if jika:
            bad44.append('作図に数値がじか書き：' + '・'.join(sorted(set(jika))[:5]))
        # 正本の系列数が作図の参照数と釣り合っていること
        import data_zuhyo as _DZ4
        nref = len(_re4.findall(r'V\("', bf))
        nser = sum(len(d['series']) for d in _DZ4.load(use_book=False).values())
        if nref < nser * 0.8:
            bad44.append(f'正本の系列{nser}に対し作図の参照が{nref}しかない')
        chk(44, '図の数値の正本が1か所であること', not bad44,
            '・'.join(bad44[:3]) if bad44
            else f'じか書きなし、表題に番号なし、正本{nser}系列を参照{nref}件')
    except Exception as e:
        chk(44, '図の数値の正本が1か所であること', False, f'照合できない（{e}）')

    # ── 45　2-8 の目標別の推移が算定と一致し、計が交付金別の得点と合うこと ──
    #    目標別の得点は該当状況調査票集計表の「Ⅰ 合計」等の行による。
    #    同じ表に明細の列もあるため、足し合わせると二重になる。
    try:
        import analyze_kofukin as AK45
        out45, tot45 = AK45.mokuhyo_suii()
        Y45 = ['令和6年度', '令和7年度', '令和8年度']
        t45 = None
        for c in SC.CH:
            for sec in c['sections']:
                if sec['no'] != '2-8':
                    continue
                for b in sec['blocks']:
                    if b['t'] == 'table' and '交付金・目標（各100点）' in b['head'][0]:
                        t45 = b
        bad45 = []
        if t45 is None:
            bad45.append('2-8に目標別の推移の表がない')
        else:
            got = {r[0]: r[1:4] for r in t45['rows']}
            # 各目標の値が算定と一致すること
            for k, t, v, _h in out45:
                key = next((x for x in got if x.startswith('%s 目標%s' % (k[:2], t))), None)
                if key is None:
                    bad45.append('%s目標%sの行がない' % (k[:2], t))
                    continue
                for i, y in enumerate(Y45):
                    if abs(float(got[key][i]) - v[y]) > 1e-9:
                        bad45.append('%s目標%s %s：算定%.0f≠素案%s'
                                     % (k[:2], t, y, v[y], got[key][i]))
            # 計が交付金別の得点の表と合うこと
            t1 = None
            for c in SC.CH:
                for sec in c['sections']:
                    if sec['no'] != '2-8':
                        continue
                    for b in sec['blocks']:
                        if b['t'] == 'table' and b['head'][:2] == ['', '令和6年度']:
                            t1 = b
            if t1:
                m = {r[0]: r[1:4] for r in t1['rows']}
                for kind, lab in [('推進交付金', '保険者機能強化推進交付金（400点）'),
                                  ('支援交付金', '介護保険保険者努力支援交付金（400点）')]:
                    if lab not in m:
                        continue
                    for i, y in enumerate(Y45):
                        if abs(tot45[(kind, y)] - float(str(m[lab][i]).replace(',', ''))) > 1e-9:
                            bad45.append('%s %s：目標別の計%.0f≠得点の表%s'
                                         % (kind[:2], y, tot45[(kind, y)], m[lab][i]))
                if '合計（800点）' in m:
                    for i, y in enumerate(Y45):
                        g = sum(tot45[(k, y)] for k in ('推進交付金', '支援交付金'))
                        if abs(g - float(str(m['合計（800点）'][i]).replace(',', ''))) > 1e-9:
                            bad45.append('合計 %s：目標別の計%.0f≠得点の表%s'
                                         % (y, g, m['合計（800点）'][i]))
        chk(45, '2-8 の目標別の推移が算定と一致し計が得点の表と合うこと', not bad45,
            '・'.join(bad45[:3]) if bad45
            else '8目標×3か年と、推進・支援・合計の計が一致')
    except Exception as e:
        chk(45, '2-8 の目標別の推移が算定と一致し計が得点の表と合うこと', False,
            f'照合できない（{e}）')

    # ── 46　図と図が同じ数値を持つところが食い違っていないこと ──────────
    #    人口の図（図2-1・図2-3・図2-4・図2-2）は同じ数値を別の切り口で持つ。
    #    片方だけを直すと静かに食い違うため、恒等式として突き合わせる。
    try:
        import data_zuhyo as _DZ46
        def _s46(fig, ser):
            return dict(zip(_DZ46.cats(fig), _DZ46.vals(fig, ser)))
        P3 = {k: _s46('fig2-3_将来推計人口', k)
              for k in ('15歳未満', '15〜64歳', '65歳以上', '総人口')}
        P4 = {k: _s46('fig2-4_前期後期高齢者', k) for k in ('65〜74歳', '75歳以上')}
        P2 = _s46('fig2-2_高齢化率推移', '北塩原村')
        P1 = {k: _s46('fig2-1_人口構成比較', k)
              for k in ('平成22年（2010年）', '令和7年（2025年）')}
        bad46, n46 = [], 0
        for y in P3['総人口']:
            n46 += 1
            g = P3['15歳未満'][y] + P3['15〜64歳'][y] + P3['65歳以上'][y]
            # 国勢調査の年は年齢不詳の分だけ3区分の計が総人口に満たない。
            # 認めるのは data_zuhyo.FUSHO に宣言した分だけである。
            g += getattr(_DZ46, 'FUSHO', {}).get(y, 0)
            if abs(g - P3['総人口'][y]) > 1e-9:
                bad46.append(f'図2-3 {y}：3区分の計{g:.0f}≠総人口{P3["総人口"][y]:.0f}')
            n46 += 1
            g = P4['65〜74歳'][y] + P4['75歳以上'][y]
            if abs(g - P3['65歳以上'][y]) > 1e-9:
                bad46.append(f'図2-4 {y}：前期＋後期{g:.0f}≠図2-3の65歳以上'
                             f'{P3["65歳以上"][y]:.0f}')
            n46 += 1
            r = P3['65歳以上'][y] / P3['総人口'][y] * 100
            if abs(round(r, 1) - P2[y]) > 0.051:
                bad46.append(f'図2-2 {y}：図2-3から求めた高齢化率{r:.1f}%'
                             f'≠{P2[y]}%')
        # 図2-1（年齢階級別）は図2-3・図2-4の同じ年と重なる
        for ser, y in (('平成22年（2010年）', '2010年'), ('令和7年（2025年）', '2025年')):
            d = P1[ser]
            for nm, got, want in (
                    ('15歳未満', d['15歳未満'], P3['15歳未満'][y]),
                    ('生産年齢人口', d['15〜39歳'] + d['40〜64歳'], P3['15〜64歳'][y]),
                    ('65〜74歳', d['65〜74歳'], P4['65〜74歳'][y]),
                    ('75歳以上', d['75歳以上'], P4['75歳以上'][y])):
                n46 += 1
                if abs(got - want) > 1e-9:
                    bad46.append(f'図2-1 {ser} {nm}：{got:.0f}≠{want:.0f}')
        chk(46, '図と図が同じ数値を持つところが一致すること', not bad46,
            '・'.join(bad46[:3]) if bad46 else f'人口の恒等式{n46}本が一致（年齢不詳は宣言した'
            f'{sum(getattr(_DZ46, "FUSHO", {}).values())}人のみ）')
    except Exception as e:
        chk(46, '図と図が同じ数値を持つところが一致すること', False, f'照合できない（{e}）')

    # ── 47　素案の表が図の数値の正本を参照していること（摂動試験）──────────
    #    正本（data_zuhyo）の値を動かしたときに素案の表の出力も動くことを確かめる。
    #    動かないなら表の側に数値がじか書きされている。
    #    data_zuhyo.LINKED が「参照している」と称している系列を1つずつ試す。
    try:
        import importlib
        import data_zuhyo as _DZ47
        def _cells():
            """素案のすべての表のセルを1本の文字列にする"""
            out = []
            for c in SC.CH:
                for sec in c['sections']:
                    for b in sec['blocks']:
                        if b.get('t') == 'table':
                            for r in b['rows']:
                                out.extend(str(x) for x in r)
            return '\u0001'.join(out)
        orig47 = _DZ47.vals
        base47 = _cells()
        bad47, n47 = [], 0
        for fig, ser in _DZ47.LINKED:
            def patched(name, series_name, use_book=True, _f=fig, _s=ser):
                v = orig47(name, series_name, use_book)
                if (name, series_name) == (_f, _s):
                    return [x + 1000 for x in v]
                return v
            _DZ47.vals = patched
            try:
                importlib.reload(SC)
                moved = _cells() != base47
            finally:
                _DZ47.vals = orig47
                importlib.reload(SC)
            n47 += 1
            if not moved:
                bad47.append(f'{fig}／{ser}：正本を動かしても表が変わらない')
        if _cells() != base47:
            bad47.append('摂動を戻したのに表が元に戻っていない')
        chk(47, '素案の表が図の数値の正本を参照していること', not bad47,
            '・'.join(bad47[:3]) if bad47
            else f'{n47}系列を1つずつ動かし、いずれも表に伝わった')
    except Exception as e:
        chk(47, '素案の表が図の数値の正本を参照していること', False,
            f'照合できない（{e}）')

    # ── 出力 ─────────────────────────────
    w = max(len(n) for _, n, _, _ in RESULTS)
    print('■ 計画素案の自己点検')
    ng = 0
    for no, name, ok, detail in RESULTS:
        mark = '適合' if ok else '不適合'
        print(f'  {no:>3}  {name:<{w}}  {mark}' + (f'  {detail}' if detail else ''))
        if not ok:
            ng += 1
    print(f'\n  {len(RESULTS)}件のうち適合{len(RESULTS)-ng}件／不適合{ng}件')
    if ng:
        print('  不適合があります。算定を改めたら素案（scripts/soan_content.py）にも反映してください。')
    return 1 if ng else 0


if __name__ == '__main__':
    sys.exit(main())
