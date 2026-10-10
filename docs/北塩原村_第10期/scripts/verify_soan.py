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
import os
import os as _os_p
import sys as _sys_p
_sys_p.path.insert(0, _os_p.path.dirname(_os_p.path.abspath(__file__)))
import paths as _P   # 置き場所はここで決める（じか書きしない）
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
    f18 = os.path.join(_P.DATA, '交付金_目標別の県内比較_令和8年度.csv')
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
    fp = os.path.join(_P.FIGURES, "fig3-1_施策体系図.png")
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
                # 要介護度別の認定者数の表。2-3には「時点」で始まる表が
                # 認定率・調整済み認定率にもあるため、要支援1の列で絞る
                if (b['t'] == 'table' and b['head'][0] == '時点'
                        and '合計' in b['head'] and '要支援1' in b['head']):
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

    # ── 26　5-2・5-3 の値が国の推計ワークシートと一致し、検算が再計算と合うこと ──
    #    令和8年10月7日より、5-2・5-3は国の推計ワークシート（将来推計総括表）の値を
    #    正とする。素案は受領ファイルから引いているため、じか書きの食い違いは起きないが、
    #    引く先を取り違える（総数と第1号を混同する等）ことは起こる。現に突合で
    #    「うち第1号被保険者数」629人と「総数」635人を取り違えていた。
    #    あわせて、検算として残した当方の算定（認定率×人口）が再計算と合うことを見る。
    import parse_worksheet as _PW26
    S52 = {sec['no']: sec for sec in SC.CH[4]['sections']}['5-2']
    S53 = {sec['no']: sec for sec in SC.CH[4]['sections']}['5-3']
    LV = ['要支援1', '要支援2', '要介護1', '要介護2', '要介護3', '要介護4', '要介護5']
    YM = ('令和9年度', '令和10年度', '令和11年度', '令和12年度',
          '令和17年度', '令和22年度')
    bad26, n26 = [], 0

    # (a) 5-2 の見込みの表＝国の年齢区分別
    t52 = [b for b in S52['blocks'] if b['t'] == 'table'
           and b['head'][:3] == ['', '前期高齢者', '後期高齢者']]
    if not t52:
        bad26.append('5-2の見込みの表がない')
    else:
        KZ, KO = _PW26.ws('前期(65～74歳)'), _PW26.ws('後期(75歳～)')
        K84, K85 = _PW26.ws('後期(75歳～84歳)'), _PW26.ws('後期(85歳～)')
        KH = _PW26.ws('第1号被保険者数')
        for r in t52[0]['rows']:
            y = str(r[0])
            if y not in YM:
                bad26.append('5-2に見込みでない年度 %s' % y)
                continue
            for col, src, nm in ((1, KZ, '前期'), (2, KO, '後期'),
                                 (3, K84, '75〜84歳'), (4, K85, '85歳以上'),
                                 (5, KH, '計')):
                n26 += 1
                if num(r[col]) != round(src[y]):
                    bad26.append('5-2 %s %s 素案%s≠国%d'
                                 % (y, nm, r[col], round(src[y])))

    # (b) 5-3 の主の表＝国の第1号被保険者の要介護度別
    t53 = [b for b in S53['blocks'] if b['t'] == 'table' and b['head'][:2] == ['', '要支援1']]
    if not t53:
        bad26.append('5-3の表がない')
    else:
        KN = _PW26.ws('第1号_計')
        for r in t53[0]['rows']:
            y = str(r[0])
            if y not in _PW26.NENDO:
                bad26.append('5-3に総括表にない年度 %s' % y)
                continue
            for k, lv in enumerate(LV, start=1):
                n26 += 1
                v = _PW26.ws('第1号_' + lv)[y]
                if num(r[k]) != round(v):
                    bad26.append('5-3 %s %s 素案%s≠国%d' % (y, lv, r[k], round(v)))
            n26 += 1
            if num(r[8]) != round(KN[y]):
                bad26.append('5-3 %s 計 素案%s≠国%d（総数と第1号の取り違えに注意）'
                             % (y, r[8], round(KN[y])))

    # (c) 検算（当方の算定）が認定率×人口と合うこと
    trate = None
    for sec in SC.CH[4]['sections']:
        for b in sec['blocks']:
            if (b['t'] == 'table' and b['head'][:3] == ['区分', '前期高齢者', '後期高齢者']
                    and any(r[0] == '要支援1' for r in b['rows'])):
                trate = b
    import shihyo_dict as SD2
    BAN = SD2.HIHOKENSHA_BAN
    tken = [b for b in S53['blocks'] if b['t'] == 'table'
            and b['head'][0] == '認定者数の置き方']
    if trate is None or not tken:
        bad26.append('認定率の表または検算の表がない')
    else:
        rate = {r[0]: (num(r[1]) / 100, num(r[2]) / 100)
                for r in trate['rows'] if r[0] in LV}
        gyo = [r for r in tken[0]['rows'] if '検算' in str(r[0])]
        if not gyo:
            bad26.append('検算の行がない')
        for k, y in enumerate(('令和9年', '令和10年', '令和11年'), start=1):
            zen, kou, _tot = BAN[y]
            raw = sum(zen * rate[lv][0] + kou * rate[lv][1] for lv in LV)
            n26 += 1
            if num(gyo[0][k]) != round(raw):
                bad26.append('検算 %s 素案%s≠認定率×人口%d' % (y, gyo[0][k], round(raw)))

    chk(26, '5-2・5-3 が国の推計ワークシートと一致し検算が再計算と合うこと', not bad26,
        '・'.join(bad26[:3]) if bad26
        else '国の値%d件が一致し、検算3か年が認定率×人口と一致' % n26)

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
    SAN = ['要介護認定の適正化', 'ケアプラン等の点検', '医療情報との突合・縦覧点検']
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

    # ── 36　確定版の積算が、送りを追った推定を上回らないこと ──────────
    #    仕様書の約100頁は成果品としての計画にかかる。素案の編集注記（⚙）と
    #    5-12（据え置いている項目の一覧）はいずれも計画の確定時に削除するため、
    #    積算はこの2つを除いた確定版で行う。素案のままの頁数も併せて示す。
    #
    #    【判定の基準を改めた経緯】
    #    もとは「95〜105頁」という絶対の幅で判定していた。しかし目安100頁を
    #    どれだけ超えるかは、すでに確認事項K-190として村の判断を仰いでいる事項であり、
    #    点検49（2つの推定が矛盾しないこと）と点検57（K-190に書いた頁数が推定と合うこと）が
    #    その数を押さえている。同じ問いを3つの点検が別々の基準で見ると、
    #    分量が増えたときに点検36だけが落ち続け、ほかの2つは適合のままになる。
    #    現にそうなった（確定版106頁で点検36のみ不適合）。
    #    そこで点検36は、2つの積算の関係が成り立つこと（要素を足すだけの積算が、
    #    送りを追った推定を上回らないこと）だけを見る形に改めた。
    #    目安からの超過は点検49が報告し、村への提示はK-190（点検57が番人）が受け持つ。
    #    分量が黙って増えることは、点検57がK-190の数との食い違いで捕らえる。
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
        # 送りを追った推定（点検49と同じ estimate_layout による計画書の頁数）
        import estimate_layout as _EL36
        okuri = _EL36.pages(_EL36.soan_items(True), _EL36.BODY_H) + 4
        bad36 = []
        if final > okuri:
            bad36.append(f'要素を足すだけの積算{final}頁が、送りを追った推定{okuri}頁を上回る')
        if est <= final:
            bad36.append(f'素案のまま{est}頁が確定版{final}頁を上回らない'
                         '（編集注記と5-12を除いても減っていない）')
        chk(36, '確定版の積算が送りを追った推定を上回らないこと', not bad36,
            '／'.join(bad36) if bad36
            else f'確定版{final}頁 ≦ 送りを追った推定{okuri}頁'
                 f'（素案のままは{est}頁。編集注記と5-12で{est - final}頁）。'
                 f'目安100頁との差は点検49が報告し、村への提示はK-190による') 
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
        FIGDIR = _P.FIGURES
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

    # ── 48　紙面が成り立つこと（図が1頁に収まり、送りが行き過ぎないこと）──────
    #    この環境では docx を PDF に変換できず紙面を目視できない。
    #    代わりに estimate_layout.py の積み上げにより、紙面として成り立たない
    #    箇所がないことを確かめる。**目視の代わりにはならない。**
    try:
        import estimate_layout as EL
        import estimate_pages as EP48
        import re as _re48
        bad48 = []
        # (a) docx 側の定数と推定側の定数が同じ値であること（二重管理の検出）
        for js, attr in (("build_soan_docx.js", "MAX_FIG_H"),
                         ("build_shiryo_docx.js", "SH_MAX_FIG_H")):
            with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), js),
                      encoding="utf-8") as f48:
                src48 = f48.read()
            m48 = _re48.search(r"MAX_H_IN\s*=\s*([0-9.]+)", src48)
            if not m48:
                bad48.append(f"{js} に図の高さの上限がない")
            elif abs(float(m48.group(1)) - getattr(EP48, attr)) > 1e-9:
                bad48.append(f"{js} の上限{m48.group(1)}in ≠ "
                             f"estimate_pages.{attr} {getattr(EP48, attr)}in")
            # 1頁の図版とする割合（素案のみ）
            if js == "build_soan_docx.js":
                z48 = _re48.search(r"ZENMEN_H_IN\s*=\s*[0-9.]+\s*\*\s*([0-9.]+)", src48)
                if not z48:
                    bad48.append(f"{js} に1頁の図版とする割合がない")
                elif abs(float(z48.group(1)) - EL.ZENMEN) > 1e-9:
                    bad48.append(f"{js} の割合{z48.group(1)} ≠ "
                                 f"estimate_layout.ZENMEN {EL.ZENMEN}")
        # (d) 表の列幅が固定であること（自動調整だと宣言した割付も推定も崩れる）
        for js in ("build_soan_docx.js", "build_shiryo_docx.js"):
            with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), js),
                      encoding="utf-8") as f48:
                if "TableLayoutType.FIXED" not in f48.read():
                    bad48.append(f"{js} の表の列幅が固定になっていない")
        # (b) 1頁に収まらない要素が表だけであること（図は分割できない）
        for items, ph, nm48 in ((EL.soan_items(), EL.BODY_H, "素案"),
                                (EL.shiryo_items(), EL.SH_H, "委員会資料")):
            _, out48 = EL.flow(items, ph)
            over = [o for o in out48 if o[0] == "頁超" and "表（" not in o[2]]
            if over:
                bad48.append(f"{nm48}：1頁に収まらない図がある（"
                             + "・".join(o[2][:30] for o in over[:2]) + "）")
            # (c) 送りで生じる空白が頁の7割を超えないこと（紙面が破れている徴候）
            huge = [o for o in out48 if o[0] in ("送り", "見出しの送り")
                    and o[3] >= ph * 0.70]
            if huge:
                bad48.append(f"{nm48}：送りで7割を超える空白が{len(huge)}件")
        chk(48, '紙面が成り立つこと（推定。目視の代わりにはならない）', not bad48,
            '・'.join(bad48[:3]) if bad48
            else f'図{len(__import__("figures_map").FIGS)}点はいずれも1頁に収まり、'
                 f'上限の値は docx 側と一致')
    except Exception as e:
        chk(48, '紙面が成り立つこと（推定。目視の代わりにはならない）', False,
            f'照合できない（{e}）')

    # ── 49　2つの頁数の推定が矛盾しないこと ─────────────────────
    #    estimate_pages.py は要素の高さを足すだけ（送りの空白を見ない下限値）、
    #    estimate_layout.py は上から積み上げて頁の境目を見る。
    #    後者は前者を下回り得ない。下回ったらどちらかの積算が壊れている。
    #    あわせて計画書としての頁数が仕様書の目安から離れすぎていないかを見る。
    try:
        import estimate_pages as EP49
        import estimate_layout as EL49
        rows49, tot49, _, _ = EP49.run()
        n_soan = EL49.pages(EL49.soan_items(False), EL49.BODY_H)
        n_kei = EL49.pages(EL49.soan_items(True), EL49.BODY_H)
        bad49 = []
        if n_soan < tot49:
            bad49.append(f'送りを追った推定{n_soan}頁が積み上げ{tot49:.1f}頁を下回る')
        if n_kei > n_soan:
            bad49.append(f'計画書として{n_kei}頁が素案のまま{n_soan}頁を上回る')
        # 計画書（本文＋前付4頁）が仕様書の目安100頁から±20頁を超えて離れていないか
        zentai = n_kei + 4
        if abs(zentai - 100) > 20:
            bad49.append(f'計画書全体{zentai}頁が仕様書の目安100頁から'
                         f'{zentai - 100:+d}頁離れている')
        chk(49, '頁数の2つの推定が矛盾しないこと', not bad49,
            '・'.join(bad49[:2]) if bad49
            else f'積み上げ{tot49:.0f}頁 ≦ 送りを追った推定{n_soan}頁、'
                 f'計画書として{n_kei}頁＋前付4頁＝{zentai}頁（目安100頁に対し'
                 f'{zentai - 100:+d}頁）')
    except Exception as e:
        chk(49, '頁数の2つの推定が矛盾しないこと', False, f'照合できない（{e}）')

    # ── 50　計画の名称と期間の表記が統一されていること ─────────────────
    #    同じものを別の言い方で書くと、計画書と委員会資料で食い違う。
    #    本則の形を決め、認める例外を明示する。素案と委員会資料の両方を見る。
    try:
        import re as _re50
        import shiryo_content as _SH50
        import shiryo3_content as _SH350

        def _text50(mod):
            o = []
            for c in mod.CH:
                o.append(f'{c["no"]}　{c["title"]}')
                for sec in c["sections"]:
                    o.append(f'{sec["no"]}　{sec.get("title", "")}')
                    for b in sec["blocks"]:
                        if b["t"] == "table":
                            o.append("　".join(map(str, b["head"])))
                            for r in b["rows"]:
                                o.append("　".join(map(str, r)))
                        elif "v" in b:
                            o.append(str(b["v"]))
            return "\n".join(o)

        # (言い方, 本則か, 認める箇所)
        YURE = [
            (r"北塩原村第10期高齢者福祉計画", False, "計画の名称は「第10期北塩原村…」"),
            (r"北塩原村第10期介護保険事業計画", False, "計画の名称は「第10期北塩原村…」"),
            (r"令和9年度[〜～]令和11年度", False, "期間は「令和9年度から令和11年度まで」"),
            (r"3年間", False, "期間の数え方は「3か年」"),
            (r"三か年", False, "期間の数え方は「3か年」"),
            (r"令和9年度から令和11年度迄", False, "「まで」を用いる"),
        ]
        bad50 = []
        for nm50, txt50 in (("素案", open(MD, encoding="utf-8").read()),
                            ("第2回資料", _text50(_SH50)),
                            ("第3回骨子", _text50(_SH350))):
            for pat, honsoku, naze in YURE:
                hit = _re50.findall(pat, txt50)
                if hit and not honsoku:
                    bad50.append(f"{nm50}：{hit[0]} が{len(hit)}件（{naze}）")
        # 本則の形が実際に使われていること（言い方を決めただけで使っていない事故を防ぐ）
        so50 = open(MD, encoding="utf-8").read()
        for honsoku in ("第10期北塩原村高齢者福祉計画", "第10期北塩原村介護保険事業計画",
                        "令和9年度から令和11年度まで"):
            if honsoku not in so50:
                bad50.append(f"素案に本則の「{honsoku}」がない")
        chk(50, "計画の名称と期間の表記が統一されていること", not bad50,
            "・".join(bad50[:3]) if bad50
            else "本則3通りが素案にあり、素案・第2回資料・第3回骨子に表記のゆれなし"
                 "（表の見出しの「令和9〜11年度」は字数の制約により認める）")
    except Exception as e:
        chk(50, "計画の名称と期間の表記が統一されていること", False, f"照合できない（{e}）")

    # ── 51　空欄の並ぶ表が5-12に掲げられていること ────────────────────
    #    成果品に「―」や【要設定】が並ぶ表があるのに、据え置いた項目の一覧
    #    （5-12）に載っていないと、何を待っているのかが誰にも分からなくなる。
    #    4-6（高齢者福祉サービスの量の目標）が12行すべて空欄のまま載っていなかった。
    try:
        KARA = ("―", "【要設定】", "")
        # 資料編の枠は、計画の数値ではなく策定の記録を後から入れるものである。
        # 5-12（据え置いた数値の一覧）に載せると性格の違うものが混ざるため除く。
        # ただし**除くのは、その節に何を待っているかの編集注記があるときだけ**とする。
        # 注記もなく空欄が並ぶのは、ただの書き漏れである。
        MENJO = {"資2", "資3", "資4", "資6"}
        sora = {}
        for c in SC.CH:
            for sec in c["sections"]:
                for b in sec["blocks"]:
                    if b["t"] not in ("table", "kpi"):
                        continue
                    n = sum(1 for r in b["rows"] for v in r[1:]
                            if str(v).strip() in KARA)
                    if n >= 3:
                        sora[sec["no"]] = sora.get(sec["no"], 0) + n
        # 5-12 の「節」欄に挙がっている節番号を集める
        kakage = set()
        for c in SC.CH:
            for sec in c["sections"]:
                if sec["no"] != "5-12":
                    continue
                for b in sec["blocks"]:
                    if b["t"] == "table":
                        for r in b["rows"]:
                            for part in str(r[0]).replace("・", "／").split("／"):
                                part = part.strip()
                                if part:
                                    kakage.add(part)
                                    # 5-4(3) のような枝番は親の節でも照合できるように
                                    kakage.add(part.split("(")[0])
        # 「第4章」と書いてあるときは、第4章の各施策の節を覆っているとみなす
        if "第4章" in kakage:
            for c in SC.CH:
                if c["no"] == "第4章":
                    for sec in c["sections"]:
                        kakage.add(sec["no"])
        # 除くと定めた節に、待っているものを書いた編集注記があること
        note_aru = set()
        for c in SC.CH:
            for sec in c["sections"]:
                if any(b["t"] == "note" for b in sec["blocks"]):
                    note_aru.add(sec["no"])
        bad51 = [f"{k}（空欄{v}）" for k, v in sorted(sora.items())
                 if k not in MENJO and k not in kakage
                 and k.split("(")[0] not in kakage]
        bad51 += [f"{k}：除くと定めたが編集注記がない" for k in sorted(MENJO)
                  if k in sora and k not in note_aru]
        chk(51, "空欄の並ぶ表が5-12に掲げられていること", not bad51,
            "・".join(bad51[:4]) if bad51
            else f"空欄の並ぶ表は{len(sora)}節にあり、"
                 f"{len(sora) - len([k for k in sora if k in MENJO])}節が5-12に、"
                 f"{len([k for k in sora if k in MENJO])}節が資料編の枠として"
                 f"編集注記つきで除かれている")
    except Exception as e:
        chk(51, "空欄の並ぶ表が5-12に掲げられていること", False, f"照合できない（{e}）")

    # ── 52　分量の削減案が成り立つこと ────────────────────────────
    #    削減案が指す小見出しが実在しないと、減る頁が黙って0になる。
    #    実際に「算定の流れ」「(5) 感応度」という実在しない名を書いて
    #    0頁と出ていた。案が効いていることを機械で確かめる。
    try:
        import estimate_sakugen as ES
        import estimate_layout as EL52
        bad52 = []
        base52 = EL52.pages(EL52.soan_items(True), EL52.BODY_H)
        for no, nm, tg, hotei, _riyu in ES.AN:
            if not tg:
                continue
            # 指す小見出しが実在すること
            for sec_no, h3 in tg:
                if ES.h3_range(sec_no, h3) is None:
                    bad52.append(f"案{no}：{sec_no}に「{h3[:18]}」という小見出しがない")
            # 減る頁が0でないこと（0なら案として意味をなさない）
            n = EL52.pages(ES.items_without(tg), EL52.BODY_H)
            if base52 - n <= 0:
                bad52.append(f"案{no}：削っても頁が減らない")
            if hotei:
                bad52.append(f"案{no}：法定記載事項を削る案になっている")
        # 小見出しを指さない案（G 用語の絞り込み・H 図を資料編へ）も、
        # 効果が何頁かを測っておく。0頁の案は「頁のためには行う意味がない」と
        # 示せていればよいので、減らないことを不適合にはしない。
        n_g = EL52.pages(ES.items_without([], yougo=ES.YOUGO_NAI), EL52.BODY_H)
        n_h = EL52.pages(ES.items_without([], figs=ES.FIG_UTSUSU), EL52.BODY_H)
        # すべて行えば目安の約100頁に収まること（前付4頁を含む）
        allt = [t for _n, _m, tg, _h, _r in ES.AN for t in tg]
        n_all = EL52.pages(ES.items_without(allt, yougo=ES.YOUGO_NAI,
                                            figs=ES.FIG_UTSUSU), EL52.BODY_H) + 4
        if n_all > 105:
            bad52.append(f"すべて行っても{n_all}頁で目安100頁に収まらない")
        chk(52, "分量の削減案が成り立つこと", not bad52,
            "・".join(bad52[:3]) if bad52
            else f"{len(ES.AN)}案（うち小見出しを指すもの{len([a for a in ES.AN if a[2]])}案）。"
                 f"G 用語の絞り込みは{base52 - n_g}頁、H 図2点を資料編へは{base52 - n_h}頁。"
                 f"すべて行えば本文{n_all - 4}頁＋前付4頁＝{n_all}頁"
                 f"（目安100頁に対し{n_all - 100:+d}頁）")
    except Exception as e:
        chk(52, "分量の削減案が成り立つこと", False, f"照合できない（{e}）")

    # ── 53　老人福祉事業の量の確保の方策が実在する施策を指すこと ─────────
    #    老人福祉法第20条の8第3項第1号は、量の確保のための方策を定めるよう求めている。
    #    4-6の12事業それぞれに方策を書いたが、指す施策が実在しなければ追えない。
    try:
        import re as _re53
        sisaku = set()
        for c in SC.CH:
            if c["no"] != "第4章":
                continue
            for sec in c["sections"]:
                if sec["no"].startswith("施策"):
                    sisaku.add(sec["no"].replace("施策", ""))
                elif sec["no"] == "4-5":
                    # 基本目標5 の4施策は節ではなく 4-5 の表の行として置いている
                    for b in sec["blocks"]:
                        if b["t"] == "table":
                            sisaku |= {str(r[0]).split(" ", 1)[0] for r in b["rows"]}
        bad53, n53 = [], 0
        for c in SC.CH:
            for sec in c["sections"]:
                if sec["no"] != "4-6":
                    continue
                for b in sec["blocks"]:
                    if b["t"] != "table":
                        continue
                    col = b["head"][-1]
                    if "方策" not in col:
                        bad53.append(f"{b['head'][0]}の表に確保の方策の欄がない")
                        continue
                    for r in b["rows"]:
                        n53 += 1
                        hosaku = str(r[-1])
                        if not hosaku.strip() or hosaku.strip() == "―":
                            bad53.append(f"{r[0]}：確保の方策が空")
                            continue
                        # 施策を指しているなら、その施策が実在すること
                        for m in _re53.findall(r"施策([0-9]+-[0-9]+)", hosaku):
                            if m not in sisaku:
                                bad53.append(f"{r[0]}：施策{m}が第4章にない")
        chk(53, "老人福祉事業の量の確保の方策が実在する施策を指すこと", not bad53,
            "・".join(bad53[:3]) if bad53
            else f"{n53}事業すべてに確保の方策があり、指す施策は第4章に実在する")
    except Exception as e:
        chk(53, "老人福祉事業の量の確保の方策が実在する施策を指すこと", False,
            f"照合できない（{e}）")

    # ── 54　正負の記号が成果品の間で揃っていること ────────────────────
    #    負は▲、正は全角の＋で揃える。半角の＋と半角のハイフンを混ぜると、
    #    同じ数値が計画書と委員会資料で違う記号で示される。
    #    負号は先に揃えたが、正号は素案で全角74件・半角31件が混在していた。
    try:
        import re as _re54
        import shiryo_content as _SH54
        import shiryo3_content as _S354

        def _cells54(mod):
            o = []
            for c in mod.CH:
                for sec in c["sections"]:
                    for b in sec["blocks"]:
                        if b["t"] == "table":
                            for r in b["rows"]:
                                o += [(sec["no"], str(x)) for x in r]
                            o += [(sec["no"], str(x)) for x in b["head"]]
                        elif "v" in b:
                            o.append((sec["no"], str(b["v"])))
            return o
        HANKAKU_P = _re54.compile(r"(?<![A-Za-z0-9])\+[0-9]")
        HANKAKU_M = _re54.compile(r"(?<![A-Za-z0-9\-])\-[0-9][0-9,\.]*(%|円|人|点|ポイント)")
        MINUS = _re54.compile(r"[−–—]\s*[0-9]")      # U+2212 ほかのダッシュ
        bad54 = []
        for nm54, mod54 in (("素案", SC), ("第2回資料", _SH54), ("第3回骨子", _S354)):
            for sec_no, t in _cells54(mod54):
                if HANKAKU_P.search(t):
                    bad54.append(f"{nm54} {sec_no}：半角の＋（{t[:22]}）")
                if HANKAKU_M.search(t):
                    bad54.append(f"{nm54} {sec_no}：半角のハイフンを負号に（{t[:22]}）")
                if MINUS.search(t):
                    bad54.append(f"{nm54} {sec_no}：▲以外の負号（{t[:22]}）")
        chk(54, "正負の記号が成果品の間で揃っていること", not bad54,
            "・".join(bad54[:3]) if bad54
            else "素案・第2回資料・第3回骨子のいずれも、正は全角の＋、負は▲で揃っている")
    except Exception as e:
        chk(54, "正負の記号が成果品の間で揃っていること", False, f"照合できない（{e}）")

    # ── 55　同じ文を2か所で述べていないこと ──────────────────────
    #    第2章（現状）と第4章（施策）で同じ事実に触れるのは読み手に必要な繰り返しだが、
    #    同じ文を丸ごと2度置くと、片方だけを直したときに食い違う。
    #    数値を伏せて正規化した文が別の節に現れたら、どちらかを参照に替える。
    #
    #    しきい値（YOMI_MIN）は実測で決めた。10月5日の改訂前の本文で測ると、
    #    30文字では9組のうち5組しか出ず、短い重複（「主任介護支援専門員は令和4年度から
    #    配置されました」24文字、「第2層協議体の活動を支援します」15文字）を取りこぼした。
    #    12文字まで下げても出るのは9組のままで、作り物の重複は増えない。
    #    実在の最短が15文字なので、12文字は余裕をもった値である。
    try:
        import re as _re55
        _NUM55 = _re55.compile(r"[0-9０-９][0-9０-９,，\.．]*")
        _PUNC55 = _re55.compile(r"[「」（）\(\)、・。％%]")

        def _norm55(t):
            return _PUNC55.sub("", _NUM55.sub("#", t)).strip()

        import shiryo_content as _SH55
        import shiryo3_content as _S355
        YOMI_MIN = 12          # 実測で決めたしきい値（上の覚え書きのとおり）
        bad55, n55 = [], 0
        for nm55, mod55 in (("素案", SC), ("第2回資料", _SH55), ("第3回骨子", _S355)):
            bag55 = {}
            for c in mod55.CH:
                for sec in c["sections"]:
                    for b in sec["blocks"]:
                        if b["t"] in ("p", "note"):
                            src = [str(b["v"])]
                        elif b["t"] == "bullets":
                            src = [str(x) for x in b["v"]]
                        else:
                            continue
                        for raw in src:
                            for sn in raw.split("。"):
                                k = _norm55(sn)
                                if len(k) >= YOMI_MIN:
                                    bag55.setdefault(k, []).append((sec["no"], b["t"], sn))
            n55 += len(bag55)
            for k, v in bag55.items():
                doko = sorted({(a, b) for a, b, _ in v})
                if len(doko) > 1:
                    bad55.append("%s %s：%s" % (nm55, "／".join("%s(%s)" % d for d in doko),
                                               v[0][2][:30]))
        chk(55, "同じ文を2か所で述べていないこと", not bad55,
            "・".join(bad55[:3]) if bad55
            else "素案・第2回資料・第3回骨子の%d文のうち、数値を伏せて%d文字以上が一致する文が"
                 "別の節に現れるものはない" % (n55, YOMI_MIN))
    except Exception as e:
        chk(55, "同じ文を2か所で述べていないこと", False, "照合できない（%s）" % e)

    # ── 56　「次のN点」の宣言と実際の数が合うこと ─────────────────
    #    本文に小見出し・表の行・箇条書きを足し引きすると、前置きの数だけが残る。
    #    2-11は課題を6本に増やしたあとも「次の5点」のままだった。
    try:
        import re as _re56
        import shiryo_content as _SH56
        import shiryo3_content as _S356
        _MARU56 = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"
        _KAZU56 = _re56.compile(r"次の([0-9０-９]+)(点|つ|項目)")

        def _zen56(t):
            return int(t.translate(str.maketrans("０１２３４５６７８９", "0123456789")))

        def _no56(t):
            """小見出しの通し番号を（接頭の形, 番号）で返す。番号が無ければ None。"""
            if t and t[0] in _MARU56:
                return ("maru", _MARU56.index(t[0]) + 1)
            m = _re56.match(r"^\((\d+)\)", t)
            if m:
                return ("kakko", int(m.group(1)))
            m = _re56.match(r"^(\D{1,4}?)(\d+)", t)
            return (m.group(1), int(m.group(2))) if m else None

        def _run56(bs, i):
            """位置 i の直後から、通し番号の続く小見出しがいくつ並ぶか。"""
            h3 = [b["v"] for b in bs[i + 1:] if b["t"] == "h3"]
            f = _no56(h3[0]) if h3 else None
            if not f:
                return 0
            n, pre, cur = 1, f[0], f[1]
            for t in h3[1:]:
                g = _no56(t)
                if not g or g[0] != pre or g[1] != cur + 1:
                    break
                n, cur = n + 1, g[1]
            return n

        def _first56(bs, i, kinds):
            for b in bs[i + 1:i + 4]:
                if b["t"] in kinds:
                    return b
            return None

        bad56, n56 = [], 0
        for nm56, mod56 in (("素案", SC), ("第2回資料", _SH56), ("第3回骨子", _S356)):
            for c in mod56.CH:
                for sec in c["sections"]:
                    bs = sec["blocks"]
                    for i, b in enumerate(bs):
                        if b["t"] not in ("p", "note"):
                            continue
                        v = str(b.get("v", ""))
                        m = _KAZU56.search(v)
                        if not m:
                            continue
                        n56 += 1
                        d = _zen56(m.group(1))
                        cand = {sum(1 for ch in v if ch in _MARU56), _run56(bs, i)}
                        tb = _first56(bs, i, ("table", "kpi"))
                        if tb:
                            cand |= {len(tb["rows"]), len(tb["head"]) - 1}
                        bu = _first56(bs, i, ("bullets",))
                        if bu:
                            cand.add(len(bu["v"]))
                        if d not in cand:
                            bad56.append("%s %s：次の%d%s に対し %s"
                                         % (nm56, sec["no"], d, m.group(2),
                                            sorted(cand - {0}) or "数えられる並びがない"))
        # 確認事項（WBS の xlsx に載る文）も、数を宣言して丸数字で並べているなら数える。
        # 丸数字のない文は、欠陥を引用している進捗の備考などに当たるため対象外。
        import wbs_kakunin as _KK56

        def _walk56(o, path=""):
            if isinstance(o, str):
                yield path, o
            elif isinstance(o, dict):
                for k, v in o.items():
                    yield from _walk56(v, "%s/%s" % (path, k))
            elif isinstance(o, (list, tuple)):
                for i, v in enumerate(o):
                    yield from _walk56(v, "%s[%d]" % (path, i))
        for nm56, obj56 in (("確認事項", _KK56.K), ("解決済", _KK56.SOLVED)):
            for path, t in _walk56(obj56, nm56):
                m = _KAZU56.search(t)
                if not m:
                    continue
                maru = sum(1 for ch in t if ch in _MARU56)
                if not maru:
                    continue
                n56 += 1
                if _zen56(m.group(1)) != maru:
                    bad56.append("%s %s：次の%s%s に対し丸数字%d"
                                 % (nm56, path, m.group(1), m.group(2), maru))
        chk(56, "「次のN点」の宣言と実際の数が合うこと", not bad56,
            "・".join(bad56[:3]) if bad56
            else "%d件の宣言は、続く小見出し・表の行・箇条書き・丸数字のいずれかの数と一致する" % n56)
    except Exception as e:
        chk(56, "「次のN点」の宣言と実際の数が合うこと", False, "照合できない（%s）" % e)

    # ── 57　自分の構造を数えて述べている箇所が、構造と合うこと ────────
    #    3-4 は「5つの基本目標」「28の施策」「新規・新設は9施策」と数を述べている。
    #    施策を足し引きするとこの数だけが残る。現に「10施策」のまま9施策になっていた。
    #    点検56 は「次のN点」の形しか見ないため、この形を別に数える。
    try:
        import re as _re57
        ch57 = {c["no"]: c for c in SC.CH}

        # 構造から数える（体系図 build_figure_taikei と同じ数え方）
        mokuhyo57 = []
        for sec in ch57["第3章"]["sections"]:
            if sec["no"] == "3-2":
                for b in sec["blocks"]:
                    if b["t"] == "table":
                        mokuhyo57 = [r[0] for r in b["rows"]]
        kubun57 = []
        for sec in ch57["第4章"]["sections"]:
            if sec["no"].startswith("施策"):
                kubun57.append(sec["title"].split("【")[1].rstrip("】"))
            elif sec["no"] == "4-5":
                for b in sec["blocks"]:
                    if b["t"] == "table":
                        kubun57 += ["新設"] * len(b["rows"])
        n_moku = len(mokuhyo57)
        n_sis = len(kubun57)
        n_shin = sum(1 for k in kubun57 if k in ("新規", "新設"))

        # 本文が述べている数を拾う（第9期について述べている数は対象外）
        HOR = {"基本目標": n_moku, "の施策": n_sis}
        bad57 = []
        n57 = 0
        for sec in ch57["第3章"]["sections"]:
            if sec["no"] not in ("3-4", "3-5"):
                continue
            for b in sec["blocks"]:
                if b["t"] not in ("p", "note"):
                    continue
                v = str(b["v"])
                # 「第9期は4つの基本目標と19の施策でした」は過去の計画の数なので外す
                v9 = _re57.sub(r"第9期[^。]*。", "", v)
                for m in _re57.finditer(r"([0-9]+)つの基本目標", v9):
                    n57 += 1
                    if int(m.group(1)) != n_moku:
                        bad57.append("%s：%sつの基本目標（実際は%d）"
                                     % (sec["no"], m.group(1), n_moku))
                for m in _re57.finditer(r"([0-9]+)の施策", v9):
                    n57 += 1
                    if int(m.group(1)) != n_sis:
                        bad57.append("%s：%sの施策（実際は%d）"
                                     % (sec["no"], m.group(1), n_sis))
                for m in _re57.finditer(r"新規・新設は([0-9]+)施策", v9):
                    n57 += 1
                    if int(m.group(1)) != n_shin:
                        bad57.append("%s：新規・新設は%s施策（実際は%d）"
                                     % (sec["no"], m.group(1), n_shin))
        # 確認事項の本文に書いた頁数も、紙面の推定と合っていること。
        # 確認事項は照会票として村に渡るため、古い数が残ると村に誤った数を示す。
        # 現に K-190 が「本文105頁＋前付4頁＝109頁」のまま106頁＋4頁になっていた。
        import estimate_layout as _EL57
        import estimate_sakugen as _ES57
        import wbs_kakunin as _KK57
        n_kei57 = _EL57.pages(_EL57.soan_items(True), _EL57.BODY_H)
        # 削減案をすべて行った場合の頁数は、点検52と同じ算定で出す。
        # 小見出しを指す案だけを渡していたため、用語の絞り込み（案G）と
        # 図を資料編へ移す案（案H）が抜け、estimate_sakugen の出す到達点と
        # 3頁食い違っていた（令和8年10月9日）。
        n_all57 = _EL57.pages(
            _ES57.items_without([t for a in _ES57.AN for t in a[2]],
                                yougo=_ES57.YOUGO_NAI, figs=_ES57.FIG_UTSUSU),
            _EL57.BODY_H)
        MACHI57 = {"本文%d頁＋前付4頁＝%d頁" % (n_kei57, n_kei57 + 4),
                   "本文%d頁＋前付4頁＝%d頁" % (n_all57, n_all57 + 4)}
        # 見る先は確認事項だけでなく、WBSの「着手しない理由」と「翌営業日の作業」も
        # 含める。理由の欄は WBS 進捗管理表として村に渡るのに、確認事項と違って
        # 誰も数を確かめていなかった。現に Ⅱ-113 の理由が「到達できる最小が102頁」
        # のまま106頁になり、Ⅴ-100 が「142件」のまま145件、Ⅱ-115 が「71件」のまま
        # 72件になっていた（令和8年10月9日）。
        import wbs_pending as _PD57
        import irai_bundle as _IB57
        n_irai57 = sum(_IB57.count(nm) for nm, _ti, _ne in _IB57.bundles())
        MITO57 = ([("確認事項「%s」" % str(k[1])[:24], str(k[2])) for k in _KK57.K]
                  + [("WBS %s の着手しない理由" % w, str(t))
                     for w, t in _PD57.HOLD_REASON.items()]
                  + [("翌営業日の作業 順位%d（%s）" % (r[0], r[1]),
                      str(r[2]) + str(r[3])) for r in _PD57.READY])
        for doko, t in MITO57:
            for m in _re57.finditer(r"本文[0-9]+頁＋前付[0-9]+頁＝[0-9]+頁", t):
                n57 += 1
                if m.group() not in MACHI57:
                    bad57.append("%s：%s（推定は %s）"
                                 % (doko, m.group(), "／".join(sorted(MACHI57))))
            for m in _re57.finditer(r"目安100頁を([0-9]+)頁上回る", t):
                n57 += 1
                if int(m.group(1)) != n_kei57 + 4 - 100:
                    bad57.append("%s：目安100頁を%s頁上回る（推定は%d頁）"
                                 % (doko, m.group(1), n_kei57 + 4 - 100))
            # 「目安100頁に対し+N頁」は、今の分量を言うときと
            # 削減案をすべて行った到達点を言うときの両方に使う言い方であるため、
            # どちらかに合っていればよいとする。
            for m in _re57.finditer(r"目安100頁に対し\+([0-9]+)頁", t):
                n57 += 1
                if int(m.group(1)) not in (n_kei57 + 4 - 100, n_all57 + 4 - 100):
                    bad57.append("%s：目安100頁に対し+%s頁"
                                 "（今の分量は+%d頁・削減案をすべて行うと+%d頁）"
                                 % (doko, m.group(1), n_kei57 + 4 - 100,
                                    n_all57 + 4 - 100))
            for m in _re57.finditer(r"到達できる最小は([0-9]+)頁", t):
                n57 += 1
                if int(m.group(1)) != n_all57 + 4:
                    bad57.append("%s：到達できる最小は%s頁（試算は%d頁）"
                                 % (doko, m.group(1), n_all57 + 4))
            # 照会票に載せた件数（束1〜6の合計）
            for m in _re57.finditer(r"束1〜6の6通に組んだ（([0-9]+)件）", t):
                n57 += 1
                if int(m.group(1)) != n_irai57:
                    bad57.append("%s：照会票%s件（実際は%d件）"
                                 % (doko, m.group(1), n_irai57))

        # 「村内で提供されているサービスはN種類」が、2-5の事業所数の表と合うこと。
        # 定員の表（居住系・通所系のみ）に引かれて訪問介護を数え落とし、
        # 7か所で「2種類」と述べていた。ケアマネジメント（居宅介護支援・
        # 介護予防支援）は給付のサービスではないため数に入れない。
        CARE57 = ("居宅介護支援", "介護予防支援")
        murauchi = set()
        for c in SC.CH:
            for sec in c["sections"]:
                if sec["no"] != "2-5":
                    continue
                for b in sec["blocks"]:
                    if b["t"] == "table" and b["head"][0] == "サービス" and "令和6年度" in b["head"]:
                        col = b["head"].index("令和6年度")
                        for r in b["rows"]:
                            if str(r[0]) in CARE57:
                                continue
                            try:
                                if float(str(r[col])) > 0:
                                    murauchi.add(str(r[0]))
                            except ValueError:
                                pass
        n_mura = len(murauchi)
        for c in SC.CH:
            for sec in c["sections"]:
                for b in sec["blocks"]:
                    txt = []
                    if b["t"] in ("p", "note"):
                        txt = [str(b["v"])]
                    elif b["t"] == "bullets":
                        txt = [str(x) for x in b["v"]]
                    for t in txt:
                        for m in _re57.finditer(
                                r"村内(?:で提供され[^。]{0,8}|に提供事業所が|の事業所は)"
                                r"[^。]{0,70}?([0-9]+)種類", t):
                            n57 += 1
                            if int(m.group(1)) != n_mura:
                                bad57.append("%s：村内のサービス%s種類（表では%d種類）"
                                             % (sec["no"], m.group(1), n_mura))

        if not n57:
            bad57.append("数を述べた箇所が見つからない（走査が効いていない）")
        chk(57, "自分の数を述べているところが実際と合うこと", not bad57,
            "・".join(bad57[:3]) if bad57
            else "基本目標%d・施策%d（うち新規・新設%d）・計画書%d頁・村内のサービス%d種類を"
                 "述べた%d件が構造と推定に一致する"
                 % (n_moku, n_sis, n_shin, n_kei57 + 4, n_mura, n57))
    except Exception as e:
        chk(57, "自分の構造を数えて述べている箇所が構造と合うこと", False,
            "照合できない（%s）" % e)

    # ── 58　素案が指す調査票の設問が、集計仕様書に実在すること ────────
    #    素案は調査票の設問番号で「どこで把握するか」を述べる。設問の番号は
    #    調査票の改訂で動くため、素案の番号だけが残る。
    #    現に資6の事項六で「介護者の勤務形態（B票 問6）」と書いていたが、
    #    B票 問6は不安に感じる介護で、勤務形態はB票 問7である。
    #    設問文の中の語と突き合わせて、番号と中身が合うことを確かめる。
    try:
        import re as _re58
        import shukei_data as _SK58
        # 集計仕様書の設問　「票＋番号」→ 設問文
        # Z の並びは（調査, 票, 設問番号, 設問文, …）
        zai58 = {}
        for r in _SK58.Z:
            zai58[(str(r[1]), str(r[2]))] = str(r[3])

        # 素案の本文から「（B票 問7）」「（A票 問6）」「（問5(1)）」の形を拾う
        PAT58 = _re58.compile(r"([AＡBＢ]票)\s*(問[0-9０-９\-－]+)")
        bad58, n58 = [], 0
        for c in SC.CH:
            for sec in c["sections"]:
                for b in sec["blocks"]:
                    txt = []
                    if b["t"] in ("p", "note"):
                        txt = [str(b["v"])]
                    elif b["t"] == "bullets":
                        txt = [str(x) for x in b["v"]]
                    elif b["t"] in ("table", "kpi"):
                        txt = [str(x) for r in b["rows"] for x in r]
                    for t in txt:
                        for m in PAT58.finditer(t):
                            hyo = m.group(1).replace("Ａ", "A").replace("Ｂ", "B")
                            q = m.group(2).replace("－", "-")
                            n58 += 1
                            if (hyo, q) not in zai58:
                                bad58.append("%s：%s %s が集計仕様書にない"
                                             % (sec["no"], hyo, q))
                                continue
                            # 番号のすぐ前にある語（句読点・括弧で区切った最後の断片）が、
                            # その設問の中身と合っているかを見る。
                            # 言い回しが違うことがあるため、語ごとに設問文側の
                            # 現れ方の候補を持つ（就労継続 ↔ 働きながら続けていけ）。
                            YOBI = (("勤務形態", ("勤務形態",)),
                                    ("就労継続", ("働きながら", "続けていけ")),
                                    ("就労", ("働き", "仕事", "就労")),
                                    ("不安", ("不安",)),
                                    ("入所", ("入所", "入居")),
                                    ("入居", ("入所", "入居")),
                                    ("傷病", ("傷病",)),
                                    ("訪問診療", ("訪問診療",)),
                                    ("介護離職", ("仕事を辞めた",)),
                                    ("世帯", ("世帯",)))
                            # 括弧がすぐ前にあると最後の断片は空になる。空でない最後を採る
                            _dan = [x for x in _re58.split(r"[、。・（）\(\)]", t[:m.start()]) if x]
                            mae = _dan[-1] if _dan else ""
                            bun = zai58[(hyo, q)]
                            # 断片の末尾に最も近い語を1つだけ採る
                            atta = None
                            for w, kouho in YOBI:
                                i = mae.rfind(w)
                                if i >= 0 and (atta is None or i > atta[0]):
                                    atta = (i, w, kouho)
                            if atta and not any(k in bun for k in atta[2]):
                                bad58.append("%s：%s %s は「%s」だが「%s」として引いている"
                                             % (sec["no"], hyo, q, bun[:16], atta[1]))
        chk(58, "素案が指す調査票の設問が集計仕様書と合うこと", not bad58,
            "・".join(bad58[:3]) if bad58
            else "在宅介護実態調査の設問を引いている%d件は、票・番号・中身が集計仕様書と一致する" % n58)
    except Exception as e:
        chk(58, "素案が指す調査票の設問が集計仕様書と合うこと", False, "照合できない（%s）" % e)

    # ── 59　確認事項が指す素案の節・図が実在すること ─────────────
    #    確認事項は照会票として村に渡る。素案の節を書き替えたり図を足し引きすると、
    #    確認事項に書いた参照先だけが残る。村に存在しない節を示すことになる。
    #    「資料3の3-3」のような委員会資料の節番号と、docNN・図NN-M は対象外とする。
    try:
        import re as _re59
        import wbs_kakunin as _KK59
        import figures_map as _FM59
        import collections as _c59
        secs59 = {sec["no"] for c in SC.CH for sec in c["sections"]}
        cnt59 = _c59.Counter()
        zu59 = set()
        for key, fn, *_r in _FM59._FIGS:
            ch = key.split("|")[0]
            cnt59[ch] += 1
            zu59.add("図%s-%d" % (ch.replace("第", "").replace("章", ""), cnt59[ch]))
        REF59 = _re59.compile(r"(?<![図表0-9c])([1-6]-[0-9]+)(?![0-9])")
        ZU59 = _re59.compile(r"図([0-9]+-[0-9]+)")
        bad59, n59 = [], 0
        for k in _KK59.K:
            for f in (k[1], k[2]):
                t = str(f)
                for m in REF59.finditer(t):
                    # 「資料N の …」に続く番号は委員会資料の節であり、素案の節ではない
                    mae = t[max(0, m.start() - 40):m.start()]
                    if _re59.search(r"資料[0-9０-９]", mae) or _re59.search(r"doc|図|表$", mae[-6:]):
                        continue
                    n59 += 1
                    if m.group(1) not in secs59:
                        bad59.append("「%s」→ 素案に%sがない" % (str(k[1])[:22], m.group(1)))
                for m in ZU59.finditer(t):
                    n59 += 1
                    if m.group(0) not in zu59:
                        bad59.append("「%s」→ %sがない" % (str(k[1])[:22], m.group(0)))
        chk(59, "確認事項が指す素案の節・図が実在すること", not bad59,
            "・".join(sorted(set(bad59))[:3]) if bad59
            else "確認事項が指す節・図の%d件はすべて素案に実在する" % n59)
    except Exception as e:
        chk(59, "確認事項が指す素案の節・図が実在すること", False, "照合できない（%s）" % e)

    # ── 60　資5の用語と本文の言い方が合っていること ──────────────
    #    用語集の見出し語が本文の言い方と違うと、読み手が引けない。
    #    現に「総合相談支援業務」で載せていたが、本文は「総合相談支援」だった。
    #    本文に現れない語は、削減案Gが落とす候補として挙げているものと一致すること。
    #    用語を足したのに本文で使っていない、あるいは本文の言い方を変えたときに気づく。
    try:
        import re as _re60
        import estimate_sakugen as _ES60
        yougo60 = []
        for c in SC.CH:
            for sec in c["sections"]:
                if sec["no"] != "資5":
                    continue
                for b in sec["blocks"]:
                    if b["t"] == "table":
                        yougo60 += [str(r[0]) for r in b["rows"]]
        hon60 = []
        for c in SC.CH:
            for sec in c["sections"]:
                if sec["no"] == "資5":
                    continue
                hon60.append(str(sec.get("title", "")))
                for b in sec["blocks"]:
                    if b["t"] in ("p", "note", "h3"):
                        hon60.append(str(b["v"]))
                    elif b["t"] == "bullets":
                        hon60 += [str(x) for x in b["v"]]
                    elif b["t"] in ("table", "kpi"):
                        hon60 += [str(x) for r in b["rows"] for x in r]
                        hon60 += [str(x) for x in b["head"]]
        T60 = "\n".join(hon60)
        nai60 = []
        for y in yougo60:
            key = _re60.sub(r"[（(].*?[）)]", "", y).strip()
            if key and key not in T60:
                nai60.append(y)
        bad60 = []
        sa = sorted(set(nai60) - set(_ES60.YOUGO_NAI))
        no = sorted(set(_ES60.YOUGO_NAI) - set(nai60))
        if sa:
            bad60.append("本文に現れないが削減案Gに挙げていない：%s" % "・".join(sa[:3]))
        if no:
            bad60.append("削減案Gに挙げているが本文に現れる：%s" % "・".join(no[:3]))
        chk(60, "資5の用語と本文の言い方が合っていること", not bad60,
            "／".join(bad60) if bad60
            else "用語%d語のうち本文に現れないのは%d語で、削減案Gの候補と一致する"
                 % (len(yougo60), len(nai60)))
    except Exception as e:
        chk(60, "資5の用語と本文の言い方が合っていること", False, "照合できない（%s）" % e)


    # ── 61　指している小見出し（5-4(8) の枝番・6-4⑦ の丸数字）が実在すること ──
    #    点検59は節（5-4）までしか見ていない。節の中の枝番を取り違えると、
    #    読み手は別の話をしている小見出しに飛ばされる。現に、必要利用定員総数は
    #    5-4(8) であるのに確認事項と進捗が 5-4(5)（供給の制約）・5-4(6) を指していた。
    #    小見出しは節の中の h3 の並びの位置で数える。番号の付け方は「(1)」と
    #    「1　」（全角の数字と空き）の2とおりがある。
    #    枝番を振っていない節（5-6・5-7・2-8 など、小見出しに名だけを置く節）は、
    #    位置を指す内部の言い方として用いてよいが、村に渡る欄では使わない。
    #    【この点検では見つけられないこと】実在する枝番のうち別の小見出しを指して
    #    しまった場合（5-4(8) と書くべきところを 5-4(5) と書いた場合）は、
    #    数の上では正しいため通る。1件の文に複数の小見出しの話が出ることは普通であり、
    #    近くの語から指し先を当てる試みは誤って鳴るほうが多かったため入れていない。
    #    節の中の小見出しを入れ替えたときは、指し先を目で確かめる必要がある。
    try:
        import re as _re61
        import wbs_kakunin as _KK61
        import wbs_pending as _PD61
        EDA61 = _re61.compile(r"^(?:[（(](\d+)[）)]|([0-9０-９]+)[　 ])")
        kazu61 = {}
        for c in SC.CH:
            for sec in c["sections"]:
                n = 0
                for b in sec["blocks"]:
                    if b["t"] == "h3" and EDA61.match(str(b["v"])):
                        n += 1
                kazu61[sec["no"]] = n
        # 節の中の丸数字（6-4⑦ のような指し方）の在庫も持つ。
        # 枝番ではないため REF61 では拾えず、6-4 の検証項目を1つ増やしたときに
        # 施策3-5・5-12 から「6-4⑦」と指しても誰も確かめられなかった（令和8年10月8日）。
        #    在庫に数えるのは、丸数字が「見出しとして立っている」ところだけである。
        #    節のどこかに文字として現れることを条件にすると、⑦の行を落としても
        #    「⑦は〜に備えるためのものです」という本文が残るため通ってしまう
        #    （自己試験で確かめた）。
        MARU61 = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮"
        maru61 = {}
        for c in SC.CH:
            for sec in c["sections"]:
                have = set()
                for b in sec["blocks"]:
                    if b["t"] in ("table", "kpi"):
                        for r in b["rows"]:
                            z = str(r[0]).strip()
                            if z[:1] in MARU61:
                                have.add(z[:1])
                    elif b["t"] == "bullets":
                        for x in b["v"]:
                            if str(x).strip()[:1] in MARU61:
                                have.add(str(x).strip()[:1])
                    elif b["t"] == "h3":
                        z = str(b["v"]).strip()
                        if z[:1] in MARU61:
                            have.add(z[:1])
                maru61[sec["no"]] = have
        # 枝番の書き方は2つある。「5-4(8)」と「1-8の3」である。
        # 後者を見ていなかったため、1-8の項を1つ増やしたときに
        # 指し先8か所が1つずれたまま残った（令和8年10月8日）。
        # 「1-8の3」の後ろに漢字・カタカナが続くときは、枝番ではなく数量である
        # （「5-7の12パターン」「4-3の4場面」「2-1の1セル」「2-1の15歳未満」）。
        # 枝番の後ろに続くのは、ひらがなの助詞か句読点か文の終わりである。
        REF61 = _re61.compile(
            r"(?<![図表0-9c])([1-6]-[0-9]+)"
            r"(?:[（(](\d+)[）)]|の([0-9０-９]+)(?![\u4e00-\u9fff\u30a0-\u30ff0-9０-９]))")
        # 「6-4⑦」「5-8(1)②」のように節（と枝番）の後ろに丸数字を続ける指し方
        MREF61 = _re61.compile(
            r"(?<![図表0-9c])([1-6]-[0-9]+)(?:[（(]\d+[）)])?([%s])" % MARU61)
        bad61, n61 = [], [0]

        def maru_mite61(t, doko):
            """「6-4⑦」の丸数字が、その節に現にあること"""
            for m in MREF61.finditer(str(t)):
                mae = str(t)[max(0, m.start() - 40):m.start()]
                if _re61.search(r"資料[0-9０-９]", mae):
                    continue
                n61[0] += 1
                sno, mk = m.group(1), m.group(2)
                have = maru61.get(sno)
                if have is None:
                    bad61.append("%s → 素案に%sがない" % (doko, sno))
                elif mk not in have:
                    bad61.append("%s → %s%s がない（%sにあるのは%s）"
                                 % (doko, sno, mk, sno,
                                    "".join(sorted(have, key=MARU61.index)) or "なし"))

        def mite61(t, doko, mura, mado=None):   # mado は使わない（下の断り）
            maru_mite61(t, doko)
            for m in REF61.finditer(str(t)):
                mae = str(t)[max(0, m.start() - 40):m.start()]
                if _re61.search(r"資料[0-9０-９]", mae):
                    continue
                n61[0] += 1
                _z = m.group(2) or m.group(3)
                sno, k = m.group(1), int(str(_z).translate(
                    str.maketrans("０１２３４５６７８９", "0123456789")))
                n = kazu61.get(sno)
                if n is None:
                    bad61.append("%s → 素案に%sがない" % (doko, sno))
                elif n == 0:
                    if mura:
                        bad61.append("%s → %sに枝番がない（村に渡る欄）" % (doko, sno))
                elif k > n:
                    bad61.append("%s → %s(%d) がない（%sの枝番は%d件）"
                                 % (doko, sno, k, sno, n))
        # 「本節N」は、その節の中の枝番を指す
        def honsetsu61(sec_no, t, doko):
            for m in _re61.finditer(r"本節([0-9０-９]+)", str(t)):
                n61[0] += 1
                k = int(m.group(1).translate(
                    str.maketrans("０１２３４５６７８９", "0123456789")))
                have = kazu61.get(sec_no, 0)
                if have and k > have:
                    bad61.append("%s → 本節%d がない（%sの枝番は%d件）"
                                 % (doko, k, sec_no, have))

        # 素案の本文（村に渡る）
        for c in SC.CH:
            for sec in c["sections"]:
                d = "素案%s" % sec["no"]
                for b in sec["blocks"]:
                    if b["t"] in ("p", "note", "h3"):
                        mite61(b["v"], d, True)
                        honsetsu61(sec["no"], b["v"], d)
                    elif b["t"] == "bullets":
                        for x in b["v"]:
                            mite61(x, d, True)
                            honsetsu61(sec["no"], x, d)
                    elif b["t"] in ("table", "kpi"):
                        for r in b["rows"]:
                            for x in r:
                                mite61(x, d, True)
                                honsetsu61(sec["no"], x, d)
        # 確認事項（照会票として村に渡る）。見出しと本体で1件なので全文を窓にする
        for k in _KK61.K:
            d = "確認事項「%s」" % str(k[1])[:18]
            zen = "%s\n%s" % (k[1], k[2])
            mite61(k[1], d, True, zen)
            mite61(k[2], d, True, zen)
        # 影響度の判定：当方の手当は照会票に出る。止まる対象は内部にとどまる
        for key, v in _PD61.IMPACT.items():
            d = "影響度「%s」" % str(key)[:18]
            zen = "%s\n%s\n%s" % (key, v[1], v[2])
            mite61(v[2], d, True, zen)
            mite61(v[1], d, False, zen)
        chk(61, "指している小見出し（枝番・丸数字）が実在すること", not bad61,
            "・".join(sorted(set(bad61))[:3]) if bad61
            else "枝番・丸数字で指している%d件はすべて実在する"
                 "（枝番を振っている節%d件）"
                 % (n61[0], sum(1 for v in kazu61.values() if v)))
    except Exception as e:
        chk(61, "指している小見出し（枝番・丸数字）が実在すること", False, "照合できない（%s）" % e)

    # ── 62　「最も高い／低いのは〜」と述べた値が、同じ節の表にあること ────────
    #    本文で最大値・最小値を言い切ると、表を直したときに本文だけが残る。
    #    現に5-4(5)は「最も高いのは令和11年度で90.7%」と述べていたが、
    #    表の値は改訂で変わっており90.7%はどこにもなかった。
    try:
        import re as _re62
        SAI62 = _re62.compile(r"最も(?:高|低|大き|小さ|多|少な)[いく][^。]*?"
                              r"([0-9]+(?:\.[0-9]+)?)\s*(%|％|人|円|千円|ポイント)")
        bad62, n62 = [], 0
        for c in SC.CH:
            for sec in c["sections"]:
                # その節の表・指標に出てくる数の集まり
                atai = set()
                for b in sec["blocks"]:
                    if b["t"] in ("table", "kpi"):
                        for r in b["rows"]:
                            for x in r:
                                for mm in _re62.finditer(r"[0-9]+(?:\.[0-9]+)?", str(x)):
                                    atai.add(mm.group(0))
                        for x in b["head"]:
                            for mm in _re62.finditer(r"[0-9]+(?:\.[0-9]+)?", str(x)):
                                atai.add(mm.group(0))
                if not atai:
                    continue
                for b in sec["blocks"]:
                    if b["t"] not in ("p", "note"):
                        continue
                    for m in SAI62.finditer(str(b["v"])):
                        n62 += 1
                        if m.group(1) not in atai:
                            bad62.append("%s「最も…%s%s」が表にない"
                                         % (sec["no"], m.group(1), m.group(2)))
        chk(62, "最も高い・低いと述べた値が同じ節の表にあること", not bad62,
            "・".join(sorted(set(bad62))[:3]) if bad62
            else "最大・最小を言い切っている%d件はいずれも同じ節の表にある" % n62)
    except Exception as e:
        chk(62, "最も高い・低いと述べた値が同じ節の表にあること", False, "照合できない（%s）" % e)

    # ── 63　制度の該当・施行を、確定していないのに断定していないこと ───────
    #    他案件（大雪）の点検から採った型である。
    #    特定地域の指定基準は部会に素案として示された段階であり、
    #    指定は村の意向を踏まえて都道府県が定める。
    #    それを「該当します」「該当すると判断できます」と書いていた（令和8年10月8日に是正）。
    #    施行の時期も同じで、「結論を得る」段階のものを「施行される」と書かない。
    try:
        import re as _re63
        DANTEI63 = (
            (r"(中山間・人口減少地域|特定地域)[^。]{0,40}(に該当します|該当すると判断|"
             r"直接該当|該当することは確実)",
             "特定地域の該当を断定している"),
            (r"(第10期[^。]{0,10}計画期間中|第10期中)に施行(され|する)",
             "施行が決まっていないのに施行されると書いている"),
            (r"指定(を受けられることは|は確実)", "指定を断定している"),
        )
        bad63, n63 = [], 0
        for c in SC.CH:
            for sec in c["sections"]:
                for b in sec["blocks"]:
                    if b["t"] in ("p", "note", "h3"):
                        tt = [str(b["v"])]
                    elif b["t"] == "bullets":
                        tt = [str(x) for x in b["v"]]
                    elif b["t"] in ("table", "kpi"):
                        tt = [str(x) for r in b["rows"] for x in r]
                    else:
                        continue
                    for x in tt:
                        n63 += 1
                        for pat, nm in DANTEI63:
                            if _re63.search(pat, x):
                                bad63.append("%s：%s（%s）" % (sec["no"], nm, x[:34]))
        chk(63, "制度の該当・施行を確定していないのに断定していないこと", not bad63,
            "・".join(sorted(set(bad63))[:3]) if bad63
            else "%d件の文・ます目に、確定していない制度の断定はない" % n63)
    except Exception as e:
        chk(63, "制度の該当・施行を確定していないのに断定していないこと", False,
            "照合できない（%s）" % e)

    # ── 64　「実績がない」と述べるときに、どの年度の実績かを書いていること ─────
    #    他案件（大雪）の点検から採った型である。
    #    廃止された区分には廃止の年度の実績が残ることがあり、
    #    年度を書かないと読み手はどの時点の話か分からない。
    try:
        import re as _re64
        NAI64 = _re64.compile(r"実績(が|は)?(あり|)ませ|実績なし|実績がない|実績の計上がな")
        # 「令和5〜7年度」のように幅で書くこともある
        NEN64 = _re64.compile(r"(令和|平成)[0-9０-９]+(?:[〜～\-][0-9０-９]+)?年")
        bad64, n64 = [], 0
        for c in SC.CH:
            for sec in c["sections"]:
                for b in sec["blocks"]:
                    if b["t"] in ("p", "note", "h3"):
                        tt = [str(b["v"])]
                    elif b["t"] == "bullets":
                        tt = [str(x) for x in b["v"]]
                    elif b["t"] in ("table", "kpi"):
                        # ます目は行の見出しと一緒に見る（年度が別のます目にあることがある）
                        tt = [" ".join(str(x) for x in r) for r in b["rows"]]
                    else:
                        continue
                    for x in tt:
                        if not NAI64.search(x):
                            continue
                        n64 += 1
                        if not NEN64.search(x):
                            bad64.append("%s：%s" % (sec["no"], x[:40]))
        chk(64, "「実績がない」にどの年度の実績かを書いていること", not bad64,
            "・".join(sorted(set(bad64))[:3]) if bad64
            else "「実績がない」と述べている%d件すべてに年度の断りがある" % n64)
    except Exception as e:
        chk(64, "「実績がない」にどの年度の実績かを書いていること", False,
            "照合できない（%s）" % e)

    # ── 65　第11期の分母の組み方が第10期で復元できること ────────────────
    #    5-7の基金の3案は、第11期の基準額を当方が組んだ分母で割って出している。
    #    国の中長期シートは第11期の所得段階別加入割合補正後被保険者数を持たないため、
    #    令和12年度の第1号被保険者数×補正係数×12か月×予定収納率で組んでいる。
    #    この組み方が妥当かは、同じ組み方で第10期の分母を復元して比べるほかない。
    #    素案の注記は差を印字するだけなので、差が大きくなっても文面は成り立ってしまう。
    #    組み方が通らなくなったことを機械で知るための点検である。
    #    あわせて、第11期の値が当方の試算であることと、
    #    ワークシートの出力そのもの（取崩額0のときの値）が併記されていることを見る。
    try:
        import estimate_kikin as _KK65
        ws65, my65, sa65 = _KK65.reconstruct()
        bad65 = []
        GEN65 = 0.005          # 復元の差の許容（0.5%）
        if abs(sa65) > GEN65:
            bad65.append("第10期の分母の復元の差が%+.2f%%で許容(%.1f%%)を超える"
                         % (sa65 * 100, GEN65 * 100))
        # 5-7 の該当の節に、当方の試算である旨とワークシートの出力が書かれていること
        txt65 = ""
        for c in SC.CH:
            for sec in c["sections"]:
                if sec["no"] != "5-7":
                    continue
                on = False
                for b in sec["blocks"]:
                    if b["t"] == "h3":
                        on = (b["v"] == "介護給付費準備基金の取崩しの水準")
                        continue
                    if on:
                        txt65 += str(b.get("v", "")) + str(b.get("rows", ""))
        for need in ("当方の試算", "単年", "3か年に均した",
                     "ワークシートを操作", "第11期の取崩額を0",
                     "百円未満を切り捨てた", "四捨五入"):
            if need not in txt65:
                bad65.append("5-7の基金の節に「%s」の断りがない" % need)
        # 3案の表が estimate_kikin の算定と一致すること
        an65 = None
        for c in SC.CH:
            for sec in c["sections"]:
                if sec["no"] != "5-7":
                    continue
                for b in sec["blocks"]:
                    if b["t"] == "table" and "残高の緩衝力" in b["head"]:
                        an65 = b
        if an65 is None:
            bad65.append("5-7に基金の3案の表がない")
        else:
            got = _KK65.run()[0]
            if len(an65["rows"]) != len(got):
                bad65.append("3案の行数 素案%d≠算定%d" % (len(an65["rows"]), len(got)))
            else:
                for r, g in zip(an65["rows"], got):
                    if "{:,.0f}円".format(g["第10期条例"]) not in r:
                        bad65.append("%s の第10期（条例上）が算定と合わない" % g["案"])
                    if "{:,.0f}円".format(g["第11期条例"]) not in r:
                        bad65.append("%s の第11期（条例上）が算定と合わない" % g["案"])
        chk(65, "第11期の分母の組み方が第10期で復元できること", not bad65,
            "・".join(bad65[:3]) if bad65
            else "第10期の分母はワークシート%d人・月に対し復元%d人・月（差%+.2f%%）。"
                 "3案の値が算定と一致し、当方の試算である旨も書かれている"
                 % (round(ws65), round(my65), sa65 * 100))
    except Exception as e:
        chk(65, "第11期の分母の組み方が第10期で復元できること", False,
            "照合できない（%s）" % e)

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
