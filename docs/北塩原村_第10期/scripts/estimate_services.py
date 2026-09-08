# -*- coding: utf-8 -*-
"""
北塩原村 第10期介護保険事業計画
サービス見込量推計の枠組み

【地域包括ケア「見える化」システムのデータ定義（本スクリプトで確定させた前提）】
  ・D1〜D4、D32系の「受給者数」は年度延べ（12か月合計）。月平均は ÷12。
    例）令和7年度 在宅サービス受給者数（要支援1）= 321人（延べ）→ 月平均 26.75人
  ・D32系の「受給率」    = 受給者数（延べ）÷ 第1号被保険者数（延べ）× 100
                          （分母は要介護度によらず第1号被保険者数の全体）
  ・D45系の「サービス利用率」= 受給者数（延べ）÷ 認定者数（要介護度別・延べ）× 100
                          （分母が要介護度別の認定者数。計画で使う「受給率」はこちら）
  ・検証：令和7年度 要支援1 在宅 321 ÷ 52.5% ÷ 12 = 51.0人 ＝ B4-a 令和8年3月末 51人
          令和5〜7年度の全要介護度で年度末認定者数と ±3人以内で一致することを確認済み。

【推計の連鎖】
  ① 将来人口（厚労省配布・住民基本台帳ベース）
  ② 第1号被保険者数（前期・後期別補正係数）        → estimate_population.py
  ③ 要介護度別認定者数                              → estimate_certified.py
  ④ サービス区分別（在宅／居住系／施設）受給者数    ← 本スクリプト
      受給者数 = 認定者数(要介護度別) × 利用率(要介護度別・直近3か年平均)
  ⑤ サービス種類別受給者数
      受給者数 = 認定者数(要介護度別) × 認定者ベース受給率(要介護度別・直近3か年平均)
  ⑥ サービス種類別利用量
      利用量   = 受給者数 × 受給者1人あたり利用日数・回数(直近3か年平均)
  ⑦ 給付費
      給付費   = 受給者数 × 受給者1人あたり給付月額(直近3か年平均) × 12
"""
import csv, os, collections

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(BASE, 'data')

# 推計に用いる認定者数の案（第2回策定委員会で確定）
ADOPTED_CASE = 'B'          # A案=厚労省配布データそのまま / B案=令和8年度実績で補正
ADOPTED_BASE = '令和8年'    # 認定率の基準（令和8年 / 3か年平均）

DEGREES = ['要支援1', '要支援2', '要介護1', '要介護2', '要介護3', '要介護4', '要介護5']
DEG_ZEN = {'要支援1': '要支援１', '要支援2': '要支援２',
           '要介護1': '要介護１', '要介護2': '要介護２', '要介護3': '要介護３',
           '要介護4': '要介護４', '要介護5': '要介護５'}
YEARS_EST = ['令和9年', '令和10年', '令和11年', '令和12年', '令和17年', '令和22年']
YEARS_ACT = ['2023', '2024', '2025']       # 実績3か年（令和5・6・7年度）
YEAR_LABEL = {'2023': '令和5年度', '2024': '令和6年度', '2025': '令和7年度'}

# 第10期の計画期間（令和9〜11年度）と中長期推計年（令和12・17・22年度）
PLAN_YEARS = ['令和9年', '令和10年', '令和11年']

# 区分別 利用率の指標コード
KUBUN = [
    ('在宅サービス',   'D4', 'D45-a'),
    ('居住系サービス', 'D3', 'D45-b'),
    ('施設サービス',   'D2', 'D45-c'),
]

# サービス種類別 指標コードの対応表
#  (サービス名, D32系=受給者数, D46系=1人あたり利用量, 利用量の単位, D17系=1人あたり給付月額, 区分)
SERVICES = [
    ('訪問介護',                         'D32-a', 'D46-a', '回', 'D17-a', '居宅'),
    ('訪問入浴介護',                     'D32-b', 'D46-b', '回', 'D17-b', '居宅'),
    ('訪問看護',                         'D32-c', 'D46-c', '回', 'D17-c', '居宅'),
    ('訪問リハビリテーション',           'D32-d', 'D46-d', '回', 'D17-d', '居宅'),
    ('居宅療養管理指導',                 'D32-e', None,    '－', 'D17-e', '居宅'),
    ('通所介護',                         'D32-f', 'D46-e', '日', 'D17-f', '居宅'),
    ('通所リハビリテーション',           'D32-g', 'D46-f', '日', 'D17-g', '居宅'),
    ('短期入所生活介護',                 'D32-h', 'D46-g', '日', 'D17-h', '居宅'),
    ('短期入所療養介護',                 'D32-i', 'D46-h', '日', 'D17-i', '居宅'),
    ('福祉用具貸与',                     'D32-j', None,    '－', 'D17-j', '居宅'),
    ('定期巡回・随時対応型訪問介護看護', 'D32-k', None,    '－', None,    '地域密着型'),
    ('夜間対応型訪問介護',               'D32-l', None,    '－', None,    '地域密着型'),
    ('認知症対応型通所介護',             'D32-m', 'D46-i', '日', 'D17-o', '地域密着型'),
    ('小規模多機能型居宅介護',           'D32-n', None,    '－', 'D17-p', '地域密着型'),
    ('看護小規模多機能型居宅介護',       'D32-o', None,    '－', None,    '地域密着型'),
    ('地域密着型通所介護',               'D32-s', 'D46-j', '回', 'D17-t', '地域密着型'),
]

# 見える化に受給者数の系列がなく、給付費のみ把握できるサービス
#  （第1号被保険者1人あたり給付月額 D13系 × 第1号被保険者数 で給付費を推計する）
SERVICES_KYUFU_ONLY = [
    ('特定施設入居者生活介護',                     'D13-q', '居住系'),
    ('認知症対応型共同生活介護',                   'D13-w', '居住系'),
    ('地域密着型特定施設入居者生活介護',           'D13-x', '居住系'),
    ('介護老人福祉施設',                           'D13-a', '施設'),
    ('地域密着型介護老人福祉施設入所者生活介護',   'D13-d', '施設'),
    ('介護老人保健施設',                           'D13-b', '施設'),
    ('介護医療院',                                 'D13-aa', '施設'),
    ('介護療養型医療施設',                         'D13-c', '施設'),
    ('特定福祉用具販売',                           'D13-o', '居宅'),
    ('住宅改修',                                   'D13-p', '居宅'),
    ('介護予防支援・居宅介護支援',                 'D13-r', '居宅'),
]


def load_mieruka():
    path = os.path.join(DATA, 'mieruka_tidy.csv')
    rows = list(csv.DictReader(open(path, encoding='utf-8-sig')))
    idx = {}
    for r in rows:
        if r['region'] != '北塩原村':
            continue
        idx.setdefault((r['code'], r['indicator'], r['year']), r['value'])
    return idx


def fval(idx, code, ind, year):
    v = idx.get((code, ind, year))
    if v in (None, '', '-', '－'):
        return None
    try:
        return float(v)
    except ValueError:
        return None


def find(idx, code, year, degree=None, total=False):
    """指標名の表記ゆれを吸収して値を取り出す。
       degree を指定した場合は「（要介護1）」「(要介護1)」「要介護1」のいずれの
       表記でも拾う。total=True のときは「合計」系を拾う。"""
    key = '合計' if total else degree
    if key is None:
        return None
    cands = []
    for (c, ind, y), v in idx.items():
        if c != code or y != year:
            continue
        norm = ind.replace('（', '(').replace('）', ')').replace(' ', '').replace('\u3000', '')
        if norm == key or norm.endswith('(' + key + ')'):
            cands.append(v)
    for v in cands:
        try:
            return float(v)
        except (TypeError, ValueError):
            continue
    return None


def avg3(vals):
    vs = [v for v in vals if v is not None]
    return sum(vs) / len(vs) if vs else None


def load_certified():
    """認定者数推計（採用案）を読み込む"""
    path = os.path.join(DATA, '第10期_要介護度別認定者数推計.csv')
    out = {}
    ins = {}
    for r in csv.DictReader(open(path, encoding='utf-8-sig')):
        if r['案'] != ADOPTED_CASE or r['認定率の基準'] != ADOPTED_BASE:
            continue
        y = r['年']
        out[y] = {d: float(r[DEG_ZEN[d]]) for d in DEGREES}
        out[y]['経過的要介護'] = float(r['経過的要介護'])
        ins[y] = float(r['第1号被保険者計'])
    return out, ins


def main():
    idx = load_mieruka()
    cert, ins = load_certified()

    # ------------------------------------------------------------------
    # ④ サービス区分別（在宅／居住系／施設）受給者数
    # ------------------------------------------------------------------
    riyoritsu = {}   # (区分, 要介護度) -> 直近3か年平均の利用率(%)
    jisseki_k = {}   # (区分, 要介護度, year) -> 実績の月平均受給者数
    for kubun, dcode, rcode in KUBUN:
        for d in DEGREES:
            rates, nums = [], {}
            for y in YEARS_ACT:
                rt = fval(idx, rcode, f'サービス利用率（{kubun}）({d})', y)
                nm = fval(idx, dcode, f'受給者数（{kubun}）（{d}）', y)
                rates.append(rt)
                if nm is not None:
                    nums[y] = nm / 12.0
            riyoritsu[(kubun, d)] = avg3(rates)
            for y, v in nums.items():
                jisseki_k[(kubun, d, y)] = v

    rows_k = []
    for kubun, _, _ in KUBUN:
        for d in DEGREES:
            rec = {'区分': kubun, '要介護度': d,
                   '利用率(3か年平均・%)': round(riyoritsu[(kubun, d)] or 0.0, 1)}
            for y in YEARS_ACT:
                rec[YEAR_LABEL[y] + '実績'] = round(jisseki_k.get((kubun, d, y), 0.0), 1)
            for y in YEARS_EST:
                n = cert[y][d] * (riyoritsu[(kubun, d)] or 0.0) / 100.0
                rec[y + '度'] = round(n, 1)
            rows_k.append(rec)
    # 合計行
    for kubun, _, _ in KUBUN:
        rec = {'区分': kubun, '要介護度': '計', '利用率(3か年平均・%)': ''}
        for key in [YEAR_LABEL[y] + '実績' for y in YEARS_ACT] + [y + '度' for y in YEARS_EST]:
            rec[key] = round(sum(r[key] for r in rows_k if r['区分'] == kubun), 1)
        rows_k.append(rec)

    cols_k = ['区分', '要介護度', '利用率(3か年平均・%)'] \
        + [YEAR_LABEL[y] + '実績' for y in YEARS_ACT] + [y + '度' for y in YEARS_EST]
    write_csv(os.path.join(DATA, '第10期_サービス区分別受給者数推計.csv'), cols_k, rows_k)

    # ------------------------------------------------------------------
    # ⑤⑥⑦ サービス種類別 受給者数・利用量・給付費
    # ------------------------------------------------------------------
    rows_s = []
    for name, c32, c46, unit, c17, kubun in SERVICES:
        # 認定者ベース受給率（要介護度別・3か年平均）
        rate_d = {}
        for d in DEGREES:
            rs = []
            for y in YEARS_ACT:
                nm = juk_num(idx, c32, name, d, y)          # 受給者数（延べ）
                cn = restore_certified(idx, d, y)           # 認定者数（延べ）
                if nm is not None and cn:
                    rs.append(nm / cn * 100.0)
            rate_d[d] = avg3(rs)

        # 1人あたり利用量（要介護度別・3か年平均）
        #   要支援の予防給付は月額包括報酬で日数・回数が計上されないサービスがあるため、
        #   合計値を使わず要介護度別に積み上げる。
        use_d = {}
        if c46:
            for d in DEGREES:
                use_d[d] = avg3([find(idx, c46, y, degree=d) for y in YEARS_ACT])
        per_use = avg3([find(idx, c46, y, total=True) for y in YEARS_ACT]) if c46 else None
        # 1人あたり給付月額（3か年平均）
        per_yen = avg3([fval(idx, c17, f'受給者1人あたり給付月額（{name}）', y)
                        for y in YEARS_ACT]) if c17 else None

        # 実績（月平均受給者数・月あたり利用量）
        act, act_q = {}, {}
        for y in YEARS_ACT:
            v = juk_num(idx, c32, name, None, y, total=True)
            act[y] = round(v / 12.0, 1) if v is not None else 0.0
            if c46:
                q = 0.0
                for d in DEGREES:
                    nd = juk_num(idx, c32, name, d, y)
                    ud = find(idx, c46, y, degree=d)
                    if nd and ud:
                        q += nd / 12.0 * ud
                act_q[y] = int(round(q))
            else:
                act_q[y] = '－'

        rec = {'区分': kubun, 'サービス種類': name, '利用量の単位': unit,
               '受給率(3か年平均・%)': round(sum(v for v in rate_d.values() if v) / 1.0, 1) if any(rate_d.values()) else 0.0,
               '1人あたり利用量(3か年平均)': round(per_use, 1) if per_use else '－',
               '_use_d': use_d,
               '1人あたり給付月額(円・3か年平均)': int(round(per_yen)) if per_yen else '－'}
        rec['受給率(3か年平均・%)'] = ''   # 要介護度別のため合計欄は空
        for y in YEARS_ACT:
            rec[YEAR_LABEL[y] + '実績(人/月)'] = act[y]
            rec[YEAR_LABEL[y] + '実績(利用量/月)'] = act_q[y]
        for y in YEARS_EST:
            n_d = {d: cert[y][d] * (rate_d[d] or 0.0) / 100.0 for d in DEGREES}
            n = sum(n_d.values())
            rec[y + '度(人/月)'] = round(n, 1)
            if c46:
                q = sum(n_d[d] * (use_d.get(d) or 0.0) for d in DEGREES)
                rec[y + '度(利用量/月)'] = int(round(q))
            else:
                rec[y + '度(利用量/月)'] = '－'
            rec[y + '度(給付費・千円/年)'] = int(round(n * per_yen * 12 / 1000)) if per_yen else '－'
        rows_s.append(rec)

    cols_s = ['区分', 'サービス種類', '利用量の単位',
              '1人あたり利用量(3か年平均)', '1人あたり給付月額(円・3か年平均)']
    for y in YEARS_ACT:
        cols_s += [YEAR_LABEL[y] + '実績(人/月)', YEAR_LABEL[y] + '実績(利用量/月)']
    for y in YEARS_EST:
        cols_s += [y + '度(人/月)', y + '度(利用量/月)', y + '度(給付費・千円/年)']
    for r in rows_s:
        r.pop('受給率(3か年平均・%)', None)
        r.pop('_use_d', None)
    write_csv(os.path.join(DATA, '第10期_サービス種類別見込量.csv'), cols_s, rows_s)

    # ------------------------------------------------------------------
    # 給付費のみ把握できるサービス（D13系 × 第1号被保険者数）
    # ------------------------------------------------------------------
    rows_y = []
    for name, c13, kubun in SERVICES_KYUFU_ONLY:
        per = avg3([fval(idx, c13, f'第１号被保険者１人あたり給付月額（{name}）', y)
                    for y in YEARS_ACT])
        rec = {'区分': kubun, 'サービス種類': name,
               '第1号1人あたり給付月額(円・3か年平均)': int(round(per)) if per else 0}
        for y in YEARS_ACT:
            v = fval(idx, c13, f'第１号被保険者１人あたり給付月額（{name}）', y)
            rec[YEAR_LABEL[y] + '実績(円/月)'] = int(v) if v is not None else 0
        for y in YEARS_EST:
            rec[y + '度(給付費・千円/年)'] = int(round((per or 0) * ins[y] * 12 / 1000))
        rows_y.append(rec)
    # 施設・居住系の種類別人数は見える化に系列がないため、区分別受給者数を
    # 給付費の構成比で按分した「参考値」を併記する（サービス種類ごとに単価が
    # 異なるため実態とずれる。確定値は村の介護保険事業状況報告の月報による）。
    for kubun_lab, kubun_key in [('居住系', '居住系サービス'), ('施設', '施設サービス')]:
        grp = [r for r in rows_y if r['区分'] == kubun_lab]
        for y in YEARS_EST:
            tot_yen = sum(r[y + '度(給付費・千円/年)'] for r in grp)
            tot_num = [x for x in rows_k
                       if x['区分'] == kubun_key and x['要介護度'] == '計'][0][y + '度']
            for r in grp:
                r[y + '度(人/月・参考)'] = (round(tot_num * r[y + '度(給付費・千円/年)'] / tot_yen, 1)
                                        if tot_yen else 0.0)
        for y in YEARS_ACT:
            tot_yen = sum(r[YEAR_LABEL[y] + '実績(円/月)'] for r in grp)
            tot_num = [x for x in rows_k
                       if x['区分'] == kubun_key and x['要介護度'] == '計'][0][YEAR_LABEL[y] + '実績']
            for r in grp:
                r[YEAR_LABEL[y] + '実績(人/月・参考)'] = (
                    round(tot_num * r[YEAR_LABEL[y] + '実績(円/月)'] / tot_yen, 1) if tot_yen else 0.0)

    cols_y = ['区分', 'サービス種類', '第1号1人あたり給付月額(円・3か年平均)'] \
        + [YEAR_LABEL[y] + '実績(円/月)' for y in YEARS_ACT] \
        + [YEAR_LABEL[y] + '実績(人/月・参考)' for y in YEARS_ACT] \
        + [y + '度(給付費・千円/年)' for y in YEARS_EST] \
        + [y + '度(人/月・参考)' for y in YEARS_EST]
    write_csv(os.path.join(DATA, '第10期_給付費のみ把握サービス.csv'), cols_y, rows_y)

    # ------------------------------------------------------------------
    # コンソール出力（検算用）
    # ------------------------------------------------------------------
    print('■ サービス区分別 受給者数（人／月）')
    hdr = ['区分'] + [YEAR_LABEL[y] for y in YEARS_ACT] + [y + '度' for y in PLAN_YEARS]
    print('  ' + ' | '.join(f'{h:>10s}' for h in hdr))
    for kubun, _, _ in KUBUN:
        r = [x for x in rows_k if x['区分'] == kubun and x['要介護度'] == '計'][0]
        vals = [r[YEAR_LABEL[y] + '実績'] for y in YEARS_ACT] + [r[y + '度'] for y in PLAN_YEARS]
        print(f"  {kubun:>10s} | " + ' | '.join(f'{v:>10.1f}' for v in vals))

    print()
    print('■ サービス種類別 見込量（令和9年度）')
    for r in rows_s:
        if r['令和9年度(人/月)'] < 0.05:
            continue
        print(f"  {r['サービス種類']:<34s} {r['令和9年度(人/月)']:>6.1f}人/月  "
              f"利用量 {str(r['令和9年度(利用量/月)']):>7s}{r['利用量の単位']}/月  "
              f"給付費 {str(r['令和9年度(給付費・千円/年)']):>9s}千円/年")

    print()
    print('■ 給付費のみ把握のサービス（令和9年度）')
    for r in rows_y:
        ref = r.get('令和9年度(人/月・参考)')
        ref_s = f"{ref:>6.1f}人/月（参考）" if ref is not None else ''
        print(f"  {r['サービス種類']:<40s} {r['令和9年度(給付費・千円/年)']:>10,d} 千円/年  {ref_s}")

    print()
    print('■ 給付費の検算（千円/年）')
    hdr = ['区分'] + [YEAR_LABEL[y] for y in YEARS_ACT] + [y + '度' for y in PLAN_YEARS]
    for y in PLAN_YEARS:
        tot = sum(r[y + '度(給付費・千円/年)'] for r in rows_s
                  if isinstance(r[y + '度(給付費・千円/年)'], int)) \
            + sum(r[y + '度(給付費・千円/年)'] for r in rows_y)
        print(f'  {y}度 サービス諸費（推計）              {tot:>12,d}')
    # 実績（介護サービス等諸費＋介護予防サービス等諸費）
    # D48系は「○年3月末」＝前年度の決算。year=2024 が令和5年度決算にあたる。
    for y, lab in [('2022', '令和3年度'), ('2023', '令和4年度'), ('2024', '令和5年度')]:
        a = fval(idx, 'D48-b', '介護サービス等諸費', y)
        b = fval(idx, 'D48-b', '介護予防サービス等諸費', y)
        if a is not None:
            print(f'  {lab} サービス諸費（決算）              {int((a + (b or 0)) / 1000):>12,d}')
    print()
    print('■ 補助給付・地域支援事業（決算・千円/年）※推計は保険料算定時に別途')
    for y, lab in [('2023', '令和4年度'), ('2024', '令和5年度')]:
        tok = fval(idx, 'D48-b', '特定入所者介護サービス等費', y)
        kog = fval(idx, 'D48-b', '高額介護サービス等費', y)
        gas = fval(idx, 'D48-b', '高額医療合算介護サービス等費', y)
        tes = fval(idx, 'D48-b', '審査支払手数料', y)
        chi = fval(idx, 'D48-c', '合計', y)
        if tok is None:
            continue
        print(f'  {lab}: 特定入所者 {int(tok/1000):>7,d} / 高額介護 {int((kog or 0)/1000):>6,d} / '
              f'高額医療合算 {int((gas or 0)/1000):>5,d} / 審査支払手数料 {int((tes or 0)/1000):>4,d} / '
              f'地域支援事業 {int((chi or 0)/1000):>7,d}')


def ins_num(idx, year):
    """第1号被保険者数（年度延べ）。D32系は延べ、D4系は月平均で格納されている。"""
    v = fval(idx, 'D4', '第1号被保険者数', year)
    return v * 12 if v is not None else None


def juk_num(idx, code, name, degree, year, total=False):
    """サービス種類別の受給者数（年度延べ）を返す。
       受給者数の系列がない指標は「受給率(%) × 第1号被保険者数（延べ）」で復元する。"""
    key = '合計' if total else degree
    # ① 受給者数の系列があればそれを使う
    for (c, ind, y), v in idx.items():
        if c != code or y != year:
            continue
        if not ind.startswith('受給者数'):
            continue
        norm = ind.replace('（', '(').replace('）', ')').replace(' ', '')
        if norm.endswith('(' + key + ')'):
            try:
                return float(v)
            except (TypeError, ValueError):
                return None
    # ② 受給率（%）から復元
    rt = find(idx, code, year, degree=degree, total=total)
    ins = ins_num(idx, year)
    if rt is not None and ins:
        return rt / 100.0 * ins
    return None


_CERT_CACHE = {}


def restore_certified(idx, degree, year):
    """要介護度別の年度延べ認定者数を復元する。
       受給者数（区分別）÷ 利用率（区分別）で復元し、3区分の平均をとる。"""
    key = (degree, year)
    if key in _CERT_CACHE:
        return _CERT_CACHE[key]
    cands = []
    for kubun, dcode, rcode in KUBUN:
        nm = fval(idx, dcode, f'受給者数（{kubun}）（{degree}）', year)
        rt = fval(idx, rcode, f'サービス利用率（{kubun}）({degree})', year)
        if nm and rt:
            cands.append(nm / (rt / 100.0))
    v = sum(cands) / len(cands) if cands else None
    _CERT_CACHE[key] = v
    return v


def write_csv(path, cols, rows):
    with open(path, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.DictWriter(f, fieldnames=cols, extrasaction='ignore')
        w.writeheader()
        for r in rows:
            w.writerow(r)
    print(f'  -> {os.path.relpath(path, BASE)} ({len(rows)}行)')


if __name__ == '__main__':
    main()
