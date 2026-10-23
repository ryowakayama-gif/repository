# -*- coding: utf-8 -*-
BASE={"13mm":1826,"20mm":1826,"25mm":3289}     # 税込・基本水量10㎥込
JURYO=284                                      # 税込・1㎥あたり（10㎥超過分）
FREE=10
RATES=[("現行",0.0),("10％改定",0.10),("20％改定",0.20),("30％改定",0.30)]
def trunc(v): return int(v)                    # 市のExcelは切捨て（1826×1.3=2373.8→2373）
print("■ むつ市公式料金表（令和8年7月〜 一般用・税込）")
print("   基本料金：13mm 1,826円／20mm 1,826円／25mm 3,289円（いずれも基本水量10㎥込）")
print(f"   従量料金：{JURYO}円/㎥（10㎥超過分）  プール用・船舶用は198円/㎥")
print("   減額支援：令和8年7〜9月使用分、口径13〜25mmの基本料金を1か月1,650円減額")
print()
print("■ 市のExcel記載の基本料金との突合")
for d in ["13mm","25mm"]:
    row=[trunc(BASE[d]*(1+r)) for _,r in RATES]
    print(f"   {d}: "+" / ".join(f"{v:,}円" for v in row))
print("   → 市のExcel記載（13,20mm: 1826→2008→2191→2373 ／ 25mm: 3289→3617→3946→4275）と一致")
print()
print("■ 月使用水量別の請求額（税込・円）と現行からの増加額")
for d in ["13mm・20mm","25mm"]:
    b=BASE["13mm"] if d.startswith("13") else BASE["25mm"]
    print(f"\n   【口径 {d}】")
    print(f"      {'区分':<10}{'基本料金':>9}{'従量単価':>9}"+"".join(f"{'月'+str(v)+'㎥':>12}" for v in (10,20,30)))
    cur={}
    for nm,r in RATES:
        bb=trunc(b*(1+r)); jj=trunc(JURYO*(1+r))
        bills={v: bb+max(0,v-FREE)*jj for v in (10,20,30)}
        if r==0: cur=dict(bills)
        s=f"      {nm:<10}{bb:>9,}{jj:>9,}"
        for v in (10,20,30):
            inc=bills[v]-cur[v]
            s+=f"{bills[v]:>7,}" + (f"(+{inc:,})" if inc else "(   — )")
        print(s)
print("\n■ 大口利用者の例（口径100mm・月1,000㎥）")
b=73150
for nm,r in RATES:
    bb=trunc(b*(1+r)); jj=trunc(JURYO*(1+r))
    bill=bb+(1000-0)*jj   # 40mm以上は基本水量の記載がないため全量従量と仮置き
    print(f"   {nm:<10} 基本{bb:>8,}円 ＋ 従量{jj:>4,}円×1,000㎥ = {bill:>10,}円")
print("   ※40mm以上は料金表に基本水量の記載がないため、全量を従量計算した仮置き。市に確認要")
