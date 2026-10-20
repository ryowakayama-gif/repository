exec(open("/tmp/claude-0/-home-user-repository/05ae8dfb-f2f7-5c38-805c-45df053cf245/scratchpad/work/recalc.py").read().split("# --- 検証")[0])
import openpyxl, warnings
warnings.filterwarnings('ignore')
SP="/tmp/claude-0/-home-user-repository/05ae8dfb-f2f7-5c38-805c-45df053cf245/scratchpad"
A=openpyxl.load_workbook(f"{SP}/pat/A_現行基準.xlsx",data_only=True)
B=openpyxl.load_workbook(f"{SP}/pat/B_是正基準.xlsx",data_only=True)
Y=["R7","R8","R9","R10","R11","R12","R13","R14","R15","R16"]

print("■ 検証1：必要改定率 B6 は是正前後で不変か")
for s in ["0%","1%","2%","3%"]:
    sn=f"SIM_R8予算置換_{s}"
    a=A[sn]["B6"].value; b=B[sn]["B6"].value
    print(f"  {sn:<22} A={a:.10f}  B={b:.10f}  {'不変' if abs(a-b)<1e-12 else '★変化'}")

print("\n■ 検証2：LibreOffice再計算値（B）と Python是正モデル の突合  SIM_R8予算置換_0%")
rate=A["SIM_R8予算置換_0%"]["B6"].value
py=run(rate, True)
ws=B["SIM_R8予算置換_0%"]
ok=True
for key,row,nm in [("r9",9,"営業収益"),("r28",28,"純損益"),("r31",31,"経常収支比率"),
                   ("r32",32,"料金回収率"),("r54",54,"4条収支不足"),("r59",59,"現預金"),("r72",72,"補填財源残高")]:
    diffs=[]
    for i,col in enumerate(range(2,12)):
        x=ws.cell(row=row,column=col).value
        if isinstance(x,(int,float)) and abs(py[key][i]-x)>0.6:
            diffs.append(f"{Y[i]}:py={py[key][i]:,.1f}/xls={x:,.1f}")
    print(f"  {nm:<12} {'一致' if not diffs else '★差異 '+str(diffs)}")
    if diffs: ok=False

print("\n■ 検証3：是正による変化（A→B）SIM_R8予算置換_0%（千円）")
for row,nm in [(9,"営業収益"),(28,"純損益"),(72,"補填財源残高"),(59,"現預金")]:
    a=[A["SIM_R8予算置換_0%"].cell(row=row,column=c).value for c in range(2,12)]
    b=[B["SIM_R8予算置換_0%"].cell(row=row,column=c).value for c in range(2,12)]
    d=[(x-y) if isinstance(x,(int,float)) and isinstance(y,(int,float)) else None for x,y in zip(a,b)]
    print(f"  {nm:<10}"+"".join(f"{v:>11,.0f}" if v is not None else f"{'-':>11}" for v in d))

print("\n■ 検証4：料金回収率・4条不足が是正前後で不変か（全4シナリオ）")
for s in ["0%","1%","2%","3%"]:
    sn=f"SIM_R8予算置換_{s}"
    same32=all(abs((A[sn].cell(row=32,column=c).value or 0)-(B[sn].cell(row=32,column=c).value or 0))<1e-6 for c in range(2,12))
    same54=all(abs((A[sn].cell(row=54,column=c).value or 0)-(B[sn].cell(row=54,column=c).value or 0))<1e-6 for c in range(2,12))
    same33=all(abs((A[sn].cell(row=33,column=c).value or 0)-(B[sn].cell(row=33,column=c).value or 0))<1e-6 for c in range(2,12))
    print(f"  {sn:<22} 料金回収率={'不変' if same32 else '★変化'} 資産維持費込={'不変' if same33 else '★変化'} 4条不足={'不変' if same54 else '★変化'}")
print(f"\n総合: {'全検証パス' if ok else '要確認'}")
