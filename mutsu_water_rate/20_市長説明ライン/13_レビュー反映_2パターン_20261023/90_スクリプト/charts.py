# -*- coding: utf-8 -*-
import openpyxl, warnings, decimal
warnings.filterwarnings('ignore')
from pptx import Presentation
from pptx.util import Pt, Cm, Emu
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_TICK_MARK
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
WD="/tmp/claude-0/-home-user-repository/05ae8dfb-f2f7-5c38-805c-45df053cf245/scratchpad"
exec(open(f"{WD}/w7/recalc.py").read().split("# --- 検証")[0])
def rnd(v): return int(decimal.Decimal(v/1000).quantize(0,rounding=decimal.ROUND_HALF_UP))
FONT="BIZ UDPゴシック"
NAVY=RGBColor(0x1C,0x3C,0x60); RED=RGBColor(0xBE,0x41,0x41)
TEAL=RGBColor(0x2E,0x86,0x8A); AMBER=RGBColor(0xC8,0x85,0x2B); INK=RGBColor(0x19,0x1C,0x23)

def ea(run):
    rPr=run._r.get_or_add_rPr()
    for t in ("a:latin","a:ea","a:cs"):
        e=rPr.find(qn(t))
        if e is None: e=rPr.makeelement(qn(t),{}); rPr.append(e)
        e.set("typeface",FONT)

def txt(slide,left,top,w,h,s,size,bold=False,color=INK):
    tb=slide.shapes.add_textbox(Cm(left),Cm(top),Cm(w),Cm(h))
    p=tb.text_frame.paragraphs[0]; r=p.add_run(); r.text=s
    r.font.size=Pt(size); r.font.bold=bold; r.font.name=FONT; r.font.color.rgb=color; ea(r)
    return tb

def chart_slide(prs,title,sub,cats,series,page,note,pct=False):
    s=prs.slides.add_slide(prs.slide_layouts[6])
    for sh in list(s.shapes): sh._element.getparent().remove(sh._element)
    txt(s,1.14,0.64,30.7,1.14,title,22,True,NAVY)
    txt(s,1.22,1.98,30.5,0.64,sub,12,False,INK)
    txt(s,32.0,18.03,1.4,0.46,str(page),10,False,INK)
    cd=CategoryChartData(); cd.categories=cats
    for nm,vals in series: cd.add_series(nm,vals)
    gf=s.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS,Cm(1.6),Cm(3.2),Cm(30.6),Cm(12.4),cd)
    ch=gf.chart
    ch.has_legend=True; ch.legend.position=XL_LEGEND_POSITION.BOTTOM
    ch.legend.include_in_layout=False
    ch.legend.font.size=Pt(12); ch.legend.font.name=FONT
    ch.font.size=Pt(12); ch.font.name=FONT
    va=ch.value_axis; va.has_major_gridlines=True
    va.tick_labels.font.size=Pt(11); va.tick_labels.font.name=FONT
    if pct: va.tick_labels.number_format='0"%"'; va.tick_labels.number_format_is_linked=False
    else:   va.tick_labels.number_format='#,##0;△#,##0'; va.tick_labels.number_format_is_linked=False
    ca=ch.category_axis
    ca.tick_labels.font.size=Pt(11); ca.tick_labels.font.name=FONT
    ca.major_tick_mark=XL_TICK_MARK.OUTSIDE
    cols=[NAVY,RED,TEAL,AMBER]
    for i,pl in enumerate(ch.series):
        pl.format.line.color.rgb=cols[i%4]; pl.format.line.width=Pt(2.5)
        pl.smooth=False
    txt(s,1.65,16.0,30.5,1.4,note,10,False,RGBColor(0x55,0x55,0x55))
    return s

for tag,fix in [("A_現行基準",False),("B_是正基準",True)]:
    wb=openpyxl.load_workbook(f"{WD}/w7/out/投資財政計画_{tag}.xlsx",data_only=True)
    sim=wb["SIM_R8予算置換_0%"]
    prs=Presentation(); prs.slide_width=Emu(12192000); prs.slide_height=Emu(6858000)
    for sl in list(prs.slides._sldIdLst): prs.slides._sldIdLst.remove(sl)

    Y9=[f"R{y}" for y in range(8,17)]
    nashi=run(0.0,False)
    hoten_n=[rnd(nashi["r72"][i]) for i in range(1,10)]
    hoten_a=[rnd(sim.cell(row=72,column=c).value) for c in range(3,12)]
    chart_slide(prs,"補填財源残高の推移","料金改定なしでは回復せず、改定後も回復はR12以降",
        Y9,[("改定なし",hoten_n),("7.3％改定あり",hoten_a)],1,
        "単位：百万円。R7末183百万円を起点とした簡便推計。改定ありは資産維持費0％ケース（必要改定率7.3％）をR11以降に反映。"
        "マイナスは追加財源・投資平準化・料金改定のいずれかが必要であることを示す。")

    Y10=[f"R{y}" for y in range(7,17)]
    chart_slide(prs,"料金回収率と経常収支比率の推移","現行料金では両指標が100％を下回り続ける",
        Y10,[("経常収支比率",[round(nashi["r31"][i],1) for i in range(10)]),
             ("料金回収率",[round(nashi["r32"][i],1) for i in range(10)])],2,
        "単位：％。現行料金ベース（料金改定なし）。R8は当初予算、R9以降は年次計画置換後。100％を下回る期間は、料金収入で費用を回収できていないことを示す。",pct=True)

    Y11=[f"R{y}" for y in range(11,17)]
    ser=[("改定なし",[rnd(nashi["r28"][i]) for i in range(4,10)])]
    for rt in (0.10,0.20,0.30):
        d=run(rt,fix); ser.append((f"{int(rt*100)}％改定",[rnd(d["r28"][i]) for i in range(4,10)]))
    chart_slide(prs,"改定率別の当期純利益","10％改定ではR16に再び赤字、20％改定で10年の黒字を維持",
        Y11,ser,3,
        f"単位：百万円。R11以降に改定率を適用した試算（{'営業収益の二重計上を是正した基準' if fix else '現行モデル基準'}）。"
        "当期純利益の黒字と、企業債償還・施設更新に必要な資金の確保は別の論点。")

    out=f"{WD}/w7/out/図表編_{tag}.pptx"
    prs.save(out); print(f"保存: 図表編_{tag}.pptx（{len(prs.slides._sldIdLst)}枚）")
