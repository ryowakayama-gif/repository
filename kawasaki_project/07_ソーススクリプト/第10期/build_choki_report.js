const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  PageNumber, PageBreak, Header, Footer
} = require('docx');

const NAVY="1F3864", BLUE="2E75B6", LBLUE="DEEBF7", MBLUE="BDD7EE";
const ORANGE="C55A11", LORANGE="FCE4D6", GREEN="548235", LGREEN="E2EFDA";
const RED="C00000", GRAY="808080", LGRAY="F2F2F2", PURPLE="7030A0", LPURPLE="E9DFF2";
const YELLOW="FFF2CC", WHITE="FFFFFF", GOLD="BF9000";
const FH="游ゴシック", FB="游明朝";
const thin=(c=GRAY)=>({style:BorderStyle.SINGLE,size:4,color:c});
const cb=(c=GRAY)=>({top:thin(c),left:thin(c),bottom:thin(c),right:thin(c)});
const M={top:60,bottom:60,left:110,right:110};

function secTitle(num,title){return new Paragraph({spacing:{before:320,after:160},shading:{type:ShadingType.CLEAR,fill:NAVY},
  border:{top:{style:BorderStyle.SINGLE,size:2,color:NAVY,space:6},bottom:{style:BorderStyle.SINGLE,size:2,color:NAVY,space:6}},
  children:[new TextRun({text:`  ${num}  `,bold:true,color:GOLD,font:FH,size:26}),new TextRun({text:title,bold:true,color:WHITE,font:FH,size:26})]});}
function sub(text){return new Paragraph({spacing:{before:200,after:90},border:{left:{style:BorderStyle.SINGLE,size:18,color:BLUE,space:8}},
  children:[new TextRun({text:` ${text}`,bold:true,color:NAVY,font:FH,size:22})]});}
function body(text,opts={}){const runs=Array.isArray(text)?text:[new TextRun({text,font:FB,size:21,color:"262626"})];
  return new Paragraph({spacing:{after:120,line:300},alignment:AlignmentType.JUSTIFIED,children:runs,...opts});}
function r(text,o={}){return new TextRun({text,font:o.f||FB,size:o.s||21,bold:o.b||false,color:o.c||"262626",italics:o.i||false});}
function find(text,fill=LORANGE,bar=ORANGE){return new Paragraph({spacing:{before:80,after:140},shading:{type:ShadingType.CLEAR,fill},
  border:{left:{style:BorderStyle.SINGLE,size:20,color:bar,space:10},top:{style:BorderStyle.SINGLE,size:2,color:fill,space:4},bottom:{style:BorderStyle.SINGLE,size:2,color:fill,space:4}},
  children:[new TextRun({text:`◆ 発見：`,bold:true,color:bar,font:FH,size:21}),new TextRun({text:text,bold:true,color:NAVY,font:FH,size:21})]});}
function callout(text,fill=LBLUE,bar=BLUE){return new Paragraph({spacing:{before:80,after:140},shading:{type:ShadingType.CLEAR,fill},
  border:{left:{style:BorderStyle.SINGLE,size:18,color:bar,space:10},top:{style:BorderStyle.SINGLE,size:2,color:fill,space:4},bottom:{style:BorderStyle.SINGLE,size:2,color:fill,space:4}},
  children:[new TextRun({text:`● ${text}`,bold:true,color:NAVY,font:FH,size:21})]});}
function caveat(text){return new Paragraph({spacing:{before:60,after:120},shading:{type:ShadingType.CLEAR,fill:YELLOW},
  border:{top:{style:BorderStyle.SINGLE,size:4,color:GOLD,space:4},bottom:{style:BorderStyle.SINGLE,size:4,color:GOLD,space:4}},
  children:[new TextRun({text:`⚠ 試算上の前提：`,bold:true,color:"7F6000",font:FH,size:18}),new TextRun({text:text,color:"7F6000",font:FB,size:18})]});}
function note(text){return new Paragraph({spacing:{after:120},children:[new TextRun({text,italics:true,color:GRAY,font:FB,size:17})]});}
function caption(text){return new Paragraph({spacing:{before:60,after:40},children:[new TextRun({text,bold:true,color:NAVY,font:FH,size:20})]});}
function chart(path,w,h){return new Paragraph({spacing:{before:80,after:80},alignment:AlignmentType.CENTER,
  children:[new ImageRun({type:"png",data:fs.readFileSync(path),transformation:{width:w,height:h}})]});}
function cell(text,{fill,c,b,align,w,font,size}={}){const runs=Array.isArray(text)?text:[new TextRun({text:String(text),bold:b||false,color:c||"262626",font:font||FB,size:size||18})];
  return new TableCell({width:w?{size:w,type:WidthType.DXA}:undefined,borders:cb(),margins:M,verticalAlign:VerticalAlign.CENTER,
    shading:fill?{type:ShadingType.CLEAR,fill}:undefined,children:[new Paragraph({alignment:align||AlignmentType.LEFT,children:runs})]});}
function hr(cells,widths,fill=NAVY){return new TableRow({tableHeader:true,children:cells.map((t,i)=>cell(t,{fill,c:WHITE,b:true,align:AlignmentType.CENTER,w:widths[i],font:FH,size:18}))});}
function table(widths,rows){return new Table({width:{size:widths.reduce((a,b)=>a+b,0),type:WidthType.DXA},columnWidths:widths,rows});}

const ch=[];
// Title
ch.push(
  new Paragraph({spacing:{before:160,after:60},alignment:AlignmentType.CENTER,children:[new TextRun({text:"川崎町高齢者保健福祉計画・第10期介護保険事業計画",bold:true,color:NAVY,font:FH,size:24})]}),
  new Paragraph({spacing:{after:40},alignment:AlignmentType.CENTER,shading:{type:ShadingType.CLEAR,fill:NAVY},
    border:{top:{style:BorderStyle.SINGLE,size:2,color:NAVY,space:8},bottom:{style:BorderStyle.SINGLE,size:2,color:NAVY,space:8}},
    children:[new TextRun({text:"長期推計・給付費・保険料試算レポート",bold:true,color:WHITE,font:FH,size:30})]}),
  new Paragraph({spacing:{before:120,after:20},alignment:AlignmentType.CENTER,children:[new TextRun({text:"〜 基金残高の確認・人口及び高齢化率の推移・人口ピーク・給付費算定・保険料試算 〜",italics:true,color:BLUE,font:FB,size:20})]}),
  new Paragraph({spacing:{before:120,after:0},alignment:AlignmentType.CENTER,children:[new TextRun({text:"令和8年6月　ビズアップ公共コンサルティング株式会社（札幌事業所）",color:GRAY,font:FB,size:18})]}),
);

// 要旨
ch.push(secTitle("要旨","本レポートの要点"));
ch.push(callout("高齢者人口は令和7年（2025年）頃にピークを迎え、以降は減少に転じる。一方で高齢化率は2050年に55.1%まで上昇を続け、後期高齢化・重度化により認定者数は横ばい〜微増となる。給付費は当面増加圧力が続く。",LBLUE,BLUE));
ch.push(callout("第9期保険料6,500円は、準備基金131,700千円から78,000千円を取り崩して実現した水準である。第10期は基金がほぼ枯渇するため、据置では基金が期間中に枯渇し、基準額は7,000円台への上昇圧力を受ける。",LORANGE,ORANGE));
ch.push(note("※ 人口・高齢化率は確定推計（社人研令和5年推計／第9期計画）を使用。給付費・保険料・基金枯渇時期は一定の前提を置いた参考試算であり、確定値は町提供データ（基金残高R8.6確定値・予定収納率・所得段階別人口）を用いてワークブックで算定する。"));

// 1. 目的・出典
ch.push(secTitle("1","分析の目的とデータ出典"));
ch.push(body("第10期計画（令和9〜11年度）の給付費・保険料を見通すため、人口・高齢化の長期構造と財政（準備基金）の持続可能性を定量的に分析する。論点を人口側（ピーク・構成変化）と財源側（基金・保険料）に分解し、それぞれを確定データと参考試算で示す。"));
ch.push(caption("表　主なデータ出典"));
ch.push(table([2600,6960],[
  hr(["区分","出典"],[2600,6960]),
  new TableRow({children:[cell("人口・高齢化率（長期）",{fill:LGRAY,b:true}),cell("国立社会保障・人口問題研究所「日本の地域別将来推計人口（令和5年推計）」（2000〜2050年・国勢調査ベース）",{})]}),
  new TableRow({children:[cell("人口・高齢化率（近年）",{fill:LGRAY,b:true}),cell("川崎町 第9期計画 第2章（住民基本台帳・コーホート変化率法、R2〜R9）",{})]}),
  new TableRow({children:[cell("認定者・給付費・基金",{fill:LGRAY,b:true}),cell("川崎町 第9期計画 第2章・第6章（認定状況・標準給付費見込額・準備基金残高・保険料算出）",{})]}),
  new TableRow({children:[cell("保険料（実績）",{fill:LGRAY,b:true}),cell("第8期6,380円／第9期6,500円（13段階・基準額78,000円/年）",{})]}),
]));

// 2. 人口の長期推移・ピーク
ch.push(secTitle("2","人口の長期推移と人口ピーク"));
ch.push(body([r("川崎町の総人口は、2000年の10,872人から2050年の4,525人へと",{}),r("約58%減少",{b:true,c:RED}),r("する見込みで、ピークは2000年以前にあり、既に長期の減少局面にある。65歳以上人口は2015年3,083人→2020年3,210人→",{}),r("2025年（令和7年）3,219人でピーク",{b:true,c:RED}),r("を迎え、以降は2030年3,149人、2050年2,494人へと減少する。第9期計画の住民基本台帳ベース推計でも、高齢者人口はR7（2025年）の3,316人をピークに減少へ転じており、両推計でピーク時期が一致する。",{})]));
ch.push(chart("choki/c1_population.png",600,314));
ch.push(find("川崎町の高齢者人口（65歳以上）は令和7年（2025年）頃が概ねピークで、第10期計画期間（R9〜11）は既に微減局面に入る。一方、15〜64歳の生産年齢人口は2020年4,381人→2050年1,795人へと約59%減少し、保険料・介護を支える担い手が急速に細る。"));
ch.push(callout("「高齢者の数が増え続ける」局面は終わりつつある。第10期以降の論点は『高齢者の増加』ではなく『高齢者人口は減るのに介護需要は減らない』という構造への対応に移る。",MBLUE,BLUE));

// 3. 高齢化率
ch.push(secTitle("3","高齢化率の長期推移（全国比較）"));
ch.push(body([r("高齢化率は2020年の38.6%から上昇を続け、2030年44.8%、2040年49.1%、",{}),r("2050年には55.1%",{b:true,c:RED}),r("に達する（全国平均37.1%を18ポイント上回る）。約10人に6人が高齢者となる水準である。",{})]));
ch.push(chart("choki/c2_koreika.png",600,300));
ch.push(find("川崎町の高齢化率は頭打ちにならず、2050年まで上昇を続ける。これは高齢者人口の減少よりも総人口（特に生産年齢人口）の減少が速いためで、分母が縮むことで比率が上がり続ける構造である。"));
ch.push(callout("高齢者人口のピークアウト（発見2）と高齢化率の上昇継続（発見3）は矛盾しない。『高齢者の絶対数は減るが、社会全体に占める割合は高まり続ける』——これが川崎町の人口構造の核心である。",MBLUE,BLUE));

// 4. 後期高齢化
ch.push(secTitle("4","後期高齢化の進行"));
ch.push(body([r("高齢者の中の後期高齢者（75歳以上）割合は、2025年の約51.5%から上昇し、団塊世代が75歳以上へ移行することで",{}),r("2035年頃に約60%でピーク",{b:true,c:RED}),r("となる見込みである。要介護認定率は加齢とともに大きく高まる（第9期計画では90歳以上75.3%・85〜89歳47.1%）ため、後期高齢化の進行は給付費の上方圧力となる。",{})]));
ch.push(chart("choki/c3_kouki.png",600,300));
ch.push(find("高齢者人口が減少に転じても、その内訳は後期高齢者へとシフトする。後期高齢者割合は2035年頃まで上昇し、一人当たりの介護必要度（重度化）が高まるため、給付費は高齢者数の減少ほどには減らない。"));
ch.push(note("※ 75歳以上の将来内訳は、2024年住基実績（75歳以上割合50.5%）と団塊世代の年齢移行を踏まえた推計値。確定値は社人研の年齢階級別推計で精査する。"));

// 5. 認定者
ch.push(secTitle("5","認定者数・認定率の推計"));
ch.push(body([r("要支援・要介護認定者数は、平成30年569人から令和5年578人とほぼ横ばいで推移している。第10期計画期間では、高齢者人口が微減する一方、後期高齢化により認定率が令和5年の17.6%から",{}),r("令和11年（2029年）には約18.9%",{b:true,c:RED}),r("へ上昇するため、認定者数は約580〜600人と横ばい〜微増で推移すると見込まれる。",{})]));
ch.push(chart("choki/c4_nintei.png",600,300));
ch.push(find("『高齢者数は減るが認定率は上がる』ため、認定者数（＝給付費の主要因）は減らない。給付費の見通しは、高齢者人口の減少ではなく、認定率上昇・重度化・介護報酬改定に左右される。"));

// 6. 給付費
ch.push(secTitle("6","介護給付費の推計（第10期）"));
ch.push(body([r("第9期計画の標準給付費見込額は、令和6年度の約10.9億円から令和8年度の約11.0億円へ緩やかに増加する。第10期（R9〜11）は、認定率上昇・重度化に加え介護報酬改定の影響を見込み、年1.2%程度（低位0.8%〜高位1.8%）の伸びを置くと、3年間累計で",{}),r("第9期比+2.9%程度",{b:true,c:RED}),r("の増加と試算される。",{})]));
ch.push(chart("choki/c5_kyufu.png",600,300));
ch.push(caveat("第10期の給付費の伸び率は、第9期実績の伸び（年約0.4%）を基準に、後期高齢化・重度化・令和9年度介護報酬改定の影響を見込んで年1.2%を中位とした仮定値。確定値は町提供データ（認定者数の実績推移・サービス利用率）を用いてワークブックで算定する。"));
ch.push(caption("表　標準給付費見込額の推計（百万円）"));
ch.push(table([2400,1790,1790,1790,1790],[
  hr(["区分","R9(2027)","R10(2028)","R11(2029)","3年計"],[2400,1790,1790,1790,1790]),
  new TableRow({children:[cell("中位（年+1.2%）",{fill:LGRAY,b:true}),cell("1,117",{align:AlignmentType.CENTER}),cell("1,130",{align:AlignmentType.CENTER}),cell("1,144",{align:AlignmentType.CENTER}),cell("3,391",{align:AlignmentType.CENTER,b:true})]}),
  new TableRow({children:[cell("低位（年+0.8%）",{fill:LGRAY}),cell("1,112",{align:AlignmentType.CENTER}),cell("1,121",{align:AlignmentType.CENTER}),cell("1,130",{align:AlignmentType.CENTER}),cell("3,364",{align:AlignmentType.CENTER})]}),
  new TableRow({children:[cell("高位（年+1.8%）",{fill:LGRAY}),cell("1,123",{align:AlignmentType.CENTER}),cell("1,144",{align:AlignmentType.CENTER}),cell("1,164",{align:AlignmentType.CENTER}),cell("3,431",{align:AlignmentType.CENTER})]}),
]));

// 7. 基金・保険料
ch.push(secTitle("7","介護給付費準備基金の確認と保険料試算"));
ch.push(sub("(1) 準備基金残高の確認"));
ch.push(body([r("第9期計画の保険料算出によれば、準備基金残高は",{}),r("131,700千円（約1億3,170万円）",{b:true,c:NAVY}),r("で、うち",{}),r("78,000千円（7,800万円）を取り崩す",{b:true,c:ORANGE}),r("ことで、取崩しをしない場合の約7,339円相当を基準額6,500円に抑制していた。計画どおり取り崩した場合、",{}),r("第10期開始時（令和8年度末）の基金残高は約53,700千円に減少",{b:true,c:RED}),r("する見込みである。",{})]));
ch.push(caveat("第10期開始時の基金残高は、第9期の取崩しが計画どおり進んだ場合の理論値（131,700−78,000＝53,700千円）。実際の給付実績により増減し、確定値（令和8年6月時点）は町に確認が必要。第9期に給付が計画を下回れば、残高はこれより大きくなる可能性がある。"));

ch.push(sub("(2) 第10期保険料の試算（3パターン）"));
ch.push(body([r("給付費の伸び（3年計+2.9%）と第1号被保険者の微減を反映し、開始基金53,700千円（計画ベース）を前提に、基金取崩額に応じた3パターンを試算した。",{})]));
ch.push(chart("choki/c6_fund_premium.png",600,274));
ch.push(caption("表　第10期保険料基準額の試算（月額）"));
ch.push(table([3200,2400,3960],[
  hr(["パターン","基準額（月額）","内容"],[3200,2400,3960]),
  new TableRow({children:[cell("（参考）第9期実績",{fill:LGRAY,b:true}),cell("6,500円",{align:AlignmentType.CENTER,b:true}),cell("基金78,000千円取崩で実現",{})]}),
  new TableRow({children:[cell("A 取崩なし",{fill:LORANGE,b:true,c:RED}),cell("約7,641円",{align:AlignmentType.CENTER,b:true,c:RED}),cell("基金を温存（第11期負担を緩和）",{})]}),
  new TableRow({children:[cell("B 50%取崩",{fill:LORANGE,b:true,c:ORANGE}),cell("約7,352円",{align:AlignmentType.CENTER,b:true,c:ORANGE}),cell("中位。基金の半分（約27百万円）を活用",{})]}),
  new TableRow({children:[cell("C 全額取崩",{fill:LGREEN,b:true,c:GREEN}),cell("約7,063円",{align:AlignmentType.CENTER,b:true,c:GREEN}),cell("基金全額（約54百万円）取崩で抑制（次期負担増）",{})]}),
]));
ch.push(find("第9期6,500円は『大きな基金（131,700千円）を取り崩して実現した水準』だった。第10期は基金がほぼ枯渇するため、同じ抑制余地がなく、給付費増も加わって基準額は7,000円台への上昇圧力を受ける。基金の枯渇が保険料上昇の最大の要因である。"));

ch.push(sub("(3) 基金の持続可能性（6,500円据置の場合）"));
ch.push(body([r("仮に保険料を6,500円に据え置くと、毎年度22〜26百万円程度の基金取崩しが必要となり、開始基金を30〜80百万円のいずれと想定しても、",{}),r("第10期計画期間中（令和9〜11年度）に基金が枯渇",{b:true,c:RED}),r("する。前頁グラフ左が示すとおり、6,500円の据置は財源的に維持できず、第10期のいずれかの時点での引上げが構造的に避けられない。",{})]));
ch.push(caveat("基金枯渇時期は、給付費の伸び（年1.2%）・予定収納率96%・第1号被保険者の減少を置いた参考試算。実際の枯渇時期は基金確定値・給付実績により変動する。"));

// 8. 総括
ch.push(secTitle("8","総括（MECE／Red Team）と委員会説明方針"));
ch.push(sub("論点の整理（MECE）"));
ch.push(body([r("本分析は論点を",{}),r("人口側（ピーク・構成変化）",{b:true,c:NAVY}),r("と",{}),r("財源側（基金・保険料）",{b:true,c:NAVY}),r("に排他的に分解した。人口側では高齢者人口は2025年頃ピークアウトするが後期高齢化・重度化で認定者・給付費は減らないこと、財源側では第9期を支えた基金が枯渇し保険料に上昇圧力がかかることを、それぞれ定量化した。",{})]));
ch.push(sub("結論の頑健性（Red Team）"));
ch.push(body([r("前提を振っても結論は変わらない。給付費の伸びを0.8〜1.8%のいずれと置いても第10期給付費は第9期を上回り、開始基金を30〜80百万円のいずれと置いても6,500円据置では期間中に基金が枯渇する。すなわち",{}),r("現行6,500円の据置は財源的に持続不可能",{b:true,c:RED}),r("であり、第10期での保険料引上げは構造的に不可避である。",{})]));
ch.push(sub("委員会説明方針（断定を避ける）"));
ch.push(table([700,8860],[
  hr(["No","説明の進め方"],[700,8860]),
  new TableRow({children:[cell("1",{fill:LGRAY,b:true,align:AlignmentType.CENTER}),cell([r("具体的な保険料額は断定せず、3パターン（A取崩なし／B50%／C全額）で提示する。基準額は基金確定値・給付実績・所得段階別人口の確定後にワークブックで算定する。",{})])]}),
  new TableRow({children:[cell("2",{fill:LGRAY,b:true,align:AlignmentType.CENTER}),cell([r("第9期6,500円が『基金取崩しで実現した抑制水準』であった事実を共有し、第10期は基金余地が小さいことを丁寧に説明する。",{})])]}),
  new TableRow({children:[cell("3",{fill:LGRAY,b:true,align:AlignmentType.CENTER}),cell([r("人口減少下でも給付費が減らない構造（高齢者数減・認定率上昇・重度化）を、長期推計グラフで視覚的に示す。",{})])]}),
  new TableRow({children:[cell("4",{fill:LGRAY,b:true,align:AlignmentType.CENTER}),cell([r("激変緩和の観点から、基金枯渇を待って急激に引き上げるのではなく、第10期で段階的・計画的に調整する選択肢を論点として提示する。",{})])]}),
]));
ch.push(new Paragraph({spacing:{before:240},border:{top:{style:BorderStyle.SINGLE,size:6,color:BLUE,space:6}},
  children:[new TextRun({text:"以上。人口・高齢化率は確定推計（社人研R5推計・第9期計画）に基づく。給付費・保険料・基金枯渇は前提を明示した参考試算であり、基金残高（R8.6確定値）・予定収納率・所得段階別人口の町提供データを受領後、保険料試算ワークブックで確定値を算定する。",italics:true,color:GRAY,font:FB,size:18})]}));

// assemble
const doc=new Document({
  styles:{default:{document:{run:{font:FB,size:21}}}},
  sections:[{
    properties:{page:{size:{width:11906,height:16838},margin:{top:1080,right:1080,bottom:1080,left:1080}}},
    headers:{default:new Header({children:[new Paragraph({alignment:AlignmentType.RIGHT,
      border:{bottom:{style:BorderStyle.SINGLE,size:4,color:MBLUE,space:4}},
      children:[new TextRun({text:"川崎町 第10期計画　長期推計・給付費・保険料試算レポート",color:GRAY,font:FH,size:15})]})]})},
    footers:{default:new Footer({children:[new Paragraph({alignment:AlignmentType.CENTER,
      children:[new TextRun({children:["- ",PageNumber.CURRENT," -"],color:GRAY,font:FB,size:16})]})]})},
    children:ch,
  }],
});
Packer.toBuffer(doc).then(buf=>{fs.writeFileSync("川崎町_長期推計・給付費・保険料試算レポート.docx",buf);console.log("written",buf.length,"bytes");});
