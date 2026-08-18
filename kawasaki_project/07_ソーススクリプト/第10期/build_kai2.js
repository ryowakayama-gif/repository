const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  PageNumber, Header, Footer
} = require('docx');

const NAVY="1F3864", BLUE="2E75B6", LBLUE="DEEBF7", MBLUE="BDD7EE";
const ORANGE="C55A11", LORANGE="FCE4D6", GREEN="548235", LGREEN="E2EFDA";
const RED="C00000", GRAY="808080", LGRAY="F2F2F2", PURPLE="7030A0", LPURPLE="E9DFF2";
const YELLOW="FFF2CC", WHITE="FFFFFF", GOLD="BF9000";
const FH="游ゴシック", FB="游明朝";
const thin=(c=GRAY)=>({style:BorderStyle.SINGLE,size:4,color:c});
const cb=(c=GRAY)=>({top:thin(c),left:thin(c),bottom:thin(c),right:thin(c)});
const M={top:60,bottom:60,left:110,right:110};

function agendaTitle(text){return new Paragraph({spacing:{before:260,after:150},shading:{type:ShadingType.CLEAR,fill:NAVY},
  border:{top:{style:BorderStyle.SINGLE,size:2,color:NAVY,space:6},bottom:{style:BorderStyle.SINGLE,size:2,color:NAVY,space:6}},
  children:[new TextRun({text:` ${text}`,bold:true,color:WHITE,font:FH,size:26})]});}
function gidai(tag,title){return new Paragraph({spacing:{before:280,after:150},shading:{type:ShadingType.CLEAR,fill:ORANGE},
  children:[new TextRun({text:`  ${tag}  `,bold:true,color:WHITE,font:FH,size:24}),new TextRun({text:`  ${title}`,bold:true,color:WHITE,font:FH,size:24})]});}
function sub(text){return new Paragraph({spacing:{before:180,after:80},border:{left:{style:BorderStyle.SINGLE,size:18,color:BLUE,space:8}},
  children:[new TextRun({text:` ${text}`,bold:true,color:NAVY,font:FH,size:22})]});}
function body(text){const runs=Array.isArray(text)?text:[new TextRun({text,font:FB,size:21,color:"262626"})];
  return new Paragraph({spacing:{after:120,line:300},alignment:AlignmentType.JUSTIFIED,children:runs});}
function r(t,o={}){return new TextRun({text:t,font:o.f||FB,size:o.s||21,bold:o.b||false,color:o.c||"262626",italics:o.i||false});}
function point(text,fill=LBLUE,bar=BLUE){return new Paragraph({spacing:{before:80,after:140},shading:{type:ShadingType.CLEAR,fill},
  border:{left:{style:BorderStyle.SINGLE,size:18,color:bar,space:10},top:{style:BorderStyle.SINGLE,size:2,color:fill,space:4},bottom:{style:BorderStyle.SINGLE,size:2,color:fill,space:4}},
  children:[new TextRun({text:`● ${text}`,bold:true,color:NAVY,font:FH,size:21})]});}
function ronten(text){return new Paragraph({spacing:{before:100,after:140},shading:{type:ShadingType.CLEAR,fill:LORANGE},
  border:{left:{style:BorderStyle.SINGLE,size:22,color:ORANGE,space:10},top:{style:BorderStyle.SINGLE,size:2,color:LORANGE,space:4},bottom:{style:BorderStyle.SINGLE,size:2,color:LORANGE,space:4}},
  children:[new TextRun({text:`◆ ご協議いただきたい点：`,bold:true,color:ORANGE,font:FH,size:21}),new TextRun({text:text,bold:true,color:NAVY,font:FH,size:21})]});}
function caveat(text){return new Paragraph({spacing:{before:60,after:120},shading:{type:ShadingType.CLEAR,fill:YELLOW},
  border:{top:{style:BorderStyle.SINGLE,size:4,color:GOLD,space:4},bottom:{style:BorderStyle.SINGLE,size:4,color:GOLD,space:4}},
  children:[new TextRun({text:`⚠ `,bold:true,color:"7F6000",font:FH,size:18}),new TextRun({text:text,color:"7F6000",font:FB,size:18})]});}
function note(text){return new Paragraph({spacing:{after:120},children:[new TextRun({text,italics:true,color:GRAY,font:FB,size:17})]});}
function caption(text){return new Paragraph({spacing:{before:80,after:40},children:[new TextRun({text,bold:true,color:NAVY,font:FH,size:20})]});}
function chart(path,w,h){return new Paragraph({spacing:{before:80,after:60},alignment:AlignmentType.CENTER,
  children:[new ImageRun({type:"png",data:fs.readFileSync(path),transformation:{width:w,height:h}})]});}
function cell(text,{fill,c,b,align,w,font,size}={}){const runs=Array.isArray(text)?text:[new TextRun({text:String(text),bold:b||false,color:c||"262626",font:font||FB,size:size||18})];
  return new TableCell({width:w?{size:w,type:WidthType.DXA}:undefined,borders:cb(),margins:M,verticalAlign:VerticalAlign.CENTER,
    shading:fill?{type:ShadingType.CLEAR,fill}:undefined,children:[new Paragraph({alignment:align||AlignmentType.LEFT,children:runs})]});}
function hr(cells,widths,fill=NAVY){return new TableRow({tableHeader:true,children:cells.map((t,i)=>cell(t,{fill,c:WHITE,b:true,align:AlignmentType.CENTER,w:widths[i],font:FH,size:18}))});}
function table(widths,rows){return new Table({width:{size:widths.reduce((a,b)=>a+b,0),type:WidthType.DXA},columnWidths:widths,rows});}
function numbered(items){return items.map((t,i)=>new Paragraph({spacing:{after:70},indent:{left:200},
  children:[new TextRun({text:`${i+1}. `,bold:true,color:NAVY,font:FH,size:21}),new TextRun({text:t,font:FB,size:21})]}));}

const ch=[];
// 表紙
ch.push(
  new Paragraph({spacing:{before:120,after:40},alignment:AlignmentType.CENTER,children:[new TextRun({text:"川崎町高齢者保健福祉計画・第10期介護保険事業計画",bold:true,color:NAVY,font:FH,size:22})]}),
  new Paragraph({spacing:{after:40},alignment:AlignmentType.CENTER,shading:{type:ShadingType.CLEAR,fill:NAVY},
    border:{top:{style:BorderStyle.SINGLE,size:2,color:NAVY,space:8},bottom:{style:BorderStyle.SINGLE,size:2,color:NAVY,space:8}},
    children:[new TextRun({text:"第2回 介護保険運営委員会　説明資料",bold:true,color:WHITE,font:FH,size:30})]}),
  new Paragraph({spacing:{before:100,after:20},alignment:AlignmentType.CENTER,children:[new TextRun({text:"〜 介護サービス見込量・介護給付費・保険料の見通し 〜",italics:true,color:BLUE,font:FB,size:20})]}),
  new Paragraph({spacing:{before:120},alignment:AlignmentType.CENTER,children:[new TextRun({text:"令和8年11月　川崎町（保健福祉課）",color:GRAY,font:FB,size:18})]}),
);

// 会議次第
ch.push(agendaTitle("会議次第"));
ch.push(...numbered([
  "開会",
  "第1回介護保険運営委員会の確認（アンケート結果・基本理念・第9期評価）",
  "【報告事項】計画素案の骨子（第2章〜第4章）",
  "【協議事項1】介護サービス見込量の推計",
  "【協議事項2】介護給付費の推計",
  "【協議事項3】介護保険料の見通し",
  "今後のスケジュール",
  "閉会",
]));
ch.push(note("※ 本資料は委員の皆様にお配りする説明資料です。数値の一部は町提供データ・国通知（夏以降）・準備基金残高（令和8年6月確定）の反映を経て、Ver.2.0で確定します。"));

// 1. 第1回の確認
ch.push(agendaTitle("1　第1回委員会の確認"));
ch.push(body("第1回委員会（令和8年8月）では、アンケート調査結果、計画の基本理念・基本目標、第9期計画の評価をご報告・ご審議いただきました。第2回では、これらを踏まえた介護サービス見込量・給付費・保険料の見通しをご協議いただきます。"));
ch.push(caption("表　第1回の主な確認事項と第2回への接続"));
ch.push(table([3200,6360],[
  hr(["第1回の事項","第2回での扱い"],[3200,6360]),
  new TableRow({children:[cell("アンケート結果（一般高齢者・認定者）",{fill:LGRAY,b:true}),cell("サービス見込量の利用意向・潜在ニーズの補正に反映（協議事項1）",{})]}),
  new TableRow({children:[cell("基本理念・基本目標",{fill:LGRAY,b:true}),cell("確定した理念・目標に沿って施策体系を整理（報告事項）",{})]}),
  new TableRow({children:[cell("第9期計画の評価（KPI達成状況）",{fill:LGRAY,b:true}),cell("第10期の課題として給付費・保険料の見通しに反映（協議事項2・3）",{})]}),
]));
ch.push(caveat("第1回でいただいたご意見への対応は、別途「委員意見反映管理シート」で整理し、次期素案（Ver.2.0）に反映します。"));

// 2. 報告事項 素案骨子
ch.push(agendaTitle("2　【報告事項】計画素案の骨子"));
ch.push(body("第10期計画（令和9〜11年度）は、認知症基本法への対応を独立章として位置づけるなど、以下の構成で編成します。"));
ch.push(caption("表　計画素案の構成"));
ch.push(table([2100,4400,3060],[
  hr(["章","主な内容","本日の関連"],[2100,4400,3060]),
  new TableRow({children:[cell("第2章 現状分析",{fill:LGRAY,b:true}),cell("人口・高齢化率・認定者・給付費の現状と将来推計",{}),cell("協議1・2の前提",{align:AlignmentType.CENTER,c:BLUE})]}),
  new TableRow({children:[cell("第3章 第9期評価",{fill:LGRAY,b:true}),cell("第9期計画の取組実績とKPI達成状況",{}),cell("報告",{align:AlignmentType.CENTER})]}),
  new TableRow({children:[cell("第4章 施策",{fill:LGRAY,b:true}),cell("5つの基本目標に沿った施策（介護予防・在宅支援・サービス基盤ほか）",{}),cell("報告",{align:AlignmentType.CENTER})]}),
  new TableRow({children:[cell("第5章 認知症施策",{fill:LGRAY,b:true}),cell("認知症基本法に基づく施策（独立章）",{}),cell("報告",{align:AlignmentType.CENTER})]}),
  new TableRow({children:[cell("第6章 サービス見込量・保険料",{fill:LORANGE,b:true,c:ORANGE}),cell("サービス見込量・給付費・保険料の算定",{}),cell("協議1・2・3",{align:AlignmentType.CENTER,c:ORANGE,b:true})]}),
]));
ch.push(point("本日は第6章の内容（サービス見込量・給付費・保険料）を中心にご協議いただきます。",MBLUE,BLUE));

// 協議事項1 サービス見込量
ch.push(gidai("協議事項1","介護サービス見込量の推計"));
ch.push(sub("(1) 見込量の算定ロジック"));
ch.push(body([r("サービス見込量は、",{}),r("将来人口 × 認定率 × 利用率 × 1人当たり利用量",{b:true,c:NAVY}),r("を基本式に、次の6ステップで年度別・サービス種類別に算定します。アンケートの利用意向・潜在ニーズで補正します。",{})]));
ch.push(caption("表　見込量算定の6ステップ"));
ch.push(table([1000,4200,4360],[
  hr(["Step","内容","データ"],[1000,4200,4360]),
  new TableRow({children:[cell("1-2",{fill:LGRAY,b:true,align:AlignmentType.CENTER}),cell("将来人口・認定者数の推計",{}),cell("社人研推計・認定実績",{})]}),
  new TableRow({children:[cell("3-4",{fill:LGRAY,b:true,align:AlignmentType.CENTER}),cell("サービス利用率・1人当たり利用量の算定",{}),cell("国保連データ・給付実績",{})]}),
  new TableRow({children:[cell("5",{fill:LGRAY,b:true,align:AlignmentType.CENTER}),cell("年度別・サービス種類別見込量の算定",{}),cell("上記の積算",{})]}),
  new TableRow({children:[cell("6",{fill:LGRAY,b:true,align:AlignmentType.CENTER}),cell("アンケートによる潜在ニーズ・利用意向の補正",{}),cell("一般高齢者・認定者調査",{})]}),
]));
ch.push(sub("(2) 川崎町のサービス構造の特徴"));
ch.push(caption("表　サービス別給付費の構成（第9期）"));
ch.push(table([3400,2400,3760],[
  hr(["サービス区分","給付費割合","特徴"],[3400,2400,3760]),
  new TableRow({children:[cell("居宅サービス",{fill:LGRAY,b:true}),cell("約38%",{align:AlignmentType.CENTER}),cell("在宅生活支援の中心",{})]}),
  new TableRow({children:[cell("地域密着型サービス",{fill:LGRAY,b:true}),cell("約16%",{align:AlignmentType.CENTER}),cell("グループホーム等",{})]}),
  new TableRow({children:[cell("施設サービス",{fill:LORANGE,b:true,c:ORANGE}),cell("約46%",{align:AlignmentType.CENTER,c:ORANGE,b:true}),cell("受給者は約29%だが給付費の半分",{})]}),
  new TableRow({children:[cell("住宅改修・福祉用具",{fill:LGRAY,b:true}),cell("約0.3%",{align:AlignmentType.CENTER}),cell("—",{})]}),
]));
ch.push(point("施設サービスは受給者数では約29%（134人）ですが、1人当たり単価が高いため給付費の約46%を占めます。施設サービスの見込みが給付費全体を大きく左右します。",MBLUE,BLUE));
ch.push(body([r("川崎町は町内施設が縮小傾向にあり、",{}),r("住所地特例24人（令和7年6月時点）",{b:true,c:NAVY}),r("が示すとおり、町外施設（柴田町・大河原町・仙台市等）の利用が相当数あります。見込量算定では、町内施設の定員のみで制約せず、町外施設利用を含めた「町民の施設サービス利用ニーズ」全体を見込みます。",{})]));
ch.push(ronten("施設サービスの見込量について、町外施設利用（住所地特例）を含めた算定方針でよろしいか。町外施設の受入余地・所在自治体の把握についてご意見をいただきたい。"));
ch.push(caveat("年度別・サービス種類別見込量の確定値は、町提供データ・国保連データ・アンケート結果・国通知の反映を経て、Ver.2.0で記載します。本日は算定方針をご協議いただきます。"));

// 協議事項2 給付費
ch.push(gidai("協議事項2","介護給付費の推計"));
ch.push(sub("(1) 人口・高齢者数の見通し"));
ch.push(body([r("川崎町の65歳以上人口は",{}),r("令和7年（2025年）頃がピーク",{b:true,c:RED}),r("で、第10期計画期間は既に微減局面に入ります。一方、総人口はそれ以上に減少し、高齢化率は上昇を続けます。",{})]));
ch.push(chart("choki/c1_population.png",560,293));
ch.push(point("「高齢者が増え続ける」局面は終わりつつありますが、後期高齢化（75歳以上の割合上昇）と重度化により、認定者数は横ばい〜微増で、介護給付費は当面減りません。",MBLUE,BLUE));
ch.push(sub("(2) 介護給付費の推計"));
ch.push(body([r("介護給付費は、サービス見込量に単価・加算率を乗じて算定します（標準給付費見込額＝総給付額＋特定入所者・高額介護等）。第9期実績を基に、後期高齢化・重度化・介護報酬改定を見込むと、第10期の標準給付費は3年間で",{}),r("第9期比+2.9%程度",{b:true,c:RED}),r("の増加が見込まれます。",{})]));
ch.push(chart("choki/c5_kyufu.png",560,278));
ch.push(caveat("給付費の伸び率は前提を置いた推計です。確定値はサービス見込量（協議事項1）と令和9年度介護報酬改定を反映して算定します。"));

// 協議事項3 保険料
ch.push(gidai("協議事項3","介護保険料の見通し"));
ch.push(sub("(1) 第9期保険料の構造"));
ch.push(body([r("第9期の保険料（基準額月額6,500円）は、",{}),r("介護給付費準備基金を約7,800万円取り崩す",{b:true,c:NAVY}),r("ことで実現した水準です。取崩しをしなければ約7,300〜7,600円相当でした。第10期開始時点では基金が大きく減少（計画ベースで約5,400万円）するため、同じ抑制余地は小さくなります。",{})]));
ch.push(sub("(2) 第10期保険料の見通し（3パターン）"));
ch.push(body("給付費の増加と基金の減少を踏まえ、基金の取崩し方に応じた3パターンを試算しました。"));
ch.push(chart("choki/c6_fund_premium.png",560,256));
ch.push(caption("表　第10期 保険料基準額の見通し（月額・試算）"));
ch.push(table([3000,2400,4160],[
  hr(["パターン","基準額（月額）","考え方"],[3000,2400,4160]),
  new TableRow({children:[cell("（参考）第9期",{fill:LGRAY,b:true}),cell("6,500円",{align:AlignmentType.CENTER,b:true}),cell("基金約7,800万円取崩で実現",{})]}),
  new TableRow({children:[cell("A 取崩なし",{fill:LORANGE,b:true,c:RED}),cell("約7,600円",{align:AlignmentType.CENTER,b:true,c:RED}),cell("基金を温存（次期の負担を緩和）",{})]}),
  new TableRow({children:[cell("B 一部取崩",{fill:LORANGE,b:true,c:ORANGE}),cell("約7,400円",{align:AlignmentType.CENTER,b:true,c:ORANGE}),cell("中位。基金の半分を活用",{})]}),
  new TableRow({children:[cell("C 全額取崩",{fill:LGREEN,b:true,c:GREEN}),cell("約7,100円",{align:AlignmentType.CENTER,b:true,c:GREEN}),cell("基金を活用し抑制（次期の負担増）",{})]}),
]));
ch.push(point("第9期6,500円を維持（据置）しようとすると基金が計画期間中に枯渇するため、財源的に維持は困難です。基金の枯渇が保険料上昇の最大の要因です。",LORANGE,ORANGE));
ch.push(ronten("保険料の水準は、基金の活用度合い（次期の負担とのバランス）と激変緩和の観点が論点となります。具体的な基準額は第3回で詳細を協議し、第4回で決定します。本日は3パターンの考え方についてご意見をいただきたい。"));

// スケジュール
ch.push(agendaTitle("3　今後のスケジュール"));
ch.push(table([2400,2600,4560],[
  hr(["委員会","時期","主な議題"],[2400,2600,4560]),
  new TableRow({children:[cell("第2回（本日）",{fill:MBLUE,b:true}),cell("令和8年11月",{align:AlignmentType.CENTER,b:true}),cell("サービス見込量・給付費・保険料の見通し",{b:true})]}),
  new TableRow({children:[cell("第3回",{fill:LGRAY,b:true}),cell("令和9年1月中旬",{align:AlignmentType.CENTER}),cell("保険料基準額の協議、計画素案（Ver.2.0）の審議",{})]}),
  new TableRow({children:[cell("第4回",{fill:LGRAY,b:true}),cell("令和9年2月",{align:AlignmentType.CENTER}),cell("保険料基準額の決定、計画最終案の確定",{})]}),
  new TableRow({children:[cell("パブリックコメント",{fill:LGRAY,b:true}),cell("令和9年1〜2月",{align:AlignmentType.CENTER}),cell("計画案の公表・意見募集",{})]}),
]));
ch.push(body([r("本日いただいたご意見と、町提供データ・国通知・準備基金残高（令和8年6月確定）を反映し、",{}),r("第3回では確定値に基づく保険料基準額をご協議",{b:true,c:NAVY}),r("いただきます。",{})]));
ch.push(new Paragraph({spacing:{before:200},border:{top:{style:BorderStyle.SINGLE,size:6,color:BLUE,space:6}},
  children:[new TextRun({text:"本日ご協議いただきたい点：①施設サービス見込量の算定方針（町外利用を含む）、②給付費推計の考え方、③保険料3パターンの考え方。いずれも第3回・第4回での確定に向けた方向性の確認です。",italics:true,color:GRAY,font:FB,size:18})]}));

const doc=new Document({
  styles:{default:{document:{run:{font:FB,size:21}}}},
  sections:[{
    properties:{page:{size:{width:11906,height:16838},margin:{top:1080,right:1080,bottom:1080,left:1080}}},
    headers:{default:new Header({children:[new Paragraph({alignment:AlignmentType.RIGHT,
      border:{bottom:{style:BorderStyle.SINGLE,size:4,color:MBLUE,space:4}},
      children:[new TextRun({text:"川崎町 第10期計画　第2回介護保険運営委員会 説明資料",color:GRAY,font:FH,size:15})]})]})},
    footers:{default:new Footer({children:[new Paragraph({alignment:AlignmentType.CENTER,
      children:[new TextRun({children:["- ",PageNumber.CURRENT," -"],color:GRAY,font:FB,size:16})]})]})},
    children:ch,
  }],
});
Packer.toBuffer(doc).then(buf=>{fs.writeFileSync("川崎町_第2回介護保険運営委員会_説明資料.docx",buf);console.log("written",buf.length,"bytes");});
