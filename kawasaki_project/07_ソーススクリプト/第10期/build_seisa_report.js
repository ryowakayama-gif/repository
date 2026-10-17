const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  PageNumber, Header, Footer
} = require('docx');

const NAVY="1F3864", BLUE="2E75B6", LBLUE="DEEBF7", MBLUE="BDD7EE";
const ORANGE="C55A11", LORANGE="FCE4D6", GREEN="548235", LGREEN="E2EFDA";
const RED="C00000", GRAY="808080", LGRAY="F2F2F2", PURPLE="7030A0";
const YELLOW="FFF2CC", WHITE="FFFFFF", GOLD="BF9000";
const FH="游ゴシック", FB="游明朝";
const thin=(c=GRAY)=>({style:BorderStyle.SINGLE,size:4,color:c});
const cb=(c=GRAY)=>({top:thin(c),left:thin(c),bottom:thin(c),right:thin(c)});
const M={top:60,bottom:60,left:110,right:110};

function secTitle(num,title){return new Paragraph({spacing:{before:300,after:150},shading:{type:ShadingType.CLEAR,fill:NAVY},
  border:{top:{style:BorderStyle.SINGLE,size:2,color:NAVY,space:6},bottom:{style:BorderStyle.SINGLE,size:2,color:NAVY,space:6}},
  children:[new TextRun({text:`  ${num}  `,bold:true,color:GOLD,font:FH,size:26}),new TextRun({text:title,bold:true,color:WHITE,font:FH,size:26})]});}
function sub(text){return new Paragraph({spacing:{before:180,after:80},border:{left:{style:BorderStyle.SINGLE,size:18,color:BLUE,space:8}},
  children:[new TextRun({text:` ${text}`,bold:true,color:NAVY,font:FH,size:22})]});}
function body(text){const runs=Array.isArray(text)?text:[new TextRun({text,font:FB,size:21,color:"262626"})];
  return new Paragraph({spacing:{after:120,line:300},alignment:AlignmentType.JUSTIFIED,children:runs});}
function r(t,o={}){return new TextRun({text:t,font:o.f||FB,size:o.s||21,bold:o.b||false,color:o.c||"262626",italics:o.i||false});}
function find(n,text){return new Paragraph({spacing:{before:100,after:140},shading:{type:ShadingType.CLEAR,fill:LORANGE},
  border:{left:{style:BorderStyle.SINGLE,size:22,color:ORANGE,space:10},top:{style:BorderStyle.SINGLE,size:2,color:LORANGE,space:4},bottom:{style:BorderStyle.SINGLE,size:2,color:LORANGE,space:4}},
  children:[new TextRun({text:`◆ 発見${n}：`,bold:true,color:ORANGE,font:FH,size:21}),new TextRun({text:text,bold:true,color:NAVY,font:FH,size:21})]});}
function pass(text){return new Paragraph({spacing:{before:80,after:120},shading:{type:ShadingType.CLEAR,fill:LGREEN},
  border:{left:{style:BorderStyle.SINGLE,size:18,color:GREEN,space:10}},
  children:[new TextRun({text:`✓ 検証OK：`,bold:true,color:GREEN,font:FH,size:21}),new TextRun({text:text,color:"262626",font:FB,size:21})]});}
function callout(text,fill=LBLUE,bar=BLUE){return new Paragraph({spacing:{before:80,after:140},shading:{type:ShadingType.CLEAR,fill},
  border:{left:{style:BorderStyle.SINGLE,size:18,color:bar,space:10},top:{style:BorderStyle.SINGLE,size:2,color:fill,space:4},bottom:{style:BorderStyle.SINGLE,size:2,color:fill,space:4}},
  children:[new TextRun({text:`● ${text}`,bold:true,color:NAVY,font:FH,size:21})]});}
function note(text){return new Paragraph({spacing:{after:120},children:[new TextRun({text,italics:true,color:GRAY,font:FB,size:17})]});}
function caption(text){return new Paragraph({spacing:{before:80,after:40},children:[new TextRun({text,bold:true,color:NAVY,font:FH,size:20})]});}
function cell(text,{fill,c,b,align,w,font,size}={}){const runs=Array.isArray(text)?text:[new TextRun({text:String(text),bold:b||false,color:c||"262626",font:font||FB,size:size||18})];
  return new TableCell({width:w?{size:w,type:WidthType.DXA}:undefined,borders:cb(),margins:M,verticalAlign:VerticalAlign.CENTER,
    shading:fill?{type:ShadingType.CLEAR,fill}:undefined,children:[new Paragraph({alignment:align||AlignmentType.LEFT,children:runs})]});}
function hr(cells,widths,fill=NAVY){return new TableRow({tableHeader:true,children:cells.map((t,i)=>cell(t,{fill,c:WHITE,b:true,align:AlignmentType.CENTER,w:widths[i],font:FH,size:18}))});}
function table(widths,rows){return new Table({width:{size:widths.reduce((a,b)=>a+b,0),type:WidthType.DXA},columnWidths:widths,rows});}

const ch=[];
ch.push(
  new Paragraph({spacing:{before:120,after:60},alignment:AlignmentType.CENTER,children:[new TextRun({text:"川崎町・金ケ崎町　第10期介護保険事業計画　策定支援",bold:true,color:NAVY,font:FH,size:22})]}),
  new Paragraph({spacing:{after:40},alignment:AlignmentType.CENTER,shading:{type:ShadingType.CLEAR,fill:NAVY},
    border:{top:{style:BorderStyle.SINGLE,size:2,color:NAVY,space:8},bottom:{style:BorderStyle.SINGLE,size:2,color:NAVY,space:8}},
    children:[new TextRun({text:"給付費・保険料算定の精査レポート",bold:true,color:WHITE,font:FH,size:30})]}),
  new Paragraph({spacing:{before:100,after:20},alignment:AlignmentType.CENTER,children:[new TextRun({text:"〜 前回計画（第9期）の算定式・国の指針／策定要領との照合 〜",italics:true,color:BLUE,font:FB,size:20})]}),
  new Paragraph({spacing:{before:100},alignment:AlignmentType.CENTER,children:[new TextRun({text:"令和8年6月　ビズアップ公共コンサルティング株式会社（札幌事業所）",color:GRAY,font:FB,size:18})]}),
);

// 要旨
ch.push(secTitle("要旨","精査結果の要点"));
ch.push(pass("保険料算定式の構造（標準給付費・地域支援事業費×第1号負担割合23%、調整交付金、準備基金取崩、予定収納率による按分）は、国の標準的算定方法と一致。川崎町第9期の各構成額も「取崩しをしない場合の収納必要額」まで完全に再現できた。"));
ch.push(find("1","保険料の按分被保険者数に用いる補正係数（所得段階別の加重平均料率）が、川崎町は実効で約0.96〜0.98。ワークブックの暫定値1.0は基準額を約3〜4%過小評価していたため、第9期計画から逆算した0.97に補正した。"));
ch.push(find("2","川崎町第9期計画の保険料算出表で「取崩しをしない場合の収納必要額（804,663千円）−取崩額（78,000千円）」と「取崩後の収納必要額（712,663千円）」に約14百万円の差がある。標準8ステップに現れない過年度精算・端数調整等と推察され、要確認。"));
ch.push(find("3","川崎町第9期の調整交付金見込額は標準5%相当を下回る約3.9%。後期高齢化・低所得の保険者は普通調整交付金が5%超となる例が多く、実際の交付実績の確認を推奨（5%なら基準額は約4%低下し、発見1の補正と相殺）。"));
ch.push(note("※ 発見1（基準額を上げる）と発見3（基準額を下げる）はほぼ相殺し、第10期基準額の中心推計7,000円台は頑健。確定値は町提供データ（所得段階別人口・調整交付金交付実績・基金残高R8.6）で算定する。"));

// 1. 目的・方法
ch.push(secTitle("1","精査の目的と方法"));
ch.push(body("第10期計画の給付費・保険料試算の妥当性を確かめるため、(1)国の標準的算定方法（厚生労働省）、(2)川崎町・金ケ崎町の前回計画（第9期）の保険料算出式、(3)当社作成の保険料試算ワークブックの3者を突き合わせ、構成額・算定式・諸係数を1項目ずつ検証した。"));
ch.push(caption("表　照合に用いた主な根拠"));
ch.push(table([3000,6560],[
  hr(["区分","出典"],[3000,6560]),
  new TableRow({children:[cell("国の標準算定式",{fill:LGRAY,b:true}),cell("厚生労働省「介護保険の保険料（第1号被保険者）財政の仕組み」、第9期保険料算定の自治体公表例（松阪市・長野市・新宿区等）",{})]}),
  new TableRow({children:[cell("調整交付金の算定式",{fill:LGRAY,b:true}),cell("介護保険最新情報Vol.1190（令和5年12月22日・厚生労働省老健局介護保険計画課）後期高齢者加入割合補正係数・所得段階別加入割合補正係数",{})]}),
  new TableRow({children:[cell("両町の前回保険料",{fill:LGRAY,b:true}),cell("厚生労働省「第9期介護保険の第1号保険料」全国一覧（川崎町6,500円／金ケ崎町4,900円）、川崎町第9期計画 第6章 保険料算出",{})]}),
]));

// 2. 国の標準算定式
ch.push(secTitle("2","国の標準的な保険料算定式（確認結果）"));
ch.push(body("国の標準的な算定は、計画期間（3年）の総費用から第1号被保険者の負担額を求め、これを予定収納率と所得段階補正後の被保険者数で按分する2段階構造である。式に表すと次のとおり。"));
ch.push(new Paragraph({spacing:{before:80,after:120},shading:{type:ShadingType.CLEAR,fill:LGRAY},border:cb(BLUE).top&&{top:{style:BorderStyle.SINGLE,size:6,color:BLUE,space:6},bottom:{style:BorderStyle.SINGLE,size:6,color:BLUE,space:6}},
  children:[new TextRun({text:"保険料基準額（月額）＝ ",bold:true,color:NAVY,font:FH,size:20}),
    new TextRun({text:"〔(標準給付費＋地域支援事業費)×23％ ＋ 調整交付金相当額(5%) − 調整交付金見込額 − 保険者機能強化交付金 − 準備基金取崩額〕",color:"262626",font:FB,size:19}),
    new TextRun({text:" ÷ 予定収納率 ÷ 所得段階補正後被保険者数(3年) ÷ 12",bold:true,color:NAVY,font:FH,size:20})]}));
ch.push(body([r("標準給付費は",{}),r("総給付額＋特定入所者介護サービス費＋高額介護サービス費＋高額医療合算＋審査支払手数料",{b:true,c:NAVY}),r("の合計、地域支援事業費を加えたものが保険料算定の基礎費用となる（長野市等の公表例で確認）。第1号負担割合23％は全国一律で、第10期も同率である。",{})]));
ch.push(callout("調整交付金は「標準5%相当額を加算し、実際の交付見込額を減算する」方式が標準。交付率が5%を下回る保険者は第1号負担が増え、上回る保険者は減る。当社ワークブックはこの正式な方式を実装済み（単純減算ではない）。",MBLUE,BLUE));

// 3. 川崎町の検証
ch.push(secTitle("3","川崎町 第9期算定式の検証"));
ch.push(sub("(1) 構成額の逐次検証（再現性）"));
ch.push(caption("表　川崎町 第9期 保険料算出の検算（3年計・千円）"));
ch.push(table([3700,2200,2200,1460],[
  hr(["項目","計画値","当社再計算","判定"],[3700,2200,2200,1460]),
  new TableRow({children:[cell("標準給付費見込額",{}),cell("3,294,553",{align:AlignmentType.RIGHT}),cell("3,294,553",{align:AlignmentType.RIGHT}),cell("一致",{align:AlignmentType.CENTER,c:GREEN,b:true})]}),
  new TableRow({children:[cell("第1号負担相当額(23%)",{}),cell("778,806",{align:AlignmentType.RIGHT}),cell("778,806",{align:AlignmentType.RIGHT,c:GREEN}),cell("一致",{align:AlignmentType.CENTER,c:GREEN,b:true})]}),
  new TableRow({children:[cell("調整交付金相当額(5%)",{}),cell("166,006",{align:AlignmentType.RIGHT}),cell("166,006",{align:AlignmentType.RIGHT}),cell("一致",{align:AlignmentType.CENTER,c:GREEN,b:true})]}),
  new TableRow({children:[cell("調整交付金見込額(3.9%)",{}),cell("129,349",{align:AlignmentType.RIGHT}),cell("129,349",{align:AlignmentType.RIGHT}),cell("一致",{align:AlignmentType.CENTER,c:GREEN,b:true})]}),
  new TableRow({children:[cell("取崩なし収納必要額",{b:true}),cell("804,663",{align:AlignmentType.RIGHT,b:true}),cell("804,663",{align:AlignmentType.RIGHT,b:true,c:GREEN}),cell("一致",{align:AlignmentType.CENTER,c:GREEN,b:true})]}),
  new TableRow({children:[cell("取崩後収納必要額",{}),cell("712,663",{align:AlignmentType.RIGHT}),cell("726,663",{align:AlignmentType.RIGHT,c:RED}),cell("△14,000",{align:AlignmentType.CENTER,c:RED,b:true})]}),
  new TableRow({children:[cell("基準額(月額)",{b:true}),cell("6,508円",{align:AlignmentType.RIGHT,b:true}),cell("約6,650円",{align:AlignmentType.RIGHT,b:true}),cell("±2%",{align:AlignmentType.CENTER,c:ORANGE,b:true})]}),
]));
ch.push(pass("第1号負担相当額（778,806千円）と取崩なし収納必要額（804,663千円）は計画値と完全一致。算定式の構造（23%、調整交付金の加減算、機能強化交付金）は正しい。"));
ch.push(find("2","取崩なし収納必要額804,663千円から取崩額78,000千円を引くと726,663千円だが、計画書は取崩後を712,663千円と記載（差14,000千円）。これは過年度精算・第2号繰越・端数調整等、標準8ステップに現れない項目と推察される。基準額の再現が±2%に留まる主因でもあり、町の算定ワークシートで確認が必要。"));
ch.push(sub("(2) 補正係数（所得段階別の加重平均料率）"));
ch.push(body([r("基準額の按分では、被保険者数を所得段階別の料率で加重した「補正後被保険者数」で割る。川崎町は非課税層が約30%（第1〜3段階・料率0.285〜0.685）と多いため、加重平均料率は1.0を下回る。第9期計画の基準額（6,508円）・取崩後収納必要額・収納率から逆算すると",{}),r("実効補正係数は約0.96〜0.98",{b:true,c:RED}),r("となる。",{})]));
ch.push(find("1","当社ワークブックは補正係数を暫定1.0としていたが、これは基準額を約3〜4%過小評価する。第9期実効値の中点0.97に補正した結果、第10期基準額の8ステップ算定値が当社の長期推計レポート（取崩なし7,641円）とほぼ一致した（下表）。"));
ch.push(sub("(3) 調整交付金見込率の妥当性"));
ch.push(body([r("国の普通調整交付金は、後期高齢者加入割合と所得段階分布で傾斜配分される。後期高齢化が進み低所得層が多い保険者ほど交付率は高くなる（国はR6改定で所得段階補正を標準13段階化し調整機能を強化）。川崎町は高齢化率42%・後期高齢者割合51.6%・非課税層30%と、",{}),r("本来は5%超の交付が見込まれる属性",{b:true,c:NAVY}),r("だが、第9期計画の見込額は約3.9%と標準5%を下回っている。",{})]));
ch.push(find("3","川崎町の調整交付金見込率3.9%は、属性から期待される水準より低い可能性がある。一人当たり給付費が全国平均を下回ること等が要因とも考えられるが、実際の交付実績（介護保険事業状況報告）の確認を推奨する。仮に5%なら第10期基準額は約4%低下し、発見1（補正係数）の引上げ分とほぼ相殺する。"));

// 4. 金ケ崎町
ch.push(secTitle("4","金ケ崎町 第9期算定式の検証"));
ch.push(body([r("金ケ崎町の第9期保険料は、第8期5,100円から",{}),r("4,900円へ3.9%引下げ",{b:true,c:GREEN}),r("（厚生労働省 全国一覧で確認）。人口減少局面で給付費の上方圧力がある中での引下げは、",{}),r("介護給付費準備基金の取崩しによる抑制",{b:true,c:NAVY}),r("で実現したものであり、川崎町（基金78,000千円取崩で6,500円を維持）と同じ構造である。",{})]));
ch.push(callout("両町とも『前回は基金取崩しで保険料を抑えた』点が共通。第10期は基金の取崩余地が縮小するため、給付費が横ばい〜微増でも基準額に上昇圧力がかかる——この構造は両町に共通する。",MBLUE,BLUE));
ch.push(note("※ 金ケ崎町第9期計画の保険料算出表（標準給付費・調整交付金・基金残高の内訳）は当該計画書原本が必要。基準額4,900円・引下げ幅・基金活用という骨格は確認済みだが、構成額の逐次検証には第9期計画書 第6〜7章の保険料算出ページの取得を推奨する。"));
ch.push(body([r("なお当社の金ケ崎町 長期推計分析は、社人研R5推計の確定値（高齢者人口は現在付近がピーク、高齢化率は約32%台で頭打ち、認定率は16.4→18.9%へ上昇）と、4,900円据置時の基金枯渇試算（仮定を明示した参考試算）に基づいており、本精査の算定式の確認と矛盾しない。",{})]));

// 5. 給付費算定の検証
ch.push(secTitle("5","給付費算定の検証"));
ch.push(sub("(1) 標準給付費の定義"));
ch.push(pass("当社ワークブックは「総給付額（居宅・地域密着・施設・住宅改修）＋補完項目（特定入所者・高額介護・高額医療合算・審査手数料）＝標準給付費見込額」の定義を採用しており、川崎町第9期計画の標準給付費（R8単年=1,103,536千円）と一致する。総給付額のみで止めると約8%過小評価となるため、補完項目を加える本定義が正しい。"));
ch.push(sub("(2) 給付費の伸び率（介護報酬改定）"));
ch.push(body([r("第9期は介護報酬改定+1.59%が織り込まれた。第10期（令和9年度〜）は新たな報酬改定が予定されるが改定率は未定である。当社推計は、後期高齢化・重度化・報酬改定を見込み",{}),r("年1.2%（低位0.8%〜高位1.8%）",{b:true,c:NAVY}),r("を置いている。これは第9期実績の伸び（年約0.4%）より高めで、後期高齢化を反映した保守的設定だが、確定値はアンケート反映後の見込量と国の報酬改定率で精緻化する。",{})]));

// 6. ワークブックへの反映
ch.push(secTitle("6","ワークブックへの反映と検証結果"));
ch.push(body("精査の結果、ワークブックの補正係数を1.0から第9期実効値0.97へ補正し、調整交付金見込率（3.9%）に交付実績確認を促す注記を追加した。第10期基準額（8ステップ算定）は次のとおりで、当社の長期推計レポートの試算と整合する。"));
ch.push(caption("表　第10期 保険料基準額（補正係数0.97・月額）"));
ch.push(table([3000,1900,1900,2760],[
  hr(["パターン","補正後(0.97)","補正前(1.0)","備考"],[3000,1900,1900,2760]),
  new TableRow({children:[cell("（参考）第9期",{fill:LGRAY,b:true}),cell("6,500円",{align:AlignmentType.CENTER,b:true}),cell("—",{align:AlignmentType.CENTER}),cell("基金78,000千円取崩で実現",{})]}),
  new TableRow({children:[cell("A 取崩なし",{fill:LORANGE,b:true,c:RED}),cell("7,632円",{align:AlignmentType.CENTER,b:true,c:RED}),cell("7,403円",{align:AlignmentType.CENTER,c:GRAY}),cell("レポート試算7,641円と一致",{})]}),
  new TableRow({children:[cell("B 50%取崩",{fill:LORANGE,b:true,c:ORANGE}),cell("7,384円",{align:AlignmentType.CENTER,b:true,c:ORANGE}),cell("7,163円",{align:AlignmentType.CENTER,c:GRAY}),cell("中位",{})]}),
  new TableRow({children:[cell("C 全額取崩",{fill:LGREEN,b:true,c:GREEN}),cell("7,137円",{align:AlignmentType.CENTER,b:true,c:GREEN}),cell("6,923円",{align:AlignmentType.CENTER,c:GRAY}),cell("基金約54百万円取崩",{})]}),
]));
ch.push(callout("補正係数の補正（基準額＋約4%）と、調整交付金が5%だった場合の低下（約−4%）は概ね相殺する。第10期基準額の中心は取崩条件に応じ7,100〜7,600円台で、6,500円据置は基金枯渇で維持困難という結論は変わらない。",LBLUE,BLUE));

// 7. 残課題
ch.push(secTitle("7","残課題と委員会説明方針"));
ch.push(caption("表　町への確認事項（確定後にワークブックで基準額を確定）"));
ch.push(table([700,4400,4460],[
  hr(["No","確認事項","基準額への影響"],[700,4400,4460]),
  new TableRow({children:[cell("1",{fill:LGRAY,b:true,align:AlignmentType.CENTER}),cell("所得段階別被保険者数（→補正係数の確定）",{}),cell("補正係数が0.97から動くと比例して増減",{})]}),
  new TableRow({children:[cell("2",{fill:LGRAY,b:true,align:AlignmentType.CENTER}),cell("調整交付金の交付実績・見込率（5%超か否か）",{}),cell("5%なら約−4%（補正係数引上げと相殺）",{})]}),
  new TableRow({children:[cell("3",{fill:LGRAY,b:true,align:AlignmentType.CENTER}),cell("準備基金残高（R8.6確定値）",{}),cell("取崩可能額が変動し3パターンが動く",{})]}),
  new TableRow({children:[cell("4",{fill:LGRAY,b:true,align:AlignmentType.CENTER}),cell("予定収納率（過去5年R4〜R7実績）",{}),cell("現在は第9期と同じ96%を暫定使用",{})]}),
  new TableRow({children:[cell("5",{fill:LGRAY,b:true,align:AlignmentType.CENTER}),cell("取崩なし→取崩後の14百万円差の内訳",{}),cell("過年度精算等。基準額の±2%要因",{})]}),
  new TableRow({children:[cell("6",{fill:LGRAY,b:true,align:AlignmentType.CENTER}),cell("金ケ崎町 第9期計画の保険料算出表",{}),cell("構成額の逐次検証に必要（原本取得）",{})]}),
]));
ch.push(new Paragraph({spacing:{before:200},border:{top:{style:BorderStyle.SINGLE,size:6,color:BLUE,space:6}},
  children:[new TextRun({text:"総括：保険料・給付費算定の構造は国の標準的方法および川崎町第9期計画と一致することを確認した。精査で見つかった補正係数（修正済）・調整交付金見込率・14百万円差は、いずれも町提供データで確定すべき項目であり、相互に相殺する要因も含むため、第10期基準額の中心推計（取崩条件に応じ7,100〜7,600円台）は頑健である。金ケ崎町は基準額・基金活用の骨格を確認したが、構成額の逐次検証には前回計画書原本の取得を推奨する。",italics:true,color:GRAY,font:FB,size:18})]}));

const doc=new Document({
  styles:{default:{document:{run:{font:FB,size:21}}}},
  sections:[{
    properties:{page:{size:{width:11906,height:16838},margin:{top:1080,right:1080,bottom:1080,left:1080}}},
    headers:{default:new Header({children:[new Paragraph({alignment:AlignmentType.RIGHT,
      border:{bottom:{style:BorderStyle.SINGLE,size:4,color:MBLUE,space:4}},
      children:[new TextRun({text:"川崎町・金ケ崎町　給付費・保険料算定の精査レポート",color:GRAY,font:FH,size:15})]})]})},
    footers:{default:new Footer({children:[new Paragraph({alignment:AlignmentType.CENTER,
      children:[new TextRun({children:["- ",PageNumber.CURRENT," -"],color:GRAY,font:FB,size:16})]})]})},
    children:ch,
  }],
});
Packer.toBuffer(doc).then(buf=>{fs.writeFileSync("川崎町・金ケ崎町_給付費・保険料算定の精査レポート.docx",buf);console.log("written",buf.length,"bytes");});
