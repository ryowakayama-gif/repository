const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  PageNumber, PageBreak, Header, Footer
} = require('docx');

const NAVY="1F3864", BLUE="2E75B6", LBLUE="DEEBF7", MBLUE="BDD7EE";
const ORANGE="C55A11", LORANGE="FCE4D6", GREEN="548235", LGREEN="E2EFDA";
const RED="C00000", GRAY="808080", LGRAY="F2F2F2", PURPLE="7030A0", LPURPLE="E9DFF2";
const YELLOW="FFF2CC", WHITE="FFFFFF", GOLD="FFC000", TEAL="2F8F83";
const FH="游ゴシック", FB="游明朝";

const thin=(c=GRAY)=>({style:BorderStyle.SINGLE,size:4,color:c});
const cb=(c=GRAY)=>({top:thin(c),bottom:thin(c),left:thin(c),right:thin(c)});
const M={top:60,bottom:60,left:110,right:110};

function sectionTitle(num,title){return new Paragraph({spacing:{before:320,after:160},shading:{type:ShadingType.CLEAR,fill:NAVY},
  border:{top:{style:BorderStyle.SINGLE,size:2,color:NAVY,space:6},bottom:{style:BorderStyle.SINGLE,size:2,color:NAVY,space:6}},
  children:[new TextRun({text:`  ${num}  `,bold:true,color:GOLD,font:FH,size:26}),new TextRun({text:title,bold:true,color:WHITE,font:FH,size:26})]});}
function shiryo(label,title,status,statusType){
  // statusType: "town"(町実績準拠)/"data"(町データ待ち)/"opt"(新規提案・任意)/"done"(本書収録)
  const m={town:[GREEN,LGREEN],data:[ORANGE,LORANGE],opt:[PURPLE,LPURPLE],done:[BLUE,LBLUE]}[statusType]||[BLUE,LBLUE];
  return new Paragraph({spacing:{before:300,after:140},shading:{type:ShadingType.CLEAR,fill:LBLUE},
    border:{left:{style:BorderStyle.SINGLE,size:20,color:BLUE,space:8},top:{style:BorderStyle.SINGLE,size:2,color:LBLUE,space:4},bottom:{style:BorderStyle.SINGLE,size:2,color:LBLUE,space:4}},
    children:[new TextRun({text:`${label}　`,bold:true,color:ORANGE,font:FH,size:24}),new TextRun({text:title,bold:true,color:NAVY,font:FH,size:24}),
      new TextRun({text:`　［${status}］`,bold:true,color:m[0],font:FH,size:18})]});}
function sub(text){return new Paragraph({spacing:{before:200,after:90},border:{left:{style:BorderStyle.SINGLE,size:18,color:BLUE,space:8}},
  children:[new TextRun({text:` ${text}`,bold:true,color:NAVY,font:FH,size:22})]});}
function body(text,opts={}){const runs=Array.isArray(text)?text:[new TextRun({text,font:FB,size:21,color:"262626"})];
  return new Paragraph({spacing:{after:120,line:300},alignment:AlignmentType.JUSTIFIED,children:runs,...opts});}
function r(text,o={}){return new TextRun({text,font:o.f||FB,size:o.s||21,bold:o.b||false,color:o.c||"262626",italics:o.i||false});}
function callout(text,fill=LBLUE,bar=BLUE){return new Paragraph({spacing:{before:80,after:140},shading:{type:ShadingType.CLEAR,fill},
  border:{left:{style:BorderStyle.SINGLE,size:18,color:bar,space:10},top:{style:BorderStyle.SINGLE,size:2,color:fill,space:4},bottom:{style:BorderStyle.SINGLE,size:2,color:fill,space:4}},
  children:[new TextRun({text:`● ${text}`,bold:true,color:NAVY,font:FH,size:21})]});}
function pending(text){return new Paragraph({spacing:{before:60,after:120},shading:{type:ShadingType.CLEAR,fill:LORANGE},
  border:{top:{style:BorderStyle.SINGLE,size:4,color:ORANGE,space:4},bottom:{style:BorderStyle.SINGLE,size:4,color:ORANGE,space:4}},
  children:[new TextRun({text:`⚠ ${text}`,bold:true,color:ORANGE,font:FH,size:19})]});}
function optnote(text){return new Paragraph({spacing:{before:60,after:120},shading:{type:ShadingType.CLEAR,fill:LPURPLE},
  border:{top:{style:BorderStyle.SINGLE,size:4,color:PURPLE,space:4},bottom:{style:BorderStyle.SINGLE,size:4,color:PURPLE,space:4}},
  children:[new TextRun({text:`◆ ${text}`,bold:true,color:PURPLE,font:FH,size:19})]});}
function note(text){return new Paragraph({spacing:{after:120},children:[new TextRun({text,italics:true,color:GRAY,font:FB,size:17})]});}
// 条文表示: 条名(太字)＋本文。号は字下げ
function jo(label,text){return new Paragraph({spacing:{before:120,after:40},children:[new TextRun({text:label,bold:true,color:NAVY,font:FH,size:20})]});}
function joText(text,indent=0){return new Paragraph({spacing:{after:30,line:280},indent:indent?{left:indent}:undefined,
  children:[new TextRun({text,font:FB,size:19,color:"262626"})]});}
function cell(text,{fill,c,b,align,w,font,size}={}){const runs=Array.isArray(text)?text:[new TextRun({text:String(text),bold:b||false,color:c||"262626",font:font||FB,size:size||19})];
  return new TableCell({width:w?{size:w,type:WidthType.DXA}:undefined,borders:cb(),margins:M,verticalAlign:VerticalAlign.CENTER,
    shading:fill?{type:ShadingType.CLEAR,fill}:undefined,children:[new Paragraph({alignment:align||AlignmentType.LEFT,children:runs})]});}
function hr(cells,widths,fill=NAVY){return new TableRow({tableHeader:true,children:cells.map((t,i)=>cell(t,{fill,c:WHITE,b:true,align:AlignmentType.CENTER,w:widths[i],font:FH,size:19}))});}
function table(widths,rows){return new Table({width:{size:widths.reduce((a,b)=>a+b,0),type:WidthType.DXA},columnWidths:widths,rows});}
function caption(text){return new Paragraph({spacing:{before:60,after:40},children:[new TextRun({text,bold:true,color:NAVY,font:FH,size:20})]});}

const children=[];

// ===== Title =====
children.push(
  new Paragraph({spacing:{before:200,after:60},alignment:AlignmentType.CENTER,children:[new TextRun({text:"川崎町高齢者保健福祉計画・第10期介護保険事業計画",bold:true,color:NAVY,font:FH,size:24})]}),
  new Paragraph({spacing:{after:40},alignment:AlignmentType.CENTER,shading:{type:ShadingType.CLEAR,fill:NAVY},
    border:{top:{style:BorderStyle.SINGLE,size:2,color:NAVY,space:8},bottom:{style:BorderStyle.SINGLE,size:2,color:NAVY,space:8}},
    children:[new TextRun({text:"資料編　構成案・素案（第9期準拠版）",bold:true,color:WHITE,font:FH,size:28})]}),
  new Paragraph({spacing:{before:120,after:20},alignment:AlignmentType.CENTER,children:[new TextRun({text:"〜 前回（第9期）計画の資料編構成に準拠し、町の実態に合わせて再構成 〜",italics:true,color:BLUE,font:FB,size:20})]}),
  new Paragraph({spacing:{before:160,after:0},alignment:AlignmentType.CENTER,children:[new TextRun({text:"令和8年6月　ビズアップ公共コンサルティング株式会社（札幌事業所）",color:GRAY,font:FB,size:18})]}),
);

// ===== Part 0: 整備方針（改訂） =====
children.push(sectionTitle("方針","資料編の整備方針"));
children.push(body([
  r("本書は、", {}),
  r("前回（第9期）川崎町計画の資料編（令和6年3月策定・全4項目）を確認", {b:true,c:NAVY}),
  r("したうえで、第10期計画の資料編構成案と素案を示すものです。前回資料編は、専用の「策定委員会」を設けず", {}),
  r("介護保険条例で設置された常設の「川崎町介護保険運営委員会」", {b:true,c:RED}),
  r("により計画を策定する形をとっており、本書はこの町の実態（house style）に準拠して構成します。", {}),
]));
children.push(callout("前回資料編は①介護保険条例（抜粋）②運営委員会規則③委員名簿④策定経過の4項目構成。本書はこれに準拠し、加えて住民利便のための用語解説・関連法令を任意の追加資料として提案する。", LBLUE, BLUE));

children.push(callout("【要・本文修正】前回計画は審議組織を「介護保険運営委員会」と表記（本文中17回）。一方、第10期計画書本文v1.6は「策定委員会」と表記（20回）。整合のため本文・委員会関係成果物の名称を「介護保険運営委員会」へ統一する必要がある。", LORANGE, ORANGE));

children.push(sub("資料編の構成案（第10期計画）"));
children.push(body("前回（第9期）の4項目を基本構成（町実績準拠）とし、用語解説・関連法令の2項目を任意の追加提案として整理します。各項目の区分を色分けで示します。"));
children.push(caption("表　資料編 構成案と区分"));
children.push(table([760, 3400, 3200, 2400], [
  hr(["資料","名称","区分・根拠","作成状況"],[760,3400,3200,2400]),
  new TableRow({children:[
    cell("資料1",{fill:LGRAY,b:true,align:AlignmentType.CENTER}),
    cell("川崎町介護保険条例（抜粋）",{b:true}),
    cell("町実績準拠／第9期資料編①",{c:GREEN,b:true}),
    cell("素案収録（要・現行条例照合）",{align:AlignmentType.CENTER,c:GREEN,b:true,fill:LGREEN}),
  ]}),
  new TableRow({children:[
    cell("資料2",{fill:LGRAY,b:true,align:AlignmentType.CENTER}),
    cell("川崎町介護保険運営委員会規則",{b:true}),
    cell("町実績準拠／第9期資料編②",{c:GREEN,b:true}),
    cell("素案収録（要・現行規則照合）",{align:AlignmentType.CENTER,c:GREEN,b:true,fill:LGREEN}),
  ]}),
  new TableRow({children:[
    cell("資料3",{fill:LGRAY,b:true,align:AlignmentType.CENTER}),
    cell("川崎町介護保険運営委員会 委員名簿",{b:true}),
    cell("町実績準拠／第9期資料編③",{c:GREEN,b:true}),
    cell("様式確定・委嘱後記入",{align:AlignmentType.CENTER,c:ORANGE,b:true,fill:LORANGE}),
  ]}),
  new TableRow({children:[
    cell("資料4",{fill:LGRAY,b:true,align:AlignmentType.CENTER}),
    cell("策定経過の概要",{b:true}),
    cell("町実績準拠／第9期資料編④",{c:GREEN,b:true}),
    cell("様式確定・策定完了後確定",{align:AlignmentType.CENTER,c:ORANGE,b:true,fill:LORANGE}),
  ]}),
  new TableRow({children:[
    cell("資料5",{fill:LGRAY,b:true,align:AlignmentType.CENTER}),
    cell("用語の解説",{b:true,c:PURPLE}),
    cell("新規提案（任意）／第9期には無し",{c:PURPLE,b:true}),
    cell("素案収録（採否は町判断）",{align:AlignmentType.CENTER,c:PURPLE,b:true,fill:LPURPLE}),
  ]}),
  new TableRow({children:[
    cell("資料6",{fill:LGRAY,b:true,align:AlignmentType.CENTER}),
    cell("関連法令・告示（抜粋）",{b:true,c:PURPLE}),
    cell("新規提案（任意）／第9期には無し",{c:PURPLE,b:true}),
    cell("素案収録（採否は町判断）",{align:AlignmentType.CENTER,c:PURPLE,b:true,fill:LPURPLE}),
  ]}),
]));
children.push(note("※ 前回（第9期）資料編では、アンケート調査・パブリックコメント・諮問答申は独立資料とせず、資料4「策定経過の概要」の中で言及する形をとっている。本書もこれに準拠し、これらを独立資料とせず策定経過に統合する。調査の詳細結果は別冊（調査結果報告書）として整理する。"));

// ===== 資料1 介護保険条例（抜粋） =====
children.push(new Paragraph({children:[new PageBreak()]}));
children.push(sectionTitle("基本構成","町実績準拠の資料（第9期資料編に対応）"));
children.push(shiryo("資料1","川崎町介護保険条例（抜粋）","町実績準拠","town"));
children.push(body("審議組織の設置根拠となる介護保険条例の関係条文（第5章 介護保険運営委員会）を掲載します。以下は前回（第9期）資料編収録の条文であり、第10期計画掲載時は現行条例との一致を確認します。"));
children.push(new Paragraph({spacing:{before:120,after:80},alignment:AlignmentType.CENTER,
  children:[new TextRun({text:"川崎町介護保険条例（平成12年川崎町条例第1号）抜粋",bold:true,color:NAVY,font:FH,size:21})]}));
children.push(new Paragraph({spacing:{after:80},alignment:AlignmentType.CENTER,
  children:[new TextRun({text:"第５章　介護保険運営委員会",bold:true,color:NAVY,font:FH,size:20})]}));
children.push(jo("（介護保険運営委員会の設置）"));
children.push(joText("第17条　介護保険に関する施策の実施を、町民の意見を十分に反映しながら円滑かつ適切に行うため、川崎町介護保険運営委員会（以下「委員会」という。）を置くことができるものとする。"));
children.push(jo("（所掌事務）"));
children.push(joText("第18条　委員会は、次に掲げる事項について調査審議する。"));
children.push(joText("(1) 法第117条第１項の規定による介護保険事業計画の策定又は変更に関する事項",300));
children.push(joText("(2) 介護保険に関する施策及び事務事業の評価に関する事項",300));
children.push(joText("(3) その他介護保険の運営に関し必要と認められる事項",300));
children.push(jo("（組織）"));
children.push(joText("第19条　委員会は、委員８人以内をもって組織する。"));
children.push(joText("２　委員は、次の各号に掲げる者のうちから、それぞれ当該各号に定める数の範囲内において、町長が任命する。"));
children.push(joText("(1) 被保険者を代表する者　３人",300));
children.push(joText("(2) 介護に関し学識又は経験を有する者　２人",300));
children.push(joText("(3) 介護サービスに関する事業に従事する者　３人",300));
children.push(joText("３　委員の任期は、２年とする（４月１日から翌々年の３月末日までとする。）。ただし、委員が欠けた場合における補欠の委員の任期は、前任者の残任期間とする。"));
children.push(joText("４　委員は、再任されることができる。"));
children.push(jo("（規則等への委任）"));
children.push(joText("第20条　前３条に定めるもののほか、介護保険運営委員会の運営に関し必要な事項は、町長が別に定める。"));
children.push(pending("第10期計画掲載時は、改正の有無を含め現行の川崎町介護保険条例と条文を照合のうえ確定する。"));

// ===== 資料2 運営委員会規則 =====
children.push(shiryo("資料2","川崎町介護保険運営委員会規則","町実績準拠","town"));
children.push(body("委員会の運営に関する規則を掲載します（条例第20条に基づき町長が定めるもの）。"));
children.push(jo("（趣旨）"));
children.push(joText("第１条　川崎町介護保険運営委員会（以下「委員会」という。）の事務については、川崎町介護保険条例（平成12年川崎町条例第１号）に定めるもののほか、この規則の定めるところによる。"));
children.push(jo("（組織）"));
children.push(joText("第２条　委員会に委員長及び副委員長を各１人置き、それぞれ委員の互選により選出する。"));
children.push(joText("２　委員長は会務を総括し、委員会を代表する。"));
children.push(joText("３　副委員長は、委員長を補佐し、委員長に事故があるとき、又は委員長が欠けたときは、その職務を代理する。"));
children.push(jo("（会議）"));
children.push(joText("第３条　委員会の会議は町長が招集し、委員長が会議の議長となる。"));
children.push(joText("２　委員会は、委員の半数以上の出席がなければ会議を開くことができない。"));
children.push(joText("３　委員会の議事は、出席委員の過半数で決し、可否同数のときは、議長の決するところによる。"));
children.push(jo("（事務局）"));
children.push(joText("第４条　委員会の事務局は、保健福祉課に置く。"));
children.push(pending("第10期計画掲載時は現行規則と条文を照合のうえ確定する。"));

// ===== 資料3 委員名簿 =====
children.push(shiryo("資料3","川崎町介護保険運営委員会 委員名簿","委嘱後記入","data"));
children.push(body([r("条例第19条に基づき、委員８人以内（被保険者代表３人・学識経験２人・介護サービス事業従事者３人）で組織します。第10期委員の氏名・役職は委嘱確定後に記入します。",{}),
  r("（参考：第9期は６名で構成）",{c:GRAY,i:true})]));
children.push(caption("表　委員名簿（第10期・様式）"));
children.push(table([700, 700, 2900, 3760, 1700], [
  hr(["No","資格\n要件","委員名","役職名","備考"],[700,700,2900,3760,1700]),
  ...[1,2,3,4,5,6,7,8].map(i=>{
    // 区分定数: 1-3=被保険者代表(1), 4-5=学識経験(2), 6-8=介護サービス事業従事者(3)
    const yoken = i<=3?"１":i<=5?"２":"３";
    const kubun = i<=3?"被保険者代表":i<=5?"学識経験":"介護サービス事業従事者";
    return new TableRow({children:[
      cell(String(i),{align:AlignmentType.CENTER,fill:LGRAY}),
      cell(yoken,{align:AlignmentType.CENTER,fill:LBLUE,b:true}),
      cell("（委嘱後記入）",{c:GRAY}),
      cell([new TextRun({text:"（委嘱後記入）　",color:GRAY,font:FB,size:18}),new TextRun({text:`〔区分：${kubun}〕`,color:BLUE,font:FB,size:16})]),
      cell("",{align:AlignmentType.CENTER}),
    ]});
  })
]));
children.push(note("※ 資格要件欄は介護保険条例第19条第２項の区分（(1)＝１、(2)＝２、(3)＝３）を表す。委員数は８人以内のため、定数の範囲で委嘱（第9期は６名）。委員長・副委員長は委員の互選（規則第２条）。"));
children.push(pending("委員名簿（氏名・役職・委員長/副委員長）は第10期委員の委嘱確定後に町担当課より入手して反映する。"));

// ===== 資料4 策定経過 =====
children.push(new Paragraph({children:[new PageBreak()]}));
children.push(shiryo("資料4","策定経過の概要","策定完了後確定","data"));
children.push(body("前回（第9期）資料編の様式に準拠し、調査の実施・各回委員会・諮問答申・パブリックコメント・議会手続を時系列で掲載します。以下は第10期の予定であり、確定日・書面/対面の別は実施後に確定します。"));
children.push(caption("表　策定経過（第10期・予定）"));
children.push(table([2600, 6960], [
  hr(["開催日（予定）","内容等"],[2600,6960]),
  new TableRow({children:[
    cell("令和8年6月",{fill:LBLUE,b:true,align:AlignmentType.CENTER}),
    cell([r("日常生活圏域ニーズ調査及び在宅介護実態調査の実施",{b:true}),r("\n対象：①川崎町在住の高齢者（65歳以上の要介護認定を受けていない高齢者、要支援認定者）：1,000人　②川崎町在住の要介護認定（要介護1〜5）を受け、在宅で暮らしている方：300人",{s:18})]),
  ]}),
  new TableRow({children:[
    cell("令和8年8月中旬",{fill:LBLUE,b:true,align:AlignmentType.CENTER}),
    cell([r("第1回川崎町介護保険運営委員会",{b:true}),r("\n・川崎町介護保険事業の現状と評価について　・第10期計画の策定について（諮問）　・ニーズ調査等の結果報告",{s:18})]),
  ]}),
  new TableRow({children:[
    cell("令和8年11月",{fill:LBLUE,b:true,align:AlignmentType.CENTER}),
    cell([r("第2回川崎町介護保険運営委員会",{b:true}),r("\n・調査結果を踏まえた骨子案　・サービス見込量　・計画案（Ver.2.0）について",{s:18})]),
  ]}),
  new TableRow({children:[
    cell("令和9年1月中旬",{fill:LBLUE,b:true,align:AlignmentType.CENTER}),
    cell([r("第3回川崎町介護保険運営委員会",{b:true}),r("\n・計画案について　・第1号保険料の試算（3パターン）について",{s:18})]),
  ]}),
  new TableRow({children:[
    cell("令和9年1月下旬\n〜2月",{fill:LBLUE,b:true,align:AlignmentType.CENTER}),
    cell([r("パブリックコメントの実施",{b:true}),r("\n（川崎町ホームページ・保健福祉課窓口）",{s:18})]),
  ]}),
  new TableRow({children:[
    cell("令和9年2月",{fill:LBLUE,b:true,align:AlignmentType.CENTER}),
    cell([r("第4回川崎町介護保険運営委員会",{b:true}),r("\n・パブリックコメントの結果について　・答申書（案）について　・保険料基準額の決定",{s:18})]),
  ]}),
  new TableRow({children:[
    cell("令和9年2月",{fill:LBLUE,b:true,align:AlignmentType.CENTER}),
    cell([r("議会への事前説明・答申",{b:true}),r("\n・総務民生常任委員会で事前説明　・介護保険運営委員会より答申書の提出　・議会全員協議会説明",{s:18})]),
  ]}),
  new TableRow({children:[
    cell("令和9年3月",{fill:LBLUE,b:true,align:AlignmentType.CENTER}),
    cell([r("川崎町議会定例会",{b:true}),r("\n・第10期計画及び介護保険条例改正案について　・計画の公表",{s:18})]),
  ]}),
]));
children.push(pending("各回の正確な開催日・出席状況・書面/対面の別・主な意見は、開催後に確定して追記する。回数・項目構成は前回（第9期）実績に準拠した予定であり、実際の運営に合わせて調整する。"));

// ===== 資料5 用語の解説（任意提案） =====
children.push(new Paragraph({children:[new PageBreak()]}));
children.push(sectionTitle("追加提案","住民利便のための任意追加資料"));
children.push(optnote("資料5・資料6は前回（第9期）資料編には無い項目です。住民・委員が計画書を読む際の参照利便を高めるため、追加を提案するものです。採否は町・委員会でご判断ください（前回どおりの簡素な構成とする場合は、本2資料を除いて資料1〜4のみで完結します）。"));
children.push(shiryo("資料5","用語の解説","新規提案（任意）","opt"));
children.push(body("本編で用いる主な専門用語を五十音順に解説します。実際の本編で使用した用語に合わせて確定時に追加・調整します。"));

const gloss = [
  ["あ","",""],
  ["ICT（情報通信技術）","あ","パソコン・スマートフォン・インターネット等を活用した情報通信の技術。介護分野では見守り機器・記録の電子化・オンライン相談等に活用される。"],
  ["8050問題","あ","80代の高齢の親が、50代の無職・ひきこもり等の子を支える状態で生じる生活・介護上の課題。親の介護と子の生活困窮が同時に起こりやすい。"],
  ["か","",""],
  ["介護給付費準備基金","か","計画期間中に介護給付費が見込みを下回った場合の剰余金を積み立てた基金。次期計画で取り崩すことで保険料の上昇を抑制できる。"],
  ["介護予防","か","高齢者が要介護状態になることを防ぎ、また状態の悪化を防ぐための取組。運動・栄養・社会参加等を通じて心身機能の維持を図る。"],
  ["介護予防・日常生活支援総合事業（総合事業）","か","要支援者等を対象に、市町村が地域の実情に応じて行う介護予防・生活支援サービスの総称。訪問型・通所型サービス、一般介護予防事業等で構成される。"],
  ["居宅サービス","か","自宅で生活する要介護者等が利用する訪問介護・通所介護・短期入所等のサービス。"],
  ["ケアプラン（介護サービス計画）","か","要介護者等の状態や希望に応じて、利用する介護サービスの種類・内容を定めた計画。介護支援専門員（ケアマネジャー）が作成する。"],
  ["後期高齢者","か","75歳以上の高齢者。要介護認定率が前期高齢者より大幅に高く、医療・介護需要の中心となる。"],
  ["高齢化率","か","総人口に占める65歳以上人口の割合。本町は令和7年3月時点で41.4%（県内5位）。"],
  ["さ","",""],
  ["在宅介護実態調査","さ","在宅で介護を受ける高齢者・家族の状況（介護負担・サービス利用・就労との両立等）を把握する調査。サービス見込量や施策検討の基礎となる。"],
  ["サービス見込量","さ","計画期間中に必要となる各介護サービスの利用量の見込み。人口推計・認定率・利用率等から算定する。"],
  ["施設サービス","さ","介護老人福祉施設（特養）・介護老人保健施設（老健）・介護医療院に入所して受けるサービス。1人当たり給付費が高い。"],
  ["住所地特例","さ","住所地以外の市町村の施設に入所した場合に、元の住所地の市町村が引き続き保険者となる仕組み。本町は対象者24人（R7.6）。"],
  ["所得段階区分","さ","第1号被保険者の保険料を所得に応じて段階分けする区分。国は第10期に向け9段階から13段階への見直しを推奨。"],
  ["調整交付金","さ","市町村間の財政力（後期高齢者割合・所得分布）の差を調整するため国が交付する交付金。"],
  ["た","",""],
  ["第1号被保険者","た","65歳以上の介護保険の被保険者。保険料は所得段階に応じて決まる。本町は3,244人（R7.6）。"],
  ["第2号被保険者","た","40歳以上65歳未満の医療保険加入者である被保険者。特定疾病により要介護等となった場合にサービスを利用できる。"],
  ["地域包括ケアシステム","た","高齢者が住み慣れた地域で自分らしく暮らし続けられるよう、医療・介護・予防・住まい・生活支援を一体的に提供する体制。"],
  ["地域包括支援センター","た","高齢者の総合相談・権利擁護・介護予防ケアマネジメント等を担う中核機関。"],
  ["地域支援事業","た","市町村が行う、要介護状態となることの予防や地域の支援体制づくりのための事業。総合事業・包括的支援事業・任意事業で構成。"],
  ["地域密着型サービス","た","住み慣れた地域での生活を支えるため、原則として当該市町村の住民が利用できるサービス。小規模多機能型居宅介護・認知症対応型共同生活介護（グループホーム）等。"],
  ["チームオレンジ","た","認知症の人やその家族の支援ニーズに、認知症サポーター等が支援チームを組んで対応する取組。"],
  ["な","",""],
  ["認知症基本法","な","共生社会の実現を推進するための認知症基本法（令和5年法律第65号、令和6年1月施行）。市町村に認知症施策推進計画の策定の努力義務を課す。"],
  ["認知症サポーター","な","認知症に関する正しい知識を持ち、認知症の人や家族を温かく見守り支援する応援者。"],
  ["認知症施策推進計画","な","認知症基本法に基づき市町村が策定する計画。本計画では独立章で位置づける。"],
  ["は","",""],
  ["フレイル","は","加齢に伴い心身の活力（筋力・認知機能等）が低下し、要介護に至る前段階の状態。適切な介入で改善が可能。"],
  ["包括的支援事業","は","地域包括支援センターの運営、在宅医療・介護連携、認知症総合支援、生活支援体制整備等を行う地域支援事業の中核。"],
  ["保険料収納率","は","賦課した保険料のうち実際に納付された割合。"],
  ["本人ミーティング","は","認知症の本人が集い、本人の体験や思いを語り合い、本人の視点を施策に反映する場。"],
  ["ま","",""],
  ["見える化システム","ま","介護保険事業（支援）計画策定支援のため厚労省が提供する地域分析・比較のためのシステム。保険料・給付費等の他団体比較に活用。"],
  ["や","",""],
  ["ヤングケアラー","や","本来大人が担うと想定される家事・家族の世話・介護等を日常的に行っている子ども。家族介護者支援の観点で配慮が必要。"],
  ["要介護認定","や","介護サービスの必要度を判定する手続き。要支援1〜2、要介護1〜5の7区分に分けられる。"],
  ["ら","",""],
  ["老老介護","ら","高齢者が高齢者を介護する状態。介護者自身の心身の負担が大きく、共倒れのリスクがある。"],
  ["LIFE（科学的介護情報システム）","ら","介護サービスの状態・ケア内容等のデータを収集・分析し、科学的に効果が裏付けられた介護を推進する国のシステム。"],
];
let grows = [ hr(["用語","解説"],[3400,6160]) ];
gloss.forEach(g=>{
  const [term, grp, def] = g;
  if(def===""){
    grows.push(new TableRow({children:[
      new TableCell({columnSpan:2,borders:cb(),margins:M,shading:{type:ShadingType.CLEAR,fill:MBLUE},
        children:[new Paragraph({children:[new TextRun({text:`【${term}行】`,bold:true,color:NAVY,font:FH,size:20})]})]}),
    ]}));
    return;
  }
  grows.push(new TableRow({children:[
    cell(term,{b:true,c:NAVY,fill:LGRAY,font:FH,size:19}),
    cell(def,{size:19}),
  ]}));
});
children.push(table([3400,6160], grows));
children.push(note("※ 五十音順。本編で新たに使用した用語があれば確定後に追加する。"));

// ===== 資料6 関連法令（任意提案） =====
children.push(new Paragraph({children:[new PageBreak()]}));
children.push(shiryo("資料6","関連法令・告示（抜粋）","新規提案（任意）","opt"));
children.push(body("本計画の根拠及び関連する主な法令・告示・指針は以下のとおりです。"));
children.push(table([3000, 5200, 1360], [
  hr(["区分","法令・指針名","略称"],[3000,5200,1360]),
  new TableRow({children:[cell("根拠法",{fill:LGREEN,b:true}),cell("介護保険法（平成9年法律第123号）第117条",{}),cell("",{align:AlignmentType.CENTER})]}),
  new TableRow({children:[cell("根拠法",{fill:LGREEN,b:true}),cell("老人福祉法（昭和38年法律第133号）第20条の8",{}),cell("",{align:AlignmentType.CENTER})]}),
  new TableRow({children:[cell("根拠法",{fill:LGREEN,b:true}),cell("共生社会の実現を推進するための認知症基本法（令和5年法律第65号）第14条",{}),cell("認知症基本法",{align:AlignmentType.CENTER})]}),
  new TableRow({children:[cell("国の指針",{fill:LBLUE,b:true}),cell("介護保険事業に係る保険給付の円滑な実施を確保するための基本指針（厚労省告示）",{}),cell("基本指針",{align:AlignmentType.CENTER})]}),
  new TableRow({children:[cell("国の計画",{fill:LBLUE,b:true}),cell("認知症施策推進基本計画（令和6年12月閣議決定）",{}),cell("",{align:AlignmentType.CENTER})]}),
  new TableRow({children:[cell("県の計画",{fill:LORANGE,b:true}),cell("宮城県高齢者保健福祉計画・介護保険事業支援計画（第10期）",{}),cell("県支援計画",{align:AlignmentType.CENTER})]}),
  new TableRow({children:[cell("町条例・規則",{fill:LGRAY,b:true}),cell("川崎町介護保険条例（平成12年川崎町条例第1号）／川崎町介護保険運営委員会規則",{}),cell("",{align:AlignmentType.CENTER})]}),
]));
children.push(note("※ 上記のうち町条例・規則の関係条文は資料1・資料2に抜粋を収録。"));

// ===== closing =====
children.push(new Paragraph({spacing:{before:300},border:{top:{style:BorderStyle.SINGLE,size:6,color:BLUE,space:6}},
  children:[new TextRun({text:"以上。資料1〜4は前回（第9期）資料編に準拠した基本構成であり、条例・規則は現行版との照合、委員名簿・策定経過は委嘱・策定の進捗に応じて確定します。資料5〜6は住民利便のための任意追加提案であり、前回どおり簡素な構成とする場合は資料1〜4のみで完結します。なお、計画書本文v1.6は審議組織を「策定委員会」と表記しているため、町の実態（介護保険運営委員会）に合わせ本文・委員会関係成果物の名称統一が別途必要です。",italics:true,color:GRAY,font:FB,size:18})]}));

// ===== assemble =====
const doc=new Document({
  styles:{default:{document:{run:{font:FB,size:21}}}},
  sections:[{
    properties:{page:{size:{width:11906,height:16838},margin:{top:1080,right:1080,bottom:1080,left:1080}}},
    headers:{default:new Header({children:[new Paragraph({alignment:AlignmentType.RIGHT,
      border:{bottom:{style:BorderStyle.SINGLE,size:4,color:MBLUE,space:4}},
      children:[new TextRun({text:"川崎町 第10期計画　資料編 構成案・素案（第9期準拠版）",color:GRAY,font:FH,size:15})]})]})},
    footers:{default:new Footer({children:[new Paragraph({alignment:AlignmentType.CENTER,
      children:[new TextRun({children:["- ",PageNumber.CURRENT," -"],color:GRAY,font:FB,size:16})]})]})},
    children,
  }],
});
Packer.toBuffer(doc).then(buf=>{fs.writeFileSync("川崎町_資料編_構成案・素案.docx",buf);console.log("written",buf.length,"bytes");});
