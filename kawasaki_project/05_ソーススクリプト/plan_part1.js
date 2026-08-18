/**
 * 川崎町第10期計画書素案v1.0 - メインビルダー
 */
const fs = require('fs');
const H = require('./plan_helpers');
const {
  C, CH, FONT,
  text, p, chapterTitle, section, subsection,
  bullet, numItem, placeholder, fact, source,
  tcell, thcell, kvTable, spacer, image, caption,
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, convertInchesToTwip,
} = H;

const S = [];

// ===========================================================
// 表紙
// ===========================================================
S.push(new Paragraph({
  spacing: { before: 4800, after: 240 },
  alignment: AlignmentType.CENTER,
  children: [text("川崎町高齢者保健福祉計画", { size: 40, bold: true, color: C.navy })],
}));
S.push(new Paragraph({
  spacing: { before: 0, after: 480 },
  alignment: AlignmentType.CENTER,
  children: [text("第10期介護保険事業計画", { size: 40, bold: true, color: C.navy })],
}));
S.push(new Paragraph({
  spacing: { before: 240, after: 240 },
  alignment: AlignmentType.CENTER,
  children: [text("〜 認知症になっても住み慣れた地域で安心して暮らせるまちづくり 〜", {
    size: 24, italics: true, color: C.blue
  })],
}));
S.push(new Paragraph({
  spacing: { before: 720, after: 240 },
  alignment: AlignmentType.CENTER,
  children: [text("計画書素案 Ver.1.0", { size: 28, bold: true, color: C.orange })],
}));
S.push(new Paragraph({
  spacing: { before: 80, after: 80 },
  alignment: AlignmentType.CENTER,
  children: [text("（アンケート結果反映前・初版）", { size: 18, color: C.gray, italics: true })],
}));
S.push(new Paragraph({
  spacing: { before: 2400, after: 80 },
  alignment: AlignmentType.CENTER,
  children: [text("計画期間：令和9年度〜令和11年度", { size: 22, color: C.gray })],
}));
S.push(new Paragraph({
  spacing: { before: 80, after: 80 },
  alignment: AlignmentType.CENTER,
  children: [text("令和8年6月", { size: 22, color: C.gray })],
}));
S.push(new Paragraph({
  spacing: { before: 80, after: 80 },
  alignment: AlignmentType.CENTER,
  children: [text("宮城県柴田郡 川崎町", { size: 22, color: C.gray })],
}));
S.push(new Paragraph({
  spacing: { before: 80, after: 80 },
  alignment: AlignmentType.CENTER,
  children: [text("（策定支援：ビズアップ公共コンサルティング株式会社）", { size: 16, color: C.gray, italics: true })],
}));

// ===========================================================
// 町長挨拶
// ===========================================================
S.push(new Paragraph({
  spacing: { before: 0, after: 0, line: 240 },
  pageBreakBefore: true,
  children: [text("", { size: 20 })],
}));
S.push(new Paragraph({
  spacing: { before: 480, after: 360, line: 320 },
  alignment: AlignmentType.CENTER,
  border: {
    top: { style: BorderStyle.DOUBLE, size: 18, color: C.navy },
    bottom: { style: BorderStyle.DOUBLE, size: 18, color: C.navy },
  },
  children: [text("ご　挨　拶", { size: 36, bold: true, color: C.navy })],
}));

// 挨拶本文（です・ます調・5段落構成）
const greetingBody = [
  "我が国では、高齢化が急速に進展する中、誰もが住み慣れた地域で安心して暮らし続けることができる社会の実現が、極めて重要な課題となっております。本町におきましても、令和7年3月31日時点で高齢化率が41.4%に達し、宮城県内35市町村中5位の高水準となるなど、高齢化の進行は加速度的に進んでおり、後期高齢者（75歳以上）の急増、独居高齢者・高齢者世帯の増加、介護人材の確保、認知症の方とご家族への支援など、多くの課題に直面しております。",
  "こうした中、本町ではこれまで第9期介護保険事業計画（令和6年度〜令和8年度）に基づき、地域包括支援センターを中心とした相談支援体制の整備、ユニバーサルサポーター制度による住民主体の地域づくり、認知症サポーター養成、福祉移送サービスやデマンドバス等による移動支援など、地域の特性を踏まえた取組を積み重ねてまいりました。",
  "このたび、令和9年度から令和11年度までの3年間を計画期間とする「川崎町高齢者保健福祉計画・第10期介護保険事業計画」を策定するにあたりましては、これまでの取組の成果を継承しつつ、令和6年1月に施行されました「共生社会の実現を推進するための認知症基本法」への対応として、認知症施策推進計画を独立章として新たに位置付けるとともに、町外医療機関との広域連携、福祉施設の地域偏在への対応、家族介護者支援、介護人材の確保といった本町固有の課題にも、正面から取り組む計画としております。",
  "本計画の基本理念である「住民が住み慣れた地域で安心して暮らせるまちづくり」のもと、認知症になっても、誰もが自分らしく暮らせる地域共生社会の実現を目指し、町民の皆様、医療・介護・福祉関係者、地域の支え合いを担うサポーターの皆様、そして関係機関と力を合わせ、本計画を着実に推進してまいる所存でございます。",
  "結びに、本計画の策定にあたり、貴重なご意見をお寄せいただきました策定委員会委員の皆様、ニーズ調査にご協力いただきました町民の皆様、ならびに関係機関の皆様に、心より感謝を申し上げます。",
];

greetingBody.forEach(g => {
  S.push(new Paragraph({
    spacing: { before: 80, after: 200, line: 380 },
    alignment: AlignmentType.LEFT,
    indent: { firstLine: 220 },
    children: [text(g, { size: 22 })],
  }));
});

// 署名（日付・役職・町長名）
S.push(new Paragraph({
  spacing: { before: 720, after: 80 },
  alignment: AlignmentType.RIGHT,
  children: [text("令和9年3月", { size: 22 })],
}));
S.push(new Paragraph({
  spacing: { before: 240, after: 80 },
  alignment: AlignmentType.RIGHT,
  children: [text("川崎町長　　　　　　　　　", { size: 24, bold: true })],
}));

// ===========================================================
// 目次
// ===========================================================
S.push(new Paragraph({
  spacing: { before: 0, after: 0, line: 240 },
  pageBreakBefore: true,
  children: [text("", { size: 20 })],
}));
S.push(new Paragraph({
  spacing: { before: 240, after: 480, line: 320 },
  alignment: AlignmentType.CENTER,
  border: { bottom: { style: BorderStyle.SINGLE, size: 24, color: C.navy } },
  children: [text("目　次", { size: 36, bold: true, color: C.navy })],
}));

const toc = [
  ["", "ご挨拶（川崎町長）", "2"],
  ["第1章", "計画の策定にあたって", "5"],
  ["", "　1-1．計画策定の背景と目的", "5"],
  ["", "　1-2．計画の位置付け", "5"],
  ["", "　1-3．計画の期間", "6"],
  ["", "　1-4．計画策定の体制", "6"],
  ["第2章", "川崎町の高齢者を取り巻く現状", "8"],
  ["", "　2-1．人口・世帯の現状", "8"],
  ["", "　2-2．高齢者の状況", "9"],
  ["", "　2-3．要介護認定とサービス利用の状況", "11"],
  ["", "　2-4．介護給付費の状況", "13"],
  ["", "　2-5．高齢者を取り巻く地域資源", "14"],
  ["第3章", "第9期計画の取組実績と評価", "16"],
  ["", "　3-1．第9期計画の体系", "16"],
  ["", "　3-2．主要施策の取組実績（6目標別・詳細）", "16"],
  ["", "　3-3．評価指標の達成状況", "20"],
  ["", "　3-4．第10期に向けた課題", "21"],
  ["第4章", "計画の基本理念と基本目標", "22"],
  ["", "　4-1．基本理念(第9期踏襲)", "22"],
  ["", "　4-2．基本目標", "22"],
  ["", "　4-3．計画の重点ポイント", "23"],
  ["", "　4-4．計画の体系図", "23"],
  ["第5章", "施策の展開", "25"],
  ["", "　5-1．健康づくり・介護予防の推進", "25"],
  ["", "　5-2．高齢者が安心して暮らせるまちづくり", "26"],
  ["", "　5-3．在宅生活継続の支援", "27"],
  ["", "　5-4．介護サービスの質の確保と人材確保", "28"],
  ["", "　5-5．地域包括ケアシステムの深化", "28"],
  ["第6章", "認知症施策推進計画(独立章)", "30"],
  ["", "　6-1．認知症基本法と川崎町の対応方針", "30"],
  ["", "　6-2．認知症施策の体系(7基本的施策)", "30"],
  ["", "　6-3．重点施策とKPI", "31"],
  ["第7章", "介護保険サービス見込量と保険料", "33"],
  ["", "　7-1．サービス見込量の推計", "33"],
  ["", "　7-2．介護給付費の見込み", "35"],
  ["", "　7-3．介護保険料の試算(8ステップ・3パターン)", "36"],
  ["第8章", "計画の推進体制と評価", "42"],
  ["", "　8-1．推進体制", "42"],
  ["", "　8-2．進行管理(PDCAサイクル)", "42"],
  ["", "　8-3．計画の評価と見直し", "43"],
];

const tocTable = new Table({
  width: { size: 96, type: WidthType.PERCENTAGE },
  rows: toc.map(([ch, t, pg]) => new TableRow({
    children: [
      tcell(ch, { width: 10, bold: !!ch, color: ch ? C.navy : C.black, size: 20, align: AlignmentType.CENTER }),
      tcell(t, { width: 80, bold: !!ch, color: ch ? C.navy : C.black, size: 20 }),
      tcell(pg, { width: 10, bold: !!ch, color: ch ? C.navy : C.black, size: 20, align: AlignmentType.RIGHT }),
    ],
  })),
});
S.push(tocTable);

S.push(spacer());
S.push(p("※本素案v1.0はアンケート調査結果反映前の状態です。調査結果（一般高齢者1,000名・要支援要介護認定者300名対象、令和8年6月下旬発送・7月末回収予定）の反映を経て、Ver.2.0として更新する予定です。",
  { size: 18, italics: true, color: C.gray, noIndent: true }));

// ===========================================================
// 第1章 計画の策定にあたって
// ===========================================================
S.push(...chapterTitle(1));

// 1-1
S.push(section(1, 1, "計画策定の背景と目的"));

S.push(subsection("介護保険制度の経緯と本計画の位置", 1));
S.push(p("介護保険制度は平成12年4月に施行されて以降、3年を1期とする介護保険事業計画の策定を市町村に義務付けています。第1期から第9期までの25年間にわたり、川崎町でもこの計画策定を継続してきました。本計画は第10期に当たり、平成12年の制度施行から数えて10サイクル目の計画策定となります。"));

S.push(subsection("我が国の高齢化と2040年問題", 1));
S.push(p("我が国の高齢化は急速に進展しており、令和22年（2040年）には総人口に占める65歳以上人口の割合が約35%に達すると見込まれています。また、団塊の世代が75歳以上となる令和7年（2025年）を経て、令和22年（2040年）には団塊ジュニア世代（昭和46〜49年生まれ）も65歳以上に到達し、医療・介護需要のさらなる増加が見込まれています。これに対し生産年齢人口は減少を続けるため、介護人材確保とサービス効率化が全国的な課題です。"));

S.push(subsection("川崎町の現状と課題", 1));
S.push(p("川崎町では、令和7年3月31日時点で高齢化率が41.4%に達し、宮城県内でも上位水準（県内5位）の高齢化が進んでいます。中山間に位置する地理的特性、人口減少、独居高齢者・高齢者世帯の増加、介護人材の不足など、本町を取り巻く課題は多岐にわたります。第9期計画策定から3年が経過する中で、令和7年3月の高齢者外出タクシー助成事業の終了、介護施設の役場周辺偏在化、認知症基本法（令和6年1月施行）への対応など、新たな政策課題も生じています。"));

S.push(subsection("本計画の目的", 1));
S.push(p("こうした状況のもと、本計画は介護保険法第117条に基づく市町村介護保険事業計画として、また老人福祉法第20条の8に基づく市町村老人福祉計画として一体的に策定し、令和9年度から令和11年度までの3年間を計画期間として、本町の高齢者施策を総合的・計画的に推進することを目的とします。"));
S.push(p("また、令和6年1月に施行された「共生社会の実現を推進するための認知症基本法」に基づき、市町村認知症施策推進計画を本計画に包含する形で、新たに独立章として位置付けることとしています。"));

S.push(fact("本計画は、介護保険法・老人福祉法・認知症基本法の3つの法律に基づく一体的計画である。"));

// 1-2
S.push(section(1, 2, "計画の位置付け"));
S.push(p("本計画は、以下の法令・上位計画・関連計画との整合を図りながら策定します。"));

S.push(subsection("法令上の根拠", 1));
S.push(numItem("①", "介護保険法第117条（市町村介護保険事業計画）"));
S.push(numItem("②", "老人福祉法第20条の8（市町村老人福祉計画）"));
S.push(numItem("③", "共生社会の実現を推進するための認知症基本法第14条（市町村認知症施策推進計画）"));

S.push(subsection("上位計画・関連計画との関係", 1));
const rcTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("区分", { width: 18 }),
      thcell("計画名", { width: 36 }),
      thcell("整合の方針", { width: 46 }),
    ]}),
    new TableRow({ children: [
      tcell("上位計画", { bold: true, fill: C.lblue }),
      tcell("第6次川崎町長期総合計画"),
      tcell("町の総合的なまちづくり方針との整合を確保"),
    ]}),
    new TableRow({ children: [
      tcell("関連計画\n（同時策定）", { bold: true, fill: C.lblue }),
      tcell("川崎町地域福祉計画（第3期）／川崎町障害者計画（同時策定）"),
      tcell("3計画同時策定のメリットを活かし、移送サービス・重層的支援体制等の重複論点を整理して整合を確保（受託：ジャパン総研）"),
    ]}),
    new TableRow({ children: [
      tcell("健康関連計画", { bold: true, fill: C.lblue }),
      tcell("健康かわさき21計画／川崎町データヘルス計画"),
      tcell("健康寿命延伸・介護予防の観点で項目相互参照"),
    ]}),
    new TableRow({ children: [
      tcell("社協計画", { bold: true, fill: C.lblue }),
      tcell("川崎町社会福祉協議会 地域福祉活動計画"),
      tcell("社協独自計画。フレイル予防・介護予防事業のタイアップで連動（整合性確認は不要と整理）"),
    ]}),
    new TableRow({ children: [
      tcell("国の指針", { bold: true, fill: C.lblue }),
      tcell("介護保険事業に係る保険給付の円滑な実施を確保するための基本指針（厚労省告示）"),
      tcell("国指針の3層整合（国→県→町）の原則に基づく"),
    ]}),
    new TableRow({ children: [
      tcell("県の計画", { bold: true, fill: C.lblue }),
      tcell("宮城県高齢者保健福祉計画・介護保険事業支援計画"),
      tcell("県計画との整合（広域連携・人材確保）"),
    ]}),
  ],
});
S.push(rcTable);

// 1-3
S.push(section(1, 3, "計画の期間"));
S.push(p("本計画の計画期間は、令和9年度から令和11年度までの3年間です。"));

const termTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("期間"),
      thcell("令和6年度"), thcell("令和7年度"), thcell("令和8年度"),
      thcell("令和9年度"), thcell("令和10年度"), thcell("令和11年度"),
      thcell("令和12年度"),
    ]}),
    new TableRow({ children: [
      tcell("第9期", { bold: true, fill: C.lblue }),
      tcell("●", { align: AlignmentType.CENTER, fill: C.lgray }),
      tcell("●", { align: AlignmentType.CENTER, fill: C.lgray }),
      tcell("●", { align: AlignmentType.CENTER, fill: C.lgray }),
      tcell("", { align: AlignmentType.CENTER }),
      tcell("", { align: AlignmentType.CENTER }),
      tcell("", { align: AlignmentType.CENTER }),
      tcell("", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("第10期\n（本計画）", { bold: true, fill: C.lorange }),
      tcell("", { align: AlignmentType.CENTER }),
      tcell("", { align: AlignmentType.CENTER }),
      tcell("", { align: AlignmentType.CENTER }),
      tcell("●", { align: AlignmentType.CENTER, fill: C.orange, color: C.white, bold: true }),
      tcell("●", { align: AlignmentType.CENTER, fill: C.orange, color: C.white, bold: true }),
      tcell("●", { align: AlignmentType.CENTER, fill: C.orange, color: C.white, bold: true }),
      tcell("", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("第11期", { bold: true, fill: C.lblue }),
      tcell("", { align: AlignmentType.CENTER }),
      tcell("", { align: AlignmentType.CENTER }),
      tcell("", { align: AlignmentType.CENTER }),
      tcell("", { align: AlignmentType.CENTER }),
      tcell("", { align: AlignmentType.CENTER }),
      tcell("", { align: AlignmentType.CENTER }),
      tcell("●", { align: AlignmentType.CENTER, fill: C.lgray }),
    ]}),
  ],
});
S.push(termTable);

// 1-4
S.push(section(1, 4, "計画策定の体制"));
S.push(p("本計画は、以下の体制で策定します。"));

S.push(subsection("策定委員会", 1));
S.push(p("学識経験者・医療関係者・介護関係者・地域包括・住民代表・認知症関係・公募委員等から構成される「川崎町高齢者保健福祉計画・介護保険事業計画策定委員会」を設置し、計画期間中4回（令和8年8月中旬／令和8年11月／令和9年1月／令和9年2月）の審議を経て策定します。"));

S.push(subsection("住民意見の反映", 1));
S.push(p("住民意見の反映として、本計画の策定にあたっては以下の取組を実施します。"));
S.push(numItem("①", "一般高齢者ニーズ調査（対象：65歳以上の高齢者1,000名）"));
S.push(numItem("②", "要支援・要介護認定者調査（対象：認定者300名）"));
S.push(numItem("③", "認知症本人・家族の意見聴取（地域包括支援センター及び国保川崎病院経由）"));
S.push(p("①②は令和8年6月下旬発送・7月末回収を予定し、結果は令和8年8月中旬の第1回策定委員会で報告します。"));

S.push(subsection("関係機関等との連携", 1));
S.push(p("計画策定にあたっては、川崎町社会福祉協議会、川崎町地域包括支援センター、国民健康保険川崎病院、町内介護事業所、民生委員等との連携・協議を行います。また、地域福祉計画・障害者計画と同時策定中であることを踏まえ、計画間の整合確保のための情報共有・調整を行います。"));

// ===========================================================
// 第2章 川崎町の高齢者を取り巻く現状
// ===========================================================
S.push(...chapterTitle(2));

// 2-1
S.push(section(2, 1, "人口・世帯の現状"));
S.push(p("川崎町は、宮城県南西部の柴田郡に位置し、蔵王連峰のふもとに広がる中山間の町です。町の総面積は約270km²で、町域の約85%を山林が占めています。日常生活圏域は1圏域とし、町内を7地区（裏丁上下・本荒町中新町／前川青根／今宿／川内／本砂金／小野／小沢支倉碁石支倉台）に区分しています。"));

S.push(subsection("第1号被保険者数の推移", 2));

const popTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("区分", { width: 30, fill: CH[2].main }),
      thcell("令和3年度末\n（2022.3）", { width: 18, fill: CH[2].main }),
      thcell("令和7年6月\n（2025.6）", { width: 18, fill: CH[2].main }),
      thcell("増減数", { width: 12, fill: CH[2].main }),
      thcell("増減率", { width: 22, fill: CH[2].main }),
    ]}),
    new TableRow({ children: [
      tcell("65〜75歳未満（前期高齢者）", { bold: true }),
      tcell("1,745人", { align: AlignmentType.RIGHT }),
      tcell("1,569人", { align: AlignmentType.RIGHT }),
      tcell("△176", { align: AlignmentType.RIGHT, color: C.red }),
      tcell("△10.1%", { align: AlignmentType.RIGHT, color: C.red }),
    ]}),
    new TableRow({ children: [
      tcell("75〜85歳未満", { bold: true }),
      tcell("905人", { align: AlignmentType.RIGHT }),
      tcell("1,070人", { align: AlignmentType.RIGHT }),
      tcell("+165", { align: AlignmentType.RIGHT, color: C.green }),
      tcell("+18.2%", { align: AlignmentType.RIGHT, color: C.green, bold: true }),
    ]}),
    new TableRow({ children: [
      tcell("85歳以上", { bold: true }),
      tcell("605人", { align: AlignmentType.RIGHT }),
      tcell("605人", { align: AlignmentType.RIGHT }),
      tcell("±0", { align: AlignmentType.RIGHT }),
      tcell("±0.0%", { align: AlignmentType.RIGHT }),
    ]}),
    new TableRow({ children: [
      tcell("計", { bold: true, fill: C.lblue }),
      tcell("3,255人", { align: AlignmentType.RIGHT, bold: true, fill: C.lblue }),
      tcell("3,244人", { align: AlignmentType.RIGHT, bold: true, fill: C.lblue }),
      tcell("△11", { align: AlignmentType.RIGHT, bold: true, fill: C.lblue }),
      tcell("△0.3%", { align: AlignmentType.RIGHT, bold: true, fill: C.lblue }),
    ]}),
    new TableRow({ children: [
      tcell("（再掲）後期高齢者75歳以上", { italics: true, fill: C.lgray }),
      tcell("1,510人", { align: AlignmentType.RIGHT, fill: C.lgray }),
      tcell("1,675人", { align: AlignmentType.RIGHT, fill: C.lgray }),
      tcell("+165", { align: AlignmentType.RIGHT, color: C.green, fill: C.lgray }),
      tcell("+10.9%", { align: AlignmentType.RIGHT, color: C.green, bold: true, fill: C.lgray }),
    ]}),
    new TableRow({ children: [
      tcell("（再掲）住所地特例", { italics: true, fill: C.lgray }),
      tcell("34人", { align: AlignmentType.RIGHT, fill: C.lgray }),
      tcell("24人", { align: AlignmentType.RIGHT, fill: C.lgray }),
      tcell("△10", { align: AlignmentType.RIGHT, fill: C.lgray }),
      tcell("△29.4%", { align: AlignmentType.RIGHT, fill: C.lgray }),
    ]}),
  ],
});
S.push(popTable);
S.push(source("介護保険事業状況報告（保険者データ R7.6月分・年報データR3年度）"));

// 図2-1：人口推移グラフ
S.push(image("/home/claude/kawasaki_work/chart_population.png", { width: 540, height: 340 }));
S.push(caption("図2-1　第1号被保険者の年齢階級別推移（R3年度末→R7.6月）"));

S.push(fact("第1号被保険者総数はほぼ横ばいだが、後期高齢者（75歳以上）は約11%増加。年齢構成の高齢化が進行している。"));

S.push(p("第1号被保険者数の総数は令和3年度末3,255人から令和7年6月3,244人と微減ですが、内訳を見ると前期高齢者（65〜75歳未満）が約10%減少する一方、後期高齢者（75歳以上）は約11%増加しており、年齢構成の高齢化が顕著に進んでいます。後期高齢者は要介護認定率が前期高齢者の約4〜5倍となるため、今後の認定者数・サービス需要の増加が見込まれます。"));

S.push(p("特に75歳以上85歳未満の階層が令和3年度末905人から令和7年6月1,070人と4年で165人（+18.2%）増加していることは、団塊の世代（昭和22〜24年生まれ）が当該年齢階層に到達したことによる影響と考えられます。今後も後期高齢者前段の人口圧力は継続するため、認知症・要介護認定者の増加とサービス需要の伸びへの対応が第10期計画の重要課題となります。"));

S.push(subsection("人口動態（令和7年6月単月）", 2));
S.push(bullet("当月中増：11人（うち65歳到達8人・転入2人・その他1人）"));
S.push(bullet("当月中減：15人（うち死亡12人・転出3人）"));
S.push(bullet("自然動態は減少局面（死亡>65歳到達）に入っており、第1号被保険者数は今後緩やかに減少する見込み"));

S.push(subsection("世帯の状況", 2));
S.push(p("令和3年度末時点で、第1号被保険者のいる世帯数は2,181世帯です。キックオフ会議では、独居高齢者・高齢者夫婦のみ世帯に加え、未婚の子と高齢の親が同居する世帯における老老介護や8050問題が町の課題として挙げられています。"));
S.push(placeholder("世帯類型別の詳細（独居・夫婦のみ・同居等の人数）は、令和8年6月下旬発送の一般高齢者ニーズ調査により把握予定です。【アンケート結果反映後追記】"));

// 2-2
S.push(section(2, 2, "高齢者の状況"));

S.push(subsection("高齢化率", 2));
S.push(p("川崎町の高齢化率は令和7年3月31日時点で41.4%に達し、宮城県内35市町村中5位の高水準となっています。これは全国平均（約29.1%）・宮城県平均（約28.5%）を大きく上回ります。"));

// 図2-5：高齢化率比較
S.push(image("/home/claude/kawasaki_work/chart_aging.png", { width: 480, height: 300 }));
S.push(caption("図2-5　高齢化率比較（令和7年3月31日時点）"));

S.push(fact("高齢化率41.4%は県内5位（令和7年3月31日時点）。中山間地域・人口減少地域としての特性を踏まえた施策展開が求められる。"));

S.push(p("川崎町は宮城県南部の中山間地域に位置し、人口減少と高齢化が同時進行する典型的な小規模町です。蔵王連峰のふもとに点在する集落構造、町域の約85%を占める山林、第二次産業従事者の減少といった地理的・社会的特性が、高齢化の急速な進行と地域偏在を生み出す要因となっています。第10期計画では、こうした地域特性を踏まえ、移動支援・地域見守り・広域連携を計画の柱に据えます。"));

S.push(subsection("所得段階別第1号被保険者の構成", 2));

const incomeTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("段階", { width: 10, fill: CH[2].main }),
      thcell("区分", { width: 40, fill: CH[2].main }),
      thcell("人数", { width: 15, fill: CH[2].main }),
      thcell("構成比", { width: 15, fill: CH[2].main }),
      thcell("分類", { width: 20, fill: CH[2].main }),
    ]}),
    new TableRow({ children: [
      tcell("第1段階", { bold: true }), tcell("世帯非課税（年金80万円以下）"),
      tcell("431人", { align: AlignmentType.RIGHT }), tcell("13.2%", { align: AlignmentType.RIGHT }),
      tcell("非課税層", { fill: C.lgreen, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("第2段階", { bold: true }), tcell("世帯非課税（年金80〜120万円）"),
      tcell("273人", { align: AlignmentType.RIGHT }), tcell("8.4%", { align: AlignmentType.RIGHT }),
      tcell("非課税層", { fill: C.lgreen, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("第3段階", { bold: true }), tcell("世帯非課税（年金120万円超）"),
      tcell("275人", { align: AlignmentType.RIGHT }), tcell("8.4%", { align: AlignmentType.RIGHT }),
      tcell("非課税層", { fill: C.lgreen, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("第4段階", { bold: true }), tcell("本人非課税・世帯課税"),
      tcell("466人", { align: AlignmentType.RIGHT }), tcell("14.3%", { align: AlignmentType.RIGHT }),
      tcell("本人非課税", { fill: C.lblue, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("第5段階", { bold: true, fill: C.lorange }),
      tcell("本人非課税・世帯課税（基準）", { fill: C.lorange }),
      tcell("697人", { align: AlignmentType.RIGHT, fill: C.lorange, bold: true }),
      tcell("21.4%", { align: AlignmentType.RIGHT, fill: C.lorange, bold: true }),
      tcell("本人非課税", { fill: C.lblue, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("第6段階", { bold: true }), tcell("本人課税（合計所得120万円未満）"),
      tcell("449人", { align: AlignmentType.RIGHT }), tcell("13.8%", { align: AlignmentType.RIGHT }),
      tcell("課税層", { fill: C.lorange, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("第7段階", { bold: true }), tcell("本人課税（120〜210万円）"),
      tcell("385人", { align: AlignmentType.RIGHT }), tcell("11.8%", { align: AlignmentType.RIGHT }),
      tcell("課税層", { fill: C.lorange, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("第8段階", { bold: true }), tcell("本人課税（210〜320万円）"),
      tcell("173人", { align: AlignmentType.RIGHT }), tcell("5.3%", { align: AlignmentType.RIGHT }),
      tcell("課税層", { fill: C.lorange, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("第9段階", { bold: true }), tcell("本人課税（320万円以上）"),
      tcell("106人", { align: AlignmentType.RIGHT }), tcell("3.3%", { align: AlignmentType.RIGHT }),
      tcell("課税層", { fill: C.lorange, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("計", { bold: true, fill: C.lblue }),
      tcell("─", { align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("3,255人", { align: AlignmentType.RIGHT, bold: true, fill: C.lblue }),
      tcell("100.0%", { align: AlignmentType.RIGHT, bold: true, fill: C.lblue }),
      tcell("─", { fill: C.lblue, align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(incomeTable);
S.push(source("年報データ2021（介護保険事業状況報告 令和3年度・様式1所得段階別）"));

// 図2-4：所得段階別人口分布グラフ
S.push(image("/home/claude/kawasaki_work/chart_income.png", { width: 560, height: 320 }));
S.push(caption("図2-4　所得段階別第1号被保険者の構成（令和3年度末・計3,255人）"));

S.push(p("非課税層（第1〜3段階）が979人（30.1%）を占め、低所得高齢者への保険料軽減・補足給付の重要性が示されています。一方、本人課税層（第6〜9段階）は1,113人（34.2%）で、相対的な保険料負担力を有する層も一定規模存在しています。"));

S.push(p("特に第1段階（年金収入80万円以下の世帯非課税層）が431人と多く、後期高齢者の増加に伴いさらに増加する可能性があります。この層は保険料軽減措置（5割軽減）の対象であり、町の一般会計繰入（低所得者保険料軽減繰入金として令和3年度実績で約13,555千円）が継続的に必要となります。"));

S.push(p("第10期保険料算定にあたっては、所得段階区分の精緻化（現行9段階から13段階等への見直し）を国の指針に沿って検討します。13段階化することで、高所得層の負担割合を厚くし、低所得層の負担を軽減する効果が期待できます。詳細は第7章で論じます。"));

// 2-3
S.push(section(2, 3, "要介護認定とサービス利用の状況"));

S.push(subsection("サービス受給者数（令和7年6月時点）", 2));

const recvTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("サービス区分", { width: 28, fill: CH[2].main }),
      thcell("要支援1", { width: 8, fill: CH[2].main }),
      thcell("要支援2", { width: 8, fill: CH[2].main }),
      thcell("要介護1", { width: 8, fill: CH[2].main }),
      thcell("要介護2", { width: 8, fill: CH[2].main }),
      thcell("要介護3", { width: 8, fill: CH[2].main }),
      thcell("要介護4", { width: 8, fill: CH[2].main }),
      thcell("要介護5", { width: 8, fill: CH[2].main }),
      thcell("計", { width: 16, fill: CH[2].main }),
    ]}),
    new TableRow({ children: [
      tcell("居宅（介護予防）サービス", { bold: true }),
      tcell("13", { align: AlignmentType.RIGHT }), tcell("50", { align: AlignmentType.RIGHT }),
      tcell("39", { align: AlignmentType.RIGHT }), tcell("77", { align: AlignmentType.RIGHT }),
      tcell("66", { align: AlignmentType.RIGHT }), tcell("23", { align: AlignmentType.RIGHT }),
      tcell("8", { align: AlignmentType.RIGHT }), tcell("276人", { align: AlignmentType.RIGHT, bold: true }),
    ]}),
    new TableRow({ children: [
      tcell("地域密着型サービス", { bold: true }),
      tcell("0", { align: AlignmentType.RIGHT }), tcell("0", { align: AlignmentType.RIGHT }),
      tcell("9", { align: AlignmentType.RIGHT }), tcell("11", { align: AlignmentType.RIGHT }),
      tcell("17", { align: AlignmentType.RIGHT }), tcell("12", { align: AlignmentType.RIGHT }),
      tcell("7", { align: AlignmentType.RIGHT }), tcell("56人", { align: AlignmentType.RIGHT, bold: true }),
    ]}),
    new TableRow({ children: [
      tcell("施設サービス計", { bold: true }),
      tcell("0", { align: AlignmentType.RIGHT }), tcell("0", { align: AlignmentType.RIGHT }),
      tcell("6", { align: AlignmentType.RIGHT }), tcell("15", { align: AlignmentType.RIGHT }),
      tcell("40", { align: AlignmentType.RIGHT }), tcell("43", { align: AlignmentType.RIGHT }),
      tcell("30", { align: AlignmentType.RIGHT }), tcell("134人", { align: AlignmentType.RIGHT, bold: true }),
    ]}),
    new TableRow({ children: [
      tcell("　うち特養", { italics: true, fill: C.lgray }),
      tcell("0", { align: AlignmentType.RIGHT, fill: C.lgray }), tcell("0", { align: AlignmentType.RIGHT, fill: C.lgray }),
      tcell("1", { align: AlignmentType.RIGHT, fill: C.lgray }), tcell("2", { align: AlignmentType.RIGHT, fill: C.lgray }),
      tcell("23", { align: AlignmentType.RIGHT, fill: C.lgray }), tcell("21", { align: AlignmentType.RIGHT, fill: C.lgray }),
      tcell("21", { align: AlignmentType.RIGHT, fill: C.lgray }), tcell("68人", { align: AlignmentType.RIGHT, fill: C.lgray }),
    ]}),
    new TableRow({ children: [
      tcell("　うち老健", { italics: true, fill: C.lgray }),
      tcell("0", { align: AlignmentType.RIGHT, fill: C.lgray }), tcell("0", { align: AlignmentType.RIGHT, fill: C.lgray }),
      tcell("5", { align: AlignmentType.RIGHT, fill: C.lgray }), tcell("13", { align: AlignmentType.RIGHT, fill: C.lgray }),
      tcell("17", { align: AlignmentType.RIGHT, fill: C.lgray }), tcell("22", { align: AlignmentType.RIGHT, fill: C.lgray }),
      tcell("9", { align: AlignmentType.RIGHT, fill: C.lgray }), tcell("66人", { align: AlignmentType.RIGHT, fill: C.lgray }),
    ]}),
    new TableRow({ children: [
      tcell("延べ合計（重複あり）", { bold: true, fill: C.lblue }),
      tcell("13", { align: AlignmentType.RIGHT, bold: true, fill: C.lblue }), tcell("50", { align: AlignmentType.RIGHT, bold: true, fill: C.lblue }),
      tcell("54", { align: AlignmentType.RIGHT, bold: true, fill: C.lblue }), tcell("103", { align: AlignmentType.RIGHT, bold: true, fill: C.lblue }),
      tcell("123", { align: AlignmentType.RIGHT, bold: true, fill: C.lblue }), tcell("78", { align: AlignmentType.RIGHT, bold: true, fill: C.lblue }),
      tcell("45", { align: AlignmentType.RIGHT, bold: true, fill: C.lblue }), tcell("466人", { align: AlignmentType.RIGHT, bold: true, fill: C.lblue }),
    ]}),
  ],
});
S.push(recvTable);
S.push(source("介護保険事業状況報告（保険者データ202506・様式1の6）"));

// 図2-2：サービス受給者構成円グラフ
S.push(image("/home/claude/kawasaki_work/chart_recipient.png", { width: 420, height: 360 }));
S.push(caption("図2-2　サービス受給者の区分別構成（令和7年6月時点・延べ466人）"));

S.push(p("令和7年6月時点で、サービス受給者は延べ466人（重複あり）です。サービス区分別構成は、居宅サービス59.2%、地域密着型12.0%、施設28.8%となっており、在宅介護が主軸を占めています。"));

S.push(p("ただし、要介護度別の利用状況を見ると、要介護3以上の重度層では施設サービス利用が中心となります。施設サービス受給者134人のうち、要介護3以上が104人（77.6%）を占めており、重度認定者の受け皿として施設機能が機能していることが分かります。一方、要介護1〜2の中軽度層では居宅サービス・地域密着型サービスでの対応が中心となっています。"));

S.push(fact("施設サービス受給者134人のうち、特養68人・老健66人がほぼ拮抗。介護療養型・介護医療院は0人で、第10期見込量算定では0据置で問題なし。"));

S.push(subsection("住所地特例（町外施設入所）の状況", 2));
S.push(p("令和7年6月時点で、住所地特例被保険者は24人となっています。キックオフ会議では、町内施設の縮小傾向により町外施設の利用が相当数あること、また福祉施設が役場周辺に偏在しており地域間でアクセス格差が見られることが課題として挙げられています。"));

S.push(placeholder("地区別の認定者・サービス利用状況、町外施設利用者の所在自治体内訳は、町担当課から経年データを入手次第追記します。"));

// 2-4
S.push(section(2, 4, "介護給付費の状況"));

S.push(subsection("年間給付費（令和3年度実績）", 2));

const kyufuTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("サービス区分", { width: 28, fill: CH[2].main }),
      thcell("要支援計", { width: 18, fill: CH[2].main }),
      thcell("要介護計", { width: 18, fill: CH[2].main }),
      thcell("合計（円）", { width: 22, fill: CH[2].main }),
      thcell("構成比", { width: 14, fill: CH[2].main }),
    ]}),
    new TableRow({ children: [
      tcell("居宅サービス計", { bold: true, fill: C.lgreen }),
      tcell("26,821,755", { align: AlignmentType.RIGHT, fill: C.lgreen }),
      tcell("339,234,895", { align: AlignmentType.RIGHT, fill: C.lgreen }),
      tcell("366,056,650", { align: AlignmentType.RIGHT, fill: C.lgreen, bold: true }),
      tcell("38.1%", { align: AlignmentType.RIGHT, fill: C.lgreen, bold: true }),
    ]}),
    new TableRow({ children: [
      tcell("　うち訪問サービス", { italics: true }),
      tcell("2,490,773", { align: AlignmentType.RIGHT }),
      tcell("54,605,942", { align: AlignmentType.RIGHT }),
      tcell("57,096,715", { align: AlignmentType.RIGHT }),
      tcell("5.9%", { align: AlignmentType.RIGHT }),
    ]}),
    new TableRow({ children: [
      tcell("　うち通所サービス", { italics: true }),
      tcell("16,765,707", { align: AlignmentType.RIGHT }),
      tcell("173,710,493", { align: AlignmentType.RIGHT }),
      tcell("190,476,200", { align: AlignmentType.RIGHT }),
      tcell("19.8%", { align: AlignmentType.RIGHT }),
    ]}),
    new TableRow({ children: [
      tcell("地域密着型計", { bold: true, fill: C.lgreen }),
      tcell("0", { align: AlignmentType.RIGHT, fill: C.lgreen }),
      tcell("152,764,596", { align: AlignmentType.RIGHT, fill: C.lgreen }),
      tcell("152,764,596", { align: AlignmentType.RIGHT, fill: C.lgreen, bold: true }),
      tcell("15.9%", { align: AlignmentType.RIGHT, fill: C.lgreen, bold: true }),
    ]}),
    new TableRow({ children: [
      tcell("施設サービス計", { bold: true, fill: C.lgreen }),
      tcell("0", { align: AlignmentType.RIGHT, fill: C.lgreen }),
      tcell("441,678,518", { align: AlignmentType.RIGHT, fill: C.lgreen }),
      tcell("441,678,518", { align: AlignmentType.RIGHT, fill: C.lgreen, bold: true }),
      tcell("46.0%", { align: AlignmentType.RIGHT, fill: C.lgreen, bold: true }),
    ]}),
    new TableRow({ children: [
      tcell("　うち特養", { italics: true }),
      tcell("0", { align: AlignmentType.RIGHT }),
      tcell("193,053,976", { align: AlignmentType.RIGHT }),
      tcell("193,053,976", { align: AlignmentType.RIGHT }),
      tcell("20.1%", { align: AlignmentType.RIGHT }),
    ]}),
    new TableRow({ children: [
      tcell("　うち老健", { italics: true }),
      tcell("0", { align: AlignmentType.RIGHT }),
      tcell("248,624,542", { align: AlignmentType.RIGHT }),
      tcell("248,624,542", { align: AlignmentType.RIGHT }),
      tcell("25.9%", { align: AlignmentType.RIGHT }),
    ]}),
    new TableRow({ children: [
      tcell("総計", { bold: true, fill: C.lblue }),
      tcell("26,821,755", { align: AlignmentType.RIGHT, bold: true, fill: C.lblue }),
      tcell("933,678,009", { align: AlignmentType.RIGHT, bold: true, fill: C.lblue }),
      tcell("960,499,764", { align: AlignmentType.RIGHT, bold: true, fill: C.lblue }),
      tcell("100.0%", { align: AlignmentType.RIGHT, bold: true, fill: C.lblue }),
    ]}),
  ],
});
S.push(kyufuTable);
S.push(source("年報データ2021（介護保険事業状況報告 令和3年度・様式2給付費）"));

// 図2-3：給付費構成横棒グラフ
S.push(image("/home/claude/kawasaki_work/chart_benefit.png", { width: 560, height: 340 }));
S.push(caption("図2-3　サービス区分別の年間給付費（令和3年度実績・総額9.6億円）"));

S.push(fact("令和3年度のサービス種類別給付費合計は約9.6億円。施設サービスが46.0%、居宅サービスが38.1%、地域密着型が15.9%の構成。"));

S.push(p("施設サービスが給付費全体の46.0%を占めていることは、川崎町の介護給付構造の大きな特徴です。受給者数では28.8%（134人）に過ぎない施設サービス利用者が、給付費の半分近くを占めるのは、施設サービスの1人当たり単価が居宅サービスより大幅に高いためです。第10期計画では、町外施設利用（住所地特例24人）を含めた施設供給見込量を精査し、給付費の中長期的な持続可能性を確保することが課題です。"));

S.push(p("一方、居宅サービス内訳では通所サービス（デイサービス・デイケア）が約1.9億円（給付費全体の19.8%）と最大の比重を占めています。これは町内の通所事業所の充実度を反映するとともに、在宅介護を支える基盤として通所サービスが重要な役割を果たしていることを示しています。地域密着型では認知症対応型共同生活介護（グループホーム）が約8,249万円と地域密着型給付費の54%を占め、認知症ケアにおける重要拠点となっています。"));

S.push(subsection("介護保険特別会計の状況（令和3年度）", 2));

const fiscalTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("区分", { width: 50, fill: CH[2].main }),
      thcell("金額（円）", { width: 28, fill: CH[2].main }),
      thcell("構成比", { width: 22, fill: CH[2].main }),
    ]}),
    new TableRow({ children: [
      tcell("保険給付費 計", { bold: true, fill: C.lorange }),
      tcell("1,042,179,056", { align: AlignmentType.RIGHT, bold: true, fill: C.lorange }),
      tcell("100.0%", { align: AlignmentType.RIGHT, bold: true, fill: C.lorange }),
    ]}),
    new TableRow({ children: [
      tcell("　うち介護サービス等諸費", { italics: true }),
      tcell("933,658,461", { align: AlignmentType.RIGHT }),
      tcell("89.6%", { align: AlignmentType.RIGHT }),
    ]}),
    new TableRow({ children: [
      tcell("　うち介護予防サービス等諸費", { italics: true }),
      tcell("26,841,303", { align: AlignmentType.RIGHT }),
      tcell("2.6%", { align: AlignmentType.RIGHT }),
    ]}),
    new TableRow({ children: [
      tcell("　うち高額介護サービス等費", { italics: true }),
      tcell("26,815,719", { align: AlignmentType.RIGHT }),
      tcell("2.6%", { align: AlignmentType.RIGHT }),
    ]}),
    new TableRow({ children: [
      tcell("　うち特定入所者介護サービス等費", { italics: true }),
      tcell("52,467,842", { align: AlignmentType.RIGHT }),
      tcell("5.0%", { align: AlignmentType.RIGHT }),
    ]}),
    new TableRow({ children: [
      tcell("地域支援事業費 計", { bold: true, fill: C.lblue }),
      tcell("32,071,763", { align: AlignmentType.RIGHT, bold: true, fill: C.lblue }),
      tcell("─", { align: AlignmentType.CENTER, fill: C.lblue }),
    ]}),
  ],
});
S.push(fiscalTable);
S.push(source("年報データ2021（介護保険事業状況報告 令和3年度・様式4）"));

S.push(subsection("保険料収納実績（令和3年度）", 2));
S.push(bullet("調定額累計：237,836,453円　収納額累計：229,136,570円"));
S.push(bullet("収納率合計：96.3%（特別徴収100%・普通徴収現年度分87.3%）"));
S.push(bullet("不納欠損額：1,207,220円　未収額：7,492,663円"));
S.push(fact("収納率96.3%は高水準。第10期保険料試算における予定収納率の根拠となる。"));

// 2-5
S.push(section(2, 5, "高齢者を取り巻く地域資源"));

S.push(subsection("地域包括支援センター", 2));
S.push(p("町内に1か所、社会福祉法人川崎町社会福祉協議会が運営する地域包括支援センターを設置しています（地域包括支援センターコード：0472200062）。"));
S.push(p("職員体制は保健師3名＋認定調査員1名の計4名（実質3名で運用）、ケアマネジャーは会計年度任用職員1名となっています。介護認定調査は1名で実施しており、業務負担が大きい状況です。一部のケアプランは外部委託も活用しています。"));

S.push(subsection("医療資源", 2));
S.push(p("町内には国民健康保険川崎病院（電話 0224-84-2119）があり、町の医療拠点となっています。町外の中核的医療機関として、みやぎ県南中核病院（大河原町）・刈田綜合病院（白石市）が近隣にあり、夜間・救急時の広域医療連携の対象となっています。"));

S.push(subsection("認知症関連資源", 2));
S.push(p("認知症サポーター養成事業を継続しており、累計550名のサポーターと73名のキャラバンメイトを育成しています。第10期では、認知症基本法対応のチームオレンジ整備・本人ミーティング等の取組強化が課題です。"));

S.push(subsection("ユニバーサルサポーター制度（川崎町独自）", 2));
S.push(p("川崎町では、地域の高齢者を支える独自の人的資源として「ユニバーサルサポーター制度」を運用しており、14種別約400名のサポーターが活動しています。主な内訳は介護予防サロン84名、スマイルサポーター40名、レクリエーションサポーター29名、傾聴サポーター24名、生活支援サポーター25名、ふれあいネットワーク活動員15名・協力員130名等です。"));

S.push(fact("ユニバーサルサポーター14種別約400名は川崎町独自の重要な地域資源。第10期の生活支援体制整備・通いの場展開の中核となる。"));

S.push(subsection("移動支援", 2));
S.push(p("町内の移動支援は、令和7年3月で従来の高齢者外出タクシー利用助成が終了し、現在は以下の3つの仕組みで実施しています。"));
S.push(numItem("①", "福祉移送サービス：社会福祉協議会及びNPO法人が町の委託事業として運行（対象：高齢者・障害者）"));
S.push(numItem("②", "デマンドバス：バス会社運行（前日17時までに予約・病院利用が多い）"));
S.push(numItem("③", "町民バス：バス会社運行（朝晩は小中学生も利用）"));
S.push(p("これらの所管は地域振興課・町民生活課・社協・NPO法人と分かれており、第10期では移動支援全体を俯瞰した整理と利用情報の住民への提供が課題です。"));

S.push(subsection("独自施策", 2));
S.push(bullet("高齢者紙おむつ等支給事業"));
S.push(bullet("高齢者世帯エアコン購入支援事業（令和7年10月開始）"));
S.push(bullet("人工透析患者通院交通費助成事業"));

// ===========================================================
// 第3章 第9期計画の取組実績と評価
// ===========================================================
S.push(...chapterTitle(3));

// 3-1
S.push(section(3, 1, "第9期計画の体系"));
S.push(p("第9期計画（令和6〜8年度）は、第8期計画の体系を踏襲しつつ、地域包括ケアシステムのさらなる深化を目指し、以下の基本理念・基本目標のもとに体系化されました。第10期計画は、この体系を基本的に継承します。"));

S.push(subsection("第9期計画策定の経緯", 3));
S.push(p("第9期計画は令和6年3月に町長へ答申され、同年6月に確定・公表されました。前期（第8期）策定業者は福祉工房であり、第9期策定では地域住民意見の反映（パブリックコメント）は実施されておらず、策定委員会での審議をもって意見集約に代える運用がとられました。"));
S.push(p("第9期計画は、地域包括ケアシステムの深化と高齢者の地域生活継続を重視し、健康づくり・介護予防・在宅生活支援・地域連携を柱とする体系で策定されました。第10期計画ではこの体系を基本的に継承するとともに、認知症基本法対応（独立章化）、移動支援の再構築、施設偏在への対応など、新たな課題を追加して体系化します。"));

S.push(placeholder("第9期計画の基本理念・基本目標の具体的記述は、第9期計画書本編からの抜粋・要約として令和8年6月中に追記します。"));

S.push(subsection("第9期計画の体系（推定）", 3));
S.push(p("（キックオフでの「9期計画を踏襲する形」との方針確認、及び従来の高齢者保健福祉計画における標準的体系から）"));
S.push(bullet("基本理念：住民が住み慣れた地域で安心して暮らせるまちづくり"));
S.push(bullet("基本目標1：健康づくりと介護予防の推進"));
S.push(bullet("基本目標2：高齢者が安心して暮らせる仕組みづくり"));
S.push(bullet("基本目標3：在宅生活の継続支援"));
S.push(bullet("基本目標4：介護サービスの質の確保と提供体制"));
S.push(bullet("基本目標5：地域包括ケアシステムの深化"));

// 3-2
S.push(section(3, 2, "主要施策の取組実績"));
S.push(p("本節では、第9期計画期間中（令和6〜8年度）に川崎町が実施した主要施策の取組実績を、第4章で示す6つの基本目標の体系に沿って整理します。各事業の数値実績（実施回数・対象者数・事業費等）の確定値は、別添「川崎町第9期介護保険事業計画 実績一覧（町記入用フォーマット）」（全7シート・35事業以上を網羅）への町担当課ご記入を経て、Ver.2.0で本表に反映します。"));

S.push(fact("第9期取組実績の確定値は、別添『川崎町_第9期実績一覧_町記入用.xlsx』との連動で管理する。本素案v1.3では事業項目と既知数値を提示する。"));

// =====================================================
// 3-2-1 基本目標1関連
// =====================================================
S.push(subsection("3-2-1　健康づくり・介護予防の推進（基本目標1関連）", 3));

S.push(p("健康づくりと介護予防の推進については、町独自のユニバーサルサポーター制度を基盤に、住民主体の通いの場・介護予防活動を展開しました。ユニバーサルサポーターの14種別約400名による多彩な活動が、町内の介護予防活動の中核となっています。"));

const g1Table = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("事業区分", { width: 22, fill: CH[3].main }),
      thcell("主な事業", { width: 30, fill: CH[3].main }),
      thcell("規模・既知実績", { width: 30, fill: CH[3].main }),
      thcell("R6/R7実績値", { width: 18, fill: CH[3].main }),
    ]}),
    new TableRow({ children: [
      tcell("一般介護予防", { bold: true, fill: C.lblue }),
      tcell("元気まんてん教室・スマイル教室・パドル運動教室", { bold: true }),
      tcell("町・社協合同で定期開催。継続実施中"),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("通いの場", { bold: true, fill: C.lblue }),
      tcell("介護予防サロン（地区サロン）", { bold: true }),
      tcell("ユニバーサルサポーター制度に基づく住民主体運営"),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("サポーター", { bold: true, fill: C.lblue }),
      tcell("ユニバーサルサポーター養成", { bold: true }),
      tcell("介護予防サロン84名、スマイル40名、レク29名、傾聴24名、SC25名等14種別約400名"),
      tcell("【町確認・最新値】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("健診・早期発見", { bold: true, fill: C.lblue }),
      tcell("特定健診・後期高齢者健診・歯科健診", { bold: true }),
      tcell("健康かわさき21・データヘルス計画と連動"),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("配食・栄養", { bold: true, fill: C.lblue }),
      tcell("配食サービス・栄養指導", { bold: true }),
      tcell("オーラルフレイル・低栄養予防の観点で実施"),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(g1Table);
S.push(source("詳細：別添『川崎町_第9期実績一覧_町記入用.xlsx』02_介護予防_健康"));

S.push(p("第10期への課題：ユニバーサルサポーター制度の質的向上、後期高齢者（75歳以上1,675人）の急増に対応した介護予防の対象拡大、中山間地域居住者の通いの場参加への移動支援との連動が課題です。"));

// =====================================================
// 3-2-2 基本目標2関連
// =====================================================
S.push(subsection("3-2-2　高齢者が安心して暮らせる仕組みづくり（基本目標2関連）", 3));

S.push(p("見守り・移動支援・住まいの確保等、独居高齢者・高齢者世帯・老老介護世帯が安心して暮らせる仕組みを整備しました。特に移動支援は、令和7年3月の高齢者外出タクシー助成終了に伴い、社協・NPO移送＋デマンドバス＋町民バスの3層構造に移行する大きな転換点を迎えています。"));

const g2Table = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("事業区分", { width: 22, fill: CH[3].main }),
      thcell("主な事業", { width: 30, fill: CH[3].main }),
      thcell("規模・既知実績", { width: 30, fill: CH[3].main }),
      thcell("R6/R7実績値", { width: 18, fill: CH[3].main }),
    ]}),
    new TableRow({ children: [
      tcell("見守り", { bold: true, fill: C.lblue }),
      tcell("ふれあいネットワーク事業", { bold: true }),
      tcell("活動員15名＋協力員130名（計145名）による地域見守り"),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("緊急支援", { bold: true, fill: C.lblue }),
      tcell("緊急通報装置設置事業", { bold: true }),
      tcell("独居高齢者・高齢者夫婦世帯への設置"),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("移動支援", { bold: true, fill: C.lorange }),
      tcell("(旧)高齢者外出タクシー助成", { bold: true }),
      tcell("令和7年3月で終了。第9期途中での制度変更", { color: C.red }),
      tcell("R7.3 終了", { color: C.red, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("移動支援", { bold: true, fill: C.lblue }),
      tcell("福祉移送サービス（社協・NPO）", { bold: true }),
      tcell("タクシー助成終了に伴い役割が拡大"),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("移動支援", { bold: true, fill: C.lblue }),
      tcell("デマンドバス・町民バス", { bold: true }),
      tcell("地域振興課・町民生活課所管。3層構造に位置付け"),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("独自支援", { bold: true, fill: C.lblue }),
      tcell("高齢者紙おむつ等支給事業", { bold: true }),
      tcell("町独自施策として継続"),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("独自支援", { bold: true, fill: C.lorange }),
      tcell("高齢者世帯エアコン購入支援", { bold: true }),
      tcell("令和7年10月開始の新規事業", { color: C.green }),
      tcell("R7.10 開始", { color: C.green, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("独自支援", { bold: true, fill: C.lblue }),
      tcell("人工透析患者通院交通費助成", { bold: true }),
      tcell("町独自施策として継続"),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(g2Table);
S.push(source("詳細：別添『川崎町_第9期実績一覧_町記入用.xlsx』03_在宅生活支援"));

S.push(p("第10期への課題：タクシー助成終了後の3層移動支援制度（社協・NPO移送／デマンドバス／町民バス）の住民への周知強化、所管が分かれている（地域振興課・町民生活課・社協・NPO）状況での利用情報の一元化、福祉施設の役場周辺偏在への対応が重要課題です。"));

// =====================================================
// 3-2-3 基本目標3関連
// =====================================================
S.push(subsection("3-2-3　在宅生活継続の支援（基本目標3関連）", 3));

S.push(p("住み慣れた自宅・地域で生活を継続したいという高齢者と家族の希望に応えるため、在宅医療・介護連携、家族介護者支援、広域医療連携を推進しました。"));

const g3Table = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("事業区分", { width: 22, fill: CH[3].main }),
      thcell("主な事業", { width: 30, fill: CH[3].main }),
      thcell("規模・既知実績", { width: 30, fill: CH[3].main }),
      thcell("R6/R7実績値", { width: 18, fill: CH[3].main }),
    ]}),
    new TableRow({ children: [
      tcell("医療連携", { bold: true, fill: C.lblue }),
      tcell("在宅医療・介護連携推進事業", { bold: true }),
      tcell("国保川崎病院を中心とした連携"),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("広域医療", { bold: true, fill: C.lblue }),
      tcell("広域医療連携", { bold: true }),
      tcell("みやぎ県南中核病院（大河原）・刈田綜合病院（白石）等との連携"),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("家族介護", { bold: true, fill: C.lblue }),
      tcell("家族介護教室・介護相談", { bold: true }),
      tcell("家族介護者の負担軽減・知識習得支援"),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("家族介護", { bold: true, fill: C.lblue }),
      tcell("レスパイト支援（短期入所活用）", { bold: true }),
      tcell("介護者の休息確保のためのショートステイ利用支援"),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("住宅・用具", { bold: true, fill: C.lblue }),
      tcell("住宅改修・福祉用具利用支援", { bold: true }),
      tcell("介護保険給付による住宅改修・福祉用具貸与"),
      tcell("R3給付費約2,857万円（参考）", { align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(g3Table);

S.push(p("第10期への課題：施設サービス利用者134人のうち77.6%が要介護3以上の重度層に集中しており、軽度〜中度層の在宅生活継続を支える体制強化が必要です。また、未婚の子と高齢親が同居する8050問題、配偶者間の老老介護世帯の増加に対応した家族介護者支援と介護離職防止が重要課題です。"));

// =====================================================
// 3-2-4 基本目標4関連
// =====================================================
S.push(subsection("3-2-4　介護サービスの質の確保と提供体制（基本目標4関連）", 3));

S.push(p("町内介護サービスの質の確保、町外施設利用を含めた供給体制の整備、深刻化する介護人材不足への対応を進めました。特に住所地特例24人（令和7年6月時点）に代表される町外施設依存と、福祉施設の役場周辺への偏在は、第9期から第10期に持ち越される重要な構造的課題です。"));

const g4Table = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("事業区分", { width: 22, fill: CH[3].main }),
      thcell("主な事業", { width: 30, fill: CH[3].main }),
      thcell("規模・既知実績", { width: 30, fill: CH[3].main }),
      thcell("R6/R7実績値", { width: 18, fill: CH[3].main }),
    ]}),
    new TableRow({ children: [
      tcell("サービス供給", { bold: true, fill: C.lblue }),
      tcell("町内介護事業所", { bold: true }),
      tcell("訪問系・通所系・施設系・地域密着型"),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("施設サービス", { bold: true, fill: C.lblue }),
      tcell("特養（68人）・老健（66人）受給", { bold: true }),
      tcell("R7.6時点。介護療養型・介護医療院は0人"),
      tcell("受給者数で参考", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("住所地特例", { bold: true, fill: C.lorange }),
      tcell("町外施設利用者", { bold: true }),
      tcell("24人（R7.6時点）。町内施設縮小に伴い町外依存"),
      tcell("R7.6＝24人", { bold: true, color: C.orange, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("質の確保", { bold: true, fill: C.lblue }),
      tcell("ケアプラン点検・事業所指導", { bold: true }),
      tcell("町・包括による継続的な質確保"),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("人材確保", { bold: true, fill: C.lblue }),
      tcell("介護人材確保・育成支援", { bold: true }),
      tcell("宮城県社協の修学資金貸付制度等の活用周知"),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(g4Table);

S.push(p("第10期への課題：施設の地域偏在（役場周辺集中）への対応、住所地特例該当者の所在自治体内訳の把握と広域連携、介護人材不足への対応（ICT・介護ロボット導入促進を含む）、町外施設も含めたサービス供給見込量の整備が重要です。"));

// =====================================================
// 3-2-5 基本目標5関連
// =====================================================
S.push(subsection("3-2-5　地域包括ケアシステムの深化（基本目標5関連）", 3));

S.push(p("地域包括支援センターを中心に、医療・介護・予防・住まい・生活支援を一体的に提供する地域包括ケアシステムを深化させました。社会福祉法人川崎町社会福祉協議会が運営する地域包括支援センター（コード0472200062）は、保健師3名＋認定調査員1名の体制で運用されており、人材体制の強化が第10期の重要課題です。"));

const g5Table = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("事業区分", { width: 22, fill: CH[3].main }),
      thcell("主な事業", { width: 30, fill: CH[3].main }),
      thcell("規模・既知実績", { width: 30, fill: CH[3].main }),
      thcell("R6/R7実績値", { width: 18, fill: CH[3].main }),
    ]}),
    new TableRow({ children: [
      tcell("包括センター", { bold: true, fill: C.lblue }),
      tcell("地域包括支援センター運営", { bold: true }),
      tcell("保健師3＋認定調査1＝計4名（社協運営・実質3名運用）"),
      tcell("R7=4名", { bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("包括センター", { bold: true, fill: C.lblue }),
      tcell("総合相談支援業務", { bold: true }),
      tcell("認定調査1名で対応、業務負担大・ケアプラン外部委託あり"),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("生活支援", { bold: true, fill: C.lblue }),
      tcell("生活支援体制整備事業", { bold: true }),
      tcell("生活支援コーディネーター（SC）25名による地域づくり"),
      tcell("R7=SC25名", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("ケア会議", { bold: true, fill: C.lblue }),
      tcell("自立支援型地域ケア会議", { bold: true }),
      tcell("ケアプラン点検・自立支援に資する会議の継続実施"),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("権利擁護", { bold: true, fill: C.lblue }),
      tcell("成年後見・高齢者虐待対応", { bold: true }),
      tcell("成年後見制度の利用促進、虐待防止"),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("地域支援事業", { bold: true, fill: C.lgreen }),
      tcell("地域支援事業費（合計）", { bold: true }),
      tcell("R3年度実績：32,071,763円（総合事業565＋一般195＋包括/任意2,447万円）"),
      tcell("R3＝3,207万円", { bold: true, color: C.green, align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(g5Table);

S.push(p("第10期への課題：地域包括支援センターの体制強化（特に認定調査員の負担軽減）、3計画同時策定（地域福祉計画・障害者計画はジャパン総研が策定）の機会を活かした重層的支援体制の整備、生活支援コーディネーターとユニバーサルサポーター制度との連動深化が重要です。"));

// =====================================================
// 3-2-6 認知症施策
// =====================================================
S.push(subsection("3-2-6　認知症施策（第10期で基本目標6として独立章化）", 3));

S.push(p("第9期計画では、認知症施策は基本目標5（地域包括ケアシステムの深化）の一部として位置付けられていましたが、令和6年1月施行の認知症基本法を踏まえ、第10期では基本目標6として独立章化（第6章）します。第9期での主な取組は以下のとおりです。"));

const g6Table = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("事業区分", { width: 22, fill: CH[6].main }),
      thcell("主な事業", { width: 30, fill: CH[6].main }),
      thcell("規模・既知実績", { width: 30, fill: CH[6].main }),
      thcell("R6/R7実績値", { width: 18, fill: CH[6].main }),
    ]}),
    new TableRow({ children: [
      tcell("普及啓発", { bold: true, fill: CH[6].sub }),
      tcell("認知症サポーター養成講座", { bold: true }),
      tcell("累計550名養成（地域福祉計画資料時点）"),
      tcell("R7=550名累計", { bold: true, color: CH[6].main, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("普及啓発", { bold: true, fill: CH[6].sub }),
      tcell("認知症キャラバンメイト養成", { bold: true }),
      tcell("累計73名養成。サポーター養成講座の講師役"),
      tcell("R7=73名累計", { bold: true, color: CH[6].main, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("社会参加", { bold: true, fill: CH[6].sub }),
      tcell("認知症カフェ「喫茶みかん」", { bold: true }),
      tcell("認知症の人と家族の交流・相談の場"),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("早期発見", { bold: true, fill: CH[6].sub }),
      tcell("もの忘れ相談・初期集中支援", { bold: true }),
      tcell("認知症初期集中支援チーム・地域支援推進員による対応"),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("医療連携", { bold: true, fill: CH[6].sub }),
      tcell("国保川崎病院との連携", { bold: true }),
      tcell("認知症診断・治療・初期段階フォロー"),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("新規(未整備)", { bold: true, fill: C.lorange }),
      tcell("チームオレンジ整備", { bold: true, color: C.orange }),
      tcell("認知症基本法対応の新規取組。第10期で整備を目標", { color: C.orange }),
      tcell("R7＝未整備", { color: C.red, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("新規(未実施)", { bold: true, fill: C.lorange }),
      tcell("認知症本人ミーティング", { bold: true, color: C.orange }),
      tcell("基本法第3条「本人の意思を尊重」対応の新設候補", { color: C.orange }),
      tcell("R7＝未実施", { color: C.red, bold: true, align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(g6Table);
S.push(source("詳細：別添『川崎町_第9期実績一覧_町記入用.xlsx』04_認知症_包括"));

S.push(p("第10期への課題：認知症基本法対応として、（1）チームオレンジの新規整備、（2）認知症本人ミーティングの新規実施、（3）本人・家族の意見聴取（地域包括センター・国保川崎病院経由）、（4）KPIの3層構造化（プロセス・アウトプット・アウトカム）が必須となります。第6章で独立章として詳述します。"));

// =====================================================
// 3-2-7 実績データ連動の枠組み
// =====================================================
S.push(subsection("3-2-7　実績データ連動の枠組み", 3));

S.push(p("本節の各事業の数値実績は、別添「川崎町第9期実績一覧（町記入用フォーマット）」（7シート構成）と連動して管理します。町担当課で記入いただいた数値が、本素案Ver.2.0で本表に反映されます。"));

const linkTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("本節の項目", { width: 30, fill: CH[3].main }),
      thcell("対応する実績一覧シート", { width: 32, fill: CH[3].main }),
      thcell("反映時期", { width: 18, fill: CH[3].main }),
      thcell("優先度", { width: 20, fill: CH[3].main }),
    ]}),
    new TableRow({ children: [
      tcell("3-2-1 健康づくり・介護予防", { bold: true }),
      tcell("02_介護予防_健康", { fill: C.lblue }),
      tcell("Ver.2.0", { align: AlignmentType.CENTER }),
      tcell("A（優先・基本目標1根拠）", { color: C.red, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("3-2-2 仕組みづくり", { bold: true }),
      tcell("03_在宅生活支援", { fill: C.lblue }),
      tcell("Ver.2.0", { align: AlignmentType.CENTER }),
      tcell("A（移動支援3層構造）", { color: C.red, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("3-2-3 在宅生活継続", { bold: true }),
      tcell("03_在宅生活支援＋04_包括", { fill: C.lblue }),
      tcell("Ver.2.0", { align: AlignmentType.CENTER }),
      tcell("A（家族介護者支援）", { color: C.red, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("3-2-4 質の確保・人材", { bold: true }),
      tcell("05_介護サービス_人材", { fill: C.lblue }),
      tcell("Ver.2.0", { align: AlignmentType.CENTER }),
      tcell("B（事業者連絡会照会含む）", { color: C.blue, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("3-2-5 包括ケア深化", { bold: true }),
      tcell("04_認知症_包括", { fill: C.lblue }),
      tcell("Ver.2.0", { align: AlignmentType.CENTER }),
      tcell("A（体制強化根拠）", { color: C.red, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("3-2-6 認知症施策", { bold: true, color: CH[6].main }),
      tcell("04_認知症_包括＋06_第10期反映", { fill: CH[6].sub }),
      tcell("Ver.2.0", { align: AlignmentType.CENTER }),
      tcell("A（基本法対応・最重要）", { color: C.red, align: AlignmentType.CENTER, bold: true }),
    ]}),
  ],
});
S.push(linkTable);

S.push(placeholder("各事業の数値実績の確定値は、別添『川崎町_第9期実績一覧_町記入用.xlsx』（35事業以上を網羅）への町担当課ご記入の完了を経て、本素案Ver.2.0で本表に反映します。記入後は東京コンサルティング若山までご返送ください。"));

// 3-3
S.push(section(3, 3, "評価指標の達成状況"));
S.push(placeholder("第9期計画で設定された評価指標（KPI）の達成状況は、町担当課での集計を経て第1回策定委員会（令和8年8月中旬）で報告し、本素案Ver.2.0に反映します。"));

// 3-4
S.push(section(3, 4, "第10期に向けた課題"));
S.push(p("キックオフ会議における自治体課題の確認及び実績データ分析の結果、第10期計画では以下の課題への対応が重要となります。"));

const issuesTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("No.", { width: 8, fill: CH[3].main }),
      thcell("課題", { width: 30, fill: CH[3].main }),
      thcell("背景・現状", { width: 62, fill: CH[3].main }),
    ]}),
    new TableRow({ children: [
      tcell("①", { align: AlignmentType.CENTER, bold: true }),
      tcell("後期高齢者の急増", { bold: true }),
      tcell("75歳以上が4年で+10.9%（1,510→1,675人）。認定率上昇・介護需要増の継続"),
    ]}),
    new TableRow({ children: [
      tcell("②", { align: AlignmentType.CENTER, bold: true }),
      tcell("町外施設依存と地域偏在", { bold: true }),
      tcell("町内施設縮小・町外住所地特例24人。施設は役場周辺に偏在し地域格差あり"),
    ]}),
    new TableRow({ children: [
      tcell("③", { align: AlignmentType.CENTER, bold: true }),
      tcell("移動支援の再構築", { bold: true }),
      tcell("R7.3でタクシー助成終了。社協・NPO移送＋デマンドバスへ移行も周知が課題"),
    ]}),
    new TableRow({ children: [
      tcell("④", { align: AlignmentType.CENTER, bold: true }),
      tcell("認知症基本法対応", { bold: true }),
      tcell("R6.1施行。認知症施策推進計画の独立章化・本人意見反映・KPI設定が必須"),
    ]}),
    new TableRow({ children: [
      tcell("⑤", { align: AlignmentType.CENTER, bold: true }),
      tcell("8050問題・老老介護", { bold: true }),
      tcell("独居・夫婦のみに加え、未婚の子と高齢親の同居世帯での介護負担が課題"),
    ]}),
    new TableRow({ children: [
      tcell("⑥", { align: AlignmentType.CENTER, bold: true }),
      tcell("地域包括支援センターの体制強化", { bold: true }),
      tcell("保健師3＋認定調査1の4名体制。認定調査1名で負担大・ケアプラン外部委託あり"),
    ]}),
    new TableRow({ children: [
      tcell("⑦", { align: AlignmentType.CENTER, bold: true }),
      tcell("3計画整合（地域福祉・障害・介護）", { bold: true }),
      tcell("同時策定（地福・障害はジャパン総研）。移送・重層的支援等の論点整理が必要"),
    ]}),
    new TableRow({ children: [
      tcell("⑧", { align: AlignmentType.CENTER, bold: true }),
      tcell("広域医療連携", { bold: true }),
      tcell("みやぎ県南中核病院・刈田綜合病院との連携の計画明記"),
    ]}),
  ],
});
S.push(issuesTable);

// ===========================================================
// 第4章 計画の基本理念と基本目標
// ===========================================================
S.push(...chapterTitle(4));

// 4-1
S.push(section(4, 1, "基本理念（第9期踏襲）"));

S.push(p("キックオフ会議における町方針「計画内容は現状の目標・方針について第9期計画を踏襲する形とし、国の方針を新たに入れていくことを優先する」を踏まえ、本計画では第9期計画の基本理念を継承します。"));

S.push(subsection("基本理念（案）", 4));
S.push(new Paragraph({
  spacing: { before: 240, after: 240, line: 360 },
  alignment: AlignmentType.CENTER,
  shading: { type: ShadingType.SOLID, fill: CH[4].sub },
  border: {
    top: { style: BorderStyle.DOUBLE, size: 18, color: CH[4].main },
    bottom: { style: BorderStyle.DOUBLE, size: 18, color: CH[4].main },
  },
  children: [text("住民が住み慣れた地域で安心して暮らせるまちづくり", { size: 32, bold: true, color: CH[4].main })],
}));

S.push(subsection("サブタイトル（第10期で新たに付加）", 4));
S.push(new Paragraph({
  spacing: { before: 200, after: 200, line: 320 },
  alignment: AlignmentType.CENTER,
  children: [text("〜 認知症になっても誰もが自分らしく暮らせる地域共生社会の実現 〜", { size: 26, italics: true, color: CH[4].main })],
}));

S.push(p("第10期計画では、令和6年1月に施行された認知症基本法を踏まえ、認知症があっても誰もが地域で自分らしく暮らせる共生社会の実現を強調するサブタイトルを新設します。"));

S.push(placeholder("基本理念の本文・サブタイトル案は、第1回策定委員会（令和8年8月中旬）でご審議いただき、確定します。"));

// 4-2
S.push(section(4, 2, "基本目標"));

S.push(p("第9期計画の基本目標体系を継承し、第10期では認知症基本法対応として基本目標6を新設します。"));

const goalTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("基本目標", { width: 22, fill: CH[4].main }),
      thcell("内容", { width: 78, fill: CH[4].main }),
    ]}),
    new TableRow({ children: [
      tcell("1．健康づくりと\n　介護予防の推進", { bold: true, fill: C.lorange }),
      tcell("健康寿命の延伸・フレイル予防・通いの場の充実・口腔と栄養の支援を推進し、高齢者がいつまでも元気に生活できる地域づくりを進めます。"),
    ]}),
    new TableRow({ children: [
      tcell("2．高齢者が安心して\n　暮らせる仕組みづくり", { bold: true, fill: C.lorange }),
      tcell("見守り・地域支え合い・移動支援・住まいの確保等、独居高齢者・高齢者世帯・老老介護世帯が安心して暮らせる仕組みを整えます。"),
    ]}),
    new TableRow({ children: [
      tcell("3．在宅生活の継続支援", { bold: true, fill: C.lorange }),
      tcell("在宅医療・介護連携、家族介護者支援、緊急時対応、住宅改修等により、住み慣れた自宅・地域での生活継続を支援します。"),
    ]}),
    new TableRow({ children: [
      tcell("4．介護サービスの\n　質の確保と提供体制", { bold: true, fill: C.lorange }),
      tcell("介護サービスの質の確保、町外施設利用も含めた供給体制の整備、介護人材の確保・育成を推進します。"),
    ]}),
    new TableRow({ children: [
      tcell("5．地域包括ケアシステム\n　の深化", { bold: true, fill: C.lorange }),
      tcell("地域包括支援センターを中心に、医療・介護・予防・住まい・生活支援を一体的に提供する地域包括ケアシステムをさらに深化させます。"),
    ]}),
    new TableRow({ children: [
      tcell("6．認知症施策の\n　総合的推進【新設】", { bold: true, fill: CH[6].sub, color: CH[6].main }),
      tcell("認知症基本法に基づき、認知症の人本人の意思を尊重し、認知症の人もそうでない人もともに暮らせる共生社会の実現を目指します（第6章で独立章として位置付け）。"),
    ]}),
  ],
});
S.push(goalTable);

// 4-3
S.push(section(4, 3, "計画の重点ポイント（4つの重点）"));

S.push(p("第10期計画では、上記基本目標を達成するため、以下の4つの重点ポイントを設定します。これらは、川崎町固有の課題（後期高齢者の急増・町外施設依存・移動支援の再構築・認知症基本法対応・8050問題）に直接対応するものです。"));

const focusTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("重点", { width: 8, fill: CH[4].main }),
      thcell("テーマ", { width: 26, fill: CH[4].main }),
      thcell("対応する基本目標・関連施策", { width: 66, fill: CH[4].main }),
    ]}),
    new TableRow({ children: [
      tcell("重点1", { align: AlignmentType.CENTER, bold: true, fill: C.lorange }),
      tcell("認知症基本法対応\nと共生社会の実現", { bold: true }),
      tcell("基本目標6を新設し、独立章（第6章）として位置付け。本人意見反映・チームオレンジ整備・認知症カフェ充実"),
    ]}),
    new TableRow({ children: [
      tcell("重点2", { align: AlignmentType.CENTER, bold: true, fill: C.lorange }),
      tcell("移動支援の再構築\nと地域格差是正", { bold: true }),
      tcell("基本目標2。タクシー助成終了後の社協・NPO移送＋デマンドバスの周知強化、施設偏在への対応"),
    ]}),
    new TableRow({ children: [
      tcell("重点3", { align: AlignmentType.CENTER, bold: true, fill: C.lorange }),
      tcell("地域包括ケアの深化\nと家族介護者支援", { bold: true }),
      tcell("基本目標3・5。8050問題・老老介護への対応、介護者レスパイト、包括センター体制強化"),
    ]}),
    new TableRow({ children: [
      tcell("重点4", { align: AlignmentType.CENTER, bold: true, fill: C.lorange }),
      tcell("広域連携と\nサービス供給体制", { bold: true }),
      tcell("基本目標4。みやぎ県南中核病院等との広域医療連携、住所地特例を含めた施設供給見込量"),
    ]}),
  ],
});
S.push(focusTable);

// 4-4
S.push(section(4, 4, "計画の体系図"));

S.push(p("本計画の基本理念・基本目標・施策の体系は以下のとおりです。"));

const taikeiTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      new TableCell({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 100, after: 100 },
          children: [text("【基本理念】住民が住み慣れた地域で安心して暮らせるまちづくり", { size: 22, bold: true, color: C.white })],
        })],
        shading: { type: ShadingType.SOLID, fill: CH[4].main },
      }),
    ]}),
    new TableRow({ children: [
      new TableCell({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 80, after: 80 },
          children: [text("〜 認知症になっても誰もが自分らしく暮らせる地域共生社会の実現 〜", { size: 20, italics: true, color: CH[4].main })],
        })],
        shading: { type: ShadingType.SOLID, fill: CH[4].sub },
      }),
    ]}),
    new TableRow({ children: [
      new TableCell({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 80, after: 80 },
          children: [text("↓", { size: 24, bold: true, color: C.gray })],
        })],
      }),
    ]}),
    new TableRow({ children: [
      new TableCell({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 80, after: 80 },
          children: [text("【6つの基本目標】", { size: 22, bold: true, color: C.navy })],
        })],
        shading: { type: ShadingType.SOLID, fill: C.lblue },
      }),
    ]}),
  ],
});
S.push(taikeiTable);

S.push(spacer());

const goalDetailTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      tcell("1", { align: AlignmentType.CENTER, bold: true, fill: CH[4].main, color: C.white, size: 22 }),
      tcell("健康づくりと介護予防の推進", { bold: true, color: CH[4].main }),
      tcell("→ 第5章 5-1", { align: AlignmentType.CENTER, color: C.gray }),
    ]}),
    new TableRow({ children: [
      tcell("2", { align: AlignmentType.CENTER, bold: true, fill: CH[4].main, color: C.white, size: 22 }),
      tcell("高齢者が安心して暮らせる仕組みづくり", { bold: true, color: CH[4].main }),
      tcell("→ 第5章 5-2", { align: AlignmentType.CENTER, color: C.gray }),
    ]}),
    new TableRow({ children: [
      tcell("3", { align: AlignmentType.CENTER, bold: true, fill: CH[4].main, color: C.white, size: 22 }),
      tcell("在宅生活の継続支援", { bold: true, color: CH[4].main }),
      tcell("→ 第5章 5-3", { align: AlignmentType.CENTER, color: C.gray }),
    ]}),
    new TableRow({ children: [
      tcell("4", { align: AlignmentType.CENTER, bold: true, fill: CH[4].main, color: C.white, size: 22 }),
      tcell("介護サービスの質の確保と提供体制", { bold: true, color: CH[4].main }),
      tcell("→ 第5章 5-4", { align: AlignmentType.CENTER, color: C.gray }),
    ]}),
    new TableRow({ children: [
      tcell("5", { align: AlignmentType.CENTER, bold: true, fill: CH[4].main, color: C.white, size: 22 }),
      tcell("地域包括ケアシステムの深化", { bold: true, color: CH[4].main }),
      tcell("→ 第5章 5-5", { align: AlignmentType.CENTER, color: C.gray }),
    ]}),
    new TableRow({ children: [
      tcell("6", { align: AlignmentType.CENTER, bold: true, fill: CH[6].main, color: C.white, size: 22 }),
      tcell("認知症施策の総合的推進（新設・独立章）", { bold: true, color: CH[6].main }),
      tcell("→ 第6章", { align: AlignmentType.CENTER, color: CH[6].main, bold: true }),
    ]}),
  ],
});
S.push(goalDetailTable);

// ----- 一旦保存 -----
module.exports = { S };
console.log("Part1 (Ch1-4) ready. Paragraphs/blocks:", S.length);
