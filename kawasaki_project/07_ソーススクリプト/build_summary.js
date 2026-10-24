/**
 * B-2: 川崎町第10期計画 概要版
 *
 * 計画素案v1.5（43頁）を委員事前読み込み用に約18頁に要約
 * 章の要点・図表抜粋・川崎町固有論点を凝縮
 *
 * 構成（18頁想定）:
 *  表紙 (p1)
 *  目次・本計画について (p2)
 *  第1章 計画の策定にあたって（要点） (p3)
 *  第2章 川崎町の高齢者を取り巻く現状（図表中心） (p4-6)
 *  第3章 第9期計画の取組実績と評価 (p7-8)
 *  第4章 計画の基本理念と基本目標 (p9-10)
 *  第5章 施策の展開 (p11-13)
 *  第6章 認知症施策推進計画 (p14-15)
 *  第7章 介護保険サービス見込量と保険料 (p16-17)
 *  第8章 計画の推進体制と評価 (p18)
 *  別添資料一覧
 */
const fs = require('fs');
const H = require('./plan_helpers');
const {
  C, CH, FONT,
  text, p, section, subsection,
  bullet, numItem, placeholder, fact, source,
  tcell, thcell, kvTable, spacer, image, caption,
  chapterTitle,
  getCurrentChapter, setCurrentChapter,
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, convertInchesToTwip,
} = H;

const S = [];

// ===========================================================
// 概要版用ヘルパー
// ===========================================================
// コンパクトな章扉（半ページ）
function chapterCompact(num, name, color) {
  setCurrentChapter(num);
  return [
    new Paragraph({
      spacing: { before: 0, after: 0, line: 240 },
      pageBreakBefore: num > 1,
      children: [text("", { size: 16 })],
    }),
    new Paragraph({
      spacing: { before: 240, after: 240, line: 360 },
      alignment: AlignmentType.CENTER,
      border: {
        top: { style: BorderStyle.DOUBLE, size: 18, color: color },
        bottom: { style: BorderStyle.DOUBLE, size: 18, color: color },
      },
      shading: { type: ShadingType.SOLID, fill: CH[num].sub },
      children: [
        text(`第${num}章　`, { size: 28, bold: true, color: color }),
        text(name, { size: 28, bold: true, color: color }),
      ],
    }),
  ];
}

// 要点ボックス
function pointBox(title, content) {
  return [
    new Paragraph({
      spacing: { before: 200, after: 0 },
      shading: { type: ShadingType.SOLID, fill: C.lblue },
      border: { top: { style: BorderStyle.SINGLE, size: 12, color: C.navy } },
      children: [
        text("◆ " + title, { size: 22, bold: true, color: C.navy }),
      ],
    }),
    new Paragraph({
      spacing: { before: 0, after: 200 },
      shading: { type: ShadingType.SOLID, fill: "FAFAFA" },
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
      children: [text("　" + content, { size: 19 })],
    }),
  ];
}

// 川崎町固有論点ハイライト
function townSpecific(title, content) {
  return new Paragraph({
    spacing: { before: 200, after: 200 },
    shading: { type: ShadingType.SOLID, fill: C.lorange },
    border: {
      top: { style: BorderStyle.SINGLE, size: 12, color: C.orange },
      bottom: { style: BorderStyle.SINGLE, size: 12, color: C.orange },
    },
    children: [
      text("【川崎町固有】 ", { size: 20, bold: true, color: C.orange }),
      text(title + "\n　", { size: 20, bold: true, color: C.navy }),
      text(content, { size: 19 }),
    ],
  });
}

// ===========================================================
// 表紙
// ===========================================================
S.push(new Paragraph({
  spacing: { before: 1800, after: 240 },
  alignment: AlignmentType.CENTER,
  children: [text("川崎町高齢者保健福祉計画", { size: 28, bold: true, color: C.navy })],
}));
S.push(new Paragraph({
  spacing: { before: 0, after: 240 },
  alignment: AlignmentType.CENTER,
  children: [text("第10期介護保険事業計画", { size: 28, bold: true, color: C.navy })],
}));
S.push(new Paragraph({
  spacing: { before: 0, after: 480 },
  alignment: AlignmentType.CENTER,
  children: [text("認知症施策推進計画", { size: 28, bold: true, color: "9333B0" })],
}));
S.push(new Paragraph({
  spacing: { before: 360, after: 360 },
  alignment: AlignmentType.CENTER,
  border: {
    top: { style: BorderStyle.DOUBLE, size: 24, color: C.navy },
    bottom: { style: BorderStyle.DOUBLE, size: 24, color: C.navy },
  },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  children: [text("計　画　概　要　版", { size: 44, bold: true, color: C.navy })],
}));
S.push(new Paragraph({
  spacing: { before: 240, after: 240 },
  alignment: AlignmentType.CENTER,
  children: [text("〜 第1回策定委員会 事前配布資料 〜", { size: 22, italics: true, color: C.blue })],
}));

S.push(new Paragraph({
  spacing: { before: 200, after: 200 },
  alignment: AlignmentType.CENTER,
  children: [text("計画期間：令和9年度〜令和11年度（3年間）", { size: 22, bold: true })],
}));

S.push(new Paragraph({
  spacing: { before: 2400, after: 80 },
  alignment: AlignmentType.CENTER,
  children: [text("令和8年6月", { size: 22, color: C.gray })],
}));
S.push(new Paragraph({
  spacing: { before: 80, after: 80 },
  alignment: AlignmentType.CENTER,
  children: [text("川崎町", { size: 24, bold: true, color: C.navy })],
}));
S.push(new Paragraph({
  spacing: { before: 80, after: 80 },
  alignment: AlignmentType.CENTER,
  children: [text("（事務局：ビズアップ公共コンサルティング株式会社）", { size: 16, italics: true, color: C.gray })],
}));

// ===========================================================
// 目次・本計画について
// ===========================================================
S.push(new Paragraph({
  spacing: { before: 0, after: 0, line: 240 },
  pageBreakBefore: true,
  children: [text("", { size: 16 })],
}));

S.push(new Paragraph({
  spacing: { before: 240, after: 240, line: 320 },
  alignment: AlignmentType.CENTER,
  border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: C.navy } },
  children: [text("本概要版について／目次", { size: 28, bold: true, color: C.navy })],
}));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 本概要版の位置付け", { size: 22, bold: true, color: C.navy })],
}));

S.push(p("本概要版は、第1回策定委員会（令和8年8月中旬予定）の事前配布資料として、計画素案v1.5（全43頁）の要点を委員の皆様が短時間で把握できるよう約18頁に要約したものです。各章の要点と川崎町固有の論点を凝縮しています。詳細は別添『計画素案v1.5（全43頁）』をご参照ください。"));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 目次", { size: 22, bold: true, color: C.navy })],
}));

const tocItems = [
  ["第1章", "計画の策定にあたって", "3", "1F3864"],
  ["第2章", "川崎町の高齢者を取り巻く現状", "4", "375623"],
  ["第3章", "第9期計画の取組実績と評価", "7", "2E75B6"],
  ["第4章", "計画の基本理念と基本目標", "9", "C55A11"],
  ["第5章", "施策の展開", "11", "7030A0"],
  ["第6章", "認知症施策推進計画（独立章）", "14", "9333B0"],
  ["第7章", "介護保険サービス見込量と保険料", "16", "0F4F73"],
  ["第8章", "計画の推進体制と評価", "18", "404040"],
];

const tocTable = new Table({
  width: { size: 90, type: WidthType.PERCENTAGE },
  rows: tocItems.map(([ch, t, pg, color]) => new TableRow({
    children: [
      tcell(ch, { width: 10, bold: true, color: color, size: 22, align: AlignmentType.CENTER }),
      tcell(t, { width: 80, bold: true, color: color, size: 22 }),
      tcell(pg, { width: 10, bold: true, color: color, size: 22, align: AlignmentType.RIGHT }),
    ],
  })),
});
S.push(tocTable);

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lorange },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.orange } },
  children: [text("　▌ 本計画の3つの基本方針", { size: 22, bold: true, color: C.orange })],
}));

S.push(numItem("①", "3層整合の原則：国基本指針・宮城県高齢者保健福祉計画と整合を取りながら策定"));
S.push(numItem("②", "課題接続型：アンケートで把握された住民ニーズと、施策・KPIを論理的に接続"));
S.push(numItem("③", "認知症基本法対応：認知症基本法を踏まえ、認知症施策を独立章として位置付け"));

// ===========================================================
// 第1章 計画の策定にあたって
// ===========================================================
S.push(...chapterCompact(1, "計画の策定にあたって", "1F3864"));

S.push(...pointBox("計画の目的",
  "本計画は、川崎町における高齢者福祉と介護保険事業の総合的な指針です。介護保険法第117条に基づく市町村介護保険事業計画と、老人福祉法第20条の8に基づく市町村老人福祉計画を一体的に策定するとともに、令和6年1月施行の認知症基本法に対応した認知症施策推進計画を本計画に内包します。"));

S.push(...pointBox("計画期間",
  "令和9年度から令和11年度までの3年間。第10期介護保険事業計画として位置付けられます。本計画策定後、令和12年度に第11期計画策定を予定。"));

S.push(...pointBox("計画策定の体制",
  "川崎町高齢者保健福祉計画策定委員会（学識経験者・医療関係者・社会福祉協議会・民生委員・住民代表等で構成）の審議を経て策定します。令和8年8月の第1回策定委員会から始まり、令和9年2月の第4回策定委員会・町長答申までを予定。"));

S.push(townSpecific("3計画同時策定（高齢×地域福祉×障害）",
  "本計画と同時期に、地域福祉計画（ジャパン総研策定）・障害者計画（同）の3計画が並行策定されます。重層的支援体制整備・移動支援連動・データ共有等で相互整合を図りながら進めます。"));

// ===========================================================
// 第2章 川崎町の高齢者を取り巻く現状
// ===========================================================
S.push(...chapterCompact(2, "川崎町の高齢者を取り巻く現状", "375623"));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: "E2EFDA" },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "375623" } },
  children: [text("　▌ 主要数値（令和7年6月時点）", { size: 22, bold: true, color: "375623" })],
}));

const dataTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("項目", { width: 30, fill: "375623" }),
      thcell("数値", { width: 26, fill: "375623" }),
      thcell("特徴・比較", { width: 44, fill: "375623" }),
    ]}),
    new TableRow({ children: [
      tcell("総人口", { bold: true, fill: "E2EFDA" }),
      tcell("約8,000人（推計）", { bold: true, color: "375623", align: AlignmentType.CENTER }),
      tcell("人口減少傾向"),
    ]}),
    new TableRow({ children: [
      tcell("第1号被保険者数", { bold: true, fill: "E2EFDA" }),
      tcell("3,244人（R7.6）", { bold: true, color: "375623", align: AlignmentType.CENTER }),
      tcell("緩やかな減少傾向"),
    ]}),
    new TableRow({ children: [
      tcell("後期高齢者（75歳以上）", { bold: true, fill: "E2EFDA" }),
      tcell("1,675人（51.6%）", { bold: true, color: C.orange, align: AlignmentType.CENTER }),
      tcell("4年で+18.2%増（団塊世代到達）", { color: C.orange }),
    ]}),
    new TableRow({ children: [
      tcell("高齢化率", { bold: true, fill: "E2EFDA" }),
      tcell("41.4%", { bold: true, color: C.red, align: AlignmentType.CENTER }),
      tcell("県内35市町村中5位・県平均28.5%・全国29.1%", { color: C.red }),
    ]}),
    new TableRow({ children: [
      tcell("要介護認定者", { bold: true, fill: "E2EFDA" }),
      tcell("約470人（推計）", { bold: true, color: "375623", align: AlignmentType.CENTER }),
      tcell("認定率約14%（全国平均約19%より低い）"),
    ]}),
    new TableRow({ children: [
      tcell("サービス受給者", { bold: true, fill: "E2EFDA" }),
      tcell("466人（R7.6）", { bold: true, color: "375623", align: AlignmentType.CENTER }),
      tcell("居宅276・地密56・施設134"),
    ]}),
    new TableRow({ children: [
      tcell("住所地特例（町外施設）", { bold: true, fill: C.lorange }),
      tcell("24人（R7.6）", { bold: true, color: C.orange, align: AlignmentType.CENTER }),
      tcell("町外依存度比較的高い", { color: C.orange }),
    ]}),
    new TableRow({ children: [
      tcell("年間給付費", { bold: true, fill: "E2EFDA" }),
      tcell("9.6億円（R3）", { bold: true, color: "375623", align: AlignmentType.CENTER }),
      tcell("施設46.0%・居宅38.1%・地密15.9%"),
    ]}),
    new TableRow({ children: [
      tcell("第1号保険料（基準月額）", { bold: true, fill: "E2EFDA" }),
      tcell("第8期6,380→第9期6,500円", { bold: true, color: "375623", align: AlignmentType.CENTER }),
      tcell("上昇率+1.9%（全国平均より抑制的）"),
    ]}),
  ],
});
S.push(dataTable);

// 高齢化率図表埋込
S.push(new Paragraph({
  spacing: { before: 240, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: "E2EFDA" },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "375623" } },
  children: [text("　▌ 高齢化率比較（川崎町：県内5位）", { size: 22, bold: true, color: "375623" })],
}));

S.push(image("/home/claude/kawasaki_work/chart_aging.png", { width: 480, height: 300 }));
S.push(caption("図1　高齢化率の比較（川崎町41.4%・宮城県28.5%・全国29.1%）"));

S.push(townSpecific("川崎町の人口・世帯特性",
  "中山間部の地理的特性・人口減少・独居高齢者・高齢者世帯の増加。介護人材不足、本町取り巻く課題は多岐。第9期計画策定から3年を経過し、令和7年3月の高齢者外出タクシー助成事業の終了、介護施設の役場周辺偏在化、認知症基本法施行への対応など、新たな政策課題が生じています。"));

// 受給者構成図表
S.push(new Paragraph({
  spacing: { before: 240, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: "E2EFDA" },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "375623" } },
  children: [text("　▌ サービス受給者の構成（R7.6時点）", { size: 22, bold: true, color: "375623" })],
}));

S.push(image("/home/claude/kawasaki_work/chart_recipient.png", { width: 480, height: 300 }));
S.push(caption("図2　サービス受給者の区分別構成（466人）"));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: "E2EFDA" },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "375623" } },
  children: [text("　▌ 川崎町の地域資源", { size: 22, bold: true, color: "375623" })],
}));

S.push(bullet("国保川崎病院（町内唯一の医療拠点）／みやぎ県南中核病院・刈田綜合病院との広域連携"));
S.push(bullet("地域包括支援センター（社協運営・保健師3+認定調査1=4名）"));
S.push(bullet("ユニバーサルサポーター制度（14種別約400名）— 川崎町独自"));
S.push(bullet("認知症サポーター累計550名・キャラバンメイト73名"));
S.push(bullet("移動支援3層構造（社協NPO移送＋デマンドバス＋町民バス）"));
S.push(bullet("独自施策：紙おむつ・エアコン購入支援（R7.10開始）・透析助成"));

// ===========================================================
// 第3章 第9期計画の取組実績と評価
// ===========================================================
S.push(...chapterCompact(3, "第9期計画の取組実績と評価", "2E75B6"));

S.push(...pointBox("第9期計画の総括",
  "第9期計画（令和6〜8年度）は、第8期の体系を踏襲しつつ地域包括ケアシステムの深化を目指し、5基本目標のもとに体系化されました。健康づくり・介護予防、見守り、在宅生活継続支援、地域連携を柱とする体系で進められました。"));

S.push(new Paragraph({
  spacing: { before: 240, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: "DAE3F3" },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "2E75B6" } },
  children: [text("　▌ 第9期主要施策の取組実績（基本目標別）", { size: 22, bold: true, color: "2E75B6" })],
}));

const phase9Table = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("基本目標", { width: 20, fill: "2E75B6" }),
      thcell("主な取組", { width: 50, fill: "2E75B6" }),
      thcell("第10期への課題", { width: 30, fill: "2E75B6" }),
    ]}),
    new TableRow({ children: [
      tcell("1. 健康・介護予防", { bold: true, fill: "DAE3F3" }),
      tcell("ユニバーサルサポーター制度14種別約400名による介護予防活動・通いの場運営・健診・栄養支援等"),
      tcell("後期高齢者急増（+18.2%）への対応、中山間地域での通いの場参加と移動支援連動"),
    ]}),
    new TableRow({ children: [
      tcell("2. 仕組みづくり", { bold: true, fill: "DAE3F3" }),
      tcell("ふれあいネットワーク145名による見守り、緊急通報装置設置、移動支援（R7.3タクシー助成終了→3層構造移行）"),
      tcell("3層移動支援の住民周知強化、所管調整、福祉施設の役場周辺偏在対応", { color: C.orange }),
    ]}),
    new TableRow({ children: [
      tcell("3. 在宅生活継続", { bold: true, fill: "DAE3F3" }),
      tcell("国保川崎病院＋広域医療連携、家族介護教室、レスパイト支援、住宅改修・福祉用具利用支援"),
      tcell("8050問題・老老介護対応、家族介護者支援強化、介護離職防止"),
    ]}),
    new TableRow({ children: [
      tcell("4. 質確保・人材", { bold: true, fill: "DAE3F3" }),
      tcell("町内事業所・施設サービス（特養68・老健66受給）、住所地特例24人で町外依存"),
      tcell("施設の地域偏在対応、住所地特例所在自治体把握、介護人材確保", { color: C.orange }),
    ]}),
    new TableRow({ children: [
      tcell("5. 包括ケア深化", { bold: true, fill: "DAE3F3" }),
      tcell("地域包括支援センター（保健師3+認定調査1）、生活支援コーディネーター25名、自立支援型地域ケア会議"),
      tcell("認定調査員の負担軽減、3計画同時策定での重層的支援体制整備"),
    ]}),
    new TableRow({ children: [
      tcell("6. 認知症施策\n（第10期新設）", { bold: true, fill: "EAD5F0", color: "9333B0" }),
      tcell("認知症サポーター累計550名、キャラバンメイト73名、認知症カフェ「喫茶みかん」、初期集中支援チーム", { color: "9333B0" }),
      tcell("基本法対応：チームオレンジ整備（未実施）、本人ミーティング新設、KPI3層構造化", { color: "9333B0", bold: true }),
    ]}),
  ],
});
S.push(phase9Table);

S.push(...pointBox("第10期に向けた5重点課題",
  "①認知症対応（基本法対応）、②移動支援3層構造の確立、③在宅生活継続支援強化（8050問題・老老介護）、④施設広域連携・人材確保、⑤広域医療連携の体制整備。これら5課題が第10期計画の基本目標・重点施策の根拠となります。"));

// ===========================================================
// 第4章 基本理念と基本目標
// ===========================================================
S.push(...chapterCompact(4, "計画の基本理念と基本目標", "C55A11"));

S.push(new Paragraph({
  spacing: { before: 240, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: "FCE4D6" },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "C55A11" } },
  children: [text("　▌ 基本理念（第9期から継承）", { size: 22, bold: true, color: "C55A11" })],
}));

S.push(new Paragraph({
  spacing: { before: 200, after: 200 },
  alignment: AlignmentType.CENTER,
  shading: { type: ShadingType.SOLID, fill: "FCE4D6" },
  border: {
    top: { style: BorderStyle.DOUBLE, size: 18, color: "C55A11" },
    bottom: { style: BorderStyle.DOUBLE, size: 18, color: "C55A11" },
  },
  children: [text("住民が住み慣れた地域で安心して暮らせるまちづくり", { size: 28, bold: true, color: "C55A11" })],
}));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: "EAD5F0" },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "9333B0" } },
  children: [text("　▌ サブタイトル（第10期で新たに付加・案・協議事項1）", { size: 22, bold: true, color: "9333B0" })],
}));

S.push(new Paragraph({
  spacing: { before: 200, after: 200 },
  alignment: AlignmentType.CENTER,
  shading: { type: ShadingType.SOLID, fill: "EAD5F0" },
  border: {
    top: { style: BorderStyle.SINGLE, size: 12, color: "9333B0" },
    bottom: { style: BorderStyle.SINGLE, size: 12, color: "9333B0" },
  },
  children: [text("〜 認知症になっても誰もが自分らしく暮らせる地域共生社会の実現 〜", { size: 22, italics: true, color: "9333B0" })],
}));

S.push(new Paragraph({
  spacing: { before: 240, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: "FCE4D6" },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "C55A11" } },
  children: [text("　▌ 基本目標（6目標体系）", { size: 22, bold: true, color: "C55A11" })],
}));

const goalsTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("No", { width: 6, fill: "C55A11" }),
      thcell("基本目標", { width: 30, fill: "C55A11" }),
      thcell("施策の方向性（概要）", { width: 44, fill: "C55A11" }),
      thcell("対応課題", { width: 12, fill: "C55A11" }),
      thcell("章", { width: 8, fill: "C55A11" }),
    ]}),
    new TableRow({ children: [
      tcell("1", { bold: true, fill: "FCE4D6", align: AlignmentType.CENTER }),
      tcell("健康づくり・介護予防の推進", { bold: true }),
      tcell("ユニバーサルサポーター制度を基盤に、後期高齢者を含む全世代の介護予防"),
      tcell("─", { align: AlignmentType.CENTER }),
      tcell("第5章", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("2", { bold: true, fill: "FCE4D6", align: AlignmentType.CENTER }),
      tcell("高齢者が安心して暮らせるまちづくり", { bold: true }),
      tcell("見守り・移動支援3層構造・住まいの確保、町独自支援（エアコン・紙おむつ等）"),
      tcell("②", { align: AlignmentType.CENTER, color: C.red, bold: true }),
      tcell("第5章", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("3", { bold: true, fill: "FCE4D6", align: AlignmentType.CENTER }),
      tcell("在宅生活継続の支援", { bold: true }),
      tcell("在宅医療・介護連携、広域医療連携、家族介護者支援、介護離職防止"),
      tcell("③⑤", { align: AlignmentType.CENTER, color: C.red, bold: true }),
      tcell("第5章", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("4", { bold: true, fill: "FCE4D6", align: AlignmentType.CENTER }),
      tcell("介護サービスの質確保と提供体制", { bold: true }),
      tcell("サービスの質確保、町外施設利用（住所地特例24人）、介護人材確保"),
      tcell("④", { align: AlignmentType.CENTER, color: C.red, bold: true }),
      tcell("第5章", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("5", { bold: true, fill: "FCE4D6", align: AlignmentType.CENTER }),
      tcell("地域包括ケアシステムの深化", { bold: true }),
      tcell("包括センター強化、生活支援体制整備、自立支援型地域ケア会議"),
      tcell("─", { align: AlignmentType.CENTER }),
      tcell("第5章", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("6", { bold: true, fill: "EAD5F0", align: AlignmentType.CENTER, color: "9333B0" }),
      tcell("認知症施策の総合的推進（新設）", { bold: true, color: "9333B0" }),
      tcell("認知症基本法対応・第6章独立章化・7基本的施策・重点5本柱", { color: "9333B0" }),
      tcell("①", { align: AlignmentType.CENTER, color: "9333B0", bold: true }),
      tcell("第6章", { align: AlignmentType.CENTER, color: "9333B0", bold: true }),
    ]}),
  ],
});
S.push(goalsTable);

S.push(townSpecific("第10期での新設・強化要素",
  "基本目標6（認知症施策の総合的推進）の新設＝認知症基本法対応、基本目標2の移動支援3層構造の整理、基本目標4の広域連携の明示、基本目標6 7基本的施策に対応した町施策体系（重点5本柱）の構築が第10期の特徴です。"));

// ===========================================================
// 第5章 施策の展開
// ===========================================================
S.push(...chapterCompact(5, "施策の展開（基本目標1〜5）", "7030A0"));

S.push(...pointBox("施策展開の考え方",
  "第5章では基本目標1〜5に対応する施策・主な事業・KPIを整理します（基本目標6認知症施策は第6章で独立章として記載）。各施策は「住民ニーズ→施策→主な事業→KPI」の課題接続型で設計しています。"));

S.push(new Paragraph({
  spacing: { before: 240, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: "E4D6F0" },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "7030A0" } },
  children: [text("　▌ 重点施策5項目（J-A〜J-E・課題接続型）", { size: 22, bold: true, color: "7030A0" })],
}));

const focusTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("No", { width: 6, fill: "7030A0" }),
      thcell("重点施策", { width: 22, fill: "7030A0" }),
      thcell("住民ニーズ（根拠）", { width: 25, fill: "7030A0" }),
      thcell("主な取組", { width: 35, fill: "7030A0" }),
      thcell("KPI候補", { width: 12, fill: "7030A0" }),
    ]}),
    new TableRow({ children: [
      tcell("J-A", { bold: true, align: AlignmentType.CENTER, fill: "EAD5F0", color: "9333B0" }),
      tcell("認知症施策の総合的推進", { bold: true, color: "9333B0" }),
      tcell("認知症相談窓口認知の低さ・基本法対応必要", { color: "9333B0" }),
      tcell("サポーター拡大・チームオレンジ新規整備・本人ミーティング新設・早期発見強化・医療連携深化", { color: "9333B0" }),
      tcell("3層KPI", { align: AlignmentType.CENTER, color: "9333B0", bold: true }),
    ]}),
    new TableRow({ children: [
      tcell("J-B", { bold: true, align: AlignmentType.CENTER, fill: "FCE4D6", color: C.orange }),
      tcell("移動支援3層構造の確立", { bold: true, color: C.orange }),
      tcell("タクシー助成終了後の3制度認知度低い・所管分散", { color: C.orange }),
      tcell("社協NPO移送・デマンドバス・町民バスの整理、住民周知の一元化、所管調整", { color: C.orange }),
      tcell("認知度70%", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("J-C", { bold: true, align: AlignmentType.CENTER, fill: "E4D6F0" }),
      tcell("家族介護者支援・介護離職防止", { bold: true }),
      tcell("介護者負担大・介護離職発生・8050問題"),
      tcell("家族介護教室・短期入所活用・8050問題対応・介護離職防止支援"),
      tcell("介護離職率減", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("J-D", { bold: true, align: AlignmentType.CENTER, fill: "E4D6F0" }),
      tcell("施設広域連携・人材確保", { bold: true }),
      tcell("住所地特例24人(R7.6)・町外依存・人材不足"),
      tcell("町外施設情報整理・住所地特例所在自治体把握・人材確保支援"),
      tcell("町外情報整備", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("J-E", { bold: true, align: AlignmentType.CENTER, fill: "E4D6F0" }),
      tcell("広域医療連携の体制整備", { bold: true }),
      tcell("町外医療機関への通院常態化"),
      tcell("みやぎ県南中核病院・刈田綜合病院連携・国保川崎病院のハブ化・救急体制"),
      tcell("連携件数増", { align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(focusTable);

S.push(townSpecific("J-Bの優先実行",
  "重点施策J-B移動支援3層構造の確立は、令和7年3月のタクシー助成終了を踏まえた最重要施策の一つ。所管が地域振興課・町民生活課・社協・NPOに分散しており、利用者目線での情報一元化と住民周知の強化が急務。第1回策定委員会の協議事項3として取り上げます。"));

// ===========================================================
// 第6章 認知症施策推進計画
// ===========================================================
S.push(...chapterCompact(6, "認知症施策推進計画（独立章）", "9333B0"));

S.push(new Paragraph({
  spacing: { before: 200, after: 200 },
  shading: { type: ShadingType.SOLID, fill: "EAD5F0" },
  border: {
    top: { style: BorderStyle.SINGLE, size: 12, color: "9333B0" },
    bottom: { style: BorderStyle.SINGLE, size: 12, color: "9333B0" },
  },
  children: [text("◆ 認知症基本法対応 ", { size: 22, bold: true, color: "9333B0" }),
    text("令和6年1月施行の認知症基本法第14条に基づき、市町村認知症施策推進計画を本計画第6章として独立章化します。本章は基本法第15条〜21条の7基本的施策に対応した町施策体系で構成します。", { size: 20 })],
}));

S.push(new Paragraph({
  spacing: { before: 240, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: "EAD5F0" },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "9333B0" } },
  children: [text("　▌ 重点施策5本柱（J-1〜J-5）", { size: 22, bold: true, color: "9333B0" })],
}));

const j5Table = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("No", { width: 6, fill: "9333B0" }),
      thcell("重点施策", { width: 30, fill: "9333B0" }),
      thcell("第9期からの位置付け", { width: 28, fill: "9333B0" }),
      thcell("第10期の取組", { width: 36, fill: "9333B0" }),
    ]}),
    new TableRow({ children: [
      tcell("J-1", { bold: true, align: AlignmentType.CENTER, fill: "EAD5F0", color: "9333B0" }),
      tcell("認知症サポーターの拡大と質向上", { bold: true }),
      tcell("継続（累計550名・キャラバンメイト73名）"),
      tcell("企業・学校サポーター拡大、ステップアップ研修"),
    ]}),
    new TableRow({ children: [
      tcell("J-2", { bold: true, align: AlignmentType.CENTER, fill: "EAD5F0", color: "9333B0" }),
      tcell("チームオレンジの整備（新規）", { bold: true, color: "9333B0" }),
      tcell("未整備", { color: C.red, bold: true }),
      tcell("新規整備、本人・家族支援を実装", { bold: true, color: "9333B0" }),
    ]}),
    new TableRow({ children: [
      tcell("J-3", { bold: true, align: AlignmentType.CENTER, fill: "EAD5F0", color: "9333B0" }),
      tcell("認知症カフェ・本人ミーティング", { bold: true }),
      tcell("「喫茶みかん」継続・本人M未実施"),
      tcell("カフェ継続＋本人ミーティング新設", { bold: true, color: "9333B0" }),
    ]}),
    new TableRow({ children: [
      tcell("J-4", { bold: true, align: AlignmentType.CENTER, fill: "EAD5F0", color: "9333B0" }),
      tcell("早期発見・早期対応の体制強化", { bold: true }),
      tcell("初期集中支援チーム稼働中"),
      tcell("もの忘れ相談継続、支援チーム強化"),
    ]}),
    new TableRow({ children: [
      tcell("J-5", { bold: true, align: AlignmentType.CENTER, fill: "EAD5F0", color: "9333B0" }),
      tcell("国保川崎病院との医療連携", { bold: true }),
      tcell("継続"),
      tcell("認知症診断・治療・初期段階フォローの地域中核"),
    ]}),
  ],
});
S.push(j5Table);

S.push(new Paragraph({
  spacing: { before: 240, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: "EAD5F0" },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "9333B0" } },
  children: [text("　▌ KPIの3層構造（プロセス・アウトプット・アウトカム）", { size: 22, bold: true, color: "9333B0" })],
}));

const kpiTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("層", { width: 18, fill: "9333B0" }),
      thcell("内容", { width: 38, fill: "9333B0" }),
      thcell("川崎町のKPI例", { width: 44, fill: "9333B0" }),
    ]}),
    new TableRow({ children: [
      tcell("プロセス\n（活動量）", { bold: true, fill: "EAD5F0", align: AlignmentType.CENTER, color: "9333B0" }),
      tcell("町・包括センター・地域での認知症施策の活動回数等"),
      tcell("サポーター養成講座開催回数、カフェ開催回数"),
    ]}),
    new TableRow({ children: [
      tcell("アウトプット\n（成果物）", { bold: true, fill: "EAD5F0", align: AlignmentType.CENTER, color: "9333B0" }),
      tcell("活動の直接的な成果物・参加者数等"),
      tcell("サポーター累計数(550名→目標)、チームオレンジ整備"),
    ]}),
    new TableRow({ children: [
      tcell("アウトカム\n（住民の変化）", { bold: true, fill: "EAD5F0", align: AlignmentType.CENTER, color: "9333B0" }),
      tcell("住民の意識・行動・QoLの変化", { color: "9333B0" }),
      tcell("認知症本人と家族の地域生活満足度（アンケート測定）", { color: "9333B0", bold: true }),
    ]}),
  ],
});
S.push(kpiTable);

// ===========================================================
// 第7章 介護保険サービス見込量と保険料
// ===========================================================
S.push(...chapterCompact(7, "介護保険サービス見込量と保険料", "0F4F73"));

S.push(...pointBox("第7章の構造",
  "本章はサービス見込量の推計（7-1）・介護給付費の見込み（7-2）・介護保険料の試算（7-3）の3節で構成。見込量は6ステップ算定、保険料は8ステップ算定で、川崎町固有の補正要素（住所地特例・移動支援3層構造影響・ユニバーサルサポーター効果等）を組み込みます。"));

S.push(new Paragraph({
  spacing: { before: 240, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: "D6E4F0" },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "0F4F73" } },
  children: [text("　▌ サービス見込量算定の6ステップ", { size: 22, bold: true, color: "0F4F73" })],
}));

S.push(numItem("Step1", "人口推計：社人研R5推計を基本に住民基本台帳との突合で補正"));
S.push(numItem("Step2", "認定者推計：性別・年齢階級別認定率を直近5年実績推移から算定"));
S.push(numItem("Step3", "利用率：国保連データからサービス種別別の利用率を算定"));
S.push(numItem("Step4", "1人当たり利用量：実績データから平均利用量を算定"));
S.push(numItem("Step5", "アンケート補正：潜在ニーズを反映"));
S.push(numItem("Step6", "見える化システム登録：仙南圏域比較で精緻化"));

S.push(new Paragraph({
  spacing: { before: 240, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: "D6E4F0" },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "0F4F73" } },
  children: [text("　▌ 保険料試算3パターン（協議事項5）", { size: 22, bold: true, color: "0F4F73" })],
}));

const patTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("パターン", { width: 12, fill: "0F4F73" }),
      thcell("基金取崩方針", { width: 22, fill: "0F4F73" }),
      thcell("特徴", { width: 44, fill: "0F4F73" }),
      thcell("月額イメージ", { width: 22, fill: "0F4F73" }),
    ]}),
    new TableRow({ children: [
      tcell("A", { bold: true, align: AlignmentType.CENTER, fill: "D6E4F0" }),
      tcell("取崩なし", { bold: true }),
      tcell("基金温存・第11期負担緩和・給付費増を保険料に直接反映"),
      tcell("最高水準", { align: AlignmentType.CENTER, color: C.red }),
    ]}),
    new TableRow({ children: [
      tcell("B", { bold: true, align: AlignmentType.CENTER, fill: C.lorange }),
      tcell("50%取崩", { bold: true }),
      tcell("給付費増を一部相殺・次期負担との均衡", { color: C.orange }),
      tcell("中位水準", { align: AlignmentType.CENTER, color: C.orange, bold: true }),
    ]}),
    new TableRow({ children: [
      tcell("C", { bold: true, align: AlignmentType.CENTER, fill: "E2EFDA" }),
      tcell("全額取崩", { bold: true }),
      tcell("住民負担最小化・第11期負担増のリスクあり"),
      tcell("最低水準", { align: AlignmentType.CENTER, color: C.green }),
    ]}),
  ],
});
S.push(patTable);

// 保険料推移図表
S.push(new Paragraph({
  spacing: { before: 240, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: "D6E4F0" },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "0F4F73" } },
  children: [text("　▌ 介護保険料の推移と第10期試算", { size: 22, bold: true, color: "0F4F73" })],
}));

S.push(image("/home/claude/kawasaki_work/chart_premium.png", { width: 480, height: 300 }));
S.push(caption("図3　介護保険料月額基準額の推移（第8期6,380円→第9期6,500円・第10期は3パターン試算）"));

S.push(townSpecific("所得段階区分の9→13段階見直し",
  "国推奨の13段階化を検討します。低所得層（第1〜3段階）約30%の負担軽減、高所得層への応分負担拡大が見込まれます。具体的影響額は第3回策定委員会（R9.1中旬）で精緻化し協議します。"));

// ===========================================================
// 第8章 計画の推進体制と評価
// ===========================================================
S.push(...chapterCompact(8, "計画の推進体制と評価", "404040"));

S.push(...pointBox("推進体制",
  "本計画は川崎町保健福祉課（介護保険係）を主担当課とし、町関係課（地域振興課・町民生活課・財政課等）と地域包括支援センター・国民健康保険川崎病院・川崎町社会福祉協議会・民生委員・ユニバーサルサポーター・広域連携先（みやぎ県南中核病院・刈田綜合病院等）の連携のもとに推進します。"));

S.push(...pointBox("PDCAサイクル",
  "毎年度のPlan-Do-Check-Actionサイクルで進捗確認。中間評価（令和10年度）と最終評価（令和11年度末）で計画見直し要否を検討。最終評価は第11期計画策定の引継ぎ材料とします。"));

S.push(...pointBox("KPI評価の3階層",
  "施策レベル（5基本目標×各重点施策）・事業レベル（個別事業の実施回数・対象者数等）・KPIレベル（数値目標の達成度）の3階層で評価。認知症施策は3層KPI（プロセス・アウトプット・アウトカム）。"));

S.push(...pointBox("情報公開と住民参画",
  "計画進捗状況は毎年度町ホームページで公表。中間評価・最終評価結果も公表。本委員会の後継機関（仮称：高齢者保健福祉計画推進委員会）で継続的に審議。"));

S.push(spacer());

S.push(new Paragraph({
  spacing: { before: 240, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: "E2E8F0" },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "404040" } },
  children: [text("　▌ 別添資料一覧（事前配布資料）", { size: 22, bold: true, color: "404040" })],
}));

S.push(bullet("計画素案v1.5（全43頁・8章構成）── 本概要版の元データ"));
S.push(bullet("第1回策定委員会資料（全12頁）── アンケート結果報告と協議事項"));
S.push(bullet("3層整合表（国基本指針→宮城県計画→川崎町計画）"));
S.push(bullet("第9期実績一覧（町記入用フォーマット）"));
S.push(bullet("アンケート集計分析テンプレート"));

S.push(spacer());

S.push(new Paragraph({
  spacing: { before: 240, after: 240 },
  alignment: AlignmentType.CENTER,
  border: {
    top: { style: BorderStyle.DOUBLE, size: 18, color: C.navy },
    bottom: { style: BorderStyle.DOUBLE, size: 18, color: C.navy },
  },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  children: [text("第1回策定委員会では、本概要版と別添資料をもとに、5つの協議事項についてご審議いただきます。委員の皆様のご意見を計画素案Ver.2.0に反映いたします。",
    { size: 22, bold: true, color: C.navy })],
}));

// ===========================================================
// ドキュメント生成
// ===========================================================
const doc = new Document({
  creator: "ビズアップ公共コンサルティング株式会社",
  title: "川崎町 第10期計画 概要版",
  description: "委員事前配布用 計画概要版",
  styles: {
    default: {
      document: { run: { font: FONT, size: 20 } },
    },
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({
            text: "川崎町第10期介護保険事業計画 概要版（委員事前配布資料）",
            font: FONT, size: 16, color: C.gray,
          })],
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "─ ", font: FONT, size: 16, color: C.gray }),
            new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: C.gray }),
            new TextRun({ text: " ─", font: FONT, size: 16, color: C.gray }),
          ],
        })],
      }),
    },
    children: S,
  }],
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/home/claude/kawasaki_work/川崎町_計画素案_概要版.docx", buffer);
  console.log("Build done. Blocks:", S.length);
}).catch(err => {
  console.error("Build error:", err);
  process.exit(1);
});
