/**
 * 川崎町_第1回策定委員会資料
 * 金ヶ崎町方式踏襲（10-12頁・統合版）
 *
 * 構成:
 *  表紙
 *  目次
 *  本日の会議について
 *  第1部 アンケート結果報告（4-7頁）
 *   1. 調査の概要
 *   2. 高齢者の心身・健康状況
 *   3. 外出・移動の状況
 *   4. 在宅介護・家族介護の状況
 *   5. 認知症に関する状況
 *   6. アンケートからみえる重点課題
 *  第2部 計画骨子案の協議事項（8-12頁）
 *   1. 計画の全体構成（章立て）
 *   2. 基本理念・基本目標6つ
 *   3. 重点施策5項目（基本目標1-5に対応）
 *   4. 認知症施策推進計画（第6章）
 *   5. 本日ご協議いただきたい事項
 *  本日の進行・次回スケジュール
 */
const fs = require('fs');
const H = require('./plan_helpers');
const {
  C, CH, FONT,
  text, p, section, subsection,
  bullet, numItem, placeholder, fact, source,
  tcell, thcell, kvTable, spacer, image, caption,
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, convertInchesToTwip,
} = H;
const { ImageRun } = require('docx');

const S = [];

// ===========================================================
// ヘルパー：協議事項ハイライト枠
// ===========================================================
function discussionItem(num, title, body) {
  return [
    new Paragraph({
      spacing: { before: 200, after: 60 },
      shading: { type: ShadingType.SOLID, fill: "FFF2CC" },
      border: {
        top: { style: BorderStyle.SINGLE, size: 24, color: C.orange },
        bottom: { style: BorderStyle.SINGLE, size: 4, color: C.orange },
      },
      children: [
        text(`【協議事項${num}】 `, { size: 22, bold: true, color: C.orange }),
        text(title, { size: 22, bold: true, color: C.navy }),
      ],
    }),
    new Paragraph({
      spacing: { before: 0, after: 200 },
      shading: { type: ShadingType.SOLID, fill: "FFF8E1" },
      border: {
        bottom: { style: BorderStyle.SINGLE, size: 12, color: C.orange },
      },
      children: [text("　" + body, { size: 20 })],
    }),
  ];
}

// 重点課題ボックス（赤系強調）
function priorityIssue(num, title, body) {
  return new Paragraph({
    spacing: { before: 160, after: 200 },
    shading: { type: ShadingType.SOLID, fill: "FFE4E4" },
    border: {
      top: { style: BorderStyle.SINGLE, size: 18, color: C.red },
      bottom: { style: BorderStyle.SINGLE, size: 18, color: C.red },
    },
    children: [
      text(`課題${num} `, { size: 22, bold: true, color: C.red }),
      text(title, { size: 22, bold: true, color: C.navy }),
      text("\n　" + body, { size: 19 }),
    ],
  });
}

// 章番号付き節タイトル
function partTitle(text_str, color) {
  color = color || C.navy;
  return [
    new Paragraph({
      spacing: { before: 0, after: 0, line: 240 },
      pageBreakBefore: true,
      children: [text("", { size: 16 })],
    }),
    new Paragraph({
      spacing: { before: 240, after: 280, line: 360 },
      alignment: AlignmentType.CENTER,
      border: {
        top: { style: BorderStyle.DOUBLE, size: 18, color: color },
        bottom: { style: BorderStyle.DOUBLE, size: 18, color: color },
      },
      shading: { type: ShadingType.SOLID, fill: C.lblue },
      children: [text(text_str, { size: 30, bold: true, color: color })],
    }),
  ];
}

function partSection(num, title) {
  return new Paragraph({
    spacing: { before: 280, after: 160, line: 320 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: C.navy } },
    children: [
      text(num + "　", { size: 24, bold: true, color: C.navy }),
      text(title, { size: 24, bold: true, color: C.navy }),
    ],
  });
}

// ===========================================================
// 表紙
// ===========================================================
S.push(new Paragraph({
  spacing: { before: 2400, after: 240 },
  alignment: AlignmentType.CENTER,
  children: [text("川崎町高齢者保健福祉計画・第10期介護保険事業計画", { size: 24, bold: true, color: C.navy })],
}));
S.push(new Paragraph({
  spacing: { before: 0, after: 0 },
  alignment: AlignmentType.CENTER,
  children: [text("認知症施策推進計画", { size: 24, bold: true, color: "9333B0" })],
}));
S.push(new Paragraph({
  spacing: { before: 480, after: 240 },
  alignment: AlignmentType.CENTER,
  border: {
    top: { style: BorderStyle.DOUBLE, size: 24, color: C.navy },
    bottom: { style: BorderStyle.DOUBLE, size: 24, color: C.navy },
  },
  children: [text("第1回 策定委員会 資料", { size: 48, bold: true, color: C.navy })],
}));
S.push(new Paragraph({
  spacing: { before: 240, after: 240 },
  alignment: AlignmentType.CENTER,
  children: [text("〜 アンケート結果報告 ／ 計画骨子案の協議事項 〜", { size: 26, bold: true, color: C.blue })],
}));

S.push(new Paragraph({
  spacing: { before: 2400, after: 80 },
  alignment: AlignmentType.CENTER,
  children: [text("令和8年　月　日（　）", { size: 22 })],
}));
S.push(new Paragraph({
  spacing: { before: 80, after: 80 },
  alignment: AlignmentType.CENTER,
  children: [text("川崎町 保健福祉課", { size: 22, bold: true, color: C.navy })],
}));
S.push(new Paragraph({
  spacing: { before: 80, after: 80 },
  alignment: AlignmentType.CENTER,
  children: [text("（事務局：ビズアップ公共コンサルティング株式会社）", { size: 18, italics: true, color: C.gray })],
}));

// ===========================================================
// 目次
// ===========================================================
S.push(new Paragraph({
  spacing: { before: 0, after: 0, line: 240 },
  pageBreakBefore: true,
  children: [text("", { size: 16 })],
}));
S.push(new Paragraph({
  spacing: { before: 360, after: 360, line: 320 },
  alignment: AlignmentType.CENTER,
  border: { bottom: { style: BorderStyle.SINGLE, size: 24, color: C.navy } },
  children: [text("目　次", { size: 32, bold: true, color: C.navy })],
}));

const toc = [
  ["", "本日の会議について", "1"],
  ["第1部", "アンケート結果報告", ""],
  ["", "　1．調査の概要", "3"],
  ["", "　2．高齢者の心身・健康の状況", "3"],
  ["", "　3．外出・移動の状況", "4"],
  ["", "　4．在宅介護・家族介護の状況", "5"],
  ["", "　5．認知症に関する状況", "5"],
  ["", "　6．アンケートからみえる重点課題", "6"],
  ["第2部", "計画骨子案の協議事項", ""],
  ["", "　1．計画の全体構成（章立て）", "7"],
  ["", "　2．基本理念・基本目標", "7"],
  ["", "　3．重点施策5項目", "8"],
  ["", "　4．認知症施策推進計画（第6章）", "9"],
  ["", "　5．本日ご協議いただきたい事項", "10"],
  ["", "本日の進行・次回スケジュール", "11"],
];

const tocTable = new Table({
  width: { size: 92, type: WidthType.PERCENTAGE },
  rows: toc.map(([part, t, pg]) => new TableRow({
    children: [
      tcell(part, { width: 12, bold: !!part, color: part ? C.blue : C.black, size: 22, align: AlignmentType.CENTER }),
      tcell(t, { width: 78, bold: !pg, color: pg ? C.black : C.navy, size: 22 }),
      tcell(pg, { width: 10, bold: !!pg, color: C.navy, size: 22, align: AlignmentType.RIGHT }),
    ],
  })),
});
S.push(tocTable);

// ===========================================================
// 本日の会議について
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
  children: [text("本日の会議について", { size: 28, bold: true, color: C.navy })],
}));

S.push(p("第1回策定委員会では、本年6月下旬に実施した町民アンケート調査の結果をご報告するとともに、第10期計画の骨子案（章立て・基本理念・重点施策の方向性）について委員の皆様にご協議いただきます。"));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 本日の進め方", { size: 22, bold: true, color: C.navy })],
}));

const flowTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("Part", { width: 10 }),
      thcell("内容", { width: 50 }),
      thcell("資料", { width: 28 }),
      thcell("時間目安", { width: 12 }),
    ]}),
    new TableRow({ children: [
      tcell("第1部", { bold: true, fill: C.lblue, align: AlignmentType.CENTER }),
      tcell("アンケート結果のご報告\n（一般高齢者1,000名・認定者300名）"),
      tcell("本資料p3-6"),
      tcell("25分", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("第2部", { bold: true, fill: C.lorange, align: AlignmentType.CENTER }),
      tcell("計画骨子案の協議\n（章立て・基本理念・重点施策・認知症章）", { color: C.orange }),
      tcell("本資料p7-10\n計画素案v1.5（別添）"),
      tcell("40分", { align: AlignmentType.CENTER, bold: true, color: C.orange }),
    ]}),
    new TableRow({ children: [
      tcell("第3部", { bold: true, fill: C.lgreen, align: AlignmentType.CENTER }),
      tcell("質疑応答・今後のスケジュール"),
      tcell("本資料p11"),
      tcell("15分", { align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(flowTable);

S.push(new Paragraph({
  spacing: { before: 240, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 本計画策定の3つの方針", { size: 22, bold: true, color: C.navy })],
}));

S.push(numItem("①", "3層整合の原則：国基本指針・宮城県高齢者保健福祉計画と整合を取りながら策定します。"));
S.push(numItem("②", "課題接続型：アンケートで把握された住民ニーズと、施策・KPIを論理的に接続します。"));
S.push(numItem("③", "認知症基本法対応：令和6年1月施行の認知症基本法を踏まえ、認知症施策を独立章として位置付けます。"));

S.push(new Paragraph({
  spacing: { before: 240, after: 240 },
  shading: { type: ShadingType.SOLID, fill: C.lorange },
  border: {
    top: { style: BorderStyle.SINGLE, size: 12, color: C.orange },
    bottom: { style: BorderStyle.SINGLE, size: 12, color: C.orange },
  },
  children: [text("●本日の最重要事項：第2部の協議事項5項目について、委員の皆様のご意見を頂戴し、計画素案Ver.2.0の確定に反映いたします。",
    { size: 22, bold: true, color: C.orange })],
}));

// ===========================================================
// 第1部 表紙
// ===========================================================
S.push(...partTitle("第1部　アンケート結果報告"));

S.push(p("令和8年6月下旬に町から発送した「川崎町 高齢者ニーズ調査」「要支援・要介護認定者調査」の集計結果をご報告いたします。本報告は単純集計・クロス集計・自由記述の3観点で構成され、第2部の協議事項の根拠データとなります。"));

S.push(fact("本資料の数値は集計確定値です。詳細は別添『アンケート集計分析シート』をご覧ください。"));

// ===========================================================
// 第1部-1 調査の概要
// ===========================================================
S.push(partSection("1", "調査の概要"));

const surveyOverviewTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("項目", { width: 18 }),
      thcell("一般高齢者ニーズ調査", { width: 41 }),
      thcell("要支援・要介護認定者調査", { width: 41 }),
    ]}),
    new TableRow({ children: [
      tcell("対象", { bold: true, fill: C.lblue }),
      tcell("65歳以上の高齢者", { align: AlignmentType.CENTER }),
      tcell("要支援・要介護認定を受けている方", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("発送数", { bold: true, fill: C.lblue }),
      tcell("1,000名", { align: AlignmentType.CENTER, bold: true, color: C.navy }),
      tcell("300名", { align: AlignmentType.CENTER, bold: true, color: C.navy }),
    ]}),
    new TableRow({ children: [
      tcell("回収数", { bold: true, fill: C.lblue }),
      tcell("【N=要記入】", { align: AlignmentType.CENTER, italics: true, color: C.orange }),
      tcell("【N=要記入】", { align: AlignmentType.CENTER, italics: true, color: C.orange }),
    ]}),
    new TableRow({ children: [
      tcell("回収率", { bold: true, fill: C.lblue }),
      tcell("【％=要記入】", { align: AlignmentType.CENTER, italics: true, color: C.orange }),
      tcell("【％=要記入】", { align: AlignmentType.CENTER, italics: true, color: C.orange }),
    ]}),
    new TableRow({ children: [
      tcell("調査期間", { bold: true, fill: C.lblue }),
      tcell("令和8年6月下旬〜7月末", { align: AlignmentType.CENTER }),
      tcell("令和8年6月下旬〜7月末", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("調査方法", { bold: true, fill: C.lblue }),
      tcell("郵送調査（無記名）", { align: AlignmentType.CENTER }),
      tcell("郵送調査（無記名）", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("設問数", { bold: true, fill: C.lblue }),
      tcell("19問（国標準16問+川崎町追加3問）", { align: AlignmentType.CENTER }),
      tcell("16問（国標準12問+川崎町追加4問）", { align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(surveyOverviewTable);

S.push(p("川崎町追加設問では、本町固有の論点（町外医療機関の利用・町独自支援制度の認知度・町外施設の利用状況・現在の生活で最も不安なこと等）を調査しました。"));

// ===========================================================
// 第1部-2 高齢者の心身・健康の状況（人口・高齢化率含む）
// ===========================================================
S.push(partSection("2", "高齢者の心身・健康の状況"));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 川崎町の高齢化（参考：保険者データ R7.6時点）", { size: 22, bold: true, color: C.navy })],
}));

// 図2-5：高齢化率比較
S.push(image("/home/claude/kawasaki_work/chart_aging.png", { width: 480, height: 300 }));
S.push(caption("図1　高齢化率比較（川崎町：県内5位・41.4%）"));

S.push(p("川崎町の高齢化率41.4%は宮城県平均（28.5%）の約1.45倍、全国平均（29.1%）の約1.42倍に達しています。県内35市町村中5位の水準です。"));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ アンケート結果（一般高齢者ニーズ調査）", { size: 22, bold: true, color: C.navy })],
}));

const healthTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("項目", { width: 50 }),
      thcell("該当割合", { width: 18 }),
      thcell("該当人数", { width: 12 }),
      thcell("リスク評価", { width: 20 }),
    ]}),
    new TableRow({ children: [
      tcell("健康状態が「あまり健康でない/健康でない」（Q1-1）", { bold: true }),
      tcell("【％=要記入】", { italics: true, color: C.orange, align: AlignmentType.CENTER }),
      tcell("【人=要】", { italics: true, color: C.orange, align: AlignmentType.CENTER }),
      tcell("注視", { color: C.orange, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("週に1回も外出しない（Q1-2）", { bold: true }),
      tcell("【％=要記入】", { italics: true, color: C.orange, align: AlignmentType.CENTER }),
      tcell("【人=要】", { italics: true, color: C.orange, align: AlignmentType.CENTER }),
      tcell("フレイルリスク", { color: C.red, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("通いの場に「参加していない」（Q1-4）", { bold: true }),
      tcell("【％=要記入】", { italics: true, color: C.orange, align: AlignmentType.CENTER }),
      tcell("【人=要】", { italics: true, color: C.orange, align: AlignmentType.CENTER }),
      tcell("社会参加リスク", { color: C.red, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("特定健診を「受診していない」", { bold: true }),
      tcell("【％=要記入】", { italics: true, color: C.orange, align: AlignmentType.CENTER }),
      tcell("【人=要】", { italics: true, color: C.orange, align: AlignmentType.CENTER }),
      tcell("早期発見機会の損失", { color: C.orange, align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(healthTable);
S.push(source("出典：川崎町第10期計画策定に係るアンケート調査（令和8年6月実施）"));

S.push(p("「通いの場参加していない」「外出頻度週1回未満」の層は、フレイル進行・要介護化のリスクが相対的に高く、第10期計画では介護予防事業の対象拡大が重要となります。"));

// ===========================================================
// 第1部-3 外出・移動の状況
// ===========================================================
S.push(partSection("3", "外出・移動の状況（川崎町固有論点）"));

S.push(p("川崎町では、令和7年3月に高齢者外出タクシー助成事業が終了し、現在は社協・NPOによる福祉移送サービス、デマンドバス、町民バスの3層構造に移行しています。本調査では、住民の交通手段・移動困難の実態と、町移動支援3制度の認知度を把握しました。"));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 主な交通手段と移動の困りごと", { size: 22, bold: true, color: C.navy })],
}));

const transportTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("項目", { width: 56 }),
      thcell("結果", { width: 18 }),
      thcell("年齢別傾向", { width: 26 }),
    ]}),
    new TableRow({ children: [
      tcell("交通手段1位「自家用車（自分で運転）」", { bold: true }),
      tcell("【％=要記入】", { italics: true, color: C.orange, align: AlignmentType.CENTER }),
      tcell("65-74歳に集中", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("交通手段2位「家族の車に同乗」", { bold: true }),
      tcell("【％=要記入】", { italics: true, color: C.orange, align: AlignmentType.CENTER }),
      tcell("75歳以上に多い", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("外出困難「公共交通機関が少ない」（Q2-1）", { bold: true }),
      tcell("【％=要記入】", { italics: true, color: C.orange, align: AlignmentType.CENTER }),
      tcell("中山間部住民で高い", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("外出困難「運転が不安・できない」", { bold: true }),
      tcell("【％=要記入】", { italics: true, color: C.orange, align: AlignmentType.CENTER }),
      tcell("80歳以上で増加", { align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(transportTable);

S.push(new Paragraph({
  spacing: { before: 240, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 町の移動支援3制度の認知度（Q2-2 川崎町追加設問）", { size: 22, bold: true, color: C.navy })],
}));

const awarenessTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("制度", { width: 40 }),
      thcell("認知度", { width: 16 }),
      thcell("所管", { width: 18 }),
      thcell("評価", { width: 26 }),
    ]}),
    new TableRow({ children: [
      tcell("町民バス", { bold: true }),
      tcell("【％=要】", { italics: true, color: C.orange, align: AlignmentType.CENTER }),
      tcell("地域振興課", { align: AlignmentType.CENTER }),
      tcell("【評価=要】", { italics: true, color: C.orange, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("デマンドバス", { bold: true }),
      tcell("【％=要】", { italics: true, color: C.orange, align: AlignmentType.CENTER }),
      tcell("町民生活課", { align: AlignmentType.CENTER }),
      tcell("【評価=要】", { italics: true, color: C.orange, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("社協・NPO福祉移送サービス", { bold: true }),
      tcell("【％=要】", { italics: true, color: C.orange, align: AlignmentType.CENTER }),
      tcell("社協・NPO", { align: AlignmentType.CENTER }),
      tcell("【評価=要】", { italics: true, color: C.orange, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("いずれも知らない", { bold: true, color: C.red }),
      tcell("【％=要】", { italics: true, color: C.red, align: AlignmentType.CENTER, bold: true }),
      tcell("─", { align: AlignmentType.CENTER }),
      tcell("周知強化必要", { color: C.red, bold: true, align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(awarenessTable);

S.push(p("「いずれの制度も知らない」層が一定数いる場合、3制度の住民周知強化と利用情報の一元化が課題となります。所管が分かれている現状の整理が必要です。"));

// ===========================================================
// 第1部-4 在宅介護・家族介護の状況
// ===========================================================
S.push(partSection("4", "在宅介護・家族介護の状況（認定者調査）"));

S.push(p("要支援・要介護認定者300名へのアンケートから、サービス利用・家族介護・町外施設利用の実態を把握しました。"));

const careTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("項目", { width: 48 }),
      thcell("結果", { width: 22 }),
      thcell("計画反映先", { width: 30 }),
    ]}),
    new TableRow({ children: [
      tcell("サービス満足度「あまり満足でない/不満」（Q1-2）", { bold: true }),
      tcell("【％=要記入】", { italics: true, color: C.orange, align: AlignmentType.CENTER }),
      tcell("第5章 5-4", { align: AlignmentType.CENTER, color: C.navy }),
    ]}),
    new TableRow({ children: [
      tcell("介護者の負担感「とても負担」（Q2-2）", { bold: true }),
      tcell("【％=要記入】", { italics: true, color: C.orange, align: AlignmentType.CENTER }),
      tcell("第5章 5-3", { align: AlignmentType.CENTER, color: C.navy }),
    ]}),
    new TableRow({ children: [
      tcell("介護のため「仕事を辞めた・転職した」（Q2-3）", { bold: true }),
      tcell("【％=要記入】", { italics: true, color: C.orange, align: AlignmentType.CENTER }),
      tcell("第5章 5-3 介護離職防止", { align: AlignmentType.CENTER, color: C.navy }),
    ]}),
    new TableRow({ children: [
      tcell("町外施設の利用「あり」（K-1 川崎町追加）", { bold: true, color: C.orange }),
      tcell("【％=要記入】", { italics: true, color: C.orange, align: AlignmentType.CENTER }),
      tcell("第7章 7-1 / 第5章 5-4", { align: AlignmentType.CENTER, color: C.navy }),
    ]}),
    new TableRow({ children: [
      tcell("町外施設の所在地内訳（K-2 川崎町追加）", { bold: true, color: C.orange }),
      tcell("【柴田/大河原/仙台等=要】", { italics: true, color: C.orange, align: AlignmentType.CENTER }),
      tcell("広域連携の検討", { align: AlignmentType.CENTER, color: C.navy }),
    ]}),
    new TableRow({ children: [
      tcell("現在の最大の不安「介護者の負担」（K-4 川崎町追加）", { bold: true, color: C.orange }),
      tcell("【％=要記入】", { italics: true, color: C.orange, align: AlignmentType.CENTER }),
      tcell("第5章 5-3 第6章", { align: AlignmentType.CENTER, color: C.navy }),
    ]}),
  ],
});
S.push(careTable);

S.push(p("住所地特例該当者24名（R7.6時点）の所在地内訳をアンケートで把握することで、町外施設依存の実態と広域連携の必要性が明確になります。"));

// ===========================================================
// 第1部-5 認知症に関する状況
// ===========================================================
S.push(partSection("5", "認知症に関する状況（基本法対応の根拠データ）"));

S.push(new Paragraph({
  spacing: { before: 200, after: 200 },
  shading: { type: ShadingType.SOLID, fill: "EAD5F0" },
  border: { top: { style: BorderStyle.SINGLE, size: 12, color: "9333B0" }, bottom: { style: BorderStyle.SINGLE, size: 12, color: "9333B0" } },
  children: [text("　令和6年1月施行の認知症基本法第14条に基づき、市町村認知症施策推進計画の策定が努力義務化されました。アンケートで認知症に関する町民の不安・意識を把握し、第6章独立章の根拠データとします。", { size: 20, color: "9333B0" })],
}));

const dementiaTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("項目", { width: 50, fill: "9333B0" }),
      thcell("結果", { width: 20, fill: "9333B0" }),
      thcell("第6章への反映", { width: 30, fill: "9333B0" }),
    ]}),
    new TableRow({ children: [
      tcell("認知症の不安「自分自身がなることが不安」（Q3-1）", { bold: true }),
      tcell("【％=要記入】", { italics: true, color: C.orange, align: AlignmentType.CENTER }),
      tcell("6-3 KPI「認知症本人と家族の生活満足度」", { color: "9333B0" }),
    ]}),
    new TableRow({ children: [
      tcell("認知症の不安「家族がなることが不安」", { bold: true }),
      tcell("【％=要記入】", { italics: true, color: C.orange, align: AlignmentType.CENTER }),
      tcell("家族支援J-2 チームオレンジ", { color: "9333B0" }),
    ]}),
    new TableRow({ children: [
      tcell("認知症の相談窓口「知っている窓口はない」（Q3-2）", { bold: true, color: C.red }),
      tcell("【％=要記入】", { italics: true, color: C.red, align: AlignmentType.CENTER, bold: true }),
      tcell("J-4 早期発見・体制強化", { color: "9333B0", bold: true }),
    ]}),
    new TableRow({ children: [
      tcell("認知症サポーター養成講座を「受講したい」（Q3-3）", { bold: true }),
      tcell("【％=要記入】", { italics: true, color: C.orange, align: AlignmentType.CENTER }),
      tcell("J-1 サポーター拡大", { color: "9333B0" }),
    ]}),
    new TableRow({ children: [
      tcell("【認定者調査】認知症の人としての社会参加希望", { bold: true }),
      tcell("【％=要記入】", { italics: true, color: C.orange, align: AlignmentType.CENTER }),
      tcell("J-3 本人ミーティング新設", { color: "9333B0" }),
    ]}),
  ],
});
S.push(dementiaTable);

S.push(p("「相談窓口を知らない」層が多い場合、地域包括支援センター・認知症カフェ「喫茶みかん」・国保川崎病院の周知強化が重点課題となります。"));

// ===========================================================
// 第1部-6 重点課題
// ===========================================================
S.push(partSection("6", "アンケートからみえる重点課題"));

S.push(p("アンケート結果と本町固有の状況（高齢化率41.4%・移動支援3層構造移行・住所地特例24人・認知症基本法対応等）を踏まえ、第10期計画で対応すべき重点課題を以下のとおり整理しました。これらの課題は第2部の基本目標・重点施策の根拠となります。"));

S.push(priorityIssue("①",
  "認知症対応の総合的推進（基本法対応）",
  "認知症の相談窓口の認知度が低く、町民の不安が大きい。基本法対応として独立章化し、サポーター拡大・チームオレンジ整備・本人ミーティング新設等を重点化する。"));

S.push(priorityIssue("②",
  "移動支援3層構造の確立と住民周知",
  "タクシー助成終了後の3制度（町民バス・デマンドバス・社協NPO移送）の認知度が低く、所管が分散している。利用情報の一元化と周知強化が急務。"));

S.push(priorityIssue("③",
  "在宅介護継続の支援強化（家族介護者・8050問題対応）",
  "介護者負担が大きく、介護離職も発生。未婚の子と高齢親の同居（8050問題）、配偶者間の老老介護への対応として、家族介護者支援とレスパイト体制を強化。"));

S.push(priorityIssue("④",
  "施設サービスの広域連携と人材確保",
  "住所地特例24人（R7.6）が示すとおり町外施設依存が常態化。広域連携と町外施設情報の整理、深刻化する介護人材不足への対応が必要。"));

S.push(priorityIssue("⑤",
  "町外医療機関との広域医療連携",
  "みやぎ県南中核病院・刈田綜合病院等への通院・連携が日常化。在宅医療と広域医療をつなぐ体制整備が課題。"));

// ===========================================================
// 第2部 表紙
// ===========================================================
S.push(...partTitle("第2部　計画骨子案の協議事項"));

S.push(p("第1部で把握した重点課題と本町固有の状況を踏まえ、第10期計画の骨子案（章立て・基本理念・重点施策・認知症章）の方向性を以下に提示します。委員の皆様には、各協議事項についてご意見を頂戴したく存じます。"));

S.push(fact("本骨子案は計画素案v1.5（別添43頁）を委員会向けに集約したものです。詳細は別添素案をご参照ください。"));

// ===========================================================
// 第2部-1 計画の全体構成（章立て）
// ===========================================================
S.push(partSection("1", "計画の全体構成（章立て）"));

S.push(p("第9期計画の体系を基本的に踏襲しつつ、認知症基本法対応として基本目標6（認知症施策の総合的推進）を新設し、第6章として独立章化します。"));

const structureTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("章", { width: 6 }),
      thcell("章タイトル", { width: 38 }),
      thcell("第9期との差異・特徴", { width: 32 }),
      thcell("対応する課題", { width: 24 }),
    ]}),
    new TableRow({ children: [
      tcell("1", { bold: true, align: AlignmentType.CENTER, fill: "DAE3F3", color: "1F3864" }),
      tcell("計画の策定にあたって", { color: "1F3864", bold: true }),
      tcell("3つの基本方針（3層整合・課題接続・基本法対応）を新設"),
      tcell("─"),
    ]}),
    new TableRow({ children: [
      tcell("2", { bold: true, align: AlignmentType.CENTER, fill: "E2EFDA", color: "375623" }),
      tcell("川崎町の高齢者を取り巻く現状", { color: "375623", bold: true }),
      tcell("図表6種を新規埋込・高齢化率41.4%等"),
      tcell("─"),
    ]}),
    new TableRow({ children: [
      tcell("3", { bold: true, align: AlignmentType.CENTER, fill: "DAE3F3", color: "2E75B6" }),
      tcell("第9期計画の取組実績と評価", { color: "2E75B6", bold: true }),
      tcell("6目標別の実績整理（独立章化された認知症含む）"),
      tcell("─"),
    ]}),
    new TableRow({ children: [
      tcell("4", { bold: true, align: AlignmentType.CENTER, fill: "FCE4D6", color: "C55A11" }),
      tcell("計画の基本理念と基本目標", { color: "C55A11", bold: true }),
      tcell("基本目標6（認知症）を新設、計6目標", { color: C.orange, bold: true }),
      tcell("─"),
    ]}),
    new TableRow({ children: [
      tcell("5", { bold: true, align: AlignmentType.CENTER, fill: "E4D6F0", color: "7030A0" }),
      tcell("施策の展開（基本目標1〜5）", { color: "7030A0", bold: true }),
      tcell("基本目標1〜5に対応した施策・主な事業・KPI"),
      tcell("課題②③④⑤"),
    ]}),
    new TableRow({ children: [
      tcell("6", { bold: true, align: AlignmentType.CENTER, fill: "EAD5F0", color: "9333B0" }),
      tcell("認知症施策推進計画（独立章）", { color: "9333B0", bold: true }),
      tcell("基本法対応の新設・7基本的施策・重点5本柱", { color: "9333B0", bold: true }),
      tcell("課題①", { color: "9333B0", bold: true }),
    ]}),
    new TableRow({ children: [
      tcell("7", { bold: true, align: AlignmentType.CENTER, fill: "D6E4F0", color: "0F4F73" }),
      tcell("介護保険サービス見込量と保険料", { color: "0F4F73", bold: true }),
      tcell("見込量6ステップ・保険料8ステップ・3パターン試算"),
      tcell("─"),
    ]}),
    new TableRow({ children: [
      tcell("8", { bold: true, align: AlignmentType.CENTER, fill: "E2E8F0", color: "404040" }),
      tcell("計画の推進体制と評価", { color: "404040", bold: true }),
      tcell("PDCAサイクル・8章独立化"),
      tcell("─"),
    ]}),
  ],
});
S.push(structureTable);

// ===========================================================
// 第2部-2 基本理念・基本目標
// ===========================================================
S.push(partSection("2", "基本理念・基本目標"));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 基本理念（第9期踏襲）", { size: 22, bold: true, color: C.navy })],
}));

S.push(new Paragraph({
  spacing: { before: 200, after: 200 },
  alignment: AlignmentType.CENTER,
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: {
    top: { style: BorderStyle.DOUBLE, size: 18, color: C.navy },
    bottom: { style: BorderStyle.DOUBLE, size: 18, color: C.navy },
  },
  children: [text("住民が住み慣れた地域で安心して暮らせるまちづくり", { size: 28, bold: true, color: C.navy })],
}));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: "EAD5F0" },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "9333B0" } },
  children: [text("　▌ サブタイトル（第10期で新たに付加・案）", { size: 22, bold: true, color: "9333B0" })],
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
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 基本目標（第10期・全6目標）", { size: 22, bold: true, color: C.navy })],
}));

const goalTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("No", { width: 8 }),
      thcell("基本目標", { width: 38 }),
      thcell("関連課題", { width: 14 }),
      thcell("対応章", { width: 16 }),
      thcell("第9期からの変化", { width: 24 }),
    ]}),
    new TableRow({ children: [
      tcell("1", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("健康づくりと介護予防の推進", { bold: true }),
      tcell("ー", { align: AlignmentType.CENTER }),
      tcell("第5章 5-1", { align: AlignmentType.CENTER }),
      tcell("継承（ユニバーサルサポーター強化）"),
    ]}),
    new TableRow({ children: [
      tcell("2", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("高齢者が安心して暮らせるまちづくり", { bold: true }),
      tcell("課題②", { color: C.red, bold: true, align: AlignmentType.CENTER }),
      tcell("第5章 5-2", { align: AlignmentType.CENTER }),
      tcell("継承（移動支援3層構造を新規)"),
    ]}),
    new TableRow({ children: [
      tcell("3", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("在宅生活継続の支援", { bold: true }),
      tcell("課題③⑤", { color: C.red, bold: true, align: AlignmentType.CENTER }),
      tcell("第5章 5-3", { align: AlignmentType.CENTER }),
      tcell("継承（家族介護者支援強化）"),
    ]}),
    new TableRow({ children: [
      tcell("4", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("介護サービスの質確保と提供体制", { bold: true }),
      tcell("課題④", { color: C.red, bold: true, align: AlignmentType.CENTER }),
      tcell("第5章 5-4", { align: AlignmentType.CENTER }),
      tcell("継承（広域連携の明示）"),
    ]}),
    new TableRow({ children: [
      tcell("5", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("地域包括ケアシステムの深化", { bold: true }),
      tcell("ー", { align: AlignmentType.CENTER }),
      tcell("第5章 5-5", { align: AlignmentType.CENTER }),
      tcell("継承（包括センター体制強化）"),
    ]}),
    new TableRow({ children: [
      tcell("6", { bold: true, align: AlignmentType.CENTER, fill: "EAD5F0", color: "9333B0" }),
      tcell("認知症施策の総合的推進", { bold: true, color: "9333B0" }),
      tcell("課題①", { color: "9333B0", bold: true, align: AlignmentType.CENTER }),
      tcell("第6章 独立", { align: AlignmentType.CENTER, color: "9333B0", bold: true }),
      tcell("新設（基本法対応）", { color: "9333B0", bold: true }),
    ]}),
  ],
});
S.push(goalTable);

// ===========================================================
// 第2部-3 重点施策5項目
// ===========================================================
S.push(partSection("3", "重点施策5項目（課題接続型）"));

S.push(p("第10期計画では、アンケートから把握した5重点課題に対応する重点施策を以下のとおり整理します。各施策は「住民ニーズ→施策→KPI」の3層構造で設計しています。"));

const policyTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("No", { width: 6 }),
      thcell("重点施策", { width: 28 }),
      thcell("住民ニーズ（根拠）", { width: 28 }),
      thcell("主な取組", { width: 26 }),
      thcell("KPI候補", { width: 12 }),
    ]}),
    new TableRow({ children: [
      tcell("J-A", { bold: true, align: AlignmentType.CENTER, fill: "EAD5F0", color: "9333B0" }),
      tcell("認知症施策の総合的推進", { bold: true, color: "9333B0" }),
      tcell("認知症相談窓口を知らない層が一定数。基本法対応必要", { color: "9333B0" }),
      tcell("サポーター拡大・チームオレンジ整備・本人ミーティング新設・初期集中支援強化", { color: "9333B0" }),
      tcell("3層KPI設定", { align: AlignmentType.CENTER, color: "9333B0", bold: true }),
    ]}),
    new TableRow({ children: [
      tcell("J-B", { bold: true, align: AlignmentType.CENTER, fill: C.lorange, color: C.orange }),
      tcell("移動支援3層構造の確立", { bold: true, color: C.orange }),
      tcell("タクシー助成終了後の3制度の認知度低い・所管分散", { color: C.orange }),
      tcell("社協NPO移送・デマンドバス・町民バスの整理・住民周知の一元化", { color: C.orange }),
      tcell("認知度70%", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("J-C", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("家族介護者支援・介護離職防止", { bold: true }),
      tcell("介護者負担大・介護離職発生・8050問題", { align: AlignmentType.LEFT }),
      tcell("家族介護教室・レスパイト体制・介護離職防止支援", { align: AlignmentType.LEFT }),
      tcell("介護離職率減", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("J-D", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("施設広域連携・人材確保", { bold: true }),
      tcell("住所地特例24人(R7.6)・町外依存・人材不足", { align: AlignmentType.LEFT }),
      tcell("町外施設情報整理・住所地特例所在自治体把握・人材確保支援", { align: AlignmentType.LEFT }),
      tcell("町外施設情報整備", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("J-E", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("広域医療連携の体制整備", { bold: true }),
      tcell("町外医療機関への通院常態化", { align: AlignmentType.LEFT }),
      tcell("みやぎ県南中核病院・刈田綜合病院との連携・国保川崎病院ハブ化", { align: AlignmentType.LEFT }),
      tcell("連携件数増", { align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(policyTable);

// ===========================================================
// 第2部-4 認知症施策推進計画（第6章）
// ===========================================================
S.push(partSection("4", "認知症施策推進計画（第6章独立章）"));

S.push(new Paragraph({
  spacing: { before: 200, after: 200 },
  shading: { type: ShadingType.SOLID, fill: "EAD5F0" },
  border: { top: { style: BorderStyle.SINGLE, size: 12, color: "9333B0" }, bottom: { style: BorderStyle.SINGLE, size: 12, color: "9333B0" } },
  children: [text("　認知症基本法第14条に基づき、市町村認知症施策推進計画を本計画第6章として独立章化します。本章は基本法第15条〜21条の7基本的施策に対応した町施策体系で構成します。", { size: 20, color: "9333B0" })],
}));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: "EAD5F0" },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "9333B0" } },
  children: [text("　▌ 重点施策 5本柱", { size: 22, bold: true, color: "9333B0" })],
}));

const dementiaPolicy = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("No", { width: 6, fill: "9333B0" }),
      thcell("重点施策", { width: 34, fill: "9333B0" }),
      thcell("第9期からの位置付け", { width: 28, fill: "9333B0" }),
      thcell("第10期の取組", { width: 32, fill: "9333B0" }),
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
      tcell("未整備", { color: C.red }),
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
S.push(dementiaPolicy);

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: "EAD5F0" },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "9333B0" } },
  children: [text("　▌ KPIの3層構造（プロセス・アウトプット・アウトカム）", { size: 22, bold: true, color: "9333B0" })],
}));

const kpi3 = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("層", { width: 18, fill: "9333B0" }),
      thcell("内容", { width: 40, fill: "9333B0" }),
      thcell("川崎町のKPI例", { width: 42, fill: "9333B0" }),
    ]}),
    new TableRow({ children: [
      tcell("プロセス\n（活動量）", { bold: true, fill: "EAD5F0", align: AlignmentType.CENTER, color: "9333B0" }),
      tcell("町・包括センター・地域での認知症施策の活動回数等"),
      tcell("サポーター養成講座開催回数・カフェ開催回数"),
    ]}),
    new TableRow({ children: [
      tcell("アウトプット\n（成果物）", { bold: true, fill: "EAD5F0", align: AlignmentType.CENTER, color: "9333B0" }),
      tcell("活動の直接的な成果物・参加者数等"),
      tcell("サポーター累計数（550名→目標）、チームオレンジ整備"),
    ]}),
    new TableRow({ children: [
      tcell("アウトカム\n（住民の変化）", { bold: true, fill: "EAD5F0", align: AlignmentType.CENTER, color: "9333B0" }),
      tcell("住民の意識・行動・QoLの変化", { color: "9333B0" }),
      tcell("認知症本人と家族の地域生活満足度（アンケート測定）", { color: "9333B0", bold: true }),
    ]}),
  ],
});
S.push(kpi3);

// ===========================================================
// 第2部-5 本日ご協議いただきたい事項
// ===========================================================
S.push(partSection("5", "本日ご協議いただきたい事項（5項目）"));

S.push(p("以下の5項目について、委員の皆様のご意見を頂戴し、計画素案Ver.2.0に反映させていただきます。"));

S.push(...discussionItem("1",
  "基本理念のサブタイトル付加について",
  "第9期から継承する基本理念「住民が住み慣れた地域で安心して暮らせるまちづくり」に、認知症基本法対応のサブタイトル「〜認知症になっても誰もが自分らしく暮らせる地域共生社会の実現〜」を新たに付加することの是非。"));

S.push(...discussionItem("2",
  "基本目標6（認知症施策）の新設と第6章独立章化について",
  "認知症基本法対応として、第9期の5基本目標に「基本目標6：認知症施策の総合的推進」を追加し、第6章として独立章化する方向性の是非。第6章では、基本法第15条〜21条の7基本的施策に対応した町施策体系を構築します。"));

S.push(...discussionItem("3",
  "移動支援3層構造の整理について",
  "令和7年3月のタクシー助成終了後、社協・NPO福祉移送／デマンドバス／町民バスの3制度に移行しました。アンケートで認知度の低さが確認された場合、第10期計画では3制度の整理・住民周知の強化を重点施策（J-B）として明示する方向性の是非。所管調整（地域振興課・町民生活課・社協・NPO）の進め方も含めご意見をいただきます。"));

S.push(...discussionItem("4",
  "認知症施策の重点施策5本柱とKPI3層構造について",
  "認知症施策の重点5本柱（J-1サポーター拡大／J-2チームオレンジ整備・新規／J-3本人ミーティング新設／J-4早期発見・体制強化／J-5医療連携）の方向性。特にJ-2チームオレンジとJ-3本人ミーティングの新規取組の進め方、及びKPI3層構造（プロセス・アウトプット・アウトカム）の設定方針へのご意見をいただきます。"));

S.push(...discussionItem("5",
  "介護保険料試算3パターン（A/B/C）の検討方針について",
  "第10期保険料は、介護給付費準備基金の取崩額により3パターン（A：取崩なし／B：50%取崩／C：全額取崩）を試算する方向性です。確定値は介護給付費準備基金残高（R8.6時点・現在確認中）の確定後、第3回策定委員会（R9.1中旬）で協議します。本日は方向性へのご意見をいただきます。また、所得段階区分の9段階から13段階への見直しの検討方針もあわせてご意見をいただきます。"));

S.push(spacer());

S.push(new Paragraph({
  spacing: { before: 240, after: 240 },
  alignment: AlignmentType.CENTER,
  shading: { type: ShadingType.SOLID, fill: C.lgreen },
  border: {
    top: { style: BorderStyle.DOUBLE, size: 12, color: C.green },
    bottom: { style: BorderStyle.DOUBLE, size: 12, color: C.green },
  },
  children: [text("以上5項目について、忌憚のないご意見を頂戴できますと幸いです。", { size: 22, bold: true, color: C.green })],
}));

// ===========================================================
// 本日の進行・次回スケジュール
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
  children: [text("本日の進行・次回以降のスケジュール", { size: 28, bold: true, color: C.navy })],
}));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 第10期計画策定の今後のスケジュール", { size: 22, bold: true, color: C.navy })],
}));

const scheduleNext = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("時期", { width: 16 }),
      thcell("マイルストーン", { width: 26 }),
      thcell("主な議題・成果物", { width: 38 }),
      thcell("段階", { width: 20 }),
    ]}),
    new TableRow({ children: [
      tcell("R8.8\n中旬", { bold: true, fill: C.lgreen, align: AlignmentType.CENTER }),
      tcell("第1回策定委員会（本日）", { bold: true, color: C.green }),
      tcell("アンケート結果・骨子案協議事項5項目"),
      tcell("骨子協議", { align: AlignmentType.CENTER, color: C.green }),
    ]}),
    new TableRow({ children: [
      tcell("R8.9-10", { bold: true, fill: C.lblue, align: AlignmentType.CENTER }),
      tcell("計画素案Ver.2.0更新", { bold: true }),
      tcell("委員意見の反映、サービス見込量精緻化"),
      tcell("素案更新", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("R8.11", { bold: true, fill: C.lblue, align: AlignmentType.CENTER }),
      tcell("第2回策定委員会", { bold: true }),
      tcell("Ver.2.0素案審議・サービス見込量方向性確定"),
      tcell("素案審議", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("R8.12-R9.1", { bold: true, fill: C.lblue, align: AlignmentType.CENTER }),
      tcell("保険料試算精緻化", { bold: true }),
      tcell("給付費・地域支援事業費・3パターン精緻化"),
      tcell("試算精緻化", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("R9.1\n中旬", { bold: true, fill: C.lorange, align: AlignmentType.CENTER }),
      tcell("第3回策定委員会", { bold: true, color: C.orange }),
      tcell("保険料3パターン協議・パターン選定"),
      tcell("保険料協議", { align: AlignmentType.CENTER, color: C.orange }),
    ]}),
    new TableRow({ children: [
      tcell("R9.2", { bold: true, fill: C.lorange, align: AlignmentType.CENTER }),
      tcell("第4回策定委員会・町長答申", { bold: true, color: C.orange }),
      tcell("保険料基準額決定・最終素案承認"),
      tcell("基準額決定", { align: AlignmentType.CENTER, color: C.orange }),
    ]}),
    new TableRow({ children: [
      tcell("R9.2-3", { bold: true, fill: C.lgreen, align: AlignmentType.CENTER }),
      tcell("パブリックコメント・3月議会上程", { bold: true, color: C.green }),
      tcell("住民意見募集・条例改正案上程"),
      tcell("確定・公表", { align: AlignmentType.CENTER, color: C.green }),
    ]}),
  ],
});
S.push(scheduleNext);

S.push(new Paragraph({
  spacing: { before: 240, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 別添資料一覧", { size: 22, bold: true, color: C.navy })],
}));

const appendix = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("No", { width: 6 }),
      thcell("資料名", { width: 50 }),
      thcell("用途", { width: 44 }),
    ]}),
    new TableRow({ children: [
      tcell("①", { bold: true, align: AlignmentType.CENTER }),
      tcell("川崎町_計画書素案_v1.5_カラー統一版.pdf（43頁）", { bold: true, color: C.navy }),
      tcell("本骨子案の詳細版・参考資料"),
    ]}),
    new TableRow({ children: [
      tcell("②", { bold: true, align: AlignmentType.CENTER }),
      tcell("川崎町_アンケート集計分析テンプレート.xlsx", { color: C.navy }),
      tcell("アンケート結果の詳細集計表"),
    ]}),
    new TableRow({ children: [
      tcell("③", { bold: true, align: AlignmentType.CENTER }),
      tcell("川崎町_第9期実績一覧_町記入用.xlsx", { color: C.navy }),
      tcell("第9期実績の根拠データ"),
    ]}),
    new TableRow({ children: [
      tcell("④", { bold: true, align: AlignmentType.CENTER }),
      tcell("策定委員会 委員名簿", { color: C.navy }),
      tcell("委員一覧（町担当課作成）"),
    ]}),
    new TableRow({ children: [
      tcell("⑤", { bold: true, align: AlignmentType.CENTER }),
      tcell("第10期計画策定スケジュール", { color: C.navy }),
      tcell("全体スケジュール表"),
    ]}),
  ],
});
S.push(appendix);

S.push(spacer());
S.push(spacer());

S.push(new Paragraph({
  spacing: { before: 480, after: 240 },
  alignment: AlignmentType.CENTER,
  border: {
    top: { style: BorderStyle.DOUBLE, size: 18, color: C.navy },
    bottom: { style: BorderStyle.DOUBLE, size: 18, color: C.navy },
  },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  children: [text("本日は貴重なお時間を頂きありがとうございます。今後ともよろしくお願い申し上げます。",
    { size: 22, bold: true, color: C.navy })],
}));

// ===========================================================
// ドキュメント生成
// ===========================================================
const doc = new Document({
  creator: "ビズアップ公共コンサルティング株式会社",
  title: "川崎町 第1回策定委員会資料",
  description: "アンケート結果報告と計画骨子案の協議事項",
  styles: {
    default: {
      document: { run: { font: FONT, size: 21 } },
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
            text: "川崎町第10期介護保険事業計画　第1回策定委員会資料",
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
  fs.writeFileSync("/home/claude/kawasaki_work/川崎町_第1回策定委員会資料.docx", buffer);
  console.log("Build done. Blocks:", S.length);
}).catch(err => {
  console.error("Build error:", err);
  process.exit(1);
});
