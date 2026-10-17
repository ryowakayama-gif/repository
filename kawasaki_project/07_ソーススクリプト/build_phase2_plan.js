/**
 * Phase2作業計画詳細書
 * R8.8中旬(第1回策定委員会)〜R8.11(第2回策定委員会)の3.5ヶ月間の詳細作業計画
 * 
 * 構成:
 *  表紙
 *  目次
 *  1. Phase 2の全体像
 *  2. アンケート集計フロー（R8.8上旬）
 *  3. 委員意見反映フロー（R8.8末-9.9月初）
 *  4. 計画素案Ver.2.0更新フロー（R8.9月中-10月）
 *  5. 第2回策定委員会準備（R8.11上旬）
 *  6. リスク管理と対応策
 *  7. 週次マイルストーン管理
 */
const fs = require('fs');
const H = require('./plan_helpers');
const {
  C, CH, FONT,
  text, p, section, subsection,
  bullet, numItem, placeholder, fact, source,
  tcell, thcell, kvTable, spacer,
  chapterTitle, getCurrentChapter, setCurrentChapter,
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, convertInchesToTwip,
} = H;

const S = [];

// ===========================================================
// ヘルパー
// ===========================================================
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
      shading: { type: ShadingType.SOLID, fill: C.lblue },
      children: [
        text(`Section ${num}　`, { size: 22, bold: true, color: color }),
        text(name, { size: 28, bold: true, color: color }),
      ],
    }),
  ];
}

function infoBox(title, content, color) {
  color = color || C.navy;
  const fill = color === C.orange ? C.lorange : (color === C.green ? C.lgreen : C.lblue);
  return [
    new Paragraph({
      spacing: { before: 200, after: 0 },
      shading: { type: ShadingType.SOLID, fill: fill },
      border: { top: { style: BorderStyle.SINGLE, size: 12, color: color } },
      children: [
        text("◆ " + title, { size: 22, bold: true, color: color }),
      ],
    }),
    new Paragraph({
      spacing: { before: 0, after: 200 },
      shading: { type: ShadingType.SOLID, fill: "FAFAFA" },
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: color } },
      children: [text("　" + content, { size: 19 })],
    }),
  ];
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
  spacing: { before: 0, after: 480 },
  alignment: AlignmentType.CENTER,
  children: [text("第10期介護保険事業計画 策定業務", { size: 28, bold: true, color: C.navy })],
}));
S.push(new Paragraph({
  spacing: { before: 240, after: 240 },
  alignment: AlignmentType.CENTER,
  border: {
    top: { style: BorderStyle.DOUBLE, size: 24, color: C.navy },
    bottom: { style: BorderStyle.DOUBLE, size: 24, color: C.navy },
  },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  children: [text("Phase 2 作業計画 詳細書", { size: 40, bold: true, color: C.navy })],
}));
S.push(new Paragraph({
  spacing: { before: 240, after: 480 },
  alignment: AlignmentType.CENTER,
  children: [text("〜 R8.8中旬 第1回委員会 → R8.11 第2回委員会 〜", { size: 22, italics: true, color: C.blue })],
}));

S.push(new Paragraph({
  spacing: { before: 480, after: 100, line: 320 },
  alignment: AlignmentType.CENTER,
  shading: { type: ShadingType.SOLID, fill: C.lorange },
  border: {
    top: { style: BorderStyle.SINGLE, size: 12, color: C.orange },
    bottom: { style: BorderStyle.SINGLE, size: 12, color: C.orange },
  },
  children: [text("対象期間：令和8年8月中旬 〜 令和8年11月（約3.5ヶ月間）", { size: 22, bold: true, color: C.orange })],
}));

S.push(new Paragraph({
  spacing: { before: 2400, after: 80 },
  alignment: AlignmentType.RIGHT,
  children: [text("令和8年6月作成", { size: 22, color: C.gray })],
}));
S.push(new Paragraph({
  spacing: { before: 80, after: 80 },
  alignment: AlignmentType.RIGHT,
  children: [text("ビズアップ公共コンサルティング株式会社", { size: 22, bold: true, color: C.navy })],
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
  spacing: { before: 240, after: 240, line: 320 },
  alignment: AlignmentType.CENTER,
  border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: C.navy } },
  children: [text("目　次", { size: 28, bold: true, color: C.navy })],
}));

const toc = [
  ["Section 1", "Phase 2の全体像", "3"],
  ["Section 2", "アンケート集計フロー（R8.8上旬）", "5"],
  ["Section 3", "委員意見反映フロー（R8.8末〜R8.9月初）", "7"],
  ["Section 4", "計画素案Ver.2.0更新フロー（R8.9月中〜R8.10末）", "9"],
  ["Section 5", "第2回策定委員会準備（R8.11上旬）", "11"],
  ["Section 6", "リスク管理と対応策", "12"],
  ["Section 7", "週次マイルストーン管理", "13"],
];

const tocTable = new Table({
  width: { size: 90, type: WidthType.PERCENTAGE },
  rows: toc.map(([sec, t, pg]) => new TableRow({
    children: [
      tcell(sec, { width: 16, bold: true, color: C.blue, size: 22, align: AlignmentType.CENTER }),
      tcell(t, { width: 74, bold: true, color: C.navy, size: 22 }),
      tcell(pg, { width: 10, bold: true, color: C.navy, size: 22, align: AlignmentType.RIGHT }),
    ],
  })),
});
S.push(tocTable);

// ===========================================================
// Section 1: Phase 2の全体像
// ===========================================================
S.push(...chapterCompact(1, "Phase 2の全体像", C.navy));

S.push(...infoBox("Phase 2の目的",
  "Phase 1で完成した計画素案v1.5を、第1回策定委員会の委員意見とアンケート集計結果を反映して計画素案Ver.2.0に発展させ、第2回策定委員会で提示することがPhase 2の目的です。"));

S.push(...infoBox("Phase 2の3つの統合作業",
  "① アンケート結果(R8.7末回収)の集計と計画素案への反映、② 第1回委員会で出された委員意見の対応方針確定と反映、③ サービス見込量算定の精緻化（6ステップ・町外利用反映）。これら3つを並行作業で進めます。"));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ Phase 2 タイムライン", { size: 22, bold: true, color: C.navy })],
}));

const phase2Timeline = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("時期", { width: 14 }),
      thcell("主要マイルストーン", { width: 26 }),
      thcell("作業内容", { width: 44 }),
      thcell("成果物", { width: 16 }),
    ]}),
    new TableRow({ children: [
      tcell("R8.8\n上旬", { bold: true, fill: C.lblue, align: AlignmentType.CENTER }),
      tcell("アンケート集計フェーズ", { bold: true, color: C.navy }),
      tcell("R8.7末回収アンケートの単純集計・クロス集計・自由記述分類"),
      tcell("集計結果一式"),
    ]}),
    new TableRow({ children: [
      tcell("R8.8\n中旬", { bold: true, fill: C.lgreen, align: AlignmentType.CENTER }),
      tcell("第1回策定委員会", { bold: true, color: C.green }),
      tcell("計画素案v1.5・アンケート結果報告・5協議事項の協議"),
      tcell("委員会議事録"),
    ]}),
    new TableRow({ children: [
      tcell("R8.8\n下旬", { bold: true, fill: C.lblue, align: AlignmentType.CENTER }),
      tcell("意見集約・対応方針協議", { bold: true, color: C.navy }),
      tcell("委員意見の本シート転記・対応方針(採用/部分/保留)の3者協議"),
      tcell("意見管理シート"),
    ]}),
    new TableRow({ children: [
      tcell("R8.9\n上中", { bold: true, fill: C.lblue, align: AlignmentType.CENTER }),
      tcell("Ver.2.0素案作成", { bold: true, color: C.navy }),
      tcell("アンケート結果＋委員意見＋見込量精緻化を計画素案に反映"),
      tcell("Ver.2.0素案"),
    ]}),
    new TableRow({ children: [
      tcell("R8.10\n通月", { bold: true, fill: C.lblue, align: AlignmentType.CENTER }),
      tcell("Ver.2.0精査・Red Team", { bold: true, color: C.navy }),
      tcell("整合性検証・MECE検証・町担当課確認・国通知反映"),
      tcell("Ver.2.0完成版"),
    ]}),
    new TableRow({ children: [
      tcell("R8.11\n上旬", { bold: true, fill: C.lorange, align: AlignmentType.CENTER }),
      tcell("第2回委員会準備", { bold: true, color: C.orange }),
      tcell("委員会資料作成・想定問答更新・サービス見込量資料整備"),
      tcell("委員会資料"),
    ]}),
    new TableRow({ children: [
      tcell("R8.11\n中旬", { bold: true, fill: C.lgreen, align: AlignmentType.CENTER }),
      tcell("第2回策定委員会", { bold: true, color: C.green }),
      tcell("Ver.2.0素案審議・サービス見込量方向性確定・継続協議事項整理"),
      tcell("委員会議事録"),
    ]}),
  ],
});
S.push(phase2Timeline);

S.push(...infoBox("Phase 2の重要KPI",
  "（1）アンケート集計完了：R8.8末まで、（2）委員意見対応方針確定：R8.9上旬まで、（3）Ver.2.0素案完成：R8.10末まで、（4）第2回委員会開催：R8.11中旬。これら4つの中間マイルストーンを順守することがPhase 3（保険料試算精緻化）への円滑な接続条件です。",
  C.orange));

// ===========================================================
// Section 2: アンケート集計フロー
// ===========================================================
S.push(...chapterCompact(2, "アンケート集計フロー（R8.8上旬）", C.navy));

S.push(...infoBox("対象アンケート",
  "一般高齢者ニーズ調査(発送1,000名)と認定者調査(発送300名)の計1,300名分。R8.6下旬発送・R8.7末回収。本Phaseでは弊社で集計を実施し、結果を計画素案Ver.2.0に反映します。"));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 集計工程（4工程・約2週間）", { size: 22, bold: true, color: C.navy })],
}));

const collectFlow = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("工程", { width: 10 }),
      thcell("内容", { width: 28 }),
      thcell("作業", { width: 38 }),
      thcell("期間", { width: 10 }),
      thcell("出力", { width: 14 }),
    ]}),
    new TableRow({ children: [
      tcell("工程1", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("データ入力", { bold: true }),
      tcell("町から受領した回収アンケートをCSV/Excelに転記。回答漏れ・誤記入の検出"),
      tcell("3日", { align: AlignmentType.CENTER }),
      tcell("生データ"),
    ]}),
    new TableRow({ children: [
      tcell("工程2", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("単純集計", { bold: true }),
      tcell("一般高齢者19問・認定者16問の全設問について件数・構成比を算出"),
      tcell("3日", { align: AlignmentType.CENTER }),
      tcell("単純集計表"),
    ]}),
    new TableRow({ children: [
      tcell("工程3", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("クロス集計", { bold: true }),
      tcell("5本のクロス集計(地区x移動・年齢x外出・世帯x介護不安・要介護度x満足度・地区x認知症窓口)"),
      tcell("3日", { align: AlignmentType.CENTER }),
      tcell("クロス集計表"),
    ]}),
    new TableRow({ children: [
      tcell("工程4", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("自由記述分類", { bold: true }),
      tcell("5分類(移動/介護/見守り/サービス不足/医療連携)で自由記述を分類・件数集計"),
      tcell("3日", { align: AlignmentType.CENTER }),
      tcell("分類結果"),
    ]}),
  ],
});
S.push(collectFlow);

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 計画素案への反映先（32項目）", { size: 22, bold: true, color: C.navy })],
}));

S.push(p("アンケート集計結果は、別添「アンケート集計分析テンプレート」の05計画反映ガイドに従って、計画素案v1.5の32箇所のプレースホルダーに反映します。具体的な反映先は以下のとおりです。"));

S.push(bullet("第2章 2-2 高齢者の状況：健康状態・外出頻度・通いの場参加・受診状況等（7項目）"));
S.push(bullet("第3章 3-3 評価指標達成状況：KPI実績値の更新（5項目）"));
S.push(bullet("第5章 5-1〜5-5 重点施策のKPI設定根拠：住民ニーズの数値化（10項目）"));
S.push(bullet("第6章 6-3 認知症施策KPI：認知症本人と家族の地域生活満足度（4項目）"));
S.push(bullet("第7章 7-1 サービス見込量算定の補正：潜在ニーズの反映（6項目）"));

S.push(...infoBox("アンケート回収率対策",
  "目標回収率は一般高齢者60-70%・認定者70-80%。R8.7中旬時点で回収率が低い場合、町担当課と協議の上、督促はがき送付・包括センター窓口での回収補完・町広報誌再周知等を実施します。",
  C.orange));

// ===========================================================
// Section 3: 委員意見反映フロー
// ===========================================================
S.push(...chapterCompact(3, "委員意見反映フロー（R8.8末〜R8.9月初）", C.navy));

S.push(...infoBox("意見集約の前提",
  "第1回策定委員会で出された全意見を、別添「委員意見反映管理シート」に転記し、事務局(町担当課)・受託者(弊社)・必要に応じて担当事業者の3者協議で対応方針を確定します。"));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 意見反映の判断基準（4分類）", { size: 22, bold: true, color: C.navy })],
}));

const judgement = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("判断", { width: 14 }),
      thcell("該当条件", { width: 50 }),
      thcell("対応", { width: 36 }),
    ]}),
    new TableRow({ children: [
      tcell("反映採用", { bold: true, color: C.green, align: AlignmentType.CENTER, fill: C.lgreen }),
      tcell("法令・制度に合致／町方針と整合／実現可能性あり／住民ニーズと整合"),
      tcell("Ver.2.0に本文反映・KPI設定根拠として明記"),
    ]}),
    new TableRow({ children: [
      tcell("部分反映", { bold: true, color: C.orange, align: AlignmentType.CENTER, fill: C.lorange }),
      tcell("方向性は合うが具体策・程度に調整必要／既存施策の強化として反映可能"),
      tcell("意見の趣旨を活かし、表現を調整して反映"),
    ]}),
    new TableRow({ children: [
      tcell("保留検討", { bold: true, color: C.orange, align: AlignmentType.CENTER, fill: "FFE4B0" }),
      tcell("制度改正待ち／予算措置必要／町方針との整合確認必要／第2回委員会で再協議"),
      tcell("第2回委員会の継続協議事項として整理"),
    ]}),
    new TableRow({ children: [
      tcell("反映なし", { bold: true, color: C.red, align: AlignmentType.CENTER, fill: C.lred }),
      tcell("法令違反／町方針に反する／実現困難／個人的見解にとどまる"),
      tcell("理由を明記の上、計画には反映しない（議事録には記載）"),
    ]}),
  ],
});
S.push(judgement);

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 3者協議のスケジュール", { size: 22, bold: true, color: C.navy })],
}));

S.push(numItem("①", "R8.8末: 委員会議事録案作成・意見管理シートへの全意見転記"));
S.push(numItem("②", "R8.9初週: 第1回3者協議(オンライン90分)- 全意見の対応方針の一次案を確定"));
S.push(numItem("③", "R8.9第2週: 第2回3者協議(オンライン60分)- 保留検討事項の対応方針確定"));
S.push(numItem("④", "R8.9第3週: 最終確認協議(メール往復可)- Ver.2.0反映先(章節)の確定"));

S.push(...infoBox("協議事項5（保険料）の特殊性",
  "協議事項5(保険料試算3パターン)については、基金残高R8.6時点確定値の入手後に精緻化するため、第1回委員会では方向性協議のみ。具体的反映は第3回委員会(R9.1)向けの保険料試算精緻化作業(Phase 3)で対応します。",
  C.orange));

// ===========================================================
// Section 4: 計画素案Ver.2.0更新フロー
// ===========================================================
S.push(...chapterCompact(4, "計画素案Ver.2.0更新フロー（R8.9月中〜R8.10末）", C.navy));

S.push(...infoBox("Ver.2.0更新の3つの柱",
  "（1）アンケート結果反映（32項目）、（2）委員意見反映（協議事項1-5＋付帯意見）、（3）サービス見込量精緻化（6ステップ）。これら3つの柱を統合してVer.2.0として完成させます。"));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 各章の更新ポイント", { size: 22, bold: true, color: C.navy })],
}));

const chapterUpdates = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("章", { width: 6 }),
      thcell("更新箇所", { width: 40 }),
      thcell("反映ソース", { width: 36 }),
      thcell("優先度", { width: 18 }),
    ]}),
    new TableRow({ children: [
      tcell("第1章", { bold: true, align: AlignmentType.CENTER, fill: "DAE3F3", color: "1F3864" }),
      tcell("策定体制・スケジュール表(委員会開催実績反映)"),
      tcell("第1回委員会開催実績"),
      tcell("A", { color: C.red, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("第2章", { bold: true, align: AlignmentType.CENTER, fill: "E2EFDA", color: "375623" }),
      tcell("高齢者の生活実態・健康状態・社会参加(アンケート反映)"),
      tcell("アンケート Q1・Q2系"),
      tcell("A", { color: C.red, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("第3章", { bold: true, align: AlignmentType.CENTER, fill: "DAE3F3", color: "2E75B6" }),
      tcell("第9期実績数値の確定値置換(町担当ご記入完了後)"),
      tcell("第9期実績一覧記入完了データ"),
      tcell("A", { color: C.red, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("第4章", { bold: true, align: AlignmentType.CENTER, fill: "FCE4D6", color: "C55A11" }),
      tcell("基本理念サブタイトル・基本目標6新設の確定/修正"),
      tcell("委員意見 協議事項1・2"),
      tcell("A", { color: C.red, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("第5章", { bold: true, align: AlignmentType.CENTER, fill: "E4D6F0", color: "7030A0" }),
      tcell("重点施策J-B(移動)の所管調整反映・KPI数値設定"),
      tcell("委員意見 協議事項3・アンケート"),
      tcell("A", { color: C.red, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("第6章", { bold: true, align: AlignmentType.CENTER, fill: "EAD5F0", color: "9333B0" }),
      tcell("認知症施策J-1〜J-5の修正・3層KPI数値設定"),
      tcell("委員意見 協議事項4・アンケート"),
      tcell("A", { color: C.red, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("第7章", { bold: true, align: AlignmentType.CENTER, fill: "D6E4F0", color: "0F4F73" }),
      tcell("サービス見込量精緻化(6ステップ)・補正反映"),
      tcell("町実績データ・アンケート補正"),
      tcell("A", { color: C.red, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("第8章", { bold: true, align: AlignmentType.CENTER, fill: "E2E8F0", color: "404040" }),
      tcell("推進体制・PDCAサイクル(委員意見で調整可能)"),
      tcell("委員意見 付帯事項E"),
      tcell("B", { color: C.blue, bold: true, align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(chapterUpdates);

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ Ver.2.0品質確保のためのRed Team検証", { size: 22, bold: true, color: C.navy })],
}));

S.push(p("Ver.2.0完成前に、以下の観点でRed Team的に検証を実施します。"));

S.push(numItem("①", "3層整合チェック：国基本指針→宮城県計画→川崎町計画の整合性再確認"));
S.push(numItem("②", "MECE検証：6基本目標・各章のテーマが網羅性・排他性を満たすか確認"));
S.push(numItem("③", "課題接続検証：アンケート課題→重点施策→KPIの論理接続が明確か確認"));
S.push(numItem("④", "数値整合検証：本文中の数値・表・図表の数値が一致しているか確認"));
S.push(numItem("⑤", "国通知反映：R8夏以降の国通知・厚労省ガイドラインの反映漏れがないか確認"));
S.push(numItem("⑥", "町方針整合：町長挨拶・3計画(高齢×地域福祉×障害)同時策定との整合確認"));

// ===========================================================
// Section 5: 第2回策定委員会準備
// ===========================================================
S.push(...chapterCompact(5, "第2回策定委員会準備（R8.11上旬）", C.navy));

S.push(...infoBox("第2回委員会の位置付け",
  "第2回策定委員会(R8.11中旬予定)は、計画素案Ver.2.0を委員に提示し審議いただくことが主目的です。第1回での協議事項の反映状況確認、サービス見込量の方向性確定、第3回委員会(R9.1)に向けた保険料試算の準備状況報告も行います。"));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 第2回委員会の準備資料", { size: 22, bold: true, color: C.navy })],
}));

const meetingDocs = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("No", { width: 6 }),
      thcell("資料名", { width: 30 }),
      thcell("内容", { width: 44 }),
      thcell("作成期限", { width: 20 }),
    ]}),
    new TableRow({ children: [
      tcell("①", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("計画素案Ver.2.0", { bold: true }),
      tcell("更新版計画素案(全体)。Ver.1.5からの主要変更点を別途まとめる"),
      tcell("R8.10末"),
    ]}),
    new TableRow({ children: [
      tcell("②", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("Ver.1.5→Ver.2.0変更点一覧", { bold: true }),
      tcell("章節別の変更点・反映元(委員意見/アンケート/見込量精緻化)・変更理由"),
      tcell("R8.11初"),
    ]}),
    new TableRow({ children: [
      tcell("③", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("第2回委員会資料(議事資料)", { bold: true }),
      tcell("委員会次第・本日の協議事項・サービス見込量資料・概要版更新"),
      tcell("R8.11初"),
    ]}),
    new TableRow({ children: [
      tcell("④", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("アンケート結果報告(完全版)", { bold: true }),
      tcell("単純集計・クロス集計5本・自由記述・計画反映状況"),
      tcell("R8.10末"),
    ]}),
    new TableRow({ children: [
      tcell("⑤", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("第1回委員意見対応状況一覧", { bold: true }),
      tcell("意見管理シートから採用・部分・保留・不採用の状況を一覧化"),
      tcell("R8.11初"),
    ]}),
    new TableRow({ children: [
      tcell("⑥", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("想定問答集(更新版)", { bold: true }),
      tcell("第1回想定問答集を基礎に、Ver.2.0関連の新規Q&A追加"),
      tcell("R8.11初"),
    ]}),
    new TableRow({ children: [
      tcell("⑦", { bold: true, align: AlignmentType.CENTER, fill: C.lorange }),
      tcell("サービス見込量算定資料", { bold: true, color: C.orange }),
      tcell("6ステップで算定した見込量・補正要素・第10期暫定見込み", { color: C.orange }),
      tcell("R8.11初", { bold: true, color: C.orange }),
    ]}),
  ],
});
S.push(meetingDocs);

S.push(...infoBox("第2回委員会の協議事項想定",
  "（1）Ver.1.5→Ver.2.0変更点の確認、（2）サービス見込量(6ステップ)の方向性確定、（3）第1回継続協議事項の整理、（4）認知症施策の3層KPI設定値の方向性、（5）次回(R9.1中旬)第3回委員会に向けた準備状況。",
  C.orange));

// ===========================================================
// Section 6: リスク管理と対応策
// ===========================================================
S.push(...chapterCompact(6, "リスク管理と対応策", C.navy));

S.push(p("Phase 2で想定される主要リスクと対応策を整理します。リスク顕在化時は速やかに町担当課・弊社で協議し、計画スケジュール全体への影響を最小化します。"));

const riskTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("No", { width: 6 }),
      thcell("リスク", { width: 28 }),
      thcell("影響", { width: 22 }),
      thcell("対応策", { width: 32 }),
      thcell("優先度", { width: 12 }),
    ]}),
    new TableRow({ children: [
      tcell("R1", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("アンケート回収率の低迷(50%以下)", { bold: true }),
      tcell("計画反映データ不足", { color: C.red }),
      tcell("督促はがき送付・包括センター窓口回収・町広報誌再周知"),
      tcell("A", { color: C.red, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("R2", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("基金残高R8.6時点の確定遅延", { bold: true }),
      tcell("保険料試算精緻化の遅延", { color: C.red }),
      tcell("町担当課に再依頼・推計値で暫定試算・第3回委員会で確定"),
      tcell("A", { color: C.red, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("R3", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("第1回委員会で大きな方向転換意見", { bold: true }),
      tcell("素案大幅修正の必要性", { color: C.orange }),
      tcell("3者協議で慎重判断・必要に応じ第2回委員会前に追加打合せ"),
      tcell("B", { color: C.blue, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("R4", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("国通知(R8夏以降)の予想外内容", { bold: true }),
      tcell("計画素案構成の見直し", { color: C.orange }),
      tcell("通知発出後速やかに影響評価・委員会への報告"),
      tcell("A", { color: C.red, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("R5", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("町担当課のキャパシティ不足", { bold: true }),
      tcell("確認・記入の遅延", { color: C.orange }),
      tcell("弊社で代替検討・優先度の高い項目に絞った確認依頼"),
      tcell("B", { color: C.blue, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("R6", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("3計画(地福・障害)との不整合発生", { bold: true }),
      tcell("計画間の矛盾", { color: C.orange }),
      tcell("ジャパン総研との情報共有強化・3計画整合確認会議"),
      tcell("B", { color: C.blue, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("R7", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("見込量算定の見える化システム不整合", { bold: true }),
      tcell("仙南圏域比較困難", { color: C.orange }),
      tcell("登録状況確認・町担当課経由で見える化担当に照会"),
      tcell("B", { color: C.blue, bold: true, align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(riskTable);

// ===========================================================
// Section 7: 週次マイルストーン管理
// ===========================================================
S.push(...chapterCompact(7, "週次マイルストーン管理", C.navy));

S.push(p("Phase 2の3.5ヶ月間を週単位で管理し、各週の主要マイルストーンを明示します。隔週でメール・Web会議による中間報告を町担当課と実施します。"));

const weeklyTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("週", { width: 8 }),
      thcell("時期", { width: 14 }),
      thcell("主要マイルストーン", { width: 36 }),
      thcell("成果物・確認事項", { width: 30 }),
      thcell("優先度", { width: 12 }),
    ]}),
    new TableRow({ children: [
      tcell("W1", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("R8.8第1週"),
      tcell("アンケート受領・集計開始(工程1-2)"),
      tcell("生データ・単純集計表"),
      tcell("A", { color: C.red, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("W2", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("R8.8第2週"),
      tcell("アンケート集計完了(工程3-4)・第1回委員会開催準備"),
      tcell("集計結果一式・委員会資料最終確認"),
      tcell("A", { color: C.red, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("W3", { bold: true, align: AlignmentType.CENTER, fill: C.lgreen }),
      tcell("R8.8第3週"),
      tcell("第1回策定委員会開催・議事録案作成"),
      tcell("議事録案・委員意見転記開始"),
      tcell("A", { color: C.red, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("W4", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("R8.8第4週"),
      tcell("委員意見集約・第1回3者協議"),
      tcell("意見管理シートVer.1・対応方針一次案"),
      tcell("A", { color: C.red, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("W5", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("R8.9第1週"),
      tcell("対応方針確定・Ver.2.0更新作業開始"),
      tcell("対応方針一覧・Ver.2.0素案(章1-2)"),
      tcell("A", { color: C.red, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("W6", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("R8.9第2週"),
      tcell("第3章・第4章更新(実績・基本理念・基本目標)"),
      tcell("Ver.2.0素案(章3-4)"),
      tcell("A", { color: C.red, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("W7", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("R8.9第3週"),
      tcell("第5章・第6章更新(重点施策・認知症)"),
      tcell("Ver.2.0素案(章5-6)"),
      tcell("A", { color: C.red, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("W8", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("R8.9第4週"),
      tcell("見込量精緻化(6ステップ)・第7章更新"),
      tcell("Ver.2.0素案(章7)・見込量資料"),
      tcell("A", { color: C.red, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("W9", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("R8.10第1週"),
      tcell("第8章更新・全体統合(Ver.2.0素案v1)"),
      tcell("Ver.2.0素案v1(全章)"),
      tcell("A", { color: C.red, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("W10", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("R8.10第2週"),
      tcell("Red Team検証(整合性・MECE・課題接続・数値・国通知)"),
      tcell("Ver.2.0素案v2(修正版)"),
      tcell("A", { color: C.red, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("W11", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("R8.10第3週"),
      tcell("町担当課確認・最終調整"),
      tcell("Ver.2.0素案v3(確認版)"),
      tcell("A", { color: C.red, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("W12", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("R8.10第4週"),
      tcell("Ver.2.0完成・第2回委員会資料作成開始"),
      tcell("Ver.2.0完成版・委員会資料素案"),
      tcell("A", { color: C.red, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("W13", { bold: true, align: AlignmentType.CENTER, fill: C.lorange }),
      tcell("R8.11第1週"),
      tcell("第2回委員会準備完了・想定問答更新"),
      tcell("第2回委員会資料一式"),
      tcell("A", { color: C.red, bold: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("W14", { bold: true, align: AlignmentType.CENTER, fill: C.lgreen }),
      tcell("R8.11第2週"),
      tcell("第2回策定委員会開催・Phase 3移行準備", { bold: true, color: C.green }),
      tcell("Ver.2.0審議結果・Phase 3計画", { color: C.green }),
      tcell("A", { color: C.red, bold: true, align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(weeklyTable);

S.push(spacer());

S.push(new Paragraph({
  spacing: { before: 240, after: 240 },
  alignment: AlignmentType.CENTER,
  border: {
    top: { style: BorderStyle.DOUBLE, size: 18, color: C.navy },
    bottom: { style: BorderStyle.DOUBLE, size: 18, color: C.navy },
  },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  children: [text("Phase 2の完了をもって、計画策定の中核作業が概ね完了します。Phase 3(保険料試算精緻化・最終確定)へ円滑に移行します。",
    { size: 22, bold: true, color: C.navy })],
}));

// ===========================================================
// ドキュメント生成
// ===========================================================
const doc = new Document({
  creator: "ビズアップ公共コンサルティング株式会社",
  title: "川崎町 Phase 2作業計画 詳細書",
  description: "R8.8中旬〜R8.11の詳細作業計画",
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
            text: "川崎町第10期介護保険事業計画 Phase 2 作業計画詳細書",
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
  fs.writeFileSync("/home/claude/kawasaki_work/川崎町_Phase2作業計画詳細書.docx", buffer);
  console.log("Build done. Blocks:", S.length);
}).catch(err => {
  console.error("Build error:", err);
  process.exit(1);
});
