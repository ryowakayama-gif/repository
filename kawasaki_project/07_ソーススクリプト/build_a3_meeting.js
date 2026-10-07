/**
 * A-3: 町担当への説明資料
 * 川崎町保健福祉課・大宮様との対面打合せで使用する説明資料
 * 計画素案v1.5の要点・川崎町固有論点・記入依頼事項・スケジュールを視覚化
 *
 * 構成（13頁想定）:
 * 1. 表紙
 * 2. 本日のアジェンダ
 * 3. 計画策定の全体スケジュール
 * 4. 計画素案v1.5 章構成
 * 5. 川崎町固有論点6つ
 * 6. 計画素案の特徴①図表6種
 * 7. 計画素案の特徴②第7章保険料試算枠組み
 * 8. 計画素案の特徴③第6章認知症基本法対応
 * 9. 町からのご記入依頼①(優先度A)
 * 10. 町からのご記入依頼②
 * 11. アンケート概要と回収協力依頼
 * 12. 次回打合せまでのアクション
 * 13. お問合せ先・閉じ
 */
const fs = require('fs');
const H = require('./plan_helpers');
const {
  C, CH, FONT,
  text, p, section, subsection,
  bullet, numItem, placeholder, fact, source,
  tcell, thcell, kvTable, spacer,
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, convertInchesToTwip,
} = H;

const S = [];

// ===========================================================
// スライド形式のヘルパー
// ===========================================================
function slideTitle(num, title, color) {
  color = color || C.navy;
  return [
    new Paragraph({
      spacing: { before: 0, after: 0, line: 240 },
      pageBreakBefore: num > 1,
      children: [text("", { size: 16 })],
    }),
    new Paragraph({
      spacing: { before: 200, after: 80, line: 320 },
      children: [
        text(`SLIDE ${num}`, { size: 16, bold: true, color: C.gray }),
      ],
    }),
    new Paragraph({
      spacing: { before: 0, after: 360, line: 360 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 24, color: color } },
      children: [text(title, { size: 32, bold: true, color: color })],
    }),
  ];
}

function bigKey(label, value, color) {
  color = color || C.navy;
  return new Paragraph({
    spacing: { before: 120, after: 120 },
    alignment: AlignmentType.CENTER,
    shading: { type: ShadingType.SOLID, fill: C.lblue },
    border: {
      top: { style: BorderStyle.SINGLE, size: 12, color: color },
      bottom: { style: BorderStyle.SINGLE, size: 12, color: color },
    },
    children: [
      text(label + "  ", { size: 20, bold: true, color: C.gray }),
      text(value, { size: 28, bold: true, color: color }),
    ],
  });
}

// ===========================================================
// SLIDE 1: 表紙
// ===========================================================
S.push(new Paragraph({
  spacing: { before: 2400, after: 240 },
  alignment: AlignmentType.CENTER,
  children: [text("川崎町高齢者保健福祉計画", { size: 30, bold: true, color: C.navy })],
}));
S.push(new Paragraph({
  spacing: { before: 0, after: 480 },
  alignment: AlignmentType.CENTER,
  children: [text("第10期介護保険事業計画 策定業務", { size: 30, bold: true, color: C.navy })],
}));
S.push(new Paragraph({
  spacing: { before: 240, after: 240 },
  alignment: AlignmentType.CENTER,
  border: {
    top: { style: BorderStyle.DOUBLE, size: 18, color: C.navy },
    bottom: { style: BorderStyle.DOUBLE, size: 18, color: C.navy },
  },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  children: [text("町担当課 ご説明資料", { size: 36, bold: true, color: C.navy })],
}));
S.push(new Paragraph({
  spacing: { before: 240, after: 240 },
  alignment: AlignmentType.CENTER,
  children: [text("〜 計画素案v1.5 の進捗ご報告と町担当ご記入依頼のお願い 〜", {
    size: 22, italics: true, color: C.blue
  })],
}));

S.push(new Paragraph({
  spacing: { before: 3600, after: 240 },
  alignment: AlignmentType.CENTER,
  children: [text("令和8年6月", { size: 22, color: C.gray })],
}));
S.push(new Paragraph({
  spacing: { before: 80, after: 80 },
  alignment: AlignmentType.CENTER,
  children: [text("ビズアップ公共コンサルティング株式会社", { size: 22, bold: true, color: C.navy })],
}));
S.push(new Paragraph({
  spacing: { before: 0, after: 80 },
  alignment: AlignmentType.CENTER,
  children: [text("札幌事業所 / 主担当：若山", { size: 20, color: C.gray })],
}));

// ===========================================================
// SLIDE 2: 本日のアジェンダ
// ===========================================================
S.push(...slideTitle(2, "本日のアジェンダ"));

S.push(new Paragraph({
  spacing: { before: 240, after: 360 },
  children: [text("本日は、以下の流れでご報告・ご協議させていただきます。所要時間は約45〜60分を想定しております。",
    { size: 22 })],
}));

const agendaTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("Part", { width: 8 }),
      thcell("内容", { width: 50 }),
      thcell("資料", { width: 30 }),
      thcell("所要", { width: 12 }),
    ]}),
    new TableRow({ children: [
      tcell("Part 1", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("計画策定スケジュールと現在地のご報告"),
      tcell("本資料 SLIDE 3"),
      tcell("5分", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("Part 2", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("計画素案 Ver.1.5 の章構成と特徴のご説明"),
      tcell("素案v1.5 (43頁)・SLIDE 4-8"),
      tcell("20分", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("Part 3", { bold: true, align: AlignmentType.CENTER, fill: C.lorange }),
      tcell("町担当課へのご記入依頼事項のご説明", { bold: true, color: C.orange }),
      tcell("記入用フォーマット3種・SLIDE 9-10", { color: C.orange }),
      tcell("15分", { align: AlignmentType.CENTER, color: C.orange, bold: true }),
    ]}),
    new TableRow({ children: [
      tcell("Part 4", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("アンケート概要と回収協力のお願い"),
      tcell("SLIDE 11"),
      tcell("10分", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("Part 5", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("次回打合せまでのアクション・質疑応答"),
      tcell("SLIDE 12"),
      tcell("10分", { align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(agendaTable);

S.push(spacer());
S.push(new Paragraph({
  spacing: { before: 240, after: 240 },
  alignment: AlignmentType.CENTER,
  shading: { type: ShadingType.SOLID, fill: C.lorange },
  border: {
    top: { style: BorderStyle.SINGLE, size: 12, color: C.orange },
    bottom: { style: BorderStyle.SINGLE, size: 12, color: C.orange },
  },
  children: [text("【本日の最重要事項】 町担当課にご記入をお願いする3つのフォーマットについて、優先順位と進め方をご確認いただきます。",
    { size: 22, bold: true, color: C.orange })],
}));

// ===========================================================
// SLIDE 3: 全体スケジュール
// ===========================================================
S.push(...slideTitle(3, "計画策定の全体スケジュール"));

S.push(bigKey("現在地：", "Phase 1 完了直前（85%）→ Phase 2 開始", C.orange));

S.push(new Paragraph({
  spacing: { before: 200, after: 200 },
  children: [text("全4フェーズで構成される本業務は、現在Phase 1の素案準備完了直前にあります。R8.6下旬の町送付パッケージ発送・R8.7末のアンケート回収を経て、R8.8中旬の第1回策定委員会が最初の大きな節目となります。",
    { size: 21 })],
}));

const scheduleTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("Phase", { width: 10 }),
      thcell("期間", { width: 16 }),
      thcell("主な作業", { width: 44 }),
      thcell("マイルストーン", { width: 18 }),
      thcell("現状", { width: 12 }),
    ]}),
    new TableRow({ children: [
      tcell("Phase 1", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("R7.11〜R8.8中旬", { align: AlignmentType.CENTER }),
      tcell("キックオフ・素案v1.5完成・町送付パッケージ整備"),
      tcell("第1回委員会"),
      tcell("85%完了", { bold: true, color: C.orange, align: AlignmentType.CENTER, fill: C.lorange }),
    ]}),
    new TableRow({ children: [
      tcell("Phase 2", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("R8.8中旬〜R8.11", { align: AlignmentType.CENTER }),
      tcell("アンケート集計・Ver.2.0更新・サービス見込量精緻化"),
      tcell("第2回委員会"),
      tcell("未着手", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("Phase 3", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("R8.11〜R9.2", { align: AlignmentType.CENTER }),
      tcell("保険料3パターン試算・13段階区分検討・町長答申"),
      tcell("第3-4回委員会"),
      tcell("未着手", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("Phase 4", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("R9.2〜R9.3", { align: AlignmentType.CENTER }),
      tcell("パブリックコメント・3月議会上程・公表"),
      tcell("計画書公表"),
      tcell("未着手", { align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(scheduleTable);

S.push(spacer());

S.push(new Paragraph({
  spacing: { before: 200, after: 200 },
  shading: { type: ShadingType.SOLID, fill: C.lorange },
  children: [text("● 直近で最も重要な日付：R8.6下旬（町送付）／R8.7末（アンケート回収）／R8.8中旬（第1回委員会）",
    { size: 22, bold: true, color: C.orange })],
}));

// ===========================================================
// SLIDE 4: 計画素案v1.5 章構成
// ===========================================================
S.push(...slideTitle(4, "計画素案 Ver.1.5 の章構成（全43頁・8章）"));

S.push(new Paragraph({
  spacing: { before: 240, after: 240 },
  children: [text("計画素案は8章構成・43頁・403ブロック。各章はチャプターカラーで識別され、節見出し・項見出しまで統一されたデザインシステムを採用しています。",
    { size: 22 })],
}));

const chapterTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("章", { width: 6 }),
      thcell("章タイトル", { width: 30 }),
      thcell("カラー", { width: 16 }),
      thcell("内容", { width: 36 }),
      thcell("頁", { width: 12 }),
    ]}),
    new TableRow({ children: [
      tcell("1", { bold: true, align: AlignmentType.CENTER, fill: "DAE3F3", color: "1F3864" }),
      tcell("計画の策定にあたって", { color: "1F3864", bold: true }),
      tcell("ダークネイビー", { color: "1F3864", align: AlignmentType.CENTER }),
      tcell("背景・目的・位置付け・期間・体制"),
      tcell("4頁", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("2", { bold: true, align: AlignmentType.CENTER, fill: "E2EFDA", color: "375623" }),
      tcell("川崎町の高齢者を取り巻く現状", { color: "375623", bold: true }),
      tcell("ダークグリーン", { color: "375623", align: AlignmentType.CENTER }),
      tcell("人口・世帯・高齢化率・認定・給付・地域資源"),
      tcell("8頁", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("3", { bold: true, align: AlignmentType.CENTER, fill: "DAE3F3", color: "2E75B6" }),
      tcell("第9期計画の取組実績と評価", { color: "2E75B6", bold: true }),
      tcell("ミドルブルー", { color: "2E75B6", align: AlignmentType.CENTER }),
      tcell("6目標別実績・評価・課題"),
      tcell("6頁", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("4", { bold: true, align: AlignmentType.CENTER, fill: "FCE4D6", color: "C55A11" }),
      tcell("計画の基本理念と基本目標", { color: "C55A11", bold: true }),
      tcell("バーンオレンジ", { color: "C55A11", align: AlignmentType.CENTER }),
      tcell("基本理念・6目標・体系図"),
      tcell("3頁", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("5", { bold: true, align: AlignmentType.CENTER, fill: "E4D6F0", color: "7030A0" }),
      tcell("施策の展開", { color: "7030A0", bold: true }),
      tcell("パープル", { color: "7030A0", align: AlignmentType.CENTER }),
      tcell("基本目標1〜5の施策・主な事業・KPI"),
      tcell("5頁", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("6", { bold: true, align: AlignmentType.CENTER, fill: "EAD5F0", color: "9333B0" }),
      tcell("認知症施策推進計画（独立章）", { color: "9333B0", bold: true }),
      tcell("マゼンタ", { color: "9333B0", align: AlignmentType.CENTER }),
      tcell("認知症基本法対応・7基本的施策・重点5本柱"),
      tcell("3頁", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("7", { bold: true, align: AlignmentType.CENTER, fill: "D6E4F0", color: "0F4F73" }),
      tcell("介護保険サービス見込量と保険料", { color: "0F4F73", bold: true }),
      tcell("ダークティール", { color: "0F4F73", align: AlignmentType.CENTER }),
      tcell("見込量6ステップ・保険料8ステップ・3パターン"),
      tcell("9頁", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("8", { bold: true, align: AlignmentType.CENTER, fill: "E2E8F0", color: "404040" }),
      tcell("計画の推進体制と評価", { color: "404040", bold: true }),
      tcell("チャコール", { color: "404040", align: AlignmentType.CENTER }),
      tcell("推進体制・PDCAサイクル・評価・公表"),
      tcell("2頁", { align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(chapterTable);

// ===========================================================
// SLIDE 5: 川崎町固有論点
// ===========================================================
S.push(...slideTitle(5, "計画に反映した 川崎町固有の6論点"));

S.push(new Paragraph({
  spacing: { before: 240, after: 240 },
  children: [text("計画素案には、キックオフ会議でご確認いただいた川崎町固有の論点を6つの重点として反映しています。",
    { size: 22 })],
}));

const issuesTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("No", { width: 6 }),
      thcell("論点", { width: 26 }),
      thcell("計画への反映箇所", { width: 22 }),
      thcell("対応の要点", { width: 46 }),
    ]}),
    new TableRow({ children: [
      tcell("①", { bold: true, align: AlignmentType.CENTER, fill: C.lorange }),
      tcell("高齢者外出タクシー助成 R7.3終了", { bold: true }),
      tcell("第5章 5-2", { align: AlignmentType.CENTER }),
      tcell("社協NPO移送・デマンドバス・町民バスの3層構造を整理。住民周知と利用情報の一元化を重点施策化"),
    ]}),
    new TableRow({ children: [
      tcell("②", { bold: true, align: AlignmentType.CENTER, fill: C.lorange }),
      tcell("認知症基本法 R6.1施行", { bold: true }),
      tcell("第6章 独立章", { align: AlignmentType.CENTER }),
      tcell("基本目標6を新設し独立章化。本人意見聴取（包括センター・国保川崎病院経由）を新規実施"),
    ]}),
    new TableRow({ children: [
      tcell("③", { bold: true, align: AlignmentType.CENTER, fill: C.lorange }),
      tcell("施設の地域偏在（役場周辺集中）", { bold: true }),
      tcell("第5章 5-4 / 7-1", { align: AlignmentType.CENTER }),
      tcell("住所地特例24人を含む町外施設利用を見込量に組み込み。広域連携の検討材料"),
    ]}),
    new TableRow({ children: [
      tcell("④", { bold: true, align: AlignmentType.CENTER, fill: C.lorange }),
      tcell("町外医療機関の利用（広域医療連携）", { bold: true }),
      tcell("第5章 5-3", { align: AlignmentType.CENTER }),
      tcell("みやぎ県南中核病院（大河原）・刈田綜合病院（白石）等との連携体制を明示"),
    ]}),
    new TableRow({ children: [
      tcell("⑤", { bold: true, align: AlignmentType.CENTER, fill: C.lorange }),
      tcell("8050問題・老老介護増加", { bold: true }),
      tcell("第5章 5-3 / 第6章", { align: AlignmentType.CENTER }),
      tcell("家族介護者支援・介護離職防止支援を強化。認知症対応と連動"),
    ]}),
    new TableRow({ children: [
      tcell("⑥", { bold: true, align: AlignmentType.CENTER, fill: C.lorange }),
      tcell("包括センター人材体制（実質3名）", { bold: true }),
      tcell("第5章 5-5", { align: AlignmentType.CENTER }),
      tcell("認定調査員1名の負担軽減策を検討。社協運営の体制強化を方針化"),
    ]}),
  ],
});
S.push(issuesTable);

// ===========================================================
// SLIDE 6: 図表6種
// ===========================================================
S.push(...slideTitle(6, "計画素案の特徴① 第2-7章に埋め込んだ図表6種"));

S.push(new Paragraph({
  spacing: { before: 240, after: 240 },
  children: [text("計画書の視覚的訴求力を高めるため、第2章・第7章にmatplotlib生成の図表6種を埋め込んでいます。すべて川崎町実数値ベースで作成しています。",
    { size: 22 })],
}));

const chartsTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("番号", { width: 10 }),
      thcell("図表名", { width: 36 }),
      thcell("出典・川崎町数値", { width: 36 }),
      thcell("章", { width: 18 }),
    ]}),
    new TableRow({ children: [
      tcell("図2-1", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("第1号被保険者の年齢階級別推移"),
      tcell("R3末→R7.6 後期高齢者+18.2%増（団塊世代到達）"),
      tcell("第2章 2-2", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("図2-2", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("サービス受給者の区分別構成（円グラフ）"),
      tcell("R7.6 居宅59.2% / 地密12.0% / 施設28.8% (計466人)"),
      tcell("第2章 2-3", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("図2-3", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("サービス区分別の年間給付費（横棒）"),
      tcell("R3年度 総額9.6億円 / 施設46.0% / 居宅38.1% / 地密15.9%"),
      tcell("第2章 2-4", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("図2-4", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("所得段階別第1号被保険者の構成"),
      tcell("R3末3,255人 / 非課税層30.1% / 第5基準697人 / 課税層34.2%"),
      tcell("第2章 2-2", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("図2-5", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("高齢化率比較（町・県・全国）"),
      tcell("川崎町41.4% / 宮城県28.5% / 全国29.1% → 県内5位"),
      tcell("第2章 2-2", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("図7-1", { bold: true, align: AlignmentType.CENTER, fill: C.lorange }),
      tcell("介護保険料月額基準額の推移（3パターン）"),
      tcell("第8期6,380 → 第9期6,500 → 第10期A/B/C 試算枠"),
      tcell("第7章 7-3", { align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(chartsTable);

// ===========================================================
// SLIDE 7: 第7章保険料試算
// ===========================================================
S.push(...slideTitle(7, "計画素案の特徴② 第7章 保険料試算の枠組み（9頁）"));

S.push(new Paragraph({
  spacing: { before: 240, after: 240 },
  children: [text("第7章は計画の核心部分です。サービス見込量の6ステップと、保険料算定の8ステップを明確化し、3パターン試算と13段階区分検討を盛り込みました。",
    { size: 22 })],
}));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ サービス見込量算定の6ステップ（7-1）", { size: 22, bold: true, color: C.navy })],
}));

const step6Table = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      tcell("Step1", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("人口推計（社人研R5推計）"),
      tcell("Step2", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("認定者推計（年齢階級別認定率×人口）"),
      tcell("Step3", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("利用率算定（国保連データ）"),
    ]}),
    new TableRow({ children: [
      tcell("Step4", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("見込量推計（認定者×利用率×利用量）"),
      tcell("Step5", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("アンケート補正（潜在ニーズ反映）"),
      tcell("Step6", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("見える化登録・仙南圏域比較"),
    ]}),
  ],
});
S.push(step6Table);

S.push(new Paragraph({
  spacing: { before: 240, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lorange },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.orange } },
  children: [text("　▌ 保険料試算の3パターン（7-3）", { size: 22, bold: true, color: C.orange })],
}));

const pat3Table = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("パターン", { width: 12 }),
      thcell("基金取崩方針", { width: 22 }),
      thcell("特徴", { width: 44 }),
      thcell("月額イメージ", { width: 22 }),
    ]}),
    new TableRow({ children: [
      tcell("A", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
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
      tcell("C", { bold: true, align: AlignmentType.CENTER, fill: C.lgreen }),
      tcell("全額取崩", { bold: true }),
      tcell("住民負担最小化・第11期負担増のリスクあり"),
      tcell("最低水準", { align: AlignmentType.CENTER, color: C.green }),
    ]}),
  ],
});
S.push(pat3Table);

// ===========================================================
// SLIDE 8: 第6章認知症基本法
// ===========================================================
S.push(...slideTitle(8, "計画素案の特徴③ 第6章 認知症基本法対応"));

S.push(new Paragraph({
  spacing: { before: 240, after: 240 },
  children: [text("令和6年1月施行の認知症基本法を踏まえ、第10期計画では認知症施策を独立章（第6章）として位置付けました。基本法第15条〜21条の7基本的施策に対応した町施策体系を構築しています。",
    { size: 22 })],
}));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: "EAD5F0" },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "9333B0" } },
  children: [text("　▌ 重点施策 5本柱（6-3）", { size: 22, bold: true, color: "9333B0" })],
}));

const j5Table = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      tcell("J-1", { bold: true, align: AlignmentType.CENTER, fill: "EAD5F0", color: "9333B0" }),
      tcell("認知症サポーターの拡大と質向上", { bold: true }),
      tcell("累計550名を基盤に企業・学校サポーター拡大"),
    ]}),
    new TableRow({ children: [
      tcell("J-2", { bold: true, align: AlignmentType.CENTER, fill: "EAD5F0", color: "9333B0" }),
      tcell("チームオレンジの整備", { bold: true }),
      tcell("新規整備。本人・家族支援を実装"),
    ]}),
    new TableRow({ children: [
      tcell("J-3", { bold: true, align: AlignmentType.CENTER, fill: "EAD5F0", color: "9333B0" }),
      tcell("認知症カフェ・本人ミーティング", { bold: true }),
      tcell("「喫茶みかん」継続＋本人ミーティング新設"),
    ]}),
    new TableRow({ children: [
      tcell("J-4", { bold: true, align: AlignmentType.CENTER, fill: "EAD5F0", color: "9333B0" }),
      tcell("早期発見・早期対応の体制強化", { bold: true }),
      tcell("初期集中支援チーム強化・もの忘れ相談継続"),
    ]}),
    new TableRow({ children: [
      tcell("J-5", { bold: true, align: AlignmentType.CENTER, fill: "EAD5F0", color: "9333B0" }),
      tcell("国保川崎病院との医療連携", { bold: true }),
      tcell("認知症診断・治療の地域中核として連携深化"),
    ]}),
  ],
});
S.push(j5Table);

S.push(spacer());

S.push(new Paragraph({
  spacing: { before: 200, after: 200 },
  shading: { type: ShadingType.SOLID, fill: "EAD5F0" },
  children: [text("【新規取組】 チームオレンジの新規整備、認知症本人ミーティングの新規実施、本人・家族の意見聴取（包括センター・国保川崎病院経由）が認知症基本法対応の必須事項です。",
    { size: 22, bold: true, color: "9333B0" })],
}));

// ===========================================================
// SLIDE 9: 町からのご記入依頼①
// ===========================================================
S.push(...slideTitle(9, "町担当課へのご記入依頼① 優先度A", C.orange));

S.push(new Paragraph({
  spacing: { before: 240, after: 240 },
  shading: { type: ShadingType.SOLID, fill: C.lorange },
  children: [text("ここからが本日の最重要事項です。3つの記入用フォーマットへのご対応をお願いします。",
    { size: 22, bold: true, color: C.orange })],
}));

S.push(new Paragraph({
  spacing: { before: 240, after: 100, line: 320 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C.orange } },
  children: [text("① 川崎町_第9期実績一覧_町記入用.xlsx", { size: 26, bold: true, color: C.orange })],
}));

S.push(p("【目的】 第10期計画素案 第3章「第9期計画の取組実績と評価」のための実績データ収集"));
S.push(p("【構成】 7シート構成・35事業以上を網羅"));

const sheet1Table = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("シート", { width: 26 }),
      thcell("内容", { width: 44 }),
      thcell("優先度", { width: 12 }),
      thcell("締切", { width: 18 }),
    ]}),
    new TableRow({ children: [
      tcell("01_KPI一覧", { bold: true, fill: C.lorange }),
      tcell("第9期計画KPI目標値・実績・達成率（5基本目標×3〜5指標）"),
      tcell("A", { bold: true, align: AlignmentType.CENTER, color: C.red, fill: C.lorange }),
      tcell("R8.6下旬", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("02_介護予防_健康", { bold: true, fill: C.lorange }),
      tcell("通いの場・ユニバーサルサポーター（7種別）・健診関連"),
      tcell("A", { bold: true, align: AlignmentType.CENTER, color: C.red, fill: C.lorange }),
      tcell("R8.6下旬", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("03_在宅生活支援", { bold: true }),
      tcell("紙おむつ・配食・緊急通報・移動支援3層構造"),
      tcell("A", { bold: true, align: AlignmentType.CENTER, color: C.red, fill: C.lorange }),
      tcell("R8.7上旬", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("04_認知症_包括", { bold: true, fill: C.lorange }),
      tcell("認知症サポーター・カフェ・初期集中支援・包括センター体制"),
      tcell("A", { bold: true, align: AlignmentType.CENTER, color: C.red, fill: C.lorange }),
      tcell("R8.6下旬", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("05_介護サービス_人材", { bold: true }),
      tcell("認定者数・施設利用町内町外・町内事業所・人材実態"),
      tcell("B", { bold: true, align: AlignmentType.CENTER, color: C.blue, fill: C.lblue }),
      tcell("R8.7上旬", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("06_第10期反映方針", { bold: true }),
      tcell("事業棚卸し（継続/拡充/縮小/廃止/新規/統合）"),
      tcell("B", { bold: true, align: AlignmentType.CENTER, color: C.blue, fill: C.lblue }),
      tcell("R8.7上旬", { align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(sheet1Table);

S.push(spacer());

S.push(new Paragraph({
  spacing: { before: 200, after: 200 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  children: [text("★ 既知数値（ユニバーサルサポーター登録者数・認知症サポーター累計550名等）は緑セルで事前投入済みです。最新値があれば右欄に追記をお願いします。",
    { size: 21, color: C.navy })],
}));

// ===========================================================
// SLIDE 10: 町からのご記入依頼②
// ===========================================================
S.push(...slideTitle(10, "町担当課へのご記入依頼② MECEデータ", C.orange));

S.push(new Paragraph({
  spacing: { before: 240, after: 100, line: 320 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C.orange } },
  children: [text("② 川崎町_必要資料_入力フォーマット.xlsx（MECE版13シート）", { size: 26, bold: true, color: C.orange })],
}));

S.push(p("【目的】 第10期計画素案 第2章・第7章のための保険給付実績データ収集"));
S.push(p("【現状】 前回送付済み・一部記入いただいた状態。残りの空欄を埋めていただく形となります。"));

const sheet2Table = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("データ区分", { width: 34 }),
      thcell("計画素案 反映先", { width: 30 }),
      thcell("締切", { width: 18 }),
      thcell("補足", { width: 18 }),
    ]}),
    new TableRow({ children: [
      tcell("第1号被保険者数の経年推移（R3〜R7）", { bold: true }),
      tcell("第2章 2-1 第7章 7-1", { align: AlignmentType.CENTER }),
      tcell("R8.6下旬", { align: AlignmentType.CENTER, fill: C.lorange }),
      tcell("最重要", { color: C.red, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("要介護認定者数（要介護度別・年度末）", { bold: true }),
      tcell("第2章 2-3 第7章 7-1", { align: AlignmentType.CENTER }),
      tcell("R8.6下旬", { align: AlignmentType.CENTER, fill: C.lorange }),
      tcell("最重要", { color: C.red, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("サービス受給者数（サービス種別×要介護度）", { bold: true }),
      tcell("第2章 2-3 第7章 7-1", { align: AlignmentType.CENTER }),
      tcell("R8.6下旬", { align: AlignmentType.CENTER, fill: C.lorange }),
      tcell("最重要", { color: C.red, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("給付費実績（標準給付費・地域支援事業費）", { bold: true }),
      tcell("第2章 2-4 第7章 7-2", { align: AlignmentType.CENTER }),
      tcell("R8.6下旬", { align: AlignmentType.CENTER, fill: C.lorange }),
      tcell("最重要", { color: C.red, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("介護給付費準備基金 残高（R8.6時点）", { bold: true, color: C.red }),
      tcell("第7章 7-3 保険料試算", { align: AlignmentType.CENTER, color: C.red, bold: true }),
      tcell("R8.6確定", { align: AlignmentType.CENTER, fill: C.lorange, bold: true, color: C.red }),
      tcell("極めて重要", { color: C.red, align: AlignmentType.CENTER, bold: true }),
    ]}),
    new TableRow({ children: [
      tcell("保険料収納率（特別徴収・普通徴収別）", { bold: true }),
      tcell("第7章 7-3", { align: AlignmentType.CENTER }),
      tcell("R8.6下旬", { align: AlignmentType.CENTER }),
      tcell("予定収納率の根拠", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("住所地特例該当者の所在自治体内訳", { bold: true }),
      tcell("第5章 5-4 第7章 7-1", { align: AlignmentType.CENTER }),
      tcell("R8.7", { align: AlignmentType.CENTER }),
      tcell("把握範囲で", { color: C.gray, align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(sheet2Table);

S.push(spacer());

S.push(new Paragraph({
  spacing: { before: 240, after: 240 },
  alignment: AlignmentType.CENTER,
  shading: { type: ShadingType.SOLID, fill: C.lorange },
  border: {
    top: { style: BorderStyle.DOUBLE, size: 12, color: C.red },
    bottom: { style: BorderStyle.DOUBLE, size: 12, color: C.red },
  },
  children: [text("最重要：介護給付費準備基金残高（R8.6時点）は、保険料試算3パターンの基準となる中核データです",
    { size: 22, bold: true, color: C.red })],
}));

// ===========================================================
// SLIDE 11: アンケート概要
// ===========================================================
S.push(...slideTitle(11, "アンケート概要と回収協力のお願い"));

S.push(new Paragraph({
  spacing: { before: 240, after: 240 },
  children: [text("計画素案Ver.2.0への反映に向け、令和8年7月末を目途にアンケートを回収します。発送・督促を町担当課にお願いします。",
    { size: 22 })],
}));

const surveyTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("項目", { width: 20 }),
      thcell("一般高齢者ニーズ調査", { width: 40 }),
      thcell("認定者調査", { width: 40 }),
    ]}),
    new TableRow({ children: [
      tcell("対象", { bold: true, fill: C.lblue }),
      tcell("65歳以上の高齢者", { align: AlignmentType.CENTER }),
      tcell("要支援・要介護認定者", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("発送数", { bold: true, fill: C.lblue }),
      tcell("1,000名", { align: AlignmentType.CENTER, bold: true }),
      tcell("300名", { align: AlignmentType.CENTER, bold: true }),
    ]}),
    new TableRow({ children: [
      tcell("設問数", { bold: true, fill: C.lblue }),
      tcell("19問（国標準+川崎町追加3問）", { align: AlignmentType.CENTER }),
      tcell("16問（国標準+川崎町追加4問）", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("川崎町追加問", { bold: true, fill: C.lorange }),
      tcell("町外医療機関利用・町独自支援認知・災害時避難不安", { color: C.orange }),
      tcell("町外施設利用・所在地・町外医療・最大不安", { color: C.orange }),
    ]}),
    new TableRow({ children: [
      tcell("発送時期", { bold: true, fill: C.lblue }),
      tcell("R8.6下旬", { align: AlignmentType.CENTER, color: C.orange, bold: true }),
      tcell("R8.6下旬", { align: AlignmentType.CENTER, color: C.orange, bold: true }),
    ]}),
    new TableRow({ children: [
      tcell("回収締切", { bold: true, fill: C.lblue }),
      tcell("R8.7末", { align: AlignmentType.CENTER, color: C.red, bold: true }),
      tcell("R8.7末", { align: AlignmentType.CENTER, color: C.red, bold: true }),
    ]}),
    new TableRow({ children: [
      tcell("集計", { bold: true, fill: C.lgreen }),
      tcell("弊社で実施（約2週間）", { align: AlignmentType.CENTER }),
      tcell("弊社で実施（約2週間）", { align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(surveyTable);

S.push(spacer());
S.push(new Paragraph({
  spacing: { before: 200, after: 200 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  children: [text("【町担当へのお願い】 アンケート発送・回収・督促のご対応をお願いします。回収後の原本・データを弊社にご送付ください。回収率向上のための工夫（広報誌掲載・再送等）もご相談させてください。",
    { size: 22, color: C.navy })],
}));

// ===========================================================
// SLIDE 12: 次回打合せまでのアクション
// ===========================================================
S.push(...slideTitle(12, "次回打合せまでのアクション", C.green));

S.push(new Paragraph({
  spacing: { before: 240, after: 240 },
  children: [text("第1回策定委員会（R8.8中旬）までの約2ヶ月間に、双方で進めるアクションを整理します。",
    { size: 22 })],
}));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lorange },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.orange } },
  children: [text("　▌ 町担当課のアクション", { size: 22, bold: true, color: C.orange })],
}));

S.push(numItem("①", "アンケート発送（R8.6下旬・1,300名）"));
S.push(numItem("②", "①第9期実績一覧の優先シート（KPI・介護予防・認知症）記入（R8.6下旬）"));
S.push(numItem("③", "③MECEデータの残空欄記入（R8.6下旬）"));
S.push(numItem("④", "介護給付費準備基金残高の確定とご連絡（R8.6時点値・最重要）"));
S.push(numItem("⑤", "アンケート回収・督促対応（R8.7末まで）"));
S.push(numItem("⑥", "①第9期実績一覧の残シート記入（R8.7上旬）"));

S.push(new Paragraph({
  spacing: { before: 240, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.blue } },
  children: [text("　▌ 弊社（ビズアップ）のアクション", { size: 22, bold: true, color: C.blue })],
}));

S.push(numItem("①", "町担当ご質問への随時対応"));
S.push(numItem("②", "アンケート受領後の集計・分析（R8.8上旬・約2週間）"));
S.push(numItem("③", "計画素案Ver.2.0の更新（実数値・アンケート結果反映）"));
S.push(numItem("④", "第1回策定委員会の資料作成（R8.8中旬）"));
S.push(numItem("⑤", "中間進捗のメール報告（隔週）"));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lgreen },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.green } },
  children: [text("　▌ 中間打合せ（メール・Web会議）", { size: 22, bold: true, color: C.green })],
}));

S.push(p("ご記入の進捗状況に応じて、隔週でメールまたはWeb会議による打合せを実施いたします。記入で迷われた箇所は遠慮なくご相談ください。"));

// ===========================================================
// SLIDE 13: お問合せ・閉じ
// ===========================================================
S.push(...slideTitle(13, "お問合せ先・本日のまとめ"));

S.push(new Paragraph({
  spacing: { before: 240, after: 200, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ お問合せ先", { size: 22, bold: true, color: C.navy })],
}));

const contactTable2 = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      tcell("受託者", { bold: true, fill: C.lblue, width: 24 }),
      tcell("ビズアップ公共コンサルティング株式会社（札幌事業所）", { bold: true, width: 76 }),
    ]}),
    new TableRow({ children: [
      tcell("主担当", { bold: true, fill: C.lblue }),
      tcell("若山（プロジェクトリーダー）"),
    ]}),
    new TableRow({ children: [
      tcell("副担当", { bold: true, fill: C.lblue }),
      tcell("髙橋・山内・河崎"),
    ]}),
    new TableRow({ children: [
      tcell("連絡方法", { bold: true, fill: C.lblue }),
      tcell("メールまたは電話（双方OK）。Web会議のご要望も承ります"),
    ]}),
  ],
});
S.push(contactTable2);

S.push(spacer());

S.push(new Paragraph({
  spacing: { before: 240, after: 200, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lgreen },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.green } },
  children: [text("　▌ 本日のまとめ", { size: 22, bold: true, color: C.green })],
}));

S.push(numItem("①", "計画素案v1.5（43頁・8章）が完成し、川崎町固有の6論点を反映済み"));
S.push(numItem("②", "町担当課への3つの記入依頼フォーマットを整備（送付パッケージ含む）"));
S.push(numItem("③", "優先度Aの記入はR8.6下旬〜R8.7上旬に対応"));
S.push(numItem("④", "アンケート発送・回収は町担当課が対応（R8.6下旬〜R8.7末）"));
S.push(numItem("⑤", "次回マイルストーンは第1回策定委員会（R8.8中旬）"));

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
  children: [text("本日はお時間をいただきありがとうございました。今後ともどうぞよろしくお願いいたします。",
    { size: 24, bold: true, color: C.navy })],
}));

// ===========================================================
// ドキュメント生成
// ===========================================================
const doc = new Document({
  creator: "ビズアップ公共コンサルティング株式会社",
  title: "川崎町第10期計画 町担当課ご説明資料",
  description: "対面打合せ用説明資料",
  styles: {
    default: {
      document: { run: { font: FONT, size: 22 } },
    },
  },
  sections: [{
    properties: {
      page: {
        size: { width: 16838, height: 11906 },  // A4横向き
        margin: { top: 720, right: 1134, bottom: 720, left: 1134 },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({
            text: "川崎町第10期介護保険事業計画策定 町担当課ご説明資料",
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
            new TextRun({ text: "ビズアップ公共コンサルティング株式会社  ─ ", font: FONT, size: 16, color: C.gray }),
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
  fs.writeFileSync("/home/claude/kawasaki_work/川崎町_町担当課ご説明資料.docx", buffer);
  console.log("Build done. Blocks:", S.length);
}).catch(err => {
  console.error("Build error:", err);
  process.exit(1);
});
