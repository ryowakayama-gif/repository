/**
 * 川崎町第10期介護保険事業計画策定支援 引継ぎメモ
 * 
 * 次セッションで作業継続するための包括的引継ぎ資料
 * - 案件全体概要
 * - 主要数値・固有条件
 * - 完成成果物一覧（5カテゴリ）
 * - 残作業と次回アクション候補
 * - 技術的注意事項
 * - 推奨アップロードファイル一覧
 * 
 * 想定18頁
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
        text(name, { size: 26, bold: true, color: color }),
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
  children: [text("第10期介護保険事業計画 策定支援業務", { size: 28, bold: true, color: C.navy })],
}));
S.push(new Paragraph({
  spacing: { before: 240, after: 240 },
  alignment: AlignmentType.CENTER,
  border: {
    top: { style: BorderStyle.DOUBLE, size: 24, color: C.navy },
    bottom: { style: BorderStyle.DOUBLE, size: 24, color: C.navy },
  },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  children: [text("引 継 ぎ メ モ", { size: 40, bold: true, color: C.navy })],
}));
S.push(new Paragraph({
  spacing: { before: 240, after: 480 },
  alignment: AlignmentType.CENTER,
  children: [text("〜 次セッションでの作業継続のための包括資料 〜", { size: 22, italics: true, color: C.blue })],
}));

S.push(new Paragraph({
  spacing: { before: 480, after: 100, line: 320 },
  alignment: AlignmentType.CENTER,
  shading: { type: ShadingType.SOLID, fill: C.lorange },
  border: {
    top: { style: BorderStyle.SINGLE, size: 12, color: C.orange },
    bottom: { style: BorderStyle.SINGLE, size: 12, color: C.orange },
  },
  children: [text("現在地：Phase 1 完了直前 / 第1回策定委員会(R8.8中旬)前", { size: 22, bold: true, color: C.orange })],
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
S.push(new Paragraph({
  spacing: { before: 0, after: 80 },
  alignment: AlignmentType.RIGHT,
  children: [text("　札幌事業所 / 主担当：若山", { size: 18, color: C.gray })],
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
  ["Section 1", "案件全体概要", "3"],
  ["Section 2", "川崎町固有の主要数値・条件", "4"],
  ["Section 3", "完成成果物一覧（5カテゴリ）", "6"],
  ["Section 4", "プロジェクトスケジュールと現在地", "10"],
  ["Section 5", "残作業と次セッションでの優先候補", "12"],
  ["Section 6", "技術的注意事項（環境・スクリプト）", "14"],
  ["Section 7", "推奨アップロードファイル一覧", "16"],
  ["Section 8", "次セッション開始時の標準フロー", "18"],
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
// Section 1: 案件全体概要
// ===========================================================
S.push(...chapterCompact(1, "案件全体概要", C.navy));

S.push(...infoBox("案件名",
  "川崎町高齢者保健福祉計画・第10期介護保険事業計画・認知症施策推進計画 策定支援業務"));

S.push(...infoBox("受託者",
  "ビズアップ公共コンサルティング株式会社（札幌事業所）。主担当：若山。副担当：髙橋・山内・河崎。"));

S.push(...infoBox("発注者",
  "川崎町（宮城県柴田郡）保健福祉課。先方主担当：大宮様。"));

S.push(...infoBox("計画期間",
  "令和9年度〜令和11年度（3年間）の第10期介護保険事業計画。本計画には認知症基本法対応の認知症施策推進計画を独立章として包含する。"));

S.push(...infoBox("業務全体期間",
  "令和7年11月キックオフ〜令和9年3月議会上程まで。現在は令和8年6月時点、Phase 1完了直前で第1回策定委員会（R8.8中旬予定）前段階。"));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ プロジェクトの4フェーズ", { size: 22, bold: true, color: C.navy })],
}));

const phasesTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("Phase", { width: 10 }),
      thcell("期間", { width: 20 }),
      thcell("主要マイルストーン", { width: 44 }),
      thcell("現状", { width: 26 }),
    ]}),
    new TableRow({ children: [
      tcell("Phase 1", { bold: true, fill: C.lblue, align: AlignmentType.CENTER }),
      tcell("R7.11〜R8.8中旬", { align: AlignmentType.CENTER }),
      tcell("素案v1.6・町記入パッケージ・第1回委員会資料整備"),
      tcell("85%完了", { bold: true, color: C.orange, align: AlignmentType.CENTER, fill: C.lorange }),
    ]}),
    new TableRow({ children: [
      tcell("Phase 2", { bold: true, fill: C.lblue, align: AlignmentType.CENTER }),
      tcell("R8.8中旬〜R8.11", { align: AlignmentType.CENTER }),
      tcell("アンケート集計・Ver.2.0更新・第2回委員会"),
      tcell("準備完了・未着手", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("Phase 3", { bold: true, fill: C.lblue, align: AlignmentType.CENTER }),
      tcell("R8.11〜R9.2", { align: AlignmentType.CENTER }),
      tcell("保険料3パターン試算・13段階区分・第3-4回委員会"),
      tcell("準備完了・未着手", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("Phase 4", { bold: true, fill: C.lblue, align: AlignmentType.CENTER }),
      tcell("R9.2〜R9.3", { align: AlignmentType.CENTER }),
      tcell("パブコメ・3月議会上程・公表"),
      tcell("未着手", { align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(phasesTable);

// ===========================================================
// Section 2: 川崎町固有の主要数値・条件
// ===========================================================
S.push(...chapterCompact(2, "川崎町固有の主要数値・条件", C.navy));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 主要数値（令和7年6月時点・確定値）", { size: 22, bold: true, color: C.navy })],
}));

const numTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("項目", { width: 36 }),
      thcell("数値", { width: 30 }),
      thcell("特徴", { width: 34 }),
    ]}),
    new TableRow({ children: [
      tcell("総人口（推計）", { bold: true, fill: C.lblue }),
      tcell("約8,000人"),
      tcell("人口減少傾向"),
    ]}),
    new TableRow({ children: [
      tcell("第1号被保険者数", { bold: true, fill: C.lblue }),
      tcell("3,244人（R7.6）"),
      tcell("微減傾向"),
    ]}),
    new TableRow({ children: [
      tcell("後期高齢者（75歳以上）", { bold: true, fill: C.lblue }),
      tcell("1,675人（51.6%）"),
      tcell("4年で+18.2%増（団塊世代）"),
    ]}),
    new TableRow({ children: [
      tcell("高齢化率", { bold: true, fill: C.lorange }),
      tcell("41.4%", { bold: true, color: C.red }),
      tcell("県内5位・全国29.1%・宮城県28.5%"),
    ]}),
    new TableRow({ children: [
      tcell("住所地特例（町外施設）", { bold: true, fill: C.lorange }),
      tcell("24人（R7.6）", { color: C.orange, bold: true }),
      tcell("町外依存度比較的高い"),
    ]}),
    new TableRow({ children: [
      tcell("サービス受給者", { bold: true, fill: C.lblue }),
      tcell("466人（R7.6）"),
      tcell("居宅276・地密56・施設134"),
    ]}),
    new TableRow({ children: [
      tcell("R3年間給付費", { bold: true, fill: C.lblue }),
      tcell("9.6億円"),
      tcell("施設46.0%・居宅38.1%・地密15.9%"),
    ]}),
    new TableRow({ children: [
      tcell("第1号保険料 第8期", { bold: true, fill: C.lblue }),
      tcell("6,380円/月"),
      tcell("─"),
    ]}),
    new TableRow({ children: [
      tcell("第1号保険料 第9期", { bold: true, fill: C.lblue }),
      tcell("6,500円/月"),
      tcell("上昇率+1.9%（抑制的水準）"),
    ]}),
    new TableRow({ children: [
      tcell("R3保険料収納率", { bold: true, fill: C.lblue }),
      tcell("総合96.3%・現年99.0%"),
      tcell("特別徴収100.0%"),
    ]}),
    new TableRow({ children: [
      tcell("所得段階別R3末", { bold: true, fill: C.lblue }),
      tcell("3,255人（基準697人）"),
      tcell("非課税層30.1%・課税層34.2%"),
    ]}),
    new TableRow({ children: [
      tcell("認知症サポーター累計", { bold: true, fill: "EAD5F0" }),
      tcell("550名", { color: "9333B0" }),
      tcell("人口比約12.5%・キャラバンメイト73名"),
    ]}),
    new TableRow({ children: [
      tcell("ユニバーサルサポーター", { bold: true, fill: C.lorange }),
      tcell("14種別約400名", { color: C.orange }),
      tcell("町独自・介護予防84・スマイル40等"),
    ]}),
    new TableRow({ children: [
      tcell("地域包括支援センター職員", { bold: true, fill: C.lblue }),
      tcell("保健師3+認定調査1=4名"),
      tcell("社協運営・実質3名・業務過多"),
    ]}),
  ],
});
S.push(numTable);

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lorange },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.orange } },
  children: [text("　▌ 川崎町固有の5重点課題", { size: 22, bold: true, color: C.orange })],
}));

S.push(numItem("①", "認知症対応（認知症基本法R6.1施行対応、独立章化、本人意見聴取の制度化）"));
S.push(numItem("②", "移動支援3層構造（R7.3タクシー助成終了→社協NPO移送＋デマンドバス＋町民バス）"));
S.push(numItem("③", "在宅生活継続（8050問題、老老介護、家族介護者支援、町外医療連携負担）"));
S.push(numItem("④", "施設広域連携・人材確保（住所地特例24人、施設地域偏在、介護人材不足）"));
S.push(numItem("⑤", "広域医療連携（みやぎ県南中核病院・刈田綜合病院、国保川崎病院ハブ化）"));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: "EAD5F0" },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "9333B0" } },
  children: [text("　▌ 認知症基本法対応（最重要新規論点）", { size: 22, bold: true, color: "9333B0" })],
}));

S.push(p("令和6年1月施行の認知症基本法に基づき、第10期計画では認知症施策を独立章（第6章）として位置付け。基本目標6を新設し、基本法第15条〜21条の7基本的施策に対応した町施策体系（重点5本柱J-1〜J-5）を構築。チームオレンジ整備・本人ミーティング新設は本町初の新規取組です。"));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 同時策定3計画（重要前提）", { size: 22, bold: true, color: C.navy })],
}));

S.push(p("本計画と同時期に、地域福祉計画・障害者計画の2計画が並行策定中。後者2計画は別ベンダー（ジャパン総研）が策定担当。重層的支援体制・移動支援連動・データ共有等で相互整合を図る必要があります。第8章では3計画整合の枠組みを論じています。"));

// ===========================================================
// Section 3: 完成成果物一覧
// ===========================================================
S.push(...chapterCompact(3, "完成成果物一覧（5カテゴリ）", C.navy));

// A群
S.push(new Paragraph({
  spacing: { before: 240, after: 120, line: 320 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C.navy } },
  children: [text("A群：受託直後の整備済成果物", { size: 24, bold: true, color: C.navy })],
}));

const aTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("No", { width: 6 }),
      thcell("成果物名", { width: 42 }),
      thcell("内容", { width: 36 }),
      thcell("形式", { width: 16 }),
    ]}),
    new TableRow({ children: [
      tcell("A1", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("R8_川崎町_キックオフ議事録_正式版", { bold: true }),
      tcell("町長挨拶含むキックオフ議事録"),
      tcell("docx/pdf", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("A2", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("F1〜F6改訂版 Ver.2.0", { bold: true }),
      tcell("施策実現性評価/アンケート設問/実施支援/スケジュール案/解説書/委員会書類ひな形"),
      tcell("xlsx/docx", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("A3", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("実績データ確認サマリー", { bold: true }),
      tcell("6シート・R3給付費等の既知値整理"),
      tcell("xlsx", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("A4", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("第9期実績一覧_町記入用", { bold: true }),
      tcell("7シート35事業以上・町担当ご記入用"),
      tcell("xlsx", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("A5", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("アンケート集計分析テンプレート", { bold: true }),
      tcell("6シート・回収後集計用・32項目反映ガイド"),
      tcell("xlsx", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("A6", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("必要資料_入力フォーマット（MECE版）", { bold: true }),
      tcell("13シート・保険給付実績データ収集"),
      tcell("xlsx", { align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(aTable);

// B群
S.push(new Paragraph({
  spacing: { before: 240, after: 120, line: 320 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C.navy } },
  children: [text("B群：計画書素案（v1.0〜v1.6）", { size: 24, bold: true, color: C.navy })],
}));

const bTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("版", { width: 10 }),
      thcell("内容", { width: 56 }),
      thcell("頁/ブロック", { width: 14 }),
      thcell("状況", { width: 20 }),
    ]}),
    new TableRow({ children: [
      tcell("v1.0", { bold: true, align: AlignmentType.CENTER }),
      tcell("基本版（8章構成）"),
      tcell("27頁", { align: AlignmentType.CENTER }),
      tcell("旧版", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("v1.1", { bold: true, align: AlignmentType.CENTER }),
      tcell("図表詳細化版（6図表埋込）"),
      tcell("33頁", { align: AlignmentType.CENTER }),
      tcell("旧版", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("v1.2", { bold: true, align: AlignmentType.CENTER }),
      tcell("町長挨拶・目次完備版"),
      tcell("34頁", { align: AlignmentType.CENTER }),
      tcell("旧版", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("v1.3", { bold: true, align: AlignmentType.CENTER }),
      tcell("第7章詳細化版（保険料8ステップ・3パターン）"),
      tcell("40頁", { align: AlignmentType.CENTER }),
      tcell("旧版", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("v1.4", { bold: true, align: AlignmentType.CENTER }),
      tcell("第3章詳細化版（6目標別実績）"),
      tcell("43頁", { align: AlignmentType.CENTER }),
      tcell("旧版", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("v1.5", { bold: true, align: AlignmentType.CENTER }),
      tcell("カラー統一版（章カラー連動・拡張性）"),
      tcell("43頁", { align: AlignmentType.CENTER }),
      tcell("旧版", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("v1.6", { bold: true, align: AlignmentType.CENTER, fill: C.lorange }),
      tcell("本文充実版（金ヶ崎町水準・5章5-3〜5-5/6章6-3拡張）", { bold: true, color: C.orange }),
      tcell("44頁/411ブロック", { align: AlignmentType.CENTER, bold: true, color: C.orange }),
      tcell("最新", { bold: true, color: C.red, align: AlignmentType.CENTER, fill: C.lred }),
    ]}),
  ],
});
S.push(bTable);

// C群
S.push(new Paragraph({
  spacing: { before: 240, after: 120, line: 320 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C.navy } },
  children: [text("C群：町担当向け資料一式（R8.6発送対象）", { size: 24, bold: true, color: C.navy })],
}));

const cTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("No", { width: 6 }),
      thcell("成果物名", { width: 38 }),
      thcell("内容", { width: 40 }),
      thcell("頁/形式", { width: 16 }),
    ]}),
    new TableRow({ children: [
      tcell("C1", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("町担当課ご記入依頼パッケージ", { bold: true }),
      tcell("カバーレター・送付物一覧・優先順位・記入ガイド・連絡先"),
      tcell("5頁/docx", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("C2", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("作業チェックリストv2.0", { bold: true }),
      tcell("7シート・5フェーズ管理・進捗ダッシュボード・14論点管理"),
      tcell("12頁/xlsx", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("C3", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("町担当課ご説明資料", { bold: true }),
      tcell("13スライド構成・対面打合せ用・A4横・全体スケジュール含む"),
      tcell("14頁/docx", { align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(cTable);

// D群
S.push(new Paragraph({
  spacing: { before: 240, after: 120, line: 320 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C.navy } },
  children: [text("D群：第1回策定委員会資料一式（R8.8中旬使用）", { size: 24, bold: true, color: C.navy })],
}));

const dTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("No", { width: 6 }),
      thcell("成果物名", { width: 38 }),
      thcell("内容", { width: 40 }),
      thcell("頁/形式", { width: 16 }),
    ]}),
    new TableRow({ children: [
      tcell("D1", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("第1回策定委員会資料", { bold: true }),
      tcell("12頁統合版・アンケート結果報告+骨子案協議事項5項目"),
      tcell("12頁/docx", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("D2", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("第1回策定委員会想定問答集", { bold: true }),
      tcell("31問8カテゴリ・事務局用・想定外質問への対応原則含む"),
      tcell("13頁/docx", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("D3", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("計画素案概要版", { bold: true }),
      tcell("v1.5から要約・委員事前配布用・12頁・3図表埋込"),
      tcell("12頁/docx", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("D4", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("3層整合表（国→県→町）", { bold: true }),
      tcell("7シート・基本理念/重点施策/認知症/保険料/推進体制の整合"),
      tcell("14頁/xlsx", { align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(dTable);

// E群
S.push(new Paragraph({
  spacing: { before: 240, after: 120, line: 320 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C.navy } },
  children: [text("E群：Phase2/Phase3準備資料（委員会後使用）", { size: 24, bold: true, color: C.navy })],
}));

const eTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("No", { width: 6 }),
      thcell("成果物名", { width: 38 }),
      thcell("内容", { width: 40 }),
      thcell("頁/形式", { width: 16 }),
    ]}),
    new TableRow({ children: [
      tcell("E1", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("委員意見反映管理シート", { bold: true }),
      tcell("8シート・5協議事項別・反映ステータス5段階・運用フロー"),
      tcell("8シート/xlsx", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("E2", { bold: true, align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("Phase2作業計画詳細書", { bold: true }),
      tcell("R8.8〜R8.11週次マイルストーン・Red Team検証6観点・リスク7"),
      tcell("10頁/docx", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("E3", { bold: true, align: AlignmentType.CENTER, fill: C.lorange }),
      tcell("保険料試算ワークブック", { bold: true, color: C.orange }),
      tcell("10シート・8ステップ・3パターン・9/13段階・129数式・IFERROR処理", { color: C.orange }),
      tcell("10シート/xlsx", { align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(eTable);

// ===========================================================
// Section 4: スケジュールと現在地
// ===========================================================
S.push(...chapterCompact(4, "プロジェクトスケジュールと現在地", C.navy));

S.push(...infoBox("現在地",
  "Phase 1完了直前・第1回策定委員会前。計画素案v1.6本文充実版完成済。町担当（大宮様）への送付パッケージ整備済。第1回委員会資料一式整備済。アンケート発送はR8.6下旬予定。基金残高R8.6時点の確定値が町担当課で集計中。"));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 全体タイムライン（R7.11〜R9.3）", { size: 22, bold: true, color: C.navy })],
}));

const tl = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("時期", { width: 16 }),
      thcell("マイルストーン", { width: 32 }),
      thcell("対応事項", { width: 36 }),
      thcell("状況", { width: 16 }),
    ]}),
    new TableRow({ children: [
      tcell("R7.11", { bold: true, fill: C.lblue, align: AlignmentType.CENTER }),
      tcell("キックオフ会議", { bold: true }),
      tcell("町長挨拶・体制確認・スケジュール確認"),
      tcell("完了", { bold: true, color: C.green, align: AlignmentType.CENTER, fill: C.lgreen }),
    ]}),
    new TableRow({ children: [
      tcell("R8.1〜4", { bold: true, fill: C.lblue, align: AlignmentType.CENTER }),
      tcell("F1〜F6改訂版整備", { bold: true }),
      tcell("施策実現性評価・アンケート設計・委員会書類ひな形"),
      tcell("完了", { bold: true, color: C.green, align: AlignmentType.CENTER, fill: C.lgreen }),
    ]}),
    new TableRow({ children: [
      tcell("R8.5〜6", { bold: true, fill: C.lblue, align: AlignmentType.CENTER }),
      tcell("計画素案v1.0〜v1.6", { bold: true }),
      tcell("8章構成・図表埋込・章カラー統一・本文充実"),
      tcell("完了", { bold: true, color: C.green, align: AlignmentType.CENTER, fill: C.lgreen }),
    ]}),
    new TableRow({ children: [
      tcell("R8.6中旬", { bold: true, fill: C.lblue, align: AlignmentType.CENTER }),
      tcell("町送付パッケージ完成", { bold: true }),
      tcell("C群・D群・E群成果物整備完了"),
      tcell("完了", { bold: true, color: C.green, align: AlignmentType.CENTER, fill: C.lgreen }),
    ]}),
    new TableRow({ children: [
      tcell("R8.6下旬", { bold: true, fill: C.lorange, align: AlignmentType.CENTER }),
      tcell("町送付・アンケート発送", { bold: true, color: C.orange }),
      tcell("町担当課に資料一式発送、アンケート1,300名発送"),
      tcell("進行中", { bold: true, color: C.orange, align: AlignmentType.CENTER, fill: C.lorange }),
    ]}),
    new TableRow({ children: [
      tcell("R8.6", { bold: true, fill: C.lred, align: AlignmentType.CENTER }),
      tcell("基金残高確定", { bold: true, color: C.red }),
      tcell("町担当課より基金残高R8.6時点の確定値を入手", { color: C.red }),
      tcell("保留", { bold: true, color: C.red, align: AlignmentType.CENTER, fill: C.lred }),
    ]}),
    new TableRow({ children: [
      tcell("R8.7末", { bold: true, fill: C.lblue, align: AlignmentType.CENTER }),
      tcell("アンケート回収", { bold: true }),
      tcell("町担当課が回収・督促対応"),
      tcell("未着手", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("R8.8上旬", { bold: true, fill: C.lblue, align: AlignmentType.CENTER }),
      tcell("アンケート集計", { bold: true }),
      tcell("弊社で単純集計・クロス集計・自由記述分類"),
      tcell("未着手", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("R8.8中旬", { bold: true, fill: C.lgreen, align: AlignmentType.CENTER }),
      tcell("【第1回策定委員会】", { bold: true, color: C.green }),
      tcell("計画素案v1.6・アンケート結果報告・骨子案協議事項5"),
      tcell("未着手", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("R8.9〜10", { bold: true, fill: C.lblue, align: AlignmentType.CENTER }),
      tcell("Ver.2.0更新作業", { bold: true }),
      tcell("委員意見+アンケート結果反映・見込量精緻化"),
      tcell("未着手", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("R8.11", { bold: true, fill: C.lgreen, align: AlignmentType.CENTER }),
      tcell("【第2回策定委員会】", { bold: true, color: C.green }),
      tcell("Ver.2.0審議・サービス見込量方向性確定"),
      tcell("未着手", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("R8.12〜R9.1", { bold: true, fill: C.lblue, align: AlignmentType.CENTER }),
      tcell("保険料試算精緻化", { bold: true }),
      tcell("3パターン精緻化・9/13段階比較・近隣比較"),
      tcell("未着手", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("R9.1中旬", { bold: true, fill: C.lgreen, align: AlignmentType.CENTER }),
      tcell("【第3回策定委員会】", { bold: true, color: C.green }),
      tcell("保険料3パターン協議・パターン選定"),
      tcell("未着手", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("R9.2", { bold: true, fill: C.lgreen, align: AlignmentType.CENTER }),
      tcell("【第4回委員会・町長答申】", { bold: true, color: C.green }),
      tcell("保険料基準額決定・最終素案承認"),
      tcell("未着手", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("R9.2〜3", { bold: true, fill: C.lblue, align: AlignmentType.CENTER }),
      tcell("パブコメ・3月議会上程", { bold: true }),
      tcell("住民意見募集・条例改正案上程・計画書公表"),
      tcell("未着手", { align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(tl);

// ===========================================================
// Section 5: 残作業と次セッション優先候補
// ===========================================================
S.push(...chapterCompact(5, "残作業と次セッションでの優先候補", C.navy));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lorange },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.orange } },
  children: [text("　▌ 直近R8.6〜R8.8の作業（緊急度A）", { size: 22, bold: true, color: C.orange })],
}));

S.push(numItem("①", "町送付パッケージのR8.6下旬発送：作成済成果物（C1〜C3＋計画素案v1.6＋A群フォーマット）を町担当（大宮様）に発送する手配。送付確認の最終チェック。"));
S.push(numItem("②", "アンケート発送支援：町担当によるR8.6下旬発送のサポート、回収率向上策の準備。"));
S.push(numItem("③", "基金残高R8.6時点確定値の入手：町担当課への再依頼。基金残高判明後、保険料試算ワークブックE3に入力可能。"));
S.push(numItem("④", "委員名簿の準備：第1回策定委員会用の委員名簿（学識経験者・医療関係者・社協・民生委員・住民代表等）の整備依頼。"));
S.push(numItem("⑤", "町長挨拶文の最終確定：計画書冒頭の挨拶文・町長氏名の本人確認依頼。"));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 第1回委員会後R8.8末〜R8.11の作業（Phase2）", { size: 22, bold: true, color: C.navy })],
}));

S.push(numItem("⑥", "アンケート集計の実施（R8.8上旬・約2週間）：単純集計19+16問・クロス集計5本・自由記述5分類。E3アンケート集計分析テンプレートを使用。"));
S.push(numItem("⑦", "委員意見の集約・反映方針確定（E1管理シート使用）：3者協議（事務局・町担当・弊社）でR8.9上旬までに対応方針確定。"));
S.push(numItem("⑧", "計画素案Ver.2.0更新（R8.9〜R8.10末）：第5章KPI数値設定・第6章認知症3層KPI・第7章サービス見込量精緻化・委員意見反映。E2 Phase2作業計画詳細書のW1-W14に従って進行。"));
S.push(numItem("⑨", "第2回策定委員会資料作成（R8.11上旬）：Ver.1.5/v1.6→Ver.2.0変更点一覧・サービス見込量資料・想定問答更新版。"));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lgreen },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.green } },
  children: [text("　▌ Phase3〜Phase4の作業（R8.11〜R9.3）", { size: 22, bold: true, color: C.green })],
}));

S.push(numItem("⑩", "保険料試算精緻化（R8.12〜R9.1）：E3保険料試算ワークブックに基金残高・収納率・見込量を入力。3パターン×9/13段階×近隣比較を確定。"));
S.push(numItem("⑪", "第3回・第4回策定委員会資料作成・運営（R9.1中旬・R9.2）：保険料協議・町長答申資料の準備。"));
S.push(numItem("⑫", "パブリックコメント実施支援（R9.2）：実施要領・住民周知資料・意見集約・反映検討。"));
S.push(numItem("⑬", "条例改正案作成支援（R9.3）：介護保険条例の保険料率・段階区分の改正案作成支援。"));
S.push(numItem("⑭", "計画書本編完成版・概要版（住民配布用）の作成（R9.3）：印刷データ作成・町HP掲載。"));

// ===========================================================
// Section 6: 技術的注意事項
// ===========================================================
S.push(...chapterCompact(6, "技術的注意事項（環境・スクリプト）", C.navy));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 作業ディレクトリと環境", { size: 22, bold: true, color: C.navy })],
}));

S.push(bullet("作業ディレクトリ：/home/claude/kawasaki_work/"));
S.push(bullet("出力先：/mnt/user-data/outputs/（永続化）"));
S.push(bullet("セッションリセットで /home/claude/ は消失するため、スクリプトは毎回再生成。出力済成果物は永続化される。"));
S.push(bullet("Node.js実行時：`export NODE_PATH=$(npm root -g)` を事前実行必須"));
S.push(bullet("日本語フォント：游ゴシック（見出し）・游明朝（本文）・FONT定数で統一"));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ ファイル構成（主要スクリプト）", { size: 22, bold: true, color: C.navy })],
}));

const filesTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("ファイル", { width: 30 }),
      thcell("用途", { width: 50 }),
      thcell("対応成果物", { width: 20 }),
    ]}),
    new TableRow({ children: [
      tcell("plan_helpers.js", { bold: true, fill: C.lblue }),
      tcell("共通ヘルパー・カラー定義・章コンテキスト記憶機能"),
      tcell("計画書v1.6"),
    ]}),
    new TableRow({ children: [
      tcell("plan_part1〜8.js", { bold: true, fill: C.lblue }),
      tcell("計画書本体（part1=Ch1-4、part2=Ch5-6、part7=Ch7、part8=Ch8）"),
      tcell("計画書v1.6"),
    ]}),
    new TableRow({ children: [
      tcell("build_plan.js", { bold: true, fill: C.lblue }),
      tcell("計画書統合・出力"),
      tcell("計画書v1.6"),
    ]}),
    new TableRow({ children: [
      tcell("build_a1_package.js", { bold: true, fill: C.lblue }),
      tcell("町送付パッケージ"),
      tcell("C1"),
    ]}),
    new TableRow({ children: [
      tcell("build_a3_meeting.js", { bold: true, fill: C.lblue }),
      tcell("町担当ご説明資料"),
      tcell("C3"),
    ]}),
    new TableRow({ children: [
      tcell("build_iinkai_1.js", { bold: true, fill: C.lblue }),
      tcell("第1回策定委員会資料"),
      tcell("D1"),
    ]}),
    new TableRow({ children: [
      tcell("build_qa.js", { bold: true, fill: C.lblue }),
      tcell("想定問答集"),
      tcell("D2"),
    ]}),
    new TableRow({ children: [
      tcell("build_summary.js", { bold: true, fill: C.lblue }),
      tcell("計画素案概要版"),
      tcell("D3"),
    ]}),
    new TableRow({ children: [
      tcell("build_phase2_plan.js", { bold: true, fill: C.lblue }),
      tcell("Phase2作業計画詳細書"),
      tcell("E2"),
    ]}),
    new TableRow({ children: [
      tcell("create_charts.py", { bold: true, fill: C.lblue }),
      tcell("6図表生成（matplotlib）"),
      tcell("計画書v1.6"),
    ]}),
    new TableRow({ children: [
      tcell("create_alignment_table.py", { bold: true, fill: C.lblue }),
      tcell("3層整合表"),
      tcell("D4"),
    ]}),
    new TableRow({ children: [
      tcell("create_opinion_mgmt.py", { bold: true, fill: C.lblue }),
      tcell("委員意見反映管理シート"),
      tcell("E1"),
    ]}),
    new TableRow({ children: [
      tcell("create_premium_workbook.py", { bold: true, fill: C.lblue }),
      tcell("保険料試算ワークブック"),
      tcell("E3"),
    ]}),
  ],
});
S.push(filesTable);

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lorange },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.orange } },
  children: [text("　▌ 重要な技術的注意点（既知の制約）", { size: 22, bold: true, color: C.orange })],
}));

S.push(bullet("docxライブラリの段落ボーダー：top/bottomのみ有効。left/rightはvalidation失敗の原因。border中のtop/left/bottom/right順序にも注意（OOXMLスキーマ要件）。"));
S.push(bullet("openpyxlでシート名にハイフン使用不可（数式エラー原因）、アンダースコアで統一。"));
S.push(bullet("openpyxlのcell.fill = Noneはエラー。条件付き代入を使用。"));
S.push(bullet("Excel数式内のシート参照でシングルクォートとf-stringが衝突する場合、ダブルクォート構築で対応。IFERROR処理で未確定値を「未算定」表示にする。"));
S.push(bullet("章コンテキスト機能（_currentChapter）：chapterTitle(N)後はsection/subsectionで章番号省略可。plan_helpers.jsで実装済み。"));
S.push(bullet("章カラー：CH定数で定義。第1章ネイビー・第2章グリーン・第3章ブルー・第4章オレンジ・第5章パープル・第6章マゼンタ・第7章ティール・第8章チャコール。"));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 標準ビルドフロー", { size: 22, bold: true, color: C.navy })],
}));

S.push(numItem("①", "`cd /home/claude/kawasaki_work && export NODE_PATH=$(npm root -g)` （環境変数設定）"));
S.push(numItem("②", "`node build_xxx.js` または `python create_xxx.py` （成果物生成）"));
S.push(numItem("③", "`python /mnt/skills/public/docx/scripts/office/validate.py xxx.docx` （docxバリデーション）"));
S.push(numItem("④", "`python /mnt/skills/public/xlsx/scripts/recalc.py xxx.xlsx` （xlsx数式検証）"));
S.push(numItem("⑤", "`python /mnt/skills/public/docx/scripts/office/soffice.py --headless --convert-to pdf xxx.docx` （PDF変換）"));
S.push(numItem("⑥", "`pdftoppm -jpeg -r 75 -f N -l N xxx.pdf /tmp/preview` （視覚確認）"));
S.push(numItem("⑦", "`cp xxx /mnt/user-data/outputs/` （永続化）"));
S.push(numItem("⑧", "`present_files` で提示"));

// ===========================================================
// Section 7: 推奨アップロードファイル一覧
// ===========================================================
S.push(...chapterCompact(7, "推奨アップロードファイル一覧", C.navy));

S.push(...infoBox("次セッション開始時にアップロードすべきファイル",
  "次セッションで作業を継続する際は、以下のファイルをアップロードすることを推奨します。本引継ぎメモと併せてアップロードすることで、Claudeが状況を迅速に把握できます。",
  C.orange));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lorange },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.orange } },
  children: [text("　▌ 必須アップロード（最低限）", { size: 22, bold: true, color: C.orange })],
}));

S.push(numItem("①", "本引継ぎメモ（川崎町_引継ぎメモ.docx または .pdf）"));
S.push(numItem("②", "計画素案v1.6 本文充実版（川崎町_計画書素案_v1.6_本文充実版.docx）"));
S.push(numItem("③", "作業チェックリストv2.0（川崎町_作業チェックリスト_v2.0.xlsx・進捗状況更新版）"));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 作業内容別の推奨アップロード", { size: 22, bold: true, color: C.navy })],
}));

const uplTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("作業内容", { width: 30 }),
      thcell("必要なファイル", { width: 50 }),
      thcell("対応No", { width: 20 }),
    ]}),
    new TableRow({ children: [
      tcell("町送付パッケージの追加修正", { bold: true, fill: C.lblue }),
      tcell("町担当課ご記入依頼パッケージ、計画素案v1.6"),
      tcell("C1, B(v1.6)", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("アンケート集計実施", { bold: true, fill: C.lblue }),
      tcell("回収アンケート原本、アンケート集計分析テンプレート"),
      tcell("A5", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("町確認データ反映", { bold: true, fill: C.lblue }),
      tcell("第9期実績一覧記入完了版、MECEデータ入力フォーマット記入完了版"),
      tcell("A4, A6", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("計画素案Ver.2.0更新", { bold: true, fill: C.lblue }),
      tcell("計画素案v1.6、第1回委員会議事録、委員意見反映管理シート（記入版）"),
      tcell("B(v1.6), E1", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("第2回委員会資料作成", { bold: true, fill: C.lblue }),
      tcell("計画素案Ver.2.0、Phase2作業計画詳細書、想定問答集（更新元）"),
      tcell("E2, D2", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("保険料試算精緻化", { bold: true, fill: C.lorange }),
      tcell("保険料試算ワークブック、基金残高R8.6確定値、過去5年収納率データ", { color: C.orange }),
      tcell("E3", { align: AlignmentType.CENTER, color: C.orange, bold: true }),
    ]}),
    new TableRow({ children: [
      tcell("パブリックコメント実施", { bold: true, fill: C.lblue }),
      tcell("計画素案最終版、計画書概要版"),
      tcell("計画素案最終, D3更新", { align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(uplTable);

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 既存uploads（過去セッションで継続利用）", { size: 22, bold: true, color: C.navy })],
}));

S.push(p("過去セッションでアップロード済の以下ファイルは、次セッションでも継続利用が想定されます。"));
S.push(bullet("R8_川崎町_打合せ議事録.xlsx（過去打合せ記録）"));
S.push(bullet("_変更後様式_保険者テ_ータ_202506川崎町.xlsx（R7.6保険者データ）"));
S.push(bullet("年報テ_ータ_2021_川崎町_修正_.xlsx（R3年報データ）"));
S.push(bullet("川崎町41.pdf、川崎町83pdf.pdf、川崎町84.pdf（町スキャン文書）"));
S.push(bullet("テンフ_レート集_計画策定支援業務全体.xlsx、引継メモ_計画策定支援業務全体.docx（全体テンプレート）"));

// ===========================================================
// Section 8: 次セッション開始時の標準フロー
// ===========================================================
S.push(...chapterCompact(8, "次セッション開始時の標準フロー", C.navy));

S.push(...infoBox("次セッション開始時の推奨指示",
  "次セッションでは、本引継ぎメモと最新の計画素案v1.6をアップロードした上で、以下のような簡潔な指示で作業を進めることが推奨されます。",
  C.orange));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 推奨指示例", { size: 22, bold: true, color: C.navy })],
}));

S.push(bullet("「引継ぎメモを確認の上、進めれる作業を提示してください」（汎用）"));
S.push(bullet("「アンケート集計に進んでください」（Phase 2移行時）"));
S.push(bullet("「保険料試算ワークブックに基金残高XXを入力して試算してください」（基金残高入手後）"));
S.push(bullet("「Ver.2.0更新作業に進んでください」（第1回委員会後）"));
S.push(bullet("「第2回委員会資料を作成してください」（Phase 2末期）"));
S.push(bullet("「金ヶ崎町と同様に〜してください」（金ヶ崎町方式踏襲指示）"));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 次セッションでの優先候補3つ", { size: 22, bold: true, color: C.navy })],
}));

S.push(numItem("①", "【最優先】基金残高R8.6確定値の入力と保険料試算3パターン精緻化：町担当課から確定値を受領次第、E3保険料試算ワークブックに入力。3パターン×9/13段階×近隣比較で具体的な保険料月額が算定される。"));

S.push(numItem("②", "【次優先】アンケート回収後の集計実施：R8.7末回収後、弊社で集計を実施。E3アンケート集計分析テンプレートを使用。32項目反映マトリクスで計画素案v1.6の【調査後設定】枠を実数値で更新。"));

S.push(numItem("③", "【継続】第1回策定委員会後のVer.2.0更新作業：E1委員意見反映管理シートとE2 Phase2作業計画詳細書を使用して、Ver.2.0素案完成までの14週間（W1〜W14）を週次マイルストーンで進める。"));

S.push(spacer());

S.push(new Paragraph({
  spacing: { before: 240, after: 240 },
  alignment: AlignmentType.CENTER,
  border: {
    top: { style: BorderStyle.DOUBLE, size: 18, color: C.navy },
    bottom: { style: BorderStyle.DOUBLE, size: 18, color: C.navy },
  },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  children: [text("引き続き、川崎町第10期計画策定の成功に向けて、円滑な作業継続を実現します。",
    { size: 22, bold: true, color: C.navy })],
}));

// ===========================================================
// ドキュメント生成
// ===========================================================
const doc = new Document({
  creator: "ビズアップ公共コンサルティング株式会社",
  title: "川崎町第10期計画 引継ぎメモ",
  description: "次セッションでの作業継続のための包括引継ぎ資料",
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
            text: "川崎町第10期介護保険事業計画策定支援 引継ぎメモ（次セッション用）",
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
  fs.writeFileSync("/home/claude/kawasaki_work/川崎町_引継ぎメモ.docx", buffer);
  console.log("Build done. Blocks:", S.length);
}).catch(err => {
  console.error("Build error:", err);
  process.exit(1);
});
