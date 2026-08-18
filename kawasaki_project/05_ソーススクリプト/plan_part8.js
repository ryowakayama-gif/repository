/**
 * 川崎町第10期計画書素案 - 第8章
 * 計画の推進体制と評価
 */
const H = require('./plan_helpers');
const {
  C, CH, FONT,
  text, p, chapterTitle, section, subsection,
  bullet, numItem, placeholder, fact, source,
  tcell, thcell, kvTable, spacer, image, caption,
  Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  convertInchesToTwip,
} = H;

const S = [];


// 第8章 計画の推進体制と評価
// ===========================================================
S.push(...chapterTitle(8));

S.push(p("本計画を着実に推進し、その効果を最大化するため、推進体制・進行管理・評価の仕組みを以下のとおり整備します。"));

// 8-1
S.push(section(8, 1, "推進体制"));

S.push(subsection("町の体制", 8));
S.push(p("本計画は、川崎町保健福祉課（介護保険係）を主担当課とし、町関係課（地域振興課・町民生活課・財政課等）の連携のもとに推進します。"));

const taiseiTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("組織", { width: 30, fill: CH[8].main }),
      thcell("役割", { width: 70, fill: CH[8].main }),
    ]}),
    new TableRow({ children: [
      tcell("川崎町保健福祉課（介護保険係）", { bold: true, fill: C.lgray }),
      tcell("計画の総合的推進、進行管理、介護保険事業の運営"),
    ]}),
    new TableRow({ children: [
      tcell("地域振興課・町民生活課", { bold: true, fill: C.lgray }),
      tcell("移動支援（町民バス・デマンドバス）の所管との連携"),
    ]}),
    new TableRow({ children: [
      tcell("川崎町地域包括支援センター（社協運営）", { bold: true, fill: C.lgray }),
      tcell("総合相談、介護予防ケアマネジメント、地域ケア会議の運営"),
    ]}),
    new TableRow({ children: [
      tcell("国民健康保険川崎病院", { bold: true, fill: C.lgray }),
      tcell("町内医療拠点としての在宅医療・介護連携、認知症診療"),
    ]}),
    new TableRow({ children: [
      tcell("川崎町社会福祉協議会", { bold: true, fill: C.lgray }),
      tcell("地域包括支援センター運営、福祉移送サービス、ユニバーサルサポーター制度"),
    ]}),
    new TableRow({ children: [
      tcell("民生委員・町内会・サポーター", { bold: true, fill: C.lgray }),
      tcell("地域見守り、ふれあいネットワーク活動、各種サポーター活動"),
    ]}),
    new TableRow({ children: [
      tcell("広域連携先（みやぎ県南中核病院・刈田綜合病院 等）", { bold: true, fill: C.lgray }),
      tcell("夜間救急・専門医療における広域連携"),
    ]}),
  ],
});
S.push(taiseiTable);

// 8-2
S.push(section(8, 2, "進行管理（PDCAサイクル）"));

S.push(p("本計画の進行管理は、計画（Plan）・実行（Do）・評価（Check）・改善（Action）のPDCAサイクルに基づき、毎年度の進捗を確認しながら推進します。"));

const pdcaTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("段階", { width: 14, fill: CH[8].main }),
      thcell("実施内容", { width: 56, fill: CH[8].main }),
      thcell("時期", { width: 30, fill: CH[8].main }),
    ]}),
    new TableRow({ children: [
      tcell("Plan\n（計画）", { bold: true, fill: C.lgray, align: AlignmentType.CENTER }),
      tcell("計画期間の事業計画・予算編成・KPI目標値設定"),
      tcell("計画期間開始前", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("Do\n（実行）", { bold: true, fill: C.lgray, align: AlignmentType.CENTER }),
      tcell("各施策・事業の実施"),
      tcell("計画期間中（通年）", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("Check\n（評価）", { bold: true, fill: C.lgray, align: AlignmentType.CENTER }),
      tcell("KPIの進捗確認、事業実績の評価、改善点の抽出"),
      tcell("毎年度末", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("Action\n（改善）", { bold: true, fill: C.lgray, align: AlignmentType.CENTER }),
      tcell("評価結果を踏まえた次年度事業計画への反映"),
      tcell("次年度予算編成時", { align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(pdcaTable);

// 8-3
S.push(section(8, 3, "計画の評価と見直し"));

S.push(subsection("評価の仕組み", 8));
S.push(p("計画の評価は以下の仕組みで実施します。"));
S.push(numItem("①", "毎年度の進捗評価：保健福祉課が事業実績・KPI進捗を取りまとめ、町長へ報告"));
S.push(numItem("②", "中間評価：計画期間の中間年（R10年度）に、策定委員会または同等の組織で中間評価"));
S.push(numItem("③", "計画期間終了時評価：第10期計画終了時（R11年度）に、第11期計画策定の基礎資料として総括評価"));

S.push(subsection("計画の見直し", 8));
S.push(p("計画期間中に、国の制度改正・大きな社会経済情勢の変化・町の重大な課題変化が生じた場合は、本計画の見直しを検討します。見直しの検討は策定委員会または同等の組織で行います。"));

S.push(subsection("情報公開と住民参画", 8));
S.push(p("本計画の推進状況は、町ホームページ・町広報誌等で住民に積極的に公開します。また、住民意見の継続的な反映のため、必要に応じてアンケート・地区説明会等を実施します。"));

// 結びの段落
S.push(spacer());
S.push(spacer());
S.push(new Paragraph({
  spacing: { before: 480, after: 240, line: 360 },
  alignment: AlignmentType.CENTER,
  border: { top: { style: BorderStyle.DOUBLE, size: 18, color: C.navy } },
  children: [text("─ 本素案 Ver.1.0 結 ─", { size: 22, bold: true, color: C.navy })],
}));
S.push(spacer());
S.push(p("本素案v1.0は、令和8年7月末のアンケート回収・8月の集計分析・国通知（夏以降）の確認を経て、Ver.2.0として更新します。Ver.2.0は令和8年11月の第2回策定委員会で素案として審議いただく予定です。",
  { size: 18, italics: true, color: C.gray, alignment: AlignmentType.CENTER, noIndent: true }));

module.exports = { S };
console.log("Part8 (Ch8) ready. Blocks:", S.length);
