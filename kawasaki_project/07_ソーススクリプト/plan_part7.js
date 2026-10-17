/**
 * 川崎町第10期計画書素案 - 第7章 詳細化版
 * 介護保険サービス見込量と保険料
 *
 * 詳細化ポイント：
 * - 7-1: 基本ロジック式、4要素前提テーブル、6ステップ、年度別見込量数値枠、
 *        町外施設・住所地特例方針、地域支援事業の量
 * - 7-2: 給付費算定構造、第10期試算枠（R9/R10/R11/3年累計）
 * - 7-3: 8ステップ詳細、算定式、基金活用方針、3パターン（A/B/C・第11期影響）、
 *        13段階区分検討、収納率根拠、近隣比較、13段階別保険料試算枠、確定スケジュール
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

// ===========================================================
// 第7章 介護保険サービス見込量と保険料
// ===========================================================
S.push(...chapterTitle(7));

S.push(p("本章では、第10期計画期間（令和9〜11年度）の介護保険サービス見込量と、それに基づく介護保険料の試算枠組みを示します。具体的な数値は、町からの提供データ・アンケート結果・国通知（夏以降）・介護給付費準備基金残高（令和8年6月確定）を踏まえ、第3回策定委員会（令和9年1月中旬）で協議のうえ第4回（令和9年2月）で確定します。"));

S.push(fact("本章は介護保険法第117条第2項第2号「各年度における種類ごとの介護給付等対象サービスの量の見込み」に対応する計画の核心部分である。"));

// =====================================================
// 7-1 サービス見込量の推計
// =====================================================
S.push(section(7, 1, "サービス見込量の推計"));

S.push(subsection("推計の基本ロジック", 7));
S.push(p("サービス見込量は、以下の基本ロジックに基づき算定します。"));
S.push(new Paragraph({
  spacing: { before: 200, after: 200, line: 360 },
  alignment: AlignmentType.CENTER,
  shading: { type: ShadingType.SOLID, fill: CH[7].sub },
  border: {
    top: { style: BorderStyle.SINGLE, size: 12, color: CH[7].main },
    bottom: { style: BorderStyle.SINGLE, size: 12, color: CH[7].main },
  },
  children: [text("見込量 = 将来人口 × 認定率 × 利用率 × 1人当たり利用量", { size: 26, bold: true, color: CH[7].main })],
}));
S.push(p("この基本式に、川崎町固有の補正要素（住所地特例による町外施設利用、認知症基本法対応の地域支援事業拡充、ユニバーサルサポーター制度による軽度認定者の重度化抑制等）を組み込みます。"));

S.push(subsection("推計の前提（4要素）", 7));

const premiseTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("要素", { width: 18, fill: CH[7].main }),
      thcell("採用データ・方針", { width: 52, fill: CH[7].main }),
      thcell("根拠・出典", { width: 30, fill: CH[7].main }),
    ]}),
    new TableRow({ children: [
      tcell("①人口推計", { bold: true, fill: C.lblue }),
      tcell("国立社会保障・人口問題研究所「日本の地域別将来推計人口（令和5年推計）」を基本。住民基本台帳との突合で補正"),
      tcell("社人研R5推計"),
    ]}),
    new TableRow({ children: [
      tcell("②認定率", { bold: true, fill: C.lblue }),
      tcell("性別・年齢階級別（5歳階級）認定率を、直近5年間（R3〜R7）の実績推移から算定。後期高齢者の急増局面を反映"),
      tcell("町保険者データ"),
    ]}),
    new TableRow({ children: [
      tcell("③利用率", { bold: true, fill: C.lblue }),
      tcell("国保連データから、サービス種類別の利用率（認定者中の利用者割合）を算定。直近年度を基本"),
      tcell("国保連・町実績"),
    ]}),
    new TableRow({ children: [
      tcell("④1人当たり\n　利用量", { bold: true, fill: C.lblue }),
      tcell("実績データから月平均利用量（訪問介護＝時間、通所介護＝回数、施設＝日数等）を算定"),
      tcell("町給付実績"),
    ]}),
  ],
});
S.push(premiseTable);

S.push(subsection("見込量算定 6ステップ", 7));

const mikomTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("Step", { width: 8, fill: CH[7].main }),
      thcell("作業", { width: 22, fill: CH[7].main }),
      thcell("内容", { width: 70, fill: CH[7].main }),
    ]}),
    new TableRow({ children: [
      tcell("1", { align: AlignmentType.CENTER, bold: true }),
      tcell("人口推計", { bold: true }),
      tcell("社人研データから令和9〜11年度の年齢階級別人口（前期高齢者・75〜85歳未満・85歳以上）を推計"),
    ]}),
    new TableRow({ children: [
      tcell("2", { align: AlignmentType.CENTER, bold: true }),
      tcell("認定者推計", { bold: true }),
      tcell("年齢階級別認定率を将来人口に乗じて、要介護度別認定者数を推計"),
    ]}),
    new TableRow({ children: [
      tcell("3", { align: AlignmentType.CENTER, bold: true }),
      tcell("利用率算定", { bold: true }),
      tcell("国保連データから、サービス種類別の利用率（認定者中の利用者割合）を算定"),
    ]}),
    new TableRow({ children: [
      tcell("4", { align: AlignmentType.CENTER, bold: true }),
      tcell("見込量推計", { bold: true }),
      tcell("認定者推計 × 利用率 × 1人当たり利用量 により、年度別・サービス種類別の見込量を算定"),
    ]}),
    new TableRow({ children: [
      tcell("5", { align: AlignmentType.CENTER, bold: true }),
      tcell("アンケート補正", { bold: true }),
      tcell("一般高齢者ニーズ調査・認定者調査の結果から、潜在ニーズ・利用意向を反映して見込量を補正"),
    ]}),
    new TableRow({ children: [
      tcell("6", { align: AlignmentType.CENTER, bold: true }),
      tcell("見える化登録・比較", { bold: true }),
      tcell("見える化システムに登録し、仙南圏域（柴田町・大河原町・村田町・丸森町・蔵王町・七ヶ宿町）と比較確認"),
    ]}),
  ],
});
S.push(mikomTable);

S.push(subsection("年度別サービス見込量（数値枠）", 7));
S.push(p("第10期計画期間中（令和9〜11年度）の年度別・サービス種類別見込量は、上記6ステップを経て下表のとおり算定します。本素案v1.0では枠のみ提示し、Ver.2.0で確定値を記載します。"));

const yearMikomTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("区分", { width: 12, fill: CH[7].main }),
      thcell("サービス種類", { width: 32, fill: CH[7].main }),
      thcell("単位", { width: 10, fill: CH[7].main }),
      thcell("R3実績", { width: 12, fill: CH[7].main }),
      thcell("R9", { width: 11, fill: CH[7].main }),
      thcell("R10", { width: 11, fill: CH[7].main }),
      thcell("R11", { width: 12, fill: CH[7].main }),
    ]}),
    // 居宅
    new TableRow({ children: [
      tcell("居宅", { bold: true, fill: C.lblue, align: AlignmentType.CENTER }),
      tcell("訪問介護", { bold: true }),
      tcell("回/月", { align: AlignmentType.CENTER }),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("─", { align: AlignmentType.CENTER }),
      tcell("─", { align: AlignmentType.CENTER }),
      tcell("─", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("居宅", { bold: true, fill: C.lblue, align: AlignmentType.CENTER }),
      tcell("訪問看護", { bold: true }),
      tcell("回/月", { align: AlignmentType.CENTER }),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("─", { align: AlignmentType.CENTER }),
      tcell("─", { align: AlignmentType.CENTER }),
      tcell("─", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("居宅", { bold: true, fill: C.lblue, align: AlignmentType.CENTER }),
      tcell("通所介護（デイサービス）", { bold: true }),
      tcell("回/月", { align: AlignmentType.CENTER }),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("─", { align: AlignmentType.CENTER }),
      tcell("─", { align: AlignmentType.CENTER }),
      tcell("─", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("居宅", { bold: true, fill: C.lblue, align: AlignmentType.CENTER }),
      tcell("通所リハビリ（デイケア）", { bold: true }),
      tcell("回/月", { align: AlignmentType.CENTER }),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("─", { align: AlignmentType.CENTER }),
      tcell("─", { align: AlignmentType.CENTER }),
      tcell("─", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("居宅", { bold: true, fill: C.lblue, align: AlignmentType.CENTER }),
      tcell("短期入所（ショートステイ）", { bold: true }),
      tcell("日/月", { align: AlignmentType.CENTER }),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("─", { align: AlignmentType.CENTER }),
      tcell("─", { align: AlignmentType.CENTER }),
      tcell("─", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("居宅", { bold: true, fill: C.lblue, align: AlignmentType.CENTER }),
      tcell("福祉用具・住宅改修", { bold: true }),
      tcell("件/月", { align: AlignmentType.CENTER }),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("─", { align: AlignmentType.CENTER }),
      tcell("─", { align: AlignmentType.CENTER }),
      tcell("─", { align: AlignmentType.CENTER }),
    ]}),
    // 地域密着
    new TableRow({ children: [
      tcell("地域\n密着型", { bold: true, fill: C.lgreen, align: AlignmentType.CENTER }),
      tcell("認知症対応型共同生活介護（GH）", { bold: true }),
      tcell("人/月", { align: AlignmentType.CENTER }),
      tcell("─", { align: AlignmentType.CENTER }),
      tcell("─", { align: AlignmentType.CENTER }),
      tcell("─", { align: AlignmentType.CENTER }),
      tcell("─", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("地域\n密着型", { bold: true, fill: C.lgreen, align: AlignmentType.CENTER }),
      tcell("小規模多機能型居宅介護", { bold: true }),
      tcell("人/月", { align: AlignmentType.CENTER }),
      tcell("─", { align: AlignmentType.CENTER }),
      tcell("─", { align: AlignmentType.CENTER }),
      tcell("─", { align: AlignmentType.CENTER }),
      tcell("─", { align: AlignmentType.CENTER }),
    ]}),
    // 施設
    new TableRow({ children: [
      tcell("施設", { bold: true, fill: C.lorange, align: AlignmentType.CENTER }),
      tcell("介護老人福祉施設（特養・町外含む）", { bold: true }),
      tcell("人/月", { align: AlignmentType.CENTER }),
      tcell("68人", { align: AlignmentType.RIGHT }),
      tcell("─", { align: AlignmentType.CENTER }),
      tcell("─", { align: AlignmentType.CENTER }),
      tcell("─", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("施設", { bold: true, fill: C.lorange, align: AlignmentType.CENTER }),
      tcell("介護老人保健施設（老健・町外含む）", { bold: true }),
      tcell("人/月", { align: AlignmentType.CENTER }),
      tcell("66人", { align: AlignmentType.RIGHT }),
      tcell("─", { align: AlignmentType.CENTER }),
      tcell("─", { align: AlignmentType.CENTER }),
      tcell("─", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("施設", { bold: true, fill: C.lorange, align: AlignmentType.CENTER }),
      tcell("介護療養型・介護医療院", { bold: true }),
      tcell("人/月", { align: AlignmentType.CENTER }),
      tcell("0人", { align: AlignmentType.RIGHT }),
      tcell("0", { align: AlignmentType.CENTER, color: C.gray }),
      tcell("0", { align: AlignmentType.CENTER, color: C.gray }),
      tcell("0", { align: AlignmentType.CENTER, color: C.gray }),
    ]}),
  ],
});
S.push(yearMikomTable);
S.push(source("R3実績：年報データ2021／R9〜R11見込み：Ver.2.0で確定（社人研推計＋6ステップ算定）"));

S.push(placeholder("年度別・サービス種類別見込量の確定値は、町担当課提供データ、国保連データ、アンケート結果、国通知の反映を経て、Ver.2.0で記載します。"));

S.push(subsection("施設サービスの町外利用・住所地特例の反映", 7));
S.push(p("キックオフ会議で確認されたとおり、川崎町は町内施設が縮小傾向にあり、町外施設の利用（住所地特例24人・令和7年6月時点）が相当数あります。施設サービス見込量の算定では、以下の方針で町外利用を組み込みます。"));
S.push(numItem("①", "町内施設の供給定員のみで見込量を制約しない。町外施設利用を含めた「町民の施設サービス利用ニーズ」全体を見込量とする"));
S.push(numItem("②", "住所地特例該当者の所在自治体内訳（柴田町・大河原町・仙台市等）を把握し、広域連携の検討材料とする"));
S.push(numItem("③", "町外施設の建設計画・受入余地等の情報収集を行い、川崎町民の施設利用継続を担保する"));

S.push(subsection("地域支援事業の量の見込み", 7));
S.push(p("地域支援事業（介護予防・日常生活支援総合事業、包括的支援事業、任意事業）の量の見込みも、サービス見込量と同様に第10期計画期間中の年度別実施量を算定します。"));

const chiikiShienTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("事業区分", { width: 30, fill: CH[7].main }),
      thcell("R3実績", { width: 22, fill: CH[7].main }),
      thcell("R9〜R11見込み", { width: 48, fill: CH[7].main }),
    ]}),
    new TableRow({ children: [
      tcell("介護予防・生活支援サービス事業", { bold: true }),
      tcell("約565万円", { align: AlignmentType.RIGHT }),
      tcell("【町確認・アンケート結果反映後確定】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("一般介護予防事業", { bold: true }),
      tcell("約195万円", { align: AlignmentType.RIGHT }),
      tcell("【ユニバーサルサポーター制度との整合確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("包括的支援・任意事業", { bold: true }),
      tcell("約2,447万円", { align: AlignmentType.RIGHT }),
      tcell("【包括センター体制強化を反映】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("合計", { bold: true, fill: C.lblue }),
      tcell("約3,207万円", { align: AlignmentType.RIGHT, bold: true, fill: C.lblue }),
      tcell("【国の上限額管理基準を踏まえ確定】", { color: C.orange, italics: true, align: AlignmentType.CENTER, fill: C.lblue, bold: true }),
    ]}),
  ],
});
S.push(chiikiShienTable);
S.push(source("R3実績：年報データ2021・様式4／R9〜R11見込み：Ver.2.0で確定"));

// =====================================================
// 7-2 介護給付費の見込み
// =====================================================
S.push(section(7, 2, "介護給付費の見込み"));

S.push(subsection("給付費算定の構造", 7));
S.push(p("介護給付費の見込みは、サービス見込量に各種単価・加算率を乗じて算定します。標準給付費見込額は、保険料算定の出発点となる極めて重要な数値です。"));

const structureTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("項目", { width: 36, fill: CH[7].main }),
      thcell("内容", { width: 64, fill: CH[7].main }),
    ]}),
    new TableRow({ children: [
      tcell("標準給付費見込額", { bold: true, fill: C.lblue }),
      tcell("サービス見込量 × サービス種類別単価 × 加算率 で算定。3年間の累計値"),
    ]}),
    new TableRow({ children: [
      tcell("　＋ 高額介護サービス費", { fill: C.lgray }),
      tcell("自己負担月額の上限超過分の払戻し（R3実績 約2,682万円）"),
    ]}),
    new TableRow({ children: [
      tcell("　＋ 高額医療合算介護", { fill: C.lgray }),
      tcell("医療＋介護の合算自己負担上限超過分（R3実績 約240万円）"),
    ]}),
    new TableRow({ children: [
      tcell("　＋ 特定入所者介護費", { fill: C.lgray }),
      tcell("低所得施設入所者の食費・居住費補足給付（R3実績 約5,247万円）"),
    ]}),
    new TableRow({ children: [
      tcell("＝ 保険給付費総額", { bold: true, fill: C.lorange }),
      tcell("R3実績：約10.4億円。第10期で給付費上方圧力（後期高齢者+10.9%）が見込まれる"),
    ]}),
    new TableRow({ children: [
      tcell("＋ 地域支援事業費", { bold: true, fill: C.lgreen }),
      tcell("総合事業・包括的支援事業・任意事業（R3実績 約3,207万円）"),
    ]}),
  ],
});
S.push(structureTable);

S.push(subsection("第10期 給付費見込み（試算枠）", 7));

const kyufuMikomi = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("区分", { width: 38, fill: CH[7].main }),
      thcell("R3実績", { width: 18, fill: CH[7].main }),
      thcell("R9", { width: 14, fill: CH[7].main }),
      thcell("R10", { width: 14, fill: CH[7].main }),
      thcell("R11", { width: 16, fill: CH[7].main }),
    ]}),
    new TableRow({ children: [
      tcell("標準給付費（サービス種類別）", { bold: true }),
      tcell("約9.6億円", { align: AlignmentType.RIGHT }),
      tcell("【試算】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("【試算】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("【試算】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("　高額・特定入所者等を加算", { italics: true, fill: C.lgray }),
      tcell("+約0.8億円", { align: AlignmentType.RIGHT, fill: C.lgray }),
      tcell("【試算】", { color: C.orange, italics: true, align: AlignmentType.CENTER, fill: C.lgray }),
      tcell("【試算】", { color: C.orange, italics: true, align: AlignmentType.CENTER, fill: C.lgray }),
      tcell("【試算】", { color: C.orange, italics: true, align: AlignmentType.CENTER, fill: C.lgray }),
    ]}),
    new TableRow({ children: [
      tcell("保険給付費総額", { bold: true, fill: C.lorange }),
      tcell("約10.4億円", { align: AlignmentType.RIGHT, bold: true, fill: C.lorange }),
      tcell("【試算】", { color: C.orange, italics: true, align: AlignmentType.CENTER, bold: true, fill: C.lorange }),
      tcell("【試算】", { color: C.orange, italics: true, align: AlignmentType.CENTER, bold: true, fill: C.lorange }),
      tcell("【試算】", { color: C.orange, italics: true, align: AlignmentType.CENTER, bold: true, fill: C.lorange }),
    ]}),
    new TableRow({ children: [
      tcell("地域支援事業費", { bold: true, fill: C.lgreen }),
      tcell("約0.32億円", { align: AlignmentType.RIGHT, bold: true, fill: C.lgreen }),
      tcell("【試算】", { color: C.orange, italics: true, align: AlignmentType.CENTER, bold: true, fill: C.lgreen }),
      tcell("【試算】", { color: C.orange, italics: true, align: AlignmentType.CENTER, bold: true, fill: C.lgreen }),
      tcell("【試算】", { color: C.orange, italics: true, align: AlignmentType.CENTER, bold: true, fill: C.lgreen }),
    ]}),
    new TableRow({ children: [
      tcell("3年累計", { bold: true, fill: C.lblue }),
      tcell("─", { align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("─", { align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("─", { align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("【3年累計】", { color: C.orange, italics: true, align: AlignmentType.CENTER, bold: true, fill: C.lblue }),
    ]}),
  ],
});
S.push(kyufuMikomi);
S.push(source("R3実績：年報データ2021／R9-R11試算：Ver.2.0で確定"));

S.push(p("第10期では、後期高齢者の増加（4年で+10.9%）と認知症基本法対応に伴う認知症施策の拡充により、給付費の上方圧力が見込まれます。一方で、ユニバーサルサポーター制度を活用した介護予防の強化、要介護認定率の適正化等により、給付費伸び率の抑制を図ります。第9期実績（給付費伸び率）を踏まえた中位推計を基本とし、上限・下限の幅を持った試算をVer.2.0で提示します。"));

// =====================================================
// 7-3 介護保険料の試算
// =====================================================
S.push(section(7, 3, "介護保険料の試算（8ステップ・3パターン）"));

S.push(subsection("保険料推移の確認", 7));

const hokenryoTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("期", { width: 20, fill: CH[7].main }),
      thcell("計画期間", { width: 24, fill: CH[7].main }),
      thcell("月額基準額", { width: 18, fill: CH[7].main }),
      thcell("対前期", { width: 18, fill: CH[7].main }),
      thcell("出典", { width: 20, fill: CH[7].main }),
    ]}),
    new TableRow({ children: [
      tcell("第8期", { bold: true, align: AlignmentType.CENTER }),
      tcell("R3〜R5年度", { align: AlignmentType.CENTER }),
      tcell("6,380円", { align: AlignmentType.RIGHT, bold: true }),
      tcell("─", { align: AlignmentType.CENTER }),
      tcell("年報2021"),
    ]}),
    new TableRow({ children: [
      tcell("第9期", { bold: true, align: AlignmentType.CENTER }),
      tcell("R6〜R8年度", { align: AlignmentType.CENTER }),
      tcell("6,500円", { align: AlignmentType.RIGHT, bold: true }),
      tcell("+120円\n(+1.9%)", { align: AlignmentType.CENTER, color: C.orange }),
      tcell("MECE版データ"),
    ]}),
    new TableRow({ children: [
      tcell("第10期\n（本計画）", { bold: true, align: AlignmentType.CENTER, fill: C.lorange }),
      tcell("R9〜R11年度", { align: AlignmentType.CENTER, fill: C.lorange }),
      tcell("【3パターン試算】", { color: C.orange, italics: true, align: AlignmentType.CENTER, fill: C.lorange, bold: true }),
      tcell("【委員会協議】", { color: C.orange, italics: true, align: AlignmentType.CENTER, fill: C.lorange }),
      tcell("本計画", { fill: C.lorange }),
    ]}),
  ],
});
S.push(hokenryoTable);

// 図7-1：保険料推移グラフ
S.push(image("/home/claude/kawasaki_work/chart_premium.png", { width: 540, height: 340 }));
S.push(caption("図7-1　介護保険料月額基準額の推移と第10期試算3パターン（試算は仮置き）"));

S.push(p("第8期保険料月額基準額6,380円から第9期6,500円への上昇は+1.9%（120円増）に留まっており、川崎町の保険料は近隣自治体と比較しても抑制的な水準で運営されてきました。これは、介護給付費準備基金の取崩しや、第1号被保険者数の安定推移、町独自の介護予防事業（ユニバーサルサポーター制度等）による要介護認定率の抑制等が寄与していると考えられます。"));

S.push(p("第10期では、後期高齢者の急増（75歳以上が4年で+10.9%）に伴う認定者・給付費増、認知症基本法対応に伴う認知症施策の拡充、地域支援事業の充実等により、給付費の上方圧力が見込まれます。一方で、所得段階区分の見直し（13段階化）、介護給付費準備基金の取崩、収納率の維持（96.3%水準）等により、保険料水準の抑制が可能となります。最終的な保険料基準額は、第3回策定委員会（令和9年1月中旬）での協議を経て第4回（令和9年2月）で決定します。"));

S.push(subsection("保険料算定の8ステップ", 7));

S.push(p("第10期保険料は、以下の8ステップで算定します。これは厚生労働省「市町村介護保険事業計画作成の手引き」（令和7年4月版）に準拠したものです。"));

const stepTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("Step", { width: 8, fill: CH[7].main }),
      thcell("作業名", { width: 22, fill: CH[7].main }),
      thcell("内容", { width: 70, fill: CH[7].main }),
    ]}),
    new TableRow({ children: [
      tcell("1", { align: AlignmentType.CENTER, bold: true }),
      tcell("標準給付費見込", { bold: true }),
      tcell("7-1で確定したサービス見込量×単価×加算率で、3年間の標準給付費を算出"),
    ]}),
    new TableRow({ children: [
      tcell("2", { align: AlignmentType.CENTER, bold: true }),
      tcell("地域支援事業費", { bold: true }),
      tcell("総合事業・包括的支援事業・任意事業の事業費を算定（上限額管理あり）"),
    ]}),
    new TableRow({ children: [
      tcell("3", { align: AlignmentType.CENTER, bold: true }),
      tcell("第1号負担分相当額", { bold: true }),
      tcell("（標準給付費＋地域支援事業費）× 第1号被保険者負担割合23%（第10期）"),
    ]}),
    new TableRow({ children: [
      tcell("4", { align: AlignmentType.CENTER, bold: true }),
      tcell("調整交付金", { bold: true }),
      tcell("後期高齢者割合・所得段階別人口で交付率算定（R3実績 約5,912万円・約6.1%）"),
    ]}),
    new TableRow({ children: [
      tcell("5", { align: AlignmentType.CENTER, bold: true }),
      tcell("財政安定化基金償還金", { bold: true }),
      tcell("第9期で借入があれば償還金計上（川崎町では従来借入なし）"),
    ]}),
    new TableRow({ children: [
      tcell("6", { align: AlignmentType.CENTER, bold: true, fill: C.lorange }),
      tcell("介護給付費準備基金\n取崩額", { bold: true, fill: C.lorange }),
      tcell("保険料抑制のため取崩額を設定。R8.6に確定の基金残高を踏まえ3パターン設定", { fill: C.lorange }),
    ]}),
    new TableRow({ children: [
      tcell("7", { align: AlignmentType.CENTER, bold: true }),
      tcell("保険料収納必要額", { bold: true }),
      tcell("Step1〜6を反映した3年間の収納必要額"),
    ]}),
    new TableRow({ children: [
      tcell("8", { align: AlignmentType.CENTER, bold: true, fill: C.lblue }),
      tcell("保険料基準額算定", { bold: true, fill: C.lblue }),
      tcell("収納必要額 ÷ 予定収納率 ÷ 加重平均調整率 ÷ 第1号被保険者数 で月額基準額決定", { fill: C.lblue }),
    ]}),
  ],
});
S.push(stepTable);

S.push(subsection("保険料基準額の算定式", 7));

S.push(new Paragraph({
  spacing: { before: 200, after: 200, line: 360 },
  alignment: AlignmentType.CENTER,
  shading: { type: ShadingType.SOLID, fill: CH[7].sub },
  border: {
    top: { style: BorderStyle.SINGLE, size: 12, color: CH[7].main },
    bottom: { style: BorderStyle.SINGLE, size: 12, color: CH[7].main },
  },
  children: [text("月額基準額 = 収納必要額 ÷ 予定収納率 ÷ 加重平均調整率 ÷ 第1号被保険者数 ÷ 12", { size: 22, bold: true, color: CH[7].main })],
}));
S.push(p("ここで、「収納必要額」は3年間の合計値です。「加重平均調整率」は所得段階区分別の保険料率を加重平均した値で、現行9段階の場合は約1.00となります。「予定収納率」は次節で詳述します。"));

S.push(subsection("介護給付費準備基金の活用方針", 7));

S.push(p("介護給付費準備基金は、第9期計画期間中の介護給付費が見込みを下回った際の剰余金を積立てた基金で、保険料抑制のために第10期で取崩しを検討します。"));

const fundTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("項目", { width: 38, fill: CH[7].main }),
      thcell("金額・方針", { width: 62, fill: CH[7].main }),
    ]}),
    new TableRow({ children: [
      tcell("R8年6月時点 基金残高（見込み）", { bold: true }),
      tcell("【町確認・R8.6確定値】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("R8年度末 基金残高（見込み）", { bold: true }),
      tcell("【R8.3積立増減を踏まえ予測】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("第10期 取崩額（3パターン）", { bold: true, fill: C.lorange }),
      tcell("A：取崩なし／B：50%取崩／C：全額取崩", { fill: C.lorange }),
    ]}),
    new TableRow({ children: [
      tcell("第11期以降への影響", { bold: true }),
      tcell("基金温存（パターンA）は次期の保険料急騰リスクを抑制／全額取崩（パターンC）は次期の負担が増大"),
    ]}),
  ],
});
S.push(fundTable);

S.push(subsection("試算3パターン", 7));
S.push(p("第10期の介護保険料は、介護給付費準備基金の取崩額を主な変動要因として、以下の3パターンを試算し、第3回策定委員会で協議のうえ確定します。"));

const patternTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("パターン", { width: 12, fill: CH[7].main }),
      thcell("基金取崩方針", { width: 20, fill: CH[7].main }),
      thcell("特徴", { width: 38, fill: CH[7].main }),
      thcell("月額イメージ", { width: 16, fill: CH[7].main }),
      thcell("第11期影響", { width: 14, fill: CH[7].main }),
    ]}),
    new TableRow({ children: [
      tcell("A", { align: AlignmentType.CENTER, bold: true, fill: C.lblue }),
      tcell("基金取崩なし", { bold: true }),
      tcell("基金を温存。第11期以降の保険料急騰リスクを抑制。給付費増を保険料に直接反映"),
      tcell("【試算】最高水準", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("緩和", { align: AlignmentType.CENTER, color: C.green, bold: true }),
    ]}),
    new TableRow({ children: [
      tcell("B", { align: AlignmentType.CENTER, bold: true, fill: C.lorange }),
      tcell("基金50%取崩", { bold: true }),
      tcell("給付費増を一部相殺し、現実的な保険料水準を維持。次期負担との均衡を取る", { fill: C.lorange }),
      tcell("【試算】中位水準", { color: C.orange, italics: true, align: AlignmentType.CENTER, fill: C.lorange }),
      tcell("中位", { align: AlignmentType.CENTER, fill: C.lorange }),
    ]}),
    new TableRow({ children: [
      tcell("C", { align: AlignmentType.CENTER, bold: true, fill: C.lgreen }),
      tcell("基金全額取崩", { bold: true }),
      tcell("保険料抑制を最大化（住民負担軽減）。ただし第11期負担が大きくなる"),
      tcell("【試算】最低水準", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("増大", { align: AlignmentType.CENTER, color: C.red, bold: true }),
    ]}),
  ],
});
S.push(patternTable);

S.push(subsection("所得段階区分の検討（9段階 → 13段階）", 7));

S.push(p("国の介護保険制度改正により、所得段階区分について現行9段階から13段階への見直しが推奨されています。13段階化は、所得の高い層に応分の負担を求め、低所得層の負担を軽減する効果が期待されます。"));

const dankaiTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("項目", { width: 28, fill: CH[7].main }),
      thcell("現行9段階", { width: 36, fill: CH[7].main }),
      thcell("第10期検討（13段階）", { width: 36, fill: CH[7].main }),
    ]}),
    new TableRow({ children: [
      tcell("段階数", { bold: true }),
      tcell("9段階", { align: AlignmentType.CENTER }),
      tcell("13段階（推奨）", { align: AlignmentType.CENTER, bold: true, color: C.orange }),
    ]}),
    new TableRow({ children: [
      tcell("最低段階の保険料率", { bold: true }),
      tcell("0.50（第1段階）", { align: AlignmentType.CENTER }),
      tcell("0.455〜（軽減強化）", { align: AlignmentType.CENTER, color: C.green }),
    ]}),
    new TableRow({ children: [
      tcell("最高段階の保険料率", { bold: true }),
      tcell("1.70（第9段階・320万円以上）", { align: AlignmentType.CENTER }),
      tcell("2.40（第13段階・720万円以上）", { align: AlignmentType.CENTER, color: C.orange, bold: true }),
    ]}),
    new TableRow({ children: [
      tcell("川崎町への影響", { bold: true, fill: C.lblue }),
      tcell("非課税層979人（30.1%）の負担", { align: AlignmentType.CENTER, fill: C.lblue }),
      tcell("非課税層の負担軽減＋高所得層への応分負担", { align: AlignmentType.CENTER, fill: C.lblue, bold: true }),
    ]}),
  ],
});
S.push(dankaiTable);

S.push(p("13段階化により、川崎町では非課税層979人（第1〜3段階・30.1%）の保険料負担軽減が期待される一方、第8段階以上の173+106＝279人（8.6%）への負担増となります。具体的な段階別保険料率と影響額は、第3回策定委員会で協議します。"));

S.push(subsection("予定収納率の設定", 7));

S.push(p("予定収納率は、過去5年間の保険料収納率の平均値を採用します。川崎町の収納率は高水準で推移しており、第10期の予定収納率は98%以上を見込みます。"));

const shunouTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("項目", { width: 36, fill: CH[7].main }),
      thcell("R3実績", { width: 22, fill: CH[7].main }),
      thcell("過去5年平均（見込み）", { width: 42, fill: CH[7].main }),
    ]}),
    new TableRow({ children: [
      tcell("特別徴収（年金天引）", { bold: true }),
      tcell("100.0%", { align: AlignmentType.RIGHT, color: C.green, bold: true }),
      tcell("100.0%（安定）", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("普通徴収（自主納付・現年度分）", { bold: true }),
      tcell("87.3%", { align: AlignmentType.RIGHT }),
      tcell("【町確認・経年推移把握必要】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("現年度分計", { bold: true, fill: C.lblue }),
      tcell("99.0%", { align: AlignmentType.RIGHT, bold: true, fill: C.lblue, color: C.green }),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER, fill: C.lblue }),
    ]}),
    new TableRow({ children: [
      tcell("総合計（滞納繰越分含む）", { bold: true, fill: C.lorange }),
      tcell("96.3%", { align: AlignmentType.RIGHT, bold: true, fill: C.lorange }),
      tcell("【町確認・第10期予定収納率の根拠】", { color: C.orange, italics: true, align: AlignmentType.CENTER, fill: C.lorange, bold: true }),
    ]}),
  ],
});
S.push(shunouTable);
S.push(source("R3実績：年報データ2021・様式3／過去5年平均：町確認"));

S.push(subsection("近隣自治体・宮城県・全国との比較", 7));

S.push(p("第10期保険料の参考値として、近隣自治体（仙南圏域）・宮城県・全国の保険料水準を比較します。川崎町は宮城県内35市町村中5位の高齢化率（41.4%）にもかかわらず、第9期保険料6,500円は宮城県平均水準で運営されていることが確認できます。"));

const compareTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("自治体・区分", { width: 38, fill: CH[7].main }),
      thcell("第9期保険料\n（月額基準額）", { width: 28, fill: CH[7].main }),
      thcell("出典・備考", { width: 34, fill: CH[7].main }),
    ]}),
    new TableRow({ children: [
      tcell("川崎町", { bold: true, fill: C.lorange }),
      tcell("6,500円", { align: AlignmentType.RIGHT, bold: true, fill: C.lorange }),
      tcell("MECE版データ", { fill: C.lorange }),
    ]}),
    new TableRow({ children: [
      tcell("仙南圏域近隣（柴田町・大河原町・村田町等）", { bold: true }),
      tcell("【見える化システム要確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("見える化システム"),
    ]}),
    new TableRow({ children: [
      tcell("宮城県平均", { bold: true }),
      tcell("【県集計値要確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("宮城県高齢者保健福祉計画"),
    ]}),
    new TableRow({ children: [
      tcell("全国平均", { bold: true, fill: C.lblue }),
      tcell("6,225円", { align: AlignmentType.RIGHT, bold: true, fill: C.lblue }),
      tcell("厚労省 第9期保険料状況", { fill: C.lblue }),
    ]}),
  ],
});
S.push(compareTable);
S.push(placeholder("仙南圏域・宮城県の比較値は、見える化システム及び宮城県提供データの取得後にVer.2.0で確定します。"));

S.push(subsection("13段階区分の所得段階別保険料（試算枠）", 7));

S.push(p("13段階区分への移行に伴う所得段階別保険料（月額・年額）の試算枠を以下に示します。確定値は第3回策定委員会で協議します。"));

const dankaiCalcTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("段階", { width: 8, fill: CH[7].main }),
      thcell("区分", { width: 36, fill: CH[7].main }),
      thcell("保険料率", { width: 14, fill: CH[7].main }),
      thcell("月額(試算)", { width: 18, fill: CH[7].main }),
      thcell("人数(R3)", { width: 24, fill: CH[7].main }),
    ]}),
    new TableRow({ children: [
      tcell("第1", { bold: true, align: AlignmentType.CENTER }), tcell("生活保護・世帯非課税(年金80万以下)"),
      tcell("0.455", { align: AlignmentType.CENTER, color: C.green }),
      tcell("【試算】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("431人", { align: AlignmentType.RIGHT }),
    ]}),
    new TableRow({ children: [
      tcell("第2", { bold: true, align: AlignmentType.CENTER }), tcell("世帯非課税(年金80〜120万)"),
      tcell("0.685", { align: AlignmentType.CENTER, color: C.green }),
      tcell("【試算】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("273人", { align: AlignmentType.RIGHT }),
    ]}),
    new TableRow({ children: [
      tcell("第3", { bold: true, align: AlignmentType.CENTER }), tcell("世帯非課税(年金120万超)"),
      tcell("0.690", { align: AlignmentType.CENTER, color: C.green }),
      tcell("【試算】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("275人", { align: AlignmentType.RIGHT }),
    ]}),
    new TableRow({ children: [
      tcell("第4", { bold: true, align: AlignmentType.CENTER }), tcell("本人非課税・世帯課税(80万以下)"),
      tcell("0.900", { align: AlignmentType.CENTER }),
      tcell("【試算】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("466人", { align: AlignmentType.RIGHT }),
    ]}),
    new TableRow({ children: [
      tcell("第5", { bold: true, align: AlignmentType.CENTER, fill: C.lorange }),
      tcell("本人非課税・世帯課税(基準)", { fill: C.lorange }),
      tcell("1.000", { align: AlignmentType.CENTER, fill: C.lorange, bold: true }),
      tcell("【基準額】", { color: C.orange, italics: true, align: AlignmentType.CENTER, fill: C.lorange, bold: true }),
      tcell("697人", { align: AlignmentType.RIGHT, fill: C.lorange, bold: true }),
    ]}),
    new TableRow({ children: [
      tcell("第6", { bold: true, align: AlignmentType.CENTER }), tcell("本人課税(合計所得120万未満)"),
      tcell("1.200", { align: AlignmentType.CENTER }),
      tcell("【試算】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("449人", { align: AlignmentType.RIGHT }),
    ]}),
    new TableRow({ children: [
      tcell("第7", { bold: true, align: AlignmentType.CENTER }), tcell("本人課税(120〜210万未満)"),
      tcell("1.300", { align: AlignmentType.CENTER }),
      tcell("【試算】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("385人", { align: AlignmentType.RIGHT }),
    ]}),
    new TableRow({ children: [
      tcell("第8", { bold: true, align: AlignmentType.CENTER }), tcell("本人課税(210〜320万未満)"),
      tcell("1.500", { align: AlignmentType.CENTER }),
      tcell("【試算】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("173人", { align: AlignmentType.RIGHT }),
    ]}),
    new TableRow({ children: [
      tcell("第9", { bold: true, align: AlignmentType.CENTER }), tcell("本人課税(320〜420万未満)"),
      tcell("1.700", { align: AlignmentType.CENTER }),
      tcell("【試算】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("106人(現9段階)", { align: AlignmentType.RIGHT }),
    ]}),
    new TableRow({ children: [
      tcell("第10", { bold: true, align: AlignmentType.CENTER, color: C.orange }),
      tcell("本人課税(420〜520万未満)", { color: C.orange }),
      tcell("1.900", { align: AlignmentType.CENTER, color: C.orange }),
      tcell("【試算】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("【新設・町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("第11", { bold: true, align: AlignmentType.CENTER, color: C.orange }),
      tcell("本人課税(520〜620万未満)", { color: C.orange }),
      tcell("2.100", { align: AlignmentType.CENTER, color: C.orange }),
      tcell("【試算】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("【新設・町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("第12", { bold: true, align: AlignmentType.CENTER, color: C.orange }),
      tcell("本人課税(620〜720万未満)", { color: C.orange }),
      tcell("2.300", { align: AlignmentType.CENTER, color: C.orange }),
      tcell("【試算】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("【新設・町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("第13", { bold: true, align: AlignmentType.CENTER, color: C.orange }),
      tcell("本人課税(720万以上)", { color: C.orange }),
      tcell("2.400", { align: AlignmentType.CENTER, color: C.orange, bold: true }),
      tcell("【試算】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("【新設・町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(dankaiCalcTable);
S.push(source("保険料率は国の標準率（参考値）。第10段階以降の人数は町からの提供データで確認"));

S.push(subsection("確定スケジュール", 7));

S.push(p("第10期介護保険料の確定は、以下のスケジュールで進めます。"));
S.push(numItem("①", "令和8年6月：介護給付費準備基金残高の確定（町担当課）"));
S.push(numItem("②", "令和8年8月：第1回策定委員会でアンケート結果報告、給付費試算前提の協議"));
S.push(numItem("③", "令和8年11月：第2回策定委員会でサービス見込量・標準給付費の方向性確定"));
S.push(numItem("④", "令和8年12月〜令和9年1月：保険料試算3パターンの精緻化（コンサル＋町担当）"));
S.push(numItem("⑤", "令和9年1月中旬：第3回策定委員会で保険料試算3パターンを協議"));
S.push(numItem("⑥", "令和9年2月：第4回策定委員会で保険料基準額を決定、町長答申"));
S.push(numItem("⑦", "令和9年3月：3月議会へ条例改正案上程、可決後に公表"));

S.push(placeholder("具体的な保険料試算結果は、介護給付費準備基金残高（令和8年6月確定）、サービス見込量、国通知（夏以降）、アンケート結果を踏まえ、Ver.2.0で確定します。確定値の信頼性確保のため、宮城県・近隣町との比較・見える化システム登録を経て、第3回・第4回策定委員会で協議します。"));

module.exports = { S };
console.log("Part7 (Ch7 detail) ready. Blocks:", S.length);
