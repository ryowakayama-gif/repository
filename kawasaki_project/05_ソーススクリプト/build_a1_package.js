/**
 * A-1: 町担当向け送付パッケージ
 * 川崎町保健福祉課 大宮様 宛の送付一式をまとめた文書
 * - カバーレター（依頼文）
 * - 送付物一覧（4つの記入用フォーマット）
 * - 記入優先順位とスケジュール
 * - 各フォーマットの記入ガイド
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
// 表紙
// ===========================================================
S.push(new Paragraph({
  spacing: { before: 2400, after: 240 },
  alignment: AlignmentType.CENTER,
  children: [text("川崎町高齢者保健福祉計画", { size: 32, bold: true, color: C.navy })],
}));
S.push(new Paragraph({
  spacing: { before: 0, after: 480 },
  alignment: AlignmentType.CENTER,
  children: [text("第10期介護保険事業計画 策定業務", { size: 32, bold: true, color: C.navy })],
}));
S.push(new Paragraph({
  spacing: { before: 240, after: 240 },
  alignment: AlignmentType.CENTER,
  border: {
    top: { style: BorderStyle.DOUBLE, size: 18, color: C.navy },
    bottom: { style: BorderStyle.DOUBLE, size: 18, color: C.navy },
  },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  children: [text("町担当課ご記入依頼パッケージ", { size: 30, bold: true, color: C.navy })],
}));
S.push(new Paragraph({
  spacing: { before: 240, after: 240 },
  alignment: AlignmentType.CENTER,
  children: [text("〜 第1回策定委員会（令和8年8月）に向けた事前準備のお願い 〜", {
    size: 20, italics: true, color: C.blue
  })],
}));

S.push(new Paragraph({
  spacing: { before: 3600, after: 80 },
  alignment: AlignmentType.RIGHT,
  children: [text("令和8年6月", { size: 22, color: C.gray })],
}));
S.push(new Paragraph({
  spacing: { before: 80, after: 80 },
  alignment: AlignmentType.RIGHT,
  children: [text("宛：川崎町保健福祉課　大宮様", { size: 22, bold: true, color: C.navy })],
}));
S.push(new Paragraph({
  spacing: { before: 80, after: 80 },
  alignment: AlignmentType.RIGHT,
  children: [text("ビズアップ公共コンサルティング株式会社", { size: 20, color: C.gray })],
}));
S.push(new Paragraph({
  spacing: { before: 0, after: 80 },
  alignment: AlignmentType.RIGHT,
  children: [text("　札幌事業所　若山・髙橋・山内・河崎", { size: 18, color: C.gray })],
}));

// ===========================================================
// 1. ご挨拶（カバーレター本文）
// ===========================================================
S.push(new Paragraph({
  spacing: { before: 0, after: 0, line: 240 },
  pageBreakBefore: true,
  children: [text("")],
}));

S.push(new Paragraph({
  spacing: { before: 240, after: 320, line: 320 },
  alignment: AlignmentType.CENTER,
  border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: C.navy } },
  children: [text("1．ご挨拶・現状報告", { size: 28, bold: true, color: C.navy })],
}));

S.push(p("時下、ますますご清祥のこととお喜び申し上げます。日頃より、川崎町第10期介護保険事業計画の策定業務に格別のご高配を賜り、誠にありがとうございます。"));

S.push(p("さて、令和7年度キックオフ会議以降、第9期計画体系の継承方針、認知症基本法対応の独立章化、川崎町固有の課題（移動支援3層構造・施設偏在・8050問題・広域医療連携等）の整理を進め、計画書素案 Ver.1.5（全43頁・8章構成）の取りまとめが完了いたしました。"));

S.push(p("つきましては、令和8年8月中旬に予定されております第1回策定委員会に向け、計画素案 Ver.2.0 への発展（実数値・アンケート結果反映）に必要な町データの収集を本格化させていただきたく、町担当課におかれましてご記入いただきたいフォーマット一式を本パッケージとしてお送りいたします。"));

S.push(p("ご多忙の折、誠に恐縮ではございますが、何卒ご協力のほどよろしくお願い申し上げます。"));

S.push(spacer());

// ===========================================================
// 2. 送付物一覧
// ===========================================================
S.push(new Paragraph({
  spacing: { before: 360, after: 240, line: 320 },
  alignment: AlignmentType.CENTER,
  border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: C.navy } },
  children: [text("2．送付物一覧", { size: 28, bold: true, color: C.navy })],
}));

S.push(p("本パッケージには、以下の4種類の記入用フォーマットと、参考資料として現在の計画素案v1.5（全43頁）が含まれます。"));

const sendingTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("No", { width: 6 }),
      thcell("送付物（ファイル名）", { width: 38 }),
      thcell("内容", { width: 38 }),
      thcell("形式", { width: 8 }),
      thcell("優先度", { width: 10 }),
    ]}),
    new TableRow({ children: [
      tcell("①", { align: AlignmentType.CENTER, bold: true }),
      tcell("川崎町_第9期実績一覧_町記入用.xlsx", { bold: true, color: C.navy }),
      tcell("第9期計画期間（R6〜R8）の主要施策実績データ。第10期計画素案 第3章への反映"),
      tcell("Excel", { align: AlignmentType.CENTER }),
      tcell("A", { bold: true, color: C.red, align: AlignmentType.CENTER, fill: C.lorange }),
    ]}),
    new TableRow({ children: [
      tcell("②", { align: AlignmentType.CENTER, bold: true }),
      tcell("川崎町_アンケート集計分析テンプレート.xlsx", { bold: true, color: C.navy }),
      tcell("一般高齢者・認定者アンケート（R8.7末回収予定）の事前集計枠組み。回収後に弊社で集計実施"),
      tcell("Excel", { align: AlignmentType.CENTER }),
      tcell("B", { bold: true, color: C.blue, align: AlignmentType.CENTER, fill: C.lblue }),
    ]}),
    new TableRow({ children: [
      tcell("③", { align: AlignmentType.CENTER, bold: true }),
      tcell("川崎町_必要資料_入力フォーマット.xlsx", { bold: true, color: C.navy }),
      tcell("MECE版データ入力フォーマット（13シート構成）。サービス見込量・給付費・保険料試算の根拠データ"),
      tcell("Excel", { align: AlignmentType.CENTER }),
      tcell("A", { bold: true, color: C.red, align: AlignmentType.CENTER, fill: C.lorange }),
    ]}),
    new TableRow({ children: [
      tcell("④", { align: AlignmentType.CENTER, bold: true }),
      tcell("川崎町_実績データ確認サマリー.xlsx", { bold: true, color: C.navy }),
      tcell("既存の保険者データ・年報データ等から既知の数値をまとめたサマリー。記入用ではなく参考用"),
      tcell("Excel", { align: AlignmentType.CENTER }),
      tcell("参考", { color: C.gray, align: AlignmentType.CENTER, fill: C.lgray }),
    ]}),
    new TableRow({ children: [
      tcell("⑤", { align: AlignmentType.CENTER, bold: true }),
      tcell("川崎町_計画書素案_v1.5_カラー統一版.pdf", { bold: true, color: C.navy }),
      tcell("計画書素案 Ver.1.5（全43頁・8章構成）。本書の前提資料・本素案を通じて記入根拠を確認可能"),
      tcell("PDF", { align: AlignmentType.CENTER }),
      tcell("参考", { color: C.gray, align: AlignmentType.CENTER, fill: C.lgray }),
    ]}),
  ],
});
S.push(sendingTable);

S.push(spacer());

S.push(fact("優先度A（赤）の①と③は、計画素案Ver.2.0への反映に必須のデータです。優先的にご対応をお願いいたします。"));

// ===========================================================
// 3. 記入優先順位とスケジュール
// ===========================================================
S.push(new Paragraph({
  spacing: { before: 360, after: 240, line: 320 },
  alignment: AlignmentType.CENTER,
  border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: C.navy } },
  children: [text("3．記入優先順位とスケジュール", { size: 28, bold: true, color: C.navy })],
}));

S.push(p("以下のスケジュールに沿ってご記入・ご返送をお願いしたく存じます。"));

const scheduleTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("時期", { width: 16 }),
      thcell("対象", { width: 24 }),
      thcell("作業内容", { width: 40 }),
      thcell("成果物への反映", { width: 20 }),
    ]}),
    new TableRow({ children: [
      tcell("令和8年6月\n下旬", { bold: true, fill: C.lorange, align: AlignmentType.CENTER }),
      tcell("①第9期実績一覧", { bold: true }),
      tcell("01_KPI一覧・02_介護予防_健康・04_認知症_包括 を優先的にご記入"),
      tcell("第3章 3-2 主要施策の取組実績", { color: C.navy }),
    ]}),
    new TableRow({ children: [
      tcell("令和8年6月\n下旬", { bold: true, fill: C.lorange, align: AlignmentType.CENTER }),
      tcell("③MECE版データ", { bold: true }),
      tcell("第1号被保険者数・要介護認定者数・サービス受給者・給付費の経年推移（直近5年分）"),
      tcell("第2章・第7章 7-1見込量", { color: C.navy }),
    ]}),
    new TableRow({ children: [
      tcell("令和8年7月\n上旬", { bold: true, fill: C.lblue, align: AlignmentType.CENTER }),
      tcell("①第9期実績一覧（続）", { bold: true }),
      tcell("03_在宅生活支援・05_介護サービス_人材・06_第10期反映方針 をご記入"),
      tcell("第3章 3-2 / 第5章施策展開", { color: C.navy }),
    ]}),
    new TableRow({ children: [
      tcell("令和8年7月\n下旬", { bold: true, fill: C.lblue, align: AlignmentType.CENTER }),
      tcell("【アンケート回収】", { bold: true, color: C.orange }),
      tcell("一般高齢者1,000名・認定者300名の回収（町実施）"),
      tcell("第2章2-2・第5章 KPI", { color: C.navy }),
    ]}),
    new TableRow({ children: [
      tcell("令和8年8月\n上旬", { bold: true, fill: C.lgreen, align: AlignmentType.CENTER }),
      tcell("②集計テンプレ", { bold: true }),
      tcell("回収アンケートを弊社で集計し、本テンプレに反映（町担当ご記入は不要）"),
      tcell("第3章・第5章・第6章", { color: C.navy }),
    ]}),
    new TableRow({ children: [
      tcell("令和8年8月\n中旬", { bold: true, fill: C.lgreen, align: AlignmentType.CENTER }),
      tcell("【第1回策定委員会】", { bold: true, color: C.orange }),
      tcell("計画素案Ver.2.0素案・アンケート結果報告・基本方針の協議"),
      tcell("計画素案Ver.2.0", { color: C.navy, bold: true }),
    ]}),
    new TableRow({ children: [
      tcell("令和8年6月\n確定", { bold: true, fill: C.lorange, align: AlignmentType.CENTER }),
      tcell("【基金残高】", { bold: true, color: C.red }),
      tcell("介護給付費準備基金残高の確定値（R8.6時点）", { color: C.red }),
      tcell("第7章 7-3保険料試算", { color: C.navy, bold: true }),
    ]}),
  ],
});
S.push(scheduleTable);

S.push(spacer());

S.push(p("※ なお、ご記入が困難な項目がございましたら、空欄のままで結構です。弊社で代替手段（年報データ・推計値の活用等）を検討いたします。"));

// ===========================================================
// 4. 各フォーマットの記入ガイド
// ===========================================================
S.push(new Paragraph({
  spacing: { before: 360, after: 240, line: 320 },
  alignment: AlignmentType.CENTER,
  pageBreakBefore: true,
  border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: C.navy } },
  children: [text("4．各フォーマットの記入ガイド", { size: 28, bold: true, color: C.navy })],
}));

// 4-1
S.push(new Paragraph({
  spacing: { before: 320, after: 140, line: 320 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C.navy } },
  children: [text("4-1．① 川崎町_第9期実績一覧_町記入用.xlsx", { size: 24, bold: true, color: C.navy })],
}));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 概要", { size: 22, bold: true, color: C.navy })],
}));

S.push(p("7シート構成（00使い方／01KPI一覧／02介護予防_健康／03在宅生活支援／04認知症_包括／05介護サービス_人材／06第10期反映方針）。第9期計画期間中の主要施策35事業以上の実績を網羅。"));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 記入の留意点", { size: 22, bold: true, color: C.navy })],
}));

S.push(numItem("①", "黄色セルが記入欄、緑セルが既知の確定値（記入不要・参考用）です。"));
S.push(numItem("②", "対象期間は令和6年度（R6.4〜R7.3）を基本とし、令和7年度の最新値もあわせてご記入ください。"));
S.push(numItem("③", "不明・該当なしの項目は空欄または「─」「0」でご対応ください。"));
S.push(numItem("④", "ユニバーサルサポーター制度（介護予防サロン84名・スマイル40名等）は既知値（緑）で事前投入済みです。最新値があれば右欄に追記ください。"));
S.push(numItem("⑤", "認知症サポーター累計550名・キャラバンメイト73名も同様に事前投入済みです。"));
S.push(numItem("⑥", "06_第10期反映方針シートは、各事業の継続・拡充・縮小・廃止・新規・統合のいずれかをご記入いただく事業棚卸しシートです。第10期方針決定の重要な根拠となります。"));

// 4-2
S.push(new Paragraph({
  spacing: { before: 320, after: 140, line: 320 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C.navy } },
  children: [text("4-2．③ 川崎町_必要資料_入力フォーマット.xlsx", { size: 24, bold: true, color: C.navy })],
}));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 概要", { size: 22, bold: true, color: C.navy })],
}));

S.push(p("MECE版データ入力フォーマット（13シート構成）。第1号被保険者数・要介護認定者数・サービス受給者・給付費・保険料の経年推移（直近5年分・R3〜R7）を収集。第10期保険料試算・サービス見込量算定の根拠データ。"));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 記入の留意点", { size: 22, bold: true, color: C.navy })],
}));

S.push(numItem("①", "本フォーマットは前回送付済みで、一部記入いただいた状態かと存じます。残りの空欄を埋めていただく形になります。"));
S.push(numItem("②", "保険給付実績・地域支援事業実績は、年報データ・介護保険事業状況報告から転記可能なものは事前投入済みです。"));
S.push(numItem("③", "第8期保険料6,380円・第9期保険料6,500円は既知値として確認済みです。"));
S.push(numItem("④", "介護給付費準備基金残高（R8.6時点）は、保険料試算の中核データとなりますので、確定次第ご連絡ください。"));
S.push(numItem("⑤", "住所地特例該当者の所在自治体内訳（柴田町・大河原町・仙台市等）が把握可能でしたら、追加でご教示ください。"));

// 4-3
S.push(new Paragraph({
  spacing: { before: 320, after: 140, line: 320 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C.navy } },
  children: [text("4-3．② 川崎町_アンケート集計分析テンプレート.xlsx", { size: 24, bold: true, color: C.navy })],
}));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 概要", { size: 22, bold: true, color: C.navy })],
}));

S.push(p("6シート構成。アンケート回収後（R8.7末予定）に弊社で集計実施するためのテンプレート。町担当のご記入は不要です。本パッケージには参考として同梱しております。"));

S.push(new Paragraph({
  spacing: { before: 200, after: 100, line: 300 },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
  children: [text("　▌ 町担当へのお願い事項", { size: 22, bold: true, color: C.navy })],
}));

S.push(numItem("①", "アンケート発送（R8.6下旬予定）・回収（R8.7末予定）のご対応をお願いします。"));
S.push(numItem("②", "回収後、原本・集計可能なデータ（紙の場合はスキャンPDF、Webの場合はCSV等）を弊社までご送付ください。"));
S.push(numItem("③", "回収率向上のための再送・督促等のご対応もお願いします。"));

// ===========================================================
// 5. お問合せ・ご返送先
// ===========================================================
S.push(new Paragraph({
  spacing: { before: 360, after: 240, line: 320 },
  alignment: AlignmentType.CENTER,
  border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: C.navy } },
  children: [text("5．お問合せ・ご返送先", { size: 28, bold: true, color: C.navy })],
}));

S.push(p("ご質問・ご相談、記入後のご返送等は、下記までお願いいたします。"));

const contactTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      tcell("受託者", { bold: true, fill: C.lblue, width: 24 }),
      tcell("ビズアップ公共コンサルティング株式会社（札幌事業所）", { bold: true, width: 76 }),
    ]}),
    new TableRow({ children: [
      tcell("主担当", { bold: true, fill: C.lblue }),
      tcell("若山　（プロジェクトリーダー）"),
    ]}),
    new TableRow({ children: [
      tcell("副担当", { bold: true, fill: C.lblue }),
      tcell("髙橋・山内・河崎"),
    ]}),
    new TableRow({ children: [
      tcell("ご返送形式", { bold: true, fill: C.lblue }),
      tcell("メール添付（Excelファイル）が望ましく、容量が大きい場合はクラウドストレージのリンクでも結構です"),
    ]}),
    new TableRow({ children: [
      tcell("中間報告", { bold: true, fill: C.lblue }),
      tcell("ご記入の進捗状況に応じて、随時メールやWeb会議で打合せをさせていただきます"),
    ]}),
  ],
});
S.push(contactTable);

S.push(spacer());
S.push(spacer());

S.push(new Paragraph({
  spacing: { before: 240, after: 240, line: 320 },
  alignment: AlignmentType.CENTER,
  border: {
    top: { style: BorderStyle.DOUBLE, size: 12, color: C.navy },
    bottom: { style: BorderStyle.DOUBLE, size: 12, color: C.navy },
  },
  shading: { type: ShadingType.SOLID, fill: C.lblue },
  children: [text("引き続き、川崎町第10期計画策定にご協力を賜りますよう、よろしくお願い申し上げます。", {
    size: 22, bold: true, color: C.navy
  })],
}));

// ===========================================================
// ドキュメント生成
// ===========================================================
const doc = new Document({
  creator: "ビズアップ公共コンサルティング株式会社",
  title: "川崎町 町担当課ご記入依頼パッケージ",
  description: "第10期介護保険事業計画策定 町担当ご記入依頼パッケージ",
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
            text: "川崎町第10期介護保険事業計画策定　町担当ご記入依頼パッケージ",
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
  fs.writeFileSync("/home/claude/kawasaki_work/川崎町_町担当課ご記入依頼パッケージ.docx", buffer);
  console.log("Build done. Blocks:", S.length);
}).catch(err => {
  console.error("Build error:", err);
  process.exit(1);
});
