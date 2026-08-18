/**
 * 川崎町第10期計画書素案v1.0 - Part2 (Ch5-8)
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
// 第5章 施策の展開
// ===========================================================
S.push(...chapterTitle(5));

S.push(p("本章では、第4章で示した6つの基本目標のうち、認知症施策（基本目標6）を除く基本目標1〜5に対応する施策を、施策体系・主な事業・成果指標（KPI）の3層構造で整理します。認知症施策（基本目標6）は第6章で独立章として詳述します。"));

S.push(placeholder("各施策の住民ニーズ根拠（一般高齢者ニーズ調査・要支援要介護認定者調査）は、令和8年7月末の回収・8月の集計後にVer.2.0へ反映します。本素案v1.0では実績データ・キックオフ確認事項・第9期計画継承の方針から施策方向を提示します。"));

// 5-1
S.push(section(5, 1, "健康づくり・介護予防の推進（基本目標1）"));

S.push(subsection("施策の方向性", 5));
S.push(p("高齢者一人ひとりが生涯にわたって健やかに暮らし続けることができるよう、健康づくり・フレイル予防・通いの場の充実・口腔と栄養の支援を一体的に推進します。"));
S.push(p("川崎町では既に介護予防サロン84名、スマイルサポーター40名、レクリエーションサポーター29名等のユニバーサルサポーター制度を運用しており、町内の通いの場活動を住民主体で展開する基盤が整っています。第10期では、この基盤を活かし、後期高齢者75歳以上1,675人の増加局面を見据えた介護予防の対象拡大と質の向上を図ります。特に、独居高齢者や中山間地域居住の方が孤立せず通いの場に参加できる移動支援との連動が課題です。"));

S.push(subsection("主な事業", 5));

const e51Table = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("No", { width: 6, fill: CH[5].main }),
      thcell("事業名", { width: 30, fill: CH[5].main }),
      thcell("内容", { width: 50, fill: CH[5].main }),
      thcell("主な実施主体", { width: 14, fill: CH[5].main }),
    ]}),
    new TableRow({ children: [
      tcell("1-1", { align: AlignmentType.CENTER, bold: true }),
      tcell("一般介護予防事業"),
      tcell("元気まんてん教室・スマイル教室・パドル運動等の継続実施"),
      tcell("町・社協", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("1-2", { align: AlignmentType.CENTER, bold: true }),
      tcell("通いの場（サロン）の展開"),
      tcell("ユニバーサルサポーター制度（介護予防サロン84名）による全町展開"),
      tcell("社協・サポーター", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("1-3", { align: AlignmentType.CENTER, bold: true }),
      tcell("フレイル予防の推進"),
      tcell("運動・栄養・社会参加の3要素を組み合わせた予防プログラム"),
      tcell("町・包括", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("1-4", { align: AlignmentType.CENTER, bold: true }),
      tcell("口腔・栄養の支援"),
      tcell("オーラルフレイル予防・栄養指導・配食サービス（独自）"),
      tcell("町・栄養士", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("1-5", { align: AlignmentType.CENTER, bold: true }),
      tcell("健診・早期発見"),
      tcell("特定健診・後期高齢者健診の受診率向上、要介護リスク早期発見"),
      tcell("町・国保川崎病院", { align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(e51Table);

S.push(subsection("成果指標（KPI）", 5));

const k51Table = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("指標", { width: 38, fill: CH[5].main }),
      thcell("現状値", { width: 20, fill: CH[5].main }),
      thcell("目標値（R11）", { width: 20, fill: CH[5].main }),
      thcell("出典・備考", { width: 22, fill: CH[5].main }),
    ]}),
    new TableRow({ children: [
      tcell("通いの場参加率（65歳以上）", { bold: true }),
      tcell("【調査後設定】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("【委員会協議】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("一般高齢者調査"),
    ]}),
    new TableRow({ children: [
      tcell("一般介護予防事業 年間実施回数", { bold: true }),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("現状＋10%", { align: AlignmentType.CENTER }),
      tcell("町実績"),
    ]}),
    new TableRow({ children: [
      tcell("特定健診受診率（65歳以上）", { bold: true }),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("【委員会協議】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("健康かわさき21との整合"),
    ]}),
    new TableRow({ children: [
      tcell("ユニバーサルサポーター活動回数", { bold: true }),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("現状維持以上", { align: AlignmentType.CENTER }),
      tcell("町独自指標"),
    ]}),
  ],
});
S.push(k51Table);

// 5-2
S.push(section(5, 2, "高齢者が安心して暮らせるまちづくり（基本目標2）"));

S.push(subsection("施策の方向性", 5));
S.push(p("独居高齢者・高齢者世帯・老老介護世帯の増加、未婚の子と高齢親が同居する8050世帯の課題を踏まえ、見守り・地域支え合い・移動支援・住まいの確保を総合的に推進します。"));
S.push(p("特に移動支援は、令和7年3月で従来の高齢者外出タクシー利用助成が終了し、現在は社協・NPO法人の福祉移送サービス、デマンドバス、町民バスの3層構造に移行しています。所管が地域振興課・町民生活課・社協・NPOと分かれているため、住民への分かりやすい周知と利用情報の一元化が第10期の重要課題です。"));
S.push(p("また、ふれあいネットワーク（活動員15名・協力員130名）による地域見守り、ユニバーサルサポーター（傾聴24名等）による傾聴・声かけ活動、緊急通報装置の運用等を組み合わせ、独居高齢者・高齢者世帯を見守る重層的な体制を維持・強化します。"));

S.push(subsection("主な事業", 5));

const e52Table = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("No", { width: 6, fill: CH[5].main }),
      thcell("事業名", { width: 30, fill: CH[5].main }),
      thcell("内容", { width: 50, fill: CH[5].main }),
      thcell("主な実施主体", { width: 14, fill: CH[5].main }),
    ]}),
    new TableRow({ children: [
      tcell("2-1", { align: AlignmentType.CENTER, bold: true }),
      tcell("地域見守り体制の強化", { bold: true }),
      tcell("ふれあいネットワーク活動員15名・協力員130名による地域見守り、独居高齢者・高齢者世帯への定期訪問"),
      tcell("社協・民生委員", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("2-2", { align: AlignmentType.CENTER, bold: true }),
      tcell("緊急通報装置の充実", { bold: true }),
      tcell("独居高齢者・高齢者夫婦世帯への緊急通報装置の設置・運用"),
      tcell("町・社協", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("2-3", { align: AlignmentType.CENTER, bold: true, fill: C.lorange }),
      tcell("移動支援の再構築【重点】", { bold: true, color: C.orange }),
      tcell("社協・NPO福祉移送、デマンドバス、町民バスの3層構造を住民に分かりやすく周知。地域振興課・町民生活課・社協・NPOと連携した利用情報の一元化"),
      tcell("町関係課・社協・NPO・バス事業者", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("2-4", { align: AlignmentType.CENTER, bold: true }),
      tcell("高齢者世帯への独自支援", { bold: true }),
      tcell("高齢者紙おむつ等支給・高齢者世帯エアコン購入支援（R7.10〜）・人工透析患者通院交通費助成"),
      tcell("町", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("2-5", { align: AlignmentType.CENTER, bold: true }),
      tcell("住まいの確保", { bold: true }),
      tcell("高齢者向け住宅情報の提供、サ高住・有料老人ホーム等の情報整理"),
      tcell("町・地域包括", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("2-6", { align: AlignmentType.CENTER, bold: true }),
      tcell("8050問題・老老介護への対応", { bold: true }),
      tcell("民生委員・地域包括による訪問把握、重層的支援体制（地域福祉計画と連動）"),
      tcell("町・包括・民生委員", { align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(e52Table);

S.push(subsection("成果指標（KPI）", 5));
const k52Table = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("指標", { width: 38, fill: CH[5].main }),
      thcell("現状値", { width: 20, fill: CH[5].main }),
      thcell("目標値（R11）", { width: 20, fill: CH[5].main }),
      thcell("出典・備考", { width: 22, fill: CH[5].main }),
    ]}),
    new TableRow({ children: [
      tcell("ふれあいネットワーク活動員数", { bold: true }),
      tcell("145名", { align: AlignmentType.CENTER }),
      tcell("現状維持", { align: AlignmentType.CENTER }),
      tcell("活動員15＋協力員130"),
    ]}),
    new TableRow({ children: [
      tcell("緊急通報装置 設置件数", { bold: true }),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("【委員会協議】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("町実績"),
    ]}),
    new TableRow({ children: [
      tcell("移動支援3制度の周知度", { bold: true }),
      tcell("【調査後設定】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("【委員会協議】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("一般高齢者調査"),
    ]}),
    new TableRow({ children: [
      tcell("外出に困っている高齢者の割合", { bold: true }),
      tcell("【調査後設定】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("削減", { align: AlignmentType.CENTER }),
      tcell("一般高齢者調査"),
    ]}),
  ],
});
S.push(k52Table);

// 5-3
S.push(section(5, 3, "在宅生活継続の支援（基本目標3）"));

S.push(subsection("施策の方向性", 5));
S.push(p("住み慣れた自宅・地域で生活を継続したいという高齢者と家族の希望に応えるため、在宅医療・介護連携、家族介護者支援、レスパイト等を充実させます。特に川崎町では、施設サービス利用者134人のうち77.6%が要介護3以上の重度層であることを踏まえ、軽度〜中度の認定者については在宅生活継続が可能な支援体制を整備します。"));
S.push(p("本町は町内唯一の医療拠点である国民健康保険川崎病院を在宅医療連携のハブとし、退院前カンファレンス・ACP（人生会議）の普及、訪問看護・訪問診療・ケアマネとの多職種連携を強化します。一方、専門医療・夜間救急については町外医療機関（みやぎ県南中核病院・刈田綜合病院）への依存度が高く、通院の移動・付添に係る家族負担が大きいことから、広域医療連携の体制整備と移動支援との連動を重点化します。"));
S.push(p("家族介護者支援については、独居高齢者・高齢者世帯・老老介護世帯の増加、未婚の子と高齢親が同居する8050世帯の課題を踏まえ、家族介護教室・介護相談・レスパイト（短期入所）活用支援を充実させます。特に、家族介護による介護離職を防ぐため、町内事業所との連携による両立支援・相談体制の整備を新規取組として位置付けます。"));

S.push(subsection("主な事業", 5));
const e53Table = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("No", { width: 6, fill: CH[5].main }),
      thcell("事業名", { width: 30, fill: CH[5].main }),
      thcell("内容", { width: 50, fill: CH[5].main }),
      thcell("主な実施主体", { width: 14, fill: CH[5].main }),
    ]}),
    new TableRow({ children: [
      tcell("3-1", { align: AlignmentType.CENTER, bold: true }),
      tcell("在宅医療・介護連携推進"),
      tcell("国保川崎病院を中心とした在宅医療・介護連携、退院支援、ACP普及"),
      tcell("町・国保川崎病院", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("3-2", { align: AlignmentType.CENTER, bold: true, fill: C.lorange }),
      tcell("広域医療連携【重点】", { bold: true, color: C.orange }),
      tcell("みやぎ県南中核病院（大河原）・刈田綜合病院（白石）等との連携、夜間救急・専門医療への広域対応"),
      tcell("町・国保病院・連携先病院", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("3-3", { align: AlignmentType.CENTER, bold: true }),
      tcell("家族介護者支援"),
      tcell("家族介護教室・介護相談・レスパイト（短期入所）の活用支援"),
      tcell("町・包括", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("3-4", { align: AlignmentType.CENTER, bold: true, fill: C.lorange }),
      tcell("介護離職防止【重点】", { bold: true, color: C.orange }),
      tcell("8050・老老介護世帯における介護と仕事の両立支援、町内事業所との連携"),
      tcell("町・包括・事業所", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("3-5", { align: AlignmentType.CENTER, bold: true }),
      tcell("住宅改修・福祉用具"),
      tcell("自宅での生活継続に必要な改修・用具の利用支援"),
      tcell("町・包括", { align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(e53Table);

// 5-4
S.push(section(5, 4, "介護サービスの質の確保と人材確保（基本目標4）"));

S.push(subsection("施策の方向性", 5));
S.push(p("町内介護サービスの質の確保、町外施設利用も含めた供給体制の整備、深刻化する介護人材不足への対応を推進します。本町の介護サービス供給体制は、特別養護老人ホーム1施設・介護老人保健施設1施設・小規模多機能型居宅介護等の地域密着型サービスが中心ですが、これらが役場周辺に偏在しており、中山間部住民にとってアクセスの障壁が大きい状況です。"));
S.push(p("住所地特例該当者24人（令和7年6月時点）が示すとおり、町外施設（柴田町・大河原町・仙台市等）への依存も常態化しています。町外施設の建設計画情報や受入余地の把握、住所地特例該当者の所在自治体内訳の把握を進め、サービス見込量算定と広域連携の根拠データとして活用します。新たな町内施設の整備は財政的に困難な状況のため、町外施設との連携深化と移動支援との連動が現実的な方策となります。"));
S.push(p("介護人材確保については、全国的・全県的に深刻化する課題であり、本町独自の対応には限界があります。宮城県社会福祉協議会の介護福祉士等修学資金貸付制度等の県事業を積極的に周知し、町内事業所と連携した職場体験・人材定着支援に取り組みます。あわせて、ICT・介護ロボット・センサー等の導入支援により、限られた人材で質の高いサービスを提供できる体制を後押しします。"));

S.push(subsection("主な事業", 5));
const e54Table = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("No", { width: 6, fill: CH[5].main }),
      thcell("事業名", { width: 30, fill: CH[5].main }),
      thcell("内容", { width: 50, fill: CH[5].main }),
      thcell("主な実施主体", { width: 14, fill: CH[5].main }),
    ]}),
    new TableRow({ children: [
      tcell("4-1", { align: AlignmentType.CENTER, bold: true }),
      tcell("介護サービスの質の確保"),
      tcell("町内事業所への研修・指導、苦情対応、ケアプラン点検"),
      tcell("町・包括", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("4-2", { align: AlignmentType.CENTER, bold: true, fill: C.lorange }),
      tcell("町外施設・住所地特例の整理【重点】", { bold: true, color: C.orange }),
      tcell("町内施設縮小・町外施設利用（住所地特例24人）の実態を踏まえた供給見込みの整備"),
      tcell("町", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("4-3", { align: AlignmentType.CENTER, bold: true }),
      tcell("福祉施設の偏在対応", { bold: true }),
      tcell("役場周辺への施設偏在を踏まえた、地域別アクセスの確保・移動支援との連動"),
      tcell("町・事業者", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("4-4", { align: AlignmentType.CENTER, bold: true }),
      tcell("介護人材の確保・育成"),
      tcell("宮城県社会福祉協議会の修学資金貸付制度の周知、町内事業所での職場体験"),
      tcell("町・県社協・事業所", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("4-5", { align: AlignmentType.CENTER, bold: true }),
      tcell("介護ロボット・ICT導入"),
      tcell("介護ロボット・ICT導入の促進、人材不足の補完"),
      tcell("町・事業所", { align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(e54Table);

// 5-5
S.push(section(5, 5, "地域包括ケアシステムの深化（基本目標5）"));

S.push(subsection("施策の方向性", 5));
S.push(p("地域包括支援センターを中心に、医療・介護・予防・住まい・生活支援を一体的に提供する地域包括ケアシステムを、川崎町の特性（中山間・7地区・後期高齢者急増）に即して深化させます。地域包括支援センターは町社会福祉協議会が運営し、保健師3名と認定調査員1名の体制で総合相談・権利擁護・包括的継続的ケアマネジメント・介護予防ケアマネジメントを担っていますが、認定調査業務との兼務で業務量が過多となっており、第10期では体制強化が課題となります。"));
S.push(p("自立支援型地域ケア会議は、要支援・要介護認定者の自立支援を多職種で検討する仕組みとして既に開催実績がありますが、第10期では多職種連携（医療・介護・地域）の深化、ケース検討から地域課題抽出への展開を重点化します。生活支援コーディネーター（SC）25名・第1層協議体・第2層協議体による生活支援体制整備事業も継続し、ユニバーサルサポーター制度との連動で地域支え合いを促進します。"));
S.push(p("地域包括ケアシステムの深化は、令和8年度から始まる地域福祉計画・障害者計画との3計画同時策定の機会を活かし、高齢×障害×子ども×生活困窮の重層的支援体制整備事業に接続することで、複合的な課題を抱える世帯（8050問題、ヤングケアラー、ダブルケア等）への包括的支援につなげます。地域福祉計画・障害者計画はジャパン総研が策定担当のため、計画間の整合確認と連動施策の調整を継続的に行います。"));

S.push(subsection("主な事業", 5));
const e55Table = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("No", { width: 6, fill: CH[5].main }),
      thcell("事業名", { width: 30, fill: CH[5].main }),
      thcell("内容", { width: 50, fill: CH[5].main }),
      thcell("主な実施主体", { width: 14, fill: CH[5].main }),
    ]}),
    new TableRow({ children: [
      tcell("5-1", { align: AlignmentType.CENTER, bold: true, fill: C.lorange }),
      tcell("地域包括支援センターの体制強化【重点】", { bold: true, color: C.orange }),
      tcell("保健師3＋認定調査1の4名体制（実質3名）の強化。認定調査の負担軽減、ケアマネジャー配置の充実"),
      tcell("町・社協", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("5-2", { align: AlignmentType.CENTER, bold: true }),
      tcell("生活支援体制整備の充実"),
      tcell("生活支援コーディネーター（SC）25名による地域づくり、ユニバーサルサポーターとの連動"),
      tcell("町・社協・SC", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("5-3", { align: AlignmentType.CENTER, bold: true }),
      tcell("重層的支援体制（3計画連動）"),
      tcell("地域福祉計画・障害者計画と連動した重層的支援体制、複合課題世帯への対応"),
      tcell("町・包括・社協", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("5-4", { align: AlignmentType.CENTER, bold: true }),
      tcell("自立支援型地域ケア会議"),
      tcell("ケアプラン点検・自立支援に資する地域ケア会議の継続実施"),
      tcell("包括", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("5-5", { align: AlignmentType.CENTER, bold: true }),
      tcell("権利擁護（成年後見・虐待対応）"),
      tcell("成年後見制度の利用促進、高齢者虐待防止"),
      tcell("町・包括", { align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(e55Table);

S.push(subsection("成果指標（KPI）まとめ", 5));
S.push(p("各施策のKPI設定は、令和8年7月末のアンケート回収後にVer.2.0で確定します。本素案v1.0では指標項目のみを示します。"));

const k55Table = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("施策分野", { width: 30, fill: CH[5].main }),
      thcell("主要KPI候補", { width: 50, fill: CH[5].main }),
      thcell("データソース", { width: 20, fill: CH[5].main }),
    ]}),
    new TableRow({ children: [
      tcell("健康・介護予防", { bold: true }),
      tcell("通いの場参加率／特定健診受診率／サポーター活動回数"),
      tcell("町実績・アンケート", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("見守り・移動支援", { bold: true }),
      tcell("ふれあいNW活動員数／移動支援3制度の周知度／外出に困っている高齢者割合"),
      tcell("町実績・アンケート", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("在宅生活支援", { bold: true }),
      tcell("在宅医療連携件数／レスパイト利用率／介護離職率"),
      tcell("町実績・アンケート", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("人材・サービス", { bold: true }),
      tcell("町内事業所職員数／処遇改善加算取得率／ICT導入事業所数"),
      tcell("事業所照会", { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("包括ケア深化", { bold: true }),
      tcell("総合相談件数／自立支援型ケア会議開催数／成年後見利用件数"),
      tcell("包括実績", { align: AlignmentType.CENTER }),
    ]}),
  ],
});
S.push(k55Table);

// ===========================================================
// 第6章 認知症施策推進計画（独立章）
// ===========================================================
S.push(...chapterTitle(6));

S.push(p("本章は、令和6年1月に施行された「共生社会の実現を推進するための認知症基本法」第14条に基づく市町村認知症施策推進計画として、本計画に独立章として位置付けるものです。"));

S.push(p("認知症基本法は、認知症の人やその家族が暮らしやすい共生社会の実現を目的とし、市町村に対して認知症施策推進計画の策定を努力義務として課しています。川崎町では、第10期計画期間（令和9〜11年度）に向けて、認知症施策を独立章として位置付けることで、認知症基本法の理念・基本的施策に対応した包括的な計画体系を構築します。"));

// 6-1
S.push(section(6, 1, "認知症基本法と川崎町の対応方針"));

S.push(subsection("認知症基本法の概要", 6));
S.push(p("認知症基本法は、認知症の人がその尊厳を保持しつつ希望をもって暮らすことができるよう、認知症施策を総合的かつ計画的に推進することを目的とする法律です。基本理念として「全ての認知症の人が、基本的人権を享有する個人として、自らの意思によって日常生活及び社会生活を営むことができるようにすること」が掲げられています（第3条）。"));

S.push(subsection("川崎町の対応方針", 6));
S.push(fact("第10期計画では、認知症基本法対応として基本目標6を新設し、独立章（第6章）として認知症施策推進計画を位置付ける。"));

S.push(p("川崎町では、第10期計画期間（R9〜R11）を通じて、以下の方針で認知症施策を推進します。"));
S.push(numItem("①", "本人意見の反映（基本法第3条）：認知症本人・家族の意見を地域包括支援センター及び国保川崎病院経由で聴取し、計画・施策に反映"));
S.push(numItem("②", "共生社会の実現：認知症があっても住み慣れた地域で自分らしく暮らせる地域づくり（サブタイトル「〜認知症になっても誰もが自分らしく暮らせる地域共生社会の実現〜」）"));
S.push(numItem("③", "7基本的施策の体系化：基本法第15〜21条の7基本的施策に対応した町施策の体系化"));
S.push(numItem("④", "KPIの3層構造化：プロセス（活動量）・アウトプット（成果物）・アウトカム（住民の変化）の3層でKPIを設定"));

// 6-2
S.push(section(6, 2, "認知症施策の体系（7基本的施策）"));

S.push(p("認知症基本法第15条から第21条までの7基本的施策と、川崎町の対応施策を以下のとおり対応付けます。"));

const kihonTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("条文", { width: 10, fill: CH[6].main }),
      thcell("基本的施策", { width: 28, fill: CH[6].main }),
      thcell("川崎町の対応", { width: 62, fill: CH[6].main }),
    ]}),
    new TableRow({ children: [
      tcell("第15条", { align: AlignmentType.CENTER, bold: true, fill: CH[6].sub }),
      tcell("認知症の人に関する国民の理解の増進等", { bold: true }),
      tcell("認知症サポーター養成講座の継続（累計550名）。企業・学校サポーター養成の拡大。世界アルツハイマー月間（9月）に合わせた啓発"),
    ]}),
    new TableRow({ children: [
      tcell("第16条", { align: AlignmentType.CENTER, bold: true, fill: CH[6].sub }),
      tcell("認知症の人の生活におけるバリアフリー化の推進", { bold: true }),
      tcell("公共施設・交通・買い物環境におけるバリアフリー化、認知症の人にやさしい地域づくり"),
    ]}),
    new TableRow({ children: [
      tcell("第17条", { align: AlignmentType.CENTER, bold: true, fill: CH[6].sub }),
      tcell("認知症の人の社会参加の機会の確保等", { bold: true }),
      tcell("認知症カフェ「喫茶みかん」の運営継続。認知症本人ミーティング・ピアサポート活動の場の確保"),
    ]}),
    new TableRow({ children: [
      tcell("第18条", { align: AlignmentType.CENTER, bold: true, fill: CH[6].sub }),
      tcell("認知症の人の意思決定の支援及び権利利益の保護", { bold: true }),
      tcell("成年後見制度の利用促進、日常生活自立支援事業の活用、意思決定支援研修の実施"),
    ]}),
    new TableRow({ children: [
      tcell("第19条", { align: AlignmentType.CENTER, bold: true, fill: CH[6].sub }),
      tcell("保健医療・福祉サービスの提供体制の整備等", { bold: true }),
      tcell("認知症初期集中支援チームの強化、認知症地域支援推進員の配置、国保川崎病院との連携強化"),
    ]}),
    new TableRow({ children: [
      tcell("第20条", { align: AlignmentType.CENTER, bold: true, fill: CH[6].sub }),
      tcell("相談体制の整備等", { bold: true }),
      tcell("地域包括支援センター（社協運営）における認知症相談窓口の充実、もの忘れ相談の継続"),
    ]}),
    new TableRow({ children: [
      tcell("第21条", { align: AlignmentType.CENTER, bold: true, fill: CH[6].sub }),
      tcell("研究等の推進", { bold: true }),
      tcell("認知症本人・家族の意見聴取（基本法第3条対応）、認知症施策への当事者参画"),
    ]}),
  ],
});
S.push(kihonTable);

// 6-3
S.push(section(6, 3, "重点施策とKPI"));

S.push(p("認知症基本法の7基本的施策（第15条〜第21条）を踏まえつつ、本町の地域資源・体制・現状を勘案して、第10期計画期間中に重点的に取り組む施策を5本柱として位置付けます。これらは、認知症サポーター累計550名・キャラバンメイト73名・認知症カフェ「喫茶みかん」の運営実績・初期集中支援チームの稼働等、本町がこれまで積み上げてきた基盤を活かしつつ、未整備のチームオレンジ・本人ミーティングを新規整備として加えることで、認知症の人本人・家族への包括的支援を実現する構成です。"));

S.push(subsection("重点施策（5本柱）", 6));

S.push(p("J-1（サポーター拡大）とJ-3（認知症カフェ等）は既存施策の質的拡充、J-2（チームオレンジ）とJ-3（本人ミーティング）の一部は基本法対応の新規整備、J-4（早期発見・対応）とJ-5（医療連携）は本町の医療資源（国民健康保険川崎病院）を活かした体制強化として位置付けます。"));

const cogJuTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("No", { width: 8, fill: CH[6].main }),
      thcell("重点施策", { width: 28, fill: CH[6].main }),
      thcell("内容", { width: 64, fill: CH[6].main }),
    ]}),
    new TableRow({ children: [
      tcell("J-1", { align: AlignmentType.CENTER, bold: true, fill: CH[6].sub }),
      tcell("認知症サポーターの拡大と質向上", { bold: true }),
      tcell("累計550名のサポーターを基盤に、企業・学校サポーターを増やし、活動の質を向上。キャラバンメイト（累計73名）の活動継続"),
    ]}),
    new TableRow({ children: [
      tcell("J-2", { align: AlignmentType.CENTER, bold: true, fill: CH[6].sub }),
      tcell("チームオレンジの整備", { bold: true }),
      tcell("認知症サポーターのうち、地域でステップアップしたメンバーによるチームオレンジを整備。本人・家族支援を実践"),
    ]}),
    new TableRow({ children: [
      tcell("J-3", { align: AlignmentType.CENTER, bold: true, fill: CH[6].sub }),
      tcell("認知症カフェ・本人ミーティング", { bold: true }),
      tcell("認知症カフェ「喫茶みかん」の運営継続。認知症本人ミーティングの新設"),
    ]}),
    new TableRow({ children: [
      tcell("J-4", { align: AlignmentType.CENTER, bold: true, fill: CH[6].sub }),
      tcell("早期発見・早期対応の体制強化", { bold: true }),
      tcell("認知症初期集中支援チームの強化、認知症地域支援推進員の活動、もの忘れ相談の継続"),
    ]}),
    new TableRow({ children: [
      tcell("J-5", { align: AlignmentType.CENTER, bold: true, fill: CH[6].sub }),
      tcell("国保川崎病院との医療連携", { bold: true }),
      tcell("認知症診断・治療における国保川崎病院との連携、認知症初期段階からの医療フォロー"),
    ]}),
  ],
});
S.push(cogJuTable);

S.push(subsection("KPIの3層構造", 6));
S.push(p("認知症施策のKPIは、プロセス（活動量）・アウトプット（成果物）・アウトカム（住民の変化）の3層で設定します。これは、活動回数や成果物の数だけを追うのではなく、最終的に「認知症の人本人と家族の地域生活満足度」がどう変化したかをアンケートで測定し、施策の真の有効性を評価するための設計です。プロセス・アウトプット指標は事業実績で測定可能ですが、アウトカム指標はアンケート結果反映後の確定となるため、目標値は委員会協議事項として整理します。"));

const cogKpiTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  rows: [
    new TableRow({ children: [
      thcell("層", { width: 14, fill: CH[6].main }),
      thcell("指標", { width: 42, fill: CH[6].main }),
      thcell("現状値", { width: 22, fill: CH[6].main }),
      thcell("目標値（R11）", { width: 22, fill: CH[6].main }),
    ]}),
    new TableRow({ children: [
      tcell("プロセス\n（活動量）", { bold: true, fill: CH[6].sub, align: AlignmentType.CENTER }),
      tcell("認知症サポーター養成講座 年間開催回数"),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("【委員会協議】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("プロセス", { bold: true, fill: CH[6].sub, align: AlignmentType.CENTER }),
      tcell("認知症カフェ 年間開催回数"),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("【委員会協議】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("アウトプット\n（成果物）", { bold: true, fill: CH[6].sub, align: AlignmentType.CENTER }),
      tcell("認知症サポーター累計養成数"),
      tcell("550名", { align: AlignmentType.CENTER }),
      tcell("【委員会協議】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("アウトプット", { bold: true, fill: CH[6].sub, align: AlignmentType.CENTER }),
      tcell("チームオレンジ整備状況"),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("整備", { align: AlignmentType.CENTER, bold: true }),
    ]}),
    new TableRow({ children: [
      tcell("アウトプット", { bold: true, fill: CH[6].sub, align: AlignmentType.CENTER }),
      tcell("認知症初期集中支援チーム 訪問件数"),
      tcell("【町確認】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("【委員会協議】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ children: [
      tcell("アウトカム\n（住民の変化）", { bold: true, fill: CH[6].sub, align: AlignmentType.CENTER }),
      tcell("認知症相談窓口を知っている人の割合"),
      tcell("【調査後設定】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("増加", { align: AlignmentType.CENTER, bold: true }),
    ]}),
    new TableRow({ children: [
      tcell("アウトカム", { bold: true, fill: CH[6].sub, align: AlignmentType.CENTER }),
      tcell("認知症の人と家族の地域生活満足度"),
      tcell("【調査後設定】", { color: C.orange, italics: true, align: AlignmentType.CENTER }),
      tcell("増加", { align: AlignmentType.CENTER, bold: true }),
    ]}),
  ],
});
S.push(cogKpiTable);

S.push(placeholder("認知症本人・家族の意見聴取は、郵送のアンケート調査とは別チャネル（地域包括支援センター・国保川崎病院経由）で実施します。聴取結果はVer.2.0で本章に反映します。"));

module.exports = { S };
console.log("Part2 (Ch5-6) ready. Blocks:", S.length);
