const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const { FaBrain, FaSearch, FaBell, FaComments, FaRocket, FaShieldAlt, FaLink, FaChartLine, FaChartBar, FaCogs, FaFileAlt, FaCheckCircle, FaExclamationTriangle, FaExclamationCircle, FaInfoCircle, FaArrowRight, FaArrowDown, FaPlay, FaPause, FaSync, FaUser, FaEnvelope, FaMicrophone, FaDatabase, FaServer, FaDesktop, FaMobile, FaPlus, FaMinus, FaTimes, FaChevronRight, FaFilter, FaDownload, FaShare, FaEdit, FaTrash } = require("react-icons/fa");

const COLORS = {
  primary: "0A1628",
  secondary: "1A365D",
  accent: "F6AD55",
  accentDark: "DD6B20",
  white: "FFFFFF",
  lightGray: "F7FAFC",
  cardBg: "FFFFFF",
  textDark: "1A202C",
  textMuted: "718096",
  lobster: "E53E3E",
  lobsterDark: "C53030",
  holmesBlue: "3182CE",
  holmesGreen: "38A169",
  holmesPurple: "805AD5",
  holmesTeal: "319795",
  success: "38A169",
  warning: "D69E2E",
  danger: "E53E3E",
  info: "3182CE",
};

function renderIconSvg(IconComponent, color = "#000000", size = 256) {
  return ReactDOMServer.renderToStaticMarkup(React.createElement(IconComponent, { color, size: String(size) }));
}

async function iconToBase64Png(IconComponent, color, size = 256) {
  const svg = renderIconSvg(IconComponent, color, size);
  const pngBuffer = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + pngBuffer.toString("base64");
}

const makeCardShadow = () => ({ type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.1 });
const makeSmallShadow = () => ({ type: "outer", color: "000000", blur: 4, offset: 1, angle: 135, opacity: 0.08 });

async function createPresentation() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = "雲端全知龍蝦 - Holmes 智能代理人完整規劃";
  pres.author = "AI大數據部門";

  const icons = {
    brain: await iconToBase64Png(FaBrain, "#FFFFFF", 256),
    search: await iconToBase64Png(FaSearch, "#FFFFFF", 256),
    bell: await iconToBase64Png(FaBell, "#FFFFFF", 256),
    comments: await iconToBase64Png(FaComments, "#FFFFFF", 256),
    rocket: await iconToBase64Png(FaRocket, "#FFFFFF", 256),
    shield: await iconToBase64Png(FaShieldAlt, "#FFFFFF", 256),
    link: await iconToBase64Png(FaLink, "#F6AD55", 256),
    chart: await iconToBase64Png(FaChartLine, "#FFFFFF", 256),
    bar: await iconToBase64Png(FaChartBar, "#FFFFFF", 256),
    cogs: await iconToBase64Png(FaCogs, "#FFFFFF", 256),
    file: await iconToBase64Png(FaFileAlt, "#FFFFFF", 256),
    check: await iconToBase64Png(FaCheckCircle, "#38A169", 256),
    warning: await iconToBase64Png(FaExclamationTriangle, "#DD6B20", 256),
    danger: await iconToBase64Png(FaExclamationCircle, "#E53E3E", 256),
    info: await iconToBase64Png(FaInfoCircle, "#3182CE", 256),
    arrow: await iconToBase64Png(FaArrowRight, "#F6AD55", 256),
    arrowDown: await iconToBase64Png(FaArrowDown, "#F6AD55", 256),
    sync: await iconToBase64Png(FaSync, "#F6AD55", 256),
    user: await iconToBase64Png(FaUser, "#FFFFFF", 256),
    envelope: await iconToBase64Png(FaEnvelope, "#FFFFFF", 256),
    mic: await iconToBase64Png(FaMicrophone, "#FFFFFF", 256),
    database: await iconToBase64Png(FaDatabase, "#FFFFFF", 256),
    server: await iconToBase64Png(FaServer, "#FFFFFF", 256),
    desktop: await iconToBase64Png(FaDesktop, "#FFFFFF", 256),
    mobile: await iconToBase64Png(FaMobile, "#FFFFFF", 256),
    plus: await iconToBase64Png(FaPlus, "#FFFFFF", 256),
    minus: await iconToBase64Png(FaMinus, "#FFFFFF", 256),
    times: await iconToBase64Png(FaTimes, "#FFFFFF", 256),
    chevronR: await iconToBase64Png(FaChevronRight, "#718096", 256),
    filter: await iconToBase64Png(FaFilter, "#718096", 256),
    download: await iconToBase64Png(FaDownload, "#3182CE", 256),
    share: await iconToBase64Png(FaShare, "#3182CE", 256),
    edit: await iconToBase64Png(FaEdit, "#38A169", 256),
    trash: await iconToBase64Png(FaTrash, "#E53E3E", 256),
  };

  // ============ SLIDE 1: Title ============
  let slide1 = pres.addSlide();
  slide1.background = { color: COLORS.primary };
  slide1.addShape(pres.shapes.OVAL, { x: 3.8, y: 0.5, w: 2.4, h: 2.4, fill: { color: COLORS.lobster }, shadow: { type: "outer", color: "000000", blur: 20, offset: 6, angle: 135, opacity: 0.35 } });
  slide1.addImage({ data: icons.brain, x: 4.2, y: 0.9, w: 1.6, h: 1.6 });
  slide1.addText("雲端全知龍蝦", { x: 0.5, y: 3.1, w: 9, h: 0.9, fontSize: 48, fontFace: "Arial Black", color: COLORS.white, bold: true, align: "center" });
  slide1.addText("Holmes 全產品線 智能代理人", { x: 0.5, y: 3.9, w: 9, h: 0.6, fontSize: 24, fontFace: "Arial", color: COLORS.accent, align: "center" });
  slide1.addText("深度理解 Holmes、串連產品關係、主動幫助決策", { x: 0.5, y: 4.6, w: 9, h: 0.5, fontSize: 14, fontFace: "Arial", color: COLORS.white, align: "center" });
  slide1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.1, w: 10, h: 0.525, fill: { color: COLORS.lobster } });
  slide1.addText("AI大數據部門  |  智能製造規劃", { x: 0.5, y: 5.15, w: 9, h: 0.45, fontSize: 12, fontFace: "Arial", color: COLORS.white, align: "center", valign: "middle" });

  // ============ SLIDE 2: Holmes Products ============
  let slide2 = pres.addSlide();
  slide2.background = { color: COLORS.lightGray };
  slide2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: COLORS.primary } });
  slide2.addText("全知龍蝦深度理解 Holmes 全產品線", { x: 0.5, y: 0.25, w: 9, h: 0.55, fontSize: 26, fontFace: "Arial Black", color: COLORS.white, bold: true, margin: 0 });

  const products = [
    { name: "智能生產 Holmes(P)", color: COLORS.holmesBlue, icon: icons.cogs, items: ["產能優化工具", "瓶頸根因分析", "OEE 追蹤", "換線時間分析", "資源利用率優化"] },
    { name: "智能品質 Holmes(Q)", color: COLORS.holmesGreen, icon: icons.chart, items: ["良品變異分析 (CTQ站)", "產品健康指數分析", "敘述統計分析 (CPK/圖表)", "不良品歸類分析", "SoH Burn-in 變異分析", "NPE 量產前驗證"] },
    { name: "知識管理 AI會議室", color: COLORS.holmesPurple, icon: icons.comments, items: ["常規會議室 (日常檢討)", "即時戰情室 (全天候監控)", "深度調研室 (專家問診)"] }
  ];

  const pW = 3.0, pH = 4.0, pStartX = 0.5, pStartY = 1.15, pGapX = 0.2;
  products.forEach((p, i) => {
    const x = pStartX + i * (pW + pGapX);
    const y = pStartY;
    slide2.addShape(pres.shapes.RECTANGLE, { x, y, w: pW, h: pH, fill: { color: COLORS.cardBg }, shadow: makeCardShadow() });
    slide2.addShape(pres.shapes.RECTANGLE, { x, y, w: pW, h: 0.65, fill: { color: p.color } });
    slide2.addImage({ data: p.icon, x: x + 0.15, y: y + 0.1, w: 0.45, h: 0.45 });
    slide2.addText(p.name, { x: x + 0.65, y: y + 0.1, w: 2.2, h: 0.45, fontSize: 11, fontFace: "Arial", color: COLORS.white, bold: true, valign: "middle", margin: 0 });
    p.items.forEach((item, j) => {
      slide2.addText([{ text: "▸ ", options: { color: p.color, bold: true } }, { text: item, options: { color: COLORS.textDark } }], { x: x + 0.15, y: y + 0.8 + j * 0.42, w: pW - 0.3, h: 0.38, fontSize: 10, fontFace: "Arial", margin: 0 });
    });
  });

  // ============ SLIDE 3: Cross-Product Value ============
  let slide3 = pres.addSlide();
  slide3.background = { color: COLORS.lightGray };
  slide3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: COLORS.primary } });
  slide3.addImage({ data: icons.link, x: 0.5, y: 0.2, w: 0.6, h: 0.6 });
  slide3.addText("核心價值：串連產品間的關係", { x: 1.2, y: 0.25, w: 8, h: 0.55, fontSize: 26, fontFace: "Arial Black", color: COLORS.white, bold: true, margin: 0 });

  slide3.addShape(pres.shapes.OVAL, { x: 4.0, y: 2.0, w: 2.0, h: 2.0, fill: { color: COLORS.lobster }, shadow: { type: "outer", color: "000000", blur: 15, offset: 5, angle: 135, opacity: 0.3 } });
  slide3.addImage({ data: icons.brain, x: 4.3, y: 2.3, w: 1.4, h: 1.4 });
  slide3.addText("全知龍蝦", { x: 3.5, y: 4.1, w: 3.0, h: 0.4, fontSize: 14, fontFace: "Arial", color: COLORS.lobster, bold: true, align: "center" });

  const rels = [
    { name: "Holmes(P)\n產能瓶頸", x: 1.0, y: 1.3, color: COLORS.holmesBlue },
    { name: "Holmes(Q)\n良品變異", x: 7.0, y: 1.3, color: COLORS.holmesGreen },
    { name: "Holmes(Q)\n不良品歸類", x: 0.3, y: 3.5, color: COLORS.holmesGreen },
    { name: "AI會議室\n戰情室", x: 7.7, y: 3.5, color: COLORS.holmesPurple },
  ];
  rels.forEach((r) => {
    slide3.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: r.x, y: r.y, w: 1.8, h: 0.9, fill: { color: r.color }, rectRadius: 0.1 });
    slide3.addText(r.name, { x: r.x, y: r.y, w: 1.8, h: 0.9, fontSize: 10, fontFace: "Arial", color: COLORS.white, bold: true, align: "center", valign: "middle" });
  });

  slide3.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.65, w: 9, h: 0.9, fill: { color: COLORS.accent, transparency: 15 } });
  slide3.addText("情境範例：當 Holmes(P) 發現產能下降 → 全知龍蝦自動查詢 Holmes(Q) 良品數據 → 發現良率下降 → 調用不良品歸類分析 → 識別是新供應商材料問題 → 自動生成報告並建議召回 → 視需要召集 AI會議室 戰情室", { x: 0.7, y: 4.75, w: 8.6, h: 0.7, fontSize: 9, fontFace: "Arial", color: COLORS.accentDark, valign: "middle" });

  // ============ SLIDE 4: Scenario 1 - Daily Report ============
  let slide4 = pres.addSlide();
  slide4.background = { color: COLORS.lightGray };
  slide4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: COLORS.primary } });
  slide4.addShape(pres.shapes.OVAL, { x: 0.4, y: 0.15, w: 0.7, h: 0.7, fill: { color: COLORS.accent } });
  slide4.addText("01", { x: 0.4, y: 0.25, w: 0.7, h: 0.5, fontSize: 18, fontFace: "Arial", color: COLORS.white, bold: true, align: "center", valign: "middle" });
  slide4.addText("情境一：每日工廠健康報告主動生成", { x: 1.2, y: 0.25, w: 8, h: 0.55, fontSize: 24, fontFace: "Arial Black", color: COLORS.white, bold: true, margin: 0 });

  slide4.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.15, w: 4.5, h: 4.4, fill: { color: COLORS.cardBg }, shadow: makeCardShadow() });
  slide4.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.15, w: 0.08, h: 4.4, fill: { color: COLORS.holmesBlue } });
  slide4.addText("觸發方式", { x: 0.75, y: 1.3, w: 4.0, h: 0.35, fontSize: 12, fontFace: "Arial", color: COLORS.holmesBlue, bold: true, margin: 0 });
  slide4.addText("每日早上 08:00 自動執行\n或 管理層透過 LINE/Email 主動要求", { x: 0.75, y: 1.6, w: 4.0, h: 0.6, fontSize: 10, fontFace: "Arial", color: COLORS.textMuted, margin: 0 });

  slide4.addText("龍蝦主動行為", { x: 0.75, y: 2.3, w: 4.0, h: 0.35, fontSize: 12, fontFace: "Arial", color: COLORS.accent, bold: true, margin: 0 });
  const s1Steps = ["1. 調用 Holmes(P) 取得昨日各線體 OEE", "2. 調用 Holmes(Q) 取得良率數據", "3. 跨廠區横比，識別異常工廠", "4. 檢查是否有設備異常預警", "5. 查詢 AI會議室 是否有待追蹤問題", "6. 整合生成圖文報告"];
  s1Steps.forEach((s, i) => { slide4.addText(s, { x: 0.75, y: 2.65 + i * 0.38, w: 4.0, h: 0.35, fontSize: 9, fontFace: "Arial", color: COLORS.textDark, margin: 0 }); });
  slide4.addText("輸出成果", { x: 0.75, y: 4.9, w: 4.0, h: 0.35, fontSize: 12, fontFace: "Arial", color: COLORS.holmesGreen, bold: true, margin: 0 });
  slide4.addText("LINE 推播 + Email 附件（PDF/Excel）", { x: 0.75, y: 5.2, w: 4.0, h: 0.3, fontSize: 9, fontFace: "Arial", color: COLORS.textMuted, margin: 0 });

  slide4.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.15, w: 4.3, h: 4.4, fill: { color: COLORS.cardBg }, shadow: makeCardShadow() });
  slide4.addText("📊 報告預覽", { x: 5.4, y: 1.3, w: 3.9, h: 0.35, fontSize: 11, fontFace: "Arial", color: COLORS.textMuted, margin: 0 });
  const bars = [{ h: 1.2, label: "工廠A", color: COLORS.holmesGreen }, { h: 0.8, label: "工廠B", color: COLORS.warning }, { h: 1.5, label: "工廠C", color: COLORS.holmesGreen }];
  bars.forEach((b, i) => {
    slide4.addShape(pres.shapes.RECTANGLE, { x: 5.5 + i * 1.2, y: 4.5 - b.h, w: 0.8, h: b.h, fill: { color: b.color } });
    slide4.addText(b.label, { x: 5.5 + i * 1.2, y: 4.55, w: 0.8, h: 0.3, fontSize: 8, fontFace: "Arial", color: COLORS.textMuted, align: "center" });
  });
  slide4.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: 1.8, w: 3.9, h: 0.9, fill: { color: COLORS.warning, transparency: 15 }, line: { color: COLORS.warning, width: 1 } });
  slide4.addImage({ data: icons.warning, x: 5.55, y: 1.95, w: 0.35, h: 0.35 });
  slide4.addText("⚠️ 工廠B 良率低於預期 3%", { x: 6.0, y: 1.95, w: 3.1, h: 0.35, fontSize: 10, fontFace: "Arial", color: COLORS.warning, bold: true, margin: 0 });
  slide4.addText("建議今日召開戰情室檢討", { x: 6.0, y: 2.3, w: 3.1, h: 0.3, fontSize: 9, fontFace: "Arial", color: COLORS.textMuted, margin: 0 });
  const stats = [{ label: "平均良率", value: "98.2%" }, { label: "平均OEE", value: "85.6%" }, { label: "異常預警", value: "2件" }];
  stats.forEach((st, i) => {
    slide4.addShape(pres.shapes.RECTANGLE, { x: 5.4 + i * 1.35, y: 2.85, w: 1.2, h: 0.7, fill: { color: COLORS.lightGray } });
    slide4.addText(st.value, { x: 5.4 + i * 1.35, y: 2.9, w: 1.2, h: 0.4, fontSize: 14, fontFace: "Arial", color: COLORS.textDark, bold: true, align: "center" });
    slide4.addText(st.label, { x: 5.4 + i * 1.35, y: 3.25, w: 1.2, h: 0.25, fontSize: 7, fontFace: "Arial", color: COLORS.textMuted, align: "center" });
  });

  // ============ SLIDE 5: Scenario 2 - Equipment Failure ============
  let slide5 = pres.addSlide();
  slide5.background = { color: COLORS.lightGray };
  slide5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: COLORS.primary } });
  slide5.addShape(pres.shapes.OVAL, { x: 0.4, y: 0.15, w: 0.7, h: 0.7, fill: { color: COLORS.accent } });
  slide5.addText("02", { x: 0.4, y: 0.25, w: 0.7, h: 0.5, fontSize: 18, fontFace: "Arial", color: COLORS.white, bold: true, align: "center", valign: "middle" });
  slide5.addText("情境二：設備故障預測與訂單調度", { x: 1.2, y: 0.25, w: 8, h: 0.55, fontSize: 24, fontFace: "Arial Black", color: COLORS.white, bold: true, margin: 0 });

  const timeline = [
    { time: "T+0", event: "邊端龍蝦偵測振動異常", detail: "工廠A 加工中心軸承高頻振動上升15%", icon: icons.warning, color: COLORS.warning },
    { time: "T+5min", event: "雲端龍蝦接收事件", detail: "分析：軸承使用72天，預測7-10天後達閾值", icon: icons.brain, color: COLORS.lobster },
    { time: "T+10min", event: "查詢跨廠產能", detail: "工廠B閒置產能60%，可接收訂單", icon: icons.search, color: COLORS.holmesBlue },
    { time: "T+15min", event: "生成應變方案", detail: "方案：訂單分流+預約維修窗口", icon: icons.rocket, color: COLORS.holmesGreen },
    { time: "T+20min", event: "管理層核准", detail: "確認：同意執行，維修安排於下週三", icon: icons.check, color: COLORS.success },
  ];
  timeline.forEach((t, i) => {
    const y = 1.2 + i * 0.85;
    slide5.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y, w: 1.1, h: 0.5, fill: { color: t.color }, rectRadius: 0.05 });
    slide5.addText(t.time, { x: 0.5, y, w: 1.1, h: 0.5, fontSize: 10, fontFace: "Arial", color: COLORS.white, bold: true, align: "center", valign: "middle" });
    if (i < timeline.length - 1) slide5.addShape(pres.shapes.LINE, { x: 1.05, y: y + 0.5, w: 0, h: 0.35, line: { color: COLORS.textMuted, width: 1.5, dashType: "dash" } });
    slide5.addShape(pres.shapes.RECTANGLE, { x: 1.8, y: y - 0.05, w: 3.8, h: 0.6, fill: { color: COLORS.cardBg }, shadow: makeSmallShadow() });
    slide5.addImage({ data: t.icon, x: 1.95, y: y + 0.05, w: 0.4, h: 0.4 });
    slide5.addText(t.event, { x: 2.45, y, w: 3.0, h: 0.35, fontSize: 10, fontFace: "Arial", color: COLORS.textDark, bold: true, margin: 0 });
    slide5.addText(t.detail, { x: 2.45, y: y + 0.3, w: 3.0, h: 0.25, fontSize: 8, fontFace: "Arial", color: COLORS.textMuted, margin: 0 });
  });

  slide5.addShape(pres.shapes.RECTANGLE, { x: 5.9, y: 1.2, w: 3.6, h: 4.2, fill: { color: COLORS.cardBg }, shadow: makeCardShadow() });
  slide5.addShape(pres.shapes.RECTANGLE, { x: 5.9, y: 1.2, w: 3.6, h: 0.6, fill: { color: COLORS.success } });
  slide5.addText("最終成果", { x: 5.9, y: 1.3, w: 3.6, h: 0.4, fontSize: 14, fontFace: "Arial", color: COLORS.white, bold: true, align: "center" });
  slide5.addShape(pres.shapes.RECTANGLE, { x: 6.1, y: 2.0, w: 3.2, h: 1.2, fill: { color: COLORS.success, transparency: 10 }, line: { color: COLORS.success, width: 1 } });
  slide5.addText("原本可能的損失", { x: 6.2, y: 2.1, w: 3.0, h: 0.3, fontSize: 9, fontFace: "Arial", color: COLORS.textMuted, margin: 0 });
  slide5.addText("2-3 天非計劃停機", { x: 6.2, y: 2.4, w: 3.0, h: 0.35, fontSize: 14, fontFace: "Arial", color: COLORS.danger, bold: true, margin: 0 });
  slide5.addText("→ 訂單延遲 + 緊急維修成本", { x: 6.2, y: 2.75, w: 3.0, h: 0.3, fontSize: 9, fontFace: "Arial", color: COLORS.textMuted, margin: 0 });
  slide5.addShape(pres.shapes.RECTANGLE, { x: 6.1, y: 3.35, w: 3.2, h: 1.2, fill: { color: COLORS.holmesGreen, transparency: 10 }, line: { color: COLORS.holmesGreen, width: 1 } });
  slide5.addText("實際結果", { x: 6.2, y: 3.45, w: 3.0, h: 0.3, fontSize: 9, fontFace: "Arial", color: COLORS.textMuted, margin: 0 });
  slide5.addText("4 小時計劃性維護", { x: 6.2, y: 3.75, w: 3.0, h: 0.35, fontSize: 14, fontFace: "Arial", color: COLORS.holmesGreen, bold: true, margin: 0 });
  slide5.addText("→ 訂單順利分流，延遲 < 0.5 天", { x: 6.2, y: 4.1, w: 3.0, h: 0.3, fontSize: 9, fontFace: "Arial", color: COLORS.textMuted, margin: 0 });

  // ============ SLIDE 6: Scenario 3 - Quality Root Cause ============
  let slide6 = pres.addSlide();
  slide6.background = { color: COLORS.lightGray };
  slide6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: COLORS.primary } });
  slide6.addShape(pres.shapes.OVAL, { x: 0.4, y: 0.15, w: 0.7, h: 0.7, fill: { color: COLORS.accent } });
  slide6.addText("03", { x: 0.4, y: 0.25, w: 0.7, h: 0.5, fontSize: 18, fontFace: "Arial", color: COLORS.white, bold: true, align: "center", valign: "middle" });
  slide6.addText("情境三：跨產品品質異常根因分析", { x: 1.2, y: 0.25, w: 8, h: 0.55, fontSize: 24, fontFace: "Arial Black", color: COLORS.white, bold: true, margin: 0 });

  slide6.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 9, h: 1.5, fill: { color: COLORS.cardBg }, shadow: makeCardShadow() });
  slide6.addText("問題：工廠B良率突然下降 2%，持續3天", { x: 0.7, y: 1.35, w: 8.6, h: 0.4, fontSize: 14, fontFace: "Arial", color: COLORS.danger, bold: true, margin: 0 });

  const pFlows = [
    { name: "Holmes(Q)\n良品變異", x: 0.8, color: COLORS.holmesGreen, result: "發現 X-bar 偏移 > 2σ" },
    { name: "邊端龍蝦\n隔離批次", x: 2.8, color: COLORS.lobster, result: "隔離 1500pcs 待檢" },
    { name: "Holmes(Q)\n不良品歸類", x: 4.8, color: COLORS.holmesGreen, result: "組裝段不良率上升" },
    { name: "時間序列\n分析", x: 6.8, color: COLORS.secondary, result: "發現原材料供應商變更" },
  ];
  pFlows.forEach((pf, i) => {
    slide6.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: pf.x, y: 1.85, w: 1.8, h: 0.7, fill: { color: pf.color }, rectRadius: 0.1 });
    slide6.addText(pf.name, { x: pf.x, y: 1.85, w: 1.8, h: 0.7, fontSize: 9, fontFace: "Arial", color: COLORS.white, bold: true, align: "center", valign: "middle" });
    if (i < pFlows.length - 1) slide6.addImage({ data: icons.arrow, x: pf.x + 1.85, y: 2.05, w: 0.4, h: 0.3 });
  });

  slide6.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.85, w: 4.5, h: 2.65, fill: { color: COLORS.cardBg }, shadow: makeCardShadow() });
  slide6.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.85, w: 0.08, h: 2.65, fill: { color: COLORS.holmesBlue } });
  slide6.addText("雲端全知龍蝦分析過程", { x: 0.75, y: 3.0, w: 4.0, h: 0.35, fontSize: 12, fontFace: "Arial", color: COLORS.holmesBlue, bold: true, margin: 0 });
  const analysis = ["▸ 横比工廠A、C同段位數據 → 正常", "▸ 時間序列追溯 → 3天前供應商變更", "▸ 知識圖譜查詢 → 新供應商材料特性", "▸ AI推理 → 膨脹係數差異導致尺寸偏差", "▸ 驗證 → 隔離批次尺寸偏大 0.05mm"];
  analysis.forEach((a, i) => { slide6.addText(a, { x: 0.75, y: 3.4 + i * 0.4, w: 4.0, h: 0.35, fontSize: 10, fontFace: "Arial", color: COLORS.textDark, margin: 0 }); });

  slide6.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 2.85, w: 4.3, h: 2.65, fill: { color: COLORS.cardBg }, shadow: makeCardShadow() });
  slide6.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 2.85, w: 0.08, h: 2.65, fill: { color: COLORS.success } });
  slide6.addText("龍蝦建議行動", { x: 5.45, y: 3.0, w: 3.8, h: 0.35, fontSize: 12, fontFace: "Arial", color: COLORS.success, bold: true, margin: 0 });
  slide6.addShape(pres.shapes.RECTANGLE, { x: 5.45, y: 3.45, w: 3.8, h: 0.6, fill: { color: COLORS.success, transparency: 10 } });
  slide6.addText("建議恢復原供應商\n或調整加工參數（緊急訂單需交付）", { x: 5.6, y: 3.55, w: 3.5, h: 0.45, fontSize: 10, fontFace: "Arial", color: COLORS.textDark, margin: 0 });
  slide6.addText("龍蝦自動執行：", { x: 5.45, y: 4.2, w: 3.8, h: 0.3, fontSize: 10, fontFace: "Arial", color: COLORS.accent, bold: true, margin: 0 });
  const actions = ["✓ 生成品質異常報告", "✓ 通知採購與供應商協商", "✓ 持續監控調整後良率"];
  actions.forEach((ac, i) => { slide6.addText(ac, { x: 5.45, y: 4.5 + i * 0.3, w: 3.8, h: 0.28, fontSize: 9, fontFace: "Arial", color: COLORS.textMuted, margin: 0 }); });

  // ============ SLIDE 7: Scenario 4 - Emergency Order ============
  let slide7 = pres.addSlide();
  slide7.background = { color: COLORS.lightGray };
  slide7.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: COLORS.primary } });
  slide7.addShape(pres.shapes.OVAL, { x: 0.4, y: 0.15, w: 0.7, h: 0.7, fill: { color: COLORS.accent } });
  slide7.addText("04", { x: 0.4, y: 0.25, w: 0.7, h: 0.5, fontSize: 18, fontFace: "Arial", color: COLORS.white, bold: true, align: "center", valign: "middle" });
  slide7.addText("情境四：緊急訂單插單仲裁", { x: 1.2, y: 0.25, w: 8, h: 0.55, fontSize: 24, fontFace: "Arial Black", color: COLORS.white, bold: true, margin: 0 });

  slide7.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.15, w: 9, h: 0.8, fill: { color: COLORS.danger, transparency: 10 }, line: { color: COLORS.danger, width: 1.5 } });
  slide7.addText("🚨 緊急訂單接收：數量 = 工廠日產能 30%，三廠都在高負載運轉", { x: 0.7, y: 1.35, w: 8.6, h: 0.4, fontSize: 13, fontFace: "Arial", color: COLORS.danger, bold: true, margin: 0 });

  const factories = [
    { name: "工廠A", load: "85%", status: "有空閒產能", issue: "需犧牲現有訂單", color: COLORS.warning },
    { name: "工廠B", load: "95%", status: "加班可以消化", issue: "額外加班成本", color: COLORS.holmesGreen },
    { name: "工廠C", load: "70%", status: "最從容", issue: "規格有限制", color: COLORS.holmesBlue },
  ];
  factories.forEach((f, i) => {
    const x = 0.5 + i * 3.1;
    slide7.addShape(pres.shapes.RECTANGLE, { x, y: 2.1, w: 2.9, h: 1.5, fill: { color: COLORS.cardBg }, shadow: makeCardShadow() });
    slide7.addShape(pres.shapes.RECTANGLE, { x, y: 2.1, w: 2.9, h: 0.5, fill: { color: f.color } });
    slide7.addText(f.name, { x, y: 2.15, w: 2.9, h: 0.4, fontSize: 14, fontFace: "Arial", color: COLORS.white, bold: true, align: "center" });
    slide7.addText(f.load + " 負載", { x, y: 2.7, w: 2.9, h: 0.35, fontSize: 16, fontFace: "Arial", color: COLORS.textDark, bold: true, align: "center" });
    slide7.addText(f.status, { x, y: 3.05, w: 2.9, h: 0.25, fontSize: 9, fontFace: "Arial", color: COLORS.textMuted, align: "center" });
    slide7.addText(f.issue, { x, y: 3.3, w: 2.9, h: 0.25, fontSize: 9, fontFace: "Arial", color: COLORS.warning, align: "center" });
  });

  slide7.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.75, w: 9, h: 1.8, fill: { color: COLORS.cardBg }, shadow: makeCardShadow() });
  slide7.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.75, w: 0.08, h: 1.8, fill: { color: COLORS.success } });
  slide7.addText("龍蝦多目標優化結果", { x: 0.75, y: 3.9, w: 8.5, h: 0.35, fontSize: 12, fontFace: "Arial", color: COLORS.success, bold: true, margin: 0 });

  const options = [
    { option: "方案A：工廠A+C聯合", delay: "0.3天", cost: "中等", impact: "低", rec: false },
    { option: "方案B：工廠C單獨", delay: "0天", cost: "最低", impact: "中", rec: false },
    { option: "方案C：工廠B加班", delay: "0天", cost: "高", impact: "低", rec: true },
  ];
  slide7.addText("方案", { x: 0.75, y: 4.3, w: 2.5, h: 0.3, fontSize: 9, fontFace: "Arial", color: COLORS.textMuted, bold: true });
  slide7.addText("延遲", { x: 3.25, y: 4.3, w: 1.5, h: 0.3, fontSize: 9, fontFace: "Arial", color: COLORS.textMuted, bold: true });
  slide7.addText("成本", { x: 4.75, y: 4.3, w: 1.5, h: 0.3, fontSize: 9, fontFace: "Arial", color: COLORS.textMuted, bold: true });
  slide7.addText("影響", { x: 6.25, y: 4.3, w: 1.5, h: 0.3, fontSize: 9, fontFace: "Arial", color: COLORS.textMuted, bold: true });
  options.forEach((o, i) => {
    const y = 4.65 + i * 0.35;
    if (o.rec) slide7.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: y - 0.05, w: 7.2, h: 0.35, fill: { color: COLORS.success, transparency: 15 } });
    slide7.addText(o.option, { x: 0.75, y, w: 2.5, h: 0.3, fontSize: 9, fontFace: "Arial", color: COLORS.textDark });
    slide7.addText(o.delay, { x: 3.25, y, w: 1.5, h: 0.3, fontSize: 9, fontFace: "Arial", color: COLORS.textDark });
    slide7.addText(o.cost, { x: 4.75, y, w: 1.5, h: 0.3, fontSize: 9, fontFace: "Arial", color: COLORS.textDark });
    slide7.addText(o.impact, { x: 6.25, y, w: 1.5, h: 0.3, fontSize: 9, fontFace: "Arial", color: COLORS.textDark });
    if (o.rec) slide7.addText("★ 推薦", { x: 7.75, y, w: 1.0, h: 0.3, fontSize: 9, fontFace: "Arial", color: COLORS.success, bold: true });
  });

  // ============ SLIDE 8: Scenario 5 - Knowledge Sharing ============
  let slide8 = pres.addSlide();
  slide8.background = { color: COLORS.lightGray };
  slide8.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: COLORS.primary } });
  slide8.addShape(pres.shapes.OVAL, { x: 0.4, y: 0.15, w: 0.7, h: 0.7, fill: { color: COLORS.accent } });
  slide8.addText("05", { x: 0.4, y: 0.25, w: 0.7, h: 0.5, fontSize: 18, fontFace: "Arial", color: COLORS.white, bold: true, align: "center", valign: "middle" });
  slide8.addText("情境五：跨廠區知識萃取與傳播", { x: 1.2, y: 0.25, w: 8, h: 0.55, fontSize: 24, fontFace: "Arial Black", color: COLORS.white, bold: true, margin: 0 });

  const kSteps = [
    { step: "1", title: "知識識別", desc: "工廠A焊接不良問題\n連續60天無復發", icon: icons.check, color: COLORS.success },
    { step: "2", title: "知識萃取", desc: "分析：預熱溫度+15°C\n焊接速度-10%", icon: icons.file, color: COLORS.holmesBlue },
    { step: "3", title: "知識傳播", desc: "向工廠B、C發送推播\n附具體建議", icon: icons.share, color: COLORS.holmesPurple },
    { step: "4", title: "落地驗證", desc: "協助B、C廠進行\n參數差異分析", icon: icons.edit, color: COLORS.warning },
  ];
  kSteps.forEach((k, i) => {
    const x = 0.5 + i * 2.4;
    slide8.addShape(pres.shapes.RECTANGLE, { x, y: 1.2, w: 2.2, h: 2.4, fill: { color: COLORS.cardBg }, shadow: makeCardShadow() });
    slide8.addShape(pres.shapes.OVAL, { x: x + 0.8, y: 1.35, w: 0.6, h: 0.6, fill: { color: k.color } });
    slide8.addText(k.step, { x: x + 0.8, y: 1.42, w: 0.6, h: 0.45, fontSize: 16, fontFace: "Arial", color: COLORS.white, bold: true, align: "center", valign: "middle" });
    slide8.addText(k.title, { x: x + 0.1, y: 2.05, w: 2.0, h: 0.35, fontSize: 12, fontFace: "Arial", color: COLORS.textDark, bold: true, align: "center" });
    slide8.addText(k.desc, { x: x + 0.1, y: 2.45, w: 2.0, h: 1.0, fontSize: 9, fontFace: "Arial", color: COLORS.textMuted, align: "center" });
    if (i < kSteps.length - 1) slide8.addImage({ data: icons.arrow, x: x + 2.2, y: 2.2, w: 0.5, h: 0.3 });
  });

  slide8.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.8, w: 9, h: 1.75, fill: { color: COLORS.cardBg }, shadow: makeCardShadow() });
  slide8.addText("知識圖譜結構", { x: 0.7, y: 3.95, w: 8.6, h: 0.35, fontSize: 11, fontFace: "Arial", color: COLORS.textMuted, margin: 0 });

  const nodes = [
    { text: "焊接不良", x: 1.2, y: 4.5, color: COLORS.lobster, size: 0.9 },
    { text: "預熱溫度", x: 2.8, y: 4.3, color: COLORS.holmesBlue, size: 0.7 },
    { text: "焊接速度", x: 2.8, y: 5.0, color: COLORS.holmesBlue, size: 0.7 },
    { text: "最佳化參數", x: 4.5, y: 4.65, color: COLORS.holmesGreen, size: 0.8 },
    { text: "工廠A經驗", x: 6.3, y: 4.5, color: COLORS.holmesPurple, size: 0.7 },
    { text: "推播至B、C", x: 8.0, y: 4.65, color: COLORS.accent, size: 0.7 },
  ];
  const connections = [[0, 1], [0, 2], [1, 3], [2, 3], [3, 4], [4, 5]];
  connections.forEach(([from, to]) => {
    const x1 = nodes[from].x + nodes[from].size / 2, y1 = nodes[from].y + nodes[from].size / 2;
    const x2 = nodes[to].x + nodes[to].size / 2, y2 = nodes[to].y + nodes[to].size / 2;
    slide8.addShape(pres.shapes.LINE, { x: x1, y: y1, w: x2 - x1, h: y2 - y1, line: { color: COLORS.textMuted, width: 1, dashType: "dash" } });
  });
  nodes.forEach((n) => {
    slide8.addShape(pres.shapes.OVAL, { x: n.x, y: n.y, w: n.size, h: n.size, fill: { color: n.color } });
    slide8.addText(n.text, { x: n.x, y: n.y, w: n.size, h: n.size, fontSize: 7, fontFace: "Arial", color: COLORS.white, bold: true, align: "center", valign: "middle" });
  });

  // ============ SLIDE 9: Scenario 6 - Energy Optimization ============
  let slide9 = pres.addSlide();
  slide9.background = { color: COLORS.lightGray };
  slide9.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: COLORS.primary } });
  slide9.addShape(pres.shapes.OVAL, { x: 0.4, y: 0.15, w: 0.7, h: 0.7, fill: { color: COLORS.accent } });
  slide9.addText("06", { x: 0.4, y: 0.25, w: 0.7, h: 0.5, fontSize: 18, fontFace: "Arial", color: COLORS.white, bold: true, align: "center", valign: "middle" });
  slide9.addText("情境六：能源消耗優化調度", { x: 1.2, y: 0.25, w: 8, h: 0.55, fontSize: 24, fontFace: "Arial Black", color: COLORS.white, bold: true, margin: 0 });

  slide9.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.15, w: 9, h: 0.7, fill: { color: COLORS.success, transparency: 10 }, line: { color: COLORS.success, width: 1 } });
  slide9.addText("🎯 目標：在保持產能的前提下，降低集團整體能耗 15%", { x: 0.7, y: 1.3, w: 8.6, h: 0.4, fontSize: 14, fontFace: "Arial", color: COLORS.success, bold: true, margin: 0 });

  slide9.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.0, w: 4.5, h: 2.0, fill: { color: COLORS.cardBg }, shadow: makeCardShadow() });
  slide9.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.0, w: 0.08, h: 2.0, fill: { color: COLORS.warning } });
  slide9.addText("龍蝦數據分析發現", { x: 0.75, y: 2.15, w: 4.0, h: 0.35, fontSize: 12, fontFace: "Arial", color: COLORS.warning, bold: true, margin: 0 });
  const findings = ["▸ 離峰時段能耗比峰值高 22%（設備空轉）", "▸ 老舊設備能耗效率低於新設備 40%", "▸ 暖通空調節能潛力最大（佔30%能耗）", "▸ 生產排程可平滑化空間大"];
  findings.forEach((f, i) => { slide9.addText(f, { x: 0.75, y: 2.55 + i * 0.35, w: 4.0, h: 0.32, fontSize: 10, fontFace: "Arial", color: COLORS.textDark, margin: 0 }); });

  slide9.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 2.0, w: 4.3, h: 2.0, fill: { color: COLORS.cardBg }, shadow: makeCardShadow() });
  slide9.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 2.0, w: 0.08, h: 2.0, fill: { color: COLORS.holmesGreen } });
  slide9.addText("分段優化策略", { x: 5.45, y: 2.15, w: 3.8, h: 0.35, fontSize: 12, fontFace: "Arial", color: COLORS.holmesGreen, bold: true, margin: 0 });
  const phases = [{ phase: "第一月", action: "排程優化", saving: "節能 5%" }, { phase: "第三月", action: "暖通優化上線", saving: "節能 8%" }, { phase: "第六月", action: "設備更換完成", saving: "節能 12%" }];
  phases.forEach((p, i) => {
    slide9.addText(p.phase, { x: 5.45, y: 2.55 + i * 0.45, w: 1.0, h: 0.35, fontSize: 9, fontFace: "Arial", color: COLORS.accent, bold: true, margin: 0 });
    slide9.addText(p.action, { x: 6.5, y: 2.55 + i * 0.45, w: 1.8, h: 0.35, fontSize: 9, fontFace: "Arial", color: COLORS.textDark, margin: 0 });
    slide9.addText(p.saving, { x: 8.3, y: 2.55 + i * 0.45, w: 1.0, h: 0.35, fontSize: 9, fontFace: "Arial", color: COLORS.success, bold: true, margin: 0 });
  });

  slide9.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.15, w: 9, h: 1.4, fill: { color: COLORS.cardBg }, shadow: makeCardShadow() });
  slide9.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.15, w: 0.08, h: 1.4, fill: { color: COLORS.holmesBlue } });
  slide9.addText("持續監控與回饋", { x: 0.75, y: 4.3, w: 8.5, h: 0.35, fontSize: 12, fontFace: "Arial", color: COLORS.holmesBlue, bold: true, margin: 0 });
  const monitors = [{ icon: icons.chart, text: "每月追蹤節能效果" }, { icon: icons.cogs, text: "自動調整策略應對季節性變化" }, { icon: icons.file, text: "生成 ESG 報告供外部披露" }, { icon: icons.bell, text: "異常超標即時告警" }];
  monitors.forEach((m, i) => {
    slide9.addImage({ data: m.icon, x: 0.85 + i * 2.2, y: 4.75, w: 0.35, h: 0.35 });
    slide9.addText(m.text, { x: 1.25 + i * 2.2, y: 4.8, w: 1.9, h: 0.6, fontSize: 9, fontFace: "Arial", color: COLORS.textDark, margin: 0 });
  });

  // ============ SLIDE 10: Scenario 7 - War Room ============
  let slide10 = pres.addSlide();
  slide10.background = { color: COLORS.lightGray };
  slide10.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: COLORS.primary } });
  slide10.addShape(pres.shapes.OVAL, { x: 0.4, y: 0.15, w: 0.7, h: 0.7, fill: { color: COLORS.accent } });
  slide10.addText("07", { x: 0.4, y: 0.25, w: 0.7, h: 0.5, fontSize: 18, fontFace: "Arial", color: COLORS.white, bold: true, align: "center", valign: "middle" });
  slide10.addText("情境七：主動召集即時戰情室", { x: 1.2, y: 0.25, w: 8, h: 0.55, fontSize: 24, fontFace: "Arial Black", color: COLORS.white, bold: true, margin: 0 });

  slide10.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.15, w: 9, h: 0.8, fill: { color: COLORS.danger, transparency: 10 }, line: { color: COLORS.danger, width: 1 } });
  slide10.addText("觸發情境：龍蝦巡查發現某線體良品率單日下降 5%，且有持續惡化趨勢", { x: 0.7, y: 1.35, w: 8.6, h: 0.4, fontSize: 13, fontFace: "Arial", color: COLORS.danger, bold: true, margin: 0 });

  slide10.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.1, w: 4.5, h: 3.45, fill: { color: COLORS.cardBg }, shadow: makeCardShadow() });
  slide10.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.1, w: 0.08, h: 3.45, fill: { color: COLORS.lobster } });
  slide10.addText("龍蝦主動行為", { x: 0.75, y: 2.25, w: 4.0, h: 0.35, fontSize: 12, fontFace: "Arial", color: COLORS.lobster, bold: true, margin: 0 });
  const wActions = ["1. 串聯 Holmes(Q) 良品變異分析", "2. 調用不良品歸類，識別不良模式", "3. 查詢歷史相似案例（知識圖譜）", "4. 分析影響範圍與客戶交付風險", "5. 評估是否需要緊急介入", "6. 判斷：達到戰情室召集標準", "7. 自動創建即時戰情室會議", "8. 通知相關人員（LINE/Email）", "9. 預先準備相關數據報告"];
  wActions.forEach((a, i) => { slide10.addText(a, { x: 0.75, y: 2.65 + i * 0.32, w: 4.0, h: 0.3, fontSize: 9, fontFace: "Arial", color: COLORS.textDark, margin: 0 }); });

  slide10.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 2.1, w: 4.3, h: 3.45, fill: { color: COLORS.cardBg }, shadow: makeCardShadow() });
  slide10.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 2.1, w: 4.3, h: 0.6, fill: { color: COLORS.holmesPurple } });
  slide10.addImage({ data: icons.comments, x: 5.4, y: 2.2, w: 0.4, h: 0.4 });
  slide10.addText("AI 即時戰情室", { x: 5.9, y: 2.25, w: 3.4, h: 0.4, fontSize: 14, fontFace: "Arial", color: COLORS.white, bold: true, margin: 0 });
  slide10.addText("會議主題", { x: 5.4, y: 2.85, w: 3.9, h: 0.25, fontSize: 9, fontFace: "Arial", color: COLORS.textMuted, margin: 0 });
  slide10.addText("【緊急】工廠A 組裝段良率異常檢討", { x: 5.4, y: 3.1, w: 3.9, h: 0.3, fontSize: 10, fontFace: "Arial", color: COLORS.textDark, bold: true, margin: 0 });
  slide10.addText("參與人員", { x: 5.4, y: 3.5, w: 3.9, h: 0.25, fontSize: 9, fontFace: "Arial", color: COLORS.textMuted, margin: 0 });
  slide10.addText("生產主管、品管主管、技術主管、採購", { x: 5.4, y: 3.75, w: 3.9, h: 0.3, fontSize: 10, fontFace: "Arial", color: COLORS.textDark, margin: 0 });
  slide10.addText("龍蝦預備資料", { x: 5.4, y: 4.15, w: 3.9, h: 0.25, fontSize: 9, fontFace: "Arial", color: COLORS.textMuted, margin: 0 });
  const prepData = ["• 過去72小時良率趨勢圖", "• 不良品分佈熱點圖", "• 疑似根因：原材料供應商變更", "• 建議決策選項"];
  prepData.forEach((p, i) => { slide10.addText(p, { x: 5.4, y: 4.4 + i * 0.28, w: 3.9, h: 0.26, fontSize: 9, fontFace: "Arial", color: COLORS.textDark, margin: 0 }); });

  // ============ SLIDE 11: UI - Main Dashboard ============
  let slide11 = pres.addSlide();
  slide11.background = { color: COLORS.lightGray };
  slide11.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: COLORS.primary } });
  slide11.addText("UI 設計：全知龍蝦主控台", { x: 0.5, y: 0.2, w: 9, h: 0.5, fontSize: 24, fontFace: "Arial Black", color: COLORS.white, bold: true, margin: 0 });

  slide11.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.0, w: 9.4, h: 4.55, fill: { color: COLORS.cardBg }, shadow: makeCardShadow() });
  slide11.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.0, w: 9.4, h: 0.4, fill: { color: "E2E8F0" } });
  const dots = ["EF4444", "F59E0B", "10B981"];
  dots.forEach((c, i) => { slide11.addShape(pres.shapes.OVAL, { x: 0.5 + i * 0.25, y: 1.12, w: 0.16, h: 0.16, fill: { color: c } }); });
  slide11.addShape(pres.shapes.RECTANGLE, { x: 1.1, y: 1.08, w: 7.5, h: 0.24, fill: { color: COLORS.white } });
  slide11.addText("lobster.holmes.ai/dashboard", { x: 1.1, y: 1.05, w: 7.5, h: 0.3, fontSize: 8, fontFace: "Arial", color: COLORS.textMuted, align: "center" });

  slide11.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.4, w: 1.8, h: 4.15, fill: { color: COLORS.secondary } });
  const sidebarItems = [{ icon: icons.desktop, label: "總覽", active: true }, { icon: icons.search, label: "查詢" }, { icon: icons.bell, label: "預警" }, { icon: icons.chart, label: "分析" }, { icon: icons.comments, label: "會議" }, { icon: icons.file, label: "報告" }, { icon: icons.cogs, label: "設定" }];
  sidebarItems.forEach((s, i) => {
    slide11.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.4 + i * 0.55, w: 1.8, h: 0.5, fill: { color: s.active ? COLORS.lobster : "transparent" } });
    slide11.addImage({ data: s.icon, x: 0.5, y: 1.5 + i * 0.55, w: 0.3, h: 0.3 });
    slide11.addText(s.label, { x: 0.9, y: 1.5 + i * 0.55, w: 1.0, h: 0.3, fontSize: 11, fontFace: "Arial", color: COLORS.white, margin: 0 });
  });

  slide11.addShape(pres.shapes.RECTANGLE, { x: 2.1, y: 1.4, w: 7.6, h: 0.5, fill: { color: COLORS.lightGray } });
  slide11.addText("工廠健康總覽", { x: 2.3, y: 1.5, w: 3.0, h: 0.3, fontSize: 16, fontFace: "Arial", color: COLORS.textDark, bold: true, margin: 0 });
  slide11.addShape(pres.shapes.OVAL, { x: 8.8, y: 1.45, w: 0.4, h: 0.4, fill: { color: COLORS.danger } });
  slide11.addText("3", { x: 8.8, y: 1.5, w: 0.4, h: 0.3, fontSize: 12, fontFace: "Arial", color: COLORS.white, bold: true, align: "center" });

  const dashStats = [{ label: "平均良率", value: "98.2%", trend: "+0.3%", up: true }, { label: "平均OEE", value: "85.6%", trend: "-1.2%", up: false }, { label: "設備異常", value: "2件", trend: "-1", up: true }, { label: "待處理預警", value: "5件", trend: "+2", up: false }];
  dashStats.forEach((ds, i) => {
    const x = 2.2 + i * 1.9;
    slide11.addShape(pres.shapes.RECTANGLE, { x, y: 2.0, w: 1.75, h: 0.85, fill: { color: COLORS.cardBg }, shadow: makeSmallShadow() });
    slide11.addText(ds.value, { x, y: 2.1, w: 1.75, h: 0.4, fontSize: 18, fontFace: "Arial", color: COLORS.textDark, bold: true, align: "center" });
    slide11.addText(ds.label, { x, y: 2.5, w: 1.75, h: 0.25, fontSize: 8, fontFace: "Arial", color: COLORS.textMuted, align: "center" });
  });

  slide11.addShape(pres.shapes.RECTANGLE, { x: 2.2, y: 3.0, w: 4.5, h: 2.3, fill: { color: COLORS.cardBg }, shadow: makeSmallShadow() });
  slide11.addText("各工廠良率趨勢", { x: 2.4, y: 3.1, w: 4.0, h: 0.3, fontSize: 11, fontFace: "Arial", color: COLORS.textDark, bold: true, margin: 0 });
  const chartPoints = [3.0, 3.3, 3.5, 3.2, 3.8, 4.0, 4.5, 4.2];
  chartPoints.forEach((y, i) => {
    slide11.addShape(pres.shapes.OVAL, { x: 2.6 + i * 0.5, y: y + 0.8, w: 0.15, h: 0.15, fill: { color: i === 7 ? COLORS.danger : COLORS.holmesBlue } });
    if (i > 0) slide11.addShape(pres.shapes.LINE, { x: 2.6 + (i-1) * 0.5 + 0.15, y: chartPoints[i-1] + 0.8 + 0.075, w: 0.35, h: (y - chartPoints[i-1]), line: { color: COLORS.holmesBlue, width: 2 } });
  });
  slide11.addText("工廠A", { x: 2.5, y: 5.05, w: 0.8, h: 0.2, fontSize: 7, color: COLORS.textMuted });

  slide11.addShape(pres.shapes.RECTANGLE, { x: 6.9, y: 3.0, w: 2.7, h: 2.3, fill: { color: COLORS.cardBg }, shadow: makeSmallShadow() });
  slide11.addText("即時預警", { x: 7.1, y: 3.1, w: 2.3, h: 0.3, fontSize: 11, fontFace: "Arial", color: COLORS.textDark, bold: true, margin: 0 });
  const alerts = [{ level: "high", text: "工廠B 良率低於預期" }, { level: "med", text: "設備A-123 需維護" }, { level: "low", text: "訂單#1234 可能延遲" }];
  alerts.forEach((a, i) => {
    const col = a.level === "high" ? COLORS.danger : a.level === "med" ? COLORS.warning : COLORS.info;
    slide11.addShape(pres.shapes.RECTANGLE, { x: 7.1, y: 3.45 + i * 0.6, w: 0.08, h: 0.45, fill: { color: col } });
    slide11.addText(a.text, { x: 7.3, y: 3.5 + i * 0.6, w: 2.2, h: 0.4, fontSize: 9, fontFace: "Arial", color: COLORS.textDark, margin: 0 });
  });

  // ============ SLIDE 12: UI - Chat Interface ============
  let slide12 = pres.addSlide();
  slide12.background = { color: COLORS.lightGray };
  slide12.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: COLORS.primary } });
  slide12.addText("UI 設計：對話式查詢介面", { x: 0.5, y: 0.2, w: 9, h: 0.5, fontSize: 24, fontFace: "Arial Black", color: COLORS.white, bold: true, margin: 0 });

  slide12.addShape(pres.shapes.RECTANGLE, { x: 1.5, y: 1.1, w: 7, h: 4.4, fill: { color: COLORS.cardBg }, shadow: makeCardShadow() });
  slide12.addShape(pres.shapes.RECTANGLE, { x: 1.5, y: 1.1, w: 7, h: 0.6, fill: { color: COLORS.lobster } });
  slide12.addShape(pres.shapes.OVAL, { x: 1.7, y: 1.2, w: 0.4, h: 0.4, fill: { color: COLORS.accent } });
  slide12.addText("全知龍蝦", { x: 2.2, y: 1.25, w: 3.0, h: 0.3, fontSize: 14, fontFace: "Arial", color: COLORS.white, bold: true, margin: 0 });
  slide12.addText("線上", { x: 7.5, y: 1.25, w: 0.8, h: 0.3, fontSize: 10, fontFace: "Arial", color: COLORS.white, margin: 0 });

  const chatMessages = [
    { from: "user", text: "為什麼這個月工廠A的良率掉了2%？" },
    { from: "bot", text: "我來幫您分析...\n\n根據 Holmes(Q) 良品變異分析：\n▸ 工廠A 組裝段良率下降 2.3%\n▸ 問題始於 4 天前\n▸ 主要不良類型：尺寸偏差\n\n正在串接其他產品數據...", thinking: true },
    { from: "bot", text: "串接 Holmes(Q) 不良品歸類 分析後發現：\n\n▸ 不良品集中在批次 #2024-0415\n▸ 該批次使用新供應商原料\n▸ 知識圖譜：該供應商材料膨脹係數較高\n\n📋 建議：恢復原供應商或調整加工參數\n\n需要我幫您生成異常報告並召集戰情室嗎？" },
  ];

  let chatY = 1.85;
  chatMessages.forEach((m) => {
    const isUser = m.from === "user";
    const msgW = 4.5;
    const msgX = isUser ? 7.5 - 0.3 - msgW : 1.8;
    if (m.thinking) {
      slide12.addShape(pres.shapes.RECTANGLE, { x: msgX, y: chatY, w: msgW + 0.5, h: 0.5, fill: { color: COLORS.lightGray } });
      slide12.addText("思考中... " + m.text, { x: msgX + 0.15, y: chatY + 0.1, w: msgW, h: 0.35, fontSize: 9, fontFace: "Arial", color: COLORS.textMuted, italic: true });
      chatY += 0.6;
    } else {
      slide12.addShape(pres.shapes.RECTANGLE, { x: msgX, y: chatY, w: msgW, h: m.text.split('\n').length * 0.28 + 0.35, fill: { color: isUser ? COLORS.holmesBlue : COLORS.lightGray } });
      slide12.addText(m.text, { x: msgX + 0.15, y: chatY + 0.1, w: msgW - 0.3, h: m.text.split('\n').length * 0.28 + 0.2, fontSize: 9, fontFace: "Arial", color: isUser ? COLORS.white : COLORS.textDark });
      chatY += m.text.split('\n').length * 0.28 + 0.5;
    }
  });

  slide12.addText("快捷指令：", { x: 1.8, y: 4.5, w: 1.2, h: 0.3, fontSize: 9, fontFace: "Arial", color: COLORS.textMuted });
  const quickActions = ["生成報告", "召集戰情室", "查看趨勢"];
  quickActions.forEach((qa, i) => {
    slide12.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 3.0 + i * 1.3, y: 4.5, w: 1.1, h: 0.35, fill: { color: COLORS.holmesBlue }, rectRadius: 0.05 });
    slide12.addText(qa, { x: 3.0 + i * 1.3, y: 4.5, w: 1.1, h: 0.35, fontSize: 8, fontFace: "Arial", color: COLORS.white, align: "center", valign: "middle" });
  });

  slide12.addShape(pres.shapes.RECTANGLE, { x: 1.5, y: 5.0, w: 7, h: 0.5, fill: { color: COLORS.lightGray } });
  slide12.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.7, y: 5.08, w: 5.8, h: 0.34, fill: { color: COLORS.white }, rectRadius: 0.05 });
  slide12.addText("輸入訊息...（Shift+Enter 換行）", { x: 1.9, y: 5.12, w: 5.4, h: 0.26, fontSize: 9, fontFace: "Arial", color: COLORS.textMuted });
  slide12.addShape(pres.shapes.OVAL, { x: 7.6, y: 5.05, w: 0.35, h: 0.35, fill: { color: COLORS.holmesBlue } });
  slide12.addImage({ data: icons.mic, x: 7.68, y: 5.13, w: 0.2, h: 0.2 });

  // ============ SLIDE 13: UI - Action Interface ============
  let slide13 = pres.addSlide();
  slide13.background = { color: COLORS.lightGray };
  slide13.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: COLORS.primary } });
  slide13.addText("UI 設計：行動執行與預警介面", { x: 0.5, y: 0.2, w: 9, h: 0.5, fontSize: 24, fontFace: "Arial Black", color: COLORS.white, bold: true, margin: 0 });

  slide13.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 4.5, h: 4.4, fill: { color: COLORS.cardBg }, shadow: makeCardShadow() });
  slide13.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 4.5, h: 0.6, fill: { color: COLORS.warning } });
  slide13.addText("⚠️  行動建議", { x: 0.7, y: 1.2, w: 4.0, h: 0.4, fontSize: 14, fontFace: "Arial", color: COLORS.white, bold: true, margin: 0 });
  slide13.addText("設備 A-123 軸承異常，建議提前維護", { x: 0.7, y: 1.85, w: 4.0, h: 0.35, fontSize: 12, fontFace: "Arial", color: COLORS.textDark, bold: true, margin: 0 });
  slide13.addText("預期效果", { x: 0.7, y: 2.3, w: 4.0, h: 0.25, fontSize: 9, fontFace: "Arial", color: COLORS.textMuted, margin: 0 });
  slide13.addText("避免 2-3 天非計劃停機，延遲 < 0.5 天", { x: 0.7, y: 2.55, w: 4.0, h: 0.25, fontSize: 10, fontFace: "Arial", color: COLORS.textDark, margin: 0 });
  slide13.addText("影響範圍", { x: 0.7, y: 2.9, w: 4.0, h: 0.25, fontSize: 9, fontFace: "Arial", color: COLORS.textMuted, margin: 0 });
  slide13.addText("工廠A 產能 -15%，持續約 4 小時", { x: 0.7, y: 3.15, w: 4.0, h: 0.25, fontSize: 10, fontFace: "Arial", color: COLORS.textDark, margin: 0 });

  slide13.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.7, y: 3.6, w: 1.8, h: 0.5, fill: { color: COLORS.success }, rectRadius: 0.05 });
  slide13.addText("✓ 核准執行", { x: 0.7, y: 3.6, w: 1.8, h: 0.5, fontSize: 11, fontFace: "Arial", color: COLORS.white, bold: true, align: "center", valign: "middle" });
  slide13.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 2.7, y: 3.6, w: 1.8, h: 0.5, fill: { color: "E2E8F0" }, rectRadius: 0.05 });
  slide13.addText("✗ 稍後處理", { x: 2.7, y: 3.6, w: 1.8, h: 0.5, fontSize: 11, fontFace: "Arial", color: COLORS.textMuted, align: "center", valign: "middle" });
  slide13.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: 4.3, w: 4.0, h: 0.06, fill: { color: COLORS.success } });
  slide13.addText("已核准 → 執行中...", { x: 0.7, y: 4.45, w: 4.0, h: 0.25, fontSize: 9, fontFace: "Arial", color: COLORS.success, margin: 0 });
  const execSteps = [{ text: "調整工廠A排程", done: true }, { text: "通知工廠B、C接收訂單", done: true }, { text: "創建維修工單", done: false }, { text: "發送通知給維修團隊", done: false }];
  execSteps.forEach((es, i) => {
    slide13.addText(es.done ? "✓" : "○", { x: 0.7, y: 4.75 + i * 0.3, w: 0.3, h: 0.25, fontSize: 10, fontFace: "Arial", color: es.done ? COLORS.success : COLORS.textMuted });
    slide13.addText(es.text, { x: 1.0, y: 4.75 + i * 0.3, w: 3.5, h: 0.25, fontSize: 9, fontFace: "Arial", color: es.done ? COLORS.textDark : COLORS.textMuted });
  });

  slide13.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.1, w: 4.3, h: 2.2, fill: { color: COLORS.cardBg }, shadow: makeCardShadow() });
  slide13.addText("風險等級判定", { x: 5.4, y: 1.25, w: 3.9, h: 0.35, fontSize: 12, fontFace: "Arial", color: COLORS.textDark, bold: true, margin: 0 });
  const riskItems = [{ label: "設備類型", value: "關鍵設備", risk: "high" }, { label: "影響產能", value: "> 10%", risk: "high" }, { label: "跨廠影響", value: "是", risk: "high" }, { label: "預測置信度", value: "92%", risk: "low" }];
  riskItems.forEach((ri, i) => {
    const col = ri.risk === "high" ? COLORS.danger : COLORS.success;
    slide13.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: 1.7 + i * 0.38, w: 0.08, h: 0.3, fill: { color: col } });
    slide13.addText(ri.label, { x: 5.6, y: 1.7 + i * 0.38, w: 1.8, h: 0.3, fontSize: 9, fontFace: "Arial", color: COLORS.textMuted, margin: 0 });
    slide13.addText(ri.value, { x: 7.4, y: 1.7 + i * 0.38, w: 1.8, h: 0.3, fontSize: 9, fontFace: "Arial", color: col, bold: true, margin: 0 });
  });

  slide13.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 3.45, w: 4.3, h: 2.05, fill: { color: COLORS.cardBg }, shadow: makeCardShadow() });
  slide13.addText("操作日誌", { x: 5.4, y: 3.6, w: 3.9, h: 0.35, fontSize: 12, fontFace: "Arial", color: COLORS.textDark, bold: true, margin: 0 });
  const logs = [{ time: "14:32", action: "龍蝦發現異常" }, { time: "14:33", action: "生成行動建議" }, { time: "14:35", action: "管理層核准" }, { time: "14:35", action: "開始執行" }];
  logs.forEach((l, i) => {
    slide13.addText(l.time, { x: 5.4, y: 4.0 + i * 0.35, w: 0.6, h: 0.3, fontSize: 8, fontFace: "Arial", color: COLORS.textMuted, margin: 0 });
    slide13.addText(l.action, { x: 6.1, y: 4.0 + i * 0.35, w: 3.2, h: 0.3, fontSize: 9, fontFace: "Arial", color: COLORS.textDark, margin: 0 });
  });

  // ============ SLIDE 14: UI - Cross-Product Analysis ============
  let slide14 = pres.addSlide();
  slide14.background = { color: COLORS.lightGray };
  slide14.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: COLORS.primary } });
  slide14.addText("UI 設計：跨產品關聯分析視圖", { x: 0.5, y: 0.2, w: 9, h: 0.5, fontSize: 24, fontFace: "Arial Black", color: COLORS.white, bold: true, margin: 0 });

  slide14.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 5.5, h: 4.4, fill: { color: COLORS.cardBg }, shadow: makeCardShadow() });
  slide14.addText("產品關聯圖譜", { x: 0.7, y: 1.25, w: 5.0, h: 0.35, fontSize: 12, fontFace: "Arial", color: COLORS.textDark, bold: true, margin: 0 });

  slide14.addShape(pres.shapes.OVAL, { x: 2.6, y: 2.7, w: 1.2, h: 1.2, fill: { color: COLORS.lobster } });
  slide14.addText("全知\n龍蝦", { x: 2.6, y: 2.95, w: 1.2, h: 0.7, fontSize: 11, fontFace: "Arial", color: COLORS.white, bold: true, align: "center", valign: "middle" });

  const productNodes = [
    { name: "Holmes(P)", sub: "產能", x: 0.8, y: 1.6, color: COLORS.holmesBlue },
    { name: "Holmes(Q)", sub: "良品", x: 4.2, y: 1.6, color: COLORS.holmesGreen },
    { name: "Holmes(Q)", sub: "不良品", x: 0.8, y: 3.8, color: COLORS.holmesGreen },
    { name: "AI會議室", sub: "戰情室", x: 4.2, y: 3.8, color: COLORS.holmesPurple },
  ];
  productNodes.forEach((n) => {
    slide14.addShape(pres.shapes.LINE, { x: 2.6 + 0.6, y: 2.7 + 0.6, w: n.x + 0.5 - (2.6 + 0.6), h: n.y + 0.3 - (2.7 + 0.6), line: { color: COLORS.accent, width: 2, dashType: "dash" } });
    slide14.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: n.x, y: n.y, w: 1.5, h: 0.7, fill: { color: n.color }, rectRadius: 0.1 });
    slide14.addText(n.name, { x: n.x, y: n.y + 0.1, w: 1.5, h: 0.3, fontSize: 10, fontFace: "Arial", color: COLORS.white, bold: true, align: "center" });
    slide14.addText(n.sub, { x: n.x, y: n.y + 0.35, w: 1.5, h: 0.25, fontSize: 8, fontFace: "Arial", color: COLORS.white, align: "center" });
  });

  slide14.addShape(pres.shapes.RECTANGLE, { x: 6.2, y: 1.1, w: 3.3, h: 4.4, fill: { color: COLORS.cardBg }, shadow: makeCardShadow() });
  slide14.addShape(pres.shapes.RECTANGLE, { x: 6.2, y: 1.1, w: 3.3, h: 0.5, fill: { color: COLORS.secondary } });
  slide14.addText("異常關聯事件鏈", { x: 6.4, y: 1.18, w: 3.0, h: 0.35, fontSize: 12, fontFace: "Arial", color: COLORS.white, bold: true, margin: 0 });

  const events = [
    { from: "P-產能", to: "Q-良品", reason: "產能↓ → 良率↓", strength: 85 },
    { from: "Q-良品", to: "Q-不良品", reason: "尺寸偏差增加", strength: 92 },
    { from: "Q-良品", to: "AI-戰情", reason: "建議召集會議", strength: 78 },
  ];
  events.forEach((e, i) => {
    const y = 1.75 + i * 1.1;
    slide14.addShape(pres.shapes.RECTANGLE, { x: 6.4, y, w: 2.9, h: 0.95, fill: { color: COLORS.lightGray } });
    slide14.addShape(pres.shapes.RECTANGLE, { x: 6.4, y, w: 0.06, h: 0.95, fill: { color: COLORS.accent } });
    slide14.addText(e.from + " → " + e.to, { x: 6.55, y: y + 0.08, w: 2.6, h: 0.3, fontSize: 10, fontFace: "Arial", color: COLORS.textDark, bold: true, margin: 0 });
    slide14.addText(e.reason, { x: 6.55, y: y + 0.38, w: 2.6, h: 0.25, fontSize: 9, fontFace: "Arial", color: COLORS.textMuted, margin: 0 });
    slide14.addText("置信度: " + e.strength + "%", { x: 6.55, y: y + 0.65, w: 2.6, h: 0.22, fontSize: 8, fontFace: "Arial", color: COLORS.holmesGreen, margin: 0 });
  });

  // ============ SLIDE 15: Summary ============
  let slide15 = pres.addSlide();
  slide15.background = { color: COLORS.primary };
  slide15.addShape(pres.shapes.OVAL, { x: 4.0, y: 0.5, w: 2.0, h: 2.0, fill: { color: COLORS.lobster }, shadow: { type: "outer", color: "000000", blur: 20, offset: 6, angle: 135, opacity: 0.35 } });
  slide15.addImage({ data: icons.brain, x: 4.3, y: 0.8, w: 1.4, h: 1.4 });
  slide15.addText("全知龍蝦的定位", { x: 0.5, y: 2.7, w: 9, h: 0.6, fontSize: 28, fontFace: "Arial Black", color: COLORS.white, bold: true, align: "center" });
  slide15.addText("Holmes 的代言人", { x: 0.5, y: 3.3, w: 9, h: 0.5, fontSize: 20, fontFace: "Arial", color: COLORS.accent, bold: true, align: "center" });

  const summaryPoints = ["深度理解 Holmes 的每一個產品", "串連產品間的關係，發現關聯異常", "主動預警，主動行動", "對話式介面，統一對外窗口", "不只是呈現數據，而是幫助人類解決問題"];
  summaryPoints.forEach((sp, i) => { slide15.addText("✓  " + sp, { x: 2.0, y: 3.95 + i * 0.32, w: 6, h: 0.3, fontSize: 12, fontFace: "Arial", color: COLORS.white, margin: 0 }); });
  slide15.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.1, w: 10, h: 0.525, fill: { color: COLORS.lobster } });
  slide15.addText("AI大數據部門  |  智能製造規劃", { x: 0.5, y: 5.15, w: 9, h: 0.45, fontSize: 12, fontFace: "Arial", color: COLORS.white, align: "center", valign: "middle" });

  await pres.writeFile({ fileName: "/Users/sam/Documents/MyProject/雲端全知龍蝦_完整規劃.pptx" });
  console.log("Presentation created successfully!");
}

createPresentation().catch(console.error);
