const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const { FaBrain, FaCogs, FaChartLine, FaShieldAlt, FaRegLightbulb, FaServer, FaMicrochip, FaLink, FaSearch, FaBell, FaFileAlt, FaChartBar, FaCheckCircle, FaExclamationTriangle, FaSyncAlt, FaArrowRight, FaComments, FaRocket, FaDatabase } = require("react-icons/fa");

// Color Palette - Deep Navy + Gold/Orange for premium feel
const COLORS = {
  primary: "0A1628",       // Deep navy
  secondary: "1A365D",     // Dark blue
  accent: "F6AD55",         // Warm orange/gold
  accentDark: "DD6B20",     // Dark orange
  white: "FFFFFF",
  lightGray: "F7FAFC",
  textDark: "1A202C",
  textMuted: "718096",
  cardBg: "FFFFFF",
  lobster: "E53E3E",       // Lobster red
  lobsterDark: "C53030",
  holmesBlue: "3182CE",    // Holmes brand blue
  holmesGreen: "38A169",    // Holmes(Q) green
  holmesPurple: "805AD5",   // Holmes(Knowledge) purple
};

// Helper function to render icon to base64 PNG
function renderIconSvg(IconComponent, color = "#000000", size = 256) {
  return ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color, size: String(size) })
  );
}

async function iconToBase64Png(IconComponent, color, size = 256) {
  const svg = renderIconSvg(IconComponent, color, size);
  const pngBuffer = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + pngBuffer.toString("base64");
}

const makeCardShadow = () => ({ type: "outer", color: "000000", blur: 8, offset: 3, angle: 135, opacity: 0.12 });

async function createPresentation() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = "雲端全知龍蝦 - Holmes 智能代理人";
  pres.author = "AI大數據部門";

  // Pre-render icons
  const icons = {
    brain: await iconToBase64Png(FaBrain, "#FFFFFF", 256),
    brainOrange: await iconToBase64Png(FaBrain, "#F6AD55", 256),
    cogs: await iconToBase64Png(FaCogs, "#FFFFFF", 256),
    chart: await iconToBase64Png(FaChartLine, "#FFFFFF", 256),
    shield: await iconToBase64Png(FaShieldAlt, "#FFFFFF", 256),
    lightbulb: await iconToBase64Png(FaRegLightbulb, "#F6AD55", 256),
    link: await iconToBase64Png(FaLink, "#F6AD55", 256),
    search: await iconToBase64Png(FaSearch, "#FFFFFF", 256),
    bell: await iconToBase64Png(FaBell, "#FFFFFF", 256),
    file: await iconToBase64Png(FaFileAlt, "#FFFFFF", 256),
    bar: await iconToBase64Png(FaChartBar, "#FFFFFF", 256),
    check: await iconToBase64Png(FaCheckCircle, "#38A169", 256),
    warning: await iconToBase64Png(FaExclamationTriangle, "#DD6B20", 256),
    sync: await iconToBase64Png(FaSyncAlt, "#F6AD55", 256),
    arrow: await iconToBase64Png(FaArrowRight, "#F6AD55", 256),
    comments: await iconToBase64Png(FaComments, "#FFFFFF", 256),
    rocket: await iconToBase64Png(FaRocket, "#FFFFFF", 256),
    database: await iconToBase64Png(FaDatabase, "#FFFFFF", 256),
  };

  // ============ SLIDE 1: Title - Holmes Agent ============
  let slide1 = pres.addSlide();
  slide1.background = { color: COLORS.primary };

  // Large lobster icon with brain - representing Holmes knowledge
  slide1.addShape(pres.shapes.OVAL, {
    x: 3.8, y: 0.6, w: 2.4, h: 2.4,
    fill: { color: COLORS.lobster },
    shadow: { type: "outer", color: "000000", blur: 20, offset: 6, angle: 135, opacity: 0.35 }
  });
  slide1.addImage({ data: icons.brain, x: 4.2, y: 1.0, w: 1.6, h: 1.6 });

  // Title
  slide1.addText("雲端全知龍蝦", {
    x: 0.5, y: 3.2, w: 9, h: 0.9,
    fontSize: 48, fontFace: "Arial Black", color: COLORS.white, bold: true, align: "center"
  });

  // Subtitle - Core positioning
  slide1.addText("Holmes 全產品線 智能代理人", {
    x: 0.5, y: 4.0, w: 9, h: 0.6,
    fontSize: 24, fontFace: "Arial", color: COLORS.accent, align: "center"
  });

  // Tagline
  slide1.addText("懂 Holmes 的每一個產品，串聯產品間的關係，成為統一的對外窗口", {
    x: 0.5, y: 4.7, w: 9, h: 0.5,
    fontSize: 14, fontFace: "Arial", color: COLORS.white, align: "center"
  });

  // ============ SLIDE 2: Current Problem - Products are Siloed ============
  let slide2 = pres.addSlide();
  slide2.background = { color: COLORS.lightGray };

  // Title bar
  slide2.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.1,
    fill: { color: COLORS.primary }
  });
  slide2.addText("現況問題：各產品線獨立運作，缺乏關聯", {
    x: 0.5, y: 0.3, w: 9, h: 0.6,
    fontSize: 28, fontFace: "Arial Black", color: COLORS.white, bold: true, margin: 0
  });

  // Left side - Problem description
  slide2.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.35, w: 4.3, h: 3.9,
    fill: { color: COLORS.cardBg },
    shadow: makeCardShadow()
  });
  slide2.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.35, w: 0.08, h: 3.9,
    fill: { color: COLORS.lobster }
  });
  slide2.addImage({ data: icons.warning, x: 0.75, y: 1.55, w: 0.5, h: 0.5 });
  slide2.addText("現況瓶頸", {
    x: 1.35, y: 1.55, w: 3.2, h: 0.5,
    fontSize: 18, fontFace: "Arial", color: COLORS.lobster, bold: true, margin: 0
  });

  const problems = [
    "各 Holmes 產品獨立分析，by線體維度",
    "產品之間沒有關聯性分析",
    "異常发生時無法快速串連根因",
    "管理層需要分別詢問多個系統",
    "跨產品趨勢無法整合判斷"
  ];
  problems.forEach((prob, i) => {
    slide2.addText([
      { text: "✗ ", options: { color: COLORS.lobster, bold: true } },
      { text: prob, options: { color: COLORS.textDark } }
    ], {
      x: 0.75, y: 2.2 + i * 0.6, w: 3.8, h: 0.5,
      fontSize: 12, fontFace: "Arial", margin: 0
    });
  });

  // Right side - What we need
  slide2.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.35, w: 4.3, h: 3.9,
    fill: { color: COLORS.cardBg },
    shadow: makeCardShadow()
  });
  slide2.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.35, w: 0.08, h: 3.9,
    fill: { color: COLORS.holmesBlue }
  });
  slide2.addImage({ data: icons.rocket, x: 5.45, y: 1.55, w: 0.5, h: 0.5 });
  slide2.addText("需要什麼", {
    x: 6.05, y: 1.55, w: 3.2, h: 0.5,
    fontSize: 18, fontFace: "Arial", color: COLORS.holmesBlue, bold: true, margin: 0
  });

  const needs = [
    "一個代理人，統一看懂所有產品",
    "能串連產品間的因果關係",
    "主動發現跨產品異常關聯",
    "單一窗口回答所有 Holmes 問題",
    "提供整合性的診斷與建議"
  ];
  needs.forEach((need, i) => {
    slide2.addText([
      { text: "✓ ", options: { color: COLORS.holmesBlue, bold: true } },
      { text: need, options: { color: COLORS.textDark } }
    ], {
      x: 5.45, y: 2.2 + i * 0.6, w: 3.8, h: 0.5,
      fontSize: 12, fontFace: "Arial", margin: 0
    });
  });

  // ============ SLIDE 3: Holmes Product Family - Deep Knowledge ============
  let slide3 = pres.addSlide();
  slide3.background = { color: COLORS.lightGray };

  // Title bar
  slide3.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.1,
    fill: { color: COLORS.primary }
  });
  slide3.addText("全知龍蝦必須深度理解 Holmes 全產品線", {
    x: 0.5, y: 0.3, w: 9, h: 0.6,
    fontSize: 26, fontFace: "Arial Black", color: COLORS.white, bold: true, margin: 0
  });

  // Product categories with their details
  const productCategories = [
    {
      name: "智能生產 Holmes(P)",
      color: COLORS.holmesBlue,
      icon: icons.cogs,
      products: [
        "產能優化工具",
        "瓶頸根因分析",
        "OEE 追蹤"
      ],
      detail: "by 線體維度"
    },
    {
      name: "智能品質 Holmes(Q)",
      color: COLORS.holmesGreen,
      icon: icons.chart,
      products: [
        "良品變異分析 (CTQ站)",
        "產品健康指數分析",
        "敘述統計分析 (CPK/圖表)",
        "不良品歸類分析",
        "SoH Burn-in 變異分析"
      ],
      detail: "by 線體維度"
    },
    {
      name: "知識管理 AI會議室",
      color: COLORS.holmesPurple,
      icon: icons.comments,
      products: [
        "常規會議室",
        "即時戰情室",
        "深度調研室"
      ],
      detail: "by 會議維度"
    }
  ];

  const catW = 3.0, catH = 3.8, catStartX = 0.5, catStartY = 1.25, catGapX = 0.2;

  productCategories.forEach((cat, i) => {
    const x = catStartX + i * (catW + catGapX);
    const y = catStartY;

    // Card
    slide3.addShape(pres.shapes.RECTANGLE, {
      x, y, w: catW, h: catH,
      fill: { color: COLORS.cardBg },
      shadow: makeCardShadow()
    });

    // Header bar
    slide3.addShape(pres.shapes.RECTANGLE, {
      x, y, w: catW, h: 0.7,
      fill: { color: cat.color }
    });
    slide3.addImage({ data: cat.icon, x: x + 0.15, y: y + 0.1, w: 0.5, h: 0.5 });
    slide3.addText(cat.name, {
      x: x + 0.7, y: y + 0.1, w: 2.2, h: 0.5,
      fontSize: 11, fontFace: "Arial", color: COLORS.white, bold: true, valign: "middle", margin: 0
    });

    // Products list
    cat.products.forEach((prod, j) => {
      slide3.addText([
        { text: "▸ ", options: { color: cat.color, bold: true } },
        { text: prod, options: { color: COLORS.textDark } }
      ], {
        x: x + 0.15, y: y + 0.85 + j * 0.42, w: catW - 0.3, h: 0.38,
        fontSize: 10, fontFace: "Arial", margin: 0
      });
    });

    // Dimension tag
    slide3.addShape(pres.shapes.RECTANGLE, {
      x: x + 0.15, y: y + 3.3, w: catW - 0.3, h: 0.35,
      fill: { color: cat.color, transparency: 15 }
    });
    slide3.addText(cat.detail, {
      x: x + 0.15, y: y + 3.3, w: catW - 0.3, h: 0.35,
      fontSize: 10, fontFace: "Arial", color: cat.color, align: "center", valign: "middle"
    });
  });

  // ============ SLIDE 4: Cross-Product Correlation - The Key Value ============
  let slide4 = pres.addSlide();
  slide4.background = { color: COLORS.lightGray };

  // Title bar
  slide4.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.1,
    fill: { color: COLORS.primary }
  });
  slide4.addImage({ data: icons.link, x: 0.5, y: 0.25, w: 0.6, h: 0.6 });
  slide4.addText("核心價值：串連產品間的關係", {
    x: 1.2, y: 0.3, w: 8, h: 0.6,
    fontSize: 28, fontFace: "Arial Black", color: COLORS.white, bold: true, margin: 0
  });

  // Center - The Lobster Agent
  slide4.addShape(pres.shapes.OVAL, {
    x: 4.0, y: 2.3, w: 2.0, h: 2.0,
    fill: { color: COLORS.lobster },
    shadow: { type: "outer", color: "000000", blur: 15, offset: 5, angle: 135, opacity: 0.3 }
  });
  slide4.addImage({ data: icons.brain, x: 4.3, y: 2.6, w: 1.4, h: 1.4 });
  slide4.addText("全知龍蝦", {
    x: 3.5, y: 4.4, w: 3.0, h: 0.4,
    fontSize: 14, fontFace: "Arial", color: COLORS.lobster, bold: true, align: "center"
  });

  // Surrounding products - showing relationships
  const relations = [
    { name: "Holmes(P)\n產能瓶頸", x: 1.2, y: 1.5, color: COLORS.holmesBlue },
    { name: "Holmes(Q)\n良品變異", x: 7.3, y: 1.5, color: COLORS.holmesGreen },
    { name: "Holmes(Q)\n不良品歸類", x: 0.5, y: 3.8, color: COLORS.holmesGreen },
    { name: "AI會議室\n戰情室", x: 8.0, y: 3.8, color: COLORS.holmesPurple },
  ];

  relations.forEach((rel) => {
    // Connection line to center
    const centerX = 5.0, centerY = 3.3;
    const relCenterX = rel.x + 0.9, relCenterY = rel.y + 0.4;
    const angle = Math.atan2(centerY - relCenterY, centerX - relCenterX);
    const lineLength = 1.2;

    slide4.addShape(pres.shapes.LINE, {
      x: relCenterX + Math.cos(angle) * 0.5,
      y: relCenterY + Math.sin(angle) * 0.5,
      w: Math.cos(angle) * lineLength,
      h: Math.sin(angle) * lineLength,
      line: { color: COLORS.accent, width: 2, dashType: "dash" }
    });

    // Product node
    slide4.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: rel.x, y: rel.y, w: 1.8, h: 0.8,
      fill: { color: rel.color },
      rectRadius: 0.1
    });
    slide4.addText(rel.name, {
      x: rel.x, y: rel.y, w: 1.8, h: 0.8,
      fontSize: 9, fontFace: "Arial", color: COLORS.white, bold: true, align: "center", valign: "middle"
    });
  });

  // Bottom - Example of correlation
  slide4.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.85, w: 9, h: 0.7,
    fill: { color: COLORS.accent, transparency: 15 }
  });
  slide4.addText("例：當 Holmes(P) 發現某線體產能下降 → 全知龍蝦自動查詢 Holmes(Q) 良品數據 → 發現良率下降 → 調用 AI會議室 召開戰情室", {
    x: 0.7, y: 4.95, w: 8.6, h: 0.5,
    fontSize: 10, fontFace: "Arial", color: COLORS.accentDark, valign: "middle"
  });

  // ============ SLIDE 5: Agent Capabilities ============
  let slide5 = pres.addSlide();
  slide5.background = { color: COLORS.lightGray };

  // Title bar
  slide5.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.1,
    fill: { color: COLORS.primary }
  });
  slide5.addText("全知龍蝦的代理人能力", {
    x: 0.5, y: 0.3, w: 9, h: 0.6,
    fontSize: 28, fontFace: "Arial Black", color: COLORS.white, bold: true, margin: 0
  });

  // Two columns of capabilities
  const leftCaps = [
    { icon: icons.search, title: "全產品理解", desc: "熟悉 Holmes 每個產品的演算法、數據結構、使用情境，能用自然語言解釋任何產品的輸出" },
    { icon: icons.link, title: "關聯分析", desc: "串連不同產品的異常信號，識別跨產品線的根因傳播鏈" },
    { icon: icons.bell, title: "主動預警", desc: "不只被動響應，而是主動監控，主動發現潛在風險並發起行動" },
  ];

  const rightCaps = [
    { icon: icons.comments, title: "對話式代理", desc: "管理者可以用自然語言提問，龍蝦整合所有產品數據回答" },
    { icon: icons.rocket, title: "協調行動", desc: "能調用 AI會議室 發起會議、觸發 Holmes(Q) 分析、生成報告" },
    { icon: icons.shield, title: "決策建議", desc: "不只是呈現數據，而是給出行動建議和預期效果" },
  ];

  const capW = 4.4, capH = 1.2, capStartY = 1.25, capGapY = 0.15;

  leftCaps.forEach((cap, i) => {
    const y = capStartY + i * (capH + capGapY);
    slide5.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y, w: capW, h: capH,
      fill: { color: COLORS.cardBg },
      shadow: makeCardShadow()
    });
    slide5.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y, w: 0.06, h: capH,
      fill: { color: COLORS.accent }
    });
    slide5.addImage({ data: cap.icon, x: 0.7, y: y + 0.35, w: 0.5, h: 0.5 });
    slide5.addText(cap.title, {
      x: 1.3, y: y + 0.15, w: 3.4, h: 0.35,
      fontSize: 14, fontFace: "Arial", color: COLORS.textDark, bold: true, margin: 0
    });
    slide5.addText(cap.desc, {
      x: 1.3, y: y + 0.5, w: 3.4, h: 0.6,
      fontSize: 9, fontFace: "Arial", color: COLORS.textMuted, margin: 0
    });
  });

  rightCaps.forEach((cap, i) => {
    const y = capStartY + i * (capH + capGapY);
    slide5.addShape(pres.shapes.RECTANGLE, {
      x: 5.1, y, w: capW, h: capH,
      fill: { color: COLORS.cardBg },
      shadow: makeCardShadow()
    });
    slide5.addShape(pres.shapes.RECTANGLE, {
      x: 5.1, y, w: 0.06, h: capH,
      fill: { color: COLORS.accent }
    });
    slide5.addImage({ data: cap.icon, x: 5.3, y: y + 0.35, w: 0.5, h: 0.5 });
    slide5.addText(cap.title, {
      x: 5.9, y: y + 0.15, w: 3.4, h: 0.35,
      fontSize: 14, fontFace: "Arial", color: COLORS.textDark, bold: true, margin: 0
    });
    slide5.addText(cap.desc, {
      x: 5.9, y: y + 0.5, w: 3.4, h: 0.6,
      fontSize: 9, fontFace: "Arial", color: COLORS.textMuted, margin: 0
    });
  });

  // ============ SLIDE 6: Use Case Examples ============
  let slide6 = pres.addSlide();
  slide6.background = { color: COLORS.lightGray };

  // Title bar
  slide6.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.1,
    fill: { color: COLORS.primary }
  });
  slide6.addText("全知龍蝦應用場景", {
    x: 0.5, y: 0.3, w: 9, h: 0.6,
    fontSize: 28, fontFace: "Arial Black", color: COLORS.white, bold: true, margin: 0
  });

  // Use case cards
  const useCases = [
    {
      q: "「為什麼這個月工廠A的良率掉了2%？」",
      a: "全知龍蝦串連 Holmes(Q) 良品數據 → 發現是組裝段問題 → 進一步調用不良品歸類分析 → 識別是新供應商材料問題 → 自動生成報告並建議召回",
      icon: icons.search
    },
    {
      q: "「這個禮拜要不要召開戰情室會議？」",
      a: "全知龍蝦主動巡查各產品線狀態 → 發現 Holmes(P) 顯示某線體產能下降15% → 判斷需要戰情室介入 → 自動創建即時戰情室並通知相關人員",
      icon: icons.bell
    },
    {
      q: "「幫我分析這個新產品的健康度」",
      a: "全知龍蝦調用 Holmes(Q) 健康指數分析 → 結合敘述統計數據 → 查詢 AI會議室 是否有類似產品問題記錄 → 給出綜合健康度報告和風險預警",
      icon: icons.file
    }
  ];

  const ucW = 2.9, ucH = 3.5, ucStartX = 0.5, ucStartY = 1.3, ucGapX = 0.2;

  useCases.forEach((uc, i) => {
    const x = ucStartX + i * (ucW + ucGapX);
    const y = ucStartY;

    // Card
    slide6.addShape(pres.shapes.RECTANGLE, {
      x, y, w: ucW, h: ucH,
      fill: { color: COLORS.cardBg },
      shadow: makeCardShadow()
    });

    // Icon
    slide6.addShape(pres.shapes.OVAL, {
      x: x + 1.05, y: y + 0.15, w: 0.8, h: 0.8,
      fill: { color: COLORS.accent }
    });
    slide6.addImage({ data: uc.icon, x: x + 1.2, y: y + 0.3, w: 0.5, h: 0.5 });

    // Question
    slide6.addText("管理者提問", {
      x: x + 0.15, y: y + 1.05, w: ucW - 0.3, h: 0.3,
      fontSize: 9, fontFace: "Arial", color: COLORS.textMuted, margin: 0
    });
    slide6.addText(uc.q, {
      x: x + 0.15, y: y + 1.3, w: ucW - 0.3, h: 0.7,
      fontSize: 10, fontFace: "Arial", color: COLORS.textDark, bold: true, margin: 0
    });

    // Answer
    slide6.addShape(pres.shapes.RECTANGLE, {
      x: x + 0.1, y: y + 2.1, w: ucW - 0.2, h: 0.06,
      fill: { color: COLORS.accent, transparency: 50 }
    });
    slide6.addText("龍蝦回應", {
      x: x + 0.15, y: y + 2.25, w: ucW - 0.3, h: 0.25,
      fontSize: 9, fontFace: "Arial", color: COLORS.accent, margin: 0
    });
    slide6.addText(uc.a, {
      x: x + 0.15, y: y + 2.5, w: ucW - 0.3, h: 0.9,
      fontSize: 8, fontFace: "Arial", color: COLORS.textMuted, margin: 0
    });
  });

  // Bottom tagline
  slide6.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 5.1, w: 10, h: 0.525,
    fill: { color: COLORS.lobster }
  });
  slide6.addText("Holmes 的代言人 — 任何問題，問龍蝦就對了", {
    x: 0.5, y: 5.15, w: 9, h: 0.45,
    fontSize: 16, fontFace: "Arial", color: COLORS.white, bold: true, align: "center", valign: "middle"
  });

  // Save the presentation
  await pres.writeFile({ fileName: "/Users/sam/Documents/MyProject/雲端全知龍蝦_Holmes代理人.pptx" });
  console.log("Presentation created successfully!");
}

createPresentation().catch(console.error);
