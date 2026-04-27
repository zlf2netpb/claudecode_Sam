const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const { FaBrain, FaCogs, FaChartLine, FaShieldAlt, FaRegLightbulb, FaServer, FaMicrochip, FaIndustry, FaWarehouse, FaSyncAlt, FaArrowRight } = require("react-icons/fa");

// Color Palette - Ocean Gradient theme for AI/Tech
const COLORS = {
  primary: "0F2942",      // Deep navy
  secondary: "1A5276",    // Ocean blue
  accent: "00BCD4",       // Cyan accent
  accentDark: "0097A7",   // Dark cyan
  white: "FFFFFF",
  lightGray: "F5F7FA",
  textDark: "1E293B",
  textMuted: "64748B",
  cardBg: "FFFFFF",
  lobster: "E74C3C",      // Lobster red for branding
  lobsterDark: "C0392B"
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

// Shadow factory to avoid mutation issues
const makeCardShadow = () => ({ type: "outer", color: "000000", blur: 8, offset: 3, angle: 135, opacity: 0.1 });

async function createPresentation() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = "雲端全知龍蝦智能體系統規劃";
  pres.author = "AI大數據部門";

  // Pre-render icons
  const icons = {
    brain: await iconToBase64Png(FaBrain, "#FFFFFF", 256),
    brainDark: await iconToBase64Png(FaBrain, "#1A5276", 256),
    cogs: await iconToBase64Png(FaCogs, "#00BCD4", 256),
    chart: await iconToBase64Png(FaChartLine, "#00BCD4", 256),
    shield: await iconToBase64Png(FaShieldAlt, "#00BCD4", 256),
    lightbulb: await iconToBase64Png(FaRegLightbulb, "#00BCD4", 256),
    server: await iconToBase64Png(FaServer, "#FFFFFF", 256),
    microchip: await iconToBase64Png(FaMicrochip, "#FFFFFF", 256),
    industry: await iconToBase64Png(FaIndustry, "#FFFFFF", 256),
    warehouse: await iconToBase64Png(FaWarehouse, "#FFFFFF", 256),
    sync: await iconToBase64Png(FaSyncAlt, "#00BCD4", 256),
    arrow: await iconToBase64Png(FaArrowRight, "#64748B", 256),
  };

  // ============ SLIDE 1: Title Slide ============
  let slide1 = pres.addSlide();
  slide1.background = { color: COLORS.primary };

  // Decorative gradient overlay shape
  slide1.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 5.625,
    fill: { color: COLORS.secondary, transparency: 40 }
  });

  // Large lobster icon circle (decorative)
  slide1.addShape(pres.shapes.OVAL, {
    x: 7.5, y: 0.8, w: 2.2, h: 2.2,
    fill: { color: COLORS.lobster },
    shadow: { type: "outer", color: "000000", blur: 15, offset: 5, angle: 135, opacity: 0.3 }
  });
  slide1.addImage({
    data: icons.brain,
    x: 7.9, y: 1.2, w: 1.4, h: 1.4
  });

  // Title
  slide1.addText("雲端全知龍蝦智能體", {
    x: 0.5, y: 1.5, w: 7, h: 1.2,
    fontSize: 44, fontFace: "Arial Black", color: COLORS.white,
    bold: true
  });

  // Subtitle
  slide1.addText("Cloud Omnipotent Lobster Agent", {
    x: 0.5, y: 2.7, w: 7, h: 0.5,
    fontSize: 18, fontFace: "Arial", color: COLORS.accent,
    italic: true
  });

  // Tagline
  slide1.addText("工廠級 AI 全局協調智能體系統", {
    x: 0.5, y: 3.5, w: 7, h: 0.5,
    fontSize: 20, fontFace: "Arial", color: COLORS.white
  });

  // Bottom info bar
  slide1.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 4.8, w: 10, h: 0.825,
    fill: { color: COLORS.accentDark }
  });
  slide1.addText("AI大數據部門  |  智能製造規劃", {
    x: 0.5, y: 4.95, w: 9, h: 0.5,
    fontSize: 14, fontFace: "Arial", color: COLORS.white,
    align: "left"
  });

  // ============ SLIDE 2: System Architecture ============
  let slide2 = pres.addSlide();
  slide2.background = { color: COLORS.lightGray };

  // Title bar
  slide2.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.1,
    fill: { color: COLORS.primary }
  });
  slide2.addText("系統架構", {
    x: 0.5, y: 0.3, w: 9, h: 0.6,
    fontSize: 32, fontFace: "Arial Black", color: COLORS.white, bold: true, margin: 0
  });

  // Cloud Lobster layer (top)
  slide2.addShape(pres.shapes.RECTANGLE, {
    x: 2.5, y: 1.4, w: 5, h: 1.0,
    fill: { color: COLORS.lobster },
    shadow: makeCardShadow()
  });
  slide2.addImage({ data: icons.brain, x: 2.7, y: 1.55, w: 0.7, h: 0.7 });
  slide2.addText("雲端全知龍蝦（旗艦智能體）", {
    x: 3.5, y: 1.55, w: 3.8, h: 0.7,
    fontSize: 16, fontFace: "Arial", color: COLORS.white, bold: true, valign: "middle"
  });

  // Arrow down
  slide2.addImage({ data: icons.arrow, x: 4.75, y: 2.45, w: 0.5, h: 0.5 });

  // Holmes Integration layer (middle)
  slide2.addShape(pres.shapes.RECTANGLE, {
    x: 1.5, y: 2.9, w: 7, h: 0.8,
    fill: { color: COLORS.secondary },
    shadow: makeCardShadow()
  });
  slide2.addText("Holmes 產品矩陣（智能生產 | 智能品質 | 知識管理）", {
    x: 1.7, y: 3.0, w: 6.6, h: 0.6,
    fontSize: 14, fontFace: "Arial", color: COLORS.white, valign: "middle"
  });

  // Arrow down
  slide2.addImage({ data: icons.arrow, x: 4.75, y: 3.75, w: 0.5, h: 0.5 });

  // Edge Lobsters layer (bottom)
  const factoryLabels = ["工廠A 龍蝦", "工廠B 龍蝦", "工廠C 龍蝦", "...N 工廠"];
  const factoryColors = [COLORS.accent, COLORS.accentDark, COLORS.accent, COLORS.textMuted];
  const factoryX = [0.5, 2.8, 5.1, 7.4];

  factoryLabels.forEach((label, i) => {
    slide2.addShape(pres.shapes.RECTANGLE, {
      x: factoryX[i], y: 4.3, w: 2.1, h: 0.9,
      fill: { color: factoryColors[i] },
      shadow: makeCardShadow()
    });
    slide2.addText(label, {
      x: factoryX[i], y: 4.45, w: 2.1, h: 0.6,
      fontSize: 13, fontFace: "Arial", color: COLORS.white, bold: true,
      align: "center", valign: "middle"
    });
  });

  // Data layer label
  slide2.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 5.25, w: 9, h: 0.35,
    fill: { color: COLORS.textDark }
  });
  slide2.addText("數據層（集中式數據湖：時序數據 | 關係型數據 | 文件數據 | 向量數據）", {
    x: 0.5, y: 5.25, w: 9, h: 0.35,
    fontSize: 10, fontFace: "Arial", color: COLORS.white, align: "center", valign: "middle"
  });

  // ============ SLIDE 3: Core Capabilities ============
  let slide3 = pres.addSlide();
  slide3.background = { color: COLORS.lightGray };

  // Title bar
  slide3.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.1,
    fill: { color: COLORS.primary }
  });
  slide3.addText("核心能力", {
    x: 0.5, y: 0.3, w: 9, h: 0.6,
    fontSize: 32, fontFace: "Arial Black", color: COLORS.white, bold: true, margin: 0
  });

  // Capability cards (2x3 grid)
  const capabilities = [
    { icon: icons.chart, title: "全局監控", desc: "跨工廠產能利用率即時視圖\n設備OEE全局追蹤", action: "自動生成每日工廠健康報告" },
    { icon: icons.microchip, title: "異常預測", desc: "設備故障預測（2-4週前）\n品質異常提前预警", action: "自動觸發預防性維護工單" },
    { icon: icons.cogs, title: "資源優化", desc: "跨廠訂單分配優化\n產能預測與瓶頸識別", action: "自動調整生產排程" },
    { icon: icons.lightbulb, title: "決策支持", desc: "多廠區資源競合仲裁\n緊急訂單插單模擬", action: "生成決策建議報告" },
    { icon: icons.brainDark, title: "知識整合", desc: "跨工廠最佳實踐萃取\n異常根因知識圖譜", action: "回答管理層各類查詢" },
    { icon: icons.shield, title: "安全審計", desc: "操作日誌完整記錄\n異常決策可解釋性追溯", action: "每月自動生成安全審計報告" },
  ];

  const cardW = 2.9, cardH = 1.35, startX = 0.5, startY = 1.3, gapX = 0.2, gapY = 0.15;

  capabilities.forEach((cap, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = startX + col * (cardW + gapX);
    const y = startY + row * (cardH + gapY);

    // Card background
    slide3.addShape(pres.shapes.RECTANGLE, {
      x, y, w: cardW, h: cardH,
      fill: { color: COLORS.cardBg },
      shadow: makeCardShadow()
    });

    // Left accent bar
    slide3.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.06, h: cardH,
      fill: { color: COLORS.accent }
    });

    // Icon
    slide3.addImage({ data: cap.icon, x: x + 0.15, y: y + 0.15, w: 0.45, h: 0.45 });

    // Title
    slide3.addText(cap.title, {
      x: x + 0.65, y: y + 0.12, w: 2.1, h: 0.35,
      fontSize: 14, fontFace: "Arial", color: COLORS.textDark, bold: true, margin: 0
    });

    // Description
    slide3.addText(cap.desc, {
      x: x + 0.15, y: y + 0.5, w: 2.6, h: 0.45,
      fontSize: 9, fontFace: "Arial", color: COLORS.textMuted, margin: 0
    });

    // Action tag
    slide3.addShape(pres.shapes.RECTANGLE, {
      x: x + 0.15, y: y + 1.0, w: 2.6, h: 0.28,
      fill: { color: COLORS.accent, transparency: 15 }
    });
    slide3.addText("→ " + cap.action, {
      x: x + 0.15, y: y + 1.0, w: 2.6, h: 0.28,
      fontSize: 8, fontFace: "Arial", color: COLORS.accentDark, valign: "middle", margin: 0
    });
  });

  // ============ SLIDE 4: Multi-Agent Collaboration ============
  let slide4 = pres.addSlide();
  slide4.background = { color: COLORS.lightGray };

  // Title bar
  slide4.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.1,
    fill: { color: COLORS.primary }
  });
  slide4.addText("多智能體協作機制", {
    x: 0.5, y: 0.3, w: 9, h: 0.6,
    fontSize: 32, fontFace: "Arial Black", color: COLORS.white, bold: true, margin: 0
  });

  // Left panel - Cloud Agent
  slide4.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.4, w: 4.3, h: 3.9,
    fill: { color: COLORS.cardBg },
    shadow: makeCardShadow()
  });
  slide4.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.4, w: 4.3, h: 0.6,
    fill: { color: COLORS.lobster }
  });
  slide4.addImage({ data: icons.brain, x: 0.7, y: 1.5, w: 0.4, h: 0.4 });
  slide4.addText("雲端全知龍蝦", {
    x: 1.2, y: 1.48, w: 3.4, h: 0.5,
    fontSize: 16, fontFace: "Arial", color: COLORS.white, bold: true, valign: "middle", margin: 0
  });

  // Cloud agent responsibilities
  const cloudTasks = [
    "全局狀態管理與資源調度",
    "跨廠區異常預測統籌",
    "策略決策引擎（多目標優化）",
    "知識蒸餾與萃取",
    "例外上報與審計"
  ];
  cloudTasks.forEach((task, i) => {
    slide4.addText([
      { text: "✓ ", options: { color: COLORS.accent, bold: true } },
      { text: task, options: { color: COLORS.textDark } }
    ], {
      x: 0.7, y: 2.15 + i * 0.55, w: 3.9, h: 0.45,
      fontSize: 13, fontFace: "Arial", margin: 0
    });
  });

  // Right panel - Edge Agents
  slide4.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.4, w: 4.3, h: 3.9,
    fill: { color: COLORS.cardBg },
    shadow: makeCardShadow()
  });
  slide4.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.4, w: 4.3, h: 0.6,
    fill: { color: COLORS.accent }
  });
  slide4.addImage({ data: icons.microchip, x: 5.4, y: 1.5, w: 0.4, h: 0.4 });
  slide4.addText("邊端龍蝦（每工廠一個）", {
    x: 5.9, y: 1.48, w: 3.4, h: 0.5,
    fontSize: 16, fontFace: "Arial", color: COLORS.white, bold: true, valign: "middle", margin: 0
  });

  // Edge agent responsibilities
  const edgeTasks = [
    "本地異常即時偵測（SPC/振動/AOI）",
    "參數微調與即時控制",
    "數據預處理與壓縮",
    "執行雲端下發指令",
    "緊急事件優先本地處置"
  ];
  edgeTasks.forEach((task, i) => {
    slide4.addText([
      { text: "✓ ", options: { color: COLORS.accent, bold: true } },
      { text: task, options: { color: COLORS.textDark } }
    ], {
      x: 5.4, y: 2.15 + i * 0.55, w: 3.9, h: 0.45,
      fontSize: 13, fontFace: "Arial", margin: 0
    });
  });

  // ============ SLIDE 5: Action Execution Framework ============
  let slide5 = pres.addSlide();
  slide5.background = { color: COLORS.lightGray };

  // Title bar
  slide5.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.1,
    fill: { color: COLORS.primary }
  });
  slide5.addText("Action 執行框架", {
    x: 0.5, y: 0.3, w: 9, h: 0.6,
    fontSize: 32, fontFace: "Arial Black", color: COLORS.white, bold: true, margin: 0
  });

  // Three risk level cards
  const riskLevels = [
    {
      level: "低風險", color: "27AE60",
      items: ["數據讀取/報表生成", "SPC管制界限調整±3%", "報警閾值調整"],
      automation: "全自動執行"
    },
    {
      level: "中風險", color: "F39C12",
      items: ["同一工廠內訂單順序調整", "批次隔離處置", "良品/不良品重新分類"],
      automation: "半自動（人類確認）"
    },
    {
      level: "高風險", color: "E74C3C",
      items: ["跨廠區訂單轉移", "緊急插單調度", "遠程停機/程序下載"],
      automation: "需要人類授權"
    }
  ];

  const riskCardW = 2.9, riskCardH = 2.8, riskStartX = 0.5, riskStartY = 1.35, riskGapX = 0.2;

  riskLevels.forEach((risk, i) => {
    const x = riskStartX + i * (riskCardW + riskGapX);
    const y = riskStartY;

    // Card background
    slide5.addShape(pres.shapes.RECTANGLE, {
      x, y, w: riskCardW, h: riskCardH,
      fill: { color: COLORS.cardBg },
      shadow: makeCardShadow()
    });

    // Top color bar
    slide5.addShape(pres.shapes.RECTANGLE, {
      x, y, w: riskCardW, h: 0.5,
      fill: { color: risk.color }
    });

    // Level title
    slide5.addText(risk.level, {
      x, y: y + 0.05, w: riskCardW, h: 0.4,
      fontSize: 16, fontFace: "Arial", color: COLORS.white, bold: true, align: "center", valign: "middle"
    });

    // Items
    risk.items.forEach((item, j) => {
      slide5.addText([
        { text: "• ", options: { color: risk.color, bold: true } },
        { text: item, options: { color: COLORS.textDark } }
      ], {
        x: x + 0.15, y: y + 0.6 + j * 0.5, w: riskCardW - 0.3, h: 0.45,
        fontSize: 11, fontFace: "Arial", margin: 0
      });
    });

    // Automation tag
    slide5.addShape(pres.shapes.RECTANGLE, {
      x: x + 0.1, y: y + 2.3, w: riskCardW - 0.2, h: 0.4,
      fill: { color: risk.color, transparency: 15 }
    });
    slide5.addText(risk.automation, {
      x: x + 0.1, y: y + 2.3, w: riskCardW - 0.2, h: 0.4,
      fontSize: 11, fontFace: "Arial", color: risk.color, bold: true, align: "center", valign: "middle"
    });
  });

  // Bottom section - Escalation logic
  slide5.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.35, w: 9, h: 1.2,
    fill: { color: COLORS.cardBg },
    shadow: makeCardShadow()
  });
  slide5.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.35, w: 0.08, h: 1.2,
    fill: { color: COLORS.secondary }
  });
  slide5.addText("例外上報觸發條件", {
    x: 0.7, y: 4.45, w: 8.6, h: 0.35,
    fontSize: 14, fontFace: "Arial", color: COLORS.textDark, bold: true, margin: 0
  });
  slide5.addText("高風險設備 | 影響範圍 > 10% 產能 | 跨廠區影響 | 涉及產品安全 | 預測置信度 < 85% | 同一問題重複報警 > 3次", {
    x: 0.7, y: 4.85, w: 8.6, h: 0.6,
    fontSize: 11, fontFace: "Arial", color: COLORS.textMuted, margin: 0
  });

  // ============ SLIDE 6: Application Scenarios ============
  let slide6 = pres.addSlide();
  slide6.background = { color: COLORS.lightGray };

  // Title bar
  slide6.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.1,
    fill: { color: COLORS.primary }
  });
  slide6.addText("關鍵應用場景", {
    x: 0.5, y: 0.3, w: 9, h: 0.6,
    fontSize: 32, fontFace: "Arial Black", color: COLORS.white, bold: true, margin: 0
  });

  // Scenario cards (2 columns, 2 rows)
  const scenarios = [
    {
      num: "01",
      title: "跨廠區設備故障預測與自動調度",
      value: "原本 2-3 天停機 → 變成計劃性 4 小時維護",
      icon: icons.microchip
    },
    {
      num: "02",
      title: "品質異常根因分析",
      value: "原本 2-3 天排查 → 縮短到 2 小時內完成",
      icon: icons.chart
    },
    {
      num: "03",
      title: "緊急訂單插單仲裁",
      value: "快速提供數據驅動決策依據，優化資源配置",
      icon: icons.cogs
    },
    {
      num: "04",
      title: "能源消耗優化調度",
      value: "系統性識別節能機會，達成碳中和目標",
      icon: icons.lightbulb
    }
  ];

  const scenW = 4.4, scenH = 1.7, scenStartX = 0.5, scenStartY = 1.3, scenGapX = 0.2, scenGapY = 0.2;

  scenarios.forEach((scen, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = scenStartX + col * (scenW + scenGapX);
    const y = scenStartY + row * (scenH + scenGapY);

    // Card
    slide6.addShape(pres.shapes.RECTANGLE, {
      x, y, w: scenW, h: scenH,
      fill: { color: COLORS.cardBg },
      shadow: makeCardShadow()
    });

    // Number badge
    slide6.addShape(pres.shapes.OVAL, {
      x: x + 0.15, y: y + 0.2, w: 0.6, h: 0.6,
      fill: { color: COLORS.accent }
    });
    slide6.addText(scen.num, {
      x: x + 0.15, y: y + 0.28, w: 0.6, h: 0.45,
      fontSize: 14, fontFace: "Arial", color: COLORS.white, bold: true, align: "center", valign: "middle"
    });

    // Title
    slide6.addText(scen.title, {
      x: x + 0.9, y: y + 0.2, w: 3.3, h: 0.45,
      fontSize: 14, fontFace: "Arial", color: COLORS.textDark, bold: true, margin: 0
    });

    // Icon
    slide6.addImage({ data: scen.icon, x: x + 3.7, y: y + 0.2, w: 0.5, h: 0.5 });

    // Value proposition
    slide6.addShape(pres.shapes.RECTANGLE, {
      x: x + 0.15, y: y + 0.9, w: 4.1, h: 0.65,
      fill: { color: COLORS.accent, transparency: 10 }
    });
    slide6.addText(scen.value, {
      x: x + 0.25, y: y + 0.95, w: 3.9, h: 0.55,
      fontSize: 11, fontFace: "Arial", color: COLORS.accentDark, valign: "middle", margin: 0
    });
  });

  // Bottom call-to-action bar
  slide6.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 5.1, w: 10, h: 0.525,
    fill: { color: COLORS.lobster }
  });
  slide6.addText("不只是警示問題，而是實際幫助人類解決問題、創造價值", {
    x: 0.5, y: 5.15, w: 9, h: 0.45,
    fontSize: 14, fontFace: "Arial", color: COLORS.white, bold: true, align: "center", valign: "middle"
  });

  // Save the presentation
  await pres.writeFile({ fileName: "/Users/sam/Documents/MyProject/雲端全知龍蝦智能體規劃.pptx" });
  console.log("Presentation created successfully!");
}

createPresentation().catch(console.error);
