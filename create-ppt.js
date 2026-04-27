const pptxgen = require("pptxgenjs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.title = '智慧製造技術架構與應用';
pres.author = 'Smart Manufacturing Research';

// Color palette - Teal Trust with custom accents
const colors = {
  primary: "065A82",      // Deep blue
  secondary: "1C7293",    // Teal
  accent: "00A896",       // Seafoam
  light: "E8F4F8",        // Light blue-gray
  dark: "0A3D62",         // Dark navy
  white: "FFFFFF",
  text: "2D3748",
  muted: "718096"
};

// ==================== SLIDE 1: Title ====================
let slide1 = pres.addSlide();
slide1.background = { path: "/Users/sam/Documents/MyProject/slide1-bg.jpeg" };

// Dark overlay for readability
slide1.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 5.625,
  fill: { color: "0A3D62", transparency: 55 }
});

// Title
slide1.addText("智慧製造", {
  x: 0.5, y: 1.8, w: 9, h: 1.2,
  fontSize: 54, fontFace: "Arial Black", color: "FFFFFF", bold: true, align: "center"
});

slide1.addText("技術架構與應用", {
  x: 0.5, y: 2.9, w: 9, h: 0.7,
  fontSize: 32, fontFace: "Arial", color: "00A896", align: "center"
});

// Subtitle
slide1.addText("從設備層到應用層的完整技術地圖", {
  x: 0.5, y: 4.0, w: 9, h: 0.5,
  fontSize: 16, fontFace: "Arial", color: "CADCFC", align: "center"
});

// Bottom accent bar
slide1.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 5.35, w: 10, h: 0.275,
  fill: { color: "00A896" }
});

// ==================== SLIDE 2: Architecture ====================
let slide2 = pres.addSlide();
slide2.background = { path: "/Users/sam/Documents/MyProject/slide2-architecture.jpeg" };

// Semi-transparent overlay
slide2.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 5.625,
  fill: { color: "FFFFFF", transparency: 70 }
});

// Title bar
slide2.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 1.0,
  fill: { color: "065A82" }
});

slide2.addText("技術架構", {
  x: 0.5, y: 0.25, w: 9, h: 0.6,
  fontSize: 32, fontFace: "Arial Black", color: "FFFFFF", bold: true, margin: 0
});

// Architecture layers - 4 boxes
const layers = [
  { name: "設備層", desc: "感測器、IoT閘道器、智慧控制器", icon: "感知層" },
  { name: "網路層", desc: "5G、TSN、工業乙太網路、MEC", icon: "連接層" },
  { name: "平台層", desc: "雲端運算、邊緣運算、資料湖", icon: "運算層" },
  { name: "應用層", desc: "AI分析、數位孿生、MES/SCADA", icon: "智慧層" }
];

const layerColors = ["1C7293", "0D9488", "0891B2", "065A82"];
const startY = 1.4;
const boxH = 0.9;
const gap = 0.15;

layers.forEach((layer, i) => {
  const y = startY + i * (boxH + gap);

  // Left accent bar
  slide2.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: y, w: 0.08, h: boxH,
    fill: { color: layerColors[i] }
  });

  // Card background
  slide2.addShape(pres.shapes.RECTANGLE, {
    x: 0.58, y: y, w: 9.0, h: boxH,
    fill: { color: "FFFFFF" },
    shadow: { type: "outer", color: "000000", blur: 4, offset: 2, angle: 135, opacity: 0.1 }
  });

  // Layer number
  slide2.addShape(pres.shapes.OVAL, {
    x: 0.75, y: y + 0.2, w: 0.5, h: 0.5,
    fill: { color: layerColors[i] }
  });
  slide2.addText(String(i + 1), {
    x: 0.75, y: y + 0.2, w: 0.5, h: 0.5,
    fontSize: 16, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle"
  });

  // Layer name
  slide2.addText(layer.name, {
    x: 1.4, y: y + 0.15, w: 2, h: 0.35,
    fontSize: 18, fontFace: "Arial", color: "0A3D62", bold: true, margin: 0
  });

  // Layer description
  slide2.addText(layer.desc, {
    x: 1.4, y: y + 0.48, w: 7.5, h: 0.35,
    fontSize: 13, fontFace: "Arial", color: "718096", margin: 0
  });

  // Icon badge
  slide2.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 8.5, y: y + 0.25, w: 0.9, h: 0.4,
    fill: { color: layerColors[i], transparency: 15 }, rectRadius: 0.05
  });
  slide2.addText(layer.icon, {
    x: 8.5, y: y + 0.25, w: 0.9, h: 0.4,
    fontSize: 10, fontFace: "Arial", color: layerColors[i], bold: true, align: "center", valign: "middle"
  });
});

// Footer note
slide2.addText("基於工業4.0 RAMI 4.0 參考架構", {
  x: 0.5, y: 5.2, w: 9, h: 0.3,
  fontSize: 10, fontFace: "Arial", color: "718096", align: "right"
});

// ==================== SLIDE 3: Applications ====================
let slide3 = pres.addSlide();
slide3.background = { path: "/Users/sam/Documents/MyProject/slide3-applications.jpeg" };

// Dark overlay
slide3.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 5.625,
  fill: { color: "065A82", transparency: 60 }
});

// Title
slide3.addText("核心應用場景", {
  x: 0.5, y: 0.3, w: 9, h: 0.7,
  fontSize: 32, fontFace: "Arial Black", color: "FFFFFF", bold: true
});

// 4 application cards in 2x2 grid
const apps = [
  { title: "預測性維護", desc: "AI 預測設備故障，提前安排維修，減少非計劃停機時間", metric: "↓ 30% 停機" },
  { title: "智慧品質管理", desc: "電腦視覺自動檢測，深度學習提升檢出率達 99%+", metric: "↑ 99% 檢出" },
  { title: "彈性製造", desc: "少量多樣、快速換線，滿足市場多元需求", metric: "↑ 50% 彈性" },
  { title: "供應鏈優化", desc: "端到端可視化，需求預測與物流路徑最佳化", metric: "↓ 20% 庫存" }
];

const cardW = 4.3;
const cardH = 1.8;
const cardStartX = 0.5;
const cardStartY = 1.2;
const cardGapX = 0.4;
const cardGapY = 0.3;

apps.forEach((app, i) => {
  const col = i % 2;
  const row = Math.floor(i / 2);
  const x = cardStartX + col * (cardW + cardGapX);
  const y = cardStartY + row * (cardH + cardGapY);

  // Card background
  slide3.addShape(pres.shapes.RECTANGLE, {
    x: x, y: y, w: cardW, h: cardH,
    fill: { color: "FFFFFF", transparency: 10 },
    line: { color: "FFFFFF", width: 1 }
  });

  // Accent top bar
  slide3.addShape(pres.shapes.RECTANGLE, {
    x: x, y: y, w: cardW, h: 0.06,
    fill: { color: "00A896" }
  });

  // Title
  slide3.addText(app.title, {
    x: x + 0.2, y: y + 0.2, w: 3.5, h: 0.4,
    fontSize: 16, fontFace: "Arial", color: "FFFFFF", bold: true, margin: 0
  });

  // Description
  slide3.addText(app.desc, {
    x: x + 0.2, y: y + 0.65, w: 3.9, h: 0.7,
    fontSize: 11, fontFace: "Arial", color: "CADCFC", margin: 0
  });

  // Metric badge
  slide3.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: x + 2.8, y: y + 1.35, w: 1.3, h: 0.35,
    fill: { color: "00A896" }, rectRadius: 0.05
  });
  slide3.addText(app.metric, {
    x: x + 2.8, y: y + 1.35, w: 1.3, h: 0.35,
    fontSize: 10, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle"
  });
});

// Bottom trend section
slide3.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 4.7, w: 9, h: 0.7,
  fill: { color: "0A3D62", transparency: 50 }
});

slide3.addText("發展趨勢：生成式AI 賦能製造 │ 自主化系統 │ 永續綠色製造", {
  x: 0.5, y: 4.7, w: 9, h: 0.7,
  fontSize: 13, fontFace: "Arial", color: "00A896", align: "center", valign: "middle"
});

// Save
pres.writeFile({ fileName: "/Users/sam/Documents/MyProject/智慧製造.pptx" })
  .then(() => console.log("PPT created: 智慧製造.pptx"))
  .catch(err => console.error(err));
