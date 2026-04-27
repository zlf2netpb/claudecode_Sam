const pptxgen = require("pptxgenjs");

let pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Holmes雲龍蝦";
pres.author = "Holmes";

// Color palette - Ocean Premium
const COLORS = {
  deepBlue: "0C4A6E",
  ocean: "0369A1",
  coral: "F97316",
  gold: "FBBF24",
  cream: "FEF3C7",
  white: "FFFFFF",
  darkText: "1E293B",
  lightText: "64748B",
};

// ========== SLIDE 1: Title Slide ==========
let slide1 = pres.addSlide();
slide1.background = { color: COLORS.deepBlue };

// Decorative top bar
slide1.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.15,
  fill: { color: COLORS.coral }
});

// Main title
slide1.addText("Holmes 雲龍蝦", {
  x: 0.5, y: 1.8, w: 9, h: 1.2,
  fontSize: 54, fontFace: "Arial Black", color: COLORS.white,
  bold: true, align: "center"
});

// Subtitle
slide1.addText("全知全能 · 雲端之選", {
  x: 0.5, y: 3.0, w: 9, h: 0.6,
  fontSize: 28, fontFace: "Arial", color: COLORS.gold,
  align: "center", charSpacing: 4
});

// Decorative lobster icon placeholder (oval shape)
slide1.addShape(pres.shapes.OVAL, {
  x: 4.25, y: 3.8, w: 1.5, h: 1.0,
  fill: { color: COLORS.coral, transparency: 30 },
  line: { color: COLORS.gold, width: 2 }
});

// Bottom tagline
slide1.addText("頂級海鮮 · 品味非凡", {
  x: 0.5, y: 5.0, w: 9, h: 0.4,
  fontSize: 14, fontFace: "Arial", color: COLORS.cream,
  align: "center", charSpacing: 2
});

// ========== SLIDE 2: Introduction ==========
let slide2 = pres.addSlide();
slide2.background = { color: COLORS.white };

// Left accent bar
slide2.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 0.12, h: 5.625,
  fill: { color: COLORS.ocean }
});

// Title
slide2.addText("關於雲龍蝦", {
  x: 0.5, y: 0.4, w: 9, h: 0.8,
  fontSize: 36, fontFace: "Arial Black", color: COLORS.deepBlue,
  bold: true
});

// Content card background
slide2.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 1.4, w: 9, h: 3.8,
  fill: { color: "F0F9FF" },
  line: { color: COLORS.ocean, width: 1 }
});

// Content text
slide2.addText([
  { text: "雲龍蝦（Cloud Dragon Lobster）", options: { bold: true, color: COLORS.ocean, breakLine: true } },
  { text: "\n", options: { breakLine: true } },
  { text: "來自深海的珍稀美味，肉質鮮嫩甜美，富有彈性。", options: { breakLine: true } },
  { text: "\n", options: { breakLine: true } },
  { text: "✨ ", options: { color: COLORS.gold } },
  { text: "生長於純淨海域，無污染、零添加", options: { breakLine: true } },
  { text: "✨ ", options: { color: COLORS.gold } },
  { text: "人工精選，品質保証", options: { breakLine: true } },
  { text: "✨ ", options: { color: COLORS.gold } },
  { text: "急速冷凍，保留最鮮風味", options: { breakLine: true } },
  { text: "✨ 全流程可追溯，安全放心", options: { breakLine: false } }
], {
  x: 0.8, y: 1.6, w: 8.4, h: 3.4,
  fontSize: 18, fontFace: "Arial", color: COLORS.darkText,
  valign: "top", lineSpaceMult: 1.5
});

// ========== SLIDE 3: Features ==========
let slide3 = pres.addSlide();
slide3.background = { color: "F8FAFC" };

// Title
slide3.addText("雲龍蝦特色", {
  x: 0.5, y: 0.3, w: 9, h: 0.8,
  fontSize: 36, fontFace: "Arial Black", color: COLORS.deepBlue,
  bold: true
});

// Feature cards - 2x2 grid
const features = [
  { title: "頂級品質", desc: "精選A级雲龍蝦\n肉質飽滿鮮嫩", color: COLORS.coral },
  { title: "營養價值", desc: "高蛋白低脂肪\n富含微量元素", color: COLORS.ocean },
  { title: "美味料理", desc: "多種烹調方式\n蒸炒煮炸皆宜", color: COLORS.gold },
  { title: "尊貴體驗", desc: "送禮自用兩相宜\n宴客首選海鮮", color: COLORS.deepBlue }
];

const cardW = 4.2;
const cardH = 1.8;
const startX = 0.5;
const startY = 1.3;
const gapX = 0.35;
const gapY = 0.3;

features.forEach((feat, i) => {
  const col = i % 2;
  const row = Math.floor(i / 2);
  const x = startX + col * (cardW + gapX);
  const y = startY + row * (cardH + gapY);

  // Card background
  slide3.addShape(pres.shapes.RECTANGLE, {
    x: x, y: y, w: cardW, h: cardH,
    fill: { color: COLORS.white },
    shadow: { type: "outer", color: "000000", blur: 8, offset: 2, angle: 135, opacity: 0.1 }
  });

  // Color accent bar on left
  slide3.addShape(pres.shapes.RECTANGLE, {
    x: x, y: y, w: 0.08, h: cardH,
    fill: { color: feat.color }
  });

  // Title
  slide3.addText(feat.title, {
    x: x + 0.25, y: y + 0.2, w: cardW - 0.4, h: 0.5,
    fontSize: 20, fontFace: "Arial", color: feat.color,
    bold: true, margin: 0
  });

  // Description
  slide3.addText(feat.desc, {
    x: x + 0.25, y: y + 0.75, w: cardW - 0.4, h: 0.9,
    fontSize: 14, fontFace: "Arial", color: COLORS.lightText,
    margin: 0
  });
});

// ========== SLIDE 4: Contact/CTA ==========
let slide4 = pres.addSlide();
slide4.background = { color: COLORS.deepBlue };

// Decorative circle
slide4.addShape(pres.shapes.OVAL, {
  x: 6.5, y: -1.5, w: 5, h: 5,
  fill: { color: COLORS.ocean, transparency: 50 }
});

slide4.addShape(pres.shapes.OVAL, {
  x: 7.5, y: 3, w: 4, h: 4,
  fill: { color: COLORS.coral, transparency: 60 }
});

// Main CTA
slide4.addText("立即品嘗", {
  x: 0.5, y: 1.5, w: 6, h: 1.0,
  fontSize: 48, fontFace: "Arial Black", color: COLORS.white,
  bold: true
});

slide4.addText("全知全能 Holmes 雲龍蝦", {
  x: 0.5, y: 2.5, w: 6, h: 0.6,
  fontSize: 24, fontFace: "Arial", color: COLORS.gold
});

slide4.addText("讓您的味蕾體驗前所未有的海洋鮮甜", {
  x: 0.5, y: 3.2, w: 6, h: 0.5,
  fontSize: 16, fontFace: "Arial", color: COLORS.cream
});

// Contact info
slide4.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 4.2, w: 5, h: 0.06,
  fill: { color: COLORS.gold }
});

slide4.addText("諮詢熱線：+886-2-XXXX-XXXX", {
  x: 0.5, y: 4.5, w: 5, h: 0.4,
  fontSize: 14, fontFace: "Arial", color: COLORS.cream
});

slide4.addText("www.holmes-lobster.com", {
  x: 0.5, y: 4.9, w: 5, h: 0.4,
  fontSize: 14, fontFace: "Arial", color: COLORS.cream
});

// Save
pres.writeFile({ fileName: "/Users/sam/Documents/MyProject/ppt-maker/Holmes雲龍蝦.pptx" })
  .then(() => console.log("PPT created: Holmes雲龍蝦.pptx"))
  .catch(err => console.error(err));
