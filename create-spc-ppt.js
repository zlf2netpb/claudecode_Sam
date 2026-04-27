const pptxgen = require("pptxgenjs");

const colors = {
  primary: "1E3A5F",
  secondary: "3B82F6",
  accent: "10B981",
  accent2: "F59E0B",
  dark: "0F172A",
  card: "1E293B",
  light: "F1F5F9",
  white: "FFFFFF",
  muted: "94A3B8"
};

const basePath = "/Users/sam/Documents/MyProject/";
const makeShadow = () => ({ type: "outer", color: "000000", blur: 8, offset: 3, angle: 135, opacity: 0.2 });

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.title = 'SPC AI Agent 專家';
pres.author = 'Claude AI';

// SLIDE 1: Cover
let slide1 = pres.addSlide();
slide1.background = { color: colors.dark };
slide1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: colors.primary, transparency: 40 } });
slide1.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.5, w: 9, h: 2.8, fill: { color: colors.card, transparency: 20 } });
slide1.addText("SPC AI Agent 專家", { x: 0.5, y: 1.7, w: 9, h: 1, fontSize: 44, fontFace: "Arial Black", color: "FFFFFF", bold: true, align: "center" });
slide1.addText("Statistical Process Control with AI Intelligence", { x: 0.5, y: 2.7, w: 9, h: 0.5, fontSize: 20, fontFace: "Arial", color: "3B82F6", align: "center" });
slide1.addText("智慧品質管制系統 - AI 驅動的 SPC 分析專家", { x: 0.5, y: 3.3, w: 9, h: 0.4, fontSize: 16, fontFace: "Arial", color: "94A3B8", align: "center" });
slide1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.35, w: 10, h: 0.275, fill: { color: "3B82F6" } });
slide1.addText("2026年4月", { x: 0.5, y: 5.0, w: 9, h: 0.3, fontSize: 12, fontFace: "Arial", color: "94A3B8", align: "center" });

// SLIDE 2: Architecture
let slide2 = pres.addSlide();
slide2.background = { color: colors.dark };
slide2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: colors.primary } });
slide2.addText("系統架構設計", { x: 0.5, y: 0.2, w: 9, h: 0.5, fontSize: 28, fontFace: "Arial Black", color: "FFFFFF", bold: true, margin: 0 });

const archLayers = [
  { name: "應用層", desc: "Dashboard / Web UI / Mobile App / API", color: "3B82F6", y: 1.1 },
  { name: "AI 分析引擎", desc: "SPC AI Agent / 異常偵測 / 預測模型", color: "8B5CF6", y: 2.0 },
  { name: "數據處理層", desc: "即時數據流處理 / 特徵工程 / 統計分析", color: "10B981", y: 2.9 },
  { name: "數據源層", desc: "感測器 / PLC / MES / Excel / CSV", color: "F59E0B", y: 3.8 }
];
archLayers.forEach((layer) => {
  slide2.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: layer.y, w: 3.5, h: 0.7, fill: { color: layer.color }, shadow: makeShadow() });
  slide2.addText(layer.name, { x: 0.5, y: layer.y, w: 3.5, h: 0.7, fontSize: 14, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
  slide2.addShape(pres.shapes.RECTANGLE, { x: 4.2, y: layer.y, w: 5.3, h: 0.7, fill: { color: colors.card } });
  slide2.addText(layer.desc, { x: 4.4, y: layer.y, w: 5.1, h: 0.7, fontSize: 12, fontFace: "Arial", color: colors.light, valign: "middle", margin: 0 });
});
for (let i = 0; i < 3; i++) {
  slide2.addShape(pres.shapes.LINE, { x: 2.25, y: archLayers[i].y + 0.7, w: 0, h: 0.2, line: { color: colors.muted, width: 2 } });
}

// Flow diagram on right side
slide2.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.6, w: 9, h: 0.8, fill: { color: colors.card } });
slide2.addText("數據流：感測器 → 即時處理 → AI 分析 → 異常告警 → Dashboard呈現", { x: 0.7, y: 4.7, w: 8.6, h: 0.6, fontSize: 13, fontFace: "Arial", color: "3B82F6", valign: "middle" });

// SLIDE 3: Prototype UI
let slide3 = pres.addSlide();
slide3.background = { color: colors.dark };
slide3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: colors.primary } });
slide3.addText("Prototype UI 介面設計", { x: 0.5, y: 0.2, w: 9, h: 0.5, fontSize: 28, fontFace: "Arial Black", color: "FFFFFF", bold: true, margin: 0 });

// UI Mockup boxes
const uiElements = [
  { title: "即時戰情中心", desc: "設備狀態 / 異常告警 / 生產良率", x: 0.5, y: 1.2, color: "3B82F6" },
  { title: "SPC 管制圖", desc: "X-bar / R / Cpk 趨勢監控", x: 3.5, y: 1.2, color: "10B981" },
  { title: "異常分析面板", desc: "根因分析 / 處置建議", x: 6.5, y: 1.2, color: "F59E0B" },
  { title: "AI 對話介面", desc: "自然語言查詢 SPC 數據", x: 0.5, y: 3.4, color: "8B5CF6" },
  { title: "報告產生器", desc: "自動生成異常分析報告", x: 3.5, y: 3.4, color: "EF4444" }
];

uiElements.forEach((elem, i) => {
  const col = i % 3;
  const row = Math.floor(i / 3);
  const x = elem.x;
  const y = elem.y;

  slide3.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: 2.8, h: 1.9, fill: { color: colors.card }, shadow: makeShadow() });
  slide3.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: 2.8, h: 0.08, fill: { color: elem.color } });

  // Simulate UI button
  slide3.addShape(pres.shapes.RECTANGLE, { x: x + 0.2, y: y + 0.3, w: 2.4, h: 0.4, fill: { color: elem.color, transparency: 70 } });
  slide3.addShape(pres.shapes.RECTANGLE, { x: x + 0.2, y: y + 0.8, w: 2.4, h: 0.15, fill: { color: colors.muted, transparency: 50 } });
  slide3.addShape(pres.shapes.RECTANGLE, { x: x + 0.2, y: y + 1.0, w: 2.4, h: 0.15, fill: { color: colors.muted, transparency: 50 } });
  slide3.addShape(pres.shapes.RECTANGLE, { x: x + 0.2, y: y + 1.2, w: 1.5, h: 0.15, fill: { color: colors.muted, transparency: 50 } });

  slide3.addText(elem.title, { x: x + 0.15, y: y + 1.5, w: 2.5, h: 0.35, fontSize: 11, fontFace: "Arial", color: "FFFFFF", bold: true, margin: 0 });
});

// SLIDE 4: Statistical Charts
let slide4 = pres.addSlide();
slide4.background = { color: colors.dark };
slide4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: colors.primary } });
slide4.addText("統計圖表分析", { x: 0.5, y: 0.2, w: 9, h: 0.5, fontSize: 28, fontFace: "Arial Black", color: "FFFFFF", bold: true, margin: 0 });

// Chart 1: X-bar Chart (simulated with shapes)
slide4.addText("X-bar 管制圖", { x: 0.5, y: 1.1, w: 4.3, h: 0.4, fontSize: 14, fontFace: "Arial", color: "3B82F6", bold: true });
// Chart background
slide4.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.5, w: 4.3, h: 2.2, fill: { color: colors.card } });
// Simulated line chart using bar shapes for data points
const xbarData = [12.5, 12.8, 11.9, 12.3, 13.1, 12.6, 12.4, 12.9, 13.2, 12.7];
const chartBaseY = 3.5;
const chartHeight = 1.0;
const pointWidth = 0.35;
xbarData.forEach((val, i) => {
  const normalizedH = ((val - 11.5) / 2) * chartHeight;
  const x = 0.65 + i * 0.4;
  const y = chartBaseY - normalizedH;
  slide4.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: pointWidth, h: normalizedH, fill: { color: "3B82F6" } });
});
// UCL/LCL lines
slide4.addShape(pres.shapes.LINE, { x: 0.5, y: 1.85, w: 4.3, h: 0, line: { color: "EF4444", width: 1, dashType: "dash" } });
slide4.addShape(pres.shapes.LINE, { x: 0.5, y: 3.45, w: 4.3, h: 0, line: { color: "EF4444", width: 1, dashType: "dash" } });
slide4.addText("UCL=13.0", { x: 4.5, y: 1.75, w: 0.7, h: 0.2, fontSize: 8, fontFace: "Arial", color: "EF4444" });
slide4.addText("LCL=12.0", { x: 4.5, y: 3.35, w: 0.7, h: 0.2, fontSize: 8, fontFace: "Arial", color: "EF4444" });

// Chart 2: Cpk Gauge (using progress bar style)
slide4.addText("製程能力 Cpk", { x: 5.2, y: 1.1, w: 4.3, h: 0.4, fontSize: 14, fontFace: "Arial", color: "10B981", bold: true });
// Progress bar background
slide4.addShape(pres.shapes.RECTANGLE, { x: 5.5, y: 1.7, w: 3.5, h: 0.5, fill: { color: "293548" } });
// Progress bar fill (83% of 2 = 1.66)
slide4.addShape(pres.shapes.RECTANGLE, { x: 5.5, y: 1.7, w: 2.9, h: 0.5, fill: { color: "10B981" } });
slide4.addText("Cpk = 1.67", { x: 5.5, y: 2.3, w: 3.5, h: 0.5, fontSize: 24, fontFace: "Arial Black", color: "FFFFFF", bold: true, align: "center" });
slide4.addText("目標 > 1.33 ✓", { x: 5.5, y: 2.8, w: 3.5, h: 0.3, fontSize: 12, fontFace: "Arial", color: "10B981", align: "center" });
slide4.addText("優異", { x: 5.5, y: 3.15, w: 3.5, h: 0.3, fontSize: 11, fontFace: "Arial", color: "10B981", align: "center" });

// Stats row
const chartStats = [
  { value: "98.5%", label: "良率", color: "3B82F6" },
  { value: "1.67", label: "Cpk", color: "10B981" },
  { value: "12", label: "異常點/日", color: "F59E0B" }
];
chartStats.forEach((stat, i) => {
  const x = 0.5 + i * 3.1;
  slide4.addShape(pres.shapes.RECTANGLE, { x: x, y: 4.0, w: 2.9, h: 1.3, fill: { color: colors.card }, shadow: makeShadow() });
  slide4.addShape(pres.shapes.RECTANGLE, { x: x, y: 4.0, w: 2.9, h: 0.08, fill: { color: stat.color } });
  slide4.addText(stat.value, { x: x, y: 4.2, w: 2.9, h: 0.7, fontSize: 28, fontFace: "Arial Black", color: stat.color, bold: true, align: "center" });
  slide4.addText(stat.label, { x: x, y: 4.85, w: 2.9, h: 0.35, fontSize: 12, fontFace: "Arial", color: colors.muted, align: "center" });
});

// SLIDE 5: AI Agent Capabilities
let slide5 = pres.addSlide();
slide5.background = { color: colors.dark };
slide5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: colors.primary } });
slide5.addText("AI Agent 核心能力", { x: 0.5, y: 0.2, w: 9, h: 0.5, fontSize: 28, fontFace: "Arial Black", color: "FFFFFF", bold: true, margin: 0 });

const capabilities = [
  { title: "即時異常偵測", desc: "自動偵測超出管制限的數據點，縮短異常發現時間 80%", color: "3B82F6" },
  { title: "智慧根因分析", desc: "AI 分析異常模式，找出影響品質的關鍵因子", color: "10B981" },
  { title: "預測性維護", desc: "預測設備異常，提前預警，避免非計劃停機", color: "F59E0B" },
  { title: "自然語言查詢", desc: "用口語方式查詢 SPC 數據，生成分析報告", color: "8B5CF6" },
  { title: "自動報告生成", desc: "AI 自動生成異常分析報告，節省 70% 人力時間", color: "EF4444" },
  { title: "多源數據整合", desc: "整合感測器、PLC、MES 等多種數據源", color: "06B6D4" }
];

capabilities.forEach((cap, i) => {
  const col = i % 3;
  const row = Math.floor(i / 3);
  const x = 0.5 + col * 3.1;
  const y = 1.2 + row * 2.0;

  slide5.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: 2.95, h: 1.8, fill: { color: colors.card }, shadow: makeShadow() });
  slide5.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: 0.08, h: 1.8, fill: { color: cap.color } });
  slide5.addText(cap.title, { x: x + 0.2, y: y + 0.15, w: 2.6, h: 0.4, fontSize: 13, fontFace: "Arial", color: "FFFFFF", bold: true, margin: 0 });
  slide5.addText(cap.desc, { x: x + 0.2, y: y + 0.6, w: 2.6, h: 1.1, fontSize: 11, fontFace: "Arial", color: colors.muted, margin: 0 });
});

// SLIDE 6: Data Flow & Integration
let slide6 = pres.addSlide();
slide6.background = { color: colors.dark };
slide6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: colors.primary } });
slide6.addText("數據流與系統整合", { x: 0.5, y: 0.2, w: 9, h: 0.5, fontSize: 28, fontFace: "Arial Black", color: "FFFFFF", bold: true, margin: 0 });

// Flow diagram
const flowSteps = [
  { text: "數據收集", x: 0.5, color: "3B82F6" },
  { text: "資料處理", x: 2.4, color: "8B5CF6" },
  { text: "AI 分析", x: 4.3, color: "10B981" },
  { text: "異常告警", x: 6.2, color: "F59E0B" },
  { text: "報告產出", x: 8.1, color: "EF4444" }
];

flowSteps.forEach((step, i) => {
  slide6.addShape(pres.shapes.RECTANGLE, { x: step.x, y: 1.3, w: 1.7, h: 0.9, fill: { color: step.color } });
  slide6.addText(step.text, { x: step.x, y: 1.3, w: 1.7, h: 0.9, fontSize: 10, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
  if (i < flowSteps.length - 1) {
    slide6.addShape(pres.shapes.LINE, { x: step.x + 1.7, y: 1.75, w: 0.5, h: 0, line: { color: colors.muted, width: 2 } });
  }
});

// Data sources
slide6.addText("數據來源", { x: 0.5, y: 2.5, w: 9, h: 0.4, fontSize: 16, fontFace: "Arial", color: "3B82F6", bold: true });

const dataSources = [
  { name: "PLC", desc: "機台運行數據" },
  { name: "感測器", desc: "溫度/壓力/振動" },
  { name: "MES", desc: "生產訂單/不良數據" },
  { name: "QMS", desc: "品質管理系統" },
  { name: "Excel", desc: "手動記錄數據" }
];

dataSources.forEach((src, i) => {
  const x = 0.5 + i * 1.9;
  slide6.addShape(pres.shapes.RECTANGLE, { x: x, y: 3.0, w: 1.7, h: 1.0, fill: { color: colors.card }, shadow: makeShadow() });
  slide6.addText(src.name, { x: x, y: 3.1, w: 1.7, h: 0.5, fontSize: 12, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center" });
  slide6.addText(src.desc, { x: x, y: 3.55, w: 1.7, h: 0.4, fontSize: 9, fontFace: "Arial", color: colors.muted, align: "center" });
});

// Output targets
slide6.addText("輸出應用", { x: 0.5, y: 4.2, w: 9, h: 0.4, fontSize: 16, fontFace: "Arial", color: "10B981", bold: true });

const outputs = [
  { name: "Dashboard", desc: "即時監控面板" },
  { name: "Alert", desc: "LINE/Email 告警" },
  { name: "Report", desc: "自動分析報告" },
  { name: "API", desc: "系統整合介面" }
];

outputs.forEach((out, i) => {
  const x = 0.5 + i * 2.35;
  slide6.addShape(pres.shapes.RECTANGLE, { x: x, y: 4.65, w: 2.15, h: 0.8, fill: { color: colors.card }, shadow: makeShadow() });
  slide6.addText(out.name, { x: x, y: 4.7, w: 2.15, h: 0.4, fontSize: 12, fontFace: "Arial", color: "10B981", bold: true, align: "center" });
  slide6.addText(out.desc, { x: x, y: 5.05, w: 2.15, h: 0.35, fontSize: 9, fontFace: "Arial", color: colors.muted, align: "center" });
});

// SLIDE 7: Expected Results
let slide7 = pres.addSlide();
slide7.background = { color: colors.dark };
slide7.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: colors.primary } });
slide7.addText("預期成效與價值", { x: 0.5, y: 0.2, w: 9, h: 0.5, fontSize: 28, fontFace: "Arial Black", color: "FFFFFF", bold: true, margin: 0 });

const results = [
  { metric: "異常發現時間", before: "4 小時", after: "15 分鐘", improvement: "94%" },
  { metric: "良率提升", before: "95%", after: "98.5%", improvement: "+3.5%" },
  { metric: "報告產出時間", before: "2 小時", after: "5 分鐘", improvement: "96%" },
  { metric: "人力節省", before: "3 人/班", after: "1 人/班", improvement: "-67%" }
];

slide7.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 9, h: 0.6, fill: { color: colors.card } });
const headers = ["指標", "導入前", "導入後", "改善幅度"];
const colWidths = [2.5, 2.5, 2.5, 1.5];
let xPos = 0.5;
headers.forEach((h, i) => {
  slide7.addText(h, { x: xPos, y: 1.2, w: colWidths[i], h: 0.6, fontSize: 12, fontFace: "Arial", color: "3B82F6", bold: true, align: "center", valign: "middle" });
  xPos += colWidths[i];
});

results.forEach((row, idx) => {
  const y = 1.8 + idx * 0.7;
  slide7.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: y, w: 9, h: 0.7, fill: { color: idx % 2 === 0 ? colors.card : "293548" } });
  let x = 0.5;
  const values = [row.metric, row.before, row.after, row.improvement];
  values.forEach((v, i) => {
    slide7.addText(v, { x: x, y: y, w: colWidths[i], h: 0.7, fontSize: 12, fontFace: "Arial", color: i === 3 ? "10B981" : colors.light, bold: i === 3, align: "center", valign: "middle" });
    x += colWidths[i];
  });
});

// ROI box
slide7.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.6, w: 9, h: 0.8, fill: { color: "10B981", transparency: 80 } });
slide7.addText("預估投資報酬率：6 個月內達成 ROI > 150%", { x: 0.5, y: 4.6, w: 9, h: 0.8, fontSize: 16, fontFace: "Arial", color: "10B981", bold: true, align: "center", valign: "middle" });

// Footer
slide7.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.35, w: 10, h: 0.275, fill: { color: "3B82F6" } });
slide7.addText("SPC AI Agent 專家 - 智慧品質管制系統", { x: 0.5, y: 5.0, w: 9, h: 0.3, fontSize: 12, fontFace: "Arial", color: "94A3B8", align: "center" });

pres.writeFile({ fileName: "/Users/sam/Documents/MyProject/SPC_AI_Agent_專家.pptx" })
  .then(() => console.log("PPT created successfully!"))
  .catch(err => console.error("Error:", err));