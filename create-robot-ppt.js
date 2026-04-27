const pptxgen = require("pptxgenjs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.title = '工廠智慧巡檢機器人解決方案';
pres.author = 'Smart Manufacturing Team';

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

// SLIDE 1: Cover
let slide1 = pres.addSlide();
slide1.background = { path: basePath + "robot-cover.jpeg" };
slide1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: "0F172A", transparency: 60 } });
slide1.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.8, w: 9, h: 2.2, fill: { color: "1E3A5F", transparency: 30 } });
slide1.addText("工廠智慧巡檢機器人", { x: 0.5, y: 1.9, w: 9, h: 1, fontSize: 44, fontFace: "Arial Black", color: "FFFFFF", bold: true, align: "center" });
slide1.addText("AI LLM 賦能的新一代工業巡檢系統", { x: 0.5, y: 2.85, w: 9, h: 0.5, fontSize: 22, fontFace: "Arial", color: "3B82F6", align: "center" });
slide1.addText("方案設計與實踐報告", { x: 0.5, y: 3.4, w: 9, h: 0.4, fontSize: 16, fontFace: "Arial", color: "94A3B8", align: "center" });
slide1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.35, w: 10, h: 0.275, fill: { color: "3B82F6" } });
slide1.addText("2026年4月", { x: 0.5, y: 5.0, w: 9, h: 0.3, fontSize: 12, fontFace: "Arial", color: "94A3B8", align: "center" });

// SLIDE 2: Executive Summary
let slide2 = pres.addSlide();
slide2.background = { color: colors.dark };
slide2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: colors.primary } });
slide2.addText("執行摘要", { x: 0.5, y: 0.2, w: 9, h: 0.5, fontSize: 28, fontFace: "Arial Black", color: "FFFFFF", bold: true, margin: 0 });

const stats = [
  { value: "85%", label: "異常偵測率", color: "3B82F6" },
  { value: "60%", label: "人力成本節省", color: "10B981" },
  { value: "24/7", label: "全天候運行", color: "F59E0B" },
  { value: "<5s", label: "異常響應時間", color: "EF4444" }
];
stats.forEach((stat, i) => {
  const x = 0.5 + i * 2.4;
  slide2.addShape(pres.shapes.RECTANGLE, { x: x, y: 1.3, w: 2.2, h: 1.4, fill: { color: colors.card }, shadow: makeShadow() });
  slide2.addShape(pres.shapes.RECTANGLE, { x: x, y: 1.3, w: 2.2, h: 0.08, fill: { color: stat.color } });
  slide2.addText(stat.value, { x: x, y: 1.5, w: 2.2, h: 0.7, fontSize: 32, fontFace: "Arial Black", color: stat.color, bold: true, align: "center" });
  slide2.addText(stat.label, { x: x, y: 2.2, w: 2.2, h: 0.4, fontSize: 12, fontFace: "Arial", color: colors.muted, align: "center" });
});

const summaryPoints = [
  "整合 AI LLM 技術，實現智能化巡檢與自然語言交互能力",
  "採用多感測器融合架構，涵蓋視覺、溫度、振動、氣體檢測",
  "基於邊緣-雲端協同計算，確保即時性與大規模數據處理",
  "已完成 Prototype 驗證，預計導入後可提升巡檢效率 300%"
];
slide2.addText(summaryPoints.map((p, i) => ({ text: p, options: { bullet: true, breakLine: i < summaryPoints.length - 1 } })), { x: 0.5, y: 3.0, w: 9, h: 2.3, fontSize: 15, color: colors.light, paraSpaceAfter: 12 });

// SLIDE 3: Background
let slide3 = pres.addSlide();
slide3.background = { color: colors.dark };
slide3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: colors.primary } });
slide3.addText("計畫背景與問題陳述", { x: 0.5, y: 0.2, w: 9, h: 0.5, fontSize: 28, fontFace: "Arial Black", color: "FFFFFF", bold: true, margin: 0 });

slide3.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 4, fill: { color: colors.card }, shadow: makeShadow() });
slide3.addText("傳統巡檢痛點", { x: 0.7, y: 1.4, w: 3.9, h: 0.4, fontSize: 16, fontFace: "Arial", color: "EF4444", bold: true });
const painPoints = ["人力成本高：需專人 24 小時輪班", "人為誤差：疲勞導致漏檢", "反應緩慢：異常發現延遲數小時", "數據孤島：缺乏整合分析能力", "危險環境：人員安全風險高"];
slide3.addText(painPoints.map((p, i) => ({ text: p, options: { bullet: true, breakLine: i < painPoints.length - 1 } })), { x: 0.7, y: 1.9, w: 3.9, h: 3, fontSize: 13, color: colors.light, paraSpaceAfter: 10 });

slide3.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 4, fill: { color: colors.card }, shadow: makeShadow() });
slide3.addText("智慧化機會", { x: 5.4, y: 1.4, w: 3.9, h: 0.4, fontSize: 16, fontFace: "Arial", color: "10B981", bold: true });
const opportunities = ["AI LLM 突破：自然語言交互成熟", "感測器成本下降，性價比提升", "5G/邊緣運算普及，低延遲傳輸", "數位孿生技術整合", "企業數位轉型剛性需求"];
slide3.addText(opportunities.map((p, i) => ({ text: p, options: { bullet: true, breakLine: i < opportunities.length - 1 } })), { x: 5.4, y: 1.9, w: 3.9, h: 3, fontSize: 13, color: colors.light, paraSpaceAfter: 10 });

// SLIDE 4: AI LLM Technology
let slide4 = pres.addSlide();
slide4.background = { path: basePath + "robot-architecture.jpeg" };
slide4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: "0F172A", transparency: 70 } });
slide4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: colors.primary } });
slide4.addText("AI LLM 核心技術", { x: 0.5, y: 0.2, w: 9, h: 0.5, fontSize: 28, fontFace: "Arial Black", color: "FFFFFF", bold: true, margin: 0 });

const techCards = [
  { title: "自然語言理解", desc: "NLU 引擎理解維修人員口語化查詢，即時回應異常狀況與處置建議" },
  { title: "多模態感知", desc: "結合視覺辨識與感測器數據，AI 理解設備運行狀態的全貌" },
  { title: "知識圖譜增強", desc: "RAG 技術整合技術文檔、維修手冊、歷史案例，生成精準回答" },
  { title: "生成式診斷", desc: "依據異常模式自動生成診斷報告與維修建議，降低人為判斷失誤" }
];
techCards.forEach((card, i) => {
  const col = i % 2;
  const row = Math.floor(i / 2);
  const x = 0.5 + col * 4.7;
  const y = 1.2 + row * 2.0;
  slide4.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: 4.5, h: 1.8, fill: { color: "FFFFFF", transparency: 10 }, line: { color: "3B82F6", width: 1 } });
  slide4.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: 4.5, h: 0.06, fill: { color: "3B82F6" } });
  slide4.addText(card.title, { x: x + 0.2, y: y + 0.2, w: 4.1, h: 0.4, fontSize: 15, fontFace: "Arial", color: "FFFFFF", bold: true, margin: 0 });
  slide4.addText(card.desc, { x: x + 0.2, y: y + 0.65, w: 4.1, h: 1, fontSize: 11, fontFace: "Arial", color: "CBD5E1", margin: 0 });
});

// SLIDE 5: System Architecture
let slide5 = pres.addSlide();
slide5.background = { color: colors.dark };
slide5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: colors.primary } });
slide5.addText("系統架構設計", { x: 0.5, y: 0.2, w: 9, h: 0.5, fontSize: 28, fontFace: "Arial Black", color: "FFFFFF", bold: true, margin: 0 });

const layers = [
  { name: "應用層", desc: "Dashboard / 維修 App / 報警系統", color: "3B82F6", y: 1.1 },
  { name: "AI 服務層", desc: "LLM 推理引擎 / 知識圖譜 / 異常偵測模型", color: "8B5CF6", y: 2.0 },
  { name: "平台層", desc: "雲端計算 / 資料湖 / 數位孿生引擎", color: "10B981", y: 2.9 },
  { name: "網路層", desc: "5G專網 / Wi-Fi 6 / OPC UA", color: "F59E0B", y: 3.8 },
  { name: "設備層", desc: "巡檢機器人 / 感測器陣列 / 邊緣運算盒", color: "EF4444", y: 4.7 }
];
layers.forEach((layer) => {
  slide5.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: layer.y, w: 3.5, h: 0.7, fill: { color: layer.color }, shadow: makeShadow() });
  slide5.addText(layer.name, { x: 0.5, y: layer.y, w: 3.5, h: 0.7, fontSize: 14, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
  slide5.addShape(pres.shapes.RECTANGLE, { x: 4.2, y: layer.y, w: 5.3, h: 0.7, fill: { color: colors.card } });
  slide5.addText(layer.desc, { x: 4.4, y: layer.y, w: 5.1, h: 0.7, fontSize: 12, fontFace: "Arial", color: colors.light, valign: "middle", margin: 0 });
});
for (let i = 0; i < 4; i++) {
  slide5.addShape(pres.shapes.LINE, { x: 2.25, y: layers[i].y + 0.7, w: 0, h: 0.2, line: { color: colors.muted, width: 2 } });
}

// SLIDE 6: Prototype
let slide6 = pres.addSlide();
slide6.background = { path: basePath + "robot-prototype.jpeg" };
slide6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: "0F172A", transparency: 65 } });
slide6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: colors.primary } });
slide6.addText("Prototype 原型設計", { x: 0.5, y: 0.2, w: 9, h: 0.5, fontSize: 28, fontFace: "Arial Black", color: "FFFFFF", bold: true, margin: 0 });

const specs = [
  { title: "尺寸與移動", items: ["尺寸：600×400×800mm", "移動速度：0.5-2.0 m/s", "續航力：8 小時", "爬坡能力：15°"] },
  { title: "感測器陣列", items: ["雙目立體相機", "熱成像攝影機", "LIDAR 雷射雷達", "振動/溫度/氣體感測器"] },
  { title: "運算單元", items: ["NVIDIA Jetson Orin NX", "16TOPS AI 算力", "8GB LPDDR5", "Wi-Fi 6 / 5G 模組"] }
];
specs.forEach((spec, i) => {
  const x = 0.5 + i * 3.15;
  slide6.addShape(pres.shapes.RECTANGLE, { x: x, y: 1.2, w: 3.0, h: 4.1, fill: { color: "FFFFFF", transparency: 10 } });
  slide6.addShape(pres.shapes.RECTANGLE, { x: x, y: 1.2, w: 3.0, h: 0.06, fill: { color: "3B82F6" } });
  slide6.addText(spec.title, { x: x + 0.15, y: 1.4, w: 2.7, h: 0.4, fontSize: 14, fontFace: "Arial", color: "FFFFFF", bold: true, margin: 0 });
  slide6.addText(spec.items.map((item, idx) => ({ text: item, options: { bullet: true, breakLine: idx < spec.items.length - 1 } })), { x: x + 0.15, y: 1.9, w: 2.7, h: 3.2, fontSize: 11, color: "CBD5E1", paraSpaceAfter: 8 });
});

// SLIDE 7: Smart Features
let slide7 = pres.addSlide();
slide7.background = { color: colors.dark };
slide7.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: colors.primary } });
slide7.addText("智慧功能介紹", { x: 0.5, y: 0.2, w: 9, h: 0.5, fontSize: 28, fontFace: "Arial Black", color: "FFFFFF", bold: true, margin: 0 });

const features = [
  { title: "智能對話查詢", desc: "維修人員可用自然語言詢問設備狀態、異常處理方式，AI 即時回應", color: "3B82F6" },
  { title: "視覺異常偵測", desc: "即時分析設備外觀，發現裂縫、漏油、變色等異常，準確率達 95%", color: "10B981" },
  { title: "溫度熱點監控", desc: "紅外熱成像即時偵測溫度異常，提前發現設備過熱問題", color: "EF4444" },
  { title: "預測性維護", desc: "AI 分析振動趨勢，預判軸承、馬達等關鍵部件故障，提前 2-4 週預警", color: "F59E0B" },
  { title: "自動路徑規劃", desc: "SLAM 技術實現自主導航，動態避障，全天候自動化巡檢", color: "8B5CF6" },
  { title: "數據整合分析", desc: "整合 MES、ERP 數據，生成戰情儀表板，支援管理決策", color: "06B6D4" }
];
features.forEach((feat, i) => {
  const col = i % 3;
  const row = Math.floor(i / 3);
  const x = 0.5 + col * 3.1;
  const y = 1.2 + row * 2.1;
  slide7.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: 2.95, h: 1.9, fill: { color: colors.card }, shadow: makeShadow() });
  slide7.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: 0.08, h: 1.9, fill: { color: feat.color } });
  slide7.addText(feat.title, { x: x + 0.2, y: y + 0.15, w: 2.6, h: 0.4, fontSize: 13, fontFace: "Arial", color: "FFFFFF", bold: true, margin: 0 });
  slide7.addText(feat.desc, { x: x + 0.2, y: y + 0.6, w: 2.6, h: 1.2, fontSize: 10, fontFace: "Arial", color: colors.muted, margin: 0 });
});

// SLIDE 8: LLM Integration
let slide8 = pres.addSlide();
slide8.background = { color: colors.dark };
slide8.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: colors.primary } });
slide8.addText("LLM 系統整合架構", { x: 0.5, y: 0.2, w: 9, h: 0.5, fontSize: 28, fontFace: "Arial Black", color: "FFFFFF", bold: true, margin: 0 });

const flowSteps = [
  { text: "感測器數據收集", x: 0.5, color: "EF4444" },
  { text: "邊緣預處理", x: 2.4, color: "F59E0B" },
  { text: "雲端 LLM 推理", x: 4.3, color: "3B82F6" },
  { text: "知識庫檢索", x: 6.2, color: "10B981" },
  { text: "生成回應", x: 8.1, color: "8B5CF6" }
];
flowSteps.forEach((step, i) => {
  slide8.addShape(pres.shapes.RECTANGLE, { x: step.x, y: 1.3, w: 1.7, h: 0.9, fill: { color: step.color } });
  slide8.addText(step.text, { x: step.x, y: 1.3, w: 1.7, h: 0.9, fontSize: 10, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
  if (i < flowSteps.length - 1) slide8.addShape(pres.shapes.LINE, { x: step.x + 1.7, y: 1.75, w: 0.5, h: 0, line: { color: colors.muted, width: 2 } });
});

slide8.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.5, w: 9, h: 2.8, fill: { color: colors.card } });
slide8.addText("RAG 檢索增強生成架構", { x: 0.7, y: 2.65, w: 8.6, h: 0.4, fontSize: 14, fontFace: "Arial", color: "3B82F6", bold: true, margin: 0 });
const ragSteps = [
  "設備數據（OPC UA）→ 轉換為結構化向量",
  "技術文檔（PDF）→ 文件解析 → 分塊 → 向量化",
  "用戶查詢 → 意圖識別 → 向量檢索 → 上下文增強",
  "LLM 生成 → 事實校驗 → 輸出對話式回應"
];
slide8.addText(ragSteps.map((s, i) => ({ text: s, options: { bullet: true, breakLine: i < ragSteps.length - 1 } })), { x: 0.7, y: 3.1, w: 8.6, h: 2, fontSize: 12, color: colors.light, paraSpaceAfter: 10 });

// SLIDE 9: Dashboard
let slide9 = pres.addSlide();
slide9.background = { path: basePath + "robot-dashboard.jpeg" };
slide9.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: "0F172A", transparency: 70 } });
slide9.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: colors.primary } });
slide9.addText("智慧戰情中心", { x: 0.5, y: 0.2, w: 9, h: 0.5, fontSize: 28, fontFace: "Arial Black", color: "FFFFFF", bold: true, margin: 0 });

const metrics = [{ value: "23", label: "在線設備", unit: "台" }, { value: "156", label: "今日巡檢點", unit: "點" }, { value: "3", label: "待處理異常", unit: "件" }, { value: "99.2%", label: "系統可用率", unit: "" }];
metrics.forEach((m, i) => {
  const x = 0.5 + i * 2.4;
  slide9.addShape(pres.shapes.RECTANGLE, { x: x, y: 1.1, w: 2.2, h: 1.1, fill: { color: "FFFFFF", transparency: 15 } });
  slide9.addText(m.value + m.unit, { x: x, y: 1.15, w: 2.2, h: 0.6, fontSize: 24, fontFace: "Arial Black", color: "3B82F6", bold: true, align: "center" });
  slide9.addText(m.label, { x: x, y: 1.75, w: 2.2, h: 0.4, fontSize: 11, fontFace: "Arial", color: colors.muted, align: "center" });
});

const dashFeatures = ["即時設備狀態監控與異常警報", "3D 工廠數位孿生視圖", "AI 對話式查詢介面", "歷史數據分析與趨勢預測", "行動 App 遠程操控支援"];
slide9.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.4, w: 9, h: 2.9, fill: { color: "FFFFFF", transparency: 10 } });
slide9.addText("系統功能", { x: 0.7, y: 2.55, w: 8.6, h: 0.4, fontSize: 14, fontFace: "Arial", color: "10B981", bold: true, margin: 0 });
slide9.addText(dashFeatures.map((f, i) => ({ text: f, options: { bullet: true, breakLine: i < dashFeatures.length - 1 } })), { x: 0.7, y: 3.0, w: 8.6, h: 2.1, fontSize: 13, color: colors.light, paraSpaceAfter: 8 });

// SLIDE 10: Implementation Plan
let slide10 = pres.addSlide();
slide10.background = { color: colors.dark };
slide10.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: colors.primary } });
slide10.addText("建置時程規劃", { x: 0.5, y: 0.2, w: 9, h: 0.5, fontSize: 28, fontFace: "Arial Black", color: "FFFFFF", bold: true, margin: 0 });

const phases = [
  { phase: "第一階段", title: "PoC 驗證", period: "Month 1-2", items: ["硬體採購與組裝", "基礎感測器整合", "LLM API 串接", "基本異常偵測"] },
  { phase: "第二階段", title: "場域試驗", period: "Month 3-4", items: ["單一產線部署", "RAG 知識庫建置", "UI/Dashboard 開發", "使用者回饋優化"] },
  { phase: "第三階段", title: "規模化推廣", period: "Month 5-6", items: ["全廠區域部署", "系統整合優化", "AI 模型訓練", "教育訓練完成"] }
];
phases.forEach((p, i) => {
  const x = 0.5 + i * 3.15;
  slide10.addShape(pres.shapes.RECTANGLE, { x: x, y: 1.2, w: 3.0, h: 4.1, fill: { color: colors.card }, shadow: makeShadow() });
  slide10.addShape(pres.shapes.RECTANGLE, { x: x, y: 1.2, w: 3.0, h: 0.8, fill: { color: "3B82F6" } });
  slide10.addText(p.phase, { x: x, y: 1.25, w: 3.0, h: 0.35, fontSize: 11, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center" });
  slide10.addText(p.title, { x: x, y: 1.55, w: 3.0, h: 0.35, fontSize: 14, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center" });
  slide10.addText(p.period, { x: x, y: 2.1, w: 3.0, h: 0.35, fontSize: 11, fontFace: "Arial", color: "F59E0B", align: "center" });
  slide10.addText(p.items.map((item, idx) => ({ text: item, options: { bullet: true, breakLine: idx < p.items.length - 1 } })), { x: x + 0.15, y: 2.5, w: 2.7, h: 2.6, fontSize: 11, color: colors.light, paraSpaceAfter: 8 });
});

// SLIDE 11: Expected Results
let slide11 = pres.addSlide();
slide11.background = { color: colors.dark };
slide11.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: colors.primary } });
slide11.addText("預期成效", { x: 0.5, y: 0.2, w: 9, h: 0.5, fontSize: 28, fontFace: "Arial Black", color: "FFFFFF", bold: true, margin: 0 });

const comparisons = [
  { metric: "巡檢頻率", before: "每 4 小時一次", after: "每 15 分鐘一次", improvement: "16x" },
  { metric: "異常發現時間", before: "平均 4 小時", after: "即時偵測", improvement: "即時" },
  { metric: "人力需求", before: "3 人/班", after: "1 人/班", improvement: "-67%" },
  { metric: "維修響應時間", before: "2 小時", after: "30 分鐘", improvement: "-75%" },
  { metric: "非計劃停機", before: "12 次/月", after: "3 次/月", improvement: "-75%" }
];

slide11.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 9, h: 0.6, fill: { color: colors.card } });
const headers = ["指標", "導入前", "導入後", "改善幅度"];
const colWidths = [2.5, 2.5, 2.5, 1.5];
let xPos = 0.5;
headers.forEach((h, i) => {
  slide11.addText(h, { x: xPos, y: 1.2, w: colWidths[i], h: 0.6, fontSize: 12, fontFace: "Arial", color: "3B82F6", bold: true, align: "center", valign: "middle" });
  xPos += colWidths[i];
});

comparisons.forEach((row, idx) => {
  const y = 1.8 + idx * 0.65;
  slide11.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: y, w: 9, h: 0.65, fill: { color: idx % 2 === 0 ? colors.card : "293548" } });
  let x = 0.5;
  const values = [row.metric, row.before, row.after, row.improvement];
  values.forEach((v, i) => {
    slide11.addText(v, { x: x, y: y, w: colWidths[i], h: 0.65, fontSize: 11, fontFace: "Arial", color: i === 3 ? "10B981" : colors.light, bold: i === 3, align: "center", valign: "middle" });
    x += colWidths[i];
  });
});

slide11.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.1, w: 9, h: 0.4, fill: { color: "10B981", transparency: 80 } });
slide11.addText("預估投資報酬率：12 個月內達成 ROI > 150%", { x: 0.5, y: 5.1, w: 9, h: 0.4, fontSize: 13, fontFace: "Arial", color: "10B981", bold: true, align: "center", valign: "middle" });

// SLIDE 12: Budget
let slide12 = pres.addSlide();
slide12.background = { color: colors.dark };
slide12.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: colors.primary } });
slide12.addText("投資預算", { x: 0.5, y: 0.2, w: 9, h: 0.5, fontSize: 28, fontFace: "Arial Black", color: "FFFFFF", bold: true, margin: 0 });

const budgetItems = [
  { category: "硬體設備", amount: "NT$ 2,500,000", percent: "45%", items: ["巡檢機器人 (2台)", "感測器陣列", "邊緣運算設備", "網路基礎建設"] },
  { category: "系統開發", amount: "NT$ 1,500,000", percent: "27%", items: ["LLM 整合開發", "Dashboard 建置", "手機 App", "系統整合"] },
  { category: "知識庫建置", amount: "NT$ 500,000", percent: "9%", items: ["文件數位化", "向量資料庫", "RAG 系統開發"] },
  { category: "教育訓練", amount: "NT$ 300,000", percent: "5%", items: ["操作訓練", "維護訓練", "AI 模型訓練"] },
  { category: "其他費用", amount: "NT$ 700,000", percent: "14%", items: ["專案管理", "系統測試", " contingency"] }
];
budgetItems.forEach((item, i) => {
  const y = 1.1 + i * 0.85;
  slide12.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: y, w: 9, h: 0.75, fill: { color: colors.card } });
  slide12.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: y, w: 0.08, h: 0.75, fill: { color: "3B82F6" } });
  slide12.addText(item.category, { x: 0.7, y: y, w: 1.8, h: 0.75, fontSize: 12, fontFace: "Arial", color: "FFFFFF", bold: true, valign: "middle", margin: 0 });
  slide12.addText(item.amount, { x: 2.5, y: y, w: 2, h: 0.75, fontSize: 14, fontFace: "Arial", color: "3B82F6", bold: true, valign: "middle", margin: 0 });
  slide12.addText(item.percent, { x: 4.5, y: y, w: 0.8, h: 0.75, fontSize: 12, fontFace: "Arial", color: "F59E0B", valign: "middle", margin: 0 });
  slide12.addText(item.items.join(" / "), { x: 5.4, y: y, w: 4, h: 0.75, fontSize: 10, fontFace: "Arial", color: colors.muted, valign: "middle", margin: 0 });
});

slide12.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.1, w: 9, h: 0.45, fill: { color: "3B82F6" } });
slide12.addText("總預算：NT$ 5,500,000", { x: 0.5, y: 5.1, w: 9, h: 0.45, fontSize: 16, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle" });

// SLIDE 13: Team
let slide13 = pres.addSlide();
slide13.background = { color: colors.dark };
slide13.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: colors.primary } });
slide13.addText("專案團隊", { x: 0.5, y: 0.2, w: 9, h: 0.5, fontSize: 28, fontFace: "Arial Black", color: "FFFFFF", bold: true, margin: 0 });

const teamMembers = [
  { role: "專案經理", name: "待定", tasks: "專案統籌 / 時程管控 / 資源協調" },
  { role: "AI/ML 工程師", name: "待定", tasks: "LLM 整合 / 模型訓練 / RAG 系統" },
  { role: "機器人軟體", name: "待定", tasks: "SLAM / 導航 / 感測器整合" },
  { role: "前端開發", name: "待定", tasks: "Dashboard / Mobile App" },
  { role: "MES/IT 整合", name: "待定", tasks: "OPC UA / API 串接" },
  { role: "設備維護", name: "待定", tasks: "硬體維護 / 異常處理" }
];
teamMembers.forEach((member, i) => {
  const col = i % 2;
  const row = Math.floor(i / 2);
  const x = 0.5 + col * 4.7;
  const y = 1.2 + row * 1.4;
  slide13.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: 4.5, h: 1.2, fill: { color: colors.card }, shadow: makeShadow() });
  slide13.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: 0.08, h: 1.2, fill: { color: "3B82F6" } });
  slide13.addText(member.role, { x: x + 0.2, y: y + 0.1, w: 2, h: 0.35, fontSize: 11, fontFace: "Arial", color: colors.muted, margin: 0 });
  slide13.addText(member.name, { x: x + 0.2, y: y + 0.4, w: 2, h: 0.35, fontSize: 14, fontFace: "Arial", color: "FFFFFF", bold: true, margin: 0 });
  slide13.addText(member.tasks, { x: x + 0.2, y: y + 0.8, w: 4.1, h: 0.35, fontSize: 10, fontFace: "Arial", color: colors.muted, margin: 0 });
});

// SLIDE 14: Risks
let slide14 = pres.addSlide();
slide14.background = { color: colors.dark };
slide14.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: colors.primary } });
slide14.addText("風險評估與對策", { x: 0.5, y: 0.2, w: 9, h: 0.5, fontSize: 28, fontFace: "Arial Black", color: "FFFFFF", bold: true, margin: 0 });

const risks = [
  { risk: "工廠環境複雜，機器人導航失效", level: "高", mitigation: "建立多重定位備援、SLAM + 視覺糾偏", color: "EF4444" },
  { risk: "LLM 幻觉產生錯誤診斷建議", level: "中", mitigation: "事實校驗機制、專家審核流程", color: "F59E0B" },
  { risk: "網路延遲影響即時性", level: "中", mitigation: "邊緣運算優先、離線緩存機制", color: "F59E0B" },
  { risk: "感測器精度不足", level: "低", mitigation: "多源感測器融合、AI 異常偵測補償", color: "10B981" },
  { risk: "人員接受度低", level: "低", mitigation: "充分教育訓練、階段性導入", color: "10B981" }
];

slide14.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 9, h: 0.5, fill: { color: colors.card } });
const riskHeaders = ["風險項目", "等級", "緩解對策"];
const riskColWidths = [3.5, 0.8, 4.7];
let riskX = 0.5;
riskHeaders.forEach((h, i) => {
  slide14.addText(h, { x: riskX, y: 1.2, w: riskColWidths[i], h: 0.5, fontSize: 11, fontFace: "Arial", color: "3B82F6", bold: true, align: "center", valign: "middle" });
  riskX += riskColWidths[i];
});

risks.forEach((r, idx) => {
  const y = 1.7 + idx * 0.7;
  slide14.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: y, w: 9, h: 0.7, fill: { color: idx % 2 === 0 ? colors.card : "293548" } });
  let rx = 0.5;
  const values = [r.risk, r.level, r.mitigation];
  values.forEach((v, i) => {
    slide14.addText(v, { x: rx + 0.1, y: y, w: riskColWidths[i] - 0.2, h: 0.7, fontSize: 10, fontFace: "Arial", color: i === 1 ? r.color : colors.light, bold: i === 1, valign: "middle" });
    rx += riskColWidths[i];
  });
});

// SLIDE 15: Conclusion
let slide15 = pres.addSlide();
slide15.background = { path: basePath + "robot-cover.jpeg" };
slide15.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: "0F172A", transparency: 65 } });
slide15.addText("結語", { x: 0.5, y: 1.0, w: 9, h: 0.6, fontSize: 36, fontFace: "Arial Black", color: "FFFFFF", bold: true, align: "center" });

const conclusions = [
  "AI LLM 技術為工廠巡檢帶來革命性突破",
  "人機協作模式大幅提升巡檢效率與安全性",
  "預測性維護能力顯著降低非計劃停機",
  "投資報酬率高，12 個月內可回收成本"
];
slide15.addText(conclusions.map((c, i) => ({ text: c, options: { bullet: true, breakLine: i < conclusions.length - 1 } })), { x: 1.5, y: 1.8, w: 7, h: 2.2, fontSize: 16, color: colors.light, paraSpaceAfter: 14, align: "left" });

slide15.addShape(pres.shapes.RECTANGLE, { x: 2, y: 4.2, w: 6, h: 0.9, fill: { color: "3B82F6" } });
slide15.addText("歡迎諮詢，共同打造智慧工廠新未來", { x: 2, y: 4.2, w: 6, h: 0.9, fontSize: 18, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
slide15.addText("2026年4月", { x: 0.5, y: 5.2, w: 9, h: 0.3, fontSize: 12, fontFace: "Arial", color: colors.muted, align: "center" });

pres.writeFile({ fileName: "/Users/sam/Documents/MyProject/工廠智慧巡檢機器人.pptx" })
  .then(() => console.log("PPT created successfully!"))
  .catch(err => console.error(err));
