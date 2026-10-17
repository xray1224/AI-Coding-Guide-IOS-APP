const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const {
  FaUserTie, FaCode, FaGraduationCap, FaLaptopCode,
  FaClipboardList, FaGavel, FaQuestionCircle, FaLightbulb, FaListUl,
  FaFileContract, FaCheckDouble, FaDraftingCompass, FaSearchPlus, FaRobot, FaUserCheck
} = require("react-icons/fa");

// ============================================================
// DIMENSIONS
// ============================================================
const pres = new pptxgen();
pres.author = "fingalyang";
pres.title = "AI Coding 保姆级指引（iOS 移动应用）";
pres.layout = "LAYOUT_16x9";
const SLIDE_W = 10, SLIDE_H = 5.625;
const MARGIN = 0.5;
const CONTENT_X = MARGIN, CONTENT_Y = MARGIN;
const CONTENT_W = SLIDE_W - 2 * MARGIN, CONTENT_H = SLIDE_H - 2 * MARGIN;

// ============================================================
// CONTAINER SYSTEM
// ============================================================
function calculateScaledImageOpts(opts) {
  const { path, w: targetW, h: targetH, x = 0, y = 0, mode = "cover", ...rest } = opts;
  if (!path || !targetW || !targetH) return opts;
  return { path, x, y, w: targetW, h: targetH, sizing: { type: mode, w: targetW, h: targetH }, ...rest };
}
function createVirtualNode(type, data, parentX = 0, parentY = 0) {
  const opts = data.opts || {};
  const node = { type, data, absX: parentX + (opts.x || 0), absY: parentY + (opts.y || 0), w: opts.w || 0, h: opts.h || 0, children: [] };
  node.addShape = function (shapeType, opts = {}) { const c = createVirtualNode("shape", { shapeType, opts }, node.absX, node.absY); node.children.push(c); return c; };
  node.addText = function (text, opts = {}) {
    const safeOpts = { fit: "shrink", ...opts };
    const bulletRe = /^(?:[•‣◦⁃∙·●○–—]\s*|-\s+)/;
    if (Array.isArray(text)) text = text.map(i => (i && i.options && i.options.bullet && typeof i.text === "string") ? { ...i, text: i.text.replace(bulletRe, "") } : i);
    const c = createVirtualNode("text", { text, opts: safeOpts }, node.absX, node.absY); node.children.push(c); return c;
  };
  node.addImage = function (opts = {}) { const c = createVirtualNode("image", { opts: calculateScaledImageOpts(opts) }, node.absX, node.absY); node.children.push(c); return c; };
  node.addTable = function (tableData, opts = {}) { const c = createVirtualNode("table", { tableData, opts }, node.absX, node.absY); node.children.push(c); return c; };
  return node;
}
function flattenNode(node, realSlide) {
  const absOpts = { ...node.data.opts, x: node.absX, y: node.absY };
  if (node.type === "shape") realSlide.addShape(node.data.shapeType, absOpts);
  else if (node.type === "text") realSlide.addText(node.data.text, absOpts);
  else if (node.type === "image") realSlide.addImage(absOpts);
  else if (node.type === "table") realSlide.addTable(node.data.tableData, absOpts);
  node.children.forEach(c => flattenNode(c, realSlide));
}
const originalAddSlide = pres.addSlide.bind(pres);
pres.addSlide = function (options) {
  const realSlide = originalAddSlide(options);
  return {
    children: [], _realSlide: realSlide,
    set background(val) { realSlide.background = val; },
    get background() { return realSlide.background; },
    addShape(t, o = {}) { const n = createVirtualNode("shape", { shapeType: t, opts: o }, 0, 0); this.children.push(n); return n; },
    addText(t, o = {}) { const n = createVirtualNode("text", { text: t, opts: { fit: "shrink", ...o } }, 0, 0); this.children.push(n); return n; },
    addImage(o = {}) { const n = createVirtualNode("image", { opts: calculateScaledImageOpts(o) }, 0, 0); this.children.push(n); return n; },
    addTable(d, o = {}) { const n = createVirtualNode("table", { tableData: d, opts: o }, 0, 0); this.children.push(n); return n; },
    addChart(t, d, o = {}) { realSlide.addChart(t, d, o); },
    render() { this.children.forEach(c => flattenNode(c, realSlide)); }
  };
};

// ============================================================
// THEME
// ============================================================
const PRIMARY = "065A82";   // deep blue
const SECONDARY = "1C7293"; // teal
const DARK = "21295C";      // midnight
const BG = "F7FAFC";
const TEXT = "1F2937";
const MUTED = "64748B";
const LIGHT_LINE = "E2E8F0";
const CARD_TEAL_BG = "EAF3F7";
const SERIF = "Songti SC";
const SANS = "PingFang SC";
const IMG = "images/";

const makeShadow = () => ({ type: "outer", blur: 12, offset: 4, angle: 135, color: "000000", opacity: 0.18 });
const makeSoftShadow = () => ({ type: "outer", blur: 10, offset: 3, angle: 135, color: "065A82", opacity: 0.12 });

function renderIconSvg(Comp, color = "#FFFFFF", size = 256) {
  return ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color, size: String(size) }));
}
async function iconPng(Comp, color) {
  const svg = renderIconSvg(Comp, color, 256);
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

function addHeader(slide, kicker, title) {
  slide.addText(kicker, { x: CONTENT_X, y: 0.36, w: CONTENT_W, h: 0.26, fontSize: 10.5, bold: true, color: SECONDARY, charSpacing: 2, fontFace: SANS, margin: 0 });
  slide.addText(title, { x: CONTENT_X, y: 0.62, w: CONTENT_W, h: 0.52, fontSize: 25, bold: true, color: DARK, fontFace: SERIF, charSpacing: 1.5, margin: 0 });
}
function addFooter(slide, n) {
  slide.addText("AI Coding 保姆级指引（iOS 移动应用）", { x: CONTENT_X, y: 5.3, w: 4.5, h: 0.2, fontSize: 8.5, color: MUTED, fontFace: SANS, margin: 0 });
  slide.addText(String(n).padStart(2, "0") + " / 16", { x: SLIDE_W - 1.6, y: 5.3, w: 1.1, h: 0.2, fontSize: 8.5, color: MUTED, fontFace: SANS, align: "right", margin: 0 });
}
function addCard(slide, x, y, w, h, fill = "FFFFFF") {
  return slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, rectRadius: 0.1, line: { color: LIGHT_LINE, width: 0.75 }, shadow: makeSoftShadow() });
}
function addChip(slide, x, y, w, h, text, opts = {}) {
  const { fill = CARD_TEAL_BG, color = PRIMARY, fontSize = 10, bold = true, lineColor = null, transparency = 0 } = opts;
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill, transparency }, rectRadius: h / 2, line: lineColor ? { color: lineColor, width: 1 } : { type: "none" } });
  slide.addText(text, { x, y, w, h, align: "center", valign: "middle", fontSize, bold, color, fontFace: SANS, margin: 0 });
}
async function addIconCircle(node, x, y, d, iconData, circleColor = PRIMARY) {
  node.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: circleColor } });
  const pad = d * 0.26;
  node.addImage({ data: iconData, x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad });
}

(async () => {
  // Pre-render icons
  const icons = {
    userTie: await iconPng(FaUserTie, "#FFFFFF"),
    code: await iconPng(FaCode, "#FFFFFF"),
    grad: await iconPng(FaGraduationCap, "#FFFFFF"),
    laptop: await iconPng(FaLaptopCode, "#FFFFFF"),
    plan: await iconPng(FaClipboardList, "#FFFFFF"),
    gavel: await iconPng(FaGavel, "#FFFFFF"),
    question: await iconPng(FaQuestionCircle, "#FFFFFF"),
    bulb: await iconPng(FaLightbulb, "#FFFFFF"),
    list: await iconPng(FaListUl, "#FFFFFF"),
    contract: await iconPng(FaFileContract, "#FFFFFF"),
    check: await iconPng(FaCheckDouble, "#FFFFFF"),
    compass: await iconPng(FaDraftingCompass, "#FFFFFF"),
    search: await iconPng(FaSearchPlus, "#FFFFFF"),
    robot: await iconPng(FaRobot, "#FFFFFF"),
    userCheck: await iconPng(FaUserCheck, "#FFFFFF"),
  };

  // ============================================================
  // S1 COVER
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { path: IMG + "bg-cover_16x9.jpg" };
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, fill: { color: "111827", transparency: 42 } });
    // kicker chip
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.62, y: 0.58, w: 1.55, h: 0.38, fill: { color: "1C7293", transparency: 45 }, rectRadius: 0.19, line: { color: "9FD8E8", width: 1 } });
    s.addText("开源项目分享", { x: 0.62, y: 0.58, w: 1.55, h: 0.38, align: "center", valign: "middle", fontSize: 11, bold: true, color: "D6EEF6", fontFace: SANS, charSpacing: 1, margin: 0 });
    // title
    s.addText("AI Coding 保姆级指引", { x: 0.6, y: 1.72, w: 8.8, h: 0.85, fontSize: 38, bold: true, color: "FFFFFF", fontFace: SERIF, charSpacing: 2.5, margin: 0 });
    s.addText("（iOS 移动应用）", { x: 0.6, y: 2.62, w: 8.8, h: 0.55, fontSize: 22, bold: true, color: "9FD8E8", fontFace: SERIF, charSpacing: 1.5, margin: 0 });
    s.addText("没有移动开发经验，也能把一款 App 送上 App Store", { x: 0.6, y: 3.5, w: 8.8, h: 0.4, fontSize: 15, color: "E6F0F5", fontFace: SANS, margin: 0 });
    // meta chips
    const chips = [["案例：份子钱银行", 1.95], ["业余时间约 3 个月", 1.95], ["Cursor 协作约 103 小时", 2.35]];
    let cx = 0.62;
    chips.forEach(([t, w]) => {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx, y: 4.32, w, h: 0.42, fill: { color: "1C7293", transparency: 45 }, rectRadius: 0.21, line: { color: "9FD8E8", width: 0.75, transparency: 40 } });
      s.addText(t, { x: cx, y: 4.32, w, h: 0.42, align: "center", valign: "middle", fontSize: 10.5, color: "FFFFFF", fontFace: SANS, margin: 0 });
      cx += w + 0.22;
    });
    s.addText("分享人：fingalyang", { x: 0.62, y: 5.08, w: 4, h: 0.28, fontSize: 10, color: "8FA3C8", fontFace: SANS, margin: 0 });
    s.render();
  }

  // ============================================================
  // S2 TOC
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: BG };
    addHeader(s, "CONTENTS", "目录");
    const items = [
      ["01", "背景与成果", "从零经验到上架 App Store"],
      ["02", "核心原则", "100% AI Coding 的人机分工"],
      ["03", "方法与技巧", "协作方法、实用技巧与问题对策"],
      ["04", "实操时间线", "三个月里的每一轮循环"],
      ["05", "量化分析", "工时、返工率与缺陷数据"],
      ["06", "常见问题", "钱、时间与适用范围"],
    ];
    const colW = 4.3, rowH = 1.02, gapX = 0.4, gapY = 0.18, startY = 1.55;
    items.forEach(([num, t, d], i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const x = CONTENT_X + col * (colW + gapX), y = startY + row * (rowH + gapY);
      const card = addCard(s, x, y, colW, rowH);
      card.addText(num, { x: 0.28, y: 0.18, w: 0.85, h: 0.66, fontSize: 26, bold: true, color: SECONDARY, fontFace: SERIF, margin: 0 });
      card.addText(t, { x: 1.15, y: 0.16, w: colW - 1.35, h: 0.36, fontSize: 15.5, bold: true, color: DARK, fontFace: SERIF, charSpacing: 1, margin: 0 });
      card.addText(d, { x: 1.15, y: 0.55, w: colW - 1.35, h: 0.32, fontSize: 10.5, color: MUTED, fontFace: SANS, margin: 0 });
    });
    addFooter(s, 2);
    s.render();
  }

  // ============================================================
  // S3 背景
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: BG };
    addHeader(s, "01 · 背景与成果", "为什么做这次实践");
    const rows = [
      [icons.userTie, "背景", "网络工程专业（2013 年毕业），长期从事信息化建设与网络信息安全管理，持有 CISSP / CISP / CDGA 证书"],
      [icons.code, "语言", "熟悉 Python、C、JavaScript、PHP——但没有大型软件开发经验，也没有移动应用开发经验"],
      [icons.grad, "路径", "先完成瑞通 AI Coding 课程，再立即动手验证；实践选择 iOS 应用开发"],
      [icons.laptop, "环境", "macOS + Xcode + Cursor；具体语言与框架不事先指定，由 Cursor 按需配置"],
    ];
    const card = addCard(s, CONTENT_X, 1.45, 5.05, 3.62);
    rows.forEach(([icon, t, d], i) => {
      const y = 0.24 + i * 0.85;
      addIconCircle(card, 0.26, y + 0.04, 0.44, icon, i % 2 === 0 ? PRIMARY : SECONDARY);
      card.addText(t, { x: 0.86, y: y - 0.02, w: 3.95, h: 0.28, fontSize: 12.5, bold: true, color: DARK, fontFace: SANS, margin: 0 });
      card.addText(d, { x: 0.86, y: y + 0.27, w: 3.95, h: 0.52, fontSize: 10, color: MUTED, fontFace: SANS, margin: 0, autoFit: false, fit: "none" });
    });
    s.addImage({ path: IMG + "image-background_4x3.jpg", x: 5.85, y: 1.85, w: 3.65, h: 2.74 });
    s.addText("先学会与 AI 协作，再让 AI 干活", { x: 5.85, y: 4.72, w: 3.65, h: 0.3, align: "center", fontSize: 10, italic: true, color: MUTED, fontFace: SANS, margin: 0 });
    addFooter(s, 3);
    s.render();
  }

  // ============================================================
  // S4 成果
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: BG };
    addHeader(s, "01 · 背景与成果", "实践成果：份子钱银行");
    // screenshot card
    const shot = addCard(s, 0.55, 1.35, 2.1, 4.05);
    shot.addImage({ path: IMG + "image-app-main.png", x: 0.11, y: 0.1, w: 1.88, h: 3.85, mode: "contain" });
    // right content
    s.addText("份子钱银行", { x: 2.95, y: 1.38, w: 4, h: 0.45, fontSize: 21, bold: true, color: DARK, fontFace: SERIF, charSpacing: 1.5, margin: 0 });
    s.addText("一款本地记账应用，记录每一份人情往来", { x: 2.95, y: 1.86, w: 6, h: 0.28, fontSize: 11, color: MUTED, fontFace: SANS, margin: 0 });
    const chipData = [["已上架 App Store", 1.72], ["支持应用内购买", 1.72], ["v1.0.2 已发布", 1.55]];
    let cx = 2.95;
    chipData.forEach(([t, w]) => { addChip(s, cx, 2.3, w, 0.38, t, { fontSize: 10 }); cx += w + 0.18; });
    s.addText("应用功能", { x: 2.95, y: 2.94, w: 3, h: 0.28, fontSize: 12, bold: true, color: TEXT, fontFace: SANS, margin: 0 });
    const feats = ["录入", "查询", "排行", "当前余额", "未来余额", "备份", "设置"];
    feats.forEach((f, i) => {
      const col = i % 4, row = Math.floor(i / 4);
      addChip(s, 2.95 + col * 1.12, 3.26 + row * 0.48, 1.0, 0.36, f, { fill: "FFFFFF", color: TEXT, fontSize: 10, bold: false, lineColor: LIGHT_LINE });
    });
    // mini stats
    const stats = [["约 3 个月", "业余时间完成"], ["103 小时", "Cursor 协作"], ["31 次", "代码提交"]];
    stats.forEach(([n, l], i) => {
      const x = 2.95 + i * 1.75;
      s.addText(n, { x, y: 4.32, w: 1.65, h: 0.32, fontSize: 16, bold: true, color: PRIMARY, fontFace: SERIF, margin: 0 });
      s.addText(l, { x, y: 4.66, w: 1.65, h: 0.24, fontSize: 9.5, color: MUTED, fontFace: SANS, margin: 0 });
    });
    s.addText("上架后同步推进推广与软件著作权申请，并按第一批用户反馈发布 1.0.2。", { x: 2.95, y: 5.0, w: 6.5, h: 0.26, fontSize: 9.5, color: MUTED, fontFace: SANS, margin: 0 });
    addFooter(s, 4);
    s.render();
  }

  // ============================================================
  // S5 核心结论 (divider)
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { path: IMG + "bg-divider_16x9.jpg" };
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, fill: { color: "111827", transparency: 35 } });
    s.addText("核心结论", { x: 0.9, y: 1.25, w: 6, h: 0.3, fontSize: 12, bold: true, color: "9FD8E8", charSpacing: 2, fontFace: SANS, margin: 0 });
    s.addText("“人工智能目前可以完成一个复杂工程。”", { x: 0.9, y: 1.95, w: 8.4, h: 1.1, fontSize: 27, bold: true, color: "FFFFFF", fontFace: SERIF, charSpacing: 1.5, margin: 0 });
    s.addText("人负责目标、取舍和验收，AI 负责执行。", { x: 0.9, y: 3.35, w: 8.4, h: 0.45, fontSize: 14.5, color: "CADCFC", fontFace: SANS, margin: 0 });
    s.render();
  }

  // ============================================================
  // S6 核心原则
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: BG };
    addHeader(s, "02 · 核心原则", "100% AI Coding：人与 AI 的分工");
    const c1 = addCard(s, CONTENT_X, 1.45, 5.15, 1.72);
    addIconCircle(c1, 0.24, 0.22, 0.44, icons.robot, PRIMARY);
    c1.addText("AI：执行者与建议者", { x: 0.82, y: 0.26, w: 4.1, h: 0.32, fontSize: 13.5, bold: true, color: DARK, fontFace: SANS, margin: 0 });
    c1.addText("代码、方案、文案、软著草稿、推广稿，尽最大可能由 AI 撰写——使用 AI 越充分，最终收益越大。", { x: 0.82, y: 0.68, w: 4.1, h: 0.85, fontSize: 10.5, color: MUTED, fontFace: SANS, margin: 0, autoFit: false, fit: "none" });
    const c2 = addCard(s, CONTENT_X, 3.33, 5.15, 1.74);
    addIconCircle(c2, 0.24, 0.22, 0.44, icons.userCheck, SECONDARY);
    c2.addText("人：产品负责人与验收者", { x: 0.82, y: 0.24, w: 4.1, h: 0.32, fontSize: 13.5, bold: true, color: DARK, fontFace: SANS, margin: 0 });
    c2.addText([
      { text: "确定目标、做取舍、验收结果", options: { bullet: { code: "2022", indent: 8 }, breakLine: true } },
      { text: "申请开发者账号，处理协议与税务", options: { bullet: { code: "2022", indent: 8 }, breakLine: true } },
      { text: "真机测试与审核回复", options: { bullet: { code: "2022", indent: 8 }, breakLine: true } },
      { text: "提交软件著作权 · 推广投稿", options: { bullet: { code: "2022", indent: 8 } } },
    ], { x: 0.82, y: 0.62, w: 4.1, h: 1.05, fontSize: 10, color: TEXT, fontFace: SANS, margin: 0, paraSpaceAfter: 3 });
    s.addImage({ path: IMG + "image-principle_4x3.jpg", x: 5.95, y: 1.45, w: 3.55, h: 2.66 });
    const q = addCard(s, 5.95, 4.3, 3.55, 0.78, CARD_TEAL_BG);
    q.addText("“人负责把要求说清楚，并做出判断。”", { x: 0.15, y: 0.1, w: 3.25, h: 0.58, align: "center", valign: "middle", fontSize: 10.5, italic: true, bold: true, color: PRIMARY, fontFace: SANS, margin: 0 });
    addFooter(s, 6);
    s.render();
  }

  // ============================================================
  // S7 协作方法
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: BG };
    addHeader(s, "03 · 方法与技巧", "协作方法：五句话的分工");
    const rows = [
      [icons.plan, "人规划，AI 执行"],
      [icons.gavel, "人决定，AI 遵守"],
      [icons.question, "人提问，AI 回答"],
      [icons.bulb, "AI 提出建议，人做决策"],
      [icons.list, "AI 列举选项，人做取舍"],
    ];
    rows.forEach(([icon, t], i) => {
      const y = 1.45 + i * 0.62;
      const card = addCard(s, CONTENT_X, y, 5.15, 0.52);
      addIconCircle(card, 0.14, 0.06, 0.4, icon, i % 2 === 0 ? PRIMARY : SECONDARY);
      card.addText(t, { x: 0.68, y: 0, w: 4.3, h: 0.52, valign: "middle", fontSize: 12.5, bold: true, color: TEXT, fontFace: SANS, margin: 0 });
    });
    const callout = addCard(s, CONTENT_X, 4.62, 5.15, 0.52, CARD_TEAL_BG);
    callout.addText("每一轮：先说明要什么 → AI 出方案 → 人确认后再动手", { x: 0.15, y: 0, w: 4.85, h: 0.52, valign: "middle", fontSize: 10.5, bold: true, color: PRIMARY, fontFace: SANS, margin: 0 });
    s.addImage({ path: IMG + "image-collab_4x3.jpg", x: 5.95, y: 1.7, w: 3.55, h: 2.66 });
    s.addText("人提要求，AI 干活；AI 提建议，人做决定", { x: 5.95, y: 4.5, w: 3.55, h: 0.3, align: "center", fontSize: 10, italic: true, color: MUTED, fontFace: SANS, margin: 0 });
    addFooter(s, 7);
    s.render();
  }

  // ============================================================
  // S8 重要技巧
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: BG };
    addHeader(s, "03 · 方法与技巧", "四个重要技巧");
    const cards = [
      [icons.contract, "提前用 Rules 约定规则", "编码规范、界面风格、兼容性与文风，提前写进每次对话 AI 都能看见的 Rules；功能完成后再补，上架前会集中返工。"],
      [icons.check, "把产品行为写到可以验收", "页面如何划分、点击何处出现什么、哪些字段必填，都写到能够验收；具体实现交给 AI。"],
      [icons.compass, "设计阶段让 AI 给出方案", "总体方向由人把握；尚未想清楚的交互，让 AI 先给几种做法并比较优缺点，再由人选择。"],
      [icons.search, "让 AI 对照代码和对话复查", "一段功能完成后复查；收尾与上架前收紧任务，只查找重大缺陷，不再增加新功能。"],
    ];
    const cw = 4.375, ch = 1.78, gx = 0.25, gy = 0.2, sy = 1.45;
    cards.forEach(([icon, t, d], i) => {
      const x = CONTENT_X + (i % 2) * (cw + gx), y = sy + Math.floor(i / 2) * (ch + gy);
      const card = addCard(s, x, y, cw, ch);
      addIconCircle(card, 0.24, 0.2, 0.44, icon, i % 2 === 0 ? PRIMARY : SECONDARY);
      card.addText(t, { x: 0.82, y: 0.26, w: cw - 1.05, h: 0.32, fontSize: 13, bold: true, color: DARK, fontFace: SANS, margin: 0 });
      card.addText(d, { x: 0.26, y: 0.78, w: cw - 0.52, h: 0.88, fontSize: 10, color: MUTED, fontFace: SANS, margin: 0, autoFit: false, fit: "none" });
    });
    addFooter(s, 8);
    s.render();
  }

  // ============================================================
  // S9 问题与对策
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: BG };
    addHeader(s, "03 · 方法与技巧", "问题与对策：八条实战经验");
    const items = [
      ["人负责产品行为，实现交给 AI", "下达可验收的指令，不介入具体实现"],
      ["总体想清楚，细节边做边定", "方向正确即可推进，不被单页细节挡住"],
      ["商店规则会改写商业设想", "买断改为免费下载 + 内购，提前确认规则"],
      ["用截图指出问题", "界面问题直接截图标注，AI 更容易定位"],
      ["同一功能反复修不好时重写", "绕进错误路径时，重写比继续修补更有效"],
      ["AI 宣称完成后，仍由人验收", "每轮把刚完成的路径点一遍，上架前做首次安装测试"],
      ["复用现成能力与已有界面", "内购用官方 StoreKit，导出页复用查询页"],
      ["按任务选择模型", "编码用 Cursor，图标交给视觉生成工具"],
    ];
    const cw = 4.375, ch = 0.8, gx = 0.25, gy = 0.17, sy = 1.42;
    items.forEach(([t, d], i) => {
      const x = CONTENT_X + (i % 2) * (cw + gx), y = sy + Math.floor(i / 2) * (ch + gy);
      const card = addCard(s, x, y, cw, ch);
      card.addShape(pres.shapes.OVAL, { x: 0.16, y: 0.22, w: 0.36, h: 0.36, fill: { color: i % 2 === 0 ? PRIMARY : SECONDARY } });
      card.addText(String(i + 1), { x: 0.16, y: 0.22, w: 0.36, h: 0.36, align: "center", valign: "middle", fontSize: 12, bold: true, color: "FFFFFF", fontFace: SERIF, margin: 0 });
      card.addText(t, { x: 0.66, y: 0.08, w: cw - 0.85, h: 0.3, fontSize: 11, bold: true, color: DARK, fontFace: SANS, margin: 0 });
      card.addText(d, { x: 0.66, y: 0.4, w: cw - 0.85, h: 0.32, fontSize: 9.5, color: MUTED, fontFace: SANS, margin: 0 });
    });
    addFooter(s, 9);
    s.render();
  }

  // ============================================================
  // S10 标准循环
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: BG };
    addHeader(s, "03 · 方法与技巧", "标准循环：每一轮都走同一条路");
    const steps = ["一轮一事", "AI 出方案", "人确认", "AI 实现", "人验收", "AI 复查", "提交 commit"];
    const n = steps.length, nw = 1.08, gap = (CONTENT_W - n * nw) / (n - 1), ny = 2.0, nh = 0.66;
    steps.forEach((t, i) => {
      const x = CONTENT_X + i * (nw + gap);
      s.addShape(pres.shapes.OVAL, { x: x + nw / 2 - 0.17, y: 1.5, w: 0.34, h: 0.34, fill: { color: "FFFFFF" }, line: { color: i % 2 === 0 ? PRIMARY : SECONDARY, width: 1.5 } });
      s.addText(String(i + 1), { x: x + nw / 2 - 0.17, y: 1.5, w: 0.34, h: 0.34, align: "center", valign: "middle", fontSize: 12, bold: true, color: i % 2 === 0 ? PRIMARY : SECONDARY, fontFace: SERIF, margin: 0 });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: ny, w: nw, h: nh, fill: { color: i % 2 === 0 ? PRIMARY : SECONDARY }, rectRadius: 0.12, shadow: makeShadow() });
      s.addText(t, { x, y: ny, w: nw, h: nh, align: "center", valign: "middle", fontSize: 10.5, bold: true, color: "FFFFFF", fontFace: SANS, margin: 0 });
      if (i < n - 1) {
        s.addShape(pres.shapes.LINE, { x: x + nw + 0.03, y: ny + nh / 2, w: gap - 0.06, h: 0, line: { color: "9AA5B1", width: 1.5, endArrowType: "triangle" } });
      }
    });
    const note = addCard(s, CONTENT_X, 3.35, CONTENT_W, 1.55);
    note.addText("为什么需要标准循环", { x: 0.3, y: 0.2, w: 8.4, h: 0.32, fontSize: 13, bold: true, color: DARK, fontFace: SANS, margin: 0 });
    note.addText("隔几天再打开工程，人很容易忘记上一轮做到哪里。每一轮只做一个页面或一件事，走同一个循环，就不靠记忆推进。", { x: 0.3, y: 0.58, w: 8.4, h: 0.4, fontSize: 10.5, color: TEXT, fontFace: SANS, margin: 0 });
    note.addText("上架、内购等隔一段时间还会再用的决定，写成文档保留。", { x: 0.3, y: 1.0, w: 8.4, h: 0.34, fontSize: 10.5, color: MUTED, fontFace: SANS, margin: 0 });
    addFooter(s, 10);
    s.render();
  }

  // ============================================================
  // S11 实操时间线
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: BG };
    addHeader(s, "04 · 实操时间线", "三个月，三段推进");
    const phases = [
      ["阶段一", "8.04 – 9.21", "功能按页面循环推进", "一轮只做一个页面或一件事，设计与编码交替，测试穿插其中", ["主界面", "头像昵称", "图标", "录入页", "查询排错", "设置页", "功能复查"]],
      ["阶段二", "9.21 – 9.30", "上架前集中处理", "开发者账号、内购、隐私政策、兼容性与验收，集中补齐", ["兼容补齐", "手工测试", "账号与内购", "正式上架"]],
      ["阶段三", "10.07 – 10.12", "上架后并行推进", "根据第一批用户反馈发布新版本，并行推进推广与软著", ["v1.0.2", "软著申请", "推广投稿"]],
    ];
    phases.forEach(([tag, date, title, desc, miles], i) => {
      const y = 1.42 + i * 1.24;
      // phase tag
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: CONTENT_X, y, w: 1.5, h: 1.1, fill: { color: i === 0 ? DARK : i === 1 ? PRIMARY : SECONDARY }, rectRadius: 0.1, shadow: makeShadow() });
      s.addText(tag, { x: CONTENT_X, y: y + 0.16, w: 1.5, h: 0.34, align: "center", fontSize: 13.5, bold: true, color: "FFFFFF", fontFace: SERIF, charSpacing: 1, margin: 0 });
      s.addText(date, { x: CONTENT_X, y: y + 0.56, w: 1.5, h: 0.3, align: "center", fontSize: 9.5, color: "CADCFC", fontFace: SANS, margin: 0 });
      // milestone card
      const card = addCard(s, CONTENT_X + 1.7, y, CONTENT_W - 1.7, 1.1);
      card.addText(title, { x: 0.22, y: 0.1, w: 3.4, h: 0.3, fontSize: 12, bold: true, color: DARK, fontFace: SANS, margin: 0 });
      card.addText(desc, { x: 3.7, y: 0.1, w: 3.4, h: 0.34, fontSize: 9, color: MUTED, fontFace: SANS, align: "right", margin: 0 });
      // milestone chips
      const innerW = CONTENT_W - 1.7 - 0.44;
      const chipW = Math.min(0.98, (innerW - (miles.length - 1) * 0.22) / miles.length);
      miles.forEach((m, j) => {
        const mx = 0.22 + j * (chipW + 0.22);
        card.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: mx, y: 0.52, w: chipW, h: 0.4, fill: { color: CARD_TEAL_BG }, rectRadius: 0.08 });
        card.addText(m, { x: mx, y: 0.52, w: chipW, h: 0.4, align: "center", valign: "middle", fontSize: 9.5, bold: true, color: PRIMARY, fontFace: SANS, margin: 0 });
        if (j < miles.length - 1) {
          card.addShape(pres.shapes.LINE, { x: mx + chipW + 0.03, y: 0.72, w: 0.16, h: 0, line: { color: "9AA5B1", width: 1.25, endArrowType: "triangle" } });
        }
      });
    });
    addFooter(s, 11);
    s.render();
  }

  // ============================================================
  // S12 量化分析
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: BG };
    addHeader(s, "05 · 量化分析", "用数据说话：103 小时的协作");
    const stats = [
      ["103 h", "Cursor 协作总时长", "功能 69.6 · 上架 26.1 · 文档 5.2"],
      ["88 次", "功能审核确认", "方案列出并确认后才动手"],
      ["30.7%", "返工率", "27 次返工 ÷ 88 次确认"],
      ["24 个", "bug", "已实现行为与既定规则不符"],
    ];
    const sw = 2.13, sh = 1.5, sgx = 0.19, sgy = 0.19, sx = CONTENT_X, sy = 1.45;
    stats.forEach(([n, l, sub], i) => {
      const x = sx + (i % 2) * (sw + sgx), y = sy + Math.floor(i / 2) * (sh + sgy);
      const card = addCard(s, x, y, sw, sh);
      card.addText(n, { x: 0.2, y: 0.16, w: sw - 0.4, h: 0.5, fontSize: 27, bold: true, color: PRIMARY, fontFace: SERIF, margin: 0 });
      card.addText(l, { x: 0.2, y: 0.72, w: sw - 0.4, h: 0.28, fontSize: 10.5, bold: true, color: DARK, fontFace: SANS, margin: 0 });
      card.addText(sub, { x: 0.2, y: 1.02, w: sw - 0.4, h: 0.34, fontSize: 8.5, color: MUTED, fontFace: SANS, margin: 0 });
    });
    s.addText("88 次功能审核确认的构成", { x: 5.15, y: 1.32, w: 4.3, h: 0.28, fontSize: 11.5, bold: true, color: TEXT, fontFace: SANS, margin: 0 });
    s.addChart(pres.charts.DOUGHNUT, [{
      name: "确认构成",
      labels: ["新功能 40", "返工 23", "缺陷修复 18", "仅复查 7"],
      values: [40, 23, 18, 7],
    }], {
      x: 5.15, y: 1.62, w: 4.35, h: 2.95, holeSize: 55,
      chartColors: ["065A82", "1C7293", "4A90A4", "8FB9CB"],
      showLegend: true, legendPos: "b", legendFontSize: 10, legendColor: "334155", legendFontFace: SANS,
      showPercent: true, dataLabelFontSize: 9.5, dataLabelColor: "FFFFFF", dataLabelFontFace: SANS,
      chartArea: { fill: { color: "FFFFFF" }, roundedCorners: true },
    });
    const strip = addCard(s, CONTENT_X, 4.72, CONTENT_W, 0.44, CARD_TEAL_BG);
    strip.addText("14 项功能中 9 项发生过返工——返工是 AI Coding 的常态，人的验收能力决定效率。", { x: 0.25, y: 0, w: 8.5, h: 0.44, valign: "middle", fontSize: 10.5, bold: true, color: PRIMARY, fontFace: SANS, margin: 0 });
    addFooter(s, 12);
    s.render();
  }

  // ============================================================
  // S13 工时分布
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: BG };
    addHeader(s, "05 · 量化分析", "功能开发工时分布（合计 69.6 小时）");
    const labels = ["其他功能", "测试与稳定性", "头像昵称与图标", "内购与免费层", "1.0.2 用户反馈", "查询页", "深色模式与兼容", "设置与选项管理", "录入页"];
    const values = [9.0, 3.7, 4.9, 5.2, 6.2, 7.2, 7.3, 7.6, 18.5];
    s.addChart(pres.charts.BAR, [{ name: "工时（小时）", labels, values }], {
      x: 0.5, y: 1.45, w: 5.95, h: 3.65, barDir: "bar",
      chartColors: ["065A82"], showLegend: false,
      catAxisLabelColor: "334155", catAxisLabelFontSize: 9.5, catAxisLabelFontFace: SANS,
      valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
      showValue: true, dataLabelPosition: "outEnd", dataLabelColor: "065A82", dataLabelFontSize: 9.5, dataLabelFontFace: SANS, dataLabelFormatCode: "0.0",
      chartArea: { fill: { color: "FFFFFF" }, roundedCorners: true },
      barGapWidthPct: 55,
    });
    const ic = addCard(s, 6.75, 1.45, 2.75, 2.5);
    ic.addText("解读", { x: 0.22, y: 0.16, w: 2.3, h: 0.3, fontSize: 12.5, bold: true, color: DARK, fontFace: SANS, margin: 0 });
    ic.addText([
      { text: "录入页占 26.7%：交互密度最高的页面最耗时", options: { bullet: { code: "2022", indent: 8 }, breakLine: true } },
      { text: "兼容性占 10.5%：若从第一天写入 Rules，可省下这笔返工", options: { bullet: { code: "2022", indent: 8 } } },
    ], { x: 0.22, y: 0.52, w: 2.35, h: 1.85, fontSize: 10, color: TEXT, fontFace: SANS, margin: 0, paraSpaceAfter: 6 });
    const nc = addCard(s, 6.75, 4.12, 2.75, 0.98, CARD_TEAL_BG);
    nc.addText("上架对话 26.1 小时、软著与推广文档 5.2 小时另计，三部分合计约 103 小时。", { x: 0.2, y: 0.12, w: 2.35, h: 0.76, fontSize: 9.5, color: PRIMARY, fontFace: SANS, margin: 0 });
    addFooter(s, 13);
    s.render();
  }

  // ============================================================
  // S14 常见问题
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: BG };
    addHeader(s, "06 · 常见问题", "三个最常被问到的问题");
    s.addImage({ path: IMG + "image-faq_4x3.jpg", x: 0.55, y: 1.75, w: 3.4, h: 2.55 });
    const qa = [
      ["需要花很多钱吗？", "不需要。两个月的模型费用共 60 美元，使用最强思考强度和 fast 模式，业余开发的用量额度没有用完。"],
      ["需要花很多时间吗？", "需要。这项工作本身具有连续性，时间主要花在一轮一轮的确认和验收上。"],
      ["可以在实际项目中使用吗？", "可以，条件是：人能定义目标、能验收结果、数据允许进入所用模型。涉密数据与无法验收的生产系统，不要交给 AI。"],
    ];
    qa.forEach(([q, a], i) => {
      const y = 1.42 + i * 1.26;
      const card = addCard(s, 4.3, y, 5.2, 1.12);
      card.addText("Q" + (i + 1), { x: 0.2, y: 0.14, w: 0.62, h: 0.4, fontSize: 17, bold: true, color: SECONDARY, fontFace: SERIF, margin: 0 });
      card.addText(q, { x: 0.82, y: 0.13, w: 4.2, h: 0.3, fontSize: 12.5, bold: true, color: DARK, fontFace: SANS, margin: 0 });
      card.addText(a, { x: 0.82, y: 0.46, w: 4.2, h: 0.58, fontSize: 9.5, color: MUTED, fontFace: SANS, margin: 0, autoFit: false, fit: "none" });
    });
    addFooter(s, 14);
    s.render();
  }

  // ============================================================
  // S15 结语 (divider)
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { path: IMG + "bg-divider_16x9.jpg" };
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, fill: { color: "111827", transparency: 35 } });
    s.addText("结语", { x: 0.9, y: 1.0, w: 6, h: 0.3, fontSize: 12, bold: true, color: "9FD8E8", charSpacing: 2, fontFace: SANS, margin: 0 });
    s.addText([
      { text: "瓶颈在想法是否值得做，", options: { breakLine: true } },
      { text: "难点在协作是否有效，", options: { breakLine: true } },
      { text: "风险在隐私能否守住。", options: {} },
    ], { x: 0.9, y: 1.6, w: 8.4, h: 2.2, fontSize: 25, bold: true, color: "FFFFFF", fontFace: SERIF, charSpacing: 1.5, margin: 0, paraSpaceAfter: 10 });
    s.addText("—— 与每一位想尝试 AI Coding 的人共勉", { x: 0.9, y: 4.1, w: 8.4, h: 0.4, fontSize: 13, color: "CADCFC", fontFace: SANS, margin: 0 });
    s.render();
  }

  // ============================================================
  // S16 谢谢
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { path: IMG + "bg-closing_16x9.jpg" };
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: SLIDE_W, h: SLIDE_H, fill: { color: "111827", transparency: 40 } });
    s.addText("谢谢观看", { x: 1, y: 1.15, w: 8, h: 0.75, align: "center", fontSize: 32, bold: true, color: "FFFFFF", fontFace: SERIF, charSpacing: 2.5, margin: 0 });
    s.addText("把目标说清楚，把验收做到位，剩下的交给 AI", { x: 1, y: 2.05, w: 8, h: 0.4, align: "center", fontSize: 13.5, color: "9FD8E8", fontFace: SANS, margin: 0 });
    const card = s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.9, y: 2.85, w: 6.2, h: 1.7, fill: { color: "1C7293", transparency: 45 }, rectRadius: 0.12, line: { color: "9FD8E8", width: 0.75, transparency: 40 } });
    card.addText([
      { text: "邮箱：fingalyang@outlook.com", options: { breakLine: true } },
      { text: "App Store：搜索「份子钱银行」", options: { breakLine: true } },
      { text: "GitHub：AI-Coding-Guide-IOS-APP", options: {} },
    ], { x: 0.4, y: 0.22, w: 5.4, h: 1.3, fontSize: 12, color: "FFFFFF", fontFace: SANS, margin: 0, paraSpaceAfter: 9 });
    s.addText("感谢陈老师开发这门课程，感谢李老师、任老师的专业工作", { x: 1, y: 4.85, w: 8, h: 0.3, align: "center", fontSize: 9.5, color: "8FA3C8", fontFace: SANS, margin: 0 });
    s.render();
  }

  await pres.writeFile({ fileName: "AI Coding 保姆级指引（iOS 移动应用）分享.pptx" });
  console.log("PPTX written");
})().catch(e => { console.error(e); process.exit(1); });
