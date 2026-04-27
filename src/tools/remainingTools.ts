import { ToolHandler } from "../types";
import { editCurrentSlideXml } from "./ooxml";
import { applyShapeStyle, shapeStyleFromInput } from "./shapeStyle";
import { normalizeGeometricShapeType } from "./shapeTypes";
import { resolveSlide, slideTargetFromInput } from "./slideTarget";

const SETTINGS_KEY = "claude-for-office.tool-settings";
const INSTRUCTIONS_KEY = "claude-for-office.user-instructions";
const BLOB_DB = "claude-office-blobs";
const BLOB_STORE = "blobs";
const DEFAULT_COLORS = ["#2F5597", "#70AD47", "#ED7D31", "#A64D79", "#5B9BD5", "#FFC000"];

export const executeOfficeJs: ToolHandler = async (input) => {
  const actions = arrayInput(input.actions, "actions");
  const results = await PowerPoint.run(async (ctx) => {
    const out: unknown[] = [];
    for (const action of actions) {
      out.push(await runOfficeAction(ctx, action));
    }
    await ctx.sync();
    return out;
  });
  return JSON.stringify({ ok: true, results }, null, 2);
};

export const editSlideChart: ToolHandler = async (input) => {
  const chartType = stringInput(input.chartType ?? input.type ?? "bar").toLowerCase();
  const categories = stringArray(input.categories);
  const series = Array.isArray(input.series) ? input.series as Array<Record<string, unknown>> : [];
  if (categories.length === 0 || series.length === 0) throw new Error("edit_slide_chart 需要 categories 和 series");

  const title = stringInput(input.title ?? "Chart");
  const left = numberOr(input.left, 80);
  const top = numberOr(input.top, 80);
  const width = numberOr(input.width, 760);
  const height = numberOr(input.height, 360);

  const result = await PowerPoint.run(async (ctx) => {
    const slide = await resolveSlide(ctx, slideTargetFromInput(input));
    const created: PowerPoint.Shape[] = [];
    const titleShape = slide.shapes.addTextBox(title, { left, top, width, height: 30 });
    applyShapeStyle(titleShape, { fontSize: 18, bold: true, textColor: "#1F2937" });
    titleShape.load("id");
    created.push(titleShape);

    if (chartType === "line") {
      created.push(...drawLineChart(slide, categories, series, left, top + 50, width, height - 70));
    } else if (chartType === "pie") {
      created.push(...drawPieChart(slide, categories, series, left, top + 50, width, height - 70));
    } else {
      created.push(...drawBarChart(slide, categories, series, left, top + 50, width, height - 70));
    }
    created.forEach((shape) => shape.load("id,name"));
    await ctx.sync();
    return {
      slideId: slide.id,
      chartType,
      shapeIds: created.map((shape) => shape.id),
    };
  });
  return `已生成 ${result.chartType} shape-based 图表，shapeIds=${result.shapeIds.join(", ")}`;
};

export const editSlideMaster: ToolHandler = async (input) => {
  const backgroundColor = typeof input.backgroundColor === "string" ? input.backgroundColor : undefined;
  const fontFamily = typeof input.fontFamily === "string" ? input.fontFamily : undefined;
  const themeColors = isObject(input.themeColors) ? input.themeColors : undefined;
  const decorations = Array.isArray(input.decorations) ? input.decorations as Array<Record<string, unknown>> : [];
  if (!backgroundColor && !fontFamily && !themeColors && decorations.length === 0) {
    throw new Error("edit_slide_master 至少需要 backgroundColor、fontFamily、themeColors 或 decorations");
  }

  const settings = readSettings();
  if (fontFamily) settings.fontFamily = fontFamily;
  if (themeColors) settings.themeColors = themeColors;
  writeSettings(settings);

  let slideId: string | undefined;
  if (backgroundColor) {
    const editResult = await editCurrentSlideXml(slideTargetFromInput(input), (doc) => setSlideBackground(doc, backgroundColor));
    slideId = editResult.newSlideId;
  }

  if (decorations.length > 0) {
    await PowerPoint.run(async (ctx) => {
      const slide = await resolveSlide(ctx, { ...slideTargetFromInput(input), slideId: slideId ?? slideTargetFromInput(input).slideId });
      for (const decoration of decorations) {
        const shapeType = normalizeGeometricShapeType(decoration.shapeType ?? "rectangle");
        const shape = slide.shapes.addGeometricShape(shapeType, {
          left: numberOr(decoration.left, 0),
          top: numberOr(decoration.top, 0),
          width: numberOr(decoration.width, 960),
          height: numberOr(decoration.height, 10),
        });
        if (typeof decoration.text === "string") shape.textFrame.textRange.text = decoration.text;
        applyShapeStyle(shape, shapeStyleFromInput(decoration));
      }
      await ctx.sync();
    });
  }

  return `已应用受控母版/主题设置${slideId ? `；新 slideId=${slideId}` : ""}。当前偏好=${JSON.stringify(settings)}`;
};

export const storeBlob: ToolHandler = async (input) => {
  const name = requireString(input.blobName ?? input.name, "blob_name/name");
  const source = stringInput(input.source ?? "base64");
  const mime = stringInput(input.mimeType ?? "application/octet-stream");
  let base64 = "";
  if (source === "text") {
    base64 = base64EncodeUtf8(stringInput(input.content ?? ""));
  } else if (source === "url") {
    const url = requireString(input.url ?? input.content, "url/content");
    const res = await fetch(url);
    if (!res.ok) throw new Error(`获取 URL 失败 (${res.status})`);
    base64 = await blobToBase64(await res.blob());
  } else {
    base64 = requireString(input.base64 ?? input.content, "base64/content");
  }
  await putBlobRecord({ name, mime, base64, createdAt: new Date().toISOString() });
  return `已保存 blob ${name} (mime=${mime}, bytes≈${Math.round(base64.length * 0.75)})`;
};

export const copyImageBetweenSlides: ToolHandler = async (input) => {
  const fromShapeId = requireString(input.fromShapeId ?? input.from_shape_id, "fromShapeId");
  const fromTarget = {
    slideId: typeof input.fromSlideId === "string" ? input.fromSlideId : typeof input.from_slide_id === "string" ? input.from_slide_id : undefined,
    slideIndex: typeof input.fromSlideIndex === "number" ? input.fromSlideIndex : typeof input.from_slide_index === "number" ? input.from_slide_index : undefined,
    pageNumber: typeof input.fromPageNumber === "number" ? input.fromPageNumber : typeof input.from_page_number === "number" ? input.from_page_number : undefined,
  };
  const toTarget = {
    slideId: typeof input.toSlideId === "string" ? input.toSlideId : typeof input.to_slide_id === "string" ? input.to_slide_id : undefined,
    slideIndex: typeof input.toSlideIndex === "number" ? input.toSlideIndex : typeof input.to_slide_index === "number" ? input.to_slide_index : undefined,
    pageNumber: typeof input.toPageNumber === "number" ? input.toPageNumber : typeof input.to_page_number === "number" ? input.to_page_number : undefined,
  };
  return await PowerPoint.run(async (ctx) => {
    const fromSlide = await resolveSlide(ctx, fromTarget);
    const toSlide = await resolveSlide(ctx, toTarget);
    const sourceShape = fromSlide.shapes.getItem(fromShapeId);
    sourceShape.load("id,left,top,width,height,type");
    const image = sourceShape.getImageAsBase64({ width: Math.max(1, numberOr(input.width, 512)) });
    await ctx.sync();
    if (!image.value) throw new Error(`形状 ${fromShapeId} 不能导出为图片`);
    const target = (toSlide.shapes as any).addImage(image.value);
    target.left = numberOr(input.left ?? input.x, sourceShape.left);
    target.top = numberOr(input.top ?? input.y, sourceShape.top);
    target.width = numberOr(input.width, sourceShape.width);
    target.height = numberOr(input.height, sourceShape.height);
    target.load("id");
    await ctx.sync();
    return `已复制图片形状 ${fromShapeId} 到 ${toSlide.id}，新 shapeId=${target.id}`;
  });
};

export const searchIcons: ToolHandler = async (input) => {
  const query = stringInput(input.query).toLowerCase();
  const top = Math.min(20, Math.max(1, numberOr(input.top, 5)));
  const results = ICONS
    .map((icon) => ({ ...icon, score: scoreIcon(query, icon) }))
    .filter((icon) => icon.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, top);
  return JSON.stringify({ query, results }, null, 2);
};

export const insertIcon: ToolHandler = async (input) => {
  const iconId = requireString(input.iconId ?? input.icon_id, "iconId");
  const icon = ICONS.find((item) => item.id === iconId);
  if (!icon) throw new Error(`未找到图标 ${iconId}，请先调用 search_icons`);
  const color = stringInput(input.color ?? "#2F5597");
  const svg = icon.svg.replace(/currentColor/g, color);
  const base64 = base64EncodeUtf8(svg);
  return await PowerPoint.run(async (ctx) => {
    const slide = await resolveSlide(ctx, slideTargetFromInput(input));
    const shape = (slide.shapes as any).addImage(base64);
    shape.left = numberOr(input.left ?? input.x, 444);
    shape.top = numberOr(input.top ?? input.y, 234);
    shape.width = numberOr(input.width, 72);
    shape.height = numberOr(input.height, 72);
    if (typeof input.description === "string") shape.name = input.description;
    shape.load("id");
    await ctx.sync();
    return `已插入图标 ${iconId}，shapeId=${shape.id}`;
  });
};

export const webSearch: ToolHandler = async (input) => {
  const query = requireString(input.query, "query");
  const top = Math.min(10, Math.max(1, numberOr(input.top, 5)));
  try {
    const res = await fetch("https://localhost:3001/__web-search", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ query, top }),
    });
    if (!res.ok) return `web_search 不可用: 本地搜索端点返回 HTTP ${res.status}`;
    return JSON.stringify(await res.json(), null, 2);
  } catch (error) {
    return `web_search 不可用: ${error instanceof Error ? error.message : String(error)}`;
  }
};

export const updateInstructions: ToolHandler = async (input) => {
  const operations = arrayInput(input.operations, "operations");
  let text = localStorage.getItem(INSTRUCTIONS_KEY) ?? "";
  for (const op of operations) {
    const oldText = typeof op.oldText === "string" ? op.oldText : typeof op.old_text === "string" ? op.old_text : "";
    const newText = typeof op.newText === "string" ? op.newText : typeof op.new_text === "string" ? op.new_text : "";
    if (!oldText) {
      text = text ? `${text}\n${newText}` : newText;
    } else {
      text = text.replace(oldText, newText);
    }
  }
  localStorage.setItem(INSTRUCTIONS_KEY, text);
  return `长期偏好已更新，当前长度=${text.length}`;
};

export const updateSetting: ToolHandler = async (input) => {
  const setting = requireString(input.setting, "setting");
  const settings = readSettings();
  if (typeof input.value === "boolean") {
    settings[setting] = input.value;
    writeSettings(settings);
    return `设置 ${setting} 已更新为 ${input.value}`;
  }
  return JSON.stringify({ setting, value: settings[setting] ?? null }, null, 2);
};

export const readSkill: ToolHandler = async (input) => {
  const skillName = requireString(input.skillName ?? input.skill_name, "skillName");
  const res = await fetch("https://localhost:3001/__skill-file", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ action: "read", name: skillName }),
  });
  if (!res.ok) return `读取技能失败: HTTP ${res.status}`;
  return JSON.stringify(await res.json(), null, 2);
};

export const createSkill: ToolHandler = async (input) => {
  const name = requireString(input.name, "name");
  const description = requireString(input.description, "description");
  const instructions = requireString(input.instructions, "instructions");
  const res = await fetch("https://localhost:3001/__skill-file", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ action: "create", name, description, instructions }),
  });
  if (!res.ok) return `创建技能失败: HTTP ${res.status}`;
  return JSON.stringify(await res.json(), null, 2);
};

export const unsupportedBridgeTool = (name: string): ToolHandler => async () =>
  `${name} 当前不可用：本地 PowerPoint add-in 未配置多 Agent/MCP 桥接协议。本工具已按“本地可用优先”注册为明确 unsupported，详见 docs/powerpoint-tools-roadmap.md。`;

export function getUserInstructions(): string {
  try {
    return localStorage.getItem(INSTRUCTIONS_KEY) ?? "";
  } catch {
    return "";
  }
}

function drawBarChart(slide: PowerPoint.Slide, categories: string[], series: Array<Record<string, unknown>>, left: number, top: number, width: number, height: number): PowerPoint.Shape[] {
  const shapes: PowerPoint.Shape[] = [];
  const values = series.flatMap((s) => numberArray(s.values));
  const max = Math.max(1, ...values);
  const plotLeft = left + 50;
  const plotTop = top + 20;
  const plotHeight = height - 80;
  const barAreaWidth = width - 110;
  const groupWidth = barAreaWidth / categories.length;
  const barWidth = Math.max(8, groupWidth / (series.length + 1));
  shapes.push(line(slide, plotLeft, plotTop + plotHeight, barAreaWidth, 0));
  shapes.push(line(slide, plotLeft, plotTop, 0, plotHeight));
  categories.forEach((cat, i) => {
    shapes.push(label(slide, cat, plotLeft + i * groupWidth, plotTop + plotHeight + 8, groupWidth, 18, 9));
    series.forEach((s, si) => {
      const value = numberArray(s.values)[i] ?? 0;
      const h = (value / max) * (plotHeight - 10);
      const bar = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
        left: plotLeft + i * groupWidth + si * barWidth + 6,
        top: plotTop + plotHeight - h,
        width: barWidth - 3,
        height: h,
      });
      applyShapeStyle(bar, { fillColor: colorAt(si), lineColor: colorAt(si) });
      shapes.push(bar);
    });
  });
  shapes.push(...legend(slide, series, left + width - 110, top + 10));
  return shapes;
}

function drawLineChart(slide: PowerPoint.Slide, categories: string[], series: Array<Record<string, unknown>>, left: number, top: number, width: number, height: number): PowerPoint.Shape[] {
  const shapes: PowerPoint.Shape[] = [];
  const values = series.flatMap((s) => numberArray(s.values));
  const max = Math.max(1, ...values);
  const plotLeft = left + 50;
  const plotTop = top + 20;
  const plotHeight = height - 80;
  const plotWidth = width - 120;
  shapes.push(line(slide, plotLeft, plotTop + plotHeight, plotWidth, 0));
  shapes.push(line(slide, plotLeft, plotTop, 0, plotHeight));
  categories.forEach((cat, i) => {
    shapes.push(label(slide, cat, plotLeft + i * (plotWidth / Math.max(1, categories.length - 1)) - 25, plotTop + plotHeight + 8, 50, 18, 9));
  });
  series.forEach((s, si) => {
    const vals = numberArray(s.values);
    let prev: { x: number; y: number } | null = null;
    vals.forEach((value, i) => {
      const x = plotLeft + i * (plotWidth / Math.max(1, categories.length - 1));
      const y = plotTop + plotHeight - (value / max) * (plotHeight - 10);
      const dot = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.ellipse, { left: x - 3, top: y - 3, width: 6, height: 6 });
      applyShapeStyle(dot, { fillColor: colorAt(si), lineColor: colorAt(si) });
      shapes.push(dot);
      if (prev) shapes.push(line(slide, prev.x, prev.y, x - prev.x, y - prev.y, colorAt(si), 2));
      prev = { x, y };
    });
  });
  shapes.push(...legend(slide, series, left + width - 110, top + 10));
  return shapes;
}

function drawPieChart(slide: PowerPoint.Slide, categories: string[], series: Array<Record<string, unknown>>, left: number, top: number, width: number, height: number): PowerPoint.Shape[] {
  const shapes: PowerPoint.Shape[] = [];
  const values = numberArray(series[0]?.values);
  const total = Math.max(1, values.reduce((sum, value) => sum + value, 0));
  const cx = left + 70;
  let y = top + 20;
  values.forEach((value, i) => {
    const w = Math.max(20, (value / total) * (width - 220));
    const slice = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, { left: cx, top: y, width: w, height: 20 });
    applyShapeStyle(slice, { fillColor: colorAt(i), lineColor: colorAt(i) });
    shapes.push(slice);
    shapes.push(label(slide, `${categories[i] ?? `Item ${i + 1}`} ${(value / total * 100).toFixed(0)}%`, cx + w + 8, y - 1, 150, 20, 10));
    y += 28;
  });
  shapes.push(label(slide, "Pie chart is represented as proportional bars for editability", left, top + height - 28, width, 20, 9));
  return shapes;
}

function legend(slide: PowerPoint.Slide, series: Array<Record<string, unknown>>, left: number, top: number): PowerPoint.Shape[] {
  const shapes: PowerPoint.Shape[] = [];
  series.forEach((s, i) => {
    const box = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, { left, top: top + i * 22, width: 12, height: 12 });
    applyShapeStyle(box, { fillColor: colorAt(i), lineColor: colorAt(i) });
    shapes.push(box);
    shapes.push(label(slide, stringInput(s.name ?? `Series ${i + 1}`), left + 18, top + i * 22 - 4, 90, 18, 9));
  });
  return shapes;
}

function label(slide: PowerPoint.Slide, text: string, left: number, top: number, width: number, height: number, fontSize: number): PowerPoint.Shape {
  const shape = slide.shapes.addTextBox(text, { left, top, width, height });
  applyShapeStyle(shape, { fontSize, textColor: "#374151" });
  return shape;
}

function line(slide: PowerPoint.Slide, left: number, top: number, width: number, height: number, color = "#6B7280", weight = 1): PowerPoint.Shape {
  const shape = slide.shapes.addLine(PowerPoint.ConnectorType.straight, { left, top, width: width || 0.01, height: height || 0.01 });
  shape.lineFormat.color = color;
  shape.lineFormat.weight = weight;
  return shape;
}

async function runOfficeAction(ctx: PowerPoint.RequestContext, action: Record<string, unknown>): Promise<Record<string, unknown>> {
  const type = requireString(action.type ?? action.action, "action.type");
  if (type === "selectSlide") {
    const slide = await resolveSlide(ctx, slideTargetFromInput(action));
    ctx.presentation.setSelectedSlides([slide.id]);
    return { type, slideId: slide.id };
  }
  const slide = await resolveSlide(ctx, slideTargetFromInput(action));
  if (type === "selectShapes") {
    const shapeIds = stringArray(action.shapeIds);
    slide.setSelectedShapes(shapeIds);
    return { type, slideId: slide.id, shapeIds };
  }
  const shapeId = requireString(action.shapeId, "shapeId");
  const shape = slide.shapes.getItem(shapeId);
  if (type === "moveShape") {
    if (typeof action.left === "number") shape.left = action.left;
    if (typeof action.top === "number") shape.top = action.top;
    return { type, shapeId };
  }
  if (type === "resizeShape") {
    if (typeof action.width === "number") shape.width = action.width;
    if (typeof action.height === "number") shape.height = action.height;
    return { type, shapeId };
  }
  if (type === "setShapeStyle") {
    applyShapeStyle(shape, shapeStyleFromInput(action));
    return { type, shapeId };
  }
  if (type === "bringToFront" || type === "sendToBack") {
    shape.setZOrder(type === "bringToFront" ? "BringToFront" : "SendToBack");
    return { type, shapeId };
  }
  throw new Error(`不支持的 execute_office_js action=${type}`);
}

function setSlideBackground(doc: Document, color: string): void {
  const pNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
  const aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const cSld = doc.getElementsByTagNameNS(pNs, "cSld")[0];
  if (!cSld) throw new Error("slide XML 中未找到 cSld");
  const existing = Array.from(cSld.childNodes).find((node) => node.nodeType === Node.ELEMENT_NODE && (node as Element).localName === "bg");
  if (existing) cSld.removeChild(existing);
  const bg = doc.createElementNS(pNs, "p:bg");
  const bgPr = doc.createElementNS(pNs, "p:bgPr");
  const solidFill = doc.createElementNS(aNs, "a:solidFill");
  const srgbClr = doc.createElementNS(aNs, "a:srgbClr");
  srgbClr.setAttribute("val", color.replace(/^#/, "").toUpperCase());
  solidFill.appendChild(srgbClr);
  bgPr.appendChild(solidFill);
  bg.appendChild(bgPr);
  cSld.insertBefore(bg, cSld.firstChild);
}

function readSettings(): Record<string, unknown> {
  try {
    return JSON.parse(localStorage.getItem(SETTINGS_KEY) ?? "{}");
  } catch {
    return {};
  }
}

function writeSettings(settings: Record<string, unknown>): void {
  localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
}

interface BlobRecord {
  name: string;
  mime: string;
  base64: string;
  createdAt: string;
}

function openBlobDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(BLOB_DB, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(BLOB_STORE, { keyPath: "name" });
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function putBlobRecord(record: BlobRecord): Promise<void> {
  const db = await openBlobDb();
  await new Promise<void>((resolve, reject) => {
    const tx = db.transaction(BLOB_STORE, "readwrite");
    tx.objectStore(BLOB_STORE).put(record);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
  db.close();
}

function blobToBase64(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result ?? "").split(",")[1] ?? "");
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(blob);
  });
}

const ICONS = [
  { id: "server", keywords: ["server", "database", "backend", "api"], description: "Server rack", svg: iconSvg('<rect x="5" y="5" width="14" height="6" rx="1"/><rect x="5" y="13" width="14" height="6" rx="1"/><circle cx="8" cy="8" r="1"/><circle cx="8" cy="16" r="1"/>') },
  { id: "database", keywords: ["database", "db", "sql", "storage"], description: "Database cylinder", svg: iconSvg('<ellipse cx="12" cy="5" rx="7" ry="3"/><path d="M5 5v10c0 1.7 3.1 3 7 3s7-1.3 7-3V5"/><path d="M5 10c0 1.7 3.1 3 7 3s7-1.3 7-3"/>') },
  { id: "user", keywords: ["user", "person", "account"], description: "User", svg: iconSvg('<circle cx="12" cy="8" r="4"/><path d="M4 21c1.5-4 4.2-6 8-6s6.5 2 8 6"/>') },
  { id: "arrow-right", keywords: ["arrow", "right", "next", "flow"], description: "Right arrow", svg: iconSvg('<path d="M5 12h14"/><path d="M13 6l6 6-6 6"/>') },
  { id: "gear", keywords: ["gear", "settings", "config"], description: "Gear", svg: iconSvg('<circle cx="12" cy="12" r="3"/><path d="M12 2v3M12 19v3M2 12h3M19 12h3M4.9 4.9l2.1 2.1M17 17l2.1 2.1M19.1 4.9 17 7M7 17l-2.1 2.1"/>') },
  { id: "search", keywords: ["search", "find", "inspect"], description: "Search", svg: iconSvg('<circle cx="10" cy="10" r="6"/><path d="M15 15l5 5"/>') },
  { id: "chart", keywords: ["chart", "bar", "analytics"], description: "Bar chart", svg: iconSvg('<path d="M4 20h16"/><rect x="6" y="11" width="3" height="6"/><rect x="11" y="7" width="3" height="10"/><rect x="16" y="4" width="3" height="13"/>') },
];

function iconSvg(content: string): string {
  return `<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${content}</svg>`;
}

function scoreIcon(query: string, icon: typeof ICONS[number]): number {
  if (!query) return 1;
  return icon.keywords.reduce((score, keyword) => score + (keyword.includes(query) || query.includes(keyword) ? 2 : 0), 0) +
    (icon.description.toLowerCase().includes(query) ? 1 : 0);
}

function arrayInput(value: unknown, name: string): Array<Record<string, unknown>> {
  if (!Array.isArray(value)) throw new Error(`${name} 必须是数组`);
  return value.filter(isObject);
}

function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function requireString(value: unknown, name: string): string {
  if (typeof value !== "string" || !value.trim()) throw new Error(`缺少 ${name}`);
  return value.trim();
}

function stringInput(value: unknown): string {
  return typeof value === "string" ? value : "";
}

function numberOr(value: unknown, fallback: number): number {
  return typeof value === "number" && Number.isFinite(value) ? value : fallback;
}

function stringArray(value: unknown): string[] {
  return Array.isArray(value) ? value.filter((item): item is string => typeof item === "string") : [];
}

function numberArray(value: unknown): number[] {
  return Array.isArray(value) ? value.map((item) => typeof item === "number" ? item : Number(item)).filter(Number.isFinite) : [];
}

function colorAt(index: number): string {
  return DEFAULT_COLORS[index % DEFAULT_COLORS.length];
}

function base64EncodeUtf8(value: string): string {
  const bytes = new TextEncoder().encode(value);
  let binary = "";
  bytes.forEach((byte) => { binary += String.fromCharCode(byte); });
  return btoa(binary);
}
