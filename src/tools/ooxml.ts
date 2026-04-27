import JSZip from "jszip";
import { withOfficeErrorContext } from "./officeErrors";
import { resolveSlide, SlideTarget } from "./slideTarget";

type SlideXmlMutator = (doc: Document) => void;
type SlideXmlReader<T> = (doc: Document, info: SlideXmlInfo) => T | Promise<T>;
type SlideZipMutator = (env: { zip: JSZip; markDirty: () => void; pptx: PptxHelper }) => void | Promise<void>;

export interface SlideXmlInfo {
  slideId: string;
  slideIndex: number;
  pageNumber: number;
}

export interface SlideXmlEditResult {
  oldSlideId: string;
  newSlideId: string;
  oldSlideIndex: number;
  newSlideIndex: number;
}

export interface SlideZipEditResult extends SlideXmlEditResult {
  dirty: boolean;
  artifacts: {
    beforeXmlPath?: string;
    beforeMetadataPath?: string;
    afterXmlPath?: string;
    afterMetadataPath?: string;
  };
}

const SLIDE_XML_PATH = "ppt/slides/slide1.xml";
const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

export interface PptxShapeRef {
  id: string;
  shapeId: string;
  alias?: string;
  left: number;
  top: number;
  width: number;
  height: number;
}

interface PptxShapeInput {
  id?: string | number;
  shapeType?: string;
  type?: string;
  text?: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  fill?: string;
  fillColor?: string;
  line?: string | { color?: string; width?: number };
  lineColor?: string;
  lineWidth?: number;
  textColor?: string;
  fontSize?: number;
  bold?: boolean;
}

interface PptxConnectorInput {
  from: string | number | PptxShapeRef;
  fromSide: "top" | "right" | "bottom" | "left";
  to: string | number | PptxShapeRef;
  toSide: "top" | "right" | "bottom" | "left";
  type?: "line" | "straight" | "elbow" | "bentConnector3";
  arrow?: "end" | "none";
  color?: string;
  width?: number;
  thickness?: number;
  dashStyle?: string;
}

interface PptxSlideHelper {
  emu: (pt: number) => number;
  nextShapeId: () => string;
  appendXml: (xml: string) => Element;
  addShape: (input: PptxShapeInput) => PptxShapeRef;
  addConnector: (input: PptxConnectorInput) => PptxShapeRef;
  save: () => void;
}

export interface PptxHelper extends PptxSlideHelper {
  openSlide: (path?: string) => Promise<PptxSlideHelper>;
}

export async function readCurrentSlideXml<T>(target: SlideTarget, reader: SlideXmlReader<T>): Promise<T> {
  return await PowerPoint.run(async (ctx) => {
    const slide = await resolveSlide(ctx, target);
    slide.load("id,index");
    const exported = slide.exportAsBase64();
    await syncWithContext(ctx, "导出目标幻灯片 exportAsBase64 失败");
    if (typeof exported.value !== "string" || exported.value.length === 0) {
      throw new Error("导出当前幻灯片失败：exportAsBase64 未返回有效 base64。请确认 Office 支持 PowerPointApi 1.8。");
    }
    const doc = await readSlideXmlFromBase64(exported.value);
    return await reader(doc, {
      slideId: slide.id,
      slideIndex: slide.index,
      pageNumber: slide.index + 1,
    });
  });
}

export async function editCurrentSlideXml(target: SlideTarget, mutator: SlideXmlMutator): Promise<SlideXmlEditResult> {
  return await PowerPoint.run(async (ctx) => {
    const presentation = ctx.presentation;
    const slide = await resolveSlide(ctx, target);
    const slides = presentation.slides;
    slides.load("items/id,index");
    slide.load("id,index");
    const exported = slide.exportAsBase64();
    await syncWithContext(ctx, "导出目标幻灯片 exportAsBase64 失败");

    if (typeof exported.value !== "string" || exported.value.length === 0) {
      throw new Error("导出当前幻灯片失败：exportAsBase64 未返回有效 base64。请确认 Office 支持 PowerPointApi 1.8。");
    }

    const oldSlideId = slide.id;
    const oldSlideIndex = slide.index;
    const existingSlideIds = new Set(slides.items.map((item) => item.id));
    const patchedBase64 = await patchSlideXml(exported.value, mutator);

    presentation.insertSlidesFromBase64(patchedBase64, {
      targetSlideId: oldSlideId,
      formatting: "KeepSourceFormatting",
    });
    await syncWithContext(ctx, `插入修正后的单页 PPTX 失败，targetSlideId=${oldSlideId}`);

    slides.load("items/id,index");
    await syncWithContext(ctx, "读取插入后的幻灯片列表失败");
    const insertedSlide = findInsertedSlide(slides.items, existingSlideIds, oldSlideIndex);
    if (!insertedSlide) {
      throw new Error("已插入修正后的幻灯片，但无法定位新 slide id。");
    }
    const newSlideId = insertedSlide.id;

    slide.delete();
    await syncWithContext(ctx, `删除旧幻灯片失败，oldSlideId=${oldSlideId}`);

    presentation.setSelectedSlides([newSlideId]);
    slides.load("items/id,index");
    await syncWithContext(ctx, `选中修正后的幻灯片失败，newSlideId=${newSlideId}`);
    const selected = slides.items.find((item) => item.id === newSlideId);
    return {
      oldSlideId,
      newSlideId,
      oldSlideIndex,
      newSlideIndex: selected?.index ?? oldSlideIndex,
    };
  });
}

export async function editCurrentSlideZip(
  target: SlideTarget,
  mutator: SlideZipMutator,
  options: {
    autosizeShapeIds?: string[];
    artifactMetadata?: Record<string, unknown>;
  } = {},
): Promise<SlideZipEditResult> {
  return await PowerPoint.run(async (ctx) => {
    const presentation = ctx.presentation;
    const slide = await resolveSlide(ctx, target);
    const slides = presentation.slides;
    slides.load("items/id,index");
    slide.load("id,index");
    const exported = slide.exportAsBase64();
    await syncWithContext(ctx, "导出目标幻灯片 exportAsBase64 失败");

    if (typeof exported.value !== "string" || exported.value.length === 0) {
      throw new Error("导出当前幻灯片失败：exportAsBase64 未返回有效 base64。请确认 Office 支持 PowerPointApi 1.8。");
    }

    const oldSlideId = slide.id;
    const oldSlideIndex = slide.index;
    const existingSlideIds = new Set(slides.items.map((item) => item.id));
    const zip = await JSZip.loadAsync(exported.value, { base64: true });
    const originalXml = await readSlideXmlFromZip(zip);
    const pptx = createPptxHelper(zip, originalXml);
    const beforeArtifact = await saveXmlArtifact("before", originalXml, {
      ...options.artifactMetadata,
      phase: "before",
      slideId: oldSlideId,
      slideIndex: oldSlideIndex,
      pageNumber: oldSlideIndex + 1,
    });

    let dirty = false;
    await mutator({
      zip,
      markDirty: () => {
        dirty = true;
      },
      pptx,
    });

    if (!dirty) {
      const executedXml = await readSlideXmlFromZip(zip).catch(() => "");
      const afterNoWriteArtifact = executedXml
        ? await saveXmlArtifact("after", executedXml, {
          ...options.artifactMetadata,
          phase: "after-no-write",
          slideId: oldSlideId,
          slideIndex: oldSlideIndex,
          pageNumber: oldSlideIndex + 1,
          dirty: false,
        })
        : null;
      return {
        oldSlideId,
        newSlideId: oldSlideId,
        oldSlideIndex,
        newSlideIndex: oldSlideIndex,
        dirty: false,
        artifacts: {
          beforeXmlPath: beforeArtifact?.path,
          beforeMetadataPath: beforeArtifact?.metadataPath,
          afterXmlPath: afterNoWriteArtifact?.path,
          afterMetadataPath: afterNoWriteArtifact?.metadataPath,
        },
      };
    }

    let modifiedXml = await readSlideXmlFromZip(zip);
    let afterArtifact = await saveXmlArtifact("after", modifiedXml, {
      ...options.artifactMetadata,
      phase: "after",
      slideId: oldSlideId,
      slideIndex: oldSlideIndex,
      pageNumber: oldSlideIndex + 1,
      autosizeShapeIds: options.autosizeShapeIds ?? [],
    });
    parseSlideXml(modifiedXml);

    if (options.autosizeShapeIds?.length) {
      await ensureNormAutofitForShapes(zip, options.autosizeShapeIds);
      modifiedXml = await readSlideXmlFromZip(zip);
      parseSlideXml(modifiedXml);
      afterArtifact = await saveXmlArtifact("after", modifiedXml, {
        ...options.artifactMetadata,
        phase: "after-autosize",
        slideId: oldSlideId,
        slideIndex: oldSlideIndex,
        pageNumber: oldSlideIndex + 1,
        autosizeShapeIds: options.autosizeShapeIds,
      }) ?? afterArtifact;
    }

    const patchedBase64 = await zip.generateAsync({
      type: "base64",
      mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    });

    presentation.insertSlidesFromBase64(patchedBase64, {
      targetSlideId: oldSlideId,
      formatting: "KeepSourceFormatting",
    });
    await syncWithContext(ctx, `插入 edit_slide_xml 修正后的单页 PPTX 失败，targetSlideId=${oldSlideId}`);

    slides.load("items/id,index");
    await syncWithContext(ctx, "读取插入后的幻灯片列表失败");
    const insertedSlide = findInsertedSlide(slides.items, existingSlideIds, oldSlideIndex);
    if (!insertedSlide) {
      throw new Error("已插入 edit_slide_xml 修正后的幻灯片，但无法定位新 slide id。");
    }
    const newSlideId = insertedSlide.id;

    slide.delete();
    await syncWithContext(ctx, `删除旧幻灯片失败，oldSlideId=${oldSlideId}`);

    presentation.setSelectedSlides([newSlideId]);
    slides.load("items/id,index");
    await syncWithContext(ctx, `选中 edit_slide_xml 修正后的幻灯片失败，newSlideId=${newSlideId}`);
    const selected = slides.items.find((item) => item.id === newSlideId);

    return {
      oldSlideId,
      newSlideId,
      oldSlideIndex,
      newSlideIndex: selected?.index ?? oldSlideIndex,
      dirty: true,
      artifacts: {
        beforeXmlPath: beforeArtifact?.path,
        beforeMetadataPath: beforeArtifact?.metadataPath,
        afterXmlPath: afterArtifact?.path,
        afterMetadataPath: afterArtifact?.metadataPath,
      },
    };
  });
}

async function syncWithContext(ctx: PowerPoint.RequestContext, message: string): Promise<void> {
  try {
    await ctx.sync();
  } catch (error) {
    throw withOfficeErrorContext(error, message);
  }
}

async function readSlideXmlFromZip(zip: JSZip): Promise<string> {
  const entry = zip.file(SLIDE_XML_PATH);
  if (!entry) throw new Error(`导出的单页 PPTX 中未找到 ${SLIDE_XML_PATH}`);
  return await entry.async("string");
}

async function patchSlideXml(base64: string, mutator: SlideXmlMutator): Promise<string> {
  const zip = await JSZip.loadAsync(base64, { base64: true });
  const xml = await readSlideXmlFromZip(zip);
  const doc = parseSlideXml(xml);
  mutator(doc);
  zip.file(SLIDE_XML_PATH, new XMLSerializer().serializeToString(doc));
  return await zip.generateAsync({
    type: "base64",
    mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  });
}

async function readSlideXmlFromBase64(base64: string): Promise<Document> {
  const zip = await JSZip.loadAsync(base64, { base64: true });
  return parseSlideXml(await readSlideXmlFromZip(zip));
}

function createPptxHelper(zip: JSZip, originalXml: string): PptxHelper {
  const slide = createSlideHelper(zip, originalXml);
  const ensureSlide = async (path?: string): Promise<PptxSlideHelper> => {
    if (path && path !== SLIDE_XML_PATH) {
      throw new Error(`edit_slide_xml 单页包内固定使用 ${SLIDE_XML_PATH}，不要猜测 ${path}`);
    }
    return slide;
  };
  return {
    openSlide: ensureSlide,
    emu: (pt: number) => pointToEmu(pt),
    nextShapeId: () => slide.nextShapeId(),
    appendXml: (xml: string) => slide.appendXml(xml),
    addShape: (input: PptxShapeInput) => slide.addShape(input),
    addConnector: (input: PptxConnectorInput) => slide.addConnector(input),
    save: () => slide.save(),
  };
}

function createSlideHelper(zip: JSZip, xml: string): PptxSlideHelper {
  const doc = parseSlideXml(xml);
  const spTree = doc.getElementsByTagNameNS(P_NS, "spTree")[0];
  if (!spTree) throw new Error("slide XML 中未找到 p:spTree");
  const aliases = new Map<string, PptxShapeRef>();
  const refs = new Map<string, PptxShapeRef>();

  const helper: PptxSlideHelper = {
    emu: (pt: number) => pointToEmu(pt),
    nextShapeId: () => nextShapeId(doc),
    appendXml: (xml: string) => appendXml(doc, spTree, xml),
    addShape: (input: PptxShapeInput) => {
      const ref = normalizeShapeInput(doc, input, aliases, refs);
      appendXml(doc, spTree, shapeXml(ref, input));
      aliases.set(ref.id, ref);
      refs.set(ref.shapeId, ref);
      if (ref.alias) aliases.set(ref.alias, ref);
      return ref;
    },
    addConnector: (input: PptxConnectorInput) => {
      const from = resolveShapeRef(input.from, aliases, refs);
      const to = resolveShapeRef(input.to, aliases, refs);
      const ref = connectorRef(doc, from, to);
      appendXml(doc, spTree, connectorXml(ref, from, to, input));
      refs.set(ref.shapeId, ref);
      return ref;
    },
    save: () => {
      zip.file(SLIDE_XML_PATH, new XMLSerializer().serializeToString(doc));
    },
  };
  return helper;
}

function normalizeShapeInput(
  doc: Document,
  input: PptxShapeInput,
  aliases: Map<string, PptxShapeRef>,
  refs: Map<string, PptxShapeRef>,
): PptxShapeRef {
  const idValue = input.id === undefined || input.id === null ? "" : String(input.id);
  const numericId = /^\d+$/.test(idValue) && !findCNvPrById(doc, idValue) ? idValue : nextShapeId(doc);
  const alias = idValue && idValue !== numericId ? idValue : undefined;
  const ref = {
    id: alias ?? numericId,
    shapeId: numericId,
    alias,
    left: numberValue(input.left, "left"),
    top: numberValue(input.top, "top"),
    width: numberValue(input.width, "width"),
    height: numberValue(input.height, "height"),
  };
  if (aliases.has(ref.id) || refs.has(ref.shapeId)) {
    throw new Error(`形状 id/alias 重复: ${ref.id}`);
  }
  return ref;
}

function connectorRef(doc: Document, from: PptxShapeRef, to: PptxShapeRef): PptxShapeRef {
  const id = nextShapeId(doc);
  const left = Math.min(from.left, to.left);
  const top = Math.min(from.top, to.top);
  return { id, shapeId: id, left, top, width: 0, height: 0 };
}

function resolveShapeRef(value: string | number | PptxShapeRef, aliases: Map<string, PptxShapeRef>, refs: Map<string, PptxShapeRef>): PptxShapeRef {
  if (typeof value === "object" && value && "shapeId" in value) return value;
  const key = String(value);
  const found = aliases.get(key) ?? refs.get(key);
  if (!found) throw new Error(`未找到连接器端点形状: ${key}`);
  return found;
}

function shapeXml(ref: PptxShapeRef, input: PptxShapeInput): string {
  const shapeType = normalizeShapeType(input.shapeType ?? input.type ?? "rect");
  const fill = fillXml(input.fill ?? input.fillColor ?? "accent1");
  const line = lineXml(input.line ?? input.lineColor ?? "dk1", input.lineWidth);
  const textColor = colorFillXml(input.textColor ?? "lt1");
  const fontSize = Math.max(6, Math.round(input.fontSize ?? 14)) * 100;
  const bold = input.bold === false ? "0" : "1";
  return `<p:sp xmlns:p="${P_NS}" xmlns:a="${A_NS}">
<p:nvSpPr><p:cNvPr id="${ref.shapeId}" name="${escapeXml(input.text || ref.alias || `Shape ${ref.shapeId}`)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="${pointToEmu(ref.left)}" y="${pointToEmu(ref.top)}"/><a:ext cx="${pointToEmu(ref.width)}" cy="${pointToEmu(ref.height)}"/></a:xfrm><a:prstGeom prst="${shapeType}"><a:avLst/></a:prstGeom>${fill}${line}</p:spPr>
${textBodyXml(input.text ?? "", fontSize, bold, textColor)}
</p:sp>`;
}

function connectorXml(ref: PptxShapeRef, from: PptxShapeRef, to: PptxShapeRef, input: PptxConnectorInput): string {
  const start = sidePoint(from, input.fromSide);
  const end = sidePoint(to, input.toSide);
  const sameAxis = Math.abs(start.x - end.x) < 0.001 || Math.abs(start.y - end.y) < 0.001;
  const requested = input.type ?? (sameAxis ? "line" : "bentConnector3");
  const prst = requested === "straight" || requested === "line" ? "line" : "bentConnector3";
  const left = Math.min(start.x, end.x);
  const top = Math.min(start.y, end.y);
  const width = Math.abs(end.x - start.x);
  const height = Math.abs(end.y - start.y);
  const flipH = start.x > end.x ? ' flipH="1"' : "";
  const flipV = start.y > end.y ? ' flipV="1"' : "";
  const color = input.color ?? "dk1";
  const weight = pointToEmu(input.width ?? input.thickness ?? 2);
  const arrow = input.arrow === "none" ? "" : '<a:tailEnd type="arrow" w="med" len="med"/>';
  const dash = dashXml(input.dashStyle);
  return `<p:cxnSp xmlns:p="${P_NS}" xmlns:a="${A_NS}">
<p:nvCxnSpPr><p:cNvPr id="${ref.shapeId}" name="Connector ${ref.shapeId}"/><p:cNvCxnSpPr><a:stCxn id="${from.shapeId}" idx="${sideIndex(input.fromSide)}"/><a:endCxn id="${to.shapeId}" idx="${sideIndex(input.toSide)}"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr>
<p:spPr><a:xfrm${flipH}${flipV}><a:off x="${pointToEmu(left)}" y="${pointToEmu(top)}"/><a:ext cx="${pointToEmu(width)}" cy="${pointToEmu(height)}"/></a:xfrm><a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom><a:ln w="${weight}" cap="flat" cmpd="sng" algn="ctr">${colorFillXml(color)}${dash}<a:miter lim="800000"/>${arrow}</a:ln></p:spPr>
</p:cxnSp>`;
}

function textBodyXml(text: string, fontSize: number, bold: string, textColor: string): string {
  const paragraphs = String(text || "").split(/\n/);
  const body = paragraphs.map((line) =>
    `<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="zh-CN" sz="${fontSize}" b="${bold}">${textColor}</a:rPr><a:t>${escapeXml(line)}</a:t></a:r><a:endParaRPr lang="zh-CN" sz="${fontSize}"/></a:p>`
  ).join("");
  return `<p:txBody><a:bodyPr wrap="square" anchor="ctr"><a:normAutofit/></a:bodyPr><a:lstStyle/>${body || "<a:p/>"}</p:txBody>`;
}

function appendXml(doc: Document, spTree: Element, xml: string): Element {
  const parsed = new DOMParser().parseFromString(`<root xmlns:p="${P_NS}" xmlns:a="${A_NS}">${xml}</root>`, "application/xml");
  const error = parsed.getElementsByTagName("parsererror")[0];
  if (error) throw new Error(`helper appendXml 解析失败: ${error.textContent?.trim() ?? "unknown parser error"}`);
  const element = Array.from(parsed.documentElement.childNodes).find((node) => node.nodeType === Node.ELEMENT_NODE) as Element | undefined;
  if (!element) throw new Error("appendXml 需要传入一个 XML 元素");
  const imported = doc.importNode(element, true) as Element;
  const extLst = firstChildByNs(spTree, P_NS, "extLst");
  if (extLst) spTree.insertBefore(imported, extLst);
  else spTree.appendChild(imported);
  return imported;
}

function nextShapeId(doc: Document): string {
  const ids = Array.from(doc.getElementsByTagNameNS(P_NS, "cNvPr"))
    .map((node) => Number(node.getAttribute("id")))
    .filter((id) => Number.isFinite(id));
  return String(Math.max(0, ...ids) + 1);
}

function findCNvPrById(doc: Document, id: string): Element | null {
  return Array.from(doc.getElementsByTagNameNS(P_NS, "cNvPr"))
    .find((node) => node.getAttribute("id") === id) ?? null;
}

function sidePoint(shape: PptxShapeRef, side: "top" | "right" | "bottom" | "left"): { x: number; y: number } {
  if (side === "top") return { x: shape.left + shape.width / 2, y: shape.top };
  if (side === "right") return { x: shape.left + shape.width, y: shape.top + shape.height / 2 };
  if (side === "bottom") return { x: shape.left + shape.width / 2, y: shape.top + shape.height };
  return { x: shape.left, y: shape.top + shape.height / 2 };
}

function sideIndex(side: "top" | "right" | "bottom" | "left"): number {
  return { top: 0, right: 1, bottom: 2, left: 3 }[side];
}

function normalizeShapeType(value: string): string {
  const key = value.replace(/[\s_-]/g, "").toLowerCase();
  const map: Record<string, string> = {
    rect: "rect",
    rectangle: "rect",
    roundrect: "roundRect",
    roundrectangle: "roundRect",
    roundedrectangle: "roundRect",
    ellipse: "ellipse",
    oval: "ellipse",
    diamond: "diamond",
    can: "can",
    cylinder: "can",
    database: "can",
    flowchartterminator: "flowChartTerminator",
    terminator: "flowChartTerminator",
    flowchartinputoutput: "flowChartInputOutput",
    inputoutput: "flowChartInputOutput",
    flowchartdecision: "flowChartDecision",
    decision: "flowChartDecision",
    flowchartprocess: "flowChartProcess",
    process: "flowChartProcess",
  };
  const normalized = map[key];
  if (!normalized) throw new Error(`不支持的 shapeType: ${value}`);
  return normalized;
}

function fillXml(color: string): string {
  const normalized = String(color || "").trim();
  if (!normalized || normalized === "none" || normalized === "transparent") return "<a:noFill/>";
  return `<a:solidFill>${colorNodeXml(normalized)}</a:solidFill>`;
}

function lineXml(line: string | { color?: string; width?: number }, width?: number): string {
  const color = typeof line === "string" ? line : line.color ?? "dk1";
  const lineWidth = typeof line === "object" && typeof line.width === "number" ? line.width : width ?? 1;
  if (color === "none" || color === "transparent") return "<a:ln><a:noFill/></a:ln>";
  return `<a:ln w="${pointToEmu(lineWidth)}">${colorFillXml(color)}</a:ln>`;
}

function colorFillXml(color: string): string {
  return `<a:solidFill>${colorNodeXml(color)}</a:solidFill>`;
}

function colorNodeXml(color: string): string {
  const raw = String(color || "").trim();
  const hex = raw.replace(/^#/, "");
  if (/^[0-9a-fA-F]{6}$/.test(hex)) return `<a:srgbClr val="${hex.toUpperCase()}"/>`;
  return `<a:schemeClr val="${escapeXml(raw || "dk1")}"/>`;
}

function dashXml(value?: string): string {
  if (!value || value === "solid") return '<a:prstDash val="solid"/>';
  const map: Record<string, string> = { dash: "dash", dot: "dot", dashDot: "dashDot" };
  return `<a:prstDash val="${map[value] ?? "solid"}"/>`;
}

function numberValue(value: unknown, name: string): number {
  if (typeof value !== "number" || !Number.isFinite(value)) throw new Error(`pptx.addShape 缺少 number 类型 ${name}`);
  return value;
}

function pointToEmu(pt: number): number {
  if (!Number.isFinite(pt)) throw new Error(`非法 pt 数值: ${pt}`);
  return Math.round(pt * 12700);
}

function escapeXml(value: string): string {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

export function parseSlideXml(xml: string): Document {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const error = doc.getElementsByTagName("parsererror")[0];
  if (error) {
    throw new Error(`解析 slide XML 失败: ${error.textContent?.trim() ?? "unknown parser error"}`);
  }
  return doc;
}

export function serializeXmlElement(element: Element): string {
  return new XMLSerializer().serializeToString(element);
}

function findInsertedSlide(
  slides: PowerPoint.Slide[],
  existingSlideIds: Set<string>,
  oldSlideIndex: number,
): PowerPoint.Slide | null {
  const newSlides = slides.filter((slide) => !existingSlideIds.has(slide.id));
  if (newSlides.length === 1) return newSlides[0];
  return slides.find((slide) => slide.index === oldSlideIndex + 1 && !existingSlideIds.has(slide.id)) ?? null;
}

async function ensureNormAutofitForShapes(zip: JSZip, shapeIds: string[]): Promise<void> {
  const idSet = new Set(shapeIds.map(String).filter(Boolean));
  if (idSet.size === 0) return;
  const doc = parseSlideXml(await readSlideXmlFromZip(zip));
  const shapes = [
    ...Array.from(doc.getElementsByTagNameNS(P_NS, "sp")),
    ...Array.from(doc.getElementsByTagNameNS(P_NS, "cxnSp")),
  ];
  for (const shape of shapes) {
    const cNvPr = shape.getElementsByTagNameNS(P_NS, "cNvPr")[0];
    const id = cNvPr?.getAttribute("id");
    if (!id || !idSet.has(id)) continue;
    const txBody = firstChildByNs(shape, P_NS, "txBody");
    if (!txBody) continue;
    let bodyPr = firstChildByNs(txBody, A_NS, "bodyPr");
    if (!bodyPr) {
      bodyPr = doc.createElementNS(A_NS, "a:bodyPr");
      txBody.insertBefore(bodyPr, txBody.firstChild);
    }
    for (const child of Array.from(bodyPr.childNodes)) {
      if (child.nodeType !== Node.ELEMENT_NODE) continue;
      const element = child as Element;
      if (element.namespaceURI === A_NS && ["normAutofit", "noAutofit", "spAutoFit"].includes(element.localName)) {
        bodyPr.removeChild(element);
      }
    }
    bodyPr.appendChild(doc.createElementNS(A_NS, "a:normAutofit"));
  }
  zip.file(SLIDE_XML_PATH, new XMLSerializer().serializeToString(doc));
}

function firstChildByNs(parent: Element, namespaceUri: string, localName: string): Element | null {
  return Array.from(parent.childNodes).find((node) =>
    node.nodeType === Node.ELEMENT_NODE &&
    (node as Element).namespaceURI === namespaceUri &&
    (node as Element).localName === localName
  ) as Element | null;
}

interface ArtifactResult {
  ok: boolean;
  path?: string;
  metadataPath?: string;
}

async function saveXmlArtifact(
  phase: "before" | "after",
  xml: string,
  metadata: Record<string, unknown>,
): Promise<ArtifactResult | null> {
  try {
    const slidePart = safeFilePart(`${metadata.slideIndex ?? "unknown"}-${metadata.slideId ?? "slide"}`);
    const stamp = new Date().toISOString().replace(/[:.]/g, "-");
    const res = await fetch("https://localhost:3001/__debug-artifact", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        kind: "edit-slide-xml",
        filename: `edit-slide-xml-${stamp}-${slidePart}-${phase}.xml`,
        mime: "application/xml",
        base64: base64EncodeUtf8(xml),
        metadata,
      }),
    });
    if (!res.ok) return null;
    return await res.json() as ArtifactResult;
  } catch {
    return null;
  }
}

function base64EncodeUtf8(value: string): string {
  const bytes = new TextEncoder().encode(value);
  let binary = "";
  bytes.forEach((byte) => { binary += String.fromCharCode(byte); });
  return btoa(binary);
}

function safeFilePart(value: string): string {
  return value
    .replace(/[^a-zA-Z0-9._-]/g, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 100) || "slide";
}
