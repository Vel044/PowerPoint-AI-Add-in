import { ToolHandler } from "../types";
import { editCurrentSlideXml, SlideXmlEditResult } from "./ooxml";
import { describeOfficeError, withOfficeErrorContext } from "./officeErrors";
import { normalizeGeometricShapeType } from "./shapeTypes";
import { resolveSlide, slideTargetFromInput } from "./slideTarget";

type LayoutMode = "vertical" | "horizontal" | "layered" | "tree";
type NodeShape = string;
export type Side = "top" | "bottom" | "left" | "right";

export interface Rect { left: number; top: number; width: number; height: number; }

interface DiagramNode {
  id: string;
  text: string;
  shape?: NodeShape;
  level?: number;
}

interface DiagramEdge {
  from: string;
  to: string;
}

interface Canvas {
  left: number;
  top: number;
  width: number;
  height: number;
}

interface Placed {
  node: DiagramNode;
  left: number;
  top: number;
  width: number;
  height: number;
}

const DEFAULT_CANVAS: Canvas = { left: 40, top: 80, width: 880, height: 420 };
const DEFAULT_NODE_W = 160;
const DEFAULT_NODE_H = 60;
const CONNECTOR_COLOR = "#2F5597";
const CONNECTOR_THICKNESS = 2;
const CONNECTOR_EPSILON = 0.01;
const POINT_TO_EMU = 12700;
const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

interface Point {
  x: number;
  y: number;
}

export interface ConnectorDrawResult {
  lineSegments: number;
  straightConnectors: number;
  elbowConnectors: number;
  arrows: number;
  shapeCount: number;
  arrowHeadsEnabled: boolean;
}

interface ConnectorOptions {
  arrow?: "none" | "end";
  color?: string;
  thickness?: number;
  dashStyle?: string;
  mode?: "direct" | "orthogonal";
}

export interface ConnectorXmlPatch {
  connectorShapeId: string;
  fromShapeId: string;
  fromSide: Side;
  toShapeId: string;
  toSide: Side;
  start: Point;
  end: Point;
  connectorType: "straight" | "elbow";
  arrow: "none" | "end";
  color: string;
  thickness: number;
  dashStyle?: string;
}

interface PendingConnectorXmlPatch extends Omit<ConnectorXmlPatch, "connectorShapeId"> {
  connectorShape: PowerPoint.Shape;
}

export interface ConnectorCreationResult extends ConnectorDrawResult {
  patches: PendingConnectorXmlPatch[];
}

function addConnectorPlaceholder(
  slide: PowerPoint.Slide,
  start: Point,
  end: Point,
  connectorType: "straight" | "elbow",
  color: string,
  thickness: number,
  dashStyle?: string,
): PowerPoint.Shape | null {
  const dx = end.x - start.x;
  const dy = end.y - start.y;
  const length = Math.hypot(dx, dy);
  if (length < 0.5) return null;

  const isElbow = connectorType === "elbow";
  const left = isElbow ? Math.min(start.x, end.x) : start.x;
  const top = isElbow ? Math.min(start.y, end.y) : start.y;
  const width = isElbow ? Math.max(Math.abs(dx), CONNECTOR_EPSILON) : (Math.abs(dx) < 0.5 ? CONNECTOR_EPSILON : dx);
  const height = isElbow ? Math.max(Math.abs(dy), CONNECTOR_EPSILON) : (Math.abs(dy) < 0.5 ? CONNECTOR_EPSILON : dy);

  const shape = slide.shapes.addLine(isElbow ? PowerPoint.ConnectorType.elbow : PowerPoint.ConnectorType.straight, {
    left,
    top,
    width,
    height,
  });
  shape.lineFormat.visible = true;
  shape.lineFormat.color = color;
  shape.lineFormat.weight = thickness;
  shape.lineFormat.dashStyle = normalizeDashStyle(dashStyle ?? "solid") as any;
  shape.lineFormat.style = "Single";
  shape.load("id");
  return shape;
}

function sideIndex(side: Side): number {
  switch (side) {
    case "top": return 0;
    case "right": return 1;
    case "bottom": return 2;
    case "left": return 3;
  }
}

export function drawConnectedLine(
  slide: PowerPoint.Slide,
  fromShapeId: string,
  fromRect: Rect,
  fromSide: Side,
  toShapeId: string,
  toRect: Rect,
  toSide: Side,
  options: ConnectorOptions = {},
): ConnectorCreationResult {
  const start = sidePoint(fromRect, fromSide);
  const end = sidePoint(toRect, toSide);
  const isAxisAligned = Math.abs(start.x - end.x) < 0.5 || Math.abs(start.y - end.y) < 0.5;
  const connectorType: "straight" | "elbow" = options.mode === "direct" || isAxisAligned ? "straight" : "elbow";
  const color = options.color ?? CONNECTOR_COLOR;
  const thickness = options.thickness ?? CONNECTOR_THICKNESS;
  const dashStyle = options.dashStyle;
  const arrow = options.arrow ?? "end";
  const shape = addConnectorPlaceholder(slide, start, end, connectorType, color, thickness, dashStyle);

  if (!shape) {
    return {
      lineSegments: 0,
      straightConnectors: 0,
      elbowConnectors: 0,
      arrows: 0,
      shapeCount: 0,
      arrowHeadsEnabled: true,
      patches: [],
    };
  }

  return {
    lineSegments: 1,
    straightConnectors: connectorType === "straight" ? 1 : 0,
    elbowConnectors: connectorType === "elbow" ? 1 : 0,
    arrows: arrow === "end" ? 1 : 0,
    shapeCount: 1,
    arrowHeadsEnabled: true,
    patches: [{
      connectorShape: shape,
      fromShapeId,
      fromSide,
      toShapeId,
      toSide,
      start,
      end,
      connectorType,
      arrow,
      color,
      thickness,
      dashStyle,
    }],
  };
}

export function resolveConnectorXmlPatches(results: ConnectorCreationResult[]): ConnectorXmlPatch[] {
  return results.flatMap((result) => result.patches.map((patch) => ({
    connectorShapeId: patch.connectorShape.id,
    fromShapeId: patch.fromShapeId,
    fromSide: patch.fromSide,
    toShapeId: patch.toShapeId,
    toSide: patch.toSide,
    start: patch.start,
    end: patch.end,
    connectorType: patch.connectorType,
    arrow: patch.arrow,
    color: patch.color,
    thickness: patch.thickness,
    dashStyle: patch.dashStyle,
  })));
}

export async function applyConnectorXmlPatches(
  target: { slideId?: string; slideIndex?: number; pageNumber?: number },
  patches: ConnectorXmlPatch[],
): Promise<SlideXmlEditResult | null> {
  if (patches.length === 0) return null;
  return await editCurrentSlideXml(target, (doc) => {
    for (const patch of patches) patchConnectorXml(doc, patch);
  });
}

export function patchConnectorXml(doc: Document, patch: ConnectorXmlPatch): void {
  const connector = findConnector(doc, patch.connectorShapeId);
  if (!connector) throw new Error(`slide XML 中未找到连接器 ${patch.connectorShapeId}`);

  const nvCxnSpPr = ensureChild(doc, connector, P_NS, "nvCxnSpPr");
  const cNvCxnSpPr = ensureChild(doc, nvCxnSpPr, P_NS, "cNvCxnSpPr");
  removeAllChildren(cNvCxnSpPr);
  const stCxn = doc.createElementNS(A_NS, "a:stCxn");
  stCxn.setAttribute("id", patch.fromShapeId);
  stCxn.setAttribute("idx", String(sideIndex(patch.fromSide)));
  const endCxn = doc.createElementNS(A_NS, "a:endCxn");
  endCxn.setAttribute("id", patch.toShapeId);
  endCxn.setAttribute("idx", String(sideIndex(patch.toSide)));
  cNvCxnSpPr.append(stCxn, endCxn);

  const spPr = ensureChild(doc, connector, P_NS, "spPr");
  setTransform(doc, spPr, patch.start, patch.end);
  setPresetGeometry(doc, spPr, patch.connectorType);
  setLine(doc, spPr, patch);
}

function findConnector(doc: Document, shapeId: string): Element | null {
  const connectors = Array.from(doc.getElementsByTagNameNS(P_NS, "cxnSp"));
  return connectors.find((connector) => {
    const cNvPr = connector.getElementsByTagNameNS(P_NS, "cNvPr")[0];
    return cNvPr?.getAttribute("id") === shapeId;
  }) ?? null;
}

function ensureChild(doc: Document, parent: Element, namespaceUri: string, localName: string): Element {
  const existing = Array.from(parent.childNodes).find((node) =>
    node.nodeType === Node.ELEMENT_NODE &&
    (node as Element).namespaceURI === namespaceUri &&
    (node as Element).localName === localName
  ) as Element | undefined;
  if (existing) return existing;
  const prefix = namespaceUri === P_NS ? "p" : "a";
  const child = doc.createElementNS(namespaceUri, `${prefix}:${localName}`);
  parent.appendChild(child);
  return child;
}

function removeAllChildren(element: Element): void {
  while (element.firstChild) element.removeChild(element.firstChild);
}

function setTransform(doc: Document, spPr: Element, start: Point, end: Point): void {
  const xfrm = ensureChild(doc, spPr, A_NS, "xfrm");
  xfrm.removeAttribute("flipH");
  xfrm.removeAttribute("flipV");
  if (end.x < start.x) xfrm.setAttribute("flipH", "1");
  if (end.y < start.y) xfrm.setAttribute("flipV", "1");

  const off = ensureChild(doc, xfrm, A_NS, "off");
  off.setAttribute("x", String(toEmu(Math.min(start.x, end.x))));
  off.setAttribute("y", String(toEmu(Math.min(start.y, end.y))));

  const ext = ensureChild(doc, xfrm, A_NS, "ext");
  ext.setAttribute("cx", String(toEmu(Math.abs(end.x - start.x))));
  ext.setAttribute("cy", String(toEmu(Math.abs(end.y - start.y))));
}

function setPresetGeometry(doc: Document, spPr: Element, connectorType: "straight" | "elbow"): void {
  const geom = ensureChild(doc, spPr, A_NS, "prstGeom");
  geom.setAttribute("prst", connectorType === "elbow" ? "bentConnector3" : "line");
  let avLst = Array.from(geom.childNodes).find((node) =>
    node.nodeType === Node.ELEMENT_NODE &&
    (node as Element).namespaceURI === A_NS &&
    (node as Element).localName === "avLst"
  ) as Element | undefined;
  if (!avLst) {
    avLst = doc.createElementNS(A_NS, "a:avLst");
    geom.appendChild(avLst);
  }
}

function setLine(doc: Document, spPr: Element, patch: ConnectorXmlPatch): void {
  const line = ensureChild(doc, spPr, A_NS, "ln");
  removeAllChildren(line);
  line.setAttribute("w", String(toEmu(patch.thickness)));
  line.setAttribute("cap", "flat");
  line.setAttribute("cmpd", "sng");
  line.setAttribute("algn", "ctr");

  const solidFill = doc.createElementNS(A_NS, "a:solidFill");
  const srgbClr = doc.createElementNS(A_NS, "a:srgbClr");
  srgbClr.setAttribute("val", normalizeColor(patch.color));
  solidFill.appendChild(srgbClr);

  const dash = doc.createElementNS(A_NS, "a:prstDash");
  dash.setAttribute("val", normalizePresetDash(patch.dashStyle ?? "solid"));

  const headEnd = doc.createElementNS(A_NS, "a:headEnd");
  headEnd.setAttribute("type", "none");

  const tailEnd = doc.createElementNS(A_NS, "a:tailEnd");
  tailEnd.setAttribute("type", patch.arrow === "end" ? "arrow" : "none");
  if (patch.arrow === "end") {
    tailEnd.setAttribute("w", "med");
    tailEnd.setAttribute("len", "med");
  }

  const miter = doc.createElementNS(A_NS, "a:miter");
  miter.setAttribute("lim", "800000");
  line.append(solidFill, dash, headEnd, tailEnd, miter);
}

function normalizeDashStyle(value: string): string {
  const normalized = value.trim().toLowerCase();
  if (normalized === "dash" || normalized === "dashed") return "Dash";
  if (normalized === "dot" || normalized === "dotted") return "Dot";
  if (normalized === "dashdot" || normalized === "dashDot") return "DashDot";
  return "Solid";
}

function normalizePresetDash(value: string): string {
  const normalized = value.trim().toLowerCase();
  if (normalized === "dash" || normalized === "dashed") return "dash";
  if (normalized === "dot" || normalized === "dotted") return "dot";
  if (normalized === "dashdot" || normalized === "dashDot") return "dashDot";
  return "solid";
}

function normalizeColor(color: string): string {
  return color.replace(/^#/, "").toUpperCase();
}

function toEmu(points: number): number {
  return Math.round(points * POINT_TO_EMU);
}

// ── 布局算法 ──────────────────────────────────────────

function assignLevelsFromEdges(nodes: DiagramNode[], edges: DiagramEdge[]): Map<string, number> {
  const levels = new Map<string, number>();
  const parents = new Map<string, string[]>();
  const children = new Map<string, string[]>();
  for (const n of nodes) {
    parents.set(n.id, []);
    children.set(n.id, []);
  }
  for (const e of edges) {
    parents.get(e.to)?.push(e.from);
    children.get(e.from)?.push(e.to);
  }
  const roots = nodes.filter((n) => (parents.get(n.id) ?? []).length === 0);
  const queue: string[] = roots.map((r) => {
    levels.set(r.id, 0);
    return r.id;
  });
  while (queue.length) {
    const id = queue.shift()!;
    const lv = levels.get(id)!;
    for (const c of children.get(id) ?? []) {
      const prev = levels.get(c);
      const next = lv + 1;
      if (prev === undefined || next > prev) {
        levels.set(c, next);
        queue.push(c);
      }
    }
  }
  for (const n of nodes) if (!levels.has(n.id)) levels.set(n.id, 0);
  return levels;
}

const MIN_GAP = 20;

function layoutTree(
  nodes: DiagramNode[],
  edges: DiagramEdge[],
  canvas: Canvas
): { placed: Map<string, Placed> } {
  const placed = new Map<string, Placed>();
  const nw = DEFAULT_NODE_W;
  const nh = DEFAULT_NODE_H;

  const levels = assignLevelsFromEdges(nodes, edges);

  const byLevel = new Map<number, string[]>();
  for (const n of nodes) {
    const lv = levels.get(n.id) ?? 0;
    if (!byLevel.has(lv)) byLevel.set(lv, []);
    byLevel.get(lv)!.push(n.id);
  }
  const sortedLevels = [...byLevel.keys()].sort((a, b) => a - b);
  const depth = sortedLevels.length;
  const rowGap = depth > 1 ? Math.max(40, (canvas.height - depth * nh) / (depth - 1)) : 0;

  const maxPerRow = Math.max(...[...byLevel.values()].map((ids) => ids.length));
  const uniformGap = maxPerRow > 1 ? Math.max(MIN_GAP, (canvas.width - maxPerRow * nw) / (maxPerRow - 1)) : 0;

  for (let li = 0; li < sortedLevels.length; li++) {
    const ids = byLevel.get(sortedLevels[li])!;
    const y = canvas.top + li * (nh + rowGap);
    const count = ids.length;
    const totalRowW = count * nw + (count - 1) * uniformGap;
    const startX = canvas.left + (canvas.width - totalRowW) / 2;

    ids.forEach((id, j) => {
      placed.set(id, {
        node: nodes.find((n) => n.id === id)!,
        left: startX + j * (nw + uniformGap),
        top: y,
        width: nw,
        height: nh,
      });
    });
  }

  return { placed };
}

function layoutNodes(
  nodes: DiagramNode[],
  edges: DiagramEdge[],
  mode: LayoutMode,
  canvas: Canvas
): { placed: Map<string, Placed> } {
  const placed = new Map<string, Placed>();
  const nw = DEFAULT_NODE_W;
  const nh = DEFAULT_NODE_H;

  if (mode === "vertical") {
    const childCount = new Map<string, number>();
    for (const e of edges) {
      childCount.set(e.from, (childCount.get(e.from) ?? 0) + 1);
    }
    if ([...childCount.values()].some(c => c > 1)) {
      return layoutTree(nodes, edges, canvas);
    }

    const n = nodes.length;
    const totalH = n * nh;
    const gap = n > 1 ? Math.max(20, (canvas.height - totalH) / (n - 1)) : 0;
    const startX = canvas.left + (canvas.width - nw) / 2;
    let y = canvas.top;
    for (const node of nodes) {
      placed.set(node.id, { node, left: startX, top: y, width: nw, height: nh });
      y += nh + gap;
    }
    return { placed };
  }

  if (mode === "horizontal") {
    const n = nodes.length;
    const totalW = n * nw;
    const gap = n > 1 ? Math.max(20, (canvas.width - totalW) / (n - 1)) : 0;
    const startY = canvas.top + (canvas.height - nh) / 2;
    let x = canvas.left;
    for (const node of nodes) {
      placed.set(node.id, { node, left: x, top: startY, width: nw, height: nh });
      x += nw + gap;
    }
    return { placed };
  }

  if (mode === "tree") {
    return layoutTree(nodes, edges, canvas);
  }

  // layered
  const levels = new Map<string, number>();
  if (nodes.every((n) => typeof n.level === "number")) {
    for (const n of nodes) levels.set(n.id, n.level!);
  } else {
    const auto = assignLevelsFromEdges(nodes, edges);
    for (const [k, v] of auto) levels.set(k, v);
  }
  const byLevel = new Map<number, DiagramNode[]>();
  for (const n of nodes) {
    const lv = levels.get(n.id) ?? 0;
    if (!byLevel.has(lv)) byLevel.set(lv, []);
    byLevel.get(lv)!.push(n);
  }
  const sortedLevels = [...byLevel.keys()].sort((a, b) => a - b);
  const depth = sortedLevels.length;
  const rowGap = depth > 1 ? Math.max(40, (canvas.height - depth * nh) / (depth - 1)) : 0;

  sortedLevels.forEach((lv, i) => {
    const row = byLevel.get(lv)!;
    const count = row.length;
    const totalW = count * nw;
    const gap = count > 1 ? Math.max(20, (canvas.width - totalW) / (count - 1)) : 0;
    const startX = count > 1 ? canvas.left : canvas.left + (canvas.width - nw) / 2;
    const y = canvas.top + i * (nh + rowGap);
    row.forEach((node, j) => {
      placed.set(node.id, { node, left: startX + j * (nw + gap), top: y, width: nw, height: nh });
    });
  });
  return { placed };
}

// ── 逐边智能选侧 ──────────────────────────────────────

function pickSides(
  src: Placed, dst: Placed, childCount: number, childIndex: number,
): { from: Side; to: Side } {
  if (childCount > 1) {
    if (childIndex === 0) return { from: "left", to: "top" };
    if (childIndex === childCount - 1) return { from: "right", to: "top" };
    return { from: "bottom", to: "top" };
  }
  const dx = (dst.left + dst.width / 2) - (src.left + src.width / 2);
  const dy = (dst.top + dst.height / 2) - (src.top + src.height / 2);
  if (Math.abs(dy) >= Math.abs(dx)) {
    return dy >= 0 ? { from: "bottom", to: "top" } : { from: "top", to: "bottom" };
  }
  return dx >= 0 ? { from: "right", to: "left" } : { from: "left", to: "right" };
}

export function sidePoint(rect: { left: number; top: number; width: number; height: number }, side: Side): { x: number; y: number } {
  switch (side) {
    case "top":    return { x: rect.left + rect.width / 2, y: rect.top };
    case "bottom": return { x: rect.left + rect.width / 2, y: rect.top + rect.height };
    case "left":   return { x: rect.left, y: rect.top + rect.height / 2 };
    case "right":  return { x: rect.left + rect.width, y: rect.top + rect.height / 2 };
  }
}

// ── createDiagram 工具 ─────────────────────────────────

export const createDiagram: ToolHandler = async (input) => {
  const layout = (input.layout as LayoutMode) ?? "vertical";
  const nodes = (input.nodes as DiagramNode[]) ?? [];
  const edges = (input.edges as DiagramEdge[]) ?? [];
  const canvas: Canvas = (input.canvas as Canvas) ?? DEFAULT_CANVAS;
  if (nodes.length === 0) throw new Error("nodes 不能为空");

  const { placed } = layoutNodes(nodes, edges, layout, canvas);

  const nodePositions = nodes.map((n) => {
    const p = placed.get(n.id)!;
    return `${n.id}(${n.text}): left=${p.left.toFixed(1)} top=${p.top.toFixed(1)} w=${p.width} h=${p.height}`;
  }).join(" | ");

  // 按水平位置排序子节点，以便正确分配 left/right 出线
  const childMap = new Map<string, string[]>();
  for (const e of edges) {
    if (!childMap.has(e.from)) childMap.set(e.from, []);
    childMap.get(e.from)!.push(e.to);
  }
  for (const [, ids] of childMap) {
    ids.sort((a, b) => (placed.get(a)?.left ?? 0) - (placed.get(b)?.left ?? 0));
  }

  const result = await PowerPoint.run(async (ctx) => {
    const slide = await resolveSlide(ctx, slideTargetFromInput(input));
    const idMap: Record<string, string> = {};
    const actualPlaced = new Map<string, Placed>();
    const logLines: string[] = [`[createDiagram] layout=${layout} canvas=${JSON.stringify(canvas)}`];
    logLines.push(`[createDiagram] node positions: ${nodePositions}`);

    // 1. 创建节点形状
    for (const node of nodes) {
      const p = placed.get(node.id)!;
      const shapeType = normalizeGeometricShapeType(node.shape ?? "rectangle");
      const shape = slide.shapes.addGeometricShape(shapeType, {
        left: p.left, top: p.top, width: p.width, height: p.height
      });
      if (node.text) shape.textFrame.textRange.text = node.text;
      shape.load("id,left,top,width,height");
      try {
        await ctx.sync();
      } catch (error) {
        throw withOfficeErrorContext(error, `创建节点 ${node.id} 失败，shape=${node.shape ?? "rectangle"}，normalized=${shapeType}`);
      }
      idMap[node.id] = shape.id;
      actualPlaced.set(node.id, {
        node,
        left: shape.left,
        top: shape.top,
        width: shape.width,
        height: shape.height,
      });
      logLines.push(`[createDiagram] node ${node.id}: planned=(${p.left.toFixed(1)},${p.top.toFixed(1)}) actual=(${shape.left},${shape.top}) id=${shape.id}`);
    }

    // 2. 画占位连接器，随后通过单页 XML export/import 写回真实连接点、肘形几何和箭头
    const edgeLogs: string[] = [];
    let lineSegments = 0;
    let straightConnectors = 0;
    let elbowConnectors = 0;
    let arrowHeads = 0;
    const connectorResults: ConnectorCreationResult[] = [];
    for (const e of edges) {
      const a = actualPlaced.get(e.from);
      const b = actualPlaced.get(e.to);
      if (!a || !b) continue;

      const children = childMap.get(e.from) ?? [];
      const childIdx = children.indexOf(e.to);
      const sides = pickSides(a, b, children.length, childIdx);
      const p1 = sidePoint(a, sides.from);
      const p2 = sidePoint(b, sides.to);

      const connector = drawConnectedLine(slide, idMap[e.from], a, sides.from, idMap[e.to], b, sides.to, { arrow: "end" });
      connectorResults.push(connector);
      lineSegments += connector.lineSegments;
      straightConnectors += connector.straightConnectors;
      elbowConnectors += connector.elbowConnectors;
      arrowHeads += connector.arrows;
      edgeLogs.push(`${e.from}→${e.to}: ${sides.from}→${sides.to} (${p1.x.toFixed(0)},${p1.y.toFixed(0)})→(${p2.x.toFixed(0)},${p2.y.toFixed(0)}), straight=${connector.straightConnectors}, elbow=${connector.elbowConnectors}, arrows=${connector.arrows}, shapes=${connector.shapeCount}`);
    }
    try {
      await ctx.sync();
    } catch (error) {
      throw withOfficeErrorContext(error, "创建连接器占位形状失败");
    }
    const connectorPatches = resolveConnectorXmlPatches(connectorResults);
    logLines.push(`[createDiagram] edges: ${edgeLogs.join(" | ")}`);

    const logMsg = logLines.join("\n");
    fetch("https://localhost:3001/__terminal-log", {
      method: "POST",
      body: JSON.stringify({ level: "info", msg: logMsg }),
      headers: { "Content-Type": "application/json" }
    }).catch(() => {});

    const mapStr = Object.entries(idMap).map(([k, v]) => `${k}=${v}`).join(", ");
    return {
      slideId: slide.id,
      slideIndex: typeof input.slideIndex === "number" ? input.slideIndex as number : undefined,
      pageNumber: typeof input.pageNumber === "number" ? input.pageNumber as number : undefined,
      connectorPatches,
      message: `已创建图：${nodes.length} 节点、${edges.length} 连线（Straight=${straightConnectors}，Elbow=${elbowConnectors}，总连接器=${lineSegments}；原生箭头=${arrowHeads}）。节点 id 映射: ${mapStr}`,
    };
  });
  let editResult: SlideXmlEditResult | null;
  try {
    editResult = await applyConnectorXmlPatches({
      slideId: result.slideId,
      slideIndex: result.slideIndex,
      pageNumber: result.pageNumber,
    }, result.connectorPatches);
  } catch (error) {
    throw new Error(`连接器 XML 修正失败: ${describeOfficeError(error) || String(error)}`);
  }
  const slideText = editResult ? `；新 slideId=${editResult.newSlideId}` : "";
  return `${result.message}。已通过单页 export/import 修正真实 PowerPoint 连接器${slideText}。`;
};
