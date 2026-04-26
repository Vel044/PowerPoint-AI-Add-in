import { ToolHandler } from "../types";
import { resolveSlide } from "./shapes";

type LayoutMode = "vertical" | "horizontal" | "layered" | "tree";
type NodeShape = "Rectangle" | "roundRectangle" | "diamond" | "ellipse" | "flowChartTerminator" | "flowChartProcess" | "flowChartDecision";
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
const ARROW_SIZE = 12;
const CONNECTOR_STUB = 18;
const CONNECTOR_EPSILON = 0.01;
const ENABLE_CONNECTOR_ARROW_HEADS = false;

interface Point {
  x: number;
  y: number;
}

export interface ConnectorDrawResult {
  lineSegments: number;
  arrows: number;
  shapeCount: number;
  arrowHeadsEnabled: boolean;
}

interface ConnectorOptions {
  arrow?: "none" | "end";
  color?: string;
  thickness?: number;
}

function directionForSide(side: Side): Point {
  switch (side) {
    case "top": return { x: 0, y: -1 };
    case "bottom": return { x: 0, y: 1 };
    case "left": return { x: -1, y: 0 };
    case "right": return { x: 1, y: 0 };
  }
}

function pointsEqual(a: Point, b: Point): boolean {
  return Math.abs(a.x - b.x) < 0.5 && Math.abs(a.y - b.y) < 0.5;
}

function pushPoint(points: Point[], point: Point): void {
  const last = points[points.length - 1];
  if (!last || !pointsEqual(last, point)) points.push(point);
}

function areCollinear(a: Point, b: Point, c: Point): boolean {
  return (Math.abs(a.x - b.x) < 0.5 && Math.abs(b.x - c.x) < 0.5)
    || (Math.abs(a.y - b.y) < 0.5 && Math.abs(b.y - c.y) < 0.5);
}

function simplifyPath(points: Point[]): Point[] {
  const deduped: Point[] = [];
  for (const p of points) pushPoint(deduped, p);
  const simplified: Point[] = [];
  for (const p of deduped) {
    simplified.push(p);
    while (simplified.length >= 3) {
      const n = simplified.length;
      const a = simplified[n - 3];
      const b = simplified[n - 2];
      const c = simplified[n - 1];
      if (!areCollinear(a, b, c)) break;
      simplified.splice(n - 2, 1);
    }
  }
  return simplified;
}

function styleConnectorShape(shape: PowerPoint.Shape, color: string): void {
  shape.fill.setSolidColor(color);
  shape.lineFormat.visible = false;
}

function addStraightConnectorSegment(
  slide: PowerPoint.Slide,
  start: Point,
  end: Point,
  color: string,
  thickness: number,
): boolean {
  const dx = end.x - start.x;
  const dy = end.y - start.y;
  const length = Math.hypot(dx, dy);
  if (length < 0.5) return false;

  const shape = slide.shapes.addLine(PowerPoint.ConnectorType.straight, {
    left: start.x,
    top: start.y,
    width: Math.abs(dx) < 0.5 ? CONNECTOR_EPSILON : dx,
    height: Math.abs(dy) < 0.5 ? CONNECTOR_EPSILON : dy,
  });
  shape.lineFormat.visible = true;
  shape.lineFormat.color = color;
  shape.lineFormat.weight = thickness;
  shape.lineFormat.dashStyle = "Solid";
  shape.lineFormat.style = "Single";
  return true;
}

function addArrowHead(
  slide: PowerPoint.Slide,
  tip: Point,
  previous: Point,
  color: string,
): boolean {
  const dx = tip.x - previous.x;
  const dy = tip.y - previous.y;
  const length = Math.hypot(dx, dy);
  if (length < 0.5) return false;

  const ux = dx / length;
  const uy = dy / length;
  const angle = Math.atan2(dy, dx) * 180 / Math.PI;
  const center = {
    x: tip.x - ux * ARROW_SIZE / 2,
    y: tip.y - uy * ARROW_SIZE / 2,
  };
  const shape = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.triangle, {
    left: center.x - ARROW_SIZE / 2,
    top: center.y - ARROW_SIZE / 2,
    width: ARROW_SIZE,
    height: ARROW_SIZE,
  });
  shape.rotation = angle + 90;
  styleConnectorShape(shape, color);
  return true;
}

function shortenEndForArrow(start: Point, end: Point, arrowSize: number): Point {
  const dx = end.x - start.x;
  const dy = end.y - start.y;
  const length = Math.hypot(dx, dy);
  if (length <= arrowSize + 0.5) return start;
  return {
    x: end.x - dx / length * arrowSize,
    y: end.y - dy / length * arrowSize,
  };
}

function orthogonalPoints(
  start: Point,
  end: Point,
  exitSide: Side,
  entrySide: Side,
  stub: number,
): Point[] {
  if (Math.abs(start.x - end.x) < 0.5 || Math.abs(start.y - end.y) < 0.5) {
    return [start, end];
  }

  const exitDir = directionForSide(exitSide);
  const entryDir = directionForSide(entrySide);
  const startOutside = {
    x: start.x + exitDir.x * stub,
    y: start.y + exitDir.y * stub,
  };
  const endOutside = {
    x: end.x + entryDir.x * stub,
    y: end.y + entryDir.y * stub,
  };

  const points: Point[] = [];
  pushPoint(points, start);
  pushPoint(points, startOutside);

  if (Math.abs(startOutside.x - endOutside.x) >= 0.5 && Math.abs(startOutside.y - endOutside.y) >= 0.5) {
    if (exitSide === "top" || exitSide === "bottom") {
      pushPoint(points, { x: startOutside.x, y: endOutside.y });
    } else {
      pushPoint(points, { x: endOutside.x, y: startOutside.y });
    }
  }

  pushPoint(points, endOutside);
  pushPoint(points, end);
  return simplifyPath(points);
}

function drawConnectorPath(
  slide: PowerPoint.Slide,
  points: Point[],
  options: ConnectorOptions = {},
): ConnectorDrawResult {
  const simplified = simplifyPath(points);
  if (simplified.length < 2) {
    return {
      lineSegments: 0,
      arrows: 0,
      shapeCount: 0,
      arrowHeadsEnabled: ENABLE_CONNECTOR_ARROW_HEADS,
    };
  }

  const color = options.color ?? CONNECTOR_COLOR;
  const thickness = options.thickness ?? CONNECTOR_THICKNESS;
  const arrow = options.arrow ?? "end";
  const shouldDrawArrow = ENABLE_CONNECTOR_ARROW_HEADS && arrow === "end";
  const linePoints = shouldDrawArrow ? simplified.slice() : simplified;
  if (shouldDrawArrow && linePoints.length >= 2) {
    const tip = linePoints[linePoints.length - 1];
    const prev = linePoints[linePoints.length - 2];
    linePoints[linePoints.length - 1] = shortenEndForArrow(prev, tip, ARROW_SIZE);
  }

  let lineSegments = 0;
  for (let i = 0; i < linePoints.length - 1; i++) {
    if (addStraightConnectorSegment(slide, linePoints[i], linePoints[i + 1], color, thickness)) lineSegments++;
  }

  let arrows = 0;
  if (shouldDrawArrow && simplified.length >= 2) {
    const tip = simplified[simplified.length - 1];
    const prev = linePoints[linePoints.length - 1];
    arrows = addArrowHead(slide, tip, prev, color) ? 1 : 0;
  }

  return {
    lineSegments,
    arrows,
    shapeCount: lineSegments + arrows,
    arrowHeadsEnabled: ENABLE_CONNECTOR_ARROW_HEADS,
  };
}

// ── 画线工具：两个模式 ────────────────────────────────

/**
 * 模式 1 — 直连：从 (x1,y1) 到 (x2,y2) 一条直线
 */
export function drawDirectLine(
  slide: PowerPoint.Slide,
  x1: number, y1: number, x2: number, y2: number,
  options: ConnectorOptions = {},
): ConnectorDrawResult {
  return drawConnectorPath(slide, [{ x: x1, y: y1 }, { x: x2, y: y2 }], options);
}

/**
 * 模式 2 — 横平竖直：从 (x1,y1) 到 (x2,y2) 由代码计算正交路径。
 * 每一段路径都使用原生 Straight connector，避免 PowerPoint Elbow 自动路由跑偏。
 */
export function drawOrthogonalLine(
  slide: PowerPoint.Slide,
  x1: number, y1: number, x2: number, y2: number,
  exitSide: Side,
  entrySide: Side = exitSide === "top" ? "bottom" : exitSide === "bottom" ? "top" : exitSide === "left" ? "right" : "left",
  options: ConnectorOptions = {},
): ConnectorDrawResult {
  const points = orthogonalPoints(
    { x: x1, y: y1 },
    { x: x2, y: y2 },
    exitSide,
    entrySide,
    CONNECTOR_STUB,
  );
  return drawConnectorPath(slide, points, options);
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

  return await PowerPoint.run(async (ctx) => {
    const slide = await resolveSlide(ctx, input.slideId as string, input.slideIndex as number);
    const idMap: Record<string, string> = {};
    const actualPlaced = new Map<string, Placed>();
    const logLines: string[] = [`[createDiagram] layout=${layout} canvas=${JSON.stringify(canvas)}`];
    logLines.push(`[createDiagram] node positions: ${nodePositions}`);

    // 1. 创建节点形状
    for (const node of nodes) {
      const p = placed.get(node.id)!;
      const shapeType = (node.shape ?? "Rectangle") as PowerPoint.GeometricShapeType;
      const shape = slide.shapes.addGeometricShape(shapeType, {
        left: p.left, top: p.top, width: p.width, height: p.height
      });
      if (node.text) shape.textFrame.textRange.text = node.text;
      shape.load("id,left,top,width,height");
      await ctx.sync();
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

    // 2. 画连线（受控正交路径，每段使用原生 Straight connector）
    const edgeLogs: string[] = [];
    let lineSegments = 0;
    let arrowHeads = 0;
    for (const e of edges) {
      const a = actualPlaced.get(e.from);
      const b = actualPlaced.get(e.to);
      if (!a || !b) continue;

      const children = childMap.get(e.from) ?? [];
      const childIdx = children.indexOf(e.to);
      const sides = pickSides(a, b, children.length, childIdx);
      const p1 = sidePoint(a, sides.from);
      const p2 = sidePoint(b, sides.to);

      const connector = drawOrthogonalLine(slide, p1.x, p1.y, p2.x, p2.y, sides.from, sides.to, { arrow: "end" });
      lineSegments += connector.lineSegments;
      arrowHeads += connector.arrows;
      edgeLogs.push(`${e.from}→${e.to}: ${sides.from}→${sides.to} (${p1.x.toFixed(0)},${p1.y.toFixed(0)})→(${p2.x.toFixed(0)},${p2.y.toFixed(0)}), segments=${connector.lineSegments}, arrows=${connector.arrows}, shapes=${connector.shapeCount}`);
    }
    await ctx.sync();
    logLines.push(`[createDiagram] edges: ${edgeLogs.join(" | ")}`);

    const logMsg = logLines.join("\n");
    fetch("https://localhost:3001/__terminal-log", {
      method: "POST",
      body: JSON.stringify({ level: "info", msg: logMsg }),
      headers: { "Content-Type": "application/json" }
    }).catch(() => {});

    const mapStr = Object.entries(idMap).map(([k, v]) => `${k}=${v}`).join(", ");
    return `已创建图：${nodes.length} 节点、${edges.length} 连线（原生 Straight 线段=${lineSegments}；箭头头=${arrowHeads}，源码开关当前关闭）。节点 id 映射: ${mapStr}`;
  });
};
