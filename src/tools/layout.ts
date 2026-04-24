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

// ── 画线工具：两个模式 ────────────────────────────────

/**
 * 模式 1 — 直连：从 (x1,y1) 到 (x2,y2) 一条直线
 */
export function drawDirectLine(
  slide: PowerPoint.Slide,
  x1: number, y1: number, x2: number, y2: number,
): void {
  slide.shapes.addLine(PowerPoint.ConnectorType.straight, {
    left: Math.min(x1, x2), top: Math.min(y1, y2),
    width: Math.abs(x2 - x1), height: Math.abs(y2 - y1),
  });
}

/**
 * 模式 2 — 横平竖直：从 (x1,y1) 到 (x2,y2) 只走水平/垂直段（L 形两段折线）
 * 每段都用 addLine(straight) + 归一化坐标（width/height ≥ 0）
 * exitSide 决定第一段的方向：top/bottom → 先垂直，left/right → 先水平
 */
export function drawOrthogonalLine(
  slide: PowerPoint.Slide,
  x1: number, y1: number, x2: number, y2: number,
  exitSide: Side,
): void {
  if (Math.abs(x2 - x1) < 2 || Math.abs(y2 - y1) < 2) {
    drawDirectLine(slide, x1, y1, x2, y2);
    return;
  }
  const vert = exitSide === "top" || exitSide === "bottom";
  if (vert) {
    // 段1: 垂直 (x1,y1)→(x1,y2)
    slide.shapes.addLine(PowerPoint.ConnectorType.straight, {
      left: x1, top: Math.min(y1, y2), width: 0, height: Math.abs(y2 - y1),
    });
    // 段2: 水平 (x1,y2)→(x2,y2)
    slide.shapes.addLine(PowerPoint.ConnectorType.straight, {
      left: Math.min(x1, x2), top: y2, width: Math.abs(x2 - x1), height: 0,
    });
  } else {
    // 段1: 水平 (x1,y1)→(x2,y1)
    slide.shapes.addLine(PowerPoint.ConnectorType.straight, {
      left: Math.min(x1, x2), top: y1, width: Math.abs(x2 - x1), height: 0,
    });
    // 段2: 垂直 (x2,y1)→(x2,y2)
    slide.shapes.addLine(PowerPoint.ConnectorType.straight, {
      left: x2, top: Math.min(y1, y2), width: 0, height: Math.abs(y2 - y1),
    });
  }
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
      logLines.push(`[createDiagram] node ${node.id}: planned=(${p.left.toFixed(1)},${p.top.toFixed(1)}) actual=(${shape.left},${shape.top}) id=${shape.id}`);
    }

    // 2. 画连线（横平竖直模式）
    const edgeLogs: string[] = [];
    for (const e of edges) {
      const a = placed.get(e.from);
      const b = placed.get(e.to);
      if (!a || !b) continue;

      const children = childMap.get(e.from) ?? [];
      const childIdx = children.indexOf(e.to);
      const sides = pickSides(a, b, children.length, childIdx);
      const p1 = sidePoint(a, sides.from);
      const p2 = sidePoint(b, sides.to);

      drawOrthogonalLine(slide, p1.x, p1.y, p2.x, p2.y, sides.from);
      edgeLogs.push(`${e.from}→${e.to}: ${sides.from}→${sides.to} (${p1.x.toFixed(0)},${p1.y.toFixed(0)})→(${p2.x.toFixed(0)},${p2.y.toFixed(0)})`);
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
    return `已创建图：${nodes.length} 节点、${edges.length} 连线（横平竖直折线）。节点 id 映射: ${mapStr}`;
  });
};
