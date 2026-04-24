import { ToolHandler } from "../types";
import { resolveSlide } from "./shapes";

type LayoutMode = "vertical" | "horizontal" | "layered" | "tree";
type NodeShape = "rectangle" | "roundRectangle" | "diamond" | "ellipse" | "flowChartTerminator" | "flowChartProcess" | "flowChartDecision";
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
): { placed: Map<string, Placed>; edgeSides: { from: Side; to: Side } } {
  const placed = new Map<string, Placed>();
  const nw = DEFAULT_NODE_W;
  const nh = DEFAULT_NODE_H;

  // Assign levels via BFS
  const levels = assignLevelsFromEdges(nodes, edges);

  // Group by level
  const byLevel = new Map<number, string[]>();
  for (const n of nodes) {
    const lv = levels.get(n.id) ?? 0;
    if (!byLevel.has(lv)) byLevel.set(lv, []);
    byLevel.get(lv)!.push(n.id);
  }
  const sortedLevels = [...byLevel.keys()].sort((a, b) => a - b);
  const depth = sortedLevels.length;
  const rowGap = depth > 1 ? Math.max(40, (canvas.height - depth * nh) / (depth - 1)) : 0;

  // Find max nodes per level to determine uniform column grid
  const maxPerRow = Math.max(...[...byLevel.values()].map((ids) => ids.length));

  // Calculate uniform gap based on widest row
  const uniformGap = maxPerRow > 1 ? Math.max(MIN_GAP, (canvas.width - maxPerRow * nw) / (maxPerRow - 1)) : 0;

  // Place each level uniformly distributed and centered
  for (let li = 0; li < sortedLevels.length; li++) {
    const ids = byLevel.get(sortedLevels[li])!;
    const y = canvas.top + li * (nh + rowGap);
    const count = ids.length;

    // Center this row: distribute `count` nodes evenly using uniformGap
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

  return { placed, edgeSides: { from: "bottom", to: "top" } };
}

function layoutNodes(
  nodes: DiagramNode[],
  edges: DiagramEdge[],
  mode: LayoutMode,
  canvas: Canvas
): { placed: Map<string, Placed>; edgeSides: { from: Side; to: Side } } {
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
    return { placed, edgeSides: { from: "bottom", to: "top" } };
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
    return { placed, edgeSides: { from: "right", to: "left" } };
  }

  if (mode === "tree") {
    return layoutTree(nodes, edges, canvas);
  }

  // layered: lay out levels top-to-bottom, nodes within level left-to-right
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
  return { placed, edgeSides: { from: "bottom", to: "top" } };
}

export const createDiagram: ToolHandler = async (input) => {
  const layout = (input.layout as LayoutMode) ?? "vertical";
  const nodes = (input.nodes as DiagramNode[]) ?? [];
  const edges = (input.edges as DiagramEdge[]) ?? [];
  const canvas: Canvas = (input.canvas as Canvas) ?? DEFAULT_CANVAS;
  if (nodes.length === 0) throw new Error("nodes 不能为空");

  const { placed, edgeSides } = layoutNodes(nodes, edges, layout, canvas);

  // Debug: log all node positions
  const nodePositions = nodes.map((n) => {
    const p = placed.get(n.id)!;
    return `${n.id}(${n.text}): left=${p.left.toFixed(1)} top=${p.top.toFixed(1)} w=${p.width} h=${p.height}`;
  }).join(" | ");

  return await PowerPoint.run(async (ctx) => {
    const slide = await resolveSlide(ctx, input.slideId as string, input.slideIndex as number);
    const idMap: Record<string, string> = {};
    const logLines: string[] = [`[createDiagram] layout=${layout} canvas=${JSON.stringify(canvas)}`];
    logLines.push(`[createDiagram] node positions: ${nodePositions}`);

    // 1. Create node shapes
    for (const node of nodes) {
      const p = placed.get(node.id)!;
      const shapeType = (node.shape ?? "rectangle") as PowerPoint.GeometricShapeType;
      const shape = slide.shapes.addGeometricShape(shapeType, {
        left: p.left, top: p.top, width: p.width, height: p.height
      });
      if (node.text) shape.textFrame.textRange.text = node.text;
      shape.load("id,left,top,width,height");
      await ctx.sync();
      idMap[node.id] = shape.id;
      logLines.push(`[createDiagram] node ${node.id}: planned=(${p.left.toFixed(1)},${p.top.toFixed(1)}) actual=(${shape.left},${shape.top}) id=${shape.id}`);
    }

    // 2. Draw edges as orthogonal polylines (横平竖直)
    const edgeLogs: string[] = [];
    for (const e of edges) {
      const a = placed.get(e.from);
      const b = placed.get(e.to);
      if (!a || !b) continue;
      const p1 = sidePoint(a, edgeSides.from);
      const p2 = sidePoint(b, edgeSides.to);

      const dx = Math.abs(p2.x - p1.x);
      const dy = Math.abs(p2.y - p1.y);

      if (dx < 2) {
        // Vertically aligned: single vertical line
        const line = slide.shapes.addLine(PowerPoint.ConnectorType.straight, {
          left: p1.x, top: p1.y, width: 0, height: p2.y - p1.y,
        });
        line.lineFormat.color = "#333333";
        line.lineFormat.weight = 1.5;
        edgeLogs.push(`${e.from}→${e.to}: vertical (${p1.x.toFixed(0)},${p1.y.toFixed(0)})→(${p2.x.toFixed(0)},${p2.y.toFixed(0)})`);
      } else {
        // Not aligned: draw 3-segment orthogonal polyline (down → horizontal → down)
        const midY = p1.y + (p2.y - p1.y) / 2;
        // Segment 1: source.bottom → midY (vertical)
        const seg1 = slide.shapes.addLine(PowerPoint.ConnectorType.straight, {
          left: p1.x, top: p1.y, width: 0, height: midY - p1.y,
        });
        seg1.lineFormat.color = "#333333";
        seg1.lineFormat.weight = 1.5;
        // Segment 2: midY horizontal from source_x to target_x
        const seg2 = slide.shapes.addLine(PowerPoint.ConnectorType.straight, {
          left: p1.x, top: midY, width: p2.x - p1.x, height: 0,
        });
        seg2.lineFormat.color = "#333333";
        seg2.lineFormat.weight = 1.5;
        // Segment 3: midY → target.top (vertical)
        const seg3 = slide.shapes.addLine(PowerPoint.ConnectorType.straight, {
          left: p2.x, top: midY, width: 0, height: p2.y - midY,
        });
        seg3.lineFormat.color = "#333333";
        seg3.lineFormat.weight = 1.5;
        edgeLogs.push(`${e.from}→${e.to}: orthogonal via midY=${midY.toFixed(0)}`);
      }
    }
    await ctx.sync();
    logLines.push(`[createDiagram] edges: ${edgeLogs.join(" | ")}`);

    // Send debug logs to terminal
    const logMsg = logLines.join("\n");
    fetch("https://localhost:3001/__terminal-log", {
      method: "POST",
      body: JSON.stringify({ level: "info", msg: logMsg }),
      headers: { "Content-Type": "application/json" }
    }).catch(() => {});

    const mapStr = Object.entries(idMap).map(([k, v]) => `${k}=${v}`).join(", ");
    return `已创建图：${nodes.length} 节点、${edges.length} 连线（纯直线，不带箭头）。节点 id 映射: ${mapStr}`;
  });
};

function sidePoint(p: Placed, side: Side): { x: number; y: number } {
  switch (side) {
    case "top": return { x: p.left + p.width / 2, y: p.top };
    case "right": return { x: p.left + p.width, y: p.top + p.height / 2 };
    case "bottom": return { x: p.left + p.width / 2, y: p.top + p.height };
    case "left": return { x: p.left, y: p.top + p.height / 2 };
  }
}

function edgePoint(src: Placed, side: Side, dst: Placed): { x: number; y: number } {
  const base = sidePoint(src, side);
  if (side === "top" || side === "bottom") {
    const srcCx = src.left + src.width / 2;
    const dstCx = dst.left + dst.width / 2;
    const offset = dstCx - srcCx;
    const maxShift = src.width * 0.4;
    base.x = srcCx + Math.max(-maxShift, Math.min(maxShift, offset));
  }
  return base;
}
