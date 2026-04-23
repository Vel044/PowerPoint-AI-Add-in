import { ToolHandler } from "../types";
import { resolveSlide } from "./shapes";

type LayoutMode = "vertical" | "horizontal" | "layered" | "tree";
type NodeShape = "rectangle" | "roundRectangle" | "diamond" | "ellipse" | "flowChartTerminator" | "flowChartProcess" | "flowChartDecision";
type ArrowMode = "none" | "end" | "both";
type Side = "top" | "bottom" | "left" | "right";

interface DiagramNode {
  id: string;
  text: string;
  shape?: NodeShape;
  level?: number;
}

interface DiagramEdge {
  from: string;
  to: string;
  arrow?: ArrowMode;
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

  // layered / tree: both lay out levels top-to-bottom, nodes within level left-to-right
  const levels = new Map<string, number>();
  if (mode === "layered" && nodes.every((n) => typeof n.level === "number")) {
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

  return await PowerPoint.run(async (ctx) => {
    const slide = await resolveSlide(ctx, input.slideId as string, input.slideIndex as number);
    const idMap: Record<string, string> = {};
    const rectMap: Record<string, { left: number; top: number; width: number; height: number }> = {};

    for (const node of nodes) {
      const p = placed.get(node.id)!;
      const shapeType = (node.shape ?? "rectangle") as PowerPoint.GeometricShapeType;
      const shape = slide.shapes.addGeometricShape(shapeType, {
        left: p.left, top: p.top, width: p.width, height: p.height
      });
      if (node.text) shape.textFrame.textRange.text = node.text;
      shape.load("id");
      await ctx.sync();
      idMap[node.id] = shape.id;
      rectMap[node.id] = { left: p.left, top: p.top, width: p.width, height: p.height };
    }

    // edges
    const arrowLog: string[] = [];
    for (const e of edges) {
      const fr = rectMap[e.from];
      const tr = rectMap[e.to];
      if (!fr || !tr) continue;
      const arrow = e.arrow ?? "end";
      // midpoints by layout's natural direction
      const fromSide = edgeSides.from;
      const toSide = edgeSides.to;
      const p1 = midpoint(fr, fromSide);
      const p2 = midpoint(tr, toSide);
      const dx = p2.x - p1.x;
      const dy = p2.y - p1.y;
      const horizontal = Math.abs(dy) < 2;
      const vertical = Math.abs(dx) < 2;
      const orthogonal = horizontal || vertical;

      if (arrow !== "none" && orthogonal) {
        const thickness = 14;
        let shapeType: PowerPoint.GeometricShapeType;
        let left: number, top: number, width: number, height: number;
        if (horizontal) {
          if (dx >= 0) {
            shapeType = "rightArrow" as PowerPoint.GeometricShapeType;
            left = p1.x; top = p1.y - thickness / 2; width = Math.max(dx, 4); height = thickness;
          } else {
            shapeType = "leftArrow" as PowerPoint.GeometricShapeType;
            left = p2.x; top = p1.y - thickness / 2; width = Math.max(-dx, 4); height = thickness;
          }
        } else {
          if (dy >= 0) {
            shapeType = "downArrow" as PowerPoint.GeometricShapeType;
            left = p1.x - thickness / 2; top = p1.y; width = thickness; height = Math.max(dy, 4);
          } else {
            shapeType = "upArrow" as PowerPoint.GeometricShapeType;
            left = p1.x - thickness / 2; top = p2.y; width = thickness; height = Math.max(-dy, 4);
          }
        }
        slide.shapes.addGeometricShape(shapeType, { left, top, width, height });
      } else {
        const connectorType: PowerPoint.ConnectorType = orthogonal
          ? ("straight" as PowerPoint.ConnectorType)
          : ("elbow" as PowerPoint.ConnectorType);
        slide.shapes.addLine(connectorType, {
          left: p1.x, top: p1.y, width: dx, height: dy
        });
        if (arrow !== "none") arrowLog.push(`${e.from}→${e.to}`);
      }
    }
    await ctx.sync();

    const mapStr = Object.entries(idMap).map(([k, v]) => `${k}=${v}`).join(", ");
    const note = arrowLog.length ? ` 斜向连线无箭头头: ${arrowLog.join(", ")}` : "";
    return `已创建图：${nodes.length} 节点、${edges.length} 连线。节点 id 映射: ${mapStr}${note}`;
  });
};

function midpoint(r: { left: number; top: number; width: number; height: number }, side: Side) {
  switch (side) {
    case "top": return { x: r.left + r.width / 2, y: r.top };
    case "bottom": return { x: r.left + r.width / 2, y: r.top + r.height };
    case "left": return { x: r.left, y: r.top + r.height / 2 };
    case "right": return { x: r.left + r.width, y: r.top + r.height / 2 };
  }
}
