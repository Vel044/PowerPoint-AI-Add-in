import { ToolHandler } from "../types";
import { resolveSlide } from "./shapes";

type LayoutMode = "vertical" | "horizontal" | "layered" | "tree";
type NodeShape = "rectangle" | "roundRectangle" | "diamond" | "ellipse" | "flowChartTerminator" | "flowChartProcess" | "flowChartDecision";
type ArrowMode = "none" | "end" | "both";
export type Side = "top" | "bottom" | "left" | "right";

export interface Rect { left: number; top: number; width: number; height: number; }

const ARROW_THICKNESS = 14;
const LINE_THICKNESS = 3;

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

export function drawOrthogonalArrowPath(
  slide: PowerPoint.Slide,
  p1: { x: number; y: number },
  p2: { x: number; y: number },
  fromSide: Side,
  toSide: Side,
  arrowMode: "end" | "both"
): void {
  const verticalPrimary = fromSide === "top" || fromSide === "bottom";

  if (verticalPrimary) {
    const midY = Math.round((p1.y + p2.y) / 2);

    // Segment 1: vertical shaft from p1 to midY
    const seg1Len = Math.abs(midY - p1.y);
    if (seg1Len > 1) {
      const top = Math.min(p1.y, midY);
      if (arrowMode === "both") {
        const isDown = midY >= p1.y;
        const arrowType = isDown ? "downArrow" : "upArrow";
        const arrowTop = isDown ? p1.y : midY;
        slide.shapes.addGeometricShape(arrowType as PowerPoint.GeometricShapeType, {
          left: p1.x - ARROW_THICKNESS / 2, top: arrowTop, width: ARROW_THICKNESS, height: Math.max(seg1Len, 4)
        });
      } else {
        slide.shapes.addGeometricShape("rectangle" as PowerPoint.GeometricShapeType, {
          left: p1.x - LINE_THICKNESS / 2, top, width: LINE_THICKNESS, height: Math.max(seg1Len, 0.5)
        });
      }
    }

    // Segment 2: horizontal shaft at midY
    const seg2Len = Math.abs(p2.x - p1.x);
    if (seg2Len > 1) {
      slide.shapes.addGeometricShape("rectangle" as PowerPoint.GeometricShapeType, {
        left: Math.min(p1.x, p2.x), top: midY - LINE_THICKNESS / 2,
        width: Math.max(seg2Len, 0.5), height: LINE_THICKNESS
      });
    }

    // Segment 3: vertical arrow from midY to p2
    const seg3Len = Math.abs(p2.y - midY);
    if (seg3Len > 1) {
      const isDown = p2.y >= midY;
      const arrowType = isDown ? "downArrow" : "upArrow";
      const arrowTop = isDown ? midY : p2.y;
      slide.shapes.addGeometricShape(arrowType as PowerPoint.GeometricShapeType, {
        left: p2.x - ARROW_THICKNESS / 2, top: arrowTop, width: ARROW_THICKNESS, height: Math.max(seg3Len, 4)
      });
    }
  } else {
    const midX = Math.round((p1.x + p2.x) / 2);

    // Segment 1: horizontal shaft from p1 to midX
    const seg1Len = Math.abs(midX - p1.x);
    if (seg1Len > 1) {
      const left = Math.min(p1.x, midX);
      if (arrowMode === "both") {
        const isRight = midX >= p1.x;
        const arrowType = isRight ? "rightArrow" : "leftArrow";
        const arrowLeft = isRight ? p1.x : midX;
        slide.shapes.addGeometricShape(arrowType as PowerPoint.GeometricShapeType, {
          left: arrowLeft, top: p1.y - ARROW_THICKNESS / 2,
          width: Math.max(seg1Len, 4), height: ARROW_THICKNESS
        });
      } else {
        slide.shapes.addGeometricShape("rectangle" as PowerPoint.GeometricShapeType, {
          left, top: p1.y - LINE_THICKNESS / 2,
          width: Math.max(seg1Len, 0.5), height: LINE_THICKNESS
        });
      }
    }

    // Segment 2: vertical shaft at midX
    const seg2Len = Math.abs(p2.y - p1.y);
    if (seg2Len > 1) {
      slide.shapes.addGeometricShape("rectangle" as PowerPoint.GeometricShapeType, {
        left: midX - LINE_THICKNESS / 2, top: Math.min(p1.y, p2.y),
        width: LINE_THICKNESS, height: Math.max(seg2Len, 0.5)
      });
    }

    // Segment 3: horizontal arrow from midX to p2
    const seg3Len = Math.abs(p2.x - midX);
    if (seg3Len > 1) {
      const isRight = p2.x >= midX;
      const arrowType = isRight ? "rightArrow" : "leftArrow";
      const arrowLeft = isRight ? midX : p2.x;
      slide.shapes.addGeometricShape(arrowType as PowerPoint.GeometricShapeType, {
        left: arrowLeft, top: p2.y - ARROW_THICKNESS / 2,
        width: Math.max(seg3Len, 4), height: ARROW_THICKNESS
      });
    }
  }
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
    for (const e of edges) {
      const fr = rectMap[e.from];
      const tr = rectMap[e.to];
      if (!fr || !tr) continue;
      const arrow = e.arrow ?? "end";
      const fromSide = edgeSides.from;
      const toSide = edgeSides.to;
      const p1 = midpoint(fr, fromSide);
      const p2 = midpoint(tr, toSide);

      if (arrow === "none") {
        slide.shapes.addLine("elbow" as PowerPoint.ConnectorType, {
          left: p1.x, top: p1.y, width: p2.x - p1.x, height: p2.y - p1.y
        });
      } else {
        drawOrthogonalArrowPath(slide, p1, p2, fromSide, toSide, arrow);
      }
    }
    await ctx.sync();

    const mapStr = Object.entries(idMap).map(([k, v]) => `${k}=${v}`).join(", ");
    return `已创建图：${nodes.length} 节点、${edges.length} 连线。节点 id 映射: ${mapStr}`;
  });
};

export function midpoint(r: { left: number; top: number; width: number; height: number }, side: Side) {
  switch (side) {
    case "top": return { x: r.left + r.width / 2, y: r.top };
    case "bottom": return { x: r.left + r.width / 2, y: r.top + r.height };
    case "left": return { x: r.left, y: r.top + r.height / 2 };
    case "right": return { x: r.left + r.width, y: r.top + r.height / 2 };
  }
}
