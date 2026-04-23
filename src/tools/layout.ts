import { ToolHandler } from "../types";
import { resolveSlide } from "./shapes";

type LayoutMode = "vertical" | "horizontal" | "layered" | "tree";
type NodeShape = "rectangle" | "roundRectangle" | "diamond" | "ellipse" | "flowChartTerminator" | "flowChartProcess" | "flowChartDecision";
type ArrowMode = "none" | "end" | "both";
export type Side = "top" | "bottom" | "left" | "right";

export interface Rect { left: number; top: number; width: number; height: number; }

const LINE_COLOR = "#333333";
const LINE_WEIGHT = 1.5;

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

const MIN_GAP = 20;

function resolveOverlaps(ids: string[], placed: Map<string, Placed>, nw: number): void {
  const sorted = [...ids].sort((a, b) => placed.get(a)!.left - placed.get(b)!.left);
  for (let j = 1; j < sorted.length; j++) {
    const prev = placed.get(sorted[j - 1])!;
    const curr = placed.get(sorted[j])!;
    const minLeft = prev.left + nw + MIN_GAP;
    if (curr.left < minLeft) {
      const shift = minLeft - curr.left;
      for (let k = j; k < sorted.length; k++) {
        placed.get(sorted[k])!.left += shift;
      }
    }
  }
}

function centerOnCanvas(ids: string[], placed: Map<string, Placed>, canvas: Canvas, nw: number): void {
  if (ids.length === 0) return;
  const sorted = [...ids].sort((a, b) => placed.get(a)!.left - placed.get(b)!.left);
  const leftmost = placed.get(sorted[0])!.left;
  const rightmost = placed.get(sorted[sorted.length - 1])!.left + nw;
  const span = rightmost - leftmost;
  const idealStart = canvas.left + (canvas.width - span) / 2;
  const offset = idealStart - leftmost;
  if (Math.abs(offset) > 0.5) {
    for (const id of ids) placed.get(id)!.left += offset;
  }
}

function layoutTree(
  nodes: DiagramNode[],
  edges: DiagramEdge[],
  canvas: Canvas
): { placed: Map<string, Placed>; edgeSides: { from: Side; to: Side } } {
  const placed = new Map<string, Placed>();
  const nw = DEFAULT_NODE_W;
  const nh = DEFAULT_NODE_H;

  // Build parent/children maps
  const parentMap = new Map<string, string[]>();
  const childrenMap = new Map<string, string[]>();
  for (const n of nodes) {
    parentMap.set(n.id, []);
    childrenMap.set(n.id, []);
  }
  for (const e of edges) {
    if (e.from === e.to) continue;
    parentMap.get(e.to)?.push(e.from);
    childrenMap.get(e.from)?.push(e.to);
  }

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

  // Node lookup
  const nodeMap = new Map(nodes.map((n) => [n.id, n]));

  // Process levels top-to-bottom
  for (let li = 0; li < sortedLevels.length; li++) {
    const lv = sortedLevels[li];
    const ids = byLevel.get(lv)!;
    const y = canvas.top + li * (nh + rowGap);

    if (li === 0) {
      // Root level: distribute evenly, centered on canvas
      const count = ids.length;
      const gap = count > 1 ? Math.max(MIN_GAP, (canvas.width - count * nw) / (count - 1)) : 0;
      const startX = count > 1
        ? canvas.left + (canvas.width - (count * nw + (count - 1) * gap)) / 2
        : canvas.left + (canvas.width - nw) / 2;
      ids.forEach((id, j) => {
        placed.set(id, { node: nodeMap.get(id)!, left: startX + j * (nw + gap), top: y, width: nw, height: nh });
      });
    } else {
      // Non-root: position children centered under their parent(s)
      // Group by parent, using average parent x for multi-parent nodes
      const parentOf = new Map<string, string>();
      for (const id of ids) {
        const parents = parentMap.get(id) ?? [];
        // Use the parent that was already placed and has the most specific relationship
        const placedParent = parents.find((p) => placed.has(p));
        if (placedParent) parentOf.set(id, placedParent);
      }

      // Group children by parent
      const groups = new Map<string, string[]>();
      const unparented: string[] = [];
      for (const id of ids) {
        const p = parentOf.get(id);
        if (p) {
          if (!groups.has(p)) groups.set(p, []);
          groups.get(p)!.push(id);
        } else {
          unparented.push(id);
        }
      }

      // Sort groups by parent x position
      const sortedGroups = [...groups.entries()].sort((a, b) => {
        return placed.get(a[0])!.left - placed.get(b[0])!.left;
      });

      // Place each group centered under its parent
      const childGap = MIN_GAP;
      for (const [parentId, childIds] of sortedGroups) {
        const p = placed.get(parentId)!;
        const parentCenterX = p.left + nw / 2;
        const groupW = childIds.length * nw + (childIds.length - 1) * childGap;
        let startX = parentCenterX - groupW / 2;
        for (let j = 0; j < childIds.length; j++) {
          placed.set(childIds[j], {
            node: nodeMap.get(childIds[j])!,
            left: startX + j * (nw + childGap),
            top: y,
            width: nw,
            height: nh
          });
        }
      }

      // Place unparented nodes after groups
      if (unparented.length > 0) {
        const rightEdge = [...placed.values()]
          .filter((p) => Math.abs(p.top - y) < 1)
          .reduce((max, p) => Math.max(max, p.left + nw), canvas.left);
        let startX = rightEdge + MIN_GAP;
        for (const id of unparented) {
          placed.set(id, { node: nodeMap.get(id)!, left: startX, top: y, width: nw, height: nh });
          startX += nw + childGap;
        }
      }

      // Resolve overlaps
      resolveOverlaps(ids, placed, nw);
      // Center on canvas
      centerOnCanvas(ids, placed, canvas, nw);
    }
  }

  // Upward feedback: adjust parents toward children center (2 rounds)
  for (let round = 0; round < 2; round++) {
    for (let li = sortedLevels.length - 1; li > 0; li--) {
      const lv = sortedLevels[li];
      const ids = byLevel.get(lv)!;
      // Group by parent again
      const groups = new Map<string, string[]>();
      for (const id of ids) {
        const parents = parentMap.get(id) ?? [];
        for (const p of parents) {
          if (placed.has(p)) {
            if (!groups.has(p)) groups.set(p, []);
            groups.get(p)!.push(id);
          }
        }
      }
      for (const [parentId, childIds] of groups.entries()) {
        const parentPlaced = placed.get(parentId)!;
        const childCenters = childIds.map((cid) => placed.get(cid)!);
        const avgChildX = childCenters.reduce((s, c) => s + c.left + nw / 2, 0) / childCenters.length;
        const targetX = avgChildX - nw / 2;
        // Move parent partially toward children center
        parentPlaced.left = parentPlaced.left + (targetX - parentPlaced.left) * 0.3;
      }
      // Re-resolve overlaps for parent level
      const parentLv = sortedLevels[li - 1];
      resolveOverlaps(byLevel.get(parentLv) ?? [], placed, nw);
      centerOnCanvas(byLevel.get(parentLv) ?? [], placed, canvas, nw);
    }
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

const ARROW_TRI_SIZE = 8;

function drawArrowHead(slide: PowerPoint.Slide, tipX: number, tipY: number, dir: "down" | "up" | "left" | "right"): void {
  const s = ARROW_TRI_SIZE;
  let left: number, top: number, rotation: number;
  switch (dir) {
    case "down":
      left = tipX - s / 2;
      top = tipY - s;
      rotation = 180;
      break;
    case "up":
      left = tipX - s / 2;
      top = tipY;
      rotation = 0;
      break;
    case "right":
      left = tipX - s;
      top = tipY - s / 2;
      rotation = 90;
      break;
    case "left":
      left = tipX;
      top = tipY - s / 2;
      rotation = -90;
      break;
  }
  const shape = slide.shapes.addGeometricShape("triangle" as PowerPoint.GeometricShapeType, {
    left, top, width: s, height: s
  });
  shape.rotation = rotation;
  shape.fill.setSolidColor(LINE_COLOR);
  shape.lineFormat.visible = false;
}

export function drawOrthogonalArrowPath(
  slide: PowerPoint.Slide,
  p1: { x: number; y: number },
  p2: { x: number; y: number },
  fromSide: Side,
  toSide: Side,
  arrowMode: "end" | "both"
): void {
  const w = p2.x - p1.x;
  const h = p2.y - p1.y;
  if (Math.abs(w) > 0.5 || Math.abs(h) > 0.5) {
    const line = slide.shapes.addLine("elbow" as PowerPoint.ConnectorType, {
      left: p1.x, top: p1.y,
      width: Math.abs(w) < 0.5 ? 0.5 : w,
      height: Math.abs(h) < 0.5 ? 0.5 : h
    });
    line.lineFormat.weight = LINE_WEIGHT;
    line.lineFormat.color = LINE_COLOR;
  }

  const toDirMap: Record<Side, "down" | "up" | "left" | "right"> = {
    top: "down", bottom: "up", left: "right", right: "left"
  };
  const fromDirMap: Record<Side, "down" | "up" | "left" | "right"> = {
    top: "up", bottom: "down", left: "left", right: "right"
  };

  if (arrowMode === "end" || arrowMode === "both") {
    drawArrowHead(slide, p2.x, p2.y, toDirMap[toSide]);
  }
  if (arrowMode === "both") {
    drawArrowHead(slide, p1.x, p1.y, fromDirMap[fromSide]);
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
