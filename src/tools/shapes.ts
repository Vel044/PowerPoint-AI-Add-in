import { ToolHandler } from "../types";
import { Side } from "./layout";

export async function resolveSlide(ctx: PowerPoint.RequestContext, slideId?: string, slideIndex?: number) {
  const slides = ctx.presentation.slides;
  slides.load("items/id");
  if (slideId === undefined && slideIndex === undefined) {
    const sel = ctx.presentation.getSelectedSlides();
    sel.load("items/id");
    await ctx.sync();
    if (sel.items.length > 0) return sel.items[0];
  }
  await ctx.sync();
  if (slideId) {
    const s = slides.items.find((x) => x.id === slideId);
    if (s) return s;
  }
  if (typeof slideIndex === "number" && slides.items[slideIndex]) return slides.items[slideIndex];
  throw new Error("无法定位目标幻灯片（未提供 slideId/slideIndex 且无选中）");
}

export const addTextBox: ToolHandler = async (input) => {
  const text = String(input.text ?? "");
  const left = (input.left as number) ?? 50;
  const top = (input.top as number) ?? 50;
  const width = (input.width as number) ?? 400;
  const height = (input.height as number) ?? 80;
  return await PowerPoint.run(async (ctx) => {
    const slide = await resolveSlide(ctx, input.slideId as string, input.slideIndex as number);
    const shape = slide.shapes.addTextBox(text, { left, top, width, height });
    shape.load("id");
    await ctx.sync();
    return `已在幻灯片 ${slide.id} 添加文本框 (id=${shape.id})`;
  });
};

export const modifyShape: ToolHandler = async (input) => {
  const shapeId = String(input.shapeId ?? "");
  if (!shapeId) throw new Error("缺少 shapeId");
  return await PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load("items/id");
    await ctx.sync();
    for (const s of slides.items) {
      const shapes = s.shapes;
      shapes.load("items/id");
      await ctx.sync();
      const target = shapes.items.find((sh) => sh.id === shapeId);
      if (!target) continue;
      if (typeof input.text === "string") {
        target.textFrame.textRange.text = input.text as string;
      }
      if (typeof input.left === "number") target.left = input.left as number;
      if (typeof input.top === "number") target.top = input.top as number;
      if (typeof input.width === "number") target.width = input.width as number;
      if (typeof input.height === "number") target.height = input.height as number;
      await ctx.sync();
      return `已更新形状 ${shapeId}`;
    }
    throw new Error(`未找到形状 ${shapeId}`);
  });
};

export const addGeometricShape: ToolHandler = async (input) => {
  const shapeType = String(input.shapeType ?? "rectangle");
  const text = typeof input.text === "string" ? (input.text as string) : "";
  const left = (input.left as number) ?? 50;
  const top = (input.top as number) ?? 50;
  const width = (input.width as number) ?? 200;
  const height = (input.height as number) ?? 60;
  return await PowerPoint.run(async (ctx) => {
    const slide = await resolveSlide(ctx, input.slideId as string, input.slideIndex as number);
    const shape = slide.shapes.addGeometricShape(shapeType as PowerPoint.GeometricShapeType, {
      left,
      top,
      width,
      height
    });
    if (text) {
      shape.textFrame.textRange.text = text;
    }
    shape.load("id");
    await ctx.sync();
    return `已在幻灯片 ${slide.id} 添加几何形状 ${shapeType} (id=${shape.id})`;
  });
};

export const addLine: ToolHandler = async (input) => {
  const lineType = String(input.lineType ?? "straight");
  const left = (input.left as number) ?? 0;
  const top = (input.top as number) ?? 0;
  const width = (input.width as number) ?? 100;
  const height = (input.height as number) ?? 0;
  return await PowerPoint.run(async (ctx) => {
    const slide = await resolveSlide(ctx, input.slideId as string, input.slideIndex as number);
    const shape = slide.shapes.addLine(lineType as PowerPoint.ConnectorType, { left, top, width, height });
    shape.load("id");
    await ctx.sync();
    return `已在幻灯片 ${slide.id} 添加连接线 ${lineType} (id=${shape.id})`;
  });
};

interface ShapeRect {
  id: string;
  left: number;
  top: number;
  width: number;
  height: number;
}

function sideToPoint(rect: ShapeRect, side: Side): { x: number; y: number } {
  switch (side) {
    case "top": return { x: rect.left + rect.width / 2, y: rect.top };
    case "right": return { x: rect.left + rect.width, y: rect.top + rect.height / 2 };
    case "bottom": return { x: rect.left + rect.width / 2, y: rect.top + rect.height };
    case "left": return { x: rect.left, y: rect.top + rect.height / 2 };
  }
}

export const connectShapes: ToolHandler = async (input) => {
  const fromShapeId = String(input.fromShapeId ?? "");
  const toShapeId = String(input.toShapeId ?? "");
  const fromSide = (input.fromSide as Side) ?? "right";
  const toSide = (input.toSide as Side) ?? "left";
  if (!fromShapeId || !toShapeId) throw new Error("缺少 fromShapeId 或 toShapeId");

  return await PowerPoint.run(async (ctx) => {
    const slide = await resolveSlide(ctx, input.slideId as string, input.slideIndex as number);
    const shapeColl = slide.shapes;
    shapeColl.load("items/id");
    await ctx.sync();
    const idSet = new Set(shapeColl.items.map((s) => s.id));
    const missing = [fromShapeId, toShapeId].filter((id) => !idSet.has(id));
    if (missing.length > 0) {
      throw new Error(`当前幻灯片（id=${slide.id}）上未找到形状: ${missing.join(", ")}。请先调用 get_current_context 确认形状 ID 与所在幻灯片，或传入 slideIndex/slideId。`);
    }
    const fromShape = slide.shapes.getItem(fromShapeId);
    const toShape = slide.shapes.getItem(toShapeId);
    fromShape.load("id,left,top,width,height");
    toShape.load("id,left,top,width,height");
    await ctx.sync();
    const shapes: Record<string, ShapeRect> = {
      [fromShapeId]: { id: fromShapeId, left: fromShape.left, top: fromShape.top, width: fromShape.width, height: fromShape.height },
      [toShapeId]: { id: toShapeId, left: toShape.left, top: toShape.top, width: toShape.width, height: toShape.height },
    };
    for (const id of [fromShapeId, toShapeId]) {
      const r = shapes[id];
      if (typeof r.left !== "number" || typeof r.top !== "number" || typeof r.width !== "number" || typeof r.height !== "number") {
        throw new Error(`形状 ${id} 的位置属性读取失败 (left=${r.left}, top=${r.top}, w=${r.width}, h=${r.height})。可能是占位符或母版继承的形状，无法作为连接端点。`);
      }
    }
    const p1 = sideToPoint(shapes[fromShapeId], fromSide);
    const p2 = sideToPoint(shapes[toShapeId], toSide);
    const dx = Math.abs(p2.x - p1.x);
    const dy = Math.abs(p2.y - p1.y);

    if (dx < 2 || dy < 2) {
      // Aligned: single straight line
      const line = slide.shapes.addLine(PowerPoint.ConnectorType.straight, {
        left: p1.x, top: p1.y, width: p2.x - p1.x, height: p2.y - p1.y,
      });
      line.lineFormat.color = "#333333";
      line.lineFormat.weight = 1.5;
    } else {
      // Not aligned: 3-segment orthogonal polyline
      const midX = p1.x + (p2.x - p1.x) / 2;
      // Segment 1: horizontal from p1 to midX
      const seg1 = slide.shapes.addLine(PowerPoint.ConnectorType.straight, {
        left: p1.x, top: p1.y, width: midX - p1.x, height: 0,
      });
      seg1.lineFormat.color = "#333333";
      seg1.lineFormat.weight = 1.5;
      // Segment 2: vertical from midY(p1.y) to p2.y
      const seg2 = slide.shapes.addLine(PowerPoint.ConnectorType.straight, {
        left: midX, top: p1.y, width: 0, height: p2.y - p1.y,
      });
      seg2.lineFormat.color = "#333333";
      seg2.lineFormat.weight = 1.5;
      // Segment 3: horizontal from midX to p2
      const seg3 = slide.shapes.addLine(PowerPoint.ConnectorType.straight, {
        left: midX, top: p2.y, width: p2.x - midX, height: 0,
      });
      seg3.lineFormat.color = "#333333";
      seg3.lineFormat.weight = 1.5;
    }
    await ctx.sync();
    return `已连接 ${fromShapeId}.${fromSide} → ${toShapeId}.${toSide}（横平竖直折线，不带箭头、不跟随形状移动）`;
  });
};

export const deleteShape: ToolHandler = async (input) => {
  const shapeId = String(input.shapeId ?? "");
  if (!shapeId) throw new Error("缺少 shapeId");
  return await PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load("items/id");
    await ctx.sync();
    for (const s of slides.items) {
      const shapes = s.shapes;
      shapes.load("items/id");
      await ctx.sync();
      const target = shapes.items.find((sh) => sh.id === shapeId);
      if (!target) continue;
      target.delete();
      await ctx.sync();
      return `已删除形状 ${shapeId}`;
    }
    throw new Error(`未找到形状 ${shapeId}`);
  });
};
