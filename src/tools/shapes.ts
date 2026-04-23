import { ToolHandler } from "../types";
import { midpoint as edgeMidpoint, Side, drawOrthogonalArrowPath } from "./layout";

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

async function findShapesByIds(
  ctx: PowerPoint.RequestContext,
  ids: string[]
): Promise<{ slide: PowerPoint.Slide; shapes: Record<string, ShapeRect> }> {
  const slides = ctx.presentation.slides;
  slides.load("items/id");
  await ctx.sync();
  const remaining = new Set(ids);
  const found: Record<string, ShapeRect> = {};
  let targetSlide: PowerPoint.Slide | null = null;
  for (const s of slides.items) {
    const shapes = s.shapes;
    shapes.load("items/id,items/left,items/top,items/width,items/height");
    await ctx.sync();
    for (const sh of shapes.items) {
      if (remaining.has(sh.id)) {
        found[sh.id] = { id: sh.id, left: sh.left, top: sh.top, width: sh.width, height: sh.height };
        remaining.delete(sh.id);
        targetSlide = s;
      }
    }
    if (remaining.size === 0) break;
  }
  if (remaining.size > 0) throw new Error(`未找到形状: ${[...remaining].join(", ")}`);
  if (!targetSlide) throw new Error("内部错误：未找到形状所在幻灯片");
  return { slide: targetSlide, shapes: found };
}

export const connectShapes: ToolHandler = async (input) => {
  const fromShapeId = String(input.fromShapeId ?? "");
  const toShapeId = String(input.toShapeId ?? "");
  const fromSide = (input.fromSide as Side) ?? "right";
  const toSide = (input.toSide as Side) ?? "left";
  const arrow = (input.arrow as "none" | "end" | "both") ?? "end";
  if (!fromShapeId || !toShapeId) throw new Error("缺少 fromShapeId 或 toShapeId");

  return await PowerPoint.run(async (ctx) => {
    const { slide, shapes } = await findShapesByIds(ctx, [fromShapeId, toShapeId]);
    const from = shapes[fromShapeId];
    const to = shapes[toShapeId];
    const p1 = edgeMidpoint(from, fromSide);
    const p2 = edgeMidpoint(to, toSide);

    if (arrow === "none") {
      const line = slide.shapes.addLine("elbow" as PowerPoint.ConnectorType, {
        left: p1.x, top: p1.y, width: p2.x - p1.x, height: p2.y - p1.y
      });
      line.load("id");
      await ctx.sync();
      return `已连接 ${fromShapeId}.${fromSide} → ${toShapeId}.${toSide}，elbow 线 (id=${line.id})`;
    }

    drawOrthogonalArrowPath(slide, p1, p2, fromSide, toSide, arrow);
    await ctx.sync();
    return `已连接 ${fromShapeId}.${fromSide} → ${toShapeId}.${toSide}，正交路径`;
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
