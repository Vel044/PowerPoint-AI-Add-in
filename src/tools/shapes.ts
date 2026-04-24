import { ToolHandler } from "../types";
import { Side, drawDirectLine, drawOrthogonalLine, sidePoint } from "./layout";

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

export const connectShapes: ToolHandler = async (input) => {
  const fromShapeId = String(input.fromShapeId ?? "");
  const toShapeId = String(input.toShapeId ?? "");
  const fromSide = (input.fromSide as Side) ?? "right";
  const toSide = (input.toSide as Side) ?? "left";
  const mode = (input.mode as string) ?? "orthogonal";
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
    const p1 = sidePoint(shapes[fromShapeId], fromSide);
    const p2 = sidePoint(shapes[toShapeId], toSide);

    if (mode === "direct") {
      drawDirectLine(slide, p1.x, p1.y, p2.x, p2.y);
    } else {
      drawOrthogonalLine(slide, p1.x, p1.y, p2.x, p2.y, fromSide);
    }
    await ctx.sync();
    return `已连接 ${fromShapeId}.${fromSide} → ${toShapeId}.${toSide}（${mode === "direct" ? "直连" : "横平竖直"}，不带箭头、不跟随形状移动）`;
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
