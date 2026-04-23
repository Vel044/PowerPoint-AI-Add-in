import { ToolHandler } from "../types";

async function resolveSlide(ctx: PowerPoint.RequestContext, slideId?: string, slideIndex?: number) {
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
