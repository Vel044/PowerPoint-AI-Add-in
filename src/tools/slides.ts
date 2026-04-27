import { ToolHandler } from "../types";
import { resolveSlide, slideTargetFromInput } from "./slideTarget";

export const addSlide: ToolHandler = async () => {
  return await PowerPoint.run(async (ctx) => {
    ctx.presentation.slides.add();
    await ctx.sync();
    return "已在演示文稿末尾新增 1 张幻灯片";
  });
};

export const deleteSlide: ToolHandler = async (input) => {
  const hasExplicitTarget =
    typeof input.slideId === "string" ||
    typeof input.index === "number" ||
    typeof input.slideIndex === "number" ||
    typeof input.pageNumber === "number";
  if (!hasExplicitTarget) throw new Error("删除幻灯片必须提供 slideId、index/slideIndex 或 pageNumber");
  return await PowerPoint.run(async (ctx) => {
    const target = await resolveSlide(ctx, {
      slideId: typeof input.slideId === "string" ? input.slideId : undefined,
      slideIndex: typeof input.index === "number" ? input.index : typeof input.slideIndex === "number" ? input.slideIndex : undefined,
      pageNumber: typeof input.pageNumber === "number" ? input.pageNumber : undefined,
    });
    target.delete();
    await ctx.sync();
    return `已删除幻灯片 (id=${target.id})`;
  });
};

export const duplicateSlide: ToolHandler = async (input) => {
  const hasExplicitTarget =
    typeof input.slideId === "string" ||
    typeof input.index === "number" ||
    typeof input.slideIndex === "number" ||
    typeof input.pageNumber === "number";
  if (!hasExplicitTarget) throw new Error("复制幻灯片必须提供 slideId、index/slideIndex 或 pageNumber");

  return await PowerPoint.run(async (ctx) => {
    const presentation = ctx.presentation;
    const slide = await resolveSlide(ctx, {
      ...slideTargetFromInput(input),
      slideIndex: typeof input.index === "number" ? input.index as number : slideTargetFromInput(input).slideIndex,
    });
    const slides = presentation.slides;
    slides.load("items/id,index");
    slide.load("id,index");
    const exported = slide.exportAsBase64();
    await ctx.sync();
    if (typeof exported.value !== "string" || exported.value.length === 0) {
      throw new Error("复制失败：exportAsBase64 未返回有效 base64。请确认 Office 支持 PowerPointApi 1.8。");
    }
    const oldIds = new Set(slides.items.map((item) => item.id));
    presentation.insertSlidesFromBase64(exported.value, {
      targetSlideId: slide.id,
      formatting: "KeepSourceFormatting",
    });
    await ctx.sync();
    slides.load("items/id,index");
    await ctx.sync();
    const inserted = slides.items.find((item) => !oldIds.has(item.id));
    if (!inserted) throw new Error("已插入复制页，但无法定位新 slide id。");
    return `已复制幻灯片 oldSlideId=${slide.id}；newSlideId=${inserted.id}；newPageNumber=${inserted.index + 1}`;
  });
};
