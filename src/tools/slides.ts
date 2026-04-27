import { ToolHandler } from "../types";
import { resolveSlide } from "./slideTarget";

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
