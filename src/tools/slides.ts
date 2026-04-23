import { ToolHandler } from "../types";

export const addSlide: ToolHandler = async () => {
  return await PowerPoint.run(async (ctx) => {
    ctx.presentation.slides.add();
    await ctx.sync();
    return "已在演示文稿末尾新增 1 张幻灯片";
  });
};

export const deleteSlide: ToolHandler = async (input) => {
  const index = input.index as number | undefined;
  const id = input.slideId as string | undefined;
  return await PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load("items/id");
    await ctx.sync();
    let target: PowerPoint.Slide | undefined;
    if (id) target = slides.items.find((s) => s.id === id);
    else if (typeof index === "number") target = slides.items[index];
    if (!target) throw new Error("未找到目标幻灯片");
    target.delete();
    await ctx.sync();
    return `已删除幻灯片 (id=${target.id})`;
  });
};
