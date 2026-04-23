import { ToolHandler } from "../types";

export const getCurrentContext: ToolHandler = async () => {
  return await PowerPoint.run(async (ctx) => {
    const pres = ctx.presentation;
    const selectedSlides = pres.getSelectedSlides();
    selectedSlides.load("items/id,items/slideId");
    const selectedShapes = pres.getSelectedShapes();
    selectedShapes.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height");
    const allSlides = pres.slides;
    allSlides.load("items/id");
    await ctx.sync();

    const slideIds = allSlides.items.map((s) => s.id);
    const currentSlideIds = selectedSlides.items.map((s) => s.id);
    const currentIndexes = currentSlideIds.map((id) => slideIds.indexOf(id));

    const shapes = selectedShapes.items.map((s) => ({
      id: s.id,
      name: s.name,
      type: s.type,
      left: s.left,
      top: s.top,
      width: s.width,
      height: s.height,
      text: safeText(s)
    }));

    return JSON.stringify(
      {
        totalSlides: slideIds.length,
        selectedSlideIds: currentSlideIds,
        selectedSlideIndexes: currentIndexes,
        selectedShapes: shapes
      },
      null,
      2
    );
  });
};

function safeText(shape: PowerPoint.Shape): string {
  try {
    const tf = (shape as any).textFrame;
    if (!tf) return "";
    return tf.textRange?.text ?? "";
  } catch {
    return "";
  }
}

export const listSlides: ToolHandler = async () => {
  return await PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load("items/id");
    await ctx.sync();
    const out = slides.items.map((s, i) => ({ index: i, id: s.id }));
    return JSON.stringify(out);
  });
};
