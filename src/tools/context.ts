import { ToolHandler } from "../types";
import { pageNumberFromIndex, resolveSlide, slideTargetFromInput } from "./slideTarget";

export const getCurrentContext: ToolHandler = async (input) => {
  try {
    return await PowerPoint.run(async (ctx) => {
      const pres = ctx.presentation;
      const selectedSlides = pres.getSelectedSlides();
      selectedSlides.load("items/id,index");
      const allSlides = pres.slides;
      allSlides.load("items/id,index");
      await ctx.sync();

      const slideIds = allSlides.items.map((s) => s.id);
      const currentSlideIds = selectedSlides.items.map((s) => s.id);
      const currentIndexes = currentSlideIds.map((id) => slideIds.indexOf(id));
      const currentPageNumbers = currentIndexes.map((index) => pageNumberFromIndex(index)).filter((n): n is number => n !== null);
      const currentSlide = selectedSlides.items[0] ?? null;
      const currentSlideId = currentSlide?.id ?? null;
      const currentSlideIndex = currentSlideId ? slideIds.indexOf(currentSlideId) : null;
      const currentPageNumber = pageNumberFromIndex(currentSlideIndex);
      const requestedTarget = slideTargetFromInput(input);
      const shouldInspectTarget =
        requestedTarget.slideId !== undefined ||
        requestedTarget.slideIndex !== undefined ||
        requestedTarget.pageNumber !== undefined;
      const inspectedSlide = shouldInspectTarget
        ? await resolveSlide(ctx, requestedTarget)
        : currentSlide;
      const inspectedSlideId = inspectedSlide?.id ?? null;
      const inspectedSlideIndex = inspectedSlideId ? slideIds.indexOf(inspectedSlideId) : null;
      const inspectedPageNumber = pageNumberFromIndex(inspectedSlideIndex);

      const allShapes = inspectedSlide
        ? (await collectShapeInfos(ctx, inspectedSlide.shapes)).map((shape) => ({
            slideId: inspectedSlideId,
            slideIndex: inspectedSlideIndex,
            pageNumber: inspectedPageNumber,
            ...shape
          }))
        : [];
      const selectedShapes = await collectSelectedShapeInfos(ctx, pres);

      return JSON.stringify(
        {
          totalSlides: slideIds.length,
          selectedSlideIds: currentSlideIds,
          selectedSlideIndexes: currentIndexes,
          selectedPageNumbers: currentPageNumbers,
          currentSlideId,
          currentSlideIndex,
          currentPageNumber,
          inspectedSlideId,
          inspectedSlideIndex,
          inspectedPageNumber,
          occupiedBounds: computeOccupiedBounds(allShapes),
          allShapes,
          selectedShapes
        },
        null,
        2
      );
    });
  } catch (err: any) {
    const msg = err?.message ?? String(err);
    if (msg.includes("0x80070057") || (err?.code === 0x80070057)) {
      return JSON.stringify({
        error: "获取上下文失败 (0x80070057)",
        detail: "线条和连接器形状可能不被当前操作支持，或选中对象不是有效形状。",
        suggestion: "请尝试选中一个普通形状（如文本框、矩形）后重试。"
      }, null, 2);
    }
    throw err;
  }
};

type ShapeInfo = {
  slideId?: string | null;
  slideIndex?: number | null;
  pageNumber?: number | null;
  id: string;
  name: string;
  type: string | null;
  text: string;
  left: number | null;
  top: number | null;
  width: number | null;
  height: number | null;
};

async function collectSelectedShapeInfos(ctx: PowerPoint.RequestContext, pres: PowerPoint.Presentation): Promise<ShapeInfo[]> {
  try {
    const selectedShapes = pres.getSelectedShapes();
    return await collectShapeInfos(ctx, selectedShapes);
  } catch {
    return [];
  }
}

async function collectShapeInfos(
  ctx: PowerPoint.RequestContext,
  collection: { load: (propertyNames?: string | string[]) => unknown; items: PowerPoint.Shape[] }
): Promise<ShapeInfo[]> {
  try {
    collection.load("items/id,name");
    await ctx.sync();
  } catch {
    return [];
  }

  const infos: ShapeInfo[] = [];
  for (const shape of collection.items) {
    infos.push(await collectShapeInfo(ctx, shape));
  }
  return infos;
}

async function collectShapeInfo(ctx: PowerPoint.RequestContext, shape: PowerPoint.Shape): Promise<ShapeInfo> {
  const info: ShapeInfo = {
    id: safeString(() => shape.id),
    name: safeString(() => shape.name),
    type: null,
    text: "",
    left: null,
    top: null,
    width: null,
    height: null
  };

  try {
    shape.load("id,name,type,left,top,width,height");
    await ctx.sync();
  } catch {
    try {
      shape.load("id,name,left,top,width,height");
      await ctx.sync();
    } catch {
      // Keep already loaded id/name and leave geometry fields null.
    }
  }

  info.id = safeString(() => shape.id) || info.id;
  info.name = safeString(() => shape.name) || info.name;
  info.type = safeValue(() => (shape as any).type) as string | null;
  info.left = safeNumber(() => shape.left);
  info.top = safeNumber(() => shape.top);
  info.width = safeNumber(() => shape.width);
  info.height = safeNumber(() => shape.height);
  info.text = await safeText(ctx, shape);
  return info;
}

function safeValue(read: () => unknown): unknown {
  try {
    return read();
  } catch {
    return null;
  }
}

function safeString(read: () => unknown): string {
  try {
    const value = read();
    return typeof value === "string" ? value : "";
  } catch {
    return "";
  }
}

function safeNumber(read: () => unknown): number | null {
  try {
    const value = read();
    return typeof value === "number" && Number.isFinite(value) ? value : null;
  } catch {
    return null;
  }
}

async function safeText(ctx: PowerPoint.RequestContext, shape: PowerPoint.Shape): Promise<string> {
  try {
    const textRange = (shape as any).textFrame?.textRange;
    if (!textRange) return "";
    textRange.load("text");
    await ctx.sync();
    return typeof textRange.text === "string" ? textRange.text : "";
  } catch {
    return "";
  }
}

function computeOccupiedBounds(shapes: ShapeInfo[]) {
  const rects = shapes.filter((s) =>
    typeof s.left === "number" &&
    typeof s.top === "number" &&
    typeof s.width === "number" &&
    typeof s.height === "number"
  );
  if (rects.length === 0) return null;
  const left = Math.min(...rects.map((s) => s.left as number));
  const top = Math.min(...rects.map((s) => s.top as number));
  const right = Math.max(...rects.map((s) => (s.left as number) + (s.width as number)));
  const bottom = Math.max(...rects.map((s) => (s.top as number) + (s.height as number)));
  return {
    left,
    top,
    width: right - left,
    height: bottom - top,
    shapeCount: rects.length
  };
}

export const listSlides: ToolHandler = async () => {
  return await PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load("items/id,index");
    await ctx.sync();
    const out = slides.items.map((s, i) => ({ index: i, pageNumber: i + 1, id: s.id }));
    return JSON.stringify(out);
  });
};
