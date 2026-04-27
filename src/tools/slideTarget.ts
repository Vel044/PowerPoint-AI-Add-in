export interface SlideTarget {
  slideId?: string;
  slideIndex?: number;
  pageNumber?: number;
}

export async function resolveSlide(
  ctx: PowerPoint.RequestContext,
  target: SlideTarget = {},
): Promise<PowerPoint.Slide> {
  const slides = ctx.presentation.slides;
  slides.load("items/id,index");

  if (
    target.slideId === undefined &&
    target.pageNumber === undefined &&
    target.slideIndex === undefined
  ) {
    const selectedSlides = ctx.presentation.getSelectedSlides();
    selectedSlides.load("items/id,index");
    await ctx.sync();
    if (selectedSlides.items.length > 0) return selectedSlides.items[0];
  } else {
    await ctx.sync();
  }

  if (target.slideId) {
    const slide = slides.items.find((item) => item.id === target.slideId);
    if (slide) return slide;
  }

  if (typeof target.pageNumber === "number") {
    const pageIndex = target.pageNumber - 1;
    const slide = slides.items.find((item) => item.index === pageIndex) ?? slides.items[pageIndex];
    if (slide) return slide;
  }

  if (typeof target.slideIndex === "number") {
    const slide = slides.items.find((item) => item.index === target.slideIndex) ?? slides.items[target.slideIndex];
    if (slide) return slide;
  }

  throw new Error(`无法定位目标幻灯片（slideId=${target.slideId ?? "none"}, slideIndex=${target.slideIndex ?? "none"}, pageNumber=${target.pageNumber ?? "none"}）`);
}

export function pageNumberFromIndex(slideIndex: number | null | undefined): number | null {
  return typeof slideIndex === "number" && Number.isFinite(slideIndex) && slideIndex >= 0 ? slideIndex + 1 : null;
}

export function slideTargetFromInput(input: Record<string, unknown>): SlideTarget {
  return {
    slideId: typeof input.slideId === "string" ? input.slideId : undefined,
    slideIndex: typeof input.slideIndex === "number" ? input.slideIndex : undefined,
    pageNumber: typeof input.pageNumber === "number" ? input.pageNumber : undefined,
  };
}
