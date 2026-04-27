import { ToolHandler } from "../types";
import { collectShapeInfos, ShapeInfo } from "./context";

const DEFAULT_SLIDE_WIDTH = 960;
const DEFAULT_SLIDE_HEIGHT = 540;
const MIN_FONT_SHAPE_HEIGHT = 10;

interface VerifyIssue {
  type: string;
  severity: "info" | "warning" | "error";
  desc: string;
  shapeIds?: string[];
}

export const verifySlides: ToolHandler = async (input) => {
  const fromPage = numberOr(input.fromPageNumber, numberOr(input.from_slide, 1));
  const explicitTo = numberOr(input.toPageNumber, numberOr(input.to_slide, NaN));
  const requestedSlideIds = Array.isArray(input.slideIds)
    ? new Set(input.slideIds.filter((id): id is string => typeof id === "string"))
    : null;

  const result = await PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load("items/id,index");
    await ctx.sync();
    const toPage = Number.isFinite(explicitTo) ? explicitTo : slides.items.length;
    const pages = [];
    for (const slide of slides.items) {
      const pageNumber = slide.index + 1;
      if (requestedSlideIds && !requestedSlideIds.has(slide.id)) continue;
      if (!requestedSlideIds && (pageNumber < fromPage || pageNumber > toPage)) continue;
      const shapes = await collectShapeInfos(ctx, slide.shapes);
      const issues = verifyShapeList(shapes);
      pages.push({
        slideId: slide.id,
        slideIndex: slide.index,
        pageNumber,
        ok: !issues.some((issue) => issue.severity === "error"),
        issueCount: issues.length,
        issues,
      });
    }
    return {
      checkedAt: new Date().toISOString(),
      pages,
      ok: pages.every((page) => page.ok),
    };
  });
  return JSON.stringify(result, null, 2);
};

function verifyShapeList(shapes: ShapeInfo[]): VerifyIssue[] {
  const issues: VerifyIssue[] = [];
  const rects = shapes.filter(hasRect);
  for (const shape of rects) {
    const right = shape.left + shape.width;
    const bottom = shape.top + shape.height;
    if (shape.left < 0 || shape.top < 0 || right > DEFAULT_SLIDE_WIDTH || bottom > DEFAULT_SLIDE_HEIGHT) {
      issues.push({
        type: "outOfBounds",
        severity: "warning",
        desc: `形状 ${shape.id} 可能越界: left=${shape.left}, top=${shape.top}, right=${right}, bottom=${bottom}`,
        shapeIds: [shape.id],
      });
    }
    if (shape.text.trim() && shape.height < MIN_FONT_SHAPE_HEIGHT) {
      issues.push({
        type: "tooSmall",
        severity: "warning",
        desc: `形状 ${shape.id} 高度过小，文字可能不可读`,
        shapeIds: [shape.id],
      });
    }
  }

  const nonLineRects = rects.filter((shape) => shape.type !== "Line");
  for (let i = 0; i < nonLineRects.length; i++) {
    for (let j = i + 1; j < nonLineRects.length; j++) {
      const a = nonLineRects[i];
      const b = nonLineRects[j];
      const area = overlapArea(a, b);
      if (area > 120) {
        issues.push({
          type: "overlap",
          severity: "error",
          desc: `形状 ${a.id} 与 ${b.id} 存在明显重叠，重叠面积约 ${Math.round(area)} pt^2`,
          shapeIds: [a.id, b.id],
        });
      }
    }
  }

  for (const shape of nonLineRects) {
    const name = shape.name.toLowerCase();
    const likelyTextShape = shape.type === "TextBox" || name.includes("text") || name.includes("rectangle");
    if (likelyTextShape && !shape.text.trim() && shape.width > 20 && shape.height > 10) {
      issues.push({
        type: "emptyText",
        severity: "info",
        desc: `形状 ${shape.id} 没有文本，可能是装饰元素，也可能是空占位符`,
        shapeIds: [shape.id],
      });
    }
  }
  return issues;
}

function hasRect(shape: ShapeInfo): shape is ShapeInfo & { left: number; top: number; width: number; height: number } {
  return typeof shape.left === "number" &&
    typeof shape.top === "number" &&
    typeof shape.width === "number" &&
    typeof shape.height === "number";
}

function overlapArea(
  a: ShapeInfo & { left: number; top: number; width: number; height: number },
  b: ShapeInfo & { left: number; top: number; width: number; height: number },
): number {
  const x = Math.max(0, Math.min(a.left + a.width, b.left + b.width) - Math.max(a.left, b.left));
  const y = Math.max(0, Math.min(a.top + a.height, b.top + b.height) - Math.max(a.top, b.top));
  return x * y;
}

function numberOr(value: unknown, fallback: number): number {
  return typeof value === "number" && Number.isFinite(value) ? value : fallback;
}
