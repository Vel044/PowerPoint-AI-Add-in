import { SlideTarget } from "./slideTarget";

export interface ShapeRef {
  slideId: string;
  shapeId: string;
}

export function makeShapeRef(slideId: string, shapeId: string): string {
  return `${slideId}:${shapeId}`;
}

export function parseShapeRef(value: unknown): ShapeRef | null {
  if (typeof value !== "string") return null;
  const index = value.lastIndexOf(":");
  if (index <= 0 || index >= value.length - 1) return null;
  return {
    slideId: value.slice(0, index),
    shapeId: value.slice(index + 1),
  };
}

export function targetFromRefOrInput(input: Record<string, unknown>): { target: SlideTarget; shapeId: string } {
  const ref = parseShapeRef(input.ref);
  if (ref) return { target: { slideId: ref.slideId }, shapeId: ref.shapeId };
  const shapeId = typeof input.shapeId === "string" ? input.shapeId : "";
  return {
    target: {
      slideId: typeof input.slideId === "string" ? input.slideId : undefined,
      slideIndex: typeof input.slideIndex === "number" ? input.slideIndex : undefined,
      pageNumber: typeof input.pageNumber === "number" ? input.pageNumber : undefined,
    },
    shapeId,
  };
}
