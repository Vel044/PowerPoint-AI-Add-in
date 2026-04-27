export interface ShapeStyle {
  fillColor?: string;
  lineColor?: string;
  lineWeight?: number;
  textColor?: string;
  fontSize?: number;
  bold?: boolean;
}

export interface LineStyle {
  color?: string;
  thickness?: number;
  dashStyle?: string;
}

export function shapeStyleFromInput(input: Record<string, unknown>): ShapeStyle {
  return {
    fillColor: stringValue(input.fillColor),
    lineColor: stringValue(input.lineColor),
    lineWeight: positiveNumber(input.lineWeight),
    textColor: stringValue(input.textColor),
    fontSize: positiveNumber(input.fontSize),
    bold: typeof input.bold === "boolean" ? input.bold : undefined,
  };
}

export function lineStyleFromInput(input: Record<string, unknown>): LineStyle {
  return {
    color: stringValue(input.color),
    thickness: positiveNumber(input.thickness),
    dashStyle: stringValue(input.dashStyle),
  };
}

export function applyShapeStyle(shape: PowerPoint.Shape, style: ShapeStyle): void {
  const anyShape = shape as any;
  if (style.fillColor && isTransparentColor(style.fillColor) && anyShape.fill) {
    anyShape.fill.transparency = 1;
  } else if (style.fillColor && anyShape.fill?.setSolidColor) {
    anyShape.fill.setSolidColor(style.fillColor);
    anyShape.fill.transparency = 0;
  }
  if (style.lineColor && isTransparentColor(style.lineColor)) {
    shape.lineFormat.visible = false;
  } else if (style.lineColor) {
    shape.lineFormat.visible = true;
    shape.lineFormat.color = style.lineColor;
  }
  if (typeof style.lineWeight === "number") {
    shape.lineFormat.visible = true;
    shape.lineFormat.weight = style.lineWeight;
  }
  const font = anyShape.textFrame?.textRange?.font;
  if (font) {
    if (style.textColor && !isTransparentColor(style.textColor)) font.color = style.textColor;
    if (typeof style.fontSize === "number") font.size = style.fontSize;
    if (typeof style.bold === "boolean") font.bold = style.bold;
  }
}

function isTransparentColor(value: string): boolean {
  const normalized = value.trim().toLowerCase().replace(/[\s_-]/g, "");
  return normalized === "transparent" || normalized === "none" || normalized === "nofill";
}

export function applyLineStyle(shape: PowerPoint.Shape, style: LineStyle): void {
  if (style.color) shape.lineFormat.color = style.color;
  if (typeof style.thickness === "number") shape.lineFormat.weight = style.thickness;
  if (style.dashStyle) (shape.lineFormat as any).dashStyle = normalizeDashStyle(style.dashStyle);
}

function stringValue(value: unknown): string | undefined {
  return typeof value === "string" && value.trim() ? value.trim() : undefined;
}

function positiveNumber(value: unknown): number | undefined {
  return typeof value === "number" && Number.isFinite(value) && value > 0 ? value : undefined;
}

function normalizeDashStyle(value: string): string {
  const normalized = value.trim().toLowerCase();
  if (normalized === "dash" || normalized === "dashed") return "Dash";
  if (normalized === "dot" || normalized === "dotted") return "Dot";
  if (normalized === "dashdot" || normalized === "dashDot") return "DashDot";
  return "Solid";
}
