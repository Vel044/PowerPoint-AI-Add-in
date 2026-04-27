import { ToolHandler } from "../types";
import { withOfficeErrorContext } from "./officeErrors";
import { applyShapeStyle, shapeStyleFromInput } from "./shapeStyle";
import { normalizeGeometricShapeType } from "./shapeTypes";
import {
  applyConnectorXmlPatches,
  ConnectorCreationResult,
  drawConnectedLine,
  resolveConnectorXmlPatches,
  Side
} from "./layout";
import { resolveSlide, slideTargetFromInput } from "./slideTarget";

interface DrawShapeInput extends Record<string, unknown> {
  id?: string;
  type?: string;
  shapeType?: string;
  text?: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
}

interface DrawConnectorInput extends Record<string, unknown> {
  from?: string;
  fromSide?: Side;
  to?: string;
  toSide?: Side;
  mode?: string;
  arrow?: string;
}

interface DrawnShapeRect {
  id: string;
  left: number;
  top: number;
  width: number;
  height: number;
}

export const drawSlideShapes: ToolHandler = async (input) => {
  const shapes = Array.isArray(input.shapes) ? input.shapes as DrawShapeInput[] : [];
  const connectors = Array.isArray(input.connectors) ? input.connectors as DrawConnectorInput[] : [];
  const title = isObject(input.title) ? input.title as DrawShapeInput : null;
  if (shapes.length === 0 && !title) throw new Error("draw_slide_shapes 需要至少一个 shape 或 title");

  const result = await PowerPoint.run(async (ctx) => {
    const slide = await resolveSlide(ctx, slideTargetFromInput(input));
    const idMap: Record<string, string> = {};
    const rects = new Map<string, DrawnShapeRect>();
    const created: PowerPoint.Shape[] = [];

    if (title && typeof title.text === "string") {
      const titleId = typeof title.id === "string" && title.id ? title.id : "__title";
      const shape = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
        left: numberOr(title.left, 40),
        top: numberOr(title.top, 40),
        width: numberOr(title.width, 880),
        height: numberOr(title.height, 36),
      });
      shape.textFrame.textRange.text = title.text;
      applyShapeStyle(shape, {
        fillColor: typeof title.fillColor === "string" ? title.fillColor : "#173B5F",
        lineColor: typeof title.lineColor === "string" ? title.lineColor : "#173B5F",
        textColor: typeof title.textColor === "string" ? title.textColor : "#FFFFFF",
        fontSize: typeof title.fontSize === "number" ? title.fontSize : 14,
        bold: typeof title.bold === "boolean" ? title.bold : true,
      });
      shape.load("id,left,top,width,height");
      created.push(shape);
      idMap[titleId] = "";
    }

    for (const item of shapes) {
      const id = requireId(item.id);
      const left = numberOr(item.left, 50);
      const top = numberOr(item.top, 50);
      const width = numberOr(item.width, 200);
      const height = numberOr(item.height, 60);
      const text = typeof item.text === "string" ? item.text : "";
      const type = typeof item.type === "string" ? item.type : "geometricShape";
      const shape = type === "textBox"
        ? slide.shapes.addTextBox(text, { left, top, width, height })
        : slide.shapes.addGeometricShape(normalizeGeometricShapeType(item.shapeType ?? "rectangle"), { left, top, width, height });
      if (text && type !== "textBox") shape.textFrame.textRange.text = text;
      applyShapeStyle(shape, shapeStyleFromInput(item));
      shape.load("id,left,top,width,height");
      created.push(shape);
      idMap[id] = "";
    }

    try {
      await ctx.sync();
    } catch (error) {
      throw withOfficeErrorContext(error, "draw_slide_shapes 创建形状或应用样式失败，请检查颜色、shapeType、尺寸等参数");
    }
    let createdIndex = 0;
    if (title && typeof title.text === "string") {
      const titleId = typeof title.id === "string" && title.id ? title.id : "__title";
      const shape = created[createdIndex++];
      idMap[titleId] = shape.id;
      rects.set(titleId, rectFromShape(shape));
    }
    for (const item of shapes) {
      const id = requireId(item.id);
      const shape = created[createdIndex++];
      idMap[id] = shape.id;
      rects.set(id, rectFromShape(shape));
    }

    const connectorResults: ConnectorCreationResult[] = [];
    const connectorLogs: string[] = [];
    for (const connector of connectors) {
      const from = requireId(connector.from);
      const to = requireId(connector.to);
      const fromRect = rects.get(from);
      const toRect = rects.get(to);
      if (!fromRect || !toRect) throw new Error(`连接器引用了不存在的 shape: ${from} -> ${to}`);
      const fromShapeId = idMap[from];
      const toShapeId = idMap[to];
      const fromSide = sideOr(connector.fromSide, "right");
      const toSide = sideOr(connector.toSide, "left");
      const drawResult = drawConnectedLine(
        slide,
        fromShapeId,
        fromRect,
        fromSide,
        toShapeId,
        toRect,
        toSide,
        {
          mode: connector.mode === "direct" ? "direct" : "orthogonal",
          arrow: connector.arrow === "none" ? "none" : "end",
          color: typeof connector.color === "string" ? connector.color : undefined,
          thickness: typeof connector.thickness === "number" ? connector.thickness : undefined,
          dashStyle: typeof connector.dashStyle === "string" ? connector.dashStyle : undefined,
        }
      );
      connectorResults.push(drawResult);
      connectorLogs.push(`${from}.${fromSide}->${to}.${toSide}`);
    }
    try {
      await ctx.sync();
    } catch (error) {
      throw withOfficeErrorContext(error, "draw_slide_shapes 创建连接器占位线失败，请检查 connector sides 和线条样式");
    }

    return {
      slideId: slide.id,
      slideIndex: typeof input.slideIndex === "number" ? input.slideIndex as number : undefined,
      pageNumber: typeof input.pageNumber === "number" ? input.pageNumber as number : undefined,
      idMap,
      connectorPatches: resolveConnectorXmlPatches(connectorResults),
      connectorLogs,
    };
  });

  const editResult = await applyConnectorXmlPatches({
    slideId: result.slideId,
    slideIndex: result.slideIndex,
    pageNumber: result.pageNumber,
  }, result.connectorPatches);
  const slideText = editResult ? `；新 slideId=${editResult.newSlideId}` : "";
  return `已自由绘制 ${Object.keys(result.idMap).length} 个形状、${result.connectorLogs.length} 条连接器${slideText}。id 映射: ${JSON.stringify(result.idMap)}`;
};

function rectFromShape(shape: PowerPoint.Shape): DrawnShapeRect {
  return {
    id: shape.id,
    left: shape.left,
    top: shape.top,
    width: shape.width,
    height: shape.height,
  };
}

function requireId(value: unknown): string {
  if (typeof value !== "string" || !value.trim()) throw new Error("shape/connectors 缺少 id");
  return value.trim();
}

function numberOr(value: unknown, fallback: number): number {
  return typeof value === "number" && Number.isFinite(value) ? value : fallback;
}

function sideOr(value: unknown, fallback: Side): Side {
  if (value === "top" || value === "bottom" || value === "left" || value === "right") return value;
  return fallback;
}

function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}
