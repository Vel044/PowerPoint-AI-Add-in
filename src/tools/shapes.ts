import { ToolHandler } from "../types";
import { withOfficeErrorContext } from "./officeErrors";
import { normalizeConnectorType, normalizeGeometricShapeType } from "./shapeTypes";
import { Side, applyConnectorXmlPatches, drawConnectedLine, resolveConnectorXmlPatches } from "./layout";
import { resolveSlide, slideTargetFromInput } from "./slideTarget";

export const addTextBox: ToolHandler = async (input) => {
  const text = String(input.text ?? "");
  const left = (input.left as number) ?? 50;
  const top = (input.top as number) ?? 50;
  const width = (input.width as number) ?? 400;
  const height = (input.height as number) ?? 80;
  return await PowerPoint.run(async (ctx) => {
    const slide = await resolveSlide(ctx, slideTargetFromInput(input));
    const shape = slide.shapes.addTextBox(text, { left, top, width, height });
    shape.load("id");
    await ctx.sync();
    return `已在幻灯片 ${slide.id} 添加文本框 (id=${shape.id})`;
  });
};

export const modifyShape: ToolHandler = async (input) => {
  const shapeId = String(input.shapeId ?? "");
  if (!shapeId) throw new Error("缺少 shapeId");
  return await PowerPoint.run(async (ctx) => {
    const slide = await resolveSlide(ctx, slideTargetFromInput(input));
    const shapes = slide.shapes;
    shapes.load("items/id");
    await ctx.sync();
    const target = shapes.items.find((sh) => sh.id === shapeId);
    if (!target) {
      throw new Error(`当前目标幻灯片（id=${slide.id}）上未找到形状 ${shapeId}。请先调用 get_current_context 查看目标页 allShapes，并使用其中的 slideId + shapeId 操作。`);
    }
    if (typeof input.text === "string") {
      target.textFrame.textRange.text = input.text as string;
    }
    if (typeof input.left === "number") target.left = input.left as number;
    if (typeof input.top === "number") target.top = input.top as number;
    if (typeof input.width === "number") target.width = input.width as number;
    if (typeof input.height === "number") target.height = input.height as number;
    await ctx.sync();
    return `已更新幻灯片 ${slide.id} 上的形状 ${shapeId}`;
  });
};

export const addGeometricShape: ToolHandler = async (input) => {
  const shapeType = normalizeGeometricShapeType(input.shapeType ?? "rectangle");
  const text = typeof input.text === "string" ? (input.text as string) : "";
  const left = (input.left as number) ?? 50;
  const top = (input.top as number) ?? 50;
  const width = (input.width as number) ?? 200;
  const height = (input.height as number) ?? 60;
  return await PowerPoint.run(async (ctx) => {
    const slide = await resolveSlide(ctx, slideTargetFromInput(input));
    const shape = slide.shapes.addGeometricShape(shapeType, {
      left,
      top,
      width,
      height
    });
    if (text) {
      shape.textFrame.textRange.text = text;
    }
    shape.load("id");
    try {
      await ctx.sync();
    } catch (error) {
      throw withOfficeErrorContext(error, `添加几何形状失败，shapeType=${input.shapeType ?? "rectangle"}，normalized=${shapeType}`);
    }
    return `已在幻灯片 ${slide.id} 添加几何形状 ${shapeType} (id=${shape.id})`;
  });
};

export const addLine: ToolHandler = async (input) => {
  const lineType = normalizeConnectorType(input.lineType ?? "straight");
  const left = (input.left as number) ?? 0;
  const top = (input.top as number) ?? 0;
  const width = (input.width as number) ?? 100;
  const height = (input.height as number) ?? 0;
  return await PowerPoint.run(async (ctx) => {
    const slide = await resolveSlide(ctx, slideTargetFromInput(input));
    const shape = slide.shapes.addLine(lineType, { left, top, width, height });
    shape.load("id");
    try {
      await ctx.sync();
    } catch (error) {
      throw withOfficeErrorContext(error, `添加连接线失败，lineType=${input.lineType ?? "straight"}，normalized=${lineType}`);
    }
    return `已在幻灯片 ${slide.id} 添加连接线 ${lineType} (id=${shape.id})`;
  });
};

interface ShapeRect {
  id: string;
  left: number;
  top: number;
  width: number;
  height: number;
}

interface ShapeBounds {
  left: number | null;
  top: number | null;
  width: number | null;
  height: number | null;
}

export const connectShapes: ToolHandler = async (input) => {
  const fromShapeId = String(input.fromShapeId ?? "");
  const toShapeId = String(input.toShapeId ?? "");
  const fromSide = (input.fromSide as Side) ?? "right";
  const toSide = (input.toSide as Side) ?? "left";
  const mode = (input.mode as string) ?? "orthogonal";
  const arrow = ((input.arrow as string) ?? "end") === "none" ? "none" : "end";
  if (!fromShapeId || !toShapeId) throw new Error("缺少 fromShapeId 或 toShapeId");

  const result = await PowerPoint.run(async (ctx) => {
    const slide = await resolveSlide(ctx, slideTargetFromInput(input));
    const shapeColl = slide.shapes;
    shapeColl.load("items/id");
    await ctx.sync();
    const idSet = new Set(shapeColl.items.map((s) => s.id));
    const missing = [fromShapeId, toShapeId].filter((id) => !idSet.has(id));
    if (missing.length > 0) {
      throw new Error(`当前目标幻灯片（id=${slide.id}）上未找到形状: ${missing.join(", ")}。请先调用 get_current_context 确认目标页形状 ID，或传入 pageNumber/slideIndex/slideId。`);
    }
    const fromShape = slide.shapes.getItem(fromShapeId);
    const toShape = slide.shapes.getItem(toShapeId);
    fromShape.load("id,left,top,width,height");
    toShape.load("id,left,top,width,height");
    await ctx.sync();
    const shapes: Record<string, ShapeRect> = {
      [fromShapeId]: { id: fromShapeId, left: fromShape.left, top: fromShape.top, width: fromShape.width, height: fromShape.height },
      [toShapeId]: { id: toShapeId, left: toShape.left, top: toShape.top, width: toShape.width, height: toShape.height },
    };
    for (const id of [fromShapeId, toShapeId]) {
      const r = shapes[id];
      if (typeof r.left !== "number" || typeof r.top !== "number" || typeof r.width !== "number" || typeof r.height !== "number") {
        throw new Error(`形状 ${id} 的位置属性读取失败 (left=${r.left}, top=${r.top}, w=${r.width}, h=${r.height})。可能是占位符或母版继承的形状，无法作为连接端点。`);
      }
    }
    const connector = drawConnectedLine(
      slide,
      fromShapeId,
      shapes[fromShapeId],
      fromSide,
      toShapeId,
      shapes[toShapeId],
      toSide,
      { arrow, mode: mode === "direct" ? "direct" : "orthogonal" },
    );
    await ctx.sync();
    return {
      slideId: slide.id,
      slideIndex: typeof input.slideIndex === "number" ? input.slideIndex as number : undefined,
      pageNumber: typeof input.pageNumber === "number" ? input.pageNumber as number : undefined,
      connectorPatches: resolveConnectorXmlPatches([connector]),
      message: `已连接 ${fromShapeId}.${fromSide} → ${toShapeId}.${toSide}（${mode === "direct" ? "直连" : "自动横平竖直"}，Straight=${connector.straightConnectors}，Elbow=${connector.elbowConnectors}，总连接器=${connector.lineSegments}，原生箭头=${connector.arrows}）`,
    };
  });
  const editResult = await applyConnectorXmlPatches({
    slideId: result.slideId,
    slideIndex: result.slideIndex,
    pageNumber: result.pageNumber,
  }, result.connectorPatches);
  const slideText = editResult ? `；新 slideId=${editResult.newSlideId}` : "";
  return `${result.message}。已通过单页 export/import 修正真实 PowerPoint 连接器${slideText}。`;
};

export const deleteShape: ToolHandler = async (input) => {
  const shapeId = String(input.shapeId ?? "");
  if (!shapeId) throw new Error("缺少 shapeId");
  return await PowerPoint.run(async (ctx) => {
    const slide = await resolveSlide(ctx, slideTargetFromInput(input));
    const shapes = slide.shapes;
    shapes.load("items/id,name");
    await ctx.sync();
    const target = shapes.items.find((sh) => sh.id === shapeId);
    if (!target) {
      throw new Error(`当前目标幻灯片（id=${slide.id}）上未找到形状 ${shapeId}。不会跨页搜索删除；请先调用 get_current_context 查看目标页 allShapes。`);
    }

    const name = safeString(() => target.name);
    const bounds = await readShapeBounds(ctx, target);
    target.delete();
    await ctx.sync();
    return `已删除幻灯片 ${slide.id} 上的形状 ${shapeId}（name=${name || "未知"}，bounds=${JSON.stringify(bounds)}）。如需替换该图形，可把这个 bounds 用作 create_diagram.canvas。`;
  });
};

async function readShapeBounds(ctx: PowerPoint.RequestContext, shape: PowerPoint.Shape): Promise<ShapeBounds> {
  try {
    shape.load("left,top,width,height");
    await ctx.sync();
  } catch {
    // Some shape-like objects may not expose geometry. Leave fields null below.
  }
  return {
    left: safeNumber(() => shape.left),
    top: safeNumber(() => shape.top),
    width: safeNumber(() => shape.width),
    height: safeNumber(() => shape.height)
  };
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
