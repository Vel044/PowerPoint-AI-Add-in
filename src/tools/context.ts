import { ToolHandler } from "../types";

export const getCurrentContext: ToolHandler = async () => {
  try {
    return await PowerPoint.run(async (ctx) => {
      const pres = ctx.presentation;
      const selectedSlides = pres.getSelectedSlides();
      selectedSlides.load("items/id");
      const selectedShapes = pres.getSelectedShapes();
      // 只加载最基本的属性，避免线条/连接器形状的某些属性导致 0x80070057
      selectedShapes.load("id,name");
      const allSlides = pres.slides;
      allSlides.load("items/id");
      await ctx.sync();

      const slideIds = allSlides.items.map((s) => s.id);
      const currentSlideIds = selectedSlides.items.map((s) => s.id);
      const currentIndexes = currentSlideIds.map((id) => slideIds.indexOf(id));

      // 逐个形状收集属性，用 try-catch 包裹以应对线条/连接器的特殊属性
      const shapes = [];
      for (const s of selectedShapes.items) {
        const info: Record<string, unknown> = {
          id: s.id,
          name: s.name
        };

        // 尝试获取 type，对于线条/连接器这可能失败
        try {
          info.type = (s as any).type;
        } catch {
          info.type = "unknown";
        }

        // 尝试获取位置和尺寸
        for (const prop of ["left", "top", "width", "height"]) {
          try {
            info[prop] = (s as any)[prop];
          } catch {
            info[prop] = null;
          }
        }

        // 判断是否为线条/连接器类型
        const type = info.type as string;
        if (type === "line" || type === "connector" || type === "unknown") {
          info.text = "";
          // 尝试获取连接器类型
          try {
            info.connectorType = (s as any).connectorType;
          } catch {
            info.connectorType = null;
          }
          try {
            info.lineType = (s as any).lineType;
          } catch {
            info.lineType = null;
          }
        } else {
          // 普通形状获取文本
          info.text = safeText(s);
        }

        shapes.push(info);
      }

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

function getShapeType(shape: PowerPoint.Shape): string {
  try {
    return (shape as any).type as string;
  } catch {
    return "unknown";
  }
}

function safeGet(shape: PowerPoint.Shape, prop: string): unknown {
  try {
    return (shape as any)[prop];
  } catch {
    return null;
  }
}

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
