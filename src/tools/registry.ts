import { ToolDefinition, ToolHandler } from "../types";
import { getCurrentContext, listSlides } from "./context";
import { addSlide, deleteSlide } from "./slides";
import { addGeometricShape, addLine, addTextBox, connectShapes, deleteShape, modifyShape } from "./shapes";
import { createDiagram } from "./layout";
import { applyPptxPatch, exportPptxXml } from "./ooxml";

export const TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "get_current_context",
    description: "读取当前 PowerPoint 上下文：总幻灯片数、当前选中的幻灯片索引/ID，以及选中的形状（含文字、位置、大小）。在任何修改操作之前都应先调用这个工具。",
    input_schema: { type: "object", properties: {}, additionalProperties: false }
  },
  {
    name: "list_slides",
    description: "列出演示文稿中所有幻灯片的索引和 ID。",
    input_schema: { type: "object", properties: {}, additionalProperties: false }
  },
  {
    name: "add_slide",
    description: "在演示文稿末尾添加一张空白幻灯片（使用默认 Master/Layout）。",
    input_schema: { type: "object", properties: {}, additionalProperties: false }
  },
  {
    name: "delete_slide",
    description: "按索引或 ID 删除一张幻灯片。",
    input_schema: {
      type: "object",
      properties: {
        index: { type: "number", description: "0-based 索引" },
        slideId: { type: "string", description: "幻灯片 ID" }
      }
    }
  },
  {
    name: "add_text_box",
    description: "在指定幻灯片（默认当前选中）上插入一个文本框。",
    input_schema: {
      type: "object",
      properties: {
        text: { type: "string" },
        slideId: { type: "string" },
        slideIndex: { type: "number" },
        left: { type: "number" },
        top: { type: "number" },
        width: { type: "number" },
        height: { type: "number" }
      },
      required: ["text"]
    }
  },
  {
    name: "add_geometric_shape",
    description: "在幻灯片上添加一个几何形状（矩形/圆角矩形/椭圆/菱形等），可以带文字。画逻辑/流程/调用链图时优先用这个工具做节点，而不是纯文本框。shapeType 支持：rectangle、roundRectangle、ellipse、diamond、triangle、rightTriangle、parallelogram、trapezoid、pentagon、hexagon、octagon、plus、rightArrow、leftArrow、upArrow、downArrow、star5、flowChartProcess、flowChartDecision、flowChartTerminator、flowChartData 等 Office GeometricShapeType 枚举值。",
    input_schema: {
      type: "object",
      properties: {
        shapeType: { type: "string", description: "几何形状类型，见 description" },
        text: { type: "string", description: "形状内显示的文字（可选）" },
        slideId: { type: "string" },
        slideIndex: { type: "number" },
        left: { type: "number" },
        top: { type: "number" },
        width: { type: "number" },
        height: { type: "number" }
      },
      required: ["shapeType"]
    }
  },
  {
    name: "add_line",
    description: "在幻灯片上添加一条线/箭头连接线，用来连接两个形状表达调用/数据流向。lineType 支持：straight、elbow、curved。坐标以 (left, top) 为起点，(left+width, top+height) 为终点。",
    input_schema: {
      type: "object",
      properties: {
        lineType: { type: "string", description: "straight / elbow / curved" },
        slideId: { type: "string" },
        slideIndex: { type: "number" },
        left: { type: "number", description: "起点 X" },
        top: { type: "number", description: "起点 Y" },
        width: { type: "number", description: "终点相对起点的 X 偏移（可为负）" },
        height: { type: "number", description: "终点相对起点的 Y 偏移（可为负）" }
      },
      required: ["lineType"]
    }
  },
  {
    name: "connect_shapes",
    description: "按形状 id 将两个现有形状的某条边中点用带箭头的连接线连起来。AI 不再自己算坐标。fromSide/toSide 指定连接哪一侧（top/bottom/left/right）。所有方向的连线都会自动生成带箭头的正交路径：正交方向用单个箭头形状，斜向方向用三段式路径（细矩形线段 + 箭头形状头）。arrow 默认 'end'。画完图后加新连接、或改现有图时用这个工具，不要再用 add_line 画裸线。",
    input_schema: {
      type: "object",
      properties: {
        fromShapeId: { type: "string", description: "起点形状 id" },
        fromSide: { type: "string", enum: ["top", "bottom", "left", "right"], description: "从哪一条边中点出发" },
        toShapeId: { type: "string", description: "终点形状 id" },
        toSide: { type: "string", enum: ["top", "bottom", "left", "right"], description: "连到哪一条边中点" },
        arrow: { type: "string", enum: ["none", "end", "both"], description: "箭头方向，默认 end" },
        slideId: { type: "string" },
        slideIndex: { type: "number" }
      },
      required: ["fromShapeId", "fromSide", "toShapeId", "toSide"]
    }
  },
  {
    name: "create_diagram",
    description: "一次性生成一张完整的图（流程图/调用链/架构图等）。传入节点和连线的抽象描述，内部自动布局、创建节点、连带箭头的连线。这是画图的首选工具 —— 不要自己逐个调用 add_geometric_shape + add_line 堆砌。layout: vertical(竖排)/horizontal(横排)/layered(按 level 分层，若节点带 level 则用之，否则按 edges 自动推断)/tree(按 edges 从根向下)。节点统一尺寸 160x60。返回每个节点的 shapeId 映射，可用于后续 modify_shape 微调。",
    input_schema: {
      type: "object",
      properties: {
        layout: { type: "string", enum: ["vertical", "horizontal", "layered", "tree"] },
        nodes: {
          type: "array",
          items: {
            type: "object",
            properties: {
              id: { type: "string", description: "节点临时 id，用于 edges 引用" },
              text: { type: "string", description: "节点内显示文字" },
              shape: { type: "string", description: "几何形状类型，如 rectangle/roundRectangle/diamond/ellipse/flowChartTerminator/flowChartProcess/flowChartDecision，默认 rectangle" },
              level: { type: "number", description: "layered 布局时的层级，0 为顶层" }
            },
            required: ["id", "text"]
          }
        },
        edges: {
          type: "array",
          items: {
            type: "object",
            properties: {
              from: { type: "string" },
              to: { type: "string" },
              arrow: { type: "string", enum: ["none", "end", "both"], description: "箭头方向，默认 end" }
            },
            required: ["from", "to"]
          }
        },
        canvas: {
          type: "object",
          description: "画布区域，默认 {left:40, top:80, width:880, height:420}",
          properties: {
            left: { type: "number" }, top: { type: "number" },
            width: { type: "number" }, height: { type: "number" }
          }
        },
        slideId: { type: "string" },
        slideIndex: { type: "number" }
      },
      required: ["layout", "nodes", "edges"]
    }
  },
  {
    name: "modify_shape",
    description: "按 shapeId 修改一个形状：文字、位置、大小。",
    input_schema: {
      type: "object",
      properties: {
        shapeId: { type: "string" },
        text: { type: "string" },
        left: { type: "number" },
        top: { type: "number" },
        width: { type: "number" },
        height: { type: "number" }
      },
      required: ["shapeId"]
    }
  },
  {
    name: "delete_shape",
    description: "删除指定 shapeId 的形状。",
    input_schema: {
      type: "object",
      properties: { shapeId: { type: "string" } },
      required: ["shapeId"]
    }
  },
  {
    name: "export_pptx_xml",
    description: "导出当前 .pptx 内部某个 XML 文件的文本（如 ppt/slides/slide1.xml）。也可传 list=true 列出所有文件。",
    input_schema: {
      type: "object",
      properties: {
        path: { type: "string", description: "包内路径，默认 ppt/slides/slide1.xml" },
        list: { type: "boolean", description: "为 true 时仅列出包内所有文件名" },
        maxChars: { type: "number", description: "返回 XML 的最大字符数，默认 4000" }
      }
    }
  },
  {
    name: "apply_pptx_patch",
    description: "对当前 .pptx 应用一组 XML 级 patch，产出新的 .pptx 文件并触发浏览器下载。注意：这不会原地修改当前打开的文件，用户需手动打开新文件。",
    input_schema: {
      type: "object",
      properties: {
        patches: {
          type: "array",
          items: {
            type: "object",
            properties: {
              path: { type: "string" },
              content: { type: "string" }
            },
            required: ["path", "content"]
          }
        }
      },
      required: ["patches"]
    }
  }
];

export const TOOL_HANDLERS: Record<string, ToolHandler> = {
  get_current_context: getCurrentContext,
  list_slides: listSlides,
  add_slide: addSlide,
  delete_slide: deleteSlide,
  add_text_box: addTextBox,
  add_geometric_shape: addGeometricShape,
  add_line: addLine,
  connect_shapes: connectShapes,
  create_diagram: createDiagram,
  modify_shape: modifyShape,
  delete_shape: deleteShape,
  export_pptx_xml: exportPptxXml,
  apply_pptx_patch: applyPptxPatch
};
