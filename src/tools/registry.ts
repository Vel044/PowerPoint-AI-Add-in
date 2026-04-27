import { ToolDefinition, ToolHandler } from "../types";
import { getCurrentContext, listSlides } from "./context";
import { addSlide, deleteSlide } from "./slides";
import { addGeometricShape, addLine, addTextBox, connectShapes, deleteShape, modifyShape } from "./shapes";
import { createDiagram } from "./layout";
import { reviewSlide } from "./review";

export const TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "get_current_context",
    description: "读取 PowerPoint 上下文。默认返回当前选中页；也可传 slideId、slideIndex(0-based) 或 pageNumber(1-based) 静默查看任意页的 allShapes、occupiedBounds，而不切换当前选中页。返回当前真实选择信息，以及 inspectedSlideId/inspectedSlideIndex/inspectedPageNumber 以区分“当前页”和“查看页”。删除/修改前应先用这个工具确认目标页的 slideId + shapeId。",
    input_schema: {
      type: "object",
      properties: {
        slideId: { type: "string" },
        slideIndex: { type: "number", description: "0-based 幻灯片索引" },
        pageNumber: { type: "number", description: "1-based 页码，如第三页传 3" }
      },
      additionalProperties: false
    }
  },
  {
    name: "list_slides",
    description: "列出演示文稿中所有幻灯片的 index(0-based)、pageNumber(1-based) 和 ID。",
    input_schema: { type: "object", properties: {}, additionalProperties: false }
  },
  {
    name: "add_slide",
    description: "在演示文稿末尾添加一张空白幻灯片（使用默认 Master/Layout）。",
    input_schema: { type: "object", properties: {}, additionalProperties: false }
  },
  {
    name: "delete_slide",
    description: "按 slideId、index(0-based) 或 pageNumber(1-based) 删除一张幻灯片。",
    input_schema: {
      type: "object",
      properties: {
        index: { type: "number", description: "0-based 索引" },
        slideIndex: { type: "number", description: "0-based 索引（与 index 等价）" },
        pageNumber: { type: "number", description: "1-based 页码" },
        slideId: { type: "string", description: "幻灯片 ID" }
      }
    }
  },
  {
    name: "add_text_box",
    description: "在指定幻灯片（默认当前选中）上插入一个文本框。目标页可用 slideId、slideIndex(0-based) 或 pageNumber(1-based) 指定。",
    input_schema: {
      type: "object",
      properties: {
        text: { type: "string" },
        slideId: { type: "string" },
        slideIndex: { type: "number" },
        pageNumber: { type: "number" },
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
    description: "在幻灯片上添加一个几何形状（矩形/圆角矩形/椭圆/菱形等），可以带文字。目标页可用 slideId、slideIndex(0-based) 或 pageNumber(1-based) 指定。画逻辑/流程/调用链图时优先用这个工具做节点，而不是纯文本框。shapeType 支持：rectangle、roundRectangle、ellipse、diamond、triangle、rightTriangle、parallelogram、trapezoid、pentagon、hexagon、octagon、plus、rightArrow、leftArrow、upArrow、downArrow、star5、flowChartProcess、flowChartDecision、flowChartTerminator、flowChartInputOutput、can 等 Office GeometricShapeType 枚举值；历史别名 flowChartData 会自动映射到 flowChartInputOutput。",
    input_schema: {
      type: "object",
      properties: {
        shapeType: { type: "string", description: "几何形状类型，见 description" },
        text: { type: "string", description: "形状内显示的文字（可选）" },
        slideId: { type: "string" },
        slideIndex: { type: "number" },
        pageNumber: { type: "number" },
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
    description: "在指定幻灯片上添加一条原生线条。目标页可用 slideId、slideIndex(0-based) 或 pageNumber(1-based) 指定。连接两个形状表达调用/数据流向时优先用 connect_shapes 或 create_diagram。lineType 支持：straight、elbow、curve（curved 会自动映射到 curve）。坐标以 (left, top) 为起点，(left+width, top+height) 为终点。",
    input_schema: {
      type: "object",
      properties: {
        lineType: { type: "string", description: "straight / elbow / curve（curved 也兼容）" },
        slideId: { type: "string" },
        slideIndex: { type: "number" },
        pageNumber: { type: "number" },
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
    description: "连接指定幻灯片上的两个形状。目标页可用 slideId、slideIndex(0-based) 或 pageNumber(1-based) 指定。mode=\"orthogonal\" 时端点同 X/Y 用真实 Straight connector，否则使用真实 bentConnector3 肘形连接器；mode=\"direct\" 强制一段真实 Straight connector。fromSide/toSide 指定起止形状的哪条边中点。arrow=\"end\" 会通过 PowerPoint 原生 tailEnd 箭头指向终点形状。工具会通过单页 export/import 修正 XML，返回的新 slideId 应用于后续操作。",
    input_schema: {
      type: "object",
      properties: {
        fromShapeId: { type: "string", description: "起点形状 id" },
        fromSide: { type: "string", enum: ["top", "bottom", "left", "right"], description: "从哪一条边中点出发" },
        toShapeId: { type: "string", description: "终点形状 id" },
        toSide: { type: "string", enum: ["top", "bottom", "left", "right"], description: "连到哪一条边中点" },
        mode: { type: "string", enum: ["orthogonal", "direct"], description: "orthogonal=同轴 Straight / 其他情况真实 bentConnector3 肘形连接器（默认），direct=一段真实 Straight connector" },
        arrow: { type: "string", enum: ["none", "end"], description: "end=末端箭头（默认），none=无箭头" },
        slideId: { type: "string" },
        slideIndex: { type: "number" },
        pageNumber: { type: "number" }
      },
      required: ["fromShapeId", "fromSide", "toShapeId", "toSide"]
    }
  },
  {
    name: "create_diagram",
    description: "一次性在指定幻灯片生成一张完整的图（流程图/调用链/架构图等）。目标页可用 slideId、slideIndex(0-based) 或 pageNumber(1-based) 指定。传入节点和连线的抽象描述，内部自动布局、创建节点，并用贴边中点的真实 PowerPoint connector 连接；端点同 X/Y 时使用 Straight，否则使用 bentConnector3 肘形连接器。连接器通过原生 tailEnd 生成箭头。工具会通过单页 export/import 修正 XML，返回的新 slideId 应用于后续操作。layout: vertical(竖排)/horizontal(横排)/layered(按 level 分层)/tree(按 edges 从根向下)。节点统一尺寸 160x60。",
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
              shape: { type: "string", description: "几何形状类型，如 rectangle/roundRectangle/diamond/ellipse/flowChartTerminator/flowChartProcess/flowChartDecision/flowChartInputOutput/can，默认 rectangle；历史别名 flowChartData 会自动映射" },
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
              to: { type: "string" }
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
        slideIndex: { type: "number" },
        pageNumber: { type: "number" }
      },
      required: ["layout", "nodes", "edges"]
    }
  },
  {
    name: "modify_shape",
    description: "在指定幻灯片（默认当前选中页）内按 shapeId 修改一个形状：文字、位置、大小。目标页可用 slideId、slideIndex(0-based) 或 pageNumber(1-based) 指定。不会跨页搜索；删除/修改前必须先用 get_current_context 查看目标页 allShapes，确认 slideId + shapeId。",
    input_schema: {
      type: "object",
      properties: {
        shapeId: { type: "string" },
        slideId: { type: "string" },
        slideIndex: { type: "number" },
        pageNumber: { type: "number" },
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
    description: "在指定幻灯片（默认当前选中页）内删除指定 shapeId 的形状。目标页可用 slideId、slideIndex(0-based) 或 pageNumber(1-based) 指定。不会跨页搜索；删除前必须先用 get_current_context 查看目标页 allShapes，确认 slideId + shapeId。返回删除前 bounds，可用于 create_diagram.canvas 原位替换。",
    input_schema: {
      type: "object",
      properties: {
        shapeId: { type: "string" },
        slideId: { type: "string" },
        slideIndex: { type: "number" },
        pageNumber: { type: "number" }
      },
      required: ["shapeId"]
    }
  },
  {
    name: "review_slide",
    description: "截取指定幻灯片的截图，保存到本地 debug-artifacts/review-slide/（同时写入 JSON 元数据并在 tool result/终端日志返回路径），再发给视觉模型做视觉检查。目标页可用 slideId、slideIndex(0-based) 或 pageNumber(1-based) 指定。返回审查结果（布局问题、重叠、歪斜、文字溢出等）。画完图/做完修改后应调用此工具检查结果，如果发现问题可以立即修正。",
    input_schema: {
      type: "object",
      properties: {
        slideId: { type: "string" },
        slideIndex: { type: "number" },
        pageNumber: { type: "number" },
        question: { type: "string", description: "自定义审查问题，默认检查布局/重叠/连线歪斜" }
      }
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
  review_slide: reviewSlide
};
