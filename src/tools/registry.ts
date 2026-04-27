import { ToolDefinition, ToolHandler } from "../types";
import { exportDeckOutline, getCurrentContext, listSlideShapes, listSlides } from "./context";
import { addSlide, deleteSlide, duplicateSlide } from "./slides";
import { addGeometricShape, addLine, addTextBox, connectShapes, deleteShape, modifyShape } from "./shapes";
import { createDiagram } from "./layout";
import { drawSlideShapes } from "./freeDraw";
import { editSlideText, editSlideXml, readSlideText } from "./richText";
import { reviewSlide } from "./review";
import { todoWrite } from "./todo";
import { verifySlides } from "./verify";

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
    name: "list_slide_shapes",
    description: "列出目标幻灯片上的所有形状。返回 ref=slideId:shapeId、id/name/type/text/left/top/width/height/bounds。编辑前优先调用它确认目标 ref，不要手写猜测 shapeId。",
    input_schema: {
      type: "object",
      properties: {
        slideId: { type: "string" },
        slideIndex: { type: "number", description: "0-based 幻灯片索引" },
        pageNumber: { type: "number", description: "1-based 页码" }
      },
      additionalProperties: false
    }
  },
  {
    name: "export_deck_outline",
    description: "导出演示文稿大纲：每页 slideId/pageNumber、标题、文本形状、shapeCount。可选保存到 debug-artifacts/deck-outline/。",
    input_schema: {
      type: "object",
      properties: {
        slides: {
          oneOf: [
            { type: "string", enum: ["all"] },
            { type: "array", items: { type: "number" } }
          ],
          description: "1-based 页码数组，或 all；默认 all"
        },
        saveArtifact: { type: "boolean", description: "是否保存 JSON artifact，默认 true" }
      },
      additionalProperties: false
    }
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
    name: "duplicate_slide",
    description: "复制指定幻灯片，副本插入在原片后面。目标页可用 slideId、slideIndex/index(0-based) 或 pageNumber(1-based) 指定。",
    input_schema: {
      type: "object",
      properties: {
        index: { type: "number", description: "0-based 索引" },
        slideIndex: { type: "number", description: "0-based 索引" },
        pageNumber: { type: "number", description: "1-based 页码" },
        slideId: { type: "string" }
      }
    }
  },
  {
    name: "add_text_box",
    description: "在指定幻灯片（默认当前选中）上插入一个文本框。支持 fillColor/lineColor/lineWeight/textColor/fontSize/bold 样式参数。目标页可用 slideId、slideIndex(0-based) 或 pageNumber(1-based) 指定。",
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
        height: { type: "number" },
        fillColor: { type: "string" },
        lineColor: { type: "string" },
        lineWeight: { type: "number" },
        textColor: { type: "string" },
        fontSize: { type: "number" },
        bold: { type: "boolean" }
      },
      required: ["text"]
    }
  },
  {
    name: "add_geometric_shape",
    description: "在幻灯片上添加一个几何形状（矩形/圆角矩形/椭圆/菱形等），可以带文字。支持 fillColor/lineColor/lineWeight/textColor/fontSize/bold。目标页可用 slideId、slideIndex(0-based) 或 pageNumber(1-based) 指定。shapeType 支持 rectangle、roundRectangle、ellipse、diamond、flowChartTerminator、flowChartDecision、flowChartInputOutput、can 等；历史别名 flowChartData 会自动映射。",
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
        height: { type: "number" },
        fillColor: { type: "string" },
        lineColor: { type: "string" },
        lineWeight: { type: "number" },
        textColor: { type: "string" },
        fontSize: { type: "number" },
        bold: { type: "boolean" }
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
        height: { type: "number", description: "终点相对起点的 Y 偏移（可为负）" },
        color: { type: "string" },
        thickness: { type: "number" },
        dashStyle: { type: "string", enum: ["solid", "dash", "dot", "dashDot"] },
        arrow: { type: "string", enum: ["none", "end"] }
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
        color: { type: "string" },
        thickness: { type: "number" },
        dashStyle: { type: "string", enum: ["solid", "dash", "dot", "dashDot"] },
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
    name: "draw_slide_shapes",
    description: "批量自由绘制复杂框图。模型可直接决定每个框的位置、尺寸、颜色、字体、标题带和连接器；适合调用链、架构图、泳道图、时间线。连接器会通过真实 PowerPoint connector + XML 修正生成，返回临时 id 到真实 shapeId 的映射和新 slideId。",
    input_schema: {
      type: "object",
      properties: {
        slideId: { type: "string" },
        slideIndex: { type: "number" },
        pageNumber: { type: "number" },
        title: {
          type: "object",
          properties: {
            id: { type: "string" },
            text: { type: "string" },
            left: { type: "number" },
            top: { type: "number" },
            width: { type: "number" },
            height: { type: "number" },
            fillColor: { type: "string" },
            lineColor: { type: "string" },
            textColor: { type: "string" },
            fontSize: { type: "number" },
            bold: { type: "boolean" }
          }
        },
        shapes: {
          type: "array",
          items: {
            type: "object",
            properties: {
              id: { type: "string", description: "临时 id，供 connectors 引用" },
              type: { type: "string", enum: ["textBox", "geometricShape"] },
              shapeType: { type: "string", description: "geometricShape 类型，默认 rectangle" },
              text: { type: "string" },
              left: { type: "number" },
              top: { type: "number" },
              width: { type: "number" },
              height: { type: "number" },
              fillColor: { type: "string" },
              lineColor: { type: "string" },
              lineWeight: { type: "number" },
              textColor: { type: "string" },
              fontSize: { type: "number" },
              bold: { type: "boolean" }
            },
            required: ["id", "left", "top", "width", "height"]
          }
        },
        connectors: {
          type: "array",
          items: {
            type: "object",
            properties: {
              from: { type: "string" },
              fromSide: { type: "string", enum: ["top", "bottom", "left", "right"] },
              to: { type: "string" },
              toSide: { type: "string", enum: ["top", "bottom", "left", "right"] },
              mode: { type: "string", enum: ["orthogonal", "direct"] },
              arrow: { type: "string", enum: ["none", "end"] },
              color: { type: "string" },
              thickness: { type: "number" },
              dashStyle: { type: "string", enum: ["solid", "dash", "dot", "dashDot"] }
            },
            required: ["from", "fromSide", "to", "toSide"]
          }
        }
      },
      required: ["shapes"]
    }
  },
  {
    name: "read_slide_text",
    description: "读取指定形状的原始 OOXML <a:p> 段落 XML。ref 推荐来自 list_slide_shapes，格式 slideId:shapeId；也可传 slideId/pageNumber + shapeId。",
    input_schema: {
      type: "object",
      properties: {
        ref: { type: "string" },
        shapeId: { type: "string" },
        slideId: { type: "string" },
        slideIndex: { type: "number" },
        pageNumber: { type: "number" }
      }
    }
  },
  {
    name: "edit_slide_text",
    description: "用新的 OOXML <a:p> 段落 XML 替换指定形状的文本内容。ref 推荐来自 list_slide_shapes。",
    input_schema: {
      type: "object",
      properties: {
        ref: { type: "string" },
        shapeId: { type: "string" },
        slideId: { type: "string" },
        slideIndex: { type: "number" },
        pageNumber: { type: "number" },
        paragraphXml: { type: "string", description: "一个或多个 <a:p> 段落 XML" },
        code: { type: "string", description: "兼容字段，等价于 paragraphXml" },
        autosize: { type: "boolean" }
      }
    }
  },
  {
    name: "edit_slide_xml",
    description: "结构化修改目标幻灯片 OOXML，不接受任意代码字符串。支持 insertShapeXml、replaceShapeXml、deleteShapeXml、patchConnector、setSlideBackground。",
    input_schema: {
      type: "object",
      properties: {
        slideId: { type: "string" },
        slideIndex: { type: "number" },
        pageNumber: { type: "number" },
        operations: { type: "array", items: { type: "object" } },
        autosizeShapeIds: { type: "array", items: { type: "string" } }
      },
      required: ["operations"]
    }
  },
  {
    name: "modify_shape",
    description: "在指定幻灯片（默认当前选中页）内按 shapeId 修改一个形状：文字、位置、大小、样式。目标页可用 slideId、slideIndex(0-based) 或 pageNumber(1-based) 指定。不会跨页搜索；删除/修改前必须先用 list_slide_shapes 或 get_current_context 确认目标。",
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
        height: { type: "number" },
        fillColor: { type: "string" },
        lineColor: { type: "string" },
        lineWeight: { type: "number" },
        textColor: { type: "string" },
        fontSize: { type: "number" },
        bold: { type: "boolean" }
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
  },
  {
    name: "verify_slide_visual",
    description: "review_slide 的 Claude 风格别名：截图指定幻灯片，保存到 debug-artifacts/review-slide/，并让视觉模型检查布局、重叠、歪斜、文字溢出等问题。",
    input_schema: {
      type: "object",
      properties: {
        slideId: { type: "string" },
        slideIndex: { type: "number" },
        pageNumber: { type: "number" },
        question: { type: "string" },
        focus: { type: "string" },
        saveArtifact: { type: "boolean" }
      }
    }
  },
  {
    name: "verify_slides",
    description: "程序化检查一批幻灯片的几何问题：形状重叠、越界、空文本、过小形状等。适合在视觉审查前做快速自检。",
    input_schema: {
      type: "object",
      properties: {
        fromPageNumber: { type: "number", description: "1-based 起始页" },
        toPageNumber: { type: "number", description: "1-based 结束页" },
        from_slide: { type: "number", description: "兼容字段，1-based 起始页" },
        to_slide: { type: "number", description: "兼容字段，1-based 结束页" },
        slideIds: { type: "array", items: { type: "string" } },
        checks: { type: "array", items: { type: "string" } }
      }
    }
  },
  {
    name: "todo_write",
    description: "创建或更新 Agent 任务列表。每次调用都是全量替换，结果会显示在对话里的工具结果中。",
    input_schema: {
      type: "object",
      properties: {
        todos: {
          type: "array",
          items: {
            type: "object",
            properties: {
              content: { type: "string" },
              activeForm: { type: "string" },
              status: { type: "string", enum: ["pending", "in_progress", "completed"] }
            },
            required: ["content", "status"]
          }
        }
      },
      required: ["todos"]
    }
  }
];

export const TOOL_HANDLERS: Record<string, ToolHandler> = {
  get_current_context: getCurrentContext,
  list_slides: listSlides,
  list_slide_shapes: listSlideShapes,
  export_deck_outline: exportDeckOutline,
  add_slide: addSlide,
  delete_slide: deleteSlide,
  duplicate_slide: duplicateSlide,
  add_text_box: addTextBox,
  add_geometric_shape: addGeometricShape,
  add_line: addLine,
  connect_shapes: connectShapes,
  create_diagram: createDiagram,
  draw_slide_shapes: drawSlideShapes,
  read_slide_text: readSlideText,
  edit_slide_text: editSlideText,
  edit_slide_xml: editSlideXml,
  modify_shape: modifyShape,
  delete_shape: deleteShape,
  review_slide: reviewSlide,
  verify_slide_visual: reviewSlide,
  verify_slides: verifySlides,
  todo_write: todoWrite
};
