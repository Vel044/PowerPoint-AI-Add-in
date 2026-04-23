import { ToolDefinition, ToolHandler } from "../types";
import { getCurrentContext, listSlides } from "./context";
import { addSlide, deleteSlide } from "./slides";
import { addTextBox, deleteShape, modifyShape } from "./shapes";
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
  modify_shape: modifyShape,
  delete_shape: deleteShape,
  export_pptx_xml: exportPptxXml,
  apply_pptx_patch: applyPptxPatch
};
