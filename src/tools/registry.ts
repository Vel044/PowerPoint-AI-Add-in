import { ToolDefinition, ToolHandler } from "../types";
import { exportDeckOutline, getCurrentContext, listSlideShapes, listSlides } from "./context";
import { addSlide, deleteSlide, duplicateSlide } from "./slides";
import { addGeometricShape, addLine, addTextBox, connectShapes, deleteShape, modifyShape } from "./shapes";
import { editSlideText, editSlideXml, readSlideText } from "./richText";
import { reviewSlide } from "./review";
import { todoWrite } from "./todo";
import { verifySlides } from "./verify";
import {
  copyImageBetweenSlides,
  createSkill,
  editSlideChart,
  editSlideMaster,
  executeOfficeJs,
  insertIcon,
  readSkill,
  searchIcons,
  storeBlob,
  unsupportedBridgeTool,
  updateInstructions,
  updateSetting,
  webSearch
} from "./remainingTools";

export const TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "get_current_context",
    description: "读取 PowerPoint 上下文。默认返回当前选中页；也可传 slideId、slideIndex(0-based) 或 pageNumber(1-based) 静默查看任意页的 allShapes、occupiedBounds，而不切换当前选中页。返回当前真实选择信息，以及 inspectedSlideId/inspectedSlideIndex/inspectedPageNumber 以区分”当前页”和”查看页”。删除/修改前应先用这个工具确认目标页的 slideId + shapeId。返回 slideWidth/slideHeight（pt）表示幻灯片画布实际尺寸，画图前必须先调用此工具确认。同时返回 themePalette/isDefaultTheme：当用户使用自定义品牌主题时，themePalette 提供 dk1/lt1/accent1-6 等主题色 hex 值（不带 #）；默认 Office 主题时 themePalette 为 null。",
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
    description: "在指定幻灯片（默认当前选中）上插入一个文本框。支持 fillColor/lineColor/lineWeight/textColor/fontSize/bold 样式参数。目标页可用 slideId、slideIndex(0-based) 或 pageNumber(1-based) 指定。省略 fillColor/lineColor/textColor 时形状继承演示文稿主题默认样式，通常效果更好。坐标单位为点(pt)；请先用 get_current_context 确认画布尺寸和已有形状位置。",
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
    description: "在幻灯片上添加一个几何形状（矩形/圆角矩形/椭圆/菱形等），可以带文字。支持 fillColor/lineColor/lineWeight/textColor/fontSize/bold。目标页可用 slideId、slideIndex(0-based) 或 pageNumber(1-based) 指定。shapeType 支持 rectangle、roundRectangle、ellipse、diamond、flowChartTerminator、flowChartDecision、flowChartInputOutput、can 等；历史别名 flowChartData 会自动映射。省略 fillColor/lineColor/textColor 时形状继承演示文稿主题默认样式，通常效果更好。坐标单位为点(pt)；请先用 get_current_context 确认画布尺寸和已有形状位置。",
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
    description: "在指定幻灯片上添加一条原生线条。目标页可用 slideId、slideIndex(0-based) 或 pageNumber(1-based) 指定。连接两个形状表达调用/数据流向时优先用 connect_shapes。lineType 支持：straight、elbow、curve（curved 会自动映射到 curve）。坐标以 (left, top) 为起点，(left+width, top+height) 为终点。坐标单位为点(pt)；请先用 get_current_context 确认画布尺寸和已有形状位置。",
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
    description: "连接指定幻灯片上的两个形状。目标页可用 slideId、slideIndex(0-based) 或 pageNumber(1-based) 指定。mode=\"orthogonal\" 时端点同 X/Y 用真实 Straight connector，否则使用真实 bentConnector3 肘形连接器；mode=\"direct\" 强制一段真实 Straight connector。fromSide/toSide 指定起止形状的哪条边中点。arrow=\"end\" 会通过 PowerPoint 原生 tailEnd 箭头指向终点形状。工具会通过单页 export/import 修正 XML，返回的新 slideId 应用于后续操作。省略 color 时连接线继承主题默认线条颜色。",
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
    description: "Claude 风格原位编辑目标幻灯片 OOXML：传入 async JS 函数体 code，工具导出单页 PPTX 为 JSZip，注入 zip、markDirty 和 pptx helper，执行后重新插回并删除旧页。普通画图优先用 pptx.addShape(...); pptx.addConnector(...); pptx.save(); markDirty();，不要手写大段 createElementNS。pptx.openSlide() 固定使用 ppt/slides/slide1.xml。仍可直接用 zip 做高级 patch。执行前后会保存 debug-artifacts/edit-slide-xml/*.xml，并返回新 slideId。",
    input_schema: {
      type: "object",
      properties: {
        slideId: { type: "string" },
        slideIndex: { type: "number" },
        pageNumber: { type: "number" },
        code: {
          type: "string",
          description: "async 函数体，不要包 async function。优先使用 pptx helper：const a=pptx.addShape({id:'a',shapeType:'roundRect',text:'A',left:40,top:40,width:160,height:60}); const b=pptx.addShape(...); pptx.addConnector({from:a,fromSide:'right',to:b,toSide:'left'}); pptx.save(); markDirty();。也可直接使用 zip、DOMParser、XMLSerializer 做高级 XML patch。"
        },
        autosize_shape_ids: {
          type: "array",
          items: { type: "string" },
          description: "可选：写回前给这些 shape id 的 bodyPr 确保 a:normAutofit"
        },
        autosizeShapeIds: {
          type: "array",
          items: { type: "string" },
          description: "兼容字段，等价于 autosize_shape_ids"
        },
        explanation: { type: "string", description: "本次 XML 编辑意图，写入 artifact metadata" }
      },
      required: ["code"]
    }
  },
  {
    name: "modify_shape",
    description: "在指定幻灯片（默认当前选中页）内按 shapeId 修改一个形状：文字、位置、大小、样式。目标页可用 slideId、slideIndex(0-based) 或 pageNumber(1-based) 指定。不会跨页搜索；删除/修改前必须先用 list_slide_shapes 或 get_current_context 确认目标。省略 fillColor/lineColor/textColor 时形状继承演示文稿主题默认样式，通常效果更好。",
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
    description: "在指定幻灯片（默认当前选中页）内删除指定 shapeId 的形状。目标页可用 slideId、slideIndex(0-based) 或 pageNumber(1-based) 指定。不会跨页搜索；删除前必须先用 get_current_context 查看目标页 allShapes，确认 slideId + shapeId。返回删除前 bounds，可用于原位替换时规划新图坐标。",
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
    name: "execute_office_js",
    description: "执行受控 Office.js 白名单动作，不接受任意代码字符串。支持 moveShape、resizeShape、setShapeStyle、bringToFront、sendToBack、selectSlide、selectShapes。用于细调位置、尺寸、样式、层级和选择状态。",
    input_schema: {
      type: "object",
      properties: {
        actions: {
          type: "array",
          items: {
            type: "object",
            properties: {
              type: { type: "string", enum: ["moveShape", "resizeShape", "setShapeStyle", "bringToFront", "sendToBack", "selectSlide", "selectShapes"] },
              slideId: { type: "string" },
              slideIndex: { type: "number" },
              pageNumber: { type: "number" },
              shapeId: { type: "string" },
              shapeIds: { type: "array", items: { type: "string" } },
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
            required: ["type"]
          }
        }
      },
      required: ["actions"],
      additionalProperties: false
    }
  },
  {
    name: "edit_slide_chart",
    description: "在目标页生成可编辑的 shape-based 图表。V1 支持 bar、line、pie 三类，参数为 title/categories/series/left/top/width/height。生成的是普通形状，后续可移动和修改；不是 native chart package。",
    input_schema: {
      type: "object",
      properties: {
        slideId: { type: "string" },
        slideIndex: { type: "number" },
        pageNumber: { type: "number" },
        chartType: { type: "string", enum: ["bar", "line", "pie"] },
        type: { type: "string", enum: ["bar", "line", "pie"] },
        title: { type: "string" },
        categories: { type: "array", items: { type: "string" } },
        series: {
          type: "array",
          items: {
            type: "object",
            properties: {
              name: { type: "string" },
              values: { type: "array", items: { type: "number" } }
            },
            required: ["values"]
          }
        },
        left: { type: "number" },
        top: { type: "number" },
        width: { type: "number" },
        height: { type: "number" }
      },
      required: ["categories", "series"],
      additionalProperties: false
    }
  },
  {
    name: "edit_slide_master",
    description: "受控设置当前页背景或本地绘图默认偏好，不开放任意 master XML。支持 backgroundColor、fontFamily、themeColors、decorations。backgroundColor 会通过单页 XML 替换，因此可能返回新 slideId。",
    input_schema: {
      type: "object",
      properties: {
        slideId: { type: "string" },
        slideIndex: { type: "number" },
        pageNumber: { type: "number" },
        backgroundColor: { type: "string" },
        fontFamily: { type: "string" },
        themeColors: { type: "object" },
        decorations: { type: "array", items: { type: "object" } }
      },
      additionalProperties: false
    }
  },
  {
    name: "store_blob",
    description: "把 base64/url/text 资源存入本地 IndexedDB，返回 blobName/mime/size。用于后续资源复用；不会上传远端。",
    input_schema: {
      type: "object",
      properties: {
        blobName: { type: "string" },
        name: { type: "string" },
        source: { type: "string", enum: ["base64", "url", "text"] },
        base64: { type: "string" },
        url: { type: "string" },
        content: { type: "string" },
        mimeType: { type: "string" }
      },
      additionalProperties: false
    }
  },
  {
    name: "copy_image_between_slides",
    description: "复制图片形状到另一页。来源必须是可导出为图片的 shape；内部使用 getImageAsBase64 + addImage，支持目标 left/top/width/height。",
    input_schema: {
      type: "object",
      properties: {
        fromSlideId: { type: "string" },
        fromSlideIndex: { type: "number" },
        fromPageNumber: { type: "number" },
        from_slide_id: { type: "string" },
        from_slide_index: { type: "number" },
        from_page_number: { type: "number" },
        fromShapeId: { type: "string" },
        from_shape_id: { type: "string" },
        toSlideId: { type: "string" },
        toSlideIndex: { type: "number" },
        toPageNumber: { type: "number" },
        to_slide_id: { type: "string" },
        to_slide_index: { type: "number" },
        to_page_number: { type: "number" },
        left: { type: "number" },
        top: { type: "number" },
        x: { type: "number" },
        y: { type: "number" },
        width: { type: "number" },
        height: { type: "number" }
      },
      required: ["fromShapeId"],
      additionalProperties: false
    }
  },
  {
    name: "search_icons",
    description: "搜索内置小型 SVG 图标库，返回 icon_id、描述和匹配度。第一版不接微软远端图标服务。插入前应先调用此工具。",
    input_schema: {
      type: "object",
      properties: {
        query: { type: "string" },
        top: { type: "number" }
      },
      required: ["query"],
      additionalProperties: false
    }
  },
  {
    name: "insert_icon",
    description: "把 search_icons 返回的内置 SVG 图标插入目标页，支持颜色、位置、尺寸和 description。目标页可用 slideId、slideIndex 或 pageNumber 指定。",
    input_schema: {
      type: "object",
      properties: {
        iconId: { type: "string" },
        icon_id: { type: "string" },
        slideId: { type: "string" },
        slideIndex: { type: "number" },
        pageNumber: { type: "number" },
        x: { type: "number" },
        y: { type: "number" },
        left: { type: "number" },
        top: { type: "number" },
        width: { type: "number" },
        height: { type: "number" },
        color: { type: "string" },
        description: { type: "string" }
      },
      additionalProperties: false
    }
  },
  {
    name: "web_search",
    description: "通过本地 dev server 的 __web-search endpoint 做轻量搜索，返回标题、链接、摘要。不可用时返回明确错误，不假装联网成功。",
    input_schema: {
      type: "object",
      properties: {
        query: { type: "string" },
        top: { type: "number" }
      },
      required: ["query"],
      additionalProperties: false
    }
  },
  {
    name: "update_instructions",
    description: "用查找替换方式修改本地长期偏好，保存到 localStorage 并拼接进后续 Agent system prompt。oldText 为空表示追加。",
    input_schema: {
      type: "object",
      properties: {
        operations: {
          type: "array",
          items: {
            type: "object",
            properties: {
              oldText: { type: "string" },
              newText: { type: "string" },
              old_text: { type: "string" },
              new_text: { type: "string" }
            }
          }
        }
      },
      required: ["operations"],
      additionalProperties: false
    }
  },
  {
    name: "update_setting",
    description: "读取或更新本地工具设置，如 web_search、review_required、debug_artifacts。传 value 时写入；不传 value 时读取当前值。",
    input_schema: {
      type: "object",
      properties: {
        setting: { type: "string" },
        value: { type: "boolean" }
      },
      required: ["setting"],
      additionalProperties: false
    }
  },
  {
    name: "read_skill",
    description: "读取仓库 skills/<skillName>.md 的技能文件。仅读取本地仓库技能，不访问远端技能市场。",
    input_schema: {
      type: "object",
      properties: {
        skillName: { type: "string" },
        skill_name: { type: "string" }
      },
      additionalProperties: false
    }
  },
  {
    name: "create_skill",
    description: "在仓库 skills/ 下创建一个 SKILL 草稿 Markdown 文件，返回文件路径。",
    input_schema: {
      type: "object",
      properties: {
        name: { type: "string" },
        description: { type: "string" },
        instructions: { type: "string" }
      },
      required: ["name", "description", "instructions"],
      additionalProperties: false
    }
  },
  {
    name: "get_connected_agents",
    description: "unsupported：当前本地 add-in 未配置多 Agent/MCP 桥接协议。调用会返回明确不可用说明。",
    input_schema: { type: "object", properties: {}, additionalProperties: false }
  },
  {
    name: "send_message",
    description: "unsupported：当前本地 add-in 未配置多 Agent/MCP 桥接协议。调用会返回明确不可用说明。",
    input_schema: {
      type: "object",
      properties: {
        agent_id: { type: "string" },
        message: { type: "string" }
      },
      additionalProperties: false
    }
  },
  {
    name: "refresh_mcp_connectors",
    description: "unsupported：当前本地 add-in 未配置 MCP 连接器刷新协议。调用会返回明确不可用说明。",
    input_schema: { type: "object", properties: {}, additionalProperties: false }
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
  read_slide_text: readSlideText,
  edit_slide_text: editSlideText,
  edit_slide_xml: editSlideXml,
  modify_shape: modifyShape,
  delete_shape: deleteShape,
  review_slide: reviewSlide,
  verify_slide_visual: reviewSlide,
  verify_slides: verifySlides,
  execute_office_js: executeOfficeJs,
  edit_slide_chart: editSlideChart,
  edit_slide_master: editSlideMaster,
  store_blob: storeBlob,
  copy_image_between_slides: copyImageBetweenSlides,
  search_icons: searchIcons,
  insert_icon: insertIcon,
  web_search: webSearch,
  update_instructions: updateInstructions,
  update_setting: updateSetting,
  read_skill: readSkill,
  create_skill: createSkill,
  get_connected_agents: unsupportedBridgeTool("get_connected_agents"),
  send_message: unsupportedBridgeTool("send_message"),
  refresh_mcp_connectors: unsupportedBridgeTool("refresh_mcp_connectors"),
  todo_write: todoWrite
};
