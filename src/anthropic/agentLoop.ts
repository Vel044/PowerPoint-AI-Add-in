import { ContentBlock, Message, ToolContext, ToolDefinition, ToolHandler } from "../types";
import {
  callMessages,
  callMessagesStream,
  isRequestCancelled,
  MessagesResponse,
  RequestCancelledError,
  StreamEvent
} from "./client";
import { ModelTier } from "../config";
import { getUserInstructions } from "../tools/remainingTools";

export interface AgentEvent {
  type: "text" | "tool_call" | "tool_result" | "error" | "done" | "cancelled";
  text?: string;
  toolName?: string;
  toolInput?: Record<string, unknown>;
  toolResult?: string;
  isError?: boolean;
}

export interface AgentOptions {
  system?: string;
  tools: ToolDefinition[];
  handlers: Record<string, ToolHandler>;
  maxIterations?: number;
  tier?: ModelTier;
  modelOverride?: string;
  signal?: AbortSignal;
  onEvent?: (ev: AgentEvent) => void;
}

const DEFAULT_SYSTEM = `你是一个嵌入在 PowerPoint 右侧任务窗格中的 AI 助手。
你可以调用工具读取当前演示文稿状态（当前幻灯片、选中的形状等），并对演示文稿做增删改查。
规则：
1. 回答使用简体中文。
2. 在执行任何修改操作前，先调用 list_slide_shapes 或 get_current_context。用户提到“第 N 页/第三页/某一页”时，应传 pageNumber 静默查看目标页，而不是要求用户手动切页。删除、修改、读富文本前必须使用目标页上下文中的 ref 或 slideId + shapeId，不得凭裸 shapeId 猜测，也不得跨页搜索。
3. 画流程图、调用链、架构图、逻辑框图时，先用 list_slide_shapes/get_current_context 看清页面，再优先用 Claude 风格 edit_slide_xml 的 code 模式批量插入形状和连接器；需要少量补充时再用 add_geometric_shape/add_text_box/connect_shapes 组合。
4. 图表优先用 edit_slide_chart；插入图标先 search_icons 再 insert_icon；需要细调已有形状的位置、尺寸、层级、选择状态时用 execute_office_js 的白名单 actions。
5. 修改现有图或少量补线时，用 add_geometric_shape/add_text_box/modify_shape/connect_shapes 组合。连接两个已有形状时优先 connect_shapes(fromShapeId, fromSide, toShapeId, toSide)，不要用普通直线假装连接器。也禁止只用一堆文本框堆砌伪缩进，那不是图。
6. **连接器**：connect_shapes 会使用真实 PowerPoint connector 画贴边中点连接线；端点同 X/Y 时优先 Straight，否则使用 bentConnector3 肘形连接器。arrow="end" 会通过 PowerPoint 原生 tailEnd 箭头指向终点形状。不要手动画三角形箭头。
7. get_connected_agents、send_message、refresh_mcp_connectors 是明确 unsupported 的外部桥接占位；不要把它们当作真实跨 Office 协作能力使用。
8. 节点统一尺寸、同类节点 shape 保持一致：过程用 rectangle、判断用 diamond、起止用 flowChartTerminator、数据/输入输出用 flowChartInputOutput；数据库也可以用 can。历史别名 flowChartData 可用，但优先不要主动生成。
9. **样式默认值**：不要指定 fillColor/lineColor/textColor 等样式参数，除非用户明确要求特定颜色。省略时形状自动继承演示文稿主题样式，效果更协调。使用 edit_slide_xml 写 OOXML 时，不要在 <a:solidFill> 中硬编码 srgbClr；如需指定填充，优先用 <a:schemeClr val="accent1"/> 等主题色引用。
10. 所有修改操作做完后，必须再生成一段纯文本回复（不再调用工具），用 2-4 句话总结你刚画了什么、用户可以怎么调整。
11. 可以多轮调用工具，但完成后一定主动停下来写总结，而不是无意义地继续调用。
12. connect_shapes 为了修正真实连接器 XML，会进行单页 export/import，执行后 slideId 可能变化；如果 tool result 返回"新 slideId=..."，后续 review_slide/修改必须使用这个新 slideId，或不传 slideId 直接操作当前选中页。
13. 替换某个现有图形时，先调用 get_current_context（必要时带 pageNumber），确认目标形状的 slideId、shapeId 和 bounds；删除时传 slideId 或 pageNumber；随后用删除前 bounds 规划 edit_slide_xml 或新增形状的坐标，确保新图落在原范围内。
14. 每次做完视觉修改（画图、添加形状、连线、edit_slide_xml/edit_slide_chart/insert_icon）后，先调用 verify_slides 做程序化检查，再调用 verify_slide_visual 或 review_slide 截图检查。截图会保存到 debug-artifacts/review-slide/ 并返回路径。审查时 question 里写"请只关注最近添加的形状，忽略旧内容"。如果审查发现问题（重叠、歪斜、溢出），**必须立即修正**（用 modify_shape 移动位置、delete_shape 删除重画等），然后再审查确认。最多自查 2 轮。**不要在审查发现问题后就停止——必须修正后才能结束。**

## 画布尺寸与坐标系

标准 16:9 幻灯片尺寸约为 960×540 点(pt)，(0,0) 在左上角。get_current_context 返回的 slideWidth/slideHeight 是当前演示文稿实际尺寸，请用它规划布局。形状宽高通常 120-200pt，间距 40-60pt，留边距 50pt。画图前必须先调用 get_current_context 确认画布尺寸和已有形状位置。

## 主题色板

get_current_context 返回 themePalette 对象（默认 Office 主题时为 null）。可用键：dk1（主文字色）、dk2（次要深色）、lt1（主背景色）、lt2（次要浅色）、accent1-6（六个强调色）、hlink/folHlink（超链接）。值为裸 hex 字符串（不带 #）。

使用规则：
1. 用户要求"用品牌色/模板配色"且 isDefaultTheme 为 false 时，从 themePalette 取对应键的 hex 值，传给需要 hex 参数的工具（如 add_geometric_shape 的 fillColor、insert_icon 的 color）。
2. 在 OOXML 写入时优先用 schemeClr 语义引用（accent1 等），themePalette 仅作为应急的 hex 兜底。
3. themePalette 为 null 或 isDefaultTheme 为 true 时，不要硬编码 hex，省略颜色参数让形状继承主题。

## 视觉上下文

每次对话开始时会附带当前幻灯片的截图。请参考截图了解幻灯片整体布局（背景、主题色、已有形状位置），以此规划新形状的放置和样式，确保与现有设计协调。

## edit_slide_xml code 模式

edit_slide_xml 是 Claude 风格的单页 OOXML 编辑工具，不再接受 operations[]。你传入的是 async JS 函数体 code，工具会把目标页导出为单页 PPTX zip，并在 code 作用域里提供：

- zip：JSZip 对象
- markDirty：调用后才会写回当前 PowerPoint
- pptx：安全绘图 helper，普通画图必须优先用它
- DOMParser / XMLSerializer：浏览器原生 XML API

普通绘图标准模板：

const start = pptx.addShape({ id: "start", shapeType: "flowChartTerminator", style: "entry", text: "开始", left: 400, top: 40, width: 160, height: 50 });
const input = pptx.addShape({ id: "input", shapeType: "flowChartInputOutput", style: "io", text: "输入用户名密码", left: 380, top: 120, width: 200, height: 60 });
pptx.addConnector({ from: start, fromSide: "bottom", to: input, toSide: "top", arrow: "end", style: "control" });
pptx.save();
markDirty();

pptx helper 规则：
1. 普通画图可直接用 pptx.addShape/addConnector/save；需要拿 session 时也可调用 await pptx.openSlide()。openSlide 固定读取 ppt/slides/slide1.xml；不要传 slide7.xml，也不要猜当前页文件名。
2. addShape 坐标单位是 pt，不是 EMU。支持 shapeType：rect、roundRect、ellipse、diamond、flowChartTerminator、flowChartInputOutput、flowChartDecision、can。
3. pptx.slideWidth / pptx.slideHeight 是当前幻灯片真实尺寸；所有 shape 必须满足 left>=0、top>=0、left+width<=slideWidth、top+height<=slideHeight。复杂图横向塞不下时必须缩小节点、换行、分组或拆页，禁止画到画布外。
4. addShape 支持 style 语义预设：entry、process、decision、success、danger/error、database、io、external、muted、note、laneHeader、title。画图时必须按语义选择 2-5 种 style，避免所有节点同一种蓝色。
5. addConnector 的 from/to 可传 addShape 返回对象或 shape id；fromSide/toSide 是 top/right/bottom/left；arrow 默认 end，使用原生 tailEnd 箭头。连接器 style 支持 control、data、dependency、success、danger、muted。
6. 样式策略：标题/泳道头用 laneHeader/title；普通过程用 process；判断用 decision；成功路径用 success；错误/失败路径用 danger；数据库用 database；注释说明用 note/muted。不要为了“好看”乱用超过 5 种颜色。
7. 只有 helper 覆盖不了的高级修改，才直接操作 zip/DOMParser。

底层 XML 模板：

const path = "ppt/slides/slide1.xml";
const xml = await zip.file(path).async("string");
const doc = new DOMParser().parseFromString(xml, "application/xml");
const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const spTree = doc.getElementsByTagNameNS(P_NS, "spTree")[0];
// 读取、创建、插入、替换 p:sp / p:cxnSp / p:pic 等节点
zip.file(path, new XMLSerializer().serializeToString(doc));
markDirty();

注意：
1. code 是函数体，不要再包一层 async function，也不要返回 Markdown。
2. 修改 slide XML 后必须调用 markDirty()；不调用则工具只执行检查，不写回。
3. 写回会通过单页 export/import 替换目标页，slideId 可能变化；后续工具必须使用返回的新 slideId 或重新 get_current_context。
4. 失败时会保存 debug-artifacts/edit-slide-xml/ 下的 before/after XML，便于复盘。

## OOXML 规范

直接写 OOXML 时必须遵守以下规则，否则 PowerPoint 会在 insertSlidesFromBase64 时抛 0x80004005 错误：

1. **schemeClr val** 只能取：dk1、lt1、accent1-6、tx1、bg1、phClr。禁止使用 txClr、fillColor、textColor、lineColor 等非 OOXML 值（会导致 PowerPoint 直接拒绝导入）。
2. **形状边框**：若不需要自定义边框，不要写 <a:ln> 元素，省略则继承主题默认边框。若需要显式 1pt 边框：<a:ln w="12700"><a:solidFill><a:schemeClr val="dk1"/></a:solidFill></a:ln>。
3. **形状填充**：省略 spPr 中的填充元素则继承主题色；透明填充用 <a:noFill/>；主题强调色用 <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>。
4. **形状 ID**：<p:cNvPr id> 必须在当前幻灯片内唯一；使用 2001、2002 ... 等大值避免与已有形状冲突。
5. **坐标单位**：<a:off x y> 和 <a:ext cx cy> 单位为 EMU（1pt = 12700 EMU）。

## 画图路线（正样本）

用户要求画图时，优先自己规划坐标和尺寸，然后用 edit_slide_xml({ code }) 里的 pptx helper 批量插入形状和连接器。少量增补或修正时再用 add_geometric_shape/add_text_box/connect_shapes/modify_shape。`;

function throwIfCancelled(signal?: AbortSignal): void {
  if (signal?.aborted) throw new RequestCancelledError();
}

function emitCancelled(emit: (ev: AgentEvent) => void): void {
  emit({ type: "cancelled", text: "已暂停" });
}

function buildSystemPrompt(system?: string): string {
  if (system) return system;
  const instructions = getUserInstructions().trim();
  return instructions ? `${DEFAULT_SYSTEM}\n\n## 用户长期偏好\n${instructions}` : DEFAULT_SYSTEM;
}

export async function runAgent(
  userMessage: string | ContentBlock[],
  history: Message[],
  options: AgentOptions
): Promise<Message[]> {
  const max = options.maxIterations ?? 50;
  const emit = options.onEvent ?? (() => {});
  const messages: Message[] = [...history, { role: "user", content: userMessage }];
  const systemPrompt = buildSystemPrompt(options.system);
  const logToTerminal = (level: string, msg: string) => {
    fetch("https://localhost:3001/__terminal-log", {
      method: "POST",
      body: JSON.stringify({ level, msg }),
      headers: { "Content-Type": "application/json" }
    }).catch(() => {});
  };
  const toolCtx: ToolContext = { log: (m: string) => { logToTerminal("info", m); emit({ type: "text", text: m }); } };

  for (let i = 0; i < max; i++) {
    let res: MessagesResponse;
    try {
      throwIfCancelled(options.signal);
      res = await callMessages({
        system: systemPrompt,
        messages,
        tools: options.tools,
        tier: options.tier,
        modelOverride: options.modelOverride,
        signal: options.signal
      });
      throwIfCancelled(options.signal);
    } catch (e) {
      if (isRequestCancelled(e)) {
        emitCancelled(emit);
        throw e;
      }
      const msg = e instanceof Error ? e.message : String(e);
      emit({ type: "error", text: msg });
      throw e;
    }

    const assistantBlocks: ContentBlock[] = [];
    for (const b of res.content as Array<Record<string, unknown>>) {
      if (b.type === "text" && typeof b.text === "string") {
        assistantBlocks.push({ type: "text", text: b.text });
      } else if (b.type === "tool_use" && typeof b.id === "string" && typeof b.name === "string") {
        assistantBlocks.push({
          type: "tool_use",
          id: b.id,
          name: b.name,
          input: (b.input as Record<string, unknown>) ?? {}
        });
      }
      // thinking / 未知类型：丢弃，不影响下一轮
    }
    messages.push({ role: "assistant", content: assistantBlocks });

    logToTerminal("info", `[Agent] stop_reason=${res.stop_reason} blocks=${assistantBlocks.length}`);
    for (const b of assistantBlocks) {
      if (b.type === "text" && b.text) emit({ type: "text", text: b.text });
      if (b.type === "tool_use") {
        const inputPreview = JSON.stringify(b.input).slice(0, 300);
        logToTerminal("info", `[Agent] → tool_use: ${b.name} input=${inputPreview}`);
        emit({ type: "tool_call", toolName: b.name, toolInput: b.input });
      }
    }

    if (res.stop_reason !== "tool_use") {
      emit({ type: "done" });
      return messages;
    }

    const toolUses = assistantBlocks.filter((b): b is Extract<ContentBlock, { type: "tool_use" }> => b.type === "tool_use");
    const results: ContentBlock[] = [];
    for (const call of toolUses) {
      try {
        throwIfCancelled(options.signal);
      } catch (e) {
        emitCancelled(emit);
        throw e;
      }
      const handler = options.handlers[call.name];
      if (!handler) {
        logToTerminal("warn", `[Agent] ← tool_result: ${call.name} (未知工具)`);
        results.push({
          type: "tool_result",
          tool_use_id: call.id,
          content: `未知工具: ${call.name}`,
          is_error: true
        });
        emit({ type: "tool_result", toolName: call.name, toolResult: "未知工具", isError: true });
        continue;
      }
      try {
        const out = await handler(call.input, toolCtx);
        throwIfCancelled(options.signal);
        const outStr = out ?? "";
        const outPreview = outStr.slice(0, 300);
        logToTerminal("info", `[Agent] ← tool_result: ${call.name} ok len=${outStr.length} preview=${outPreview}`);
        results.push({ type: "tool_result", tool_use_id: call.id, content: out });
        emit({ type: "tool_result", toolName: call.name, toolResult: out });
      } catch (e) {
        if (isRequestCancelled(e)) {
          emitCancelled(emit);
          throw e;
        }
        const msg = e instanceof Error ? e.message : String(e);
        logToTerminal("error", `[Agent] ← tool_result: ${call.name} FAILED: ${msg}`);
        results.push({ type: "tool_result", tool_use_id: call.id, content: msg, is_error: true });
        emit({ type: "tool_result", toolName: call.name, toolResult: msg, isError: true });
      }
    }
    try {
      throwIfCancelled(options.signal);
    } catch (e) {
      emitCancelled(emit);
      throw e;
    }
    messages.push({ role: "user", content: results });
  }

  emit({ type: "error", text: `已达到最大循环次数 ${max}，自动停止。` });
  return messages;
}

export async function runAgentStream(
  userMessage: string | ContentBlock[],
  history: Message[],
  options: AgentOptions
): Promise<Message[]> {
  const max = options.maxIterations ?? 50;
  const emit = options.onEvent ?? (() => {});
  const messages: Message[] = [...history, { role: "user", content: userMessage }];
  const systemPrompt = buildSystemPrompt(options.system);
  const logToTerminal = (level: string, msg: string) => {
    fetch("https://localhost:3001/__terminal-log", {
      method: "POST",
      body: JSON.stringify({ level, msg }),
      headers: { "Content-Type": "application/json" }
    }).catch(() => {});
  };
  const toolCtx: ToolContext = { log: (m: string) => { logToTerminal("info", m); emit({ type: "text", text: m }); } };

  for (let i = 0; i < max; i++) {
    let textBuffer = "";
    let pendingToolUse: { id: string; name: string; input: Record<string, unknown> } | null = null;

    const streamPromise = new Promise<void>((resolve, reject) => {
      try {
        throwIfCancelled(options.signal);
      } catch (e) {
        reject(e);
        return;
      }
      callMessagesStream(
        {
          system: systemPrompt,
          messages,
          tools: options.tools,
          tier: options.tier,
          modelOverride: options.modelOverride,
          signal: options.signal
        },
        (ev: StreamEvent) => {
          if (ev.type === "text" && ev.text) {
            textBuffer += ev.text;
            emit({ type: "text", text: ev.text });
          } else if (ev.type === "tool_use" && ev.toolName) {
            pendingToolUse = {
              id: Math.random().toString(36).slice(2),
              name: ev.toolName,
              input: ev.toolInput ?? {}
            };
            emit({ type: "tool_call", toolName: ev.toolName, toolInput: ev.toolInput });
          } else if (ev.type === "done") {
            logToTerminal("info", "[Agent] 流收到 done 事件");
            resolve();
          } else if (ev.type === "error") {
            logToTerminal("error", `[Agent] 流收到 error 事件: ${ev.error}`);
            reject(new Error(ev.error));
          }
        }
      ).catch((err) => {
        logToTerminal("error", `[Agent] callMessagesStream 异常: ${err}`);
        reject(err);
      });
    });

    try {
      await streamPromise;
      throwIfCancelled(options.signal);
    } catch (e) {
      if (isRequestCancelled(e)) {
        emitCancelled(emit);
        throw e;
      }
      throw e;
    }

    const assistantBlocks: ContentBlock[] = [];
    if (textBuffer) {
      assistantBlocks.push({ type: "text", text: textBuffer });
    }
    if (pendingToolUse) {
      assistantBlocks.push(pendingToolUse);
    }

    if (assistantBlocks.length === 0) {
      return messages;
    }

    messages.push({ role: "assistant", content: assistantBlocks });

    if (assistantBlocks[0].type === "text") {
      emit({ type: "done" });
      return messages;
    }

    if (assistantBlocks[0].type !== "tool_use") {
      emit({ type: "done" });
      return messages;
    }

    const toolUses = assistantBlocks.filter((b): b is Extract<ContentBlock, { type: "tool_use" }> => b.type === "tool_use");
    const results: ContentBlock[] = [];
    for (const call of toolUses) {
      try {
        throwIfCancelled(options.signal);
      } catch (e) {
        emitCancelled(emit);
        throw e;
      }
      const handler = options.handlers[call.name];
      if (!handler) {
        results.push({
          type: "tool_result",
          tool_use_id: call.id,
          content: `未知工具: ${call.name}`,
          is_error: true
        });
        emit({ type: "tool_result", toolName: call.name, toolResult: "未知工具", isError: true });
        continue;
      }
      try {
        const out = await handler(call.input, toolCtx);
        throwIfCancelled(options.signal);
        results.push({ type: "tool_result", tool_use_id: call.id, content: out });
        emit({ type: "tool_result", toolName: call.name, toolResult: out });
      } catch (e) {
        if (isRequestCancelled(e)) {
          emitCancelled(emit);
          throw e;
        }
        const msg = e instanceof Error ? e.message : String(e);
        results.push({ type: "tool_result", tool_use_id: call.id, content: msg, is_error: true });
        emit({ type: "tool_result", toolName: call.name, toolResult: msg, isError: true });
      }
    }
    messages.push({ role: "user", content: results });
  }

  emit({ type: "error", text: `已达到最大循环次数 ${max}，自动停止。` });
  return messages;
}
