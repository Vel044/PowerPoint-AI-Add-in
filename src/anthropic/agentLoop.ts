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
2. 在执行任何修改操作前，先调用 get_current_context 了解用户当前处于哪张幻灯片、当前页 allShapes 里有哪些形状、每个形状的位置/尺寸/文字，以及选中了什么。删除或修改形状时必须使用当前页上下文中的 slideId + shapeId，不得凭裸 shapeId 猜测，也不得跨页搜索。
3. 画流程图/调用链/架构图/逻辑框图等"图"时，**首选 create_diagram**（一次传入节点和连线的抽象描述，内部自动布局、节点不重叠、连接线自动贴边中点）。只有修改现有图、或往已有图上增删节点时，才用 add_geometric_shape + connect_shapes 组合。
4. **不要自己手算坐标去连两个形状**——用 connect_shapes(fromShapeId, fromSide, toShapeId, toSide) 让工具算边中点。你只需要想清楚"A 的哪条边连到 B 的哪条边"。也禁止只用一堆 add_text_box 堆砌伪缩进，那不是图。文本框只用来放标题或长段说明。
5. **连接器**：connect_shapes 和 create_diagram 会使用真实 PowerPoint connector 画贴边中点连接线；端点同 X/Y 时优先 Straight，否则使用 bentConnector3 肘形连接器。arrow="end" 会通过 PowerPoint 原生 tailEnd 箭头指向终点形状。不要手动画三角形箭头。
6. 节点统一尺寸、同类节点 shape 保持一致：过程用 rectangle、判断用 diamond、起止用 flowChartTerminator、数据/输入输出用 flowChartInputOutput；数据库也可以用 can。历史别名 flowChartData 可用，但优先不要主动生成。
7. 所有修改操作做完后，必须再生成一段纯文本回复（不再调用工具），用 2-4 句话总结你刚画了什么、用户可以怎么调整。
8. 可以多轮调用工具，但完成后一定主动停下来写总结，而不是无意义地继续调用。
9. connect_shapes/create_diagram 为了修正真实连接器 XML，会进行单页 export/import，执行后 slideId 可能变化；如果 tool result 返回"新 slideId=..."，后续 review_slide/修改必须使用这个新 slideId，或不传 slideId 直接操作当前选中页。
10. 替换某个现有图形时，先调用 get_current_context，确认目标形状的 slideId、shapeId 和 bounds；删除时传 slideId；随后把删除前 bounds 作为 create_diagram.canvas，确保新图落在原范围内。
11. 每次做完视觉修改（画图、添加形状、连线、create_diagram）后，**必须调用 review_slide 截图检查**。review_slide 会把截图保存到本地 debug-artifacts/review-slide/ 并在结果里返回路径，便于复盘模型到底看到了什么。审查时 question 里写"请只关注最近添加的形状，忽略旧内容"。如果审查发现问题（重叠、歪斜、溢出），**必须立即修正**（用 modify_shape 移动位置、delete_shape 删除重画等），然后再 review_slide 确认。最多自查 2 轮。**不要在审查发现问题后就停止——必须修正后才能结束。**

## 画图示例（正样本）

用户："画一个登录流程图"。正确做法是**一次 create_diagram 调用**：

\`\`\`
create_diagram({
  layout: "vertical",
  nodes: [
    {id:"n1", text:"开始", shape:"flowChartTerminator"},
    {id:"n2", text:"输入用户名/密码", shape:"rectangle"},
    {id:"n3", text:"验证通过？", shape:"diamond"},
    {id:"n4", text:"进入首页", shape:"rectangle"},
    {id:"n5", text:"提示错误", shape:"rectangle"},
    {id:"n6", text:"结束", shape:"flowChartTerminator"}
  ],
  edges: [
    {from:"n1", to:"n2"}, {from:"n2", to:"n3"},
    {from:"n3", to:"n4"}, {from:"n3", to:"n5"},
    {from:"n4", to:"n6"}, {from:"n5", to:"n6"}
  ]
})
\`\`\`

分层架构图/调用链用 layout:"layered" 并给每个节点指定 level（0 为最顶层）。树状展开用 layout:"tree"，工具会按 edges 自动推断父子层级。`;

function throwIfCancelled(signal?: AbortSignal): void {
  if (signal?.aborted) throw new RequestCancelledError();
}

function emitCancelled(emit: (ev: AgentEvent) => void): void {
  emit({ type: "cancelled", text: "已暂停" });
}

export async function runAgent(
  userMessage: string,
  history: Message[],
  options: AgentOptions
): Promise<Message[]> {
  const max = options.maxIterations ?? 50;
  const emit = options.onEvent ?? (() => {});
  const messages: Message[] = [...history, { role: "user", content: userMessage }];
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
        system: options.system ?? DEFAULT_SYSTEM,
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
  userMessage: string,
  history: Message[],
  options: AgentOptions
): Promise<Message[]> {
  const max = options.maxIterations ?? 50;
  const emit = options.onEvent ?? (() => {});
  const messages: Message[] = [...history, { role: "user", content: userMessage }];
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
          system: options.system ?? DEFAULT_SYSTEM,
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
