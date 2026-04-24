import { ContentBlock, Message, ToolContext, ToolDefinition, ToolHandler } from "../types";
import { callMessages, callMessagesStream, MessagesResponse, StreamEvent } from "./client";
import { ModelTier } from "../config";

export interface AgentEvent {
  type: "text" | "tool_call" | "tool_result" | "error" | "done";
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
  onEvent?: (ev: AgentEvent) => void;
}

const DEFAULT_SYSTEM = `你是一个嵌入在 PowerPoint 右侧任务窗格中的 AI 助手。
你可以调用工具读取当前演示文稿状态（当前幻灯片、选中的形状等），并对演示文稿做增删改查。
规则：
1. 回答使用简体中文。
2. 在执行任何修改操作前，先调用 get_current_context 了解用户当前处于哪张幻灯片、选中了什么。
3. 画流程图/调用链/架构图/逻辑框图等"图"时，**首选 create_diagram**（一次传入节点和连线的抽象描述，内部自动布局、节点不重叠、箭头自动贴边中点）。只有修改现有图、或往已有图上增删节点时，才用 add_geometric_shape + connect_shapes 组合。
4. **不要自己手算坐标去连两个形状**——用 connect_shapes(fromShapeId, fromSide, toShapeId, toSide) 让工具算边中点。你只需要想清楚"A 的哪条边连到 B 的哪条边"。也禁止只用一堆 add_text_box 堆砌伪缩进，那不是图。文本框只用来放标题或长段说明。
5. **连接线限制**：受 Office.js 限制，connect_shapes 和 create_diagram 画出的都是**纯直线**，不带箭头、也不会随形状移动自动跟随。如果用户要求"带箭头"、"流向"、"方向"，要明确告知这一限制，不要虚假承诺箭头效果。
6. 节点统一尺寸、同类节点 shape 保持一致：过程用 rectangle、判断用 diamond、起止用 flowChartTerminator、数据用 flowChartData。
7. 所有修改操作做完后，必须再生成一段纯文本回复（不再调用工具），用 2-4 句话总结你刚画了什么、用户可以怎么调整。
8. 可以多轮调用工具，但完成后一定主动停下来写总结，而不是无意义地继续调用。
9. 涉及底层 XML 操作（改 theme、master、复杂动画）时用 export_pptx_xml 导出查看、用 apply_pptx_patch 产出新文件。
10. 每次做完视觉修改（画图、添加形状、连线、create_diagram）后，**必须调用 review_slide 截图检查**。如果审查发现问题，立即修正后再次检查。最多自查 2 轮。

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
      res = await callMessages({
        system: options.system ?? DEFAULT_SYSTEM,
        messages,
        tools: options.tools,
        tier: options.tier,
        modelOverride: options.modelOverride
      });
    } catch (e) {
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
        const outStr = out ?? "";
        const outPreview = outStr.slice(0, 300);
        logToTerminal("info", `[Agent] ← tool_result: ${call.name} ok len=${outStr.length} preview=${outPreview}`);
        results.push({ type: "tool_result", tool_use_id: call.id, content: out });
        emit({ type: "tool_result", toolName: call.name, toolResult: out });
      } catch (e) {
        const msg = e instanceof Error ? e.message : String(e);
        logToTerminal("error", `[Agent] ← tool_result: ${call.name} FAILED: ${msg}`);
        results.push({ type: "tool_result", tool_use_id: call.id, content: msg, is_error: true });
        emit({ type: "tool_result", toolName: call.name, toolResult: msg, isError: true });
      }
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
      callMessagesStream(
        {
          system: options.system ?? DEFAULT_SYSTEM,
          messages,
          tools: options.tools,
          tier: options.tier,
          modelOverride: options.modelOverride
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

    await streamPromise;

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
        results.push({ type: "tool_result", tool_use_id: call.id, content: out });
        emit({ type: "tool_result", toolName: call.name, toolResult: out });
      } catch (e) {
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
