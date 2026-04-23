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
3. 修改完成后用一句话告诉用户做了什么。
4. 如果涉及底层 XML 级别的操作（如改 theme、master、复杂动画），使用 export_pptx_xml 导出查看，用 apply_pptx_patch 产出新文件让用户另存。
5. 优先使用 Office.js 级工具（add_slide / modify_shape 等）做实时编辑，体验最流畅。`;

export async function runAgent(
  userMessage: string,
  history: Message[],
  options: AgentOptions
): Promise<Message[]> {
  const max = options.maxIterations ?? 10;
  const emit = options.onEvent ?? (() => {});
  const messages: Message[] = [...history, { role: "user", content: userMessage }];
  const toolCtx: ToolContext = { log: (m) => emit({ type: "text", text: m }) };

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

    const assistantBlocks: ContentBlock[] = res.content.map((b) => {
      if (b.type === "text") return { type: "text", text: b.text };
      return { type: "tool_use", id: b.id, name: b.name, input: b.input };
    });
    messages.push({ role: "assistant", content: assistantBlocks });

    for (const b of assistantBlocks) {
      if (b.type === "text" && b.text) emit({ type: "text", text: b.text });
      if (b.type === "tool_use") emit({ type: "tool_call", toolName: b.name, toolInput: b.input });
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

export async function runAgentStream(
  userMessage: string,
  history: Message[],
  options: AgentOptions
): Promise<Message[]> {
  const max = options.maxIterations ?? 10;
  const emit = options.onEvent ?? (() => {});
  const messages: Message[] = [...history, { role: "user", content: userMessage }];
  const toolCtx: ToolContext = { log: (m) => emit({ type: "text", text: m }) };

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
            resolve();
          } else if (ev.type === "error") {
            reject(new Error(ev.error));
          }
        }
      ).catch(reject);
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
