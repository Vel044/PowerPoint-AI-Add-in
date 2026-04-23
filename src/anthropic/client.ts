import {
  getActiveProvider,
  getTimeoutMs,
  loadConfig,
  ModelTier,
  Provider,
  resolveModel
} from "../config";
import { Message, ToolDefinition } from "../types";

export interface MessagesRequest {
  system?: string;
  messages: Message[];
  tools?: ToolDefinition[];
  max_tokens?: number;
  tier?: ModelTier;
  modelOverride?: string;
}

export interface MessagesResponse {
  id: string;
  type: "message";
  role: "assistant";
  model: string;
  stop_reason: "end_turn" | "tool_use" | "max_tokens" | "stop_sequence" | string;
  content: Array<
    | { type: "text"; text: string }
    | { type: "tool_use"; id: string; name: string; input: Record<string, unknown> }
  >;
  usage?: { input_tokens: number; output_tokens: number };
}

export async function callMessages(req: MessagesRequest): Promise<MessagesResponse> {
  const config = await loadConfig();
  const provider = getActiveProvider(config);
  return callMessagesWithProvider(provider, req);
}

export interface StreamEvent {
  type: "text" | "tool_use" | "done" | "error";
  text?: string;
  toolName?: string;
  toolInput?: Record<string, unknown>;
  error?: string;
}

export async function callMessagesStream(
  req: MessagesRequest,
  onEvent: (ev: StreamEvent) => void
): Promise<void> {
  const config = await loadConfig();
  const provider = getActiveProvider(config);
  const base = provider.env.ANTHROPIC_BASE_URL.replace(/\/$/, "");
  const url = `${base}/v1/messages`;
  const model = req.modelOverride ?? resolveModel(provider, req.tier ?? "sonnet");
  const token = provider.env.ANTHROPIC_AUTH_TOKEN;
  if (!token) throw new Error("当前 Provider 未配置 ANTHROPIC_AUTH_TOKEN");

  const headers: Record<string, string> = {
    "Content-Type": "application/json",
    "anthropic-version": "2023-06-01",
    "x-api-key": token,
    Authorization: `Bearer ${token}`
  };

  const body: Record<string, unknown> = {
    model,
    max_tokens: req.max_tokens ?? 4096,
    messages: req.messages,
    stream: true
  };
  if (req.system) body.system = req.system;
  if (req.tools && req.tools.length > 0) body.tools = req.tools;

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), getTimeoutMs(provider));

  try {
    const res = await fetch(url, {
      method: "POST",
      headers,
      body: JSON.stringify(body),
      signal: controller.signal
    });
    if (!res.ok) {
      const text = await res.text();
      throw new Error(`API ${res.status}: ${text.slice(0, 500)}`);
    }

    const reader = (res.body as ReadableStream<Uint8Array>).getReader();
    const decoder = new TextDecoder();
    let buffer = "";

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;

      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split("\n");
      buffer = lines.pop() ?? "";

      for (const line of lines) {
        if (line.startsWith("data: ")) {
          const data = line.slice(6);
          if (data === "[DONE]") {
            onEvent({ type: "done" });
            return;
          }
          try {
            const event = JSON.parse(data);
            if (event.type === "content_block_delta") {
              if (event.delta.type === "text_delta") {
                onEvent({ type: "text", text: event.delta.text });
              } else if (event.delta.type === "input_json_delta") {
                // tool use partial - ignore for now
              }
            } else if (event.type === "message_delta" && event.delta.stop_reason) {
              onEvent({ type: "done" });
            } else if (event.type === "error") {
              onEvent({ type: "error", error: event.error });
            }
          } catch {
            // ignore parse errors
          }
        }
      }
    }
    onEvent({ type: "done" });
  } finally {
    clearTimeout(timeout);
  }
}

export async function callMessagesWithProvider(
  provider: Provider,
  req: MessagesRequest
): Promise<MessagesResponse> {
  const base = provider.env.ANTHROPIC_BASE_URL.replace(/\/$/, "");
  const url = `${base}/v1/messages`;
  const model = req.modelOverride ?? resolveModel(provider, req.tier ?? "sonnet");
  const token = provider.env.ANTHROPIC_AUTH_TOKEN;
  if (!token) throw new Error("当前 Provider 未配置 ANTHROPIC_AUTH_TOKEN");

  const headers: Record<string, string> = {
    "Content-Type": "application/json",
    "anthropic-version": "2023-06-01",
    "x-api-key": token,
    Authorization: `Bearer ${token}`
  };

  const body: Record<string, unknown> = {
    model,
    max_tokens: req.max_tokens ?? 4096,
    messages: req.messages
  };
  if (req.system) body.system = req.system;
  if (req.tools && req.tools.length > 0) body.tools = req.tools;

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), getTimeoutMs(provider));

  try {
    const res = await fetch(url, {
      method: "POST",
      headers,
      body: JSON.stringify(body),
      signal: controller.signal
    });
    if (!res.ok) {
      const text = await res.text();
      throw new Error(`API ${res.status}: ${text.slice(0, 500)}`);
    }
    return (await res.json()) as MessagesResponse;
  } finally {
    clearTimeout(timeout);
  }
}
