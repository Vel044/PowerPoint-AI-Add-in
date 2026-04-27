export type Role = "user" | "assistant";

export interface TextBlock { type: "text"; text: string }
export interface ToolUseBlock {
  type: "tool_use";
  id: string;
  name: string;
  input: Record<string, unknown>;
}
export interface ToolResultBlock {
  type: "tool_result";
  tool_use_id: string;
  content: string;
  is_error?: boolean;
}
export interface ImageBlock {
  type: "image";
  source: {
    type: "base64";
    media_type: "image/png";
    data: string;
  };
}
export type ContentBlock = TextBlock | ToolUseBlock | ToolResultBlock | ImageBlock;

export interface Message {
  role: Role;
  content: string | ContentBlock[];
}

export interface ToolDefinition {
  name: string;
  description: string;
  input_schema: Record<string, unknown>;
}

export interface ToolContext {
  log: (msg: string) => void;
}

export type ToolHandler = (
  input: Record<string, unknown>,
  ctx: ToolContext
) => Promise<string>;
