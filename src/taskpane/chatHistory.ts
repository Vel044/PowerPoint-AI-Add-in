import { Message } from "../types";

const STORAGE_KEY = "claude-for-office.chats";
const MAX_SESSIONS = 50;
const TITLE_MAX_LEN = 60;

export interface ChatSession {
  id: string;
  title: string;
  createdAt: number;
  updatedAt: number;
  messages: Message[];
}

function readAll(): ChatSession[] {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return [];
  try {
    const arr = JSON.parse(raw) as ChatSession[];
    return Array.isArray(arr) ? arr : [];
  } catch {
    return [];
  }
}

function writeAll(sessions: ChatSession[]): void {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(sessions));
}

export function listSessions(): ChatSession[] {
  return readAll().sort((a, b) => b.updatedAt - a.updatedAt);
}

export function getSession(id: string): ChatSession | undefined {
  return readAll().find((s) => s.id === id);
}

export function createSession(): ChatSession {
  const now = Date.now();
  return {
    id: `${now}-${Math.random().toString(36).slice(2, 8)}`,
    title: "新对话",
    createdAt: now,
    updatedAt: now,
    messages: []
  };
}

export function saveSession(session: ChatSession): void {
  const all = readAll();
  const idx = all.findIndex((s) => s.id === session.id);
  const updated = { ...session, updatedAt: Date.now() };
  if (idx >= 0) {
    all[idx] = updated;
  } else {
    all.push(updated);
  }
  // 只保留最新的 MAX_SESSIONS 条
  all.sort((a, b) => b.updatedAt - a.updatedAt);
  writeAll(all.slice(0, MAX_SESSIONS));
}

export function deleteSession(id: string): void {
  writeAll(readAll().filter((s) => s.id !== id));
}

export function deriveTitle(messages: Message[]): string {
  for (const m of messages) {
    if (m.role !== "user") continue;
    const text = extractUserText(m);
    if (text) {
      const trimmed = text.replace(/\s+/g, " ").trim();
      return trimmed.length > TITLE_MAX_LEN ? trimmed.slice(0, TITLE_MAX_LEN) + "…" : trimmed;
    }
  }
  return "新对话";
}

// 用户消息可能是纯字符串（附带 [当前上下文] 前缀）或 tool_result 数组，
// 这里只抽取真实用户文字用于标题
export function extractUserText(m: Message): string {
  if (typeof m.content !== "string") return "";
  const marker = "[用户消息]\n";
  const i = m.content.indexOf(marker);
  return i >= 0 ? m.content.slice(i + marker.length) : m.content;
}

export function formatRelativeTime(ts: number, now = Date.now()): string {
  const diff = Math.max(0, now - ts);
  const sec = Math.floor(diff / 1000);
  if (sec < 60) return "just now";
  const min = Math.floor(sec / 60);
  if (min < 60) return `${min}m`;
  const hr = Math.floor(min / 60);
  if (hr < 24) return `${hr}h`;
  const day = Math.floor(hr / 24);
  if (day < 30) return `${day}d`;
  const mo = Math.floor(day / 30);
  if (mo < 12) return `${mo}mo`;
  return `${Math.floor(mo / 12)}y`;
}
