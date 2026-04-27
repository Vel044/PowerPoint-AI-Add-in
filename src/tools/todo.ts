import { ToolHandler } from "../types";

type TodoStatus = "pending" | "in_progress" | "completed";

interface TodoItem {
  content?: string;
  activeForm?: string;
  status?: TodoStatus;
}

export const todoWrite: ToolHandler = async (input) => {
  const todos = Array.isArray(input.todos) ? input.todos as TodoItem[] : [];
  if (todos.length === 0) throw new Error("todo_write 需要 todos 数组");
  const lines = todos.map((todo, index) => {
    const status = normalizeStatus(todo.status);
    const marker = status === "completed" ? "[x]" : status === "in_progress" ? "[>]" : "[ ]";
    const text = status === "in_progress" && todo.activeForm ? todo.activeForm : todo.content;
    return `${marker} ${index + 1}. ${text ?? "未命名任务"}`;
  });
  return `任务进度已更新:\n${lines.join("\n")}`;
};

function normalizeStatus(value: unknown): TodoStatus {
  if (value === "completed" || value === "in_progress" || value === "pending") return value;
  return "pending";
}
