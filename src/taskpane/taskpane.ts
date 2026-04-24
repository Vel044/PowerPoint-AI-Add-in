import "./styles.css";
import { runAgent, AgentEvent } from "../anthropic/agentLoop";
import {
  getActiveProvider,
  loadConfig,
  ProvidersConfig,
  saveConfig,
  clearCache
} from "../config";
import { TOOL_DEFINITIONS, TOOL_HANDLERS } from "../tools/registry";
import { ContentBlock, Message } from "../types";
import { addBubble, setContextBar } from "./ui";
import {
  ChatSession,
  createSession,
  deleteSession,
  deriveTitle,
  extractUserText,
  formatRelativeTime,
  getSession,
  listSessions,
  saveSession
} from "./chatHistory";

let history: Message[] = [];
let currentSession: ChatSession = createSession();
let config: ProvidersConfig;
let modelOverride = "";
let officeReady = false;

async function init() {
  try {
    config = await loadConfig();
  } catch (e) {
    addBubble("error", e instanceof Error ? e.message : String(e));
    return;
  }
  bindUI();

  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const OfficeApi: any = (window as unknown as { Office?: unknown }).Office;
    if (OfficeApi) {
      OfficeApi.onReady(() => {
        officeReady = true;
        setContextBar(`已连接 PowerPoint · Provider: ${config.providers[config.activeProvider]?.name ?? config.activeProvider}`);
        setupSelectionListener();
        startSelectionPolling();
      });
    } else {
      setContextBar("浏览器调试模式（未加载 Office.js）");
    }
  } catch (e) {
    setContextBar("Office.js 未就绪");
  }

  addBubble("assistant", "您好！我是 PowerPoint AI Add-in。请告诉我你想对当前演示文稿做什么。");
}

async function setupSelectionListener() {
  try {
    await PowerPoint.run(async (ctx) => {
      const pres = ctx.presentation;
      const allSlides = pres.slides;
      allSlides.load("items/id");
      const selectedSlides = pres.getSelectedSlides();
      const selectedShapes = pres.getSelectedShapes();
      selectedSlides.load("items/id");
      selectedShapes.load("items/id");
      await ctx.sync();
      const slideIds = allSlides.items.map((s) => s.id);
      const selIds = selectedSlides.items.map((s) => s.id);
      const indexes = selIds.map((id) => slideIds.indexOf(id)).filter((i) => i >= 0);
      updateSelectionStatus(indexes, selectedShapes.items.length);
    });
  } catch {
    // Office API 不可用，静默跳过
  }
}

// 定期刷新选择状态
let selectionInterval: number | null = null;
function startSelectionPolling() {
  if (selectionInterval) return;
  selectionInterval = window.setInterval(async () => {
    try {
      await PowerPoint.run(async (ctx) => {
        const pres = ctx.presentation;
        const allSlides = pres.slides;
        allSlides.load("items/id");
        const selectedSlides = pres.getSelectedSlides();
        const selectedShapes = pres.getSelectedShapes();
        selectedSlides.load("items/id");
        selectedShapes.load("items/id");
        await ctx.sync();
        const slideIds = allSlides.items.map((s) => s.id);
        const selIds = selectedSlides.items.map((s) => s.id);
        const indexes = selIds.map((id) => slideIds.indexOf(id)).filter((i) => i >= 0);
        const shapeCount = selectedShapes.items.length;
        updateSelectionStatus(indexes, shapeCount);
      });
    } catch {
      // ignore polling errors
    }
  }, 1000);
}

function updateSelectionStatus(slideIndexes: number[], shapeCount: number) {
  const el = document.getElementById("ctx-text");
  if (el) {
    const slideText = slideIndexes.length === 0
      ? "未选中幻灯片"
      : slideIndexes.length === 1
        ? `Slide ${slideIndexes[0] + 1} selected`
        : `Slides ${slideIndexes.map((i) => i + 1).join(", ")} selected`;
    const shapeText = shapeCount === 0 ? "无形状" : `${shapeCount} shape${shapeCount > 1 ? "s" : ""}`;
    el.textContent = `${slideText} · ${shapeText}`;
  }
}

function bindUI() {
  document.getElementById("btn-config")!.addEventListener("click", showSettings);
  document.getElementById("btn-back")!.addEventListener("click", showMain);
  document.getElementById("btn-send")!.addEventListener("click", onSend);
  document.getElementById("btn-new-chat")!.addEventListener("click", onNewChat);
  document.getElementById("btn-history")!.addEventListener("click", toggleHistoryPopup);
  document.getElementById("btn-close-history")!.addEventListener("click", () => {
    document.getElementById("history-popup")!.classList.add("hidden");
  });
  // 点击弹窗外部自动关闭
  document.addEventListener("click", (e) => {
    const popup = document.getElementById("history-popup");
    const trigger = document.getElementById("btn-history");
    if (!popup || popup.classList.contains("hidden")) return;
    const target = e.target as Node;
    if (popup.contains(target) || trigger?.contains(target)) return;
    popup.classList.add("hidden");
  });
  document.getElementById("input")!.addEventListener("keydown", (e) => {
    const ke = e as KeyboardEvent;
    if (ke.key === "Enter" && !ke.shiftKey) {
      e.preventDefault();
      onSend();
    }
  });
  document.getElementById("btn-save-config")!.addEventListener("click", onSaveConfig);
  document.getElementById("btn-reset-config")!.addEventListener("click", async () => {
    localStorage.removeItem("claude-for-office.providers");
    clearCache();
    config = await loadConfig();
    renderSettingsView();
    addBubble("assistant", "已重置为配置文件。");
  });
  document.getElementById("provider-select")!.addEventListener("change", (e) => {
    config.activeProvider = (e.target as HTMLSelectElement).value;
    renderSettingsView();
  });
  document.getElementById("model-select")!.addEventListener("change", (e) => {
    const sel = e.target as HTMLSelectElement;
    const customLabel = document.querySelector(".custom-model-label") as HTMLElement;
    if (sel.value === "__custom__") {
      customLabel.classList.remove("hidden");
    } else {
      customLabel.classList.add("hidden");
    }
  });
  document.getElementById("btn-refresh-ctx")!.addEventListener("click", async () => {
    try {
      const out = await TOOL_HANDLERS.get_current_context({}, { log: () => {} });
      addBubble("tool", out, "当前上下文");
    } catch (e) {
      addBubble("error", e instanceof Error ? e.message : String(e));
    }
  });
}

function showSettings() {
  document.getElementById("main-view")!.classList.add("view-hidden");
  document.getElementById("settings-view")!.classList.remove("hidden");
  renderSettingsView();
}

function showMain() {
  document.getElementById("settings-view")!.classList.add("hidden");
  document.getElementById("main-view")!.classList.remove("view-hidden");
}

function renderSettingsView() {
  const sel = document.getElementById("provider-select") as HTMLSelectElement;
  sel.innerHTML = "";
  for (const [key, p] of Object.entries(config.providers)) {
    const opt = document.createElement("option");
    opt.value = key;
    opt.textContent = `${p.name} (${key})`;
    if (key === config.activeProvider) opt.selected = true;
    sel.appendChild(opt);
  }
  const active = getActiveProvider(config);
  (document.getElementById("provider-key") as HTMLInputElement).value = active.env.ANTHROPIC_AUTH_TOKEN ?? "";
  (document.getElementById("provider-url") as HTMLInputElement).value = active.env.ANTHROPIC_BASE_URL ?? "";

  const modelSel = document.getElementById("model-select") as HTMLSelectElement;
  modelSel.innerHTML = "";
  const tiers: { value: string; label: string; model?: string }[] = [
    { value: "", label: "使用 Provider 默认" },
    { value: "opus", label: `Opus（推荐）`, model: active.env.ANTHROPIC_DEFAULT_OPUS_MODEL },
    { value: "sonnet", label: "Sonnet", model: active.env.ANTHROPIC_DEFAULT_SONNET_MODEL },
    { value: "haiku", label: "Haiku", model: active.env.ANTHROPIC_DEFAULT_HAIKU_MODEL },
    { value: "__custom__", label: "自定义..." },
  ];
  for (const tier of tiers) {
    const opt = document.createElement("option");
    opt.value = tier.value;
    opt.textContent = tier.model ? `${tier.label}：${tier.model}` : tier.label;
    modelSel.appendChild(opt);
  }

  const customLabel = document.querySelector(".custom-model-label") as HTMLElement;
  if (modelOverride && !modelSel.querySelector(`option[value="${modelOverride}"]`)) {
    modelSel.value = "__custom__";
    customLabel.classList.remove("hidden");
    (document.getElementById("model-override") as HTMLInputElement).value = modelOverride;
  } else {
    modelSel.value = modelOverride || "";
    customLabel.classList.add("hidden");
  }
}

function onSaveConfig() {
  const active = getActiveProvider(config);
  active.env.ANTHROPIC_AUTH_TOKEN = (document.getElementById("provider-key") as HTMLInputElement).value.trim();
  active.env.ANTHROPIC_BASE_URL = (document.getElementById("provider-url") as HTMLInputElement).value.trim();
  const modelSel = (document.getElementById("model-select") as HTMLSelectElement).value;
  if (modelSel === "__custom__") {
    modelOverride = (document.getElementById("model-override") as HTMLInputElement).value.trim();
  } else {
    modelOverride = modelSel;
  }
  saveConfig(config);
  showMain();
  addBubble("assistant", `配置已保存，当前 Provider: ${active.name}`);
  setContextBar(`Provider: ${active.name}${modelOverride ? ` · 模型: ${modelOverride}` : ""}`);
}

async function onSend() {
  const ta = document.getElementById("input") as HTMLTextAreaElement;
  const text = ta.value.trim();
  if (!text) return;
  ta.value = "";

  let contextInfo = "";
  try {
    contextInfo = await TOOL_HANDLERS.get_current_context({}, { log: () => {} });
  } catch {
    // ignore context fetch errors
  }

  const fullMessage = contextInfo
    ? `[当前上下文]\n${contextInfo}\n\n[用户消息]\n${text}`
    : text;

  addBubble("user", text);
  const sendBtn = document.getElementById("btn-send") as HTMLButtonElement;
  sendBtn.disabled = true;

  let assistantBubble: HTMLElement | null = null;
  let assistantText = "";

  try {
    const updated = await runAgent(fullMessage, history, {
      tools: TOOL_DEFINITIONS,
      handlers: TOOL_HANDLERS,
      modelOverride: modelOverride || undefined,
      onEvent: (ev: AgentEvent) => {
        if (ev.type === "text" && ev.text) {
          assistantText += ev.text;
          if (!assistantBubble) {
            assistantBubble = addBubble("assistant", assistantText);
          } else {
            const body = assistantBubble.querySelector("div:last-child");
            if (body) body.textContent = assistantText;
          }
        } else if (ev.type === "tool_call") {
          assistantBubble = null;
          assistantText = "";
          addBubble("tool", JSON.stringify(ev.toolInput, null, 2), `→ ${ev.toolName}`);
        } else if (ev.type === "tool_result") {
          addBubble(ev.isError ? "error" : "tool", ev.toolResult ?? "", `← ${ev.toolName}`);
        } else if (ev.type === "error" && ev.text) {
          addBubble("error", ev.text);
        }
      }
    });
    history = updated;
    persistCurrentSession();
  } catch (e) {
    addBubble("error", e instanceof Error ? e.message : String(e));
  } finally {
    sendBtn.disabled = false;
  }
}

function persistCurrentSession() {
  currentSession.messages = history;
  currentSession.title = deriveTitle(history);
  saveSession(currentSession);
}

function onNewChat() {
  // 如果当前会话有内容就先保存；空的就不留痕
  if (history.length > 0) {
    persistCurrentSession();
  }
  history = [];
  currentSession = createSession();
  clearMessagesUI();
  addBubble("assistant", "新对话已开始。请告诉我你想对当前演示文稿做什么。");
  document.getElementById("history-popup")!.classList.add("hidden");
}

function clearMessagesUI() {
  const container = document.getElementById("messages");
  if (container) container.innerHTML = "";
}

function toggleHistoryPopup() {
  const popup = document.getElementById("history-popup")!;
  if (popup.classList.contains("hidden")) {
    renderHistoryList();
    popup.classList.remove("hidden");
  } else {
    popup.classList.add("hidden");
  }
}

function renderHistoryList() {
  const list = document.getElementById("history-list")!;
  const empty = document.getElementById("history-empty")!;
  list.innerHTML = "";
  const sessions = listSessions();
  if (sessions.length === 0) {
    empty.classList.remove("hidden");
    list.classList.add("hidden");
    return;
  }
  empty.classList.add("hidden");
  list.classList.remove("hidden");
  const now = Date.now();
  for (const s of sessions) {
    const item = document.createElement("div");
    item.className = "history-item" + (s.id === currentSession.id ? " active" : "");

    const title = document.createElement("span");
    title.className = "history-item-title";
    title.textContent = s.title || "新对话";

    const time = document.createElement("span");
    time.className = "history-item-time";
    time.textContent = formatRelativeTime(s.updatedAt, now);

    const del = document.createElement("button");
    del.className = "history-item-del";
    del.textContent = "✕";
    del.title = "删除";
    del.addEventListener("click", (e) => {
      e.stopPropagation();
      deleteSession(s.id);
      if (s.id === currentSession.id) {
        history = [];
        currentSession = createSession();
        clearMessagesUI();
        addBubble("assistant", "会话已删除。");
      }
      renderHistoryList();
    });

    item.addEventListener("click", () => {
      loadSessionIntoView(s.id);
    });

    item.appendChild(title);
    item.appendChild(time);
    item.appendChild(del);
    list.appendChild(item);
  }
}

function loadSessionIntoView(id: string) {
  // 切换前保存当前会话（如果有内容）
  if (history.length > 0 && id !== currentSession.id) {
    persistCurrentSession();
  }
  const s = getSession(id);
  if (!s) return;
  currentSession = s;
  history = s.messages.slice();
  rebuildMessagesUI(history);
  document.getElementById("history-popup")!.classList.add("hidden");
}

function rebuildMessagesUI(messages: Message[]) {
  clearMessagesUI();
  // 先建一个 tool_use_id → name 的映射，用于 tool_result 回显
  const toolNameById = new Map<string, string>();
  for (const m of messages) {
    if (m.role === "assistant" && Array.isArray(m.content)) {
      for (const b of m.content) {
        if (b.type === "tool_use") toolNameById.set(b.id, b.name);
      }
    }
  }
  for (const m of messages) {
    if (m.role === "user") {
      if (typeof m.content === "string") {
        addBubble("user", extractUserText(m));
      } else {
        for (const b of m.content) {
          if (b.type === "tool_result") {
            const name = toolNameById.get(b.tool_use_id) ?? "tool";
            addBubble(b.is_error ? "error" : "tool", b.content, `← ${name}`);
          }
        }
      }
    } else if (m.role === "assistant") {
      if (typeof m.content === "string") {
        addBubble("assistant", m.content);
      } else {
        for (const b of m.content as ContentBlock[]) {
          if (b.type === "text") {
            if (b.text) addBubble("assistant", b.text);
          } else if (b.type === "tool_use") {
            addBubble("tool", JSON.stringify(b.input, null, 2), `→ ${b.name}`);
          }
        }
      }
    }
  }
}

init();
