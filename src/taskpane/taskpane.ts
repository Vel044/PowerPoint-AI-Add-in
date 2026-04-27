import "./styles.css";
import { runAgent, AgentEvent } from "../anthropic/agentLoop";
import { isRequestCancelled, RequestCancelledError } from "../anthropic/client";
import {
  getActiveProvider,
  loadConfig,
  ModelTier,
  ProvidersConfig,
  resolveModel,
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
let isAgentRunning = false;
let activeRunController: AbortController | null = null;
const MODEL_TIERS = new Set<ModelTier>(["opus", "sonnet", "haiku"]);

function hasPowerPointApi(): boolean {
  return typeof (window as unknown as { PowerPoint?: unknown }).PowerPoint !== "undefined";
}

function requirePowerPointApi(): boolean {
  if (officeReady && hasPowerPointApi()) return true;
  const message = "当前没有检测到 PowerPoint API。请在 PowerPoint 桌面端通过功能区按钮打开此任务窗格；如果是在普通浏览器/localhost 页面里调试，绘图工具无法执行。";
  setContextBar("未连接 PowerPoint");
  addBubble("error", message);
  return false;
}

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
  document.getElementById("btn-send")!.addEventListener("click", onSendButtonClick);
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
      if (!isAgentRunning) onSend();
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
    if (!requirePowerPointApi()) return;
    try {
      const out = await TOOL_HANDLERS.get_current_context({}, { log: () => {} });
      addBubble("tool", out, "当前上下文");
    } catch (e) {
      addBubble("error", e instanceof Error ? e.message : String(e));
    }
  });
}

function onSendButtonClick() {
  if (isAgentRunning) {
    pauseAgent();
    return;
  }
  onSend();
}

function pauseAgent() {
  if (!activeRunController || activeRunController.signal.aborted) return;
  activeRunController.abort();
  const sendBtn = document.getElementById("btn-send") as HTMLButtonElement;
  sendBtn.textContent = "暂停中...";
  sendBtn.disabled = true;
}

function setAgentRunning(running: boolean) {
  isAgentRunning = running;
  const sendBtn = document.getElementById("btn-send") as HTMLButtonElement | null;
  if (!sendBtn) return;
  sendBtn.disabled = false;
  sendBtn.textContent = running ? "暂停" : "发送";
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
  setContextBar(`Provider: ${active.name}${modelOverride ? ` · 模型: ${describeSelectedModel(active)}` : ""}`);
}

async function onSend() {
  if (isAgentRunning) return;
  const ta = document.getElementById("input") as HTMLTextAreaElement;
  const text = ta.value.trim();
  if (!text) return;
  if (!requirePowerPointApi()) return;
  ta.value = "";
  const runController = new AbortController();
  activeRunController = runController;
  setAgentRunning(true);
  addBubble("user", text);

  let contextInfo = "";
  let slidesInfo = "";
  try {
    contextInfo = await TOOL_HANDLERS.get_current_context({}, { log: () => {} });
    slidesInfo = await TOOL_HANDLERS.list_slides({}, { log: () => {} });
    if (runController.signal.aborted) throw new RequestCancelledError();
  } catch {
    if (runController.signal.aborted) {
      addBubble("assistant", "已暂停。");
      cleanupAgentRun(runController);
      return;
    }
    // ignore context fetch errors
  }

  let slideScreenshot: string | null = null;
  try {
    slideScreenshot = await PowerPoint.run(async (ctx) => {
      const slide = ctx.presentation.getSelectedSlides().getItemAt(0);
      const img = (slide as any).getImageAsBase64({ width: 640 });
      await ctx.sync();
      return typeof img.value === "string" ? img.value : null;
    });
  } catch {
    // API 不支持时静默忽略
  }

  const userContent = buildUserContent(contextInfo, slidesInfo, text, slideScreenshot);

  let assistantBubble: HTMLElement | null = null;
  let assistantText = "";

  try {
    const modelRequest = getModelRequest();
    const updated = await runAgent(userContent, history, {
      tools: TOOL_DEFINITIONS,
      handlers: TOOL_HANDLERS,
      tier: modelRequest.tier,
      modelOverride: modelRequest.modelOverride,
      signal: runController.signal,
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
        } else if (ev.type === "cancelled") {
          addBubble("assistant", ev.text ?? "已暂停。");
        }
      }
    });
    history = updated;
    persistCurrentSession();
  } catch (e) {
    if (!isRequestCancelled(e)) {
      addBubble("error", e instanceof Error ? e.message : String(e));
    }
  } finally {
    cleanupAgentRun(runController);
  }
}

function cleanupAgentRun(controller: AbortController) {
  if (activeRunController === controller) {
    activeRunController = null;
    setAgentRunning(false);
  }
}

function buildUserContent(
  contextInfo: string,
  slidesInfo: string,
  userText: string,
  slideScreenshot: string | null,
): string | ContentBlock[] {
  const context = parseJsonObject(contextInfo);
  const slides = parseJsonArray(slidesInfo);
  const textBlock = context
    ? `${formatInitialState(context, slides)}\n\n${formatUserContext(context)}\n\n<user_message>\n${userText}\n</user_message>`
    : `<user_message>\n${userText}\n</user_message>`;

  if (!slideScreenshot) return textBlock;
  return [
    { type: "text", text: textBlock },
    { type: "text", text: "以下是当前幻灯片的截图，请参考其布局、背景和主题色规划新内容：" },
    { type: "image", source: { type: "base64", media_type: "image/png", data: slideScreenshot } }
  ];
}

function formatInitialState(context: Record<string, unknown>, slides: unknown[]): string {
  const slidesMetadata = slides
    .filter((item): item is Record<string, unknown> => Boolean(item) && typeof item === "object")
    .map((slide) => ({
      slideId: slide.id,
      position: slide.pageNumber,
      index: slide.index,
    }));
  return `<initial_state>\n${JSON.stringify({
    slidesMetadata,
    totalSlides: context.totalSlides,
    slideWidth: context.slideWidth,
    slideHeight: context.slideHeight,
    themePalette: context.themePalette,
    isDefaultTheme: context.isDefaultTheme,
    themeName: context.themeName,
  }, null, 2)}\n</initial_state>`;
}

function formatUserContext(context: Record<string, unknown>): string {
  return `<user_context>\n${JSON.stringify({
    selectedSlideIds: context.selectedSlideIds,
    selectedSlideIndexes: context.selectedSlideIndexes,
    selectedPageNumbers: context.selectedPageNumbers,
    currentSlideId: context.currentSlideId,
    currentSlideIndex: context.currentSlideIndex,
    currentPageNumber: context.currentPageNumber,
    inspectedSlideId: context.inspectedSlideId,
    inspectedSlideIndex: context.inspectedSlideIndex,
    inspectedPageNumber: context.inspectedPageNumber,
    occupiedBounds: context.occupiedBounds,
    allShapes: context.allShapes,
    selectedShapes: context.selectedShapes,
  }, null, 2)}\n</user_context>`;
}

function parseJsonObject(value: string): Record<string, unknown> | null {
  try {
    const parsed = value ? JSON.parse(value) : null;
    return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed as Record<string, unknown> : null;
  } catch {
    return null;
  }
}

function parseJsonArray(value: string): unknown[] {
  try {
    const parsed = value ? JSON.parse(value) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function getModelRequest(): { tier?: ModelTier; modelOverride?: string } {
  if (MODEL_TIERS.has(modelOverride as ModelTier)) {
    return { tier: modelOverride as ModelTier };
  }
  return { modelOverride: modelOverride || undefined };
}

function describeSelectedModel(active = getActiveProvider(config)): string {
  if (MODEL_TIERS.has(modelOverride as ModelTier)) {
    const tier = modelOverride as ModelTier;
    return `${tier} / ${resolveModel(active, tier)}`;
  }
  return modelOverride;
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
