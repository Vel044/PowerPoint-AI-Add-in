import "./styles.css";
import { runAgentStream, AgentEvent } from "../anthropic/agentLoop";
import {
  getActiveProvider,
  loadConfig,
  ProvidersConfig,
  saveConfig,
  clearCache
} from "../config";
import { TOOL_DEFINITIONS, TOOL_HANDLERS } from "../tools/registry";
import { Message } from "../types";
import { addBubble, setContextBar } from "./ui";

const history: Message[] = [];
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
  renderConfigPanel();

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
      const selectedSlides = pres.getSelectedSlides();
      const selectedShapes = pres.getSelectedShapes();
      selectedSlides.load("items/id");
      selectedShapes.load("items/id");
      await ctx.sync();
      updateSelectionStatus(selectedSlides.items.length, selectedShapes.items.length);
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
        const selectedSlides = pres.getSelectedSlides();
        const selectedShapes = pres.getSelectedShapes();
        selectedSlides.load("items/id");
        selectedShapes.load("items/id");
        await ctx.sync();
        updateSelectionStatus(selectedSlides.items.length, selectedShapes.items.length);
      });
    } catch {
      // ignore polling errors
    }
  }, 1000);
}

function updateSelectionStatus(slideCount: number, shapeCount: number) {
  const el = document.getElementById("ctx-text");
  if (el) {
    const slideText = slideCount === 0 ? "未选中幻灯片" : `Slide ${slideCount} selected`;
    const shapeText = shapeCount === 0 ? "无形状" : `${shapeCount} shape${shapeCount > 1 ? "s" : ""}`;
    el.textContent = `${slideText} · ${shapeText}`;
  }
}

function bindUI() {
  document.getElementById("btn-config")!.addEventListener("click", () => {
    document.getElementById("config-panel")!.classList.toggle("hidden");
  });
  document.getElementById("btn-send")!.addEventListener("click", onSend);
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
    renderConfigPanel();
    addBubble("assistant", "已重置为配置文件。");
  });
  document.getElementById("provider-select")!.addEventListener("change", (e) => {
    config.activeProvider = (e.target as HTMLSelectElement).value;
    renderConfigPanel();
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

function renderConfigPanel() {
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
  (document.getElementById("model-override") as HTMLInputElement).value = modelOverride;
}

function onSaveConfig() {
  const active = getActiveProvider(config);
  active.env.ANTHROPIC_AUTH_TOKEN = (document.getElementById("provider-key") as HTMLInputElement).value.trim();
  active.env.ANTHROPIC_BASE_URL = (document.getElementById("provider-url") as HTMLInputElement).value.trim();
  modelOverride = (document.getElementById("model-override") as HTMLInputElement).value.trim();
  saveConfig(config);
  addBubble("assistant", `配置已保存，当前 Provider: ${active.name}`);
  setContextBar(`Provider: ${active.name}${modelOverride ? ` · 模型: ${modelOverride}` : ""}`);
  document.getElementById("config-panel")!.classList.add("hidden");
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
    const updated = await runAgentStream(fullMessage, history, {
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
    history.length = 0;
    history.push(...updated);
  } catch (e) {
    addBubble("error", e instanceof Error ? e.message : String(e));
  } finally {
    sendBtn.disabled = false;
  }
}

init();
