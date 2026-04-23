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
    history.length = 0;
    history.push(...updated);
  } catch (e) {
    addBubble("error", e instanceof Error ? e.message : String(e));
  } finally {
    sendBtn.disabled = false;
  }
}

init();
