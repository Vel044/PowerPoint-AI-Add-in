export type BubbleKind = "user" | "assistant" | "tool" | "error";

export function addBubble(kind: BubbleKind, text: string, meta?: string): HTMLElement {
  const container = document.getElementById("messages")!;
  const el = document.createElement("div");
  el.className = `msg ${kind}`;
  if (meta) {
    const m = document.createElement("div");
    m.className = "meta";
    m.textContent = meta;
    el.appendChild(m);
  }
  const body = document.createElement("div");
  body.textContent = text;
  el.appendChild(body);

  // Tool messages are collapsible by default
  if (kind === "tool" || kind === "error") {
    el.classList.add("collapsible");
    const preview = document.createElement("div");
    preview.className = "preview";
    preview.textContent = text.length > 100 ? text.slice(0, 100) + "..." : text;
    el.insertBefore(preview, body);
    body.style.display = "none";

    el.addEventListener("click", () => {
      el.classList.toggle("expanded");
      if (el.classList.contains("expanded")) {
        preview.style.display = "none";
        body.style.display = "block";
      } else {
        preview.style.display = "block";
        body.style.display = "none";
      }
    });
  }

  container.appendChild(el);
  container.scrollTop = container.scrollHeight;
  return el;
}

export function setContextBar(text: string): void {
  const el = document.getElementById("ctx-text");
  if (el) el.textContent = text;
}
