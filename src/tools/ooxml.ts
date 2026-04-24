import JSZip from "jszip";
import { ToolHandler } from "../types";
import { ARROW_NAME_PREFIX } from "./layout";

let cachedZip: { zip: JSZip; loadedAt: number } | null = null;
const CACHE_MS = 30000;

async function getDocumentBytes(): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(
      Office.FileType.Compressed,
      { sliceSize: 65536 },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(result.error?.message ?? "getFileAsync 失败"));
          return;
        }
        const file = result.value;
        const sliceCount = file.sliceCount;
        const chunks: Uint8Array[] = [];
        let received = 0;
        const getSlice = (i: number) => {
          file.getSliceAsync(i, (r) => {
            if (r.status !== Office.AsyncResultStatus.Succeeded) {
              file.closeAsync();
              reject(new Error(r.error?.message ?? "getSliceAsync 失败"));
              return;
            }
            const data = r.value.data as unknown;
            const bytes = data instanceof Uint8Array ? data : new Uint8Array(data as ArrayBuffer | number[]);
            chunks[i] = bytes;
            received++;
            if (received === sliceCount) {
              file.closeAsync();
              const total = chunks.reduce((n, c) => n + c.length, 0);
              const out = new Uint8Array(total);
              let off = 0;
              for (const c of chunks) { out.set(c, off); off += c.length; }
              resolve(out);
            } else {
              getSlice(i + 1 < sliceCount ? i + 1 : i);
            }
          });
        };
        for (let i = 0; i < sliceCount; i++) getSlice(i);
      }
    );
  });
}

async function getZip(): Promise<JSZip> {
  const now = Date.now();
  if (cachedZip && now - cachedZip.loadedAt < CACHE_MS) return cachedZip.zip;
  const bytes = await getDocumentBytes();
  const zip = await JSZip.loadAsync(bytes);
  cachedZip = { zip, loadedAt: now };
  return zip;
}

export const exportPptxXml: ToolHandler = async (input) => {
  const wanted = (input.path as string) || "ppt/slides/slide1.xml";
  const maxChars = (input.maxChars as number) ?? 4000;
  const zip = await getZip();
  if (input.list === true) {
    const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir);
    return JSON.stringify({ files: names.slice(0, 200) });
  }
  const entry = zip.file(wanted);
  if (!entry) {
    const candidates = Object.keys(zip.files).filter((n) => n.includes(wanted));
    return `未找到 ${wanted}。候选: ${candidates.slice(0, 20).join(", ")}`;
  }
  const xml = await entry.async("string");
  if (xml.length > maxChars) {
    return xml.slice(0, maxChars) + `\n...[已截断，原长 ${xml.length}]`;
  }
  return xml;
};

export const applyPptxPatch: ToolHandler = async (input) => {
  const patches = input.patches as Array<{ path: string; content: string }> | undefined;
  if (!patches || !Array.isArray(patches) || patches.length === 0) {
    throw new Error("patches 必须是 {path, content}[] 非空数组");
  }
  const zip = await getZip();
  for (const p of patches) {
    if (!p.path || typeof p.content !== "string") throw new Error("patch 缺少 path 或 content");
    zip.file(p.path, p.content);
  }
  const blob = await zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `claude-patched-${Date.now()}.pptx`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
  cachedZip = null;
  return `已生成修补后的 .pptx，共修改 ${patches.length} 个文件，已触发浏览器下载。请用户打开下载的新文件。`;
};

const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export const finalizeArrows: ToolHandler = async (_input) => {
  cachedZip = null;
  const zip = await getZip();

  const slideFiles = Object.keys(zip.files)
    .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n));
  if (slideFiles.length === 0) return "未找到任何 slide 文件";

  const parser = new DOMParser();
  const serializer = new XMLSerializer();

  let patchCount = 0;
  const patches: Array<{ path: string; content: string }> = [];

  for (const path of slideFiles) {
    const xml = await zip.file(path)!.async("string");
    const doc = parser.parseFromString(xml, "application/xml");
    const cxns = doc.getElementsByTagNameNS(P_NS, "cxnSp");
    let slideChanged = false;

    for (let i = 0; i < cxns.length; i++) {
      const cxn = cxns[i];
      // cNvPr sits inside nvCxnSpPr → cNvPr
      const cNvPrs = cxn.getElementsByTagNameNS(P_NS, "cNvPr");
      const cNvPr = cNvPrs.length > 0 ? cNvPrs[0] : null;
      const name = cNvPr?.getAttribute("name") ?? "";
      if (!name.startsWith(ARROW_NAME_PREFIX + "_")) continue;

      const mode = name.endsWith("_BOTH") ? "both"
                 : name.endsWith("_END")  ? "end"
                 : null;
      if (!mode) continue;

      const spPrs = cxn.getElementsByTagNameNS(P_NS, "spPr");
      const spPr = spPrs.length > 0 ? spPrs[0] : null;
      if (!spPr) continue;

      let ln = spPr.getElementsByTagNameNS(A_NS, "ln")[0];
      if (!ln) {
        ln = doc.createElementNS(A_NS, "a:ln");
        spPr.appendChild(ln);
      }

      // Idempotent: remove existing headEnd/tailEnd
      const existingHead = ln.getElementsByTagNameNS(A_NS, "headEnd");
      for (let h = existingHead.length - 1; h >= 0; h--) ln.removeChild(existingHead[h]);
      const existingTail = ln.getElementsByTagNameNS(A_NS, "tailEnd");
      for (let h = existingTail.length - 1; h >= 0; h--) ln.removeChild(existingTail[h]);

      if (mode === "both") {
        const headEnd = doc.createElementNS(A_NS, "a:headEnd");
        headEnd.setAttribute("type", "triangle");
        headEnd.setAttribute("w", "med");
        headEnd.setAttribute("len", "med");
        ln.appendChild(headEnd);
      }
      const tailEnd = doc.createElementNS(A_NS, "a:tailEnd");
      tailEnd.setAttribute("type", "triangle");
      tailEnd.setAttribute("w", "med");
      tailEnd.setAttribute("len", "med");
      ln.appendChild(tailEnd);

      slideChanged = true;
      patchCount++;
    }

    if (slideChanged) {
      patches.push({ path, content: serializer.serializeToString(doc) });
    }
  }

  if (patches.length === 0) {
    return `未发现带 ${ARROW_NAME_PREFIX}_* 标记的连接线，无需后处理。`;
  }

  for (const p of patches) {
    zip.file(p.path, p.content);
  }
  const blob = await zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `claude-arrows-${Date.now()}.pptx`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
  cachedZip = null;
  return `已注入 ${patchCount} 条真箭头，共修改 ${patches.length} 张幻灯片。新 .pptx 已触发下载，请打开新文件查看效果。`;
};
