import JSZip from "jszip";
import { ToolHandler } from "../types";

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
