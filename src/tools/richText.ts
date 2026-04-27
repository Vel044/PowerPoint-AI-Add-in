import { ToolHandler } from "../types";
import { targetFromRefOrInput } from "./refs";
import { editCurrentSlideXml, editCurrentSlideZip, readCurrentSlideXml, serializeXmlElement } from "./ooxml";

const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

export const readSlideText: ToolHandler = async (input) => {
  const { target, shapeId } = targetFromRefOrInput(input);
  if (!shapeId) throw new Error("缺少 ref 或 shapeId");
  return await readCurrentSlideXml(target, (doc, info) => {
    const shape = findShape(doc, shapeId);
    if (!shape) throw new Error(`slide XML 中未找到形状 ${shapeId}`);
    const txBody = firstChildByNs(shape, P_NS, "txBody");
    if (!txBody) throw new Error(`形状 ${shapeId} 没有文本体 txBody`);
    const paragraphs = Array.from(txBody.getElementsByTagNameNS(A_NS, "p"))
      .map((p) => serializeXmlElement(p));
    return JSON.stringify({
      ref: `${info.slideId}:${shapeId}`,
      slideId: info.slideId,
      slideIndex: info.slideIndex,
      pageNumber: info.pageNumber,
      shapeId,
      paragraphXml: paragraphs.join("\n"),
      paragraphs,
    }, null, 2);
  });
};

export const editSlideText: ToolHandler = async (input) => {
  const { target, shapeId } = targetFromRefOrInput(input);
  const paragraphXml = typeof input.paragraphXml === "string"
    ? input.paragraphXml
    : typeof input.code === "string"
      ? input.code
      : "";
  if (!shapeId) throw new Error("缺少 ref 或 shapeId");
  if (!paragraphXml.trim()) throw new Error("缺少 paragraphXml");
  const result = await editCurrentSlideXml(target, (doc) => {
    const shape = findShape(doc, shapeId);
    if (!shape) throw new Error(`slide XML 中未找到形状 ${shapeId}`);
    const txBody = ensureTxBody(doc, shape);
    const paragraphs = parseParagraphs(paragraphXml);
    const oldParagraphs = Array.from(txBody.getElementsByTagNameNS(A_NS, "p"));
    for (const p of oldParagraphs) p.parentNode?.removeChild(p);
    for (const p of paragraphs) txBody.appendChild(doc.importNode(p, true));
  });
  return `已替换形状 ${shapeId} 的富文本段落；新 slideId=${result.newSlideId}`;
};

export const editSlideXml: ToolHandler = async (input) => {
  const code = typeof input.code === "string" ? input.code : "";
  if (!code.trim()) throw new Error("缺少 code。edit_slide_xml 现在只支持 Claude 风格 code:string，不再支持 operations[]。");

  const autosizeShapeIds = stringArray(input.autosize_shape_ids ?? input.autosizeShapeIds);
  const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor as new (...args: string[]) => (env: {
    zip: unknown;
    markDirty: () => void;
    pptx: unknown;
  }) => Promise<unknown>;
  const fn = new AsyncFunction("{ zip, markDirty, pptx }", code);
  const result = await editCurrentSlideZip({
    slideId: typeof input.slideId === "string" ? input.slideId : undefined,
    slideIndex: typeof input.slideIndex === "number" ? input.slideIndex : undefined,
    pageNumber: typeof input.pageNumber === "number" ? input.pageNumber : undefined,
  }, async ({ zip, markDirty, pptx }) => {
    await fn({ zip, markDirty, pptx });
  }, {
    autosizeShapeIds,
    artifactMetadata: {
      tool: "edit_slide_xml",
      explanation: typeof input.explanation === "string" ? input.explanation : "",
      timestamp: new Date().toISOString(),
    },
  });
  const paths = [
    result.artifacts.beforeXmlPath ? `before=${result.artifacts.beforeXmlPath}` : "",
    result.artifacts.afterXmlPath ? `after=${result.artifacts.afterXmlPath}` : "",
  ].filter(Boolean).join("；");
  if (!result.dirty) {
    return `code 已执行，但没有调用 markDirty()，未写回 PowerPoint。oldSlideId=${result.oldSlideId}${paths ? `；artifact: ${paths}` : ""}`;
  }
  return `已执行 Claude 风格 edit_slide_xml code，并通过单页 export/import 写回当前 PowerPoint；oldSlideId=${result.oldSlideId}；newSlideId=${result.newSlideId}；newSlideIndex=${result.newSlideIndex}；artifact: ${paths || "未保存"}`;
};

function findShape(doc: Document, shapeId: string): Element | null {
  const candidates = [
    ...Array.from(doc.getElementsByTagNameNS(P_NS, "sp")),
    ...Array.from(doc.getElementsByTagNameNS(P_NS, "cxnSp")),
    ...Array.from(doc.getElementsByTagNameNS(P_NS, "pic")),
    ...Array.from(doc.getElementsByTagNameNS(P_NS, "graphicFrame")),
  ];
  return candidates.find((shape) => {
    const cNvPr = shape.getElementsByTagNameNS(P_NS, "cNvPr")[0];
    return cNvPr?.getAttribute("id") === shapeId;
  }) ?? null;
}

function ensureTxBody(doc: Document, shape: Element): Element {
  const existing = firstChildByNs(shape, P_NS, "txBody");
  if (existing) return existing;
  const txBody = doc.createElementNS(P_NS, "p:txBody");
  txBody.appendChild(doc.createElementNS(A_NS, "a:bodyPr"));
  txBody.appendChild(doc.createElementNS(A_NS, "a:lstStyle"));
  shape.appendChild(txBody);
  return txBody;
}

function firstChildByNs(parent: Element, namespaceUri: string, localName: string): Element | null {
  return Array.from(parent.childNodes).find((node) =>
    node.nodeType === Node.ELEMENT_NODE &&
    (node as Element).namespaceURI === namespaceUri &&
    (node as Element).localName === localName
  ) as Element | null;
}

function parseParagraphs(xml: string): Element[] {
  const doc = new DOMParser().parseFromString(
    `<root xmlns:a="${A_NS}" xmlns:p="${P_NS}">${xml}</root>`,
    "application/xml"
  );
  const error = doc.getElementsByTagName("parsererror")[0];
  if (error) throw new Error(`paragraphXml 解析失败: ${error.textContent?.trim() ?? "unknown parser error"}`);
  const paragraphs = Array.from(doc.documentElement.childNodes).filter((node) =>
    node.nodeType === Node.ELEMENT_NODE &&
    (node as Element).localName === "p"
  ) as Element[];
  if (paragraphs.length === 0) throw new Error("paragraphXml 必须包含至少一个 <a:p>");
  return paragraphs;
}

function stringArray(value: unknown): string[] {
  if (!Array.isArray(value)) return [];
  return value.filter((item): item is string => typeof item === "string" && item.trim().length > 0);
}
