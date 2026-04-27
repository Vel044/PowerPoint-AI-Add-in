import { ToolHandler } from "../types";
import { ConnectorXmlPatch, patchConnectorXml } from "./layout";
import { parseShapeRef, targetFromRefOrInput } from "./refs";
import { editCurrentSlideXml, readCurrentSlideXml, serializeXmlElement } from "./ooxml";

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
  const operations = Array.isArray(input.operations) ? input.operations as Array<Record<string, unknown>> : [];
  if (operations.length === 0) throw new Error("缺少 operations");
  const result = await editCurrentSlideXml({
    slideId: typeof input.slideId === "string" ? input.slideId : undefined,
    slideIndex: typeof input.slideIndex === "number" ? input.slideIndex : undefined,
    pageNumber: typeof input.pageNumber === "number" ? input.pageNumber : undefined,
  }, (doc) => {
    for (const operation of operations) applyXmlOperation(doc, operation);
  });
  return `已执行 ${operations.length} 个结构化 XML 操作；oldSlideId=${result.oldSlideId}；newSlideId=${result.newSlideId}`;
};

function applyXmlOperation(doc: Document, operation: Record<string, unknown>): void {
  const type = String(operation.type ?? "");
  if (type === "insertShapeXml") {
    const xml = requireXml(operation.xml);
    const spTree = doc.getElementsByTagNameNS(P_NS, "spTree")[0];
    if (!spTree) throw new Error("slide XML 中未找到 spTree");
    spTree.appendChild(doc.importNode(parseSingleElement(xml), true));
    return;
  }
  if (type === "replaceShapeXml") {
    const shapeId = requireString(operation.shapeId, "shapeId");
    const shape = findShape(doc, shapeId);
    if (!shape) throw new Error(`slide XML 中未找到形状 ${shapeId}`);
    shape.parentNode?.replaceChild(doc.importNode(parseSingleElement(requireXml(operation.xml)), true), shape);
    return;
  }
  if (type === "deleteShapeXml") {
    const shapeId = requireString(operation.shapeId, "shapeId");
    const shape = findShape(doc, shapeId);
    if (!shape) throw new Error(`slide XML 中未找到形状 ${shapeId}`);
    shape.parentNode?.removeChild(shape);
    return;
  }
  if (type === "patchConnector") {
    patchConnectorXml(doc, connectorPatchFromOperation(operation));
    return;
  }
  if (type === "setSlideBackground") {
    setSlideBackground(doc, requireString(operation.color, "color"));
    return;
  }
  throw new Error(`不支持的 edit_slide_xml operation.type=${type || "empty"}`);
}

function connectorPatchFromOperation(operation: Record<string, unknown>): ConnectorXmlPatch {
  return {
    connectorShapeId: requireString(operation.connectorShapeId, "connectorShapeId"),
    fromShapeId: requireString(operation.fromShapeId, "fromShapeId"),
    fromSide: sideValue(operation.fromSide, "fromSide"),
    toShapeId: requireString(operation.toShapeId, "toShapeId"),
    toSide: sideValue(operation.toSide, "toSide"),
    start: pointValue(operation.start, "start"),
    end: pointValue(operation.end, "end"),
    connectorType: String(operation.connectorType ?? "elbow") === "straight" ? "straight" : "elbow",
    arrow: String(operation.arrow ?? "end") === "none" ? "none" : "end",
    color: typeof operation.color === "string" ? operation.color : "#2F5597",
    thickness: typeof operation.thickness === "number" ? operation.thickness : 2,
    dashStyle: typeof operation.dashStyle === "string" ? operation.dashStyle : undefined,
  };
}

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

function parseSingleElement(xml: string): Element {
  const doc = new DOMParser().parseFromString(
    `<root xmlns:a="${A_NS}" xmlns:p="${P_NS}">${xml}</root>`,
    "application/xml"
  );
  const error = doc.getElementsByTagName("parsererror")[0];
  if (error) throw new Error(`XML 解析失败: ${error.textContent?.trim() ?? "unknown parser error"}`);
  const element = Array.from(doc.documentElement.childNodes).find((node) => node.nodeType === Node.ELEMENT_NODE) as Element | undefined;
  if (!element) throw new Error("XML 操作必须包含一个元素");
  return element;
}

function setSlideBackground(doc: Document, color: string): void {
  const cSld = doc.getElementsByTagNameNS(P_NS, "cSld")[0];
  if (!cSld) throw new Error("slide XML 中未找到 cSld");
  const oldBg = firstChildByNs(cSld, P_NS, "bg");
  if (oldBg) cSld.removeChild(oldBg);
  const bg = doc.createElementNS(P_NS, "p:bg");
  const bgPr = doc.createElementNS(P_NS, "p:bgPr");
  const solidFill = doc.createElementNS(A_NS, "a:solidFill");
  const srgbClr = doc.createElementNS(A_NS, "a:srgbClr");
  srgbClr.setAttribute("val", color.replace(/^#/, "").toUpperCase());
  solidFill.appendChild(srgbClr);
  bgPr.appendChild(solidFill);
  bg.appendChild(bgPr);
  cSld.insertBefore(bg, cSld.firstChild);
}

function requireXml(value: unknown): string {
  const xml = typeof value === "string" ? value.trim() : "";
  if (!xml) throw new Error("XML 操作缺少 xml");
  return xml;
}

function requireString(value: unknown, name: string): string {
  if (typeof value !== "string" || !value.trim()) throw new Error(`缺少 ${name}`);
  return value.trim();
}

function sideValue(value: unknown, name: string): "top" | "bottom" | "left" | "right" {
  if (value === "top" || value === "bottom" || value === "left" || value === "right") return value;
  throw new Error(`${name} 必须是 top/bottom/left/right`);
}

function pointValue(value: unknown, name: string): { x: number; y: number } {
  const point = value as { x?: unknown; y?: unknown };
  if (typeof point?.x !== "number" || typeof point?.y !== "number") {
    throw new Error(`${name} 必须包含 number 类型的 x/y`);
  }
  return { x: point.x, y: point.y };
}
