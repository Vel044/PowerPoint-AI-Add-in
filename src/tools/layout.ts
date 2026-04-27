import { editCurrentSlideXml, SlideXmlEditResult } from "./ooxml";

export type Side = "top" | "bottom" | "left" | "right";

export interface Rect { left: number; top: number; width: number; height: number; }

const CONNECTOR_COLOR = "#2F5597";
const CONNECTOR_THICKNESS = 2;
const CONNECTOR_EPSILON = 0.01;
const POINT_TO_EMU = 12700;
const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

interface Point {
  x: number;
  y: number;
}

export interface ConnectorDrawResult {
  lineSegments: number;
  straightConnectors: number;
  elbowConnectors: number;
  arrows: number;
  shapeCount: number;
  arrowHeadsEnabled: boolean;
}

interface ConnectorOptions {
  arrow?: "none" | "end";
  color?: string;
  thickness?: number;
  dashStyle?: string;
  mode?: "direct" | "orthogonal";
}

export interface ConnectorXmlPatch {
  connectorShapeId: string;
  fromShapeId: string;
  fromSide: Side;
  toShapeId: string;
  toSide: Side;
  start: Point;
  end: Point;
  connectorType: "straight" | "elbow";
  arrow: "none" | "end";
  color: string;
  thickness: number;
  dashStyle?: string;
}

interface PendingConnectorXmlPatch extends Omit<ConnectorXmlPatch, "connectorShapeId"> {
  connectorShape: PowerPoint.Shape;
}

export interface ConnectorCreationResult extends ConnectorDrawResult {
  patches: PendingConnectorXmlPatch[];
}

function addConnectorPlaceholder(
  slide: PowerPoint.Slide,
  start: Point,
  end: Point,
  connectorType: "straight" | "elbow",
  color: string,
  thickness: number,
  dashStyle?: string,
): PowerPoint.Shape | null {
  const dx = end.x - start.x;
  const dy = end.y - start.y;
  const length = Math.hypot(dx, dy);
  if (length < 0.5) return null;

  const isElbow = connectorType === "elbow";
  const left = isElbow ? Math.min(start.x, end.x) : start.x;
  const top = isElbow ? Math.min(start.y, end.y) : start.y;
  const width = isElbow ? Math.max(Math.abs(dx), CONNECTOR_EPSILON) : (Math.abs(dx) < 0.5 ? CONNECTOR_EPSILON : dx);
  const height = isElbow ? Math.max(Math.abs(dy), CONNECTOR_EPSILON) : (Math.abs(dy) < 0.5 ? CONNECTOR_EPSILON : dy);

  const shape = slide.shapes.addLine(isElbow ? PowerPoint.ConnectorType.elbow : PowerPoint.ConnectorType.straight, {
    left,
    top,
    width,
    height,
  });
  shape.lineFormat.visible = true;
  shape.lineFormat.color = color;
  shape.lineFormat.weight = thickness;
  shape.lineFormat.dashStyle = normalizeDashStyle(dashStyle ?? "solid") as any;
  shape.lineFormat.style = "Single";
  shape.load("id");
  return shape;
}

function sideIndex(side: Side): number {
  switch (side) {
    case "top": return 0;
    case "right": return 1;
    case "bottom": return 2;
    case "left": return 3;
  }
}

export function drawConnectedLine(
  slide: PowerPoint.Slide,
  fromShapeId: string,
  fromRect: Rect,
  fromSide: Side,
  toShapeId: string,
  toRect: Rect,
  toSide: Side,
  options: ConnectorOptions = {},
): ConnectorCreationResult {
  const start = sidePoint(fromRect, fromSide);
  const end = sidePoint(toRect, toSide);
  const isAxisAligned = Math.abs(start.x - end.x) < 0.5 || Math.abs(start.y - end.y) < 0.5;
  const connectorType: "straight" | "elbow" = options.mode === "direct" || isAxisAligned ? "straight" : "elbow";
  const color = options.color ?? CONNECTOR_COLOR;
  const thickness = options.thickness ?? CONNECTOR_THICKNESS;
  const dashStyle = options.dashStyle;
  const arrow = options.arrow ?? "end";
  const shape = addConnectorPlaceholder(slide, start, end, connectorType, color, thickness, dashStyle);

  if (!shape) {
    return {
      lineSegments: 0,
      straightConnectors: 0,
      elbowConnectors: 0,
      arrows: 0,
      shapeCount: 0,
      arrowHeadsEnabled: true,
      patches: [],
    };
  }

  return {
    lineSegments: 1,
    straightConnectors: connectorType === "straight" ? 1 : 0,
    elbowConnectors: connectorType === "elbow" ? 1 : 0,
    arrows: arrow === "end" ? 1 : 0,
    shapeCount: 1,
    arrowHeadsEnabled: true,
    patches: [{
      connectorShape: shape,
      fromShapeId,
      fromSide,
      toShapeId,
      toSide,
      start,
      end,
      connectorType,
      arrow,
      color,
      thickness,
      dashStyle,
    }],
  };
}

export function resolveConnectorXmlPatches(results: ConnectorCreationResult[]): ConnectorXmlPatch[] {
  return results.flatMap((result) => result.patches.map((patch) => ({
    connectorShapeId: patch.connectorShape.id,
    fromShapeId: patch.fromShapeId,
    fromSide: patch.fromSide,
    toShapeId: patch.toShapeId,
    toSide: patch.toSide,
    start: patch.start,
    end: patch.end,
    connectorType: patch.connectorType,
    arrow: patch.arrow,
    color: patch.color,
    thickness: patch.thickness,
    dashStyle: patch.dashStyle,
  })));
}

export async function applyConnectorXmlPatches(
  target: { slideId?: string; slideIndex?: number; pageNumber?: number },
  patches: ConnectorXmlPatch[],
): Promise<SlideXmlEditResult | null> {
  if (patches.length === 0) return null;
  return await editCurrentSlideXml(target, (doc) => {
    for (const patch of patches) patchConnectorXml(doc, patch);
  });
}

export function patchConnectorXml(doc: Document, patch: ConnectorXmlPatch): void {
  const connector = findConnector(doc, patch.connectorShapeId);
  if (!connector) throw new Error(`slide XML 中未找到连接器 ${patch.connectorShapeId}`);

  const nvCxnSpPr = ensureChild(doc, connector, P_NS, "nvCxnSpPr");
  const cNvCxnSpPr = ensureChild(doc, nvCxnSpPr, P_NS, "cNvCxnSpPr");
  removeAllChildren(cNvCxnSpPr);
  const stCxn = doc.createElementNS(A_NS, "a:stCxn");
  stCxn.setAttribute("id", patch.fromShapeId);
  stCxn.setAttribute("idx", String(sideIndex(patch.fromSide)));
  const endCxn = doc.createElementNS(A_NS, "a:endCxn");
  endCxn.setAttribute("id", patch.toShapeId);
  endCxn.setAttribute("idx", String(sideIndex(patch.toSide)));
  cNvCxnSpPr.append(stCxn, endCxn);

  const spPr = ensureChild(doc, connector, P_NS, "spPr");
  setTransform(doc, spPr, patch.start, patch.end);
  setPresetGeometry(doc, spPr, patch.connectorType);
  setLine(doc, spPr, patch);
}

function findConnector(doc: Document, shapeId: string): Element | null {
  const connectors = Array.from(doc.getElementsByTagNameNS(P_NS, "cxnSp"));
  return connectors.find((connector) => {
    const cNvPr = connector.getElementsByTagNameNS(P_NS, "cNvPr")[0];
    return cNvPr?.getAttribute("id") === shapeId;
  }) ?? null;
}

function ensureChild(doc: Document, parent: Element, namespaceUri: string, localName: string): Element {
  const existing = Array.from(parent.childNodes).find((node) =>
    node.nodeType === Node.ELEMENT_NODE &&
    (node as Element).namespaceURI === namespaceUri &&
    (node as Element).localName === localName
  ) as Element | undefined;
  if (existing) return existing;
  const prefix = namespaceUri === P_NS ? "p" : "a";
  const child = doc.createElementNS(namespaceUri, `${prefix}:${localName}`);
  parent.appendChild(child);
  return child;
}

function removeAllChildren(element: Element): void {
  while (element.firstChild) element.removeChild(element.firstChild);
}

function setTransform(doc: Document, spPr: Element, start: Point, end: Point): void {
  const xfrm = ensureChild(doc, spPr, A_NS, "xfrm");
  xfrm.removeAttribute("flipH");
  xfrm.removeAttribute("flipV");
  if (end.x < start.x) xfrm.setAttribute("flipH", "1");
  if (end.y < start.y) xfrm.setAttribute("flipV", "1");

  const off = ensureChild(doc, xfrm, A_NS, "off");
  off.setAttribute("x", String(toEmu(Math.min(start.x, end.x))));
  off.setAttribute("y", String(toEmu(Math.min(start.y, end.y))));

  const ext = ensureChild(doc, xfrm, A_NS, "ext");
  ext.setAttribute("cx", String(toEmu(Math.abs(end.x - start.x))));
  ext.setAttribute("cy", String(toEmu(Math.abs(end.y - start.y))));
}

function setPresetGeometry(doc: Document, spPr: Element, connectorType: "straight" | "elbow"): void {
  const geom = ensureChild(doc, spPr, A_NS, "prstGeom");
  geom.setAttribute("prst", connectorType === "elbow" ? "bentConnector3" : "line");
  let avLst = Array.from(geom.childNodes).find((node) =>
    node.nodeType === Node.ELEMENT_NODE &&
    (node as Element).namespaceURI === A_NS &&
    (node as Element).localName === "avLst"
  ) as Element | undefined;
  if (!avLst) {
    avLst = doc.createElementNS(A_NS, "a:avLst");
    geom.appendChild(avLst);
  }
}

function setLine(doc: Document, spPr: Element, patch: ConnectorXmlPatch): void {
  const line = ensureChild(doc, spPr, A_NS, "ln");
  removeAllChildren(line);
  line.setAttribute("w", String(toEmu(patch.thickness)));
  line.setAttribute("cap", "flat");
  line.setAttribute("cmpd", "sng");
  line.setAttribute("algn", "ctr");

  const solidFill = doc.createElementNS(A_NS, "a:solidFill");
  const srgbClr = doc.createElementNS(A_NS, "a:srgbClr");
  srgbClr.setAttribute("val", normalizeColor(patch.color));
  solidFill.appendChild(srgbClr);

  const dash = doc.createElementNS(A_NS, "a:prstDash");
  dash.setAttribute("val", normalizePresetDash(patch.dashStyle ?? "solid"));

  const headEnd = doc.createElementNS(A_NS, "a:headEnd");
  headEnd.setAttribute("type", "none");

  const tailEnd = doc.createElementNS(A_NS, "a:tailEnd");
  tailEnd.setAttribute("type", patch.arrow === "end" ? "arrow" : "none");
  if (patch.arrow === "end") {
    tailEnd.setAttribute("w", "med");
    tailEnd.setAttribute("len", "med");
  }

  const miter = doc.createElementNS(A_NS, "a:miter");
  miter.setAttribute("lim", "800000");
  line.append(solidFill, dash, headEnd, tailEnd, miter);
}

function normalizeDashStyle(value: string): string {
  const normalized = value.trim().toLowerCase();
  if (normalized === "dash" || normalized === "dashed") return "Dash";
  if (normalized === "dot" || normalized === "dotted") return "Dot";
  if (normalized === "dashdot" || normalized === "dashDot") return "DashDot";
  return "Solid";
}

function normalizePresetDash(value: string): string {
  const normalized = value.trim().toLowerCase();
  if (normalized === "dash" || normalized === "dashed") return "dash";
  if (normalized === "dot" || normalized === "dotted") return "dot";
  if (normalized === "dashdot" || normalized === "dashDot") return "dashDot";
  return "solid";
}

function normalizeColor(color: string): string {
  return color.replace(/^#/, "").toUpperCase();
}

function toEmu(points: number): number {
  return Math.round(points * POINT_TO_EMU);
}

// ── 布局算法 ──────────────────────────────────────────

export function sidePoint(rect: { left: number; top: number; width: number; height: number }, side: Side): { x: number; y: number } {
  switch (side) {
    case "top":    return { x: rect.left + rect.width / 2, y: rect.top };
    case "bottom": return { x: rect.left + rect.width / 2, y: rect.top + rect.height };
    case "left":   return { x: rect.left, y: rect.top + rect.height / 2 };
    case "right":  return { x: rect.left + rect.width, y: rect.top + rect.height / 2 };
  }
}
