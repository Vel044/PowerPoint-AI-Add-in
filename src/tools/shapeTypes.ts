function token(value: string): string {
  return value.replace(/[\s_-]/g, "").toLowerCase();
}

const GEOMETRIC_SHAPE_ALIASES: Record<string, string> = {
  rectangle: "Rectangle",
  rect: "Rectangle",
  roundrectangle: "RoundRectangle",
  roundedrectangle: "RoundRectangle",
  roundrect: "RoundRectangle",
  ellipse: "Ellipse",
  oval: "Ellipse",
  diamond: "Diamond",
  decision: "FlowChartDecision",
  triangle: "Triangle",
  righttriangle: "RightTriangle",
  parallelogram: "Parallelogram",
  trapezoid: "Trapezoid",
  pentagon: "Pentagon",
  hexagon: "Hexagon",
  octagon: "Octagon",
  plus: "Plus",
  rightarrow: "RightArrow",
  leftarrow: "LeftArrow",
  uparrow: "UpArrow",
  downarrow: "DownArrow",
  star5: "Star5",
  can: "Can",
  cylinder: "Can",
  database: "Can",
  db: "Can",
  datastore: "Can",
  process: "FlowChartProcess",
  terminator: "FlowChartTerminator",
  startend: "FlowChartTerminator",
  flowchartprocess: "FlowChartProcess",
  flowchartdecision: "FlowChartDecision",
  flowchartterminator: "FlowChartTerminator",
  flowchartdata: "FlowChartInputOutput",
  data: "FlowChartInputOutput",
  inputoutput: "FlowChartInputOutput",
  io: "FlowChartInputOutput",
  flowchartinputoutput: "FlowChartInputOutput",
  flowchartpredefinedprocess: "FlowChartPredefinedProcess",
  flowchartinternalstorage: "FlowChartInternalStorage",
  flowchartdocument: "FlowChartDocument",
  flowchartmultidocument: "FlowChartMultidocument",
  flowchartpreparation: "FlowChartPreparation",
  flowchartmanualinput: "FlowChartManualInput",
  flowchartmanualoperation: "FlowChartManualOperation",
  flowchartconnector: "FlowChartConnector",
  flowchartpunchedcard: "FlowChartPunchedCard",
  flowchartpunchedtape: "FlowChartPunchedTape",
  flowchartsummingjunction: "FlowChartSummingJunction",
  flowchartor: "FlowChartOr",
  flowchartcollate: "FlowChartCollate",
  flowchartsort: "FlowChartSort",
  flowchartextract: "FlowChartExtract",
  flowchartmerge: "FlowChartMerge",
  flowchartofflinestorage: "FlowChartOfflineStorage",
  flowchartonlinestorage: "FlowChartOnlineStorage",
  flowchartmagnetictape: "FlowChartMagneticTape",
  flowchartmagneticdisk: "FlowChartMagneticDisk",
  flowchartmagneticdrum: "FlowChartMagneticDrum",
  flowchartdisplay: "FlowChartDisplay",
  flowchartdelay: "FlowChartDelay",
  flowchartalternateprocess: "FlowChartAlternateProcess",
  flowchartoffpageconnector: "FlowChartOffpageConnector",
};

const CONNECTOR_ALIASES: Record<string, string> = {
  straight: "Straight",
  line: "Straight",
  direct: "Straight",
  elbow: "Elbow",
  orthogonal: "Elbow",
  curve: "Curve",
  curved: "Curve",
};

export function normalizeGeometricShapeType(input: unknown): PowerPoint.GeometricShapeType {
  const raw = String(input ?? "rectangle").trim();
  return normalizeOfficeEnumValue(raw, "GeometricShapeType", GEOMETRIC_SHAPE_ALIASES) as PowerPoint.GeometricShapeType;
}

export function normalizeConnectorType(input: unknown): PowerPoint.ConnectorType {
  const raw = String(input ?? "straight").trim();
  return normalizeOfficeEnumValue(raw, "ConnectorType", CONNECTOR_ALIASES) as PowerPoint.ConnectorType;
}

function normalizeOfficeEnumValue(
  raw: string,
  enumName: "GeometricShapeType" | "ConnectorType",
  aliases: Record<string, string>,
): string {
  if (!raw) raw = enumName === "ConnectorType" ? "straight" : "rectangle";

  const normalized = token(raw);
  const enumValue = findRuntimeEnumValue(enumName, raw, normalized);
  if (enumValue) return enumValue;

  const alias = aliases[normalized];
  if (alias) return alias;

  const fallback = raw[0].toUpperCase() + raw.slice(1);
  const fallbackMatch = findRuntimeEnumValue(enumName, fallback, token(fallback));
  if (fallbackMatch) return fallbackMatch;

  const supported = Object.keys(aliases).slice(0, 16).join(", ");
  throw new Error(`不支持的 ${enumName}: ${raw}。常用可用值包括：${supported}`);
}

function findRuntimeEnumValue(
  enumName: "GeometricShapeType" | "ConnectorType",
  raw: string,
  normalized: string,
): string | undefined {
  const maybePowerPoint = (globalThis as unknown as {
    PowerPoint?: {
      GeometricShapeType?: Record<string, string>;
      ConnectorType?: Record<string, string>;
    };
  }).PowerPoint;
  const enumObj = maybePowerPoint?.[enumName];
  if (!enumObj) return undefined;

  if (typeof enumObj[raw] === "string") return enumObj[raw];
  const exactValue = Object.values(enumObj).find((value) => value === raw);
  if (exactValue) return exactValue;
  const normalizedMatch = Object.entries(enumObj).find(([key, value]) =>
    token(key) === normalized || token(value) === normalized
  );
  return normalizedMatch?.[1];
}
