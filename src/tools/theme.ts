import JSZip from "jszip";

export interface ThemePalette {
  dk1?: string; dk2?: string; lt1?: string; lt2?: string;
  accent1?: string; accent2?: string; accent3?: string;
  accent4?: string; accent5?: string; accent6?: string;
  hlink?: string; folHlink?: string;
}

export interface ThemeInfo {
  themePalette: ThemePalette | null;
  isDefaultTheme: boolean;
  themeName: string | null;
  slideWidth: number | null;
  slideHeight: number | null;
}

const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const THEME_XML_PATH = "ppt/theme/theme1.xml";
const PRESENTATION_XML_PATH = "ppt/presentation.xml";
const EMU_PER_POINT = 12700;
const COLOR_KEYS = [
  "dk1", "dk2", "lt1", "lt2",
  "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
  "hlink", "folHlink",
];

const EMPTY_THEME: ThemeInfo = {
  themePalette: null,
  isDefaultTheme: true,
  themeName: null,
  slideWidth: null,
  slideHeight: null,
};

let cached: ThemeInfo | null = null;

export async function readThemePalette(forceRefresh = false): Promise<ThemeInfo> {
  if (cached && !forceRefresh) return cached;
  cached = await PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load("items/id");
    await ctx.sync();
    const firstSlide = slides.items[0];
    if (!firstSlide) return EMPTY_THEME;
    const exported = firstSlide.exportAsBase64();
    await ctx.sync();
    if (typeof exported.value !== "string" || exported.value.length === 0) return EMPTY_THEME;
    return await parseThemeFromBase64(exported.value);
  });
  return cached;
}

export function clearThemeCache(): void {
  cached = null;
}

async function parseThemeFromBase64(base64: string): Promise<ThemeInfo> {
  const zip = await JSZip.loadAsync(base64, { base64: true });

  const themeFields = await parseClrScheme(zip);
  const sizeFields = await parseSlideSize(zip);

  return { ...themeFields, ...sizeFields };
}

async function parseClrScheme(zip: JSZip): Promise<Pick<ThemeInfo, "themePalette" | "isDefaultTheme" | "themeName">> {
  const entry = zip.file(THEME_XML_PATH);
  if (!entry) return { themePalette: null, isDefaultTheme: true, themeName: null };
  const xml = await entry.async("string");
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  if (doc.getElementsByTagName("parsererror").length > 0) return { themePalette: null, isDefaultTheme: true, themeName: null };
  const scheme = doc.getElementsByTagNameNS(A_NS, "clrScheme")[0];
  if (!scheme) return { themePalette: null, isDefaultTheme: true, themeName: null };

  const themeName = scheme.getAttribute("name");
  const palette: ThemePalette = {};
  for (const key of COLOR_KEYS) {
    const node = scheme.getElementsByTagNameNS(A_NS, key)[0];
    if (!node) continue;
    const srgb = node.getElementsByTagNameNS(A_NS, "srgbClr")[0];
    const sysClr = node.getElementsByTagNameNS(A_NS, "sysClr")[0];
    const val = srgb?.getAttribute("val") ?? sysClr?.getAttribute("lastClr");
    if (val) (palette as Record<string, string>)[key] = val.toUpperCase();
  }
  const hasAny = Object.keys(palette).length > 0;
  const isDefaultTheme = themeName === "Office" || themeName === null || themeName === "";
  return {
    themePalette: hasAny ? palette : null,
    isDefaultTheme,
    themeName,
  };
}

async function parseSlideSize(zip: JSZip): Promise<Pick<ThemeInfo, "slideWidth" | "slideHeight">> {
  const entry = zip.file(PRESENTATION_XML_PATH);
  if (!entry) return { slideWidth: null, slideHeight: null };
  const xml = await entry.async("string");
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  if (doc.getElementsByTagName("parsererror").length > 0) return { slideWidth: null, slideHeight: null };
  const sldSz = doc.getElementsByTagNameNS(P_NS, "sldSz")[0];
  if (!sldSz) return { slideWidth: null, slideHeight: null };
  const cx = Number(sldSz.getAttribute("cx"));
  const cy = Number(sldSz.getAttribute("cy"));
  return {
    slideWidth: Number.isFinite(cx) ? Math.round(cx / EMU_PER_POINT) : null,
    slideHeight: Number.isFinite(cy) ? Math.round(cy / EMU_PER_POINT) : null,
  };
}
