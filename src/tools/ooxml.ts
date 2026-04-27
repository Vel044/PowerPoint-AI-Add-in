import JSZip from "jszip";
import { withOfficeErrorContext } from "./officeErrors";
import { resolveSlide, SlideTarget } from "./slideTarget";

type SlideXmlMutator = (doc: Document) => void;

export interface SlideXmlEditResult {
  oldSlideId: string;
  newSlideId: string;
  oldSlideIndex: number;
  newSlideIndex: number;
}

const SLIDE_XML_PATH = "ppt/slides/slide1.xml";

export async function editCurrentSlideXml(target: SlideTarget, mutator: SlideXmlMutator): Promise<SlideXmlEditResult> {
  return await PowerPoint.run(async (ctx) => {
    const presentation = ctx.presentation;
    const slide = await resolveSlide(ctx, target);
    const slides = presentation.slides;
    slides.load("items/id,index");
    slide.load("id,index");
    const exported = slide.exportAsBase64();
    await syncWithContext(ctx, "导出目标幻灯片 exportAsBase64 失败");

    if (typeof exported.value !== "string" || exported.value.length === 0) {
      throw new Error("导出当前幻灯片失败：exportAsBase64 未返回有效 base64。请确认 Office 支持 PowerPointApi 1.8。");
    }

    const oldSlideId = slide.id;
    const oldSlideIndex = slide.index;
    const existingSlideIds = new Set(slides.items.map((item) => item.id));
    const patchedBase64 = await patchSlideXml(exported.value, mutator);

    presentation.insertSlidesFromBase64(patchedBase64, {
      targetSlideId: oldSlideId,
      formatting: "KeepSourceFormatting",
    });
    await syncWithContext(ctx, `插入修正后的单页 PPTX 失败，targetSlideId=${oldSlideId}`);

    slides.load("items/id,index");
    await syncWithContext(ctx, "读取插入后的幻灯片列表失败");
    const insertedSlide = findInsertedSlide(slides.items, existingSlideIds, oldSlideIndex);
    if (!insertedSlide) {
      throw new Error("已插入修正后的幻灯片，但无法定位新 slide id。");
    }
    const newSlideId = insertedSlide.id;

    slide.delete();
    await syncWithContext(ctx, `删除旧幻灯片失败，oldSlideId=${oldSlideId}`);

    presentation.setSelectedSlides([newSlideId]);
    slides.load("items/id,index");
    await syncWithContext(ctx, `选中修正后的幻灯片失败，newSlideId=${newSlideId}`);
    const selected = slides.items.find((item) => item.id === newSlideId);
    return {
      oldSlideId,
      newSlideId,
      oldSlideIndex,
      newSlideIndex: selected?.index ?? oldSlideIndex,
    };
  });
}

async function syncWithContext(ctx: PowerPoint.RequestContext, message: string): Promise<void> {
  try {
    await ctx.sync();
  } catch (error) {
    throw withOfficeErrorContext(error, message);
  }
}

async function patchSlideXml(base64: string, mutator: SlideXmlMutator): Promise<string> {
  const zip = await JSZip.loadAsync(base64, { base64: true });
  const entry = zip.file(SLIDE_XML_PATH);
  if (!entry) throw new Error(`导出的单页 PPTX 中未找到 ${SLIDE_XML_PATH}`);

  const xml = await entry.async("string");
  const doc = parseSlideXml(xml);
  mutator(doc);
  zip.file(SLIDE_XML_PATH, new XMLSerializer().serializeToString(doc));
  return await zip.generateAsync({
    type: "base64",
    mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  });
}

function parseSlideXml(xml: string): Document {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const error = doc.getElementsByTagName("parsererror")[0];
  if (error) {
    throw new Error(`解析 slide XML 失败: ${error.textContent?.trim() ?? "unknown parser error"}`);
  }
  return doc;
}

function findInsertedSlide(
  slides: PowerPoint.Slide[],
  existingSlideIds: Set<string>,
  oldSlideIndex: number,
): PowerPoint.Slide | null {
  const newSlides = slides.filter((slide) => !existingSlideIds.has(slide.id));
  if (newSlides.length === 1) return newSlides[0];
  return slides.find((slide) => slide.index === oldSlideIndex + 1 && !existingSlideIds.has(slide.id)) ?? null;
}
