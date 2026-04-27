import { ToolHandler } from "../types";
import { getActiveProvider, loadConfig, resolveModel } from "../config";
import { pageNumberFromIndex, resolveSlide, slideTargetFromInput } from "./slideTarget";

export const reviewSlide: ToolHandler = async (input, toolCtx) => {
  const target = slideTargetFromInput(input);
  const question =
    (input.question as string) ||
    '请只检查这张幻灯片上最近新添加的形状和连线，忽略旧内容。检查是否有：1) 形状重叠 2) 连线歪斜（没从一个形状中心连到另一个形状中心）3) 文字溢出形状边界 4) 布局明显不对齐。请用 JSON 回复：{"ok": true/false, "issues": [{"type": "重叠/歪斜/溢出/不对齐", "desc": "描述", "fix": "修正建议"}]}。没问题则 {"ok": true, "issues": []}';

  return await PowerPoint.run(async (ctx) => {
    const slide = await resolveSlide(ctx, target);

    let base64Png: string;
    let actualSlideIndex: number | null = null;
    let actualPageNumber: number | null = null;
    try {
      const imgResult = (slide as any).getImageAsBase64({ width: 1280 });
      await ctx.sync();
      base64Png = imgResult.value;
      const slides = ctx.presentation.slides;
      slides.load("items/id,index");
      await ctx.sync();
      actualSlideIndex = slides.items.findIndex((s) => s.id === slide.id);
      actualPageNumber = pageNumberFromIndex(actualSlideIndex);
    } catch {
      return "截图功能不可用（当前 PowerPoint 版本不支持 getImageAsBase64）。请用 get_current_context 检查形状位置来自查。";
    }

    const artifact = await saveReviewScreenshot(base64Png, {
      tool: "review_slide",
      slideId: slide.id,
      slideIndex: actualSlideIndex,
      pageNumber: actualPageNumber,
      requestedSlideId: target.slideId ?? null,
      requestedSlideIndex: target.slideIndex ?? null,
      requestedPageNumber: target.pageNumber ?? null,
      question,
      timestamp: new Date().toISOString()
    });
    const artifactPrefix = artifact?.path ? `截图已保存: ${artifact.path}\n` : "";
    if (artifact?.path) {
      toolCtx.log(`[review_slide] 截图已保存: ${artifact.path}`);
    }

    const config = await loadConfig();
    const provider = getActiveProvider(config);
    const model = resolveModel(provider, "sonnet");
    const base = provider.env.ANTHROPIC_BASE_URL.replace(/\/$/, "");
    const token = provider.env.ANTHROPIC_AUTH_TOKEN;
    if (!token) throw new Error("当前 Provider 未配置 ANTHROPIC_AUTH_TOKEN");

    const headers: Record<string, string> = {
      "Content-Type": "application/json",
      Authorization: `Bearer ${token}`,
    };
    if (/(^|\.)anthropic\.com$/i.test(new URL(base).hostname)) {
      headers["anthropic-version"] = "2023-06-01";
      headers["x-api-key"] = token;
    }

    const res = await fetch(`${base}/v1/messages`, {
      method: "POST",
      headers,
      body: JSON.stringify({
        model,
        max_tokens: 1024,
        messages: [
          {
            role: "user",
            content: [
              {
                type: "image",
                source: {
                  type: "base64",
                  media_type: "image/png",
                  data: base64Png,
                },
              },
              { type: "text", text: question },
            ],
          },
        ],
      }),
    });

    if (!res.ok) {
      const text = await res.text();
      return `${artifactPrefix}视觉审查 API 调用失败 (${res.status}): ${text.slice(0, 300)}`;
    }

    const json = await res.json();
    const textBlocks = (json.content ?? []).filter(
      (b: Record<string, unknown>) => b.type === "text"
    );
    if (textBlocks.length === 0) return `${artifactPrefix}视觉审查未返回文本结果。`;
    return artifactPrefix + textBlocks
      .map((b: Record<string, unknown>) => (b as any).text ?? "")
      .join("\n");
  });
};

interface ArtifactResult {
  ok: boolean;
  path?: string;
  metadataPath?: string;
}

async function saveReviewScreenshot(base64Png: string, metadata: Record<string, unknown>): Promise<ArtifactResult | null> {
  try {
    const slidePart = safeFilePart(`${metadata.slideIndex ?? "unknown"}-${metadata.slideId ?? "slide"}`);
    const stamp = new Date().toISOString().replace(/[:.]/g, "-");
    const suffix = Math.random().toString(36).slice(2, 8);
    const res = await fetch("https://localhost:3001/__debug-artifact", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        kind: "review-slide",
        filename: `review-slide-${stamp}-${slidePart}-${suffix}.png`,
        mime: "image/png",
        base64: base64Png,
        metadata
      })
    });
    if (!res.ok) return null;
    return await res.json() as ArtifactResult;
  } catch {
    return null;
  }
}

function safeFilePart(value: string): string {
  return value.replace(/[^a-zA-Z0-9._-]/g, "_").slice(0, 80) || "slide";
}
