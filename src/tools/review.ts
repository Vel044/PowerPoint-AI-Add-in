import { ToolHandler } from "../types";
import { resolveSlide } from "./shapes";
import { getActiveProvider, loadConfig, resolveModel } from "../config";

export const reviewSlide: ToolHandler = async (input) => {
  const slideIndex = input.slideIndex as number | undefined;
  const slideId = input.slideId as string | undefined;
  const question =
    (input.question as string) ||
    "请检查这张幻灯片上的形状布局是否有问题：重叠、对齐偏移、连线歪斜、文字溢出等。如果有问题，说明具体是什么以及如何修正。如果没问题，直接说'没问题'。";

  return await PowerPoint.run(async (ctx) => {
    const slide = await resolveSlide(ctx, slideId, slideIndex);

    let base64Png: string;
    try {
      const imgResult = (slide as any).getImageAsBase64({ width: 1280 });
      await ctx.sync();
      base64Png = imgResult.value;
    } catch {
      return "截图功能不可用（当前 PowerPoint 版本不支持 getImageAsBase64）。请用 get_current_context 检查形状位置来自查。";
    }

    const config = await loadConfig();
    const provider = getActiveProvider(config);
    const model = resolveModel(provider, "opus");
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
      return `视觉审查 API 调用失败 (${res.status}): ${text.slice(0, 300)}`;
    }

    const json = await res.json();
    const textBlocks = (json.content ?? []).filter(
      (b: Record<string, unknown>) => b.type === "text"
    );
    if (textBlocks.length === 0) return "视觉审查未返回文本结果。";
    return textBlocks
      .map((b: Record<string, unknown>) => (b as any).text ?? "")
      .join("\n");
  });
};
