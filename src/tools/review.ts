import { ToolHandler } from "../types";
import { resolveSlide } from "./shapes";
import { getActiveProvider, loadConfig, resolveModel } from "../config";

export const reviewSlide: ToolHandler = async (input) => {
  const slideIndex = input.slideIndex as number | undefined;
  const slideId = input.slideId as string | undefined;
  const question =
    (input.question as string) ||
    '请只检查这张幻灯片上最近新添加的形状和连线，忽略旧内容。检查是否有：1) 形状重叠 2) 连线歪斜（没从一个形状中心连到另一个形状中心）3) 文字溢出形状边界 4) 布局明显不对齐。请用 JSON 回复：{"ok": true/false, "issues": [{"type": "重叠/歪斜/溢出/不对齐", "desc": "描述", "fix": "修正建议"}]}。没问题则 {"ok": true, "issues": []}';

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
