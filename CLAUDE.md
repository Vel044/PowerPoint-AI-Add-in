# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

始终用简体中文回答所有问题。

## 项目概述

PowerPoint 任务窗格插件（Office Add-in），在右侧面板中提供 AI 助手，通过 Anthropic 兼容 API 的 tool_use 循环操作演示文稿。支持多 Provider（Z.AI / MiniMax / Anthropic 等），配置从 `config/providers.json` 加载，可在 UI 中切换并写入 localStorage 覆盖。

## 常用命令

```bash
npm install                                    # 安装依赖
npx office-addin-dev-certs install             # 首次：安装本地 HTTPS 开发证书
cp config/providers.example.json config/providers.json  # 创建配置并填入 API Key
npm run dev                                    # 启动 HTTPS dev server @ localhost:3000
npm run build                                  # 生产构建 → dist/
npm run sideload                               # 注册 manifest 到 Mac PowerPoint 并启动
npm run validate                               # 校验 manifest.xml
```

## 架构

入口为 `src/taskpane/taskpane.ts`，webpack 打包到 `dist/taskpane.js`，由 `taskpane.html` 加载。

**数据流**：用户输入 → `onSend()` 自动附加当前上下文 → `runAgentStream()` 驱动 agent 循环 → `callMessagesStream()` 调用 `/v1/messages` → 返回 tool_use 时查找 `TOOL_HANDLERS` 执行 → 将 tool_result 追加回消息历史 → 循环直到 `stop_reason !== "tool_use"`。

**关键模块**：

- `src/config.ts` — Provider 配置加载（fetch → localStorage 缓存）、模型按 tier 解析
- `src/anthropic/client.ts` — 流式/非流式调用 Anthropic Messages API，SSE 解析
- `src/anthropic/agentLoop.ts` — tool_use 循环（最多 10 轮），含 system prompt 定义
- `src/tools/registry.ts` — 工具定义（`TOOL_DEFINITIONS`）与处理器（`TOOL_HANDLERS`）的中央注册
- `src/tools/context.ts` — Office.js 读取当前上下文（幻灯片、选中形状）
- `src/tools/slides.ts` — 增删幻灯片
- `src/tools/shapes.ts` — 增删改文本框等形状，`resolveSlide()` 定位目标幻灯片
- `src/tools/ooxml.ts` — 内部单页 export/import XML 编辑桥接，用于修正真实 PowerPoint 连接器
- `src/types.ts` — `ContentBlock`、`Message`、`ToolHandler` 等共享类型

**操作模型**：
1. Office.js 层（`context.ts`、`slides.ts`、`shapes.ts`）— 实时编辑当前打开的演示文稿
2. 单页 XML 层（`ooxml.ts`）— 通过 PowerPointApi 1.8 `Slide.exportAsBase64()` 导出单页 PPTX，JSZip 修改 `ppt/slides/slide1.xml`，再用 `insertSlidesFromBase64()` 插回并删除原页

**配置系统**：`config/providers.json` 定义多个 provider 及其 env（base URL、API key、模型映射）。运行时通过 `localStorage("claude-for-office.providers")` 缓存，UI 可切换 provider 并覆盖 key。

## 开发注意事项

- 所有工具处理器签名为 `(input: Record<string, unknown>, ctx: ToolContext) => Promise<string>`，新增工具在 `registry.ts` 同时注册定义和处理器
- Office.js API 通过 `PowerPoint.run(callback)` 调用，回调内需 `load()` 属性再 `ctx.sync()`
- 单页 XML 能力只作为内部 helper 使用，不作为 Agent 可见工具注册；执行后 slide id 会变化，工具结果会返回新 slide id
- `manifest.xml` 中 `SourceLocation` 和图标 URL 指向 `https://localhost:3000`，部署时需替换
- TypeScript 编译目标 ES2020，模块解析用 Bundler 模式
