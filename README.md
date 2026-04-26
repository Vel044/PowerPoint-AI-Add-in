# PowerPoint AI Add-in — PowerPoint 任务窗格插件

## 快速开始（Mac 桌面版 PowerPoint）

```bash
npm install
npx office-addin-dev-certs install      # 首次：安装本地开发证书
cp config/providers.example.json config/providers.json  # 若尚未创建
# 编辑 config/providers.json，填入 ANTHROPIC_AUTH_TOKEN
lsof -ti:3001 | xargs kill -9
npm run sideload
npm run dev
```

启动后 PowerPoint 的「开始」选项卡会出现「打开 AI 助手」按钮，点击即可在右侧调出任务窗格。

- 支持在任务窗格里切换 Provider（z.ai / minimax / anthropic 等），修改值会写入 `localStorage` 覆盖文件
- `Cmd+Enter` 发送消息
- 模型通过 `get_current_context / add_slide / modify_shape / delete_shape / add_text_box / connect_shapes / create_diagram` 等工具实际操作演示文稿

## 目录
- `manifest.xml` — Office Add-in 清单
- `src/taskpane/` — 任务窗格 UI（TS + CSS + HTML）
- `src/anthropic/` — `/v1/messages` 客户端 + tool_use 循环
- `src/tools/` — 基于 Office.js 与单页 export/import XML 编辑能力的工具实现
- `config/providers.json` — Provider 切换配置（沿用 cc-switch 格式）

连接器 XML 修正通过 PowerPointApi 1.8 `Slide.exportAsBase64()` 导出单页 PPTX，JSZip 修改 `ppt/slides/slide1.xml`，再用 `insertSlidesFromBase64()` 插回并删除原页；执行后 slide id 会变化，工具结果会返回新 slide id。
