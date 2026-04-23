# PowerPoint AI Add-in — PowerPoint 任务窗格插件

## 快速开始（Mac 桌面版 PowerPoint）

```bash
npm install
npx office-addin-dev-certs install      # 首次：安装本地开发证书
cp config/providers.example.json config/providers.json  # 若尚未创建
# 编辑 config/providers.json，填入 ANTHROPIC_AUTH_TOKEN

npm run dev          # 另开终端：HTTPS dev server @ https://localhost:3000
npm run sideload     # 自动把 manifest.xml 注册到 Mac PowerPoint 并启动
```

启动后 PowerPoint 的「开始」选项卡会出现「打开 AI 助手」按钮，点击即可在右侧调出任务窗格。

- 支持在任务窗格里切换 Provider（z.ai / minimax / anthropic 等），修改值会写入 `localStorage` 覆盖文件
- `Cmd+Enter` 发送消息
- 模型通过 `get_current_context / add_slide / modify_shape / delete_shape / add_text_box / export_pptx_xml / apply_pptx_patch` 等工具实际操作演示文稿

## 目录
- `manifest.xml` — Office Add-in 清单
- `src/taskpane/` — 任务窗格 UI（TS + CSS + HTML）
- `src/anthropic/` — `/v1/messages` 客户端 + tool_use 循环
- `src/tools/` — 基于 Office.js 与 JSZip 的工具实现
- `config/providers.json` — Provider 切换配置（沿用 cc-switch 格式）

