# 测试用例

绘图和工具能力的手动回归测试。每个文件是一组可以复制到 PowerPoint 右侧 AI 面板的中文 prompt，并列出预期 tool call 顺序和检查点。

## 怎么用

1. `npm run dev` 启动本地开发服务器。
2. PowerPoint 打开任一 `.pptx`，加载插件。
3. 打开一个 testcase 文件，把 **输入 prompt** 复制到输入框发送。
4. 对照 **预期 tool call** 和 **检查点** 验收。

## 文件索引

| 文件 | 场景 | 主要验证 |
| --- | --- | --- |
| [01-simple-flowchart.md](01-simple-flowchart.md) | 登录流程图 | `edit_slide_xml`、真实连接器、视觉验证 |
| [02-call-chain.md](02-call-chain.md) | 函数调用链 | 复杂图优先 Claude 风格 `edit_slide_xml({ code })` |
| [03-architecture.md](03-architecture.md) | 系统架构图 | `edit_slide_xml` 树状/分层 |
| [04-connect-shapes.md](04-connect-shapes.md) | 已有形状连线 | `connect_shapes` 贴边连接 |
| [05-mixed-edit.md](05-mixed-edit.md) | 在已有图上加节点 | `add_geometric_shape`、`modify_shape`、`connect_shapes` |
| [06-regression.md](06-regression.md) | 基础功能回归 | 页面、上下文、TODO、大纲、白名单动作 |
| [07-hard-case.md](07-hard-case.md) | 截图复现场景 | 自由排版调用链框图 |
| [08-v1-free-draw.md](08-v1-free-draw.md) | 自由框图 | `edit_slide_xml`、`verify_slides`、`verify_slide_visual` |
| [09-rich-text-ooxml.md](09-rich-text-ooxml.md) | 富文本与 OOXML | `read_slide_text`、`edit_slide_text`、`edit_slide_xml` |
| [10-chart-and-master.md](10-chart-and-master.md) | 图表与主题 | `edit_slide_chart`、`edit_slide_master` |
| [11-images-icons-blobs.md](11-images-icons-blobs.md) | 图片、图标、资源 | `store_blob`、`copy_image_between_slides`、`search_icons`、`insert_icon` |
| [12-settings-skills-search.md](12-settings-skills-search.md) | 设置、技能、搜索 | `web_search`、`update_instructions`、`update_setting`、`read_skill`、`create_skill` |
| [13-unsupported-bridges.md](13-unsupported-bridges.md) | 外部桥接占位 | unsupported 工具返回明确不可用 |

## 工具覆盖矩阵

| Tool | Testcase | 类型 | 预期顺序/说明 |
| --- | --- | --- | --- |
| `get_current_context` | 06 | manual-ppt | 读取当前页或指定页形状 |
| `list_slides` | 06 | manual-ppt | 列出 `index/pageNumber/id` |
| `list_slide_shapes` | 04, 09 | manual-ppt | 编辑前确认 `ref/id/bounds` |
| `export_deck_outline` | 06 | artifact | 导出大纲，可保存 artifact |
| `add_slide` | 06 | manual-ppt | 新增空白页 |
| `delete_slide` | 06 | manual-ppt | 删除指定页 |
| `duplicate_slide` | 06 | manual-ppt | 复制指定页并追踪新 id |
| `add_text_box` | 06 | manual-ppt | 添加标题文本框 |
| `add_geometric_shape` | 05 | manual-ppt | 增加普通节点 |
| `add_line` | 04 | manual-ppt | 基础线条回归 |
| `connect_shapes` | 04, 05 | manual-ppt | 真实连接器 |
| `read_slide_text` | 09 | manual-ppt | 读取 `<a:p>` |
| `edit_slide_text` | 09 | manual-ppt | 替换富文本 XML |
| `edit_slide_xml` | 01, 02, 03, 07, 08, 09 | manual-ppt | Claude 风格 `code:string`，普通绘图用 `pptx` helper，高级 patch 才直接编辑 slide1.xml |
| `modify_shape` | 05, 06 | manual-ppt | 修改文字、位置、样式 |
| `delete_shape` | 06 | manual-ppt | 指定页内删除 |
| `review_slide` | 01, 06 | artifact | 截图留档到 `debug-artifacts/` |
| `verify_slide_visual` | 08 | artifact | `review_slide` 别名 |
| `verify_slides` | 01, 08, 10 | manual-ppt | 程序化检查重叠/越界 |
| `todo_write` | 06 | manual-ppt | 输出任务步骤 |
| `execute_office_js` | 06 | manual-ppt | 白名单 move/resize/style/select |
| `edit_slide_chart` | 10 | manual-ppt | shape-based bar/line/pie |
| `edit_slide_master` | 10 | manual-ppt | 背景、字体偏好、装饰 |
| `store_blob` | 11 | manual-ppt | IndexedDB 资源保存 |
| `copy_image_between_slides` | 11 | manual-ppt | 图片形状复制 |
| `search_icons` | 11 | manual-ppt | 内置图标搜索 |
| `insert_icon` | 11 | manual-ppt | 插入 SVG 图标 |
| `web_search` | 12 | artifact | 本地 endpoint 轻量搜索 |
| `update_instructions` | 12 | manual-ppt | localStorage 长期偏好 |
| `update_setting` | 12 | manual-ppt | localStorage 设置 |
| `read_skill` | 12 | artifact | 读取 `skills/*.md` |
| `create_skill` | 12 | artifact | 创建技能草稿 |
| `get_connected_agents` | 13 | unsupported-expected | 明确不可用 |
| `send_message` | 13 | unsupported-expected | 明确不可用 |
| `refresh_mcp_connectors` | 13 | unsupported-expected | 明确不可用 |

## 通用验收关键词

- **节点不重叠**：任意两个框边界不相交，除非 prompt 明确要求叠放。
- **连接线贴边中点**：连接线端点指向框的上/下/左/右边中点，没有飞线。
- **图形自由排版**：流程图、调用链、架构细节、泳道图优先用 Claude 风格 `edit_slide_xml({ code })`，code 里优先调用 `pptx` helper。
- **修改前先读上下文**：删除、修改、富文本读取前必须用 `list_slide_shapes` 或 `get_current_context` 确认目标。
- **修改后必须验证**：绘图、图表、图标、连线后至少调用 `verify_slides`；视觉类任务再调用 `verify_slide_visual` 或 `review_slide`。
- **调试目录不入 git**：`debug-artifacts/` 必须被 `.gitignore` 忽略。
