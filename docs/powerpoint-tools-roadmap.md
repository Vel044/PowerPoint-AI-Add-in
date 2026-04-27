# PowerPoint 工具路线图

这份文档记录 Claude 风格 PowerPoint 工具体系的长期建设方向。实现时优先复用现有 `slideTarget`、`context`、`review_slide`、`ooxml`、`create_diagram/connect_shapes` 和 agent loop，不暴露未完成工具给模型。

## 状态说明

| 状态 | 含义 |
| --- | --- |
| `existing` | 当前已有能力，可能需要改名或补参数 |
| `v1` | 第一版要落地的核心工具 |
| `v2` | 第二批本地可用工具，已注册并可手测 |
| `later` | 后续增强 |
| `deferred` | 暂缓，不做假实现 |

## 读取与上下文

| 工具 | 状态 | 说明 |
| --- | --- | --- |
| `list_slides` | existing | 列出 slide index、pageNumber、slideId |
| `get_current_context` | existing | 读取当前页或任意页上下文 |
| `list_slide_shapes` | v1 | 读取某页所有形状，返回 `ref/id/name/type/text/bounds` |
| `read_slide_text` | v1 | 读取指定 shape 的 `<a:p>` 富文本 XML |
| `export_deck_outline` | v1 | 导出整份 PPT 的文本和形状大纲 |

## 编辑与自由绘图

| 工具 | 状态 | 说明 |
| --- | --- | --- |
| `add_text_box` | existing | 添加文本框，补样式参数 |
| `add_geometric_shape` | existing | 添加几何图形，补样式参数 |
| `modify_shape` | existing | 修改形状文字、位置、尺寸、样式 |
| `delete_shape` | existing | 删除指定页内指定形状 |
| `connect_shapes` | existing | 连接两个形状，补线条样式参数 |
| `create_diagram` | existing | 自动布局图，后续保留为快速入口 |
| `draw_slide_shapes` | v1 | 批量自由绘制复杂框图 |
| `edit_slide_text` | v1 | 用 OOXML 替换 shape 富文本段落 |

## OOXML 与底层能力

| 工具 | 状态 | 说明 |
| --- | --- | --- |
| `edit_slide_xml` | v1 | 结构化 OOXML 操作，不开放任意代码字符串 |
| `execute_office_js` | v2 | 白名单 Office.js 动作，不开放任意代码 |
| `edit_slide_chart` | v2 | 插入 shape-based 基础图表 |
| `edit_slide_master` | v2 | 受控修改背景、字体偏好、主题色、装饰元素 |

## 验证与调试

| 工具 | 状态 | 说明 |
| --- | --- | --- |
| `review_slide` | existing | 截图并调用视觉模型审查 |
| `verify_slide_visual` | v1 | `review_slide` 的 Claude 风格别名 |
| `verify_slides` | v1 | 程序化检查重叠、越界、空文本等问题 |
| `todo_write` | v1 | 将 Agent 任务步骤显示为工具结果 |

## 幻灯片管理

| 工具 | 状态 | 说明 |
| --- | --- | --- |
| `add_slide` | existing | 添加空白页 |
| `delete_slide` | existing | 删除指定页 |
| `duplicate_slide` | v1 | 复制指定页到原页后 |

## 图片、图标、资源

| 工具 | 状态 | 说明 |
| --- | --- | --- |
| `store_blob` | v2 | 保存 base64/url/text 资源到本地 IndexedDB |
| `copy_image_between_slides` | v2 | 用 `getImageAsBase64` + `addImage` 跨页复制图片 |
| `search_icons` | v2 | 搜索内置 SVG 图标索引 |
| `insert_icon` | v2 | 插入内置 SVG 图标，可改颜色和尺寸 |

## 偏好、技能、多 Agent

| 工具 | 状态 | 说明 |
| --- | --- | --- |
| `web_search` | v2 | 通过本地 dev server endpoint 轻量搜索 |
| `update_instructions` | v2 | 保存长期偏好到 localStorage，并拼接进 system prompt |
| `update_setting` | v2 | 管理本地功能设置 |
| `read_skill` | v2 | 读取仓库 `skills/*.md` |
| `create_skill` | v2 | 在仓库 `skills/` 生成技能草稿 |
| `get_connected_agents` | deferred | 需要外部多 Agent 协议 |
| `send_message` | deferred | 需要外部多 Agent 协议 |
| `refresh_mcp_connectors` | deferred | 需要外部连接器协议 |

## 第二批实现约束

- `execute_office_js` 只接受结构化 `actions[]`，不暴露任意 `code:string`。
- `edit_slide_chart` 第一版生成普通 shape-based 图表，优先可编辑和稳定，后续再升级为 native chart package。
- `edit_slide_master` 不开放任意 master XML，只做当前页背景和本地默认偏好。
- 图标库使用内置 SVG，不接微软远端图标服务。
- 多 Agent/MCP 类工具注册为明确 unsupported，避免模型误以为已经具备跨 Office 协作协议。
