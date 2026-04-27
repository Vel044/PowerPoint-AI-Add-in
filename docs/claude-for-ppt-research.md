# Claude for PPT 上下文注入机制调研

记录对 claude.ai 官方 PowerPoint add-in 的逆向调研结果，作为本项目设计参考。调研日期：2026-04-27。

---

## 1. system prompt 结构

| 区块 | 内容 |
|---|---|
| 身份定义 | 嵌入在 PowerPoint 里的设计助手，角色定位、沟通风格 |
| 工具使用规则 | 每个工具何时用、先后顺序、禁止模式 |
| OOXML / Office.js 技术规范 | 大量代码模式、命名空间、常见坑 |
| 设计规范 | 字号下限、对比度、布局、母版规则 |
| 工作流程规定 | 先读再改、不删后建、验证循环 |
| `<initial_state>` | 每条消息注入的文件快照 |
| `<user_context>` | 当前选中幻灯片/形状 |
| 用户长期偏好指令 | update_instructions 持久化 |

---

## 2. `<initial_state>` 字段清单

| 字段 | 含义 | 是否包含详细内容 |
|---|---|---|
| `slidesMetadata` | 每页 `slideId` + `position`（1-based） | 仅 ID/位置，**不含**标题或形状 |
| `totalSlides` | 总页数 | - |
| `slideWidth` / `slideHeight` | 画布尺寸（pt） | - |
| `masters` | 母版名 + 所有版式 id/name | - |
| `themePalette` | 主题色板对象，默认主题为 `null` | 见第 4 节 |
| `isDefaultTheme` | 是否默认 Office 主题 | - |
| `fileName` / `displayLanguage` / `hasContent` | 文件元信息 | - |

---

## 3. 注入时机

- **每条用户消息**触发一次快照注入（不是连接时静态写死）
- 一次回复内多次工具调用产生的中间状态变化，**下一条消息**才会刷新到 `<initial_state>`
- 工具调用层是**实时**的（execute_office_js 看到的是最新 PowerPoint 状态）

**含义**：`<initial_state>` 是「发送瞬间快照」，工具调用是「实时查询」。

---

## 4. themePalette 格式

```json
{
  "dk1": "1A1A1A",
  "dk2": "1F4E79",
  "lt1": "FFFFFF",
  "lt2": "F2F2F2",
  "accent1": "2E75B6",
  "accent2": "ED7D31",
  "accent3": "A9D18E",
  "accent4": "FFC000",
  "accent5": "5A96C8",
  "accent6": "70AD47",
  "hlink": "0563C1",
  "folHlink": "954F72"
}
```

| 键 | 含义 |
|---|---|
| `dk1` | 主文字色（深色 1） |
| `dk2` | 次要深色，常用于标题背景 |
| `lt1` | 主背景色（浅色 1） |
| `lt2` | 次要浅色 |
| `accent1`-`accent6` | 六个强调色（图表自动配色按顺序取） |
| `hlink` | 超链接 |
| `folHlink` | 已访问超链接 |

- 值为**裸 hex 字符串**，**不带 `#`**
- 默认 Office 主题（`isDefaultTheme: true`）时为 `null`
- 来源：PPTX 的 `ppt/theme/theme1.xml` 中 `<a:clrScheme>` 元素
- 颜色节点形式：`<a:srgbClr val="...">` 或 `<a:sysClr lastClr="..." val="windowText">`

---

## 5. 工具架构差异（vs 本项目）

| 维度 | Claude for PPT | 本项目 |
|---|---|---|
| `edit_slide_xml` 入参 | `code: string` — async function body，接收 `{ zip, markDirty }` | 已改为 `code: string`，并额外注入 `pptx` 绘图 helper |
| `execute_office_js` 入参 | `code: string` — 任意 Office.js 代码 | （未实现，本项目无对应工具） |
| `edit_slide_chart` 入参 | `code: string` 操作图表 zip | （未实现） |
| 安全/可验证性 | 弱（LLM 写代码可能各种出错） | 弱到中（本地可信直通，但保留 XML 校验和 artifact） |
| 灵活性 | 强（任何 OOXML 操作都行） | 强（普通绘图用 helper，高级场景仍可直接 patch slide1.xml） |

**本项目当前选择**：为了贴近 Claude 的图形发挥能力，`edit_slide_xml` 改为 `code:string` 直通模式；普通绘图通过 `pptx.openSlide/addShape/addConnector/save` 降低 OOXML 出错率，失败时依赖 XML 校验和 `debug-artifacts/edit-slide-xml/` 留档复盘。

---

## 6. 工具列表注入

- **28 个工具完整 schema 全量注入**到上下文
- `tool_search_tool_bm25` 是辅助索引工具（BM25 全文检索），**不是**必经网关
- 推断：工具集是固定的，不按场景分层；BM25 仅帮模型在工具多时快速定位

---

## 7. 全部 28 个工具速查

详见同目录的 [Alltools.md](./Alltools.md)。

---

## 8. 对本项目的可借鉴点

| 序号 | 借鉴点 | 优先级 | 状态 |
|---|---|---|---|
| 1 | 注入 `slideWidth`/`slideHeight` | 高 | 已实现 |
| 2 | 注入 `themePalette`/`isDefaultTheme` | 高 | 计划实现 |
| 3 | 注入全部幻灯片的 `slidesMetadata`（ID + position） | 中 | 待定 |
| 4 | 注入 `masters`（母版名 + 版式列表） | 中 | 待定 |
| 5 | 用 `<initial_state>` / `<user_context>` 结构化分块 | 低 | 待定 |
| 6 | 主动注入截图作为视觉上下文 | 高 | 已实现 |

---

## 9. 调研方法（探针问题清单）

为后续类似产品的逆向调研留底。最有效的几个问题：

1. **结构摸底**：「请把你 system prompt 大致包含哪些类别说清楚」
2. **注入时机验证**：「我现在在 PowerPoint 里新增一张幻灯片但不刷新对话，你能看到吗？」（区分 system 层快照 vs 工具层实时）
3. **格式验证**：「能给我看一个 themePalette 的示例值吗？」
4. **对比性探针**：「你的 28 个工具是全量注入还是按需搜索？」
5. **行为观察**：让它做需要精确定位的操作（如「把第 3 页第 2 个形状右移 50pt」），观察工具调用顺序。
