# 06 · 基础功能回归

确保改造没破坏原有能力。

## 输入 prompt 清单

### A. 文本框

```
在当前幻灯片顶部加一个标题文本框"季度回顾 2026 Q1"，大号字。
```

预期：`add_text_box`，不涉及图工具。

### B. 修改形状文字

先选中一个矩形，然后：
```
把选中形状的文字改成"已完成"。
```

预期：先 `get_current_context`，再 `modify_shape`，并使用当前页的 `slideId + shapeId`。

### C. 删除形状

```
删除 第三页SQL 的形状。
```

预期：先 `get_current_context({ pageNumber: 3 })` 或 `list_slide_shapes({ pageNumber: 3 })` 查看第三页，再 `delete_shape({ pageNumber: 3, shapeId })`；不会跨页搜索同名/同 id 形状。

### C2. 视觉审查截图留档

```
检查当前页最近添加的图形是否有重叠。有的话移动一下让他们分开
```

预期：`review_slide`，tool result 和终端日志都包含 `debug-artifacts/review-slide/*.png` 路径，并生成同名 `.json` 元数据。

### D. 新增/删除幻灯片

```
在末尾加一张新幻灯片。
```

```
删除第 3 张幻灯片。
```

预期：`add_slide` / `delete_slide`。

### E. 复制页与大纲

```
复制第 2 页，然后导出整份 PPT 的文字大纲。
```

预期：`duplicate_slide({ pageNumber: 2 })`，再 `export_deck_outline`。大纲结果包含每页 `pageNumber/slideId/textShapes`。

### F. TODO 进度

```
把任务拆成三步显示：读取现状、修改图形、验证效果。第一步标记已完成，第二步进行中，第三步待办。
```

预期：调用 `todo_write`，返回完整 todos。

### G. 白名单 Office.js 动作

先用 `list_slide_shapes` 找到一个形状，然后：

```
把刚才那个形状向右移动 20pt，并置于顶层。
```

预期：先 `list_slide_shapes`，再 `execute_office_js({ actions:[{type:"moveShape"...},{type:"bringToFront"...}] })`。不得传任意 `code` 字符串。

## 检查点

- [ ] 所有原工具正常工作
- [ ] AI 不会在非绘图场景误用绘图工具
- [ ] AI 不会在简单"画个框"场景误走批量 XML（应当用 `add_geometric_shape`）
- [ ] `get_current_context` 返回当前页 `allShapes` 和 `occupiedBounds`
- [ ] `list_slides` 返回 `index/pageNumber/id`
- [ ] `duplicate_slide` 后新页能被后续工具定位
- [ ] `execute_office_js` 只接受白名单 actions
