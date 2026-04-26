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
删除 shape id 为 xxx 的形状。
```

预期：先 `get_current_context` 确认当前页 `allShapes` 中存在该形状，再 `delete_shape({ slideId, shapeId })`；不会跨页搜索同名/同 id 形状。

### C2. 视觉审查截图留档

```
检查当前页最近添加的图是否有重叠。
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

## 检查点

- [ ] 所有原工具正常工作
- [ ] AI 不会在非绘图场景误调用 `create_diagram`
- [ ] AI 不会在简单"画个框"场景误调用 `create_diagram`（应当用 `add_geometric_shape`）
- [ ] `get_current_context` 返回当前页 `allShapes` 和 `occupiedBounds`
