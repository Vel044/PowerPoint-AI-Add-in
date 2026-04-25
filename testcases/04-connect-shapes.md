# 04 · 已有形状连线（connect_shapes）

## 前置

先手动在幻灯片上画两个矩形 A 和 B（或让 AI 画）：
- A 在左侧，B 在右侧，大致水平对齐

## 输入 prompt

```
把 A 矩形的右边连到 B 矩形的左边，带箭头指向 B。
```

## 预期行为

- AI 调用 `get_current_context` 或 `list_slides` 获取形状 id
- 调用 `connect_shapes({ fromShapeId, fromSide: "right", toShapeId, toSide: "left", arrow: "end" })`
- 不应该出现 `add_line` 调用

## 检查点

- [ ] 箭头从 A 右边中点出发，指向 B 左边中点
- [ ] 箭头是几何连接器（细矩形线段 + 箭头头，可以选中、看到填充色）
- [ ] 箭头长度正好是两个形状之间的水平距离
- [ ] 箭头不穿入矩形内部

## 变体

追加 prompt：
```
再把 B 的下边连到 A 的下边
```
预期：斜向连接 → `connect_shapes` 返回信息里有"斜向连接无法加箭头头"提示，画出的是 elbow 无头线。
