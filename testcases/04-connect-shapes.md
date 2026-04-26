# 04 · 已有形状连线（connect_shapes）

## 前置

先手动在幻灯片上画两个矩形 A 和 B（或让 AI 画）：
- A 在左侧，B 在右侧，大致水平对齐

## 输入 prompt

```
把 A 矩形的右边连到 B 矩形的左边。
```

## 预期行为

- AI 调用 `get_current_context` 或 `list_slides` 获取形状 id
- 调用 `connect_shapes({ fromShapeId, fromSide: "right", toShapeId, toSide: "left", arrow: "end" })`
- 不应该出现 `add_line` 调用

## 检查点

- [ ] 线条从 A 右边中点出发，连到 B 左边中点
- [ ] 线条主体是真实 PowerPoint Straight connector，并通过 stCxn/endCxn 绑定两端形状
- [ ] 线条长度正好是两个形状之间的水平距离
- [ ] 线条不穿入矩形内部
- [ ] 箭头为 PowerPoint 原生 tailEnd 箭头，不出现三角形形状模拟箭头

## 变体

追加 prompt：
```
再把 B 的下边连到 A 的左边
```
预期：端点不同 X/Y → `connect_shapes` 返回信息里显示真实 bentConnector3 肘形连接器，并且用原生 tailEnd 箭头指向终点。
