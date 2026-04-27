# 05 · 在已有图上加节点

## 前置

先跑 [01-simple-flowchart.md](01-simple-flowchart.md)，生成登录流程图。

## 输入 prompt

```
在"判断是否验证通过"的菱形右边加一个"记录登录日志"的矩形，并把菱形的右边连到它的左边。
```

## 预期行为

- AI 调用 `get_current_context` 获得当前页 `slideId`、`allShapes`、已有节点 id 和位置
- 调用 `add_geometric_shape` 加一个 rectangle，坐标在菱形右侧
- 调用 `connect_shapes(fromSide:"right", toSide:"left")`
- 不应重建整张图

## 检查点

- [ ] 新矩形在菱形右侧、纵坐标对齐
- [ ] 新连接线贴菱形右边中点和新矩形左边中点
- [ ] 原有节点和连线没动
- [ ] AI 明确知道"在已有图上加"，没有重画整张图
- [ ] 若需要删除/修改旧形状，tool input 必须带当前页 `slideId`，不跨页凭裸 `shapeId` 操作
