# 03 · 系统架构图（tree）

## 输入 prompt

```
画一个简单的 Web 应用架构：前端（React）→ API 网关 → 两个后端服务（用户服务、订单服务）→ 各自连数据库（Postgres、MongoDB）。用树状布局。
```

## 预期行为

- `create_diagram({ layout: "tree", ... })`
- 节点 shape 混用：前端/网关用 `roundRectangle`，服务用 `rectangle`，数据库用 `flowChartData` 或 `ellipse`
- edges 构成 3 层树

## 检查点

- [ ] 三层清晰，从上到下展开
- [ ] 同父节点的子节点并排居中在父节点下方
- [ ] 每条 edge 都是带箭头的几何形状
- [ ] 两条数据库分支不交叉
- [ ] 整图水平居中
