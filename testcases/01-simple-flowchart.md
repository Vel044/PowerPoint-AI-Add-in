# 01 · 登录流程图（vertical）

## 输入 prompt

```
在当前幻灯片画一个登录流程图：开始 → 输入用户名密码 → 判断是否验证通过 → 通过则进入首页、不通过则提示错误 → 结束。
```

## 预期行为

- AI 调用 `get_current_context`
- 调用一次 `create_diagram`，参数大致为：
  - `layout: "vertical"` 或 `"tree"`
  - 6 个节点，起止用 `flowChartTerminator`、判断用 `diamond`、过程用 `rectangle`
  - 6 条 edges（包括判断节点到两条分支再汇合到结束）

## 检查点

- [ ] 节点垂直等间距排列，水平居中
- [ ] 同类形状尺寸一致
- [ ] 每条连接线主体是原生 PowerPoint Straight connector 线段（不是细矩形伪线，也不是 Elbow 自动路由）
- [ ] 当前源码箭头开关关闭，不出现三角形箭头头
- [ ] 菱形的两条分支（通过/不通过）连接方向合理
- [ ] AI 最后有 2-4 句中文总结
