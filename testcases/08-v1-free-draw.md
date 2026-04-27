# 08 · 自由框图

覆盖 `draw_slide_shapes`、`verify_slides`、`verify_slide_visual`。

## 输入 prompt

```
在当前幻灯片画一个“订单创建调用链”框图：左到右是层次、上到下是时间。包含 API Controller、OrderService、InventoryService、PaymentService、OrderDB 五个框；主调用从 Controller 到 OrderService，再分支到库存和支付，最后写 OrderDB。请自己决定坐标、颜色和连接线，画完后做程序化检查和截图检查。
```

## 预期 tool call

1. `list_slide_shapes` 或 `get_current_context`
2. `draw_slide_shapes`
3. `verify_slides`
4. `verify_slide_visual`

## 检查点

- [ ] 复杂框图使用 `draw_slide_shapes`，不是 `create_diagram`
- [ ] 节点不重叠，连接线方向清楚
- [ ] `verify_slide_visual` 返回截图 artifact 路径
