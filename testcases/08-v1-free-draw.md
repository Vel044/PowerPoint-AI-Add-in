# 08 · XML 自由框图

覆盖 `edit_slide_xml`、`verify_slides`、`verify_slide_visual`。

## 输入 prompt

```
在当前幻灯片画一个“订单创建调用链”框图：左到右是层次、上到下是时间。包含 API Controller、OrderService、InventoryService、PaymentService、OrderDB 五个框；主调用从 Controller 到 OrderService，再分支到库存和支付，最后写 OrderDB。请自己决定坐标、颜色和连接线，画完后做程序化检查和截图检查。
```

## 预期 tool call

1. `list_slide_shapes` 或 `get_current_context`
2. `edit_slide_xml`
3. `verify_slides`
4. `verify_slide_visual`

## 检查点

- [ ] 复杂框图使用 Claude 风格 `edit_slide_xml({ code })`，且 code 里使用 `pptx` helper 而不是大段 `createElementNS`
- [ ] 节点按语义使用不同 `style`，例如服务、数据库、错误/警告、注释不要全部同色
- [ ] 节点不重叠，连接线方向清楚
- [ ] `verify_slide_visual` 返回截图 artifact 路径
