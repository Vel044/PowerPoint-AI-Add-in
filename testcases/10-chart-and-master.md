# 10 · 图表与主题

覆盖 `edit_slide_chart`、`edit_slide_master`。

## 输入 prompt

```
在当前页生成一个 2026 Q1-Q4 收入图表，categories 是 Q1、Q2、Q3、Q4，series 是 Revenue=[12,18,16,24] 和 Cost=[8,11,10,13]，用柱状图。再把当前页背景设成 #F7F9FC，加一条顶部深蓝装饰条。完成后检查重叠和越界。
```

## 预期 tool call

1. `get_current_context` 或 `list_slide_shapes`
2. `edit_slide_chart`
3. `edit_slide_master`
4. `verify_slides`

## 检查点

- [ ] 图表由可编辑形状组成
- [ ] 有标题、坐标轴、图例和数据形状
- [ ] 背景和装饰条生效
- [ ] 程序化验证没有明显重叠/越界
