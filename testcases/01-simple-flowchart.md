# 01 · 登录流程图（vertical）

## 输入 prompt

```
在当前幻灯片画一个登录流程图：开始 → 输入用户名密码 → 判断是否验证通过 → 通过则进入首页、不通过则提示错误 → 结束。
```

## 预期行为

- AI 调用 `get_current_context`
- 调用 Claude 风格 `edit_slide_xml({ code })`，code 应使用 `pptx.openSlide/addShape/addConnector/save` 批量插入 6 个节点和 6 条真实连接器：
  - 起止用 terminator 或圆角矩形风格、判断用 diamond、过程用 rectangle
  - 连接器包含开始到输入、输入到判断、判断到两条分支、两条分支汇合到结束
- 画完后调用 `verify_slides` 和截图检查

## 检查点

- [ ] 节点垂直等间距排列，水平居中
- [ ] 同类形状尺寸一致
- [ ] 每条连接线主体是真实 PowerPoint connector：同轴为 Straight，分支/汇合为 bentConnector3 肘形连接器
- [ ] 箭头为 PowerPoint 原生 tailEnd 箭头，不出现三角形形状模拟箭头
- [ ] 菱形的两条分支（通过/不通过）连接方向合理
- [ ] AI 最后有 2-4 句中文总结
