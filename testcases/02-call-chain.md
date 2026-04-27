# 02 · 函数调用链（XML 自由框图）

## 输入 prompt

```
画一个 Python 程序的函数调用链：main 调用 load_config 和 run_pipeline；run_pipeline 内部先 fetch_data 再 transform 最后 save。请画成框图：从左到右是层次，从上到下是时间，每个函数一个框，连接线带箭头。
```

## 预期行为

- 复杂调用链优先结构化 `edit_slide_xml`
- AI 自己决定 OOXML 中的坐标、尺寸、颜色和连接线
- 画完后调用 `verify_slides`；视觉检查任务还应调用 `verify_slide_visual`

## 检查点

- [ ] 从左到右层次明显
- [ ] 从上到下时间顺序明显
- [ ] 连接线方向体现调用关系
- [ ] run_pipeline 下方的三个子节点并排
- [ ] 无节点重叠、无交叉连接线错位
