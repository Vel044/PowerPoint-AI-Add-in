# 02 · 函数调用链（layered）

## 输入 prompt

```
画一个 Python 程序的函数调用链：main 调用 load_config 和 run_pipeline；run_pipeline 内部先 fetch_data 再 transform 最后 save。用分层布局，每层自上而下。
```

## 预期行为

- `create_diagram` 参数：
  - `layout: "layered"`
  - 节点带 `level` 字段（0: main, 1: load_config / run_pipeline, 2: fetch_data / transform / save）
  - 或不带 level 让工具按 edges 自动推断
  - edges 体现调用关系

## 检查点

- [ ] 分层明显，每层节点在同一 y 坐标
- [ ] 同层节点水平等间距
- [ ] 连接线方向全部自上而下（parent.bottom → child.top）
- [ ] run_pipeline 下方的三个子节点并排
- [ ] 无节点重叠、无交叉连接线错位
