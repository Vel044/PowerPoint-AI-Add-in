# 07 · 截图复现 · 舵机写总线调用链

对比改造前后的效果。改造前的截图是 17 节点 14 连线，框子重叠、线没贴边。

## 输入 prompt

```
把下面这段调用链画成框图，用分层布局，每层自上而下：

motors_bus._sync_write(addr=42, length=2, ids_values={1:1891,...,6:34})
  └─ _setup_sync_writer(ids_values, addr=42, length=2)
       ├─ sync_writer.clearParam()
       ├─ sync_writer.start_addr = 42
       ├─ sync_writer.data_length = 2
       └─ for id_, value in ids_values:
            ├─ _serialize_data(value=1891) → [0x83, 0x07]
            └─ sync_writer.addParam(id_=1, data=[0x83,0x07])
                 └─ makeParam()
  └─ sync_writer.txPacket()
       └─ ph.syncWriteTxOnly(port, start_addr, data_length, param, len)
            └─ txRxPacket(port, txpacket, rxpacket=None)
                 └─ txPacket(port, txpacket)
                      ├─ port.clearPort()
                      └─ port.writePort(txpacket)    ★ 写入物理帧 FF FF FE 16 83 2A 02 [id+data×6] CS
```

## 预期行为

- 一次 `create_diagram({ layout: "layered" or "tree", ... })`
- 大约 15-17 个节点，带层级
- edges 体现父子调用关系
- 节点 shape：菱形用于 `for` 循环，其他用矩形，终点用 `flowChartTerminator`

## 检查点（对比改造前）

| 维度 | 改造前（截图） | 改造后预期 |
|---|---|---|
| 节点重叠 | 有，大量 | 无 |
| 箭头贴边 | 否，随意连 | 是，贴边中点 |
| 节点尺寸 | 不一致 | 完全一致（默认 160×60）|
| 布局方向 | 混乱 | 明显分层 |
| 箭头头 | 无（全裸线）| 正交方向有箭头头 |

## 验收

如果改造后依然重叠或箭头不贴边，回头查：
- `create_diagram` 的 canvas 是否太小装不下 17 节点 → 考虑缩小默认 node 尺寸或扩大 canvas
- 是否有同层节点过多（>5）导致横向塞不下 → 可以让 AI 拆成两张幻灯片
