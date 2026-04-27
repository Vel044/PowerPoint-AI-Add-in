# 07 · 截图复现 · 舵机写总线调用链

对比改造前后的效果。改造前的截图是 17 节点 14 连线，框子重叠、线没贴边。

## 输入 prompt

```
把下面这段调用链画成框图，从左到右是层次，从上而下是时间，一个框是一个文件：

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

- 优先一次结构化 `edit_slide_xml`，由模型自由决定每个框的坐标、尺寸、颜色、字体和连接器
- 可以用标题带标注“左至右：层次 / 上至下：时间”
- edges/连接器体现父子调用关系
- `for` 循环、序列化、物理帧写入可以用不同颜色分组
- 画完必须 `verify_slides`，再用 `verify_slide_visual` 或 `review_slide` 做截图检查

## 检查点（对比改造前）

| 维度       | 改造前（截图） | 改造后预期                                                  |
| ---------- | -------------- | ----------------------------------------------------------- |
| 节点重叠   | 有，大量       | 无                                                          |
| 连接线贴边 | 否，随意连     | 是，贴边中点                                                |
| 节点尺寸   | 不一致         | 同类节点尺寸一致，长文本框可按内容加宽                     |
| 布局方向   | 混乱           | 左到右是层次，上到下是时间                                  |
| 箭头头     | 无（全裸线）   | 使用 PowerPoint 原生 tailEnd 箭头，不出现三角形形状模拟箭头 |

## 验收

如果改造后依然重叠或连接线不贴边，回头查：
- `edit_slide_xml` 是否给了足够大的画布和合适的节点宽度
- 是否有同层节点过多导致横向塞不下 → 可以让 AI 拆成两张幻灯片
