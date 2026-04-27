# PowerPoint 工具 V1 实施说明

V1 目标是先把“读 -> 自由画图 -> 富文本/XML 修正 -> 验证”的闭环跑通，让模型能像 Claude PowerPoint 工具一样更自由地构建框图，同时保留现有稳定入口。

## V1 工具清单

| 工具 | 作用 |
| --- | --- |
| `list_slide_shapes` | 列出目标页全部形状，返回 `ref/id/name/type/text/bounds` |
| `read_slide_text` | 读取指定 shape 的 `<a:p>` 富文本 XML |
| `edit_slide_text` | 替换指定 shape 的富文本段落 XML |
| `draw_slide_shapes` | 批量自由绘制框图 |
| `edit_slide_xml` | 执行结构化 OOXML 操作 |
| `verify_slides` | 程序化检查几何问题 |
| `verify_slide_visual` | `review_slide` 的别名 |
| `duplicate_slide` | 复制指定幻灯片 |
| `export_deck_outline` | 导出演示文稿文本大纲 |
| `todo_write` | 展示 Agent 任务步骤 |

## 统一参数

所有按页工具支持：

| 参数 | 类型 | 说明 |
| --- | --- | --- |
| `slideId` | string | 精确指定幻灯片 |
| `slideIndex` | number | 0-based 索引 |
| `pageNumber` | number | 1-based 页码 |

优先级固定为：`slideId > pageNumber > slideIndex > 当前选中页`。

形状引用 `ref` 统一为 `slideId:shapeId`，应来自 `list_slide_shapes` 的返回值。

## `draw_slide_shapes`

适合复杂框图、调用链、泳道图、时间线。模型负责坐标、尺寸、颜色和层次，工具负责批量落图和真实连接器。

```json
{
  "pageNumber": 5,
  "shapes": [
    {
      "id": "entry",
      "type": "geometricShape",
      "shapeType": "rectangle",
      "text": "motors_bus._sync_write(...)",
      "left": 50,
      "top": 80,
      "width": 250,
      "height": 56,
      "fillColor": "#173B5F",
      "lineColor": "#173B5F",
      "textColor": "#FFFFFF",
      "fontSize": 12,
      "bold": true
    }
  ],
  "connectors": [
    {
      "from": "entry",
      "fromSide": "right",
      "to": "setup",
      "toSide": "left",
      "mode": "direct",
      "arrow": "end"
    }
  ]
}
```

## `edit_slide_xml`

只接受结构化操作，不开放任意 `code:string`。

V1 支持：

| `type` | 参数 |
| --- | --- |
| `insertShapeXml` | `xml` |
| `replaceShapeXml` | `shapeId`, `xml` |
| `deleteShapeXml` | `shapeId` |
| `patchConnector` | `connectorShapeId`, `fromShapeId`, `fromSide`, `toShapeId`, `toSide`, `start`, `end`, `connectorType`, `arrow`, `color`, `thickness` |
| `setSlideBackground` | `color` |

## 验收标准

- `list_slide_shapes({ pageNumber: 5 })` 能静默读取第 5 页。
- `draw_slide_shapes` 能复刻调用链框图，文字不明显溢出，连线不飞。
- `read_slide_text` 能返回指定 shape 的 `<a:p>` XML。
- `edit_slide_text` 能替换指定 shape 的富文本段落。
- `verify_slides` 能报告重叠和越界。
- `verify_slide_visual` 能保存截图并返回视觉审查结果。
- 原有 `create_diagram/connect_shapes/review_slide` 行为不回退。
