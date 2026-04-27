# 09 · 富文本与 OOXML

覆盖 `list_slide_shapes`、`read_slide_text`、`edit_slide_text`、`edit_slide_xml`。

## 输入 prompt

```
先列出当前页所有形状。选一个有文字的形状，读取它的富文本 XML，然后把第一段文字改成“富文本已更新”，其中“已更新”加粗并改成蓝色。最后把当前页背景改成很浅的灰色。
```

## 预期 tool call

1. `list_slide_shapes`
2. `read_slide_text`
3. `edit_slide_text`
4. `edit_slide_xml`（`code` 读取 `ppt/slides/slide1.xml` 并修改背景 XML）
5. `verify_slides`

## 检查点

- [ ] `read_slide_text` 返回 `<a:p>` 片段
- [ ] `edit_slide_text` 使用 `ref` 或 `slideId + shapeId`
- [ ] `edit_slide_xml` 使用 `code:string`，修改后调用 `markDirty()` 并返回 before/after XML artifact
- [ ] 背景修改后仍能定位新 slide id
