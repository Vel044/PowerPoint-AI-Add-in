# 11 · 图片、图标、资源

覆盖 `store_blob`、`copy_image_between_slides`、`search_icons`、`insert_icon`。

## 输入 prompt

```
保存一个文本资源，名字叫 demo-note，内容是“hello blob”。然后搜索 database 图标，把找到的图标插入当前页左上角，颜色 #2F5597，尺寸 64x64。
```

## 预期 tool call

1. `store_blob`
2. `search_icons`
3. `insert_icon`
4. `verify_slides`

## 图片复制补充 prompt

先准备一页里有图片形状，再执行：

```
列出第一页和第二页的形状，把第一页的图片复制到第二页，放在 left=100 top=100 width=200 height=120。
```

## 图片复制预期

1. `list_slide_shapes({ pageNumber: 1 })`
2. `list_slide_shapes({ pageNumber: 2 })`
3. `copy_image_between_slides`

## 检查点

- [ ] `store_blob` 返回 blob 名、mime 和大小
- [ ] `search_icons` 返回内置图标结果
- [ ] `insert_icon` 插入的是可见图片形状
- [ ] 非图片来源调用 `copy_image_between_slides` 时返回明确错误
