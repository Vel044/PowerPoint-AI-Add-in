# 14 · 主题色板注入

覆盖 `readThemePalette` 与 `get_current_context` 的 `themePalette`/`isDefaultTheme` 字段，验证模型在自定义品牌主题下能用上正确的 hex。

## 准备

- 用例 A：用 PowerPoint 默认 Office 主题打开任意空白演示文稿
- 用例 B：打开一个带自定义品牌主题的 .pptx（比如自定义了 accent1 为品牌主色）

## 输入 prompt

**用例 A（默认主题）**：
```
读取当前演示文稿上下文，告诉我有没有自定义主题。
```

**用例 B（自定义主题）**：
```
读取当前演示文稿，然后在中间画一个矩形，颜色用模板的主色调。
```

## 预期 tool call

用例 A：
1. `get_current_context`

用例 B：
1. `get_current_context`
2. `add_geometric_shape`，`fillColor` 字段取自 themePalette.accent1

## 检查点

### 用例 A（默认主题）
- [ ] `get_current_context` 返回 JSON 包含 `themePalette`/`isDefaultTheme`/`themeName` 三个字段
- [ ] `isDefaultTheme: true`
- [ ] `themeName: "Office"`（或 null）
- [ ] 模型不在 OOXML 里硬编码 hex，如需画形状则**省略颜色参数**让其继承主题
- [ ] 浏览器 DevTools 控制台能看到首次 onSend 触发了一次 export-as-base64；后续 onSend 不再触发（缓存生效）

### 用例 B（自定义主题）
- [ ] `themePalette` 包含 dk1/lt1/accent1-6 等键，值为 6 位 hex 字符串（不带 `#`）
- [ ] `isDefaultTheme: false`
- [ ] `themeName` 不为 "Office"
- [ ] 模型主动从 themePalette 取 accent1 的 hex 值，传给 `add_geometric_shape` 的 `fillColor` 参数
- [ ] 形状颜色与模板品牌主色一致

### 缓存行为
- [ ] 多次 onSend 不会重复 export+JSZip（除非显式调用 `clearThemeCache`）
- [ ] 用户切换演示文稿后，缓存可能仍是旧主题（已知限制，需要重启 add-in 或手动清缓存）

## 异常处理

- [ ] PowerPointApi 1.8 不可用时（`exportAsBase64` 失败），`themePalette: null`、`isDefaultTheme: true`，不影响 `get_current_context` 其余字段
- [ ] 演示文稿一张幻灯片都没有时，`themePalette: null`，不抛错
- [ ] theme1.xml 缺失或解析失败时，graceful fallback 到 `themePalette: null`
