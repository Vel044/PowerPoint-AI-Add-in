# 12 · 设置、技能、搜索

覆盖 `web_search`、`update_instructions`、`update_setting`、`read_skill`、`create_skill`。

## 输入 prompt

```
搜索 PowerPoint Office.js addImage 的相关资料。然后把长期偏好追加一条：“复杂框图优先使用深色标题带”。开启 review_required 设置。创建一个名为 diagram-style 的技能草稿，描述是“统一框图风格”，内容包含三条：先读上下文、再自由绘图、最后验证。创建后读取这个技能。
```

## 预期 tool call

1. `web_search`
2. `update_instructions`
3. `update_setting`
4. `create_skill`
5. `read_skill`

## 检查点

- [ ] `web_search` 可用时返回标题/链接/摘要；不可用时返回明确错误
- [ ] 长期偏好写入 localStorage，并影响后续 system prompt
- [ ] `update_setting` 能写入和读取
- [ ] `skills/diagram-style.md` 被创建且可读取
