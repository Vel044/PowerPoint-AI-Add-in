# 13 · 外部桥接占位

覆盖 `get_connected_agents`、`send_message`、`refresh_mcp_connectors`。

## 输入 prompt

```
列出可连接的其他 Office Agent，然后发消息给 Word Agent 让它总结当前 PPT，最后刷新 MCP 连接器。
```

## 预期 tool call

1. `get_connected_agents`
2. 可选 `send_message`
3. 可选 `refresh_mcp_connectors`

## 检查点

- [ ] 工具返回“当前本地 add-in 未配置多 Agent/MCP 桥接协议”
- [ ] Agent 不把 unsupported 结果当成成功协作
- [ ] 最终回复应该解释当前不能跨 Office 委托，而不是编造结果
