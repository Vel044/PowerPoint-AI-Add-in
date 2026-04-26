# 测试用例

绘图能力的手动回归测试。每个文件是一个场景，里面是一段可以直接复制粘贴到 PowerPoint 右侧 AI 面板输入框的中文 prompt，以及预期效果和检查点。

## 怎么用

1. `npm run dev` 启动本地开发服务器
2. PowerPoint 打开任一 .pptx，加载插件
3. 打开一个 testcase 文件，把 **输入 prompt** 复制到输入框，发送
4. 对照 **预期** 和 **检查点** 验收

## 文件索引

| 文件 | 场景 | 主要验证 |
|---|---|---|
| [01-simple-flowchart.md](01-simple-flowchart.md) | 登录流程图（vertical） | `create_diagram` 竖排布局、判断菱形分叉、连接线贴边中点 |
| [02-call-chain.md](02-call-chain.md) | 函数调用链（layered） | `create_diagram` 分层布局，3 层节点 |
| [03-architecture.md](03-architecture.md) | 系统架构图（tree） | `create_diagram` 树状，按 edges 自动推断层级 |
| [04-connect-shapes.md](04-connect-shapes.md) | 已有形状连线 | `connect_shapes` 按 fromSide/toSide 贴边连接 |
| [05-mixed-edit.md](05-mixed-edit.md) | 在已有图上加节点 | `add_geometric_shape` + `connect_shapes` 组合修改 |
| [06-regression.md](06-regression.md) | 基础功能回归 | 文本框、页内删改、上下文几何信息、截图留档 |
| [07-hard-case.md](07-hard-case.md) | 截图复现场景 | 对比改造前后，motors_bus._sync_write 调用链 |

## 验收关键词（用于对比改造前效果）

- **节点不重叠**：任意两个矩形边框不相交
- **连接线贴边中点**：连接线端点精确指向框四条边的中点（上/下/左/右），没有歪斜或错位
- **同类节点同尺寸**：所有 rectangle 一样宽高，所有 diamond 一样大
- **连线方向正确**：流程方向与语义一致（开始→结束自上而下）
- **布局居中**：整个图水平居中于幻灯片画布
- **页内形状操作**：删除/修改形状必须基于当前页 `slideId + shapeId`，不跨页搜索
- **截图可复盘**：`review_slide` 结果中包含本地 `debug-artifacts/review-slide/` 截图路径
