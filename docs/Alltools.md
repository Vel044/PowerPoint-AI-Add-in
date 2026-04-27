📐 list_slide_shapes
作用：列出某张幻灯片上所有形状的元数据，是编辑前的必要步骤。

参数	类型	必填	说明
slide_id	string	✅	幻灯片 ID，来自 slidesMetadata[].slideId
explanation	string	❌	操作说明（最多50字，显示在工具图标旁）
📖 read_slide_text
作用：读取某个形状的原始 OOXML <a:p> 段落 XML，保留粗体/颜色/字号等所有格式。

参数	类型	必填	说明
ref	string	✅	形状引用，格式如 "258:5"，只能来自 list_slide_shapes 的输出，不能手动构造
explanation	string	❌	操作说明
✏️ edit_slide_text
作用：用新的 OOXML <a:p> 段落替换某个形状的全部文本内容，保留形状的 <a:bodyPr> 和 <a:lstStyle>。

参数	类型	必填	说明
ref	string	✅	形状引用，来自 list_slide_shapes
code	string	✅	原始 OOXML <a:p> 段落 XML，替换形状的全部文本
explanation	string	❌	操作说明
🔧 edit_slide_xml
作用：直接操作幻灯片的原始 OOXML ZIP 包，适合复杂布局、多形状批量修改、绘制图形等底层操作。

参数	类型	必填	说明
slideId / slideIndex / pageNumber	string / number	❌	目标页；不传默认当前页
code	string	✅	async 函数体，接收作用域变量 `zip`、`markDirty`、`pptx`。普通绘图优先 `pptx.addShape(...); pptx.addConnector(...); pptx.save(); markDirty();`；`pptx.slideWidth/slideHeight` 是真实画布尺寸，越界 shape 会报错；可用 `style` 语义预设（entry/process/decision/success/danger/database/io/laneHeader/title 等）做合理区分
autosize_shape_ids	string[]	❌	修改了文本的形状的 cNvPr ID 列表，工具会重新应用 AutoSize 触发重排
explanation	string	❌	操作说明
📊 edit_slide_chart
作用：通过操作原始 OOXML 向幻灯片添加或修改图表（柱状图、折线图、饼图等）。

参数	类型	必填	说明
slide_id	string	✅	幻灯片 ID
code	string	✅	异步函数体，接收 { zip, markDirty }，内部操作图表 XML
autosize_shape_ids	string[]	❌	修改了文本的形状 ID，触发重排
explanation	string	❌	操作说明
🎨 edit_slide_master
作用：编辑幻灯片母版和版式，设置背景、主题色、字体方案、装饰形状等全局样式，影响所有幻灯片。

参数	类型	必填	说明
code	string	✅	异步函数体，接收 { zip, markDirty }，zip 包含完整 PPTX 结构（ppt/slideMasters/、ppt/slideLayouts/ 等）
explanation	string	❌	操作说明
💻 execute_office_js
作用：在 PowerPoint.run() 内执行任意 Office.js 代码，适合移动/缩放形状、读取属性、添加文本框/表格、插入图片等操作。

参数	类型	必填	说明
code	string	✅	异步函数体，接收 context: PowerPoint.RequestContext，必须调用 context.sync() 执行批量操作，返回 JSON 可序列化结果
explanation	string	❌	操作说明
📋 verify_slides
作用：检查幻灯片的几何问题，包括形状重叠、越界、空占位符、低对比度文字，返回每张幻灯片的结构化结果。

参数	类型	必填	说明
from_slide	number	❌	起始幻灯片索引（0-based，含），默认为 0
to_slide	number	❌	结束幻灯片索引（0-based，含），默认为最后一张
explanation	string	❌	操作说明
👁️ verify_slide_visual
作用：截图后发给独立 AI 审阅（无对话记忆，消除确认偏差），客观反馈视觉问题。仅用于验证完成的工作，不用于编辑前检查。

参数	类型	必填	说明
slide_id	string	✅	幻灯片 ID
explanation	string	❌	操作说明
📄 duplicate_slide
作用：复制指定幻灯片，副本插入在原片紧后面。

参数	类型	必填	说明
slide_id	string	✅	要复制的幻灯片 ID
explanation	string	❌	操作说明
🖼️ copy_image_between_slides
作用：将一张幻灯片上的图片形状复制到另一张，自动处理媒体文件注册和关系绑定，保留裁剪/滤镜/替代文字。

参数	类型	必填	说明
from_slide_id	string	✅	来源幻灯片 ID
from_shape_id	string	✅	来源图片形状的 cNvPr id，来自 list_slide_shapes 的 id 字段
to_slide_id	string	✅	目标幻灯片 ID
x	number	❌	目标位置左边距（pt），省略则保留原位置
y	number	❌	目标位置上边距（pt），省略则保留原位置
width	number	❌	目标宽度（pt），省略则保留原尺寸
height	number	❌	目标高度（pt），省略则保留原尺寸
explanation	string	❌	操作说明
💾 store_blob
作用：从 Files API 把文件字节拉取到浏览器 blob 存储，供后续 execute_office_js 使用（如从 .pptx 模板插入幻灯片）。

参数	类型	必填	说明
file_id	string	✅	Files API 文件 ID
blob_name	string	✅	存储键名，之后用 blobs.getBase64(name) 等方法访问
📤 export_deck_outline
作用：把整个演示文稿的大纲（每张幻灯片标题 + 所有文本框内容）导出为 deck-outline.json，写入共享文件系统供其他 Agent 使用。

参数	类型	必填	说明
slides	number[] 或 "all"	❌	要导出的幻灯片位置（1-based），默认 "all" 全部导出
explanation	string	❌	操作说明
🔍 search_icons
作用：在微软图标库中搜索矢量图标，返回 id、描述、风格、搜索评分等元数据，再用 insert_icon 插入。

参数	类型	必填	说明
query	string	✅	搜索词，建议用单个名词（如 "rocket"），多词组合通常返回零结果
top	number	❌	最多返回条数，默认 5，最大 20
explanation	string	❌	操作说明
🏷️ insert_icon
作用：把 search_icons 找到的图标以矢量 SVG（含 PNG 备用）形式插入幻灯片，行为与 PowerPoint 原生"插入图标"完全一致。

参数	类型	必填	说明
icon_id	string	✅	图标 ID，来自 search_icons 结果
slide_id	string	✅	目标幻灯片 ID
x	number	❌	左边距（pt），省略则居中
y	number	❌	上边距（pt），省略则居中
width	number	❌	宽度（pt），默认 72（1英寸）
height	number	❌	高度（pt），默认 72
color	string	❌	统一填充颜色，十六进制如 "#FF5733"
description	string	❌	无障碍替代文字
explanation	string	❌	操作说明
🌐 web_search
作用：搜索互联网，获取实时或最新信息。

参数	类型	必填	说明
query	string	✅	搜索关键词，建议简短精准（1-6个词）
🤝 get_connected_agents
作用：列出当前连接的其他 Agent（如 Excel、Word）及其 ID、标签、能力描述。无参数。

📨 send_message
作用：向另一个 Agent 发送消息，委托它在对应 Office 应用里完成操作（异步，发出即返回）。

参数	类型	必填	说明
agent_id	string	✅	目标 Agent 的 ID，来自 get_connected_agents
message	string	✅	消息内容，描述要做什么
🔄 refresh_mcp_connectors
作用：重新拉取用户在 claude.ai 上配置的 MCP 连接器，解决连接失败或新增连接器未显示的问题。无参数。

🖥️ bash
作用：在共享工作区运行受限 shell 命令，读取其他 Agent 的文件或对话记录（只读沙箱，不能联网）。

参数	类型	必填	说明
command	string	✅	Shell 命令，仅支持 cat、head、tail、grep、jq、ls、find 等列表内命令，不支持 awk、sed、curl 等
📝 update_instructions
作用：用查找替换方式修改你的长期偏好指令（字体、配色、布局习惯等），跨对话持久保存。

参数	类型	必填	说明
operations	object[]	✅	操作数组，每项包含 old_text（要找的文本，空字符串=追加）和 new_text（替换内容，空字符串=删除）
⚙️ update_setting
作用：开关功能设置，需用户确认。

参数	类型	必填	说明
setting	enum	✅	"cross_file_access"（跨文件访问）、"web_search"（网络搜索）、"mcp_connectors"（MCP连接器）
value	boolean	❌	省略则读取当前状态，填写则尝试修改
📚 read_skill
作用：读取某个预置技能的完整指令文件，执行技能前必须先调用。

参数	类型	必填	说明
skill_name	string	✅	技能名称，来自 <available_skills> 列表，如 "competitive-analysis"
🛠️ create_skill
作用：提议创建新技能，生成 SKILL.md 草稿作为审核卡片展示给用户，用户点击 Apply 后才保存。

参数	类型	必填	说明
name	string	✅	技能名，slug 格式（小写字母、数字、连字符），最长64字符，如 "summary-slide"
description	string	✅	一句话描述技能用途，显示在技能列表和斜杠菜单中
instructions	string	✅	SKILL.md 正文（frontmatter 之后的部分），用 Markdown 标题组织工作流
❓ ask_user_question
作用：向用户展示可点击的选项卡片，代替在对话里打字回答，适合收集偏好或决策。

参数	类型	必填	说明
questions	object[]	✅	问题数组，每项包含：question（问题文本）、header（卡片标题）、options（选项数组，每项有 label 和 description，2-4个）、multiSelect（是否多选）
✅ todo_write
作用：创建或更新多步任务列表，在侧边栏 Steps 面板显示进度。每次调用都是全量替换。

参数	类型	必填	说明
todos	object[]	✅	完整任务列表，每项包含：content（任务描述）、activeForm（进行中显示的文字）、status（"pending" / "in_progress" / "completed"）
✂️ context_snip
作用：标记可压缩的对话区间，在上下文窗口压力到达约60%时自动压缩（内部管理用）。

参数	类型	必填	说明
from_id	string	✅	区间起始用户消息的 [id:xxxxxx] 标签，完整复制
to_id	string	✅	区间结束用户消息的 [id:xxxxxx] 标签
summary	string	✅	压缩后留存的摘要，应包含完成的工作、关键数值和未完成事项
🔎 retrieve_snipped
作用：从已压缩的区间存档中检索特定内容，避免重新调用原始工具。

参数	类型	必填	说明
from_id	string	✅	已应用的 snip 的 from_id
search	string	❌	可选子字符串过滤，只返回包含该词的记录并截取匹配窗口
max_chars	number	❌	返回内容上限，默认 4000，范围 500-20000
🔍 tool_search_tool_bm25
作用：用 BM25 算法在所有可用工具中搜索，找到最匹配某个需求描述的工具。

参数	类型	必填	说明
query	string	✅	自然语言搜索词，最长500字符，支持多词自动分词和词干处理
limit	number	❌	最多返回条数，默认5，最大10000
总计 28 个工具，涵盖读取→编辑→验证→协作→偏好管理的完整工作流。
