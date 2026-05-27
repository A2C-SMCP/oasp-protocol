# 变更日志

本文档记录 OASP 协议的所有重要变更。

格式基于 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.0.0/)，
版本号遵循 [语义化版本](https://semver.org/lang/zh-CN/)。

---

## [Unreleased]

_暂无未发布变更。_

---

## [0.3.0] - 2026-05-26

### 架构转变：Python MCP Server 从「纯中转」升级为「具备生产能力」

0.3.0 是 OASP 的一次**模型转变**，也是本次 MINOR 升级（0.2 → 0.3）的根本原因——不是单纯加事件：

- **此前（≤0.2.x）**：Python MCP Server 仅做消息中转，**所有动作都在 AddIn（Office.js）实现**；AddIn 未连接即无法操作文档。
- **从 0.3.0 起**：Server **也具备生产能力**——当 AddIn 未打开、或 AddIn 能力无法实现某操作时，Server 直接介入生产（借助 `python-pptx` 等 OOXML 工具）；**未来 Server 将在 AddIn 完全未连接的情况下实现全部工具操作**。

本版所有变更都服务于这一转变、互为支撑：

- **剥离实现技术 / Normative-Informative 分层** → 让同一套协议可被 Server 或 AddIn **任一端**满足，是「双端皆可生产」的地基；
- **图表双路径线缆可观测项（`3016` / `elementId` 不透明性 / 客户端路径 Informative）** → Server-生产 与 AddIn-生产 之间的**路由**与反应式降级；
- **通用幻灯片 OOXML 搬运事件（`ppt:get:slideOoxml` / `ppt:insert:slidesOoxml`）** → AddIn 开着时也能把整页 OOXML 交给 Server 生产再回插；
- **协议版本握手** → 既然**两端都生产**，Server ↔ AddIn 的版本错位代价更高，故在连接期强制校验版本。

### 新增（协议版本握手）

为避免新旧 Server / AddIn 协议错位，新增连接握手阶段的协议版本校验机制。AddIn 在 `auth` 中声明 `oaspVersion`，Server 在 `connect` handler 校验兼容性（**版本先于业务参数**），不兼容即拒绝连接。

| 变更类型 | 位置 | 说明 |
|----------|------|------|
| 新增（握手参数） | `connection.md` | `auth.oaspVersion`（必填，SemVer）；新增「协议版本握手」节：校验时机/顺序（**oaspVersion 先行**）、`HandshakeRejection` 扁平拒绝结构、Server 校验示例、AddIn 收到拒绝后 MUST 主动断开 |
| 新增（连接确认字段） | `connection.md` | `connection:established` 增加 `serverVersion` 字段（仅诊断用） |
| 新增（版本治理） | `conventions.md` | 「版本兼容」升级为「协议版本与兼容性」：MAJOR/MINOR/PATCH 语义与触发条件、`is_compatible` 判定规则（v0.x 严格 MAJOR.MINOR；v1.0+ Server 向后兼容）、Python 参考实现、载体选型理由、**非目标**与「**版本握手 ≠ 运行时能力**」正交说明 |
| 新增错误码 | `error-handling.md` | `2006 PROTOCOL_VERSION_MISMATCH`（连接/认证段）；复用 `2003 HANDSHAKE_FAILED` 处理缺失/非法 `oaspVersion` |

**兼容性**: 引入新的**必填**握手参数 `oaspVersion`——采用本机制的 Server 会拒绝不声明版本的旧 AddIn，需 Server 与 AddIn **协同升级**。属 v0.x 阶段的 MINOR 级变更（v0.x 阶段 MINOR 可含破坏性）。

**设计来源**: 对齐 A2C-SMCP `versioning.md` 的版本握手规范，按 OASP 两方单连接模型简化——校验放 `connect` handler（非 HTTP 中间件），载体放 `auth`（非 URL query），拒绝走 `ConnectionRefusedError`（非 HTTP 400）。

### 变更（规范澄清，非破坏性）

**将实现技术规定从规范层剥离，确立 Normative / Informative 分层治理原则**

线缆契约（事件名、请求/响应形状、`ChartData` / `ChartType` / `CategoricalSeries` / `ScatterSeries` 判别联合、错误码）**全部保持不变**；本次仅调整规范文本，使 `/ppt` 图表事件可被服务端或客户端任意路径实现。

| 变更类型 | 位置 | 说明 |
|----------|------|------|
| 移除（实现泄漏） | `events-ppt.md` chart 事件顶部 admonition | 删除「不经 Add-In Office.js 路径」「由 Server 端使用 OOXML 工具（如 `python-pptx`）修改 .pptx」等执行位置/技术选型表述 |
| 重构 | `events-ppt.md` chart 事件 | 拆分为两块：① 实现中立的**线缆契约**（请求-响应语义、`success` 后变更对 `ppt:get:chart` 可见、并发顺序约定、显式声明"协议不规定由哪一端/何种技术实现"）；② `!!! info "实现提示（非规范）"` 收纳仅服务端离线路径相关的副作用（延迟 >1s、调用前 `save()`、完成后重载文档、离线写入串行化） |
| 中立化 | `events-ppt.md` / `error-handling.md` 错误码触发条件 | `3004 OPERATION_FAILED` 描述由「OOXML 写入失败」改为「图表写入失败」；`3003 DOCUMENT_READ_ONLY` 由「文档为只读模式」改为「目标文档不可写入」，描述线缆可观测状态而非特定实现 |
| 移除（实现泄漏） | `events-ppt.md` 事件列表表格 | 三个 chart 事件行去掉「（Server OOXML 离线处理）」括注 |
| 新增（治理原则） | `conventions.md` | 新增「规范分层（Normative / Informative）」章节：规范层只含线缆可观测内容（事件名/方向、字段形状与语义、错误码及实现中立触发条件、顺序/幂等/可见性保证）；用什么库、服务端还是客户端、性能特征、特定实现副作用一律归为非规范实现提示 |
| 新增（设计原则） | `error-handling.md` | 错误码设计原则补充「与实现无关」：触发条件描述线缆可观测状态而非实现技术/执行位置 |

**动机**: 消费方 office4ai（OF4AI-21）需按文档连接状态在服务端 OOXML 路径与客户端 Office.js 路径之间路由——同一套事件/payload 两种方式均可实现，接口形状无需改动。原 v0.2.0 文本把「用 `python-pptx` / 服务端离线 / 不走 Office.js」写进规范，等于协议层替消费方做了实现决策，阻碍多路径复用。

**兼容性**: 纯文档/规范澄清，无任何线缆契约变更；对已部署消费方零影响。`/ppt` 整体仍处 📋 Draft 阶段。

**相关工单**: [OF4AI-21](https://turingfocus.atlassian.net/browse/OF4AI-21)（PPT 图表双路径路由）

### 新增（图表双路径线缆可观测项）

在上述减法基础上，为支撑 OF4AI-21 双路径路由追加少量线缆可观测项；事件名、请求/响应形状、`ChartData` 判别联合、现有错误码仍不变。

| 变更类型 | 位置 | 说明 |
|----------|------|------|
| 新增错误码 | `error-handling.md` + `ppt:insert:chart` / `ppt:get:chart` / `ppt:update:chart` 错误码表 | `3016 API_NOT_SUPPORTED`：目标操作在当前客户端/平台不可用（如所需 PowerPointApi requirement set 不满足、宿主不支持）。语义实现中立，供调用方**反应式降级**到另一路径或提示用户 |
| 新增（Normative） | `data-structures.md` 新增「元素标识符不透明性」（锚点 `#element-id-opacity`） | 明确 `SlideElement.id` / chart `elementId` 为服务端分配的**不透明字符串**，消费方不得解析其结构。防御整页 round-trip 重置 native id 的实现细节泄漏到协议 |
| 新增（Informative） | `events-ppt.md` chart 事件 `!!! info` 注记 | 补「路径 B — 客户端 Office.js 整页 round-trip」实现提示（`exportAsBase64` → 服务端改 → `insertSlidesFromBase64` 整页回插），并列其非规范副作用：选区/滚动重置、撤销非原子、母版累积、requirement set 门槛（新页 1.2 / 现有页·get·update 1.8，皆不满足回退路径 A） |

**确认不新增**: `CHART_DATA_NOT_READABLE`——`exportAsBase64` + 服务端解析可读回任意图表数据（含未保存态、含非本系统生成图表），仅平台 <1.8 读不到，已归入 `3016 API_NOT_SUPPORTED`。

**能力声明**: 本轮采用**反应式**路由（直接尝试 → 失败返回 `3016` / `3003` 即保证正确性），不新增事件。主动式 `ppt:capabilities`（握手声明 `isSetSupported` 矩阵）作为后续优化单独立项，不阻塞本轮。

### 新增（通用幻灯片 OOXML 搬运事件）

为支撑 OF4AI-21 双路径的路径 B（客户端整页 round-trip），`/ppt`「幻灯片管理类」新增 2 个 **Server→AddIn 请求-响应**事件（📋 Draft）。图表无关的低层传输原语，未来"服务端编辑打开中文档"类能力可复用。纯加法、向后兼容。

| 变更类型 | 事件 / 类型 | 说明 |
|----------|-------------|------|
| 新增 | `ppt:get:slideOoxml` | 导出指定页当前 OOXML（含未保存态）为单页 `.pptx` base64；响应含不透明 `slideId` |
| 新增 | `ppt:insert:slidesOoxml` | 应用 OOXML 页包；可选 `replaceSlideId` + `finalSlideIndex` 把「插入→删旧→复位」做成**尽力顺序复合（非原子）round-trip**；响应可回报命名元素（如图表）的最终几何 |
| 扩展 | `data-structures.md` 标识符不透明性 | `#element-id-opacity` 节从"元素标识符"扩展到涵盖 `slideId` |
| 扩展 | 错误码 `3010 ELEMENT_NOT_FOUND` | 描述扩展为「元素或幻灯片未找到」，供 `replaceSlideId` 复用，不新增错误码 |
| 文档基建 | `mkdocs.yml` | 启用 `attr_list`（支撑 `#element-id-opacity` 等锚点；顺带修复既有 `#slideelementelementextensions` 失效链接） |

**命名**: 取 `get:slideOoxml` / `insert:slidesOoxml` 中立对称对——按内容格式（OOXML）命名，不把 `base64` 传输编码或 Office.js 方法名写进规范事件名（遵循本周期确立的事件名实现中立原则）。

**消费方验证（office-editor4ai cross-ask，无 P0）**: 据 Add-In 实现可行性反馈做 round-N 修订——
- `insert:slidesOoxml` 措辞由"原子"放宽为**"尽力顺序复合（非原子，无回滚保证）"**：Office.js 无事务/回滚；部分失败时 `error.details` 给出 `{ stage, partiallyApplied, createdSlideId }` 供服务端对账补偿。
- 命名元素回报（`elements[]`）的定位机制下沉为 Informative：**主路径** `cNvPr/@name` + `shape.name`（`@name` 穿越 round-trip 的存活性已由 [office-editor4ai#34](https://github.com/JIAQIA/office-editor4ai/issues/34) 实测确认——Mac/单次：存活、几何可读、`masterLeak: 0`），`customXmlParts` 注册表降为**防御后备**；线缆契约不依赖具体机制。另记：图表可能位于占位符内（`type==="Placeholder"` + `containedType==="Chart"`），按不透明 `elementId` 定位不受影响，但勿用 `type==="Chart"` 过滤。
- `formatting` 保持 camelCase（与其它 ppt 事件一致），Add-In 内部映射到 Office.js PascalCase 枚举；`finalSlideIndex` 任意复位需 1.8、原位替换 1.2 即可；`3016` 经 `isSetSupported` 预检主动返回。

> Add-In spike（[office-editor4ai#34](https://github.com/JIAQIA/office-editor4ai/issues/34)）已确认 `@name` 穿越 round-trip 存活（Mac / 单次 round-trip：存活、几何可读、无母版累积），**主路径成立**；`customXmlParts` 注册表作防御后备。边界：Web/Windows 与多次连续 round-trip 的母版累积建议后续各补一次抽测。

---

## [0.2.0] - 2026-04-30

### 新增

**`/ppt` 命名空间补齐图表（Chart）能力（3 个 Draft 事件 + 1 个跨命名空间数据结构组）**

| 变更类型 | 事件 / 类型 | 说明 |
|----------|-------------|------|
| 新增 | `ppt:insert:chart` | 📋 Draft，按 `ChartData` discriminated union 在指定幻灯片插入图表 |
| 新增 | `ppt:get:chart` | 📋 Draft，按 elementId 读取图表完整数据，响应 `chart` 字段为同一 union |
| 新增 | `ppt:update:chart` | 📋 Draft，按 elementId 更新图表数据/类型/标题/展示选项；`chart.chartType` 为必需 discriminator |
| 提升为跨命名空间 | `ChartType` | 从 `data-structures.md` 的 `Excel 相关` 段移至新建的 `图表相关` 段，并扩展为 10 个 Office.js `Excel.ChartType` 对齐值（新增 `ColumnClustered` / `ColumnStacked` / `BarClustered` / `LineMarkers` / `Radar`，原 `Column` / `Bar` 高层值移除以避免歧义）。`ChartType` 进一步细分为 `CategoricalChartType`（9 个）+ `ScatterChartType`（"Scatter"）。破坏性变更，但 Excel 命名空间整体仍处 📋 Draft 阶段，影响面可控 |
| 新增 | `CategoricalSeries` / `ScatterSeries` / `ScatterPoint` | 数据结构，分别承载分类型图表的数据系列（name / values / color）与散点型图表的数据系列（name / points / color）。取代原始 `ChartSeries` 单一定义 |
| 新增 | `ChartData` discriminated union | `ChartData = CategoricalChartData \| ScatterChartData`，由 `chartType` 字段判别。Categorical 含 `categories: string[]` + `series: CategoricalSeries[]`；Scatter 不含 `categories`，X 由 `series[].points[].x` 提供。供 PPT / Excel 共用 |
| 新增 | 错误码 `3015 INVALID_CHART_DATA` | 多用途：categorical 维度不一致 / scatter `points` 为空或含非法值 / 跨 variant 切换未补齐 `series` |
| 新增（说明） | `ppt:delete:element` 描述 | 追加 `亦适用于图表（Chart）元素删除` admonition；不新增独立 `ppt:delete:chart`，复用通用删除入口 |

**动机**: PPT 是图表高频使用场景（业绩报告、产品方案）。当前 OASP `/ppt` 命名空间已能插入文本框、形状、图片、表格，**唯独不能插入图表**——`Chart` 仅在 `SlideElementType` 中作为枚举出现并可被 `ppt:get:slideElements.includeCharts` 过滤，但没有任何对图表本身的增改查事件。AI 体验上只能让用户手动插入图表后再让 AI 改文字，断裂明显。

**设计取舍**:

- **Server-handled OOXML 路径**：PowerPoint Office.js 当前不暴露图表创建与数据更新接口（参见 [office-js#5463](https://github.com/OfficeDev/office-js/issues/5463)），本批 chart 类事件由 OASP Server 端通过 OOXML 离线工具（如 `python-pptx`）直接修改 .pptx 文件后再让 Add-In 重新加载文档。事件文档顶部 admonition 明确标注延迟 >1s、要求 Add-In 调用前 `save()`、并发需串行化等副作用约束
- **`ChartData` 采用 discriminated union（`chartType` 为 discriminator）**：`Scatter` 与分类型图表的数据形状本质不同（X 是连续数值 vs 离散标签）。如果硬塞同一 schema（让 `categories` 在 Scatter 时存数字字符串），LLM 在 MCP 工具的 JSON Schema 里看不到这个隐式约束，调用错误率高。Discriminated union 让 LLM 一选定 `chartType`，schema 就限定可填字段；JSON Schema 用 `oneOf` + `discriminator: { propertyName: "chartType" }` 落地，Pydantic / FastMCP 用 `Annotated[Union[...], Field(discriminator="chartType")]`
- **`update:chart` 强制要求 `chartType`**：让 schema 层能选中正确 variant；AI 调用前先 `get:chart` 读取并回填，比让 schema 层放任所有字段都 optional 更可靠
- **复用 `ppt:delete:element`**：删除统一走通用入口，避免新增 `ppt:delete:chart` 造成 API 表面冗余
- **`ChartType` 与 Office.js 对齐**：枚举值与 `Excel.ChartType` 完全一致（如 `ColumnClustered` / `LineMarkers`），消费方可直接 `cast`，无需维护额外映射表
- **`ChartType` 跨命名空间提升**：未来 Excel `excel:insert:chart` 可直接复用同一枚举与 `ChartData` 联合，避免双套语义
- **不新增 `update:element` 字段塞图表**：`ppt:update:element` 仅承担几何属性，图表的「数据/类型/标题」是图表特有语义，独立 `update:chart` 边界更清晰、错误恢复粒度更细

**兼容性**:

- 新增事件 + 新增可选字段，不影响 `/word` 等其它命名空间
- `ChartType` 枚举值变更属破坏性变更，但 Excel 侧 `excel:insert:chart` 尚未定义、PPT 侧本就无 chart 事件，实际无在用消费方
- discriminated union 形态在本 [Unreleased] 周期内确定（含中途 PR #4 → PR #5 的 schema 重构），无对外发版的旧 shape 残留
- 新事件初始标记 📋 Draft，待消费方（office4ai）落地稳定后再统一转 ✅ Stable

**相关 Issue / 工单**:

- 协议侧：[oasp-protocol#2](https://github.com/A2C-SMCP/oasp-protocol/issues/2)（Milestone: PPT Chart Capabilities）
- 协议 PR：[#4](https://github.com/A2C-SMCP/oasp-protocol/pull/4)（首版 categorical-only schema）→ [#5](https://github.com/A2C-SMCP/oasp-protocol/pull/5)（重构为 discriminated union，提升 LLM 可发现性）
- 消费方实现：[office4ai#9](https://github.com/JIAQIA/office4ai/issues/9)（Server 端 python-pptx OOXML 拦截）

---

**`/word` 命名空间补齐表格操作能力（4 个 Draft 事件）**

| 变更类型 | 事件 / 类型 | 说明 |
|----------|-------------|------|
| 新增 | `word:merge:cells` | 📋 Draft，合并表格中任意矩形单元格区域 |
| 新增 | `word:update:tableCell` | 📋 Draft，批量更新单元格文本与格式 |
| 新增 | `word:update:tableRowColumn` | 📋 Draft，按行/列批量更新单元格内容 |
| 新增 | `word:update:tableFormat` | 📋 Draft，更新整表样式、边框、列宽、对齐 |
| 新增 | `CellFormat` 数据结构 | 定义于 `data-structures.md`，承载单元格对齐/字体/底色等格式属性，供 `word:update:tableCell` 等事件复用 |
| 新增 | 错误码 `3013 NO_TABLE_AT_CURSOR` / `3014 ALREADY_MERGED` | 表格类事件专用错误码，便于 AI 区分"表格 ID 错"、"光标位置错"、"已合并冲突"三类失败 |
| 新增（Stable 事件向后兼容扩展） | `word:get:documentStructure` 增加可选 `tables: TableSummary[]` | 用于"重新发现"现有表格，配合 `precedingHeading` 启发式定位；旧客户端忽略该字段不受影响 |

**动机**: Add-In 内部 `word-tools/table.ts` 已实现合并单元格、单元格内容/格式更新、整表样式/边框/列宽更新等能力，协议层此前仅暴露 `word:insert:table`，导致 AI 无法调用细粒度表格操作（如生成"蓝底居中表头 + 灰底加粗标签"等常见合同/报告美化效果，或合并表头横跨多列）。本批新增事件按"协议先行"原则补齐该能力面，统一以 `tableId`（与 `word:insert:table` 响应一致）作为定位标识；缺省 `tableId` 时取当前光标所在表格。

**与 Office.js 对齐**（消费方 office-editor4ai 验证后修订）:

- `CellFormat.horizontalAlignment` / `verticalAlignment` 枚举值与 `Word.Alignment` / `Word.VerticalAlignment` 完全一致——`"Centered"` / `"Justified"` 用过去分词形，`"Center"` 表示垂直居中（不再是 `"Middle"`），Add-In 可直接 `cast` 不做映射
- `CellFormat.cellPadding` **未纳入**——Word.js 单元格内边距是表级 API（`Word.Table.setCellPadding`），无法逐单元格设置；改放到 `word:update:tableFormat.styleOptions.cellPadding: { top, bottom, left, right }`
- `word:merge:cells` 响应字段从 `mergedCells: number` 改为 `requestedRange: { rowCount, columnCount }`——Word.js 不暴露"被合并的原子单元格总数"，乘法计算在已部分合并区域会失真，返回请求矩形的尺寸语义更清晰
- `word:update:tableFormat.borderOptions.location?: "all" | "inside" | "outside"`（默认 `"all"`）——覆盖"内细外粗"高频场景，对应 `Word.BorderLocation`
- `word:update:tableFormat` **移除 `data` 字段**——避免与 `word:update:tableRowColumn` 职责重叠，强制写数据与改样式分两步调用，错误恢复粒度更细
- `word:update:tableFormat.styleOptions.styleType` 不存在的样式名错误码使用 `3011 STYLE_NOT_FOUND`（已存在），非 `OPERATION_FAILED`
- `word:update:tableFormat.columnWidths` 长度策略放宽为"≤ 列数"——过短只覆盖前缀列，避免 AI 边缘列少传一个就触发失败
- 全部表格事件错误码修正：`3010 ELEMENT_NOT_FOUND`（表格未找到）+ `3013 NO_TABLE_AT_CURSOR`（缺省 tableId + 光标不在表格内）替代此前误用的 `3003 OPERATION_FAILED`，与现有 `error-handling.md` 错误码定义对齐

**新增警示**：

- `word:insert:table` 响应中 `tableId` 当前为**临时索引**，跨会话或经过结构变更的场景调用方应通过 `word:get:documentStructure` 重新发现表格（响应文档新增 `tableId 稳定性` 警告）。基于 Content Control 的稳定 ID 方案为单独工单跟进，不阻塞本批 Draft 事件转 Stable

**兼容性**: 全部为新增事件 + 新增可选字段，不影响现有 Stable 事件；`/word` 命名空间稳定性整体保持。新事件初始标记 📋 Draft，待消费方双端落地稳定后再统一转 ✅ Stable。

**相关 Issue / 工单**:

- 协议侧：[oasp-protocol#1](https://github.com/A2C-SMCP/oasp-protocol/issues/1)（Milestone: Word Table Capabilities）
- 消费方：[OF4AI-10](https://turingfocus.atlassian.net/browse/OF4AI-10)（mergeCells）、[OF4AI-11](https://turingfocus.atlassian.net/browse/OF4AI-11)（updateCell / updateTable）

---

## [0.1.9] - 2026-04-17

### 新增

**word:insert:table 新增 `insertLocation` 字段**

| 变更类型 | 字段 | 说明 |
|----------|------|------|
| 新增 | `TableInsertOptions.insertLocation` | 可选枚举字段，取值 `Start` / `End` / `Before` / `After` / `Replace`，未传时默认 `"End"`（向后兼容） |
| 新增 | `TableInsertLocation` 类型定义 | 表格插入位置枚举类型 |

**动机**: Add-In 内部 `insertTable` 函数已支持五种插入位置，但此前协议层仅定义 `rows` / `columns` / `data` / `style`，导致 AI 无法指定表格插入位置，所有表格默认插入文档末尾。此次变更为协议层能力补齐，不影响现有客户端。

**相关工单**: [OF4AI-9](https://turingfocus.atlassian.net/browse/OF4AI-9)

---

## [0.1.8] - 2026-02-05

### 变更

**word:get:styles 响应结构调整**

基于 Add-In 实际实现进行协议规范对齐：

| 变更类型 | 字段 | 说明 |
|----------|------|------|
| 修改 | `timestamp` (请求) | 从必需改为可选 |
| 删除 | `duration` (响应) | Add-In 未实现，移除 |
| 修改 | `success` (响应) | 从字面量 `true` 改为 `boolean` |
| 新增 | `error` (响应) | 添加可选错误响应字段 |

---

## [0.1.7] - 2026-02-05

### 变更

**word:get:documentStats 请求与响应结构调整**

基于 Add-In 实际实现进行协议规范对齐：

| 变更类型 | 字段 | 说明 |
|----------|------|------|
| 新增 | 请求结构 | 补充完整的请求定义（原协议缺失） |
| 修改 | `timestamp` (请求) | 从必需改为可选 |
| 删除 | `duration` (响应) | Add-In 未实现，移除 |
| 修改 | `success` (响应) | 从字面量 `true` 改为 `boolean` |
| 新增 | `error` (响应) | 添加可选错误响应字段 |
| 修改 | `characterCount` | 明确语义为"不含空格" |
| 新增 | `characterCountWithSpaces` | 含空格的字符数 |
| 新增 | `pageCount` | 页数（可选） |
| 重命名 | `DocumentStats` → `DocumentStatsResult` | 统一命名风格 |

---

## [0.1.6] - 2026-02-05

### 变更

**word:get:documentStructure 请求与响应结构调整**

基于 Add-In 实际实现进行协议规范对齐：

| 变更类型 | 字段 | 说明 |
|----------|------|------|
| 修改 | `timestamp` (请求) | 从必需改为可选 |
| 删除 | `duration` (响应) | Add-In 未实现，移除 |
| 修改 | `success` (响应) | 从字面量 `true` 改为 `boolean` |
| 新增 | `error` (响应) | 添加可选错误响应字段 |
| 修改 | `data` 字段顺序 | 调整为 `sectionCount → paragraphCount → tableCount → imageCount` |

---

## [0.1.5] - 2026-02-05

### 变更

**word:get:visibleContent 请求与响应结构调整**

基于 Add-In 实际实现进行协议规范对齐：

| 变更类型 | 字段 | 说明 |
|----------|------|------|
| 修改 | `timestamp` (请求) | 从必需改为可选 |
| 新增 | `options` (请求) | 支持 GetContentOptions（与 selectedContent 一致） |
| 删除 | `duration` (响应) | Add-In 未实现，移除 |
| 删除 | `data.startPosition` | Add-In 未实现，移除 |
| 删除 | `data.endPosition` | Add-In 未实现，移除 |
| 新增 | `data.elements` | 内容元素数组，带类型映射 |
| 新增 | `data.metadata` | 统计元数据 |

**新增 VisibleContentElement 结构**：包含 `type`（映射后类型）和 `content`（原始元素）。

---

## [0.1.4] - 2026-02-05

### 变更

**word:get:selectedContent 请求与响应结构调整**

基于 Add-In 实际实现进行协议规范对齐：

| 变更类型 | 字段 | 说明 |
|----------|------|------|
| 修改 | `timestamp` (请求) | 从必需改为可选 |
| 重构 | `options` | 从 `includeFormatting` 改为 6 个独立选项 |
| 删除 | `duration` (响应) | Add-In 未实现，移除 |
| 删除 | `data.html` | Add-In 未实现，移除 |
| 删除 | `data.format` | Add-In 未实现，移除 |
| 新增 | `data.elements` | 内容元素数组（段落、表格、图片、内容控件） |
| 新增 | `data.metadata` | 统计元数据（字符数、段落数等） |

**新增 GetContentOptions**：`includeText`、`includeImages`、`includeTables`、`includeContentControls`、`detailedMetadata`、`maxTextLength`

**新增内容元素类型**：`ParagraphElement`、`TableElement`、`InlinePictureElement`、`ContentControlElement`

---

## [0.1.3] - 2026-02-05

### 变更

**word:get:selection 请求与响应结构调整**

基于 Add-In 实际实现进行协议规范对齐：

| 变更类型 | 字段 | 说明 |
|----------|------|------|
| 修改 | `timestamp` (请求) | 从必需改为可选 |
| 删除 | `duration` (响应) | Add-In 未实现，移除 |
| 完善 | `data.type` | 补充完整枚举值：`NoSelection`、`InsertionPoint`、`Normal` |
| 修改 | `data.start` | 从必需改为可选（仅选区非空时存在） |
| 修改 | `data.end` | 从必需改为可选（仅选区非空时存在） |
| 修改 | `data.text` | 从必需改为可选（仅选区非空时存在） |

**新增响应示例**：覆盖"有选区"、"光标点"、"无选区"三种场景。

---

## [0.1.2] - 2026-02-05

### 变更

**word:event:selectionChanged 事件结构调整**

基于 Add-In 实际实现进行协议规范对齐：

| 变更类型 | 字段 | 说明 |
|----------|------|------|
| 新增 | `eventType` | 事件类型标识，固定值 `"selectionChanged"` |
| 新增 | `clientId` | 发送事件的客户端标识 |
| 修改 | `selection` → `data` | 字段名变更，简化结构 |
| 删除 | `selection.isEmpty` | 实现未提供 |
| 删除 | `selection.type` | 实现未提供 |
| 删除 | `selection.start` | 实现未提供 |
| 删除 | `selection.end` | 实现未提供 |
| 新增 | `data.text` | 选中的文本内容 |
| 新增 | `data.length` | 选中文本的字符长度 |

**word:event:documentModified 事件结构调整**

基于 Add-In 实际实现进行协议规范对齐：

| 变更类型 | 字段 | 说明 |
|----------|------|------|
| 新增 | `eventType` | 事件类型标识，固定值 `"documentModified"` |
| 新增 | `clientId` | 发送事件的客户端标识 |
| 新增 | `data` | 事件数据对象 |
| 新增 | `data.modificationType` | 修改类型：`insert`、`delete`、`update` |

**word:get:styles 请求参数调整**

基于 Add-In 实际实现进行协议规范对齐：

| 变更类型 | 参数 | 说明 |
|----------|------|------|
| 保留 | `includeBuiltIn` | 无变化 |
| 保留 | `includeCustom` | 无变化 |
| 删除 | `typeFilter` | Add-In 未实现，移除以保持一致性 |
| 新增 | `includeUnused` | 控制是否返回文档中未使用的样式，默认 false |
| 新增 | `detailedInfo` | 控制是否返回 description 字段，默认 false |

**StyleInfo.description 字段说明**

- `description` 字段现仅在请求 `detailedInfo=true` 时返回
- 此功能依赖 WordApi BETA，在部分环境中可能不可用

---

## [0.1.1] - 2026-02-02

### 变更

**事件命名规范化**

为保持命名一致性，统一采用 `{namespace}:{action}:{target}` 格式：

- PPT 事件:
  - `ppt:slide:add` → `ppt:add:slide`
  - `ppt:slide:delete` → `ppt:delete:slide`
  - `ppt:slide:move` → `ppt:move:slide`
  - `ppt:slide:goto` → `ppt:goto:slide`

- Excel 事件:
  - `excel:get:rangeValues` → `excel:get:range`
  - `excel:set:rangeValues` → `excel:set:range`
  - `excel:sheet:add` → `excel:add:sheet`
  - `excel:sheet:delete` → `excel:delete:sheet`
  - `excel:sheet:rename` → `excel:rename:sheet`
  - `excel:sheet:activate` → `excel:activate:sheet`

**错误码调整**

- `2002` 从 `CONNECTION_LOST` 改为 `INVALID_TOKEN`（令牌无效）
- 新增 `2005 CONNECTION_LOST`（连接丢失）

**数据类型简化**

- `ChartType`: 简化为 Column, Bar, Line, Pie, Area, Scatter, Doughnut
- `ShapeType`: 合并两端实现，现包含 Rectangle, RoundedRectangle, Circle, Oval, Triangle, Diamond, Pentagon, Hexagon, Line, Arrow, Star, TextBox

### 新增

**PPT 事件**

- `ppt:insert:table` - 在幻灯片中插入表格
- `ppt:update:textBox` - 更新幻灯片中的文本框

---

## [0.1.0] - 2026-02-02

### 新增

**核心协议**

- 定义了协议名称: OASP (Office AddIn Socket Protocol)
- 定义了两角色通信模型 (Server ↔ AddIn)
- 定义了三个命名空间: `/word`, `/ppt`, `/excel`
- 定义了事件命名规范: `{namespace}:{action}:{target}`
- 定义了请求-响应模式和事件报告模式

**连接与握手**

- 定义了握手参数: `clientId`, `documentUri`
- 定义了连接确认事件: `connection:established`
- 定义了断开连接和重连机制

**Word 事件 (✅ Stable)**

- `word:event:selectionChanged` - 选区变化通知
- `word:event:documentModified` - 文档修改通知
- `word:get:selection` - 获取选区位置信息
- `word:get:selectedContent` - 获取选中内容
- `word:get:visibleContent` - 获取可见内容
- `word:get:documentStructure` - 获取文档结构
- `word:get:documentStats` - 获取文档统计
- `word:get:styles` - 获取文档样式列表
- `word:insert:text` - 插入文本
- `word:replace:selection` - 替换选中内容
- `word:replace:text` - 查找替换
- `word:select:text` - 查找并选中文本

**Word 事件 (📋 Draft)**

- `word:append:text` - 追加文本
- `word:insert:image` - 插入图片
- `word:insert:table` - 插入表格
- `word:insert:equation` - 插入公式
- `word:insert:toc` - 插入目录
- `word:export:content` - 导出内容

**PPT 事件 (📋 Draft)**

- `ppt:get:currentSlideElements` - 获取当前幻灯片元素
- `ppt:get:slideElements` - 获取指定幻灯片元素
- `ppt:get:slideScreenshot` - 获取幻灯片截图
- `ppt:insert:text` - 插入文本
- `ppt:insert:shape` - 插入形状
- `ppt:insert:image` - 插入图片
- `ppt:slide:add` - 添加幻灯片
- `ppt:slide:delete` - 删除幻灯片
- `ppt:slide:move` - 移动幻灯片
- `ppt:slide:goto` - 跳转到幻灯片

**Excel 事件 (📋 Draft)**

- `excel:get:selectedRange` - 获取选中范围
- `excel:get:usedRange` - 获取已使用范围
- `excel:get:cellValue` - 获取单元格值
- `excel:get:rangeValues` - 获取范围内的值
- `excel:set:cellValue` - 设置单元格值
- `excel:set:rangeValues` - 设置范围内的值
- `excel:insert:table` - 插入表格
- `excel:insert:chart` - 插入图表
- `excel:sheet:add` - 添加工作表
- `excel:sheet:delete` - 删除工作表
- `excel:sheet:rename` - 重命名工作表
- `excel:sheet:activate` - 激活工作表

**数据结构**

- 定义了基础请求/响应结构
- 定义了 `SelectionInfo`, `TextFormat`, `StyleInfo` 等核心类型
- 定义了 PPT 和 Excel 相关类型

**错误处理**

- 定义了错误码分类 (1xxx-4xxx)
- 定义了通用错误码 (1xxx)
- 定义了连接错误码 (2xxx)
- 定义了文档操作错误码 (3xxx)
- 定义了参数验证错误码 (4xxx)

**通用约定**

- 定义了时间戳格式 (Unix 毫秒)
- 定义了字段命名规范 (camelCase)
- 定义了颜色值格式 (#RRGGBB)
- 定义了单位规范 (磅、像素)

**文档**

- 创建了完整的协议文档结构
- 创建了术语表
- 创建了变更日志

---

## 版本号说明

- **主版本号 (Major)**: 不兼容的 API 变更
- **次版本号 (Minor)**: 向后兼容的功能新增
- **修订号 (Patch)**: 向后兼容的问题修复

协议当前处于 `0.x` 初始开发阶段，API 可能随时变更。
