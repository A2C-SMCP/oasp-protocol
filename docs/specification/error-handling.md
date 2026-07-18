# 错误处理

## 概述

本章定义 OASP 协议的错误码体系和错误处理规范。

## 设计原则 {#design-principles}

OASP 的错误码设计遵循以下原则：

1. **与应用无关**: 错误码描述的是「操作失败」或「状态异常」，而非特定于 Word、PPT 或 Excel
2. **与实现无关**: 触发条件描述**线缆可观测的状态**（如「目标文档不可写入」），而非特定实现技术或执行位置（如「服务端 OOXML 写入失败」）。同一错误码可由任意实现路径在等价条件下抛出
3. **语义清晰**: 每个错误码有明确的含义，便于定位问题
4. **数字分段**: 使用数字范围区分错误类别，便于程序化处理

## 错误响应格式

所有错误响应遵循统一格式：

```typescript
interface ErrorResponse {
  requestId: string;
  success: false;
  error: {
    code: string;          // 错误码
    message: string;       // 人类可读的错误消息
    details?: object;      // 附加详情（可选）
  };
  timestamp: number;
  duration?: number;
}
```

**示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": false,
  "error": {
    "code": "SELECTION_EMPTY",
    "message": "Selection is empty, cannot perform replace operation",
    "details": {
      "operation": "replace:selection"
    }
  },
  "timestamp": 1704067200500,
  "duration": 10
}
```

## 错误码分类

### 1xxx - 通用错误

适用于所有操作的通用错误。

| 错误码 | 名称 | 说明 |
|--------|------|------|
| `1000` | `UNKNOWN` | 未知错误 |
| `1001` | `INVALID_REQUEST` | 请求格式无效 |
| `1002` | `TIMEOUT` | 操作超时 |
| `1003` | `NOT_IMPLEMENTED` | 功能未实现 |
| `1004` | `INTERNAL_ERROR` | 内部错误 |
| `1005` | `RATE_LIMITED` | 请求过于频繁 |

### 2xxx - 连接与认证错误

与连接和认证相关的错误。

| 错误码 | 名称 | 说明 |
|--------|------|------|
| `2000` | `UNAUTHORIZED` | 未授权 |
| `2001` | `TOKEN_EXPIRED` | 令牌已过期 |
| `2002` | `INVALID_TOKEN` | 令牌无效 |
| `2003` | `HANDSHAKE_FAILED` | 握手失败 |
| `2004` | `SESSION_INVALID` | 会话无效 |
| `2005` | `CONNECTION_LOST` | 连接丢失 |
| `2006` | `PROTOCOL_VERSION_MISMATCH` | 协议版本不兼容（握手阶段拒绝） |

### 3xxx - 文档与操作错误

与文档状态和操作执行相关的错误。

| 错误码 | 名称 | 说明 |
|--------|------|------|
| `3000` | `DOCUMENT_ERROR` | 文档操作错误（通用） |
| `3001` | `DOCUMENT_NOT_FOUND` | 文档不存在或未打开 |
| `3002` | `SELECTION_EMPTY` | 选区为空（需要非空选区的操作） |
| `3003` | `DOCUMENT_READ_ONLY` | 目标不可写入（文档只读、被锁定，或其中的工作表/区域受保护） |
| `3004` | `OPERATION_FAILED` | 操作执行失败 |
| `3006` | `CONTENT_TOO_LARGE` | 内容超过大小限制 |
| `3007` | `FORMAT_NOT_SUPPORTED` | 载荷/目标格式可识别但不受支持（apply-time，区别于载荷不可解码 `4002`） |
| `3008` | `POSITION_INVALID` | 序号位置相对当前文档状态无效（apply-time，区别于静态越界 `4004`） |
| `3009` | `RANGE_INVALID` | 范围/区域地址无效（如非法的 A1 表示法、越界区域） |
| `3010` | `ELEMENT_NOT_FOUND` | 具名/具 id 对象未找到（元素、幻灯片、工作表、表格、图表、透视表等，按 id 或名称定位失败；具体类型见 `details.kind`） |
| `3011` | `STYLE_NOT_FOUND` | 样式未找到 |
| `3012` | `SEARCH_NO_MATCH` | 搜索无匹配结果 |
| `3013` | `NO_TABLE_AT_CURSOR` | 缺省 `tableId` 时光标未落在任何表格内 |
| `3014` | `ALREADY_MERGED` | 目标合并区域内已存在合并冲突，无法再次合并 |
| `3015` | `INVALID_CHART_DATA` | 图表数据非法（categorical: series.values 长度与 categories 不匹配；scatter: points 为空或含非法值；或跨类型切换未补齐数据） |
| `3016` | `API_NOT_SUPPORTED` | 目标操作在当前客户端/平台不可用（如所需 requirement set 不满足、宿主不支持该能力，或该平台不提供此功能如透视表） |
| `3017` | `FORMULA_ERROR` | 公式语法错误或引用无效 |
| `3018` | `DATA_TYPE_MISMATCH` | 写入值与目标单元格/列的数据类型不兼容（apply-time，区别于线缆参数类型错误 `4003`） |

### 4xxx - 参数验证错误

请求参数验证相关的错误。

| 错误码 | 名称 | 说明 |
|--------|------|------|
| `4000` | `VALIDATION_ERROR` | 参数验证失败（通用） |
| `4001` | `MISSING_PARAM` | 缺少必填参数 |
| `4002` | `INVALID_PARAM` | 参数值无效 |
| `4003` | `INVALID_PARAM_TYPE` | 参数类型错误 |
| `4004` | `PARAM_OUT_OF_RANGE` | 参数超出范围 |

## 错误码详解

### TIMEOUT (1002)

**触发场景**: 请求在指定时间内未收到响应。

**处理建议**:
- 服务端应在超时后自动标记请求失败
- 不进行自动重试（避免重复操作）
- 客户端可选择向用户报告并允许手动重试

**示例**:

```json
{
  "error": {
    "code": "TIMEOUT",
    "message": "Operation timed out after 30000ms",
    "details": {
      "timeoutMs": 30000,
      "operation": "word:get:documentStats"
    }
  }
}
```

### PROTOCOL_VERSION_MISMATCH (2006)

**触发场景**: AddIn 在握手 `auth.oaspVersion` 中声明的协议版本与 Server 不兼容（`is_compatible` 判定为假，规则见[通用约定 · 兼容性判定规则](conventions.md#compatibility-rule)）。

**特殊性**: 该错误发生在**连接握手阶段**，不属于请求-响应周期——没有 `requestId`，**不走**标准 `ErrorResponse`。Server 在 `connect` handler 中抛 `ConnectionRefusedError`，**扁平**拒绝数据经 Socket.IO 送达 AddIn `connect_error` 的 `error.data`（见[连接与握手 · 协议版本握手](connection.md#protocol-version-handshake)）。

**拒绝数据**（扁平，非 `ErrorResponse`）:

```json
{
  "code": "PROTOCOL_VERSION_MISMATCH",
  "message": "Protocol version mismatch",
  "serverVersion": "0.3.0",
  "clientVersion": "0.2.0",
  "minSupported": "0.3.0",
  "maxSupported": "0.3.999"
}
```

**处理建议**:

- AddIn **MUST** 主动 `disconnect()` 并停止自动重连，再上抛明确异常——不得静默重试（重连只会再次触发同一拒绝，进入死循环）
- 据 `serverVersion` / `clientVersion` 提示用户升级 AddIn 或切换到匹配的 Server 实例

**错误码复用说明**:

- **复用 `2003 HANDSHAKE_FAILED`**: 缺少或格式非法的 `oaspVersion`（与缺少 `clientId` / `documentUri` 同类，属「握手参数不合法」）
- **新增 `2006 PROTOCOL_VERSION_MISMATCH`**: 参数合法但版本语义不兼容（专码以便 AddIn 区分「参数错」与「版本错」——前者改参数，后者须升级端点）

### SELECTION_EMPTY (3002)

**触发场景**: 执行需要非空选区的操作（如 `replace:selection`）时，当前选区为空。

**处理建议**:
- 提示用户先选中内容
- 或改用其他不依赖选区的操作

**示例**:

```json
{
  "error": {
    "code": "SELECTION_EMPTY",
    "message": "Selection is empty, cannot perform replace operation",
    "details": {
      "operation": "word:replace:selection",
      "hint": "Please select content before replacing"
    }
  }
}
```

### DOCUMENT_READ_ONLY (3003)

**触发场景**: 目标不可写入——文档只读、被锁定，或其中的工作表/区域受保护，导致修改无法应用。该判断只描述「目标可写与否」这一线缆可观测状态，不绑定任何具体实现技术。受保护的范围可经 `details.scope`（如 `"document"` / `"worksheet"`）指明。

**处理建议**:
- 检查文档是否被其他程序锁定
- 检查用户是否有编辑权限，或目标工作表是否受保护
- 提示用户保存文档副本，或解除工作表保护后重试

### OPERATION_FAILED (3004)

**触发场景**: 操作在执行阶段失败。

**`run:script` 的二级分流**: 承载 `{namespace}:run:script`（[通用约定 · 脚本执行](conventions.md#run-script)）的执行期失败时，`details.fault` 区分失败来源——`"script"`（脚本自身抛错：TypeError / ReferenceError / 主动 throw，附 `name` / `stack`）或 `"office"`（脚本调用 Office.js 失败，附 `officeCode` / `debugInfo?`）。AI 据此决定**重写脚本** vs **调整操作**。脚本各阶段（compile / execute / serialize）→ 错误码的完整映射见[通用约定 · 脚本执行 · 错误映射](conventions.md#run-script-errors)。

### FORMAT_NOT_SUPPORTED (3007)

**触发场景**: 事件本身可用，但**所请求的这一种格式变体**不受支持。两侧对称：

- **输入侧**（载荷格式）：数据完整、可解码，但其格式目标拒绝接受（如解码后是宿主拒绝嵌入的图片格式）
- **输出侧**（请求产出的格式）：`format` 枚举合法，但目标无法产出该格式（如无法导出为 Markdown）

**判法划界**（三码易混用，按**失败的粒度**区分）:

| 条件（线缆可观测） | 错误码 | 调用方应做 |
|---|---|---|
| 载荷在线缆层即不合法（base64 无法解码、字段缺失） | `4002 INVALID_PARAM` | 修正参数后重发 |
| 事件可用，但**这一种格式**不受支持 | `3007 FORMAT_NOT_SUPPORTED` | **换格式**重发：输入侧转码载荷，输出侧改请求另一格式 |
| **整个操作/能力**在当前宿主不可用 | `3016 API_NOT_SUPPORTED` | 换路径或换平台（换格式无用） |

`3007` 与 `3016` 的分界是**失败的粒度**，不是「谁的锅」——两者都可能源于宿主。`3007` 说「这个事件能用，只是不吃这一种格式」：**同一宿主、同一路径换个格式即可成功**；`3016` 说「这个能力在这儿整个不存在」：换格式无济于事，须换路径或平台。误报 `3016` 会把调用方引向「换平台」这条对格式问题无效的路，而误报 `3007` 会让它在一个根本不存在的能力上徒劳试遍所有格式。

本码承 `4003 INVALID_PARAM_TYPE` ↔ `3018 DATA_TYPE_MISMATCH` 的「线缆层 / apply-time」配对先例——`4002` / `3007` 是同一划分在**格式**轴上的应用：格式值在线缆上合法，仅在应用时被拒。

**处理建议**:

- **输入侧**：转码到通用格式（如图片转 PNG）后重发
- **输出侧**：改请求另一个格式（如 `markdown` → `html`）
- `details` 可回带被拒格式与目标可接受的格式列表，便于调用方一次选对

### POSITION_INVALID (3008)

**触发场景**: 按**序号**定位的位置相对**当前文档状态**无效——如 `slideIndex: 15` 发往一个 10 页的演示文稿。该值在线缆上完全合法（非负整数、类型正确），仅因文档此刻的状态而不可用。

!!! warning "过渡期：本码的规范层效力暂仅及于显式列出它的事件"
    下表判法是**目标态**，但全量接线尚未完成——规范内仍有 ~22 处「序号越界」判 `4002`（`/ppt`、`/word`）或 `4004`（`/excel`）。**在 [#24](https://github.com/A2C-SMCP/oasp-protocol/issues/24) 清扫完成前**，实现**以各事件「可能的错误」表所列为准**：仅 `ppt:delete:slide` / `ppt:goto:slide` 按 `3008` 判定，其余沿用其表内现码。清扫完成后本 admonition 移除，下表即对全部序号越界条件生效。

**判法划界**:

| 条件（线缆可观测） | 错误码 | 调用方应做 |
|---|---|---|
| 参数值本身非法（枚举外取值、类型错） | `4002 INVALID_PARAM` / `4003 INVALID_PARAM_TYPE` | 修正参数后重发 |
| 超出**静态声明**的边界（如 `timeoutMs` 上限） | `4004 PARAM_OUT_OF_RANGE` | 修正参数后重发 |
| 序号超出**当前文档**的实际范围 | `3008 POSITION_INVALID` | **重读文档状态**后重发 |
| 按 id / 名称定位对象失败 | `3010 ELEMENT_NOT_FOUND` | 先列举可用对象再重发 |
| A1 表示法的区域地址非法或越界 | `3009 RANGE_INVALID` | 修正地址后重发 |

`3008` 与 `4004` 的分界是**边界从哪来**：`4004` 的边界由规范**静态声明**（改参数即可），`3008` 的边界是**文档运行时状态**（参数本身可能没错，须重读状态）。归入 4xxx 会让调用方按本文「错误处理最佳实践 · AddIn 端」的「参数错误 → 不可重试」处理，而正确恢复恰是**重读后以原值重试**（文档增长后同一序号即有效）。

`3008` 是 `3009 RANGE_INVALID` 在**序号**轴上的对应物——后者管 A1 区域地址，前者管 `slideIndex` / `rowIndex` / `columnIndex` / `insertIndex` 一类序数。

**`details`**: 回带请求序号与实时边界，如 `{ "index": 15, "total": 10, "kind": "slide" }`。

`kind` **复用** [`3010` 的 `details.kind` 词表](#element_not_found-3010)——该 token 标明**对象类型**，而**定位方式由错误码本身区分**（`3010` 必为按 id / 名称查找，`3008` 必为按序号索引），故两码同名同义、**不分叉**。全量接线（[#24](https://github.com/A2C-SMCP/oasp-protocol/issues/24)）若需 `row` / `column` 等新 token，须并入该词表统一定义。

**处理建议**:

- 先用对应的 `get:*` 读取当前实际数量，再以有效序号重发

### API_NOT_SUPPORTED (3016)

**触发场景**: 目标操作所依赖的能力在当前客户端或平台上不可用——例如所需的 requirement set 未满足（如 PowerPointApi、ExcelApi），或宿主环境（如移动端、Web 端、老版本永久授权 Office）不提供该能力（如部分平台不支持透视表 `excel:insert:pivotTable`）。

**处理建议**:

- **反应式降级**：调用方据此切换到另一条实现路径（如客户端 Office.js 路径不可用时回退服务端离线路径），或提示用户更换环境。直接尝试 → 失败返回本码即可保证路由正确性，无需预先声明能力。
- 同一路径上重试无意义——须更换路径或平台后再发起。

**示例**:

```json
{
  "error": {
    "code": "API_NOT_SUPPORTED",
    "message": "Required PowerPointApi requirement set 1.8 is not available on this host",
    "details": {
      "operation": "ppt:update:chart",
      "requiredApiSet": "PowerPointApi 1.8"
    }
  }
}
```

### ELEMENT_NOT_FOUND (3010)

**触发场景**: 按 id 或名称定位的目标对象不存在。该码跨命名空间复用，具体对象类型由 `details.kind` 区分（同一事件可能对不同 `kind` 分别返回 3010，如 Excel 图表事件先查工作表再查图表）。

**`details.kind` 取值**（规范层，双端据此对齐断言）:

| `kind` | 命名空间 | 对象类型 | 本码语境下的定位方式 |
|--------|----------|----------|----------------------|
| `element` | `/word`、`/ppt` | 内容元素 | 按 `elementId` |
| `slide` | `/ppt` | 幻灯片 | 按 `slideId` |
| `worksheet` | `/excel` | 工作表 | 按名称 |
| `table` | `/excel` | 表格 | 按名称 / id |
| `chart` | `/excel` | 图表 | 按名称 |
| `pivotTable` | `/excel` | 透视表 | 按名称 |

`details` 另可回带被查标识（如 `id` / `name`）便于定位。取值集合随命名空间演进而扩充，但已列值语义固定、不复用。

token 承载的是**对象类型**，**定位方式则由错误码本身承载**（故末列限定「本码语境下」——`3010` 必为按 id / 名称查找）。同一批 token 由 [`3008 POSITION_INVALID`](#position_invalid-3008) 复用以标明**被索引**的对象类型（其定位方式必为按序号），两码下 `slide` 都指「幻灯片」这一对象类型，故词表**同名同义、不分叉**。

**处理建议**:

- 先用对应的 `get:*` 列举可用对象，再以有效 id/名称重发
- 据 `details.kind` 定位是哪一类对象未找到

### STYLE_NOT_FOUND (3011)

**触发场景**: 使用不存在的样式名称。

**处理建议**:
- 先使用 `get:styles` 获取可用样式列表
- 使用返回的样式名称进行操作

**示例**:

```json
{
  "error": {
    "code": "STYLE_NOT_FOUND",
    "message": "Style 'Custom Heading' not found in document",
    "details": {
      "requestedStyle": "Custom Heading",
      "availableStyles": ["标题 1", "标题 2", "正文"]
    }
  }
}
```

### SEARCH_NO_MATCH (3012)

**触发场景**: 以**搜索文本定位目标范围**的操作未找到任何匹配，操作因而失去锚点、无法执行。

!!! warning "规范层（Normative）"
    本节为**规范层**（见[通用约定 · 规范分层](conventions.md#normative-layering)）。判据是**零匹配是否仍是一个良定义的结果**——实现须按下表判定，**不得**混用。

| 零匹配时 | 语义 | 判定 | 事件 |
|---|---|---|---|
| 仍是**良定义的零元结果**（0 次替换 / 0 个匹配 / 空列表） | 「没找到」是一个**有效答案** | 正常结果 → `success: true`，**不得**返回 `3012` | `word:replace:text`（`replaceCount: 0`）、`word:select:text`（`matchCount: 0`）、`excel:find:values`（`matches: []`） |
| 使操作**无法产生所请求的效果**（批注无处可附） | 「没找到」= **前置条件不满足** | **MUST** 返回 `3012` | `word:insert:comment`（`target.type: "searchText"`） |

**判据不是「搜索是否为变更锚点」**——`word:replace:text` 的搜索同样是变更的锚点，但「替换了 0 处」是一个**完整、真实、可消费**的答案，报成错误只会迫使调用方从错误处理路径消费正常结果。真正的分界是**零元有没有对应的良定义结果**：`replaceCount: 0` 有，而「把批注附到 0 个范围上」没有——批注要么附上了，要么这次请求**什么也没做成**。

后者若降级为通用 `3000 DOCUMENT_ERROR`，调用方将无法区分「文档操作失败」（通常应上报）与「搜索无匹配」（应改搜索词重试或改用选区模式）——恢复动作完全不同。

**处理建议**:

- 放宽搜索选项（如关闭 `matchCase` / `matchWholeWord`）后重试
- 或改用 `type: "selection"` 模式，由用户先选中目标范围
- `details` 可回带 `searchText` 与实际生效的搜索选项

### NO_TABLE_AT_CURSOR (3013)

**触发场景**: 表格类事件（如 `word:update:tableCell`、`word:merge:cells` 等）在请求中未提供 `tableId`，且当前光标也不在任何表格内。

**处理建议**:

- 提示用户先把光标点击到目标表格内
- 或调用 `word:get:documentStructure` 列出所有 `tables[]`，再以显式 `tableId` 重发请求

**示例**:

```json
{
  "error": {
    "code": "NO_TABLE_AT_CURSOR",
    "message": "tableId not provided and cursor is not inside any table",
    "details": {
      "operation": "word:update:tableCell"
    }
  }
}
```

### ALREADY_MERGED (3014)

**触发场景**: 合并请求的矩形区域与已有合并单元格冲突（如目标矩形跨越了已合并区域的一部分），无法执行二次合并。适用于 Word 表格单元格（`word:merge:cells`）与 Excel 区域（`excel:merge:cells`）。

**处理建议**:

- 调用方先确认目标区域当前合并状态，必要时调整起止索引
- 如需重新合并，先取消已有合并（Word 暂未提供拆分事件；Excel 可先 `excel:unmerge:cells`）

**示例**:

```json
{
  "error": {
    "code": "ALREADY_MERGED",
    "message": "Target range overlaps with existing merged cells",
    "details": {
      "tableId": "table-0",
      "requestedRange": { "startRowIndex": 0, "startColumnIndex": 0, "endRowIndex": 1, "endColumnIndex": 3 }
    }
  }
}
```

### FORMULA_ERROR (3017)

**触发场景**: 写入的公式存在语法错误或引用无效（如括号不匹配、函数名拼写错误、引用了不存在的名称/区域）。适用于承载公式的操作（如 `excel:set:formula`）。

**处理建议**:

- 校正公式字符串后重发（`details` 可回带宿主报告的具体公式错误）
- 确认引用的单元格/区域/命名对象存在

### DATA_TYPE_MISMATCH (3018)

**触发场景**: 写入值与目标单元格/列在应用时（apply-time）的数据类型不兼容——值在线缆上类型合法，但无法被目标接受（如向强类型表格列写入不可转换的值）。区别于 `4003 INVALID_PARAM_TYPE`（请求参数在**线缆层**类型即不合法）。

**处理建议**:

- 调整写入值类型以匹配目标列/单元格
- 如需覆盖类型，先清除目标格式或改用可接受的表示

### MISSING_PARAM (4001)

**触发场景**: 请求缺少必填参数。

**示例**:

```json
{
  "error": {
    "code": "MISSING_PARAM",
    "message": "Missing required parameter: text",
    "details": {
      "missingParams": ["text"],
      "operation": "word:insert:text"
    }
  }
}
```

## 错误处理最佳实践

### Server 端

1. **始终返回 requestId**: 便于关联请求和响应
2. **提供有意义的 message**: 便于调试和日志记录
3. **使用 details 提供上下文**: 帮助定位具体问题
4. **记录错误日志**: 包含完整的请求信息

### AddIn 端

1. **优雅降级**: 遇到错误时显示友好提示
2. **区分可重试和不可重试错误**:
   - 可重试: `TIMEOUT`, `CONNECTION_LOST`
   - 不可重试: `VALIDATION_ERROR`, `SELECTION_EMPTY`
3. **向 Server 报告错误**: 便于服务端监控和分析

### AI Agent 端

1. **解析错误码**: 根据错误码决定下一步操作
2. **利用 details 信息**: 如 `availableStyles` 可用于重新选择样式
3. **避免无限重试**: 对于参数错误等，需要修正参数后重试

## 扩展错误码

如需扩展错误码，请遵循以下规则：

1. **保持分段**: 新增错误码应在对应分类范围内
2. **语义独立**: 新错误码应有明确的独立语义
3. **文档同步**: 更新本文档并发布新版本
4. **向后兼容**: 不修改已有错误码的语义
