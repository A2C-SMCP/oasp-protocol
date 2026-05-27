# 错误处理

## 概述

本章定义 OASP 协议的错误码体系和错误处理规范。

## 设计原则

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
| `3003` | `DOCUMENT_READ_ONLY` | 文档为只读模式 |
| `3004` | `OPERATION_FAILED` | 操作执行失败 |
| `3005` | `RESOURCE_NOT_ACCESSIBLE` | 资源不可访问 |
| `3006` | `CONTENT_TOO_LARGE` | 内容超过大小限制 |
| `3007` | `FORMAT_NOT_SUPPORTED` | 格式不支持 |
| `3008` | `POSITION_INVALID` | 位置无效 |
| `3009` | `RANGE_INVALID` | 范围无效 |
| `3010` | `ELEMENT_NOT_FOUND` | 元素或幻灯片未找到（按 elementId / slideId 定位失败） |
| `3011` | `STYLE_NOT_FOUND` | 样式未找到 |
| `3012` | `SEARCH_NO_MATCH` | 搜索无匹配结果 |
| `3013` | `NO_TABLE_AT_CURSOR` | 缺省 `tableId` 时光标未落在任何表格内 |
| `3014` | `ALREADY_MERGED` | 目标合并区域内已存在合并冲突，无法再次合并 |
| `3015` | `INVALID_CHART_DATA` | 图表数据非法（categorical: series.values 长度与 categories 不匹配；scatter: points 为空或含非法值；或跨类型切换未补齐数据） |
| `3016` | `API_NOT_SUPPORTED` | 目标操作在当前客户端/平台不可用（如所需 requirement set 不满足或宿主不支持该能力） |

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

**触发场景**: 目标文档当前不可写入——只读、被锁定或其它原因导致修改无法应用。该判断只描述「文档可写与否」这一线缆可观测状态，不绑定任何具体实现技术。

**处理建议**:
- 检查文档是否被其他程序锁定
- 检查用户是否有编辑权限
- 提示用户保存文档副本

### API_NOT_SUPPORTED (3016)

**触发场景**: 目标操作所依赖的能力在当前客户端或平台上不可用——例如所需的 PowerPointApi requirement set 未满足，或宿主环境（如移动端、老版本永久授权 Office）不支持该能力。

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

**触发场景**: `word:merge:cells` 请求的矩形区域与已有合并单元格冲突（如目标矩形跨越了已合并区域的一部分），Word.js 无法执行二次合并。

**处理建议**:

- 调用方先确认目标区域当前合并状态，必要时调整起止索引
- 如需重新合并，先取消已有合并（暂未提供拆分事件）

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
