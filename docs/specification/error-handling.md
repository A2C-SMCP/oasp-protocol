# 错误处理

## 概述

本章定义 OASP 协议的错误码体系和错误处理规范。

## 设计原则

OASP 的错误码设计遵循以下原则：

1. **与应用无关**: 错误码描述的是「操作失败」或「状态异常」，而非特定于 Word、PPT 或 Excel
2. **语义清晰**: 每个错误码有明确的含义，便于定位问题
3. **数字分段**: 使用数字范围区分错误类别，便于程序化处理

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
| `2002` | `CONNECTION_LOST` | 连接丢失 |
| `2003` | `HANDSHAKE_FAILED` | 握手失败 |
| `2004` | `SESSION_INVALID` | 会话无效 |

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
| `3010` | `ELEMENT_NOT_FOUND` | 元素未找到 |
| `3011` | `STYLE_NOT_FOUND` | 样式未找到 |
| `3012` | `SEARCH_NO_MATCH` | 搜索无匹配结果 |

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

**触发场景**: 尝试修改只读文档。

**处理建议**:
- 检查文档是否被其他程序锁定
- 检查用户是否有编辑权限
- 提示用户保存文档副本

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
