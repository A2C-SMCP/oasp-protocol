# 通用约定

## 概述

本章定义 OASP 协议中的通用约定，包括数据格式、命名规则和序列化规范。所有实现**必须**遵循这些约定。

## 时间戳

### 格式

所有时间戳使用 **Unix 毫秒时间戳**（自 1970-01-01 00:00:00 UTC 以来的毫秒数）。

**正确示例**:
```json
{
  "timestamp": 1704067200000
}
```

**错误示例**:
```json
{
  "timestamp": "2024-01-01T00:00:00Z"
}
```

### 时区

时间戳总是表示 **UTC 时间**，不包含时区信息。客户端负责根据需要转换为本地时区。

### 精度

虽然使用毫秒精度，但实际精度取决于系统实现。通常精度在 1-10 毫秒范围内。

## 字段命名

### 传输层命名

**Socket.IO 传输的 JSON 数据使用 camelCase 命名**。

**正确示例**:
```json
{
  "requestId": "abc123",
  "documentUri": "file:///path/to/doc.docx",
  "isEmpty": true,
  "paragraphCount": 10
}
```

**错误示例**:
```json
{
  "request_id": "abc123",
  "document_uri": "file:///path/to/doc.docx"
}
```

### 命名规则总结

| 场景 | 命名风格 | 示例 |
|------|----------|------|
| JSON 字段名 | camelCase | `documentUri`, `requestId` |
| 事件名 | kebab-with-colon | `word:get:selection` |
| 错误码 | SCREAMING_SNAKE_CASE | `SELECTION_EMPTY` |
| 枚举值 | PascalCase | `InsertionPoint`, `Paragraph` |

## 请求 ID

### 格式

请求 ID 使用 **UUID v4 格式**。

**正确示例**:
```
a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d
```

### 生成

- Server 端生成请求 ID
- 每个请求必须有唯一的 ID
- AddIn 在响应中必须返回相同的请求 ID

### 用途

1. **请求-响应关联**: 将响应与对应的请求匹配
2. **去重**: 识别重复请求
3. **日志追踪**: 跨系统追踪请求链路
4. **超时处理**: 标识超时的请求

## 文档 URI

### 格式

文档 URI 使用 **file:// 协议**。

**格式**:
```
file:///{path}
```

**示例**:
```
file:///Users/john/Documents/report.docx
file:///C:/Users/john/Documents/report.docx
```

### 编码

- 路径中的特殊字符使用 URL 编码
- 空格编码为 `%20`

**示例**:
```
file:///Users/john/My%20Documents/report.docx
```

### 大小写

- macOS/Linux: 区分大小写
- Windows: 不区分大小写

建议实现时统一转换为小写进行比较（在 Windows 上）。

## 字符编码

### 文本内容

所有文本内容使用 **UTF-8 编码**。

### JSON 数据

JSON 数据使用 **UTF-8 编码**，不使用 BOM。

### 换行符

文本内容中的换行符：
- Windows: `\r\n` (CRLF)
- macOS/Linux: `\n` (LF)

建议：AddIn 应保留文档原有的换行符风格，不做自动转换。

## 颜色值

### 格式

颜色使用 **十六进制格式**，带 `#` 前缀。

**支持的格式**:
```
#RRGGBB    // 6 位格式
#RGB       // 3 位简写格式（可选支持）
```

**示例**:
```json
{
  "color": "#FF0000",
  "highlightColor": "#FFFF00"
}
```

### 大小写

颜色值不区分大小写，但建议使用大写字母。

## 数值单位

### 字号

字号使用 **磅 (point)** 为单位。

```json
{
  "fontSize": 12
}
```

### 位置和尺寸

PPT 中的位置和尺寸使用 **磅 (point)** 为单位。

```json
{
  "position": {
    "left": 100,
    "top": 50,
    "width": 200,
    "height": 150
  }
}
```

**换算关系**:
- 1 英寸 = 72 磅
- 1 厘米 ≈ 28.35 磅

### 像素

图片尺寸使用 **像素 (pixel)** 为单位。

```json
{
  "image": {
    "width": 800,
    "height": 600
  }
}
```

## 可选字段

### 空值处理

可选字段为空时：
- 推荐：**省略字段**
- 可接受：设置为 `null`
- 不推荐：设置为空字符串 `""`

**推荐**:
```json
{
  "text": "Hello",
  "format": {
    "bold": true
  }
}
```

**可接受**:
```json
{
  "text": "Hello",
  "format": {
    "bold": true,
    "italic": null
  }
}
```

**不推荐**:
```json
{
  "text": "Hello",
  "format": {
    "bold": true,
    "fontName": ""
  }
}
```

### 默认值

当可选字段省略时，使用文档中指定的默认值。各事件定义中会说明默认值。

## 数组

### 空数组

空数组使用 `[]`，不使用 `null`。

**正确**:
```json
{
  "styles": []
}
```

**不推荐**:
```json
{
  "styles": null
}
```

### 索引

数组索引从 **0** 开始。

```json
{
  "slideIndex": 0,
  "selectIndex": 0
}
```

## 布尔值

布尔值使用 JSON 原生的 `true` / `false`。

**正确**:
```json
{
  "isEmpty": true,
  "matchCase": false
}
```

**错误**:
```json
{
  "isEmpty": "true",
  "matchCase": 0
}
```

## 超时约定

### 默认超时

| 操作类型 | 默认超时 |
|----------|----------|
| 简单查询 | 10 秒 |
| 复杂查询 | 30 秒 |
| 修改操作 | 30 秒 |
| 批量操作 | 60 秒 |

### 超时处理

1. Server 端应在超时后标记请求为失败
2. 不进行自动重试
3. 返回 `TIMEOUT` 错误码
4. AddIn 收到超时后的响应应忽略

## 版本兼容

### 向后兼容原则

1. **新增字段**: 可以添加新的可选字段
2. **新增事件**: 可以添加新的事件类型
3. **新增枚举值**: 可以添加新的枚举值

### 不兼容变更

以下变更需要升级协议主版本号：

1. 删除或重命名字段
2. 更改字段类型
3. 更改字段语义
4. 删除事件或错误码
