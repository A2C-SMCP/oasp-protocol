# Word 事件定义

## 概述

本章定义 `/word` 命名空间下的所有事件。Word 事件用于操作 Microsoft Word 文档。

## 事件列表

### 事件报告类（AddIn → Server，单向）

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [word:event:selectionChanged](#wordeventselectionchanged) | ✅ Stable | 选区变化通知 |
| [word:event:documentModified](#wordeventdocumentmodified) | ✅ Stable | 文档修改通知 |

### 内容检索类（Server → AddIn，请求-响应）

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [word:get:selection](#wordgetselection) | ✅ Stable | 获取选区位置信息 |
| [word:get:selectedContent](#wordgetselectedcontent) | ✅ Stable | 获取选中内容（完整） |
| [word:get:visibleContent](#wordgetvisiblecontent) | ✅ Stable | 获取可见内容 |
| [word:get:documentStructure](#wordgetdocumentstructure) | ✅ Stable | 获取文档结构 |
| [word:get:documentStats](#wordgetdocumentstats) | ✅ Stable | 获取文档统计 |
| [word:get:styles](#wordgetstyles) | ✅ Stable | 获取文档样式列表 |

### 文本操作类（Server → AddIn，请求-响应）

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [word:insert:text](#wordinserttext) | ✅ Stable | 插入文本 |
| [word:replace:selection](#wordreplaceselection) | ✅ Stable | 替换选中内容 |
| [word:replace:text](#wordreplacetext) | ✅ Stable | 查找替换 |
| [word:select:text](#wordselecttext) | ✅ Stable | 查找并选中文本 |
| [word:append:text](#wordappendtext) | 📋 Draft | 追加文本 |

### 多媒体操作类（Server → AddIn，请求-响应）

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [word:insert:image](#wordinsertimage) | 📋 Draft | 插入图片 |
| [word:insert:table](#wordinserttable) | 📋 Draft | 插入表格 |
| [word:insert:equation](#wordinsertequation) | 📋 Draft | 插入公式 |

### 高级功能类（Server → AddIn，请求-响应）

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [word:insert:toc](#wordinserttoc) | 📋 Draft | 插入目录 |
| [word:export:content](#wordexportcontent) | 📋 Draft | 导出内容 |

---

## 事件报告类

### word:event:selectionChanged

**方向**: AddIn → Server（单向通知）

**状态**: ✅ Stable

**说明**: 当用户在 Word 中更改选区时触发。

**事件数据**:

```typescript
interface SelectionChangedEvent {
  eventType: "selectionChanged";  // 事件类型标识
  clientId: string;               // 客户端标识
  documentUri: string;            // 文档 URI
  timestamp: number;              // 事件发生时间（毫秒）
  data: {
    text: string;                 // 选中的文本内容
    length: number;               // 选中文本的长度
  };
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `eventType` | string | ✅ | 固定值 `"selectionChanged"`，用于事件类型识别 |
| `clientId` | string | ✅ | 客户端唯一标识，用于区分多客户端场景 |
| `documentUri` | string | ✅ | 文档 URI（如 `file:///path/to/doc.docx`） |
| `timestamp` | number | ✅ | Unix 时间戳（毫秒） |
| `data.text` | string | ✅ | 当前选中的文本内容，无选中时为空字符串 |
| `data.length` | number | ✅ | 选中文本的字符长度 |

**示例**:

```json
{
  "eventType": "selectionChanged",
  "clientId": "word-addin-abc123",
  "documentUri": "file:///Users/john/Documents/report.docx",
  "timestamp": 1704067200000,
  "data": {
    "text": "Hello World",
    "length": 11
  }
}
```

---

### word:event:documentModified

**方向**: AddIn → Server（单向通知）

**状态**: ✅ Stable

**说明**: 当文档内容被修改时触发。

**请求数据**:

```typescript
interface DocumentModifiedEvent {
  documentUri: string;   // 文档 URI
  timestamp: number;     // 事件发生时间（毫秒）
}
```

**示例**:

```json
{
  "documentUri": "file:///Users/john/Documents/report.docx",
  "timestamp": 1704067200000
}
```

---

## 内容检索类

### word:get:selection

**方向**: Server → AddIn（请求-响应）

**状态**: ✅ Stable

**说明**: 获取当前选区的位置信息（轻量级）。

**请求数据**:

```typescript
interface GetSelectionRequest {
  requestId: string;     // 请求 ID (UUID)
  documentUri: string;   // 文档 URI
  timestamp: number;     // 请求时间戳（毫秒）
}
```

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/report.docx",
  "timestamp": 1704067200000
}
```

**响应数据**:

```typescript
interface GetSelectionResponse {
  requestId: string;
  success: true;
  data: SelectionInfo;
  timestamp: number;
  duration: number;      // 操作耗时（毫秒）
}
```

**响应示例（成功）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "isEmpty": false,
    "type": "Normal",
    "start": 100,
    "end": 150,
    "text": "Hello World"
  },
  "timestamp": 1704067200500,
  "duration": 50
}
```

---

### word:get:selectedContent

**方向**: Server → AddIn（请求-响应）

**状态**: ✅ Stable

**说明**: 获取选中区域的完整内容，包括格式信息。

**请求数据**:

```typescript
interface GetSelectedContentRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  options?: {
    includeFormatting?: boolean;  // 是否包含格式信息，默认 true
  };
}
```

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/report.docx",
  "timestamp": 1704067200000,
  "options": {
    "includeFormatting": true
  }
}
```

**响应数据**:

```typescript
interface GetSelectedContentResponse {
  requestId: string;
  success: true;
  data: {
    text: string;
    html?: string;       // HTML 格式内容
    format?: TextFormat; // 格式信息
  };
  timestamp: number;
  duration: number;
}
```

**响应示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "text": "Hello World",
    "html": "<p><b>Hello</b> World</p>",
    "format": {
      "bold": true,
      "fontSize": 12,
      "fontName": "Calibri"
    }
  },
  "timestamp": 1704067200500,
  "duration": 80
}
```

---

### word:get:visibleContent

**方向**: Server → AddIn（请求-响应）

**状态**: ✅ Stable

**说明**: 获取当前视口中可见的内容。

**请求数据**:

```typescript
interface GetVisibleContentRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
}
```

**响应数据**:

```typescript
interface GetVisibleContentResponse {
  requestId: string;
  success: true;
  data: {
    text: string;
    startPosition: number;
    endPosition: number;
  };
  timestamp: number;
  duration: number;
}
```

---

### word:get:documentStructure

**方向**: Server → AddIn（请求-响应）

**状态**: ✅ Stable

**说明**: 获取文档的结构统计信息。

**请求数据**:

```typescript
interface GetDocumentStructureRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
}
```

**响应数据**:

```typescript
interface GetDocumentStructureResponse {
  requestId: string;
  success: true;
  data: DocumentStructure;
  timestamp: number;
  duration: number;
}
```

**响应示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "paragraphCount": 25,
    "tableCount": 3,
    "imageCount": 5,
    "sectionCount": 4
  },
  "timestamp": 1704067200500,
  "duration": 120
}
```

---

### word:get:documentStats

**方向**: Server → AddIn（请求-响应）

**状态**: ✅ Stable

**说明**: 获取文档的字数统计。

**响应数据**:

```typescript
interface GetDocumentStatsResponse {
  requestId: string;
  success: true;
  data: DocumentStats;
  timestamp: number;
  duration: number;
}
```

**响应示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "wordCount": 1500,
    "characterCount": 8500,
    "paragraphCount": 25
  },
  "timestamp": 1704067200500,
  "duration": 150
}
```

---

### word:get:styles

**方向**: Server → AddIn（请求-响应）

**状态**: ✅ Stable

**说明**: 获取文档中可用的样式列表。

**请求数据**:

```typescript
interface GetStylesRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  options?: {
    includeBuiltIn?: boolean;  // 是否包含内置样式，默认 true
    includeCustom?: boolean;   // 是否包含自定义样式，默认 true
    includeUnused?: boolean;   // 是否包含未使用的样式，默认 false
    detailedInfo?: boolean;    // 是否返回详细信息（description），默认 false
  };
}
```

**请求参数说明**:

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `includeBuiltIn` | boolean | true | 是否包含 Word 内置样式 |
| `includeCustom` | boolean | true | 是否包含用户自定义样式 |
| `includeUnused` | boolean | false | 是否包含文档中未使用的样式。设为 false 时仅返回 inUse=true 的样式 |
| `detailedInfo` | boolean | false | 是否返回样式的详细描述。设为 true 时返回 description 字段（依赖 WordApi BETA，部分环境可能不可用） |

**响应数据**:

```typescript
interface GetStylesResponse {
  requestId: string;
  success: true;
  data: {
    styles: StyleInfo[];
  };
  timestamp: number;
  duration: number;
}
```

**响应示例**（默认参数，仅返回正在使用的样式）:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "styles": [
      {
        "name": "标题 1",
        "type": "Paragraph",
        "builtIn": true,
        "inUse": true
      },
      {
        "name": "正文",
        "type": "Paragraph",
        "builtIn": true,
        "inUse": true
      }
    ]
  },
  "timestamp": 1704067200500,
  "duration": 200
}
```

**响应示例**（`detailedInfo=true` 时返回 description 字段）:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "styles": [
      {
        "name": "标题 1",
        "type": "Paragraph",
        "builtIn": true,
        "inUse": true,
        "description": "用于主要章节标题"
      }
    ]
  },
  "timestamp": 1704067200500,
  "duration": 200
}
```

---

## 文本操作类

### word:insert:text

**方向**: Server → AddIn（请求-响应）

**状态**: ✅ Stable

**说明**: 在当前光标位置插入文本。

!!! important "样式优先级规则"
    当同时指定直接格式（如 `bold`、`fontSize`）和 `styleName` 时，**直接格式优先级高于样式名**。

    即：先应用 `styleName` 指定的样式，再覆盖应用直接格式属性。

**请求数据**:

```typescript
interface InsertTextRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  text: string;          // 要插入的文本
  format?: TextFormat;   // 可选的格式设置
}
```

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/report.docx",
  "timestamp": 1704067200000,
  "text": "这是新插入的文本",
  "format": {
    "bold": true,
    "fontSize": 14,
    "fontName": "微软雅黑",
    "color": "#FF0000"
  }
}
```

**响应数据**:

```typescript
interface InsertTextResponse {
  requestId: string;
  success: true;
  data: {
    insertedLength: number;  // 插入的字符数
  };
  timestamp: number;
  duration: number;
}
```

**响应示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "insertedLength": 8
  },
  "timestamp": 1704067200500,
  "duration": 100
}
```

---

### word:replace:selection

**方向**: Server → AddIn（请求-响应）

**状态**: ✅ Stable

**说明**: 替换当前选中的内容。

!!! warning "前置条件"
    选区必须非空。如果选区为空，将返回错误码 `SELECTION_EMPTY` (3002)。

**请求数据**:

```typescript
interface ReplaceSelectionRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  content: ReplaceContent;  // 替换内容
}
```

**请求示例（文本替换）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/report.docx",
  "timestamp": 1704067200000,
  "content": {
    "text": "新的替换文本",
    "format": {
      "bold": true
    }
  }
}
```

**响应示例（成功）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "replacedLength": 5,
    "newLength": 6
  },
  "timestamp": 1704067200500,
  "duration": 80
}
```

**响应示例（选区为空）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": false,
  "error": {
    "code": "SELECTION_EMPTY",
    "message": "Selection is empty, cannot replace"
  },
  "timestamp": 1704067200500,
  "duration": 10
}
```

---

### word:replace:text

**方向**: Server → AddIn（请求-响应）

**状态**: ✅ Stable

**说明**: 查找并替换文档中的文本。

**请求数据**:

```typescript
interface ReplaceTextRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  searchText: string;        // 要查找的文本
  replaceText: string;       // 替换为的文本
  options?: {
    matchCase?: boolean;     // 区分大小写，默认 false
    matchWholeWord?: boolean; // 全词匹配，默认 false
    replaceAll?: boolean;    // 替换全部，默认 false
  };
}
```

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/report.docx",
  "timestamp": 1704067200000,
  "searchText": "旧文本",
  "replaceText": "新文本",
  "options": {
    "matchCase": true,
    "replaceAll": true
  }
}
```

**响应数据**:

```typescript
interface ReplaceTextResponse {
  requestId: string;
  success: true;
  data: {
    matchCount: number;      // 找到的匹配数
    replacedCount: number;   // 实际替换的数量
  };
  timestamp: number;
  duration: number;
}
```

**响应示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "matchCount": 5,
    "replacedCount": 5
  },
  "timestamp": 1704067200500,
  "duration": 150
}
```

---

### word:select:text

**方向**: Server → AddIn（请求-响应）

**状态**: ✅ Stable

**说明**: 查找并选中文档中的文本。

**请求数据**:

```typescript
interface SelectTextRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  searchText: string;        // 要查找的文本
  options?: {
    selectionMode?: "select" | "start" | "end";  // 选择模式，默认 "select"
    selectIndex?: number;    // 选择第几个匹配项，默认 0（第一个）
  };
}
```

**选择模式说明**:

| 模式 | 说明 |
|------|------|
| `select` | 选中整个匹配文本 |
| `start` | 将光标移动到匹配文本的开头 |
| `end` | 将光标移动到匹配文本的末尾 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/report.docx",
  "timestamp": 1704067200000,
  "searchText": "目标文本",
  "options": {
    "selectionMode": "select",
    "selectIndex": 0
  }
}
```

**响应数据**:

```typescript
interface SelectTextResponse {
  requestId: string;
  success: true;
  data: {
    matchCount: number;      // 总匹配数
    selectedIndex: number;   // 选中的是第几个
    selectedText: string;    // 选中的文本
  };
  timestamp: number;
  duration: number;
}
```

**响应示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "matchCount": 3,
    "selectedIndex": 0,
    "selectedText": "目标文本"
  },
  "timestamp": 1704067200500,
  "duration": 100
}
```

---

### word:append:text

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 在文档末尾追加文本。

**请求数据**:

```typescript
interface AppendTextRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  text: string;
  format?: TextFormat;
}
```

---

## 多媒体操作类

### word:insert:image

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 在当前光标位置插入图片。

**请求数据**:

```typescript
interface InsertImageRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  image: ImageData;
}
```

---

### word:insert:table

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 在当前光标位置插入表格。

**请求数据**:

```typescript
interface InsertTableRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  options: TableInsertOptions;
}
```

---

### word:insert:equation

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 在当前光标位置插入数学公式。

**请求数据**:

```typescript
interface InsertEquationRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  latex: string;         // LaTeX 格式的公式
}
```

---

## 高级功能类

### word:insert:toc

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 在当前光标位置插入目录。

**请求数据**:

```typescript
interface InsertTOCRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  options?: {
    levels?: number;     // 包含的标题级别，默认 3
  };
}
```

---

### word:export:content

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 导出文档内容为指定格式。

**请求数据**:

```typescript
interface ExportContentRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  format: "text" | "html" | "markdown";
}
```

**响应数据**:

```typescript
interface ExportContentResponse {
  requestId: string;
  success: true;
  data: {
    content: string;     // 导出的内容
    format: string;      // 导出格式
  };
  timestamp: number;
  duration: number;
}
```
