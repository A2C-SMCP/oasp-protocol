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

**事件数据**:

```typescript
interface DocumentModifiedEvent {
  eventType: "documentModified";  // 事件类型标识
  clientId: string;               // 客户端标识
  documentUri: string;            // 文档 URI
  timestamp: number;              // 事件发生时间（毫秒）
  data: {
    modificationType: "insert" | "delete" | "update";  // 修改类型
  };
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `eventType` | string | ✅ | 固定值 `"documentModified"`，用于事件类型识别 |
| `clientId` | string | ✅ | 客户端唯一标识，用于区分多客户端场景 |
| `documentUri` | string | ✅ | 文档 URI（如 `file:///path/to/doc.docx`） |
| `timestamp` | number | ✅ | Unix 时间戳（毫秒） |
| `data.modificationType` | string | ✅ | 修改类型：`"insert"` 插入、`"delete"` 删除、`"update"` 更新 |

**示例**:

```json
{
  "eventType": "documentModified",
  "clientId": "word-addin-abc123",
  "documentUri": "file:///Users/john/Documents/report.docx",
  "timestamp": 1704067200000,
  "data": {
    "modificationType": "insert"
  }
}
```

---

## 内容检索类

### word:get:selection

**方向**: Server → AddIn（请求-响应）

**状态**: ✅ Stable

**说明**: 获取当前选区的位置信息（轻量级查询，不含完整内容）。

**请求数据**:

```typescript
interface GetSelectionRequest {
  requestId: string;      // 请求 ID (UUID)
  documentUri: string;    // 文档 URI
  timestamp?: number;     // 请求时间戳（毫秒），可选
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
  success: boolean;
  data?: SelectionInfo;
  error?: ErrorResponse;
  timestamp: number;
}

interface SelectionInfo {
  /** 选区是否为空（光标点） */
  isEmpty: boolean;
  /** 选区类型 */
  type: "NoSelection" | "InsertionPoint" | "Normal";
  /** 起始位置（字符偏移量），选区非空时存在 */
  start?: number;
  /** 结束位置（字符偏移量），选区非空时存在 */
  end?: number;
  /** 选区文本，选区非空时存在 */
  text?: string;
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `isEmpty` | boolean | ✅ | 选区是否为空（光标点或无选区） |
| `type` | string | ✅ | 选区类型：`NoSelection`、`InsertionPoint`、`Normal` |
| `start` | number | ❌ | 起始位置（字符偏移量），仅选区非空时存在 |
| `end` | number | ❌ | 结束位置（字符偏移量），仅选区非空时存在 |
| `text` | string | ❌ | 选区文本，仅选区非空时存在 |

**选区类型说明**:

| 类型 | 说明 |
|------|------|
| `NoSelection` | 文档中没有活动选区 |
| `InsertionPoint` | 光标处于一个点（`start === end`） |
| `Normal` | 有文本被选中 |

**响应示例（成功 - 有选区）**:

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
  "timestamp": 1704067200500
}
```

**响应示例（成功 - 光标点）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "isEmpty": true,
    "type": "InsertionPoint",
    "start": 100,
    "end": 100
  },
  "timestamp": 1704067200500
}
```

**响应示例（成功 - 无选区）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "isEmpty": true,
    "type": "NoSelection"
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |

---

### word:get:selectedContent

**方向**: Server → AddIn（请求-响应）

**状态**: ✅ Stable

**说明**: 获取选中区域的完整内容，包括文本、段落、表格、图片、内容控件等元素。

**请求数据**:

```typescript
interface GetSelectedContentRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  options?: GetContentOptions;
}

interface GetContentOptions {
  includeText?: boolean;            // 是否包含文本，默认 true
  includeImages?: boolean;          // 是否包含图片，默认 true
  includeTables?: boolean;          // 是否包含表格，默认 true
  includeContentControls?: boolean; // 是否包含内容控件，默认 true
  detailedMetadata?: boolean;       // 是否包含详细元数据，默认 false
  maxTextLength?: number;           // 文本最大长度，超出截断
}
```

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/report.docx",
  "options": {
    "includeText": true,
    "detailedMetadata": true
  }
}
```

**响应数据**:

```typescript
interface GetSelectedContentResponse {
  requestId: string;
  success: boolean;
  data?: ContentInfo;
  error?: ErrorResponse;
  timestamp: number;
}

interface ContentInfo {
  text: string;                    // 纯文本内容
  elements: ContentElement[];      // 内容元素数组
  metadata?: ContentMetadata;      // 统计元数据
}

interface ContentMetadata {
  isEmpty: boolean;
  characterCount: number;
  paragraphCount: number;
  tableCount?: number;
  imageCount?: number;
}

type ContentElement = ParagraphElement | TableElement | InlinePictureElement | ContentControlElement;
```

**元素类型定义**:

```typescript
interface ParagraphElement {
  id: string;
  type: "Paragraph";
  text?: string;
  style?: string;
  alignment?: string;
  // detailedMetadata=true 时返回以下字段
  firstLineIndent?: number;
  leftIndent?: number;
  rightIndent?: number;
  lineSpacing?: number;
  spaceBefore?: number;
  spaceAfter?: number;
  isListItem?: boolean;
}

interface TableElement {
  id: string;
  type: "Table";
  rowCount: number;
  columnCount: number;
  cells?: TableCellInfo[][];
}

interface InlinePictureElement {
  id: string;
  type: "InlinePicture";
  width: number;
  height: number;
  altText?: string;
  hyperlink?: string;
}

interface ContentControlElement {
  id: string;
  type: "ContentControl";
  text?: string;
  title?: string;
  tag?: string;
  controlType: string;
  cannotDelete?: boolean;
  cannotEdit?: boolean;
  placeholderText?: string;
}
```

**响应示例（有选区）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "text": "Hello World\nThis is a paragraph.",
    "elements": [
      { "id": "p-0", "type": "Paragraph", "text": "Hello World", "style": "Normal" },
      { "id": "p-1", "type": "Paragraph", "text": "This is a paragraph." }
    ],
    "metadata": {
      "isEmpty": false,
      "characterCount": 32,
      "paragraphCount": 2
    }
  },
  "timestamp": 1704067200500
}
```

**响应示例（空选区）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "text": "",
    "elements": [],
    "metadata": { "isEmpty": true, "characterCount": 0, "paragraphCount": 0 }
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |

---

### word:get:visibleContent

**方向**: Server → AddIn（请求-响应）

**状态**: ✅ Stable

**说明**: 获取当前视口中可见的内容，包括文本、段落、表格、图片、内容控件等元素。

**请求数据**:

```typescript
interface GetVisibleContentRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  options?: GetContentOptions;  // 与 word:get:selectedContent 相同
}
```

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/report.docx",
  "options": {
    "includeText": true,
    "detailedMetadata": false
  }
}
```

**响应数据**:

```typescript
interface GetVisibleContentResponse {
  requestId: string;
  success: boolean;
  data?: VisibleContentInfo;
  error?: ErrorResponse;
  timestamp: number;
}

interface VisibleContentInfo {
  text: string;                         // 可见区域纯文本（页面间用 \n\n 分隔）
  elements: VisibleContentElement[];    // 内容元素数组
  metadata?: ContentMetadata;           // 统计元数据
}

interface VisibleContentElement {
  type: "text" | "image" | "table" | "other";  // 元素类型（映射后）
  content: ContentElement;                      // 原始元素内容
}
```

**元素类型映射**:

| Word 原始类型 | 协议 type 值 |
|---------------|--------------|
| Paragraph | `"text"` |
| InlinePicture | `"image"` |
| Table | `"table"` |
| ContentControl | `"other"` |

**响应示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "text": "Hello World\n\nThis is visible content.",
    "elements": [
      { "type": "text", "content": { "id": "para-1-0", "type": "Paragraph", "text": "Hello World" } },
      { "type": "text", "content": { "id": "para-1-1", "type": "Paragraph", "text": "This is visible content." } }
    ],
    "metadata": {
      "isEmpty": false,
      "characterCount": 36,
      "paragraphCount": 2
    }
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |

!!! note "与 word:get:selectedContent 的关系"
    本事件与 `word:get:selectedContent` 使用相同的 `GetContentOptions` 和元素类型定义。

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
  timestamp?: number;
}
```

**响应数据**:

```typescript
interface GetDocumentStructureResponse {
  requestId: string;
  success: boolean;
  data?: DocumentStructureResult;
  error?: ErrorResponse;
  timestamp: number;
}

interface DocumentStructureResult {
  sectionCount: number;    // 章节数量
  paragraphCount: number;  // 段落数量
  tableCount: number;      // 表格数量
  imageCount: number;      // 图片数量
}
```

**响应示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "sectionCount": 4,
    "paragraphCount": 25,
    "tableCount": 3,
    "imageCount": 5
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |

---

### word:get:documentStats

**方向**: Server → AddIn（请求-响应）

**状态**: ✅ Stable

**说明**: 获取文档的字数统计。

**请求数据**:

```typescript
interface GetDocumentStatsRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
}
```

**响应数据**:

```typescript
interface GetDocumentStatsResponse {
  requestId: string;
  success: boolean;
  data?: DocumentStatsResult;
  error?: ErrorResponse;
  timestamp: number;
}

interface DocumentStatsResult {
  characterCount: number;           // 字符数（不含空格）
  characterCountWithSpaces: number; // 字符数（含空格）
  wordCount: number;                // 单词数
  paragraphCount: number;           // 段落数
  pageCount?: number;               // 页数（可选）
}
```

**响应示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "characterCount": 7200,
    "characterCountWithSpaces": 8500,
    "wordCount": 1500,
    "paragraphCount": 25,
    "pageCount": 8
  },
  "timestamp": 1704067200500
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
  timestamp?: number;
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
  success: boolean;
  data?: {
    styles: StyleInfo[];
  };
  error?: ErrorResponse;
  timestamp: number;
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
  "timestamp": 1704067200500
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
  "timestamp": 1704067200500
}
```

---

## 文本操作类

### word:insert:text

**方向**: Server → AddIn（请求-响应）

**状态**: ✅ Stable

**说明**: 在指定位置插入文本。

!!! important "样式优先级规则"
    当同时指定直接格式（如 `bold`、`fontSize`）和 `styleName` 时，**直接格式优先级高于样式名**。

    即：先应用 `styleName` 指定的样式，再覆盖应用直接格式属性。

**请求数据**:

```typescript
interface InsertTextRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  text: string;                              // 要插入的文本
  location?: "Cursor" | "Start" | "End";     // 插入位置，默认 "Cursor"
  format?: TextFormat;                       // 可选的格式设置
}

interface TextFormat {
  bold?: boolean;
  italic?: boolean;
  fontSize?: number;
  fontName?: string;
  color?: string;        // hex 颜色值，如 "#FF0000"
  underline?: string;    // 下划线类型，如 "Single", "Double", "None"
  styleName?: string;    // Word 样式名，如 "Heading 1", "Normal"
}
```

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/report.docx",
  "text": "这是新插入的文本",
  "location": "Cursor",
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
  success: boolean;
  data?: {
    inserted: boolean;
  };
  error?: ErrorResponse;
  timestamp: number;
}
```

**响应示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "inserted": true
  },
  "timestamp": 1704067200500
}
```

---

### word:replace:selection

**方向**: Server → AddIn（请求-响应）

**状态**: ✅ Stable

**说明**: 替换当前选中的内容。

!!! warning "前置条件"
    选区必须非空。如果选区为空，将返回错误码 `SELECTION_EMPTY` (3002)。

!!! important "格式优先级规则"
    - `format`（最高优先级）：包含直接格式属性和 `format.styleName`
    - `styleName`（仅在 `format` 未提供时使用）
    - 默认保持选区原有格式

**请求数据**:

```typescript
interface ReplaceSelectionRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  content: ReplaceContent;  // 替换内容
}

interface ReplaceContent {
  text?: string;            // 替换文本
  images?: ImageData[];     // 替换图片（可插入多张）
  format?: TextFormat;      // 文本格式（最高优先级）
  styleName?: string;       // Word 样式名（仅在 format 未提供时使用）
}
```

**请求示例（文本替换）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/report.docx",
  "content": {
    "text": "新的替换文本",
    "format": {
      "bold": true
    }
  }
}
```

**请求示例（含图片替换）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/report.docx",
  "content": {
    "text": "替换文本",
    "images": [
      {
        "base64": "data:image/png;base64,iVBORw0...",
        "width": 200,
        "height": 150,
        "altText": "示例图片"
      }
    ],
    "styleName": "Heading 1"
  }
}
```

**响应数据**:

```typescript
interface ReplaceSelectionResponse {
  requestId: string;
  success: boolean;
  data?: {
    replaced: boolean;       // 是否成功替换
    characterCount: number;  // 替换后的字符数
  };
  error?: ErrorResponse;
  timestamp: number;
}
```

**响应示例（成功）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "replaced": true,
    "characterCount": 6
  },
  "timestamp": 1704067200500
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
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3002 | `SELECTION_EMPTY` - 选区为空 |
| 4001 | `VALIDATION_ERROR` - 请求参数校验失败 |
| 3999 | `OFFICE_API_ERROR` - Office API 调用错误 |

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
  timestamp?: number;
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
  success: boolean;
  data?: {
    replaceCount: number;    // 实际替换的数量
  };
  error?: ErrorResponse;
  timestamp: number;
}
```

**响应示例（成功）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "replaceCount": 5
  },
  "timestamp": 1704067200500
}
```

**响应示例（失败）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": false,
  "error": {
    "code": 3001,
    "message": "Document not found"
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4001 | `VALIDATION_ERROR` - Schema 校验失败 |
| 4002 | `MISSING_PARAM` - 缺少必要参数 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3999 | `OFFICE_API_ERROR` - 通用 Office API 错误（兜底） |

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
  timestamp?: number;
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
  success: boolean;
  data?: {
    matchCount: number;      // 总匹配数
    selectedIndex: number;   // 选中的是第几个
    selectedText: string;    // 选中的文本
  };
  error?: ErrorResponse;
  timestamp: number;
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
  "timestamp": 1704067200500
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
