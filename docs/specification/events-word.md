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
| [word:append:text](#wordappendtext) | ✅ Stable | 追加文本 |

### 多媒体操作类（Server → AddIn，请求-响应）

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [word:insert:image](#wordinsertimage) | ✅ Stable | 插入图片 |
| [word:insert:table](#wordinserttable) | ✅ Stable | 插入表格 |
| [word:insert:equation](#wordinsertequation) | ✅ Stable | 插入公式 |

### 表格操作类（Server → AddIn，请求-响应）

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [word:merge:cells](#wordmergecells) | 📋 Draft | 合并表格中任意矩形单元格区域 |
| [word:update:tableCell](#wordupdatetablecell) | 📋 Draft | 批量更新单元格文本与格式 |
| [word:update:tableRowColumn](#wordupdatetablerowcolumn) | 📋 Draft | 按行/列批量更新单元格内容 |
| [word:update:tableFormat](#wordupdatetableformat) | 📋 Draft | 更新整表样式、边框、列宽、对齐 |

### 高级功能类（Server → AddIn，请求-响应）

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [word:insert:toc](#wordinserttoc) | ✅ Stable | 插入目录 |
| [word:export:content](#wordexportcontent) | ✅ Stable | 导出内容 |

### 批注操作类（Server → AddIn，请求-响应）

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [word:get:comments](#wordgetcomments) | ✅ Stable | 获取批注列表 |
| [word:insert:comment](#wordinsertcomment) | ✅ Stable | 在选区上插入批注 |
| [word:delete:comment](#worddeletecomment) | ✅ Stable | 按 ID 删除批注 |
| [word:reply:comment](#wordreplycomment) | ✅ Stable | 回复已有批注 |
| [word:resolve:comment](#wordresolvecomment) | ✅ Stable | 解决/取消解决批注 |

### 脚本执行类（Server → AddIn，请求-响应）

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [word:run:script](#wordrunscript) | 📋 Draft | 执行原始 Office.js 脚本（封装层逃生舱） |

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
  tables?: TableSummary[]; // 表格清单（可选，用于"重新发现"现有表格）
}
```

`TableSummary` 详见 [data-structures.md#documentstructureresult](data-structures.md#documentstructureresult)。

!!! note "tables 字段"
    可选字段。若 Add-In 提供，则按文档中表格出现顺序返回；调用方可基于 `precedingHeading` 启发式找到目标表格，再以显式 `tableId` 调用 [`word:update:tableCell`](#wordupdatetablecell) 等表格操作事件，避免依赖 `tableId` 临时索引的脆弱性。

**响应示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "sectionCount": 4,
    "paragraphCount": 25,
    "tableCount": 3,
    "imageCount": 5,
    "tables": [
      { "tableId": "table-0", "rowCount": 3, "columnCount": 4, "precedingHeading": "甲方信息" },
      { "tableId": "table-1", "rowCount": 5, "columnCount": 3, "precedingHeading": "车辆信息" },
      { "tableId": "table-2", "rowCount": 2, "columnCount": 2 }
    ]
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
    当同时指定 `format.font`（[`WordFont`](data-structures.md#wordfont)）和 `styleName` 时，**`font` 优先级高于样式名**。

    即：先应用 `styleName` 指定的样式，再覆盖应用 `format.font`。

**请求数据**:

```typescript
interface InsertTextRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  text: string;                              // 要插入的文本
  location?: "Cursor" | "Start" | "End";     // 插入位置，默认 "Cursor"
  format?: TextFormat;                       // 可选的格式设置，定义见 data-structures.md#textformat
}
```

`TextFormat` 数据结构定义见 [data-structures.md#textformat](data-structures.md#textformat)（权威单一来源；`underline` 为 [`UnderlineStyle`](data-structures.md#underlinestyle) 枚举）。

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/report.docx",
  "text": "这是新插入的文本",
  "location": "Cursor",
  "format": {
    "font": { "bold": true, "size": 14, "name": "微软雅黑", "color": "#FF0000" }
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
    - `format`（最高优先级）：包含 `format.font`（字体）和 `format.styleName`
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
      "font": { "bold": true }
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
| 4000 | `VALIDATION_ERROR` - 请求参数校验失败 |
| 3000 | `DOCUMENT_ERROR` - 文档操作错误（通用） |

---

### word:replace:text

**方向**: Server → AddIn（请求-响应）

**状态**: ✅ Stable

**说明**: 查找并替换文档中的文本。

!!! important "样式优先级规则"
    当同时指定 `format.font`（[`WordFont`](data-structures.md#wordfont)）和 `styleName` 时，**`font` 优先级高于样式名**。
    即：先应用 `styleName` 指定的样式，再覆盖应用 `format.font`。

!!! note "特殊字符搜索"
    `searchText` 和 `replaceText` 支持 Word 特殊字符记号来匹配不可见字符。**不要使用 `\n` 等转义字符**，而应使用下表中的 Word 记号：

    | 要查找/替换 | 记号 | 说明 |
    |------------|------|------|
    | 段落标记 (Enter) | `^p` | 跨段落搜索时使用 |
    | 手动换行 (Shift+Enter) | `^l` | 段内软换行 |
    | 制表符 | `^t` | Tab 字符 |

    示例：搜索跨两个段落的文本 `"line1^pline2"`，替换为 `"new^pcontent"`。
    完整特殊字符列表参见 [Word Search Options Guidance](https://learn.microsoft.com/en-us/office/dev/add-ins/word/search-option-guidance)。

**请求数据**:

```typescript
interface ReplaceTextRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  searchText: string;        // 要查找的文本（支持 ^p ^l ^t 等特殊字符记号）
  replaceText: string;       // 替换为的文本（同样支持特殊字符记号）
  options?: {
    matchCase?: boolean;     // 区分大小写，默认 false
    matchWholeWord?: boolean; // 全词匹配，默认 false
    replaceAll?: boolean;    // 替换全部，默认 false
  };
  format?: TextFormat;       // 可选的格式设置（定义见 data-structures.md#textformat）
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
  },
  "format": {
    "font": { "bold": true, "size": 14, "color": "#FF0000" }
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
| 4000 | `VALIDATION_ERROR` - Schema 校验失败 |
| 4001 | `MISSING_PARAM` - 缺少必要参数 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3000 | `DOCUMENT_ERROR` - 文档操作错误（通用，兜底） |

---

### word:select:text

**方向**: Server → AddIn（请求-响应）

**状态**: ✅ Stable

**说明**: 查找并选中文档中的文本。

!!! note "特殊字符搜索"
    `searchText` 支持 Word 特殊字符记号（如 `^p` 段落标记、`^l` 手动换行、`^t` 制表符）来匹配不可见字符。**不要使用 `\n` 等转义字符**。详见 [word:replace:text](#wordreplacetext) 的特殊字符说明。

**请求数据**:

```typescript
interface SelectTextRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  searchText: string;                              // 要查找的文本（支持 ^p ^l ^t 等特殊字符记号）
  searchOptions?: {
    matchCase?: boolean;                           // 区分大小写，默认 false
    matchWholeWord?: boolean;                      // 全词匹配，默认 false
    matchWildcards?: boolean;                      // 通配符匹配，默认 false
  };
  selectionMode?: "select" | "start" | "end";      // 选择模式，默认 "select"
  selectIndex?: number;                            // 选择第几个匹配项（1-based），默认 1
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
  "searchText": "目标文本",
  "searchOptions": {
    "matchCase": true
  },
  "selectionMode": "select",
  "selectIndex": 1
}
```

**响应数据**:

```typescript
interface SelectTextResponse {
  requestId: string;
  success: boolean;
  data?: SelectTextResult;
  error?: ErrorResponse;
  timestamp: number;
}

interface SelectTextResult {
  success: boolean;              // 是否找到并选中了文本
  matchCount: number;            // 总匹配数
  selectedIndex: number;         // 选中的是第几个（1-based）
  selectedText: string;          // 选中的文本
  selectionInfo?: {              // 选中后的选区详细信息
    type: "Normal" | "NoSelection" | "InsertionPoint";  // 选区类型
    start?: number;              // 起始位置
    end?: number;                // 结束位置
    text?: string;               // 选区文本
    isEmpty: boolean;            // 是否为空
  };
}
```

**响应示例（成功）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "success": true,
    "matchCount": 3,
    "selectedIndex": 1,
    "selectedText": "目标文本",
    "selectionInfo": {
      "type": "Normal",
      "start": 100,
      "end": 104,
      "text": "目标文本",
      "isEmpty": false
    }
  },
  "timestamp": 1704067200500
}
```

**响应示例（未找到匹配）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "success": false,
    "matchCount": 0,
    "selectedIndex": 0,
    "selectedText": ""
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4000 | `VALIDATION_ERROR` - 请求参数校验失败 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3000 | `DOCUMENT_ERROR` - 文档操作错误（通用） |

---

### word:append:text

**方向**: Server → AddIn（请求-响应）

**状态**: ✅ Stable

**说明**: 在文档开头或末尾追加文本。

**请求数据**:

```typescript
interface AppendTextRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  text: string;                              // 要追加的文本
  location?: "Start" | "End";                // 追加位置，默认 "End"
  format?: TextFormat;                       // 可选的格式设置
}
```

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/report.docx",
  "text": "这是追加的文本",
  "location": "End",
  "format": {
    "font": { "bold": true, "size": 12 }
  }
}
```

**响应数据**:

```typescript
interface AppendTextResponse {
  requestId: string;
  success: boolean;
  data?: {
    appended: boolean;    // 是否成功追加
    length: number;       // 追加的字符数
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
    "appended": true,
    "length": 7
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4000 | `VALIDATION_ERROR` - 请求参数校验失败 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3000 | `DOCUMENT_ERROR` - 文档操作错误（通用） |

---

## 多媒体操作类

### word:insert:image

**方向**: Server → AddIn（请求-响应）

**状态**: ✅ Stable

**说明**: 在文档中插入图片，支持内联和浮动布局。

**请求数据**:

```typescript
interface InsertImageRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  image: ImageData;                          // 图片数据
  location?: InsertLocation;                 // 插入位置
  wrapType?: "Inline" | "Square" | "Tight" | "Behind" | "InFront";  // 文字环绕方式
}

interface ImageData {
  base64: string;       // Base64 编码的图片数据
  width?: number;       // 图片宽度（磅）
  height?: number;      // 图片高度（磅）
  altText?: string;     // 替代文本
}

interface InsertLocation {
  type: "Cursor" | "Start" | "End" | "BeforeBookmark" | "AfterBookmark";
  bookmarkName?: string;  // 当 type 为 BeforeBookmark/AfterBookmark 时必需
}
```

**文字环绕方式说明**:

| 类型 | 说明 |
|------|------|
| `Inline` | 嵌入型（默认），图片作为文字的一部分 |
| `Square` | 四周型环绕 |
| `Tight` | 紧密型环绕 |
| `Behind` | 衬于文字下方 |
| `InFront` | 浮于文字上方 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/report.docx",
  "image": {
    "base64": "data:image/png;base64,iVBORw0KGgoAAAANS...",
    "width": 200,
    "height": 150,
    "altText": "示例图片"
  },
  "location": {
    "type": "Cursor"
  },
  "wrapType": "Square"
}
```

**响应数据**:

```typescript
interface InsertImageResponse {
  requestId: string;
  success: boolean;
  data?: {
    inserted: boolean;   // 是否成功插入
    imageId: string;     // 图片标识符
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
    "inserted": true,
    "imageId": "shape-12345"
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4000 | `VALIDATION_ERROR` - 请求参数校验失败 |
| 4002 | `INVALID_PARAM` - `image.base64` 无法解码（线缆层） |
| 3007 | `FORMAT_NOT_SUPPORTED` - 图片可解码但格式不受支持（见[判法划界](error-handling.md#format_not_supported-3007)） |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3000 | `DOCUMENT_ERROR` - 文档操作错误（通用） |

---

### word:insert:table

**方向**: Server → AddIn（请求-响应）

**状态**: ✅ Stable

**说明**: 在文档中插入表格。

**请求数据**:

```typescript
interface InsertTableRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  options: TableInsertOptions;
}

interface TableInsertOptions {
  rows: number;              // 行数（必需）
  columns: number;           // 列数（必需）
  data?: string[][];         // 表格数据（可选，按行列顺序填充）
  style?: string;            // 表格样式名称
  insertLocation?: TableInsertLocation;  // 插入位置，默认 "End"
}

type TableInsertLocation =
  | "Start"    // 文档开头
  | "End"      // 文档末尾（默认）
  | "Before"   // 当前选区/光标之前
  | "After"    // 当前选区/光标之后
  | "Replace"; // 替换当前选区
```

**插入位置语义**:

| 值 | 语义 |
|----|------|
| `Start` | 插入到文档起始位置 |
| `End` | 插入到文档末尾（默认，兼容旧客户端未传此字段的行为）|
| `Before` | 插入到当前选区/光标所在段落之前 |
| `After` | 插入到当前选区/光标所在段落之后 |
| `Replace` | 以新表格替换当前选区内容 |

未传 `insertLocation` 时，Add-In 必须回退至 `"End"`（向后兼容）。

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/report.docx",
  "options": {
    "rows": 3,
    "columns": 4,
    "data": [
      ["姓名", "年龄", "城市", "职业"],
      ["张三", "28", "北京", "工程师"],
      ["李四", "32", "上海", "设计师"]
    ],
    "style": "Grid Table 1 Light",
    "insertLocation": "After"
  }
}
```

**响应数据**:

```typescript
interface InsertTableResponse {
  requestId: string;
  success: boolean;
  data?: {
    tableId: string;       // 表格标识符（详见下方 tableId 稳定性说明）
    rowCount: number;      // 行数
    columnCount: number;   // 列数
  };
  error?: ErrorResponse;
  timestamp: number;
}
```

!!! warning "tableId 稳定性"
    当前 `tableId`（如 `"table-0"`）为**临时索引**，按文档中表格的当前顺序生成。其稳定性依赖文档表格顺序不变——如果在它之前插入或删除其它表格，旧的 `tableId` 会失效或指向错误对象。**跨会话或经过结构变更的场景，调用方应在使用前通过 [`word:get:documentStructure`](#wordgetdocumentstructure) 重新发现表格列表。**协议未来计划提供基于 `Content Control` 的稳定 ID 方案，届时会在不破坏现有响应结构的前提下追加新字段。

**响应示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "tableId": "table-0",
    "rowCount": 3,
    "columnCount": 4
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4000 | `VALIDATION_ERROR` - 请求参数校验失败 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3000 | `DOCUMENT_ERROR` - 文档操作错误（通用） |

---

### word:insert:equation

**方向**: Server → AddIn（请求-响应）

**状态**: ✅ Stable

**说明**: 在文档中插入数学公式（LaTeX 格式）。

**请求数据**:

```typescript
interface InsertEquationRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  latex: string;                 // LaTeX 格式的公式
  options?: {
    inline?: boolean;            // 是否内联显示，默认 false
  };
}
```

**支持的 LaTeX 语法**:

| 语法 | 说明 | 示例 |
|------|------|------|
| `^{}` | 上标 | `x^{2}` → x² |
| `_{}` | 下标 | `x_{i}` → xᵢ |
| `\frac{}{}` | 分数 | `\frac{a}{b}` → a/b |
| `\sqrt{}` | 平方根 | `\sqrt{x}` → √x |
| `\sum_{}^{}` | 求和 | `\sum_{i=1}^{n}` |
| `\int_{}^{}` | 积分 | `\int_{0}^{1}` |
| 希腊字母 | α, β, γ 等 | `\alpha`, `\beta` |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/report.docx",
  "latex": "E = mc^{2}",
  "options": {
    "inline": true
  }
}
```

**响应数据**:

```typescript
interface InsertEquationResponse {
  requestId: string;
  success: boolean;
  data?: {
    equationId: string;    // 公式标识符
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
    "equationId": "eq-001"
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4000 | `VALIDATION_ERROR` - 请求参数校验失败 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3000 | `DOCUMENT_ERROR` - 文档操作错误（通用） |

---

## 表格操作类

本节定义在 [`word:insert:table`](#wordinserttable) 已创建表格之上的精细操作能力，覆盖合并单元格、单元格内容/格式批量更新、整表样式调整等场景。

**通用约定**：

- `tableId` 与 [`word:insert:table`](#wordinserttable) 响应中的 `tableId`（如 `"table-0"`）一致；缺省时表示当前光标所在表格。
- 行列索引均从 `0` 开始。
- `CellFormat` 数据结构定义见 [data-structures.md#cellformat](data-structures.md#cellformat)。

### word:merge:cells

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 将表格中由 `[startRowIndex, startColumnIndex]` 与 `[endRowIndex, endColumnIndex]` 圈定的矩形区域合并为单个单元格。

**请求数据**:

```typescript
interface MergeCellsRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  tableId?: string;          // 表格标识符；缺省时取当前光标所在表格
  startRowIndex: number;     // 矩形区域起始行（含）
  startColumnIndex: number;  // 矩形区域起始列（含）
  endRowIndex: number;       // 矩形区域结束行（含）
  endColumnIndex: number;    // 矩形区域结束列（含）
}
```

!!! note "矩形区域要求"
    起止索引必须满足 `start <= end`，且整个矩形必须落在表格范围内。合并后由原区域左上角单元格承载所有内容，其它原单元格内容会被丢弃。

**请求示例**（合并首行 4 列形成横跨表头）:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/contract.docx",
  "tableId": "table-0",
  "startRowIndex": 0,
  "startColumnIndex": 0,
  "endRowIndex": 0,
  "endColumnIndex": 3
}
```

**响应数据**:

```typescript
interface MergeCellsResponse {
  requestId: string;
  success: boolean;
  data?: {
    tableId: string;
    requestedRange: {
      rowCount: number;       // 矩形行跨度 = endRowIndex - startRowIndex + 1
      columnCount: number;    // 矩形列跨度 = endColumnIndex - startColumnIndex + 1
    };
  };
  error?: ErrorResponse;
  timestamp: number;
}
```

!!! note "为何返回 requestedRange 而非"实际被合并的原子单元格数""
    Word.js 不暴露"被合并掉的原始单元格总数"——若区域内已存在跨多列/多行的合并单元格，按 `rowCount × columnCount` 朴素相乘会失真。返回请求矩形的尺寸语义清晰、可验证，不强迫 Add-In 做无意义的扫描。

**响应示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "tableId": "table-0",
    "requestedRange": { "rowCount": 1, "columnCount": 4 }
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4001 | `MISSING_PARAM` - 缺少必需字段 |
| 4002 | `INVALID_PARAM` - 行/列索引超出表格范围或矩形非法（`start > end`） |
| 3010 | `ELEMENT_NOT_FOUND` - 指定的 `tableId` 在文档中未找到 |
| 3013 | `NO_TABLE_AT_CURSOR` - 缺省 `tableId` 时光标未在任何表格内 |
| 3014 | `ALREADY_MERGED` - 目标矩形与现有合并单元格冲突 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3000 | `DOCUMENT_ERROR` - 文档操作错误（通用） |

---

### word:update:tableCell

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 批量更新指定单元格的文本与/或格式。`text` 与 `format` 至少传一个；都不传的条目视为参数错误。

**请求数据**:

```typescript
interface UpdateTableCellRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  tableId?: string;          // 缺省时取当前光标所在表格
  cells: Array<{
    rowIndex: number;
    columnIndex: number;
    text?: string;           // 新文本内容（可选）
    format?: CellFormat;     // 单元格格式（可选）
  }>;
}
```

`CellFormat` 详见 [data-structures.md#cellformat](data-structures.md#cellformat)。

**请求示例**（首行表头：蓝底白字居中）:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/contract.docx",
  "tableId": "table-0",
  "cells": [
    {
      "rowIndex": 0,
      "columnIndex": 0,
      "text": "甲方信息",
      "format": {
        "horizontalAlignment": "Centered",
        "backgroundColor": "#4472C4",
        "font": { "color": "#FFFFFF", "bold": true }
      }
    }
  ]
}
```

**响应数据**:

```typescript
interface UpdateTableCellResponse {
  requestId: string;
  success: boolean;
  data?: {
    tableId: string;
    cellsUpdated: number;    // 成功更新的单元格数量
    rowCount: number;        // 表格总行数
    columnCount: number;     // 表格总列数
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
    "tableId": "table-0",
    "cellsUpdated": 1,
    "rowCount": 5,
    "columnCount": 4
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4001 | `MISSING_PARAM` - 缺少 `cells`，或某条目 `text` 与 `format` 均未提供 |
| 4002 | `INVALID_PARAM` - `rowIndex` / `columnIndex` 超出范围或 `format` 字段值不合法 |
| 3010 | `ELEMENT_NOT_FOUND` - 指定的 `tableId` 在文档中未找到 |
| 3013 | `NO_TABLE_AT_CURSOR` - 缺省 `tableId` 时光标未在任何表格内 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3000 | `DOCUMENT_ERROR` - 文档操作错误（通用） |

---

### word:update:tableRowColumn

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 按行或按列批量写入文本，适用于一次性填充整行或整列的数据。`rows` 与 `columns` 至少提供一个。

**请求数据**:

```typescript
interface UpdateTableRowColumnRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  tableId?: string;          // 缺省时取当前光标所在表格
  rows?: Array<{
    rowIndex: number;
    values: string[];        // 该行各列的文本值，按列顺序排列
  }>;
  columns?: Array<{
    columnIndex: number;
    values: string[];        // 该列各行的文本值，按行顺序排列
  }>;
}
```

!!! note "行列同时提供"
    当 `rows` 与 `columns` 同时提供时，先处理 `rows`，再处理 `columns`；后者会覆盖前者对相同单元格的写入。

!!! note "values 长度"
    `values` 长度应等于目标行/列的列/行数；过短则只覆盖前缀单元格，过长则触发 `INVALID_PARAM`。

**请求示例**（覆盖前两行）:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/contract.docx",
  "tableId": "table-0",
  "rows": [
    { "rowIndex": 0, "values": ["姓名", "年龄", "城市", "职业"] },
    { "rowIndex": 1, "values": ["张三", "28", "北京", "工程师"] }
  ]
}
```

**响应数据**:

```typescript
interface UpdateTableRowColumnResponse {
  requestId: string;
  success: boolean;
  data?: {
    tableId: string;
    cellsUpdated: number;    // 成功更新的单元格总数
    rowCount: number;
    columnCount: number;
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
    "tableId": "table-0",
    "cellsUpdated": 8,
    "rowCount": 5,
    "columnCount": 4
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4001 | `MISSING_PARAM` - `rows` 与 `columns` 均未提供 |
| 4002 | `INVALID_PARAM` - `rowIndex` / `columnIndex` 超出范围，或 `values` 长度超过表格列/行数 |
| 3010 | `ELEMENT_NOT_FOUND` - 指定的 `tableId` 在文档中未找到 |
| 3013 | `NO_TABLE_AT_CURSOR` - 缺省 `tableId` 时光标未在任何表格内 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3000 | `DOCUMENT_ERROR` - 文档操作错误（通用） |

---

### word:update:tableFormat

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 更新整表样式（内置/自定义样式名）、边框、列宽、整体对齐等表级属性。所有字段均为可选，按需传递。

**请求数据**:

```typescript
interface UpdateTableFormatRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  tableId?: string;            // 缺省时取当前光标所在表格
  styleOptions?: {
    styleType?: string;        // 内置或自定义表格样式名（如 "Grid Table 1 Light"）
    firstColumn?: boolean;     // 突出首列
    lastColumn?: boolean;      // 突出末列
    totalRow?: boolean;        // 启用汇总行
    bandedRows?: boolean;      // 行间隔条带
    bandedColumns?: boolean;   // 列间隔条带
    cellPadding?: {            // 表级单元格内边距（磅，对应 Word.Table.setCellPadding）
      top?: number;
      bottom?: number;
      left?: number;
      right?: number;
    };
  };
  borderOptions?: {
    location?: "all" | "inside" | "outside";  // 应用范围，默认 "all"
    style?: "Single" | "Double" | "Dashed" | "Dotted" | "None";
    width?: number;            // 边框宽度（磅，建议 0.25–6）
    color?: string;            // 边框颜色（十六进制）
  };
  columnWidths?: number[];     // 每列宽度（磅）；长度建议等于表格列数（以未合并状态计）
  alignment?: "Left" | "Centered" | "Right";  // 表格在页面中的水平对齐（对应 Word.Alignment）
}
```

!!! note "cellPadding 为表级"
    Word JavaScript API 仅提供表级 `Word.Table.setCellPadding`，不支持逐单元格设置。如需差异化单元格内边距，目前无法通过协议表达。

!!! note "borderOptions.location"
    - `"all"`（默认）：刷新所有边框（外框 + 内部分隔线）
    - `"outside"`：仅刷新外框，保留内部分隔线
    - `"inside"`：仅刷新内部分隔线，保留外框
    映射到 Office.js 的 `Word.BorderLocation.all` / `outside` / `inside`。

!!! note "columnWidths 长度策略"
    `columnWidths` 长度应不大于表格列数；过短时只覆盖前缀列（其它列保持原宽），过长时抛 `INVALID_PARAM`。"列数"以表格未合并状态计，等同于 `word:get:documentStructure` 中 `tables[].columnCount`。

!!! note "数据写入分离"
    本事件不接受表内容覆写——如需写入文本，请单独调用 [`word:update:tableRowColumn`](#wordupdatetablerowcolumn) 或 [`word:update:tableCell`](#wordupdatetablecell)。这样可以保证错误恢复粒度更细，且事件职责清晰内聚。

**请求示例**（首行作表头并设置统一列宽 + 内细外粗）:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/contract.docx",
  "tableId": "table-0",
  "styleOptions": {
    "styleType": "Grid Table 1 Light",
    "totalRow": false,
    "bandedRows": true,
    "cellPadding": { "top": 4, "bottom": 4, "left": 6, "right": 6 }
  },
  "borderOptions": {
    "location": "outside",
    "style": "Single",
    "width": 1.5,
    "color": "#000000"
  },
  "columnWidths": [120, 100, 120, 120],
  "alignment": "Centered"
}
```

**响应数据**:

```typescript
interface UpdateTableFormatResponse {
  requestId: string;
  success: boolean;
  data?: {
    tableId: string;
    rowCount: number;
    columnCount: number;
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
    "tableId": "table-0",
    "rowCount": 5,
    "columnCount": 4
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4001 | `MISSING_PARAM` - 至少需要传递一个可更新字段 |
| 4002 | `INVALID_PARAM` - `columnWidths` 长度大于表格列数，或 `borderOptions.width` 等值非法 |
| 3010 | `ELEMENT_NOT_FOUND` - 指定的 `tableId` 在文档中未找到 |
| 3011 | `STYLE_NOT_FOUND` - `styleOptions.styleType` 在文档中不存在 |
| 3013 | `NO_TABLE_AT_CURSOR` - 缺省 `tableId` 时光标未在任何表格内 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3000 | `DOCUMENT_ERROR` - 文档操作错误（通用） |

---

## 高级功能类

### word:insert:toc

**方向**: Server → AddIn（请求-响应）

**状态**: ✅ Stable

**说明**: 在文档中插入目录（Table of Contents）。

**请求数据**:

```typescript
interface InsertTOCRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  options?: {
    maxLevel?: number;     // 包含的最大标题级别（1-9），默认 3
    styles?: string[];     // 自定义样式名称列表
  };
}
```

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/report.docx",
  "options": {
    "maxLevel": 3
  }
}
```

**响应数据**:

```typescript
interface InsertTOCResponse {
  requestId: string;
  success: boolean;
  data?: {
    inserted: boolean;   // 是否成功插入
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

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4000 | `VALIDATION_ERROR` - 请求参数校验失败 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3000 | `DOCUMENT_ERROR` - 文档操作错误（通用） |

---

### word:export:content

**方向**: Server → AddIn（请求-响应）

**状态**: ✅ Stable

**说明**: 导出文档内容为指定格式。

**请求数据**:

```typescript
interface ExportContentRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  format: "text" | "html" | "markdown";   // 导出格式
  options?: {
    includeImages?: boolean;              // 是否包含图片，默认 true
    includeTables?: boolean;              // 是否包含表格，默认 true
  };
}
```

**导出格式说明**:

| 格式 | 说明 |
|------|------|
| `text` | 纯文本格式 |
| `html` | HTML 格式，保留基本格式 |
| `markdown` | Markdown 格式 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/report.docx",
  "format": "markdown",
  "options": {
    "includeImages": true,
    "includeTables": true
  }
}
```

**响应数据**:

```typescript
interface ExportContentResponse {
  requestId: string;
  success: boolean;
  data?: {
    content: string;     // 导出的内容
    format: string;      // 导出格式
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
    "content": "# 标题\n\n这是文档内容...",
    "format": "markdown"
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4000 | `VALIDATION_ERROR` - 请求参数校验失败 |
| 3007 | `FORMAT_NOT_SUPPORTED` - `format` 枚举合法但目标不支持产出该格式；改请求另一格式重发（见[判法划界](error-handling.md#format_not_supported-3007)） |
| 3016 | `API_NOT_SUPPORTED` - 该事件在当前宿主/平台整体不可用（换格式无用，须换路径或平台） |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3000 | `DOCUMENT_ERROR` - 文档操作错误（通用） |

---

## 批注操作类

### word:get:comments

**方向**: Server → AddIn（请求-响应）

**状态**: ✅ Stable

**说明**: 获取文档或当前选区中的批注列表。如果有选区，优先返回选区范围内的批注；否则返回整个文档的批注。

**请求数据**:

```typescript
interface GetCommentsRequest {
  requestId: string;      // 请求 ID (UUID)
  documentUri: string;    // 文档 URI
  timestamp?: number;     // 请求时间戳（毫秒），可选
  options?: {
    includeResolved?: boolean;       // 是否包含已解决的批注，默认 true
    includeReplies?: boolean;        // 是否包含批注回复，默认 true
    includeAssociatedText?: boolean; // 是否包含关联文本，默认 true
    detailedMetadata?: boolean;      // 是否包含详细元数据（作者、日期），默认 false
    maxTextLength?: number;          // 文本内容最大长度，超出截断
  };
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `requestId` | string | ✅ | 请求唯一标识（UUID） |
| `documentUri` | string | ✅ | 文档 URI |
| `options.includeResolved` | boolean | ❌ | 是否包含已解决的批注，默认 `true` |
| `options.includeReplies` | boolean | ❌ | 是否包含批注回复，默认 `true` |
| `options.includeAssociatedText` | boolean | ❌ | 是否包含关联文本，默认 `true` |
| `options.detailedMetadata` | boolean | ❌ | 是否包含作者和日期，默认 `false` |
| `options.maxTextLength` | number | ❌ | 文本内容最大长度，超出截断 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/report.docx",
  "options": {
    "includeResolved": false,
    "includeReplies": true,
    "detailedMetadata": true
  }
}
```

**响应数据**:

```typescript
interface GetCommentsResponse {
  requestId: string;
  success: boolean;
  data?: {
    comments: CommentData[];
  };
  error?: ErrorResponse;
  timestamp: number;
}

interface CommentData {
  id: string;                  // 批注唯一 ID
  content: string;             // 批注文本内容
  authorName?: string;         // 作者名称（detailedMetadata=true 时返回）
  authorEmail?: string;        // 作者邮箱（detailedMetadata=true 时返回）
  creationDate?: string;       // 创建时间 ISO 8601（detailedMetadata=true 时返回）
  resolved: boolean;           // 是否已解决
  associatedText?: string;     // 关联文本（被批注的文本）
  replies?: CommentReplyData[];// 回复列表
}

interface CommentReplyData {
  id: string;                  // 回复唯一 ID
  content: string;             // 回复文本内容
  authorName?: string;         // 回复作者名称
  authorEmail?: string;        // 回复作者邮箱
  creationDate?: string;       // 回复创建时间 ISO 8601
}
```

**响应示例（成功）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "comments": [
      {
        "id": "comment-1",
        "content": "请修改这段文字",
        "authorName": "张三",
        "authorEmail": "zhangsan@example.com",
        "creationDate": "2025-01-15T10:30:00.000Z",
        "resolved": false,
        "associatedText": "需要修改的文字",
        "replies": [
          {
            "id": "reply-1",
            "content": "已修改",
            "authorName": "李四",
            "creationDate": "2025-01-15T11:00:00.000Z"
          }
        ]
      }
    ]
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4000 | `VALIDATION_ERROR` - 请求参数校验失败 |
| 3000 | `DOCUMENT_ERROR` - 文档操作错误（通用） |

---

### word:insert:comment

**方向**: Server → AddIn（请求-响应）

**状态**: ✅ Stable

**说明**: 在指定目标范围上插入一条新批注。支持在当前选区或通过搜索文本定位的范围上插入。

**请求数据**:

```typescript
type InsertCommentTarget =
  | { type: "selection" }
  | {
      type: "searchText";
      searchText: string;          // 要搜索的文本（1-255 字符）
      searchOptions?: {
        matchCase?: boolean;       // 区分大小写，默认 false
        matchWholeWord?: boolean;  // 全字匹配，默认 false
      };
    };

interface InsertCommentRequest {
  requestId: string;      // 请求 ID (UUID)
  documentUri: string;    // 文档 URI
  timestamp?: number;     // 请求时间戳（毫秒），可选
  text: string;           // 批注文本内容（不能为空）
  target?: InsertCommentTarget;  // 目标范围（默认当前选区）
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `requestId` | string | ✅ | 请求唯一标识（UUID） |
| `documentUri` | string | ✅ | 文档 URI |
| `text` | string | ✅ | 批注文本内容，至少 1 个字符 |
| `target` | InsertCommentTarget | ❌ | 目标范围，不传或 `type: "selection"` 使用当前选区；`type: "searchText"` 通过搜索文本定位 |
| `target.searchText` | string | 条件 | 搜索文本（`type: "searchText"` 时必需，1-255 字符） |
| `target.searchOptions` | object | ❌ | 搜索选项 |
| `target.searchOptions.matchCase` | boolean | ❌ | 区分大小写，默认 `false` |
| `target.searchOptions.matchWholeWord` | boolean | ❌ | 全字匹配，默认 `false` |

**请求示例（当前选区）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/report.docx",
  "text": "这段文字需要修改"
}
```

**请求示例（搜索文本）**:

```json
{
  "requestId": "b2c3d4e5-f6a7-5b6c-9d8e-0f1a2b3c4d5e",
  "documentUri": "file:///Users/john/Documents/report.docx",
  "text": "建议将此处改为更准确的表述",
  "target": {
    "type": "searchText",
    "searchText": "需要审核的段落",
    "searchOptions": {
      "matchCase": false,
      "matchWholeWord": false
    }
  }
}
```

**响应数据**:

```typescript
interface InsertCommentResponse {
  requestId: string;
  success: boolean;
  data?: {
    commentId: string;    // 新创建的批注 ID
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
    "commentId": "comment-12345"
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4000 | `VALIDATION_ERROR` - 请求参数校验失败（text 为空、searchText 超过 255 字符等） |
| 3002 | `SELECTION_EMPTY` - 没有活动选区（`target` 为 `selection` 模式时） |
| 3012 | `SEARCH_NO_MATCH` - 搜索文本未找到匹配（`target` 为 `searchText` 模式时，见[判法划界](error-handling.md#search_no_match-3012)） |
| 3000 | `DOCUMENT_ERROR` - 文档操作错误（通用） |

---

### word:delete:comment

**方向**: Server → AddIn（请求-响应）

**状态**: ✅ Stable

**说明**: 按 ID 删除一条批注。

**请求数据**:

```typescript
interface DeleteCommentRequest {
  requestId: string;      // 请求 ID (UUID)
  documentUri: string;    // 文档 URI
  timestamp?: number;     // 请求时间戳（毫秒），可选
  commentId: string;      // 要删除的批注 ID（不能为空）
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `requestId` | string | ✅ | 请求唯一标识（UUID） |
| `documentUri` | string | ✅ | 文档 URI |
| `commentId` | string | ✅ | 要删除的批注 ID |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/report.docx",
  "commentId": "comment-12345"
}
```

**响应数据**:

```typescript
interface DeleteCommentResponse {
  requestId: string;
  success: boolean;
  data?: {
    deleted: boolean;     // 是否删除成功
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
    "deleted": true
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4000 | `VALIDATION_ERROR` - 请求参数校验失败（commentId 为空） |
| 3000 | `DOCUMENT_ERROR` - 文档操作错误（如批注不存在） |

---

### word:reply:comment

**方向**: Server → AddIn（请求-响应）

**状态**: ✅ Stable

**说明**: 回复已有批注。

**请求数据**:

```typescript
interface ReplyCommentRequest {
  requestId: string;      // 请求 ID (UUID)
  documentUri: string;    // 文档 URI
  timestamp?: number;     // 请求时间戳（毫秒），可选
  commentId: string;      // 要回复的批注 ID（不能为空）
  text: string;           // 回复文本内容（不能为空）
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `requestId` | string | ✅ | 请求唯一标识（UUID） |
| `documentUri` | string | ✅ | 文档 URI |
| `commentId` | string | ✅ | 要回复的批注 ID |
| `text` | string | ✅ | 回复文本内容，至少 1 个字符 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/report.docx",
  "commentId": "comment-12345",
  "text": "已按照建议修改"
}
```

**响应数据**:

```typescript
interface ReplyCommentResponse {
  requestId: string;
  success: boolean;
  data?: {
    replyId: string;      // 新创建的回复 ID
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
    "replyId": "reply-67890"
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4000 | `VALIDATION_ERROR` - 请求参数校验失败（commentId 或 text 为空） |
| 3000 | `DOCUMENT_ERROR` - 文档操作错误（如批注不存在） |

---

### word:resolve:comment

**方向**: Server → AddIn（请求-响应）

**状态**: ✅ Stable

**说明**: 解决或取消解决一条批注。

**请求数据**:

```typescript
interface ResolveCommentRequest {
  requestId: string;      // 请求 ID (UUID)
  documentUri: string;    // 文档 URI
  timestamp?: number;     // 请求时间戳（毫秒），可选
  commentId: string;      // 要操作的批注 ID（不能为空）
  resolved: boolean;      // true=解决，false=取消解决
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `requestId` | string | ✅ | 请求唯一标识（UUID） |
| `documentUri` | string | ✅ | 文档 URI |
| `commentId` | string | ✅ | 要操作的批注 ID |
| `resolved` | boolean | ✅ | `true` 表示解决，`false` 表示取消解决 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/report.docx",
  "commentId": "comment-12345",
  "resolved": true
}
```

**响应数据**:

```typescript
interface ResolveCommentResponse {
  requestId: string;
  success: boolean;
  data?: {
    resolved: boolean;    // 当前解决状态
  };
  error?: ErrorResponse;
  timestamp: number;
}
```

**响应示例（成功 - 解决）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "resolved": true
  },
  "timestamp": 1704067200500
}
```

**响应示例（成功 - 取消解决）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "resolved": false
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4000 | `VALIDATION_ERROR` - 请求参数校验失败（commentId 为空） |
| 3000 | `DOCUMENT_ERROR` - 文档操作错误（如批注不存在） |

---

## 脚本执行类

### word:run:script

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

!!! note "新增事件（纯加法）"
    `/word` 命名空间为 ✅ Stable，但本事件为**纯加法**——新增事件**不修改**任何既有事件 / 字段 / 错误码的语义与契约，不影响 `/word` 既有 Stable 事件的兼容性承诺。事件本身以 📋 Draft 进入（接口可能在正式发布前微调）。

**说明**: **封装层逃生舱**——对实时文档执行一段原始 Office.js 脚本。**应优先使用 typed `word:*` 事件**，仅在其未覆盖某能力、或某 typed 事件有 Bug 阻塞时使用。共享执行语义、大小限制、超时、安全模型、非原子性与错误映射见[通用约定 · 脚本执行](conventions.md#run-script)。

**请求数据**（共享 [`RunScriptRequest`](data-structures.md#script-execution)）:

```typescript
interface RunScriptRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  script: string;                    // async 函数体；注入 context(Word.RequestContext)/args/console，可 return
  args?: Record<string, unknown>;    // 注入脚本，脚本内经 `args` 读取；必须可 JSON 序列化
  timeoutMs?: number;                // 执行超时（毫秒），缺省「脚本执行」档 60000，无硬上限
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `script` | string | ✅ | 待执行 JS 源码，async 函数体语义 |
| `args` | Record<string, unknown> | ❌ | 注入脚本的参数，脚本内经 `args` 读取；必须可 JSON 序列化 |
| `timeoutMs` | number | ❌ | 执行超时（毫秒），缺省 60000，无硬上限 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/report.docx",
  "script": "const paras = context.document.body.paragraphs;\nparas.load('text');\nawait context.sync();\nreturn paras.items.map(p => p.text);"
}
```

**响应数据**（`data` 为共享 [`ScriptResult`](data-structures.md#script-execution)）:

```typescript
interface RunScriptResponse {
  requestId: string;
  success: boolean;
  data?: ScriptResult;        // { result, logs, durationMs, logsTruncated }
  error?: ErrorResponse;
  timestamp: number;
  duration?: number;
}
```

**响应示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "result": ["标题", "第一段正文。"],
    "logs": [],
    "durationMs": 51,
    "logsTruncated": false
  },
  "timestamp": 1704067200500,
  "duration": 54
}
```

!!! danger "安全模型与非原子性"
    `run:script` 以敞开全局对象的原生 JS 执行、**非沙箱**，信任边界即 Socket.IO 握手鉴权；失败时文档可能残留**部分修改**（无原子性 / 回滚）。执行前的人在环确认由**上层 Agent 层**承担，OASP 为纯能力提供方。完整规范见[通用约定 · 脚本执行](conventions.md#run-script)。

**可能的错误**（完整映射见[通用约定 · 脚本执行 · 错误映射](conventions.md#run-script-errors)）:

| 错误码 | 说明 |
|--------|------|
| 4001 | `MISSING_PARAM` - 缺少 `script` |
| 4003 | `INVALID_PARAM_TYPE` - `args` 不可 JSON 序列化 |
| 4004 | `PARAM_OUT_OF_RANGE` - `timeoutMs` 越界 |
| 4002 | `INVALID_PARAM` - 脚本语法错（`phase:"compile"`） |
| 3016 | `API_NOT_SUPPORTED` - 宿主不支持动态代码构造（`phase:"compile"`） |
| 3004 | `OPERATION_FAILED` - 脚本自身抛错（`fault:"script"`）或 Office.js 调用失败（`fault:"office"`） |
| 1002 | `TIMEOUT` - 执行超时 |
| 3006 | `CONTENT_TOO_LARGE` - result 过大（`phase:"serialize"`） |
