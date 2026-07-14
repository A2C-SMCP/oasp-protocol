# Excel 事件定义

!!! warning "Draft 状态"
    本文档中的所有事件处于 **Draft** 状态，接口可能在正式发布前发生变更。

## 概述

本章定义 `/excel` 命名空间下的所有事件。Excel 事件用于操作 Microsoft Excel 工作簿。

共 **37 个事件**，覆盖状态感知、Range 操作、格式样式、合并单元格、工作表管理、表格操作、图表操作、透视表操作、查找筛选和公式设置。

## 事件列表

### 状态感知类（Server → AddIn，请求-响应）

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [excel:get:workbookInfo](#excelgetworkbookinfo) | 📋 Draft | 获取工作簿信息 |
| [excel:get:worksheetInfo](#excelgetworksheetinfo) | 📋 Draft | 获取工作表详细信息 |
| [excel:get:selectedRange](#excelgetselectedrange) | 📋 Draft | 获取选中范围 |

### Range 操作类（Server → AddIn，请求-响应）

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [excel:get:range](#excelgetrange) | 📋 Draft | 获取范围内的值 |
| [excel:set:range](#excelsetrange) | 📋 Draft | 设置范围内的值 |
| [excel:clear:range](#excelclearrange) | 📋 Draft | 清除范围内容/格式 |
| [excel:copy:range](#excelcopyrange) | 📋 Draft | 复制范围到目标位置 |
| [excel:delete:range](#exceldeleterange) | 📋 Draft | 删除范围并移动单元格 |
| [excel:insert:range](#excelinsertrange) | 📋 Draft | 插入范围并移动单元格 |

### 格式与样式类（Server → AddIn，请求-响应）

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [excel:get:rangeFormat](#excelgetrangeformat) | 📋 Draft | 获取范围格式信息 |
| [excel:set:rangeFormat](#excelsetrangeformat) | 📋 Draft | 设置范围格式 |
| [excel:add:conditionalFormat](#exceladdconditionalformat) | 📋 Draft | 添加条件格式 |
| [excel:clear:conditionalFormat](#excelclearconditionalformat) | 📋 Draft | 清除条件格式 |

### 合并单元格类（Server → AddIn，请求-响应）

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [excel:merge:cells](#excelmergecells) | 📋 Draft | 合并单元格 |
| [excel:unmerge:cells](#excelunmergecells) | 📋 Draft | 取消合并单元格 |

### 工作表管理类（Server → AddIn，请求-响应）

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [excel:get:worksheets](#excelgetworksheets) | 📋 Draft | 获取所有工作表列表 |
| [excel:add:worksheet](#exceladdworksheet) | 📋 Draft | 添加工作表 |
| [excel:delete:worksheet](#exceldeleteworksheet) | 📋 Draft | 删除工作表 |
| [excel:rename:worksheet](#excelrenameworksheet) | 📋 Draft | 重命名工作表 |
| [excel:activate:worksheet](#excelactivateworksheet) | 📋 Draft | 激活工作表 |

### Table 操作类（Server → AddIn，请求-响应）

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [excel:insert:table](#excelinserttable) | 📋 Draft | 创建结构化表格 |
| [excel:get:table](#excelgettable) | 📋 Draft | 获取表格详细信息 |
| [excel:get:tables](#excelgettables) | 📋 Draft | 获取工作表中所有表格 |
| [excel:add:tableRow](#exceladdtablerow) | 📋 Draft | 追加表格行 |
| [excel:delete:tableRow](#exceldeletetablerow) | 📋 Draft | 删除表格行 |
| [excel:sort:table](#excelsorttable) | 📋 Draft | 对表格排序 |

### Chart 操作类（Server → AddIn，请求-响应）

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [excel:insert:chart](#excelinsertchart) | 📋 Draft | 创建图表 |
| [excel:get:charts](#excelgetcharts) | 📋 Draft | 获取所有图表 |
| [excel:update:chart](#excelupdatechart) | 📋 Draft | 更新图表属性 |
| [excel:delete:chart](#exceldeletechart) | 📋 Draft | 删除图表 |

### PivotTable 操作类（Server → AddIn，请求-响应）

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [excel:insert:pivotTable](#excelinsertpivottable) | 📋 Draft | 创建透视表 |
| [excel:get:pivotTables](#excelgetpivottables) | 📋 Draft | 获取所有透视表 |
| [excel:delete:pivotTable](#exceldeletepivottable) | 📋 Draft | 删除透视表 |

### 查找与筛选类（Server → AddIn，请求-响应）

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [excel:find:values](#excelfindvalues) | 📋 Draft | 查找值 |
| [excel:set:autoFilter](#excelsetautofilter) | 📋 Draft | 设置自动筛选 |
| [excel:clear:autoFilter](#excelclearautofilter) | 📋 Draft | 清除自动筛选 |

### 公式类（Server → AddIn，请求-响应）

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [excel:set:formula](#excelsetformula) | 📋 Draft | 设置单元格公式 |

---

## Excel 错误码映射（规范层）

`/excel` **不使用命名空间专属错误码块**，与 `/word`、`/ppt` 一致，全部复用[错误处理](error-handling.md)定义的共享注册表（1xxx ~ 4xxx）。错误码描述的是**线缆可观测的失败条件**，与具体应用无关（见[错误处理 · 设计原则](error-handling.md#design-principles)）。

!!! warning "规范层（Normative）"
    本节为**规范层**（见[通用约定 · 规范分层](conventions.md#normative-layering)）。当出现下表所列**线缆可观测条件**时，实现 **MUST** 返回对应错误码，**不得**降级为通用 `3000 DOCUMENT_ERROR`。多个同类对象（工作表 / 表格 / 图表 / 透视表）共用 `3010`，由 [`details.kind`](error-handling.md#element_not_found-3010) 区分（取值集合见该处规范定义）。

| 触发条件（线缆可观测） | 错误码 | 名称 | `details` | 典型事件 |
|---|---|---|---|---|
| 按名称/id 定位工作表失败 | `3010` | `ELEMENT_NOT_FOUND` | `kind:"worksheet"` | 所有工作表作用域事件（`get:worksheetInfo`、`delete:worksheet`、`get:range` …） |
| 按名称/id 定位表格失败 | `3010` | `ELEMENT_NOT_FOUND` | `kind:"table"` | `get:table`、`add:tableRow`、`delete:tableRow`、`sort:table` |
| 按名称定位图表失败 | `3010` | `ELEMENT_NOT_FOUND` | `kind:"chart"` | `update:chart`、`delete:chart` |
| 按名称定位透视表失败 | `3010` | `ELEMENT_NOT_FOUND` | `kind:"pivotTable"` | `delete:pivotTable` |
| 区域/单元格地址非法（语法或越界） | `3009` | `RANGE_INVALID` | — | `get:range`、`set:range`、`insert:chart`(source)、`insert:pivotTable`(source/target)、`find:values`、`set:formula` … |
| 合并区域与现有合并冲突 | `3014` | `ALREADY_MERGED` | — | `merge:cells` |
| 目标工作表/区域受保护不可写 | `3003` | `DOCUMENT_READ_ONLY` | `scope:"worksheet"` | `set:range`、`clear:range`、`delete:range`、`insert:range`、`set:rangeFormat`、`merge:cells`、`insert:table`、`delete:worksheet`、`set:formula` … |
| 公式语法错误或引用无效 | `3017` | `FORMULA_ERROR` | — | `set:formula` |
| 写入值与目标类型不兼容（apply-time） | `3018` | `DATA_TYPE_MISMATCH` | — | `set:range`、`add:tableRow` |
| 平台不提供该功能（如透视表） | `3016` | `API_NOT_SUPPORTED` | `requiredApiSet` / 平台 | `insert:pivotTable` |
| 文档未打开/找不到 | `3001` | `DOCUMENT_NOT_FOUND` | — | `get:workbookInfo`、`get:worksheets` … |
| 参数缺失/非法/越界 | `4001`/`4002`/`4004` | `MISSING_PARAM` / `INVALID_PARAM` / `PARAM_OUT_OF_RANGE` | — | `add:conditionalFormat`、`insert:chart`(type)、`sort:table`、`delete:tableRow`、`set:autoFilter` |

!!! note "从早期 5xxx 草案迁移"
    早期草案曾为 `/excel` 定义 `5001`~`5010` 专属码，现已全部收敛至上述通用码（对齐 `/word`、`/ppt`）：`5001/5006/5007/5008 → 3010`（`details.kind`）、`5002 → 3009`、`5003 → 3014`、`5004 → 3003`、`5005 → 3017`、`5009 → 3018`、`5010 → 3016`。`5xxx` 不再使用。

---

## 通用说明

### worksheetName 参数

大部分事件支持可选的 `worksheetName` 参数。省略时默认操作当前活动工作表。

### address 格式

Excel 地址支持以下格式：

- 单元格：`"A1"`、`"B2"`
- 范围：`"A1:C3"`、`"A:A"`（整列）、`"1:1"`（整行）
- 带工作表前缀：`"Sheet1!A1:C3"`

---

## 状态感知类

### excel:get:workbookInfo

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 获取当前工作簿的基本信息，包括所有工作表列表、当前活动工作表和文件名。

**请求数据**:

```typescript
interface GetWorkbookInfoRequest {
  requestId: string;      // 请求 ID (UUID)
  documentUri: string;    // 文档 URI
  timestamp?: number;     // 请求时间戳（毫秒），可选
}
```

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "timestamp": 1704067200000
}
```

**响应数据**:

```typescript
interface GetWorkbookInfoResponse {
  requestId: string;
  success: boolean;
  data?: {
    sheets: SheetInfo[];     // 所有工作表信息
    activeSheet: string;     // 当前活动工作表名称
    fileName: string;        // 文件名
  };
  error?: ErrorResponse;
  timestamp: number;
  duration?: number;
}

interface SheetInfo {
  name: string;              // 工作表名称
  index: number;             // 工作表索引（从 0 开始）
  isActive: boolean;         // 是否为活动工作表
  isHidden: boolean;         // 是否隐藏
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `sheets` | SheetInfo[] | ✅ | 工作簿中所有工作表的信息数组 |
| `sheets[].name` | string | ✅ | 工作表名称 |
| `sheets[].index` | number | ✅ | 工作表索引（从 0 开始） |
| `sheets[].isActive` | boolean | ✅ | 是否为当前活动工作表 |
| `sheets[].isHidden` | boolean | ✅ | 是否隐藏 |
| `activeSheet` | string | ✅ | 当前活动工作表名称 |
| `fileName` | string | ✅ | 文件名 |

**响应示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "sheets": [
      { "name": "Sheet1", "index": 0, "isActive": true, "isHidden": false },
      { "name": "Sheet2", "index": 1, "isActive": false, "isHidden": false },
      { "name": "Hidden", "index": 2, "isActive": false, "isHidden": true }
    ],
    "activeSheet": "Sheet1",
    "fileName": "data.xlsx"
  },
  "timestamp": 1704067200500,
  "duration": 30
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |

---

### excel:get:worksheetInfo

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 获取指定工作表的详细信息，包括已使用范围、表格数量和图表数量。

**请求数据**:

```typescript
interface GetWorksheetInfoRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  worksheetName?: string;    // 工作表名称，省略时使用活动工作表
}
```

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "worksheetName": "Sheet1"
}
```

**响应数据**:

```typescript
interface GetWorksheetInfoResponse {
  requestId: string;
  success: boolean;
  data?: {
    name: string;                         // 工作表名称
    usedRange: {
      address: string;                    // 已使用范围地址
      rowCount: number;                   // 行数
      columnCount: number;                // 列数
    };
    tableCount: number;                   // 表格数量
    chartCount: number;                   // 图表数量
  };
  error?: ErrorResponse;
  timestamp: number;
  duration?: number;
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `name` | string | ✅ | 工作表名称 |
| `usedRange.address` | string | ✅ | 已使用范围的地址（如 `"A1:D10"`） |
| `usedRange.rowCount` | number | ✅ | 已使用范围的行数 |
| `usedRange.columnCount` | number | ✅ | 已使用范围的列数 |
| `tableCount` | number | ✅ | 工作表中的结构化表格数量 |
| `chartCount` | number | ✅ | 工作表中的图表数量 |

**响应示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "name": "Sheet1",
    "usedRange": {
      "address": "Sheet1!A1:D10",
      "rowCount": 10,
      "columnCount": 4
    },
    "tableCount": 1,
    "chartCount": 2
  },
  "timestamp": 1704067200500,
  "duration": 45
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3010 | `ELEMENT_NOT_FOUND` - 工作表不存在（`details.kind:"worksheet"`） |

---

### excel:get:selectedRange

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 获取当前选中的单元格范围信息及其值。

**请求数据**:

```typescript
interface GetSelectedRangeRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
}
```

**响应数据**:

```typescript
interface GetSelectedRangeResponse {
  requestId: string;
  success: boolean;
  data?: {
    address: string;          // 选中范围地址
    values: unknown[][];      // 二维值数组
    rowCount: number;         // 行数
    columnCount: number;      // 列数
  };
  error?: ErrorResponse;
  timestamp: number;
  duration?: number;
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `address` | string | ✅ | 选中范围地址（如 `"Sheet1!A1:C3"`） |
| `values` | unknown[][] | ✅ | 二维数组，行优先 |
| `rowCount` | number | ✅ | 选中范围的行数 |
| `columnCount` | number | ✅ | 选中范围的列数 |

**响应示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "address": "Sheet1!A1:C3",
    "values": [
      ["姓名", "年龄", "城市"],
      ["张三", 25, "北京"],
      ["李四", 30, "上海"]
    ],
    "rowCount": 3,
    "columnCount": 3
  },
  "timestamp": 1704067200500,
  "duration": 50
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |

---

## Range 操作类

### excel:get:range

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 获取指定范围内的所有值，可选包含格式信息。

**请求数据**:

```typescript
interface GetRangeRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  address: string;            // 范围地址，如 "A1:C3"
  worksheetName?: string;     // 工作表名称，省略时使用活动工作表
  includeFormat?: boolean;    // 是否包含格式信息，默认 false
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `address` | string | ✅ | 范围地址 |
| `worksheetName` | string | ❌ | 工作表名称，省略时使用活动工作表 |
| `includeFormat` | boolean | ❌ | 是否返回格式信息，默认 `false` |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "address": "A1:C3",
  "includeFormat": true
}
```

**响应数据**:

```typescript
interface GetRangeResponse {
  requestId: string;
  success: boolean;
  data?: {
    address: string;
    values: unknown[][];
    rowCount: number;
    columnCount: number;
    format?: RangeFormatInfo;   // 仅当 includeFormat=true 时返回
  };
  error?: ErrorResponse;
  timestamp: number;
  duration?: number;
}

interface RangeFormatInfo {
  font: {
    name: string;               // 字体名称
    size: number;               // 字号（磅）
    bold: boolean;              // 粗体
    italic: boolean;            // 斜体
    color: string;              // 字体颜色 (#RRGGBB)
    underline: string;          // 下划线样式
  };
  fill: {
    color: string;              // 填充颜色 (#RRGGBB)
  };
  horizontalAlignment: string;  // 水平对齐
  verticalAlignment: string;    // 垂直对齐
  wrapText: boolean;            // 自动换行
  numberFormat: string[][];     // 数字格式（二维数组，每个单元格独立）
}
```

**响应示例（含格式）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "address": "Sheet1!A1:C3",
    "values": [
      ["姓名", "年龄", "城市"],
      ["张三", 25, "北京"],
      ["李四", 30, "上海"]
    ],
    "rowCount": 3,
    "columnCount": 3,
    "format": {
      "font": {
        "name": "等线",
        "size": 11,
        "bold": false,
        "italic": false,
        "color": "#000000",
        "underline": "None"
      },
      "fill": { "color": "#FFFFFF" },
      "horizontalAlignment": "General",
      "verticalAlignment": "Bottom",
      "wrapText": false,
      "numberFormat": [["General", "General", "General"], ["General", "0", "General"], ["General", "0", "General"]]
    }
  },
  "timestamp": 1704067200500,
  "duration": 80
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 工作表不存在（`details.kind:"worksheet"`） |
| 3009 | `RANGE_INVALID` - 无效的区域地址 |

---

### excel:set:range

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 设置指定范围内的值。支持标量值（填充整个范围）或二维数组。

**请求数据**:

```typescript
interface SetRangeRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  address: string;                    // 范围地址
  values: unknown | unknown[][];      // 标量或二维数组
  worksheetName?: string;
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `address` | string | ✅ | 目标范围地址 |
| `values` | unknown \| unknown[][] | ✅ | 标量值（填充整个范围）或二维数组 |
| `worksheetName` | string | ❌ | 工作表名称 |

**请求示例（二维数组）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "address": "A1:C3",
  "values": [
    ["姓名", "年龄", "城市"],
    ["张三", 25, "北京"],
    ["李四", 30, "上海"]
  ]
}
```

**请求示例（标量填充）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "address": "A1:C3",
  "values": 0
}
```

**响应数据**:

```typescript
interface SetRangeResponse {
  requestId: string;
  success: boolean;
  data?: {
    address: string;            // 被操作的范围地址
  };
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
    "address": "Sheet1!A1:C3"
  },
  "timestamp": 1704067200500,
  "duration": 60
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 工作表不存在（`details.kind:"worksheet"`） |
| 3009 | `RANGE_INVALID` - 无效的区域地址 |
| 3003 | `DOCUMENT_READ_ONLY` - 工作表受保护（`details.scope:"worksheet"`） |
| 3018 | `DATA_TYPE_MISMATCH` - 数据类型不匹配 |

---

### excel:clear:range

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 清除指定范围的内容、格式或全部。

**请求数据**:

```typescript
interface ClearRangeRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  address: string;
  clearType: "contents" | "formats" | "all";
  worksheetName?: string;
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `address` | string | ✅ | 范围地址 |
| `clearType` | string | ✅ | 清除类型：`"contents"` 仅内容、`"formats"` 仅格式、`"all"` 全部 |
| `worksheetName` | string | ❌ | 工作表名称 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "address": "A1:C3",
  "clearType": "contents"
}
```

**响应数据**:

```typescript
interface ClearRangeResponse {
  requestId: string;
  success: boolean;
  data?: { address: string };
  error?: ErrorResponse;
  timestamp: number;
  duration?: number;
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 工作表不存在（`details.kind:"worksheet"`） |
| 3009 | `RANGE_INVALID` - 无效的区域地址 |
| 3003 | `DOCUMENT_READ_ONLY` - 工作表受保护（`details.scope:"worksheet"`） |

---

### excel:copy:range

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 将源范围的内容复制到目标位置。

**请求数据**:

```typescript
interface CopyRangeRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  sourceAddress: string;      // 源范围地址
  targetAddress: string;      // 目标范围地址
  worksheetName?: string;
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `sourceAddress` | string | ✅ | 源范围地址 |
| `targetAddress` | string | ✅ | 目标范围地址 |
| `worksheetName` | string | ❌ | 工作表名称 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "sourceAddress": "A1:C3",
  "targetAddress": "E1:G3"
}
```

**响应数据**:

```typescript
interface CopyRangeResponse {
  requestId: string;
  success: boolean;
  data?: { address: string };   // 目标范围地址
  error?: ErrorResponse;
  timestamp: number;
  duration?: number;
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 工作表不存在（`details.kind:"worksheet"`） |
| 3009 | `RANGE_INVALID` - 无效的区域地址 |
| 3003 | `DOCUMENT_READ_ONLY` - 工作表受保护（`details.scope:"worksheet"`） |

---

### excel:delete:range

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 删除指定范围并将周围单元格向上或向左移动填充空位。

**请求数据**:

```typescript
interface DeleteRangeRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  address: string;
  shiftDirection: "up" | "left";    // 移动方向
  worksheetName?: string;
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `address` | string | ✅ | 要删除的范围地址 |
| `shiftDirection` | string | ✅ | 删除后周围单元格的移动方向：`"up"` 向上、`"left"` 向左 |
| `worksheetName` | string | ❌ | 工作表名称 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "address": "B2:B5",
  "shiftDirection": "up"
}
```

**响应数据**:

```typescript
interface DeleteRangeResponse {
  requestId: string;
  success: boolean;
  data?: { address: string };
  error?: ErrorResponse;
  timestamp: number;
  duration?: number;
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 工作表不存在（`details.kind:"worksheet"`） |
| 3009 | `RANGE_INVALID` - 无效的区域地址 |
| 3003 | `DOCUMENT_READ_ONLY` - 工作表受保护（`details.scope:"worksheet"`） |

---

### excel:insert:range

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 在指定位置插入空白单元格，将现有单元格向下或向右移动。

**请求数据**:

```typescript
interface InsertRangeRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  address: string;
  shiftDirection: "down" | "right";   // 移动方向
  worksheetName?: string;
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `address` | string | ✅ | 要插入的范围地址 |
| `shiftDirection` | string | ✅ | 现有单元格的移动方向：`"down"` 向下、`"right"` 向右 |
| `worksheetName` | string | ❌ | 工作表名称 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "address": "B2:B5",
  "shiftDirection": "down"
}
```

**响应数据**:

```typescript
interface InsertRangeResponse {
  requestId: string;
  success: boolean;
  data?: { address: string };
  error?: ErrorResponse;
  timestamp: number;
  duration?: number;
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 工作表不存在（`details.kind:"worksheet"`） |
| 3009 | `RANGE_INVALID` - 无效的区域地址 |
| 3003 | `DOCUMENT_READ_ONLY` - 工作表受保护（`details.scope:"worksheet"`） |

---

## 格式与样式类

### excel:get:rangeFormat

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 获取指定范围的完整格式信息，包括字体、填充、边框、对齐和数字格式。

**请求数据**:

```typescript
interface GetRangeFormatRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  address: string;
  worksheetName?: string;
}
```

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "address": "A1:C3"
}
```

**响应数据**:

```typescript
interface GetRangeFormatResponse {
  requestId: string;
  success: boolean;
  data?: {
    address: string;
    format: RangeFormatInfo;     // 见 excel:get:range 中的 RangeFormatInfo 定义
  };
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
    "address": "Sheet1!A1:C3",
    "format": {
      "font": {
        "name": "等线",
        "size": 11,
        "bold": true,
        "italic": false,
        "color": "#000000",
        "underline": "None"
      },
      "fill": { "color": "#FFFF00" },
      "horizontalAlignment": "Center",
      "verticalAlignment": "Center",
      "wrapText": true,
      "numberFormat": [["General", "0.00", "#,##0"]]
    }
  },
  "timestamp": 1704067200500,
  "duration": 40
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 工作表不存在（`details.kind:"worksheet"`） |
| 3009 | `RANGE_INVALID` - 无效的区域地址 |

---

### excel:set:rangeFormat

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 设置指定范围的格式。所有格式属性均为可选，仅传入的属性会被修改。

**请求数据**:

```typescript
interface SetRangeFormatRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  address: string;
  format: {
    font?: {
      name?: string;           // 字体名称
      size?: number;           // 字号（磅）
      bold?: boolean;          // 粗体
      italic?: boolean;        // 斜体
      color?: string;          // 颜色 (#RRGGBB)
      underline?: string;      // 下划线：None, Single, Double 等
    };
    fill?: {
      color?: string;          // 填充颜色 (#RRGGBB)
    };
    borders?: {
      top?: BorderFormat;
      bottom?: BorderFormat;
      left?: BorderFormat;
      right?: BorderFormat;
    };
    alignment?: {
      horizontal?: string;     // Left, Center, Right, Fill, Justify, General 等
      vertical?: string;       // Top, Center, Bottom, Justify 等
      wrapText?: boolean;      // 自动换行
      indentLevel?: number;    // 缩进级别（0 起）
      textOrientation?: number; // 文本方向角度
    };
    numberFormat?: string;     // 数字格式字符串，如 "0.00", "#,##0", "yyyy-mm-dd"
  };
  worksheetName?: string;
}

interface BorderFormat {
  style?: string;              // None, Thin, Medium, Thick, Dashed, Dotted 等
  color?: string;              // 颜色 (#RRGGBB)
  weight?: string;             // Hairline, Thin, Medium, Thick
}
```

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "address": "A1:C1",
  "format": {
    "font": { "bold": true, "size": 14, "color": "#FFFFFF" },
    "fill": { "color": "#4472C4" },
    "alignment": { "horizontal": "Center" }
  }
}
```

**响应数据**:

```typescript
interface SetRangeFormatResponse {
  requestId: string;
  success: boolean;
  data?: { address: string };
  error?: ErrorResponse;
  timestamp: number;
  duration?: number;
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 工作表不存在（`details.kind:"worksheet"`） |
| 3009 | `RANGE_INVALID` - 无效的区域地址 |
| 3003 | `DOCUMENT_READ_ONLY` - 工作表受保护（`details.scope:"worksheet"`） |

---

### excel:add:conditionalFormat

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 为指定范围添加条件格式规则。规则对象采用透传模式，工具只负责应用。

**请求数据**:

```typescript
interface AddConditionalFormatRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  address: string;
  rule: {
    type: string;              // 规则类型，如 "cellValue", "colorScale", "dataBar", "iconSet"
    [key: string]: unknown;    // 规则参数（透传）
  };
  worksheetName?: string;
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `address` | string | ✅ | 范围地址 |
| `rule.type` | string | ✅ | 条件格式规则类型 |
| `rule.*` | unknown | ❌ | 规则具体参数，按 `type` 不同而不同 |
| `worksheetName` | string | ❌ | 工作表名称 |

**请求示例（单元格值条件）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "address": "B2:B100",
  "rule": {
    "type": "cellValue",
    "operator": "greaterThan",
    "value": 90,
    "format": { "font": { "color": "#006100" }, "fill": { "color": "#C6EFCE" } }
  }
}
```

**响应数据**:

```typescript
interface AddConditionalFormatResponse {
  requestId: string;
  success: boolean;
  data?: { address: string };
  error?: ErrorResponse;
  timestamp: number;
  duration?: number;
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 工作表不存在（`details.kind:"worksheet"`） |
| 3009 | `RANGE_INVALID` - 无效的区域地址 |
| 4002 | `INVALID_PARAM` - 规则参数无效 |

---

### excel:clear:conditionalFormat

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 清除指定范围上的所有条件格式。

**请求数据**:

```typescript
interface ClearConditionalFormatRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  address: string;
  worksheetName?: string;
}
```

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "address": "B2:B100"
}
```

**响应数据**:

```typescript
interface ClearConditionalFormatResponse {
  requestId: string;
  success: boolean;
  data?: { address: string };
  error?: ErrorResponse;
  timestamp: number;
  duration?: number;
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 工作表不存在（`details.kind:"worksheet"`） |
| 3009 | `RANGE_INVALID` - 无效的区域地址 |

---

## 合并单元格类

### excel:merge:cells

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 将指定范围的单元格合并为一个单元格。合并后保留左上角单元格的值。

**请求数据**:

```typescript
interface MergeCellsRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  address: string;
  worksheetName?: string;
}
```

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "address": "A1:C1"
}
```

**响应数据**:

```typescript
interface MergeCellsResponse {
  requestId: string;
  success: boolean;
  data?: { address: string };
  error?: ErrorResponse;
  timestamp: number;
  duration?: number;
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 工作表不存在（`details.kind:"worksheet"`） |
| 3009 | `RANGE_INVALID` - 无效的区域地址 |
| 3014 | `ALREADY_MERGED` - 与现有合并区域冲突 |
| 3003 | `DOCUMENT_READ_ONLY` - 工作表受保护（`details.scope:"worksheet"`） |

---

### excel:unmerge:cells

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 取消指定范围的单元格合并。

**请求数据**:

```typescript
interface UnmergeCellsRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  address: string;
  worksheetName?: string;
}
```

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "address": "A1:C1"
}
```

**响应数据**:

```typescript
interface UnmergeCellsResponse {
  requestId: string;
  success: boolean;
  data?: { address: string };
  error?: ErrorResponse;
  timestamp: number;
  duration?: number;
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 工作表不存在（`details.kind:"worksheet"`） |
| 3009 | `RANGE_INVALID` - 无效的区域地址 |
| 3003 | `DOCUMENT_READ_ONLY` - 工作表受保护（`details.scope:"worksheet"`） |

---

## 工作表管理类

### excel:get:worksheets

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 获取工作簿中所有工作表的列表。

**请求数据**:

```typescript
interface GetWorksheetsRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
}
```

**响应数据**:

```typescript
interface GetWorksheetsResponse {
  requestId: string;
  success: boolean;
  data?: {
    worksheets: SheetInfo[];    // 见 excel:get:workbookInfo 中的 SheetInfo 定义
  };
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
    "worksheets": [
      { "name": "Sheet1", "index": 0, "isActive": true, "isHidden": false },
      { "name": "Sheet2", "index": 1, "isActive": false, "isHidden": false }
    ]
  },
  "timestamp": 1704067200500,
  "duration": 20
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |

---

### excel:add:worksheet

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 添加新工作表。

**请求数据**:

```typescript
interface AddWorksheetRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  name?: string;               // 工作表名称，省略时自动生成
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `name` | string | ❌ | 新工作表名称，省略时由 Excel 自动命名 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "name": "数据分析"
}
```

**响应数据**:

```typescript
interface AddWorksheetResponse {
  requestId: string;
  success: boolean;
  data?: {
    name: string;              // 新工作表的实际名称
    index: number;             // 新工作表的索引
  };
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
    "name": "数据分析",
    "index": 2
  },
  "timestamp": 1704067200500,
  "duration": 50
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3004 | `OPERATION_FAILED` - 操作失败（如名称已存在） |

---

### excel:delete:worksheet

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 删除指定工作表。

**请求数据**:

```typescript
interface DeleteWorksheetRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  worksheetName: string;       // 要删除的工作表名称
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `worksheetName` | string | ✅ | 要删除的工作表名称 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "worksheetName": "Sheet3"
}
```

**响应数据**:

```typescript
interface DeleteWorksheetResponse {
  requestId: string;
  success: boolean;
  data?: { deleted: boolean };
  error?: ErrorResponse;
  timestamp: number;
  duration?: number;
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 工作表不存在（`details.kind:"worksheet"`） |
| 3003 | `DOCUMENT_READ_ONLY` - 工作表受保护（`details.scope:"worksheet"`） |

---

### excel:rename:worksheet

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 重命名工作表。

**请求数据**:

```typescript
interface RenameWorksheetRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  currentName: string;         // 当前名称
  newName: string;             // 新名称
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `currentName` | string | ✅ | 当前工作表名称 |
| `newName` | string | ✅ | 新工作表名称 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "currentName": "Sheet1",
  "newName": "销售数据"
}
```

**响应数据**:

```typescript
interface RenameWorksheetResponse {
  requestId: string;
  success: boolean;
  data?: { name: string };     // 重命名后的名称
  error?: ErrorResponse;
  timestamp: number;
  duration?: number;
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 工作表不存在（`details.kind:"worksheet"`） |
| 3004 | `OPERATION_FAILED` - 操作失败（如新名称已存在） |

---

### excel:activate:worksheet

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 激活（切换到）指定工作表。

**请求数据**:

```typescript
interface ActivateWorksheetRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  worksheetName: string;       // 要激活的工作表名称
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `worksheetName` | string | ✅ | 要激活的工作表名称 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "worksheetName": "Sheet2"
}
```

**响应数据**:

```typescript
interface ActivateWorksheetResponse {
  requestId: string;
  success: boolean;
  data?: { activated: boolean };
  error?: ErrorResponse;
  timestamp: number;
  duration?: number;
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 工作表不存在（`details.kind:"worksheet"`） |

---

## Table 操作类

### excel:insert:table

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 在指定范围创建结构化表格（Excel Table）。可选提供初始数据和样式。

**请求数据**:

```typescript
interface InsertTableRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  address: string;             // 表格范围地址
  hasHeaders: boolean;         // 第一行是否为标题
  data?: unknown[][];          // 初始数据（可选）
  styleName?: string;          // 表格样式名称（可选）
  worksheetName?: string;
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `address` | string | ✅ | 表格范围地址 |
| `hasHeaders` | boolean | ✅ | 第一行是否为标题行 |
| `data` | unknown[][] | ❌ | 初始数据（二维数组） |
| `styleName` | string | ❌ | 表格样式名称（如 `"TableStyleMedium2"`） |
| `worksheetName` | string | ❌ | 工作表名称 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "address": "A1:C4",
  "hasHeaders": true,
  "data": [
    ["姓名", "年龄", "城市"],
    ["张三", 25, "北京"],
    ["李四", 30, "上海"],
    ["王五", 28, "广州"]
  ],
  "styleName": "TableStyleMedium2"
}
```

**响应数据**:

```typescript
interface InsertTableResponse {
  requestId: string;
  success: boolean;
  data?: {
    name: string;              // 表格名称（自动生成，如 "Table1"）
    address: string;           // 表格实际范围
  };
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
    "name": "Table1",
    "address": "Sheet1!A1:C4"
  },
  "timestamp": 1704067200500,
  "duration": 100
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 工作表不存在（`details.kind:"worksheet"`） |
| 3009 | `RANGE_INVALID` - 无效的区域地址 |
| 3003 | `DOCUMENT_READ_ONLY` - 工作表受保护（`details.scope:"worksheet"`） |

---

### excel:get:table

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 获取指定表格的详细信息。

**请求数据**:

```typescript
interface GetTableRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  tableId: string;             // 表格名称或 ID
  worksheetName?: string;
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `tableId` | string | ✅ | 表格名称或 ID |
| `worksheetName` | string | ❌ | 工作表名称 |

**响应数据**:

```typescript
interface GetTableResponse {
  requestId: string;
  success: boolean;
  data?: {
    name: string;
    id: string;
    address: string;
    rowCount: number;          // 数据行数（不含标题）
    columnCount: number;
    columns: Array<{
      name: string;            // 列名
      index: number;           // 列索引（从 0 开始）
    }>;
    styleName: string;         // 表格样式名称
    showHeaders: boolean;      // 是否显示标题行
  };
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
    "name": "Table1",
    "id": "{12345678-1234-1234-1234-123456789012}",
    "address": "Sheet1!A1:C4",
    "rowCount": 3,
    "columnCount": 3,
    "columns": [
      { "name": "姓名", "index": 0 },
      { "name": "年龄", "index": 1 },
      { "name": "城市", "index": 2 }
    ],
    "styleName": "TableStyleMedium2",
    "showHeaders": true
  },
  "timestamp": 1704067200500,
  "duration": 40
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 工作表不存在（`details.kind:"worksheet"`） |
| 3010 | `ELEMENT_NOT_FOUND` - 表格不存在（`details.kind:"table"`） |

---

### excel:get:tables

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 获取工作表中所有表格的概要列表。

**请求数据**:

```typescript
interface GetTablesRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  worksheetName?: string;
}
```

**响应数据**:

```typescript
interface GetTablesResponse {
  requestId: string;
  success: boolean;
  data?: {
    tables: Array<{
      name: string;            // 表格名称
      id: string;              // 表格 ID
      address: string;         // 表格范围
    }>;
  };
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
    "tables": [
      { "name": "Table1", "id": "{12345678-...}", "address": "Sheet1!A1:C4" },
      { "name": "Table2", "id": "{87654321-...}", "address": "Sheet1!E1:G10" }
    ]
  },
  "timestamp": 1704067200500,
  "duration": 30
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 工作表不存在（`details.kind:"worksheet"`） |

---

### excel:add:tableRow

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 向指定表格末尾追加一行数据。

**请求数据**:

```typescript
interface AddTableRowRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  tableId: string;             // 表格名称或 ID
  values: unknown[];           // 行数据（一维数组，按列顺序）
  worksheetName?: string;
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `tableId` | string | ✅ | 表格名称或 ID |
| `values` | unknown[] | ✅ | 一维数组，元素顺序与列对应 |
| `worksheetName` | string | ❌ | 工作表名称 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "tableId": "Table1",
  "values": ["赵六", 35, "深圳"]
}
```

**响应数据**:

```typescript
interface AddTableRowResponse {
  requestId: string;
  success: boolean;
  data?: { tableId: string };
  error?: ErrorResponse;
  timestamp: number;
  duration?: number;
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 表格不存在（`details.kind:"table"`） |
| 3018 | `DATA_TYPE_MISMATCH` - 数据类型不匹配 |

---

### excel:delete:tableRow

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 删除表格中指定索引的行。

**请求数据**:

```typescript
interface DeleteTableRowRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  tableId: string;
  rowIndex: number;            // 要删除的行索引（从 0 开始，不含标题行）
  worksheetName?: string;
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `tableId` | string | ✅ | 表格名称或 ID |
| `rowIndex` | number | ✅ | 行索引（从 0 开始，不含标题行） |
| `worksheetName` | string | ❌ | 工作表名称 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "tableId": "Table1",
  "rowIndex": 0
}
```

**响应数据**:

```typescript
interface DeleteTableRowResponse {
  requestId: string;
  success: boolean;
  data?: { deleted: boolean };
  error?: ErrorResponse;
  timestamp: number;
  duration?: number;
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 表格不存在（`details.kind:"table"`） |
| 4004 | `PARAM_OUT_OF_RANGE` - 行索引超出范围 |

---

### excel:sort:table

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 对表格按指定列排序，支持多级排序。

**请求数据**:

```typescript
interface SortTableRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  tableId: string;
  sortFields: Array<{
    columnIndex: number;       // 列索引（从 0 开始）
    ascending?: boolean;       // 升序（默认 true）
  }>;
  worksheetName?: string;
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `tableId` | string | ✅ | 表格名称或 ID |
| `sortFields` | Array | ✅ | 排序字段数组，按优先级排列 |
| `sortFields[].columnIndex` | number | ✅ | 列索引（从 0 开始） |
| `sortFields[].ascending` | boolean | ❌ | 是否升序，默认 `true` |
| `worksheetName` | string | ❌ | 工作表名称 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "tableId": "Table1",
  "sortFields": [
    { "columnIndex": 1, "ascending": false },
    { "columnIndex": 0, "ascending": true }
  ]
}
```

**响应数据**:

```typescript
interface SortTableResponse {
  requestId: string;
  success: boolean;
  data?: { sorted: boolean };
  error?: ErrorResponse;
  timestamp: number;
  duration?: number;
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 表格不存在（`details.kind:"table"`） |
| 4004 | `PARAM_OUT_OF_RANGE` - 列索引超出范围 |

---

## Chart 操作类

### excel:insert:chart

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 根据数据范围创建图表。

**请求数据**:

```typescript
interface InsertChartRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  sourceAddress: string;       // 数据源范围
  chartType: string;           // 图表类型
  title?: string;              // 图表标题
  position?: {
    top?: number;              // 顶部位置（磅）
    left?: number;             // 左侧位置（磅）
    width?: number;            // 宽度（磅）
    height?: number;           // 高度（磅）
  };
  worksheetName?: string;
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `sourceAddress` | string | ✅ | 数据源范围地址 |
| `chartType` | string | ✅ | 图表类型（见下方类型列表） |
| `title` | string | ❌ | 图表标题 |
| `position` | object | ❌ | 图表位置和尺寸（单位：磅） |
| `worksheetName` | string | ❌ | 工作表名称 |

**常见图表类型**:

| 类型值 | 说明 |
|--------|------|
| `ColumnClustered` | 簇状柱形图 |
| `ColumnStacked` | 堆积柱形图 |
| `BarClustered` | 簇状条形图 |
| `Line` | 折线图 |
| `LineMarkers` | 带数据标记的折线图 |
| `Pie` | 饼图 |
| `Doughnut` | 圆环图 |
| `Area` | 面积图 |
| `XYScatter` | 散点图 |
| `Radar` | 雷达图 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "sourceAddress": "A1:C4",
  "chartType": "ColumnClustered",
  "title": "销售数据对比",
  "position": { "top": 200, "left": 300, "width": 400, "height": 300 }
}
```

**响应数据**:

```typescript
interface InsertChartResponse {
  requestId: string;
  success: boolean;
  data?: { name: string };      // 图表名称
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
  "data": { "name": "Chart 1" },
  "timestamp": 1704067200500,
  "duration": 150
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 工作表不存在（`details.kind:"worksheet"`） |
| 3009 | `RANGE_INVALID` - 无效的数据源范围 |
| 4002 | `INVALID_PARAM` - 无效的图表类型 |

---

### excel:get:charts

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 获取工作表中所有图表的信息。

**请求数据**:

```typescript
interface GetChartsRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  worksheetName?: string;
}
```

**响应数据**:

```typescript
interface GetChartsResponse {
  requestId: string;
  success: boolean;
  data?: {
    charts: Array<{
      name: string;            // 图表名称
      chartType: string;       // 图表类型
      title: string;           // 图表标题
      top: number;             // 顶部位置（磅）
      left: number;            // 左侧位置（磅）
      width: number;           // 宽度（磅）
      height: number;          // 高度（磅）
    }>;
  };
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
    "charts": [
      {
        "name": "Chart 1",
        "chartType": "ColumnClustered",
        "title": "销售数据",
        "top": 200,
        "left": 300,
        "width": 400,
        "height": 300
      }
    ]
  },
  "timestamp": 1704067200500,
  "duration": 35
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 工作表不存在（`details.kind:"worksheet"`） |

---

### excel:update:chart

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 更新图表属性。所有属性均为可选，仅传入的属性会被修改。

**请求数据**:

```typescript
interface UpdateChartRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  chartName: string;           // 图表名称
  properties: {
    title?: string;            // 新标题
    chartType?: string;        // 新图表类型
    sourceAddress?: string;    // 新数据源范围
    position?: {
      top?: number;
      left?: number;
      width?: number;
      height?: number;
    };
  };
  worksheetName?: string;
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `chartName` | string | ✅ | 图表名称 |
| `properties` | object | ✅ | 要更新的属性（全部可选） |
| `worksheetName` | string | ❌ | 工作表名称 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "chartName": "Chart 1",
  "properties": {
    "title": "2024年销售数据",
    "chartType": "Line"
  }
}
```

**响应数据**:

```typescript
interface UpdateChartResponse {
  requestId: string;
  success: boolean;
  data?: { name: string };
  error?: ErrorResponse;
  timestamp: number;
  duration?: number;
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 工作表不存在（`details.kind:"worksheet"`） |
| 3010 | `ELEMENT_NOT_FOUND` - 图表不存在（`details.kind:"chart"`） |
| 4002 | `INVALID_PARAM` - 无效的图表类型 |

---

### excel:delete:chart

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 删除指定图表。

**请求数据**:

```typescript
interface DeleteChartRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  chartName: string;
  worksheetName?: string;
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `chartName` | string | ✅ | 图表名称 |
| `worksheetName` | string | ❌ | 工作表名称 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "chartName": "Chart 1"
}
```

**响应数据**:

```typescript
interface DeleteChartResponse {
  requestId: string;
  success: boolean;
  data?: { deleted: boolean };
  error?: ErrorResponse;
  timestamp: number;
  duration?: number;
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 工作表不存在（`details.kind:"worksheet"`） |
| 3010 | `ELEMENT_NOT_FOUND` - 图表不存在（`details.kind:"chart"`） |

---

## PivotTable 操作类

### excel:insert:pivotTable

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 基于数据源范围创建透视表。

**请求数据**:

```typescript
interface InsertPivotTableRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  sourceAddress: string;       // 数据源范围
  targetAddress: string;       // 透视表放置位置
  name?: string;               // 透视表名称
  worksheetName?: string;
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `sourceAddress` | string | ✅ | 数据源范围地址 |
| `targetAddress` | string | ✅ | 透视表放置位置的起始单元格 |
| `name` | string | ❌ | 透视表名称，省略时自动生成 |
| `worksheetName` | string | ❌ | 工作表名称 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "sourceAddress": "A1:D100",
  "targetAddress": "F1",
  "name": "销售汇总"
}
```

**响应数据**:

```typescript
interface InsertPivotTableResponse {
  requestId: string;
  success: boolean;
  data?: { name: string };     // 透视表名称
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
  "data": { "name": "销售汇总" },
  "timestamp": 1704067200500,
  "duration": 200
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 工作表不存在（`details.kind:"worksheet"`） |
| 3009 | `RANGE_INVALID` - 无效的范围地址 |
| 3016 | `API_NOT_SUPPORTED` - 当前平台不支持透视表操作 |

---

### excel:get:pivotTables

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 获取工作表中所有透视表的列表。

**请求数据**:

```typescript
interface GetPivotTablesRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  worksheetName?: string;
}
```

**响应数据**:

```typescript
interface GetPivotTablesResponse {
  requestId: string;
  success: boolean;
  data?: {
    pivotTables: Array<{
      name: string;            // 透视表名称
      id: string;              // 透视表 ID
    }>;
  };
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
    "pivotTables": [
      { "name": "销售汇总", "id": "{abcd-...}" },
      { "name": "PivotTable2", "id": "{efgh-...}" }
    ]
  },
  "timestamp": 1704067200500,
  "duration": 25
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 工作表不存在（`details.kind:"worksheet"`） |

---

### excel:delete:pivotTable

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 删除指定透视表。

**请求数据**:

```typescript
interface DeletePivotTableRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  pivotTableName: string;      // 透视表名称
  worksheetName?: string;
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `pivotTableName` | string | ✅ | 透视表名称 |
| `worksheetName` | string | ❌ | 工作表名称 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "pivotTableName": "销售汇总"
}
```

**响应数据**:

```typescript
interface DeletePivotTableResponse {
  requestId: string;
  success: boolean;
  data?: { deleted: boolean };
  error?: ErrorResponse;
  timestamp: number;
  duration?: number;
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 工作表不存在（`details.kind:"worksheet"`） |
| 3010 | `ELEMENT_NOT_FOUND` - 透视表不存在（`details.kind:"pivotTable"`） |

---

## 查找与筛选类

### excel:find:values

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 在指定范围内搜索值，返回所有匹配的单元格地址和值。

**请求数据**:

```typescript
interface FindValuesRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  searchText: string;          // 搜索文本
  address?: string;            // 搜索范围，省略时搜索整个工作表
  worksheetName?: string;
  matchCase?: boolean;         // 是否区分大小写，默认 false
  matchEntireCell?: boolean;   // 是否完全匹配，默认 false
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `searchText` | string | ✅ | 搜索文本 |
| `address` | string | ❌ | 搜索范围，省略时搜索整个工作表 |
| `worksheetName` | string | ❌ | 工作表名称 |
| `matchCase` | boolean | ❌ | 是否区分大小写，默认 `false` |
| `matchEntireCell` | boolean | ❌ | 是否要求完全匹配单元格内容，默认 `false` |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "searchText": "张三",
  "matchEntireCell": true
}
```

**响应数据**:

```typescript
interface FindValuesResponse {
  requestId: string;
  success: boolean;
  data?: {
    matches: Array<{
      address: string;         // 匹配单元格地址
      value: unknown;          // 单元格值
    }>;
  };
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
    "matches": [
      { "address": "Sheet1!A2", "value": "张三" },
      { "address": "Sheet1!A15", "value": "张三" }
    ]
  },
  "timestamp": 1704067200500,
  "duration": 120
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 工作表不存在（`details.kind:"worksheet"`） |
| 3009 | `RANGE_INVALID` - 无效的搜索范围 |

---

### excel:set:autoFilter

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 对指定范围应用自动筛选。

**请求数据**:

```typescript
interface SetAutoFilterRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  address: string;             // 筛选范围
  criteria: Array<{
    columnIndex: number;       // 列索引（从 0 开始）
    filterOn: string;          // 筛选方式：Values, CellColor, FontColor 等
    values?: string[];         // 筛选值列表（filterOn 为 "Values" 时使用）
  }>;
  worksheetName?: string;
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `address` | string | ✅ | 筛选范围地址 |
| `criteria` | Array | ✅ | 筛选条件数组 |
| `criteria[].columnIndex` | number | ✅ | 列索引（从 0 开始） |
| `criteria[].filterOn` | string | ✅ | 筛选方式 |
| `criteria[].values` | string[] | ❌ | 筛选值列表 |
| `worksheetName` | string | ❌ | 工作表名称 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "address": "A1:C10",
  "criteria": [
    { "columnIndex": 2, "filterOn": "Values", "values": ["北京", "上海"] }
  ]
}
```

**响应数据**:

```typescript
interface SetAutoFilterResponse {
  requestId: string;
  success: boolean;
  data?: { address: string };
  error?: ErrorResponse;
  timestamp: number;
  duration?: number;
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 工作表不存在（`details.kind:"worksheet"`） |
| 3009 | `RANGE_INVALID` - 无效的范围地址 |
| 4004 | `PARAM_OUT_OF_RANGE` - 列索引超出范围 |

---

### excel:clear:autoFilter

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 清除工作表上的自动筛选。

**请求数据**:

```typescript
interface ClearAutoFilterRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  worksheetName?: string;
}
```

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx"
}
```

**响应数据**:

```typescript
interface ClearAutoFilterResponse {
  requestId: string;
  success: boolean;
  data?: { cleared: boolean };
  error?: ErrorResponse;
  timestamp: number;
  duration?: number;
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 工作表不存在（`details.kind:"worksheet"`） |

---

## 公式类

### excel:set:formula

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 设置单元格的公式。公式字符串由 AI 直接构造，工具透传应用。

**请求数据**:

```typescript
interface SetFormulaRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  address: string;             // 单元格地址
  formula: string;             // 公式字符串（含 = 号）
  worksheetName?: string;
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `address` | string | ✅ | 目标单元格地址 |
| `formula` | string | ✅ | 公式字符串，如 `"=SUM(A1:A10)"`、`"=VLOOKUP(A1,B:C,2,FALSE)"` |
| `worksheetName` | string | ❌ | 工作表名称 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "address": "D1",
  "formula": "=SUM(A1:C1)"
}
```

**响应数据**:

```typescript
interface SetFormulaResponse {
  requestId: string;
  success: boolean;
  data?: { address: string };
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
  "data": { "address": "Sheet1!D1" },
  "timestamp": 1704067200500,
  "duration": 30
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3010 | `ELEMENT_NOT_FOUND` - 工作表不存在（`details.kind:"worksheet"`） |
| 3009 | `RANGE_INVALID` - 无效的单元格地址 |
| 3003 | `DOCUMENT_READ_ONLY` - 工作表受保护（`details.scope:"worksheet"`） |
| 3017 | `FORMULA_ERROR` - 公式语法错误或引用无效 |
