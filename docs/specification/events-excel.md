# Excel 事件定义

!!! warning "Draft 状态"
    本文档中的所有事件处于 **Draft** 状态，接口可能在正式发布前发生变更。

## 概述

本章定义 `/excel` 命名空间下的所有事件。Excel 事件用于操作 Microsoft Excel 工作簿。

## 事件列表

### 内容检索类

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [excel:get:selectedRange](#excelgetselectedrange) | 📋 Draft | 获取选中范围 |
| [excel:get:usedRange](#excelgetusedrange) | 📋 Draft | 获取已使用范围 |
| [excel:get:cellValue](#excelgetcellvalue) | 📋 Draft | 获取单元格值 |
| [excel:get:rangeValues](#excelgetrangevalues) | 📋 Draft | 获取范围内的值 |

### 内容操作类

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [excel:set:cellValue](#excelsetcellvalue) | 📋 Draft | 设置单元格值 |
| [excel:set:rangeValues](#excelsetrangevalues) | 📋 Draft | 设置范围内的值 |
| [excel:insert:table](#excelInserttable) | 📋 Draft | 插入表格 |
| [excel:insert:chart](#excelinsertchart) | 📋 Draft | 插入图表 |

### 工作表管理类

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [excel:sheet:add](#excelsheetadd) | 📋 Draft | 添加工作表 |
| [excel:sheet:delete](#excelsheetdelete) | 📋 Draft | 删除工作表 |
| [excel:sheet:rename](#excelsheetrename) | 📋 Draft | 重命名工作表 |
| [excel:sheet:activate](#excelsheetactivate) | 📋 Draft | 激活工作表 |

---

## 内容检索类

### excel:get:selectedRange

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 获取当前选中的单元格范围信息。

**请求数据**:

```typescript
interface GetSelectedRangeRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
}
```

**响应数据**:

```typescript
interface GetSelectedRangeResponse {
  requestId: string;
  success: true;
  data: {
    range: RangeInfo;
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
    "range": {
      "address": "Sheet1!A1:C3",
      "rowCount": 3,
      "columnCount": 3,
      "worksheet": "Sheet1"
    }
  },
  "timestamp": 1704067200500,
  "duration": 50
}
```

---

### excel:get:usedRange

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 获取工作表中已使用的范围。

**请求数据**:

```typescript
interface GetUsedRangeRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  options?: {
    worksheet?: string;      // 工作表名称，默认当前活动工作表
  };
}
```

**响应数据**:

```typescript
interface GetUsedRangeResponse {
  requestId: string;
  success: true;
  data: {
    range: RangeInfo;
    values?: any[][];        // 范围内的值（可选）
  };
  timestamp: number;
  duration: number;
}
```

---

### excel:get:cellValue

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 获取指定单元格的值。

**请求数据**:

```typescript
interface GetCellValueRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  address: string;           // 单元格地址，如 "A1" 或 "Sheet1!A1"
}
```

**响应数据**:

```typescript
interface GetCellValueResponse {
  requestId: string;
  success: true;
  data: {
    address: string;
    value: any;              // 单元格值
    formattedValue: string;  // 格式化后的显示值
    formula?: string;        // 公式（如果有）
    type: CellValueType;     // 值类型
  };
  timestamp: number;
  duration: number;
}
```

---

### excel:get:rangeValues

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 获取指定范围内的所有值。

**请求数据**:

```typescript
interface GetRangeValuesRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  address: string;           // 范围地址，如 "A1:C3"
}
```

**响应数据**:

```typescript
interface GetRangeValuesResponse {
  requestId: string;
  success: true;
  data: {
    address: string;
    values: any[][];         // 二维数组
    rowCount: number;
    columnCount: number;
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
  "duration": 80
}
```

---

## 内容操作类

### excel:set:cellValue

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 设置指定单元格的值。

**请求数据**:

```typescript
interface SetCellValueRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  address: string;           // 单元格地址
  value: any;                // 要设置的值
  options?: {
    numberFormat?: string;   // 数字格式
  };
}
```

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "timestamp": 1704067200000,
  "address": "A1",
  "value": "Hello World"
}
```

**响应数据**:

```typescript
interface SetCellValueResponse {
  requestId: string;
  success: true;
  data: {
    address: string;
  };
  timestamp: number;
  duration: number;
}
```

---

### excel:set:rangeValues

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 设置指定范围内的值。

**请求数据**:

```typescript
interface SetRangeValuesRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  address: string;           // 范围起始地址
  values: any[][];           // 二维数组
}
```

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/data.xlsx",
  "timestamp": 1704067200000,
  "address": "A1",
  "values": [
    ["姓名", "年龄", "城市"],
    ["张三", 25, "北京"],
    ["李四", 30, "上海"]
  ]
}
```

---

### excel:insert:table

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 将指定范围转换为表格。

**请求数据**:

```typescript
interface ExcelInsertTableRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  address: string;           // 表格范围
  options?: {
    hasHeaders?: boolean;    // 第一行是否为标题，默认 true
    name?: string;           // 表格名称
    style?: string;          // 表格样式
  };
}
```

---

### excel:insert:chart

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 根据数据创建图表。

**请求数据**:

```typescript
interface InsertChartRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  dataRange: string;         // 数据范围
  chartType: ChartType;      // 图表类型
  options?: {
    title?: string;          // 图表标题
    position?: {
      left: number;
      top: number;
      width: number;
      height: number;
    };
  };
}
```

---

## 工作表管理类

### excel:sheet:add

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 添加新工作表。

**请求数据**:

```typescript
interface AddSheetRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  options?: {
    name?: string;           // 工作表名称
    position?: number;       // 插入位置
  };
}
```

---

### excel:sheet:delete

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 删除工作表。

**请求数据**:

```typescript
interface DeleteSheetRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  name: string;              // 要删除的工作表名称
}
```

---

### excel:sheet:rename

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 重命名工作表。

**请求数据**:

```typescript
interface RenameSheetRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  oldName: string;           // 原名称
  newName: string;           // 新名称
}
```

---

### excel:sheet:activate

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 激活指定工作表。

**请求数据**:

```typescript
interface ActivateSheetRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  name: string;              // 工作表名称
}
```

**响应数据**:

```typescript
interface ActivateSheetResponse {
  requestId: string;
  success: true;
  data: {
    name: string;            // 当前活动工作表名称
  };
  timestamp: number;
  duration: number;
}
```
