# PPT 事件定义

!!! warning "Draft 状态"
    本文档中的所有事件处于 **Draft** 状态，接口可能在正式发布前发生变更。

## 概述

本章定义 `/ppt` 命名空间下的所有事件。PPT 事件用于操作 Microsoft PowerPoint 演示文稿。

## 事件列表

### 事件报告类（AddIn → Server，单向）

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [ppt:event:slideChanged](#ppteventslidechanged) | 📋 Draft | 幻灯片切换通知 |

### 内容检索类（Server → AddIn，请求-响应）

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [ppt:get:currentSlideElements](#pptgetcurrentslideelements) | 📋 Draft | 获取当前幻灯片元素 |
| [ppt:get:slideElements](#pptgetslideelements) | 📋 Draft | 获取指定幻灯片元素 |
| [ppt:get:slideScreenshot](#pptgetslidescreenshot) | 📋 Draft | 获取幻灯片截图 |

### 内容操作类（Server → AddIn，请求-响应）

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [ppt:insert:text](#pptinserttext) | 📋 Draft | 插入文本框 |
| [ppt:insert:shape](#pptinsertshape) | 📋 Draft | 插入形状 |
| [ppt:insert:image](#pptinsertimage) | 📋 Draft | 插入图片 |
| [ppt:insert:table](#pptinserttable) | 📋 Draft | 插入表格 |
| [ppt:update:textBox](#pptupdatetextbox) | 📋 Draft | 更新文本框 |

### 幻灯片管理类（Server → AddIn，请求-响应）

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [ppt:add:slide](#pptaddslide) | 📋 Draft | 添加幻灯片 |
| [ppt:delete:slide](#pptdeleteslide) | 📋 Draft | 删除幻灯片 |
| [ppt:move:slide](#pptmoveslide) | 📋 Draft | 移动幻灯片 |
| [ppt:goto:slide](#pptgotoslide) | 📋 Draft | 跳转到幻灯片 |

---

## 事件报告类

### ppt:event:slideChanged

**方向**: AddIn → Server（单向通知）

**状态**: 📋 Draft

**说明**: 当用户在 PowerPoint 中切换幻灯片时触发。

**事件数据**:

```typescript
interface SlideChangedEvent {
  eventType: "slideChanged";      // 事件类型标识
  clientId: string;               // 客户端标识
  documentUri: string;            // 文档 URI
  timestamp: number;              // 事件发生时间（毫秒）
  data: {
    fromIndex: number;            // 切换前的幻灯片索引（从 0 开始）
    toIndex: number;              // 切换后的幻灯片索引（从 0 开始）
  };
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `eventType` | string | ✅ | 固定值 `"slideChanged"`，用于事件类型识别 |
| `clientId` | string | ✅ | 客户端唯一标识，用于区分多客户端场景 |
| `documentUri` | string | ✅ | 文档 URI（如 `file:///path/to/presentation.pptx`） |
| `timestamp` | number | ✅ | Unix 时间戳（毫秒） |
| `data.fromIndex` | number | ✅ | 切换前的幻灯片索引（从 0 开始） |
| `data.toIndex` | number | ✅ | 切换后的幻灯片索引（从 0 开始） |

**示例**:

```json
{
  "eventType": "slideChanged",
  "clientId": "ppt-addin-abc123",
  "documentUri": "file:///Users/john/Documents/presentation.pptx",
  "timestamp": 1704067200000,
  "data": {
    "fromIndex": 0,
    "toIndex": 2
  }
}
```

---

## 内容检索类

### ppt:get:currentSlideElements

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 获取当前显示幻灯片上的所有元素信息，包括文本框、图片、形状、占位符等。

**请求数据**:

```typescript
interface GetCurrentSlideElementsRequest {
  requestId: string;      // 请求 ID (UUID)
  documentUri: string;    // 文档 URI
  timestamp?: number;     // 请求时间戳（毫秒），可选
}
```

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/presentation.pptx",
  "timestamp": 1704067200000
}
```

**响应数据**:

```typescript
interface GetCurrentSlideElementsResponse {
  requestId: string;
  success: boolean;
  data?: {
    slideIndex: number;          // 当前幻灯片索引（从 0 开始）
    elements: SlideElement[];    // 元素数组
  };
  error?: ErrorResponse;
  timestamp: number;
}

interface SlideElement {
  id: string;                    // 元素唯一标识
  type: string;                  // 元素类型
  left: number;                  // X 坐标（磅）
  top: number;                   // Y 坐标（磅）
  width: number;                 // 宽度（磅）
  height: number;                // 高度（磅）
  name?: string;                 // 元素名称
  text?: string;                 // 文本内容（仅文本类元素）
  placeholderType?: string;      // 占位符类型（仅占位符元素）
}
```

**元素类型说明**:

| 类型 | 说明 |
|------|------|
| `TextBox` | 文本框 |
| `Image` | 图片 |
| `GeometricShape` | 几何形状 |
| `Placeholder` | 占位符（标题、正文等） |
| `Table` | 表格 |
| `Chart` | 图表 |

**占位符类型说明**:

| 类型 | 说明 |
|------|------|
| `Title` | 标题 |
| `Body` | 正文 |
| `Picture` | 图片占位符 |
| `SlideNumber` | 页码 |
| `Footer` | 页脚 |
| `Header` | 页眉 |

**响应示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "slideIndex": 0,
    "elements": [
      {
        "id": "shape-001",
        "type": "Placeholder",
        "left": 50,
        "top": 30,
        "width": 600,
        "height": 60,
        "name": "Title 1",
        "text": "演示文稿标题",
        "placeholderType": "Title"
      },
      {
        "id": "shape-002",
        "type": "Placeholder",
        "left": 50,
        "top": 120,
        "width": 600,
        "height": 300,
        "name": "Content Placeholder 2",
        "text": "正文内容...",
        "placeholderType": "Body"
      },
      {
        "id": "shape-003",
        "type": "Image",
        "left": 400,
        "top": 200,
        "width": 200,
        "height": 150,
        "name": "Picture 3"
      }
    ]
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3999 | `OFFICE_API_ERROR` - Office API 调用错误 |

---

### ppt:get:slideElements

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 获取指定幻灯片上的所有元素信息，支持按元素类型过滤。

**请求数据**:

```typescript
interface GetSlideElementsRequest {
  requestId: string;      // 请求 ID (UUID)
  documentUri: string;    // 文档 URI
  timestamp?: number;     // 请求时间戳（毫秒），可选
  slideIndex: number;     // 幻灯片索引（从 0 开始）
  options?: SlideElementsOptions;
}

interface SlideElementsOptions {
  includeText?: boolean;      // 是否包含文本内容，默认 true
  includeImages?: boolean;    // 是否包含图片元素，默认 true
  includeShapes?: boolean;    // 是否包含形状元素，默认 true
  includeTables?: boolean;    // 是否包含表格元素，默认 true
  includeCharts?: boolean;    // 是否包含图表元素，默认 true
}
```

**请求参数说明**:

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `slideIndex` | number | - | 目标幻灯片索引（从 0 开始），必填 |
| `includeText` | boolean | true | 是否在元素中包含文本内容 |
| `includeImages` | boolean | true | 是否返回图片类型元素 |
| `includeShapes` | boolean | true | 是否返回形状类型元素 |
| `includeTables` | boolean | true | 是否返回表格类型元素 |
| `includeCharts` | boolean | true | 是否返回图表类型元素 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/presentation.pptx",
  "slideIndex": 2,
  "options": {
    "includeText": true,
    "includeImages": true
  }
}
```

**响应数据**:

```typescript
interface GetSlideElementsResponse {
  requestId: string;
  success: boolean;
  data?: {
    slideIndex: number;          // 幻灯片索引
    elements: SlideElement[];    // 元素数组（与 ppt:get:currentSlideElements 相同）
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
    "slideIndex": 2,
    "elements": [
      {
        "id": "shape-010",
        "type": "TextBox",
        "left": 100,
        "top": 200,
        "width": 300,
        "height": 50,
        "name": "TextBox 1",
        "text": "自定义文本"
      }
    ]
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 4002 | `INVALID_PARAM` - slideIndex 超出范围 |
| 3999 | `OFFICE_API_ERROR` - Office API 调用错误 |

!!! note "与 ppt:get:currentSlideElements 的关系"
    本事件与 `ppt:get:currentSlideElements` 返回相同的 `SlideElement` 结构。
    区别在于本事件可以指定任意幻灯片索引，并支持通过 `options` 过滤元素类型。

---

### ppt:get:slideScreenshot

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 获取幻灯片的截图，返回 Base64 编码的图片数据。

**请求数据**:

```typescript
interface GetSlideScreenshotRequest {
  requestId: string;      // 请求 ID (UUID)
  documentUri: string;    // 文档 URI
  timestamp?: number;     // 请求时间戳（毫秒），可选
  slideIndex: number;     // 幻灯片索引（从 0 开始）
  options?: ScreenshotOptions;
}

interface ScreenshotOptions {
  format?: "png" | "jpeg";    // 图片格式，默认 "png"
  quality?: number;           // 图片质量（0-100），仅 jpeg 有效
}
```

**请求参数说明**:

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `slideIndex` | number | - | 目标幻灯片索引（从 0 开始），必填 |
| `format` | string | `"png"` | 输出图片格式 |
| `quality` | number | - | JPEG 图片质量（0-100） |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/presentation.pptx",
  "slideIndex": 0,
  "options": {
    "format": "png"
  }
}
```

**响应数据**:

```typescript
interface GetSlideScreenshotResponse {
  requestId: string;
  success: boolean;
  data?: {
    base64: string;      // Base64 编码的图片数据（不含 data URL 前缀）
    format: string;      // 图片格式（"png" 或 "jpeg"）
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
    "base64": "iVBORw0KGgoAAAANSUhEUgAA...",
    "format": "png"
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 4002 | `INVALID_PARAM` - slideIndex 超出范围 |
| 3999 | `OFFICE_API_ERROR` - Office API 调用错误 |

---

## 内容操作类

### ppt:insert:text

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 在指定幻灯片上插入文本框。

**请求数据**:

```typescript
interface InsertTextRequest {
  requestId: string;         // 请求 ID (UUID)
  documentUri: string;       // 文档 URI
  timestamp?: number;        // 请求时间戳（毫秒），可选
  text: string;              // 要插入的文本内容
  options?: TextInsertOptions;
}

interface TextInsertOptions {
  slideIndex?: number;       // 目标幻灯片索引（从 0 开始），默认当前幻灯片
  left?: number;             // X 坐标（磅）
  top?: number;              // Y 坐标（磅）
  width?: number;            // 文本框宽度（磅），默认 300
  height?: number;           // 文本框高度（磅），默认 100
  fontSize?: number;         // 字号
  fontName?: string;         // 字体名称
  color?: string;            // 文字颜色（十六进制，如 "#FF0000"）
  fillColor?: string;        // 填充颜色（十六进制）
}
```

**请求参数说明**:

| 参数 | 类型 | 必需 | 默认值 | 说明 |
|------|------|------|--------|------|
| `text` | string | ✅ | - | 要插入的文本内容 |
| `slideIndex` | number | ❌ | 当前幻灯片 | 目标幻灯片索引（从 0 开始） |
| `left` | number | ❌ | - | X 坐标（磅），未指定时使用默认位置 |
| `top` | number | ❌ | - | Y 坐标（磅），未指定时使用默认位置 |
| `width` | number | ❌ | 300 | 文本框宽度（磅） |
| `height` | number | ❌ | 100 | 文本框高度（磅） |
| `fontSize` | number | ❌ | - | 字号 |
| `fontName` | string | ❌ | - | 字体名称 |
| `color` | string | ❌ | - | 文字颜色（十六进制） |
| `fillColor` | string | ❌ | - | 文本框填充颜色（十六进制） |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/presentation.pptx",
  "text": "这是新插入的文本",
  "options": {
    "slideIndex": 0,
    "left": 100,
    "top": 200,
    "width": 400,
    "height": 80,
    "fontSize": 18,
    "fontName": "微软雅黑",
    "color": "#333333"
  }
}
```

**响应数据**:

```typescript
interface InsertTextResponse {
  requestId: string;
  success: boolean;
  data?: {
    elementId: string;       // 创建的文本框元素 ID
    slideIndex: number;      // 插入的幻灯片索引
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
    "elementId": "shape-015",
    "slideIndex": 0
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4001 | `MISSING_PARAM` - 缺少 text 参数 |
| 4002 | `INVALID_PARAM` - slideIndex 超出范围 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3999 | `OFFICE_API_ERROR` - Office API 调用错误 |

---

### ppt:insert:shape

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 在指定幻灯片上插入几何形状。

**请求数据**:

```typescript
interface InsertShapeRequest {
  requestId: string;         // 请求 ID (UUID)
  documentUri: string;       // 文档 URI
  timestamp?: number;        // 请求时间戳（毫秒），可选
  shapeType: ShapeType;      // 形状类型
  options?: ShapeInsertOptions;
}

type ShapeType =
  | "Rectangle"
  | "RoundedRectangle"
  | "Circle"
  | "Oval"
  | "Triangle"
  | "Line"
  | "Arrow"
  | "Star"
  | "TextBox";

interface ShapeInsertOptions {
  slideIndex?: number;       // 目标幻灯片索引（从 0 开始），默认当前幻灯片
  left?: number;             // X 坐标（磅），默认居中
  top?: number;              // Y 坐标（磅），默认居中
  width?: number;            // 宽度（磅），默认 100
  height?: number;           // 高度（磅），默认 100
  fillColor?: string;        // 填充颜色（十六进制），默认 "#4472C4"
  borderColor?: string;      // 边框颜色（十六进制），默认 "#2E5090"
  borderWidth?: number;      // 边框宽度（磅），默认 2
  text?: string;             // 形状内文本
}
```

**形状类型说明**:

| 类型 | 说明 |
|------|------|
| `Rectangle` | 矩形 |
| `RoundedRectangle` | 圆角矩形 |
| `Circle` | 圆形 |
| `Oval` | 椭圆 |
| `Triangle` | 三角形 |
| `Line` | 线条 |
| `Arrow` | 箭头 |
| `Star` | 星形 |
| `TextBox` | 文本框 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/presentation.pptx",
  "shapeType": "RoundedRectangle",
  "options": {
    "slideIndex": 0,
    "left": 200,
    "top": 150,
    "width": 200,
    "height": 100,
    "fillColor": "#4472C4",
    "borderColor": "#2E5090",
    "text": "点击这里"
  }
}
```

**响应数据**:

```typescript
interface InsertShapeResponse {
  requestId: string;
  success: boolean;
  data?: {
    shapeId: string;         // 创建的形状元素 ID
    slideIndex: number;      // 插入的幻灯片索引
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
    "shapeId": "shape-020",
    "slideIndex": 0
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4001 | `MISSING_PARAM` - 缺少 shapeType 参数 |
| 4002 | `INVALID_PARAM` - shapeType 不支持或 slideIndex 超出范围 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3999 | `OFFICE_API_ERROR` - Office API 调用错误 |

---

### ppt:insert:image

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 在指定幻灯片上插入图片。

**请求数据**:

```typescript
interface InsertImageRequest {
  requestId: string;         // 请求 ID (UUID)
  documentUri: string;       // 文档 URI
  timestamp?: number;        // 请求时间戳（毫秒），可选
  image: SlideImageData;     // 图片数据
  options?: ElementInsertOptions;
}

interface SlideImageData {
  base64: string;            // Base64 编码的图片数据（含或不含 data URL 前缀）
}

interface ElementInsertOptions {
  slideIndex?: number;       // 目标幻灯片索引（从 0 开始），默认当前幻灯片
  left?: number;             // X 坐标（磅）
  top?: number;              // Y 坐标（磅）
  width?: number;            // 宽度（磅），默认 200
  height?: number;           // 高度（磅），默认 150
}
```

**请求参数说明**:

| 参数 | 类型 | 必需 | 默认值 | 说明 |
|------|------|------|--------|------|
| `image.base64` | string | ✅ | - | Base64 编码的图片数据 |
| `slideIndex` | number | ❌ | 当前幻灯片 | 目标幻灯片索引 |
| `left` | number | ❌ | - | X 坐标（磅） |
| `top` | number | ❌ | - | Y 坐标（磅） |
| `width` | number | ❌ | 200 | 图片宽度（磅） |
| `height` | number | ❌ | 150 | 图片高度（磅） |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/presentation.pptx",
  "image": {
    "base64": "iVBORw0KGgoAAAANSUhEUgAA..."
  },
  "options": {
    "slideIndex": 0,
    "left": 300,
    "top": 200,
    "width": 400,
    "height": 300
  }
}
```

**响应数据**:

```typescript
interface InsertImageResponse {
  requestId: string;
  success: boolean;
  data?: {
    imageId: string;         // 创建的图片元素 ID
    slideIndex: number;      // 插入的幻灯片索引
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
    "imageId": "shape-025",
    "slideIndex": 0
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4001 | `MISSING_PARAM` - 缺少 image.base64 参数 |
| 4002 | `INVALID_PARAM` - Base64 数据无效或 slideIndex 超出范围 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3999 | `OFFICE_API_ERROR` - Office API 调用错误 |

---

### ppt:insert:table

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 在指定幻灯片上插入表格，支持初始数据填充。

**请求数据**:

```typescript
interface InsertTableRequest {
  requestId: string;         // 请求 ID (UUID)
  documentUri: string;       // 文档 URI
  timestamp?: number;        // 请求时间戳（毫秒），可选
  options: SlideTableInsertOptions;
}

interface SlideTableInsertOptions {
  rows: number;              // 行数（1-100）
  columns: number;           // 列数（1-50）
  slideIndex?: number;       // 目标幻灯片索引（从 0 开始），默认当前幻灯片
  left?: number;             // X 坐标（磅），默认居中
  top?: number;              // Y 坐标（磅），默认居中
  data?: string[][];         // 初始数据（二维数组，按行列顺序）
}
```

**请求参数说明**:

| 参数 | 类型 | 必需 | 默认值 | 说明 |
|------|------|------|--------|------|
| `rows` | number | ✅ | - | 行数（1-100） |
| `columns` | number | ✅ | - | 列数（1-50） |
| `slideIndex` | number | ❌ | 当前幻灯片 | 目标幻灯片索引 |
| `left` | number | ❌ | 居中 | X 坐标（磅） |
| `top` | number | ❌ | 居中 | Y 坐标（磅） |
| `data` | string[][] | ❌ | - | 初始数据，维度需与 rows/columns 匹配 |

!!! warning "数据维度校验"
    当提供 `data` 参数时，数组维度必须与 `rows` × `columns` 精确匹配，否则返回校验错误。

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/presentation.pptx",
  "options": {
    "rows": 3,
    "columns": 4,
    "slideIndex": 0,
    "data": [
      ["姓名", "年龄", "城市", "职业"],
      ["张三", "28", "北京", "工程师"],
      ["李四", "32", "上海", "设计师"]
    ]
  }
}
```

**响应数据**:

```typescript
interface InsertTableResponse {
  requestId: string;
  success: boolean;
  data?: {
    elementId: string;       // 创建的表格元素 ID
    rowCount: number;        // 行数
    columnCount: number;     // 列数
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
    "elementId": "shape-030",
    "rowCount": 3,
    "columnCount": 4
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4000 | `VALIDATION_ERROR` - data 维度与 rows/columns 不匹配 |
| 4001 | `MISSING_PARAM` - 缺少 rows 或 columns |
| 4002 | `INVALID_PARAM` - rows 超过 100 或 columns 超过 50 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3999 | `OFFICE_API_ERROR` - Office API 调用错误 |

---

### ppt:update:textBox

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 更新幻灯片中现有文本框的内容或样式。支持更新 TextBox、Placeholder、GeometricShape 类型的元素。

**请求数据**:

```typescript
interface UpdateTextBoxRequest {
  requestId: string;         // 请求 ID (UUID)
  documentUri: string;       // 文档 URI
  timestamp?: number;        // 请求时间戳（毫秒），可选
  elementId: string;         // 要更新的元素 ID
  updates: TextBoxUpdates;
}

interface TextBoxUpdates {
  text?: string;             // 新文本内容
  fontSize?: number;         // 字号
  fontName?: string;         // 字体名称
  color?: string;            // 文字颜色（十六进制，如 "#FF0000"）
  fillColor?: string;        // 填充颜色（十六进制）
  bold?: boolean;            // 粗体
  italic?: boolean;          // 斜体
}
```

**请求参数说明**:

| 参数 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `elementId` | string | ✅ | 要更新的元素 ID（可通过 `ppt:get:slideElements` 获取） |
| `text` | string | ❌ | 新文本内容 |
| `fontSize` | number | ❌ | 字号 |
| `fontName` | string | ❌ | 字体名称 |
| `color` | string | ❌ | 文字颜色（十六进制） |
| `fillColor` | string | ❌ | 文本框填充颜色（十六进制） |
| `bold` | boolean | ❌ | 是否粗体 |
| `italic` | boolean | ❌ | 是否斜体 |

!!! note "支持的元素类型"
    仅支持 `TextBox`、`Placeholder`、`GeometricShape` 类型的元素。
    对不支持文本的元素类型将返回错误。

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/presentation.pptx",
  "elementId": "shape-001",
  "updates": {
    "text": "更新后的标题",
    "fontSize": 28,
    "bold": true,
    "color": "#333333"
  }
}
```

**响应数据**:

```typescript
interface UpdateTextBoxResponse {
  requestId: string;
  success: boolean;
  data?: {
    elementId: string;       // 更新的元素 ID
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
    "elementId": "shape-001"
  },
  "timestamp": 1704067200500
}
```

**响应示例（元素未找到）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": false,
  "error": {
    "code": "3003",
    "message": "Element not found: shape-999"
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4001 | `MISSING_PARAM` - 缺少 elementId |
| 3003 | `OPERATION_FAILED` - 元素未找到或元素类型不支持文本编辑 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3999 | `OFFICE_API_ERROR` - Office API 调用错误 |

---

## 幻灯片管理类

### ppt:add:slide

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

!!! note "类型定义状态"
    本事件尚未在 `socketio-types.ts` 中定义 Request/Response 类型。以下接口为规划设计，待实现时同步添加。

**说明**: 添加新幻灯片，支持指定版式。

**请求数据**:

```typescript
interface AddSlideRequest {
  requestId: string;         // 请求 ID (UUID)
  documentUri: string;       // 文档 URI
  timestamp?: number;        // 请求时间戳（毫秒），可选
  options?: {
    insertIndex?: number;    // 插入位置索引（从 0 开始），默认末尾
    layout?: string;         // 版式名称（如 "Title Slide", "Blank"）
  };
}
```

**请求参数说明**:

| 参数 | 类型 | 必需 | 默认值 | 说明 |
|------|------|------|--------|------|
| `insertIndex` | number | ❌ | 末尾 | 插入位置索引（从 0 开始） |
| `layout` | string | ❌ | - | 版式名称，可通过幻灯片母版获取可用版式 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/presentation.pptx",
  "options": {
    "insertIndex": 2,
    "layout": "Title Slide"
  }
}
```

**响应数据**:

```typescript
interface AddSlideResponse {
  requestId: string;
  success: boolean;
  data?: {
    slideIndex: number;      // 新幻灯片的索引
    slideId: string;         // 新幻灯片的 ID
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
    "slideIndex": 2,
    "slideId": "slide-003"
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4002 | `INVALID_PARAM` - insertIndex 超出范围或 layout 不存在 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3999 | `OFFICE_API_ERROR` - Office API 调用错误 |

---

### ppt:delete:slide

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 删除指定的幻灯片。

**请求数据**:

```typescript
interface DeleteSlideRequest {
  requestId: string;         // 请求 ID (UUID)
  documentUri: string;       // 文档 URI
  timestamp?: number;        // 请求时间戳（毫秒），可选
  slideIndex: number;        // 要删除的幻灯片索引（从 0 开始）
}
```

**请求参数说明**:

| 参数 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `slideIndex` | number | ✅ | 要删除的幻灯片索引（从 0 开始） |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/presentation.pptx",
  "slideIndex": 3
}
```

**响应数据**:

```typescript
interface DeleteSlideResponse {
  requestId: string;
  success: boolean;
  data?: {
    deleted: boolean;        // 是否成功删除
    totalSlides: number;     // 删除后的总幻灯片数
  };
  error?: ErrorResponse;
  timestamp: number;
}
```

!!! note "类型定义状态"
    本事件的 Response 类型尚未在 `socketio-types.ts` 中定义，待实现时同步添加。

**响应示例（成功）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "deleted": true,
    "totalSlides": 9
  },
  "timestamp": 1704067200500
}
```

**响应示例（索引超出范围）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": false,
  "error": {
    "code": "4002",
    "message": "Slide index 15 out of range, total slides: 10"
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4001 | `MISSING_PARAM` - 缺少 slideIndex |
| 4002 | `INVALID_PARAM` - slideIndex 超出范围 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3999 | `OFFICE_API_ERROR` - Office API 调用错误 |

---

### ppt:move:slide

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 移动幻灯片到指定位置。

**请求数据**:

```typescript
interface MoveSlideRequest {
  requestId: string;         // 请求 ID (UUID)
  documentUri: string;       // 文档 URI
  timestamp?: number;        // 请求时间戳（毫秒），可选
  fromIndex: number;         // 原位置索引（从 0 开始）
  toIndex: number;           // 目标位置索引（从 0 开始）
}
```

**请求参数说明**:

| 参数 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `fromIndex` | number | ✅ | 原位置索引（从 0 开始） |
| `toIndex` | number | ✅ | 目标位置索引（从 0 开始） |

!!! warning "索引约束"
    - `fromIndex` 和 `toIndex` 不能相同
    - 两个索引都必须在有效范围内（0 至 totalSlides-1）

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/presentation.pptx",
  "fromIndex": 0,
  "toIndex": 3
}
```

**响应数据**:

```typescript
interface MoveSlideResponse {
  requestId: string;
  success: boolean;
  data?: {
    fromIndex: number;       // 原位置索引
    toIndex: number;         // 目标位置索引
    totalSlides: number;     // 总幻灯片数
  };
  error?: ErrorResponse;
  timestamp: number;
}
```

!!! note "类型定义状态"
    本事件的 Response 类型尚未在 `socketio-types.ts` 中定义，待实现时同步添加。

**响应示例（成功）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "fromIndex": 0,
    "toIndex": 3,
    "totalSlides": 10
  },
  "timestamp": 1704067200500
}
```

**响应示例（相同位置）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": false,
  "error": {
    "code": "4002",
    "message": "fromIndex and toIndex cannot be the same"
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4001 | `MISSING_PARAM` - 缺少 fromIndex 或 toIndex |
| 4002 | `INVALID_PARAM` - 索引超出范围或 fromIndex === toIndex |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3999 | `OFFICE_API_ERROR` - Office API 调用错误 |

---

### ppt:goto:slide

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

!!! note "类型定义状态"
    本事件尚未在 `socketio-types.ts` 中定义 Request/Response 类型。以下接口为规划设计，待实现时同步添加。

**说明**: 跳转到指定幻灯片，使其成为当前显示的幻灯片。

**请求数据**:

```typescript
interface GotoSlideRequest {
  requestId: string;         // 请求 ID (UUID)
  documentUri: string;       // 文档 URI
  timestamp?: number;        // 请求时间戳（毫秒），可选
  slideIndex: number;        // 目标幻灯片索引（从 0 开始）
}
```

**请求参数说明**:

| 参数 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `slideIndex` | number | ✅ | 目标幻灯片索引（从 0 开始） |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/presentation.pptx",
  "slideIndex": 5
}
```

**响应数据**:

```typescript
interface GotoSlideResponse {
  requestId: string;
  success: boolean;
  data?: {
    slideIndex: number;      // 当前幻灯片索引
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
    "slideIndex": 5
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4001 | `MISSING_PARAM` - 缺少 slideIndex |
| 4002 | `INVALID_PARAM` - slideIndex 超出范围 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3999 | `OFFICE_API_ERROR` - Office API 调用错误 |
