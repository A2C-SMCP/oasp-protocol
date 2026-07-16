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
| [ppt:get:slideInfo](#pptgetslideinfo) | 📋 Draft | 获取演示文稿/幻灯片基本信息 |
| [ppt:get:slideLayouts](#pptgetslidelayouts) | 📋 Draft | 获取可用幻灯片版式列表 |

### 内容操作类（Server → AddIn，请求-响应）

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [ppt:insert:text](#pptinserttext) | 📋 Draft | 插入文本框 |
| [ppt:insert:shape](#pptinsertshape) | 📋 Draft | 插入形状 |
| [ppt:insert:image](#pptinsertimage) | 📋 Draft | 插入图片 |
| [ppt:insert:table](#pptinserttable) | 📋 Draft | 插入表格 |
| [ppt:insert:chart](#pptinsertchart) | 📋 Draft | 插入图表（柱形/折线/饼/散点等） |
| [ppt:update:textBox](#pptupdatetextbox) | 📋 Draft | 更新文本框内容/样式 |
| [ppt:delete:element](#pptdeleteelement) | 📋 Draft | 删除指定元素（亦适用于图表删除） |
| [ppt:update:image](#pptupdateimage) | 📋 Draft | 替换图片内容 |
| [ppt:update:tableCell](#pptupdatetablecell) | 📋 Draft | 更新表格单元格 |
| [ppt:update:tableRowColumn](#pptupdatetablerowcolumn) | 📋 Draft | 更新表格行/列内容 |
| [ppt:get:chart](#pptgetchart) | 📋 Draft | 读取图表数据 |
| [ppt:update:chart](#pptupdatechart) | 📋 Draft | 更新图表数据/类型/标题 |

### 样式格式类（Server → AddIn，请求-响应）

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [ppt:update:tableFormat](#pptupdatetableformat) | 📋 Draft | 更新表格样式 |

### 布局操作类（Server → AddIn，请求-响应）

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [ppt:update:element](#pptupdateelement) | 📋 Draft | 更新元素位置/尺寸 |
| [ppt:reorder:element](#pptreorderelement) | 📋 Draft | 调整元素层叠顺序 |

### 幻灯片管理类（Server → AddIn，请求-响应）

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [ppt:add:slide](#pptaddslide) | 📋 Draft | 添加幻灯片 |
| [ppt:delete:slide](#pptdeleteslide) | 📋 Draft | 删除幻灯片 |
| [ppt:move:slide](#pptmoveslide) | 📋 Draft | 移动幻灯片 |
| [ppt:goto:slide](#pptgotoslide) | 📋 Draft | 跳转到幻灯片 |
| [ppt:get:slideOoxml](#pptgetslideooxml) | 📋 Draft | 导出单页实时 OOXML（base64），含未保存态 |
| [ppt:insert:slidesOoxml](#pptinsertslidesooxml) | 📋 Draft | 应用 OOXML 页包；可选替换旧页 + 复位（顺序复合 round-trip，非原子） |

!!! info "关于 ppt:insert:video"
    PowerPoint JavaScript API **不支持**插入视频/音频元素。此功能标记为 🚫 Not Feasible，不在事件列表中定义。

### 脚本执行类（Server → AddIn，请求-响应）

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [ppt:run:script](#pptrunscript) | 📋 Draft | 执行原始 Office.js 脚本（封装层逃生舱） |

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
  rotation?: number;             // 旋转角度（度）
  zOrder?: number;               // 层叠顺序
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
| 3000 | `OFFICE_API_ERROR` - Office API 调用错误 |

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
| 3000 | `OFFICE_API_ERROR` - Office API 调用错误 |

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
| 3000 | `OFFICE_API_ERROR` - Office API 调用错误 |

---

### ppt:get:slideInfo

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 获取演示文稿的基本信息（幻灯片总数、尺寸），或获取指定幻灯片的详细布局信息（版式、元素列表、背景等）。AI 进行布局计算时必须先知道幻灯片尺寸、总页数和元素分布。

**请求数据**:

```typescript
interface GetSlideInfoRequest {
  requestId: string;      // 请求 ID (UUID)
  documentUri: string;    // 文档 URI
  timestamp?: number;     // 请求时间戳（毫秒），可选
  slideIndex?: number;    // 幻灯片索引（从 0 开始），可选。指定时返回该幻灯片详细信息
}
```

**请求参数说明**:

| 参数 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `slideIndex` | number | ❌ | 幻灯片索引（从 0 开始）。指定时响应中包含 `slideInfo` |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/presentation.pptx",
  "slideIndex": 0
}
```

**响应数据**:

```typescript
interface GetSlideInfoResponse {
  requestId: string;
  success: boolean;
  data?: {
    slideCount: number;          // 幻灯片总数
    dimensions: {
      width: number;             // 幻灯片宽度（磅）
      height: number;            // 幻灯片高度（磅）
      aspectRatio: string;       // 宽高比，如 "16:9", "4:3"
      isFromAPI: boolean;        // 是否通过 Office API 获取（false 表示使用默认值降级）
    };
    currentSlideIndex: number;   // 当前显示的幻灯片索引
    slideInfo?: {                // 仅当请求中指定 slideIndex 时返回
      slideIndex: number;
      slideId: string;
      layout: {
        name: string;            // 版式名称（如 "Title Slide", "Blank"）
        type: string;            // 版式类型
      };
      elements: SlideElement[];  // 该幻灯片上的所有元素
      background?: SlideBackgroundInfo;  // 幻灯片背景信息
    };
  };
  error?: ErrorResponse;
  timestamp: number;
}

interface SlideBackgroundInfo {
  type: "solid" | "gradient" | "image" | "pattern" | "none" | "unknown";
  color?: string;            // 背景颜色（仅 solid 类型）
  imageData?: string;        // Base64 或 URL（仅 image 类型）
}
```

!!! note "SlideElement 扩展字段"
    当通过 `ppt:get:slideInfo` 返回时，`SlideElement` 可包含以下扩展字段（详见[元素类型说明](#slideelementelementextensions)）：
    `relativePosition`（相对位置百分比）、`textInfo`（结构化文本信息）、`imageInfo`（图片信息）、`fillInfo`（填充信息）。

#### SlideElement 扩展字段说明 {#slideelementelementextensions}

```typescript
interface SlideElement {
  // 基础字段（所有 get 事件通用）
  id: string;                    // 元素唯一标识
  type: string;                  // 元素类型
  left: number;                  // X 坐标（磅）
  top: number;                   // Y 坐标（磅）
  width: number;                 // 宽度（磅）
  height: number;                // 高度（磅）
  name?: string;                 // 元素名称
  text?: string;                 // 文本内容（纯文本，仅文本类元素）
  placeholderType?: string;      // 占位符类型（仅占位符元素）
  rotation?: number;             // 旋转角度（度）
  zOrder?: number;               // 层叠顺序
  // 扩展字段（ppt:get:slideInfo 可返回）
  relativePosition?: {           // 相对于幻灯片的百分比位置
    leftPercent: number;
    topPercent: number;
    widthPercent: number;
    heightPercent: number;
  };
  textInfo?: {                   // 结构化文本信息（比 text 字段更详细）
    content: string;
    fontSize?: number;
    fontFamily?: string;
    color?: string;
    alignment?: string;
  };
  imageInfo?: {                  // 图片信息（仅图片类元素）
    format: "picture" | "picture-placeholder" | "picture-fill";
    data?: string;               // Base64 编码（需显式请求）
    url?: string;                // 外部链接（如有）
  };
  fillInfo?: {                   // 填充信息
    type: "solid" | "gradient" | "image" | "none" | "unknown";
    color?: string;
  };
}
```

**响应示例（不指定 slideIndex）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "slideCount": 10,
    "dimensions": {
      "width": 960,
      "height": 540,
      "aspectRatio": "16:9",
      "isFromAPI": true
    },
    "currentSlideIndex": 2
  },
  "timestamp": 1704067200500
}
```

**响应示例（指定 slideIndex）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "slideCount": 10,
    "dimensions": {
      "width": 960,
      "height": 540,
      "aspectRatio": "16:9",
      "isFromAPI": true
    },
    "currentSlideIndex": 2,
    "slideInfo": {
      "slideIndex": 0,
      "slideId": "slide-001",
      "layout": { "name": "Title Slide", "type": "TitleSlide" },
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
          "placeholderType": "Title",
          "zOrder": 0,
          "relativePosition": {
            "leftPercent": 5.21,
            "topPercent": 5.56,
            "widthPercent": 62.50,
            "heightPercent": 11.11
          },
          "textInfo": {
            "content": "演示文稿标题",
            "fontSize": 36,
            "fontFamily": "微软雅黑"
          }
        },
        {
          "id": "shape-003",
          "type": "Image",
          "left": 400,
          "top": 200,
          "width": 200,
          "height": 150,
          "name": "Picture 3",
          "zOrder": 2,
          "relativePosition": {
            "leftPercent": 41.67,
            "topPercent": 37.04,
            "widthPercent": 20.83,
            "heightPercent": 27.78
          },
          "imageInfo": {
            "format": "picture"
          }
        }
      ]
    }
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 4002 | `INVALID_PARAM` - slideIndex 超出范围 |
| 3000 | `OFFICE_API_ERROR` - Office API 调用错误 |

---

### ppt:get:slideLayouts

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 获取当前演示文稿中所有可用的幻灯片版式（Layout）列表。AI 在调用 `ppt:add:slide` 时需要先知道有哪些可用版式。

**请求数据**:

```typescript
interface GetSlideLayoutsRequest {
  requestId: string;      // 请求 ID (UUID)
  documentUri: string;    // 文档 URI
  timestamp?: number;     // 请求时间戳（毫秒），可选
  options?: {
    includePlaceholders?: boolean;  // 是否包含占位符详细信息，默认 true
  };
}
```

**请求参数说明**:

| 参数 | 类型 | 必需 | 默认值 | 说明 |
|------|------|------|--------|------|
| `includePlaceholders` | boolean | ❌ | true | 是否返回每个版式的占位符类型信息 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/presentation.pptx"
}
```

**响应数据**:

```typescript
interface GetSlideLayoutsResponse {
  requestId: string;
  success: boolean;
  data?: {
    layouts: SlideLayoutTemplate[];
  };
  error?: ErrorResponse;
  timestamp: number;
}

interface SlideLayoutTemplate {
  id: string;                    // 版式 ID
  name: string;                  // 版式名称（如 "Title Slide", "Blank"）
  type: string;                  // 版式类型
  placeholderCount: number;      // 占位符数量
  placeholderTypes: string[];    // 占位符类型列表（如 ["title", "body"]）
  isCustom: boolean;             // 是否为自定义版式
}
```

**响应示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "layouts": [
      {
        "id": "layout-001",
        "name": "Title Slide",
        "type": "title",
        "placeholderCount": 2,
        "placeholderTypes": ["title", "body"],
        "isCustom": false
      },
      {
        "id": "layout-002",
        "name": "Blank",
        "type": "blank",
        "placeholderCount": 0,
        "placeholderTypes": [],
        "isCustom": false
      },
      {
        "id": "layout-003",
        "name": "Title and Content",
        "type": "titleAndContent",
        "placeholderCount": 2,
        "placeholderTypes": ["title", "body"],
        "isCustom": false
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
| 3000 | `OFFICE_API_ERROR` - Office API 调用错误 |

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
  fillColor?: string;        // 文本框填充色（十六进制）
  font?: PptFont;            // 插入文本的字体格式，见 data-structures.md#pptfont
}
```

!!! note "插入即带字体格式"
    `options.font`（`PptFont`）在插入文本框的同时应用字体格式，语义等价「插入 + 格式」的一次性复合操作。各属性的 requirement set 与**反应式 3016 全或无**降级同 [`PptFont`](data-structures.md#pptfont)。

**请求参数说明**:

| 参数 | 类型 | 必需 | 默认值 | 说明 |
|------|------|------|--------|------|
| `text` | string | ✅ | - | 要插入的文本内容 |
| `slideIndex` | number | ❌ | 当前幻灯片 | 目标幻灯片索引（从 0 开始） |
| `left` | number | ❌ | - | X 坐标（磅），未指定时使用默认位置 |
| `top` | number | ❌ | - | Y 坐标（磅），未指定时使用默认位置 |
| `width` | number | ❌ | 300 | 文本框宽度（磅） |
| `height` | number | ❌ | 100 | 文本框高度（磅） |
| `fillColor` | string | ❌ | - | 文本框填充色（十六进制） |
| `font` | [`PptFont`](data-structures.md#pptfont) | ❌ | - | 插入文本的字体格式 |

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
    "font": { "size": 18, "name": "微软雅黑", "color": "#333333" }
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
    left: number;            // 实际 X 坐标（磅）
    top: number;             // 实际 Y 坐标（磅）
    width: number;           // 实际宽度（磅）
    height: number;          // 实际高度（磅）
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
    "slideIndex": 0,
    "left": 100,
    "top": 200,
    "width": 400,
    "height": 80
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4001 | `MISSING_PARAM` - 缺少 text 参数 |
| 4002 | `INVALID_PARAM` - slideIndex 超出范围，或 `font.underline` 不在枚举内 |
| 3016 | `API_NOT_SUPPORTED` - `font` 所含属性所需 PowerPointApi requirement set（删除线/上下标/大写=1.8）在当前宿主不满足；`details.requiredApiSet` 标注所需版本 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3000 | `OFFICE_API_ERROR` - Office API 调用错误 |

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
  text?: string;             // 形状内文本（仅 text-capable 形状）
  font?: PptFont;            // 插入形状的字体格式（仅 text-capable 形状），见 data-structures.md#pptfont
}
```

!!! note "插入即带字体格式（仅 text-capable 形状）"
    `options.font`（`PptFont`）在插入形状的同时应用字体格式，语义等价「插入 + 格式」的一次性复合操作，与 [`ppt:insert:text`](#pptinserttext) 的 `font` 字段 / 语义一致。

    **仅适用 text-capable 形状**——`TextBox` 及具备文本框的几何形状（`Rectangle` / `RoundedRectangle` / `Circle` / `Oval` / `Triangle` / `Star` / `Arrow`）。对**无文本框**的形状（`Line`）传入 `font`（或 `text`）→ 前置 `4002 INVALID_PARAM`（语义误用，静态拒绝、非静默忽略）。

    当 `text` 与 `font` 并存时，施加顺序为 `text` → `font`——**`font` 作用于插入后的最终文本内容**。未提供 `text` 时，`font` 作用于形状文本框的默认字体（空文本框，后续输入继承）。

    各属性的 requirement set（基础属性 1.4 / 删除线·上下标·大写 1.8）与降级同 [`PptFont`](data-structures.md#pptfont)：宿主 requirement set 不满足时走**反应式** [`3016 API_NOT_SUPPORTED`](error-handling.md#api_not_supported-3016) **整体失败（全或无）**——能力不足（`3016`）与语义误用（`4002`）语义分离。

**形状类型说明**:

> 除 `Line` 外的所有形状均为 text-capable（接受 `text` / `font`）；`Line` 无文本框。

| 类型 | 说明 |
|------|------|
| `Rectangle` | 矩形 |
| `RoundedRectangle` | 圆角矩形 |
| `Circle` | 圆形 |
| `Oval` | 椭圆 |
| `Triangle` | 三角形 |
| `Line` | 线条（无文本框，不接受 `text` / `font`） |
| `Arrow` | 箭头（块状箭头，含文本框，接受 `text` / `font`） |
| `Star` | 星形 |
| `TextBox` | 文本框 |

**请求参数说明**:

| 参数 | 类型 | 必需 | 默认值 | 说明 |
|------|------|------|--------|------|
| `shapeType` | `ShapeType` | ✅ | - | 形状类型（见上「形状类型说明」） |
| `slideIndex` | number | ❌ | 当前幻灯片 | 目标幻灯片索引（从 0 开始） |
| `left` | number | ❌ | 居中 | X 坐标（磅） |
| `top` | number | ❌ | 居中 | Y 坐标（磅） |
| `width` | number | ❌ | 100 | 宽度（磅） |
| `height` | number | ❌ | 100 | 高度（磅） |
| `fillColor` | string | ❌ | `#4472C4` | 填充颜色（十六进制） |
| `borderColor` | string | ❌ | `#2E5090` | 边框颜色（十六进制） |
| `borderWidth` | number | ❌ | 2 | 边框宽度（磅） |
| `text` | string | ❌ | - | 形状内文本（仅 text-capable 形状；用于 `Line` → `4002`） |
| `font` | [`PptFont`](data-structures.md#pptfont) | ❌ | - | 插入形状的字体格式（仅 text-capable 形状；用于 `Line` → `4002`） |

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
    "text": "点击这里",
    "font": { "size": 18, "name": "微软雅黑", "color": "#FFFFFF", "bold": true }
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
    left: number;            // 实际 X 坐标（磅）
    top: number;             // 实际 Y 坐标（磅）
    width: number;           // 实际宽度（磅）
    height: number;          // 实际高度（磅）
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
    "slideIndex": 0,
    "left": 200,
    "top": 150,
    "width": 200,
    "height": 100
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4001 | `MISSING_PARAM` - 缺少 shapeType 参数 |
| 4002 | `INVALID_PARAM` - shapeType 不支持、slideIndex 超出范围、`font` / `text` 用于无文本框的形状（如 `Line`），或 `font.underline` 不在枚举内 |
| 3016 | `API_NOT_SUPPORTED` - `font` 所含属性所需 PowerPointApi requirement set（删除线 / 上下标 / 大写 = 1.8）在当前宿主不满足；`details.requiredApiSet` 标注所需版本 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3000 | `OFFICE_API_ERROR` - Office API 调用错误 |

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
    left: number;            // 实际 X 坐标（磅）
    top: number;             // 实际 Y 坐标（磅）
    width: number;           // 实际宽度（磅）
    height: number;          // 实际高度（磅）
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
    "slideIndex": 0,
    "left": 300,
    "top": 200,
    "width": 400,
    "height": 300
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
| 3000 | `OFFICE_API_ERROR` - Office API 调用错误 |

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
    left: number;            // 实际 X 坐标（磅）
    top: number;             // 实际 Y 坐标（磅）
    width: number;           // 实际宽度（磅）
    height: number;          // 实际高度（磅）
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
    "columnCount": 4,
    "left": 100,
    "top": 150,
    "width": 520,
    "height": 200
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
| 3000 | `OFFICE_API_ERROR` - Office API 调用错误 |

---

### ppt:insert:chart

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 在指定幻灯片插入图表（柱形/折线/饼/散点等），含数据系列与基础展示选项。

!!! note "契约说明（实现中立，适用于 ppt:insert:chart / ppt:get:chart / ppt:update:chart）"
    chart 类事件遵循标准的请求-响应语义。以下为**线缆可观测契约**，任意实现路径都必须满足，与「在哪一端、用何种技术实现」无关：

    - **持久化与可见性**：当响应 `success: true` 时，图表变更已持久化到目标文档，且对后续 `ppt:get:chart` / `ppt:get:slideElements` 可见。
    - **并发顺序**：协议不保证对同一文档的并发 chart 写入之间的相互可见顺序；若顺序有意义，调用方应串行发起。
    - **elementId 不透明**：响应中的 `elementId` 是服务端分配的不透明字符串，调用方原样回传、**不得**解析其内部结构（详见 [data-structures.md 元素标识符不透明性](data-structures.md#element-id-opacity)）。
    - 本协议**不规定**这些事件由哪一端、用何种技术实现——服务端离线处理或客户端图表 API 均可。接口形状、字段语义、错误码及上述契约对各实现路径一致。
    - 相关数据结构（`ChartData` discriminated union / `ChartType` / `CategoricalSeries` / `ScatterSeries`）定义见 [data-structures.md#ChartType](data-structures.md#charttype)。

!!! info "实现提示（非规范 / Informative）"
    chart 类事件可由不同路径实现。以下特征**均不属于线缆契约**，仅供实现者参考——消费方可按文档连接状态在两条路径间路由。

    **路径 A — 服务端离线（OOXML）**（适用于文档关闭、未被宿主独占时）：

    - 截至撰写时，PowerPoint JavaScript API 未暴露图表创建与数据更新接口（参见 [office-js#5463](https://github.com/OfficeDev/office-js/issues/5463)），因此一种可行实现是由服务端使用 OOXML 工具（如 `python-pptx`）离线修改 `.pptx` 后通知 Add-In 重新加载文档。
    - 该路径预期延迟 >1s（取决于文档大小与磁盘 I/O）。
    - 调用前 Add-In 应已 `save()`，否则未保存的本地修改可能被离线写入覆盖。
    - 完成后 Add-In 需重新打开/刷新文档以呈现新图表。
    - 同一文档的并发离线写入应串行化，避免文件写入冲突。

    **路径 B — 客户端 Office.js 整页 round-trip**（适用于文档已在宿主中打开时）：

    - 一种可行实现：客户端用 `Slide.exportAsBase64` 取目标页实时 OOXML → 服务端改图表 → `insertSlidesFromBase64(keepSourceFormatting)` 整页回插 + 删旧页 + `moveTo` 复位。
    - 整页替换会变更该页元素的 native id——故 `elementId` 不应绑定 native id（线缆契约已要求其不透明，见上）。
    - 操作会重置当前选区与滚动位置。
    - 撤销非原子：整页 round-trip 是多步操作，宿主单次 `Ctrl+Z` 回不到操作前状态——建议把 chart 增改视为「AI 显式操作」，由客户端自备撤销/确认。
    - 反复整页写入可能在文档内累积母版/版式副本。
    - requirement set 门槛：仅插入新页需 PowerPointApi **1.2**；「插入到现有页 / get / update」需 **1.8**；两者皆不满足的平台（如 iPad、老永久版 Office）应回退路径 A。该路径不可用时返回 `3016 API_NOT_SUPPORTED`，调用方据此降级。

**请求数据**:

```typescript
interface InsertChartRequest {
  requestId: string;            // 请求 ID (UUID)
  documentUri: string;          // 文档 URI
  timestamp?: number;           // 请求时间戳（毫秒），可选
  chart: ChartData;             // discriminated union: CategoricalChartData | ScatterChartData
  options?: ChartInsertOptions; // 几何与目标位置（可选）
}

interface ChartInsertOptions {
  slideIndex?: number;          // 目标幻灯片索引（从 0 开始），默认当前幻灯片
  left?: number;                // X 坐标（磅），默认居中
  top?: number;                 // Y 坐标（磅），默认居中
  width?: number;               // 宽度（磅），默认 480
  height?: number;              // 高度（磅），默认 320
}
```

!!! info "schema 形状由 chartType 决定"
    `chart.chartType` 是 discriminator：
    
    - `chartType ∈ CategoricalChartType`（9 个分类型）→ `chart` 形状为 `CategoricalChartData`，需 `categories: string[]` + `series: CategoricalSeries[]`（每条 `series.values` 长度必须等于 `categories.length`）
    - `chartType === "Scatter"` → `chart` 形状为 `ScatterChartData`，需 `series: ScatterSeries[]`（每条 series 自带 `points: { x, y }[]`，**不传 `categories`**）
    
    JSON Schema 应使用 `oneOf` + `discriminator: { propertyName: "chartType" }` 表达，让消费方（LLM）一选定 `chartType` 就能看到对应的必填字段。

**请求示例 — 分类型图表（柱形图）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/q1-report.pptx",
  "chart": {
    "chartType": "ColumnClustered",
    "categories": ["Jan", "Feb", "Mar"],
    "series": [
      { "name": "Revenue", "values": [120, 135, 158] },
      { "name": "Cost",    "values": [80,  90,  102] }
    ],
    "title": "Q1 业绩",
    "showLegend": true
  },
  "options": { "slideIndex": 1, "width": 520, "height": 340 }
}
```

**请求示例 — 散点图**:

```json
{
  "requestId": "b2c3d4e5-f6a7-4b6c-9d8e-0f1a2b3c4d5e",
  "documentUri": "file:///Users/john/Documents/ads-analysis.pptx",
  "chart": {
    "chartType": "Scatter",
    "series": [
      {
        "name": "Ads vs Sales",
        "points": [
          { "x": 1000, "y": 50  },
          { "x": 1500, "y": 80  },
          { "x": 3000, "y": 200 },
          { "x": 5000, "y": 350 },
          { "x": 8000, "y": 600 }
        ]
      }
    ],
    "title": "广告投入 vs 销售额"
  },
  "options": { "slideIndex": 2 }
}
```

**响应数据**:

```typescript
interface InsertChartResponse {
  requestId: string;
  success: boolean;
  data?: {
    elementId: string;          // 创建的图表元素 ID
    slideIndex: number;         // 实际插入的幻灯片索引
    chartType: ChartType;       // 实际写入的图表类型
    seriesCount: number;        // 系列数
    left: number;               // 实际 X 坐标（磅）
    top: number;                // 实际 Y 坐标（磅）
    width: number;              // 实际宽度（磅）
    height: number;             // 实际高度（磅）
  };
  error?: ErrorResponse;
  timestamp: number;
}
```

> 详细的 categories / points 内容如需读回，请调用 `ppt:get:chart`。

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4001 | `MISSING_PARAM` - 缺少 `chart.chartType`，或所选 variant 缺必填字段（categorical 缺 `categories`/`series`，scatter 缺 `series`） |
| 4002 | `INVALID_PARAM` - `chartType` 不在枚举内，或字段形状与 `chartType` 所选 variant 不匹配（例如 `chartType: "Scatter"` 却传了 `categories`） |
| 3015 | `INVALID_CHART_DATA` - categorical: `series[].values.length !== categories.length`；scatter: `points` 为空 / 含 NaN / Infinity |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3003 | `DOCUMENT_READ_ONLY` - 目标文档不可写入（只读或被锁定，无法应用修改） |
| 3004 | `OPERATION_FAILED` - 图表写入失败 |
| 3016 | `API_NOT_SUPPORTED` - 目标操作在当前客户端/平台不可用（如所需 PowerPointApi requirement set 不满足）；调用方应降级到另一实现路径或提示用户 |

---

### ppt:get:chart

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 读取指定 `elementId` 图表的当前数据（类型、分类/数据点、系列、标题、展示选项）。

> 该事件的可见性/顺序契约及实现提示见 [`ppt:insert:chart` 顶部说明](#pptinsertchart)。

**请求数据**:

```typescript
interface GetChartRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  elementId: string;            // 图表元素 ID（必需）
  slideIndex?: number;          // 可选，仅作为加速定位提示；不一致以 elementId 为准
}
```

**响应数据**:

```typescript
interface GetChartResponse {
  requestId: string;
  success: boolean;
  data?: {
    elementId: string;
    slideIndex: number;
    chart: ChartData;           // discriminated union；形状由其中的 chartType 字段决定
    left: number;
    top: number;
    width: number;
    height: number;
  };
  error?: ErrorResponse;
  timestamp: number;
}
```

> 调用方在读取 `chart` 时应先看 `chart.chartType`：若属于 `CategoricalChartType`，存在 `categories` 与 `series[].values`；若为 `"Scatter"`，无 `categories`，数据在 `series[].points[]`。

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4001 | `MISSING_PARAM` - 缺少 elementId |
| 3010 | `ELEMENT_NOT_FOUND` - 指定 elementId 不存在或不是图表元素 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3016 | `API_NOT_SUPPORTED` - 读取所需能力（PowerPointApi 1.8）在当前客户端/平台不可用；调用方应降级到另一实现路径或提示用户 |

---

### ppt:update:chart

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 更新已存在图表的数据、类型、标题或展示选项。除 `chartType`（discriminator，必需）外，其它字段均可选 — 只更新提供的字段。

> 该事件的可见性/顺序契约及实现提示见 [`ppt:insert:chart` 顶部说明](#pptinsertchart)。

**请求数据**:

```typescript
interface UpdateChartRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  elementId: string;            // 图表元素 ID（必需）
  chart: ChartUpdate;           // discriminated union
}

type ChartUpdate = CategoricalChartUpdate | ScatterChartUpdate;

interface CategoricalChartUpdate {
  chartType: CategoricalChartType;     // 必需，作为 discriminator
  categories?: string[];               // 整体替换 X 轴标签（提供即覆盖）
  series?: CategoricalSeries[];        // 整体替换数据系列（提供即覆盖）
  title?: string | null;               // null 表示删除标题
  showLegend?: boolean;
  showDataLabels?: boolean;
}

interface ScatterChartUpdate {
  chartType: "Scatter";                // 必需，作为 discriminator
  series?: ScatterSeries[];            // 整体替换数据系列
  title?: string | null;
  showLegend?: boolean;
  showDataLabels?: boolean;
}
```

!!! warning "chartType 必须显式提供"
    AI 调用前应先用 `ppt:get:chart` 读取当前 `chartType` 并照原值回填。这样 schema 层就能选中正确的 variant、限定可更新字段，避免「以为在改 categorical 实际图表是 scatter」的乌龙。

!!! warning "跨 variant 切换需补齐数据"
    - 同 variant 内换 chartType（如 `Pie → ColumnClustered`，皆为 categorical）：可不传 `categories`/`series`，沿用原值
    - 跨 variant 换 chartType（如 `Pie → Scatter`，或 `Scatter → Line`）：原数据形状不再适用，**必须同时传新 variant 的 `series`**（scatter 需 `series[].points`，categorical 需 `series[].values`），否则返回 `3015 INVALID_CHART_DATA`
    - 同一 variant 内同时替换 `categories` + `series` 时，每条 `series.values.length` 必须等于新的 `categories.length`

**请求示例 — 仅改标题**:

```json
{
  "requestId": "c3d4e5f6-a7b8-4c7d-9e8f-1a2b3c4d5e6f",
  "documentUri": "file:///Users/john/Documents/q1-report.pptx",
  "elementId": "chart-001",
  "chart": {
    "chartType": "ColumnClustered",
    "title": "Q1 业绩（修订版）"
  }
}
```

**请求示例 — 跨 variant 切换（Column → Scatter，必须补 points）**:

```json
{
  "requestId": "d4e5f6a7-b8c9-4d8e-9f0a-2b3c4d5e6f7a",
  "documentUri": "file:///Users/john/Documents/q1-report.pptx",
  "elementId": "chart-001",
  "chart": {
    "chartType": "Scatter",
    "series": [
      {
        "name": "Cost vs Revenue",
        "points": [
          { "x": 80, "y": 120 },
          { "x": 90, "y": 135 },
          { "x": 102, "y": 158 }
        ]
      }
    ]
  }
}
```

**响应数据**:

```typescript
interface UpdateChartResponse {
  requestId: string;
  success: boolean;
  data?: {
    elementId: string;
    chartType: ChartType;       // 更新后的 chartType
    updatedFields: string[];    // 实际生效的字段列表（如 ["title","series"]）
  };
  error?: ErrorResponse;
  timestamp: number;
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4001 | `MISSING_PARAM` - 缺少 `elementId` 或 `chart.chartType` |
| 4002 | `INVALID_PARAM` - `chartType` 不在枚举内，或字段形状与所选 variant 不匹配 |
| 3010 | `ELEMENT_NOT_FOUND` - 指定 elementId 不存在或不是图表元素 |
| 3015 | `INVALID_CHART_DATA` - 跨 variant 切换未补齐 `series`；或更新后 categorical 维度不一致 / scatter `points` 含非法值 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3003 | `DOCUMENT_READ_ONLY` - 目标文档不可写入（只读或被锁定，无法应用修改） |
| 3004 | `OPERATION_FAILED` - 图表写入失败 |
| 3016 | `API_NOT_SUPPORTED` - 目标操作在当前客户端/平台不可用（如所需 PowerPointApi requirement set 不满足）；调用方应降级到另一实现路径或提示用户 |

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
  text?: string;                     // 新文本内容（整框）
  fillColor?: string;                // 文本框填充色（十六进制，框级）
  font?: PptFont;                    // 整框级字体格式（铺底），见 data-structures.md#pptfont
  runs?: PptTextRun[];               // run 级局部格式（区间覆盖，后者胜）
  paragraphs?: PptParagraphStyle[];  // 段落级格式（项目符号 bulletFormat）
}
```

!!! note "字段结构：font 子对象 + 区间覆盖"
    - **整框级**用 `font`（`PptFont`）铺底；**run 级**用 `runs[]`（每段 `{start, length, font}`）覆盖指定字符区间；**段落级** bullet 用 `paragraphs[]`（`{start, length, bulletFormat}`）。
    - **应用顺序**：先 `text`（整框内容替换，如提供）→ 再 `font`（整框铺底）→ 再 `runs`（按数组序覆盖重叠区间）→ 最后 `paragraphs`（bullet）。因内容替换在先，`runs` / `paragraphs` 的 `start` / `length` **一律对齐替换后的最终文本**（详见下）。
    - `runs` / `paragraphs` 的 `start` / `length` 口径（UTF-16 code unit、含 `\r` / `\v`、对齐最终文本、越界 → `4002`）见 [字符区间口径](data-structures.md#pptparagraphstyle)。
    - 各字体属性的 requirement set 及**反应式 3016 全或无**降级见 [`PptFont`](data-structures.md#pptfont)。

**请求参数说明**:

| 参数 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `elementId` | string | ✅ | 要更新的元素 ID（可通过 `ppt:get:slideElements` 获取） |
| `updates.text` | string | ❌ | 新文本内容（整框） |
| `updates.fillColor` | string | ❌ | 文本框填充色（十六进制，框级） |
| `updates.font` | [`PptFont`](data-structures.md#pptfont) | ❌ | 整框级字体格式 |
| `updates.runs` | [`PptTextRun[]`](data-structures.md#ppttextrun) | ❌ | run 级局部格式（字符区间寻址，可多段） |
| `updates.paragraphs` | [`PptParagraphStyle[]`](data-structures.md#pptparagraphstyle) | ❌ | 段落级项目符号（字符区间寻址） |

!!! note "支持的元素类型"
    仅支持 `TextBox`、`Placeholder`、`GeometricShape` 类型的元素。
    对不支持文本的元素类型将返回错误。

**请求示例（font 子对象 + run 级局部格式 + 段落 bullet）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/presentation.pptx",
  "elementId": "shape-001",
  "updates": {
    "text": "更新后的标题",
    "font": { "size": 28, "bold": true, "color": "#333333", "underline": "WavyHeavy" },
    "runs": [
      { "start": 0, "length": 4, "font": { "color": "#C00000", "allCaps": true } },
      { "start": 10, "length": 6, "font": { "underline": "Dotted", "subscript": true } }
    ],
    "paragraphs": [
      { "start": 0, "length": 20, "bulletFormat": { "visible": true, "type": "Numbered", "style": "ArabicNumeralPeriod" } }
    ]
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
    left: number;            // 实际 X 坐标（磅）
    top: number;             // 实际 Y 坐标（磅）
    width: number;           // 实际宽度（磅）
    height: number;          // 实际高度（磅）
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
    "elementId": "shape-001",
    "left": 50,
    "top": 30,
    "width": 600,
    "height": 60
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
| 4002 | `INVALID_PARAM` - `runs` / `paragraphs` 的 `start`/`length` 越界（`start < 0` 或 `start + length > text.length`），或 `underline` / `bulletFormat.type` 不在枚举内 |
| 3003 | `OPERATION_FAILED` - 元素未找到或元素类型不支持文本编辑 |
| 3016 | `API_NOT_SUPPORTED` - 本次请求所含字体/项目符号属性所需 PowerPointApi requirement set（删除线/上下标/大写=1.8、bullet type/style=1.10）在当前宿主不满足；`details.requiredApiSet` 标注所需版本，调用方应仅发送受支持的属性或提示用户 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3000 | `OFFICE_API_ERROR` - Office API 调用错误 |

---

### ppt:delete:element

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 删除幻灯片上的指定元素。支持单个或批量删除。

!!! info "亦适用于图表（Chart）元素删除"
    无独立的 `ppt:delete:chart` 事件。删除图表元素请直接传入 chart 对应的 `elementId`（可由 `ppt:get:slideElements` / `ppt:insert:chart` 响应获取）。

**请求数据**:

```typescript
interface DeleteElementRequest {
  requestId: string;         // 请求 ID (UUID)
  documentUri: string;       // 文档 URI
  timestamp?: number;        // 请求时间戳（毫秒），可选
  elementId?: string;        // 要删除的元素 ID（单个删除）
  elementIds?: string[];     // 要删除的元素 ID 列表（批量删除）
  slideIndex?: number;       // 幻灯片索引（从 0 开始），默认当前幻灯片
}
```

**请求参数说明**:

| 参数 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `elementId` | string | ❌ | 要删除的元素 ID，与 `elementIds` 二选一 |
| `elementIds` | string[] | ❌ | 要批量删除的元素 ID 列表，与 `elementId` 二选一 |
| `slideIndex` | number | ❌ | 幻灯片索引（从 0 开始），默认当前幻灯片 |

!!! warning "参数约束"
    `elementId` 和 `elementIds` 必须提供其中一个。如果都提供，`elementIds` 优先。

**请求示例（单个删除）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/presentation.pptx",
  "elementId": "shape-015",
  "slideIndex": 0
}
```

**请求示例（批量删除）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/presentation.pptx",
  "elementIds": ["shape-015", "shape-016", "shape-017"],
  "slideIndex": 0
}
```

**响应数据**:

```typescript
interface DeleteElementResponse {
  requestId: string;
  success: boolean;
  data?: {
    deletedCount: number;    // 成功删除的元素数量
    slideIndex: number;      // 操作的幻灯片索引
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
    "deletedCount": 3,
    "slideIndex": 0
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4001 | `MISSING_PARAM` - 缺少 elementId 或 elementIds |
| 3003 | `OPERATION_FAILED` - 元素未找到 |
| 4002 | `INVALID_PARAM` - slideIndex 超出范围 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3000 | `OFFICE_API_ERROR` - Office API 调用错误 |

---

### ppt:update:image

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 替换幻灯片中现有图片的内容。可选择保持原有尺寸或指定新尺寸。

**请求数据**:

```typescript
interface UpdateImageRequest {
  requestId: string;         // 请求 ID (UUID)
  documentUri: string;       // 文档 URI
  timestamp?: number;        // 请求时间戳（毫秒），可选
  elementId: string;         // 要替换的图片元素 ID
  image: {
    base64: string;          // Base64 编码的新图片数据
  };
  options?: {
    keepDimensions?: boolean; // 是否保持原有尺寸，默认 true
    width?: number;          // 新宽度（磅），仅 keepDimensions=false 时生效
    height?: number;         // 新高度（磅），仅 keepDimensions=false 时生效
  };
}
```

**请求参数说明**:

| 参数 | 类型 | 必需 | 默认值 | 说明 |
|------|------|------|--------|------|
| `elementId` | string | ✅ | - | 要替换的图片元素 ID |
| `image.base64` | string | ✅ | - | Base64 编码的新图片数据 |
| `keepDimensions` | boolean | ❌ | true | 是否保持原有尺寸 |
| `width` | number | ❌ | - | 新宽度（磅），仅 keepDimensions=false 时生效 |
| `height` | number | ❌ | - | 新高度（磅），仅 keepDimensions=false 时生效 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/presentation.pptx",
  "elementId": "shape-025",
  "image": {
    "base64": "iVBORw0KGgoAAAANSUhEUgAA..."
  },
  "options": {
    "keepDimensions": true
  }
}
```

**响应数据**:

```typescript
interface UpdateImageResponse {
  requestId: string;
  success: boolean;
  data?: {
    elementId: string;       // 更新的图片元素 ID
    slideIndex: number;      // 所在幻灯片索引
    left: number;            // 实际 X 坐标（磅）
    top: number;             // 实际 Y 坐标（磅）
    width: number;           // 实际宽度（磅）
    height: number;          // 实际高度（磅）
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
    "elementId": "shape-025",
    "slideIndex": 0,
    "left": 300,
    "top": 200,
    "width": 400,
    "height": 300
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4001 | `MISSING_PARAM` - 缺少 elementId 或 image.base64 |
| 3003 | `OPERATION_FAILED` - 元素未找到或不是图片类型 |
| 4002 | `INVALID_PARAM` - Base64 数据无效 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3000 | `OFFICE_API_ERROR` - Office API 调用错误 |

---

### ppt:update:tableCell

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 更新表格中指定单元格的文本内容，支持批量更新多个单元格。

**请求数据**:

```typescript
interface UpdateTableCellRequest {
  requestId: string;         // 请求 ID (UUID)
  documentUri: string;       // 文档 URI
  timestamp?: number;        // 请求时间戳（毫秒），可选
  elementId: string;         // 表格元素 ID
  cells: Array<{
    rowIndex: number;        // 行索引（从 0 开始）
    columnIndex: number;     // 列索引（从 0 开始）
    text: string;            // 新文本内容
  }>;
}
```

**请求参数说明**:

| 参数 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `elementId` | string | ✅ | 表格元素 ID（可通过 `ppt:get:slideElements` 获取） |
| `cells` | Array | ✅ | 要更新的单元格列表 |
| `cells[].rowIndex` | number | ✅ | 行索引（从 0 开始） |
| `cells[].columnIndex` | number | ✅ | 列索引（从 0 开始） |
| `cells[].text` | string | ✅ | 新文本内容 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/presentation.pptx",
  "elementId": "shape-030",
  "cells": [
    { "rowIndex": 0, "columnIndex": 0, "text": "姓名" },
    { "rowIndex": 0, "columnIndex": 1, "text": "年龄" },
    { "rowIndex": 1, "columnIndex": 0, "text": "张三" },
    { "rowIndex": 1, "columnIndex": 1, "text": "28" }
  ]
}
```

**响应数据**:

```typescript
interface UpdateTableCellResponse {
  requestId: string;
  success: boolean;
  data?: {
    elementId: string;       // 表格元素 ID
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
    "elementId": "shape-030",
    "cellsUpdated": 4,
    "rowCount": 3,
    "columnCount": 4
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4001 | `MISSING_PARAM` - 缺少 elementId 或 cells |
| 3003 | `OPERATION_FAILED` - 元素未找到或不是表格类型 |
| 4002 | `INVALID_PARAM` - rowIndex 或 columnIndex 超出范围 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3000 | `OFFICE_API_ERROR` - Office API 调用错误 |

---

### ppt:update:tableRowColumn

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 按行或按列批量更新表格内容。适用于一次性填充整行或整列数据。

**请求数据**:

```typescript
interface UpdateTableRowColumnRequest {
  requestId: string;         // 请求 ID (UUID)
  documentUri: string;       // 文档 URI
  timestamp?: number;        // 请求时间戳（毫秒），可选
  elementId: string;         // 表格元素 ID
  rows?: Array<{
    rowIndex: number;        // 行索引（从 0 开始）
    values: string[];        // 该行各列的值，按列顺序排列
  }>;
  columns?: Array<{
    columnIndex: number;     // 列索引（从 0 开始）
    values: string[];        // 该列各行的值，按行顺序排列
  }>;
}
```

**请求参数说明**:

| 参数 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `elementId` | string | ✅ | 表格元素 ID |
| `rows` | Array | ❌ | 按行更新，与 `columns` 至少提供一个 |
| `rows[].rowIndex` | number | ✅ | 行索引（从 0 开始） |
| `rows[].values` | string[] | ✅ | 该行各列的值 |
| `columns` | Array | ❌ | 按列更新，与 `rows` 至少提供一个 |
| `columns[].columnIndex` | number | ✅ | 列索引（从 0 开始） |
| `columns[].values` | string[] | ✅ | 该列各行的值 |

!!! note "行列同时提供"
    当 `rows` 和 `columns` 同时提供时，先处理 `rows` 再处理 `columns`，后者可能覆盖前者对相同单元格的修改。

**请求示例（按行更新）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/presentation.pptx",
  "elementId": "shape-030",
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
    elementId: string;       // 表格元素 ID
    cellsUpdated: number;    // 成功更新的单元格总数
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
    "elementId": "shape-030",
    "cellsUpdated": 8,
    "rowCount": 3,
    "columnCount": 4
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4001 | `MISSING_PARAM` - 缺少 elementId，或 rows 和 columns 均未提供 |
| 3003 | `OPERATION_FAILED` - 元素未找到或不是表格类型 |
| 4002 | `INVALID_PARAM` - rowIndex/columnIndex 超出范围，或 values 长度不匹配 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3000 | `OFFICE_API_ERROR` - Office API 调用错误 |

---

## 样式格式类

### ppt:update:tableFormat

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 更新表格的样式格式，支持按单元格、按行、按列设置格式。

**请求数据**:

```typescript
interface UpdateTableFormatRequest {
  requestId: string;         // 请求 ID (UUID)
  documentUri: string;       // 文档 URI
  timestamp?: number;        // 请求时间戳（毫秒），可选
  elementId: string;         // 表格元素 ID
  cellFormats?: Array<{
    rowIndex: number;                                    // 行索引（从 0 开始）
    columnIndex: number;                                 // 列索引（从 0 开始）
    backgroundColor?: string;                            // 单元格填充色（十六进制）
    font?: PptFont;                                      // 单元格字体格式（整格级，见 data-structures.md#pptfont）
    horizontalAlignment?: ParagraphHorizontalAlignment;  // 水平对齐（7 值枚举）
    verticalAlignment?: TextVerticalAlignment;           // 垂直对齐（6 值枚举）
  }>;
  rowFormats?: Array<{
    rowIndex: number;          // 行索引（从 0 开始）
    height?: number;           // 行高（磅），行级 native
    backgroundColor?: string;  // 背景色（逐格施加，见下）
    font?: PptFont;            // 字体格式（逐格施加，见下）
  }>;
  columnFormats?: Array<{
    columnIndex: number;       // 列索引（从 0 开始）
    width?: number;            // 列宽（磅），列级 native
    backgroundColor?: string;  // 背景色（逐格施加，见下）
    font?: PptFont;            // 字体格式（逐格施加，见下）
  }>;
}
```

!!! note "版本门槛与降级（全或无，统一 1.9）"
    本事件的字体 / 对齐 / 行高列宽 / 单元格填充均经 office.js `TableCell` / `TableRow` / `TableColumn`，其中 `TableCell.font` 访问器门槛为 **PowerPointApi 1.9**，会把复用的 [`PptFont`](data-structures.md#pptfont) 底层 1.4/1.8 属性**整体抬平到 1.9**。故本事件**有效门槛统一为 1.9**——不套用文本框的 1.4/1.8 分档，降级判定为单一 `isSetSupported("PowerPointApi","1.9")`：不满足 → 反应式 [`3016 API_NOT_SUPPORTED`](error-handling.md#api_not_supported-3016)（`details.requiredApiSet: "1.9"`），全或无、不做部分应用。
    单元格**无 run 级**：office.js 未暴露单元格内字符区间寻址，`font` 作用于**整格**，不支持 `runs`。

!!! warning "目标格前置校验（全或无，含合并单元格）"
    应用前须前置校验本次涉及的所有单元格坐标（含由 `rowFormats` / `columnFormats` 展开到的单元格）：任一坐标越界、或命中**合并单元格的非左上格**（office.js `getCellOrNullObject` 返回空对象）→ 整请求按 [`4002 INVALID_PARAM`](error-handling.md) 失败，**不写入任何单元格**。不采用"跳过非法格"的部分成功语义，与文本框口径一致。

**请求参数说明**:

| 参数 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `elementId` | string | ✅ | 表格元素 ID |
| `cellFormats` | Array | ❌ | 按单元格设置格式 |
| `rowFormats` | Array | ❌ | 按行设置格式（应用到整行所有单元格） |
| `columnFormats` | Array | ❌ | 按列设置格式（应用到整列所有单元格） |

!!! note "行/列级即逐格扇出 + 优先级"
    office.js 无"整行 / 整列字体"原生 API：`rowFormats[].font` / `columnFormats[].font` / `backgroundColor` 均为**逐格施加**到该行 / 列每个单元格的语法糖（唯有 `height` / `width` 是行 / 列 native 属性）。因此行 / 列级字体与单元格级字体走同一 `PptFont` 实体、同一 1.9 门槛。
    当多级格式命中同一单元格时优先级：**`cellFormats` > `columnFormats` > `rowFormats`**（细粒度覆盖粗粒度），与 `backgroundColor` 现有口径一致。

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/presentation.pptx",
  "elementId": "shape-030",
  "rowFormats": [
    { "rowIndex": 0, "backgroundColor": "#4472C4", "font": { "size": 14, "bold": true, "color": "#FFFFFF" } }
  ],
  "cellFormats": [
    { "rowIndex": 1, "columnIndex": 0, "font": { "color": "#333333", "underline": "Single" }, "horizontalAlignment": "Center" }
  ]
}
```

**响应数据**:

```typescript
interface UpdateTableFormatResponse {
  requestId: string;
  success: boolean;
  data?: {
    elementId: string;       // 表格元素 ID
    cellsFormatted: number;  // 受影响的单元格数量
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
    "cellsFormatted": 5
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4001 | `MISSING_PARAM` - 缺少 elementId |
| 3003 | `OPERATION_FAILED` - 元素未找到或不是表格类型 |
| 4002 | `INVALID_PARAM` - `rowIndex`/`columnIndex` 超出范围、命中合并单元格非左上格（空对象），或 `font.underline` / 对齐值不在枚举内 |
| 3016 | `API_NOT_SUPPORTED` - 表格单元格字体 / 对齐 / 行高列宽需 PowerPointApi **1.9**，当前宿主不满足；`details.requiredApiSet` 标注所需版本 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3000 | `OFFICE_API_ERROR` - Office API 调用错误 |

---

## 布局操作类

### ppt:update:element

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 更新元素的位置、尺寸或旋转角度。这是布局优化的核心操作，用于移动和缩放元素。

!!! note "与 ppt:update:textBox 的关系"
    `ppt:update:textBox` 用于更新文本**内容和样式**（text、`font` 字体格式、`runs` 局部格式、bullet 等），
    `ppt:update:element` 用于更新**几何属性**（left, top, width, height, rotation）。
    两者互补，分别处理不同维度的更新。

**请求数据**:

```typescript
interface UpdateElementRequest {
  requestId: string;         // 请求 ID (UUID)
  documentUri: string;       // 文档 URI
  timestamp?: number;        // 请求时间戳（毫秒），可选
  elementId: string;         // 要更新的元素 ID
  slideIndex?: number;       // 幻灯片索引（从 0 开始），默认当前幻灯片
  updates: {
    left?: number;           // 新 X 坐标（磅）
    top?: number;            // 新 Y 坐标（磅）
    width?: number;          // 新宽度（磅）
    height?: number;         // 新高度（磅）
    rotation?: number;       // 新旋转角度（度，0-360）
  };
}
```

**请求参数说明**:

| 参数 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `elementId` | string | ✅ | 要更新的元素 ID |
| `slideIndex` | number | ❌ | 幻灯片索引（从 0 开始），默认当前幻灯片 |
| `updates.left` | number | ❌ | 新 X 坐标（磅） |
| `updates.top` | number | ❌ | 新 Y 坐标（磅） |
| `updates.width` | number | ❌ | 新宽度（磅） |
| `updates.height` | number | ❌ | 新高度（磅） |
| `updates.rotation` | number | ❌ | 新旋转角度（度，0-360） |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/presentation.pptx",
  "elementId": "shape-015",
  "slideIndex": 0,
  "updates": {
    "left": 200,
    "top": 150,
    "width": 300,
    "height": 200
  }
}
```

**响应数据**:

```typescript
interface UpdateElementResponse {
  requestId: string;
  success: boolean;
  data?: {
    elementId: string;       // 更新的元素 ID
    slideIndex: number;      // 所在幻灯片索引
    left: number;            // 实际 X 坐标（磅）
    top: number;             // 实际 Y 坐标（磅）
    width: number;           // 实际宽度（磅）
    height: number;          // 实际高度（磅）
    rotation: number;        // 实际旋转角度（度）
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
    "slideIndex": 0,
    "left": 200,
    "top": 150,
    "width": 300,
    "height": 200,
    "rotation": 0
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4001 | `MISSING_PARAM` - 缺少 elementId 或 updates |
| 3003 | `OPERATION_FAILED` - 元素未找到 |
| 4002 | `INVALID_PARAM` - slideIndex 超出范围或尺寸值无效 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3000 | `OFFICE_API_ERROR` - Office API 调用错误 |

---

### ppt:reorder:element

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 调整元素的层叠顺序（Z 轴顺序），支持移至最前、最后、上移一层、下移一层。

**请求数据**:

```typescript
interface ReorderElementRequest {
  requestId: string;         // 请求 ID (UUID)
  documentUri: string;       // 文档 URI
  timestamp?: number;        // 请求时间戳（毫秒），可选
  elementId: string;         // 要调整的元素 ID
  slideIndex?: number;       // 幻灯片索引（从 0 开始），默认当前幻灯片
  action: "bringToFront" | "sendToBack" | "bringForward" | "sendBackward";
}
```

**请求参数说明**:

| 参数 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `elementId` | string | ✅ | 要调整的元素 ID |
| `slideIndex` | number | ❌ | 幻灯片索引（从 0 开始），默认当前幻灯片 |
| `action` | string | ✅ | 调整动作 |

**动作说明**:

| 动作 | 说明 |
|------|------|
| `bringToFront` | 移至最前（置于顶层） |
| `sendToBack` | 移至最后（置于底层） |
| `bringForward` | 上移一层 |
| `sendBackward` | 下移一层 |

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/presentation.pptx",
  "elementId": "shape-015",
  "slideIndex": 0,
  "action": "bringToFront"
}
```

**响应数据**:

```typescript
interface ReorderElementResponse {
  requestId: string;
  success: boolean;
  data?: {
    elementId: string;       // 调整的元素 ID
    slideIndex: number;      // 所在幻灯片索引
    zOrder: number;          // 调整后的层叠顺序值
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
    "slideIndex": 0,
    "zOrder": 5
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4001 | `MISSING_PARAM` - 缺少 elementId 或 action |
| 3003 | `OPERATION_FAILED` - 元素未找到 |
| 4002 | `INVALID_PARAM` - action 值无效或 slideIndex 超出范围 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3000 | `OFFICE_API_ERROR` - Office API 调用错误 |

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
| 3000 | `OFFICE_API_ERROR` - Office API 调用错误 |

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
| 3000 | `OFFICE_API_ERROR` - Office API 调用错误 |

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
| 3000 | `OFFICE_API_ERROR` - Office API 调用错误 |

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
| 3000 | `OFFICE_API_ERROR` - Office API 调用错误 |

---

### ppt:get:slideOoxml

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

!!! note "低层传输原语"
    本事件与 `ppt:insert:slidesOoxml` 是**整页 OOXML 搬运原语**，主要供「双路径」实现搬运幻灯片内容（如 [`ppt:insert:chart` 实现提示路径 B](#pptinsertchart)），而非 AI 常规直接调用。

**说明**: 导出指定幻灯片的当前 OOXML（Office Open XML）为单页 `.pptx` 包，base64 编码。导出反映文档**实时状态（含未保存的本地修改）**，供服务端离线解析或改写后再回插。

!!! info "实现提示（非规范 / Informative）"
    客户端 Office.js 路径的一种可行实现（office-editor4ai 已核对）：`slide.exportAsBase64()` 返回**单页** `.pptx` 的 base64（需 PowerPointApi 1.8）。Office.js 操作内存中的实时文档模型（不读磁盘），故导出含未保存修改——建议实现方以一次 spike 落锤确认（改标题不 save → export → 解压看 XML 含改动）。能力不满足时按 [3016 判定](#pptinsertslidesooxml)用 `isSetSupported` 预检主动返回。

**请求数据**:

```typescript
interface GetSlideOoxmlRequest {
  requestId: string;       // 请求 ID (UUID)
  documentUri: string;     // 文档 URI
  timestamp?: number;      // 请求时间戳（毫秒），可选
  slideIndex: number;      // 目标幻灯片索引（从 0 开始）
}
```

**请求示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/presentation.pptx",
  "slideIndex": 2
}
```

**响应数据**:

```typescript
interface GetSlideOoxmlResponse {
  requestId: string;
  success: boolean;
  data?: {
    slideIndex: number;    // 回显目标页索引
    slideId: string;       // 不透明幻灯片标识符（供后续 insert:slidesOoxml 定位/替换）
    base64: string;        // 单页 .pptx 整包，base64（无 data URL 前缀）
  };
  error?: ErrorResponse;
  timestamp: number;
}
```

> `slideId` 为服务端分配的不透明字符串，调用方原样回传、不得解析（见 [标识符不透明性](data-structures.md#element-id-opacity)）。

**响应示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "slideIndex": 2,
    "slideId": "slide-003",
    "base64": "UEsDBBQABgAIAAAAIQ..."
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
| 3004 | `OPERATION_FAILED` - 导出失败 |
| 3016 | `API_NOT_SUPPORTED` - 导出所需能力（如 PowerPointApi 1.8）在当前客户端/平台不可用；调用方应降级到另一实现路径 |

---

### ppt:insert:slidesOoxml

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

!!! note "低层传输原语"
    见 [`ppt:get:slideOoxml`](#pptgetslideooxml) 顶部说明。

**说明**: 将一个 OOXML 页包（≥1 页，base64 编码的 `.pptx`）插入到演示文稿。可选地在插入后**替换旧页并复位**，把「导出 → 改 → 回插 → 删旧 → 复位」做成一次顺序复合的 round-trip（尽力、非原子，详见下方说明）。

**请求数据**:

```typescript
interface InsertSlidesOoxmlRequest {
  requestId: string;          // 请求 ID (UUID)
  documentUri: string;        // 文档 URI
  timestamp?: number;         // 请求时间戳（毫秒），可选
  base64: string;             // 待插入的 .pptx 页包（≥1 页），base64（无 data URL 前缀）
  formatting?: "keepSourceFormatting" | "useDestinationTheme";  // 默认 "keepSourceFormatting"
  targetSlideIndex?: number;  // 插到此页之后（从 0 开始）；缺省 = 文档末尾
  replaceSlideId?: string;    // 可选：插入完成后删除此旧页（slideId 来自 ppt:get:slideOoxml）
  finalSlideIndex?: number;   // 可选：把插入页移动到此索引（从 0 开始）以复位
}
```

**字段说明**:

| 字段 | 类型 | 必需 | 默认值 | 说明 |
|------|------|------|--------|------|
| `base64` | string | ✅ | - | 待插入 .pptx 页包（≥1 页） |
| `formatting` | string | ❌ | `"keepSourceFormatting"` | 沿用源格式或套用目标主题 |
| `targetSlideIndex` | number | ❌ | 末尾 | 插入位置（插到该索引之后） |
| `replaceSlideId` | string | ❌ | - | round-trip：插入后删除的旧页 |
| `finalSlideIndex` | number | ❌ | - | round-trip：插入页复位到的目标索引。任意复位需 PowerPointApi 1.8；省略、或等于被替换页索引（原位替换）时 1.2 即可 |

!!! warning "顺序复合执行（尽力，非原子）"
    当提供 `replaceSlideId` / `finalSlideIndex` 时，本事件按 **[插入 → 删除 `replaceSlideId` → 移动到 `finalSlideIndex`]** 有序复合执行，语义上视为一次逻辑操作，但**不保证原子性**——底层平台可能不支持事务/回滚，实现应尽量合批以最小化可见中间态，但无法保证零中间态。

    - 成功（`success: true`）：各步全部生效，`insertedSlideIndices` 反映最终真实索引。
    - 部分失败（`success: false`）：`error.details` 给出 `{ stage, partiallyApplied, createdSlideId }`——`stage` 标明失败阶段（`resolve` / `insert` / `delete` / `move` / `readback`）；`partiallyApplied` 表示是否已有副作用；`createdSlideId` 标明已插入但未清理的新页，供调用方（服务端）**对账补偿**（如删除残留新页以恢复操作前状态）。调用方不得假设失败即无副作用。

!!! info "实现提示（非规范 / Informative）"
    客户端 Office.js 路径的一种可行实现（office-editor4ai 已核对 API 可行性）：

    - **调用**：`presentation.insertSlidesFromBase64(base64, { formatting, targetSlideId })`（需 PowerPointApi 1.2），再按需删旧页 + `slide.moveTo(finalSlideIndex)`（需 1.8）。整页插入会重置选区/滚动、撤销非原子——详见 [`ppt:insert:chart` 实现提示路径 B](#pptinsertchart)。
    - **formatting 大小写映射**：协议沿用 camelCase（`keepSourceFormatting` / `useDestinationTheme`，与其它 ppt 事件一致），但 Office.js `PowerPoint.InsertSlideFormatting` 枚举字符串值为 PascalCase（`KeepSourceFormatting` / `UseDestinationTheme`），**不接受 camelCase**。由 Add-In 内部映射到枚举成员，无需改协议。
    - **targetSlideIndex 解析**：Add-In 先 `slides.load("items/id,items/index")` 把 `targetSlideIndex` 解析为 `targetSlideId`；缺省（末尾）时解析为最后一页 id（注意 Office.js 原生缺省 `targetSlideId` 是插到**开头**，故缺省语义须由 Add-In 显式落到末尾）。建议回插时用 `sourceSlideIds` 把源 pin 成单页避免歧义。
    - **命名元素回报（`elements[]`）的定位机制**：因 `elementId` 不透明（见 [标识符不透明性](data-structures.md#element-id-opacity)），定位方式属实现细节，规范不绑定。**主路径**——服务端写 `cNvPr/@name`（如 `oasp-chart-<uuid>`），Add-In 插入后按 `shape.name` 查回（`Shape.name` 1.4；几何单位为磅）。`@name` 穿越 `keepSourceFormatting` round-trip 的存活性已由 Add-In spike 实测确认（[office-editor4ai#34](https://github.com/JIAQIA/office-editor4ai/issues/34)；Mac PowerPoint、单次 round-trip：`@name` 存活、几何可读、`masterLeak: 0`）。**防御后备**——若某平台 / 多次连续 round-trip 下 `@name` 被宿主重写，可改用服务端在 OOXML 写 presentation 级 `customXmlParts` 注册表（`oaspId → {slideId, 页内序号}`，Add-In 经 `presentation.customXmlParts` 读回，需 1.7），不依赖 `@name` 存活；切换后**线缆契约（`elements[]`）不变**。（边界：spike 仅覆盖 Mac/单次 round-trip；Web/Windows 与多次 round-trip 的母版累积建议各补一次抽测。）
    - **占位符内图表的识别**：图表可能承载于内容占位符内，此时 `Shape.type` 报告为 `"Placeholder"`（而非 `"Chart"`），图表性需经 `placeholderFormat.containedType === "Chart"` 判别（office-editor4ai#34 实测）。`elements[]` 按不透明 `elementId` 定位**不受影响**；但调用方**不应**用 `type === "Chart"` 来找图表，否则会漏掉占位符内图表（`ppt:get:slideElements` 的 `includeCharts` 过滤同理）。
    - **3016 判定**：建议用 `Office.context.requirements.isSetSupported('PowerPointApi', '1.2'|'1.8')` **预检后主动返回**（同步、零副作用，避免"insert 成功才发现 moveTo 不支持"留下半成品），少数"API 存在但平台行为不支持"的边缘再由 `OfficeExtension.Error` 兜底映射。

**请求示例 — 整页 round-trip（替换并复位）**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///Users/john/Documents/q1-report.pptx",
  "base64": "UEsDBBQABgAIAAAAIQ...",
  "formatting": "keepSourceFormatting",
  "targetSlideIndex": 2,
  "replaceSlideId": "slide-003",
  "finalSlideIndex": 2
}
```

**响应数据**:

```typescript
interface InsertSlidesOoxmlResponse {
  requestId: string;
  success: boolean;
  data?: {
    insertedSlideIndices: number[];   // 插入页最终索引（从 0 开始）
    insertedSlideIds: string[];       // 插入页的不透明标识符
    elements?: Array<{                // 可选：回报命名元素的实际几何（如图表）
      elementId: string;
      type: string;
      left: number;
      top: number;
      width: number;
      height: number;
    }>;
  };
  error?: ErrorResponse;
  timestamp: number;
}
```

> `insertedSlideIds` 与 `elements[].elementId` 均为不透明标识符（见 [标识符不透明性](data-structures.md#element-id-opacity)）。`elements` 用于把服务端在 OOXML 中命名的元素（如图表）的最终几何回报给调用方，无需额外往返。

**响应示例**:

```json
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": {
    "insertedSlideIndices": [2],
    "insertedSlideIds": ["slide-009"],
    "elements": [
      { "elementId": "oasp-chart-7f3a", "type": "Chart", "left": 60, "top": 120, "width": 480, "height": 320 }
    ]
  },
  "timestamp": 1704067200500
}
```

**可能的错误**:

| 错误码 | 说明 |
|--------|------|
| 4001 | `MISSING_PARAM` - 缺少 base64 |
| 4002 | `INVALID_PARAM` - base64 非法，或 targetSlideIndex / finalSlideIndex 超出范围 |
| 3010 | `ELEMENT_NOT_FOUND` - replaceSlideId 指定的幻灯片不存在 |
| 3001 | `DOCUMENT_NOT_FOUND` - 文档未找到 |
| 3003 | `DOCUMENT_READ_ONLY` - 目标文档不可写入（只读或被锁定，无法应用修改） |
| 3004 | `OPERATION_FAILED` - 插入/替换/复位失败 |
| 3016 | `API_NOT_SUPPORTED` - 所需能力（插入需 PowerPointApi 1.2、finalSlideIndex 复位需 1.8）在当前客户端/平台不满足；调用方应降级到另一实现路径 |

---

## 脚本执行类

### ppt:run:script

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: **封装层逃生舱**——对实时演示文稿执行一段原始 Office.js 脚本。**应优先使用 typed `ppt:*` 事件**，仅在其未覆盖某能力、或某 typed 事件有 Bug 阻塞时使用。共享执行语义、大小限制、超时、安全模型、非原子性与错误映射见[通用约定 · 脚本执行](conventions.md#run-script)。

**请求数据**（共享 [`RunScriptRequest`](data-structures.md#script-execution)）:

```typescript
interface RunScriptRequest {
  requestId: string;
  documentUri: string;
  timestamp?: number;
  script: string;                    // async 函数体；注入 context(PowerPoint.RequestContext)/args/console，可 return
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
  "documentUri": "file:///Users/john/Documents/deck.pptx",
  "script": "const slides = context.presentation.slides;\nslides.load('items');\nawait context.sync();\nreturn slides.items.length;"
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
    "result": 12,
    "logs": [],
    "durationMs": 38,
    "logsTruncated": false
  },
  "timestamp": 1704067200500,
  "duration": 40
}
```

!!! info "PPT 宿主能力自检（实现提示 / 非规范）"
    `PowerPointApi` requirement set 起步较晚，脚本调用当前宿主**不存在**的 API 时，代理对象上的方法为 `undefined` → 抛 `TypeError`，会被错误映射记为 `fault:"script"`（而真因是宿主能力不足）。脚本**建议**先用全局 `Office.context.requirements.isSetSupported("PowerPointApi", "1.x")` 自检目标能力再调用，以获得更准确的失败归因。

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
