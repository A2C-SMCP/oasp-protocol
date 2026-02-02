# PPT 事件定义

!!! warning "Draft 状态"
    本文档中的所有事件处于 **Draft** 状态，接口可能在正式发布前发生变更。

## 概述

本章定义 `/ppt` 命名空间下的所有事件。PPT 事件用于操作 Microsoft PowerPoint 演示文稿。

## 事件列表

### 内容检索类

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [ppt:get:currentSlideElements](#pptgetcurrentslideelements) | 📋 Draft | 获取当前幻灯片元素 |
| [ppt:get:slideElements](#pptgetslideelements) | 📋 Draft | 获取指定幻灯片元素 |
| [ppt:get:slideScreenshot](#pptgetslidescreenshot) | 📋 Draft | 获取幻灯片截图 |

### 内容操作类

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [ppt:insert:text](#pptinserttext) | 📋 Draft | 插入文本 |
| [ppt:insert:shape](#pptinsertshape) | 📋 Draft | 插入形状 |
| [ppt:insert:image](#pptinsertimage) | 📋 Draft | 插入图片 |
| [ppt:insert:table](#pptinserttable) | 📋 Draft | 插入表格 |
| [ppt:update:textBox](#pptupdatetextbox) | 📋 Draft | 更新文本框 |

### 幻灯片管理类

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [ppt:add:slide](#pptaddslide) | 📋 Draft | 添加幻灯片 |
| [ppt:delete:slide](#pptdeleteslide) | 📋 Draft | 删除幻灯片 |
| [ppt:move:slide](#pptmoveslide) | 📋 Draft | 移动幻灯片 |
| [ppt:goto:slide](#pptgotoslide) | 📋 Draft | 跳转到幻灯片 |

---

## 内容检索类

### ppt:get:currentSlideElements

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 获取当前显示幻灯片上的所有元素信息。

**请求数据**:

```typescript
interface GetCurrentSlideElementsRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  options?: {
    includeHidden?: boolean;  // 是否包含隐藏元素，默认 false
  };
}
```

**响应数据**:

```typescript
interface GetCurrentSlideElementsResponse {
  requestId: string;
  success: true;
  data: {
    slideIndex: number;      // 当前幻灯片索引（从 0 开始）
    slideId: string;         // 幻灯片 ID
    elements: SlideElement[];
  };
  timestamp: number;
  duration: number;
}
```

---

### ppt:get:slideElements

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 获取指定幻灯片上的所有元素信息。

**请求数据**:

```typescript
interface GetSlideElementsRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  slideIndex: number;        // 幻灯片索引（从 0 开始）
}
```

---

### ppt:get:slideScreenshot

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 获取幻灯片的截图。

**请求数据**:

```typescript
interface GetSlideScreenshotRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  slideIndex?: number;       // 幻灯片索引，默认当前幻灯片
  options?: {
    width?: number;          // 输出宽度（像素）
    height?: number;         // 输出高度（像素）
    format?: "png" | "jpeg"; // 图片格式，默认 "png"
  };
}
```

**响应数据**:

```typescript
interface GetSlideScreenshotResponse {
  requestId: string;
  success: true;
  data: {
    slideIndex: number;
    imageBase64: string;     // Base64 编码的图片数据
    mimeType: string;        // MIME 类型
    width: number;
    height: number;
  };
  timestamp: number;
  duration: number;
}
```

---

## 内容操作类

### ppt:insert:text

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 在当前幻灯片插入文本框。

**请求数据**:

```typescript
interface PPTInsertTextRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  text: string;
  position?: {
    left: number;            // 左边距（点）
    top: number;             // 上边距（点）
    width?: number;          // 宽度（点）
    height?: number;         // 高度（点）
  };
  format?: TextFormat;
}
```

---

### ppt:insert:shape

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 在当前幻灯片插入形状。

**请求数据**:

```typescript
interface InsertShapeRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  shapeType: ShapeType;      // 形状类型
  position: {
    left: number;
    top: number;
    width: number;
    height: number;
  };
  options?: {
    fillColor?: string;      // 填充颜色
    lineColor?: string;      // 边框颜色
    text?: string;           // 形状内文本
  };
}
```

---

### ppt:insert:image

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 在当前幻灯片插入图片。

**请求数据**:

```typescript
interface PPTInsertImageRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  image: ImageData;
  position?: {
    left: number;
    top: number;
  };
}
```

---

### ppt:insert:table

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 在幻灯片中插入表格。

**请求数据**:

```typescript
interface PPTInsertTableRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  options: {
    rows: number;            // 行数（>= 1）
    columns: number;         // 列数（>= 1）
    slideIndex?: number;     // 幻灯片索引，默认当前幻灯片
    left?: number;           // 左边距（点）
    top?: number;            // 上边距（点）
    data?: string[][];       // 初始数据（二维数组）
  };
}
```

**响应数据**:

```typescript
interface PPTInsertTableResponse {
  requestId: string;
  success: true;
  data: {
    elementId: string;       // 创建的表格元素 ID
  };
  timestamp: number;
  duration: number;
}
```

---

### ppt:update:textBox

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 更新幻灯片中现有文本框的内容或样式。

**请求数据**:

```typescript
interface PPTUpdateTextBoxRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  elementId: string;         // 要更新的文本框元素 ID
  updates: {
    text?: string;           // 新文本内容
    fontSize?: number;       // 字号
    fontName?: string;       // 字体名称
    color?: string;          // 文字颜色（十六进制）
    fillColor?: string;      // 填充颜色（十六进制）
    bold?: boolean;          // 粗体
    italic?: boolean;        // 斜体
  };
}
```

**响应数据**:

```typescript
interface PPTUpdateTextBoxResponse {
  requestId: string;
  success: true;
  data: {
    elementId: string;       // 更新的元素 ID
  };
  timestamp: number;
  duration: number;
}
```

---

## 幻灯片管理类

### ppt:add:slide

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 添加新幻灯片。

**请求数据**:

```typescript
interface AddSlideRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  options?: {
    insertIndex?: number;    // 插入位置，默认末尾
    layout?: string;         // 版式名称
  };
}
```

---

### ppt:delete:slide

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 删除幻灯片。

**请求数据**:

```typescript
interface DeleteSlideRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  slideIndex: number;        // 要删除的幻灯片索引
}
```

---

### ppt:move:slide

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 移动幻灯片位置。

**请求数据**:

```typescript
interface MoveSlideRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  fromIndex: number;         // 原位置
  toIndex: number;           // 目标位置
}
```

---

### ppt:goto:slide

**方向**: Server → AddIn（请求-响应）

**状态**: 📋 Draft

**说明**: 跳转到指定幻灯片。

**请求数据**:

```typescript
interface GotoSlideRequest {
  requestId: string;
  documentUri: string;
  timestamp: number;
  slideIndex: number;        // 目标幻灯片索引
}
```

**响应数据**:

```typescript
interface GotoSlideResponse {
  requestId: string;
  success: true;
  data: {
    slideIndex: number;      // 当前幻灯片索引
  };
  timestamp: number;
  duration: number;
}
```
