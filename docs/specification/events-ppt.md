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

### 幻灯片管理类

| 事件名 | 状态 | 说明 |
|--------|------|------|
| [ppt:slide:add](#pptslideadd) | 📋 Draft | 添加幻灯片 |
| [ppt:slide:delete](#pptslidedelete) | 📋 Draft | 删除幻灯片 |
| [ppt:slide:move](#pptslidemove) | 📋 Draft | 移动幻灯片 |
| [ppt:slide:goto](#pptslidegoto) | 📋 Draft | 跳转到幻灯片 |

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

## 幻灯片管理类

### ppt:slide:add

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

### ppt:slide:delete

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

### ppt:slide:move

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

### ppt:slide:goto

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
