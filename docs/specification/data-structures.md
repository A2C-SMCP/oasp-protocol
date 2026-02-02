# 数据结构

## 概述

本章定义 OASP 协议中使用的所有通用数据结构。这些数据结构被多个事件共享。

## 基础请求/响应结构

### BaseRequest

所有请求的基础结构。

```typescript
interface BaseRequest {
  requestId: string;       // UUID v4 格式的请求 ID
  documentUri: string;     // 文档 URI，格式为 file:///path/to/document
  timestamp: number;       // 请求发起时间（Unix 毫秒时间戳）
}
```

### BaseResponse

所有成功响应的基础结构。

```typescript
interface BaseResponse {
  requestId: string;       // 对应请求的 ID
  success: true;
  data: object;            // 具体的返回数据
  timestamp: number;       // 响应时间（Unix 毫秒时间戳）
  duration?: number;       // 操作耗时（毫秒）
}
```

### ErrorResponse

错误响应结构。

```typescript
interface ErrorResponse {
  requestId: string;       // 对应请求的 ID
  success: false;
  error: {
    code: string;          // 错误码，见错误处理章节
    message: string;       // 错误消息
    details?: object;      // 附加详情（可选）
  };
  timestamp: number;
  duration?: number;
}
```

---

## 选区相关

### SelectionInfo

选区位置信息。

```typescript
interface SelectionInfo {
  isEmpty: boolean;        // 选区是否为空
  type: SelectionType;     // 选区类型
  start?: number;          // 开始位置（字符偏移）
  end?: number;            // 结束位置（字符偏移）
  text?: string;           // 选中的文本内容
}
```

### SelectionType

选区类型枚举。

```typescript
type SelectionType =
  | "NoSelection"          // 无选区
  | "InsertionPoint"       // 光标（无选中内容）
  | "Normal";              // 正常选区（有选中内容）
```

---

## 文本格式

### TextFormat

文本格式定义。

```typescript
interface TextFormat {
  bold?: boolean;          // 粗体
  italic?: boolean;        // 斜体
  underline?: UnderlineStyle;  // 下划线样式
  fontSize?: number;       // 字号（磅）
  fontName?: string;       // 字体名称
  color?: string;          // 文字颜色（十六进制，如 "#FF0000"）
  highlightColor?: string; // 高亮颜色
  styleName?: string;      // Word 样式名称
}
```

!!! important "样式优先级"
    当 `styleName` 与直接格式属性同时存在时，**直接格式属性优先级更高**。

    处理顺序：
    1. 先应用 `styleName` 指定的样式
    2. 再用直接格式属性覆盖

### UnderlineStyle

下划线样式枚举。

```typescript
type UnderlineStyle =
  | "none"                 // 无下划线
  | "single"               // 单下划线
  | "double"               // 双下划线
  | "dotted"               // 点线
  | "dashed"               // 虚线
  | "thick"                // 粗下划线
  | "wave";                // 波浪线
```

---

## 样式相关

### StyleInfo

文档样式信息。

```typescript
interface StyleInfo {
  name: string;            // 样式名称（本地化名称）
  type: StyleType;         // 样式类型
  builtIn: boolean;        // 是否为内置样式
  inUse: boolean;          // 是否在文档中使用
  description?: string;    // 样式描述（可选）
}
```

### StyleType

样式类型枚举。

```typescript
type StyleType =
  | "Paragraph"            // 段落样式
  | "Character"            // 字符样式
  | "Table"                // 表格样式
  | "List";                // 列表样式
```

---

## 文档统计

### DocumentStructure

文档结构统计。

```typescript
interface DocumentStructure {
  paragraphCount: number;  // 段落数量
  tableCount: number;      // 表格数量
  imageCount: number;      // 图片数量
  sectionCount: number;    // 章节数量
}
```

### DocumentStats

文档字数统计。

```typescript
interface DocumentStats {
  wordCount: number;       // 字数
  characterCount: number;  // 字符数（包括空格和标点）
  paragraphCount: number;  // 段落数（包括空段落）
}
```

---

## 替换内容

### ReplaceContent

替换操作的内容定义。

```typescript
interface ReplaceContent {
  text?: string;           // 文本内容
  images?: ImageData[];    // 图片内容（替换为图片）
  format?: TextFormat;     // 格式设置（仅对文本有效）
}
```

---

## 图片相关

### ImageData

图片数据定义。

```typescript
interface ImageData {
  base64: string;          // Base64 编码的图片数据
  mimeType?: string;       // MIME 类型，如 "image/png"
  width?: number;          // 宽度（像素或点）
  height?: number;         // 高度（像素或点）
  altText?: string;        // 替代文本
}
```

---

## 表格相关

### TableInsertOptions

表格插入选项。

```typescript
interface TableInsertOptions {
  rows: number;            // 行数（>= 1）
  columns: number;         // 列数（>= 1）
  data?: string[][];       // 初始数据（二维数组）
  style?: string;          // 表格样式名称
}
```

---

## PPT 相关

### SlideElement

幻灯片元素信息。

```typescript
interface SlideElement {
  id: string;              // 元素 ID
  type: SlideElementType;  // 元素类型
  position: {
    left: number;          // 左边距（点）
    top: number;           // 上边距（点）
    width: number;         // 宽度（点）
    height: number;        // 高度（点）
  };
  text?: string;           // 文本内容（如适用）
  zIndex: number;          // 层级
}
```

### SlideElementType

幻灯片元素类型。

```typescript
type SlideElementType =
  | "TextBox"              // 文本框
  | "Shape"                // 形状
  | "Image"                // 图片
  | "Table"                // 表格
  | "Chart"                // 图表
  | "SmartArt"             // SmartArt 图形
  | "Video"                // 视频
  | "Audio";               // 音频
```

### ShapeType

形状类型（常用）。

```typescript
type ShapeType =
  | "Rectangle"            // 矩形
  | "RoundedRectangle"     // 圆角矩形
  | "Circle"               // 圆形
  | "Oval"                 // 椭圆
  | "Triangle"             // 三角形
  | "Diamond"              // 菱形
  | "Pentagon"             // 五边形
  | "Hexagon"              // 六边形
  | "Line"                 // 直线
  | "Arrow"                // 箭头
  | "Star"                 // 星形
  | "TextBox";             // 文本框
```

---

## Excel 相关

### RangeInfo

Excel 范围信息。

```typescript
interface RangeInfo {
  address: string;         // 范围地址，如 "Sheet1!A1:C3"
  rowCount: number;        // 行数
  columnCount: number;     // 列数
  worksheet: string;       // 所属工作表名称
}
```

### CellValueType

单元格值类型。

```typescript
type CellValueType =
  | "String"               // 字符串
  | "Number"               // 数字
  | "Boolean"              // 布尔值
  | "Date"                 // 日期
  | "Error"                // 错误值
  | "Empty";               // 空值
```

### ChartType

图表类型（常用）。

```typescript
type ChartType =
  | "Column"               // 柱形图
  | "Bar"                  // 条形图
  | "Line"                 // 折线图
  | "Pie"                  // 饼图
  | "Area"                 // 面积图
  | "Scatter"              // 散点图
  | "Doughnut";            // 圆环图
```

---

## 通用枚举

### InsertLocation

插入位置枚举。

```typescript
type InsertLocation =
  | "Before"               // 在目标之前
  | "After"                // 在目标之后
  | "Start"                // 在目标开头
  | "End"                  // 在目标末尾
  | "Replace";             // 替换目标
```
