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
  description?: string;    // 样式描述（仅当请求 detailedInfo=true 时返回）
}
```

!!! note "关于 description 字段"
    `description` 字段仅在 `word:get:styles` 请求中设置 `detailedInfo=true` 时返回。
    此功能依赖 WordApi BETA，在部分环境中可能不可用。

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

### DocumentStructureResult

文档结构统计。

```typescript
interface DocumentStructureResult {
  sectionCount: number;    // 章节数量
  paragraphCount: number;  // 段落数量
  tableCount: number;      // 表格数量
  imageCount: number;      // 图片数量
  tables?: TableSummary[]; // 表格清单（可选，用于"重新发现"现有表格）
}

interface TableSummary {
  tableId: string;            // 临时索引（详见 word:insert:table 稳定性说明）
  rowCount: number;           // 表格行数（未合并状态）
  columnCount: number;        // 表格列数（未合并状态）
  precedingHeading?: string;  // 表格前最近的标题文本，便于 AI 通过启发式定位
}
```

!!! note "tables 字段为可选"
    旧 Add-In 版本可能不返回 `tables`；调用方应做存在性判断。如缺省，可退化为按 `tableCount` 配合 `tableId = "table-{i}"` 推断（仅在文档结构未变更时可靠）。

### DocumentStatsResult

文档字数统计。

```typescript
interface DocumentStatsResult {
  characterCount: number;           // 字符数（不含空格）
  characterCountWithSpaces: number; // 字符数（含空格）
  wordCount: number;                // 单词数
  paragraphCount: number;           // 段落数
  pageCount?: number;               // 页数（可选）
}
```

---

## 替换内容

### ReplaceContent

替换操作的内容定义。

```typescript
interface ReplaceContent {
  text?: string;           // 替换文本
  images?: ImageData[];    // 替换图片（可插入多张）
  format?: TextFormat;     // 文本格式（最高优先级）
  styleName?: string;      // Word 样式名（仅在 format 未提供时使用）
}
```

!!! important "格式优先级"
    - `format`（最高优先级）：包含直接格式属性和 `format.styleName`
    - `styleName`（仅在 `format` 未提供时使用）
    - 默认保持选区原有格式

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
  insertLocation?: TableInsertLocation;  // 插入位置，默认 "End"
}

type TableInsertLocation =
  | "Start"    // 文档开头
  | "End"      // 文档末尾（默认）
  | "Before"   // 当前选区/光标之前
  | "After"    // 当前选区/光标之后
  | "Replace"; // 替换当前选区
```

### CellFormat

表格单元格格式属性。可由 `word:update:tableCell` 等事件复用，承载单元格的对齐、字体与底色等可视样式。所有字段均为可选——未传字段保持原状，便于增量更新。

```typescript
interface CellFormat {
  horizontalAlignment?: "Left" | "Centered" | "Right" | "Justified";  // 对应 Word.Alignment
  verticalAlignment?: "Top" | "Center" | "Bottom";                    // 对应 Word.VerticalAlignment
  backgroundColor?: string;       // 背景颜色（十六进制，如 "#4472C4"）— 对应 cell.shadingColor
  fontName?: string;              // 字体名称（应用到整个单元格 body，会覆盖原段落字体）
  fontSize?: number;              // 字号（磅）
  fontColor?: string;             // 字体颜色（十六进制，如 "#333333"）
  bold?: boolean;                 // 粗体
  italic?: boolean;               // 斜体
}
```

!!! note "命名对齐 Office.js"
    枚举值刻意与 Word JavaScript API 的 `Word.Alignment` / `Word.VerticalAlignment` 完全一致（注意 `Centered` / `Justified` 是过去分词形），便于 Add-In 直接 `cast` 不做映射。

!!! note "整刷语义"
    `fontName` / `fontSize` / `fontColor` / `bold` / `italic` 作用于整个单元格 body，会覆盖单元格内所有段落与 Run 的字体设置。这是 Word.js 的固有行为，不是协议保留的灵活度。

!!! note "单元格内边距"
    Word JavaScript API 的单元格内边距是**表级 API**（`Word.Table.setCellPadding`），无法逐单元格设置。如需调整内边距，请通过 `word:update:tableFormat.styleOptions.cellPadding` 在表级配置。

> 该结构当前仅 `/word` 命名空间引用；`/ppt` 已有的 `ppt:update:tableFormat` 后续如需统一字段命名，可平滑迁移到本结构。

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

### 标识符不透明性 {#element-id-opacity}

幻灯片元素与幻灯片标识符——`SlideElement.id`、各 chart 事件的 `elementId`、以及 `ppt:get:slideOoxml` / `ppt:insert:slidesOoxml` 的 `slideId`——都是**服务端分配的不透明字符串**。消费方**不得**解析、推断或依赖其内部结构与格式，只能将其作为整体令牌原样回传。

这是规范层约束，独立于实现：实现可自由更换内部 id 生成方式（例如整页 round-trip 后改用稳定 UUID 并写入 OOXML `cNvPr/@name`），只要同一对象在其生命周期内返回稳定、可回传的标识符即可——协议不绑定任何具体格式。

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

---

## 图表相关

> 跨命名空间通用图表数据结构。供 PPT (`ppt:insert:chart` / `ppt:get:chart` / `ppt:update:chart`) 与未来 Excel (`excel:insert:chart` 等) 共用。

### ChartType

图表类型总枚举，命名与 [`Excel.ChartType`](https://learn.microsoft.com/en-us/javascript/api/excel/excel.charttype) 保持一致，便于消费方直接 cast。`ChartType` 在数据形状上分两类，`ChartData` 据此采用 discriminated union 表达。

```typescript
// 分类型图表：X 轴是离散标签，Y 是数值
type CategoricalChartType =
  | "ColumnClustered"      // 簇状柱形图
  | "ColumnStacked"        // 堆积柱形图
  | "BarClustered"         // 簇状条形图
  | "Line"                 // 折线图
  | "LineMarkers"          // 带标记折线图
  | "Pie"                  // 饼图
  | "Doughnut"             // 圆环图
  | "Area"                 // 面积图
  | "Radar";               // 雷达图

// 散点型图表：每个数据点自带 (x, y) 数对
type ScatterChartType = "Scatter";

type ChartType = CategoricalChartType | ScatterChartType;
```

### CategoricalSeries

分类型图表的单条数据系列（柱形/折线/饼图等使用）。

```typescript
interface CategoricalSeries {
  name: string;            // 系列名称（图例显示）
  values: number[];        // Y 数值数组，长度需 = CategoricalChartData.categories.length
  color?: string;          // 系列颜色（hex，如 "#4472C4"），缺省使用主题色
}
```

### ScatterSeries

散点型图表的单条数据系列。

```typescript
interface ScatterSeries {
  name: string;            // 系列名称
  points: ScatterPoint[];  // 数据点数组（≥1）
  color?: string;          // 系列颜色（hex），缺省使用主题色
}

interface ScatterPoint {
  x: number;               // X 坐标（连续数值轴）
  y: number;               // Y 坐标
}
```

### ChartData

图表的逻辑数据与展示选项（不含几何位置）。`ChartData` 是 discriminated union，由 `chartType` 字段决定具体形状：

```typescript
type ChartData = CategoricalChartData | ScatterChartData;

interface CategoricalChartData {
  chartType: CategoricalChartType;     // 9 个分类型枚举之一（不含 "Scatter"）
  categories: string[];                // X 轴离散标签（如 ["Jan","Feb","Mar"]）
  series: CategoricalSeries[];         // 数据系列（≥1）
  title?: string;                      // 图表标题
  showLegend?: boolean;                // 是否显示图例，默认 true
  showDataLabels?: boolean;            // 是否显示数据标签，默认 false
}

interface ScatterChartData {
  chartType: "Scatter";                // 字面量
  series: ScatterSeries[];             // 数据系列（≥1）；X 由 series[].points[].x 提供，无 categories
  title?: string;
  showLegend?: boolean;
  showDataLabels?: boolean;
}
```

!!! info "为什么用 discriminated union"
    `ChartType` 在数据形状上不一致：分类型图表用 `categories + series.values`（X 是离散标签），散点图用 `series.points`（X 是连续数值，每个点自带 (x,y)）。把它们硬塞进同一个 schema（如让 `categories` 在 Scatter 时存数字字符串）会让 LLM 在 MCP 工具 schema 里看不到这个隐式约束。Discriminated union 让 LLM 一选定 `chartType`，schema 就自动限定可填字段。

    JSON Schema 落地建议使用 `oneOf` + `discriminator: { propertyName: "chartType" }`；Pydantic / FastMCP 使用 `Annotated[Union[...], Field(discriminator="chartType")]`。

!!! warning "数据维度校验"
    - `CategoricalChartData`：每个 `CategoricalSeries.values` 长度必须等于 `categories.length`，否则返回 `3015 INVALID_CHART_DATA`
    - `ScatterChartData`：每个 `ScatterSeries.points` 长度 ≥ 1，`x`/`y` 必须为有限数（非 `NaN` / `Infinity`），否则返回 `3015 INVALID_CHART_DATA`

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
