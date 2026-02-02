# 术语表

## 概述

本术语表解释 OASP 协议文档中使用的 Office 相关术语和技术概念。

---

## 通用术语

### AddIn

Office 加载项，也称为 Office Add-in。一种扩展 Microsoft Office 功能的 Web 应用程序，运行在 Office 应用程序内部的沙盒环境中。

### Office.js

Microsoft 提供的 JavaScript API 库，用于与 Office 应用程序（Word、Excel、PowerPoint 等）进行交互。是开发 Office AddIn 的核心技术。

### Socket.IO

一个实时双向通信库，基于 WebSocket 协议，并提供回退机制（如 HTTP 长轮询）。OASP 使用 Socket.IO 作为传输层。

### 命名空间 (Namespace)

Socket.IO 的概念，用于在单个连接上实现逻辑分离。OASP 使用不同的命名空间区分 Word (`/word`)、PPT (`/ppt`) 和 Excel (`/excel`) 的通信。

---

## Word 相关术语

### Selection（选区）

用户当前选中的文档内容区域。可以是：

- **无选区** (NoSelection): 没有任何选中内容
- **插入点** (InsertionPoint): 光标位置，没有选中文本
- **正常选区** (Normal): 选中了一段文本内容

### Range（范围）

文档中的一段连续区域，由起始位置和结束位置定义。Range 是 Office.js 中操作文档内容的基本单位。

### Paragraph（段落）

以段落标记（回车符）分隔的文档内容块。每个段落可以有独立的格式设置。

### Section（节）

文档中的一个分区，用于控制页面布局设置（如页边距、页眉页脚、纸张方向）。一个文档可以包含多个节。

### Style（样式）

预定义的格式集合，包括字体、字号、颜色、段落间距等。Office 中的样式分为：

- **段落样式** (Paragraph Style): 应用于整个段落
- **字符样式** (Character Style): 应用于选中的文字
- **表格样式** (Table Style): 应用于表格
- **列表样式** (List Style): 应用于列表

### Content Control（内容控件）

Word 中的容器元素，用于存放和管理文档内容。可以限制用户编辑、设置占位符文本等。

### Body（正文）

文档的主体内容区域，不包括页眉、页脚和脚注。

### Header / Footer（页眉/页脚）

出现在每页顶部或底部的内容区域，通常用于显示页码、文档标题等。

---

## PowerPoint 相关术语

### Slide（幻灯片）

PPT 演示文稿的单个页面。每个 Slide 包含各种元素（文本框、图片、形状等）。

### Slide Index（幻灯片索引）

幻灯片在演示文稿中的位置序号，从 0 开始计数。

### Master Slide（母版）

定义幻灯片外观和布局的模板。修改母版会影响所有使用该母版的幻灯片。

### Layout（版式）

预定义的幻灯片布局模板，如"标题幻灯片"、"标题和内容"等。

### Shape（形状）

PPT 中的基本图形元素，包括矩形、圆形、箭头等。形状可以包含文本。

### TextBox（文本框）

用于在幻灯片上放置文本的容器。文本框是一种特殊的形状。

### Placeholder（占位符）

版式中预定义的内容区域，如标题占位符、内容占位符。

### Z-Index（层级）

元素的堆叠顺序。Z-Index 值大的元素显示在上层。

---

## Excel 相关术语

### Workbook（工作簿）

Excel 文件，可以包含多个工作表。

### Worksheet（工作表）

工作簿中的单个表格页面，也称为"Sheet"。每个工作表是一个二维单元格网格。

### Cell（单元格）

工作表中的最小数据单元，由列（A、B、C...）和行（1、2、3...）的交叉点标识，如"A1"、"B2"。

### Range（范围）

一组连续的单元格区域，用起始和结束单元格表示，如"A1:C3"表示从 A1 到 C3 的矩形区域。

### Address（地址）

单元格或范围的位置标识符。格式为 `Sheet!Cell` 或 `Sheet!Range`，如"Sheet1!A1:C3"。

### Used Range（已使用范围）

工作表中包含数据或格式的区域。从第一个非空单元格到最后一个非空单元格的矩形区域。

### Table（表格）

Excel 中的结构化数据区域，具有列标题和自动筛选功能。不同于普通的单元格区域。

### Chart（图表）

数据的可视化呈现，如柱形图、折线图、饼图等。

### Formula（公式）

以 `=` 开头的表达式，用于计算单元格的值。如 `=SUM(A1:A10)`。

---

## 技术术语

### DTO (Data Transfer Object)

数据传输对象，用于在不同系统或层之间传递数据的对象。OASP 中定义了各种请求和响应的 DTO。

### UUID (Universally Unique Identifier)

通用唯一标识符，用于生成全局唯一的 ID。OASP 使用 UUID v4 作为请求 ID。

### ACK (Acknowledgement)

确认，Socket.IO 中的一种机制，允许接收方确认消息已收到并可选地返回数据。

### Callback

回调函数，Socket.IO 中用于处理 ACK 响应的函数。

### Base64

一种二进制数据的文本编码方式。OASP 中用于传输图片等二进制数据。

### MIME Type

媒体类型，用于标识数据格式。如 `image/png`、`application/json`。

### URI (Uniform Resource Identifier)

统一资源标识符，用于唯一标识资源。OASP 中用于标识文档位置。

### Point（磅）

印刷排版单位，1 英寸 = 72 磅。Office 中用于表示字号和元素尺寸。

---

## 缩写对照

| 缩写 | 全称 | 说明 |
|------|------|------|
| OASP | Office AddIn Socket Protocol | 本协议名称 |
| API | Application Programming Interface | 应用程序编程接口 |
| JSON | JavaScript Object Notation | JavaScript 对象表示法 |
| UTF-8 | 8-bit Unicode Transformation Format | 8 位 Unicode 转换格式 |
| HTTP | Hypertext Transfer Protocol | 超文本传输协议 |
| WS | WebSocket | WebSocket 协议 |
| URL | Uniform Resource Locator | 统一资源定位符 |
| PPT | PowerPoint | Microsoft PowerPoint |
