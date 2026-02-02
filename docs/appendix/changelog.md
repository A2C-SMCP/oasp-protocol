# 变更日志

本文档记录 OASP 协议的所有重要变更。

格式基于 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.0.0/)，
版本号遵循 [语义化版本](https://semver.org/lang/zh-CN/)。

---

## [0.1.0] - 2026-02-02

### 新增

**核心协议**

- 定义了协议名称: OASP (Office AddIn Socket Protocol)
- 定义了两角色通信模型 (Server ↔ AddIn)
- 定义了三个命名空间: `/word`, `/ppt`, `/excel`
- 定义了事件命名规范: `{namespace}:{action}:{target}`
- 定义了请求-响应模式和事件报告模式

**连接与握手**

- 定义了握手参数: `clientId`, `documentUri`
- 定义了连接确认事件: `connection:established`
- 定义了断开连接和重连机制

**Word 事件 (✅ Stable)**

- `word:event:selectionChanged` - 选区变化通知
- `word:event:documentModified` - 文档修改通知
- `word:get:selection` - 获取选区位置信息
- `word:get:selectedContent` - 获取选中内容
- `word:get:visibleContent` - 获取可见内容
- `word:get:documentStructure` - 获取文档结构
- `word:get:documentStats` - 获取文档统计
- `word:get:styles` - 获取文档样式列表
- `word:insert:text` - 插入文本
- `word:replace:selection` - 替换选中内容
- `word:replace:text` - 查找替换
- `word:select:text` - 查找并选中文本

**Word 事件 (📋 Draft)**

- `word:append:text` - 追加文本
- `word:insert:image` - 插入图片
- `word:insert:table` - 插入表格
- `word:insert:equation` - 插入公式
- `word:insert:toc` - 插入目录
- `word:export:content` - 导出内容

**PPT 事件 (📋 Draft)**

- `ppt:get:currentSlideElements` - 获取当前幻灯片元素
- `ppt:get:slideElements` - 获取指定幻灯片元素
- `ppt:get:slideScreenshot` - 获取幻灯片截图
- `ppt:insert:text` - 插入文本
- `ppt:insert:shape` - 插入形状
- `ppt:insert:image` - 插入图片
- `ppt:slide:add` - 添加幻灯片
- `ppt:slide:delete` - 删除幻灯片
- `ppt:slide:move` - 移动幻灯片
- `ppt:slide:goto` - 跳转到幻灯片

**Excel 事件 (📋 Draft)**

- `excel:get:selectedRange` - 获取选中范围
- `excel:get:usedRange` - 获取已使用范围
- `excel:get:cellValue` - 获取单元格值
- `excel:get:rangeValues` - 获取范围内的值
- `excel:set:cellValue` - 设置单元格值
- `excel:set:rangeValues` - 设置范围内的值
- `excel:insert:table` - 插入表格
- `excel:insert:chart` - 插入图表
- `excel:sheet:add` - 添加工作表
- `excel:sheet:delete` - 删除工作表
- `excel:sheet:rename` - 重命名工作表
- `excel:sheet:activate` - 激活工作表

**数据结构**

- 定义了基础请求/响应结构
- 定义了 `SelectionInfo`, `TextFormat`, `StyleInfo` 等核心类型
- 定义了 PPT 和 Excel 相关类型

**错误处理**

- 定义了错误码分类 (1xxx-4xxx)
- 定义了通用错误码 (1xxx)
- 定义了连接错误码 (2xxx)
- 定义了文档操作错误码 (3xxx)
- 定义了参数验证错误码 (4xxx)

**通用约定**

- 定义了时间戳格式 (Unix 毫秒)
- 定义了字段命名规范 (camelCase)
- 定义了颜色值格式 (#RRGGBB)
- 定义了单位规范 (磅、像素)

**文档**

- 创建了完整的协议文档结构
- 创建了术语表
- 创建了变更日志

---

## 版本号说明

- **主版本号 (Major)**: 不兼容的 API 变更
- **次版本号 (Minor)**: 向后兼容的功能新增
- **修订号 (Patch)**: 向后兼容的问题修复

当前版本 `0.1.0` 表示协议处于初始开发阶段，API 可能随时变更。
