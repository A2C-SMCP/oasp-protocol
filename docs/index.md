# OASP - Office AddIn Socket Protocol

> Office AddIn 与后端服务之间的实时通信协议规范

## 什么是 OASP？

OASP (Office AddIn Socket Protocol) 是一个基于 Socket.IO 的通信协议，专为 **AI Agent 控制和操作 Microsoft Office 文档**而设计。

协议定义了：

- **连接与握手** - 如何建立和管理 Socket.IO 连接
- **事件格式** - 请求/响应的数据结构
- **错误处理** - 统一的错误码体系
- **业务规则** - 各类操作的行为规范

## 设计目标

1. **AI Agent 友好** - 为程序化操作 Office 文档优化，而非人工交互
2. **类型安全** - 所有数据结构都有严格的类型定义
3. **跨应用统一** - Word、PPT、Excel 使用一致的协议模式
4. **可扩展** - 易于添加新的事件和功能

## 支持的应用

| 应用 | 命名空间 | 状态 |
|------|----------|------|
| Microsoft Word | `/word` | ✅ Stable |
| Microsoft PowerPoint | `/ppt` | 📋 Draft |
| Microsoft Excel | `/excel` | 📋 Draft |

## 快速导航

- **[协议概述](specification/index.md)** - 了解协议的核心概念
- **[架构设计](specification/architecture.md)** - 角色、通信模型、数据流
- **[连接与握手](specification/connection.md)** - 如何建立连接
- **[Word 事件](specification/events-word.md)** - Word 命名空间的所有事件
- **[数据结构](specification/data-structures.md)** - 通用的数据类型定义
- **[错误处理](specification/error-handling.md)** - 错误码和异常处理
- **[术语表](appendix/glossary.md)** - Office 相关术语解释

## 版本

当前协议版本: **0.1.0**

查看[变更日志](appendix/changelog.md)了解版本历史。
