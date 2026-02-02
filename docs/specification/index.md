# 协议概述

## 协议简介

OASP (Office AddIn Socket Protocol) 是一个应用层协议，基于 Socket.IO 传输层实现 AI Agent 与 Office AddIn 之间的双向实时通信。

## 核心概念

### 角色模型

OASP 采用**两角色模型**：

```
┌─────────────────┐                    ┌─────────────────┐
│     Server      │◄──── Socket.IO ────►│     AddIn       │
│  (Python 后端)  │                    │  (Office 插件)   │
└─────────────────┘                    └─────────────────┘
       ▲
       │ API/MCP
       ▼
┌─────────────────┐
│    AI Agent     │
└─────────────────┘
```

| 角色 | 说明 |
|------|------|
| **Server** | Python 后端服务，接收 AI Agent 指令，通过 Socket.IO 与 AddIn 通信 |
| **AddIn** | Office 插件客户端，执行 Office 操作，上报文档状态变化 |
| **AI Agent** | 智能体，通过 MCP/API 调用 Server 提供的能力（不直接参与协议） |

### 命名空间

每个 Office 应用对应一个独立的 Socket.IO 命名空间：

| 命名空间 | 应用 | 状态 |
|----------|------|------|
| `/word` | Microsoft Word | ✅ Stable |
| `/ppt` | Microsoft PowerPoint | 📋 Draft |
| `/excel` | Microsoft Excel | 📋 Draft |

### 事件分类

OASP 事件分为两大类：

#### 1. 请求-响应事件

Server 发起请求，AddIn 返回响应。用于查询数据或执行操作。

```
Server ──[event:request]──► AddIn
Server ◄──[callback/ack]─── AddIn
```

#### 2. 事件报告

AddIn 单向通知 Server，无需响应。用于上报文档状态变化。

```
Server ◄──[event:report]─── AddIn
```

### 事件命名规范

事件名采用 `{namespace}:{action}:{target}` 格式：

- `{namespace}` - 应用标识：`word`、`ppt`、`excel`
- `{action}` - 动作类型：`get`、`insert`、`replace`、`select`、`event`、`export`
- `{target}` - 操作目标：`selection`、`text`、`styles` 等

示例：
```
word:get:selection       # 获取选区位置
word:insert:text         # 插入文本
word:event:selectionChanged  # 选区变化事件
```

## 通信模式

### 请求-响应模式

所有请求-响应事件都遵循统一的模式：

1. Server 发送请求，包含 `requestId`
2. AddIn 执行操作
3. AddIn 通过 Socket.IO 的 ack 机制返回响应
4. 如果超时，Server 标记请求失败

```json
// 请求
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "documentUri": "file:///path/to/document.docx",
  "timestamp": 1704067200000,
  // ... 业务参数
}

// 响应
{
  "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
  "success": true,
  "data": { /* 返回数据 */ },
  "timestamp": 1704067200500,
  "duration": 500
}
```

### 事件报告模式

事件报告是单向的，AddIn 发送后不等待响应：

```json
{
  "documentUri": "file:///path/to/document.docx",
  "timestamp": 1704067200000,
  // ... 事件数据
}
```

## 连接约束

- **单客户端模式**: 每个文档同时只有一个 AddIn 客户端连接
- **文档隔离**: 不同文档的操作相互独立
- **自动超时**: 请求超时后自动标记失败，不进行重试

## 状态标记

协议文档中的事件使用以下状态标记：

| 标记 | 含义 |
|------|------|
| ✅ Stable | 已实现且稳定，可在生产环境使用 |
| 📋 Draft | 已定义但尚未实现，接口可能变更 |
| ⚠️ Deprecated | 已废弃，将在未来版本移除 |
