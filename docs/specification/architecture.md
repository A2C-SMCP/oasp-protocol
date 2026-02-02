# 架构设计

## 系统架构

### 整体架构图

```
┌──────────────────────────────────────────────────────────────┐
│                        AI Agent Layer                        │
│  ┌────────────────────────────────────────────────────────┐  │
│  │                      AI Agent                          │  │
│  │         (LLM / AutoGPT / Custom Agent)                │  │
│  └────────────────────────────────────────────────────────┘  │
│                            │ MCP / API                       │
└────────────────────────────┼─────────────────────────────────┘
                             ▼
┌──────────────────────────────────────────────────────────────┐
│                        Server Layer                          │
│  ┌────────────────────────────────────────────────────────┐  │
│  │              Python Backend (Office4AI)               │  │
│  │  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐ │  │
│  │  │ MCP Server   │  │  Socket.IO   │  │  Connection  │ │  │
│  │  │   Tools      │  │   Server     │  │   Manager    │ │  │
│  │  └──────────────┘  └──────────────┘  └──────────────┘ │  │
│  └────────────────────────────────────────────────────────┘  │
│                            │ Socket.IO (OASP)                │
└────────────────────────────┼─────────────────────────────────┘
                             ▼
┌──────────────────────────────────────────────────────────────┐
│                       AddIn Layer                            │
│  ┌──────────────────────────────────────────────────────────┐│
│  │                   Office AddIn                           ││
│  │  ┌────────────┐  ┌────────────┐  ┌────────────────────┐ ││
│  │  │ Socket.IO  │  │   Event    │  │    Office.js       │ ││
│  │  │  Client    │  │  Handler   │  │      API           │ ││
│  │  └────────────┘  └────────────┘  └────────────────────┘ ││
│  └──────────────────────────────────────────────────────────┘│
│                            │ Office.js API                   │
│  ┌──────────────────────────────────────────────────────────┐│
│  │              Microsoft Office Application                ││
│  │          (Word / PowerPoint / Excel)                     ││
│  └──────────────────────────────────────────────────────────┘│
└──────────────────────────────────────────────────────────────┘
```

## 数据流

### 查询操作数据流

以「获取选区内容」为例：

```
AI Agent                Server                  AddIn                   Office
    │                      │                      │                       │
    │──[1] get_selection──►│                      │                       │
    │                      │──[2] word:get:──────►│                       │
    │                      │    selection         │──[3] Office.js ──────►│
    │                      │                      │    context.document   │
    │                      │                      │    .getSelection()    │
    │                      │                      │◄─[4] Selection ───────│
    │                      │◄─[5] ack ────────────│                       │
    │◄─[6] SelectionInfo───│                      │                       │
    │                      │                      │                       │
```

### 修改操作数据流

以「插入文本」为例：

```
AI Agent                Server                  AddIn                   Office
    │                      │                      │                       │
    │──[1] insert_text ───►│                      │                       │
    │    (text, format)    │                      │                       │
    │                      │──[2] word:insert:───►│                       │
    │                      │    text              │──[3] Office.js ──────►│
    │                      │                      │    range.insertText() │
    │                      │                      │    range.font.bold    │
    │                      │                      │◄─[4] Success ─────────│
    │                      │◄─[5] ack ────────────│                       │
    │◄─[6] {success:true}──│                      │                       │
    │                      │                      │                       │
```

### 事件报告数据流

以「选区变化通知」为例：

```
AI Agent                Server                  AddIn                   Office
    │                      │                      │                       │
    │                      │                      │      [用户操作]        │
    │                      │                      │◄─[1] onSelectionChanged│
    │                      │◄─[2] word:event:────│                       │
    │                      │    selectionChanged  │                       │
    │   [可选处理]          │                      │                       │
    │                      │                      │                       │
```

## 组件职责

### Server 组件

| 组件 | 职责 |
|------|------|
| **MCP Server** | 向 AI Agent 暴露工具接口，处理 MCP 协议 |
| **Socket.IO Server** | 管理与 AddIn 的 Socket.IO 连接 |
| **Connection Manager** | 维护连接状态、文档到连接的映射 |
| **Request Wrapper** | 将业务参数封装为标准请求格式 |
| **DTO Registry** | 管理事件与 DTO 类型的映射 |

### AddIn 组件

| 组件 | 职责 |
|------|------|
| **Socket.IO Client** | 管理与 Server 的 Socket.IO 连接 |
| **Event Handler** | 处理各类事件，调用 Office.js API |
| **Office.js API** | Microsoft 官方 API，操作 Office 文档 |

## 连接管理

### 连接映射

Server 维护以下映射关系：

```
┌─────────────────────────────────────────────────────────────┐
│                    Connection Manager                       │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  _clients: Dict[socket_id → ClientInfo]                    │
│     管理所有活跃连接                                          │
│                                                             │
│  _document_to_sockets: Dict[document_uri → Set[socket_id]] │
│     按文档查找连接（一个文档可能有多个历史连接）                  │
│                                                             │
│  _client_id_to_socket: Dict[client_id → socket_id]        │
│     按客户端 ID 查找连接                                      │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

### ClientInfo 数据结构

```python
class ClientInfo:
    socket_id: str       # Socket.IO 会话 ID
    client_id: str       # 客户端提供的唯一标识
    document_uri: str    # 正在处理的文档 URI
    namespace: str       # 命名空间 (/word, /ppt, /excel)
    connected_at: float  # 连接时间戳
```

## 命名空间隔离

每个 Office 应用使用独立的 Socket.IO 命名空间：

```
Socket.IO Server
├── /word    ──► WordNamespace    ──► Word AddIn
├── /ppt     ──► PPTNamespace     ──► PowerPoint AddIn
└── /excel   ──► ExcelNamespace   ──► Excel AddIn
```

命名空间之间完全隔离：
- 事件不会跨命名空间传播
- 连接状态独立管理
- 错误不会相互影响
