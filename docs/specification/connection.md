# 连接与握手

## 概述

本章定义 AddIn 与 Server 之间建立连接的完整流程，包括握手参数、认证机制和连接生命周期管理。

## 连接流程

```
AddIn                                    Server
  │                                        │
  │──[1] connect(auth, headers)───────────►│
  │                                        │  验证握手参数
  │                                        │  注册到 ConnectionManager
  │◄─[2] connection:established───────────│
  │                                        │
  │         ... 正常通信 ...                 │
  │                                        │
  │──[N] disconnect ──────────────────────►│
  │                                        │  从 ConnectionManager 移除
  │                                        │
```

## 握手参数

### 必填参数

AddIn 连接时**必须**在 `auth` 对象中提供以下参数：

| 参数 | 类型 | 说明 |
|------|------|------|
| `clientId` | `string` | 客户端唯一标识，用于标识同一个 AddIn 实例 |
| `documentUri` | `string` | 当前文档的 URI，格式为 `file:///path/to/document.ext` |

### 连接示例

=== "TypeScript (AddIn)"

    ```typescript
    import { io, Socket } from "socket.io-client";

    const socket: Socket = io("http://127.0.0.1:3000/word", {
      auth: {
        clientId: "addin-instance-uuid",
        documentUri: "file:///Users/john/Documents/report.docx"
      },
      reconnection: true,
      reconnectionAttempts: 5,
      reconnectionDelay: 2000,
      reconnectionDelayMax: 10000
    });

    socket.on("connection:established", (data) => {
      console.log("Connected successfully:", data);
    });

    socket.on("connect_error", (error) => {
      console.error("Connection failed:", error);
    });
    ```

=== "Python (Server)"

    ```python
    from socketio import AsyncServer

    sio = AsyncServer(cors_allowed_origins=["https://localhost:3002"])

    @sio.on("connect", namespace="/word")
    async def handle_connect(sid, environ, auth):
        # 验证必填参数
        if not auth or "clientId" not in auth or "documentUri" not in auth:
            raise ConnectionRefusedError("Missing required auth parameters")

        client_id = auth["clientId"]
        document_uri = auth["documentUri"]

        # 注册连接
        connection_manager.register_client(
            socket_id=sid,
            client_id=client_id,
            document_uri=document_uri,
            namespace="/word"
        )

        # 发送确认
        await sio.emit(
            "connection:established",
            {"socketId": sid, "timestamp": int(time.time() * 1000)},
            room=sid,
            namespace="/word"
        )
    ```

## 连接确认事件

### connection:established

**方向**: Server → AddIn

**触发时机**: 握手成功后立即发送

**数据结构**:

```typescript
interface ConnectionEstablishedData {
  socketId: string;      // 分配的 Socket.IO 会话 ID
  timestamp: number;     // 服务器时间戳（毫秒）
}
```

**示例**:

```json
{
  "socketId": "abc123xyz",
  "timestamp": 1704067200000
}
```

## 握手失败处理

### 错误场景

| 场景 | 错误码 | 说明 |
|------|--------|------|
| 缺少 `clientId` | `HANDSHAKE_FAILED` | auth 中未提供 clientId |
| 缺少 `documentUri` | `HANDSHAKE_FAILED` | auth 中未提供 documentUri |
| 无效的 `documentUri` | `HANDSHAKE_FAILED` | URI 格式不正确 |
| 服务器内部错误 | `UNKNOWN` | 服务器处理异常 |

### 错误响应

握手失败时，Socket.IO 会触发 `connect_error` 事件：

```typescript
socket.on("connect_error", (error) => {
  // error.message 包含错误原因
  console.error("Handshake failed:", error.message);
});
```

## 断开连接

### 正常断开

AddIn 主动断开连接时：

```typescript
socket.disconnect();
```

Server 端会收到 `disconnect` 事件并自动清理连接信息。

### 异常断开

网络中断等异常情况下，Socket.IO 的心跳机制会检测到连接断开：

- **ping_timeout**: 60 秒
- **ping_interval**: 25 秒

超过 ping_timeout 未收到心跳响应时，Server 认定连接已断开。

## 重连机制

### AddIn 端重连

建议配置：

```typescript
const socket = io(url, {
  reconnection: true,           // 启用自动重连
  reconnectionAttempts: 5,      // 最多尝试 5 次
  reconnectionDelay: 2000,      // 初始延迟 2 秒
  reconnectionDelayMax: 10000   // 最大延迟 10 秒
});
```

### 重连后的处理

重连成功后，AddIn 需要：

1. 重新发送握手参数（Socket.IO 自动处理）
2. 等待 `connection:established` 确认
3. 重新订阅需要的事件监听（如有）

!!! warning "重连期间的请求"
    重连期间 Server 发送的请求将无法到达 AddIn，会触发超时失败。

## Server 端配置

### 推荐配置

```python
sio = AsyncServer(
    cors_allowed_origins=[
        # Python Socket.IO 客户端
        "http://localhost:3000",
        "http://127.0.0.1:3000",
        # Word AddIn (HTTPS)
        "https://localhost:3002",
        "https://127.0.0.1:3002",
        # Excel AddIn
        "https://localhost:3001",
        "https://127.0.0.1:3001",
        # PowerPoint AddIn
        "https://localhost:3003",
        "https://127.0.0.1:3003",
    ],
    ping_timeout=60,        # 心跳超时 60 秒
    ping_interval=25,       # 心跳间隔 25 秒
    max_http_buffer_size=1024 * 1024  # 最大消息 1MB
)
```

## 连接状态查询

### 检查文档是否有活跃连接

```python
is_active = connection_manager.is_document_active("file:///path/to/doc.docx")
```

### 获取处理特定文档的连接

```python
socket_id = connection_manager.get_socket_by_document("file:///path/to/doc.docx")
```

### 获取连接统计

```python
connection_count = connection_manager.get_connection_count()
document_count = connection_manager.get_document_count()
```
