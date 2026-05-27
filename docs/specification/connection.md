# 连接与握手

## 概述

本章定义 AddIn 与 Server 之间建立连接的完整流程，包括握手参数、认证机制和连接生命周期管理。

## 连接流程

```
AddIn                                    Server
  │                                        │
  │──[1] connect(auth, headers)───────────►│
  │                                        │  验证握手参数（含 oaspVersion 兼容性）
  │                                        │  不兼容 → 拒绝（PROTOCOL_VERSION_MISMATCH）
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
| `oaspVersion` | `string` | AddIn 实现的 OASP 协议版本（SemVer，如 `0.3.0`）。Server 据此判定兼容性，不兼容则拒绝连接。详见[协议版本握手](#protocol-version-handshake) |

### 连接示例

=== "TypeScript (AddIn)"

    ```typescript
    import { io, Socket } from "socket.io-client";

    const socket: Socket = io("http://127.0.0.1:3000/word", {
      auth: {
        clientId: "addin-instance-uuid",
        documentUri: "file:///Users/john/Documents/report.docx",
        oaspVersion: "0.3.0"   // AddIn SDK 自动从内置版本常量注入
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
        # ① 协议版本校验先行（协议层先于业务层），详见「协议版本握手」节
        raw_ver = (auth or {}).get("oaspVersion")
        if not raw_ver:
            raise ConnectionRefusedError(
                {"code": "HANDSHAKE_FAILED", "message": "Missing oaspVersion"}
            )
        try:
            client_ver = OaspVersion.parse(raw_ver)
        except ValueError:
            raise ConnectionRefusedError(
                {"code": "HANDSHAKE_FAILED", "message": f"Invalid oaspVersion: {raw_ver}"}
            )
        if not is_compatible(client_ver, SERVER_VERSION):
            raise ConnectionRefusedError({
                "code": "PROTOCOL_VERSION_MISMATCH",
                "message": "Protocol version mismatch",
                "serverVersion": str(SERVER_VERSION),
                "clientVersion": str(client_ver),
                "minSupported": str(SERVER_MIN_SUPPORTED),
                "maxSupported": str(SERVER_MAX_SUPPORTED),
            })

        # ② 业务参数校验（版本兼容后再查）——缺失 → HANDSHAKE_FAILED
        if "clientId" not in auth or "documentUri" not in auth:
            raise ConnectionRefusedError(
                {"code": "HANDSHAKE_FAILED", "message": "Missing required auth parameters"}
            )

        # ③ 注册连接
        connection_manager.register_client(
            socket_id=sid,
            client_id=auth["clientId"],
            document_uri=auth["documentUri"],
            namespace="/word"
        )

        # 发送确认（含 serverVersion 供诊断）
        await sio.emit(
            "connection:established",
            {"socketId": sid, "serverVersion": str(SERVER_VERSION), "timestamp": int(time.time() * 1000)},
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
  serverVersion: string; // Server 的 OASP 协议版本（SemVer），供 AddIn 展示与诊断
  timestamp: number;     // 服务器时间戳（毫秒）
}
```

!!! note "serverVersion 仅供诊断"
    握手通过即代表版本已兼容（见[协议版本握手](#protocol-version-handshake)），`serverVersion` 不用于 AddIn 端二次校验，仅用于日志、UI 展示与故障定位。

**示例**:

```json
{
  "socketId": "abc123xyz",
  "serverVersion": "0.3.0",
  "timestamp": 1704067200000
}
```

## 协议版本握手 {#protocol-version-handshake}

为避免新旧两端协议错位，AddIn **MUST** 在 `auth.oaspVersion` 中声明自己实现的协议版本，Server 在 `connect` handler 中校验兼容性。版本号语义与兼容性判定规则见[通用约定 · 协议版本与兼容性](conventions.md#versioning)。

### 校验时机与位置

- 校验在 Socket.IO 的 `connect` handler 内完成，**先于任何业务注册**（注册 ConnectionManager、emit `connection:established` 之前）
- 不兼容时 **MUST** 抛 `ConnectionRefusedError` 并携带结构化拒绝数据；连接不被建立，AddIn 不进入正常通信
- OASP 为两方单连接模型，无 A2C 三方房间的传递性需求，故**无需** HTTP 中间件层校验——`connect` handler 已是业务代码的最前沿，且 `ConnectionRefusedError` 的数据会原样送达客户端 `connect_error`，对 polling 与 websocket 两种 transport 一致生效（无需 A2C 的 polling-first 约束与 WS close code 回退）

### 校验顺序

版本（协议层）**先于**业务参数（业务层）校验——不兼容时直接报版本错，比"缺 clientId"更切中根因：

```
1. 缺少 oaspVersion / 格式非法      → 拒绝，code HANDSHAKE_FAILED（复用 2003）
2. 版本不兼容（is_compatible 假）    → 拒绝，code PROTOCOL_VERSION_MISMATCH（2006）+ 结构化版本字段
3. 缺少 clientId / documentUri      → 拒绝，code HANDSHAKE_FAILED
4. 全部通过                         → 注册连接，emit connection:established（含 serverVersion）
```

### 拒绝数据结构

版本握手的拒绝数据是**扁平结构**（非标准 `ErrorResponse`——握手期没有 `requestId`），经 `ConnectionRefusedError` 送达客户端 `connect_error` 的 `error.data`：

```typescript
interface HandshakeRejection {
  code: string;            // "PROTOCOL_VERSION_MISMATCH" 或 "HANDSHAKE_FAILED"
  message: string;         // 人类可读
  serverVersion?: string;  // 仅 PROTOCOL_VERSION_MISMATCH：Server 当前版本
  clientVersion?: string;  // 仅 PROTOCOL_VERSION_MISMATCH：从 auth.oaspVersion 读取
  minSupported?: string;   // 仅 PROTOCOL_VERSION_MISMATCH：Server 支持的最低版本
  maxSupported?: string;   // 仅 PROTOCOL_VERSION_MISMATCH：Server 支持的最高版本
}
```

版本不兼容示例：

```json
{
  "code": "PROTOCOL_VERSION_MISMATCH",
  "message": "Protocol version mismatch",
  "serverVersion": "0.3.0",
  "clientVersion": "0.2.0",
  "minSupported": "0.3.0",
  "maxSupported": "0.3.999"
}
```

### Server 校验示例（Python）

```python
# is_compatible / OaspVersion 参考实现见 conventions.md#compatibility-rule
SERVER_VERSION = OaspVersion(0, 3, 0)
SERVER_MIN_SUPPORTED = OaspVersion(0, 3, 0)
SERVER_MAX_SUPPORTED = OaspVersion(0, 3, 999)

@sio.on("connect", namespace="/word")
async def handle_connect(sid, environ, auth):
    raw = (auth or {}).get("oaspVersion")
    if not raw:
        raise ConnectionRefusedError({"code": "HANDSHAKE_FAILED", "message": "Missing oaspVersion"})
    try:
        client_ver = OaspVersion.parse(raw)
    except ValueError:
        raise ConnectionRefusedError({"code": "HANDSHAKE_FAILED", "message": f"Invalid oaspVersion: {raw}"})
    if not is_compatible(client_ver, SERVER_VERSION):
        raise ConnectionRefusedError({
            "code": "PROTOCOL_VERSION_MISMATCH",
            "message": "Protocol version mismatch",
            "serverVersion": str(SERVER_VERSION),
            "clientVersion": str(client_ver),
            "minSupported": str(SERVER_MIN_SUPPORTED),
            "maxSupported": str(SERVER_MAX_SUPPORTED),
        })
    # ... 校验通过后注册连接并 emit connection:established
```

### AddIn 收到拒绝后的行为

收到 `connect_error` 且 `error.data.code === "PROTOCOL_VERSION_MISMATCH"` 时，AddIn **MUST**：

- **主动断开**（`socket.disconnect()`）并**关闭自动重连**——不得静默重试。Socket.IO 客户端默认 `reconnection: true`，握手被拒后会立即重连再次触发同一拒绝，进入死循环烧 CPU 与 Server 资源
- 将错误上抛为明确的版本不匹配异常（携带 `serverVersion` / `clientVersion`），交上层决定（升级 AddIn / 提示用户 / 上报运维）

```typescript
socket.on("connect_error", (err) => {
  const payload = (err as any).data;          // ConnectionRefusedError 携带的数据
  if (payload?.code === "PROTOCOL_VERSION_MISMATCH") {
    socket.disconnect();                       // MUST：阻断自动重连死循环
    throw new ProtocolVersionError(payload);   // 交上层处理（携带 server/client 版本）
  }
});
```

## 握手失败处理

### 错误场景

| 场景 | 错误码 | 说明 |
|------|--------|------|
| 缺少 `clientId` | `HANDSHAKE_FAILED` | auth 中未提供 clientId |
| 缺少 `documentUri` | `HANDSHAKE_FAILED` | auth 中未提供 documentUri |
| 无效的 `documentUri` | `HANDSHAKE_FAILED` | URI 格式不正确 |
| 缺少/非法 `oaspVersion` | `HANDSHAKE_FAILED` | auth 中未提供或格式非法 |
| 协议版本不兼容 | `PROTOCOL_VERSION_MISMATCH` | `oaspVersion` 与 Server 不兼容，详见[协议版本握手](#protocol-version-handshake) |
| 服务器内部错误 | `UNKNOWN` | 服务器处理异常 |

### 错误响应

握手失败时，Socket.IO 会触发 `connect_error` 事件：

```typescript
socket.on("connect_error", (error) => {
  // error.message 包含错误原因；结构化拒绝数据在 error.data（见上方 HandshakeRejection）
  console.error("Handshake failed:", error.message, (error as any).data);
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
