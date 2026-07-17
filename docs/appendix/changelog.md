# 变更日志

本文档记录 OASP 协议的所有重要变更。

格式基于 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.0.0/)，
版本号遵循 [语义化版本](https://semver.org/lang/zh-CN/)。

---

## [Unreleased]

### 裁决四个 3xxx 孤儿错误码 —— `3012` 收敛 / `3007`·`3008` 接线 / `3005` 退役（订正非合规降级）

裁决 [#23](https://github.com/A2C-SMCP/oasp-protocol/issues/23)：`3012 SEARCH_NO_MATCH` 注册表已定义但**全规范零引用**，而 `word:insert:comment` 反把「搜索文本未找到匹配」判给通用 `3000`——正是 `events-excel.md` 规范层「**不得**降级为通用 `3000 DOCUMENT_ERROR`」所禁止的模式。

**根因**：守护器**只校验配对、不校验可达性**。#21 引入的 `inv docs.check-error-codes` 断言每处 `(编号, 名称)` 精确命中注册表，但注册表里「**有码无人引用**」它抓不到——`3000` / `DOCUMENT_ERROR` 是合法配对，故守护器放行合理，#23 系人工审查**偶然**发现。排查证实问题比 #23 描述的更系统：**注册表 37 码中 14 码零引用**，其中 3xxx 有 `3005` / `3007` / `3008` / `3012` 四码，各需独立裁决（1xxx 通用失败 / 2xxx 连接生命周期属跨切面，非事件表可枚举项，豁免合理）。

| 码 | 裁决 | 依据 |
|---|---|---|
| `3012` `SEARCH_NO_MATCH` | **收敛**（#23 本体） | 有且仅有一个合法归属：`word:insert:comment` 的 `searchText` **定位**模式。`3000` 行恢复为纯通用兜底 |
| `3007` `FORMAT_NOT_SUPPORTED` | **接上** | 有真实触发点，现被降级进 `4002`（"Base64 数据无效"）或 `3000` 兜底。承 `4003`↔`3018` 已确立的「线缆层 / apply-time」配对先例，在**格式**轴上的应用 |
| `3008` `POSITION_INVALID` | **接上**（本次立规 + `/ppt` 样板） | 「apply-time 序号位置无效」真空位：`slideIndex: 15` 对 10 页片**线缆上完全合法**，仅相对文档状态无效；正确恢复是「重读状态后原值重试」，而 4xxx 暗示的是「改参数」。`4002` / `4004` / `3010` / `3009` 表达它都要对成因说谎 |
| `3005` `RESOURCE_NOT_ACCESSIBLE` | **退役** | **无指涉对象**：全协议无任何请求参数引用可拉取的外部资源——`documentUri` 是唯一 URI 参数且已由 `3001`/`3003` 覆盖；图片等载荷一律 inline base64（自包含，无需解引用）；所有 URL / 链接字段均为**响应**字段，不可能触发请求错误 |

**新增规范层判法——「搜索即定位」vs「搜索即查询」**（`error-handling.md` §`SEARCH_NO_MATCH (3012)`）。搜索在协议中承担两种角色，`3012` **只适用于前者**：

| 角色 | 无匹配时 | 事件 |
|---|---|---|
| **定位**（locator）——搜索结果是变更操作的**锚点** | 操作无法进行 → **MUST** 返回 `3012` | `word:insert:comment`（`target.type: "searchText"`） |
| **查询**（query）——搜索结果**本身**即返回值 | 是**正常空结果** → `success: true`，**不得**返回 `3012` | `word:replace:text`（`replaceCount: 0`）、`word:select:text`（`matchCount: 0`）、`excel:find:values`（`matches: []`） |

故 #23 建议方向 2（其余搜索事件是否也列 `3012`）判**否**且有原则支撑，方向 3（退役 `3012`）判**否**。方向 4（`/cross-ask office-editor4ai` 确认宿主侧能否区分「无匹配」与「操作失败」）**无需执行**——`word:select:text` 规范早已把 `matchCount: 0` 写成**正常成功响应**，即仓内自证该区分在线缆上可观测。

| 变更类型 | 位置 | 说明 |
|----------|------|------|
| 收敛（语义订正） | `events-word.md` | `word:insert:comment` 的「搜索文本未找到匹配」`3000` → `3012`；`3000` 行恢复为纯通用兜底 |
| **移除（码退役）** | `error-handling.md` | 删除 `3005 RESOURCE_NOT_ACCESSIBLE`，**不留别名** |
| 新增（判法立规） | `error-handling.md` | 三个触发场景小节：`3012`（定位/查询划界，规范层）、`3007`（与 `4002` 载荷不可解码、`3016` 宿主能力缺失的划界）、`3008`（与 `4004` 静态边界、`3010` 具名查找的划界；`details` 带 `{ index, total, kind }`）。均为**散文体、不含 JSON 示例**——绕开尚未裁决的 [#22](https://github.com/A2C-SMCP/oasp-protocol/issues/22)（`error.code` 装名称还是数字），不新增线缆读法债务 |
| 接线（`3007`） | `events-{word,ppt}.md`（5 处） | `word:insert:image`、`ppt:insert:image`、`ppt:update:image` 拆 `4002`（不可解码，线缆层）+ `3007`（可解码但格式不受支持）；`word:export:content`、`ppt:get:slideScreenshot` 补 `3007`（枚举合法但宿主无法产出/渲染，原落 `3000` 兜底） |
| 接线（`3008` 样板） | `events-ppt.md`（2 处） | `ppt:delete:slide`、`ppt:goto:slide` 的 `slideIndex 超出范围` `4002` → `3008`（并补 `details`）。**仅取 `/ppt`（Draft）两处最无歧义样板**，余下横扫重指另开 issue |
| 新增（守护） | `scripts/docs/tasks.py` | `check_error_codes` 补**可达性闸口**：3xxx/4xxx 码**必须**被 `error-handling.md` **之外**的至少一处引用（「之外」是关键——否则 `### TIMEOUT (1002)` 这类详解小节标题会让码自我满足），否则报孤儿；1xxx/2xxx **按码段豁免**（跨切面，强求引用只会逼出噪声）。`_orphans()` 抽为可测函数，补 5 个可达性自检样本。**无白名单**——四码裁决后 3xxx/4xxx 全部可达，以硬闸口落地。修复前报 4 个孤儿、修复后 400 处引用全绿 + 23 个 3xxx/4xxx 码均可达（13 个 1xxx/2xxx 豁免） |

**兼容性**：定性为**订正非合规降级**，非破坏性契约变更——

- `3012`：`word:insert:comment`（✅ Stable）「搜索无匹配」由 `3000` → `3012`。向 `3000` 降级本就违反 `events-excel.md` 规范层 **MUST**，属订正不合规实现；调用方原先**无法**区分该条件与「文档操作失败」，收敛后信号严格增强
- `3007`：多为**新增行**（该条件原落 `3000` 兜底）；`/ppt` 图片事件属 `4002` 拆分，Draft 无稳定性承诺
- `3008`：本次仅动 `/ppt`（Draft）。`/word`（Stable）3 处与 `/excel` 3 处横扫重指留待专项 issue，届时逐条评估
- `3005`：退役零风险——`git log -S` 证实该码自初始提交 `069a2d8` 起**从未被任何事件引用**，无实现可能发出它。对齐 #20 退役 `3999`、#17 退役 `5xxx` 先例

跨仓跟进——office-editor4ai：`word:insert:comment` 的 searchText 无匹配分支改发 `3012`；图片事件按「不可解码 / 格式不受支持」二分 `4002` / `3007`；`ppt:{delete,goto}:slide` 越界改发 `3008`（带 `details`）。office4ai：同步 `SEARCH_NO_MATCH` / `FORMAT_NOT_SUPPORTED` / `POSITION_INVALID`，移除 `RESOURCE_NOT_ACCESSIBLE`。

**遗留**（均另开 issue，不在本次范围）：

1. **`3008` 横扫重指**：余 ~24 行「index 超出范围」仍判 `4002`（`/ppt` 16 处、`/word` Stable 3 处）或 `4004`（`/excel` 3 处），另有 ~7 处静默缺口。两处**保留区**须一并裁决：`ppt:update:textBox` 的 `start`/`length`（字符偏移，`data-structures.md` 钉在 `4002`）、`word:insert:image` 的 `bookmarkName`（属**具名查找** → 宜 `3010` 新增 `kind:"bookmark"`，非 `3008`）
2. **同条件跨命名空间判法不一致**：「index 超出范围」在 `/ppt`+`/word` 判 `4002`、在 `/excel` 判 `4004`，规范未陈述理由
3. **`ImageData` 重复定义且语义分叉**：`events-word.md` 为 `word:insert:image` 本地重声明 `ImageData` 并**丢掉 `mimeType`**，与权威 `data-structures.md` 冲突
4. **`imageInfo.data` 不可达**：`events-ppt.md` 标注「需显式请求」，但 `GetSlideInfoRequest` 无对应参数
5. **`excel:set:rangeFormat` 的 `numberFormat`**：非法/不受支持归 `3007` 还是 `4002` 待裁决（属格式**字符串**而非格式**种类**，故本次未接线，`events-excel.md` 规范层映射表暂不加 `3007` 行）

### 事件表错误码与权威注册表对账 —— 退役 `OFFICE_API_ERROR` 与 `3999`（文档订正，线缆影响取决于读法裁决）

裁决 [#20](https://github.com/A2C-SMCP/oasp-protocol/issues/20)：`OFFICE_API_ERROR` 在 `events-word.md` 内部前半用 `3999`、后半用 `3000`，而权威注册表（`error-handling.md`）**两个号都不认**——`3000` 名为 `DOCUMENT_ERROR`，`3999` 从未定义。

**根因**：规范从未声明 `| 3000 | DOCUMENT_ERROR |` 这个二元组里**哪一列是线缆值**（`ErrorResponse.code` 声明为 `string`；注册表表头把数字叫「错误码」；但 `error-handling.md` 全部 9 处示例把名称放进 `code`）。两列都不被声明为规范，于是**两列各自独立漂移**——[#17](https://github.com/A2C-SMCP/oasp-protocol/issues/17)（`RANGE_INVALID` 5002↔3009）是第一次复发，#20 是第二次。本次**只修行级配对**（在两种读法下都正确），线缆列裁决另开 issue 处理。

排查中对全部事件表做了「(编号, 名称) 配对是否精确命中注册表」的机械对账，命中 **5 类 57 行**违规——#20 只点了前两类，故一并扫净。下表按**缺陷面**分组，故处数不可加总（13 行 `3999 OFFICE_API_ERROR` 同时具备「名称非法」与「编号非法」两种缺陷，在前两行**重复计入**；去重后为 57 行）：

| 变更类型 | 位置 | 说明 |
|----------|------|------|
| **移除（名称退役）** | `events-{word,ppt}.md`（39 处） | `OFFICE_API_ERROR` → `3000 DOCUMENT_ERROR`，**不留别名**。该名称点名实现技术（Office API），违反注册表**设计原则 #2「与实现无关」**——不只是未注册，是**不可注册**。注册表不新增任何条目 |
| **移除（编号退役）** | `events-word.md`（13 处，含于上行 39 处内） | `3999` 退役（注册表从未定义、无任何码占用，删除零风险） |
| 订正（编号漂移） | `events-word.md`（10 处） | `4001 VALIDATION_ERROR` → `4000`（9 处）、`4002 MISSING_PARAM` → `4001`（1 处）。名称与描述本就自洽，漂的是编号 |
| 收敛（语义订正） | `events-ppt.md`（8 处） | `3003 OPERATION_FAILED - 元素未找到` 系**三方打架**（编号 `3003` 实为 `DOCUMENT_READ_ONLY`、名称 `OPERATION_FAILED` 实为 `3004`、描述实为 `3010`）。按描述收敛为 `3010 ELEMENT_NOT_FOUND`；复合描述行拆为 `3010`（元素不存在）+ `3004`（元素类型不符）两条，承 #17「`ELEMENT_NOT_FOUND` 为该类收敛靶」之势 |
| 收敛（判法自洽） | `events-ppt.md`（2 处） | 图表事件原把「不存在**或不是图表元素**」揉进单个 `3010`（配对合法，故不在上述 57 行内）。若只拆上一行的 5 处，「元素类型不符」在 `/ppt` 内将分裂为两答案（图表判 `3010`、其余判 `3004`），故一并按同一判法拆分，使 `/ppt` 恢复自洽 |
| 新增（守护） | `scripts/docs/tasks.py`、`.github/workflows/check.yml` | `inv docs.check-error-codes`：解析注册表与规范内**全部五种引用体例**（事件表行、规范层映射表、同格 `` `4001 MISSING_PARAM` ``、`名称 (编号)`、锚点 `#name-code`），断言每处 (编号, 名称) 精确命中注册表。含防脱节保护：解析处数低于下限、或出现「像配对却未被识别」的行时**大声失败**，不以「零违规」放行。接入 PR/push CI，并前置于 `docs.build` / `docs.deploy`（本仓发布走本地 inv，不经 CI）。另附 `--self-test` 自检（合成样本断言各体例「应解析/应失败」，与主逻辑共用同一组正则）。本次修复前报 57、修复后 369 处引用全绿 |

**漂移沿文件年龄边界分布**（可自证）：`events-word.md` 较新段落（2290+，批注类事件）用的是**正确**的 `4000 VALIDATION_ERROR`，较老段落（984–2163）用**过期**的 `4001`；`4001 MISSING_PARAM` 亦在 5 处正确、仅 1 处残留 `4002`。即老段落照着一份**旧注册表**写成，注册表后来重排过编号，事件表没跟。

**兼容性**：**取决于尚未裁决的线缆读法**（见下「遗留」），此处不作无条件承诺——

- 在**「数字上线缆」读法**下（= office-editor4ai 现网实现：`OFFICE_API_ERROR` 已是 `@deprecated` 别名常量、值即 `"3000"`）：**线缆零变更**。13 处标 `3999` 的事件（含 `word:merge:cells`）线上实收 `3000`，本次属**文档追认现实**，`/word`（Stable）不受破坏。
- 在**「名称上线缆」读法**下（= 本规范 9 处示例的字面写法）：`OFFICE_API_ERROR` → `DOCUMENT_ERROR` 且不留别名，对 `/word`（Stable）构成一次**破坏性重命名**，须与线缆裁决 issue 一并处理。

跨仓跟进——office-editor4ai：线缆裁决落地后再定 `OFFICE_API_ERROR` 别名常量的去留；office4ai：视裁决结果确认。

**遗留**（均另开 issue，不在本次范围）：

1. **线缆读法未裁决（根因，高优）**：`error.code` 究竟装「名称」还是「数字」（规范 9 处示例发名称、office-editor4ai 实测发数字字符串、注册表表头又管数字叫「错误码」，三者互斥；另有 `events-word.md` 一处示例为裸数字 `"code": 3001`，违反其自身 `string` 类型）。不裁决则该类漂移仍会复发，须跨 3 仓对齐。
2. ~~**`3012 SEARCH_NO_MATCH` 系孤儿码**~~ —— 已由 [#23](https://github.com/A2C-SMCP/oasp-protocol/issues/23) 裁决，见上「裁决四个 3xxx 孤儿错误码」。

### /excel 错误码收敛回通用注册表（Draft 破坏性收敛）

裁决 [#17](https://github.com/A2C-SMCP/oasp-protocol/issues/17)：`RANGE_INVALID` 在 `error-handling.md`（`3009`）与 `events-excel.md`（`5002`）间编号打架。全局排查发现根因不止于此——OASP 错误码的既定架构是「一张与应用无关的共享注册表（1xxx–4xxx）」（`error-handling.md` 设计原则 #1），`/word`（Stable）与 `/ppt` 连单命名空间语义都放进通用 3xxx 表；**唯 `/excel` 例外**开了 `5xxx` 专属块，且 10 码中 7 码重复了已有通用码。故**退役整个 `5xxx` 块**，逐码收敛回通用注册表，对齐 word/ppt。

| 变更类型 | 位置 | 说明 |
|----------|------|------|
| **移除（Draft 破坏性）** | `events-excel.md` | 删除「Excel 专属错误码」`5001`~`5010` 块，替换为「Excel 错误码映射（规范层）」——一张「触发条件 × 通用错误码 × 涉及事件 × `details`」映射表 |
| 收敛（错误码复用） | `events-excel.md` 全部逐事件「可能的错误」表 | `5001/5006/5007/5008 → 3010 ELEMENT_NOT_FOUND`（由 `details.kind` 区分 worksheet/table/chart/pivotTable）、`5002 → 3009 RANGE_INVALID`、`5003 → 3014 ALREADY_MERGED`、`5004 → 3003 DOCUMENT_READ_ONLY`（`details.scope`）、`5010 → 3016 API_NOT_SUPPORTED` |
| 新增（通用错误码） | `error-handling.md` | `3017 FORMULA_ERROR`、`3018 DATA_TYPE_MISMATCH`（两码无干净通用孪生，按 3011/3013 先例落入通用 3xxx；`5005 → 3017`、`5009 → 3018`） |
| 澄清（描述扩容，非破坏） | `error-handling.md` | `3003`/`3010`/`3014`/`3016` 表述与触发场景去 Word/PPT 专有化，覆盖 Excel 条件（语义不变，既有 word/ppt 用法仍有效） |
| 澄清（规范层归属） | `events-excel.md` / `conventions.md` | 引用 `conventions.md` §规范分层（补 `{#normative-layering}` 锚点）明确 Excel 错误码为**规范层（MUST）**：出现所列线缆可观测条件时实现 **MUST** 返回对应码，不得降级为通用 `3000` |

**Normative 级别（#17 Problem 3）**：无需新增级别——`conventions.md` §规范分层 早已把「错误码 + 实现中立触发条件」列入规范层（MUST）。本次仅把该既有框架显式应用到 Excel 条件并补齐每码的线缆可观测触发条件。

**跨命名空间**：本次为**收敛**（错误码是与应用无关的共享 wire 语义，应复用）——与字体枚举「跨命名空间零映射不复用」（对齐不同宿主 API、故不共享）方向相反、各自适用，不冲突。

**兼容性**：`/excel` 为 Draft（无稳定性承诺），退役 5xxx 属可接受的破坏性收敛；通用码描述扩容对 `/word`（Stable）/`/ppt` 向后兼容。跨仓跟进——office4ai e2e（`manual_tests/excel/test_excel_e2e.py`）按通用码 reassert（`*_NOT_FOUND` 以 `details.kind` 区分）；office-editor4ai（`packages/shared/src/error-codes.ts`）port 通用 3xxx 而非 5xxx。

### `{excel,word,ppt}:run:script` —— 通用 Office.js 脚本执行接口（封装层逃生舱）

裁决 [#18](https://github.com/A2C-SMCP/oasp-protocol/issues/18)：三命名空间各新增一个 `run:script` 事件——调用方下发 JS 源码，AddIn 注入宿主 `RequestContext` 后按 Office.js 语义执行，回传返回值 + 日志。定位是**封装层逃生舱**（typed 事件未覆盖某能力、或某 typed 事件有 Bug 阻塞时的止血路径 + 新能力孵化器），**不替代** typed 事件。

| 变更类型 | 位置 | 说明 |
|----------|------|------|
| 新增（事件 ×3） | `events-{excel,word,ppt}.md` | `excel:run:script` / `word:run:script` / `ppt:run:script`，三端同构，📋 Draft |
| 新增（共享数据结构） | `data-structures.md` | `RunScriptRequest`（`script`/`args?`/`timeoutMs?`）+ `ScriptResult`（`result`/`logs`/`durationMs`/`logsTruncated`）——**宿主无关执行信封**，三命名空间共享（对齐 `BaseRequest`/`BaseResponse`；含「有意跨命名空间相同」admonition） |
| 新增（共享语义节） | `conventions.md` §脚本执行 | 执行语义（AsyncFunction / 注入 `context`·`args`·`console` / 补 `sync` / JSON 序列化）+ 大小限制（`result` ≤ 512KB、`logs` ≤ 100KB，**UTF-8 字节**）+ 超时 + 安全模型 + 非原子性 + 错误映射 |
| 新增（超时档） | `conventions.md` §超时约定 | 新增「脚本执行」档：缺省 60000ms，`timeoutMs` 可覆盖，**无硬上限** |
| 错误码 | —— | **零新增**：复用 `4001/4002/4003/4004`（参数）·`3016`（宿主禁动态代码，`phase:compile`）·`3004`（`fault:script`/`fault:office`）·`1002`（超时）·`3006`（`result` 过大） |

**安全模型（规范正文，MUST 阅读）**：`run:script` 以 `AsyncFunction` 执行、**非沙箱**（沙箱与「直接调原生 Office.js」的目的在定义上互斥），全局对象敞开，信任边界 = Socket.IO 握手鉴权。如实记录**本事件独有的攻击面**（恶意文档提示词注入 → AI 生成含 `fetch` 的脚本 → 外泄，typed 事件不可达）。人在环「执行前确认」由**上层 Agent 层**的通用工具二次确认承担，OASP 为**纯能力提供方**，不自建确认门、不提供安全边界。

**非原子性（规范正文）**：与 typed 事件不同，`run:script` **不保证原子性/回滚**——失败时文档可能处于任意中间状态（跨 sync 已落盘不可回滚 + 单 sync 内部半批执行）。Office.js 无事务 API；需对账的调用方应在脚本内自行 read-back。

**待裁决收敛**（原 issue 6 项 → 3 项由架构判断消除）：确认门移至上层 Agent 层后，「用户拒绝执行」新码（原候选 `3019`）**取消**（拒绝发生在 Agent 层、请求不到线缆）、确认门握手能力位**取消**、「能力被关闭」码收敛为既有 `3016`（仅剩宿主真的跑不了这一种）。动作词取 `run`（对齐 Office.js `Excel.run`/`Word.run`）；错误码取「compile→`4002` / execute→`3004` 且 `details.fault` 区分」而非专用码（承 #17 收敛号段之势）；不引入 `dryRun`。

**兼容性**：纯加法 MINOR。`/excel`·`/ppt` 为 Draft；`/word` 为 Stable 但本事件为**纯加法**（新增事件不改任何既有契约），符合 Stable 变更门槛。跨仓跟进——office4ai：每命名空间一个 MCP 工具（`{ns}_run_script`，ESCAPE HATCH 定位 + 优先 typed 的 description，与既有离线 `office_run_script` 互指）、`{ns}:run:script` 派生 `server_timeout=(timeoutMs??60s)+GRACE`；office-editor4ai：三端 `AsyncFunction` 执行器（宿主无关核心进 `shared`）、`client.ts` 双重 ack 修复、**发版前**补测 Windows/WebView2 + Office-on-web 探针。

## [0.4.0] - 2026-07-13

### /word 字体字段收敛为 WordFont + 修正 UnderlineStyle（含 Stable 破坏性变更）

延续 #10 的字体抽象统一，把 `/word` 散落在 `TextFormat` / `CellFormat` 上的字体字段收敛为统一实体 **`WordFont`**（`bold`/`italic`/`underline`/`size`/`name`/`color`/`highlightColor`，两处复用）。经 `/cross-ask office-editor4ai` 复核可行（全 `Word.Font`、文本场景 WordApi 1.1），并**纠正了现协议 `UnderlineStyle` 的严重错误**。

| 变更类型 | 位置 | 说明 |
|----------|------|------|
| 新增（数据结构） | `data-structures.md` | `WordFont`（7 属性字体子对象，对齐 `Word.Font`、零映射） |
| **修正（错误值）** | `data-structures.md` `UnderlineStyle` | 原 7 值小写（`single`/`double`/…）**三重错误**：大小写全错、`dashed` 为**非法值**（Word 用 `DashLine`）、且为漏项子集。改为对齐 `Word.UnderlineType` 的 **18 值 PascalCase**（`None` + 17 可 SET 值，排除仅回读 `Mixed` 与废弃 `Hidden`/`DotLine`）。已部署 Add-In 本就用 PascalCase 全枚举，本次收敛掉文档与实现的漂移 |
| **移除 + 收敛（Stable 破坏性，clean-break）** | `TextFormat`（`word:insert:text` / `replace:selection` / `replace:text`） | 扁平 `bold`/`italic`/`underline`/`fontSize`/`fontName`/`color`/`highlightColor` **删除**，收敛为 `TextFormat = { font?: WordFont, styleName? }`。经确认无外部已部署发送方，直接 clean-break、不留兼容窗口 |
| **移除 + 收敛（Draft 破坏性）** | `CellFormat`（`word:update:tableCell`） | 扁平 `fontName`/`fontSize`/`fontColor`/`bold`/`italic` **删除**，收敛为 `font?: WordFont`（`fontColor`→`color`）；单元格因此获得 `underline`/`highlightColor`。对齐 / `backgroundColor` 保留顶层 |
| 澄清（既有一致性） | `data-structures.md` | `TextFormat.styleName` 经 `styleBuiltIn`（中文宿主必需，实际 WordApi 1.3）；`highlightColor` 桌面端仅 15 预置色、非自由 RGB、`null` 清除；`CellFormat` 全字段有效门槛 WordApi 1.3 |

**版本门槛**：文本字体 **WordApi 1.1**（几乎恒满足）；表格单元格字体经 `cell.body.font` 抬至 **1.3**（与 `CellFormat` 其余字段同档，不额外抬高）；`styleName` 可靠应用（`styleBuiltIn`）需 **1.3**。降级同 `/ppt` 反应式 `3016`。

**跨命名空间**：`/word` `UnderlineStyle`（18 值，对齐 `Word.UnderlineType`）与 `/ppt` `ShapeFontUnderlineStyle`（17 值，对齐 office.js）字面量不同（虚线 `DashLine` vs `Dash`），**有意不复用**。

**兼容性**：`TextFormat`（Stable）与 `CellFormat`（Draft）均破坏性收敛；因确认无外部消费方，`/word` 亦采 clean-break。工具侧（`JIAQIA/office-editor4ai`）跟进消费——Add-In 现状已是 PascalCase 全枚举，underline 侧零改动，主要改动为扁平→嵌套 schema 归一化。目标 `0.4.0`。

### /ppt 文字格式能力补齐 + 字段收敛（Draft 破坏性收敛）

为 `/ppt` 文字工具补齐一批 office.js 原生支持的常规文字格式能力：run 级局部格式、删除线/双删除线、上标/下标、全大写/小型大写、多下划线样式、项目符号、插入即带字体。**字段结构统一收敛为 `font` 子对象**：`insert:text` / `update:textBox` 原有的扁平 `fontSize`/`fontName`/`color`/`bold`/`italic` **直接移除**（不保留 Deprecated 别名），改由 `font.*` 承载——整框级与 run 级复用同一 `PptFont`。因 `/ppt` 为 Draft，允许此破坏性收敛，避免协议表面长期背负散堆的扁平字段。

**可行性来源**：office4ai#45 提出；经 `/cross-ask office-editor4ai`（Add-In / office.js）逐项复核可落地，并**纠正了 issue 的版本假设**（见下「requirement set 分档」）——无 P0，无需 Server 离线回退。

| 变更类型 | 位置 | 说明 |
|----------|------|------|
| 新增（数据结构） | `data-structures.md` | `PptFont`（整框级 / run 级共用的字体子对象，12 属性含删除线/上下标/大写/下划线）、`ShapeFontUnderlineStyle`（17 值 PascalCase 枚举）、`PptTextRun`（`{start,length,font}` run 级寻址）、`PptParagraphStyle`（`{start,length,bulletFormat}` 段落级寻址）、`BulletFormat`（`visible`/`type`/`style`，**无 `character`**）、`BulletType` 枚举 |
| 变更（事件字段） | `events-ppt.md` `ppt:update:textBox` | `TextBoxUpdates` 新增 `font` / `runs[]` / `paragraphs[]`；应用顺序 `text`→`font`→`runs`→`paragraphs`（内容替换在先，区间寻址对齐替换后最终文本；后者覆盖重叠区间） |
| 变更（事件字段） | `events-ppt.md` `ppt:insert:text` | `TextInsertOptions` 新增 `font`（插入即带字体，语义等价「插入 + 格式」复合操作） |
| 变更（事件字段） | `events-ppt.md` `ppt:insert:shape` | `ShapeInsertOptions` 新增 `font`（插入即带字体，复用 `PptFont`，字段 / 语义与 `insert:text` 一致）；`font` / `text` 仅适用 text-capable 形状（`TextBox` + 除 `Line` 外的几何形状），用于无文本框的 `Line` → 前置 `4002 INVALID_PARAM`（语义误用，与能力不足 `3016` 分离）；`text` 与 `font` 并存施加顺序 `text`→`font` |
| **移除（Draft 破坏性收敛）** | `events-ppt.md` | `TextBoxUpdates` 与 `TextInsertOptions` 的扁平 `fontSize`/`fontName`/`color`/`bold`/`italic` **删除**，统一改用 `font.*`；`/ppt` 为 Draft，不保留 Deprecated 别名 |
| 复用（无新错误码） | `events-ppt.md` / `error-handling.md` | 区间越界 / 枚举非法复用 [`4002 INVALID_PARAM`](../specification/error-handling.md)；requirement set 不满足复用 [`3016 API_NOT_SUPPORTED`](../specification/error-handling.md#api_not_supported-3016)（`details.requiredApiSet` 标注版本） |

**requirement set 分档（经 Add-In 真机复核，纠正 issue 假设）**：

| 能力 | 最低 PowerPointApi | 备注 |
|------|-------------------|------|
| 下划线全枚举、run 级寻址（`getSubstring`）、run 级 bold/italic/size/name/color、bullet `visible` | **1.4** | issue 曾误标下划线/大写/删除线为 1.4、getSubstring 为 1.5 |
| 删除线 / 双删除线、上标 / 下标、全大写 / 小型大写（含其 run 级） | **1.8** | issue 误标为 1.4；整批有效门槛 = 1.8 |
| bullet `type` / `style`（编号列表） | **1.10** | `visible` 仅需 1.4 |

**降级语义**：宿主不满足**本次请求所含属性**的最高 requirement set 时，事件按**反应式 `3016` 整体失败（全或无）**，与现有 `/ppt` 事件（`insert:chart` / `slidesOoxml`）一致；调用方靠「只发受支持的属性」控制粒度。不引入「部分成功」响应语义。

**字符区间口径**：`runs` / `paragraphs` 的 `start` / `length` 以 **UTF-16 code unit** 计，含段落分隔 `\r` 与软换行 `\v`；与 `updates.text` 并存时**先替换内容、offset 对齐替换后的最终文本**；越界 → `4002`。

**跨命名空间说明**：PPT `ShapeFontUnderlineStyle`（17 值 PascalCase，对齐 office.js PowerPoint 枚举、零映射）与 `/word` `UnderlineStyle`（7 值小写、映射 Word.js）**有意不同**，服务不同宿主 API，不复用。

**枚举承载分档**：宿主枚举按**集合大小与稳定性**择一承载——小而封闭者就地全枚举（`ShapeFontUnderlineStyle` 17 值，可校验）；大而易变者直通 `string`（bullet `style` 40+ 值，由宿主校验、非法值 → `4002`）。二者同为"零映射对齐宿主"，形态分档是一致原则而非遗漏。

**兼容性**：新增可选字段属向后兼容；但**移除扁平 `fontSize`/`fontName`/`color`/`bold`/`italic` 对 `/ppt` 是破坏性变更**——调用方须改用 `font.*`。因 `/ppt` 为 Draft（无稳定性承诺），此收敛可接受。已部署 AddIn 忽略未知字段即可。属 v0.x 阶段变更（目标 `0.4.0`）。

**范围边界（不含）**：段落级排版（行距 / 缩进 / 段间距）、字间距、超链接写入、竖排——office.js PowerPoint API 无原生支持，走 OOXML 路径另案；bullet 自定义字符 / 字体 / 颜色——office.js 不提供，协议不设字段。Add-In 侧既有实现缺口（`insert:text` 字体哑参数 / `update:textBox` `fillColor` 未透传 / underline 退化实现）由 office-editor4ai 单独立项修复，非协议变更。

### /ppt 表格单元格字体收敛（`ppt:update:tableFormat`，Draft 破坏性收敛）

延续上条的字体抽象统一：`ppt:update:tableFormat` 的 `cellFormats` / `rowFormats` / `columnFormats` 原本平铺 `fontSize`/`fontColor`/`bold`/`italic`，属同类散堆，现**收敛复用同一 [`PptFont`](../specification/data-structures.md#pptfont)**（12 属性全量可用，含 17 值下划线 / 删除线 / 上下标 / 大写），使全 `/ppt` 字体抽象一致。

**可行性来源**：经 `/cross-ask office-editor4ai` 复核——`TableCell.font` 即标准 `ShapeFont`，`PptFont` 12 属性零丢失、无需另立精简结构。

| 变更类型 | 位置 | 说明 |
|----------|------|------|
| **移除 + 收敛（Draft 破坏性）** | `ppt:update:tableFormat` | 三级 formats 的扁平 `fontSize`/`fontColor`/`bold`/`italic` **删除**，统一改用 `font?: PptFont`（`fontColor`→`color` 语义一致、安全改名） |
| 新增（行/列级字体，方案 A） | `rowFormats[]` / `columnFormats[]` | 由「仅 `fontSize`」升级为完整 `font?: PptFont`，消除「行/列只能设字号」的割裂；`height`/`width` 保留为行/列 native |
| 新增（数据结构） | `data-structures.md` | `ParagraphHorizontalAlignment`（7 值）、`TextVerticalAlignment`（6 值）枚举，替换原 `string` 并放开到 office.js 全枚举 |
| 复用（新增错误码引用） | `ppt:update:tableFormat` | 门槛不满足 → [`3016`](../specification/error-handling.md#api_not_supported-3016)；越界 / 合并单元格非左上格 / 枚举非法 → [`4002`](../specification/error-handling.md) |

**版本门槛（与文本框不同，单列）**：表格单元格字体 / 对齐 / 行高列宽**统一 1.9**——因 `TableCell.font` 访问器门槛为 1.9，把复用的 `PptFont` 底层 1.4/1.8 属性整体抬平；降级判定为单一 `isSetSupported("PowerPointApi","1.9")`，不按属性取 max。现工具已依赖 `cell.font`（1.9），扩到 12 属性**零额外版本成本**。

**单元格无 run 级**：office.js 未暴露单元格内字符区间寻址，`font` 作用于整格，不支持 `runs`。

**行/列级即逐格扇出**：office.js 无「整行/整列字体」原生 API，`rowFormats`/`columnFormats` 的 `font`/`backgroundColor` 均为逐格施加到该行/列每个单元格的糖；命中同一格优先级 `cellFormats` > `columnFormats` > `rowFormats`（沿用既有 `backgroundColor` 口径）。

**全或无（含合并单元格）**：应用前前置校验所有目标格（含行/列展开），任一越界或命中合并单元格非左上格（`getCellOrNullObject` 空对象）→ 整请求 `4002` 失败、不写入任何格；与文本框「全或无」口径一致（纠正现工具「跳过非法格」的部分成功语义，由 office-editor4ai 跟进实现）。

**兼容性**：移除三级 formats 的扁平字体字段对 `/ppt` 为破坏性变更；因 `/ppt` 为 Draft 可接受。目标 `0.4.0`。

---

## [0.3.0] - 2026-05-26

### 架构转变：Python MCP Server 从「纯中转」升级为「具备生产能力」

0.3.0 是 OASP 的一次**模型转变**，也是本次 MINOR 升级（0.2 → 0.3）的根本原因——不是单纯加事件：

- **此前（≤0.2.x）**：Python MCP Server 仅做消息中转，**所有动作都在 AddIn（Office.js）实现**；AddIn 未连接即无法操作文档。
- **从 0.3.0 起**：Server **也具备生产能力**——当 AddIn 未打开、或 AddIn 能力无法实现某操作时，Server 直接介入生产（借助 `python-pptx` 等 OOXML 工具）；**未来 Server 将在 AddIn 完全未连接的情况下实现全部工具操作**。

本版所有变更都服务于这一转变、互为支撑：

- **剥离实现技术 / Normative-Informative 分层** → 让同一套协议可被 Server 或 AddIn **任一端**满足，是「双端皆可生产」的地基；
- **图表双路径线缆可观测项（`3016` / `elementId` 不透明性 / 客户端路径 Informative）** → Server-生产 与 AddIn-生产 之间的**路由**与反应式降级；
- **通用幻灯片 OOXML 搬运事件（`ppt:get:slideOoxml` / `ppt:insert:slidesOoxml`）** → AddIn 开着时也能把整页 OOXML 交给 Server 生产再回插；
- **协议版本握手** → 既然**两端都生产**，Server ↔ AddIn 的版本错位代价更高，故在连接期强制校验版本。

### 新增（协议版本握手）

为避免新旧 Server / AddIn 协议错位，新增连接握手阶段的协议版本校验机制。AddIn 在 `auth` 中声明 `oaspVersion`，Server 在 `connect` handler 校验兼容性（**版本先于业务参数**），不兼容即拒绝连接。

| 变更类型 | 位置 | 说明 |
|----------|------|------|
| 新增（握手参数） | `connection.md` | `auth.oaspVersion`（必填，SemVer）；新增「协议版本握手」节：校验时机/顺序（**oaspVersion 先行**）、`HandshakeRejection` 扁平拒绝结构、Server 校验示例、AddIn 收到拒绝后 MUST 主动断开 |
| 新增（连接确认字段） | `connection.md` | `connection:established` 增加 `serverVersion` 字段（仅诊断用） |
| 新增（版本治理） | `conventions.md` | 「版本兼容」升级为「协议版本与兼容性」：MAJOR/MINOR/PATCH 语义与触发条件、`is_compatible` 判定规则（v0.x 严格 MAJOR.MINOR；v1.0+ Server 向后兼容）、Python 参考实现、载体选型理由、**非目标**与「**版本握手 ≠ 运行时能力**」正交说明 |
| 新增错误码 | `error-handling.md` | `2006 PROTOCOL_VERSION_MISMATCH`（连接/认证段）；复用 `2003 HANDSHAKE_FAILED` 处理缺失/非法 `oaspVersion` |

**兼容性**: 引入新的**必填**握手参数 `oaspVersion`——采用本机制的 Server 会拒绝不声明版本的旧 AddIn，需 Server 与 AddIn **协同升级**。属 v0.x 阶段的 MINOR 级变更（v0.x 阶段 MINOR 可含破坏性）。

**设计来源**: 对齐 A2C-SMCP `versioning.md` 的版本握手规范，按 OASP 两方单连接模型简化——校验放 `connect` handler（非 HTTP 中间件），载体放 `auth`（非 URL query），拒绝走 `ConnectionRefusedError`（非 HTTP 400）。

### 变更（规范澄清，非破坏性）

**将实现技术规定从规范层剥离，确立 Normative / Informative 分层治理原则**

线缆契约（事件名、请求/响应形状、`ChartData` / `ChartType` / `CategoricalSeries` / `ScatterSeries` 判别联合、错误码）**全部保持不变**；本次仅调整规范文本，使 `/ppt` 图表事件可被服务端或客户端任意路径实现。

| 变更类型 | 位置 | 说明 |
|----------|------|------|
| 移除（实现泄漏） | `events-ppt.md` chart 事件顶部 admonition | 删除「不经 Add-In Office.js 路径」「由 Server 端使用 OOXML 工具（如 `python-pptx`）修改 .pptx」等执行位置/技术选型表述 |
| 重构 | `events-ppt.md` chart 事件 | 拆分为两块：① 实现中立的**线缆契约**（请求-响应语义、`success` 后变更对 `ppt:get:chart` 可见、并发顺序约定、显式声明"协议不规定由哪一端/何种技术实现"）；② `!!! info "实现提示（非规范）"` 收纳仅服务端离线路径相关的副作用（延迟 >1s、调用前 `save()`、完成后重载文档、离线写入串行化） |
| 中立化 | `events-ppt.md` / `error-handling.md` 错误码触发条件 | `3004 OPERATION_FAILED` 描述由「OOXML 写入失败」改为「图表写入失败」；`3003 DOCUMENT_READ_ONLY` 由「文档为只读模式」改为「目标文档不可写入」，描述线缆可观测状态而非特定实现 |
| 移除（实现泄漏） | `events-ppt.md` 事件列表表格 | 三个 chart 事件行去掉「（Server OOXML 离线处理）」括注 |
| 新增（治理原则） | `conventions.md` | 新增「规范分层（Normative / Informative）」章节：规范层只含线缆可观测内容（事件名/方向、字段形状与语义、错误码及实现中立触发条件、顺序/幂等/可见性保证）；用什么库、服务端还是客户端、性能特征、特定实现副作用一律归为非规范实现提示 |
| 新增（设计原则） | `error-handling.md` | 错误码设计原则补充「与实现无关」：触发条件描述线缆可观测状态而非实现技术/执行位置 |

**动机**: 消费方 office4ai（OF4AI-21）需按文档连接状态在服务端 OOXML 路径与客户端 Office.js 路径之间路由——同一套事件/payload 两种方式均可实现，接口形状无需改动。原 v0.2.0 文本把「用 `python-pptx` / 服务端离线 / 不走 Office.js」写进规范，等于协议层替消费方做了实现决策，阻碍多路径复用。

**兼容性**: 纯文档/规范澄清，无任何线缆契约变更；对已部署消费方零影响。`/ppt` 整体仍处 📋 Draft 阶段。

**相关工单**: [OF4AI-21](https://turingfocus.atlassian.net/browse/OF4AI-21)（PPT 图表双路径路由）

### 新增（图表双路径线缆可观测项）

在上述减法基础上，为支撑 OF4AI-21 双路径路由追加少量线缆可观测项；事件名、请求/响应形状、`ChartData` 判别联合、现有错误码仍不变。

| 变更类型 | 位置 | 说明 |
|----------|------|------|
| 新增错误码 | `error-handling.md` + `ppt:insert:chart` / `ppt:get:chart` / `ppt:update:chart` 错误码表 | `3016 API_NOT_SUPPORTED`：目标操作在当前客户端/平台不可用（如所需 PowerPointApi requirement set 不满足、宿主不支持）。语义实现中立，供调用方**反应式降级**到另一路径或提示用户 |
| 新增（Normative） | `data-structures.md` 新增「元素标识符不透明性」（锚点 `#element-id-opacity`） | 明确 `SlideElement.id` / chart `elementId` 为服务端分配的**不透明字符串**，消费方不得解析其结构。防御整页 round-trip 重置 native id 的实现细节泄漏到协议 |
| 新增（Informative） | `events-ppt.md` chart 事件 `!!! info` 注记 | 补「路径 B — 客户端 Office.js 整页 round-trip」实现提示（`exportAsBase64` → 服务端改 → `insertSlidesFromBase64` 整页回插），并列其非规范副作用：选区/滚动重置、撤销非原子、母版累积、requirement set 门槛（新页 1.2 / 现有页·get·update 1.8，皆不满足回退路径 A） |

**确认不新增**: `CHART_DATA_NOT_READABLE`——`exportAsBase64` + 服务端解析可读回任意图表数据（含未保存态、含非本系统生成图表），仅平台 <1.8 读不到，已归入 `3016 API_NOT_SUPPORTED`。

**能力声明**: 本轮采用**反应式**路由（直接尝试 → 失败返回 `3016` / `3003` 即保证正确性），不新增事件。主动式 `ppt:capabilities`（握手声明 `isSetSupported` 矩阵）作为后续优化单独立项，不阻塞本轮。

### 新增（通用幻灯片 OOXML 搬运事件）

为支撑 OF4AI-21 双路径的路径 B（客户端整页 round-trip），`/ppt`「幻灯片管理类」新增 2 个 **Server→AddIn 请求-响应**事件（📋 Draft）。图表无关的低层传输原语，未来"服务端编辑打开中文档"类能力可复用。纯加法、向后兼容。

| 变更类型 | 事件 / 类型 | 说明 |
|----------|-------------|------|
| 新增 | `ppt:get:slideOoxml` | 导出指定页当前 OOXML（含未保存态）为单页 `.pptx` base64；响应含不透明 `slideId` |
| 新增 | `ppt:insert:slidesOoxml` | 应用 OOXML 页包；可选 `replaceSlideId` + `finalSlideIndex` 把「插入→删旧→复位」做成**尽力顺序复合（非原子）round-trip**；响应可回报命名元素（如图表）的最终几何 |
| 扩展 | `data-structures.md` 标识符不透明性 | `#element-id-opacity` 节从"元素标识符"扩展到涵盖 `slideId` |
| 扩展 | 错误码 `3010 ELEMENT_NOT_FOUND` | 描述扩展为「元素或幻灯片未找到」，供 `replaceSlideId` 复用，不新增错误码 |
| 文档基建 | `mkdocs.yml` | 启用 `attr_list`（支撑 `#element-id-opacity` 等锚点；顺带修复既有 `#slideelementelementextensions` 失效链接） |

**命名**: 取 `get:slideOoxml` / `insert:slidesOoxml` 中立对称对——按内容格式（OOXML）命名，不把 `base64` 传输编码或 Office.js 方法名写进规范事件名（遵循本周期确立的事件名实现中立原则）。

**消费方验证（office-editor4ai cross-ask，无 P0）**: 据 Add-In 实现可行性反馈做 round-N 修订——
- `insert:slidesOoxml` 措辞由"原子"放宽为**"尽力顺序复合（非原子，无回滚保证）"**：Office.js 无事务/回滚；部分失败时 `error.details` 给出 `{ stage, partiallyApplied, createdSlideId }` 供服务端对账补偿。
- 命名元素回报（`elements[]`）的定位机制下沉为 Informative：**主路径** `cNvPr/@name` + `shape.name`（`@name` 穿越 round-trip 的存活性已由 [office-editor4ai#34](https://github.com/JIAQIA/office-editor4ai/issues/34) 实测确认——Mac/单次：存活、几何可读、`masterLeak: 0`），`customXmlParts` 注册表降为**防御后备**；线缆契约不依赖具体机制。另记：图表可能位于占位符内（`type==="Placeholder"` + `containedType==="Chart"`），按不透明 `elementId` 定位不受影响，但勿用 `type==="Chart"` 过滤。
- `formatting` 保持 camelCase（与其它 ppt 事件一致），Add-In 内部映射到 Office.js PascalCase 枚举；`finalSlideIndex` 任意复位需 1.8、原位替换 1.2 即可；`3016` 经 `isSetSupported` 预检主动返回。

> Add-In spike（[office-editor4ai#34](https://github.com/JIAQIA/office-editor4ai/issues/34)）已确认 `@name` 穿越 round-trip 存活（Mac / 单次 round-trip：存活、几何可读、无母版累积），**主路径成立**；`customXmlParts` 注册表作防御后备。边界：Web/Windows 与多次连续 round-trip 的母版累积建议后续各补一次抽测。

---

## [0.2.0] - 2026-04-30

### 新增

**`/ppt` 命名空间补齐图表（Chart）能力（3 个 Draft 事件 + 1 个跨命名空间数据结构组）**

| 变更类型 | 事件 / 类型 | 说明 |
|----------|-------------|------|
| 新增 | `ppt:insert:chart` | 📋 Draft，按 `ChartData` discriminated union 在指定幻灯片插入图表 |
| 新增 | `ppt:get:chart` | 📋 Draft，按 elementId 读取图表完整数据，响应 `chart` 字段为同一 union |
| 新增 | `ppt:update:chart` | 📋 Draft，按 elementId 更新图表数据/类型/标题/展示选项；`chart.chartType` 为必需 discriminator |
| 提升为跨命名空间 | `ChartType` | 从 `data-structures.md` 的 `Excel 相关` 段移至新建的 `图表相关` 段，并扩展为 10 个 Office.js `Excel.ChartType` 对齐值（新增 `ColumnClustered` / `ColumnStacked` / `BarClustered` / `LineMarkers` / `Radar`，原 `Column` / `Bar` 高层值移除以避免歧义）。`ChartType` 进一步细分为 `CategoricalChartType`（9 个）+ `ScatterChartType`（"Scatter"）。破坏性变更，但 Excel 命名空间整体仍处 📋 Draft 阶段，影响面可控 |
| 新增 | `CategoricalSeries` / `ScatterSeries` / `ScatterPoint` | 数据结构，分别承载分类型图表的数据系列（name / values / color）与散点型图表的数据系列（name / points / color）。取代原始 `ChartSeries` 单一定义 |
| 新增 | `ChartData` discriminated union | `ChartData = CategoricalChartData \| ScatterChartData`，由 `chartType` 字段判别。Categorical 含 `categories: string[]` + `series: CategoricalSeries[]`；Scatter 不含 `categories`，X 由 `series[].points[].x` 提供。供 PPT / Excel 共用 |
| 新增 | 错误码 `3015 INVALID_CHART_DATA` | 多用途：categorical 维度不一致 / scatter `points` 为空或含非法值 / 跨 variant 切换未补齐 `series` |
| 新增（说明） | `ppt:delete:element` 描述 | 追加 `亦适用于图表（Chart）元素删除` admonition；不新增独立 `ppt:delete:chart`，复用通用删除入口 |

**动机**: PPT 是图表高频使用场景（业绩报告、产品方案）。当前 OASP `/ppt` 命名空间已能插入文本框、形状、图片、表格，**唯独不能插入图表**——`Chart` 仅在 `SlideElementType` 中作为枚举出现并可被 `ppt:get:slideElements.includeCharts` 过滤，但没有任何对图表本身的增改查事件。AI 体验上只能让用户手动插入图表后再让 AI 改文字，断裂明显。

**设计取舍**:

- **Server-handled OOXML 路径**：PowerPoint Office.js 当前不暴露图表创建与数据更新接口（参见 [office-js#5463](https://github.com/OfficeDev/office-js/issues/5463)），本批 chart 类事件由 OASP Server 端通过 OOXML 离线工具（如 `python-pptx`）直接修改 .pptx 文件后再让 Add-In 重新加载文档。事件文档顶部 admonition 明确标注延迟 >1s、要求 Add-In 调用前 `save()`、并发需串行化等副作用约束
- **`ChartData` 采用 discriminated union（`chartType` 为 discriminator）**：`Scatter` 与分类型图表的数据形状本质不同（X 是连续数值 vs 离散标签）。如果硬塞同一 schema（让 `categories` 在 Scatter 时存数字字符串），LLM 在 MCP 工具的 JSON Schema 里看不到这个隐式约束，调用错误率高。Discriminated union 让 LLM 一选定 `chartType`，schema 就限定可填字段；JSON Schema 用 `oneOf` + `discriminator: { propertyName: "chartType" }` 落地，Pydantic / FastMCP 用 `Annotated[Union[...], Field(discriminator="chartType")]`
- **`update:chart` 强制要求 `chartType`**：让 schema 层能选中正确 variant；AI 调用前先 `get:chart` 读取并回填，比让 schema 层放任所有字段都 optional 更可靠
- **复用 `ppt:delete:element`**：删除统一走通用入口，避免新增 `ppt:delete:chart` 造成 API 表面冗余
- **`ChartType` 与 Office.js 对齐**：枚举值与 `Excel.ChartType` 完全一致（如 `ColumnClustered` / `LineMarkers`），消费方可直接 `cast`，无需维护额外映射表
- **`ChartType` 跨命名空间提升**：未来 Excel `excel:insert:chart` 可直接复用同一枚举与 `ChartData` 联合，避免双套语义
- **不新增 `update:element` 字段塞图表**：`ppt:update:element` 仅承担几何属性，图表的「数据/类型/标题」是图表特有语义，独立 `update:chart` 边界更清晰、错误恢复粒度更细

**兼容性**:

- 新增事件 + 新增可选字段，不影响 `/word` 等其它命名空间
- `ChartType` 枚举值变更属破坏性变更，但 Excel 侧 `excel:insert:chart` 尚未定义、PPT 侧本就无 chart 事件，实际无在用消费方
- discriminated union 形态在本 [Unreleased] 周期内确定（含中途 PR #4 → PR #5 的 schema 重构），无对外发版的旧 shape 残留
- 新事件初始标记 📋 Draft，待消费方（office4ai）落地稳定后再统一转 ✅ Stable

**相关 Issue / 工单**:

- 协议侧：[oasp-protocol#2](https://github.com/A2C-SMCP/oasp-protocol/issues/2)（Milestone: PPT Chart Capabilities）
- 协议 PR：[#4](https://github.com/A2C-SMCP/oasp-protocol/pull/4)（首版 categorical-only schema）→ [#5](https://github.com/A2C-SMCP/oasp-protocol/pull/5)（重构为 discriminated union，提升 LLM 可发现性）
- 消费方实现：[office4ai#9](https://github.com/JIAQIA/office4ai/issues/9)（Server 端 python-pptx OOXML 拦截）

---

**`/word` 命名空间补齐表格操作能力（4 个 Draft 事件）**

| 变更类型 | 事件 / 类型 | 说明 |
|----------|-------------|------|
| 新增 | `word:merge:cells` | 📋 Draft，合并表格中任意矩形单元格区域 |
| 新增 | `word:update:tableCell` | 📋 Draft，批量更新单元格文本与格式 |
| 新增 | `word:update:tableRowColumn` | 📋 Draft，按行/列批量更新单元格内容 |
| 新增 | `word:update:tableFormat` | 📋 Draft，更新整表样式、边框、列宽、对齐 |
| 新增 | `CellFormat` 数据结构 | 定义于 `data-structures.md`，承载单元格对齐/字体/底色等格式属性，供 `word:update:tableCell` 等事件复用 |
| 新增 | 错误码 `3013 NO_TABLE_AT_CURSOR` / `3014 ALREADY_MERGED` | 表格类事件专用错误码，便于 AI 区分"表格 ID 错"、"光标位置错"、"已合并冲突"三类失败 |
| 新增（Stable 事件向后兼容扩展） | `word:get:documentStructure` 增加可选 `tables: TableSummary[]` | 用于"重新发现"现有表格，配合 `precedingHeading` 启发式定位；旧客户端忽略该字段不受影响 |

**动机**: Add-In 内部 `word-tools/table.ts` 已实现合并单元格、单元格内容/格式更新、整表样式/边框/列宽更新等能力，协议层此前仅暴露 `word:insert:table`，导致 AI 无法调用细粒度表格操作（如生成"蓝底居中表头 + 灰底加粗标签"等常见合同/报告美化效果，或合并表头横跨多列）。本批新增事件按"协议先行"原则补齐该能力面，统一以 `tableId`（与 `word:insert:table` 响应一致）作为定位标识；缺省 `tableId` 时取当前光标所在表格。

**与 Office.js 对齐**（消费方 office-editor4ai 验证后修订）:

- `CellFormat.horizontalAlignment` / `verticalAlignment` 枚举值与 `Word.Alignment` / `Word.VerticalAlignment` 完全一致——`"Centered"` / `"Justified"` 用过去分词形，`"Center"` 表示垂直居中（不再是 `"Middle"`），Add-In 可直接 `cast` 不做映射
- `CellFormat.cellPadding` **未纳入**——Word.js 单元格内边距是表级 API（`Word.Table.setCellPadding`），无法逐单元格设置；改放到 `word:update:tableFormat.styleOptions.cellPadding: { top, bottom, left, right }`
- `word:merge:cells` 响应字段从 `mergedCells: number` 改为 `requestedRange: { rowCount, columnCount }`——Word.js 不暴露"被合并的原子单元格总数"，乘法计算在已部分合并区域会失真，返回请求矩形的尺寸语义更清晰
- `word:update:tableFormat.borderOptions.location?: "all" | "inside" | "outside"`（默认 `"all"`）——覆盖"内细外粗"高频场景，对应 `Word.BorderLocation`
- `word:update:tableFormat` **移除 `data` 字段**——避免与 `word:update:tableRowColumn` 职责重叠，强制写数据与改样式分两步调用，错误恢复粒度更细
- `word:update:tableFormat.styleOptions.styleType` 不存在的样式名错误码使用 `3011 STYLE_NOT_FOUND`（已存在），非 `OPERATION_FAILED`
- `word:update:tableFormat.columnWidths` 长度策略放宽为"≤ 列数"——过短只覆盖前缀列，避免 AI 边缘列少传一个就触发失败
- 全部表格事件错误码修正：`3010 ELEMENT_NOT_FOUND`（表格未找到）+ `3013 NO_TABLE_AT_CURSOR`（缺省 tableId + 光标不在表格内）替代此前误用的 `3003 OPERATION_FAILED`，与现有 `error-handling.md` 错误码定义对齐

**新增警示**：

- `word:insert:table` 响应中 `tableId` 当前为**临时索引**，跨会话或经过结构变更的场景调用方应通过 `word:get:documentStructure` 重新发现表格（响应文档新增 `tableId 稳定性` 警告）。基于 Content Control 的稳定 ID 方案为单独工单跟进，不阻塞本批 Draft 事件转 Stable

**兼容性**: 全部为新增事件 + 新增可选字段，不影响现有 Stable 事件；`/word` 命名空间稳定性整体保持。新事件初始标记 📋 Draft，待消费方双端落地稳定后再统一转 ✅ Stable。

**相关 Issue / 工单**:

- 协议侧：[oasp-protocol#1](https://github.com/A2C-SMCP/oasp-protocol/issues/1)（Milestone: Word Table Capabilities）
- 消费方：[OF4AI-10](https://turingfocus.atlassian.net/browse/OF4AI-10)（mergeCells）、[OF4AI-11](https://turingfocus.atlassian.net/browse/OF4AI-11)（updateCell / updateTable）

---

## [0.1.9] - 2026-04-17

### 新增

**word:insert:table 新增 `insertLocation` 字段**

| 变更类型 | 字段 | 说明 |
|----------|------|------|
| 新增 | `TableInsertOptions.insertLocation` | 可选枚举字段，取值 `Start` / `End` / `Before` / `After` / `Replace`，未传时默认 `"End"`（向后兼容） |
| 新增 | `TableInsertLocation` 类型定义 | 表格插入位置枚举类型 |

**动机**: Add-In 内部 `insertTable` 函数已支持五种插入位置，但此前协议层仅定义 `rows` / `columns` / `data` / `style`，导致 AI 无法指定表格插入位置，所有表格默认插入文档末尾。此次变更为协议层能力补齐，不影响现有客户端。

**相关工单**: [OF4AI-9](https://turingfocus.atlassian.net/browse/OF4AI-9)

---

## [0.1.8] - 2026-02-05

### 变更

**word:get:styles 响应结构调整**

基于 Add-In 实际实现进行协议规范对齐：

| 变更类型 | 字段 | 说明 |
|----------|------|------|
| 修改 | `timestamp` (请求) | 从必需改为可选 |
| 删除 | `duration` (响应) | Add-In 未实现，移除 |
| 修改 | `success` (响应) | 从字面量 `true` 改为 `boolean` |
| 新增 | `error` (响应) | 添加可选错误响应字段 |

---

## [0.1.7] - 2026-02-05

### 变更

**word:get:documentStats 请求与响应结构调整**

基于 Add-In 实际实现进行协议规范对齐：

| 变更类型 | 字段 | 说明 |
|----------|------|------|
| 新增 | 请求结构 | 补充完整的请求定义（原协议缺失） |
| 修改 | `timestamp` (请求) | 从必需改为可选 |
| 删除 | `duration` (响应) | Add-In 未实现，移除 |
| 修改 | `success` (响应) | 从字面量 `true` 改为 `boolean` |
| 新增 | `error` (响应) | 添加可选错误响应字段 |
| 修改 | `characterCount` | 明确语义为"不含空格" |
| 新增 | `characterCountWithSpaces` | 含空格的字符数 |
| 新增 | `pageCount` | 页数（可选） |
| 重命名 | `DocumentStats` → `DocumentStatsResult` | 统一命名风格 |

---

## [0.1.6] - 2026-02-05

### 变更

**word:get:documentStructure 请求与响应结构调整**

基于 Add-In 实际实现进行协议规范对齐：

| 变更类型 | 字段 | 说明 |
|----------|------|------|
| 修改 | `timestamp` (请求) | 从必需改为可选 |
| 删除 | `duration` (响应) | Add-In 未实现，移除 |
| 修改 | `success` (响应) | 从字面量 `true` 改为 `boolean` |
| 新增 | `error` (响应) | 添加可选错误响应字段 |
| 修改 | `data` 字段顺序 | 调整为 `sectionCount → paragraphCount → tableCount → imageCount` |

---

## [0.1.5] - 2026-02-05

### 变更

**word:get:visibleContent 请求与响应结构调整**

基于 Add-In 实际实现进行协议规范对齐：

| 变更类型 | 字段 | 说明 |
|----------|------|------|
| 修改 | `timestamp` (请求) | 从必需改为可选 |
| 新增 | `options` (请求) | 支持 GetContentOptions（与 selectedContent 一致） |
| 删除 | `duration` (响应) | Add-In 未实现，移除 |
| 删除 | `data.startPosition` | Add-In 未实现，移除 |
| 删除 | `data.endPosition` | Add-In 未实现，移除 |
| 新增 | `data.elements` | 内容元素数组，带类型映射 |
| 新增 | `data.metadata` | 统计元数据 |

**新增 VisibleContentElement 结构**：包含 `type`（映射后类型）和 `content`（原始元素）。

---

## [0.1.4] - 2026-02-05

### 变更

**word:get:selectedContent 请求与响应结构调整**

基于 Add-In 实际实现进行协议规范对齐：

| 变更类型 | 字段 | 说明 |
|----------|------|------|
| 修改 | `timestamp` (请求) | 从必需改为可选 |
| 重构 | `options` | 从 `includeFormatting` 改为 6 个独立选项 |
| 删除 | `duration` (响应) | Add-In 未实现，移除 |
| 删除 | `data.html` | Add-In 未实现，移除 |
| 删除 | `data.format` | Add-In 未实现，移除 |
| 新增 | `data.elements` | 内容元素数组（段落、表格、图片、内容控件） |
| 新增 | `data.metadata` | 统计元数据（字符数、段落数等） |

**新增 GetContentOptions**：`includeText`、`includeImages`、`includeTables`、`includeContentControls`、`detailedMetadata`、`maxTextLength`

**新增内容元素类型**：`ParagraphElement`、`TableElement`、`InlinePictureElement`、`ContentControlElement`

---

## [0.1.3] - 2026-02-05

### 变更

**word:get:selection 请求与响应结构调整**

基于 Add-In 实际实现进行协议规范对齐：

| 变更类型 | 字段 | 说明 |
|----------|------|------|
| 修改 | `timestamp` (请求) | 从必需改为可选 |
| 删除 | `duration` (响应) | Add-In 未实现，移除 |
| 完善 | `data.type` | 补充完整枚举值：`NoSelection`、`InsertionPoint`、`Normal` |
| 修改 | `data.start` | 从必需改为可选（仅选区非空时存在） |
| 修改 | `data.end` | 从必需改为可选（仅选区非空时存在） |
| 修改 | `data.text` | 从必需改为可选（仅选区非空时存在） |

**新增响应示例**：覆盖"有选区"、"光标点"、"无选区"三种场景。

---

## [0.1.2] - 2026-02-05

### 变更

**word:event:selectionChanged 事件结构调整**

基于 Add-In 实际实现进行协议规范对齐：

| 变更类型 | 字段 | 说明 |
|----------|------|------|
| 新增 | `eventType` | 事件类型标识，固定值 `"selectionChanged"` |
| 新增 | `clientId` | 发送事件的客户端标识 |
| 修改 | `selection` → `data` | 字段名变更，简化结构 |
| 删除 | `selection.isEmpty` | 实现未提供 |
| 删除 | `selection.type` | 实现未提供 |
| 删除 | `selection.start` | 实现未提供 |
| 删除 | `selection.end` | 实现未提供 |
| 新增 | `data.text` | 选中的文本内容 |
| 新增 | `data.length` | 选中文本的字符长度 |

**word:event:documentModified 事件结构调整**

基于 Add-In 实际实现进行协议规范对齐：

| 变更类型 | 字段 | 说明 |
|----------|------|------|
| 新增 | `eventType` | 事件类型标识，固定值 `"documentModified"` |
| 新增 | `clientId` | 发送事件的客户端标识 |
| 新增 | `data` | 事件数据对象 |
| 新增 | `data.modificationType` | 修改类型：`insert`、`delete`、`update` |

**word:get:styles 请求参数调整**

基于 Add-In 实际实现进行协议规范对齐：

| 变更类型 | 参数 | 说明 |
|----------|------|------|
| 保留 | `includeBuiltIn` | 无变化 |
| 保留 | `includeCustom` | 无变化 |
| 删除 | `typeFilter` | Add-In 未实现，移除以保持一致性 |
| 新增 | `includeUnused` | 控制是否返回文档中未使用的样式，默认 false |
| 新增 | `detailedInfo` | 控制是否返回 description 字段，默认 false |

**StyleInfo.description 字段说明**

- `description` 字段现仅在请求 `detailedInfo=true` 时返回
- 此功能依赖 WordApi BETA，在部分环境中可能不可用

---

## [0.1.1] - 2026-02-02

### 变更

**事件命名规范化**

为保持命名一致性，统一采用 `{namespace}:{action}:{target}` 格式：

- PPT 事件:
  - `ppt:slide:add` → `ppt:add:slide`
  - `ppt:slide:delete` → `ppt:delete:slide`
  - `ppt:slide:move` → `ppt:move:slide`
  - `ppt:slide:goto` → `ppt:goto:slide`

- Excel 事件:
  - `excel:get:rangeValues` → `excel:get:range`
  - `excel:set:rangeValues` → `excel:set:range`
  - `excel:sheet:add` → `excel:add:sheet`
  - `excel:sheet:delete` → `excel:delete:sheet`
  - `excel:sheet:rename` → `excel:rename:sheet`
  - `excel:sheet:activate` → `excel:activate:sheet`

**错误码调整**

- `2002` 从 `CONNECTION_LOST` 改为 `INVALID_TOKEN`（令牌无效）
- 新增 `2005 CONNECTION_LOST`（连接丢失）

**数据类型简化**

- `ChartType`: 简化为 Column, Bar, Line, Pie, Area, Scatter, Doughnut
- `ShapeType`: 合并两端实现，现包含 Rectangle, RoundedRectangle, Circle, Oval, Triangle, Diamond, Pentagon, Hexagon, Line, Arrow, Star, TextBox

### 新增

**PPT 事件**

- `ppt:insert:table` - 在幻灯片中插入表格
- `ppt:update:textBox` - 更新幻灯片中的文本框

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

协议当前处于 `0.x` 初始开发阶段，API 可能随时变更。
