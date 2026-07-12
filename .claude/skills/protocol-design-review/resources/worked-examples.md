# 范本案例 — 协议接口收敛（对照校准用）

> 这些是 oasp-protocol 已落地的真实案例（源自 #10 `/ppt` 字体收敛）。评审时对照它们校准"收敛得好"长什么样。每个案例给 **❌ 反模式 → ✅ 收敛后 → 为什么**。

---

## 案例 1：散堆扁平字段 → 可复用实体（Step 2.1 / 2.3）

**场景**：`/ppt` 文本框要支持字号、字体、颜色、粗体、斜体、下划线、删除线、上下标、大小写……

❌ **反模式**（被节奏压着走：需求逐条翻成字段）：
```typescript
interface TextBoxUpdates {
  fontSize?: number; fontName?: string; color?: string;
  bold?: boolean; italic?: boolean; underline?: ...;
  strikethrough?: boolean; superscript?: boolean; /* …再加 6 个 */
}
// run 级要局部格式？→ 又平铺一遍相同字段。散堆 × 2。
```

✅ **收敛后**：
```typescript
interface PptFont {          // 一个高内聚字体实体，12 属性
  size?; name?; color?; bold?; italic?; underline?: ShapeFontUnderlineStyle;
  strikethrough?; doubleStrikethrough?; superscript?; subscript?; allCaps?; smallCaps?;
}
interface TextBoxUpdates {
  font?: PptFont;            // 整框级
  runs?: PptTextRun[];       // run 级：{start, length, font: PptFont} — 复用同一实体
}
```
**为什么**：12 个字段高内聚（都是"字体"），收进 `PptFont` 后，整框级 / run 级 / 表格单元格**三处复用同一实体**。接口只挂实体、不挂散字段。复用度 = 抽象价值。

---

## 案例 2：更少接口 + 更富嵌套 vs 更多接口（Step 2.2）

**场景**：要支持 run 级局部格式、段落级项目符号。

❌ **反模式**：新开 `ppt:update:run`、`ppt:update:paragraph` 两个事件 → 接口膨胀。

✅ **收敛后**：挂在既有 `ppt:update:textBox` 上，用统一的 `{start, length, payload}` 区间寻址：
```typescript
runs?: PptTextRun[];              // {start, length, font}
paragraphs?: PptParagraphStyle[]; // {start, length, bulletFormat}
```
**为什么**：能力用区间寻址就能承载时，别为它新增顶层接口。同一 `{start,length,...}` 模式在 run/paragraph 复用，一致且省接口。

---

## 案例 3：跨命名空间——有意不复用，零映射对齐宿主（Step 4）

**场景**：`/ppt` 要下划线枚举，`/word` 已有 `UnderlineStyle`。

❌ **反模式**：直接复用 `/word` 的 `UnderlineStyle`（7 值小写）——"反正都是下划线"。

✅ **收敛后**：`/ppt` 另立 `ShapeFontUnderlineStyle`（17 值 PascalCase），并写 admonition：
> 与 Word 的 `UnderlineStyle`（7 值小写、映射 Word.js）**有意不同**；PPT 对齐 office.js PowerPoint 同名枚举、零映射。服务不同宿主 API，故不复用——这是有意的跨命名空间差异，非命名不一致。

**为什么**：Word.js 与 office.js 是不同宿主 API，枚举字面量不同。"抄"会制造映射 bug。有意不同必须写成文档，否则后人当 bug 修回去。

---

## 案例 4：枚举承载分档（Step 5.4）

**场景**：`underline` 17 个值、bullet `style` 40+ 个值。

✅ **决策**：
- `underline`（17 值，小而封闭稳定）→ **就地全枚举** `type ShapeFontUnderlineStyle = "None" | "Single" | …`，可校验。
- bullet `style`（40+ 值，随宿主版本增补）→ **直通 `string`**，宿主校验、非法 → `4002`。

**为什么**：不是"underline 认真、style 偷懒"，而是按**集合规模与稳定性**分档。二者同为"零映射对齐宿主"。**关键**：把这个分档理由写进文档，否则会被当成"不一致"。

---

## 案例 5：Draft 命名空间 clean-break，不留 Deprecated 别名（Step 6.2）

**场景**：收敛后旧的扁平 `fontSize`/`fontName`/… 怎么办？`/ppt` 是 Draft。

❌ **反模式**（求稳，怕破坏）：保留旧扁平字段标 `⚠️ Deprecated`，加一条"同时提供时 `font.*` 优先"的合并规则。→ 协议表面**同时背着**实体化和散堆两套写法 + 一条双端都要实现的合并规则。

✅ **收敛后**：`/ppt` 是 Draft（无稳定性承诺）→ **直接删除**扁平字段，只留 `font`。changelog 如实标"破坏性收敛（Draft 允许）"，**不写成"向后兼容"**。

**为什么**：Draft 阶段留兼容别名 = 给协议表面留永久债。Stable（`/word`）则相反：破坏性变更要独立 cross-ask + 独立 PR。

---

## 案例 6：降级全或无，消除"部分成功"（Step 5.3）

**场景**：表格批量设单元格字体，遇到合并单元格的非左上格 / 越界坐标。

❌ **反模式**：`console.warn + continue` 跳过非法格 → **部分成功**语义（设了一半）。

✅ **收敛后**：前置校验所有目标格（含行/列展开），任一越界或命中合并单元格空对象（`getCellOrNullObject`）→ 整请求 `4002` 失败、**不写任何格**；否则全部排入单次 `sync()` 提交。

**为什么**：协议不引入"部分成功"——调用方无法可靠预期状态。全或无 = 可预期。这条常在"赶紧让它跑起来"时被违反。

---

## 案例 7：反推需求方的事实假设（Step 3）

**场景**：需求来源 office4ai#45 标注了一批 requirement set 版本。

✅ **动作**：`/cross-ask office-editor4ai` 逐项复核，**纠正**了多处：删除线/上下标/大写实为 1.8（非 1.4）、`getSubstring` 实为 1.4（非 1.5）、下划线 17 值（非 15）、`bulletFormat` 无 `character`。

**为什么**：Requestor 的版本/事实假设可能错。以宿主复核为准，别把错误假设固化进协议。
