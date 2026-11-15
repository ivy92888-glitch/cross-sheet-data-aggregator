# 你手里有两张表要合成一张？Cross-Sheet Data Aggregator 帮你一次性搞定

![](./examples/sample-data-preview.svg)

有一天领导给你两个 Excel 文件，说"把它们合成一张给我"。

你打开一看——左边是主表（Source A），有 89 行数据，带链接；右边是辅助数据源（Source B），有 55 行补充数据，也带链接。两张表都长这样：

> **Source A（主表）**
> - C 列：中文名 "支架" / "笔筒" / "花盆架"
> - E 列：标识符
> - J 列：单品净重 5.64 / 0.6 / 0.7 kg
> - R 列：销售价 39.99 / 16.99 / 22.99 欧元
> - T 列：百分比参数 0.00% / 0.00% / 6.50%
> - AB 列：链接 `https://www.amazon.de/dp/B0BNZXBGJ1?psc=1&pf_rd_w=...`（塞满追踪参数）
>
> **Source B（辅助数据源）**
> - 结构类似，但数据是补充的

你叹了口气，开始想："先把两张表的链接列都过一遍吧——同一个链接如果两张表都有，怎么处理？完全相同的删重复？同一个产品但中文名不同的（"支架" 也叫 "TV Stand"）要合并？不同产品共用一个链接的能保留吗？"

还有另一种情况：领导要你把主表的链接和一张参考库表（Source C）关联，用链接里的 **实体标识符** 作为跨表的桥梁——主表里只有链接，参考库里才有那个标识符对应的编码和百分比。得先从链接里正则提取出标识符，再在参考库里找匹配。

**Cross-Sheet Data Aggregator 把这两种跨表需求都做成了一键 VBA 宏。**

---

## 🏗️ 它是怎么做到的（两个独立模块，都跑一遍也可以）

### 模块 1：跨表合并 + 三级去重（Link Collector）

**干的事**：Source A（主表）+ Source B（辅助数据源）→ 一张干净的 MergedLinks 表。

**三级去重决策树**（从 89 + 55 = 144 条原始数据里，最终只留下约 110 条唯一记录）：

```
发现同一个链接在两张表里都出现了
  │
  ├─ 规则 1：B 列（中文名）+ D~H 列（标识符、百分比、描述、净重、价格）
  │          全部一致？
  │          └─ Yes → 删除重复的那行，保留第一次出现的 ✂️
  │
  ├─ 规则 2：D~H 列全部一致，但 B 列（中文名）不一样？
  │          （同一个产品，中英文两个名字）
  │          └─ 合并：把两个中文名拼到第一行，用 "/" 分隔 ✏️
  │
  └─ 规则 3：D~H 列也不一样？
             （不同产品共用同一个链接）
             └─ 两个都保留，各自成行 ✅
```

**数字例子**：
> Source A 第 3 行和 Source B 第 57 行都有链接 `https://amazon.de/dp/B0BNZXBGJ1`
> 
> 规则 1：检查——两张表的中文名都是 "支架"，标识符都是 9403208000，净重都是 5.64，价格都是 39.99 → **完全一致，删掉 Source B 那条**
> 
> 另一条：Source A 有中文名 "花盆架"，Source B 里同一个产品（同一个标识符）也有一条，但中文名写的是 "Plant Shelf" → **规则 2 命中，合并成 "花盆架 / Plant Shelf"，只留一行**

### 模块 2：实体标识符关联采集（Identifier Collector）

**干的事**：Source A（主表，有链接）+ Source C（参考库，有标识符对应的编码和百分比）→ 按实体标识符关联，一张 MergedCodes 表。

**核心桥梁**：从链接里正则提取实体标识符（`[A-Z][0-9A-Z]{9}`），用它当字典键。

```
Source A 链接: https://www.amazon.de/dp/B0BNZXBGJ1?psc=1...
                    ↓ 正则提取 [A-Z][0-9A-Z]{9}
Source A 索引: "B0BNZXBGJ1" → 第 3 行
                    ↓ Source C 有没有一行的链接也能提取出 "B0BNZXBGJ1"？
Source C 匹配: 找到！第 17 行的标识符列写着 "9403208000"，百分比 "0.00%"
                    ↓
合并写入 MergedCodes 表：
  中文名 "支架" | 编码 9403208000 | 百分比 0.00% | 净重 5.64 | 链接（已清洗）
```

**性能对比**：
- **手工**：打开 Source A 的 89 行，逐个复制链接里的标识符，在 Source C 的 500+ 行里 Ctrl+F 搜索匹配……一条一条来 ≈ 20-30 分钟
- **宏**：2-3 秒

---

## 🔥 适用场景（至少 6 种不同行业）

### 场景 1 — 跨境电商产品链接采集（原始场景）
主表 + 辅助数据源 → 合并成一张干净的产品链接追踪表，带百分比参数和批次标识符（Batch ID）标记，超链接可点击。

### 场景 2 — 多来源地址池合并
地址池被多个系统维护：系统 A 一份、手动 Excel 一份、历史导出一份 → 全部合并，按去重规则处理，生成统一的地址主数据表。

### 场景 3 — 财务交易记录跨表聚合
交易流水来自两个来源：ERP 导出一份、支付平台导出一份 → 按订单号匹配合并，去重后生成总账表。

### 场景 4 — HR 员工档案统一
员工基础信息在一张表、薪资明细在另一张表 → 用工号或邮件当关联键，合并成一张完整档案表。

### 场景 5 — 制造业部件选型推荐
历史部件使用记录在 Source A，供应商报价在 Source C → 按部件型号关联，生成一张"部件 + 历史使用频率 + 当前报价"的决策表。

### 场景 6 — 任意两张有关联键的表
只要满足这个条件：
> 表 A 有一个列能提取出某个唯一标识符（比如 URL 里的 ID、工号、SKU、订单号），表 B 也有这个标识符 → 就能用 Identifier Collector 合并。

---

## 🚀 快速上手（4 步）

### 1. 导入代码
VBA 编辑器（Alt + F11）→ 导入 `src/CrossSheetAggregator.bas`。

### 2. 准备你的三张表（列布局可调，见 docs/Configuration.md）

| 工作表 | 必填列 | 干什么 |
|---|---|---|
| **Source A** | A-J （默认） | 主数据表——含实体链接的那张 |
| **Source B** | 类似 Source A | 备用数据源 |
| **Source C** | A-C （默认） | 参考库（标识符 + 编码 + 百分比）——模块 2 用 |

> 示例数据见 `examples/sample-data.xlsx`（虚构的家居工具领域）。

### 3. 运行

```
一键跑两个模块:      RunAllAggregations()
只跑模块 1:          RunModule1()
只跑模块 2:          RunModule2()
```

### 4. 看输出

- 模块 1 输出 → **MergedLinks** 工作表（批次标识符（Batch ID）、中文名、英文名、编码、百分比、净重、价格、清洗后的链接 + 可点击超链接）
- 模块 2 输出 → **MergedCodes** 工作表（按标识符关联后的完整记录）

两张表的表头都已经：深蓝色底白字、自动列宽、百分比格式化（0.00%）、去重、冻结首行。

---

## 🎯 去重规则速查表

| 规则 | 触发条件 | 处理方式 | 典型案例 |
|---|---|---|---|
| **规则 1 — 全同删重** | 同链接 + 中文名 + 编码 + 百分比 + 净重 + 价格 全一致 | 保留第一行，删掉当前行 | 完全相同的产品记录在两张表里都存在 |
| **规则 2 — 同品异名合并** | 同链接 + 编码 + 百分比 + 净重 + 价格一致，但中文名不同 | 合并中文名到第一行，用 "/" 分隔 | "支架" 也叫 "TV Stand" —— 两个名字拼起来 |
| **规则 3 — 异品同链保留** | 只有链接相同，其他字段不同 | 两个都保留，各自成行 | 同一个电商页面里带两个产品变体 |

---

## 🛠️ 可调常量（源码顶部，不要在文档抄数值）

所有 Sheet 名、列字母、正则模式都在 `src/CrossSheetAggregator.bas` 顶部的 `Private Const` 里。改完保存即生效。**源码是唯一真源**。

| 常量 | 干什么 | 什么时候改 |
|---|---|---|
| SHEET_SOURCE_A / B / C | 三张源表的表名 | 你的表不叫 SourceA / SourceB / SourceC 时 |
| COL_LINK | 哪一列放实体链接 | 你的链接列不是 I 列时 |
| BATCH_ID_REGEX | 批次标识符（Batch ID）/批次号的正则（默认 4 大写字母 + 7 数字） | 你的批次号格式不同时 |
| ENTITY_ID_REGEX | 实体标识符的正则（默认 `[A-Z][0-9A-Z]{9}`） | 你的链接 ID 格式不同时 |
| TARGET_DOMAIN | 清洗后统一的域名后缀 | 不要 amazon.example-domain 时 |

完整说明见 `docs/Configuration.md`。

---

## ⚖️ 合规与免责

- 本项目由作者独立 **clean-room 重写**。不引用或拷贝任何前雇主的专有代码或业务文档。
- MIT 许可证仅覆盖作者本人贡献的实现（src/ 目录）。不构成对任何第三方权利的处分。
- 示例数据为完全虚构、中性的演示数据，与任何真实业务无关联。
- 与任何电子表格软件厂商无关联。各类产品名称仅作为常见运行环境提及，不代表背书。

---

## 📦 文件清单

```
cross-sheet-data-aggregator/
├── src/CrossSheetAggregator.bas    ← 唯一代码文件，2 个模块 + 8 个工具函数
├── examples/sample-data.xlsx         ← 虚构演示工作簿（Source A/B/C 三张表）
├── docs/Configuration.md             ← 常量说明 + 表结构模板 + 第二领域示例
├── README.md
├── LICENSE                           ← MIT
├── .gitignore
└── .gitattributes
```

---

## 📅 版本

- **v1.0.0** — 2026-11-15 — 初始发布：Link Collector（跨表合并 + 三级去重）+ Identifier Collector（实体标识符关联采集）

---

## Quick Summary (English)

**Cross-Sheet Data Aggregator** is a single-file Excel VBA toolkit that
merges data from multiple source sheets into standardized output
sheets. Two independent modules:

- **Link Collector** — Merge rows from Source A + Source B by product link,
  with 3-tier dedup (identical → skip, same product / diff name → merge,
  different product / same link → keep both). Output: `MergedLinks` sheet.
- **Identifier Collector** — Match rows from Source A + Source C by
  entity ID extracted from links (`[A-Z][0-9A-Z]{9}` regex). Output:
  `MergedCodes` sheet.

Both modules sanitize entity links (strip query params → normalize domain
suffix), extract batch IDs from text (4 uppercase + 7 digits), and add
clickable hyperlinks. Designed as a generic framework — drop into any
domain where multiple worksheets need to be consolidated by a shared key.

## Maintenance Note

This README is intentionally **Chinese-primary**. The English section is
a concise summary only. All parameters and constants live in source code —
**do not copy numbers into documentation**. Source is the single source
of truth.
