# Langium 用于本插件语法分析的评估

调研日期：2026-10-29。调研对象：Langium（`eclipse-langium/langium`，npm `langium`）。

本文只记录调研结论与依据，不改变需求，也**不是迁移决定**。需求与验收目标仍以 [requirements.md](requirements.md) 为准；当前实现见 [implementation.md](implementation.md)。

## 1. 结论摘要

| 问题 | 结论 | 置信度 |
|---|---|---|
| Langium 能否解析 VBA 语法 | **能**。可用 `.langium` 语法声明生成 parser、AST、TextMate 高亮 | 高 |
| 能否自定义词法器处理 VBA 行结构（续行 `_`、冒号分隔 `:`） | **能**。Langium 3.2.0+ 支持通过 DI 覆盖 `parser.TokenBuilder` 与 `parser.Lexer`，官方 `IndentationAwareTokenBuilder` 即为先例 | 高 |
| 能否自定义词法器处理 `#` 歧义（日期字面量 / 预处理指令 / 文件号） | **能**，需自定义 TokenBuilder 的 matcher 函数（可带状态、可访问上下文） | 中高（机制存在，未实测 VBA 场景） |
| 是否满足本项目的容错要求（缺 `End Sub` 仍可补全/导航） | **PoC 已验证通过**（2026-10-29）：缺 `End Sub` 与语句级错误后，后续过程仍进入 AST，交叉引用仍解析成功 | 高（合成用例实测） |
| 是否支持条件编译（`#If` 按常量求值裁剪分支） | 官方**不提供预处理器**；但**PoC 已验证通过**：自研预处理层约 100 行即可实现分支裁剪，且诊断位置与原始文件一致 | 高（机制确认 + 实测） |
| 是否满足性能目标（1000 文件/10 万行 ≤ 10 秒） | **PoC 实测 4718 ms**（当前实现 1106 ms，**约 4.3 倍**），且 Langium 口径更轻；参考硬件下很可能逼近或超出 10 秒目标 | 中高（合成负载实测，参考硬件未测） |
| 当前是否建议替换现有解析器 | **不建议**（三项 PoC 完成后置信度提高）：容错与条件编译均验证通过，但性能约为当前实现的 4.3 倍且很可能不满足 PERF-02，另需自研工程隔离、推断候选重命名，格式化策略存在冲突；迁移需重写核心并重新验证全部 | 高 |

## 2. Langium 官方能力（对应本项目需求）

来源：`langium.org/docs/features/`、`langium.org/docs/reference/configuration-services/`、`langium.org/docs/recipes/` 各页面。

| 本项目需求 | Langium 官方支持情况 |
|---|---|
| LSP-05/06 补全 | `CompletionProvider` 服务，可覆盖；基于语法与作用域自动提供 |
| LSP-07 悬停 | `HoverProvider`，默认取前置注释，可覆盖 |
| LSP-09 定义跳转 | 交叉引用 `property=[Type:TOKEN]` + `ScopeProvider` 自动链接 |
| LSP-10 引用查询 | `ReferenceFinder` 服务 |
| LSP-11 文档符号 | `DocumentSymbolProvider` 服务 |
| LSP-12 工作区符号 | `indexManager` + 全局作用域（`getGlobalScope`），跨文件 |
| LSP-08 签名提示 | 需自行实现（官方文档未列为默认服务） |
| LSP-13 折叠 | `FoldingRangeProvider` 服务 |
| REN-01 重命名 | `RenameHandler`（Prepare + Rename），基于交叉引用 |
| FMT-01…03 格式化 | 有 Formatting recipe（`LangiumFormatter`），**基于 AST** |
| LSP-02 语义高亮 | `SemanticTokenProvider`，可覆盖 |
| SYN 大小写不敏感 | `caseInsensitive: true`（关键字）；标识符/交叉引用需 `MapScope/StreamScope` 的 `{ caseInsensitive: true }` 或覆盖 `ScopeProvider` |
| 关键字作标识符 | 支持但**需逐关键字显式枚举**（`name=(ID \| 'Keyword')`）；官方明确警示“possible, but has some costs” |
| COM-07 外部/内建符号 | Builtin Libraries recipe：`loadAdditionalDocuments` + 自定义 `FileSystemProvider`（`builtin://` scheme），支持导航 |

维护与许可：MIT；Eclipse 基金会；npm `langium` 最新版 4.4.0；依赖 8 个；Node >= 20.10.0。来源：npm 包页（第三方镜像 `npmx.dev` 汇总，中等置信度）。

## 3. VBA 特有难点与 Langium 的匹配

### 3.1 可通过自定义词法器解决（机制已确认）

- **`_` 续行与 `:` 冒号分隔语句**：Langium 默认 token 流无“逻辑行/物理行”概念。可通过覆盖 `parser.TokenBuilder` 与 `parser.Lexer` 实现（官方 `IndentationAwareTokenBuilder` 证明可用自定义 matcher 替代正则，并带状态检测层级变化）。
- **`#` 歧义**：`#` 同时用于日期字面量（`#1/1/2020#`）、预处理指令（`#If`）、文件号（`#1`）。可用自定义 matcher 按上下文区分。
- **类型后缀**（`$ % & ! # @ ^`）：词法正则可覆盖，但 `#` 需与上述歧义一并处理。

### 3.2 Langium 不提供，必须手写

- **条件编译求值（SYN-02、LIM-02）**：需按常量值决定哪些分支参与语义分析，未知常量分支须整段排除并给出 Information 级提示。Langium 语法层无法按常量值动态裁剪分支，需在解析前加预处理层，并维护**位置映射**使诊断仍指向原始文件行。
- **工程隔离（PRJ-02）**：Langium 工作区默认把全部 DSL 文件纳入全局作用域。本项目要求多工程符号隔离、嵌套工程父不串扫子文件。需覆盖 `WorkspaceManager` 与 `ScopeProvider`。
- **推断候选重命名（REN-02）**：`Application.Run` / `OnAction` 字符串候选须以 QuickPick 展示、**默认不勾选**，用户显式选择后才修改。Langium `RenameHandler` 只基于交叉引用，不支持字符串候选，需完全自定义。
- **COM 可移植类型索引与显式版本校验（COM-02/03）**：`PortableTypeIndex` 的 GUID/版本/LCID/位数比对是业务规则。可映射到 builtin library 机制，但校验逻辑须自建。
- **编码检测（INP-03）**：非 UTF-8 文件须暂停语义分析。不属 Langium 职责。
- **URI 归一**：当前 `canonicalUri()` 处理 Windows 盘符与转义差异。Langium 用 `vscode-uri`，是否需保留自有归一逻辑待验证。

### 3.3 存在策略冲突

- **格式化（FMT-04）**：当前 `formatVba` 是**物理行扫描、不依赖解析器**，因此语法错误时仍可格式化，并逐字保留注释/字符串、保留 BOM 与逐行 CRLF。Langium Formatter 基于 AST，语法错误可能导致格式化不可用，且与“不依赖解析器、保留物理行字节”的策略冲突。

### 3.4 容错能力（LSP-03、LIM-01）：PoC 已验证通过

调研初期的判断是“容错风险最高、必须实测”，原因是 Chevrotain 为 LL(k)/ALL(\*) 解析器，与当前逐行扫描策略不同。2026-10-29 完成 PoC 后，该判断被**推翻**：详见 §5。

要点：缺 `End Sub` 与过程内 `If` 缺 `Then` 两种错误下，Langium 均只报 1 个语法错误，且**后续过程仍完整进入 AST**，交叉引用 `Compute(1)` 仍解析到目标过程（意味着定义跳转可用）。

## 4. 当前实现对照

| 维度 | 当前实现（已验证） | Langium |
|---|---|---|
| 单元测试 | 117/117 通过 | — |
| VS Code API 检查 | 31 项通过（`output/vscode-api-check.json`） | — |
| 真实样例 | 61 文件只读普查通过，未改源码（`output/sample-check.json`） | 未验证 |
| 性能（1000 文件/10 万行） | 冷启动索引 ~1.1 秒（`output/lsp-benchmark.json`，16 核/28 GB） | 未验证 |
| 语法定义方式 | 手写 TypeScript 词法 + 结构解析 | 声明式 `.langium` 语法 |

迁移意味着重写核心解析器，并新增预处理器、覆盖 workspace/scope/rename/formatter、重新验证全部。Langium 的增量价值是声明式语法与成熟交叉引用框架；风险是容错、性能、条件编译三项均未验证。

## 5. PoC 记录（2026-10-29）

PoC 位于 `poc/langium/`，独立于主实现，**未改动 `src/` 下任何已验证代码**，也不进入发布包（见 `.vscodeignore`）。

### 5.1 方法

- 用 `langium-cli` 从最小 `.langium` 语法生成 parser、AST 与 TextMate 语法（`poc/langium/src/vba-mini.langium`）。
- 语法覆盖：模块成员（过程、模块级声明、`Option`、用户定义类型）、过程与参数、`Dim` 声明、赋值、调用、块结构（`If/ElseIf/Else`、`For`、`While`、`Do`、`With`）与表达式；开启 `caseInsensitive`。
- 用 Langium 核心服务 + `parseHelper` 解析，统计**语法错误数、AST 中保留的过程、交叉引用解析成功数**。
- 结果写入 `poc/langium/poc-result.json`。

### 5.2 结果

| 场景 | 语法错误 | AST 中过程 | 后续过程是否保留 | 交叉引用 |
|---|---:|---|---|---|
| A 正确基线（三过程完整） | 0 | First / Second / Compute | —（无错误） | 1 解析、0 未解析 |
| **B 首过程缺 `End Sub`** | **1** | First / Second / Compute | **是** | **1 解析、0 未解析** |
| **C 过程内 `If` 缺 `Then`** | **1** | First / Second / Compute | **是** | **1 解析、0 未解析** |

B 的错误为 `Expecting token of type 'End' but found 'Public'`；C 为 `Expecting token of type 'Then' but found '\n'`。

**结论**：在“缺 `End Sub`”与“语句级语法错误”下，Langium 均能在错误后恢复，后续过程仍进入 AST，且对后续过程的调用引用仍能解析——即定义跳转在有语法错误时仍可用。这直接回应了 §3.4 原列的“最高风险”。

### 5.3 真实样例未能验证（覆盖度问题，非容错问题）

对真实样例 `CQ24_AdvancedSort.bas`（1018 行，含 11 个 `Sub`、8 个 `Function`）测试结果：`procedureCount: 0`、`parserErrors: 2`。

阻塞点已精确定位到第 21 行的 `Private ...` 模块级声明：VBA 模块级私有变量写作 `Private X As Y`（**无 `Dim` 关键字**），而 PoC 语法只识别 `Private Dim ...`，解析器在此报“期望文件结束”后停止。

该文件还含大量 PoC 语法未覆盖的构造：`On Error`（25 处）、`Exit`（32 处）、`GoTo`（19 处）、行标签（17 处）、`Select Case`（1 处）、`ReDim`（4 处）。**任何解析器都需要相应语法覆盖才能解析**，因此该结果反映的是 PoC 语法覆盖度，不能作为 Langium 容错能力的反证。

若要验证真实文件，需要先实现接近完整的 VBA 语法（含无 `Dim` 的模块级声明、数组类型、`On Error`、`GoTo`、行标签、`Select Case`、`Exit`、续行、条件编译等）——这已属于迁移实施工作量，超出本次 PoC 范围。

### 5.4 PoC 过程中暴露的一个真实语法难点

Chevrotain 报出**歧义**：VBA 的普通赋值 `x = 1` 与函数返回语句 `Compute = a` 语法前缀完全相同（均为 `Name = 表达式`），LL(k) 无法区分，只能靠“是否在同名 `Function` 内”的语义判断。PoC 通过合并为一条规则绕过；真实实现需自行处理该区分（当前手写解析器已处理，见 `parser.ts` 的返回类型逻辑）。

### 5.5 性能 PoC（2026-10-29）

脚本 `poc/langium/src/bench.ts`，结果 `poc/langium/bench-result.json`。负载与主实现基准**完全一致**：1000 个文件、每文件 100 行、共 10 万行，合成内容为 `Option Explicit` + `Function` + `Dim` + 93 行 `result = result + 1`（与 `scripts/benchmark-server.cjs` 相同，且能被 PoC 语法解析，冷启动解析错误数为 0、1000 个过程全部识别）。

| 指标 | Langium（3 次中位数） | 当前实现（同机） | 说明 |
|---|---:|---:|---|
| 冷启动索引（磁盘 + 解析 + 建索引） | **4717.84 ms** | 1106 ms（`output/lsp-benchmark.json`） | Langium **不含**进程启动与 JSON-RPC；当前实现**含**进程启动与 JSON-RPC |
| 单文档增量重建 | 23.51 ms | 160 ms（增量诊断通知） | 当前实现含诊断发布与 JSON-RPC 通知；Langium 仅解析 + 索引 |

**口径差异很重要**：Langium 数字不含进程启动与 JSON-RPC，若补齐会更慢；当前实现的 1106 ms 已包含这两项。因此实际差距大于 4.3 倍。

**对照需求 PERF-02（首次索引 ≤ 10 秒，参考硬件 4 核 / 16 GB）**：本机为 16 核 / 28 GB（Ryzen 7 7840HS），明显强于参考硬件。Langium 在本机已用约 4.7 秒，按核数与性能缩放，在 4 核 / 16 GB 上很可能逼近或超出 10 秒；未经参考硬件实测，不能下"必然达标"或"必然不达标"的结论，但**风险显著**。当前实现 1106 ms 在参考硬件上仍有充足富余。

增量方面 Langium 更快（23.5 ms），但该口径不含诊断发布与通知，与当前实现不完全可比。

### 5.6 条件编译 + 位置映射 PoC（2026-10-29）

脚本 `poc/langium/src/condpoc.ts`，结果 `poc/langium/cond-result.json`。

**技术**：预处理层按常量求值 `#Const` / `#If` / `#ElseIf` / `#Else` / `#End If`，将非生效分支的行内容**清空但保留行数与行序**。由此生效分支内的行列位置与原始文件一一对应，诊断位置无需额外映射表。

**结果**：

| 检查 | 结果 |
|---|---|
| `#Const DEBUG = 1` 常量提取 | 成功 |
| 非生效分支（`#Else` 内）中的语法错误 | **未产生任何诊断**（`inactiveBranchErrorReported: false`） |
| 生效分支内故意语法错误（0 基第 10 行 `result = result +`） | **报出且位置正确**（Chevrotain 1 基 startLine=11 对应 0 基 10） |
| 行数保持 / 过程解析 | 保持 / `Test` 过程正常识别 |

**结论**：条件编译的分支裁剪与位置映射可行，自研预处理层约 100 行。

**局限**：PoC 求值器仅支持单一常量的 `NAME = value` / `NAME` 与 `#Const` 数字字面量；完整受限表达式（`And/Or/Not/Xor/Eqv/Imp`、`defined()`、字符串比较等）未在 PoC 中实现——当前手写实现 `parser.ts` 的 `CondExpressionParser` 已完整支持，迁移时需移植。

### 5.7 仍未验证的项

1. **真实文件解析**：需先补齐语法覆盖度（见 §5.3）。
2. **补全、格式化、工程隔离、推断候选重命名**：均未验证（调研阶段已判断需大量自研，见 §3.2/§3.3）。
3. **参考硬件性能**：未在 4 核 / 16 GB 机器上复测（见 §5.5）。

## 6. 来源

- Langium 文档首页与导航：https://langium.org/docs/
- Features：https://langium.org/docs/features/
- Grammar Language 参考：https://langium.org/docs/reference/grammar-language/
- Configuration via Services：https://langium.org/docs/reference/configuration-services/
- Recipe 大小写不敏感：https://langium.org/docs/recipes/lexing/case-insensitive-languages/
- Recipe 关键字作标识符：https://langium.org/docs/recipes/keywords-as-identifiers/
- Recipe 内建库：https://langium.org/docs/recipes/builtin-library/
- Recipe 缩进敏感（自定义 TokenBuilder/Lexer 机制）：https://langium.org/docs/recipes/lexing/indentation-sensitive-languages/
- Langium 1.0 公告（架构、Chevrotain、ALL(\*)）：https://www.typefox.io/blog/langium-1.0-a-mature-language-toolkit/
- GitHub：https://github.com/eclipse-langium/langium

## 7. 本文未做的事

- 未做 PoC，未安装或运行 Langium。
- 未验证 4.4.0 是否实现 1.0 路线图中的 “lexer modes”（该概念在 1.0 公告中列为未来功能；本次未单独查证当前版本状态）。
- 未评估 Langium 的 formatter 能否满足 FMT-04 的逐字节保留约束（仅指出策略差异）。
- 未比较 Langium 与 ANTLR、tree-sitter 等其他方案。
