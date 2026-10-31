# VBAdminx 0.2.3 实现说明

本文记录 vbadminx-vscode 0.2.3 **已实现且经本机验证**的架构与行为。需求与验收目标以 [requirements.md](requirements.md) 为准（本文不修改需求）；AI 协议细节见 [ai-prototype.md](ai-prototype.md)；术语见 [GLOSSARY.md](../GLOSSARY.md)；操作步骤见 [README](../README.md)。

## 1. 定位

- VS Code 扩展：`vbadminx-vscode` 0.1.0，`engines.vscode: ^1.104.0`，`@types/vscode` 1.104 API 基线，本机实测宿主 VS Code 1.143.1。
- 三块能力，全部离线可运行（BND-06）：
  1. **VBA 语言服务器（LSP）**：解析、诊断、补全、导航、重命名、格式化、折叠、语义高亮。
  2. **COM 可移植类型索引**：Windows 提取 + 全平台导入/导出/校验 + 工程显式引用。
  3. **AI 只读上下文工具与可选 Go 模型适配器**：见 [ai-prototype.md](ai-prototype.md)。
- 边界不变：不连接 Office、不运行宏、不改源码（BND-05）；源码工程导入导出由其他工具负责（BND-03）。

## 2. 架构与代码组成

```
VS Code 扩展宿主
├─ src/extension.ts      activate()：AI provider/工具/命令、COM 命令、编码检查、
│                        startLanguageClient()、状态栏（vbadminx.projectStatus）
├─ src/languageClient.ts LanguageClient（IPC）→ dist/server.bundle.js
│                        rename middleware、vba-preview/vba-library 虚拟文档、
│                        文件监视（**/*.{bas,cls,frm,json}）、配置同步 vbadminx
├─ src/encoding.ts       打开时磁盘字节 UTF-8 检测
├─ src/com.ts / src/contextTool.ts / src/provider.ts / src/goClient.ts
└─ 语言服务器（独立进程，bundle 内含 vscode-languageserver）
   ├─ src/server.ts            LSP 连接、能力声明、诊断发布（100ms 防抖）
   ├─ src/language/project.ts  ProjectStore：工程发现、配置解析、索引加载、UTF-8 校验
   ├─ src/language/parser.ts   词法 + 结构级容错解析（条件编译、保护行、折叠）
   ├─ src/language/service.ts  VbaService：解析消歧、补全/悬停/导航/诊断/语义 token
   └─ src/language/formatter.ts formatVba：物理行级格式化（不依赖解析器）
```

- 构建：`npm run compile` = `tsc -p .` + `scripts/build.cjs`（esbuild 打出 `dist/extension.bundle.js`、`dist/server.bundle.js`，platform node、target node20、cjs；扩展 bundle 仅 `vscode` 为 external）。F5 与 VSIX 都用 bundle，改源码后必须重编译。
- 打包：`npm run package` → `vsce package --no-dependencies … --out output/vbadminx-vscode-0.1.0.vsix`。`main` 指向 bundle，**运行时不需要 node_modules**；`.vscodeignore` 排除 `node_modules/**`、`src/**`、`test/**`、`output/**`、`dist/test/**`、`dist/src/**`，但保留 **`scripts/export-com.ps1`**（发布包必须包含该脚本，Windows 提取依赖它）。
- URI 归一：所有文件 URI 经 `canonicalUri()` 归一（Windows 盘符大写 + `pathToFileURL` 重编码），避免盘符与转义差异产生重复模块；有单测覆盖（`test/regression.test.ts`）。
- 信任：`capabilities.untrustedWorkspaces.supported: false`；`extensionKind: ["workspace"]`（远程场景在扩展主机运行）。

## 3. LSP 能力（`src/server.ts` onInitialize 声明）

| LSP 能力 | 触发/参数 | 实现 |
|---|---|---|
| textDocumentSync | Incremental | `TextDocuments`；didOpen/didChange 重建模块并重算上下文内模块诊断 |
| completion | 触发字符 `.` | `VbaService.completions`：可见符号（作用域 0 / 工程 1 / 外部 2）+ 关键字（3）；`With` 隐式成员、成员补全；注释与字符串内返回空 |
| hover | — | 签名 markdown + 文档；外部/内建符号标注来源（LSP-07） |
| definition | — | 同工程跳声明 `selectionRange`；外部符号跳 `vba-library:` 只读虚拟文档（LSP-09） |
| references | includeDeclaration | 遍历上下文内全部模块的 identifier 并按解析结果过滤（LSP-10） |
| documentSymbol / workspaceSymbol | — | 过滤过程级作用域符号；工作区查询跨工程聚合但不共享解析作用域（LSP-11/LSP-12） |
| signatureHelp | 触发 `(` `,` | 解析调用名 → 参数表 + 按逗号计活动参数（LSP-08） |
| rename + prepare | — | 见 §6.5；工程冲突/外部符号/模块类名拒绝（REN-01/04、CFG-05） |
| document/range Formatting | — | 见 §6.6（FMT-01…04） |
| foldingRange | — | 解析产生的块区域与注释连续段（LSP-13） |
| codeAction | QuickFix | `name-case` 修正、`undeclared` 插入声明，均先预览（QF-01/QF-02），见 §6.7 |
| semanticTokens | full，legend 9 类型 2 修饰 | 只为**可解析**符号着色，未解析引用不赋确定语义色（LSP-02） |
| workspace folders | 增删通知 | 变更后重建 `ProjectStore` |

自定义请求：`vba/projectStatus`（状态栏与命令）、`vba/renameCandidates`（重命名推断候选）、`vba/context`（AI 工具的 `symbols`/`references` 模式）。

## 4. 工程发现与配置

- 发现（`ProjectStore.discover`）：任一层存在 `vba-project.json` 即为工程（显式配置可在更深目录）；无配置时，仅工作区根或直接子目录同时出现 `模块/`、`类模块/`、`表格对象/`、`窗体/` 四类目录才按约定识别（PRJ-03/04）。忽略 `node_modules`、`.git`、`.svn`、`.vscode-test`、`dist`、`output` 与点目录。
- 隔离与单文件：多工程符号隔离（PRJ-02，`context()` 只含同工程模块）；不归属工程的文件按单文件模式（PRJ-05）。嵌套工程时父扫描跳过子工程根。
- 配置字段（`schemas/vba-project.schema.json` + `parseProjectConfig` 运行时校验，`additionalProperties: false`）：

| 字段 | 类型 | 说明 |
|---|---|---|
| `name` / `host` | string | 工程名、宿主名（如 Excel） |
| `targetBits` | 32 \| 64，默认 64 | 决定 `Win32`/`Win64` 条件分支 |
| `constants` | 标量映射 | 条件编译常量（CFG-01） |
| `modules` | `{ [相对路径]: { name?, kind? } }` | 模块名/类型显式映射，`kind ∈ standard/class/document/form` |
| `typeIndexes` | string[] | 类型索引相对路径，**须手工加入，导入命令不静默改写**（COM-03） |
| `references` | 数组，**必填** `name/guid/major/minor/lcid/bitness` | 外部引用的显式版本，与索引逐项比对（GUID 忽略大小写与花括号，版本/位数区分） |

- 冲突（`projectDiagnostics` + `conflictedModules`，CFG-04/05）：
  - 显式模块名与 `Attribute VB_Name` 不符 → `project-conflict` Error，该模块入 `conflictedModules`；
  - 显式模块类型（`modules[].kind`）与导出头（`VERSION`/`Attribute`）推定类型不符 → 同上；
  - 同工程重名模块 → `project-conflict` Error；
  - 工程内任一模块存在上述冲突时，**同工程其他文件也会收到** `工程内存在导出头或模块类型冲突` 的 Error，`prepareRename`/`onRenameRequest` 据此**全工程拒绝重命名**；
  - 引用声明但缺索引 → `类型信息缺失`（Information）；索引与引用版本不匹配 → 报错且不入列。
- 示例：`examples/vba-project.json`（含 `modules`、`constants`、`typeIndexes`、`references` 结构）；命令 `vbadminx.createProject` 生成同结构初始文件，已存在不覆盖。
- 模块名优先级（CFG-02/03）：显式配置 > VBE 头 > 目录约定 > 扩展名/文件名推导。

## 5. 解析与诊断

### 5.1 解析范围（`src/language/parser.ts`）

- 词法类别：identifier/keyword/string/number/comment/operator；支持 `_` 续行合并逻辑行、冒号分隔语句、类型后缀（`$%&!#@`）、`Rem` 注释、日期字面量、`[方括号标识符]`。
- 结构级容错解析：过程/`If`/`With`/`For`/`Do`/`While`/`Select`/`Type`/`Enum` 块配对；单行 `If ... Then 语句:语句` 不作为块起点；多行参数；`Function ... As 类型()` 数组返回与关键字同名形参正确保留；`Declare PtrSafe`、`Event`、`WithEvents`、事件过程、`Implements`；窗体/文档头识别与设计段保护行（`protectedLines`）。
- 条件编译（LIM-02 相关）：环境内置 `VBA7=true`、`VBA6=false`、`Win32/Win64` 按 `targetBits`、`Win16/Mac=false`，叠加 `constants` 与 `#Const`；只分析生效分支；**无法判定的分支整段从语义分析排除并给 `cond.unknown`（Information）**。
- 未定义常量/未知表达式的分支不参与解析与诊断；保护行（Attribute、设计段）不参与语义诊断。

### 5.2 诊断码

| 来源 | 码 | 级别 | 含义 |
|---|---|---|---|
| 词法 | `lex.unterminatedString` | Error | 字符串缺结束引号 |
| 词法 | `lex.lineTooLong` | Error | 物理行超过 1023 字符 |
| 词法 | `lex.invalidCharacter` | Error | 存在 VBA 词法不合法的字符（如独立的 `@@@`） |
| 词法 | `lex.continuationComment` | Error | 行连接符后跟注释 |
| 词法 | `lex.tooManyContinuations` | Error | 行连接超过 24 个（逻辑行最多 25 个物理行） |
| 结构 | `syntax.missingEnd` / `syntax.unexpectedEnd` / `syntax.mismatchedEnd` / `syntax.missingThen` / `syntax.unmatchedElse` / `syntax.missingName` / `syntax.malformedDeclare` / `syntax.malformedProperty` / `syntax.nestedType` | Error | 块与语句结构错误（LSP-03） |
| 结构 | `syntax.optionPosition` / `syntax.optionDuplicate` | Error | Option 语句必须位于声明前、不得重复（PRE-01） |
| 结构 | `syntax.deftypePosition` | Error | Deftype 必须位于声明前（PRE-09） |
| 结构 | `syntax.exitWhile` | Error | VBA 不存在 Exit While（STMT-08b） |
| 语义 | `ambiguous-name` | Error | 模块级 Sub/Function 重名（SEM-03/05） |
| 语义 | `param-optional-order` / `param-array-position` / `param-array-unique` / `param-array-byref` | Error | 参数规则：Optional 次序、ParamArray 位置与唯一性、数组参数须 ByRef（SEM-06） |
| 语义 | `type-public-in-class` | Error | 类模块中用户定义类型必须 Private（SEM-08） |
| 语义 | `withevents-non-object` | Error | WithEvents 变量必须为对象类型（SEM-11） |
| 语义 | `friend-in-standard` | Error | 标准模块中不允许 Friend（STMT-01b） |
| 条件编译 | `cond.unknown` | **Information（可配）** | 常量未定义，分支排除，非确定错误（LIM-01/LIM-02）；与 `missing-*` 同受 `vbadminx.diagnostics.missingInformation` 控制，`hint`/`off` 同样生效 |
| 条件编译 | `cond.missingExpression` / `unmatchedElseIf` / `unmatchedElse` / `malformedConst` / `unmatchedEndIf` / `unclosedIf` | Error | 指令结构错误 |
| 语义 | `missing-type` / `missing-host` / `missing-external-symbol` | Information（可配） | 缺类型、缺宿主、缺外部符号类型信息，源码级能力仍可用，只提示不报确定错误（COM-06/LIM-01） |
| 语义 | `undeclared` | Error | **仅 `Option Explicit` 模块**，且工程没有缺失的外部类型信息；宿主对象名、`xl/mso/ad/wd/ac` 前缀名、窗体/文档模块一律不报；无 Explicit 不报（LIM-03）。工程缺外部类型时同一位置降为 `missing-external-symbol` |
| 语义 | `name-case` | Hint | 引用与声明大小写不一致（QF-02） |
| 语义 | `type-mismatch` | Error | 字面量与已声明类型不可判定转换（仅证据确定时，LSP-04） |
| 语义 | `argument-count` | Error | 已知过程的实参数量不符（LSP-04） |
| 工程 | `project-conflict` | Error／Information | 见 §4 |
| 编码 | `encoding` | Error | 非 UTF-8，语义分析暂停（INP-03，见 §8） |

- 发布规则：每次发布语义诊断上限 300 条；`vbadminx.diagnostics.missingInformation`（information/hint/off）作用于所有 `missing-*` 码**及 `cond.unknown`**；`vbadminx.diagnostics.nameCase`（bool）可关闭 `name-case`。诊断随编辑 100ms 防抖重算，发布目标 URI 经 `canonicalUri` 归一。

### 5.3 明确不做（保守性）

- **解析是容错的结构级分析，不是完整 VBA 编译器**：不做全表达式类型推导；`CreateObject`/`GetObject` 返回 Object，其成员需类型索引或补充声明才可解析（dynamic 对象成员**不确定**时不给确定结论，也不报确定错误）。
- `type-mismatch` 只覆盖字面量赋值等可判定场景，不做运算表达式类型检查。
- **不做表达式语法校验**：`a = = 1`、`a, b = 1`、`Foo (1, 2)`、`a = (1, 2)` 等表达式层非法构造不报错（见 `output/vba-cases-report.json` 的 STMT-03b/04b、EXPR-03b/11、DIAG-02）。
- Default 属性唯一性（SEM-07b）与 Const 常量表达式合法性（SEM-09b）未校验。

### 5.4 用例回归（2026-10-30）

以《VBA解析器测试用例.md》v1.0（106 项，依据 MS-VBAL）建立自动判定：`npm run cases:vba`，报告写 `output/vba-cases-report.json` / `.md`。首轮全量 75/105，修复后 **97/105 通过、误报清零**；剩余 8 项漏报全部属于 §5.3 声明的表达式校验边界与 SEM-07b/09b 两个未校验语义项。修复内容：`Next a, b` 多层关闭、Default 修饰符、行连接上限与续行后注释、1023 字符行、非法字符、Option/Deftype 位置、Exit While、`Dim` 缺名、参数规则（Optional 次序/ParamArray/数组 ByRef）、模块级过程重名、类模块 Public Type、WithEvents 非对象、标准模块 Friend、方括号标识符语句路由。

## 6. 语义服务行为（`VbaService`）

1. **解析消歧**：`resolve()` 按 声明位 → 成员访问（`.` 前缀、`With` 隐式）→ 作用域可见集（局部遮蔽 → 本模块 → 共享模块 → 外部/内建）逐级消歧；大小写不敏感；候选不唯一则不解析（不猜）。
2. **补全**：作用域可见符号 + 内建符号（`MsgBox`、`Len`、`vbCrLf` 等）+ 关键字；`对象.` 与 `With` 内 `.成员` 走 `membersOf`（本工程模块成员 + 类型索引成员 + 内建 `Debug/Err/Collection` 成员）。
   属性链、方法返回值和 `With obj.Child` 逐级使用类型；`Library.Type` 保留库名，避免合并同名 COM 类型。命名实参绑定到被调用过程的参数。
3. **悬停/签名**：内建函数表（`builtinData`）与类型索引提供签名、参数、文档。
4. **引用/符号**：按解析结果过滤，字符串与注释不参与（REN-01）。
5. **重命名流程**（原生 F2，`provideRenameEdits` middleware）：
   `next()` 取确定引用编辑 → 请求 `vba/renameCandidates` → **推断候选（`Run`/`OnAction` 字符串、类型未知的成员调用）以 QuickPick 展示且默认不勾选** → 模态确认（公共入口/事件过程提示源码外调用，REN-03）→ 逐文件打开 `vba-preview:` 差异预览（`vscode.diff`）→ 二次模态确认后才返回编辑（REN-02）。
   前置拒绝：非标识符/关键字新名、外部符号、模块或类名、工程符号冲突（CFG-05）。缺失类型索引不阻止已确定的局部变量重命名。差异预览检查工作区归属和文档版本；预览期间源码变化时取消旧修改。类型后缀保留，属性访问器成组修改。
6. **格式化**（`formatVba`，物理行扫描、不依赖解析器）：默认 4 空格 + 关键字标准大小写；`vbadminx.format.lineWrap`（120 列内仅安全逗号处续行）、`reorderDeclarations`（仅安全连续 `Dim`）、`formatDesign`（仅设计段排版）三项默认关闭；保留 BOM、逐行换行形式、字符串/注释/条件编译语义；未闭合字符串行原样保留；幂等（测试覆盖）。`[vba]` 默认 `formatOnSave: false`、指定默认格式化器。
7. **快速修复**（均经 `vbadminx.previewEdit`：工作区成员校验 → 差异预览 → 模态“应用修改”）：
   - `name-case` → 统一为声明名；
   - `undeclared` → 在所属过程声明行（含续行跳过）后插入 `    Dim <name> As Variant`，仅当该行是 `<name> =` 赋值形态。
8. **语义 token**：仅解析成功的 identifier；legend 类型 `namespace/class/enum/parameter/variable/property/enumMember/function/method`，修饰 `declaration/readonly`。

## 7. COM 类型索引

- 格式：`PortableTypeIndex`，`formatVersion: 1`；library 记录 `name/guid/major/minor/lcid/bitness`；类型 kind 仅 `class/interface/enum/module`，成员 kind 仅 `method/property/constant/event`。校验（`src/typeIndex.ts`）拒绝未知字段与超限（name 256、documentation 20000、types 50000、members 10000、parameters 1024），单文件 ≤ 32 MB，严格 UTF-8。契约与 `schemas/type-index.schema.json` 由单测保持一致。
- 命令与平台：

| 命令 | 平台 | 说明 |
|---|---|---|
| `vbadminx.extractComTypeIndex` | **仅 Windows 扩展主机**（COM-01） | 调 `scripts/export-com.ps1`（PowerShell 5.1，`LoadTypeLibEx(REGKIND_NONE)`：不注册类型库、不加载执行、不启动 Office）；参数逐项传递 `shell:false`；120 秒超时/可取消；位数（32/64）必须显式选择并回验（COM-02） |
| `vbadminx.importTypeIndex` / `exportTypeIndex` | 全平台（COM-02/05） | 校验后读写，写出为无 BOM UTF-8；覆盖需确认；默认目录 `.vba/type-indexes` |
| 导入/提取后的提示 | 全平台 | 若索引在工作区内，提示**手工**加入 `vba-project.json` 的 `typeIndexes`；工作区外则明示不会被工程使用（COM-03，不静默改配置） |

- 提取脚本行为：`TKIND_RECORD`/`UNION`/`ALIAS` 在 `PortableTypeIndex` 中无对应表示，**不输出**（record/alias 未表达）；coclass（`TKIND_COCLASS`）映射为 `class` 并展开其非 source 接口、source 接口作为事件输出。
- **实测**：本机 Windows 成功提取 `C:/Windows/System32/stdole2.tlb`（stdole 2.0、64 位），当前报告 `output/com-check.json` 为 **13 types / 90 members**（`coclassCompletion: true`），并抽查确认 `New StdFont` 的 `Name`、`Size` 等成员可用。类型/成员计数随展开策略可能变化，**以最新一次提取报告为准**。
- 区分呈现：VBA 内建符号、类型索引外部符号、源码符号在补全/悬停中分别标注（COM-07）。

## 8. 编码处理（INP-02/INP-03）

- 扩展侧（`src/encoding.ts`）：打开 `.bas/.cls/.frm` 时读磁盘字节做严格 UTF-8 解码；失败 → 警告“语义分析已暂停”，不改文件；成功且文件未修改、编码与打开编码不一致 → 以 `{ encoding: 'utf8' | 'utf8bom' }` 重新打开，**保留原 BOM 状态**。
- 服务器侧（`ProjectStore`）：同一严格解码，失败入 `invalidFiles`，该文件只保留一条 `encoding` Error，跳过语义诊断/补全/语义 token/格式化/快速修复；关闭文档后从磁盘重读。
- 限制（如实声明）：与 VS Code 自身保存能力协作，**不承诺拦截用户手动“另存为”导致的转码**；非 UTF-8 文件需用户在扩展外转换（INP-03 要求的提示已实现，转换动作本身不属本扩展）。

## 9. AI 接入现状

- 模型适配器与只读工具**保留**：默认 `magpie` 连接、默认模型 `opencode-go/deepseek-v4.1-flash`（`scope: application` 设置）。
- `get_vba_context` 输入模式已扩为 `selection | diagnostics | symbols | references`（`symbols` 可带 `query`），返回体字段 `semanticAnalysis: true`；`symbols/references` 经 `vba/context` 请求由语言服务器回答，超 24000 字符截断。信任、工作区成员、活动文件类型检查不变。
- 协议、密钥、SSE 约束与安全说明见 [ai-prototype.md](ai-prototype.md)（其“无 LSP”表述已标注为历史状态）。

## 10. 验证状态（本机 Windows 实测）

### 已通过

| 项 | 结果 |
|---|---|
| `npm run typecheck` | 通过 |
| `npm test` | 单元回归测试通过，覆盖解析、服务、格式化、工程、COM 索引、AI 协议和修改安全边界。最终数量见交付报告 |
| `npm run test:host`（VS Code **1.144.0**） | 31 项检查通过，含基础语言与 AI 接入、COM 成员补全、BOM/CRLF 保存、坏编码暂停语义修改、文件增删、单次撤销跨文件重命名和范围格式化边界。明细见 `output/vscode-api-check.json`（`failed: []`） |
| VSIX 安装验证 | 0.1.0 时已用 `output/vbadminx-vscode-0.1.0.vsix` 安装到**隔离的 extensions 目录**（`VBA_INSTALLED_EXTENSIONS` 模式）验证补全与命令面板，截图 `output/screenshots/01-03` 为 installed-vsix 模式产出；**未改动用户日常 VS Code 配置**。0.2.0 已重新构建 `output/vbadminx-vscode-0.2.0.vsix`，包内容由 `npm run package` 校验 |
| `node scripts/check-sample.cjs <导出目录>`（只读普查，样例路径见 requirements 第 9 节） | 61 文件、1 工程、0 坏编码。唯一 Error 为 `CQ24_AdvancedSort.bas` 第 567 行对应过程真实缺 `End Sub`；其余类型、宿主与大小写提示的计数见 `output/sample-check.json`。未改源码，报告不含源码文本，未加载 Office 或外部类型索引 |
| `npm run benchmark:lsp` | `output/lsp-benchmark.json`：真实 LSP 子进程、磁盘扫描、JSON-RPC 与诊断通知，1000 文件、10 万行。实际数值以报告为准；不含 VS Code UI 启动，`referenceHardwareMatched: false` |
| `npm run benchmark` | `output/benchmark.json`：in-process 解析与语言服务 CPU 口径的对照报告，数字以报告文件为准 |
| COM 提取 | stdole 提取成功，`output/com-check.json`：**13 types / 90 members**，抽查确认 `New StdFont` 的 `Name`/`Size`（见 §7） |
| `npm run smoke:go`（Magpie 真实模型） | 此前通过两轮工具往返；接入协议未变（该脚本消耗订阅额度，只发合成样例） |

**性能口径声明（必须连同数字一起引用）**：`lsp-benchmark` 覆盖真实子进程、磁盘与 JSON-RPC，但仍**不含 VS Code UI 启动**；机器为 16 核 / 28 GB（AMD Ryzen 7 7840HS）、Node v22.23.2、win32，与需求 PERF-01 的 4 核 16 GB 参考硬件**不匹配**（`referenceHardwareMatched: false`），故**不能声称 PERF-01/PERF-04 达标**，也无 PERF-05 的远程主机记录。三项数值相对目标（10000 / 500 / 200 ms）的裕量以报告为记录，不作达标结论。

### 未验证（不得笼统声称通过）

- 官方 `direct` 模式未实连（无独立 API Key）。
- macOS、Linux、Remote SSH/WSL/dev container：实现路径存在（`extensionKind: workspace`、纯 Node、非 Windows 拒绝提取），但**无实际环境验收**。
- AI 真实聊天 UI 未人工验收；不受信任工作区单独启动验收未做。
- `.frx` 二进制资源不解析；窗体预览只读取 `.frm` 文本设计段，图片与部分列表内容以占位呈现。
- 语言服务器诊断文本与内置符号说明仍为简体中文；界面多语言覆盖扩展侧文案与清单。


## 11. 0.2.1 宿主模型

| 项 | 说明 |
|---|---|
| 模块 | `src/language/hostModel.ts`（内建 Excel 16.0 近似声明） |
| 注入条件 | 工程存在且 `host` 为 Excel 或未配置；单文件模式与 Word/Outlook 宿主不注入 |
| 合并策略 | `ProjectStore.context()` 把内建索引追加到工程索引之后，并按类型名过滤；显式索引提供同名类型时不重复注入。宿主全局符号（ThisWorkbook、ActiveSheet、xlUp 等）仅在未被显式索引覆盖时加入 `LanguageContext.hostGlobals` |
| 服务行为 | `VbaService.visible()` 追加宿主全局符号；解析 `Worksheet`/`Workbook`/`Range` 等类型与成员时与普通类型索引一致 |
| 边界 | 近似模型，不覆盖全部成员；非 Excel 宿主不注入；不替代 `references` 的显式版本约束 |

## 12. 0.2.0 新增模块

| 模块 | 职责 | 关键约束 |
|---|---|---|
| `src/comDiscovery.ts` + `scripts/discover-com.ps1` | 按 `vbadminx.com.softwareList` 只读扫描 `HKCR\\TypeLib`，多选后批量调用 `export-com.ps1` 写入 `.vba/type-indexes/`，可选更新 `vba-project.json` 的 `typeIndexes` | 仅 Windows；不注册类型库、不执行组件；PowerShell 参数逐项传递（`shell: false`） |
| `src/i18n.ts` + `src/i18n/catalog.ts` | 五语言文案与 `auto` 解析；`package.nls.*.json` 负责清单本地化 | 文案回调不依赖语言服务器；服务器诊断暂不本地化 |
| `src/snippetFile.ts` + `src/snippets.ts` + `snippets/vba.json` | 内置片段经 `contributes.snippets`；自定义片段按设置路径读取并以补全项注入 | 片段文件允许 JSONC 注释；保存后按 mtime 重新读取；不可执行任意代码 |
| `src/formDesign.ts` + `src/formHtml.ts` + `src/formPreview.ts` | 解析 `.frm` 设计段并近似渲染到 WebView，编辑时 250 ms 防抖热更新 | 只读；不解析 `.frx`；HTML 文本与属性均转义；CSP 仅允许内联样式与 nonce 脚本 |
| `src/formInfer.ts`（0.2.2） | 无设计段的 `.frm`：扫描 `With Me` / `With .控件`、属性赋值与事件过程，推断控件类型与布局 | 近似推断，预览顶部显示提示；只读；完整设计请导出含设计段的 `.frm` |
| `src/symbolsView.ts` | 打开 VBA 工程自动展开一次；标题栏齿轮打开扩展设置 | 每个扩展会话只强制展开一次 |
