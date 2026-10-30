# VBAdminx

[![CI](https://github.com/lilixxs/vbadminx-vscode/actions/workflows/ci.yml/badge.svg)](https://github.com/lilixxs/vbadminx-vscode/actions/workflows/ci.yml)
[![Release](https://github.com/lilixxs/vbadminx-vscode/actions/workflows/release.yml/badge.svg)](https://github.com/lilixxs/vbadminx-vscode/actions/workflows/release.yml)

VS Code 扩展（`vbadminx-vscode` 0.2.0，显示名 **VBAdminx**）：为导出的 VBA 源文件提供**离线语言服务器（LSP）**、**COM 可移植类型索引（含从已安装软件自动查找）**、**多语言界面**、**内置/自定义代码片段**、**窗体预览**与**只读 AI 上下文工具 / Go 模型适配器**。

- 首版需求：[docs/requirements.md](docs/requirements.md) ｜ 实现与验证细节：[docs/implementation.md](docs/implementation.md)
- 功能总览与截图：[docs/features.md](docs/features.md) ｜ Langium 评估：[docs/langium-evaluation.md](docs/langium-evaluation.md)
- CI/CD（自动构建与发布）：[docs/ci-cd.md](docs/ci-cd.md)
- AI 协议与安全约束：[docs/ai-prototype.md](docs/ai-prototype.md) ｜ 术语：[GLOSSARY.md](GLOSSARY.md)
- 不连接 Office，不运行宏。只读 AI 工具不改写源码；重命名和快速修复经预览确认后修改源码。无 AI 时语言服务照常可用。

## 0.2.0 新增

- **COM 自动查找**：设置 `vbadminx.com.softwareList`（默认 Excel/Word/ADODB/Scripting/MSXML2 等），命令 `VBAdminx: 从已安装软件查找 COM 类型库` 只读扫描注册表，多选后批量提取到 `.vba/type-indexes/`，可确认加入 `typeIndexes`。仅 Windows 扩展主机。
- **符号面板增强**：打开 VBA 工程后左侧“VBA 符号”自动展开一次；标题栏新增设置齿轮（`vbadminx.openSettings`）。
- **多语言**：`vbadminx.language` = `auto` / `en` / `zh-cn` / `zh-tw` / `ja` / `ko`；`auto` 跟随 VS Code，未匹配回退英文。命令面板、视图名与设置说明随 VS Code 语言本地化。
- **代码片段**：内置 `Sub`/`If`/`For`/`Select Case`/`With`/`On Error` 等片段；`vbadminx.snippets.customPath` 可加载自定义片段文件，命令 `VBAdminx: 编辑 VBA 代码片段` 创建并打开。
- **窗体预览**：`.frm` 编辑器右上角预览按钮，WebView 近似还原 VBA 窗体并在编辑时热更新；只读，不解析 `.frx`。

## 环境要求

| 项 | 要求 |
|----|------|
| Node.js | 22（编译与 AI 客户端依赖 Node 22 原生 `fetch`） |
| npm | 随 Node 22 安装 |
| VS Code | `^1.104.0`（`package.json` engines），宿主测试实测 1.144.0；安装扩展需 ≥ 1.104 |
| 网络 | `npm ci` 需要网络；普通测试用本机模拟服务，`smoke:go` 访问模型上游 |

PowerShell 若拦截 `npm.ps1`，改用 `npm.cmd`。

## 快速开始

```powershell
npm ci
npm run typecheck
npm run compile   # tsc + esbuild 打出 dist/extension.bundle.js、dist/server.bundle.js
# 在 VS Code 中按 F5 启动扩展开发宿主
```

F5 使用 `.vscode/launch.json` 中的「运行 VBAdminx」，先执行 `npm: compile` 再打开宿主。**改源码后必须重编译**：F5 与 VSIX 都运行 bundle。

离线体验语言功能：在开发宿主打开 `examples/`（含 `vba-project.json`、`Demo.bas`、`Caller.bas`），即可试补全、悬停、跨文件跳转、重命名预览、格式化与命令面板中的 `VBAdminx:` 命令。截图见 `output/screenshots/`。

## 语言功能（已实现）

补全（含成员补全与 `With` 隐式成员）、悬停、定义跳转、引用查询、文档/工作区符号、签名提示、确定引用重命名（推断候选默认不勾选 + 差异预览后应用）、文档与选区格式化、代码折叠、语义高亮、语法与语义诊断、快速修复（统一大小写、插入局部声明，均先预览）。

- 解析器是**容错的结构级分析器，不是 VBA 编译器**：诊断不等同编译结论；无法判定的条件编译分支整段排除并给 Information；缺类型/宿主信息只提示不报确定错误。
- 工程以 `vba-project.json` 隔离（显式配置优先，无配置时识别 `模块/类模块/表格对象/窗体` 四目录布局）；文件 URI 经 `canonicalUri()` 归一，避免 Windows 盘符差异产生重复模块。
- 仅支持 UTF-8（可带 BOM）：非 UTF-8 文件暂停语义分析并提示在扩展外转换；打开时按磁盘字节保留 BOM 状态。不承诺拦截用户手动另存转码。

详见 [docs/implementation.md](docs/implementation.md)。

## 工程配置与 COM 类型索引

- `vbadminx: 创建工程配置` 生成 `vba-project.json`（字段：`name`、`host`、`targetBits`、`constants`、`modules`、`typeIndexes`、`references`；`references` 必填 `name/guid/major/minor/lcid/bitness` 显式版本）。JSON 校验绑定 `schemas/vba-project.schema.json`，示例见 `examples/vba-project.json`。
- `导入 COM 类型索引` / `导出 COM 类型索引`：全平台可用，写出无 BOM UTF-8；**导入后需手工把索引路径加入 `typeIndexes`，扩展不静默改工程配置**。
- `从 Windows 类型库提取 COM 索引`：仅 Windows 扩展主机，运行 `scripts/export-com.ps1`（`LoadTypeLibEx(REGKIND_NONE)`，不注册、不执行、不启动 Office），位数必须显式选择。其他平台只读取可移植索引。
- 当前实测报告 `output/com-check.json`：stdole2.tlb 提取成功，**13 types / 90 members**（coclass 展开后数字，以最新一次提取报告为准）。

## 模型接入

扩展通过 `languageModelChatProviders` 注册 vendor `vbadminx-go`，聊天中显示为「DeepSeek V4.1 Flash · OpenCode Go」。设置项全部 `scope: application`（**用户设置**），模板：`examples/user-settings.magpie.json`、`examples/user-settings.direct.json`。

| 模式 | 地址 | 模型 ID | 密钥 |
|------|------|---------|------|
| `magpie`（默认） | `http://127.0.0.1:3425/v1` | `opencode-go/deepseek-v4.1-flash` | 固定占位值 `magpie`，不是用户密钥 |
| `direct` | `https://opencode.ai/inference/go/openai/v1` | `deepseek-v4.1-flash`（不带前缀） | 仅 SecretStorage 或 `OPENCODE_GO_API_KEY`，不写入设置文件 |

- Magpie 使用本机现有代理，启动 VS Code 前需先运行代理。
- `配置 OpenCode Go API Key` 保存密钥并切到 `direct`；`删除保存的 OpenCode Go API Key` 只删 SecretStorage。
- 不复用其他工具登录 token，不从 OpenCode 认证文件读凭据。协议细节见 [docs/ai-prototype.md](docs/ai-prototype.md)。

## 聊天与只读工具

1. 启动 Magpie 或配置 `direct`，打开工作区内的 VBA 文件并信任工作区。
2. 聊天中选择「DeepSeek V4.1 Flash · OpenCode Go」，用 `#vbaContext` 引用工具或允许自动调用。
3. 工具 `get_vba_context` 支持四种模式：`selection`（选区，12000 字符截断）、`diagnostics`（最多 100 条）、`symbols`（工程符号，可带 `query`）、`references`（光标处确定引用）；后两者由语言服务器回答，超 24000 字符截断。调用前有确认框，工具只读。

## 命令

| 命令 ID | 标题 | 网络 |
|---------|------|------|
| `vbadminx.projectStatus` | 查看工程索引 | 无 |
| `vbadminx.createProject` | 创建工程配置 | 无 |
| `vbadminx.importTypeIndex` | 导入 COM 类型索引 | 无 |
| `vbadminx.exportTypeIndex` | 导出 COM 类型索引 | 无 |
| `vbadminx.extractComTypeIndex` | 从 Windows 类型库提取 COM 索引 | 无（本地 PowerShell） |
| `vbadminx.discoverCom` | 从已安装软件查找 COM 类型库 | 无（本地 PowerShell） |
| `vbadminx.openSettings` | 打开本扩展设置页 | 无 |
| `vbadminx.editSnippets` | 编辑 VBA 代码片段 | 无 |
| `vbadminx.previewForm` | 预览 `.frm` 窗体 | 无 |
| `vbadminx.configureGo` | 配置 OpenCode Go API Key | 无 |
| `vbadminx.clearGoKey` | 删除保存的 OpenCode Go API Key | 无 |
| `vbadminx.testGo` | 验证 Go 模型连接 | 有（合成问题） |
| `vbadminx.previewContext` | 预览只读 VBA 上下文 | 无 |

另有内部命令 `vbadminx.previewEdit`（快速修复差异预览，由诊断快速修复触发，不出现在命令面板贡献项中）。

## 脚本

| 命令 | 作用 | 网络/额度 |
|------|------|-----------|
| `npm run typecheck` | `tsc --noEmit` | 无 |
| `npm test` | 编译后运行 `dist/test/*.test.js`（本机模拟服务） | 无 |
| `npm run test:host` | 隔离配置宿主内集成检查，报告写 `output/vscode-api-check.json` | 测试请求仅发本机 |
| `npm run benchmark` | in-process 解析/语言服务 CPU 计时，写 `output/benchmark.json` | 无 |
| `npm run benchmark:lsp` | 真实 LSP 子进程 + 磁盘 + JSON-RPC 计时，写 `output/lsp-benchmark.json` | 无 |
| `npm run capture` | Playwright 截图到 `output/screenshots/` | 无 |
| `npm run preview:form` | 渲染 `examples/LoginForm.frm` 的窗体预览截图到 `output/screenshots/09-form-preview.png` | 无 |
| `npm run package` | VSIX → `output/`（文件名由 vsce 按包名与版本自动生成；`.vscodeignore` 排除 `node_modules` 与 `poc/`，保留 `scripts/export-com.ps1`、`scripts/discover-com.ps1` 与 `snippets/vba.json`） | 无 |
| `npm run verify:package` | 对比安装包、当前构建与隔离安装运行时，生成交付验收报告 | 无 |
| `npm run smoke:go` | 真实模型两轮工具往返 | **消耗订阅额度**，只发合成示例 |
| `node scripts/check-sample.cjs <导出目录>` | 只读诊断普查，写 `output/sample-check.json`（不含源码文本） | 无 |

`test:host` 与 `capture` 可用 `VSCODE_EXECUTABLE_PATH` 指定新版 VS Code，避免 PATH 中旧版 `code`；未指定且无标准安装时下载 1.104.3 对齐 API 基线。官方直连验证用 `node scripts/smoke-go.cjs --direct`（需先编译并设 `OPENCODE_GO_API_KEY`）。

## 验证状态

已通过（本机 Windows 实测）：

- `npm run typecheck` 和单元回归测试通过（134 / 134），新增覆盖窗体设计解析/渲染、片段文件与五语言文案一致性。
- `npm run test:host`（VS Code **1.144.0**）：31 项检查通过；新增命令（`discoverCom`、`openSettings`、`editSnippets`、`previewForm`）均已注册。结果见 `output/vscode-api-check.json`（`failed: []`）。
- COM 自动查找实测（Windows 注册表，只读）：`Excel;ADODB;Scripting;MSXML2` 返回 4 个类型库；对 `scrrun.dll` 提取得到 17 个类型，链路可用。
- 窗体预览、代码片段与多语言目前由单元测试与宿主命令注册检查覆盖；未做逐语言人工走查与真实 `UserForm` 截图验收。
- **VSIX 安装验证（0.1.0）**：`output/vbadminx-vscode-0.1.0.vsix` 安装到隔离 extensions 目录后，用 installed-vsix 模式重新截图验证补全与 `VBAdminx:` 命令面板（`output/screenshots/01/02/03/04`）；**未改动用户日常配置**。0.2.0 已重新构建 `output/vbadminx-vscode-0.2.0.vsix`，安装包内容由 `npm run package` 校验，未重复做隔离安装截图。
- `npm run benchmark:lsp` 已测量真实 LSP 子进程、磁盘扫描和 JSON-RPC。负载为 1000 文件、10 万行，实际数值见 `output/lsp-benchmark.json`。不含 VS Code UI 启动；本机 16 核/28 GB 与需求参考硬件不匹配，不作参考硬件达标结论。
- 只读样例普查 `output/sample-check.json`：61 文件、1 工程、0 坏编码；唯一 Error 为真实缺少 `End Sub`。未加载外部索引时会提示缺少类型和宿主信息，计数见报告。未改样例源码；样例路径见 [requirements 第 9 节](docs/requirements.md)。
- stdole 提取成功：13 types / 90 members，`New StdFont` 的 `Name`/`Size` 抽查可用（`output/com-check.json`）。
- `npm run smoke:go`（Magpie 真实模型）：此前两轮工具往返通过；接入协议未变。

未验证，不得笼统声称通过：

- 官方 `direct` 模式未实连（无独立 API Key）。
- macOS、Linux、Remote SSH/WSL/dev container 无实际环境验收（实现路径存在，见 [docs/implementation.md](docs/implementation.md)）。
- AI 真实聊天 UI 未人工验收；不受信任工作区尚未单独启动验收。
- 性能无远程主机记录（PERF-05），且无与需求一致的 4 核 16 GB 硬件复测。

## 已知限制

- 容错结构级解析 ≠ VBA 编译器结论；语义分析为结构级有限推断，不做全表达式类型检查。
- COM 索引不表达 record/alias 类型；`CreateObject`/动态对象成员在无类型索引时**不确定**，只给信息不足提示，不报确定错误。
- 窗体预览为**近似渲染**：`.frx` 二进制资源（图片、部分列表项）以占位呈现，控件视觉与系统主题存在差异；不提供设计编辑。
- 多语言覆盖扩展界面与清单本地化；语言服务器诊断与内置符号说明暂为简体中文。

## 参考

- VS Code AI 工具：https://code.visualstudio.com/api/extension-guides/ai/tools
- VS Code 语言模型聊天 provider：https://code.visualstudio.com/api/extension-guides/ai/language-model-chat-provider
- LSP 规范：https://microsoft.github.io/language-server-protocol/
- OpenCode Go：https://opencode.ai/v2/docs/console/go
- OpenCode API：https://opencode.ai/v2/docs/console/api
