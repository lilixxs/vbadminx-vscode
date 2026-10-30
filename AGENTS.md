# 项目指令

## 目标

- 本工程用于开发 VS Code 扩展，为 VBA 语言提供 LSP（语言服务器协议）支持。

## 需求与术语

- 设计、实现或验收功能前，阅读 [首版需求](docs/requirements.md)。功能范围、平台限制与验收目标以该文档为准。
- 编写需求或命名领域对象时，使用 [领域术语](GLOSSARY.md)。

## 当前实现与验证

- `src/server.ts` 通过 LSP 提供语言功能。`src/language/` 包含解析、工程索引、符号服务和格式化。解析器是容错结构分析器，不能将诊断等同于 VBA 编译器结论。
- 模型接入与远程主机配置见 [AI 接入原型](docs/ai-prototype.md)。Magpie 使用带 `opencode-go/` 前缀的模型 ID；官方 Go 直连使用不带该前缀的 ID。
- 使用 Node.js 22。`npm ci` 后运行 `npm run typecheck`、`npm test`。Windows PowerShell 若拦截 `npm.ps1`，使用 `npm.cmd`。
- `npm run test:host` 在隔离配置的 VS Code 宿主中运行。可用 `VSCODE_EXECUTABLE_PATH` 指定新版可执行文件，避免 PATH 中旧版 `code`。
- 宿主测试会输出实际 VS Code 版本。未指定可执行文件且没有标准安装时，下载 1.104.3，以匹配 `engines` 与 `@types/vscode` 的 1.104 API 基线。
- `npm run smoke:go` 访问真实模型并消耗订阅额度。它只发送合成示例；普通测试使用本机模拟服务。
- `src/extension.ts` 注册命令和 AI 工具，`src/languageClient.ts` 启动 IPC 服务器并确认跨文件修改。工具不得绕过工作区与信任检查。
- 工程以 `vba-project.json` 隔离。文件 URI 必须经 `canonicalUri()` 归一化，避免 Windows 盘符与转义差异产生重复模块。
- `npm run compile` 先执行 TypeScript 编译，再生成 `dist/extension.bundle.js` 和 `dist/server.bundle.js`。F5 与 VSIX 使用 bundle，修改源码后必须重编译。
- `npm run package` 将 VSIX 写入 `output/`。`npm run capture` 保存 VS Code 截图，`npm run test:host` 生成 `output/vscode-api-check.json`。
- COM 元数据提取仅在 Windows 运行 `scripts/export-com.ps1`。发布包必须包含该脚本；其他平台只读取可移植类型索引。

## Langium 评估结论（2026-10-29）

- 已完成调研与三项 PoC（容错、条件编译+位置映射、性能），证据与数据见 [评估文档](docs/langium-evaluation.md)，PoC 代码在 `poc/langium/`。
- 结论：**不用 Langium 替换现有解析器**。容错与条件编译均验证可行，但冷启动索引（1000 文件/10 万行）约 4.7 秒，为当前实现（约 1.1 秒）的 4.3 倍，且口径不含进程启动与 JSON-RPC，很可能不满足 PERF-02。
- 主实现不引用 Langium；`langium`/`langium-cli` 仅为 PoC 的 devDependency，PoC 已排除出发布包。若未来需求演进到完整表达式类型推导，可基于该文档重新评估。
