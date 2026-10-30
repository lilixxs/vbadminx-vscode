# AI 接入原型说明（协议与安全细节）

> **状态说明**：本文保留 AI 模型接入的协议、密钥、HTTP/SSE 与安全约束细节——这些部分仍然有效。
> 其中标注为“历史”的状态描述写于语言服务器实现之前（如“本扩展无 VBA LSP”），**已过时**：
> vbadminx-vscode 0.1.0 已由 `src/server.ts` 提供首版语言功能的对应处理器，工具的 `symbols`/`references`
> 模式由语言服务器回答。**当前架构、语言功能与验证状态一律以 [docs/implementation.md](implementation.md) 为准。**

本文说明 AI 接入的实现范围、边界与安全约束。操作步骤见 [README](../README.md)，产品需求见 [docs/requirements.md](requirements.md)，术语见 [GLOSSARY.md](../GLOSSARY.md)。

## 1. 定位与边界

- 注册一个只读上下文工具 `get_vba_context`，返回选区、编辑器诊断，或**经语言服务器**取得的工程符号与确定引用。
- 注册一个模型 provider（vendor `vbadminx-go`），经 OpenAI Chat Completions 兼容协议访问模型。
- **（历史 → 现状）**“这是接入原型，不是 VBA LSP”“不含 VBA 解析器、语法/语义分析、补全、重命名、格式化”“工具返回体固定含 `semanticAnalysis: false`”等表述均为原型期状态；现状是语言服务已实现、工具返回 `semanticAnalysis: true`，见 [implementation.md](implementation.md)。
- 诊断仍来自编辑器中的诊断提供器（现含 VBAdminx 的 LSP 诊断）；空列表不证明代码正确，诊断不等同编译器结论。
- 不连接 Office、不运行宏、不改源码。
- INP-02/INP-03 的 UTF-8 检测与保存协作**已实现**，见 [implementation.md](implementation.md) 第 8 节。

## 2. 代码组成

| 文件 | 职责 |
|------|------|
| `src/extension.ts` | `activate()`：注册 provider、工具、4 个命令、配置变更监听；`guarded()` 统一错误提示 |
| `src/provider.ts` | `getConnection()` 连接解析；`GoProvider` 消息转换、会话关联、reasoning 缓存、token 估算 |
| `src/goClient.ts` | `completeChat()`：Node 22 原生 `fetch` + SSE 流式解析，无第三方运行时依赖 |
| `src/contextTool.ts` | `readContext()` 只读取逻辑；`VbaContextTool` 工具包装与确认文案 |
| `scripts/test-host.cjs` | 隔离宿主集成检查启动器 |
| `scripts/smoke-go.cjs` | 真实模型两轮工具往返冒烟 |
| `test/goClient.test.ts` | 14 项协议测试（本机模拟服务） |
| `test/host/index.ts` | 宿主内断言：激活、工具、选区、诊断、守卫、provider 往返 |
| `examples/` | `Demo.bas` 样例与两个用户设置模板 |

输出目录为 `dist/`，由 `npm run compile` 生成。

> 注：上表只列 AI 接入相关文件。`src/extension.ts` 现在还注册语言客户端、COM 命令与编码检查；语言服务器、工程索引与格式化的文件组成见 [implementation.md](implementation.md) 第 2 节。

## 3. 扩展注册内容（`package.json`）

- 语言：`vba`，扩展名 `.bas`、`.cls`、`.frm`；语法 `syntaxes/vba.tmLanguage.json`；`[vba]` 默认 4 空格与默认格式化器。
- 激活事件：`onLanguage:vba`、`onLanguageModelChatProvider:vbadminx-go`。
- `capabilities.untrustedWorkspaces.supported: false`：工作区不信任时不激活。
- `extensionKind: ["workspace"]`：在扩展主机（可能为远程主机）运行。
- 模型 provider：vendor `vbadminx-go`，显示名「VBAdminx · OpenCode Go」，管理命令 `vbadminx.configureGo`。
- 工具：`get_vba_context`，`toolReferenceName: vbaContext`，`canBeReferencedInPrompt: true`，`when: isWorkspaceTrusted`；输入模式枚举已扩为 `selection | diagnostics | symbols | references`（`symbols` 可带 `query`），`modelDescription` 也已按四种模式更新。
- AI 相关配置 4 项，全部 `scope: "application"`：`vbadminx.go.enabled`、`connection`（`magpie`/`direct`）、`baseUrl`、`model`。另有 `vbadminx.format.*`、`vbadminx.diagnostics.*`（工作区作用域，供语言服务器使用）与 `vba-project.json` 的 JSON 校验、9 个贡献命令，见 [implementation.md](implementation.md)。

设置为 application scope，因此必须写入**用户设置**；workspace settings 不生效。模板：`examples/user-settings.magpie.json`、`examples/user-settings.direct.json`。

## 4. 模型接入

### 4.1 连接解析（`getConnection()`，`src/provider.ts`）

| 项 | `magpie`（默认） | `direct` |
|----|------------------|----------|
| `baseUrl` | `vbadminx.go.baseUrl`，默认 `http://127.0.0.1:3425/v1` | 固定 `https://opencode.ai/inference/go/openai/v1` |
| `model` | `vbadminx.go.model`，默认 `opencode-go/deepseek-v4.1-flash` | 固定 `deepseek-v4.1-flash`（不带 `opencode-go/` 前缀） |
| `apiKey` | 固定占位值 `magpie`，**不是**用户密钥 | SecretStorage 键 `openCodeGoApiKey`，否则 `OPENCODE_GO_API_KEY` |

- 密钥不写入任何设置文件。
- `vbadminx.configureGo`：输入框保存密钥到 SecretStorage，并把 `connection` 切到 `direct`（Global 目标），随后 `provider.refresh()`。
- `vbadminx.clearGoKey`：只删除 SecretStorage，不删除环境变量；界面提示明确说明。
- 不复用其他工具的登录 token，不从 OpenCode 认证文件读取凭据。
- `vbadminx.go.enabled: false` 时 `provideLanguageModelChatInformation()` 返回空数组，模型不出现在选择器中。
- 连接信息读取失败时同样返回空数组，不向用户暴露底层错误细节。

### 4.2 模型信息（`provideLanguageModelChatInformation()`）

| 字段 | 值 |
|------|----|
| `id` | 当前连接的 `model` |
| `name` | `DeepSeek V4.1 Flash · OpenCode Go` |
| `family` | `deepseek-v4.1-flash` |
| `version` | `prototype-1` |
| `maxInputTokens` / `maxOutputTokens` | `16000` / `4096` |
| `capabilities` | `toolCalling: true`，`imageInput: false` |
| `detail` | 原型：文本与工具调用 |
| `tooltip` | 保守上下文上限与 UTF-8 字节估算，不代表官方最大上下文 |

这些是**原型预算**，不是模型官方最大上下文。配置变更后 `onDidChangeLanguageModelChatInformation` 触发刷新；若模型 ID 与当前连接不一致，请求报错并要求重新选择模型。

### 4.3 响应流程（`provideLanguageModelChatResponse()`）

1. 信任检查：工作区不信任直接抛错。
2. 消息转换：`convertMessages()`。
3. 会话关联：`sessionKey = SHA-256(JSON(首条 user 消息))`，映射到 `ses_<uuid>`，经 `x-opencode-session` 请求头发送。
4. 有工具时，`toolMode === Required` 对应 `tool_choice: "required"`，否则为 `auto`。无工具时不发送 `tool_choice`。
5. 文本增量经 `progress` 报告为 `LanguageModelTextPart`；工具调用报告为 `LanguageModelToolCallPart`。
6. 取消经 `AbortController` 转发；取消时抛 `CancellationError`。

**会话启发式限制**：

- 会话按首条 user 消息哈希关联。两个首条消息相同的不同对话可能共享同一后端 session。
- 缓存仅在内存中，扩展宿主重启后 session 变化。
- 这是启发式关联，不是宿主提供的会话标识。

**reasoning 处理**：

- SSE 中的 `reasoning_content` 不输出给聊天界面，只进入 `result.reasoning`。
- 按工具调用 `id` 存入内存 Map，上限 **128 条**，超限按插入顺序淘汰。
- 仅用于后续轮次工具调用消息的 `reasoning_content` 续传；重启即丢失。

**token 估算（`provideTokenCount()`）**：无模型 tokenizer，用 UTF-8 字节数（字符串或 `JSON.stringify` 整条消息）作保守估计，避免英文四字符假设低估中文。这是估算，不是官方计数。

### 4.4 消息转换限制（`convertMessages()`）

- 只接受 `LanguageModelTextPart`、`LanguageModelToolCallPart`、`LanguageModelToolResultPart`。
- 纯文本与工具消息之外的未知 part **抛错拒绝**，不静默忽略。不支持图片。
- 工具结果只接受文本 part，否则抛「原型仅支持文本工具结果」。
- `tool_calls` 必须位于 assistant 消息，否则抛错。

## 5. 只读工具 `get_vba_context`

### 5.1 输入

`{ "mode": "selection" | "diagnostics" | "symbols" | "references", "query"?: string }`，`mode` 必填，`additionalProperties: false`；`query` 仅用于 `symbols` 模式过滤符号名。

### 5.2 守卫（`readContext()`）

按顺序检查：

1. 取消令牌已取消 → `CancellationError`。
2. `vscode.workspace.isTrusted` 为 false → 抛错，要求先信任工作区。
3. `mode` 非法 → 抛错。
4. 活动编辑器为空，或文件扩展名不是 `.bas`/`.cls`/`.frm` → 抛错。
5. 活动文件不属于当前工作区文件夹 → 抛错（工作区外单文件被拒绝）。

`symbols`/`references` 模式先复用上述检查（以 `diagnostics` 路径触发），再经语言客户端向服务器发 `vba/context` 请求；语言服务器未启动时抛错。

### 5.3 返回体

公共字段：`semanticAnalysis: true`、`file`（相对路径）、`documentVersion`。

| 模式 | 内容 | 上限 |
|------|------|------|
| `selection` | `range`（1 起行号）、`code`、`truncated` | 选区非空要求；超过 **12000 字符**截断并置 `truncated: true` |
| `diagnostics` | `note`、`diagnostics[]`、`truncated` | 最多 **100 条**；`message` 截到 1000 字符；`truncated: true` 表示被裁剪 |
| `symbols` / `references` | `{ semanticAnalysis, mode, result }`，`result` 来自服务器 | 序列化后超 **24000 字符**时整体替换为 `{ truncated: true, message: "结果过多，请缩小查询。" }` |

诊断项字段：`line`、`character`（均 1 起）、`severity`、`message`、`source`。`note` 现为：「这些诊断来自编辑器中的诊断提供器，包括 VBAdminx。空列表不证明代码完全正确；未知类型可能限制分析。」（原型期“本原型未实现 VBA LSP”措辞已随实现更新。）

### 5.4 确认与只读性

- `prepareInvocation()` 返回确认文案：按 `mode` 分别询问「读取选中代码？」或「读取已有诊断？」，并说明结果将交给当前聊天模型。
- 工具只读：不写文件、不修改源码、不触发保存。

## 6. 命令（AI 相关）

| 命令 ID | 行为 | 网络 |
|---------|------|------|
| `vbadminx.previewContext` | `readContext({mode:'selection'})` 结果作为 JSON 文档在旁侧打开 | 无 |
| `vbadminx.testGo` | 信任检查 → 模态确认 → 进度条内发送一条合成英文 VBA 问题（`maxTokens: 128`），结果写入输出通道「VBAdminx」 | 有 |
| `vbadminx.configureGo` | 保存密钥并切 `direct` | 无 |
| `vbadminx.clearGoKey` | 删除 SecretStorage 密钥 | 无 |

`vbadminx.testGo` 只发送合成问题，不发送工作区代码；确认框文案明确写出这一点。其余 6 个贡献命令（工程索引、工程配置、COM 导入/导出/提取）与内部命令 `vbadminx.previewEdit` 见 [implementation.md](implementation.md)。

## 7. HTTP/SSE 客户端约束（`src/goClient.ts`）

- 依赖：Node 22 原生 `fetch`，无第三方运行时包。
- URL：仅 `https`（任意主机）或 `http`（仅 `127.0.0.1`、`localhost`、`::1`）；拒绝 userinfo、查询串、片段；末尾去斜杠后拼 `/chat/completions`。
- 请求头：`authorization: Bearer <key>`、`x-opencode-session`、`accept: text/event-stream`、`user-agent: vbadminx-vscode-prototype/0.0.1`。
- 请求体：`model`、`messages`、`stream: true`、`max_tokens`（默认 4096）、可选 `tools` 与 `tool_choice`。
- DeepSeek 的 Thinking 模式会拒绝强制工具选择。`tool_choice: required` 时显式发送 `thinking: { type: "disabled" }`，保留强制调用语义。该组合已通过本机 Magpie 路由实测。
- **60 秒**超时；外部取消信号经 `AbortController` 转发；`redirect: 'error'` 不跟随重定向（避免 `Authorization` 被转发到其他主机）。
- SSE 解析：跨 chunk 累计 `tool_calls` 分片；忽略注释行与未知字段；`[DONE]` 为完成标记（EOF 时容忍缺尾空行的最后事件）。
- 明确失败：`finish_reason` 为 `length`、`content_filter`、`aborted` 或 `insufficient_system_resource`；无完成标记；HTTP 200 错误信封；`event: error`；非流式响应；工具参数非 JSON 对象。无完成标记指既无 `[DONE]`，也无 `finish_reason`。
- 错误归一化：HTTP 错误只暴露状态码；传输错误统一为「请求失败：与接口的通信中断」；错误信息不含密钥、请求体、响应正文、用户源码。

### 最小请求示例

Magpie 模式发送 `POST http://127.0.0.1:3425/v1/chat/completions`。请求头使用 `Authorization: Bearer magpie`、`Content-Type: application/json`、`x-opencode-session: ses_<当前对话标识>` 和 `User-Agent: vbadminx-vscode-prototype/0.0.1`。

```json
{
  "model": "opencode-go/deepseek-v4.1-flash",
  "messages": [{ "role": "user", "content": "请调用工具读取我选中的 VBA 代码。" }],
  "stream": true,
  "max_tokens": 4096,
  "thinking": { "type": "disabled" },
  "tool_choice": "required",
  "tools": [{
    "type": "function",
    "function": {
      "name": "get_vba_context",
      "description": "读取当前工作区中活动 VBA 文件的选中代码。",
      "parameters": {
        "type": "object",
        "properties": { "mode": { "type": "string", "enum": ["selection"] } },
        "required": ["mode"],
        "additionalProperties": false
      }
    }
  }]
}
```

官方直连将地址改为 `https://opencode.ai/inference/go/openai/v1/chat/completions`，将模型改为 `deepseek-v4.1-flash`，并使用自己的 Go API Key。仅开通订阅或登录 OpenCode 不等于已取得直连 API Key。

模型返回工具调用后，由 VS Code 聊天宿主确认并执行工具。模型提供器只转换协议，不自行执行工具。完整两轮请求示例见 `scripts/smoke-go.cjs`；该脚本使用合成工具结果，不访问编辑器或用户源码。

## 8. 安全与隐私

- HTTP 客户端不记录请求体与响应体，不写日志文件。连接验证在输出通道显示合成问题的模型回答；冒烟脚本在终端显示合成示例的结果。
- 连接验证与 `smoke:go` 只发送合成 VBA 例子。
- 密钥只存 SecretStorage 或环境变量；设置文件、示例模板中均无密钥字段（`magpie` 的 `magpie` 是固定占位值）。
- 删除 SecretStorage 不删除 `OPENCODE_GO_API_KEY`；两者互不影响。
- 不读取 OpenCode 认证文件，不复用任何登录 token。
- 工具调用前有工作区信任检查与用户确认框；工具本身不改源码。
- `npm run smoke:go` 真实访问模型，**计入订阅额度**；普通测试使用本机模拟服务，不消耗额度。

## 9. 远程环境

- `extensionKind: workspace`：扩展在扩展主机运行。远程 SSH 场景下即在远程主机。
- `baseUrl` 中的回环地址指向**扩展主机**的回环，不指向 Windows 本机。
- 不自动桥接 Windows 代理。需要时由用户自行做 SSH 端口转发或在远程主机配置代理。
- 文档不引导把服务放开到公网。
- 三平台与远程环境**未实测**（见第 11 节）。

## 10. 与需求的对照

| 需求 | 现状 |
|------|------|
| AI-01 只读上下文工具（符号、引用、诊断、用户选择的代码） | 已实现：四种模式齐备，`symbols`/`references` 经语言服务器 `vba/context` 回答（宿主测试项 `aiSymbols` 通过） |
| AI-02 只访问当前工作区，遵循宿主权限与信任流程 | 已实现：工作区成员检查、`isTrusted` 检查、调用前确认框 |
| AI-03 语言服务和只读工具不管理模型密钥、不自建聊天界面 | 已遵守：只读工具不接触密钥；可选模型适配器按 AI-06 管理独立凭据 |
| AI-04 不承诺兼容全部第三方 AI | 已遵守：需聊天宿主支持模型 provider 与工具调用 |
| AI-05 AI 修改代码走宿主审阅；只读工具不改源码 | 已遵守：工具只读，无任何写文件路径 |
| AI-06 开发验证用模型适配器 | 已写入需求文档：默认复用 Magpie，官方直连使用独立 API Key，不读取 OpenCode 登录凭据 |

## 11. 验证状态

> 语言功能与整体验证矩阵见 [implementation.md](implementation.md) 第 10 节。此处只记 AI 接入相关项。

### 已通过（本机 Windows 实测）

| 项 | 结果 |
|----|------|
| `npm test` 中 `goClient.test.ts` | 14 项协议测试通过（单测总数与 LSP 部分见 implementation.md） |
| `npm run test:host`（VS Code 1.143.1） | 模型发现、工具注册、选区/诊断读取、`aiSymbols`、取消与工作区守卫、provider 两轮模拟服务往返、`vscode.lm.sendRequest` 通过。当前完整插件的 31 项宿主检查见 `output/vscode-api-check.json` |
| `npm run smoke:go`（Magpie 真实模型） | 此前两轮 `get_vba_context` 工具调用及结果解释通过；接入协议未变 |

`test:host` 每次使用新建用户配置目录，并输出实际 VS Code 版本。打开 VBA 文件前先调用 `vscode.lm.selectChatModels()`，验证模型发现与扩展激活。本机模拟 SSE 服务随后完成两轮往返，检查会话头一致、推理内容续传、工具结果 ID 对齐。另经已注册模型的 `sendRequest()` 收到流式文本。

宿主检查还覆盖适配器禁用后返回空模型列表，以及预先取消的 provider 请求不访问网络并返回 `CancellationError`。

### 未验证（不得笼统声称通过）

- 官方 `direct` 模式未测试：无独立 API key，`https://opencode.ai/inference/go/openai/v1` 未实连。
- macOS、Linux、远程（SSH/WSL/dev container）未实测。
- 聊天 UI 端到端未人工验收（模型选择器、`#vbaContext` 引用、确认框在真实聊天中的表现）。
- 不受信任工作区的禁用声明已配置，尚未单独启动受限宿主验收。
- VSIX 产物 `output/vbadminx-vscode-0.1.0.vsix` 已打包，并安装到**隔离 extensions 目录**经 installed-vsix 截图复验（补全与 `VBAdminx:` 命令可用）；未改动用户日常 VS Code 配置。
- 未编造已安装的其他 AI 扩展，兼容性仅对已声明路径负责。

## 12. 不包含

- OpenCode 独立 CLI 不能直接消费 VS Code 注册的工具。MCP 接入是后续独立任务，**不在本范围**。
- 聊天 UI、模型计费、额度管理。
- 从宿主复用登录凭据。
- （历史）“VBA LSP 未实现”已失效：解析、诊断、补全、导航、重命名、格式化现由 `src/server.ts` 提供，能力与限制见 [implementation.md](implementation.md)。

## 13. 参考（官方文档）

- VS Code AI 工具：https://code.visualstudio.com/api/extension-guides/ai/tools
- VS Code 语言模型聊天 provider：https://code.visualstudio.com/api/extension-guides/ai/language-model-chat-provider
- OpenCode Go：https://opencode.ai/v2/docs/console/go
- OpenCode Harnesses：https://opencode.ai/v2/docs/console/harnesses
- OpenCode API：https://opencode.ai/v2/docs/console/api
- DeepSeek 工具选择与 Thinking 参数：https://api-docs.deepseek.com/api/create-chat-completion

不引用套餐额度与价格（易过时，且未实测）。
