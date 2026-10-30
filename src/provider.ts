import { createHash, randomUUID } from 'node:crypto';
import * as vscode from 'vscode';
import { completeChat, type ChatMessage } from './goClient';

export const VENDOR = 'vbadminx-go';
export const KEY_NAME = 'openCodeGoApiKey';

export interface Connection {
  baseUrl: string;
  model: string;
  apiKey: string;
}

export async function getConnection(context: vscode.ExtensionContext): Promise<Connection> {
  const config = vscode.workspace.getConfiguration('vbadminx.go');
  if (config.get<string>('connection', 'magpie') === 'direct') {
    const key = await context.secrets.get(KEY_NAME) || process.env.OPENCODE_GO_API_KEY;
    if (!key?.trim()) { throw new Error('先执行“VBAdminx: 配置 OpenCode Go API Key”，或设置 OPENCODE_GO_API_KEY。'); }
    return {
      baseUrl: 'https://opencode.ai/inference/go/openai/v1',
      model: 'deepseek-v4.1-flash',
      apiKey: key.trim(),
    };
  }
  return {
    baseUrl: config.get<string>('baseUrl', 'http://127.0.0.1:3425/v1'),
    model: config.get<string>('model', 'opencode-go/deepseek-v4.1-flash'),
    // 本机代理配置使用固定占位值；它不是官方 Go API Key。
    apiKey: 'magpie',
  };
}

export class GoProvider implements vscode.LanguageModelChatProvider, vscode.Disposable {
  private readonly changed = new vscode.EventEmitter<void>();
  readonly onDidChangeLanguageModelChatInformation = this.changed.event;
  private readonly sessions = new Map<string, string>();
  private readonly reasoning = new Map<string, string>();

  constructor(private readonly context: vscode.ExtensionContext) {}

  refresh(): void { this.changed.fire(); }
  dispose(): void { this.changed.dispose(); }

  async provideLanguageModelChatInformation(): Promise<vscode.LanguageModelChatInformation[]> {
    if (!vscode.workspace.getConfiguration('vbadminx.go').get<boolean>('enabled', true)) { return []; }
    try {
      const connection = await getConnection(this.context);
      return [{
        id: connection.model,
        name: 'DeepSeek V4.1 Flash · OpenCode Go',
        family: 'deepseek-v4.1-flash',
        version: 'prototype-1',
        maxInputTokens: 16_000,
        maxOutputTokens: 4096,
        detail: '原型：文本与工具调用',
        tooltip: '原型使用保守的上下文上限和 UTF-8 字节数估算，不代表模型的官方最大上下文。',
        capabilities: { toolCalling: true, imageInput: false },
      }];
    } catch { return []; }
  }

  async provideLanguageModelChatResponse(
    model: vscode.LanguageModelChatInformation,
    messages: readonly vscode.LanguageModelChatRequestMessage[],
    options: vscode.ProvideLanguageModelChatResponseOptions,
    progress: vscode.Progress<vscode.LanguageModelResponsePart>,
    token: vscode.CancellationToken,
  ): Promise<void> {
    if (!vscode.workspace.isTrusted) { throw new Error('先信任工作区，再发送模型请求。'); }
    if (!vscode.workspace.getConfiguration('vbadminx.go').get<boolean>('enabled', true)) { throw new Error('Go 模型适配器已禁用。'); }
    const connection = await getConnection(this.context);
    if (model.id !== connection.model) { throw new Error('模型配置已更改，请重新选择模型。'); }
    const converted = this.convertMessages(messages);
    const firstMessage = converted.find(item => item.role === 'user') ?? converted[0];
    const sessionKey = createHash('sha256').update(JSON.stringify(firstMessage ?? {})).digest('hex');
    let sessionId = this.sessions.get(sessionKey);
    if (!sessionId) {
      sessionId = `ses_${randomUUID().replaceAll('-', '')}`;
      this.remember(this.sessions, sessionKey, sessionId);
    }
    const abort = new AbortController();
    const subscription = token.onCancellationRequested(() => abort.abort());
    if (token.isCancellationRequested) { abort.abort(); }
    try {
      const result = await completeChat({
        ...connection,
        sessionId,
        messages: converted,
        tools: options.tools?.length ? options.tools.map(tool => ({
          type: 'function',
          function: { name: tool.name, description: tool.description, parameters: tool.inputSchema ?? { type: 'object', properties: {} } },
        })) : undefined,
        toolChoice: options.tools?.length
          ? options.toolMode === vscode.LanguageModelChatToolMode.Required ? 'required' : 'auto'
          : undefined,
      }, text => progress.report(new vscode.LanguageModelTextPart(text)), abort.signal);
      for (const call of result.toolCalls) {
        if (result.reasoning) { this.remember(this.reasoning, call.id, result.reasoning); }
        progress.report(new vscode.LanguageModelToolCallPart(call.id, call.name, call.input));
      }
    } catch (error) {
      if (token.isCancellationRequested) { throw new vscode.CancellationError(); }
      throw error;
    } finally { subscription.dispose(); }
  }

  async provideTokenCount(_model: vscode.LanguageModelChatInformation, text: string | vscode.LanguageModelChatRequestMessage): Promise<number> {
    // 无模型 tokenizer。用 UTF-8 字节数作保守估算，避免英文四字符假设低估中文。
    return Buffer.byteLength(typeof text === 'string' ? text : JSON.stringify(text), 'utf8');
  }

  private convertMessages(messages: readonly vscode.LanguageModelChatRequestMessage[]): ChatMessage[] {
    const converted: ChatMessage[] = [];
    for (const message of messages) {
      if (message.role !== vscode.LanguageModelChatMessageRole.User && message.role !== vscode.LanguageModelChatMessageRole.Assistant) {
        throw new Error('原型收到不支持的消息角色。');
      }
      const text: string[] = [];
      const calls: NonNullable<ChatMessage['tool_calls']> = [];
      const results: ChatMessage[] = [];
      for (const part of message.content) {
        if (part instanceof vscode.LanguageModelTextPart) { text.push(part.value); }
        else if (part instanceof vscode.LanguageModelToolCallPart) {
          calls.push({ id: part.callId, type: 'function', function: { name: part.name, arguments: JSON.stringify(part.input) } });
        } else if (part instanceof vscode.LanguageModelToolResultPart) {
          const resultText = part.content.map(value => {
            if (!(value instanceof vscode.LanguageModelTextPart)) { throw new Error(`工具调用 ${part.callId} 含非文本结果。原型仅支持文本工具结果。`); }
            return value.value;
          }).join('\n');
          results.push({ role: 'tool', tool_call_id: part.callId, content: resultText });
        } else {
          throw new Error('原型仅支持文本和工具消息，不能忽略未知消息类型。');
        }
      }
      if (results.length > 0 && message.role !== vscode.LanguageModelChatMessageRole.User) {
        throw new Error('工具结果必须位于 user 消息。');
      }
      // OpenAI 要求工具结果紧跟前一条 assistant 工具调用，再发送新用户文本。
      converted.push(...results);
      if (text.length > 0 || calls.length > 0) {
        const role = message.role === vscode.LanguageModelChatMessageRole.Assistant ? 'assistant' : 'user';
        const output: ChatMessage = { role, content: text.join('\n') || null };
        if (calls.length > 0) {
          if (role !== 'assistant') { throw new Error('工具调用必须位于 assistant 消息。'); }
          output.tool_calls = calls;
          const reasoning = this.reasoning.get(calls[0]!.id);
          if (reasoning) { output.reasoning_content = reasoning; }
        }
        converted.push(output);
      }
    }
    return converted;
  }

  private remember(map: Map<string, string>, key: string, value: string): void {
    if (map.size >= 128) { map.delete(map.keys().next().value!); }
    map.set(key, value);
  }
}
