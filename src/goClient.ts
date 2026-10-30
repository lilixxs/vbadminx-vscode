/**
 * 独立 HTTP 客户端：OpenAI Chat Completions 兼容协议（SSE 流式）。
 *
 * 仅依赖 Node 22 原生 fetch。Magpie 本地地址与官方 Go 使用同一协议。
 * 本模块不读取凭据、不记录请求体与响应体，错误信息不包含密钥、原文或响应正文。
 */

export interface ChatMessage {
  role: 'system' | 'user' | 'assistant' | 'tool';
  content: string | null;
  tool_calls?: Array<{
    id: string;
    type: 'function';
    function: { name: string; arguments: string };
  }>;
  tool_call_id?: string;
  reasoning_content?: string;
}

export interface ToolDefinition {
  type: 'function';
  function: { name: string; description: string; parameters: object };
}

export interface CompletionRequest {
  baseUrl: string;
  apiKey: string;
  model: string;
  sessionId: string;
  messages: ChatMessage[];
  tools?: ToolDefinition[];
  toolChoice?: 'auto' | 'required';
  maxTokens?: number;
}

export interface CompletionResult {
  text: string;
  toolCalls: Array<{ id: string; name: string; input: object }>;
  reasoning?: string;
}

const DEFAULT_MAX_TOKENS = 4096;
const REQUEST_TIMEOUT_MS = 60_000;
const USER_AGENT = 'vba-coder-prototype/0.0.1';

/** 本模块主动产生的协议错误；外部传输错误会被归一化为同类错误，避免泄露原始信息。 */
class ChatClientError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ChatClientError';
  }
}

function abortError(): Error {
  const error = new Error('请求已取消');
  error.name = 'AbortError';
  return error;
}

function timeoutError(): Error {
  const error = new Error('请求超时（60 秒）');
  error.name = 'TimeoutError';
  return error;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

/** 仅允许 https；http 仅限本机回环地址。拒绝 userinfo、查询串与片段。 */
function buildChatCompletionsUrl(baseUrl: string): string {
  if (typeof baseUrl !== 'string' || baseUrl.trim().length === 0) {
    throw new ChatClientError('baseUrl 不能为空');
  }
  const trimmed = baseUrl.trim();
  if (trimmed.includes('?') || trimmed.includes('#')) {
    throw new ChatClientError('baseUrl 不能包含查询串或片段');
  }

  let parsed: URL;
  try {
    parsed = new URL(trimmed);
  } catch {
    throw new ChatClientError('baseUrl 不是合法 URL');
  }

  if (parsed.username.length > 0 || parsed.password.length > 0) {
    throw new ChatClientError('baseUrl 不能包含用户信息');
  }

  const host = parsed.hostname.toLowerCase();
  if (parsed.protocol === 'https:') {
    // 允许任意 https 主机。
  } else if (parsed.protocol === 'http:') {
    const isLoopback = host === '127.0.0.1' || host === 'localhost' || host === '::1' || host === '[::1]';
    if (!isLoopback) {
      throw new ChatClientError('baseUrl 仅允许 https，或指向本机回环地址的 http');
    }
  } else {
    throw new ChatClientError('baseUrl 仅支持 http 或 https');
  }

  return `${trimmed.replace(/\/+$/, '')}/chat/completions`;
}

function assertHeaderValue(value: string, label: string): void {
  if (/[\r\n]/.test(value)) {
    throw new ChatClientError(`${label} 包含非法控制字符`);
  }
}

function buildHeaders(request: CompletionRequest): Record<string, string> {
  assertHeaderValue(request.apiKey, 'apiKey');
  assertHeaderValue(request.sessionId, 'sessionId');
  return {
    'content-type': 'application/json',
    accept: 'text/event-stream',
    authorization: `Bearer ${request.apiKey}`,
    'x-opencode-session': request.sessionId,
    'user-agent': USER_AGENT,
  };
}

function buildBody(request: CompletionRequest): Record<string, unknown> {
  const maxTokens = request.maxTokens ?? DEFAULT_MAX_TOKENS;
  if (!Number.isInteger(maxTokens) || maxTokens <= 0) {
    throw new ChatClientError('maxTokens 必须是正整数');
  }

  const body: Record<string, unknown> = {
    model: request.model,
    messages: request.messages,
    stream: true,
    max_tokens: maxTokens,
  };
  if (request.tools !== undefined) {
    body.tools = request.tools;
  }
  if (request.toolChoice !== undefined) {
    body.tool_choice = request.toolChoice;
  }
  // DeepSeek Thinking 模式不支持 required。保持强制工具语义，显式关闭该请求的 Thinking。
  if (request.toolChoice === 'required' && /(?:^|\/)deepseek(?:-|$)/i.test(request.model)) {
    body.thinking = { type: 'disabled' };
  }
  return body;
}

interface ToolCallAccumulator {
  id: string;
  name: string;
  arguments: string;
}

interface StreamState {
  text: string;
  reasoning: string;
  finishReason: string | null;
  sawDone: boolean;
  toolCalls: Map<number, ToolCallAccumulator>;
}

function createStreamState(): StreamState {
  return {
    text: '',
    reasoning: '',
    finishReason: null,
    sawDone: false,
    toolCalls: new Map<number, ToolCallAccumulator>(),
  };
}

function applyToolCallDeltas(state: StreamState, rawToolCalls: unknown): void {
  if (rawToolCalls === undefined || rawToolCalls === null) {
    return;
  }
  if (!Array.isArray(rawToolCalls)) {
    throw new ChatClientError('接口返回了不支持的 tool_calls 结构');
  }

  for (const rawCall of rawToolCalls) {
    if (!isRecord(rawCall)) {
      throw new ChatClientError('接口返回了不支持的 tool_calls 结构');
    }
    const index = rawCall.index;
    if (typeof index !== 'number' || !Number.isInteger(index) || index < 0) {
      throw new ChatClientError('工具调用缺少有效的 index');
    }
    let accumulator = state.toolCalls.get(index);
    if (accumulator === undefined) {
      accumulator = { id: '', name: '', arguments: '' };
      state.toolCalls.set(index, accumulator);
    }

    const id = rawCall.id;
    if (id !== undefined && id !== null) {
      if (typeof id !== 'string') {
        throw new ChatClientError('接口返回了不支持的 tool_calls.id');
      }
      accumulator.id += id;
    }

    const type = rawCall.type;
    if (type !== undefined && type !== null && type !== 'function') {
      throw new ChatClientError('不支持的工具调用类型');
    }

    const fn = rawCall.function;
    if (fn !== undefined && fn !== null) {
      if (!isRecord(fn)) {
        throw new ChatClientError('接口返回了不支持的 tool_calls.function');
      }
      const name = fn.name;
      if (name !== undefined && name !== null) {
        if (typeof name !== 'string') {
          throw new ChatClientError('接口返回了不支持的函数名');
        }
        accumulator.name += name;
      }
      const args = fn.arguments;
      if (args !== undefined && args !== null) {
        if (typeof args !== 'string') {
          throw new ChatClientError('接口返回了不支持的函数参数结构');
        }
        accumulator.arguments += args;
      }
    }
  }
}

function applyDelta(
  state: StreamState,
  delta: Record<string, unknown>,
  onText: (text: string) => void,
): void {
  const content = delta.content;
  if (content !== undefined && content !== null) {
    if (typeof content !== 'string') {
      throw new ChatClientError('收到不支持的 content 结构');
    }
    if (content.length > 0) {
      state.text += content;
      onText(content);
    }
  }

  // reasoning_content 只进入 result.reasoning，不交给 onText。
  const reasoning = delta.reasoning_content;
  if (reasoning !== undefined && reasoning !== null) {
    if (typeof reasoning !== 'string') {
      throw new ChatClientError('收到不支持的 reasoning_content 结构');
    }
    state.reasoning += reasoning;
  }

  applyToolCallDeltas(state, delta.tool_calls);
}

function applyPayload(
  raw: string,
  state: StreamState,
  onText: (text: string) => void,
): void {
  let payload: unknown;
  try {
    payload = JSON.parse(raw);
  } catch {
    throw new ChatClientError('接口返回了非法 JSON 数据');
  }
  if (!isRecord(payload)) {
    throw new ChatClientError('接口返回了非对象 JSON 数据');
  }
  if (payload.error !== undefined && payload.error !== null) {
    throw new ChatClientError('接口返回了错误响应');
  }

  const choices = payload.choices;
  if (choices === undefined || choices === null) {
    return;
  }
  if (!Array.isArray(choices)) {
    throw new ChatClientError('接口返回了不支持的 choices 结构');
  }
  const choice = choices[0];
  if (choice === undefined) {
    return;
  }
  if (!isRecord(choice)) {
    throw new ChatClientError('接口返回了不支持的 choice 结构');
  }

  const finishReason = choice.finish_reason;
  if (finishReason !== undefined && finishReason !== null) {
    if (typeof finishReason !== 'string') {
      throw new ChatClientError('接口返回了不支持的 finish_reason');
    }
    state.finishReason = finishReason;
  }

  const delta = choice.delta;
  if (delta === undefined || delta === null) {
    return;
  }
  if (!isRecord(delta)) {
    throw new ChatClientError('接口返回了不支持的 delta 结构');
  }
  applyDelta(state, delta, onText);
}

function finalizeToolCalls(state: StreamState): CompletionResult['toolCalls'] {
  const ordered = [...state.toolCalls.entries()].sort((a, b) => a[0] - b[0]);
  return ordered.map(([index, accumulator]) => {
    if (accumulator.id.length === 0) {
      throw new ChatClientError(`工具调用 #${index} 缺少 id`);
    }
    if (accumulator.name.length === 0) {
      throw new ChatClientError(`工具调用 #${index} 缺少函数名`);
    }
    let parsed: unknown;
    try {
      parsed = JSON.parse(accumulator.arguments);
    } catch {
      throw new ChatClientError(`工具调用 #${index} 的参数不是合法 JSON`);
    }
    if (!isRecord(parsed)) {
      throw new ChatClientError(`工具调用 #${index} 的参数必须是 JSON 对象`);
    }
    return { id: accumulator.id, name: accumulator.name, input: parsed };
  });
}

interface LineBreak {
  index: number;
  length: number;
}

function findLineBreak(text: string): LineBreak | null {
  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    if (char === '\n') {
      return { index: i, length: 1 };
    }
    if (char === '\r') {
      if (i + 1 >= text.length) {
        // 可能是跨 chunk 的 CRLF，等待下一块。
        return null;
      }
      return { index: i, length: text[i + 1] === '\n' ? 2 : 1 };
    }
  }
  return null;
}

/** 处理没有 SSE 字段前缀的 JSON（HTTP 200 错误信封或非流式响应）。不泄露正文。 */
function rejectRawJson(line: string): never {
  let parsed: unknown;
  try {
    parsed = JSON.parse(line);
  } catch {
    throw new ChatClientError('接口返回了无法解析的响应');
  }
  if (isRecord(parsed) && parsed.error !== undefined && parsed.error !== null) {
    throw new ChatClientError('接口返回了错误响应（HTTP 200）');
  }
  throw new ChatClientError('接口返回了非流式响应');
}

async function readStream(
  reader: ReadableStreamDefaultReader<Uint8Array>,
  onText: (text: string) => void,
  abortSignal: AbortSignal,
  isTimedOut: () => boolean,
): Promise<CompletionResult> {
  const decoder = new TextDecoder('utf-8');
  const state = createStreamState();
  let buffer = '';
  let dataLines: string[] = [];
  let eventName = '';

  const dispatch = (): void => {
    const name = eventName;
    eventName = '';
    if (dataLines.length === 0) {
      return;
    }
    const data = dataLines.join('\n');
    dataLines = [];
    if (data.length === 0) {
      return;
    }
    if (name === 'error') {
      throw new ChatClientError('接口返回了错误事件');
    }
    if (data.trim() === '[DONE]') {
      state.sawDone = true;
      return;
    }
    applyPayload(data, state, onText);
  };

  const handleLine = (line: string): void => {
    if (line.length === 0) {
      dispatch();
      return;
    }
    if (line.startsWith(':')) {
      return;
    }
    if (line.startsWith('{')) {
      rejectRawJson(line);
    }
    const colon = line.indexOf(':');
    const field = colon === -1 ? line : line.slice(0, colon);
    let value = colon === -1 ? '' : line.slice(colon + 1);
    if (value.startsWith(' ')) {
      value = value.slice(1);
    }
    if (field === 'data') {
      dataLines.push(value);
    } else if (field === 'event') {
      eventName = value;
    }
    // id/retry 及未知字段按 SSE 规范忽略。
  };

  let ended = false;
  while (!ended && !state.sawDone) {
    let result: ReadableStreamReadResult<Uint8Array>;
    try {
      result = await reader.read();
    } catch (error) {
      if (abortSignal.aborted) {
        throw isTimedOut() ? timeoutError() : abortError();
      }
      throw error;
    }

    if (result.done) {
      buffer += decoder.decode();
      ended = true;
    } else {
      buffer += decoder.decode(result.value, { stream: true });
    }

    let lineBreak = findLineBreak(buffer);
    while (lineBreak !== null) {
      const line = buffer.slice(0, lineBreak.index);
      buffer = buffer.slice(lineBreak.index + lineBreak.length);
      handleLine(line);
      if (state.sawDone) {
        break;
      }
      lineBreak = findLineBreak(buffer);
    }
  }

  if (!state.sawDone) {
    // EOF 时容忍没有末尾空行的最后一条事件。
    if (buffer.length > 0) {
      const tail = buffer.endsWith('\r') ? buffer.slice(0, -1) : buffer;
      buffer = '';
      for (const line of tail.split(/\r\n|\n|\r/)) {
        handleLine(line);
        if (state.sawDone) {
          break;
        }
      }
    }
    if (!state.sawDone) {
      dispatch();
    }
  }

  if (state.finishReason === 'length') {
    throw new ChatClientError('响应因达到 max_tokens 上限被截断');
  }
  if (state.finishReason === 'content_filter') {
    throw new ChatClientError('响应被内容过滤器拦截');
  }
  if (state.finishReason === 'aborted' || state.finishReason === 'insufficient_system_resource') {
    throw new ChatClientError('上游中断了响应，不能将部分结果视为完成');
  }
  if (!state.sawDone && state.finishReason === null) {
    throw new ChatClientError('响应流提前结束：未收到完成标记');
  }

  const result: CompletionResult = {
    text: state.text,
    toolCalls: finalizeToolCalls(state),
  };
  if (state.reasoning.length > 0) {
    result.reasoning = state.reasoning;
  }
  return result;
}

async function discardBody(response: Response): Promise<void> {
  try {
    await response.body?.cancel();
  } catch {
    // 忽略清理失败。
  }
}

export async function completeChat(
  request: CompletionRequest,
  onText: (text: string) => void,
  signal?: AbortSignal,
): Promise<CompletionResult> {
  if (signal?.aborted) {
    throw abortError();
  }

  const url = buildChatCompletionsUrl(request.baseUrl);
  const headers = buildHeaders(request);
  const body = JSON.stringify(buildBody(request));

  const controller = new AbortController();
  let timedOut = false;
  const timer = setTimeout(() => {
    timedOut = true;
    controller.abort();
  }, REQUEST_TIMEOUT_MS);
  const forwardAbort = (): void => {
    controller.abort();
  };
  if (signal !== undefined) {
    signal.addEventListener('abort', forwardAbort, { once: true });
  }

  let reader: ReadableStreamDefaultReader<Uint8Array> | null = null;
  try {
    const response = await fetch(url, {
      method: 'POST',
      headers,
      body,
      redirect: 'error',
      signal: controller.signal,
    });

    if (!response.ok) {
      const status = response.status;
      await discardBody(response);
      throw new ChatClientError(`接口返回 HTTP ${status}`);
    }

    const contentType = (response.headers.get('content-type') ?? '').toLowerCase();
    if (contentType.includes('application/json')) {
      let raw = '';
      try {
        raw = await response.text();
      } catch {
        raw = '';
      }
      let parsed: unknown;
      try {
        parsed = JSON.parse(raw);
      } catch {
        parsed = undefined;
      }
      if (isRecord(parsed) && parsed.error !== undefined && parsed.error !== null) {
        throw new ChatClientError('接口返回了错误响应（HTTP 200）');
      }
      throw new ChatClientError('接口返回了非流式 JSON 响应');
    }

    if (response.body === null) {
      throw new ChatClientError('接口响应缺少可读的流式正文');
    }
    reader = response.body.getReader();
    return await readStream(reader, onText, controller.signal, () => timedOut);
  } catch (error) {
    if (controller.signal.aborted) {
      throw timedOut ? timeoutError() : abortError();
    }
    if (error instanceof ChatClientError) {
      throw error;
    }
    // 外部错误可能携带 URL/主机信息；统一归一化，避免泄露。
    throw new ChatClientError('请求失败：与接口的通信中断');
  } finally {
    clearTimeout(timer);
    if (signal !== undefined) {
      signal.removeEventListener('abort', forwardAbort);
    }
    if (reader !== null) {
      try {
        await reader.cancel();
      } catch {
        // 流已关闭或被中止。
      }
    }
  }
}
