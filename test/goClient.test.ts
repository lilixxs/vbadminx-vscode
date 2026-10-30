import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createServer, type IncomingHttpHeaders, type IncomingMessage, type ServerResponse } from 'node:http';
import { completeChat, type CompletionRequest, type ToolDefinition } from '../src/goClient.js';

interface MockServer {
  baseUrl: string;
  close: () => Promise<void>;
}

async function startMock(
  handler: (req: IncomingMessage, res: ServerResponse) => void | Promise<void>,
): Promise<MockServer> {
  const server = createServer((req, res) => {
    req.on('error', () => {});
    res.on('error', () => {});
    Promise.resolve()
      .then(() => handler(req, res))
      .catch(() => {
        if (!res.headersSent) {
          try {
            res.writeHead(500);
          } catch {
            // 连接已销毁。
          }
        }
        try {
          res.end();
        } catch {
          // 连接已销毁。
        }
      });
  });

  await new Promise<void>((resolve, reject) => {
    const onError = (error: Error): void => reject(error);
    server.once('error', onError);
    server.listen(0, '127.0.0.1', () => {
      server.off('error', onError);
      resolve();
    });
  });

  const address = server.address();
  if (address === null || typeof address === 'string') {
    throw new Error('测试服务器未取得端口');
  }

  return {
    baseUrl: `http://127.0.0.1:${address.port}/v1`,
    close: async () => {
      server.closeAllConnections();
      await new Promise<void>((resolve) => server.close(() => resolve()));
    },
  };
}

async function readRequestBody(req: IncomingMessage): Promise<string> {
  const chunks: Buffer[] = [];
  for await (const chunk of req) {
    chunks.push(typeof chunk === 'string' ? Buffer.from(chunk, 'utf8') : Buffer.from(chunk as Uint8Array));
  }
  return Buffer.concat(chunks).toString('utf8');
}

function sseEvent(payload: unknown): string {
  return `data: ${JSON.stringify(payload)}\n\n`;
}

function writeSplit(res: ServerResponse, payload: Buffer, cutPoints: number[]): void {
  const points = [...new Set(cutPoints)].filter((point) => point > 0 && point < payload.length).sort((a, b) => a - b);
  let start = 0;
  for (const point of points) {
    res.write(payload.subarray(start, point));
    start = point;
  }
  res.write(payload.subarray(start));
  res.end();
}

function makeRequest(baseUrl: string, overrides: Partial<CompletionRequest> = {}): CompletionRequest {
  return {
    baseUrl,
    apiKey: 'test-key',
    model: 'test-model',
    sessionId: 'session-1',
    messages: [{ role: 'user', content: '你好' }],
    ...overrides,
  };
}

const toolDefinition: ToolDefinition = {
  type: 'function',
  function: {
    name: 'get_vba_context',
    description: '读取 VBA 上下文',
    parameters: { type: 'object', properties: { mode: { type: 'string' } }, required: ['mode'] },
  },
};

const wait = (ms: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, ms));

test('请求头/请求体正确，SSE 文本、推理与工具调用跨分片累计', { timeout: 10_000 }, async () => {
  let capturedHeaders: IncomingHttpHeaders | undefined;
  let capturedMethod = '';
  let capturedUrl = '';
  let capturedBody: unknown;

  const mock = await startMock(async (req, res) => {
    capturedHeaders = req.headers;
    capturedMethod = req.method ?? '';
    capturedUrl = req.url ?? '';
    capturedBody = JSON.parse(await readRequestBody(req)) as unknown;

    const text = [
      ': keepalive\r\n\r\n',
      `data: ${JSON.stringify({ id: 'x', choices: [{ index: 0, delta: { role: 'assistant', reasoning_content: '先分析' } }] })}\r\n\r\n`,
      `data: ${JSON.stringify({ choices: [{ index: 0, delta: { content: '你好' } }] })}\r\n\r\n`,
      'data: {"choices":\n',
      `data: ${JSON.stringify([{ index: 0, delta: { content: '世界' } }])}}\r\n\r\n`,
      `data: ${JSON.stringify({
        choices: [
          {
            index: 0,
            delta: {
              tool_calls: [
                { index: 0, id: 'call_', type: 'function', function: { name: 'get_vba_', arguments: '{"mode":' } },
              ],
            },
          },
        ],
      })}\r\n\r\n`,
      `data: ${JSON.stringify({
        choices: [
          {
            index: 0,
            delta: { tool_calls: [{ index: 0, id: '1', function: { name: 'context', arguments: '"selection"}' } }] },
          },
        ],
      })}\r\n\r\n`,
      `data: ${JSON.stringify({ choices: [{ index: 0, delta: {}, finish_reason: 'tool_calls' }] })}\r\n\r\n`,
      `data: ${JSON.stringify({ choices: [], usage: { prompt_tokens: 3, completion_tokens: 5, total_tokens: 8 } })}\r\n\r\n`,
      'data: [DONE]\r\n\r\n',
    ].join('');

    res.writeHead(200, { 'content-type': 'text/event-stream; charset=utf-8' });
    const raw = Buffer.from(text, 'utf8');
    const splitPoints = [
      raw.indexOf(Buffer.from('你')) + 1,
      raw.indexOf(Buffer.from('世界')) + 2,
      raw.indexOf(Buffer.from('get_vba_')) + 3,
    ];
    writeSplit(res, raw, splitPoints);
  });

  try {
    const textChunks: string[] = [];
    const result = await completeChat(
      makeRequest(mock.baseUrl, { tools: [toolDefinition], toolChoice: 'auto' }),
      (chunk) => textChunks.push(chunk),
    );

    assert.equal(capturedMethod, 'POST');
    assert.equal(capturedUrl, '/v1/chat/completions');
    assert.ok(capturedHeaders !== undefined);
    assert.equal(capturedHeaders['authorization'], 'Bearer test-key');
    assert.equal(capturedHeaders['x-opencode-session'], 'session-1');
    assert.equal(capturedHeaders['user-agent'], 'vba-coder-prototype/0.0.1');
    assert.match(String(capturedHeaders['content-type']), /application\/json/);
    assert.match(String(capturedHeaders['accept']), /text\/event-stream/);

    assert.ok(capturedBody !== null && typeof capturedBody === 'object');
    const body = capturedBody as Record<string, unknown>;
    assert.equal(body['model'], 'test-model');
    assert.equal(body['stream'], true);
    assert.equal(body['max_tokens'], 4096);
    assert.deepEqual(body['messages'], [{ role: 'user', content: '你好' }]);
    assert.deepEqual(body['tools'], [toolDefinition]);
    assert.equal(body['tool_choice'], 'auto');

    assert.deepEqual(textChunks, ['你好', '世界']);
    assert.equal(result.text, '你好世界');
    assert.equal(result.reasoning, '先分析');
    assert.deepEqual(result.toolCalls, [{ id: 'call_1', name: 'get_vba_context', input: { mode: 'selection' } }]);
  } finally {
    await mock.close();
  }
});

test('DeepSeek 强制工具调用关闭 Thinking，并保留 maxTokens 与 toolChoice', { timeout: 10_000 }, async () => {
  let capturedBody: unknown;

  const mock = await startMock(async (req, res) => {
    capturedBody = JSON.parse(await readRequestBody(req)) as unknown;
    res.writeHead(200, { 'content-type': 'text/event-stream' });
    res.write(
      sseEvent({
        choices: [
          {
            index: 0,
            delta: {
              tool_calls: [
                { index: 0, id: 'call_9', type: 'function', function: { name: 'get_vba_context', arguments: '{}' } },
              ],
            },
          },
        ],
      }),
    );
    res.write(sseEvent({ choices: [{ index: 0, delta: {}, finish_reason: 'tool_calls' }] }));
    res.end('data: [DONE]\n\n');
  });

  try {
    const result = await completeChat(
      makeRequest(mock.baseUrl, {
        model: 'opencode-go/deepseek-v4.1-flash',
        tools: [toolDefinition],
        toolChoice: 'required',
        maxTokens: 128,
      }),
      () => {},
    );

    assert.ok(capturedBody !== null && typeof capturedBody === 'object');
    const body = capturedBody as Record<string, unknown>;
    assert.equal(body['max_tokens'], 128);
    assert.equal(body['tool_choice'], 'required');
    assert.deepEqual(body['thinking'], { type: 'disabled' });

    assert.equal(result.text, '');
    assert.equal(result.reasoning, undefined);
    assert.deepEqual(result.toolCalls, [{ id: 'call_9', name: 'get_vba_context', input: {} }]);
  } finally {
    await mock.close();
  }
});

test('预取消的 signal 不发起任何请求', { timeout: 10_000 }, async () => {
  let hits = 0;
  const mock = await startMock((_req, res) => {
    hits += 1;
    res.writeHead(200, { 'content-type': 'text/event-stream' });
    res.end('data: [DONE]\n\n');
  });

  try {
    const controller = new AbortController();
    controller.abort();
    await assert.rejects(completeChat(makeRequest(mock.baseUrl), () => {}, controller.signal), (error: unknown) => {
      assert.ok(error instanceof Error);
      assert.equal(error.name, 'AbortError');
      return true;
    });
    await wait(50);
    assert.equal(hits, 0);
  } finally {
    await mock.close();
  }
});

test('运行中取消会中止读取并保留已收到的正文', { timeout: 10_000 }, async () => {
  let pendingResponse: ServerResponse | undefined;
  const mock = await startMock((_req, res) => {
    pendingResponse = res;
    res.writeHead(200, { 'content-type': 'text/event-stream' });
    res.write(sseEvent({ choices: [{ index: 0, delta: { content: '开头' } }] }));
    // 不结束响应，等待客户端取消。
  });

  try {
    const controller = new AbortController();
    const chunks: string[] = [];
    const pending = completeChat(
      makeRequest(mock.baseUrl),
      (chunk) => {
        chunks.push(chunk);
        controller.abort();
      },
      controller.signal,
    );

    await assert.rejects(pending, (error: unknown) => {
      assert.ok(error instanceof Error);
      assert.equal(error.name, 'AbortError');
      return true;
    });
    assert.deepEqual(chunks, ['开头']);
    pendingResponse?.end();
  } finally {
    pendingResponse?.end();
    await mock.close();
  }
});

test('非 2xx 只暴露状态码，不回显响应正文或密钥', { timeout: 10_000 }, async () => {
  const secret = 'sk-live-secret-123';
  const source = 'Attribute VB_Name = "秘密源码"';
  const mock = await startMock((_req, res) => {
    res.writeHead(401, { 'content-type': 'application/json' });
    res.end(JSON.stringify({ error: { message: secret }, source }));
  });

  try {
    await assert.rejects(completeChat(makeRequest(mock.baseUrl), () => {}), (error: unknown) => {
      assert.ok(error instanceof Error);
      assert.match(error.message, /401/);
      assert.equal(error.message.includes(secret), false);
      assert.equal(error.message.includes(source), false);
      assert.equal(error.message.includes('test-key'), false);
      return true;
    });
  } finally {
    await mock.close();
  }
});

test('HTTP 200 的 JSON 错误信封明确失败且不回显正文', { timeout: 10_000 }, async () => {
  const secret = 'sk-live-secret-456';
  const mock = await startMock((_req, res) => {
    res.writeHead(200, { 'content-type': 'application/json' });
    res.end(JSON.stringify({ error: { message: secret, type: 'invalid_request_error' } }));
  });

  try {
    await assert.rejects(completeChat(makeRequest(mock.baseUrl), () => {}), (error: unknown) => {
      assert.ok(error instanceof Error);
      assert.match(error.message, /错误响应/);
      assert.equal(error.message.includes(secret), false);
      return true;
    });
  } finally {
    await mock.close();
  }
});

test('SSE error 载荷明确失败且不回显正文', { timeout: 10_000 }, async () => {
  const secret = 'sk-live-secret-789';
  const mock = await startMock((_req, res) => {
    res.writeHead(200, { 'content-type': 'text/event-stream' });
    res.end(`data: ${JSON.stringify({ error: { message: secret } })}\n\n`);
  });

  try {
    await assert.rejects(completeChat(makeRequest(mock.baseUrl), () => {}), (error: unknown) => {
      assert.ok(error instanceof Error);
      assert.match(error.message, /错误响应/);
      assert.equal(error.message.includes(secret), false);
      return true;
    });
  } finally {
    await mock.close();
  }
});

test('没有 [DONE] 且没有 finish_reason 视为截断', { timeout: 10_000 }, async () => {
  const mock = await startMock((_req, res) => {
    res.writeHead(200, { 'content-type': 'text/event-stream' });
    res.write(sseEvent({ choices: [{ index: 0, delta: { content: '半截' } }] }));
    res.end();
  });

  try {
    await assert.rejects(completeChat(makeRequest(mock.baseUrl), () => {}), /未收到完成标记/);
  } finally {
    await mock.close();
  }
});

test('finish_reason=stop 但没有 [DONE] 仍视为完成', { timeout: 10_000 }, async () => {
  const mock = await startMock((_req, res) => {
    res.writeHead(200, { 'content-type': 'text/event-stream' });
    res.write(sseEvent({ choices: [{ index: 0, delta: { content: 'ok' } }] }));
    res.write(sseEvent({ choices: [{ index: 0, delta: {}, finish_reason: 'stop' }] }));
    res.end();
  });

  try {
    const result = await completeChat(makeRequest(mock.baseUrl), () => {});
    assert.equal(result.text, 'ok');
    assert.deepEqual(result.toolCalls, []);
  } finally {
    await mock.close();
  }
});

test('截断、过滤、上游中断与资源不足均明确失败', { timeout: 10_000 }, async () => {
  const cases = [
    { finishReason: 'length', pattern: /max_tokens/ },
    { finishReason: 'content_filter', pattern: /内容过滤器/ },
    { finishReason: 'aborted', pattern: /上游中断/ },
    { finishReason: 'insufficient_system_resource', pattern: /上游中断/ },
  ] as const;

  for (const item of cases) {
    const mock = await startMock((_req, res) => {
      res.writeHead(200, { 'content-type': 'text/event-stream' });
      res.write(sseEvent({ choices: [{ index: 0, delta: {}, finish_reason: item.finishReason }] }));
      res.end('data: [DONE]\n\n');
    });
    try {
      await assert.rejects(completeChat(makeRequest(mock.baseUrl), () => {}), item.pattern);
    } finally {
      await mock.close();
    }
  }
});

test('工具参数必须是合法 JSON 对象', { timeout: 10_000 }, async () => {
  const cases = [
    { args: '{"mode":', pattern: /不是合法 JSON/ },
    { args: '[]', pattern: /必须是 JSON 对象/ },
    { args: 'null', pattern: /必须是 JSON 对象/ },
    { args: '"text"', pattern: /必须是 JSON 对象/ },
  ] as const;

  for (const item of cases) {
    const mock = await startMock((_req, res) => {
      res.writeHead(200, { 'content-type': 'text/event-stream' });
      res.write(
        sseEvent({
          choices: [
            {
              index: 0,
              delta: {
                tool_calls: [
                  { index: 0, id: 'call_1', type: 'function', function: { name: 'fn', arguments: item.args } },
                ],
              },
            },
          ],
        }),
      );
      res.write(sseEvent({ choices: [{ index: 0, delta: {}, finish_reason: 'tool_calls' }] }));
      res.end('data: [DONE]\n\n');
    });
    try {
      await assert.rejects(completeChat(makeRequest(mock.baseUrl), () => {}), item.pattern);
    } finally {
      await mock.close();
    }
  }
});

test('不支持的 content 结构与非法 SSE JSON 明确失败', { timeout: 10_000 }, async () => {
  const arrayContentMock = await startMock((_req, res) => {
    res.writeHead(200, { 'content-type': 'text/event-stream' });
    res.write(sseEvent({ choices: [{ index: 0, delta: { content: [{ type: 'text', text: 'hi' }] } }] }));
    res.write(sseEvent({ choices: [{ index: 0, delta: {}, finish_reason: 'stop' }] }));
    res.end('data: [DONE]\n\n');
  });
  try {
    await assert.rejects(completeChat(makeRequest(arrayContentMock.baseUrl), () => {}), /不支持的 content/);
  } finally {
    await arrayContentMock.close();
  }

  const badJsonMock = await startMock((_req, res) => {
    res.writeHead(200, { 'content-type': 'text/event-stream' });
    res.end('data: {not-json}\n\n');
  });
  try {
    await assert.rejects(completeChat(makeRequest(badJsonMock.baseUrl), () => {}), /非法 JSON/);
  } finally {
    await badJsonMock.close();
  }
});

test('拒绝不安全或非法的 baseUrl', { timeout: 10_000 }, async () => {
  const unsafe = [
    '',
    '   ',
    'http://example.com/v1',
    'http://localhost.evil.com/v1',
    'https://user:pass@example.com/v1',
    'https://example.com/v1?x=1',
    'https://example.com/v1#frag',
    'ftp://example.com/v1',
    'not-a-url',
  ];

  for (const baseUrl of unsafe) {
    await assert.rejects(completeChat(makeRequest(baseUrl), () => {}), /baseUrl/);
  }
});

test('不跟随重定向，避免把 Authorization 转发到其他主机', { timeout: 10_000 }, async () => {
  let targetHits = 0;
  const target = await startMock((_req, res) => {
    targetHits += 1;
    res.writeHead(200, { 'content-type': 'text/event-stream' });
    res.end('data: [DONE]\n\n');
  });
  const origin = await startMock((_req, res) => {
    res.writeHead(302, { location: `${target.baseUrl}/chat/completions` });
    res.end();
  });

  try {
    await assert.rejects(completeChat(makeRequest(origin.baseUrl), () => {}), /请求失败/);
    await wait(100);
    assert.equal(targetHits, 0);
  } finally {
    await origin.close();
    await target.close();
  }
});
