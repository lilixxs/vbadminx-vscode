const { randomUUID } = require('node:crypto');
const { completeChat } = require('../dist/src/goClient.js');

async function main() {
  const direct = process.argv.includes('--direct');
  const apiKey = direct ? process.env.OPENCODE_GO_API_KEY : 'magpie';
  if (!apiKey) { throw new Error('direct 模式需要 OPENCODE_GO_API_KEY。'); }
  const connection = {
    baseUrl: direct ? 'https://opencode.ai/inference/go/openai/v1' : 'http://127.0.0.1:3425/v1',
    model: direct ? 'deepseek-v4.1-flash' : 'opencode-go/deepseek-v4.1-flash',
    apiKey,
    sessionId: `ses_${randomUUID().replaceAll('-', '')}`,
  };
  const messages = [{ role: 'user', content: 'Use get_vba_context once to read the synthetic selection. Then explain in one short sentence what the function does.' }];
  const tools = [{ type: 'function', function: {
    name: 'get_vba_context',
    description: 'Read the synthetic VBA selection for this connectivity test.',
    parameters: { type: 'object', properties: { mode: { type: 'string', enum: ['selection'] } }, required: ['mode'], additionalProperties: false },
  } }];
  const first = await completeChat({ ...connection, messages, tools, toolChoice: 'required', maxTokens: 1024 }, () => {});
  if (first.toolCalls.length !== 1 || first.toolCalls[0].name !== 'get_vba_context' || first.toolCalls[0].input.mode !== 'selection') {
    throw new Error('模型未返回预期工具调用。');
  }
  const call = first.toolCalls[0];
  messages.push({
    role: 'assistant', content: first.text || null,
    tool_calls: [{ id: call.id, type: 'function', function: { name: call.name, arguments: JSON.stringify(call.input) } }],
    ...(first.reasoning ? { reasoning_content: first.reasoning } : {}),
  });
  messages.push({ role: 'tool', tool_call_id: call.id, content: JSON.stringify({ code: 'Public Function AddNumbers(a As Long, b As Long) As Long\nAddNumbers = a + b\nEnd Function', synthetic: true }) });
  const second = await completeChat({ ...connection, messages, maxTokens: 512 }, () => {});
  if (!second.text.trim()) { throw new Error('工具结果回传后没有文本响应。'); }
  console.log(JSON.stringify({ connection: direct ? 'direct' : 'magpie', model: connection.model, toolCall: call.name, roundTrip: 'passed', response: second.text }, null, 2));
}

main().catch(error => { console.error(error.message); process.exitCode = 1; });
