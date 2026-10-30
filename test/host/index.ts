import assert from 'node:assert/strict';
import { createServer } from 'node:http';
import * as vscode from 'vscode';
import { readContext } from '../../src/contextTool';
import { GoProvider } from '../../src/provider';
import { runExtended } from './extended';

export async function run(): Promise<void> {
  console.log(`HOST VERSION: ${vscode.version}`);
  const coldModels = await vscode.lm.selectChatModels({ vendor: 'vbadminx-go' });
  assert.ok(coldModels.some(model => model.id === 'opencode-go/deepseek-v4.1-flash'), '打开 VBA 文件前可以通过提供器发现模型');
  const extension = vscode.extensions.getExtension<{ provider: GoProvider; client: import('vscode-languageclient/node').LanguageClient }>('vbadminx.vbadminx-vscode');
  assert.ok(extension, '开发扩展必须存在');
  const api = await extension.activate();
  const status = await api.client.sendRequest<{ modules: number }>('vba/projectStatus');
  assert.ok(status.modules >= 2, '语言服务器索引示例工程');
  assert.ok(vscode.lm.tools.some(tool => tool.name === 'get_vba_context'), '工具已注册');
  const discovered = await vscode.lm.selectChatModels({ vendor: 'vbadminx-go' });
  assert.ok(discovered.some(model => model.id === 'opencode-go/deepseek-v4.1-flash'), 'VS Code 可以发现提供器模型');
  const commands = await vscode.commands.getCommands(true);
  assert.ok(commands.includes('vbadminx.testGo'));
  for (const command of ['vbadminx.importTypeIndex', 'vbadminx.exportTypeIndex', 'vbadminx.extractComTypeIndex', 'vbadminx.projectStatus', 'vbadminx.createProject']) { assert.ok(commands.includes(command), `命令已注册：${command}`); }
  const folder = vscode.workspace.workspaceFolders?.[0];
  assert.ok(folder);
  const document = await vscode.workspace.openTextDocument(vscode.Uri.joinPath(folder.uri, 'Demo.bas'));
  const editor = await vscode.window.showTextDocument(document);
  const caller = await vscode.workspace.openTextDocument(vscode.Uri.joinPath(folder.uri, 'Caller.bas'));
  await vscode.window.showTextDocument(caller);
  // 与请求队列同步，确保 didOpen 文档已送到服务器。
  await api.client.sendRequest('vba/projectStatus');
  const definition = await vscode.commands.executeCommand<vscode.Location[]>('vscode.executeDefinitionProvider', caller.uri, new vscode.Position(4, 16));
  assert.ok(definition?.some(location => location.uri.path.endsWith('/Demo.bas')), '跨文件定义跳转');
  const references = await vscode.commands.executeCommand<vscode.Location[]>('vscode.executeReferenceProvider', document.uri, new vscode.Position(2, 20));
  assert.ok(references && references.length >= 2, '跨文件引用');
  const completions = await vscode.commands.executeCommand<vscode.CompletionList>('vscode.executeCompletionItemProvider', caller.uri, new vscode.Position(4, 13));
  assert.ok(completions?.items.some(item => (typeof item.label === 'string' ? item.label : item.label.label) === 'AddNumbers'), '工程符号补全');
  const hover = await vscode.commands.executeCommand<vscode.Hover[]>('vscode.executeHoverProvider', caller.uri, new vscode.Position(4, 16));
  assert.ok(hover?.length, '悬停信息');
  const symbols = await vscode.commands.executeCommand<vscode.DocumentSymbol[]>('vscode.executeDocumentSymbolProvider', document.uri);
  assert.ok(symbols?.some(symbol => symbol.name === 'AddNumbers'), '文档符号');
  const signatures = await vscode.commands.executeCommand<vscode.SignatureHelp>('vscode.executeSignatureHelpProvider', caller.uri, new vscode.Position(4, 25));
  assert.ok(signatures?.signatures.length, '签名提示');
  const folds = await vscode.commands.executeCommand<vscode.FoldingRange[]>('vscode.executeFoldingRangeProvider', document.uri);
  assert.ok(folds?.length, '代码折叠');
  const semantic = await vscode.commands.executeCommand<vscode.SemanticTokens>('vscode.provideDocumentSemanticTokens', document.uri);
  assert.ok(semantic?.data.length, '语义高亮');
  const rename = await api.client.sendRequest<{ changes: Record<string, unknown[]> }>('textDocument/rename', { textDocument: { uri: document.uri.toString() }, position: { line: 2, character: 20 }, newName: 'SumNumbers' });
  assert.ok(Object.keys(rename.changes).length === 2, '确定引用重命名覆盖两个文件');
  const allSymbols = await vscode.commands.executeCommand<vscode.SymbolInformation[]>('vscode.executeWorkspaceSymbolProvider', 'AddNumbers');
  assert.ok(allSymbols?.length, '工作区符号');
  const formatEdits = await vscode.commands.executeCommand<vscode.TextEdit[]>('vscode.executeFormatDocumentProvider', caller.uri, { tabSize: 4, insertSpaces: true });
  assert.ok(formatEdits === undefined || Array.isArray(formatEdits), '已格式化文档允许返回空编辑');
  const rangeEdits = await vscode.commands.executeCommand<vscode.TextEdit[]>('vscode.executeFormatRangeProvider', caller.uri, new vscode.Range(2, 0, 5, 0), { tabSize: 4, insertSpaces: true });
  assert.ok(rangeEdits === undefined || Array.isArray(rangeEdits), '已格式化选区允许返回空编辑');
  const aiSymbols = await api.client.sendRequest<unknown[]>('vba/context', { uri: caller.uri.toString(), mode: 'symbols', query: 'AddNumbers' });
  assert.ok(aiSymbols.length, 'AI 符号上下文来自语言服务器');
  await vscode.window.showTextDocument(document);
  const extendedChecks = await runExtended(api, folder.uri);
  const original = caller.getText();
  try {
    const callerEditor = await vscode.window.showTextDocument(caller);
    await callerEditor.edit(builder => builder.replace(new vscode.Range(3, 0, 3, caller.lineAt(3).text.length), 'dim result as long'));
    const formatting = await vscode.commands.executeCommand<vscode.TextEdit[]>('vscode.executeFormatDocumentProvider', caller.uri, { tabSize: 4, insertSpaces: true });
    assert.ok(formatting?.length, '未格式化代码应返回修改');
    const formattedEdit = new vscode.WorkspaceEdit();
    formattedEdit.set(caller.uri, formatting);
    assert.equal(await vscode.workspace.applyEdit(formattedEdit), true);
    assert.equal(caller.lineAt(3).text, '    Dim result As Long', '实际应用格式化修正缩进与关键字');
    await callerEditor.edit(builder => builder.insert(new vscode.Position(3, 0), '    missingValue = 1\n'));
    const deadline = Date.now() + 5000;
    while (!vscode.languages.getDiagnostics(caller.uri).some(item => item.code === 'undeclared') && Date.now() < deadline) { await new Promise(resolve => setTimeout(resolve, 50)); }
    assert.ok(vscode.languages.getDiagnostics(caller.uri).some(item => item.code === 'undeclared'), '增量未声明符号诊断');
    const actions = await vscode.commands.executeCommand<vscode.CodeAction[]>('vscode.executeCodeActionProvider', caller.uri, new vscode.Range(3, 4, 3, 16));
    assert.ok(actions?.some(action => action.title.includes('missingValue')), '局部声明快速修复');
  } finally {
    const callerEditor = await vscode.window.showTextDocument(caller);
    await callerEditor.edit(builder => builder.replace(new vscode.Range(caller.positionAt(0), caller.positionAt(caller.getText().length)), original));
  }
  await vscode.window.showTextDocument(document);
  const source = new vscode.CancellationTokenSource();
  const diagnostics = vscode.languages.createDiagnosticCollection('prototype-host-test');
  let server: ReturnType<typeof createServer> | undefined;
  const config = vscode.workspace.getConfiguration('vbadminx.go');
  const previousUrl = config.inspect<string>('baseUrl')?.globalValue;
  const previousEnabled = config.inspect<boolean>('enabled')?.globalValue;
  try {
    const activeEditor = await vscode.window.showTextDocument(document);
    activeEditor.selection = new vscode.Selection(2, 0, 4, 12);
    const selected = JSON.parse(readContext({ mode: 'selection' }, source.token));
    assert.equal(selected.semanticAnalysis, true);
    assert.match(selected.code, /AddNumbers/);
    diagnostics.set(document.uri, [new vscode.Diagnostic(new vscode.Range(0, 0, 0, 6), 'synthetic diagnostic')]);
    const existing = JSON.parse(readContext({ mode: 'diagnostics' }, source.token));
    assert.ok(existing.diagnostics.some((item: { message: string }) => item.message === 'synthetic diagnostic'));
    activeEditor.selection = new vscode.Selection(0, 0, 0, 0);
    assert.throws(() => readContext({ mode: 'selection' }, source.token), /选择/);
    source.cancel();
    assert.throws(() => readContext({ mode: 'diagnostics' }, source.token));
    const untitled = await vscode.workspace.openTextDocument({ language: 'vba', content: 'Sub Outside()\nEnd Sub' });
    await vscode.window.showTextDocument(untitled);
    const tokenSource = new vscode.CancellationTokenSource();
    assert.throws(() => readContext({ mode: 'diagnostics' }, tokenSource.token), /工作区|打开/);

    const requests: Array<{ headers: Record<string, unknown>; body: any }> = [];
    server = createServer(async (request, response) => {
      let body = '';
      for await (const chunk of request) { body += chunk; }
      requests.push({ headers: request.headers, body: JSON.parse(body) });
      response.writeHead(200, { 'Content-Type': 'text/event-stream' });
      if (requests.length === 1) {
        response.write(`data: ${JSON.stringify({ choices: [{ delta: { reasoning_content: 'synthetic reasoning', tool_calls: [{ index: 0, id: 'call_host', type: 'function', function: { name: 'get_vba_context', arguments: '{"mode":"selection"}' } }] }, finish_reason: 'tool_calls' }] })}\n\n`);
      } else {
        response.write(`data: ${JSON.stringify({ choices: [{ delta: { content: 'host round trip passed' }, finish_reason: 'stop' }] })}\n\n`);
      }
      response.end('data: [DONE]\n\n');
    });
    await new Promise<void>(resolve => server!.listen(0, '127.0.0.1', resolve));
    const address = server.address();
    assert.ok(address && typeof address !== 'string');
    await config.update('baseUrl', `http://127.0.0.1:${address.port}/v1`, vscode.ConfigurationTarget.Global);
    const model = (await api.provider.provideLanguageModelChatInformation())[0];
    assert.ok(model);
    await config.update('enabled', false, vscode.ConfigurationTarget.Global);
    assert.deepEqual(await api.provider.provideLanguageModelChatInformation(), []);
    await config.update('enabled', true, vscode.ConfigurationTarget.Global);
    const messages: vscode.LanguageModelChatRequestMessage[] = [{
      role: vscode.LanguageModelChatMessageRole.User, name: undefined,
      content: [new vscode.LanguageModelTextPart('synthetic VBA tool request')],
    }];
    const parts: vscode.LanguageModelResponsePart[] = [];
    const options: vscode.ProvideLanguageModelChatResponseOptions = {
      toolMode: vscode.LanguageModelChatToolMode.Required,
      tools: [{ name: 'get_vba_context', description: 'synthetic context', inputSchema: { type: 'object' } }],
    };
    const cancelled = new vscode.CancellationTokenSource();
    cancelled.cancel();
    await assert.rejects(
      api.provider.provideLanguageModelChatResponse(model, messages, options, { report: () => {} }, cancelled.token),
      error => error instanceof vscode.CancellationError,
    );
    assert.equal(requests.length, 0, 'provider 预取消不发请求');
    cancelled.dispose();
    await api.provider.provideLanguageModelChatResponse(model, messages, options, { report: part => parts.push(part) }, tokenSource.token);
    const call = parts.find(part => part instanceof vscode.LanguageModelToolCallPart);
    assert.ok(call instanceof vscode.LanguageModelToolCallPart);
    assert.deepEqual(call.input, { mode: 'selection' });
    messages.push({ role: vscode.LanguageModelChatMessageRole.Assistant, name: undefined, content: [call] });
    messages.push({ role: vscode.LanguageModelChatMessageRole.User, name: undefined, content: [
      new vscode.LanguageModelTextPart('Explain the synthetic result.'),
      new vscode.LanguageModelToolResultPart(call.callId, [new vscode.LanguageModelTextPart('synthetic selection')]),
    ] });
    await api.provider.provideLanguageModelChatResponse(model, messages, { toolMode: vscode.LanguageModelChatToolMode.Auto }, { report: part => parts.push(part) }, tokenSource.token);
    assert.equal(requests[0]!.headers['x-opencode-session'], requests[1]!.headers['x-opencode-session']);
    assert.equal(requests[1]!.body.messages[1].reasoning_content, 'synthetic reasoning');
    assert.equal(requests[1]!.body.messages[2].tool_call_id, 'call_host');
    assert.equal(requests[1]!.body.messages[3].role, 'user', '工具结果必须先于同一输入消息中的新用户文本');
    assert.equal(requests[1]!.body.tool_choice, undefined, '没有工具时不能传 tool_choice');
    assert.ok(parts.some(part => part instanceof vscode.LanguageModelTextPart && /passed/.test(part.value)));
    const registeredModel = (await vscode.lm.selectChatModels({ vendor: 'vbadminx-go' }))[0];
    assert.ok(registeredModel);
    const registeredResponse = await registeredModel.sendRequest([
      vscode.LanguageModelChatMessage.User('Explain Option Explicit briefly.'),
    ], {}, tokenSource.token);
    let registeredText = '';
    for await (const fragment of registeredResponse.text) { registeredText += fragment; }
    assert.match(registeredText, /host round trip passed/, '经 vscode.lm 模型请求管线取得响应');
    tokenSource.dispose();
    const report = { vscodeVersion: vscode.version, passed: ['activation', 'definition', 'references', 'completion', 'hover', 'signatureHelp', 'documentSymbols', 'workspaceSymbols', 'folding', 'semanticTokens', 'rename', 'formatDocument', 'formatRange', 'incrementalDiagnostics', 'quickFix', 'aiSymbols', 'modelDiscovery', 'toolRegistration', 'selection', 'existingDiagnostics', 'cancellation', 'workspaceGuards', 'modelToolRoundTrip', 'vscode.lm.sendRequest', ...extendedChecks], failed: [] };
    const reportUri = vscode.Uri.joinPath(extension.extensionUri, 'output', 'vscode-api-check.json');
    await vscode.workspace.fs.createDirectory(vscode.Uri.joinPath(extension.extensionUri, 'output'));
    await vscode.workspace.fs.writeFile(reportUri, Buffer.from(JSON.stringify(report, null, 2) + '\n'));
    console.log(`HOST CHECKS PASSED: ${report.passed.join(', ')}`);
  } finally {
    diagnostics.dispose();
    source.dispose();
    await config.update('baseUrl', previousUrl, vscode.ConfigurationTarget.Global);
    await config.update('enabled', previousEnabled, vscode.ConfigurationTarget.Global);
    if (server) { await new Promise<void>(resolve => server!.close(() => resolve())); }
  }
}
