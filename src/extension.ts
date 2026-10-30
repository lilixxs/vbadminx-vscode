import { randomUUID } from 'node:crypto';
import * as vscode from 'vscode';
import { readContext, VbaContextTool } from './contextTool';
import { completeChat } from './goClient';
import { getConnection, GoProvider, KEY_NAME, VENDOR } from './provider';
import { startLanguageClient } from './languageClient';
import { registerComCommands } from './com';
import type { LanguageClient } from 'vscode-languageclient/node';
import { registerEncodingChecks } from './encoding';
import { registerSymbolsView } from './symbolsView';

export async function activate(context: vscode.ExtensionContext): Promise<{ provider: GoProvider; client: LanguageClient }> {
  const provider = new GoProvider(context);
  const output = vscode.window.createOutputChannel('VBAdminx');
  const guarded = (action: () => Promise<unknown>) => async () => {
    try { return await action(); }
    catch (error) {
      if (error instanceof vscode.CancellationError || (error instanceof Error && error.name === 'AbortError')) { return; }
      const message = error instanceof Error ? error.message : '操作失败。';
      void vscode.window.showErrorMessage(`VBAdminx: ${message}`);
    }
  };
  context.subscriptions.push(
    output,
    provider,
    vscode.lm.registerLanguageModelChatProvider(VENDOR, provider),
    vscode.lm.registerTool('get_vba_context', new VbaContextTool()),
    vscode.workspace.onDidChangeConfiguration(event => {
      if (event.affectsConfiguration('vbadminx.go')) { provider.refresh(); }
    }),
    vscode.commands.registerCommand('vbadminx.configureGo', guarded(async () => {
      const key = await vscode.window.showInputBox({
        title: '保存官方 Go API Key',
        prompt: '仅 direct 模式使用。密钥保存在当前扩展主机的 SecretStorage，不写入设置文件。',
        password: true,
        ignoreFocusOut: true,
        validateInput: value => value.trim() ? undefined : '请输入 API Key。',
      });
      if (key === undefined) { return; }
      await context.secrets.store(KEY_NAME, key.trim());
      await vscode.workspace.getConfiguration('vbadminx.go').update('connection', 'direct', vscode.ConfigurationTarget.Global);
      provider.refresh();
      void vscode.window.showInformationMessage('已保存官方 Go API Key，并切换到 direct 模式。');
    })),
    vscode.commands.registerCommand('vbadminx.clearGoKey', guarded(async () => {
      await context.secrets.delete(KEY_NAME);
      provider.refresh();
      void vscode.window.showInformationMessage('已删除保存的 Go API Key。环境变量中的密钥不受影响。');
    })),
    vscode.commands.registerCommand('vbadminx.previewContext', guarded(async () => {
      const source = new vscode.CancellationTokenSource();
      try {
        const result = readContext({ mode: 'selection' }, source.token);
        const document = await vscode.workspace.openTextDocument({ language: 'json', content: result });
        await vscode.window.showTextDocument(document, { preview: true, viewColumn: vscode.ViewColumn.Beside });
      } finally { source.dispose(); }
    })),
    vscode.commands.registerCommand('vbadminx.testGo', guarded(async () => {
      if (!vscode.workspace.isTrusted) { throw new Error('先信任工作区，再验证连接。'); }
      const connection = await getConnection(context);
      const answer = await vscode.window.showInformationMessage(
        `向 ${new URL(connection.baseUrl).origin} 发送一条合成的 VBA 问题，验证 ${connection.model}？不发送工作区代码。`,
        { modal: true }, '验证连接',
      );
      if (answer !== '验证连接') { return; }
      await vscode.window.withProgress({ location: vscode.ProgressLocation.Notification, title: '验证 Go 模型连接', cancellable: true }, async (_progress, token) => {
        const abort = new AbortController();
        const subscription = token.onCancellationRequested(() => abort.abort());
        try {
          const result = await completeChat({
            ...connection,
            sessionId: `ses_${randomUUID().replaceAll('-', '')}`,
            messages: [{ role: 'user', content: 'In VBA, what does Option Explicit do? Answer in one short sentence.' }],
            maxTokens: 128,
          }, () => {}, abort.signal);
          output.appendLine(`连接成功。模型：${connection.model}\n${result.text}`);
          output.show(true);
        } finally { subscription.dispose(); }
      });
    })),
  );
  context.subscriptions.push(...registerComCommands(context));
  registerEncodingChecks(context);
  const client = await startLanguageClient(context);
  // 工程符号查询：转发给语言服务器，供“VBA 符号”视图使用。
  context.subscriptions.push(vscode.commands.registerCommand('vbadminx.workspaceSymbols', async (query: string) => {
    const uri = vscode.window.activeTextEditor?.document.uri;
    if (!uri || !/vba/i.test(vscode.window.activeTextEditor?.document.languageId ?? '')) { return []; }
    if (!vscode.workspace.isTrusted) { return []; }
    return await client.sendRequest('vba/context', { uri: uri.toString(), mode: 'symbols', query: query ?? '' });
  }));
  registerSymbolsView(context);
  const status = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Right, 100);
  status.text = '$(symbol-method) VBAdminx';
  status.tooltip = 'VBA Language Server 已启动。点击查看工程索引。';
  status.command = 'vbadminx.projectStatus';
  status.show();
  context.subscriptions.push(status);
  return { provider, client };
}
