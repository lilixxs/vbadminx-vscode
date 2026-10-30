import { randomUUID } from 'node:crypto';
import * as vscode from 'vscode';
import { readContext, VbaContextTool } from './contextTool';
import { completeChat } from './goClient';
import { getConnection, GoProvider, KEY_NAME, VENDOR } from './provider';
import { startLanguageClient } from './languageClient';
import { registerComCommands } from './com';
import { registerComDiscovery } from './comDiscovery';
import type { LanguageClient } from 'vscode-languageclient/node';
import { registerEncodingChecks } from './encoding';
import { registerSymbolsView } from './symbolsView';
import { registerSnippetSupport } from './snippets';
import { registerFormPreview } from './formPreview';
import { activateLocale, localeLabel, t, watchLanguage } from './i18n';

export async function activate(context: vscode.ExtensionContext): Promise<{ provider: GoProvider; client: LanguageClient }> {
  // 先确定界面语言，后续所有提示都按当前语言输出。
  activateLocale();
  const provider = new GoProvider(context);
  const output = vscode.window.createOutputChannel('VBAdminx');
  const guarded = (action: () => Promise<unknown>) => async () => {
    try { return await action(); }
    catch (error) {
      if (error instanceof vscode.CancellationError || (error instanceof Error && error.name === 'AbortError')) { return; }
      const message = error instanceof Error ? error.message : t('error.generic');
      void vscode.window.showErrorMessage(t('error.prefix', message));
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
        title: t('go.saveKey.title'),
        prompt: t('go.saveKey.prompt'),
        password: true,
        ignoreFocusOut: true,
        validateInput: value => value.trim() ? undefined : t('go.saveKey.required'),
      });
      if (key === undefined) { return; }
      await context.secrets.store(KEY_NAME, key.trim());
      await vscode.workspace.getConfiguration('vbadminx.go').update('connection', 'direct', vscode.ConfigurationTarget.Global);
      provider.refresh();
      void vscode.window.showInformationMessage(t('go.saveKey.saved'));
    })),
    vscode.commands.registerCommand('vbadminx.clearGoKey', guarded(async () => {
      await context.secrets.delete(KEY_NAME);
      provider.refresh();
      void vscode.window.showInformationMessage(t('go.clearKey.done'));
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
      if (!vscode.workspace.isTrusted) { throw new Error(t('go.test.untrusted')); }
      const connection = await getConnection(context);
      const answer = await vscode.window.showInformationMessage(
        t('go.test.confirm', new URL(connection.baseUrl).origin, connection.model),
        { modal: true }, t('go.test.button'),
      );
      if (answer !== t('go.test.button')) { return; }
      await vscode.window.withProgress({ location: vscode.ProgressLocation.Notification, title: t('go.test.progress'), cancellable: true }, async (_progress, token) => {
        const abort = new AbortController();
        const subscription = token.onCancellationRequested(() => abort.abort());
        try {
          const result = await completeChat({
            ...connection,
            sessionId: `ses_${randomUUID().replaceAll('-', '')}`,
            messages: [{ role: 'user', content: 'In VBA, what does Option Explicit do? Answer in one short sentence.' }],
            maxTokens: 128,
          }, () => {}, abort.signal);
          output.appendLine(t('go.test.success', connection.model, result.text));
          output.show(true);
        } finally { subscription.dispose(); }
      });
    })),
  );
  context.subscriptions.push(...registerComCommands(context));
  context.subscriptions.push(registerComDiscovery(context));
  registerEncodingChecks(context);
  registerSnippetSupport(context);
  registerFormPreview(context);
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
  status.tooltip = t('status.tooltip');
  status.command = 'vbadminx.projectStatus';
  status.show();
  context.subscriptions.push(status);
  // 语言设置变化：刷新符号视图分组名与状态栏提示。
  watchLanguage(context, () => {
    status.tooltip = t('status.tooltip');
    void vscode.commands.executeCommand('vbadminx.refreshSymbols');
    output.appendLine(t('language.changed', localeLabel()));
  });
  return { provider, client };
}
