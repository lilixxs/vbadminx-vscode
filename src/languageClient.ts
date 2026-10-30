import path from 'node:path';
import * as vscode from 'vscode';
import { LanguageClient, TransportKind, type LanguageClientOptions, type ServerOptions } from 'vscode-languageclient/node';

interface SerializedEdit { changes?: Record<string, Array<{ range: { start: { line: number; character: number }; end: { line: number; character: number } }; newText: string }>>; }
let current: LanguageClient | undefined;
export const getLanguageClient = (): LanguageClient | undefined => current;

function toEdit(serialized: SerializedEdit): vscode.WorkspaceEdit {
  const edit = new vscode.WorkspaceEdit();
  for (const [uri, changes] of Object.entries(serialized.changes ?? {})) {
    edit.set(vscode.Uri.parse(uri), changes.map(change => new vscode.TextEdit(new vscode.Range(change.range.start.line, change.range.start.character, change.range.end.line, change.range.end.character), change.newText)));
  }
  return edit;
}

export async function startLanguageClient(context: vscode.ExtensionContext): Promise<LanguageClient> {
  const serverModule = context.asAbsolutePath(path.join('dist', 'server.bundle.js'));
  const serverOptions: ServerOptions = { run: { module: serverModule, transport: TransportKind.ipc }, debug: { module: serverModule, transport: TransportKind.ipc, options: { execArgv: ['--nolazy', '--inspect=6010'] } } };
  const watcher = vscode.workspace.createFileSystemWatcher('**/*.{bas,cls,frm,json}');
  const originalContent = new Map<string, string>();
  const previews = vscode.workspace.registerTextDocumentContentProvider('vba-preview', { provideTextDocumentContent: uri => originalContent.get(uri.toString()) ?? '' });
  context.subscriptions.push(vscode.workspace.onDidCloseTextDocument(document => { if (document.uri.scheme === 'vba-preview') { originalContent.delete(document.uri.toString()); } }));
  const clientOptions: LanguageClientOptions = {
    documentSelector: [{ scheme: 'file', language: 'vba' }],
    synchronize: { fileEvents: watcher, configurationSection: 'vbadminx' },
    middleware: {
      provideRenameEdits: async (document, position, newName, token, next) => {
        const requestedVersion = document.version;
        const openedVersions = new Map(vscode.workspace.textDocuments.map(document => [document.uri.toString(), document.version]));
        const edits = await next(document, position, newName, token);
        if (!edits) { return edits; }
        if (document.version !== requestedVersion) { throw new Error('源码在计算重命名时已更改，请重新执行。'); }
        const versions = new Map<string, number>();
        const captureVersion = async (uri: vscode.Uri): Promise<void> => {
          const version = (await vscode.workspace.openTextDocument(uri)).version;
          const previous = openedVersions.get(uri.toString());
          if (previous !== undefined && previous !== version) { throw new Error('目标文件在计算修改时已更改，请重新执行。'); }
          versions.set(uri.toString(), version);
        };
        for (const [uri] of edits.entries()) { await captureVersion(uri); }
        const extra = await current?.sendRequest<{ externalCalls: boolean; candidates: Array<{ uri: string; range: { start: { line: number; character: number }; end: { line: number; character: number } }; newText: string; reason: string }> }>('vba/renameCandidates', { uri: document.uri.toString(), position, newName });
        if (extra?.candidates.length) {
          for (const candidate of extra.candidates) { if (!versions.has(vscode.Uri.parse(candidate.uri).toString())) { await captureVersion(vscode.Uri.parse(candidate.uri)); } }
          const chosen = await vscode.window.showQuickPick(extra.candidates.map((candidate, index) => ({ label: `${vscode.workspace.asRelativePath(vscode.Uri.parse(candidate.uri))}:${candidate.range.start.line + 1}`, description: candidate.reason, picked: false, index })), { title: '选择推断候选；默认不修改', canPickMany: true, ignoreFocusOut: true });
          if (chosen === undefined) { return undefined; }
          for (const item of chosen) {
            const candidate = extra.candidates[item.index]!;
            edits.replace(vscode.Uri.parse(candidate.uri), new vscode.Range(candidate.range.start.line, candidate.range.start.character, candidate.range.end.line, candidate.range.end.character), candidate.newText);
          }
        }
        await assertVersions(versions);
        const count = edits.entries().reduce((sum, [, items]) => sum + items.length, 0);
        const decision = await vscode.window.showInformationMessage(`重命名将修改 ${edits.entries().length} 个文件的 ${count} 处。${extra?.externalCalls ? '该符号可能存在源码外调用，请同步检查事件和外部入口。' : ''}`, { modal: true }, '预览并继续', '取消');
        if (decision !== '预览并继续') { return undefined; }
        await showDiffs(edits, versions);
        const confirm = await vscode.window.showInformationMessage('已打开修改前后差异。应用全部已选修改？', { modal: true }, '应用修改');
        await assertVersions(versions);
        return confirm === '应用修改' ? edits : undefined;
      },
    },
  };
  async function assertVersions(versions: Map<string, number>): Promise<void> {
    for (const [uri, version] of versions) {
      if ((await vscode.workspace.openTextDocument(vscode.Uri.parse(uri))).version !== version) { throw new Error('源码在预览期间已更改。已取消旧修改，请重新执行。'); }
    }
  }
  async function showDiffs(edit: vscode.WorkspaceEdit, versions = new Map<string, number>()): Promise<Map<string, number>> {
    for (const [uri, changes] of edit.entries()) {
      if (!vscode.workspace.getWorkspaceFolder(uri)) { throw new Error('修改目标不属于当前工作区。'); }
      const document = await vscode.workspace.openTextDocument(uri);
      if (!versions.has(uri.toString())) { versions.set(uri.toString(), document.version); }
      let after = document.getText();
      for (const change of [...changes].sort((a, b) => document.offsetAt(b.range.start) - document.offsetAt(a.range.start))) {
        after = after.slice(0, document.offsetAt(change.range.start)) + change.newText + after.slice(document.offsetAt(change.range.end));
      }
      const preview = vscode.Uri.parse(`vba-preview:/${encodeURIComponent(path.basename(uri.fsPath))}?${Date.now()}-${Math.random()}`);
      originalContent.set(preview.toString(), after);
      if (originalContent.size > 32) { originalContent.delete(originalContent.keys().next().value!); }
      await vscode.commands.executeCommand('vscode.diff', uri, preview, `${path.basename(uri.fsPath)}：修改预览`, { preview: false });
    }
    return versions;
  }
  const client = new LanguageClient('vbadminx', 'VBAdminx Language Server', serverOptions, clientOptions);
  current = client;
  context.subscriptions.push(watcher, previews, { dispose: () => { current = undefined; void client.stop(); } });
  context.subscriptions.push(vscode.commands.registerCommand('vbadminx.previewEdit', async (serialized: SerializedEdit) => {
    if (!vscode.workspace.isTrusted) { throw new Error('请先信任工作区。'); }
    const edit = toEdit(serialized);
    const versions = await showDiffs(edit);
    if (await vscode.window.showInformationMessage('应用预览中的快速修复？', { modal: true }, '应用修改') === '应用修改') { await assertVersions(versions); await vscode.workspace.applyEdit(edit); }
  }));
  context.subscriptions.push(vscode.commands.registerCommand('vbadminx.projectStatus', async () => {
    const data = await client.sendRequest('vba/projectStatus');
    const document = await vscode.workspace.openTextDocument({ language: 'json', content: JSON.stringify(data, null, 2) });
    await vscode.window.showTextDocument(document, { preview: true });
  }));
  context.subscriptions.push(vscode.commands.registerCommand('vbadminx.createProject', async () => {
    const roots = vscode.workspace.workspaceFolders;
    if (!roots?.length) { void vscode.window.showErrorMessage('先打开导出工程目录。'); return; }
    const root = roots.length === 1 ? roots[0]! : (await vscode.window.showWorkspaceFolderPick());
    if (!root) { return; }
    const target = vscode.Uri.joinPath(root.uri, 'vba-project.json');
    try { await vscode.workspace.fs.stat(target); void vscode.window.showWarningMessage('vba-project.json 已存在，未覆盖。'); return; } catch {}
    await vscode.workspace.fs.writeFile(target, Buffer.from(JSON.stringify({ name: root.name, host: 'Excel', targetBits: 64, constants: { VBA7: true, Win64: true }, modules: {}, typeIndexes: [], references: [] }, null, 2) + '\n'));
    await vscode.window.showTextDocument(await vscode.workspace.openTextDocument(target));
  }));
  // 外部库定义以只读虚拟文档展示，避免导航到不存在的文件。
  context.subscriptions.push(vscode.workspace.registerTextDocumentContentProvider('vba-library', { provideTextDocumentContent: uri => client.sendRequest<string>('vba/libraryDocument', { uri: uri.toString() }) }));
  await client.start();
  return client;
}
