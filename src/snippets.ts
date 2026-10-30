import * as vscode from 'vscode';
import { t } from './i18n';
import { CUSTOM_SNIPPET_TEMPLATE, matchSnippets, parseCustomSnippets, type CustomSnippet } from './snippetFile';

/**
 * 自定义代码片段：
 * - 内置片段由 `contributes.snippets`（snippets/vba.json）提供；
 * - 本模块加载设置 `vbadminx.snippets.customPath` 指向的片段文件，并注册补全项；
 * - `vbadminx.editSnippets` 打开（必要时创建）该文件，方便在设置之外直接编辑。
 */
const SECTION = 'vbadminx';
const SETTING = 'snippets.customPath';
const DEFAULT_RELATIVE = '.vba/snippets.json';

interface CacheEntry {
  uri: string;
  mtime: number;
  snippets: CustomSnippet[];
}

let cache: CacheEntry | undefined;
let reportedError: string | undefined;

function resolveUri(value: string): vscode.Uri | undefined {
  const trimmed = value.trim();
  if (!trimmed) { return undefined; }
  if (/^[A-Za-z]:[\\/]/.test(trimmed) || trimmed.startsWith('/') || trimmed.startsWith('\\\\')) {
    return vscode.Uri.file(trimmed);
  }
  const folder = vscode.workspace.workspaceFolders?.[0];
  if (!folder) { return undefined; }
  return vscode.Uri.joinPath(folder.uri, ...trimmed.split(/[\\/]+/).filter(Boolean));
}

function configuredUri(): vscode.Uri | undefined {
  return resolveUri(vscode.workspace.getConfiguration(SECTION).get<string>(SETTING, ''));
}

function defaultUri(): vscode.Uri | undefined {
  const folder = vscode.workspace.workspaceFolders?.[0];
  return folder ? vscode.Uri.joinPath(folder.uri, ...DEFAULT_RELATIVE.split('/')) : undefined;
}

async function loadSnippets(): Promise<CustomSnippet[]> {
  const uri = configuredUri();
  if (!uri) { return []; }
  let stat: vscode.FileStat;
  try { stat = await vscode.workspace.fs.stat(uri); }
  catch { return cache?.uri === uri.toString() ? cache.snippets : []; }
  if (cache && cache.uri === uri.toString() && cache.mtime === stat.mtime) { return cache.snippets; }
  let text: string;
  try {
    text = new TextDecoder('utf-8').decode(await vscode.workspace.fs.readFile(uri));
  } catch {
    return [];
  }
  try {
    const snippets = parseCustomSnippets(text);
    cache = { uri: uri.toString(), mtime: stat.mtime, snippets };
    reportedError = undefined;
    return snippets;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    // 同一文件同一错误只提示一次，避免每次补全都弹窗。
    const stamp = `${uri.toString()}#${stat.mtime}`;
    if (reportedError !== stamp) {
      reportedError = stamp;
      void vscode.window.showWarningMessage(t('snippets.invalid', message));
    }
    return [];
  }
}

function createProvider(): vscode.CompletionItemProvider {
  return {
    async provideCompletionItems(document, position) {
      const range = document.getWordRangeAtPosition(position, /[A-Za-z0-9_]+/);
      if (!range) { return undefined; }
      const word = document.getText(range);
      if (!word) { return undefined; }
      const snippets = await loadSnippets();
      const matches = matchSnippets(snippets, word);
      if (!matches.length) { return undefined; }
      return matches.map(snippet => {
        const item = new vscode.CompletionItem(snippet.prefixes[0] ?? snippet.name, vscode.CompletionItemKind.Snippet);
        item.insertText = new vscode.SnippetString(snippet.body);
        item.detail = t('snippets.detail');
        item.documentation = snippet.description ? new vscode.MarkdownString(snippet.description) : undefined;
        item.filterText = `${snippet.prefixes.join(' ')} ${snippet.name}`;
        item.sortText = `00${snippet.prefixes[0] ?? snippet.name}`;
        item.range = range;
        return item;
      });
    },
  };
}

async function editSnippets(): Promise<void> {
  if (!vscode.workspace.isTrusted) { throw new Error(t('snippets.trusted')); }
  const config = vscode.workspace.getConfiguration(SECTION);
  const configured = config.get<string>(SETTING, '').trim();
  const uri = configured ? resolveUri(configured) : defaultUri();
  if (!uri) { throw new Error(t('snippets.noWorkspace')); }
  let stat: vscode.FileStat | undefined;
  try { stat = await vscode.workspace.fs.stat(uri); } catch { stat = undefined; }
  if (!stat) {
    await vscode.workspace.fs.createDirectory(vscode.Uri.joinPath(uri, '..'));
    await vscode.workspace.fs.writeFile(uri, new TextEncoder().encode(CUSTOM_SNIPPET_TEMPLATE));
    void vscode.window.showInformationMessage(t('snippets.created', vscode.workspace.asRelativePath(uri, true)));
  }
  if (!configured) {
    const folder = vscode.workspace.getWorkspaceFolder(uri);
    const relative = folder ? vscode.workspace.asRelativePath(uri, false) : uri.fsPath;
    await config.update(SETTING, relative, vscode.ConfigurationTarget.Workspace);
    void vscode.window.showInformationMessage(t('snippets.setting', relative));
  }
  try {
    const document = await vscode.workspace.openTextDocument(uri);
    await vscode.window.showTextDocument(document);
  } catch (error) {
    throw new Error(t('snippets.openFailed', uri.fsPath, error instanceof Error ? error.message : String(error)));
  }
}

export function registerSnippetSupport(context: vscode.ExtensionContext): void {
  context.subscriptions.push(
    vscode.languages.registerCompletionItemProvider('vba', createProvider(), ' ', '_'),
    vscode.commands.registerCommand('vbadminx.editSnippets', async () => {
      try { await editSnippets(); }
      catch (error) {
        if (error instanceof vscode.CancellationError) { return; }
        void vscode.window.showErrorMessage(t('error.prefix', error instanceof Error ? error.message : t('error.generic')));
      }
    }),
    vscode.workspace.onDidChangeConfiguration(event => {
      if (event.affectsConfiguration(`${SECTION}.${SETTING}`)) { cache = undefined; }
    }),
  );
}
