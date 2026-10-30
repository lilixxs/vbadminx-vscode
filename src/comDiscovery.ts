import { randomUUID } from 'node:crypto';
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import * as vscode from 'vscode';
import { t } from './i18n';
import {
  assertTrusted,
  assertWindowsHost,
  defaultIndexFolder,
  exists,
  pickBitness,
  readTypeIndexFile,
  resolvePowerShell,
  runPowerShell,
  writeTypeIndexFile,
} from './com';
import { typeIndexFileName } from './typeIndex';

/**
 * 自动查找 COM 组件：
 * 1. 按设置 `vbadminx.com.softwareList` 的关键词扫描注册表中的 TypeLib；
 * 2. 让用户多选要提取的库；
 * 3. 逐个调用 export-com.ps1，把可移植类型索引写入 `.vba/type-indexes/`；
 * 4. 可选地写入 `vba-project.json` 的 `typeIndexes`。
 *
 * 只读取注册表与类型库公开元数据，不注册、不执行任何组件。
 */

export const SOFTWARE_SETTING = 'com.softwareList';

export const DEFAULT_SOFTWARE_LIST: readonly string[] = [
  'Excel', 'Word', 'PowerPoint', 'Outlook', 'Access', 'Office',
  'ADODB', 'Scripting', 'MSXML2', 'VBScript', 'WScript', 'Shell', 'DAO',
];

interface DiscoveredLibrary {
  libraryName: string;
  guid: string;
  version: string;
  platform: string;
  path: string;
}

function readSoftwareList(): string[] {
  const configured = vscode.workspace.getConfiguration('vbadminx').get<unknown>(SOFTWARE_SETTING);
  const values = Array.isArray(configured) ? configured : DEFAULT_SOFTWARE_LIST;
  const keywords = values
    .filter((item): item is string => typeof item === 'string')
    .map(item => item.trim())
    .filter(Boolean);
  return [...new Set(keywords)];
}

function parseDiscovered(text: string): DiscoveredLibrary[] {
  let parsed: unknown;
  try { parsed = JSON.parse(text.replace(/^\uFEFF/, '').trim()); }
  catch { throw new Error(t('com.discover.invalidJson')); }
  if (!Array.isArray(parsed)) { throw new Error(t('com.discover.invalidJson')); }
  const libraries: DiscoveredLibrary[] = [];
  const seen = new Set<string>();
  for (const item of parsed) {
    if (item === null || typeof item !== 'object') { continue; }
    const entry = item as Record<string, unknown>;
    const libraryName = typeof entry['libraryName'] === 'string' ? entry['libraryName'] : '';
    const guid = typeof entry['guid'] === 'string' ? entry['guid'] : '';
    const version = typeof entry['version'] === 'string' ? entry['version'] : '';
    const platform = typeof entry['platform'] === 'string' ? entry['platform'] : '';
    const filePath = typeof entry['path'] === 'string' ? entry['path'] : '';
    if (!libraryName || !guid || !version || !filePath) { continue; }
    const key = `${guid}|${version}|${platform}|${filePath.toLowerCase()}`;
    if (seen.has(key)) { continue; }
    seen.add(key);
    libraries.push({ libraryName, guid, version, platform, path: filePath });
  }
  return libraries;
}

async function addToProject(relativePaths: string[]): Promise<void> {
  if (!relativePaths.length) { return; }
  const folder = vscode.workspace.workspaceFolders?.[0];
  if (!folder) { return; }
  const projectUri = vscode.Uri.joinPath(folder.uri, 'vba-project.json');
  if (!(await exists(projectUri))) {
    void vscode.window.showInformationMessage(t('com.discover.noProject'));
    return;
  }
  const answer = await vscode.window.showInformationMessage(
    t('com.discover.addProject', relativePaths.length),
    t('com.discover.addButton'),
  );
  if (answer !== t('com.discover.addButton')) { return; }
  let config: Record<string, unknown>;
  try {
    config = JSON.parse(new TextDecoder('utf-8').decode(await vscode.workspace.fs.readFile(projectUri))) as Record<string, unknown>;
  } catch {
    void vscode.window.showWarningMessage(t('com.discover.noProject'));
    return;
  }
  const existing = Array.isArray(config['typeIndexes'])
    ? (config['typeIndexes'] as unknown[]).filter((item): item is string => typeof item === 'string')
    : [];
  const merged = [...existing];
  for (const item of relativePaths) { if (!merged.includes(item)) { merged.push(item); } }
  config['typeIndexes'] = merged;
  await vscode.workspace.fs.writeFile(projectUri, new TextEncoder().encode(`${JSON.stringify(config, null, 2)}\n`));
  void vscode.window.showInformationMessage(t('com.discover.added', vscode.workspace.asRelativePath(projectUri, true)));
}

async function discover(context: vscode.ExtensionContext): Promise<void> {
  assertTrusted(t('action.discoverCom'));
  assertWindowsHost();
  const keywords = readSoftwareList();
  if (!keywords.length) { throw new Error(t('com.discover.noSoftware')); }
  const script = vscode.Uri.joinPath(context.extensionUri, 'scripts', 'discover-com.ps1');
  if (!(await exists(script))) { throw new Error(t('com.discover.scriptMissing', script.fsPath)); }
  const exportScript = vscode.Uri.joinPath(context.extensionUri, 'scripts', 'export-com.ps1');
  if (!(await exists(exportScript))) { throw new Error(t('com.scriptMissing', exportScript.fsPath)); }
  const bitness = await pickBitness();
  if (!bitness) { return; }

  const powershell = await resolvePowerShell();
  const found = parseDiscovered(await vscode.window.withProgress(
    { location: vscode.ProgressLocation.Notification, title: t('com.discover.title'), cancellable: true },
    async (progress, token) => {
      progress.report({ message: t('com.discover.progress') });
      return await runPowerShell({
        powershell,
        script: script.fsPath,
        args: ['-Keywords', keywords.join(';'), '-Bitness', String(bitness)],
        token,
        maxCapture: 4 * 1024 * 1024,
      });
    },
  ));
  if (!found.length) {
    void vscode.window.showInformationMessage(t('com.discover.none'));
    return;
  }

  const picked = await vscode.window.showQuickPick(
    found.map(library => ({
      label: library.libraryName,
      description: `${library.version} · ${library.platform}`,
      detail: library.path,
      library,
    })),
    {
      title: t('com.discover.pickTitle'),
      placeHolder: t('com.discover.pickPlaceholder'),
      canPickMany: true,
      ignoreFocusOut: true,
      matchOnDescription: true,
      matchOnDetail: true,
    },
  );
  if (!picked?.length) { return; }

  const folder = defaultIndexFolder();
  if (!folder) { throw new Error(t('com.discover.noProject')); }
  await vscode.workspace.fs.createDirectory(folder);

  const saved = await vscode.window.withProgress(
    { location: vscode.ProgressLocation.Notification, title: t('com.discover.title'), cancellable: true },
    async (progress, token) => {
      const results: string[] = [];
      for (let index = 0; index < picked.length; index += 1) {
        if (token.isCancellationRequested) { throw new vscode.CancellationError(); }
        const library = picked[index]!.library;
        progress.report({ message: t('com.discover.extractProgress', library.libraryName, index + 1, picked.length) });
        const temp = path.join(os.tmpdir(), `vba-coder-discover-${randomUUID()}.json`);
        try {
          await runPowerShell({
            powershell,
            script: exportScript.fsPath,
            args: ['-InputPath', library.path, '-OutputPath', temp, '-Bitness', String(bitness)],
            token,
          });
          const extracted = await readTypeIndexFile(vscode.Uri.file(temp), t('com.index.extractedLabel'));
          if (extracted.library.bitness !== bitness) { continue; }
          const target = vscode.Uri.joinPath(folder, typeIndexFileName(extracted));
          await writeTypeIndexFile(target, extracted);
          results.push(vscode.workspace.asRelativePath(target, false));
        } finally {
          await fs.rm(temp, { force: true }).catch(() => undefined);
        }
      }
      return results;
    },
  );

  if (!saved.length) {
    void vscode.window.showWarningMessage(t('com.discover.none'));
    return;
  }
  void vscode.window.showInformationMessage(t('com.discover.done', saved.length));
  await addToProject(saved);
}

export function registerComDiscovery(context: vscode.ExtensionContext): vscode.Disposable {
  return vscode.commands.registerCommand('vbadminx.discoverCom', async () => {
    try { await discover(context); }
    catch (error) {
      if (error instanceof vscode.CancellationError || (error instanceof Error && error.name === 'AbortError')) { return; }
      const message = error instanceof Error ? error.message : t('error.generic');
      void vscode.window.showErrorMessage(t('error.prefix', message));
    }
  });
}
