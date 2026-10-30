import { spawn } from 'node:child_process';
import { randomUUID } from 'node:crypto';
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import * as vscode from 'vscode';
import type { PortableTypeIndex } from './language/model';
import { t } from './i18n';
import {
  parseTypeIndex,
  stringifyTypeIndex,
  typeIndexFileName,
  typeIndexLabel,
} from './typeIndex';

/**
 * COM 类型索引命令。导入、导出与校验在全部平台可用（COM-02/COM-05）；
 * 提取只在 Windows 扩展主机调用本地脚本读取公开类型元数据（COM-01/COM-04），
 * 不注册类型库、不执行组件、不连接 Office。
 *
 * 本模块同时导出脚本执行与索引读写工具，供“自动查找 COM 组件”复用。
 */

export const INDEX_FILTER: Record<string, string[]> = { 'Type index JSON': ['json'] };
export const MAX_INDEX_BYTES = 32 * 1024 * 1024;
const EXTRACTION_TIMEOUT_MS = 120_000;
const MAX_CAPTURE = 64 * 1024;
const INDEX_FOLDER = ['.vba', 'type-indexes'] as const;

export function assertTrusted(action: string): void {
  if (!vscode.workspace.isTrusted) { throw new Error(t('error.trust', action)); }
}

export function assertWindowsHost(): void {
  if (process.platform !== 'win32') { throw new Error(t('com.windowsOnly')); }
}

export function defaultIndexFolder(): vscode.Uri | undefined {
  const folder = vscode.workspace.workspaceFolders?.[0];
  return folder ? vscode.Uri.joinPath(folder.uri, ...INDEX_FOLDER) : undefined;
}

export function defaultIndexUri(index: PortableTypeIndex): vscode.Uri {
  const folder = defaultIndexFolder();
  return folder
    ? vscode.Uri.joinPath(folder, typeIndexFileName(index))
    : vscode.Uri.file(path.join(os.homedir(), typeIndexFileName(index)));
}

export async function exists(uri: vscode.Uri): Promise<boolean> {
  try { await vscode.workspace.fs.stat(uri); return true; } catch { return false; }
}

export async function readTypeIndexFile(uri: vscode.Uri, label: string): Promise<PortableTypeIndex> {
  let stat: vscode.FileStat;
  try { stat = await vscode.workspace.fs.stat(uri); }
  catch { throw new Error(t('com.import.readFailed', label, uri.fsPath)); }
  if ((stat.type & vscode.FileType.Directory) !== 0) { throw new Error(t('com.import.notFile', label)); }
  if (stat.size > MAX_INDEX_BYTES) { throw new Error(t('com.import.tooLarge', label)); }
  const bytes = await vscode.workspace.fs.readFile(uri);
  let text: string;
  try { text = new TextDecoder('utf-8', { fatal: true }).decode(bytes); }
  catch { throw new Error(t('com.import.invalidUtf8', label)); }
  try { return parseTypeIndex(text); }
  catch (error) {
    throw new Error(t('com.import.invalid', label, error instanceof Error ? error.message : t('com.import.invalidDetail')));
  }
}

async function confirmOverwrite(uri: vscode.Uri): Promise<boolean> {
  if (!(await exists(uri))) { return true; }
  const choice = await vscode.window.showWarningMessage(
    t('com.overwrite.title', uri.fsPath),
    { modal: true },
    t('com.overwrite.button'),
  );
  return choice === t('com.overwrite.button');
}

export async function writeTypeIndexFile(uri: vscode.Uri, index: PortableTypeIndex): Promise<void> {
  if (uri.scheme !== 'file') { throw new Error(t('com.export.localOnly')); }
  const directory = vscode.Uri.file(path.dirname(uri.fsPath));
  await vscode.workspace.fs.createDirectory(directory);
  // stringifyTypeIndex 会再次校验，写出的字节始终是无 BOM 的 UTF-8。
  await vscode.workspace.fs.writeFile(uri, new TextEncoder().encode(stringifyTypeIndex(index)));
}

function relativeToWorkspace(uri: vscode.Uri): string | undefined {
  return vscode.workspace.getWorkspaceFolder(uri)
    ? vscode.workspace.asRelativePath(uri, false)
    : undefined;
}

export function nextStep(uri: vscode.Uri): string {
  const relative = relativeToWorkspace(uri);
  return relative
    ? t('com.nextStep.relative', relative)
    : t('com.nextStep.outside');
}

async function importTypeIndex(): Promise<void> {
  assertTrusted(t('action.importTypeIndex'));
  const picked = await vscode.window.showOpenDialog({
    canSelectMany: false,
    openLabel: t('com.import.openLabel'),
    title: t('com.import.title'),
    filters: INDEX_FILTER,
    defaultUri: defaultIndexFolder(),
  });
  const source = picked?.[0];
  if (!source) { return; }
  const index = await readTypeIndexFile(source, t('com.index.readLabel'));
  const target = await vscode.window.showSaveDialog({
    saveLabel: t('com.import.saveLabel'),
    title: t('com.save.title', typeIndexLabel(index)),
    filters: INDEX_FILTER,
    defaultUri: defaultIndexUri(index),
  });
  if (!target) { return; }
  if (!(await confirmOverwrite(target))) { return; }
  await writeTypeIndexFile(target, index);
  void vscode.window.showInformationMessage(t('com.import.saved', typeIndexLabel(index), nextStep(target)));
}

async function exportTypeIndex(): Promise<void> {
  assertTrusted(t('action.exportTypeIndex'));
  const picked = await vscode.window.showOpenDialog({
    canSelectMany: false,
    openLabel: t('com.export.openLabel'),
    title: t('com.export.title'),
    filters: INDEX_FILTER,
    defaultUri: defaultIndexFolder(),
  });
  const source = picked?.[0];
  if (!source) { return; }
  if (source.scheme !== 'file') { throw new Error(t('com.export.localOnly')); }
  const index = await readTypeIndexFile(source, t('com.index.readLabel'));
  const target = await vscode.window.showSaveDialog({
    saveLabel: t('com.export.saveLabel'),
    title: t('com.export.saveLabel'),
    filters: INDEX_FILTER,
    defaultUri: vscode.Uri.file(path.join(path.dirname(source.fsPath), typeIndexFileName(index))),
  });
  if (!target) { return; }
  if (target.scheme === 'file' && path.resolve(source.fsPath) === path.resolve(target.fsPath)) {
    throw new Error(t('com.export.sameFile'));
  }
  if (!(await confirmOverwrite(target))) { return; }
  await writeTypeIndexFile(target, index);
  void vscode.window.showInformationMessage(t('com.export.saved', typeIndexLabel(index)));
}

export async function resolvePowerShell(): Promise<string> {
  const systemRoot = process.env['SystemRoot'] ?? process.env['windir'] ?? 'C:\\Windows';
  const bundled = path.join(systemRoot, 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');
  try { await fs.access(bundled); return bundled; } catch { return 'powershell.exe'; }
}

export async function pickBitness(): Promise<32 | 64 | undefined> {
  const defaultBitness: 32 | 64 = process.arch === 'x64' || process.arch === 'arm64' ? 64 : 32;
  const items: Array<vscode.QuickPickItem & { bitness: 32 | 64 }> = [
    { label: `$(chip) ${t('com.bitness.64')}`, description: defaultBitness === 64 ? t('com.bitness.default') : t('com.bitness.64hint'), bitness: 64 },
    { label: `$(chip) ${t('com.bitness.32')}`, description: defaultBitness === 32 ? t('com.bitness.default') : t('com.bitness.32hint'), bitness: 32 },
  ];
  const picked = await vscode.window.showQuickPick(items, {
    title: t('com.bitness.title'),
    placeHolder: t('com.bitness.placeholder'),
    ignoreFocusOut: true,
  });
  return picked?.bitness;
}

export interface PowerShellRunOptions {
  powershell: string;
  script: string;
  args: string[];
  token: vscode.CancellationToken;
  timeoutMs?: number;
  /** 标准输出捕获上限；自动查找的 JSON 可能较大，默认 64 KB。 */
  maxCapture?: number;
}

/**
 * 运行 PowerShell 脚本并返回标准输出。
 * 参数逐项传递，shell 始终为 false；文件名中的空格、引号不会被解释。
 */
export function runPowerShell(options: PowerShellRunOptions): Promise<string> {
  const timeoutMs = options.timeoutMs ?? EXTRACTION_TIMEOUT_MS;
  const maxCapture = options.maxCapture ?? MAX_CAPTURE;
  const args = [
    '-NoLogo', '-NoProfile', '-NonInteractive',
    '-ExecutionPolicy', 'Bypass',
    '-File', options.script,
    ...options.args,
  ];
  return new Promise<string>((resolve, reject) => {
    const child = spawn(options.powershell, args, { shell: false, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] });
    let stdout = '';
    let stderr = '';
    let settled = false;
    let termination: Error | undefined;
    let timer: ReturnType<typeof setTimeout> | undefined;
    let cancellation: vscode.Disposable | undefined;

    const settle = (error?: Error): void => {
      if (settled) { return; }
      settled = true;
      if (timer) { clearTimeout(timer); }
      cancellation?.dispose();
      if (error) { reject(error); } else { resolve(stdout); }
    };
    const stop = (error: Error): void => {
      termination ??= error;
      child.kill();
      // 进程不响应 kill 时也必须结束命令，避免永不返回。
      const guard = setTimeout(() => { if (termination) { settle(termination); } }, 5_000);
      guard.unref?.();
    };

    timer = setTimeout(() => stop(new Error(t('com.timeout', Math.round(timeoutMs / 1000)))), timeoutMs);
    cancellation = options.token.onCancellationRequested(() => stop(new vscode.CancellationError()));

    child.stdout?.setEncoding('utf8');
    child.stderr?.setEncoding('utf8');
    child.stdout?.on('data', (chunk: string) => {
      if (stdout.length < maxCapture) { stdout += chunk.slice(0, maxCapture - stdout.length); }
    });
    child.stderr?.on('data', (chunk: string) => {
      if (stderr.length < MAX_CAPTURE) { stderr += chunk.slice(0, MAX_CAPTURE - stderr.length); }
    });
    child.on('error', (error: Error) => settle(new Error(t('com.powershell', error.message))));
    child.on('close', (code, signal) => {
      if (termination) { settle(termination); return; }
      if (options.token.isCancellationRequested) { settle(new vscode.CancellationError()); return; }
      if (code !== 0) {
        const detail = (stderr.trim() || stdout.trim()).slice(-4_000) || t('com.noOutput');
        settle(new Error(t('com.failed', code ?? '?', signal ? t('com.signal', signal) : '', detail)));
        return;
      }
      settle();
    });
  });
}

async function extractComTypeIndex(context: vscode.ExtensionContext): Promise<void> {
  assertTrusted(t('action.extractTypeIndex'));
  assertWindowsHost();
  const script = vscode.Uri.joinPath(context.extensionUri, 'scripts', 'export-com.ps1');
  if (!(await exists(script))) { throw new Error(t('com.scriptMissing', script.fsPath)); }
  const picked = await vscode.window.showOpenDialog({
    canSelectMany: false,
    openLabel: t('com.extract.openLabel'),
    title: t('com.extract.title'),
    filters: { [t('com.extract.filters')]: ['tlb', 'olb', 'dll'] },
    defaultUri: defaultIndexFolder(),
  });
  const source = picked?.[0];
  if (!source) { return; }
  if (source.scheme !== 'file') { throw new Error(t('com.extract.localOnly')); }
  const bitness = await pickBitness();
  if (!bitness) { return; }

  const index = await vscode.window.withProgress(
    { location: vscode.ProgressLocation.Notification, title: t('com.extract.progress', path.basename(source.fsPath)), cancellable: true },
    async (progress, token) => {
      progress.report({ message: t('com.extract.reading') });
      const temp = path.join(os.tmpdir(), `vba-coder-type-index-${randomUUID()}.json`);
      try {
        await runPowerShell({
          powershell: await resolvePowerShell(),
          script: script.fsPath,
          args: ['-InputPath', source.fsPath, '-OutputPath', temp, '-Bitness', String(bitness)],
          token,
        });
        if (token.isCancellationRequested) { throw new vscode.CancellationError(); }
        const extracted = await readTypeIndexFile(vscode.Uri.file(temp), t('com.index.extractedLabel'));
        if (extracted.library.bitness !== bitness) { throw new Error(t('com.extract.mismatch')); }
        return extracted;
      } finally {
        await fs.rm(temp, { force: true }).catch(() => undefined);
      }
    },
  );

  const target = await vscode.window.showSaveDialog({
    saveLabel: t('com.import.saveLabel'),
    title: t('com.save.title', typeIndexLabel(index)),
    filters: INDEX_FILTER,
    defaultUri: defaultIndexUri(index),
  });
  if (!target) { return; }
  if (!(await confirmOverwrite(target))) { return; }
  await writeTypeIndexFile(target, index);
  void vscode.window.showInformationMessage(t('com.extract.saved', typeIndexLabel(index), nextStep(target)));
}

/** 调用方负责把返回的 Disposable 加入 context.subscriptions。 */
export function registerComCommands(context: vscode.ExtensionContext): vscode.Disposable[] {
  const guarded = (action: () => Promise<void>) => async (): Promise<void> => {
    try { await action(); }
    catch (error) {
      if (error instanceof vscode.CancellationError || (error instanceof Error && error.name === 'AbortError')) { return; }
      const message = error instanceof Error ? error.message : t('error.generic');
      void vscode.window.showErrorMessage(t('error.prefix', message));
    }
  };
  return [
    vscode.commands.registerCommand('vbadminx.importTypeIndex', guarded(importTypeIndex)),
    vscode.commands.registerCommand('vbadminx.exportTypeIndex', guarded(exportTypeIndex)),
    vscode.commands.registerCommand('vbadminx.extractComTypeIndex', guarded(() => extractComTypeIndex(context))),
  ];
}
