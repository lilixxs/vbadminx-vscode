import { spawn } from 'node:child_process';
import { randomUUID } from 'node:crypto';
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import * as vscode from 'vscode';
import type { PortableTypeIndex } from './language/model';
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
 */

const INDEX_FILTER: Record<string, string[]> = { '类型索引 JSON': ['json'] };
const MAX_INDEX_BYTES = 32 * 1024 * 1024;
const EXTRACTION_TIMEOUT_MS = 120_000;
const MAX_CAPTURE = 64 * 1024;
const INDEX_FOLDER = ['.vba', 'type-indexes'] as const;

function assertTrusted(action: string): void {
  if (!vscode.workspace.isTrusted) { throw new Error(`先信任工作区，再${action}。`); }
}

function defaultIndexFolder(): vscode.Uri | undefined {
  const folder = vscode.workspace.workspaceFolders?.[0];
  return folder ? vscode.Uri.joinPath(folder.uri, ...INDEX_FOLDER) : undefined;
}

function defaultIndexUri(index: PortableTypeIndex): vscode.Uri {
  const folder = defaultIndexFolder();
  return folder
    ? vscode.Uri.joinPath(folder, typeIndexFileName(index))
    : vscode.Uri.file(path.join(os.homedir(), typeIndexFileName(index)));
}

async function exists(uri: vscode.Uri): Promise<boolean> {
  try { await vscode.workspace.fs.stat(uri); return true; } catch { return false; }
}

async function readTypeIndexFile(uri: vscode.Uri, label: string): Promise<PortableTypeIndex> {
  let stat: vscode.FileStat;
  try { stat = await vscode.workspace.fs.stat(uri); }
  catch { throw new Error(`${label}无法读取：${uri.fsPath}`); }
  if ((stat.type & vscode.FileType.Directory) !== 0) { throw new Error(`${label}指向目录，不是文件。`); }
  if (stat.size > MAX_INDEX_BYTES) { throw new Error(`${label}超过 32 MB，已拒绝读取。`); }
  const bytes = await vscode.workspace.fs.readFile(uri);
  let text: string;
  try { text = new TextDecoder('utf-8', { fatal: true }).decode(bytes); }
  catch { throw new Error(`${label}不是有效 UTF-8。`); }
  try { return parseTypeIndex(text); }
  catch (error) { throw new Error(`${label}无效：${error instanceof Error ? error.message : '无法解析。'}`); }
}

async function confirmOverwrite(uri: vscode.Uri): Promise<boolean> {
  if (!(await exists(uri))) { return true; }
  const choice = await vscode.window.showWarningMessage(
    `目标文件已存在，覆盖？\n${uri.fsPath}`,
    { modal: true },
    '覆盖',
  );
  return choice === '覆盖';
}

async function writeTypeIndexFile(uri: vscode.Uri, index: PortableTypeIndex): Promise<void> {
  if (uri.scheme !== 'file') { throw new Error('类型索引只能写入本地文件系统。'); }
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

function nextStep(uri: vscode.Uri): string {
  const relative = relativeToWorkspace(uri);
  return relative
    ? `请在 vba-project.json 的 typeIndexes 中显式加入 ${relative}（COM-03）。`
    : '该文件不在当前工作区，工程不会自动使用它。';
}

async function importTypeIndex(): Promise<void> {
  assertTrusted('导入类型索引');
  const picked = await vscode.window.showOpenDialog({
    canSelectMany: false,
    openLabel: '导入',
    title: '选择可移植类型索引 JSON',
    filters: INDEX_FILTER,
    defaultUri: defaultIndexFolder(),
  });
  const source = picked?.[0];
  if (!source) { return; }
  const index = await readTypeIndexFile(source, '选中的类型索引');
  const target = await vscode.window.showSaveDialog({
    saveLabel: '保存到工作区',
    title: `导入 ${typeIndexLabel(index)}`,
    filters: INDEX_FILTER,
    defaultUri: defaultIndexUri(index),
  });
  if (!target) { return; }
  if (!(await confirmOverwrite(target))) { return; }
  await writeTypeIndexFile(target, index);
  void vscode.window.showInformationMessage(`已导入 ${typeIndexLabel(index)}。${nextStep(target)}`);
}

async function exportTypeIndex(): Promise<void> {
  assertTrusted('导出类型索引');
  const picked = await vscode.window.showOpenDialog({
    canSelectMany: false,
    openLabel: '选择要导出的索引',
    title: '选择已保存的可移植类型索引 JSON',
    filters: INDEX_FILTER,
    defaultUri: defaultIndexFolder(),
  });
  const source = picked?.[0];
  if (!source) { return; }
  if (source.scheme !== 'file') { throw new Error('仅支持本地文件系统中的类型索引。'); }
  const index = await readTypeIndexFile(source, '选中的类型索引');
  const target = await vscode.window.showSaveDialog({
    saveLabel: '导出',
    title: `导出 ${typeIndexLabel(index)}`,
    filters: INDEX_FILTER,
    defaultUri: vscode.Uri.file(path.join(path.dirname(source.fsPath), typeIndexFileName(index))),
  });
  if (!target) { return; }
  if (target.scheme === 'file' && path.resolve(source.fsPath) === path.resolve(target.fsPath)) {
    throw new Error('导出目标与源文件相同。请另选路径。');
  }
  if (!(await confirmOverwrite(target))) { return; }
  await writeTypeIndexFile(target, index);
  void vscode.window.showInformationMessage(`已导出 ${typeIndexLabel(index)}。`);
}

function assertWindowsHost(): void {
  if (process.platform !== 'win32') {
    throw new Error('类型库提取只在 Windows 扩展主机可用（COM-01）。请在 Windows 机器上提取，再用导入/导出类型索引交付（COM-04）。');
  }
}

async function resolvePowerShell(): Promise<string> {
  const systemRoot = process.env['SystemRoot'] ?? process.env['windir'] ?? 'C:\\Windows';
  const bundled = path.join(systemRoot, 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');
  try { await fs.access(bundled); return bundled; } catch { return 'powershell.exe'; }
}

async function pickBitness(): Promise<32 | 64 | undefined> {
  const defaultBitness: 32 | 64 = process.arch === 'x64' || process.arch === 'arm64' ? 64 : 32;
  const items: Array<vscode.QuickPickItem & { bitness: 32 | 64 }> = [
    { label: '$(chip) 64 位', description: defaultBitness === 64 ? '默认（当前扩展主机架构）' : 'VBA7 64 位宿主', bitness: 64 },
    { label: '$(chip) 32 位', description: defaultBitness === 32 ? '默认（当前扩展主机架构）' : 'VBA7 32 位宿主', bitness: 32 },
  ];
  const picked = await vscode.window.showQuickPick(items, {
    title: '索引目标位数',
    placeHolder: '选择 VBA7 工程的目标位数。类型库本身不携带位数，必须显式记录（COM-02/COM-03）。',
    ignoreFocusOut: true,
  });
  return picked?.bitness;
}

interface ExtractionOptions {
  powershell: string;
  script: string;
  input: string;
  output: string;
  bitness: 32 | 64;
  token: vscode.CancellationToken;
}

function runExtraction(options: ExtractionOptions): Promise<void> {
  // 参数逐项传递，shell 始终为 false；文件名中的空格、引号不会被解释。
  const args = [
    '-NoLogo', '-NoProfile', '-NonInteractive',
    '-ExecutionPolicy', 'Bypass',
    '-File', options.script,
    '-InputPath', options.input,
    '-OutputPath', options.output,
    '-Bitness', String(options.bitness),
  ];
  return new Promise<void>((resolve, reject) => {
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
      if (error) { reject(error); } else { resolve(); }
    };
    const stop = (error: Error): void => {
      termination ??= error;
      child.kill();
      // 进程不响应 kill 时也必须结束命令，避免永不返回。
      const guard = setTimeout(() => { if (termination) { settle(termination); } }, 5_000);
      guard.unref?.();
    };

    timer = setTimeout(() => stop(new Error(`提取超时（${Math.round(EXTRACTION_TIMEOUT_MS / 1000)} 秒）。`)), EXTRACTION_TIMEOUT_MS);
    cancellation = options.token.onCancellationRequested(() => stop(new vscode.CancellationError()));

    child.stdout?.setEncoding('utf8');
    child.stderr?.setEncoding('utf8');
    child.stdout?.on('data', (chunk: string) => {
      if (stdout.length < MAX_CAPTURE) { stdout += chunk.slice(0, MAX_CAPTURE - stdout.length); }
    });
    child.stderr?.on('data', (chunk: string) => {
      if (stderr.length < MAX_CAPTURE) { stderr += chunk.slice(0, MAX_CAPTURE - stderr.length); }
    });
    child.on('error', (error: Error) => settle(new Error(`无法启动 PowerShell：${error.message}`)));
    child.on('close', (code, signal) => {
      if (termination) { settle(termination); return; }
      if (options.token.isCancellationRequested) { settle(new vscode.CancellationError()); return; }
      if (code !== 0) {
        const detail = (stderr.trim() || stdout.trim()).slice(-4_000) || '没有错误输出。';
        settle(new Error(`提取脚本失败（退出码 ${code ?? '未知'}${signal ? `，信号 ${signal}` : ''}）。${detail}`));
        return;
      }
      settle();
    });
  });
}

async function extractComTypeIndex(context: vscode.ExtensionContext): Promise<void> {
  assertTrusted('提取 COM 类型索引');
  assertWindowsHost();
  const script = vscode.Uri.joinPath(context.extensionUri, 'scripts', 'export-com.ps1');
  if (!(await exists(script))) { throw new Error(`找不到提取脚本：${script.fsPath}`); }
  const picked = await vscode.window.showOpenDialog({
    canSelectMany: false,
    openLabel: '提取',
    title: '选择 TypeLib、OLB 或 DLL。只读取公开类型元数据，不注册、不执行。',
    filters: { '类型库与二进制': ['tlb', 'olb', 'dll'] },
    defaultUri: defaultIndexFolder(),
  });
  const source = picked?.[0];
  if (!source) { return; }
  if (source.scheme !== 'file') { throw new Error('类型库必须位于本地文件系统。'); }
  const bitness = await pickBitness();
  if (!bitness) { return; }

  const index = await vscode.window.withProgress(
    { location: vscode.ProgressLocation.Notification, title: `提取 ${path.basename(source.fsPath)}`, cancellable: true },
    async (progress, token) => {
      progress.report({ message: '正在读取公开类型元数据…' });
      const temp = path.join(os.tmpdir(), `vba-coder-type-index-${randomUUID()}.json`);
      try {
        await runExtraction({
          powershell: await resolvePowerShell(),
          script: script.fsPath,
          input: source.fsPath,
          output: temp,
          bitness,
          token,
        });
        if (token.isCancellationRequested) { throw new vscode.CancellationError(); }
        const extracted = await readTypeIndexFile(vscode.Uri.file(temp), '提取结果');
        if (extracted.library.bitness !== bitness) { throw new Error('提取结果的目标位数与选择不一致。'); }
        return extracted;
      } finally {
        await fs.rm(temp, { force: true }).catch(() => undefined);
      }
    },
  );

  const target = await vscode.window.showSaveDialog({
    saveLabel: '保存',
    title: `保存 ${typeIndexLabel(index)}`,
    filters: INDEX_FILTER,
    defaultUri: defaultIndexUri(index),
  });
  if (!target) { return; }
  if (!(await confirmOverwrite(target))) { return; }
  await writeTypeIndexFile(target, index);
  void vscode.window.showInformationMessage(`已提取 ${typeIndexLabel(index)}。${nextStep(target)}`);
}

/** 调用方负责把返回的 Disposable 加入 context.subscriptions。 */
export function registerComCommands(context: vscode.ExtensionContext): vscode.Disposable[] {
  const guarded = (action: () => Promise<void>) => async (): Promise<void> => {
    try { await action(); }
    catch (error) {
      if (error instanceof vscode.CancellationError || (error instanceof Error && error.name === 'AbortError')) { return; }
      const message = error instanceof Error ? error.message : '操作失败。';
      void vscode.window.showErrorMessage(`VBAdminx: ${message}`);
    }
  };
  return [
    vscode.commands.registerCommand('vbadminx.importTypeIndex', guarded(importTypeIndex)),
    vscode.commands.registerCommand('vbadminx.exportTypeIndex', guarded(exportTypeIndex)),
    vscode.commands.registerCommand('vbadminx.extractComTypeIndex', guarded(() => extractComTypeIndex(context))),
  ];
}
