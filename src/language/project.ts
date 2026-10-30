import { promises as fs } from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { TextDecoder } from 'node:util';
import type { Diagnostic } from 'vscode-languageserver/node';
import type { ModuleKind, ParsedModule, PortableTypeIndex, ParseOptions } from './model';
import { parseModule } from './parser';
import { validateTypeIndex } from '../typeIndex';

export interface ProjectConfig {
  name?: string;
  host?: string;
  targetBits?: 32 | 64;
  constants?: Record<string, boolean | number | string>;
  modules?: Record<string, { name?: string; kind?: ModuleKind }>;
  typeIndexes?: string[];
  references?: Array<{ name: string; guid: string; major: number; minor: number; lcid: number; bitness: 32 | 64 }>;
}
export interface Project {
  root: string;
  name: string;
  config: ProjectConfig;
  indexes: PortableTypeIndex[];
  files: Set<string>;
  errors: string[];
}
const kinds = new Set(['standard', 'class', 'document', 'form']);
const folders = ['模块', '类模块', '表格对象', '窗体'];
const ignored = new Set(['node_modules', '.git', '.svn', '.vscode-test', 'dist', 'output']);
const decoder = new TextDecoder('utf-8', { fatal: true, ignoreBOM: false });
const atStart = { start: { line: 0, character: 0 }, end: { line: 0, character: 1 } };
export const DIAGNOSTIC_SOURCE = 'VBAdminx';

export function canonicalUri(uri: string): string {
  try {
    let target = fileURLToPath(uri);
    if (process.platform === 'win32') { target = target.replace(/^[a-z]:/, value => value.toUpperCase()); }
    return pathToFileURL(target).href;
  } catch { return uri; }
}
export function isIgnoredUri(uri: string): boolean {
  try { return fileURLToPath(uri).split(/[\\/]/).some(segment => ignored.has(segment)); } catch { return false; }
}

export function parseProjectConfig(value: unknown): ProjectConfig {
  if (!value || typeof value !== 'object' || Array.isArray(value)) { throw new Error('工程配置必须为 JSON 对象。'); }
  const data = value as Record<string, unknown>;
  const allowed = new Set(['$schema', 'name', 'host', 'targetBits', 'constants', 'modules', 'typeIndexes', 'references']);
  if (Object.keys(data).some(field => !allowed.has(field))) { throw new Error('工程配置包含未知字段，请检查拼写。'); }
  if (data.targetBits !== undefined && data.targetBits !== 32 && data.targetBits !== 64) { throw new Error('targetBits 必须为 32 或 64。'); }
  if (data.name !== undefined && typeof data.name !== 'string') { throw new Error('name 必须为字符串。'); }
  if (data.host !== undefined && typeof data.host !== 'string') { throw new Error('host 必须为字符串。'); }
  if (data.typeIndexes !== undefined && (!Array.isArray(data.typeIndexes) || data.typeIndexes.some(item => typeof item !== 'string'))) { throw new Error('typeIndexes 必须为路径数组。'); }
  if (data.constants !== undefined) {
    if (!data.constants || typeof data.constants !== 'object' || Array.isArray(data.constants) || Object.values(data.constants).some(item => !['boolean', 'number', 'string'].includes(typeof item) || typeof item === 'number' && !Number.isFinite(item))) { throw new Error('constants 只能包含布尔值、数字或字符串。'); }
  }
  if (data.modules !== undefined) {
    if (!data.modules || typeof data.modules !== 'object' || Array.isArray(data.modules)) { throw new Error('modules 必须为映射。'); }
    for (const entry of Object.values(data.modules)) {
      if (!entry || typeof entry !== 'object' || Array.isArray(entry)) { throw new Error('模块映射必须为对象。'); }
      const item = entry as Record<string, unknown>;
      if (item.name !== undefined && (typeof item.name !== 'string' || !/^[\p{L}_][\p{L}\p{N}_]*$/u.test(item.name))) { throw new Error('模块名无效。'); }
      if (item.kind !== undefined && !kinds.has(String(item.kind))) { throw new Error('模块类型无效。'); }
    }
  }
  if (data.references !== undefined) {
    if (!Array.isArray(data.references)) { throw new Error('references 必须为数组。'); }
    for (const reference of data.references) {
      if (!reference || typeof reference !== 'object' || typeof reference.name !== 'string' || typeof reference.guid !== 'string' || ![32, 64].includes(reference.bitness) || !['major', 'minor', 'lcid'].every(key => Number.isInteger(reference[key]) && reference[key] >= 0)) { throw new Error('外部引用需指定 name、guid、major、minor、lcid 和 bitness。'); }
    }
  }
  return data as ProjectConfig;
}

async function entries(root: string): Promise<import('node:fs').Dirent[]> {
  try { return await fs.readdir(root, { withFileTypes: true }); } catch { return []; }
}

export class ProjectStore {
  revision = 0;
  readonly modules = new Map<string, ParsedModule>();
  readonly projects: Project[] = [];
  readonly invalidFiles = new Set<string>();
  readonly conflictedModules = new Set<string>();
  private readonly open = new Map<string, string>();
  constructor(readonly roots: string[]) {}

  async initialize(): Promise<void> {
    this.revision++;
    this.modules.clear();
    this.invalidFiles.clear();
    this.conflictedModules.clear();
    this.projects.length = 0;
    for (const root of this.roots) { await this.discover(root, 0); }
    this.projects.sort((a, b) => b.root.length - a.root.length);
    for (const project of this.projects) {
      await this.loadIndexes(project);
      await this.scan(project.root, project);
    }
    for (const [uri, text] of this.open) { this.update(uri, text); }
  }

  private async discover(root: string, depth: number): Promise<void> {
    const list = await entries(root);
    const configFile = list.find(item => item.isFile() && item.name === 'vba-project.json');
    const automatic = depth <= 1 && folders.every(name => list.some(item => item.isDirectory() && item.name === name));
    if (configFile || automatic) {
      let config: ProjectConfig = {};
      const errors: string[] = [];
      if (configFile) {
        try { config = parseProjectConfig(JSON.parse(decoder.decode(await fs.readFile(path.join(root, configFile.name))))); }
        catch (error) { errors.push(error instanceof Error ? error.message : '配置读取失败。'); }
      }
      if (!this.projects.some(item => item.root === root)) { this.projects.push({ root, name: config.name ?? path.basename(root), config, indexes: [], files: new Set(), errors }); }
    }
    // 显式配置可位于更深目录；约定布局只检查根和直接子目录。
    for (const item of list) {
      if (item.isDirectory() && !ignored.has(item.name) && !item.name.startsWith('.')) { await this.discover(path.join(root, item.name), depth + 1); }
    }
  }

  private async loadIndexes(project: Project): Promise<void> {
    for (const relative of project.config.typeIndexes ?? []) {
      try {
        const target = path.resolve(project.root, relative);
        const content = await fs.readFile(target);
        if (content.length > 32 * 1024 * 1024) { throw new Error('类型索引超过 32 MB。'); }
        const index = validateTypeIndex(JSON.parse(decoder.decode(content)));
        if (project.indexes.some(existing => existing.library.name.toLowerCase() === index.library.name.toLowerCase() || existing.library.guid.replace(/[{}]/g, '').toLowerCase() === index.library.guid.replace(/[{}]/g, '').toLowerCase())) {
          project.errors.push(`同一类型库存在重复或多版本索引：${index.library.name}。请显式保留一个版本。`);
          continue;
        }
        if (project.config.targetBits && project.config.targetBits !== index.library.bitness) { project.errors.push(`类型信息缺失：${index.library.name} 索引位数与 targetBits 不匹配。`); continue; }
        const reference = project.config.references?.find(item => item.name.toLowerCase() === index.library.name.toLowerCase());
        if (reference && (reference.guid.replace(/[{}]/g, '').toLowerCase() !== index.library.guid.replace(/[{}]/g, '').toLowerCase() || reference.major !== index.library.major || reference.minor !== index.library.minor || reference.lcid !== index.library.lcid || reference.bitness !== index.library.bitness)) {
          project.errors.push(`类型索引 ${relative} 与显式引用版本不匹配。`);
          continue;
        }
        project.indexes.push(index);
      } catch (error) { project.errors.push(`无法加载类型索引 ${relative}：${error instanceof Error ? error.message : '无效文件'}`); }
    }
    for (const reference of project.config.references ?? []) {
      if (!project.indexes.some(item => item.library.name.toLowerCase() === reference.name.toLowerCase())) { project.errors.push(`类型信息缺失：${reference.name} ${reference.major}.${reference.minor}`); }
    }
  }

  private async scan(root: string, project: Project): Promise<void> {
    for (const item of await entries(root)) {
      const target = path.join(root, item.name);
      if (item.isDirectory() && !ignored.has(item.name) && !item.name.startsWith('.')) {
        if (target !== project.root && this.projects.some(other => other.root === target)) { continue; }
        await this.scan(target, project);
      } else if (item.isFile() && /\.(bas|cls|frm)$/i.test(item.name)) { await this.loadFile(pathToFileURL(target).href); }
    }
  }

  projectFor(uri: string): Project | undefined {
    if (isIgnoredUri(uri)) { return undefined; }
    let target: string;
    try { target = fileURLToPath(uri); } catch { return undefined; }
    return this.projects.find(project => {
      const relative = path.relative(project.root, target);
      return relative !== '' && relative !== '..' && !relative.startsWith(`..${path.sep}`) && !path.isAbsolute(relative);
    });
  }

  private options(uri: string): ParseOptions {
    const project = this.projectFor(uri);
    if (!project) { return {}; }
    const relative = path.relative(project.root, fileURLToPath(uri)).replaceAll('\\', '/');
    const mapping = project.config.modules?.[relative];
    return { moduleName: mapping?.name, moduleKind: mapping?.kind, constants: project.config.constants, targetBits: project.config.targetBits };
  }

  update(uri: string, text: string, isOpen = true): ParsedModule {
    uri = canonicalUri(uri);
    this.revision++;
    if (isOpen) { this.open.set(uri, text); }
    const parsed = parseModule(uri, text, this.options(uri));
    this.conflictedModules.delete(uri);
    const project = this.projectFor(uri);
    project?.files.add(uri);
    // VS Code 打开非 UTF-8 文件时会把无效字节替换为 U+FFFD。已打开文件只能靠替换字符判断，
    // 因此下次从磁盘重读成功前，必须保持“编码无效”状态。
    if (text.includes('\ufffd')) { this.invalidFiles.add(uri); }
    if (this.invalidFiles.has(uri)) {
      parsed.diagnostics = [{ range: atStart, severity: 1, code: 'encoding', source: DIAGNOSTIC_SOURCE, message: '文件不是有效 UTF-8。请先在扩展外转换；语义分析已暂停。' }];
    }
    const header = /^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"/im.exec(text)?.[1];
    if (header && this.options(uri).moduleName && header.toLowerCase() !== parsed.name.toLowerCase()) {
      parsed.diagnostics.push({ range: atStart, severity: 1, code: 'project-conflict', source: DIAGNOSTIC_SOURCE, message: '显式模块名与 Attribute VB_Name 冲突。重命名已禁用。' });
      this.conflictedModules.add(uri);
    }
    const explicitKind = this.options(uri).moduleKind;
    if (explicitKind && /^\s*(?:VERSION|Attribute)\b/im.test(text)) {
      const inferred = parseModule(uri, text).kind;
      if (inferred !== explicitKind) { parsed.diagnostics.push({ range: atStart, severity: 1, code: 'project-conflict', source: DIAGNOSTIC_SOURCE, message: `显式模块类型 ${explicitKind} 与导出头推定类型 ${inferred} 冲突。重命名已禁用。` }); this.conflictedModules.add(uri); }
    }
    this.modules.set(uri, parsed);
    return parsed;
  }

  async loadFile(uri: string): Promise<void> {
    uri = canonicalUri(uri);
    try {
      const data = await fs.readFile(fileURLToPath(uri));
      let text: string;
      try { text = decoder.decode(data); this.invalidFiles.delete(uri); }
      catch { text = data.toString('utf8'); this.invalidFiles.add(uri); }
      this.update(uri, text, false);
    } catch { this.remove(uri); }
  }

  remove(uri: string): void { this.revision++; uri = canonicalUri(uri); this.modules.delete(uri); this.invalidFiles.delete(uri); this.conflictedModules.delete(uri); this.open.delete(uri); for (const project of this.projects) { project.files.delete(uri); } }
  async close(uri: string): Promise<void> { uri = canonicalUri(uri); this.open.delete(uri); await this.loadFile(uri); }
  context(uri: string): { modules: ParsedModule[]; typeIndexes: PortableTypeIndex[] } {
    uri = canonicalUri(uri);
    const project = this.projectFor(uri);
    return {
      modules: [...this.modules.values()].filter(module => !this.invalidFiles.has(module.uri) && (project ? this.projectFor(module.uri) === project : module.uri === uri)),
      typeIndexes: project?.indexes ?? [],
    };
  }
  projectDiagnostics(uri: string): Diagnostic[] {
    uri = canonicalUri(uri);
    const project = this.projectFor(uri);
    const parsed = this.modules.get(uri);
    if (!project || !parsed) { return []; }
    const same = this.context(uri).modules.filter(item => item.name.toLowerCase() === parsed.name.toLowerCase());
    const messages = [...project.errors];
    if ([...this.conflictedModules].some(target => this.projectFor(target) === project) && !this.conflictedModules.has(uri)) { messages.push('工程内存在导出头或模块类型冲突。请先解决配置问题再重命名。'); }
    if (same.length > 1) { messages.push(`模块名 ${parsed.name} 重复。重命名已禁用。`); }
    return messages.map(message => ({ range: atStart, severity: message.startsWith('类型信息缺失') ? 3 : 1, code: /类型索引|类型库|类型信息/.test(message) ? 'project-types' : 'project-conflict', source: DIAGNOSTIC_SOURCE, message }));
  }
}
