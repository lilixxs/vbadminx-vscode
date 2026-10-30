import {
  createConnection, ProposedFeatures, TextDocuments, TextDocumentSyncKind, ResponseError, ErrorCodes, LSPErrorCodes,
  CodeActionKind, type Position, type TextEdit, type Diagnostic, type CancellationToken,
} from 'vscode-languageserver/node';
import { TextDocument } from 'vscode-languageserver-textdocument';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { ProjectStore, canonicalUri, isIgnoredUri, DIAGNOSTIC_SOURCE } from './language/project';
import { VbaService } from './language/service';
import { formatVba, type FormatOptions } from './language/formatter';

const connection = createConnection(ProposedFeatures.all);
const documents = new TextDocuments(TextDocument);
let store = new ProjectStore([]);
let settings: { format?: FormatOptions; diagnostics?: { missingInformation?: 'information' | 'hint' | 'off'; nameCase?: boolean } } = {};
let initialized: Promise<void> = Promise.resolve();
let supportsWorkspaceFolders = false;
const timers = new Map<string, NodeJS.Timeout>();
const abort = (token: CancellationToken): void => { if (token.isCancellationRequested) { throw new ResponseError(LSPErrorCodes.RequestCancelled, '请求已取消。'); } };
const moduleFor = (uri: string) => store.invalidFiles.has(canonicalUri(uri)) ? undefined : store.modules.get(canonicalUri(uri));
let cachedRevision = -1;
let cachedStore = store;
const serviceCache = new Map<string, VbaService>();
const serviceFor = (uri: string): VbaService => {
  if (cachedStore !== store || cachedRevision !== store.revision) { serviceCache.clear(); cachedStore = store; cachedRevision = store.revision; }
  const key = store.projectFor(uri)?.root ?? canonicalUri(uri);
  let service = serviceCache.get(key);
  if (!service) { service = new VbaService(store.context(uri)); serviceCache.set(key, service); }
  return service;
};

connection.onInitialize(params => {
  supportsWorkspaceFolders = params.capabilities.workspace?.workspaceFolders === true;
  const roots = (params.workspaceFolders?.map(folder => folder.uri) ?? (params.rootUri ? [params.rootUri] : [])).flatMap(uri => { try { return [fileURLToPath(uri)]; } catch { return []; } });
  store = new ProjectStore(roots);
  initialized = store.initialize();
  return {
    serverInfo: { name: 'VBAdminx Language Server', version: '0.1.0' },
    capabilities: {
      textDocumentSync: TextDocumentSyncKind.Incremental,
      completionProvider: { triggerCharacters: ['.'] },
      hoverProvider: true,
      definitionProvider: true,
      referencesProvider: true,
      documentSymbolProvider: true,
      workspaceSymbolProvider: true,
      signatureHelpProvider: { triggerCharacters: ['(', ','] },
      renameProvider: { prepareProvider: true },
      documentFormattingProvider: true,
      documentRangeFormattingProvider: true,
      foldingRangeProvider: true,
      codeActionProvider: { codeActionKinds: [CodeActionKind.QuickFix] },
      semanticTokensProvider: { legend: { tokenTypes: ['namespace', 'class', 'enum', 'parameter', 'variable', 'property', 'enumMember', 'function', 'method'], tokenModifiers: ['declaration', 'readonly'] }, full: true },
      workspace: { workspaceFolders: { supported: true, changeNotifications: true } },
    },
  };
});

async function publish(uri: string): Promise<void> {
  await initialized;
  const module = store.modules.get(canonicalUri(uri));
  if (!module) { connection.sendDiagnostics({ uri, diagnostics: [] }); return; }
  let diagnostics: Diagnostic[] = [...module.diagnostics, ...store.projectDiagnostics(uri)];
  if (!store.invalidFiles.has(canonicalUri(uri))) { diagnostics.push(...serviceFor(uri).diagnostics(module)); }
  diagnostics = diagnostics.filter(diagnostic => {
    if (diagnostic.code === 'name-case' && settings.diagnostics?.nameCase === false) { return false; }
    if (String(diagnostic.code).startsWith('missing-') || diagnostic.code === 'cond.unknown') {
      if (settings.diagnostics?.missingInformation === 'off') { return false; }
      if (settings.diagnostics?.missingInformation === 'hint') { diagnostic.severity = 4; }
    }
    return true;
  });
  const openDocument = documents.get(uri) ?? documents.all().find(document => canonicalUri(document.uri) === canonicalUri(uri));
  connection.sendDiagnostics({ uri: openDocument?.uri ?? uri, version: openDocument?.version, diagnostics });
}
function schedule(uri: string): void {
  clearTimeout(timers.get(uri));
  timers.set(uri, setTimeout(() => { timers.delete(uri); void publish(uri); }, 100));
}
connection.onInitialized(async () => {
  await initialized;
  for (const document of documents.all()) { schedule(document.uri); }
  if (supportsWorkspaceFolders) { connection.workspace.onDidChangeWorkspaceFolders(async event => {
    const roots = store.roots.filter(root => !event.removed.some(folder => { try { return fileURLToPath(folder.uri) === root; } catch { return false; } }));
    for (const folder of event.added) { try { roots.push(fileURLToPath(folder.uri)); } catch {} }
    store = new ProjectStore(roots); initialized = store.initialize(); await initialized;
    for (const document of documents.all()) { store.update(document.uri, document.getText()); schedule(document.uri); }
  }); }
});
documents.onDidOpen(async event => {
  await initialized;
  // 编辑器文本是打开时的真相（可能含未保存修改）；磁盘字节校验由 src/encoding.ts 负责，
  // 非 UTF-8 的替换字符由 ProjectStore.update 检测。此处不再调用 loadFile，避免同一文件双重解析。
  store.update(event.document.uri, event.document.getText());
  schedule(event.document.uri);
});
documents.onDidChangeContent(async event => {
  await initialized;
  store.update(event.document.uri, event.document.getText());
  for (const document of documents.all()) { if (store.projectFor(document.uri) === store.projectFor(event.document.uri)) { schedule(document.uri); } }
});
documents.onDidClose(async event => { await store.close(event.document.uri); schedule(event.document.uri); });
connection.onDidChangeConfiguration(event => { settings = event.settings?.vbadminx ?? {}; for (const document of documents.all()) { schedule(document.uri); } });
connection.onDidChangeWatchedFiles(async event => {
  const configurationChanged = event.changes.some(item => {
    if (/\/vba-project\.json$/i.test(item.uri)) { return true; }
    return store.projects.some(project => project.config.typeIndexes?.some(relative => {
      try { return path.resolve(project.root, relative).toLowerCase() === fileURLToPath(item.uri).toLowerCase(); } catch { return false; }
    }));
  });
  if (configurationChanged) {
    const old = [...store.modules.keys()];
    store.modules.clear(); initialized = store.initialize(); await initialized;
    for (const uri of old) { schedule(uri); }
  } else {
    for (const change of event.changes) { if (!/\.(bas|cls|frm)$/i.test(change.uri)) { continue; } if (change.type === 3) { store.remove(change.uri); connection.sendDiagnostics({ uri: change.uri, diagnostics: [] }); } else { await store.loadFile(change.uri); } }
  }
  for (const document of documents.all()) { schedule(document.uri); }
});

connection.onCompletion(async (params, token) => { await initialized; abort(token); const module = moduleFor(params.textDocument.uri); return module && !store.invalidFiles.has(module.uri) ? serviceFor(module.uri).completions(module, params.position) : []; });
connection.onHover(async (params, token) => { await initialized; abort(token); const module = moduleFor(params.textDocument.uri); return module ? serviceFor(module.uri).hover(module, params.position) : null; });
connection.onDefinition(async (params, token) => { await initialized; abort(token); const module = moduleFor(params.textDocument.uri); return module ? serviceFor(module.uri).definition(module, params.position) : []; });
connection.onReferences(async (params, token) => {
  await initialized; abort(token); const module = moduleFor(params.textDocument.uri); if (!module) { return []; }
  const service = serviceFor(module.uri); const target = service.tokenAt(module, params.position); const symbol = target && service.resolve(module, target);
  if (!symbol) { return []; }
  // 引用查询遍历工程内全部模块的全部 token，是最长的请求循环；逐模块响应取消。
  return symbol ? service.references(symbol, params.context.includeDeclaration, () => abort(token)) : [];
});
connection.onDocumentSymbol(async (params, token) => { await initialized; abort(token); const module = moduleFor(params.textDocument.uri); return module ? serviceFor(module.uri).documentSymbols(module) : []; });
connection.onWorkspaceSymbol(async (params, token) => { await initialized; abort(token); return new VbaService({ modules: [...store.modules.values()].filter(module => !store.invalidFiles.has(module.uri) && !isIgnoredUri(module.uri)), typeIndexes: [] }).workspaceSymbols(params.query); });
connection.onSignatureHelp(async (params, token) => { await initialized; abort(token); const module = moduleFor(params.textDocument.uri); return module ? serviceFor(module.uri).signature(module, params.position) : null; });
connection.onPrepareRename(async (params, token) => {
   await initialized; abort(token); const module = moduleFor(params.textDocument.uri); if (!module || store.projectDiagnostics(module.uri).some(diagnostic => diagnostic.code === 'project-conflict' && diagnostic.severity === 1) || module.diagnostics.some(diagnostic => diagnostic.code === 'project-conflict')) { throw new ResponseError(ErrorCodes.InvalidRequest, '工程有配置冲突或文件无效，不能重命名。'); }
  const service = serviceFor(module.uri); const target = service.tokenAt(module, params.position); const symbol = target && service.resolve(module, target);
  if (!symbol || symbol.external || ['module', 'class'].includes(symbol.kind)) { throw new ResponseError(ErrorCodes.InvalidRequest, '只能重命名已解析的源码符号。'); }
  return { range: { start: { line: target!.line, character: target!.start }, end: { line: target!.line, character: target!.end } }, placeholder: symbol.name };
});
connection.onRenameRequest(async (params, token) => {
  await initialized; abort(token); const module = moduleFor(params.textDocument.uri); if (!module || store.invalidFiles.has(module.uri) || store.projectDiagnostics(module.uri).some(diagnostic => diagnostic.code === 'project-conflict' && diagnostic.severity === 1) || module.diagnostics.some(diagnostic => diagnostic.code === 'project-conflict')) { throw new ResponseError(ErrorCodes.InvalidRequest, '工程有配置冲突，不能重命名。'); }
  try {
    const service = serviceFor(module.uri);
    const selected = service.tokenAt(module, params.position);
    const symbol = selected && service.resolve(module, selected);
    if (symbol && store.projectDiagnostics(symbol.uri).some(diagnostic => diagnostic.code === 'project-conflict' && diagnostic.severity === 1)) { throw new Error('声明所在模块有配置冲突，不能重命名。'); }
    return { changes: service.rename(module, params.position, params.newName) };
  }
  catch (error) { throw new ResponseError(ErrorCodes.InvalidRequest, error instanceof Error ? error.message : '重命名失败。'); }
});
connection.onRequest('vba/renameCandidates', async (params: { uri: string; position: Position; newName: string }) => {
  await initialized; const module = moduleFor(params.uri); if (!module) { return { candidates: [], externalCalls: false }; }
  const service = serviceFor(module.uri); const token = service.tokenAt(module, params.position); const symbol = token && service.resolve(module, token);
  return symbol ? { candidates: service.renameCandidates(symbol, params.newName), externalCalls: symbol.access === 'public' || symbol.eventHandler === true } : { candidates: [], externalCalls: false };
});

function format(uri: string, options: { tabSize: number; insertSpaces: boolean }, lineRange?: { start: number; end: number }): TextEdit[] {
  const document = documents.get(uri) ?? documents.all().find(item => canonicalUri(item.uri) === canonicalUri(uri)); if (!document || store.invalidFiles.has(canonicalUri(uri))) { return []; }
  const before = document.getText(); const after = formatVba(before, { ...settings.format, ...options }, lineRange);
  return before === after ? [] : [{ range: { start: { line: 0, character: 0 }, end: document.positionAt(before.length) }, newText: after }];
}
connection.onDocumentFormatting(async (params, token) => { await initialized; abort(token); return format(params.textDocument.uri, params.options); });
connection.onDocumentRangeFormatting(async (params, token) => { await initialized; abort(token); return format(params.textDocument.uri, params.options, { start: params.range.start.line, end: params.range.end.character === 0 && params.range.end.line > params.range.start.line ? params.range.end.line - 1 : params.range.end.line }); });
connection.onFoldingRanges(async (params, token) => { await initialized; abort(token); return moduleFor(params.textDocument.uri)?.folds ?? []; });
connection.languages.semanticTokens.on(async (params, token) => { await initialized; abort(token); const module = moduleFor(params.textDocument.uri); return module && !store.invalidFiles.has(module.uri) ? serviceFor(module.uri).semanticTokens(module) : { data: [] }; });
connection.onCodeAction(async (params, token) => {
  await initialized; abort(token); const module = moduleFor(params.textDocument.uri); if (!module || store.invalidFiles.has(module.uri)) { return []; }
  const fixes = [];
  for (const diagnostic of params.context.diagnostics) {
    const data = diagnostic.data as { name?: string; scope?: string } | undefined;
    if (!data?.name) { continue; }
    let edits: TextEdit[] = [];
    if (diagnostic.code === 'name-case') { edits = [{ range: diagnostic.range, newText: data.name }]; }
    else if (diagnostic.code === 'undeclared') {
      const procedure = module.symbols.find(symbol => symbol.id === data.scope);
      const line = module.lines[diagnostic.range.start.line] ?? '';
      const declaredName = data.name.replace(/[$%&!#@^]$/, '');
      if (declaredName !== data.name) { continue; } // 类型后缀已规定类型，不能用 Variant 声明改变语义。
      const escapedName = data.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      if (!procedure || !/^[\p{L}_][\p{L}\p{N}_]*$/u.test(declaredName) || !new RegExp(`^\\s*(?:Let\\s+|Set\\s+)?${escapedName}\\s*=`, 'i').test(line)) { continue; }
      let declarationLine = procedure.selectionRange.start.line;
      // 声明行可能由多行续行组成。
      while (/\s_\s*(?:'.*)?$/.test(module.lines[declarationLine] ?? '')) { declarationLine++; }
      edits = [{ range: { start: { line: declarationLine + 1, character: 0 }, end: { line: declarationLine + 1, character: 0 } }, newText: `    Dim ${declaredName} As Variant${module.text.includes('\r\n') ? '\r\n' : '\n'}` }];
    }
    if (edits.length) { fixes.push({ title: diagnostic.code === 'name-case' ? `统一名称为 ${data.name}` : `声明局部变量 ${data.name} As Variant`, kind: CodeActionKind.QuickFix, diagnostics: [diagnostic], command: { title: '预览快速修复', command: 'vbadminx.previewEdit', arguments: [{ changes: { [module.uri]: edits } }] } }); }
  }
  return fixes;
});
connection.onRequest('vba/projectStatus', async () => { await initialized; return { projects: store.projects.map(project => ({ name: project.name, root: project.root, files: project.files.size, indexes: project.indexes.map(index => index.library), errors: project.errors })), modules: store.modules.size, invalidUtf8: [...store.invalidFiles] }; });
connection.onRequest('vba/libraryDocument', async (params: { uri: string }) => {
  await initialized;
  const target = new URL(params.uri);
  const [libraryName, typeName] = target.pathname.split('/').filter(Boolean).map(decodeURIComponent);
  const index = store.projects.flatMap(project => project.indexes).find(index => index.library.name.toLowerCase() === libraryName?.toLowerCase());
  const type = index?.types.find(type => type.name.toLowerCase() === typeName?.toLowerCase());
  if (!index || !type) { return `' 类型来源：${libraryName ?? ''}.${typeName ?? ''}\n' 当前工程未加载该类型的声明。\n`; }
  return [`' ${index.library.name} ${index.library.major}.${index.library.minor} / ${index.library.bitness}-bit`, `' ${type.documentation ?? ''}`, `${type.kind} ${type.name}`, ...type.members.map(member => `    ${member.kind} ${member.name}${member.parameters ? `(${member.parameters.map(parameter => `${parameter.optional ? 'Optional ' : ''}${parameter.name} As ${parameter.typeName}`).join(', ')})` : ''} As ${member.returnType ?? 'Variant'}`)].join('\n') + '\n';
});
connection.onRequest('vba/context', async (params: { uri: string; query?: string; position?: Position; mode: 'symbols' | 'references' }) => {
  await initialized; const module = moduleFor(params.uri); if (!module) { return []; }
  const service = serviceFor(module.uri);
  if (params.mode === 'symbols') { return service.workspaceSymbols(params.query ?? ''); }
  const token = params.position && service.tokenAt(module, params.position); const symbol = token && service.resolve(module, token);
  return symbol ? service.references(symbol) : [];
});
connection.onShutdown(() => { for (const timer of timers.values()) { clearTimeout(timer); } });
documents.listen(connection);
connection.listen();
