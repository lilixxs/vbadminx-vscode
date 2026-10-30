import * as vscode from 'vscode';
import { t } from './i18n';
import type { MessageKey } from './i18n/catalog';

/**
 * “VBA 符号”侧边栏视图：列出工程或当前文件的函数、过程与变量。
 * - 点击项：在编辑器中打开对应文件并选中该符号（名称范围）。
 * - 光标移动：在树中同步显示并选中当前所在符号。
 * - 标题栏按钮：刷新、视图选项（范围 / 排序 / 过滤）与设置。
 * - 打开 VBA 工程后自动展开一次（每个扩展会话）。
 */

export const SYMBOLS_VIEW_ID = 'vbadminxSymbols';

type Scope = 'currentFile' | 'workspace';
type SortMode = 'position' | 'name';
type FilterMode = 'all' | 'functions' | 'variables';
type GroupKind = 'procedures' | 'variables' | 'types';

interface Options {
  scope: Scope;
  sort: SortMode;
  filter: FilterMode;
}

interface SymbolEntry {
  name: string;
  detail: string;
  icon: string;
  uri: vscode.Uri;
  /** 名称范围，用于“选中该符号”。 */
  selectionRange: vscode.Range;
  /** 完整范围，用于判断光标落在哪个符号内。 */
  range: vscode.Range;
  kind: 'procedure' | 'variable' | 'constant' | 'type' | 'property';
  children: SymbolEntry[];
  parent?: GroupNode;
}

type GroupNode = {
  kind: 'group';
  group: GroupKind;
  label: string;
  icon: string;
  children: SymbolEntry[];
};

type Node = GroupNode | SymbolEntry;

const isGroup = (node: Node): node is GroupNode => (node as GroupNode).kind === 'group';

const PROCEDURE_KINDS = new Set([vscode.SymbolKind.Function, vscode.SymbolKind.Method, vscode.SymbolKind.Property, vscode.SymbolKind.Event, vscode.SymbolKind.Constructor]);
const VARIABLE_KINDS = new Set([vscode.SymbolKind.Variable, vscode.SymbolKind.Field, vscode.SymbolKind.Constant, vscode.SymbolKind.String, vscode.SymbolKind.Number, vscode.SymbolKind.Boolean, vscode.SymbolKind.Array, vscode.SymbolKind.Key]);
const TYPE_KINDS = new Set([vscode.SymbolKind.Class, vscode.SymbolKind.Struct, vscode.SymbolKind.Enum, vscode.SymbolKind.Interface, vscode.SymbolKind.EnumMember]);
const MAX_ITEMS = 800;

function classify(symbol: vscode.DocumentSymbol): SymbolEntry['kind'] {
  if (PROCEDURE_KINDS.has(symbol.kind)) { return symbol.kind === vscode.SymbolKind.Property ? 'property' : 'procedure'; }
  if (VARIABLE_KINDS.has(symbol.kind)) { return symbol.kind === vscode.SymbolKind.Constant ? 'constant' : 'variable'; }
  return 'type';
}

function groupKindOf(entry: SymbolEntry): GroupKind {
  if (entry.kind === 'procedure' || entry.kind === 'property') { return 'procedures'; }
  if (entry.kind === 'variable' || entry.kind === 'constant') { return 'variables'; }
  return 'types';
}

function groupIconOf(kind: GroupKind): string {
  return kind === 'procedures' ? 'symbol-method' : kind === 'variables' ? 'symbol-variable' : 'symbol-struct';
}

const GROUP_LABEL_KEY: Record<GroupKind, MessageKey> = {
  procedures: 'symbols.group.procedures',
  variables: 'symbols.group.variables',
  types: 'symbols.group.types',
};

function entryIconOf(kind: SymbolEntry['kind']): string {
  return kind === 'procedure' ? 'symbol-method'
    : kind === 'property' ? 'symbol-property'
      : kind === 'variable' ? 'symbol-variable'
        : kind === 'constant' ? 'symbol-constant' : 'symbol-struct';
}

function convert(
  symbol: vscode.DocumentSymbol,
  uri: vscode.Uri,
  depth: number,
): SymbolEntry {
  const kind = classify(symbol);
  const entry: SymbolEntry = {
    name: symbol.name,
    detail: symbol.detail ?? '',
    icon: entryIconOf(kind),
    uri,
    selectionRange: symbol.selectionRange,
    range: symbol.range,
    kind,
    children: [],
  };
  // 只展开一层：过程内部的参数与局部变量。
  if (depth === 0) {
    for (const child of symbol.children ?? []) {
      const converted = convert(child, uri, depth + 1);
      converted.parent = undefined;
      entry.children.push(converted);
    }
  }
  return entry;
}

export class SymbolsViewProvider implements vscode.TreeDataProvider<Node> {
  private readonly emitter = new vscode.EventEmitter<Node | undefined>();
  readonly onDidChangeTreeData = this.emitter.event;
  private entries: SymbolEntry[] = [];
  private groups: GroupNode[] = [];
  private options: Options;
  private suppressReveal = false;
  private timer: ReturnType<typeof setTimeout> | undefined;

  constructor(private readonly state: vscode.Memento) {
    this.options = {
      scope: (state.get<string>('symbols.scope') as Scope) ?? 'currentFile',
      sort: (state.get<string>('symbols.sort') as SortMode) ?? 'position',
      filter: (state.get<string>('symbols.filter') as FilterMode) ?? 'all',
    };
  }

  getOptions(): Options { return { ...this.options }; }

  async setOptions(patch: Partial<Options>): Promise<void> {
    this.options = { ...this.options, ...patch };
    for (const [key, value] of Object.entries(this.options)) { await this.state.update(`symbols.${key}`, value); }
    await this.refresh();
  }

  refresh(): void {
    this.emitter.fire(undefined);
  }

  getTreeItem(node: Node): vscode.TreeItem {
    if (isGroup(node)) {
      const item = new vscode.TreeItem(node.label, vscode.TreeItemCollapsibleState.Expanded);
      item.iconPath = new vscode.ThemeIcon(node.icon);
      item.contextValue = 'group';
      return item;
    }
    const collapsible = node.children.length > 0
      ? vscode.TreeItemCollapsibleState.Collapsed
      : vscode.TreeItemCollapsibleState.None;
    const item = new vscode.TreeItem(node.name, collapsible);
    item.iconPath = new vscode.ThemeIcon(node.icon);
    item.description = node.detail;
    item.resourceUri = node.uri;
    item.tooltip = `${node.name}${node.detail ? ` — ${node.detail}` : ''}\n${vscode.workspace.asRelativePath(node.uri, true)}:${node.selectionRange.start.line + 1}`;
    item.contextValue = 'symbol';
    // 点击项：跳转并选中符号（范围编码进 command 参数，避免依赖树实例比较）。
    item.command = {
      command: 'vbadminx.symbolsGoto',
      title: 'Goto symbol',
      arguments: [node.uri.toString(), node.selectionRange.start.line, node.selectionRange.start.character, node.selectionRange.end.line, node.selectionRange.end.character],
    };
    return item;
  }

  async getChildren(node?: Node): Promise<Node[]> {
    if (node === undefined) { await this.load(); return this.groups as Node[]; }
    if (isGroup(node)) { return node.children as Node[]; }
    return node.children as Node[];
  }

  getParent(node: Node): Node | undefined {
    if (isGroup(node)) { return undefined; }
    if (node.parent !== undefined) { return node.parent; }
    return this.groups.find(group => group.children.includes(node));
  }

  /** 找到包含该位置的符号树项，用于“左侧同步显示当前选中”。 */
  entryAt(uri: vscode.Uri, position: vscode.Position): SymbolEntry | undefined {
    const sameFile = this.entries.filter(entry => entry.uri.toString() === uri.toString());
    let best: SymbolEntry | undefined;
    let bestSize = Number.MAX_SAFE_INTEGER;
    const consider = (entry: SymbolEntry): void => {
      if (!entry.range.contains(position)) { return; }
      const size = entry.range.end.line - entry.range.start.line;
      if (size < bestSize) { bestSize = size; best = entry; }
    };
    for (const entry of sameFile) {
      consider(entry);
      for (const child of entry.children) { consider(child); }
    }
    return best;
  }

  private async load(): Promise<void> {
    const editor = vscode.window.activeTextEditor;
    if (this.options.scope === 'currentFile') {
      const document = editor?.document;
      if (!document || !/vba/i.test(document.languageId)) {
        this.entries = []; this.groups = []; return;
      }
      const symbols = await vscode.commands.executeCommand<vscode.DocumentSymbol[]>(
        'vscode.executeDocumentSymbolProvider', document.uri,
      ) ?? [];
      this.entries = symbols.map(symbol => convert(symbol, document.uri, 0));
      this.buildGroups();
      return;
    }
    // 整个工程：用语言服务器的自定义请求聚合各工程符号。
    const uri = editor?.document.uri;
    const raw = uri
      ? await vscode.commands.executeCommand<Array<{ name: string; kind: number; location: { uri: string; range: { start: { line: number; character: number }; end: { line: number; character: number } } }; containerName?: string }>>(
        'vbadminx.workspaceSymbols', '',
      ) ?? []
      : [];
    this.entries = [];
    for (const item of raw.slice(0, MAX_ITEMS)) {
      let targetUri: vscode.Uri;
      try { targetUri = vscode.Uri.parse(item.location.uri); } catch { continue; }
      const range = new vscode.Range(
        item.location.range.start.line, item.location.range.start.character,
        item.location.range.end.line, item.location.range.end.character,
      );
      const kind = PROCEDURE_KINDS.has(item.kind) ? 'procedure'
        : VARIABLE_KINDS.has(item.kind) ? 'variable'
          : TYPE_KINDS.has(item.kind) ? 'type' : 'variable';
      this.entries.push({
        name: item.name, detail: item.containerName ?? '', icon: entryIconOf(kind),
        uri: targetUri, selectionRange: range, range, kind, children: [],
      });
    }
    this.buildGroups(true);
  }

  private buildGroups(byModule = false): void {
    let visible = this.entries;
    if (this.options.filter === 'functions') { visible = visible.filter(entry => entry.kind === 'procedure' || entry.kind === 'property'); }
    if (this.options.filter === 'variables') { visible = visible.filter(entry => entry.kind === 'variable' || entry.kind === 'constant'); }
    if (this.options.sort === 'name') {
      visible = [...visible].sort((a, b) => a.name.toLowerCase().localeCompare(b.name.toLowerCase()));
    } else {
      visible = [...visible].sort((a, b) => a.range.start.line - b.range.start.line || a.range.start.character - b.range.start.character);
    }
    visible = visible.slice(0, MAX_ITEMS);
    if (byModule) {
      const map = new Map<string, SymbolEntry[]>();
      for (const entry of visible) {
        const key = vscode.workspace.asRelativePath(entry.uri, true);
        const list = map.get(key) ?? [];
        list.push(entry);
        map.set(key, list);
      }
      this.groups = [...map.entries()].map(([label, children]) => ({
        kind: 'group' as const, group: 'procedures' as const, label, icon: 'file-code', children,
      }));
      for (const group of this.groups) { for (const child of group.children) { child.parent = group; } }
      return;
    }
    const map = new Map<GroupKind, SymbolEntry[]>();
    for (const entry of visible) {
      const kind = groupKindOf(entry);
      const list = map.get(kind) ?? [];
      list.push(entry);
      map.set(kind, list);
    }
    const order: GroupKind[] = ['procedures', 'variables', 'types'];
    this.groups = order.filter(kind => map.has(kind)).map(kind => {
      const group: GroupNode = {
        kind: 'group', group: kind, label: t(GROUP_LABEL_KEY[kind]), icon: groupIconOf(kind), children: map.get(kind) ?? [],
      };
      for (const child of group.children) { child.parent = group; }
      return group;
    });
  }

  /** 供外部在文档变化后防抖刷新。 */
  scheduleRefresh(): void {
    if (this.timer) { clearTimeout(this.timer); }
    this.timer = setTimeout(() => { this.timer = undefined; this.refresh(); }, 400);
  }

  /** 同步期间抑制自身触发的回环。 */
  async withSuppressedReveal(action: () => Promise<void>): Promise<void> {
    this.suppressReveal = true;
    try { await action(); } finally { this.suppressReveal = false; }
  }

  get isSuppressed(): boolean { return this.suppressReveal; }

  dispose(): void {
    if (this.timer) { clearTimeout(this.timer); }
    this.emitter.dispose();
  }
}

function hasVbaEditor(): boolean {
  return vscode.window.activeTextEditor?.document.languageId === 'vba';
}

async function workspaceHasProject(): Promise<boolean> {
  if (!vscode.workspace.workspaceFolders?.length) { return false; }
  try {
    const found = await vscode.workspace.findFiles('**/vba-project.json', '**/node_modules/**', 1);
    return found.length > 0;
  } catch { return false; }
}

/**
 * 打开 VBA 工程或 VBA 文件后自动展开一次“VBA 符号”视图。
 * 每个扩展会话只展开一次，避免用户收起后又被强制打开。
 */
function autoExpandView(context: vscode.ExtensionContext): void {
  let expanded = false;
  const expand = async (): Promise<void> => {
    if (expanded) { return; }
    if (!hasVbaEditor() && !(await workspaceHasProject())) { return; }
    expanded = true;
    // 等视图完成首次数据加载再聚焦，避免空列表闪烁。
    setTimeout(() => { void vscode.commands.executeCommand(`${SYMBOLS_VIEW_ID}.focus`); }, 250);
  };
  context.subscriptions.push(
    vscode.window.onDidChangeActiveTextEditor(() => { void expand(); }),
    vscode.workspace.onDidOpenTextDocument(document => { if (document.languageId === 'vba') { void expand(); } }),
  );
  void expand();
}

export function registerSymbolsView(context: vscode.ExtensionContext): void {
  const provider = new SymbolsViewProvider(context.workspaceState);
  const view = vscode.window.createTreeView(SYMBOLS_VIEW_ID, {
    treeDataProvider: provider,
    showCollapseAll: true,
    canSelectMany: false,
  });
  context.subscriptions.push(view, provider);

  // 点击项：打开文档并选中符号名称。
  context.subscriptions.push(vscode.commands.registerCommand(
    'vbadminx.symbolsGoto',
    async (uriText: string, startLine: number, startChar: number, endLine: number, endChar: number) => {
      let uri: vscode.Uri;
      try { uri = vscode.Uri.parse(uriText); } catch { return; }
      const document = await vscode.workspace.openTextDocument(uri);
      const editor = await vscode.window.showTextDocument(document, { preview: true, preserveFocus: false });
      const range = new vscode.Range(startLine, startChar, endLine, endChar);
      await provider.withSuppressedReveal(async () => {
        editor.selection = new vscode.Selection(range.start, range.end);
        editor.revealRange(range, vscode.TextEditorRevealType.InCenterIfOutsideViewport);
      });
      await revealCurrent(provider, view, document.uri, range.start);
    },
  ));

  context.subscriptions.push(vscode.commands.registerCommand('vbadminx.refreshSymbols', () => provider.refresh()));

  // 聚焦（并滚动到）“VBA 符号”视图；VS Code 为每个视图自动生成 <viewId>.focus 命令。
  context.subscriptions.push(vscode.commands.registerCommand('vbadminx.focusSymbols', async () => {
    await vscode.commands.executeCommand(`${SYMBOLS_VIEW_ID}.focus`);
  }));

  // 标题栏“设置”：直接打开本扩展的设置页。
  context.subscriptions.push(vscode.commands.registerCommand('vbadminx.openSettings', async () => {
    try {
      await vscode.commands.executeCommand('workbench.action.openSettings', `@ext:${context.extension.id}`);
    } catch (error) {
      void vscode.window.showErrorMessage(t('settings.openFailed', error instanceof Error ? error.message : t('error.generic')));
    }
  }));

  // 标题栏“视图选项”：分步菜单式快捷选择（范围 / 排序 / 过滤）。
  context.subscriptions.push(vscode.commands.registerCommand('vbadminx.symbolsOptions', async () => {
    const current = provider.getOptions();
    const pick = async (title: string, items: Array<{ label: string; value: string }>, active: string): Promise<string | undefined> => {
      const chosen = await vscode.window.showQuickPick(
        items.map(item => ({ label: item.label, value: item.value, description: item.value === active ? t('symbols.options.current') : undefined })),
        { title, placeHolder: title, ignoreFocusOut: true },
      );
      return chosen?.value;
    };
    const scope = await pick(t('symbols.options.scope.title'), [
      { label: t('symbols.options.scope.currentFile'), value: 'currentFile' },
      { label: t('symbols.options.scope.workspace'), value: 'workspace' },
    ], current.scope);
    if (scope === undefined) { return; }
    const sort = await pick(t('symbols.options.sort.title'), [
      { label: t('symbols.options.sort.position'), value: 'position' },
      { label: t('symbols.options.sort.name'), value: 'name' },
    ], current.sort);
    if (sort === undefined) { return; }
    const filter = await pick(t('symbols.options.filter.title'), [
      { label: t('symbols.options.filter.all'), value: 'all' },
      { label: t('symbols.options.filter.functions'), value: 'functions' },
      { label: t('symbols.options.filter.variables'), value: 'variables' },
    ], current.filter);
    if (filter === undefined) { return; }
    await provider.setOptions({ scope: scope as Scope, sort: sort as SortMode, filter: filter as FilterMode });
  }));

  // 树项右键菜单动作。
  context.subscriptions.push(vscode.commands.registerCommand('vbadminx.symbolsFindReferences', async () => {
    const editor = vscode.window.activeTextEditor;
    if (!editor) { return; }
    await vscode.commands.executeCommand('editor.action.referenceSearch.trigger');
  }));
  context.subscriptions.push(vscode.commands.registerCommand('vbadminx.symbolsRename', async () => {
    await vscode.commands.executeCommand('editor.action.rename');
  }));

  // 编辑器变化与光标移动 → 树中同步显示当前符号。
  const syncFrom = async (): Promise<void> => {
    const editor = vscode.window.activeTextEditor;
    if (!editor || provider.isSuppressed) { return; }
    await revealCurrent(provider, view, editor.document.uri, editor.selection.active);
  };
  context.subscriptions.push(
    vscode.window.onDidChangeActiveTextEditor(() => { provider.scheduleRefresh(); void syncFrom(); }),
    vscode.window.onDidChangeTextEditorSelection(() => { void syncFrom(); }),
    vscode.workspace.onDidChangeTextDocument(() => provider.scheduleRefresh()),
    vscode.workspace.onDidSaveTextDocument(() => provider.refresh()),
  );
  autoExpandView(context);
  void syncFrom();
}

async function revealCurrent(
  provider: SymbolsViewProvider,
  view: vscode.TreeView<Node>,
  uri: vscode.Uri,
  position: vscode.Position,
): Promise<void> {
  // getChildren 会重建缓存，先触发一次读取。
  await provider.getChildren();
  const entry = provider.entryAt(uri, position);
  if (!entry) { return; }
  try {
    await view.reveal(entry, { select: true, focus: false, expand: true });
  } catch {
    // 视图未聚焦或项不可见时忽略。
  }
}
