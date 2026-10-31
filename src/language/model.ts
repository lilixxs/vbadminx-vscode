import type { Range, Diagnostic, FoldingRange } from 'vscode-languageserver/node';

export type ModuleKind = 'standard' | 'class' | 'document' | 'form';
export type VbaSymbolKind = 'module' | 'class' | 'sub' | 'function' | 'property' | 'variable' | 'parameter' | 'constant' | 'enum' | 'enumMember' | 'type' | 'field' | 'event';
export interface VbaToken {
  text: string;
  kind: 'identifier' | 'keyword' | 'string' | 'number' | 'comment' | 'operator';
  line: number;
  start: number;
  end: number;
  /** 方括号转义标识符（`[Dim]`），词法文本不含括号。 */
  escaped?: boolean;
}
export interface VbaParameter { name: string; typeName: string; optional: boolean; byRef: boolean; paramArray?: boolean; array?: boolean; }
export interface VbaSymbol {
  id: string;
  name: string;
  kind: VbaSymbolKind;
  uri: string;
  range: Range;
  selectionRange: Range;
  scope?: string;
  container?: string;
  typeName?: string;
  parameters?: VbaParameter[];
  access: 'public' | 'private' | 'friend';
  signature: string;
  documentation?: string;
  eventHandler?: boolean;
  external?: boolean;
}
export interface ParseOptions {
  moduleName?: string;
  moduleKind?: ModuleKind;
  constants?: Record<string, boolean | number | string>;
  targetBits?: 32 | 64;
}
export interface ParsedModule {
  uri: string;
  text: string;
  lines: string[];
  name: string;
  kind: ModuleKind;
  explicit: boolean;
  tokens: VbaToken[];
  symbols: VbaSymbol[];
  diagnostics: Diagnostic[];
  folds: FoldingRange[];
  protectedLines: Set<number>;
  inactiveLines: Set<number>;
  unknownConditional: boolean;
}
export interface PortableTypeIndex {
  formatVersion: 1;
  library: { name: string; guid: string; major: number; minor: number; lcid: number; bitness: 32 | 64; };
  types: Array<{
    name: string;
    kind: 'class' | 'interface' | 'enum' | 'module';
    documentation?: string;
    members: Array<{ name: string; kind: 'method' | 'property' | 'constant' | 'event'; returnType?: string; parameters?: VbaParameter[]; documentation?: string; }>;
  }>;
}
export interface LanguageContext {
  modules: ParsedModule[];
  typeIndexes: PortableTypeIndex[];
  /** 宿主全局对象与常量（例如 Excel 的 ThisWorkbook、xlUp），由宿主模型提供。 */
  hostGlobals?: VbaSymbol[];
}
