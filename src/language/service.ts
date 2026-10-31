import {
  CompletionItemKind, SymbolKind, DiagnosticSeverity, InsertTextFormat,
  type Position, type Range, type Diagnostic, type CompletionItem, type Location, type Hover,
  type SignatureHelp, type DocumentSymbol, type SymbolInformation, type TextEdit,
} from 'vscode-languageserver/node';
import type { ParsedModule, VbaSymbol, VbaToken, LanguageContext, VbaParameter } from './model';
import { KEYWORDS } from './parser';

const key = (name: string): string => name.replace(/[$%&!#@^]$/, '').toLowerCase();
const rangeOf = (token: VbaToken): Range => ({ start: { line: token.line, character: token.start }, end: { line: token.line, character: token.end } });
const primitive = new Set(['byte', 'boolean', 'integer', 'long', 'longlong', 'longptr', 'single', 'double', 'currency', 'decimal', 'date', 'string', 'variant', 'object', 'any']);
const builtinData: Array<[string, string, string]> = [
  ['MsgBox', 'Function MsgBox(Prompt As Variant, Optional Buttons As Long, Optional Title As String) As Long', '显示消息框。'],
  ['InputBox', 'Function InputBox(Prompt As String, Optional Title As String, Optional Default As String) As String', '读取用户输入。'],
  ['Len', 'Function Len(Expression As Variant) As Long', '返回字符串或变量的长度。'],
  ['CStr', 'Function CStr(Expression As Variant) As String', '转换为字符串。'],
  ['CLng', 'Function CLng(Expression As Variant) As Long', '转换为 Long。'],
  ['CInt', 'Function CInt(Expression As Variant) As Integer', '转换为 Integer。'],
  ['CDbl', 'Function CDbl(Expression As Variant) As Double', '转换为 Double。'],
  ['CBool', 'Function CBool(Expression As Variant) As Boolean', '转换为 Boolean。'],
  ['IsNull', 'Function IsNull(Expression As Variant) As Boolean', '检查 Null。'],
  ['IsMissing', 'Function IsMissing(ArgName As Variant) As Boolean', '检查可选 Variant 参数是否缺省。'],
  ['EOF', 'Function EOF(FileNumber As Integer) As Boolean', '检查文件结束。'],
  ['Kill', 'Sub Kill(PathName As String)', '删除文件。'],
  ['IsEmpty', 'Function IsEmpty(Expression As Variant) As Boolean', '检查未初始化的 Variant。'],
  ['IsNumeric', 'Function IsNumeric(Expression As Variant) As Boolean', '检查数值表达式。'],
  ['LBound', 'Function LBound(ArrayName As Variant, Optional Dimension As Long) As Long', '数组下界。'],
  ['UBound', 'Function UBound(ArrayName As Variant, Optional Dimension As Long) As Long', '数组上界。'],
  ['Array', 'Function Array(ParamArray Arguments As Variant) As Variant', '创建数组。'],
  ['CreateObject', 'Function CreateObject(Class As String, Optional ServerName As String) As Object', '创建 Automation 对象；成员类型可能需要补充声明。'],
  ['GetObject', 'Function GetObject(Optional PathName As String, Optional Class As String) As Object', '取得 Automation 对象。'],
  ['Left', 'Function Left(String As String, Length As Long) As String', '返回左侧字符。'],
  ['Right', 'Function Right(String As String, Length As Long) As String', '返回右侧字符。'],
  ['Mid', 'Function Mid(String As String, Start As Long, Optional Length As Long) As String', '返回子字符串。'],
  ['Trim', 'Function Trim(String As String) As String', '去除两端空格。'],
  ['Replace', 'Function Replace(Expression As String, Find As String, Replace As String) As String', '替换字符串内容。'],
  ['Split', 'Function Split(Expression As String, Optional Delimiter As String) As Variant', '拆分字符串。'],
  ['Join', 'Function Join(SourceArray As Variant, Optional Delimiter As String) As String', '合并字符串数组。'],
  ['InStr', 'Function InStr(Start As Long, String1 As String, String2 As String) As Long', '查找子字符串。'],
  ['Format', 'Function Format(Expression As Variant, Optional Format As String) As String', '格式化值。'],
  ['Debug', 'VBA.Debug', 'VBA 调试对象。'], ['Err', 'VBA.Err', 'VBA 错误对象。'],
  ['Collection', 'Class VBA.Collection', 'VBA 内建集合。'],
  ['vbCrLf', 'Const vbCrLf As String', '回车换行。'], ['vbCr', 'Const vbCr As String', '回车。'], ['vbLf', 'Const vbLf As String', '换行。'],
  ['vbTab', 'Const vbTab As String', '制表符。'], ['vbNullString', 'Const vbNullString As String', '空字符串常量。'],
  ['vbOKOnly', 'Const vbOKOnly As Long', '仅确定按钮。'], ['vbInformation', 'Const vbInformation As Long', '信息图标。'],
  ['vbExclamation', 'Const vbExclamation As Long', '警告图标。'], ['vbYesNo', 'Const vbYesNo As Long', '是与否按钮。'],
];
for (const [name, signature] of [
  ['IIf', 'Function IIf(Expression As Boolean, TruePart As Variant, FalsePart As Variant) As Variant'],
  ['Now', 'Function Now() As Date'], ['Timer', 'Function Timer() As Single'], ['Date', 'Function Date() As Date'], ['Time', 'Function Time() As Date'],
  ['DoEvents', 'Function DoEvents() As Integer'], ['UCase', 'Function UCase(String As String) As String'], ['LCase', 'Function LCase(String As String) As String'],
  ['Chr', 'Function Chr(CharCode As Long) As String'], ['ChrW', 'Function ChrW(CharCode As Long) As String'], ['Asc', 'Function Asc(String As String) As Integer'], ['AscW', 'Function AscW(String As String) As Integer'],
  ['Int', 'Function Int(Number As Double) As Double'], ['Fix', 'Function Fix(Number As Double) As Double'], ['Round', 'Function Round(Number As Double, Optional NumDigitsAfterDecimal As Long) As Double'],
  ['Abs', 'Function Abs(Number As Double) As Double'], ['Sqr', 'Function Sqr(Number As Double) As Double'], ['Rnd', 'Function Rnd(Optional Number As Single) As Single'],
  ['Dir', 'Function Dir(Optional PathName As String, Optional Attributes As Long) As String'], ['FreeFile', 'Function FreeFile(Optional RangeNumber As Variant) As Integer'],
  ['Str', 'Function Str(Number As Variant) As String'], ['Val', 'Function Val(String As String) As Double'], ['TypeName', 'Function TypeName(VarName As Variant) As String'],
  ['VarType', 'Function VarType(VarName As Variant) As Integer'], ['IsObject', 'Function IsObject(Expression As Variant) As Boolean'], ['IsArray', 'Function IsArray(VarName As Variant) As Boolean'],
  ['CDate', 'Function CDate(Expression As Variant) As Date'], ['CSng', 'Function CSng(Expression As Variant) As Single'], ['CByte', 'Function CByte(Expression As Variant) As Byte'],
  ['DateAdd', 'Function DateAdd(Interval As String, Number As Double, Date As Date) As Date'], ['DateDiff', 'Function DateDiff(Interval As String, Date1 As Date, Date2 As Date) As Long'],
  ['InStrRev', 'Function InStrRev(StringCheck As String, StringMatch As String, Optional Start As Long) As Long'], ['StrComp', 'Function StrComp(String1 As String, String2 As String, Optional Compare As Long) As Integer'],
  ['LTrim', 'Function LTrim(String As String) As String'], ['RTrim', 'Function RTrim(String As String) As String'], ['Space', 'Function Space(Number As Long) As String'],
] as Array<[string, string]>) { builtinData.push([name, signature, 'VBA 标准库内建函数。']); }
for (const name of ['vbCritical', 'vbQuestion', 'vbOKCancel', 'vbYesNoCancel', 'vbDefaultButton1', 'vbDefaultButton2', 'vbOK', 'vbCancel', 'vbYes', 'vbNo', 'vbTextCompare', 'vbBinaryCompare', 'vbUseCompareOption', 'vbNullChar', 'vbBack', 'vbFormFeed', 'vbVerticalTab', 'vbObjectError', 'vbTrue', 'vbFalse', 'vbUseDefault', 'vbNormal', 'vbReadOnly', 'vbHidden', 'vbSystem', 'vbDirectory', 'vbArchive']) { builtinData.push([name, `Const ${name} As Variant`, 'VBA 标准库内建常量。']); }
const externalHint = /^(?:xl[A-Z]|mso[A-Z]|ad[A-Z]|wd[A-Z]|ac[A-Z])/;
const hostNames = new Set(['application', 'thisworkbook', 'activeworkbook', 'activesheet', 'activewindow', 'workbooks', 'worksheets', 'sheets', 'range', 'cells', 'selection', 'rows', 'columns', 'worksheetfunction', 'charts', 'names', 'vba']);
const DIAGNOSTIC_SOURCE = 'VBAdminx';
function parameters(signature: string): VbaParameter[] {
  const raw = /\((.*)\)/.exec(signature)?.[1];
  return raw ? raw.split(',').map(part => ({ name: part.trim().split(/\s+/).filter(word => !/^(optional|byval|byref|paramarray)$/i.test(word))[0] ?? '', typeName: /\bAs\s+(\w+)/i.exec(part)?.[1] ?? 'Variant', optional: /\bOptional\b/i.test(part), byRef: !/\bByVal\b/i.test(part), paramArray: /\bParamArray\b/i.test(part) })) : [];
}
const builtins: VbaSymbol[] = builtinData.map(([name, signature, documentation]) => ({
  id: `builtin:${name.toLowerCase()}`, name, signature, documentation, kind: /^Const/.test(signature) ? 'constant' : /^Function/.test(signature) ? 'function' : /^Sub/.test(signature) ? 'sub' : 'class',
  uri: `vba-library:/VBA/${name}`, range: { start: { line: 0, character: 0 }, end: { line: 0, character: 0 } }, selectionRange: { start: { line: 0, character: 0 }, end: { line: 0, character: name.length } },
  access: 'public', typeName: /\) As (\w+)/.exec(signature)?.[1] ?? name, parameters: /^(?:Function|Sub)/.test(signature) ? parameters(signature) : undefined, external: true,
}));

const extraMembers: Record<string, Array<[string, string, string]>> = {
  debug: [['Print', 'Sub Print(ParamArray OutputList As Variant)', '输出到立即窗口。'], ['Assert', 'Sub Assert(Expression As Boolean)', '检查断言。']],
  err: [['Number', 'Property Number As Long', '错误编号。'], ['Description', 'Property Description As String', '错误说明。'], ['Clear', 'Sub Clear()', '清除错误。'], ['Raise', 'Sub Raise(Number As Long, Optional Source As String, Optional Description As String)', '引发错误。']],
  collection: [['Add', 'Sub Add(Item As Variant, Optional Key As String)', '加入成员。'], ['Item', 'Function Item(Index As Variant) As Variant', '取得成员。'], ['Count', 'Property Count As Long', '成员数量。'], ['Remove', 'Sub Remove(Index As Variant)', '删除成员。']],
};

function symbolicKind(kind: VbaSymbol['kind']): SymbolKind {
  return ({ module: SymbolKind.Module, class: SymbolKind.Class, sub: SymbolKind.Method, function: SymbolKind.Function, property: SymbolKind.Property, variable: SymbolKind.Variable, parameter: SymbolKind.Variable, constant: SymbolKind.Constant, enum: SymbolKind.Enum, enumMember: SymbolKind.EnumMember, type: SymbolKind.Struct, field: SymbolKind.Field, event: SymbolKind.Event })[kind];
}
function completionKind(kind: VbaSymbol['kind']): CompletionItemKind {
  return ({ module: CompletionItemKind.Module, class: CompletionItemKind.Class, sub: CompletionItemKind.Method, function: CompletionItemKind.Function, property: CompletionItemKind.Property, variable: CompletionItemKind.Variable, parameter: CompletionItemKind.Variable, constant: CompletionItemKind.Constant, enum: CompletionItemKind.Enum, enumMember: CompletionItemKind.EnumMember, type: CompletionItemKind.Struct, field: CompletionItemKind.Field, event: CompletionItemKind.Event })[kind];
}
function inside(position: Position, range: Range): boolean {
  return (position.line > range.start.line || position.line === range.start.line && position.character >= range.start.character) && (position.line < range.end.line || position.line === range.end.line && position.character <= range.end.character);
}
const procedureKinds = new Set(['sub', 'function', 'property']);

export class VbaService {
  private externalCache?: VbaSymbol[];
  private readonly lineCache = new WeakMap<ParsedModule, Map<number, VbaToken[]>>();
  private readonly visibleCache = new Map<string, VbaSymbol[]>();
  private readonly memberCache = new Map<string, VbaSymbol[]>();
  private readonly resolved = new WeakMap<VbaToken, VbaSymbol | null>();
  constructor(readonly context: LanguageContext) {}

  private lineTokens(module: ParsedModule, line: number): VbaToken[] {
    let lines = this.lineCache.get(module);
    if (!lines) {
      lines = new Map();
      for (const token of module.tokens) { if (token.kind !== 'comment') { const list = lines.get(token.line) ?? []; list.push(token); lines.set(token.line, list); } }
      this.lineCache.set(module, lines);
    }
    return lines.get(line) ?? [];
  }
  private withOwner(module: ParsedModule, line: number): { expression: string; line: number } | undefined {
    const stack: Array<{ expression: string; line: number }> = [];
    for (let index = 0; index <= line; index++) {
      if (module.inactiveLines.has(index) || module.protectedLines.has(index)) { continue; }
      const statement = (module.lines[index] ?? '').trim();
      const opening = /^With\s+([\p{L}_][\p{L}\p{N}_.]*)\s*(?:'.*)?$/iu.exec(statement);
      if (/^End\s+With\b/i.test(statement)) { stack.pop(); }
      else if (opening) { stack.push({ expression: opening[1]!, line: index }); }
    }
    return stack.at(-1);
  }

  tokenAt(module: ParsedModule, position: Position): VbaToken | undefined {
    return module.tokens.find(token => token.line === position.line && token.start <= position.character && token.end >= position.character && token.kind === 'identifier');
  }
  scope(module: ParsedModule, position: Position): VbaSymbol | undefined {
    return module.symbols.find(symbol => procedureKinds.has(symbol.kind) && inside(position, symbol.range));
  }
  private external(): VbaSymbol[] {
    if (this.externalCache) { return this.externalCache; }
    const values: VbaSymbol[] = [];
    for (const index of this.context.typeIndexes) {
      for (const type of index.types) {
        const uri = `vba-library:/${encodeURIComponent(index.library.name)}/${encodeURIComponent(type.name)}`;
        const base = { uri, range: { start: { line: 0, character: 0 }, end: { line: 0, character: 0 } }, selectionRange: { start: { line: 0, character: 0 }, end: { line: 0, character: 0 } }, access: 'public' as const, external: true };
        values.push({ ...base, id: `${uri}:type`, name: type.name, typeName: type.name, kind: type.kind === 'enum' ? 'enum' : 'class', signature: `${type.kind} ${index.library.name}.${type.name}`, documentation: type.documentation });
        for (const member of type.members) {
          values.push({ ...base, id: `${uri}:${member.name}`, name: member.name, kind: member.kind === 'property' ? 'property' : member.kind === 'constant' ? 'constant' : member.kind === 'event' ? 'event' : 'function', container: type.name,
            signature: `${member.name}${member.parameters ? `(${member.parameters.map(param => `${param.optional ? 'Optional ' : ''}${param.name} As ${param.typeName}`).join(', ')})` : ''} As ${member.returnType ?? 'Variant'}`,
            typeName: member.returnType, parameters: member.parameters, documentation: `${index.library.name} ${index.library.major}.${index.library.minor}\n${member.documentation ?? ''}` });
        }
      }
    }
    this.externalCache = values;
    return values;
  }
  private membersOf(typeName: string, module: ParsedModule): VbaSymbol[] {
    const cleanName = typeName.replace(/\(\)$/, '');
    const typeKey = key(cleanName.split('.').at(-1) ?? cleanName);
    const library = cleanName.includes('.') ? key(cleanName.split('.')[0]!) : undefined;
    const cacheKey = `${module.uri}:${key(cleanName)}`;
    const cached = this.memberCache.get(cacheKey);
    if (cached) { return cached; }
    if (!library && this.context.typeIndexes.filter(index => index.types.some(type => key(type.name) === typeKey)).length > 1 && !this.context.modules.some(item => key(item.name) === typeKey)) {
      this.memberCache.set(cacheKey, []);
      return [];
    }
    const sourceModule = this.context.modules.find(item => key(item.name) === typeKey);
    const members = sourceModule?.symbols.filter(symbol => !symbol.scope && !['module', 'class'].includes(symbol.kind) && (symbol.access !== 'private' || sourceModule.uri === module.uri)) ?? [];
    for (const item of this.context.modules) {
      members.push(...item.symbols.filter(symbol => symbol.container && key(symbol.container) === typeKey && (!symbol.scope || symbol.kind === 'field' || symbol.kind === 'enumMember')));
    }
    members.push(...this.external().filter(symbol => symbol.container && key(symbol.container) === typeKey && (!library || key(decodeURIComponent(new URL(symbol.uri).pathname.split('/')[1] ?? '')) === library)));
    for (const [name, signature, documentation] of extraMembers[typeKey] ?? []) {
      members.push({ ...builtins[0]!, id: `builtin:${typeKey}.${name}`, name, signature, documentation, kind: signature.startsWith('Property') ? 'property' : signature.startsWith('Sub') ? 'sub' : 'function', typeName: /As (\w+)$/.exec(signature)?.[1], parameters: parameters(signature) });
    }
    this.memberCache.set(cacheKey, members);
    return members;
  }
  private chooseMember(typeName: string, name: string, module: ParsedModule): VbaSymbol | undefined {
    const choices = this.membersOf(typeName, module).filter(symbol => key(symbol.name) === key(name));
    if (choices.length === 1) { return choices[0]; }
    if (choices.length > 1 && choices.every(symbol => symbol.kind === 'property' && symbol.uri === choices[0]!.uri)) {
      return choices.find(symbol => /\bProperty\s+Get\b/i.test(symbol.signature)) ?? choices[0];
    }
    return undefined;
  }
  private valueType(symbol: VbaSymbol): string {
    const typeName = symbol.typeName ?? symbol.name;
    // COM 返回类型在所属库内查找，避免 Excel.Application 与其他库同名类型混淆。
    if (symbol.external && !typeName.includes('.') && symbol.uri.startsWith('vba-library:')) {
      const library = decodeURIComponent(new URL(symbol.uri).pathname.split('/')[1] ?? '');
      const index = this.context.typeIndexes.find(item => key(item.library.name) === key(library));
      if (index?.types.some(item => key(item.name) === key(typeName))) { return `${library}.${typeName}`; }
    }
    return typeName;
  }
  private expressionOwner(module: ParsedModule, tokens: VbaToken[], end: number, position: Position, depth = 0): VbaSymbol | undefined {
    if (depth > 24 || end < 0) { return undefined; }
    const last = tokens[end]!;
    if (last.text === ')') {
      let nesting = 1;
      for (let index = end - 1; index >= 0; index--) {
        if (tokens[index]!.text === ')') { nesting++; }
        else if (tokens[index]!.text === '(' && --nesting === 0) { return this.expressionOwner(module, tokens, index - 1, position, depth + 1); }
      }
      return undefined;
    }
    if (last.text === '.') {
      const withBlock = this.withOwner(module, position.line);
      if (!withBlock) { return undefined; }
      const expression = this.lineTokens(module, withBlock.line).filter(token => token.start >= (module.lines[withBlock.line] ?? '').indexOf(withBlock.expression));
      return this.expressionOwner(module, expression, expression.length - 1, { line: withBlock.line, character: 0 }, depth + 1);
    }
    if (key(last.text) === 'me') { return module.symbols.find(symbol => ['module', 'class'].includes(symbol.kind)); }
    if (tokens[end - 1]?.text === '.') {
      const owner = this.expressionOwner(module, tokens, end - 2, position, depth + 1);
      if (!owner) { return undefined; }
      return this.chooseMember(this.valueType(owner), last.text, module);
    }
    const visible = this.visible(module, position).filter(symbol => key(symbol.name) === key(last.text));
    const local = visible.find(symbol => !!symbol.scope) ?? visible.find(symbol => symbol.uri === module.uri);
    if (local) { return local; }
    const unique = [...new Map(visible.map(symbol => [symbol.id, symbol])).values()];
    if (unique.length === 1) { return unique[0]; }
    // 已限定的类型库名称可以作为表达式的命名空间属主。
    if (this.context.typeIndexes.some(index => key(index.library.name) === key(last.text))) {
      return { ...builtins[0]!, id: `library:${last.text}`, name: last.text, kind: 'module', typeName: `@library:${last.text}` };
    }
    return undefined;
  }
  visible(module: ParsedModule, position: Position): VbaSymbol[] {
    const scope = this.scope(module, position);
    const cacheKey = `${module.uri}:${scope?.id ?? ''}`;
    const cached = this.visibleCache.get(cacheKey);
    if (cached) { return cached; }
    const local = module.symbols.filter(symbol => symbol.scope && (symbol.scope === scope?.id || key(symbol.scope) === key(scope?.name ?? '')));
    const own = module.symbols.filter(symbol => !symbol.scope);
    const shared = this.context.modules.filter(item => item.uri !== module.uri).flatMap(item => item.symbols.filter(symbol => !symbol.scope && (symbol.kind === 'module' || symbol.kind === 'class' || item.kind === 'standard' && symbol.access === 'public' || ['enum', 'type'].includes(symbol.kind) && symbol.access === 'public')));
    const externalVisible = this.external().filter(symbol => !symbol.container || symbol.kind === 'constant');
    const hostGlobals = (this.context.hostGlobals ?? []).filter(symbol =>
      ![local, own, shared, externalVisible].some(list => list.some(item => key(item.name) === key(symbol.name))));
    const result = [...local, ...own, ...shared, ...externalVisible, ...hostGlobals, ...builtins];
    this.visibleCache.set(cacheKey, result);
    return result;
  }

  resolve(module: ParsedModule, token: VbaToken): VbaSymbol | undefined {
    if (this.resolved.has(token)) { return this.resolved.get(token) ?? undefined; }
    const symbol = this.resolveUncached(module, token);
    this.resolved.set(token, symbol ?? null);
    return symbol;
  }
  private resolveUncached(module: ParsedModule, token: VbaToken): VbaSymbol | undefined {
    if (module.inactiveLines.has(token.line) || module.protectedLines.has(token.line)) { return undefined; }
    const declaration = module.symbols.find(symbol => symbol.selectionRange.start.line === token.line && symbol.selectionRange.start.character === token.start);
    if (declaration) { return declaration; }
    const sameLine = this.lineTokens(module, token.line);
    const index = sameLine.indexOf(token);
    // 命名实参属于被调用过程的参数，不能绑定到调用方的同名局部变量。
    if (sameLine[index + 1]?.text === ':=') {
      let depth = 0;
      for (let offset = index - 1; offset >= 0; offset--) {
        if (sameLine[offset]!.text === ')') { depth++; }
        if (sameLine[offset]!.text === '(') {
          if (depth > 0) { depth--; continue; }
          const callee = sameLine[offset - 1];
          const procedure = callee && this.resolve(module, callee);
          const declaringModule = procedure && this.context.modules.find(item => item.uri === procedure.uri);
          return declaringModule?.symbols.find(symbol => symbol.kind === 'parameter' && symbol.scope === procedure?.id && key(symbol.name) === key(token.text));
        }
      }
      return undefined;
    }
    if (sameLine[index - 1]?.text === '.') {
      const owner = index === 1 || ['=', '(', ',', ':'].includes(sameLine[index - 2]?.text ?? '')
        ? this.expressionOwner(module, sameLine, index - 1, { line: token.line, character: token.start })
        : this.expressionOwner(module, sameLine, index - 2, { line: token.line, character: token.start });
      if (!owner) { return undefined; }
      if (owner.typeName?.startsWith('@library:')) {
        const library = key(owner.typeName.slice(9));
        return this.external().find(symbol => !symbol.container && key(symbol.name) === key(token.text) && key(decodeURIComponent(new URL(symbol.uri).pathname.split('/')[1] ?? '')) === library);
      }
      return this.chooseMember(this.valueType(owner), token.text, module);
    }
    const candidates = this.visible(module, { line: token.line, character: token.start }).filter(symbol => key(symbol.name) === key(token.text));
    const scoped = candidates.filter(symbol => symbol.scope);
    if (scoped.length) { return scoped[0]; }
    const own = candidates.filter(symbol => symbol.uri === module.uri);
    if (own.length) { return own.find(symbol => symbol.kind !== 'property') ?? own.find(symbol => /\bProperty\s+Get\b/i.test(symbol.signature)) ?? own[0]; }
    const unique = [...new Map(candidates.map(symbol => [symbol.id, symbol])).values()];
    return unique.length === 1 ? unique[0] : undefined;
  }

  completions(module: ParsedModule, position: Position): CompletionItem[] {
    const prefix = (module.lines[position.line] ?? '').slice(0, position.character);
    if (module.tokens.some(token => token.line === position.line && token.start < position.character && token.end >= position.character && ['comment', 'string'].includes(token.kind))) { return []; }
    const member = /\.([\p{L}\p{N}_]*)$/u.exec(prefix);
    let symbols: VbaSymbol[];
    if (member) {
      const dot = prefix.length - member[1]!.length - 1;
      const before = this.lineTokens(module, position.line).filter(token => token.end <= dot);
      const implicit = before.length === 0 || ['=', '(', ',', ':'].includes(before.at(-1)?.text ?? '');
      const dotToken = this.lineTokens(module, position.line).find(token => token.start === dot);
      const owner = implicit && dotToken
        ? this.expressionOwner(module, [dotToken], 0, position)
        : this.expressionOwner(module, before, before.length - 1, position);
      symbols = owner ? owner.typeName?.startsWith('@library:')
        ? this.external().filter(symbol => !symbol.container && key(decodeURIComponent(new URL(symbol.uri).pathname.split('/')[1] ?? '')) === key(owner.typeName!.slice(9)))
        : this.membersOf(this.valueType(owner), module) : [];
    } else { symbols = this.visible(module, position); }
    const dedup = new Map<string, CompletionItem>();
    for (const symbol of symbols) {
      if (!dedup.has(key(symbol.name))) { dedup.set(key(symbol.name), { label: symbol.name, kind: completionKind(symbol.kind), detail: symbol.signature, documentation: symbol.documentation, insertText: symbol.name, insertTextFormat: InsertTextFormat.PlainText, sortText: `${symbol.scope ? '0' : symbol.external ? '2' : '1'}${symbol.name}` }); }
    }
    if (!member) { for (const keyword of KEYWORDS) { if (!dedup.has(keyword)) { dedup.set(keyword, { label: keyword[0]!.toUpperCase() + keyword.slice(1), kind: CompletionItemKind.Keyword, sortText: `3${keyword}` }); } } }
    return [...dedup.values()];
  }

  hover(module: ParsedModule, position: Position): Hover | null {
    const token = this.tokenAt(module, position);
    const symbol = token && this.resolve(module, token);
    return symbol ? { contents: { kind: 'markdown', value: `\`\`\`vb\n${symbol.signature}\n\`\`\`\n${symbol.documentation ?? ''}\n\n${symbol.external ? '外部库 / VBA 内建符号' : this.context.modules.find(item => item.uri === symbol.uri)?.name ?? module.name}` }, range: rangeOf(token!) } : null;
  }
  definition(module: ParsedModule, position: Position): Location[] {
    const token = this.tokenAt(module, position);
    const symbol = token && this.resolve(module, token);
    return symbol ? [{ uri: symbol.uri, range: symbol.selectionRange }] : [];
  }
  references(symbol: VbaSymbol, includeDeclaration = true, onModuleStart?: () => void): Location[] {
    const result: Location[] = [];
    for (const module of this.context.modules) {
      onModuleStart?.();
      for (const token of module.tokens) {
        if (token.kind !== 'identifier' || key(token.text) !== key(symbol.name)) { continue; }
        const resolved = this.resolve(module, token);
        if (resolved?.id !== symbol.id) { continue; }
        if (!includeDeclaration && symbol.uri === module.uri && token.line === symbol.selectionRange.start.line && token.start === symbol.selectionRange.start.character) { continue; }
        result.push({ uri: module.uri, range: rangeOf(token) });
      }
    }
    return result;
  }
  rename(module: ParsedModule, position: Position, newName: string): Record<string, TextEdit[]> {
    if (!/^[\p{L}_][\p{L}\p{N}_]*$/u.test(newName) || KEYWORDS.has(key(newName))) { throw new Error('新名称必须是有效的 VBA 标识符，且不能是关键字。'); }
    const token = this.tokenAt(module, position);
    const symbol = token && this.resolve(module, token);
    if (!symbol || symbol.external || ['module', 'class'].includes(symbol.kind)) { throw new Error('该符号不能安全重命名。模块文件改名需单独同步导出信息。'); }
    const declaringModule = this.context.modules.find(item => item.uri === symbol.uri) ?? module;
    if (this.visible(declaringModule, symbol.selectionRange.start).some(other => other.id !== symbol.id && key(other.name) === key(newName) && (other.scope === symbol.scope || !other.scope && !symbol.scope))) { throw new Error('新名称与已有符号冲突。'); }
    const changes: Record<string, TextEdit[]> = {};
    const group = symbol.kind === 'property' ? declaringModule.symbols.filter(item => item.kind === 'property' && key(item.name) === key(symbol.name)) : [symbol];
    const seen = new Set<string>();
    for (const accessor of group) {
      for (const location of this.references(accessor)) {
        const identity = `${location.uri}:${location.range.start.line}:${location.range.start.character}`;
        if (seen.has(identity)) { continue; }
        seen.add(identity);
        const sourceModule = this.context.modules.find(item => item.uri === location.uri);
        if (sourceModule && this.visible(sourceModule, location.range.start).some(other => other.id !== symbol.id && !(symbol.kind === 'property' && other.kind === 'property' && other.uri === symbol.uri && key(other.name) === key(symbol.name)) && key(other.name) === key(newName))) {
          throw new Error('新名称会在某个引用位置被同名局部变量或其他符号遮蔽，已取消重命名。');
        }
        const originalToken = sourceModule?.tokens.find(item => item.line === location.range.start.line && item.start === location.range.start.character);
        const suffix = /[$%&!#@^]$/.exec(originalToken?.text ?? '')?.[0] ?? '';
        (changes[location.uri] ??= []).push({ range: location.range, newText: newName + suffix });
      }
    }
    return changes;
  }
  renameCandidates(symbol: VbaSymbol, newName: string): Array<{ uri: string; range: Range; newText: string; reason: string }> {
    const candidates: Array<{ uri: string; range: Range; newText: string; reason: string }> = [];
    for (const module of this.context.modules) {
      for (const token of module.tokens) {
        if (token.kind === 'string' && /\b(?:Run|OnAction)\b/i.test(module.lines[token.line] ?? '')) {
          const value = token.text.slice(1, -1).replaceAll('""', '"');
          const match = /^(.*[.!])?([^.!]+)$/.exec(value);
          if (key(match?.[2] ?? '') === key(symbol.name)) { candidates.push({ uri: module.uri, range: rangeOf(token), newText: `"${(match?.[1] ?? '') + newName}"`, reason: 'Run/OnAction 字符串可能指向该过程；无法从源码证明。' }); }
        } else if (token.kind === 'identifier' && key(token.text) === key(symbol.name) && !this.resolve(module, token)) {
          const before = (module.lines[token.line] ?? '').slice(0, token.start);
          if (/\.\s*$/.test(before)) { candidates.push({ uri: module.uri, range: rangeOf(token), newText: newName, reason: '成员类型未知，同名成员可能指向其他声明。' }); }
        }
      }
    }
    return candidates;
  }

  signature(module: ParsedModule, position: Position): SignatureHelp | null {
    const line = (module.lines[position.line] ?? '').slice(0, position.character);
    const match = /([\p{L}_][\p{L}\p{N}_]*)\s*\(([^()]*)$/u.exec(line) ?? /^\s*(?:Call\s+)?([\p{L}_][\p{L}\p{N}_]*)\s+([^()]*)$/u.exec(line);
    if (!match) { return null; }
    const name = match[1]!;
    const token = module.tokens.find(item => item.line === position.line && key(item.text) === key(name));
    const symbol = token && this.resolve(module, token);
    if (!symbol?.parameters) { return null; }
    let commas = 0; let quoted = false;
    for (const char of match[2] ?? '') { if (char === '"') { quoted = !quoted; } else if (char === ',' && !quoted) { commas++; } }
    return { signatures: [{ label: symbol.signature, documentation: symbol.documentation, parameters: symbol.parameters.map(param => ({ label: `${param.name} As ${param.typeName}` })) }], activeSignature: 0, activeParameter: Math.min(commas, Math.max(0, symbol.parameters.length - 1)) };
  }

  documentSymbols(module: ParsedModule): DocumentSymbol[] {
    return module.symbols.filter(symbol => !symbol.scope && !['module', 'class'].includes(symbol.kind)).map(symbol => ({
      name: symbol.name, detail: symbol.signature, kind: symbolicKind(symbol.kind), range: symbol.range, selectionRange: symbol.selectionRange,
      children: module.symbols.filter(child => child.scope === symbol.id || child.scope && key(child.scope) === key(symbol.name) || child.container === symbol.name && ['field', 'enumMember'].includes(child.kind)).map(child => ({ name: child.name, detail: child.typeName, kind: symbolicKind(child.kind), range: child.range, selectionRange: child.selectionRange })),
    }));
  }
  workspaceSymbols(query: string): SymbolInformation[] {
    return this.context.modules.flatMap(module => module.symbols.filter(symbol => !symbol.scope && key(symbol.name).includes(key(query))).map(symbol => ({ name: symbol.name, kind: symbolicKind(symbol.kind), location: { uri: symbol.uri, range: symbol.selectionRange }, containerName: module.name })));
  }

  /**
   * 声明级语义校验（SEM-03/05/06/08/11、STMT-01b/02b）。
   * 只在证据确定时报错：同名检查仅针对模块级 Sub/Function（Property 各访问器同名合法）。
   */
  private declarationDiagnostics(module: ParsedModule): Diagnostic[] {
    const diagnostics: Diagnostic[] = [];
    const seenProcedures = new Map<string, VbaSymbol>();
    for (const symbol of module.symbols) {
      if (!symbol.scope && (symbol.kind === 'sub' || symbol.kind === 'function')) {
        const k = key(symbol.name);
        const previous = seenProcedures.get(k);
        if (previous !== undefined) {
          diagnostics.push({ range: symbol.selectionRange, severity: DiagnosticSeverity.Error, code: 'ambiguous-name', source: DIAGNOSTIC_SOURCE, message: `发现二义性名称：${symbol.name}。VBA 不支持重载。` });
          continue;
        }
        seenProcedures.set(k, symbol);
      }
      if (module.kind === 'standard' && symbol.access === 'friend' && !symbol.scope) {
        diagnostics.push({ range: symbol.selectionRange, severity: DiagnosticSeverity.Error, code: 'friend-in-standard', source: DIAGNOSTIC_SOURCE, message: 'Friend 成员只能出现在类模块中。' });
      }
      if (module.kind === 'class' && symbol.kind === 'type' && symbol.access === 'public') {
        diagnostics.push({ range: symbol.selectionRange, severity: DiagnosticSeverity.Error, code: 'type-public-in-class', source: DIAGNOSTIC_SOURCE, message: '类模块中用户定义类型必须声明为 Private。' });
      }
      if (symbol.scope === undefined && symbol.signature.includes('WithEvents')) {
        const typeName = key(symbol.typeName ?? '');
        if (typeName !== '' && primitive.has(typeName)) {
          diagnostics.push({ range: symbol.selectionRange, severity: DiagnosticSeverity.Error, code: 'withevents-non-object', source: DIAGNOSTIC_SOURCE, message: `WithEvents 变量 ${symbol.name} 必须为对象类型。` });
        }
      }
      const params = symbol.parameters;
      if (!params || params.length === 0) { continue; }
      let optionalSeen = false;
      let paramArraySeen = false;
      params.forEach((param, index) => {
        if (param.paramArray) {
          if (paramArraySeen) {
            diagnostics.push({ range: symbol.selectionRange, severity: DiagnosticSeverity.Error, code: 'param-array-unique', source: DIAGNOSTIC_SOURCE, message: `${symbol.name} 只能有一个 ParamArray 参数。` });
          }
          if (index !== params.length - 1) {
            diagnostics.push({ range: symbol.selectionRange, severity: DiagnosticSeverity.Error, code: 'param-array-position', source: DIAGNOSTIC_SOURCE, message: `${symbol.name} 的 ParamArray 参数必须是最后一个参数。` });
          }
          paramArraySeen = true;
          return;
        }
        if (param.optional) { optionalSeen = true; }
        else if (optionalSeen) {
          diagnostics.push({ range: symbol.selectionRange, severity: DiagnosticSeverity.Error, code: 'param-optional-order', source: DIAGNOSTIC_SOURCE, message: `${symbol.name} 的必选参数 ${param.name} 不能出现在 Optional 参数之后。` });
        }
        if (param.array === true && !param.byRef) {
          diagnostics.push({ range: symbol.selectionRange, severity: DiagnosticSeverity.Error, code: 'param-array-byref', source: DIAGNOSTIC_SOURCE, message: `数组参数 ${param.name} 必须按引用传递（不能使用 ByVal）。` });
        }
      });
    }
    return diagnostics;
  }

  diagnostics(module: ParsedModule): Diagnostic[] {
    const diagnostics: Diagnostic[] = [];
    const tokens = module.tokens;
    const missingTypes = new Set<string>();
    const missingSymbols = new Set<string>();

    // 声明级语义校验（SEM-03/05/06/08/11、STMT-01b/02b）。
    diagnostics.push(...this.declarationDiagnostics(module));

    for (const symbol of module.symbols) {
      const typeName = symbol.typeName?.replace(/\(\)$/, '');
      if (typeName && !primitive.has(key(typeName)) && !this.context.modules.some(item => key(item.name) === key(typeName)) && !this.external().some(item => !item.container && key(item.name) === key(typeName.split('.').at(-1)!)) && !builtins.some(item => key(item.name) === key(typeName)) && !module.symbols.some(item => ['type', 'enum'].includes(item.kind) && key(item.name) === key(typeName))) {
        if (!missingTypes.has(key(typeName))) { diagnostics.push({ range: symbol.selectionRange, severity: DiagnosticSeverity.Information, code: 'missing-type', source: DIAGNOSTIC_SOURCE, message: `类型信息缺失：${typeName}。请添加匹配版本的类型索引或补充声明。` }); missingTypes.add(key(typeName)); }
      }
    }
    for (const token of tokens) {
      if (token.kind !== 'identifier' || module.inactiveLines.has(token.line) || module.protectedLines.has(token.line)) { continue; }
      const lineTokens = this.lineTokens(module, token.line);
      const index = lineTokens.indexOf(token);
      const previous = lineTokens[index - 1]?.text.toLowerCase();
      const next = lineTokens[index + 1]?.text;
      if (previous === '.' || previous === 'as' || previous === 'new' || previous === 'implements' || previous === 'goto' || previous === 'gosub' || previous === 'resume' || next === ':' || next === ':=' || /^(?:attribute|version|option|#)/i.test((module.lines[token.line] ?? '').trim()) || ['output', 'append', 'binary', 'access'].includes(key(token.text)) && /\bOpen\b.+\bFor\b/i.test(module.lines[token.line] ?? '')) { continue; }
      const resolved = this.resolve(module, token);
      if (!resolved && module.explicit && this.scope(module, { line: token.line, character: token.start })) {
        if (externalHint.test(token.text) || hostNames.has(key(token.text))) {
          if (!missingSymbols.has(key(token.text))) { diagnostics.push({ range: rangeOf(token), severity: DiagnosticSeverity.Information, code: 'missing-external-symbol', source: DIAGNOSTIC_SOURCE, message: `外部符号 ${token.text} 的类型信息缺失。请加载对应宿主或引用库索引。` }); missingSymbols.add(key(token.text)); }
          continue;
        }
        if (module.kind === 'form' || module.kind === 'document') { continue; }
        if (missingTypes.size > 0) {
          if (!missingSymbols.has(key(token.text))) { diagnostics.push({ range: rangeOf(token), severity: DiagnosticSeverity.Information, code: 'missing-external-symbol', source: DIAGNOSTIC_SOURCE, message: `无法确认符号 ${token.text}：工程缺少外部类型信息。请先加载引用库后重新分析。` }); missingSymbols.add(key(token.text)); }
          continue;
        }
        diagnostics.push({ range: rangeOf(token), severity: DiagnosticSeverity.Error, code: 'undeclared', source: DIAGNOSTIC_SOURCE, message: `变量或符号 ${token.text} 未声明。`, data: { name: token.text, scope: this.scope(module, { line: token.line, character: token.start })?.id } });
      }
      const tokenSuffix = /[$%&!#@^]$/.exec(token.text)?.[0] ?? '';
      const tokenBase = tokenSuffix ? token.text.slice(0, -1) : token.text;
      if (resolved && !resolved.external && key(resolved.name) === key(token.text) && resolved.name !== tokenBase) {
        diagnostics.push({ range: rangeOf(token), severity: DiagnosticSeverity.Hint, code: 'name-case', source: DIAGNOSTIC_SOURCE, message: `声明中的名称为 ${resolved.name}。`, data: { name: resolved.name + tokenSuffix } });
      }
      if (resolved && ['variable', 'parameter'].includes(resolved.kind) && next === '=') {
        const value = lineTokens[index + 2];
        const type = key(resolved.typeName ?? 'Variant');
        if (value && value.kind === 'string' && lineTokens.length === index + 3 && ['long', 'integer', 'double', 'single', 'byte', 'currency'].includes(type) && !Number.isFinite(Number(value.text.slice(1, -1)))) {
          diagnostics.push({ range: rangeOf(value), severity: DiagnosticSeverity.Error, code: 'type-mismatch', source: DIAGNOSTIC_SOURCE, message: `该字面量不能确定转换为 ${resolved.typeName}。` });
        }
      }
      if (resolved?.parameters && next === '(' && !resolved.external && token.line !== resolved.selectionRange.start.line) {
        const parenToken = lineTokens[index + 1]!;
        // 空格调用形式的括号（`Foo (1), 2`）只包裹首实参，后续逗号在括号外，无法可靠计数——按 LIM-01 跳过。
        const spaceInvoked = parenToken.start > token.end;
        if (!spaceInvoked) {
        const after = lineTokens.slice(index + 2);
        let depth = 0; let count = 0; let found = false; let nonempty = false;
        for (const part of after) {
          if (part.text === ')' && depth === 0) { found = true; break; }
          if (part.text === '(') { depth++; } else if (part.text === ')') { depth--; } else if (part.text === ',' && depth === 0) { count++; }
          nonempty = true;
        }
        if (found) {
          const supplied = nonempty ? count + 1 : 0;
          const required = resolved.parameters.filter(param => !param.optional && !param.paramArray).length;
          if (supplied < required || supplied > resolved.parameters.length && !resolved.parameters.some(param => param.paramArray)) { diagnostics.push({ range: rangeOf(token), severity: DiagnosticSeverity.Error, code: 'argument-count', source: DIAGNOSTIC_SOURCE, message: `${resolved.name} 参数数量不符，需要 ${required}–${resolved.parameters.length} 个。` }); }
        }
        }
      }
    }
    if ((module.kind === 'form' || module.kind === 'document') && this.context.typeIndexes.length === 0) {
      diagnostics.unshift({ range: { start: { line: 0, character: 0 }, end: { line: 0, character: 1 } }, severity: DiagnosticSeverity.Information, code: 'missing-host', source: DIAGNOSTIC_SOURCE, message: '宿主或控件类型信息缺失。源码级能力仍可用；请配置类型索引和补充声明。' });
    }
    return diagnostics.slice(0, 300);
  }

  semanticTokens(module: ParsedModule): { data: number[] } {
    const values: Array<{ token: VbaToken; type: number; modifiers: number }> = [];
    const typeKinds: Partial<Record<VbaSymbol['kind'], number>> = { module: 0, class: 1, type: 1, enum: 2, parameter: 3, variable: 4, constant: 4, field: 5, enumMember: 6, function: 7, sub: 8, property: 5, event: 8 };
    for (const token of module.tokens) {
      if (token.kind !== 'identifier') { continue; }
      const symbol = this.resolve(module, token);
      if (!symbol) { continue; }
      const declaration = symbol.uri === module.uri && symbol.selectionRange.start.line === token.line && symbol.selectionRange.start.character === token.start;
      values.push({ token, type: typeKinds[symbol.kind] ?? 4, modifiers: (declaration ? 1 : 0) | (symbol.kind === 'constant' ? 2 : 0) });
    }
    const data: number[] = []; let previousLine = 0; let previousStart = 0;
    for (const value of values.sort((a, b) => a.token.line - b.token.line || a.token.start - b.token.start)) {
      const { token } = value;
      const deltaLine = token.line - previousLine;
      data.push(deltaLine, deltaLine ? token.start : token.start - previousStart, token.end - token.start, value.type, value.modifiers);
      previousLine = token.line; previousStart = token.start;
    }
    return { data };
  }
}
