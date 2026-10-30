import { writeFileSync } from 'node:fs';
import { EmptyFileSystem } from 'langium';
import { createDefaultCoreModule, createDefaultSharedCoreModule, inject, URI } from 'langium';
import { parseHelper } from 'langium/test';
import { VbaMiniGeneratedModule, VbaMiniGeneratedSharedModule } from './generated/module.js';
import { isProcedure, type Module } from './generated/ast.js';

type Constant = number | string | boolean;

interface Branch { active: boolean; taken: boolean; }

/**
 * 条件编译预处理：清空非生效分支的行内容，但**保留行数与行序**。
 * 这样生效分支内的行列位置与原始文件一一对应，诊断位置无需额外映射。
 */
function preprocess(text: string, environment: Record<string, Constant>): {
  processed: string;
  constants: Record<string, Constant>;
  activeLineCount: number;
  inactiveLineCount: number;
  unknownConstant: boolean;
} {
  const constants: Record<string, Constant> = { ...environment };
  const lines = text.split(/\r\n|\n|\r/);
  const stack: Branch[] = [];
  let activeLineCount = 0;
  let inactiveLineCount = 0;
  let unknownConstant = false;

  const evaluate = (expression: string): boolean | undefined => {
    const match = /^(\w+)\s*(=|<>)\s*(\S+)$/.exec(expression.trim());
    if (match) {
      const [, name, operator, rawValue] = match;
      const value = constants[name!];
      if (value === undefined) { unknownConstant = true; return undefined; }
      const literal = /^-?\d+$/.test(rawValue!) ? Number(rawValue) : rawValue!.replace(/^"|"$/g, '');
      return operator === '=' ? value === literal : value !== literal;
    }
    const sole = /^(\w+)$/.exec(expression.trim());
    if (sole) {
      const value = constants[sole[1]!];
      if (value === undefined) { unknownConstant = true; return undefined; }
      return Boolean(value);
    }
    unknownConstant = true;
    return undefined;
  };

  const output = lines.map(line => {
    const trimmed = line.trim();
    const constMatch = /^#Const\s+(\w+)\s*=\s*(.+)$/i.exec(trimmed);
    if (constMatch) {
      const raw = constMatch[2]!.trim();
      constants[constMatch[1]!] = /^-?\d+$/.test(raw) ? Number(raw) : raw.replace(/^"|"$/g, '');
      return '';
    }
    const ifMatch = /^#If\s+(.+?)\s+Then$/i.exec(trimmed) ?? /^#If\s+(.+)$/i.exec(trimmed);
    if (ifMatch) {
      const result = evaluate(ifMatch[1]!);
      const active = result === true;
      stack.push({ active, taken: active });
      return '';
    }
    const elifMatch = /^#ElseIf\s+(.+?)\s+Then$/i.exec(trimmed) ?? /^#ElseIf\s+(.+)$/i.exec(trimmed);
    if (elifMatch) {
      const current = stack[stack.length - 1];
      if (!current) { return ''; }
      if (current.taken) { current.active = false; }
      else {
        const result = evaluate(elifMatch[1]!);
        current.active = result === true;
        current.taken = current.active;
      }
      return '';
    }
    if (/^#Else$/i.test(trimmed)) {
      const current = stack[stack.length - 1];
      if (current) { if (current.taken) { current.active = false; } else { current.active = true; current.taken = true; } }
      return '';
    }
    if (/^#End\s*If$/i.test(trimmed)) { stack.pop(); return ''; }

    const active = stack.every(item => item.active);
    if (active) { activeLineCount += 1; return line; }
    inactiveLineCount += 1;
    return '';
  });

  return { processed: output.join('\n'), constants, activeLineCount, inactiveLineCount, unknownConstant };
}

function buildServices() {
  const shared = inject(createDefaultSharedCoreModule(EmptyFileSystem), VbaMiniGeneratedSharedModule);
  const services = inject(createDefaultCoreModule({ shared }), VbaMiniGeneratedModule);
  shared.ServiceRegistry.register(services);
  return services;
}

async function main(): Promise<void> {
  const services = buildServices();
  const parse = parseHelper(services);

  // 行号（0 基）：
  // 7-8 位于 #Else 非生效分支内的“错误”应被排除；第 11 行生效区的语法错误应报出且行号不变。
  const source = [
    "' header comment",
    '#Const DEBUG = 1',
    'Public Sub Test()',
    '#If DEBUG = 1 Then',
    '    Dim x As Long',
    '    x = 1',
    '#Else',
    '    Dim y As Long',
    '    y = This is broken',
    '#End If',
    '    result = result +',
    'End Sub',
  ].join('\n');

  const pre = preprocess(source, {});
  const doc = await parse(pre.processed, { documentUri: URI.parse('file:///cond-1.bas') });
  const root = doc.parseResult.value as Module;
  const procedures = (root.members ?? []).filter(isProcedure);
  const errorLines = doc.parseResult.parserErrors.map(error => (error as { token?: { startLine?: number } }).token?.startLine);

  // 原始文本中 `result = result +` 位于 0 基第 10 行；Chevrotain 的 token.startLine 为 1 基，故期望 11。
  const expectedZeroBasedLine = 10;
  const expectedOneBasedLine = expectedZeroBasedLine + 1;
  const originalLines = source.split('\n');
  const processedLines = pre.processed.split('\n');
  const report = {
    measured: true,
    langium: '4.4.0',
    technique: '清空非生效分支行内容但保留行数，令生效分支行列位置与原始文件一致',
    constants: pre.constants,
    activeLineCount: pre.activeLineCount,
    inactiveLineCount: pre.inactiveLineCount,
    unknownConstant: pre.unknownConstant,
    procedureCount: procedures.length,
    procedureNames: procedures.map(item => item.name),
    parserErrors: doc.parseResult.parserErrors.length,
    errorLines,
    errorMessages: doc.parseResult.parserErrors.map(error => error.message),
    lineBase: 'Chevrotain token.startLine 为 1 基；LSP 行号为 0 基',
    inactiveBranchErrorReported: errorLines.includes(9), // 非生效分支行（0 基 8）对应 1 基 9
    activeErrorLineMatchesOriginal: errorLines.includes(expectedOneBasedLine),
    originalLineAtExpected: originalLines[expectedZeroBasedLine],
    processedLineAtExpected: processedLines[expectedZeroBasedLine],
    lineCountPreserved: originalLines.length === processedLines.length,
  };
  writeFileSync('poc/langium/cond-result.json', JSON.stringify(report, null, 2) + '\n');
  console.log(JSON.stringify(report, null, 2));
}

main().catch(error => { console.error(error); process.exitCode = 1; });
