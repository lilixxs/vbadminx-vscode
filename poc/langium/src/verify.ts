import { readFileSync, writeFileSync } from 'node:fs';
import { EmptyFileSystem, type LangiumCoreServices } from 'langium';
import { createDefaultCoreModule, createDefaultSharedCoreModule, inject } from 'langium';
import { parseHelper } from 'langium/test';
import { VbaMiniGeneratedModule, VbaMiniGeneratedSharedModule } from './generated/module.js';
import { isProcedure, isCallStatement, isFunctionCall, isAssignment, type Module, type Procedure } from './generated/ast.js';

function buildServices(): LangiumCoreServices {
  const shared = inject(createDefaultSharedCoreModule(EmptyFileSystem), VbaMiniGeneratedSharedModule);
  const services = inject(createDefaultCoreModule({ shared }), VbaMiniGeneratedModule);
  shared.ServiceRegistry.register(services);
  return services;
}

const services = buildServices();
const parse = parseHelper(services);

interface ScenarioResult {
  name: string;
  purpose: string;
  parserErrors: number;
  procedureCount: number;
  procedureNames: string[];
  expectedProcedures: string[];
  missingProcedures: string[];
  recovered: boolean;
  resolvedCalls: number;
  unresolvedCalls: number;
  firstErrors: string[];
}

function proceduresOf(root: Module): Procedure[] {
  return (root.members ?? []).filter(isProcedure);
}

let scenarioIndex = 0;

async function scenario(name: string, purpose: string, source: string, expectedProcedures: string[]): Promise<ScenarioResult> {
  scenarioIndex += 1;
  const doc = await parse(source, { documentUri: `file:///poc-${scenarioIndex}.bas` });
  const root = doc.parseResult.value as Module;
  const procedures = proceduresOf(root);
  const names = procedures.map(item => item.name);
  let resolvedCalls = 0;
  let unresolvedCalls = 0;
  const countRef = (ref: unknown): void => { if (ref) { resolvedCalls += 1; } else { unresolvedCalls += 1; } };
  // 递归遍历全部 AST 节点：调用节点可能被 BinaryExpression 等中间节点包裹。
  // 跳过 `$` 前缀的内部引用（$container / $cstNode / $document），避免循环。
  const seen = new Set<unknown>();
  const walk = (node: unknown): void => {
    if (!node || typeof node !== 'object' || seen.has(node)) { return; }
    seen.add(node);
    if (isCallStatement(node)) { countRef(node.callee?.ref); }
    else if (isFunctionCall(node)) { countRef(node.callee?.ref); }
    for (const key of Object.keys(node as Record<string, unknown>)) {
      if (key.startsWith('$')) { continue; }
      const value = (node as Record<string, unknown>)[key];
      if (Array.isArray(value)) { for (const item of value) { walk(item); } }
      else if (value && typeof value === 'object') { walk(value); }
    }
  };
  walk(root);
  const found = new Set(names);
  const missingProcedures = expectedProcedures.filter(item => !found.has(item));
  return {
    name,
    purpose,
    parserErrors: doc.parseResult.parserErrors.length,
    procedureCount: procedures.length,
    procedureNames: names,
    expectedProcedures,
    missingProcedures,
    recovered: missingProcedures.length === 0,
    resolvedCalls,
    unresolvedCalls,
    firstErrors: doc.parseResult.parserErrors.slice(0, 5).map(error => error.message),
  };
}

const computeTail = `Public Function Compute(a As Long) As Long
    Compute = a
End Function
`;

async function main(): Promise<void> {
  const results: ScenarioResult[] = [];

  // A 基线：全部过程完整，语法正确。
  results.push(await scenario(
    'A-correct-baseline',
    '基线：三个完整过程，验证解析与交叉引用可用',
    `Public Sub First()
    Dim x As Long
    x = 1
End Sub

Public Function Second() As Long
    Second = Compute(1)
End Function

${computeTail}`,
    ['First', 'Second', 'Compute'],
  ));

  // B 容错核心：First 缺 End Sub，其后两个过程完整。
  results.push(await scenario(
    'B-missing-End-Sub',
    '容错核心：首过程缺 End Sub，检查后续过程是否仍进入 AST',
    `Public Sub First()
    Dim x As Long
    x = 1

Public Function Second() As Long
    Second = Compute(1)
End Function

${computeTail}`,
    ['Second', 'Compute'],
  ));

  // C 语句级错误：First 内 If 缺 Then，其后过程完整。
  results.push(await scenario(
    'C-statement-error',
    '语句级错误：过程内 If 缺 Then，检查后续过程是否仍进入 AST',
    `Public Sub First()
    If x > 1
        x = 2
    End If
End Sub

Public Function Second() As Long
    Second = Compute(1)
End Function

${computeTail}`,
    ['Second', 'Compute'],
  ));

  // D 真实样例：工程中确实缺 End Sub 的文件（只读，不修改）。
  const sample = 'D:/SVN_CQ24/trunk/doc/process/BAS(节能)/02-生成程序&生成结果备份/02-生成程序/X89-辅助小工具/vba_sync/X89-辅助小工具-V1_4.xlsm/模块/CQ24_AdvancedSort.bas';
  let realSample: unknown = { skipped: '样例不可读' };
  try {
    const text = readFileSync(sample, 'utf8');
    const doc = await parse(text, { documentUri: 'file:///CQ24_AdvancedSort.bas' });
    const root = doc.parseResult.value as Module;
    const procedures = proceduresOf(root);
    realSample = {
      note: '仅供参考：mini 语法仅覆盖过程/块/表达式，未覆盖 Type 之外的 Enum/Declare/On Error/GoTo/行标签/Select Case/Exit/条件编译等；错误数与过程数不代表 Langium 能力',
      file: 'CQ24_AdvancedSort.bas',
      parserErrors: doc.parseResult.parserErrors.length,
      procedureCount: procedures.length,
      procedureNames: procedures.slice(0, 10).map(item => item.name),
      // 只输出错误消息与行号，不输出业务源码内容。
      firstErrors: doc.parseResult.parserErrors.slice(0, 6).map(error => ({
        message: error.message,
        line: (error as { token?: { startLine?: number } }).token?.startLine,
        offset: (error as { token?: { startOffset?: number } }).token?.startOffset,
      })),
    };
  } catch (error) {
    realSample = { error: error instanceof Error ? error.message : '读取或解析失败' };
  }

  const report = { measured: true, langium: '4.4.0', generated: 'langium-cli 4.4.0', scenarios: results, realSample };
  writeFileSync('poc/langium/poc-result.json', JSON.stringify(report, null, 2) + '\n');
  console.log(JSON.stringify(report, null, 2));
}

main().catch(error => { console.error(error); process.exitCode = 1; });
