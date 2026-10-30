import test from 'node:test';
import assert from 'node:assert/strict';
import { parseModule } from '../src/language/parser';
import { VbaService } from '../src/language/service';
import { canonicalUri } from '../src/language/project';

test('命名实参不能当作调用方局部变量参与重命名', () => {
  const module = parseModule('file:///Named.bas', 'Option Explicit\nPublic Sub Run()\nDim count As Long\ncount = 1\nCall Other(count:=2)\nEnd Sub\nPublic Sub Other(ByVal count As Long)\nEnd Sub');
  const service = new VbaService({ modules: [module], typeIndexes: [] });
  const changes = service.rename(module, { line: 3, character: 2 }, 'total');
  assert.equal(changes[module.uri]?.length, 2);
  assert.ok(!changes[module.uri]?.some(edit => edit.range.start.line === 4));
  assert.ok(!service.diagnostics(module).some(diagnostic => diagnostic.range.start.line === 4 && diagnostic.code === 'undeclared'));
});

test('属性 Get Let 访问器作为一个名称重命名', () => {
  const module = parseModule('file:///Box.cls', 'Private value As Long\nPublic Property Get Size() As Long\nSize = value\nEnd Property\nPublic Property Let Size(ByVal input As Long)\nvalue = input\nEnd Property');
  const service = new VbaService({ modules: [module], typeIndexes: [] });
  const changes = service.rename(module, { line: 1, character: 22 }, 'Length');
  const lines = changes[module.uri]?.map(edit => edit.range.start.line);
  assert.deepEqual(lines?.sort((a, b) => a - b), [1, 2, 4]);
});

test('Windows 文件 URI 盘符与编码差异被归一化', () => {
  if (process.platform !== 'win32') { return; }
  assert.equal(canonicalUri('file:///d%3A/Folder/Test.bas'), canonicalUri('file:///D:/Folder/Test.bas'));
});

test('缺少外部库时外部常量和 VBA 内建函数不报确定的未声明错误', () => {
  const module = parseModule('file:///External.bas', 'Option Explicit\nPublic Sub Main()\nDim conn As ADODB.Connection\nMsgBox IIf(True, "a", "b"), vbCritical\nconn.State = adStateOpen\nEnd Sub');
  const diagnostics = new VbaService({ modules: [module], typeIndexes: [] }).diagnostics(module);
  assert.ok(diagnostics.some(diagnostic => diagnostic.code === 'missing-type'));
  assert.ok(diagnostics.some(diagnostic => diagnostic.code === 'missing-external-symbol'));
  assert.ok(!diagnostics.some(diagnostic => diagnostic.code === 'undeclared'));
});

test('重命名保留 VBA 类型后缀', () => {
  const module = parseModule('file:///Suffix.bas', 'Option Explicit\nSub Main()\nDim label$\nlabel$ = "x"\nDim counter^\ncounter^ = 1\nEnd Sub');
  const service = new VbaService({ modules: [module], typeIndexes: [] });
  const strings = service.rename(module, { line: 3, character: 2 }, 'caption');
  assert.ok(strings[module.uri]?.every(edit => edit.newText === 'caption$'));
  const integers = service.rename(module, { line: 5, character: 2 }, 'total');
  assert.ok(integers[module.uri]?.every(edit => edit.newText === 'total^'));
  assert.ok(!service.diagnostics(module).some(diagnostic => diagnostic.code === 'undeclared'));
});

test('同名 With 属主按当前过程的局部类型解析', () => {
  const first = parseModule('file:///First.cls', 'Public Sub One()\nEnd Sub');
  const second = parseModule('file:///Second.cls', 'Public Sub Two()\nEnd Sub');
  const module = parseModule('file:///Scopes.bas', 'Sub A()\nDim obj As First\nWith obj\n.One\nEnd With\nEnd Sub\nSub B()\nDim obj As Second\nWith obj\n.Two\nEnd With\nEnd Sub');
  const service = new VbaService({ modules: [module, first, second], typeIndexes: [] });
  const completions = service.completions(module, { line: 9, character: 1 });
  assert.ok(completions.some(item => item.label === 'Two'));
  assert.ok(!completions.some(item => item.label === 'One'));
  assert.equal(service.definition(module, { line: 9, character: 2 })[0]?.uri, second.uri);
});

test('链式属性、方法返回值与 With 使用逐级类型补全', () => {
  const child = parseModule('file:///Child.cls', 'Public Property Get Name() As String\nEnd Property');
  const parent = parseModule('file:///Parent.cls', 'Public Property Get Child() As Child\nEnd Property\nPublic Function GetChild() As Child\nEnd Function');
  const source = parseModule('file:///Chain.bas', 'Sub Main()\nDim obj As Parent\nobj.Child.\nobj.GetChild().\nWith obj.Child\n.Name\nEnd With\nEnd Sub');
  const service = new VbaService({ modules: [source, parent, child], typeIndexes: [] });
  assert.ok(service.completions(source, { line: 2, character: 10 }).some(item => item.label === 'Name'));
  assert.ok(service.completions(source, { line: 3, character: 15 }).some(item => item.label === 'Name'));
  assert.equal(service.definition(source, { line: 5, character: 3 })[0]?.uri, child.uri);
});

test('同名 COM 类型按库名限定，未知属主不能猜测到全局同名类', () => {
  const source = parseModule('file:///Qualified.bas', 'Sub Main()\nDim obj As First.Application\nobj.\nApplication.\nDim ambiguous As Application\nambiguous.\nEnd Sub');
  const library = (name: string, member: string): import('../src/language/model').PortableTypeIndex => ({ formatVersion: 1, library: { name, guid: name, major: 1, minor: 0, lcid: 0, bitness: 64 }, types: [{ name: 'Application', kind: 'class', members: [{ name: member, kind: 'property', returnType: 'String' }] }] });
  const service = new VbaService({ modules: [source], typeIndexes: [library('First', 'FirstMember'), library('Second', 'SecondMember')] });
  assert.deepEqual(service.completions(source, { line: 2, character: 4 }).map(item => item.label), ['FirstMember']);
  assert.deepEqual(service.completions(source, { line: 3, character: 12 }), []);
  assert.deepEqual(service.completions(source, { line: 5, character: 10 }), []);
});

test('参数重命名包含命名实参，调用方同名变量保持不变', () => {
  const module = parseModule('file:///Arguments.bas', 'Public Function Compute(ByVal count As Long) As Long\nCompute = count\nEnd Function\nSub Main()\nDim count As Long\ncount = Compute(count:=2)\nEnd Sub');
  const service = new VbaService({ modules: [module], typeIndexes: [] });
  const changes = service.rename(module, { line: 0, character: 33 }, 'quantity')[module.uri]!;
  assert.equal(changes.length, 3);
  assert.equal(changes.filter(edit => edit.range.start.line === 5).length, 1);
  assert.equal(changes.find(edit => edit.range.start.line === 5)?.range.start.character, 16);
});

test('VBA 数字赋给 Boolean 可隐式转换，不能误报类型不匹配', () => {
  const module = parseModule('file:///Boolean.bas', 'Option Explicit\nSub Main()\nDim enabled As Boolean\nenabled = 1\nEnd Sub');
  assert.ok(!new VbaService({ modules: [module], typeIndexes: [] }).diagnostics(module).some(diagnostic => diagnostic.code === 'type-mismatch'));
});

test('名称大小写修复保留类型后缀，相同名称不产生多余提示', () => {
  const module = parseModule('file:///Case.bas', 'Sub Main()\nDim Caption$\nCaption$ = "a"\ncaption$ = "b"\nEnd Sub');
  const cases = new VbaService({ modules: [module], typeIndexes: [] }).diagnostics(module).filter(diagnostic => diagnostic.code === 'name-case');
  assert.equal(cases.length, 1);
  assert.equal(cases[0]!.data.name, 'Caption$');
});

test('公共符号重命名不能被调用方的局部变量捕获', () => {
  const exported = parseModule('file:///Exported.bas', 'Public Function Compute() As Long\nCompute = 1\nEnd Function');
  const caller = parseModule('file:///Caller.bas', 'Sub Main()\nDim result As Long\nresult = Compute()\nEnd Sub');
  const service = new VbaService({ modules: [exported, caller], typeIndexes: [] });
  assert.throws(() => service.rename(exported, { line: 0, character: 20 }, 'result'), /遮蔽/);
});
