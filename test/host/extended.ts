import assert from 'node:assert/strict';
import { randomUUID } from 'node:crypto';
import * as vscode from 'vscode';
import type { LanguageClient } from 'vscode-languageclient/node';

async function until<T>(read: () => PromiseLike<T>, accept: (value: T) => boolean, message: string): Promise<T> {
  const deadline = Date.now() + 8000;
  let value = await read();
  while (!accept(value) && Date.now() < deadline) { await new Promise(resolve => setTimeout(resolve, 80)); value = await read(); }
  assert.ok(accept(value), message);
  return value;
}
interface ServerEdit { changes: Record<string, Array<{ range: { start: vscode.Position; end: vscode.Position }; newText: string }>>; }
interface Status { projects: Array<{ name: string; files: number; indexes: unknown[]; errors: string[] }>; invalidUtf8: string[]; }

export async function runExtended(api: { client: LanguageClient }, root: vscode.Uri): Promise<string[]> {
  const id = randomUUID().slice(0, 8);
  const name = `HostChecks${id}`;
  const directory = vscode.Uri.joinPath(root, `_host_checks_${id}`);
  const file = (name: string) => vscode.Uri.joinPath(directory, name);
  const write = async (name: string, content: string | Uint8Array) => vscode.workspace.fs.writeFile(file(name), typeof content === 'string' ? Buffer.from(content) : content);
  const projectConfig = { name, targetBits: 64, typeIndexes: ['types.json', 'missing.json'] };
  const status = () => api.client.sendRequest<Status>('vba/projectStatus');
  const results: string[] = [];
  try {
    await vscode.workspace.fs.createDirectory(directory);
    await write('types.json', JSON.stringify({ formatVersion: 1, library: { name: 'HostTest', guid: '{12345678-1234-1234-1234-123456789012}', major: 1, minor: 0, lcid: 0, bitness: 64 }, types: [{ name: 'StdFont', kind: 'class', members: [{ name: 'Name', kind: 'property', returnType: 'String' }, { name: 'Size', kind: 'property', returnType: 'Double' }] }] }));
    await write('Font.bas', 'Option Explicit\nPublic Sub FontCheck()\nDim font As StdFont\nfont.\nEnd Sub\n');
    await write('Bom.bas', '\ufeffoption explicit\r\npublic sub BomCheck()\r\ndim count as long\r\ncount = 1\r\nend sub\r\n');
    await write('RenameA.bas', 'Option Explicit\nPublic Function AtomicName() As Long\nAtomicName = 1\nEnd Function\n');
    await write('RenameB.bas', 'Option Explicit\nPublic Sub CallAtomic()\nDim answer As Long\nanswer = AtomicName()\nEnd Sub\n');
    await write('vba-project.json', JSON.stringify(projectConfig));
    await until(status, value => value.projects.some(project => project.name === name && project.files >= 4), '新增嵌套工程由文件监视器发现');

    const font = await vscode.workspace.openTextDocument(file('Font.bas'));
    await vscode.window.showTextDocument(font);
    await api.client.sendRequest('vba/projectStatus');
    const completion = await vscode.commands.executeCommand<vscode.CompletionList>('vscode.executeCompletionItemProvider', font.uri, new vscode.Position(3, 5));
    assert.ok(completion?.items.some(item => (typeof item.label === 'string' ? item.label : item.label.label) === 'Name'), '已加载的 COM 类型索引提供成员补全');
    const externalDoc = await vscode.workspace.openTextDocument(vscode.Uri.parse('vba-library:/HostTest/StdFont'));
    assert.ok(externalDoc.getText().includes('property Name As String'), '外部定义以只读虚拟声明文档显示');
    results.push('portableComCompletion');

    const bom = await vscode.workspace.openTextDocument(file('Bom.bas'));
    await vscode.window.showTextDocument(bom);
    await until(async () => bom.encoding, encoding => encoding === 'utf8bom', 'BOM 编码保持为 utf8bom');
    const edits = await vscode.commands.executeCommand<vscode.TextEdit[]>('vscode.executeFormatDocumentProvider', bom.uri, { tabSize: 4, insertSpaces: true });
    assert.ok(edits?.length);
    const edit = new vscode.WorkspaceEdit(); edit.set(bom.uri, edits);
    assert.ok(await vscode.workspace.applyEdit(edit));
    assert.ok(await bom.save());
    const saved = await vscode.workspace.fs.readFile(bom.uri);
    assert.deepEqual([...saved.slice(0, 3)], [0xef, 0xbb, 0xbf]);
    const decoded = Buffer.from(saved).toString('utf8');
    assert.ok(decoded.includes('    Dim count As Long'));
    assert.ok(!/(?<!\r)\n/.test(decoded), '保存保留 CRLF');
    results.push('utf8BomAndCrlfSave');

    // 缺失索引不得使已确定的源码局部变量也无法重命名。
    const local = await api.client.sendRequest<ServerEdit>('textDocument/rename', { textDocument: { uri: bom.uri.toString() }, position: { line: 3, character: 6 }, newName: 'total' });
    assert.equal(local.changes[Object.keys(local.changes)[0]!]!.length, 2);
    results.push('localRenameWithoutExternalIndex');

    const first = await vscode.workspace.openTextDocument(file('RenameA.bas'));
    const second = await vscode.workspace.openTextDocument(file('RenameB.bas'));
    await vscode.window.showTextDocument(first);
    await vscode.window.showTextDocument(second);
    await api.client.sendRequest('vba/projectStatus');
    const beforeFirst = first.getText(); const beforeSecond = second.getText();
    const rename = await api.client.sendRequest<ServerEdit>('textDocument/rename', { textDocument: { uri: first.uri.toString() }, position: { line: 1, character: 19 }, newName: 'AtomicRenamed' });
    const atomic = new vscode.WorkspaceEdit();
    for (const [uri, items] of Object.entries(rename.changes)) { atomic.set(vscode.Uri.parse(uri), items.map(item => new vscode.TextEdit(new vscode.Range(item.range.start.line, item.range.start.character, item.range.end.line, item.range.end.character), item.newText))); }
    assert.ok(await vscode.workspace.applyEdit(atomic));
    assert.ok(first.getText().includes('AtomicRenamed') && second.getText().includes('AtomicRenamed'));
    await vscode.commands.executeCommand('undo');
    assert.equal(first.getText(), beforeFirst, '一次撤销恢复声明文件');
    assert.equal(second.getText(), beforeSecond, '一次撤销恢复调用文件');
    results.push('crossFileRenameSingleUndo');

    await write('Bad.bas', Buffer.concat([Buffer.from('Option Explicit\nSub Bad()\n'), Buffer.from([0xd6, 0xd0, 0xce, 0xc4]), Buffer.from(' = 1\nEnd Sub\n')]));
    const bad = await vscode.workspace.openTextDocument(file('Bad.bas'));
    await vscode.window.showTextDocument(bad);
    const statusSnapshot = await status();
    await until(async () => vscode.languages.getDiagnostics(bad.uri).filter(item => item.source === 'VBAdminx'), items => items.some(item => item.code === 'encoding'), `坏编码诊断到达；工程状态：${JSON.stringify(statusSnapshot.projects.map(project => ({ name: project.name, files: project.files, errors: project.errors })))}；invalidUtf8=${statusSnapshot.invalidUtf8.length}；当前诊断：${JSON.stringify(vscode.languages.getDiagnostics(bad.uri).map(item => ({ source: item.source, code: item.code })))}`);
    assert.ok(vscode.languages.getDiagnostics(bad.uri).filter(item => item.source === 'VBAdminx').every(item => item.code === 'encoding' || item.code === 'project-types'));
    const badCompletion = await api.client.sendRequest<unknown[]>('textDocument/completion', { textDocument: { uri: bad.uri.toString() }, position: { line: 2, character: 1 } });
    assert.deepEqual(badCompletion, []);
    const badFormat = await api.client.sendRequest<unknown[]>('textDocument/formatting', { textDocument: { uri: bad.uri.toString() }, options: { tabSize: 4, insertSpaces: true } });
    assert.deepEqual(badFormat, []);
    results.push('invalidUtf8SuspendsSemanticEdits');

    const symbolName = `NewSymbol${id}`;
    await write('Added.bas', `Public Sub ${symbolName}()\nEnd Sub\n`);
    const query = () => vscode.commands.executeCommand<vscode.SymbolInformation[]>('vscode.executeWorkspaceSymbolProvider', symbolName);
    await until(query, items => !!items?.length, '新增文件刷新工作区符号');
    await vscode.workspace.fs.delete(file('Added.bas'));
    await until(query, items => !items?.length, '删除文件移除工作区符号');
    results.push('fileWatcherCreateDelete');

    await write('Range.bas', 'public sub RangeCheck()\ndim first as long\ndim second as long\nend sub\n');
    const range = await vscode.workspace.openTextDocument(file('Range.bas'));
    const rangeEditor = await vscode.window.showTextDocument(range);
    const partial = await api.client.sendRequest<Array<{ range: { start: vscode.Position; end: vscode.Position }; newText: string }>>('textDocument/rangeFormatting', { textDocument: { uri: range.uri.toString() }, range: { start: { line: 1, character: 0 }, end: { line: 2, character: 0 } }, options: { tabSize: 4, insertSpaces: true } });
    assert.ok(partial.length);
    await rangeEditor.edit(builder => { for (const item of partial) builder.replace(new vscode.Range(item.range.start.line, item.range.start.character, item.range.end.line, item.range.end.character), item.newText); });
    assert.equal(range.lineAt(1).text, '    Dim first As Long');
    assert.equal(range.lineAt(2).text, 'dim second as long', '末端字符为 0 的下一行不应改变');
    results.push('rangeFormattingExclusiveEnd');
    return results;
  } finally {
    for (const document of vscode.workspace.textDocuments.filter(document => document.uri.path.startsWith(directory.path))) {
      if (document.isDirty) {
        await vscode.window.showTextDocument(document);
        await vscode.commands.executeCommand('workbench.action.files.revert');
      }
    }
    await vscode.workspace.fs.delete(directory, { recursive: true, useTrash: false });
  }
}
