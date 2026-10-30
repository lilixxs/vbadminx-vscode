import * as vscode from 'vscode';

/** 不修改源文件；检查磁盘字节，拒绝让 LSP 把乱码当源码分析。 */
export function registerEncodingChecks(context: vscode.ExtensionContext): void {
  const checked = new Set<string>();
  const check = async (document: vscode.TextDocument): Promise<void> => {
    if (document.uri.scheme !== 'file' || !/\.(bas|cls|frm)$/i.test(document.uri.path) || checked.has(document.uri.toString())) { return; }
    checked.add(document.uri.toString());
    let bytes: Uint8Array;
    try { bytes = await vscode.workspace.fs.readFile(document.uri); } catch { return; }
    try { new TextDecoder('utf-8', { fatal: true }).decode(bytes); }
    catch {
      void vscode.window.showWarningMessage(`${vscode.workspace.asRelativePath(document.uri)} 不是有效 UTF-8。语义分析已暂停。请在外部工具转换副本后重新打开，避免直接保存导致转码。`);
      return;
    }
    const encoding = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf ? 'utf8bom' : 'utf8';
    if (!document.isDirty && document.encoding !== encoding) {
      try { await vscode.workspace.openTextDocument(document.uri, { encoding }); } catch {}
    }
  };
  context.subscriptions.push(vscode.workspace.onDidOpenTextDocument(document => { void check(document); }));
  context.subscriptions.push(vscode.workspace.onDidCloseTextDocument(document => checked.delete(document.uri.toString())));
  for (const document of vscode.workspace.textDocuments) { void check(document); }
}
