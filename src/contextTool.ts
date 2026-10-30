import * as vscode from 'vscode';
import { getLanguageClient } from './languageClient';

export interface ContextInput {
  mode: 'selection' | 'diagnostics' | 'symbols' | 'references';
  query?: string;
}

export function readContext(input: ContextInput, token: vscode.CancellationToken): string {
  if (token.isCancellationRequested) { throw new vscode.CancellationError(); }
  if (!vscode.workspace.isTrusted) { throw new Error('先信任工作区，再读取代码上下文。'); }
  if (input.mode !== 'selection' && input.mode !== 'diagnostics') {
    throw new Error('mode 必须为 selection 或 diagnostics。');
  }
  const editor = vscode.window.activeTextEditor;
  if (!editor || !/\.(bas|cls|frm)$/i.test(editor.document.uri.path)) {
    throw new Error('先打开工作区内的 .bas、.cls 或 .frm 文件。');
  }
  if (!vscode.workspace.getWorkspaceFolder(editor.document.uri)) {
    throw new Error('活动文件不属于当前工作区。');
  }
  const common = {
    semanticAnalysis: true,
    file: vscode.workspace.asRelativePath(editor.document.uri, true),
    documentVersion: editor.document.version,
  };
  if (input.mode === 'selection') {
    if (editor.selection.isEmpty) { throw new Error('先选择要提供给 AI 的代码。'); }
    const selected = editor.document.getText(editor.selection);
    const limit = 12_000;
    return JSON.stringify({
      ...common,
      range: { startLine: editor.selection.start.line + 1, endLine: editor.selection.end.line + 1 },
      code: selected.slice(0, limit),
      truncated: selected.length > limit,
    });
  }
  const all = vscode.languages.getDiagnostics(editor.document.uri);
  return JSON.stringify({
    ...common,
    note: '这些诊断来自编辑器中的诊断提供器，包括 VBAdminx。空列表不证明代码完全正确；未知类型可能限制分析。',
    diagnostics: all.slice(0, 100).map(item => ({
      line: item.range.start.line + 1,
      character: item.range.start.character + 1,
      severity: item.severity,
      message: item.message.slice(0, 1000),
      source: item.source,
    })),
    truncated: all.length > 100,
  });
}

export class VbaContextTool implements vscode.LanguageModelTool<ContextInput> {
  prepareInvocation(options: vscode.LanguageModelToolInvocationPrepareOptions<ContextInput>): vscode.PreparedToolInvocation {
    return {
      invocationMessage: '读取 VBA 上下文',
      confirmationMessages: {
        title: '读取 VBA 上下文',
        message: options.input.mode === 'selection'
          ? '允许聊天工具读取当前 VBA 文件的选中代码？结果将交给当前聊天模型。'
          : `允许聊天工具读取当前 VBA 文件的 ${options.input.mode} 上下文？结果将交给当前聊天模型。`,
      },
    };
  }

  async invoke(options: vscode.LanguageModelToolInvocationOptions<ContextInput>, token: vscode.CancellationToken): Promise<vscode.LanguageModelToolResult> {
    if (!['selection', 'diagnostics', 'symbols', 'references'].includes(options.input.mode)) { throw new Error('不支持的上下文模式。'); }
    if (options.input.mode === 'selection' || options.input.mode === 'diagnostics') {
      return new vscode.LanguageModelToolResult([new vscode.LanguageModelTextPart(readContext(options.input, token))]);
    }
    // 复用相同的工作区、信任与活动文件检查。
    readContext({ mode: 'diagnostics' }, token);
    const editor = vscode.window.activeTextEditor!;
    const client = getLanguageClient();
    if (!client) { throw new Error('VBA 语言服务器尚未启动。'); }
    const result = await client.sendRequest('vba/context', { uri: editor.document.uri.toString(), mode: options.input.mode, query: options.input.query, position: editor.selection.active }, token);
    const text = JSON.stringify({ semanticAnalysis: true, mode: options.input.mode, result });
    return new vscode.LanguageModelToolResult([new vscode.LanguageModelTextPart(text.length > 24000 ? JSON.stringify({ truncated: true, message: '结果过多，请缩小查询。' }) : text)]);
  }
}
