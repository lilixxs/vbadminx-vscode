import { randomUUID } from 'node:crypto';
import * as vscode from 'vscode';
import { t } from './i18n';
import { parseFormDesign } from './formDesign';
import { buildRenderPayload, renderFormPage } from './formHtml';

/**
 * 窗体预览：在侧边打开 WebView，按 .frm 的设计段近似还原 UserForm。
 * - 只读，不执行 VBA，不读取 .frx 二进制资源。
 * - 文档内容变化后 250 ms 内热更新（重新解析并替换 stage 内容）。
 */
const MAX_BYTES = 2 * 1024 * 1024;
const DEBOUNCE_MS = 250;

function isFormDocument(document: vscode.TextDocument): boolean {
  return document.uri.scheme === 'file' && /\.frm$/i.test(document.uri.path);
}

function baseName(uri: vscode.Uri): string {
  return uri.path.split('/').pop() ?? uri.fsPath;
}

export class FormPreviewManager implements vscode.Disposable {
  private panel: vscode.WebviewPanel | undefined;
  private documentUri: string | undefined;
  private timer: ReturnType<typeof setTimeout> | undefined;
  private disposed = false;

  isShowing(uri: vscode.Uri): boolean {
    return this.panel !== undefined && this.documentUri === uri.toString();
  }

  async show(resource?: vscode.Uri): Promise<void> {
    const document = resource
      ? await vscode.workspace.openTextDocument(resource)
      : await this.activeFormDocument();
    if (!isFormDocument(document)) {
      throw new Error(t('form.openFailed', resource?.fsPath ?? '', t('form.noDesign')));
    }
    if (Buffer.byteLength(document.getText(), 'utf8') > MAX_BYTES) {
      void vscode.window.showWarningMessage(t('form.tooLarge'));
      return;
    }
    const title = t('form.title', baseName(document.uri));
    if (!this.panel) {
      this.panel = vscode.window.createWebviewPanel(
        'vbadminxFormPreview',
        title,
        { viewColumn: vscode.ViewColumn.Beside, preserveFocus: true },
        { enableScripts: true, retainContextWhenHidden: true, localResourceRoots: [] },
      );
      this.panel.onDidDispose(() => this.reset());
      this.panel.onDidChangeViewState(() => {
        if (this.panel?.visible) { this.updateNow(); }
      });
    } else {
      this.panel.title = title;
      this.panel.reveal(vscode.ViewColumn.Beside, true);
    }
    this.documentUri = document.uri.toString();
    this.render(document);
  }

  /** 语言或设置变化后重建页面（保留同一面板）。 */
  reload(): void {
    if (!this.documentUri || !this.panel) { return; }
    const document = vscode.workspace.textDocuments.find(item => item.uri.toString() === this.documentUri);
    if (document) { this.render(document); }
  }

  private render(document: vscode.TextDocument): void {
    if (!this.panel) { return; }
    const title = t('form.title', baseName(document.uri));
    const nonce = randomUUID().replaceAll('-', '');
    this.panel.title = title;
    this.panel.webview.html = renderFormPage(parseFormDesign(document.getText()), {
      nonce,
      title,
      hint: t('form.hint'),
      emptyText: t('form.noDesign'),
    });
  }

  scheduleUpdate(document: vscode.TextDocument): void {
    if (!this.panel || this.documentUri !== document.uri.toString()) { return; }
    if (this.timer) { clearTimeout(this.timer); }
    this.timer = setTimeout(() => { this.timer = undefined; this.updateNow(); }, DEBOUNCE_MS);
  }

  /** 关闭与指定文档绑定的预览面板（文档关闭时调用）。 */
  closeFor(document: vscode.TextDocument): void {
    if (this.documentUri === document.uri.toString()) { this.panel?.dispose(); }
  }

  updateNow(): void {
    const panel = this.panel;
    if (!panel || !this.documentUri) { return; }
    const document = vscode.workspace.textDocuments.find(item => item.uri.toString() === this.documentUri);
    if (!document) { return; }
    const model = parseFormDesign(document.getText());
    const title = t('form.title', baseName(document.uri));
    const payload = buildRenderPayload(model, title, t('form.noDesign'));
    void panel.webview.postMessage({ type: 'render', ...payload });
    panel.title = title;
  }

  private async activeFormDocument(): Promise<vscode.TextDocument> {
    const editor = vscode.window.activeTextEditor;
    if (editor && isFormDocument(editor.document)) { return editor.document; }
    throw new Error(t('form.openFailed', '', t('form.noDesign')));
  }

  private reset(): void {
    if (this.timer) { clearTimeout(this.timer); this.timer = undefined; }
    this.panel = undefined;
    this.documentUri = undefined;
  }

  dispose(): void {
    if (this.disposed) { return; }
    this.disposed = true;
    if (this.timer) { clearTimeout(this.timer); this.timer = undefined; }
    this.panel?.dispose();
    this.panel = undefined;
    this.documentUri = undefined;
  }
}

export function registerFormPreview(context: vscode.ExtensionContext): void {
  const manager = new FormPreviewManager();
  const report = (error: unknown): void => {
    if (error instanceof vscode.CancellationError) { return; }
    const message = error instanceof Error ? error.message : t('error.generic');
    void vscode.window.showErrorMessage(t('error.prefix', message));
  };
  context.subscriptions.push(
    manager,
    vscode.commands.registerCommand('vbadminx.previewForm', async (resource?: vscode.Uri) => {
      try { await manager.show(resource instanceof vscode.Uri ? resource : undefined); }
      catch (error) { report(error); }
    }),
    vscode.workspace.onDidChangeTextDocument(event => manager.scheduleUpdate(event.document)),
    vscode.workspace.onDidCloseTextDocument(document => manager.closeFor(document)),
    vscode.workspace.onDidChangeConfiguration(event => {
      if (event.affectsConfiguration('vbadminx.language')) { manager.reload(); }
    }),
  );
}
