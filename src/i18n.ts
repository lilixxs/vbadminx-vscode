import * as vscode from 'vscode';
import {
  LOCALES,
  LOCALE_LABELS,
  normalizeLanguageTag,
  translate,
  type Locale,
  type MessageKey,
} from './i18n/catalog';

/**
 * 界面语言解析。
 *
 * 设置项 `vbadminx.language`：
 * - `auto`：跟随 VS Code 的显示语言；VS Code 语言不在支持列表时使用英文。
 * - `en` / `zh-cn` / `zh-tw` / `ja` / `ko`：强制使用该语言。
 */
export const LANGUAGE_SECTION = 'vbadminx';
export const LANGUAGE_KEY = 'language';

let active: Locale = 'en';

export function resolveLocale(
  setting: string | undefined = vscode.workspace.getConfiguration(LANGUAGE_SECTION).get<string>(LANGUAGE_KEY, 'auto'),
  vscodeLanguage: string = vscode.env.language,
): Locale {
  if (setting && setting !== 'auto' && (LOCALES as readonly string[]).includes(setting)) {
    return setting as Locale;
  }
  return normalizeLanguageTag(vscodeLanguage);
}

/** 在扩展激活时调用，确定当前界面语言。 */
export function activateLocale(): Locale {
  active = resolveLocale();
  return active;
}

export function currentLocale(): Locale {
  return active;
}

export function localeLabel(locale: Locale = active): string {
  return LOCALE_LABELS[locale];
}

/** 取当前语言文案；参数按 `{0}`、`{1}` 占位替换。 */
export function t(key: MessageKey, ...args: Array<string | number>): string {
  return translate(active, key, ...args);
}

/**
 * 监听语言设置变化：更新当前语言，切换时提示一次，并调用回调（刷新视图与预览）。
 */
export function watchLanguage(context: vscode.ExtensionContext, onChange?: () => void): void {
  context.subscriptions.push(vscode.workspace.onDidChangeConfiguration(event => {
    if (!event.affectsConfiguration(`${LANGUAGE_SECTION}.${LANGUAGE_KEY}`)) { return; }
    const previous = active;
    active = resolveLocale();
    if (active !== previous) {
      void vscode.window.showInformationMessage(t('language.changed', localeLabel(active)));
      onChange?.();
    }
  }));
}
