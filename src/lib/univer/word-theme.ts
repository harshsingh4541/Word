import { defaultTheme } from "@univerjs/presets";
import type { ILocales } from "@univerjs/core";

// Word's accent blue, expanded into the 50-900 ramp Univer's design tokens
// expect. 600 is the one that matters most: it drives the active ribbon
// tab, focus rings, selection chrome and every primary button.
export const WORD_THEME = {
  ...defaultTheme,
  primary: {
    50: "#EFF3FB",
    100: "#DCE6F6",
    200: "#B9CDEC",
    300: "#8FB0E0",
    400: "#5C8ED2",
    500: "#2F6FC4",
    600: "#185ABD",
    700: "#14479A",
    800: "#103A7D",
    900: "#0C2C5E",
  },
};

type LocaleTree = { [key: string]: string | LocaleTree };

/**
 * `mergeLocales` is a shallow Object.assign over top-level namespaces, so
 * overriding a single key (say `ui.ribbon.start`) with it would drop every
 * other string in that namespace. This merges key by key instead.
 */
export function deepMergeLocale<T extends LocaleTree>(base: T, override: LocaleTree): T {
  const result: LocaleTree = { ...base };
  for (const [key, value] of Object.entries(override)) {
    const existing = result[key];
    if (existing && typeof existing === "object" && value && typeof value === "object") {
      result[key] = deepMergeLocale(existing, value);
    } else {
      result[key] = value;
    }
  }
  return result as T;
}

export function buildWordLocale(base: ILocales[keyof ILocales], override: LocaleTree): ILocales[keyof ILocales] {
  return deepMergeLocale(base as LocaleTree, override) as ILocales[keyof ILocales];
}
