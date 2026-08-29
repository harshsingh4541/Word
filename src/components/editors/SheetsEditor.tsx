"use client";

import { useEffect, useRef } from "react";
import { createUniver, LocaleType, mergeLocales } from "@univerjs/presets";
import { UniverSheetsCorePreset } from "@univerjs/preset-sheets-core";
import UniverPresetSheetsCoreEnUS from "@univerjs/preset-sheets-core/locales/en-US";
import { UniverSheetsFindReplacePreset } from "@univerjs/preset-sheets-find-replace";
import UniverPresetSheetsFindReplaceEnUS from "@univerjs/preset-sheets-find-replace/locales/en-US";
import { UniverSheetsConditionalFormattingPreset } from "@univerjs/preset-sheets-conditional-formatting";
import UniverPresetSheetsConditionalFormattingEnUS from "@univerjs/preset-sheets-conditional-formatting/locales/en-US";
import { UniverSheetsDataValidationPreset } from "@univerjs/preset-sheets-data-validation";
import UniverPresetSheetsDataValidationEnUS from "@univerjs/preset-sheets-data-validation/locales/en-US";
import { UniverSheetsFilterPreset } from "@univerjs/preset-sheets-filter";
import UniverPresetSheetsFilterEnUS from "@univerjs/preset-sheets-filter/locales/en-US";
import { UniverSheetsSortPreset } from "@univerjs/preset-sheets-sort";
import UniverPresetSheetsSortEnUS from "@univerjs/preset-sheets-sort/locales/en-US";
import { ICommandService } from "@univerjs/core";
import type { Injector, IWorkbookData } from "@univerjs/core";
import { loadSnapshot, saveSnapshot } from "@/lib/univer/persistence";
import { buildWordLocale, WORD_THEME } from "@/lib/univer/word-theme";

import "@univerjs/preset-sheets-core/lib/index.css";
import "@univerjs/preset-sheets-find-replace/lib/index.css";
import "@univerjs/preset-sheets-conditional-formatting/lib/index.css";
import "@univerjs/preset-sheets-data-validation/lib/index.css";
import "@univerjs/preset-sheets-filter/lib/index.css";
import "@univerjs/preset-sheets-sort/lib/index.css";

// Office names the first ribbon tab "Home"; the rest of Univer's sheet tabs
// (Insert / Formulas / Data / View) already match Excel's.
const SHEETS_UI_LOCALE = {
  ui: {
    ribbon: {
      start: "Home",
      startDesc: "Fonts, number formats and cell styles.",
    },
  },
};

const STORAGE_KEY = "sheets-default";
const AUTOSAVE_DELAY_MS = 600;

export default function SheetsEditor() {
  const containerRef = useRef<HTMLDivElement>(null);
  const disposedRef = useRef(false);

  useEffect(() => {
    if (!containerRef.current || disposedRef.current) return;
    disposedRef.current = true;

    const { univer, univerAPI } = createUniver({
      // Same Office chrome as the docs editor: Word blue, and the tabbed
      // grid ribbon rather than the single toolbar row.
      theme: WORD_THEME,
      locale: LocaleType.EN_US,
      locales: {
        [LocaleType.EN_US]: buildWordLocale(
          mergeLocales(
            UniverPresetSheetsCoreEnUS,
            UniverPresetSheetsFindReplaceEnUS,
            UniverPresetSheetsConditionalFormattingEnUS,
            UniverPresetSheetsDataValidationEnUS,
            UniverPresetSheetsFilterEnUS,
            UniverPresetSheetsSortEnUS,
          ),
          SHEETS_UI_LOCALE,
        ),
      },
      presets: [
        UniverSheetsCorePreset({
          container: containerRef.current,
          ribbonType: "grid",
        }),
        UniverSheetsFindReplacePreset(),
        UniverSheetsConditionalFormattingPreset(),
        UniverSheetsDataValidationPreset(),
        UniverSheetsFilterPreset(),
        UniverSheetsSortPreset(),
      ],
    });

    const saved = loadSnapshot<Partial<IWorkbookData>>(STORAGE_KEY);
    const fWorkbook = univerAPI.createWorkbook(saved ?? {});

    const injector = univer.__getInjector() as Injector;
    const commandService = injector.get(ICommandService);

    let saveTimeout: ReturnType<typeof setTimeout> | undefined;
    const flushSave = () => saveSnapshot(STORAGE_KEY, fWorkbook.save());
    const commandSubscription = commandService.onCommandExecuted(() => {
      clearTimeout(saveTimeout);
      saveTimeout = setTimeout(flushSave, AUTOSAVE_DELAY_MS);
    });
    window.addEventListener("beforeunload", flushSave);

    return () => {
      commandSubscription.dispose();
      window.removeEventListener("beforeunload", flushSave);
      clearTimeout(saveTimeout);
      flushSave();

      // See DocsEditor.tsx's identical dispose() guard: a fast unmount
      // before Univer's async preset init reaches its "steady" lifecycle
      // stage makes an internal firstValueFrom(lifecycle$...) reject with
      // RxJS's EmptyError once disposal completes the source stream.
      // Harmless — swallow only that specific error.
      const swallowEmptyError = (event: PromiseRejectionEvent) => {
        if (event.reason?.name === "EmptyError") event.preventDefault();
      };
      window.addEventListener("unhandledrejection", swallowEmptyError);

      // Same race, different symptom: dispose() can synchronously unmount
      // an internal React root Univer owns (its own toolbar/canvas overlay)
      // while THIS component's own unmount is still mid-render for the same
      // commit. React reports that via console.error, not a thrown
      // exception, so the try/catch below can't see it — only a scoped
      // console.error filter can. Restored synchronously right after
      // dispose() returns, so no unrelated error in this window gets lost.
      const originalConsoleError = console.error;
      console.error = (...args: unknown[]) => {
        if (typeof args[0] === "string" && args[0].includes("synchronously unmount a root")) return;
        originalConsoleError(...args);
      };
      try {
        univer.dispose();
      } catch (err) {
        if ((err as Error)?.name !== "EmptyError") throw err;
      } finally {
        console.error = originalConsoleError;
        setTimeout(() => window.removeEventListener("unhandledrejection", swallowEmptyError), 0);
      }

      disposedRef.current = false;
    };
  }, []);

  return <div ref={containerRef} className="h-full w-full" />;
}
