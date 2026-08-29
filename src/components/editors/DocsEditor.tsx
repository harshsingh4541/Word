"use client";

import { useEffect, useImperativeHandle, useRef, useState } from "react";
import { createUniver, LocaleType, mergeLocales } from "@univerjs/presets";
import { UniverDocsCorePreset } from "@univerjs/preset-docs-core";
import UniverPresetDocsCoreEnUS from "@univerjs/preset-docs-core/locales/en-US";
import { UniverDocsDrawingPreset } from "@univerjs/preset-docs-drawing";
import UniverPresetDocsDrawingEnUS from "@univerjs/preset-docs-drawing/locales/en-US";
import { UniverDocsHyperLinkPreset } from "@univerjs/preset-docs-hyper-link";
import UniverPresetDocsHyperLinkEnUS from "@univerjs/preset-docs-hyper-link/locales/en-US";
import { UniverDocsThreadCommentPreset } from "@univerjs/preset-docs-thread-comment";
import UniverPresetDocsThreadCommentEnUS from "@univerjs/preset-docs-thread-comment/locales/en-US";
import { UniverDocsFindReplacePlugin } from "@univerjs/docs-find-replace";
import { DocumentFlavor, ICommandService, UniverInstanceType, validateDocumentStructure } from "@univerjs/core";
import type { DocumentDataModel, IDocumentData, Injector, Nullable } from "@univerjs/core";
import { IUniverInstanceService } from "@univerjs/core";
import { DocSelectionManagerService, DocSkeletonManagerService, SetTextSelectionsOperation } from "@univerjs/docs";
import { IRenderManagerService } from "@univerjs/engine-render";
import {
  ALL_TABLE_STYLE_COMMANDS,
  clearRememberedTableRange,
  resolveLiveTableRange,
} from "@/lib/univer/table-style-commands";
import { SetBorderPenCommand } from "@/lib/univer/border-pen";
import { loadSnapshot, saveSnapshot, clearSnapshot } from "@/lib/univer/persistence";
import {
  createWordCommands,
  SetIndentCommandId,
  SetPageMarginsCommandId,
  SetZoomCommandId,
} from "@/lib/univer/word-commands";
import WordRuler, { type RulerGeometry } from "./WordRuler";
import WordVerticalRuler from "./WordVerticalRuler";
import { BuiltInUIPart, IUIPartsService } from "@univerjs/ui";
import { installWordRibbon, RELOCATED_UNIVER_MENU_ITEMS, WORD_UI_LOCALE } from "@/lib/univer/word-ribbon";
import { createTableResizeInteraction } from "@/lib/univer/table-resize";
import { hidePageMarginMarks } from "@/lib/univer/page-chrome";
import { disableSlashMenu } from "@/lib/univer/slash-key";
import { restoreFocusAfterDialogs } from "@/lib/univer/editor-focus";
import { keepCaretOnRibbonPointerDown } from "@/lib/univer/ribbon-focus";
import { persistInputStyleAcrossMutation } from "@/lib/univer/persistent-input-style";
import { createWordFeatureCommands } from "@/lib/univer/word-features";
import { createSpellCheckCommand, createSpellChecker } from "@/lib/univer/spell-check";
import { createTrackChanges, createTrackChangesCommands } from "@/lib/univer/track-changes";
import { createWatermarkCommand } from "@/lib/univer/watermark";
import { buildWordLocale, WORD_THEME } from "@/lib/univer/word-theme";

const STORAGE_KEY = "docs-default";
const AUTOSAVE_DELAY_MS = 600;
const DEFAULT_DOCUMENT_NAME = "Untitled document";
const STATUS_REFRESH_DELAY_MS = 400;
// A4 at 96 DPI. Traditional flavor is what unlocks Word-compatible real
// pagination (page breaks, ruler-visible page bounds) and header/footer
// editing — both crash on creation-time documentStyle in Univer 0.25.x but
// work cleanly as of 1.0.0-beta.2.
const DEFAULT_DOCUMENT_STYLE = {
  pageSize: { width: 794, height: 1123 },
  documentFlavor: DocumentFlavor.TRADITIONAL,
};

import "@univerjs/preset-docs-core/lib/index.css";
import "@univerjs/preset-docs-drawing/lib/index.css";
import "@univerjs/preset-docs-hyper-link/lib/index.css";
import "@univerjs/preset-docs-thread-comment/lib/index.css";

/** What the Word-style title bar and status bar display. */
export type WordDocumentStatus = {
  name: string;
  wordCount: number;
  pageCount: number;
  currentPage: number;
  zoom: number;
};

/**
 * What the surrounding Word chrome can do to the document. Everything else
 * — formatting, layout, export — is a ribbon command inside Univer.
 */
export type DocsEditorHandle = {
  setName: (name: string) => void;
  setZoom: (zoom: number) => void;
  /** Live page geometry for the ruler, or null before the doc renders. */
  getRulerGeometry: () => RulerGeometry | null;
  setIndents: (indents: { indentStart?: number; indentEnd?: number; indentFirstLine?: number }) => void;
  setMargins: (margins: { marginLeft?: number; marginRight?: number }) => void;
};

export default function DocsEditor({
  apiRef,
  onStatusChange,
}: {
  apiRef?: React.RefObject<DocsEditorHandle | null>;
  onStatusChange?: (status: WordDocumentStatus) => void;
}) {
  const containerRef = useRef<HTMLDivElement>(null);
  const disposedRef = useRef(false);
  const commandServiceRef = useRef<ICommandService | null>(null);
  const rulerGeometryRef = useRef<() => RulerGeometry | null>(() => null);
  const documentNameRef = useRef<(name: string) => void>(() => {});
  const statusListenerRef = useRef(onStatusChange);
  const [ready, setReady] = useState(false);

  // The editor is created once; the callback identity may change on every
  // parent render, so it is read through a ref rather than re-running setup.
  useEffect(() => {
    statusListenerRef.current = onStatusChange;
  }, [onStatusChange]);

  useEffect(() => {
    if (!containerRef.current || disposedRef.current) return;
    disposedRef.current = true;

    // Word types "/" as a character; Univer's block menu steals the key.
    disableSlashMenu();

    const { univer, univerAPI } = createUniver({
      theme: WORD_THEME,
      locale: LocaleType.EN_US,
      locales: {
        [LocaleType.EN_US]: buildWordLocale(
          mergeLocales(
            UniverPresetDocsCoreEnUS,
            UniverPresetDocsDrawingEnUS,
            UniverPresetDocsHyperLinkEnUS,
            UniverPresetDocsThreadCommentEnUS,
          ),
          WORD_UI_LOCALE,
        ),
      },
      presets: [
        UniverDocsCorePreset({
          container: containerRef.current,
          // Word's ribbon: a tab strip over grouped, two-row controls.
          ribbonType: "grid",
          // Univer's own footer is replaced by a Word status bar that also
          // reports the page count.
          footer: false,
          menu: RELOCATED_UNIVER_MENU_ITEMS,
        }),
        UniverDocsDrawingPreset(),
        UniverDocsHyperLinkPreset(),
        UniverDocsThreadCommentPreset(),
      ],
      plugins: [UniverDocsFindReplacePlugin],
    });

    // Docs saved before the 1.0.0-beta.2 upgrade won't have a documentStyle
    // (it used to crash at creation time in 0.25.x — see git history), so
    // they'd silently lose pagination/header-footer on load. Backfill it
    // for any saved doc that predates this, without touching its content.
    let saved = loadSnapshot<Partial<IDocumentData>>(STORAGE_KEY);

    // 1.0.0-beta.2 added a strict structural-integrity check that now runs
    // on every edit (table start/end tokens, section IDs, etc.) and throws
    // if violated — Univer 0.25.x never validated this, so a doc edited
    // under the old version (in particular through our own dataStream-
    // editing MergeTableCellsCommand) can carry corruption that only
    // surfaces now, crashing on the very first edit after load. Check
    // before handing anything to createDocument(): a corrupt snapshot is
    // backed up under its own key (nothing is silently destroyed) and the
    // editor falls back to a fresh document instead of hard-crashing.
    if (saved?.body) {
      const issues = validateDocumentStructure(saved as Pick<IDocumentData, "body" | "headers" | "footers">);
      if (issues.length > 0) {
        console.warn("[DocKaro] Saved document failed structure validation, starting fresh:", issues);
        saveSnapshot(`${STORAGE_KEY}.corrupted.${Date.now()}`, saved);
        clearSnapshot(STORAGE_KEY);
        saved = null;
      }
    }

    const initialData: Partial<IDocumentData> = saved
      ? { ...saved, documentStyle: { ...DEFAULT_DOCUMENT_STYLE, ...saved.documentStyle } }
      : { documentStyle: DEFAULT_DOCUMENT_STYLE };
    // Word names a new document rather than leaving it blank, and this name
    // is what the title bar shows and what the export is filed under.
    if (!initialData.title) initialData.title = DEFAULT_DOCUMENT_NAME;
    const fDoc = univerAPI.createDocument(initialData);

    const injector = univer.__getInjector() as Injector;
    const commandService = injector.get(ICommandService);
    const spellChecker = createSpellChecker(injector, fDoc, () => containerRef.current);
    const trackChanges = createTrackChanges(injector, fDoc);
    const registrations = [
      SetBorderPenCommand,
      ...ALL_TABLE_STYLE_COMMANDS,
      ...createWordCommands({ doc: fDoc, getContainer: () => containerRef.current }),
      ...createWordFeatureCommands(fDoc),
      createSpellCheckCommand(spellChecker),
      ...createTrackChangesCommands(trackChanges),
      createWatermarkCommand(fDoc),
    ].map((command) => commandService.registerCommand(command));
    commandServiceRef.current = commandService;
    documentNameRef.current = (name: string) => {
      fDoc.setName(name);
      saveSnapshot(STORAGE_KEY, fDoc.save());
      void refreshStatus();
    };

    const wordRibbon = installWordRibbon(injector);

    // Word puts its ruler between the ribbon and the page. Univer renders a
    // header slot in exactly that spot, so the ruler goes in as a UI part
    // rather than a sibling element that would sit above the ribbon.
    function DocumentRuler() {
      return (
        <WordRuler
          getGeometry={() => rulerGeometryRef.current()}
          handlers={{
            onIndentChange: (indents) => void commandService.executeCommand(SetIndentCommandId, indents),
            onMarginChange: (margins) => void commandService.executeCommand(SetPageMarginsCommandId, margins),
          }}
        />
      );
    }
    const rulerPart = injector.get(IUIPartsService).registerComponent(BuiltInUIPart.HEADER, () => DocumentRuler);

    // The ruler needs the page's on-screen position, which is the document
    // component's own offset inside the scene, shifted by the horizontal
    // scroll and multiplied by the zoom.
    rulerGeometryRef.current = () => {
      const container = containerRef.current;
      const renderUnit = renderManagerService.getRenderUnitById(fDoc.getId());
      const canvas = container?.querySelector("canvas");
      if (!container || !renderUnit || !canvas) return null;

      const documents = renderUnit.mainComponent as unknown as { left: number; top: number } | undefined;
      const scene = renderUnit.scene;
      const scale = scene.getAncestorScale().scaleX || 1;
      const scrollX = scene.getViewport("viewMain")?.viewportScrollX ?? 0;
      const canvasOffset = canvas.getBoundingClientRect().left - container.getBoundingClientRect().left;

      const docModel = univerInstanceService.getCurrentUnitOfType<DocumentDataModel>(UniverInstanceType.UNIVER_DOC);
      const style = docModel?.getDocumentStyle();
      if (!documents || !style?.pageSize?.width) return null;

      const paragraphStyle = currentParagraphStyle(docModel);
      const canvasRect = canvas.getBoundingClientRect();
      const containerRect = container.getBoundingClientRect();
      const scrollY = scene.getViewport("viewMain")?.viewportScrollY ?? 0;
      return {
        pageLeft: canvasOffset + (documents.left - scrollX) * scale,
        pageTop: canvasRect.top - containerRect.top + (documents.top - scrollY) * scale,
        pageWidth: style.pageSize.width * scale,
        pageHeight: (style.pageSize.height ?? 1123) * scale,
        marginLeft: style.marginLeft ?? 72,
        marginRight: style.marginRight ?? 72,
        marginTop: style.marginTop ?? 72,
        marginBottom: style.marginBottom ?? 72,
        indentStart: paragraphStyle?.indentStart?.v ?? 0,
        indentEnd: paragraphStyle?.indentEnd?.v ?? 0,
        indentFirstLine: paragraphStyle?.indentFirstLine?.v ?? 0,
        scale,
      };
    };
    // Word's table borders are draggable; Univer's have no such interaction.
    const tableResize = createTableResizeInteraction(injector, fDoc.getId(), () => containerRef.current);
    const pageChrome = hidePageMarginMarks(injector, fDoc.getId());
    const dialogFocus = restoreFocusAfterDialogs(injector, fDoc.getId());
    // Clicking the ribbon must not drop the caret, or a format toggled for the
    // next typed characters (bold, italic, font, colour...) is discarded.
    const ribbonFocus = keepCaretOnRibbonPointerDown(() => containerRef.current);
    // Repopulate DocMenuStyleService's input-style cache after each inline-
    // format command — its own re-render mutation clears the cache before
    // the next keystroke can pick it up.
    const persistInputStyle = persistInputStyleAcrossMutation(injector);

    const renderManagerService = injector.get(IRenderManagerService);
    const docSelectionManagerService = injector.get(DocSelectionManagerService);
    const univerInstanceService = injector.get(IUniverInstanceService);

    /** The paragraph the cursor is in, whose indents the ruler shows. */
    const currentParagraphStyle = (docModel: Nullable<DocumentDataModel>) => {
      const offset = docSelectionManagerService.getActiveTextRange()?.startOffset;
      if (offset == null) return undefined;
      const paragraphs = docModel?.getBody()?.paragraphs ?? [];
      return paragraphs.find((paragraph) => paragraph.startIndex >= offset)?.paragraphStyle;
    };

    // Word's status bar: which page the cursor is on, how many pages there
    // are, the word count and the zoom level.
    let statusTimeout: ReturnType<typeof setTimeout> | undefined;
    const refreshStatus = async () => {
      const listener = statusListenerRef.current;
      if (!listener) return;
      const docModel = univerInstanceService.getCurrentUnitOfType<DocumentDataModel>(UniverInstanceType.UNIVER_DOC);
      if (!docModel) return;

      const skeleton = renderManagerService.getRenderUnitById(fDoc.getId())?.with(DocSkeletonManagerService)?.getSkeleton();
      const pages = skeleton?.getSkeletonData()?.pages ?? [];
      const cursor = docSelectionManagerService.getActiveTextRange()?.startOffset ?? 0;
      const currentIndex = pages.findIndex((page) => cursor >= page.st && cursor <= page.ed);

      let wordCount = 0;
      try {
        wordCount = (await docModel.getStatistics()).words;
      } catch {
        // Statistics are best-effort: an aborted run (fast typing) must not
        // blank out the rest of the status bar.
      }

      listener({
        name: fDoc.getName(),
        wordCount,
        pageCount: Math.max(pages.length, 1),
        currentPage: currentIndex >= 0 ? currentIndex + 1 : 1,
        zoom: Math.round((docModel.zoomRatio || 1) * 100),
      });
    };
    const scheduleStatusRefresh = () => {
      clearTimeout(statusTimeout);
      statusTimeout = setTimeout(() => void refreshStatus(), STATUS_REFRESH_DELAY_MS);
    };

    // Autosave: debounce so a fast typist doesn't hit localStorage on every
    // keystroke, and flush immediately on refresh/close so the last edit
    // isn't lost (React's unmount cleanup never runs on a hard refresh).
    let saveTimeout: ReturnType<typeof setTimeout> | undefined;
    const flushSave = () => saveSnapshot(STORAGE_KEY, fDoc.save());
    // Word shows its Table Design tab whenever the cursor is inside a
    // table. The caret's offset against the document's own table ranges is
    // the reliable test: the selection's node path is empty right after a
    // table mutation (a merge, say), and `textSelection$` alone misses
    // pointer-driven moves, so the selection operation Univer's own toolbar
    // items listen to drives this too.
    const isCursorInsideTable = (): boolean | null => {
      const docDataModel = univerInstanceService.getCurrentUnitOfType<DocumentDataModel>(
        UniverInstanceType.UNIVER_DOC,
      );
      if (resolveLiveTableRange(docSelectionManagerService, docDataModel)) return true;
      const offset = docSelectionManagerService.getActiveTextRange()?.startOffset;
      // No selection at all says nothing about where the user is (a table
      // mutation clears it), so the tab keeps whatever state it had.
      if (offset == null) return null;
      const tables = docDataModel?.getBody()?.tables;
      return Boolean(tables?.some((table) => offset > table.startIndex && offset < table.endIndex));
    };
    const refreshTableContext = () => {
      const inside = isCursorInsideTable();
      if (inside !== null) wordRibbon.setTableContextActive(inside);
    };

    const commandSubscription = commandService.onCommandExecuted((command) => {
      // Using a table tool keeps the tab up even though the mutation clears
      // the cell selection it was applied to; the next selection change
      // decides again, exactly as in Word.
      if (command.id.startsWith("dockaro.command.table-")) wordRibbon.setTableContextActive(true);
      else if (command.id === SetTextSelectionsOperation.id) refreshTableContext();
      clearTimeout(saveTimeout);
      saveTimeout = setTimeout(flushSave, AUTOSAVE_DELAY_MS);
      scheduleStatusRefresh();
    });
    window.addEventListener("beforeunload", flushSave);

    const subscription = docSelectionManagerService.textSelection$.subscribe(() => {
      // Reflect the CURRENT selection exactly, like Word's Table Design tab:
      // show it only while the selection is actually inside a table, and
      // drop it the instant it isn't.
      refreshTableContext();
      scheduleStatusRefresh();
    });

    setReady(true);
    void refreshStatus();

    return () => {
      subscription.unsubscribe();
      commandSubscription.dispose();
      registrations.forEach((registration) => registration.dispose());
      wordRibbon.dispose();
      rulerPart.dispose();
      tableResize.dispose();
      pageChrome.dispose();
      dialogFocus.dispose();
      ribbonFocus.dispose();
      persistInputStyle.dispose();
      spellChecker.dispose();
      trackChanges.dispose();
      window.removeEventListener("beforeunload", flushSave);
      clearTimeout(saveTimeout);
      clearTimeout(statusTimeout);
      flushSave();
      clearRememberedTableRange();

      // univer.dispose() torn down while Univer's async preset init hasn't
      // yet reached its "steady" lifecycle stage (unmounting/navigating away
      // very quickly after mount) leaves an internal
      // firstValueFrom(lifecycle$...) with nothing left to emit once
      // disposal completes the source stream — RxJS rejects that with
      // EmptyError ("no elements in sequence"), surfaced by V8's async
      // stack traces as if thrown right here. Harmless: the instance is
      // being torn down either way. Swallow only this specific error so a
      // fast unmount doesn't crash the dev overlay / bubble as an uncaught
      // rejection, while any other dispose failure still surfaces.
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
      commandServiceRef.current = null;
      documentNameRef.current = () => {};
      rulerGeometryRef.current = () => null;
      setReady(false);
    };
  }, []);

  useImperativeHandle(apiRef, () => ({
    setName: (name: string) => documentNameRef.current(name),
    setZoom: (zoom: number) => {
      void commandServiceRef.current?.executeCommand(SetZoomCommandId, { value: zoom });
    },
    getRulerGeometry: () => rulerGeometryRef.current(),
    setIndents: (indents) => {
      void commandServiceRef.current?.executeCommand(SetIndentCommandId, indents);
    },
    setMargins: (margins) => {
      void commandServiceRef.current?.executeCommand(SetPageMarginsCommandId, margins);
    },
  }));

  return (
    <div ref={containerRef} className="relative h-full min-h-0 w-full flex-1">
      {ready && (
        <WordVerticalRuler
          getGeometry={() => rulerGeometryRef.current()}
          onMarginChange={(margins) => {
            void commandServiceRef.current?.executeCommand(SetPageMarginsCommandId, margins);
          }}
        />
      )}
    </div>
  );
}
