import { CommandType, ICommandService, IUniverInstanceService, UniverInstanceType } from "@univerjs/core";
import type { DocumentDataModel, ICommand, IDisposable, Injector } from "@univerjs/core";
import { RichTextEditingMutation } from "@univerjs/docs";
import { calcDocRangePositions } from "@univerjs/docs-ui";
import { IRenderManagerService } from "@univerjs/engine-render";
import type { FDocument } from "@univerjs/docs/facade";

// Word underlines misspellings with a red squiggle and offers corrections on
// right-click. Univer has neither, and the browser's own spell checker never
// sees the text because the document is painted on a canvas.
//
// So this does what Word does, in the same two parts: a hunspell dictionary
// (the same one Firefox and LibreOffice ship) decides what is misspelled,
// and an overlay draws the squiggles over the canvas. Nothing is written to
// the document — the marks are a view of it, so they never end up in a save
// or an export.

const DICTIONARY_BASE = "/dictionaries/en";
/** Letters, plus the apostrophes and hyphens that live inside words. */
const WORD_PATTERN = /[A-Za-z][A-Za-z'’-]*/g;
/** Below this a "misspelling" is almost always an initial or a variable. */
const MIN_WORD_LENGTH = 3;
/** Re-checking on every keystroke is wasted work; this is a typing pause. */
const RESCAN_DELAY_MS = 450;
/** Enough to cover a screen many times over without scanning a novel. */
const MAX_MARKS = 400;
const MAX_SUGGESTIONS = 5;

type Misspelling = { word: string; startOffset: number; endOffset: number };

type Speller = { correct: (word: string) => boolean; suggest: (word: string) => string[] };

let spellerPromise: Promise<Speller | null> | null = null;

/**
 * Loads the dictionary once per session, on first use rather than at start-up
 * — it is half a megabyte, and a document nobody spell-checks should not pay
 * for it.
 */
function loadSpeller(): Promise<Speller | null> {
  spellerPromise ??= (async () => {
    try {
      const [{ default: NSpell }, aff, dic] = await Promise.all([
        import("nspell"),
        fetch(`${DICTIONARY_BASE}/en.aff`).then((response) => response.text()),
        fetch(`${DICTIONARY_BASE}/en.dic`).then((response) => response.text()),
      ]);
      return NSpell(aff, dic) as Speller;
    } catch {
      // No dictionary means no squiggles, which is better than a broken
      // editor; the toggle simply appears to do nothing.
      return null;
    }
  })();
  return spellerPromise;
}

function findMisspellings(text: string, speller: Speller): Misspelling[] {
  const found: Misspelling[] = [];
  WORD_PATTERN.lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = WORD_PATTERN.exec(text)) !== null) {
    const word = match[0];
    if (word.length < MIN_WORD_LENGTH) continue;
    // Word leaves ALL-CAPS acronyms alone, and so does this.
    if (word === word.toUpperCase()) continue;
    if (speller.correct(word)) continue;
    found.push({ word, startOffset: match.index, endOffset: match.index + word.length });
    if (found.length >= MAX_MARKS) break;
  }
  return found;
}

/** A red squiggle, drawn once and repeated along the width of a word. */
const SQUIGGLE = `url("data:image/svg+xml;utf8,${encodeURIComponent(
  '<svg xmlns="http://www.w3.org/2000/svg" width="6" height="3" viewBox="0 0 6 3">' +
    '<path d="M0 2 L1.5 0.5 L3 2 L4.5 0.5 L6 2" fill="none" stroke="#d13438" stroke-width="1"/>' +
    "</svg>",
)}")`;

export interface SpellCheckHandle extends IDisposable {
  setEnabled: (enabled: boolean) => void;
}

export function createSpellChecker(
  injector: Injector,
  doc: FDocument,
  getContainer: () => HTMLElement | null,
): SpellCheckHandle {
  const unitId = doc.getId();
  const renderManagerService = injector.get(IRenderManagerService);
  const univerInstanceService = injector.get(IUniverInstanceService);

  let enabled = false;
  let disposed = false;
  let marks: Misspelling[] = [];
  let rescanTimer: ReturnType<typeof setTimeout> | null = null;
  let overlay: HTMLElement | null = null;
  let popup: HTMLElement | null = null;

  const ensureOverlay = (): HTMLElement | null => {
    if (!getContainer()) return null;
    if (!overlay) {
      overlay = document.createElement("div");
      overlay.style.position = "fixed";
      overlay.style.inset = "0";
      overlay.style.pointerEvents = "none";
      // On the body, not in the editor container: Univer's canvas sits in
      // its own stacking context, so an overlay nested beside it is painted
      // underneath however high its z-index goes.
      overlay.style.zIndex = "100";
      document.body.appendChild(overlay);
    }
    return overlay;
  };

  const closePopup = () => {
    popup?.remove();
    popup = null;
  };

  const clearOverlay = () => {
    if (overlay) overlay.textContent = "";
    closePopup();
  };

  /** Where each misspelling sits on screen right now. */
  const markRects = (): { mark: Misspelling; rect: DOMRect }[] => {
    const render = renderManagerService.getRenderUnitById(unitId);
    if (!render) return [];
    const placed: { mark: Misspelling; rect: DOMRect }[] = [];
    for (const mark of marks) {
      const bounds = calcDocRangePositions(
        { startOffset: mark.startOffset, endOffset: mark.endOffset, collapsed: false },
        render,
      );
      for (const bound of bounds ?? []) {
        placed.push({
          mark,
          rect: new DOMRect(bound.left, bound.top, bound.right - bound.left, bound.bottom - bound.top),
        });
      }
    }
    return placed;
  };

  const paint = () => {
    const target = ensureOverlay();
    if (!target) return;
    target.textContent = "";
    if (!enabled) return;

    const canvasRect = getContainer()?.querySelector("canvas")?.getBoundingClientRect();
    for (const { rect } of markRects()) {
      // A word scrolled out of the page area must not leave its squiggle
      // floating over the ribbon.
      if (canvasRect && (rect.bottom < canvasRect.top || rect.top > canvasRect.bottom)) continue;
      const squiggle = document.createElement("div");
      squiggle.style.position = "fixed";
      squiggle.style.left = `${rect.left}px`;
      squiggle.style.top = `${rect.bottom - 2}px`;
      squiggle.style.width = `${rect.width}px`;
      squiggle.style.height = "3px";
      squiggle.style.backgroundImage = SQUIGGLE;
      squiggle.style.backgroundRepeat = "repeat-x";
      target.appendChild(squiggle);
    }
  };

  const rescan = async () => {
    if (!enabled || disposed) return;
    const speller = await loadSpeller();
    if (!speller || disposed || !enabled) return;
    const body = univerInstanceService
      .getUnit<DocumentDataModel>(unitId, UniverInstanceType.UNIVER_DOC)
      ?.getBody();
    marks = body ? findMisspellings(body.dataStream, speller) : [];
    paint();
  };

  const scheduleRescan = () => {
    if (rescanTimer) clearTimeout(rescanTimer);
    rescanTimer = setTimeout(() => void rescan(), RESCAN_DELAY_MS);
  };

  // Right-clicking a squiggle offers corrections, exactly as Word does; a
  // right-click anywhere else falls through to Univer's own menu.
  const hitTest = (clientX: number, clientY: number) =>
    markRects().find(
      ({ rect }) =>
        clientX >= rect.left && clientX <= rect.right && clientY >= rect.top && clientY <= rect.bottom,
    );

  // Univer opens its own context menu from the pointer press, not from the
  // `contextmenu` event, so a right-press on a misspelling has to be caught
  // there too or both menus appear.
  const onPointerDown = (event: PointerEvent | MouseEvent) => {
    if (!enabled || event.button !== 2) return;
    if (!hitTest(event.clientX, event.clientY)) return;
    event.preventDefault();
    event.stopPropagation();
  };

  const onContextMenu = (event: MouseEvent) => {
    if (!enabled) return;
    const hit = markRects().find(
      ({ rect }) =>
        event.clientX >= rect.left &&
        event.clientX <= rect.right &&
        event.clientY >= rect.top &&
        event.clientY <= rect.bottom,
    );
    if (!hit) return;
    event.preventDefault();
    event.stopPropagation();
    void showSuggestions(hit.mark, event.clientX, event.clientY);
  };

  const showSuggestions = async (mark: Misspelling, x: number, y: number) => {
    const speller = await loadSpeller();
    if (!speller) return;
    closePopup();

    const menu = document.createElement("div");
    menu.style.position = "fixed";
    menu.style.left = `${x}px`;
    menu.style.top = `${y}px`;
    menu.style.zIndex = "1000";
    menu.style.minWidth = "160px";
    menu.style.padding = "4px";
    menu.style.borderRadius = "6px";
    menu.style.background = "#ffffff";
    menu.style.color = "#201f1e";
    menu.style.font = "13px system-ui, sans-serif";
    menu.style.boxShadow = "0 4px 16px rgba(0,0,0,0.18)";
    menu.style.pointerEvents = "auto";

    const addItem = (label: string, onPick: () => void, italic = false) => {
      const item = document.createElement("div");
      item.textContent = label;
      item.style.padding = "5px 10px";
      item.style.borderRadius = "4px";
      item.style.cursor = "pointer";
      if (italic) item.style.fontStyle = "italic";
      item.addEventListener("mouseenter", () => (item.style.background = "#f3f2f1"));
      item.addEventListener("mouseleave", () => (item.style.background = "transparent"));
      item.addEventListener("mousedown", (event) => {
        event.preventDefault();
        onPick();
        closePopup();
      });
      menu.appendChild(item);
    };

    const suggestions = speller.suggest(mark.word).slice(0, MAX_SUGGESTIONS);
    if (suggestions.length === 0) addItem("No suggestions", () => {}, true);
    for (const suggestion of suggestions) {
      addItem(suggestion, () => {
        doc.getTextRange(mark.startOffset, mark.endOffset).setText(suggestion);
        scheduleRescan();
      });
    }

    addItem("Ignore", () => {
      marks = marks.filter((other) => other.startOffset !== mark.startOffset);
      paint();
    }, true);

    document.body.appendChild(menu);
    popup = menu;
  };

  const commandDisposable = injector.get(ICommandService).onCommandExecuted((command) => {
    if (command.id === RichTextEditingMutation.id) scheduleRescan();
  });

  // The squiggles are drawn in screen space, so anything that moves the page
  // under them has to move them too.
  const onViewportChange = () => {
    closePopup();
    if (enabled) paint();
  };
  window.addEventListener("scroll", onViewportChange, true);
  window.addEventListener("resize", onViewportChange);
  document.addEventListener("contextmenu", onContextMenu, true);
  document.addEventListener("pointerdown", onPointerDown, true);
  document.addEventListener("mousedown", onPointerDown, true);
  document.addEventListener("mousedown", closePopup);

  return {
    setEnabled: (next: boolean) => {
      if (next === enabled) return;
      enabled = next;
      if (enabled) void rescan();
      else clearOverlay();
    },
    dispose: () => {
      disposed = true;
      if (rescanTimer) clearTimeout(rescanTimer);
      window.removeEventListener("scroll", onViewportChange, true);
      window.removeEventListener("resize", onViewportChange);
      document.removeEventListener("contextmenu", onContextMenu, true);
      document.removeEventListener("pointerdown", onPointerDown, true);
      document.removeEventListener("mousedown", onPointerDown, true);
      document.removeEventListener("mousedown", closePopup);
      clearOverlay();
      overlay?.remove();
      overlay = null;
      commandDisposable.dispose();
    },
  };
}

/** Wraps the checker in a command so the ribbon can drive it. */
export const ToggleSpellCheckCommandId = "dockaro.command.spell-check";

export interface IToggleSpellCheckParams {
  enabled: boolean;
}

export function createSpellCheckCommand(handle: SpellCheckHandle): ICommand<IToggleSpellCheckParams> {
  return {
    id: ToggleSpellCheckCommandId,
    type: CommandType.COMMAND,
    handler: async (_accessor, params) => {
      if (!params) return false;
      handle.setEnabled(params.enabled);
      return true;
    },
  };
}
