import {
  BooleanNumber,
  CommandType,
  ICommandService,
  IUniverInstanceService,
  TextDecoration,
  UniverInstanceType,
} from "@univerjs/core";
import type { DocumentDataModel, ICommand, IDisposable, Injector, ITextRun, ITextStyle } from "@univerjs/core";
import { DocSelectionManagerService, RichTextEditingMutation } from "@univerjs/docs";
import type { FDocument } from "@univerjs/docs/facade";

// Word's Track Changes: while it is on, what you type is marked as an
// insertion and what you delete is struck through instead of disappearing,
// and every change can then be accepted or rejected.
//
// Univer has no revision model, so the marks live in the document's own
// character formatting — Word's revision colours, underline and
// strike-through. A tracked document therefore survives a save, a reload and
// a round trip through .docx with no sidecar state, and the marks are
// exactly what a reader of the exported file sees.

/** A typing pause, after which a run of new text is marked in one go. */
const MARK_DELAY_MS = 220;

/** Word's own revision colours. */
const INSERT_COLOR = "#2b579a";
const DELETE_COLOR = "#c00000";

const INSERT_STYLE: ITextStyle = {
  cl: { rgb: INSERT_COLOR },
  ul: { s: BooleanNumber.TRUE, t: TextDecoration.SINGLE },
};
const DELETE_STYLE: ITextStyle = {
  cl: { rgb: DELETE_COLOR },
  st: { s: BooleanNumber.TRUE, t: TextDecoration.SINGLE },
};
/** What accepting or rejecting a change leaves behind. */
const CLEARED_STYLE: ITextStyle = {
  cl: { rgb: "#000000" },
  ul: { s: BooleanNumber.FALSE },
  st: { s: BooleanNumber.FALSE },
};

/**
 * Paragraph marks, table tokens and the rest of Univer's structural stream
 * live below U+0020 and carry no character formatting, so there is nothing
 * to colour or strike on them.
 */
function hasStructuralToken(text: string): boolean {
  for (const character of text) {
    if (character.charCodeAt(0) < 0x20) return true;
  }
  return false;
}

type RevisionKind = "insert" | "delete";
type Revision = { kind: RevisionKind; startOffset: number; endOffset: number };

function isInsertion(run: ITextRun): boolean {
  return run.ts?.cl?.rgb === INSERT_COLOR && run.ts?.ul?.s === BooleanNumber.TRUE;
}

function isDeletion(run: ITextRun): boolean {
  return run.ts?.cl?.rgb === DELETE_COLOR && run.ts?.st?.s === BooleanNumber.TRUE;
}

export interface TrackChangesHandle extends IDisposable {
  setEnabled: (enabled: boolean) => void;
  resolveAll: (action: "accept" | "reject") => void;
}

export function createTrackChanges(injector: Injector, doc: FDocument): TrackChangesHandle {
  const unitId = doc.getId();
  const univerInstanceService = injector.get(IUniverInstanceService);
  const commandService = injector.get(ICommandService);
  const docSelectionManagerService = injector.get(DocSelectionManagerService);

  let enabled = false;
  /** Guards against re-entering on the mutations this module itself makes. */
  let applying = false;
  let previousStream = "";
  let markTimer: ReturnType<typeof setTimeout> | null = null;

  const body = () =>
    univerInstanceService.getUnit<DocumentDataModel>(unitId, UniverInstanceType.UNIVER_DOC)?.getBody();

  const withGuard = (work: () => void) => {
    applying = true;
    try {
      work();
    } finally {
      applying = false;
    }
  };

  /**
   * What the last edit added, found by trimming the parts of the stream that
   * did not move. Typing and pasting both produce one contiguous run, which
   * is all this needs to recognise.
   */
  const insertedRange = (before: string, after: string): { start: number; end: number } | null => {
    if (after.length <= before.length) return null;
    let prefix = 0;
    while (prefix < before.length && before[prefix] === after[prefix]) prefix++;
    let suffix = 0;
    while (
      suffix < before.length - prefix &&
      before[before.length - 1 - suffix] === after[after.length - 1 - suffix]
    ) {
      suffix++;
    }
    const start = prefix;
    const end = after.length - suffix;
    if (end <= start) return null;
    if (hasStructuralToken(after.slice(start, end))) return null;
    return { start, end };
  };

  const markInsertion = () => {
    markTimer = null;
    const stream = body()?.dataStream ?? "";
    const range = insertedRange(previousStream, stream);
    previousStream = stream;
    if (!range) return;
    const caret = docSelectionManagerService.getActiveTextRange();
    withGuard(() => {
      doc.getTextRange(range.start, range.end).setTextStyle(INSERT_STYLE);
      // Styling a range selects it, and the next keystroke would replace
      // what was just typed. Put the caret back where the typist left it.
      if (caret) doc.setSelection(caret.startOffset, caret.endOffset);
      else doc.setSelection(range.end, range.end);
    });
    previousStream = body()?.dataStream ?? "";
  };

  const mutationDisposable = commandService.onCommandExecuted((command) => {
    if (command.id !== RichTextEditingMutation.id) return;
    if (!enabled || applying) {
      if (!markTimer) previousStream = body()?.dataStream ?? "";
      return;
    }
    // A delete is put back immediately — waiting would let the next
    // keystroke land in the gap. Typing is marked once per burst instead of
    // once per character, since each mark is itself a document mutation.
    const stream = body()?.dataStream ?? "";
    if (stream.length < previousStream.length) {
      if (markTimer) {
        clearTimeout(markTimer);
        markTimer = null;
      }
      markDeletion();
      return;
    }
    if (markTimer) clearTimeout(markTimer);
    markTimer = setTimeout(markInsertion, MARK_DELAY_MS);
  });

  /**
   * What the last edit removed, the mirror of `insertedRange`. Fighting
   * Univer's delete for control of the key press turned out to be a losing
   * game — its own handler is registered first and its editable surface
   * commits the change before a later listener can stop it — so a tracked
   * delete lets the delete happen and then puts the text back, struck
   * through, which is what the reader is meant to see anyway.
   */
  const deletedRange = (before: string, after: string): { start: number; text: string } | null => {
    if (after.length >= before.length) return null;
    let prefix = 0;
    while (prefix < after.length && before[prefix] === after[prefix]) prefix++;
    let suffix = 0;
    while (
      suffix < after.length - prefix &&
      before[before.length - 1 - suffix] === after[after.length - 1 - suffix]
    ) {
      suffix++;
    }
    const text = before.slice(prefix, before.length - suffix);
    if (!text || hasStructuralToken(text)) return null;
    return { start: prefix, text };
  };

  const markDeletion = () => {
    const stream = body()?.dataStream ?? "";
    const removed = deletedRange(previousStream, stream);
    previousStream = stream;
    if (!removed) return;
    withGuard(() => {
      doc.insertText(removed.start, removed.text);
      const end = removed.start + removed.text.length;
      doc.getTextRange(removed.start, end).setTextStyle(DELETE_STYLE);
      // Backspace leaves the caret before the struck text, Delete after it,
      // exactly as Word does.
      const caret = lastDeleteKey === "Delete" ? end : removed.start;
      doc.setSelection(caret, caret);
    });
    previousStream = body()?.dataStream ?? "";
  };

  let lastDeleteKey: string | null = null;
  const onKeyDown = (event: KeyboardEvent) => {
    lastDeleteKey = event.key === "Delete" || event.key === "Backspace" ? event.key : null;
  };
  window.addEventListener("keydown", onKeyDown, true);

  const revisions = (): Revision[] => {
    const found: Revision[] = [];
    for (const run of body()?.textRuns ?? []) {
      const kind: RevisionKind | null = isInsertion(run) ? "insert" : isDeletion(run) ? "delete" : null;
      if (kind) found.push({ kind, startOffset: run.st, endOffset: run.ed });
    }
    return found;
  };

  const resolveAll = (action: "accept" | "reject") => {
    // Back to front, so removing text never shifts a range not yet handled.
    const pending = revisions().sort((a, b) => b.startOffset - a.startOffset);
    withGuard(() => {
      for (const revision of pending) {
        const removes =
          (action === "accept" && revision.kind === "delete") ||
          (action === "reject" && revision.kind === "insert");
        const range = doc.getTextRange(revision.startOffset, revision.endOffset);
        if (removes) range.setText("");
        else range.setTextStyle(CLEARED_STYLE);
      }
    });
    previousStream = body()?.dataStream ?? "";
  };

  return {
    setEnabled: (next: boolean) => {
      enabled = next;
      previousStream = body()?.dataStream ?? "";
    },
    resolveAll,
    dispose: () => {
      if (markTimer) clearTimeout(markTimer);
      window.removeEventListener("keydown", onKeyDown, true);
      mutationDisposable.dispose();
    },
  };
}

export const ToggleTrackChangesCommandId = "dockaro.command.track-changes";
export const ResolveTrackedChangesCommandId = "dockaro.command.resolve-changes";

export interface IToggleTrackChangesParams {
  enabled: boolean;
}

export interface IResolveTrackedChangesParams {
  action: "accept" | "reject";
}

export function createTrackChangesCommands(handle: TrackChangesHandle): ICommand[] {
  const toggle: ICommand<IToggleTrackChangesParams> = {
    id: ToggleTrackChangesCommandId,
    type: CommandType.COMMAND,
    handler: async (_accessor, params) => {
      if (!params) return false;
      handle.setEnabled(params.enabled);
      return true;
    },
  };

  const resolve: ICommand<IResolveTrackedChangesParams> = {
    id: ResolveTrackedChangesCommandId,
    type: CommandType.COMMAND,
    handler: async (_accessor, params) => {
      if (!params) return false;
      handle.resolveAll(params.action);
      return true;
    },
  };

  return [toggle, resolve];
}
