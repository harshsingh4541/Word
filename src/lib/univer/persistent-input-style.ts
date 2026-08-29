import {
  BaselineOffset,
  BooleanNumber,
  ICommandService,
  IUniverInstanceService,
  toDisposable,
  UniverInstanceType,
} from "@univerjs/core";
import type { DocumentDataModel, IDisposable, Injector, ITextStyle } from "@univerjs/core";
import { DocSelectionManagerService, InsertTextCommand } from "@univerjs/docs";
import {
  getStyleInTextRange,
  IMEInputCommand,
  ResetInlineFormatTextBackgroundColorCommand,
  ResetInlineFormatTextColorCommand,
  SetDocInputStyleCommand,
  SetInlineFormatBoldCommand,
  SetInlineFormatFontFamilyCommand,
  SetInlineFormatFontSizeCommand,
  SetInlineFormatItalicCommand,
  SetInlineFormatStrikethroughCommand,
  SetInlineFormatSubscriptCommand,
  SetInlineFormatSuperscriptCommand,
  SetInlineFormatTextBackgroundColorCommand,
  SetInlineFormatTextColorCommand,
  SetInlineFormatUnderlineCommand,
} from "@univerjs/docs-ui";

// At a collapsed caret, Univer's SetInlineFormatCommand records the intended
// format in DocMenuStyleService's "input style" cache — DocInputController
// reads that cache on the next keystroke to style the inserted characters.
// The command then runs a RichTextEditingMutation to force a re-render, but
// that mutation re-emits the current selection, and DocMenuStyleService
// clears the cache on ANY selection emission. So the mutation the command
// just fired throws away the cache the command just set, before the user's
// keystroke can pick it up.
//
// SetDocInputStyleCommand is the exported command that writes the same
// cache, so this repopulates it right after each inline-format command runs.
// A small in-memory replica accumulates the pending style across chained
// clicks (Bold then Italic then type) and gets cleared when the caret moves
// or the user types — the same lifecycle the real cache follows in Univer's
// design, minus the self-inflicted wipe.

/** DocMenuStyleService's own default; matches DEFAULT_TEXT_STYLE in docs-ui. */
const DEFAULT_TEXT_STYLE: ITextStyle = { ff: "Arial", fs: 11 };

/** Which inline-format command writes which text-style property. */
const COMMAND_TO_STYLE_KEY: Readonly<Record<string, "bl" | "it" | "ul" | "st" | "va" | "fs" | "ff" | "cl" | "bg">> = {
  [SetInlineFormatBoldCommand.id]: "bl",
  [SetInlineFormatItalicCommand.id]: "it",
  [SetInlineFormatUnderlineCommand.id]: "ul",
  [SetInlineFormatStrikethroughCommand.id]: "st",
  [SetInlineFormatSubscriptCommand.id]: "va",
  [SetInlineFormatSuperscriptCommand.id]: "va",
  [SetInlineFormatFontSizeCommand.id]: "fs",
  [SetInlineFormatFontFamilyCommand.id]: "ff",
  [SetInlineFormatTextColorCommand.id]: "cl",
  [SetInlineFormatTextBackgroundColorCommand.id]: "bg",
  [ResetInlineFormatTextColorCommand.id]: "cl",
  [ResetInlineFormatTextBackgroundColorCommand.id]: "bg",
};

/** Commands whose completion consumes the pending style, so tracking resets. */
const INPUT_COMMAND_IDS: readonly string[] = [InsertTextCommand.id, IMEInputCommand.id];

type StyleReplica = Partial<Record<string, unknown>>;

/**
 * Re-populate DocMenuStyleService's input-style cache after every inline-
 * format command so the format applied at a collapsed caret survives the
 * command's own re-render mutation.
 */
export function persistInputStyleAcrossMutation(injector: Injector): IDisposable {
  const commandService = injector.get(ICommandService);
  const univerInstanceService = injector.get(IUniverInstanceService);
  const docSelectionManagerService = injector.get(DocSelectionManagerService);

  // Our replica of what DocMenuStyleService's cache would hold if it weren't
  // being wiped by every mutation. Merged, not replaced, so chained clicks
  // (Bold then Italic) accumulate correctly.
  let tracked: StyleReplica = {};
  let lastOffset: number | null = null;

  // Reset on real caret movement. A mutation re-emitting the same offset
  // isn't a caret move and leaves tracking alone.
  const selectionSub = docSelectionManagerService.textSelection$.subscribe(() => {
    const range = docSelectionManagerService.getActiveTextRange();
    const offset = range && range.collapsed ? range.startOffset : null;
    if (offset !== lastOffset) {
      tracked = {};
      lastOffset = offset;
    }
  });

  // Reset when the user actually types. The style has been consumed on the
  // typed character; subsequent characters inherit from the previous text
  // run, exactly like Word.
  const executed = commandService.onCommandExecuted((command) => {
    if (INPUT_COMMAND_IDS.includes(command.id)) {
      tracked = {};
      return;
    }

    const key = COMMAND_TO_STYLE_KEY[command.id];
    if (!key) return;

    const range = docSelectionManagerService.getActiveTextRange();
    if (!range || !range.collapsed) return;

    lastOffset = range.startOffset;

    const doc = univerInstanceService.getCurrentUnitOfType<DocumentDataModel>(UniverInstanceType.UNIVER_DOC);
    const body = doc?.getSelfOrHeaderFooterModel(range.segmentId ?? "")?.getBody();
    if (!body) return;

    // Toggle commands need to know whether the property is already set —
    // both in the underlying text (Univer's own logic) and in our tracked
    // pending style (chained clicks). Value commands ignore this.
    const textStyle = getStyleInTextRange(body, range, DEFAULT_TEXT_STYLE) as ITextStyle;
    const effective = { ...textStyle, ...tracked };

    const value = computeStyleValue(command.id, key, effective, command.params as { value?: unknown } | undefined);
    if (value === undefined) return;

    tracked = { ...tracked, [key]: value };

    void commandService.executeCommand(SetDocInputStyleCommand.id, { style: { ...tracked } });
  });

  return toDisposable(() => {
    selectionSub.unsubscribe();
    executed.dispose();
  });
}

/**
 * Replicates DocMenuStyleService's own decision for what the input-style
 * cache should hold after each inline-format command runs at a collapsed
 * caret. Kept small and inline; the toggle table mirrors getReverseFormatValue
 * from @univerjs/docs-ui.
 */
function computeStyleValue(
  commandId: string,
  key: "bl" | "it" | "ul" | "st" | "va" | "fs" | "ff" | "cl" | "bg",
  effective: ITextStyle & Record<string, unknown>,
  params: { value?: unknown } | undefined,
): unknown {
  switch (key) {
    case "bl":
    case "it":
      return effective[key] === BooleanNumber.TRUE ? BooleanNumber.FALSE : BooleanNumber.TRUE;
    case "ul":
    case "st": {
      const current = effective[key] as { s?: BooleanNumber } | undefined;
      return current?.s === BooleanNumber.TRUE ? { s: BooleanNumber.FALSE } : { s: BooleanNumber.TRUE };
    }
    case "va":
      if (commandId === SetInlineFormatSubscriptCommand.id) {
        return effective.va === BaselineOffset.SUBSCRIPT ? BaselineOffset.NORMAL : BaselineOffset.SUBSCRIPT;
      }
      return effective.va === BaselineOffset.SUPERSCRIPT ? BaselineOffset.NORMAL : BaselineOffset.SUPERSCRIPT;
    case "cl":
    case "bg":
      if (
        commandId === ResetInlineFormatTextColorCommand.id ||
        commandId === ResetInlineFormatTextBackgroundColorCommand.id
      ) {
        return { rgb: null };
      }
      return { rgb: params?.value };
    case "fs":
    case "ff":
      return params?.value;
  }
}
