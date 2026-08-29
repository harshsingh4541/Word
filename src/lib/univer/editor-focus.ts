import { ICommandService, toDisposable } from "@univerjs/core";
import type { IDisposable, Injector } from "@univerjs/core";
import { CreateDocTableCommand, DocSelectionRenderService } from "@univerjs/docs-ui";
import { IRenderManagerService } from "@univerjs/engine-render";

// Commands that are reached through one of Univer's modal dialogs. The
// dialog's OK button holds DOM focus when it closes, so the document's
// hidden input never gets it back: the caret is in the right place (Univer's
// create-table command does set the text range to the first cell) but the
// next keystroke goes nowhere. Word puts you straight into the new table's
// first cell, typing immediately.
//
// Insert > Table > "Insert table..." was the reproducer; the size-grid path
// beside it never lost focus, which is why this is scoped to the command
// rather than applied blindly after every ribbon action.
const AFTER_DIALOG_COMMANDS: readonly string[] = [
  CreateDocTableCommand.id,
  // Univer's own paragraph and page dialogs, same shape.
  "doc-paragraph-setting.command",
  "docs.command.page-setup",
];

/**
 * Return keyboard focus to the document after a dialog-driven command, so
 * the user can type where the command left the caret.
 */
export function restoreFocusAfterDialogs(injector: Injector, unitId: string): IDisposable {
  const renderManagerService = injector.get(IRenderManagerService);

  return toDisposable(
    injector.get(ICommandService).onCommandExecuted((command) => {
      if (!AFTER_DIALOG_COMMANDS.includes(command.id)) return;
      // The dialog is torn down in the same tick as the command; focus has
      // to be claimed after React has removed it, or it bounces back.
      requestAnimationFrame(() => {
        renderManagerService.getRenderUnitById(unitId)?.with(DocSelectionRenderService).focus();
      });
    }).dispose,
  );
}
