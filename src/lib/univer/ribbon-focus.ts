import { toDisposable } from "@univerjs/core";
import type { IDisposable } from "@univerjs/core";

// Clicking a ribbon control moves DOM focus off the document canvas.
// Univer's live text selection is scoped to that focus, and DocMenuStyleService
// drops its pending-input-style cache on any selection emission — the cache
// is exactly what makes "click Bold, then type" work at a collapsed caret,
// so losing it loses the format for bold, italic, underline, strikethrough,
// subscript, superscript, font, size and colour alike.
//
// Every rich text editor solves this the same way: preventDefault the toolbar
// container's mousedown so the browser never shifts focus in the first place.
// Word and Univer's own toolbar buttons both do this internally. This applies
// the same behaviour to this app's ribbon.
//
// mousedown is important, not pointerdown: the "move focus to <button>"
// default action is a default of mousedown, and canceling pointerdown does
// not cancel the compat mousedown that follows it.

/** Controls that legitimately need DOM focus of their own. */
const FOCUS_OWNERS = [
  "input",
  "textarea",
  "select",
  "[contenteditable='']",
  "[contenteditable='true']",
].join(",");

/**
 * Keeps the document's caret while the ribbon is clicked.
 *
 * @param getContainer Returns the element Univer renders the editor into.
 */
export function keepCaretOnRibbonPointerDown(getContainer: () => HTMLElement | null): IDisposable {
  const onMouseDown = (event: MouseEvent) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;

    // The canvas owns its own pointer handling — placing the caret depends on
    // it, so it must keep the default behaviour.
    if (target.closest("canvas")) return;

    // A real text field has to be allowed to take focus (the row-height box,
    // the font-size box, a find/replace field...).
    if (target.closest(FOCUS_OWNERS)) return;

    // Cancelling only the default focus shift leaves click handling untouched,
    // so the ribbon's own buttons and menus still fire normally.
    event.preventDefault();
  };

  const container = getContainer();
  container?.addEventListener("mousedown", onMouseDown, true);

  return toDisposable(() => {
    container?.removeEventListener("mousedown", onMouseDown, true);
  });
}
