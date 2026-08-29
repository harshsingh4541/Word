import { DocParagraphMenuService } from "@univerjs/docs-ui";

// Univer ships a Notion-style block menu that opens when you type "/" at a
// collapsed cursor, and it swallows the keystroke: `_openSlashMenu` calls
// `preventDefault()` + `stopPropagation()`, so the character never reaches
// the document. That makes dates (03/12/2026), fractions, URLs and file
// paths impossible to type.
//
// Word has no slash menu — "/" is just a character — so the gate is closed
// here. Both entry points are neutered (the keydown path and the
// input/IME path); everything else about the paragraph menu, including the
// hover handle beside each block, is untouched: with the gate shut a "/"
// keydown falls through to Univer's normal `hideParagraphMenu` branch,
// exactly like any other printable character.
const GATES = ["_shouldOpenSlashMenu", "_shouldOpenSlashMenuFromInput"] as const;

let patched = false;

/**
 * Make "/" type a slash instead of opening Univer's block menu. Patches the
 * service prototype, so it covers the body, header/footer editors and any
 * document opened later; calling it more than once is a no-op.
 */
export function disableSlashMenu(): void {
  if (patched) return;

  // Private members, reached deliberately: there is no config flag for this
  // and no public API to unsubscribe the handler.
  const proto = DocParagraphMenuService.prototype as unknown as Record<string, unknown>;
  if (!GATES.every((gate) => typeof proto[gate] === "function")) {
    // A Univer upgrade renamed or removed the slash menu. Leave it alone
    // rather than half-patching it.
    return;
  }

  patched = true;
  for (const gate of GATES) proto[gate] = () => false;
}
