import { LifecycleService, LifecycleStages, toDisposable } from "@univerjs/core";
import type { IDisposable, Injector } from "@univerjs/core";
import { getDocObjectById } from "@univerjs/docs-ui";
import { IRenderManagerService } from "@univerjs/engine-render";

// Univer paints little L-shaped brackets at the four corners of a page's
// text area. Word only draws those when "Text boundaries" is switched on in
// its options - a blank Word page is clean - so they are turned off here to
// match. The page edge itself stays, since Word shows that too.
//
// The marks are pure canvas chrome and were never part of any export; this
// only changes what the editor draws.
export function hidePageMarginMarks(injector: Injector, unitId: string): IDisposable {
  let disposed = false;

  void injector
    .get(LifecycleService)
    .onStage(LifecycleStages.Rendered)
    .then(() => {
      if (disposed) return;
      const docObject = getDocObjectById(unitId, injector.get(IRenderManagerService));
      // Only the fourth colour is overridden; the others keep Univer's own
      // workspace, page fill and page border.
      docObject?.docBackground.setFillColors(undefined, undefined, undefined, "transparent");
    })
    .catch(() => {
      // A document torn down before it rendered keeps Univer's default
      // chrome; there is nothing to undo.
    });

  return toDisposable(() => {
    disposed = true;
  });
}
