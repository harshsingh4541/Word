import {
  ICommandService,
  IUniverInstanceService,
  LifecycleService,
  LifecycleStages,
  UniverInstanceType,
  toDisposable,
} from "@univerjs/core";
import type { DocumentDataModel, IDisposable, Injector, ITableCellMargin } from "@univerjs/core";
import { DocEventManagerService } from "@univerjs/docs-ui";
import { CURSOR_TYPE, getTableIdAndSliceIndex, IRenderManagerService } from "@univerjs/engine-render";
import { DocSkeletonManagerService } from "@univerjs/docs";
import { ResizeTableColumnCommandId, ResizeTableRowCommandId } from "./table-style-commands";

// Word lets you drag a table's borders with the mouse. Univer's open-source
// Docs has no such interaction at all (there is no resize command anywhere
// in docs-ui - only an internal fit-to-page helper), so this adds it:
// hovering a cell border switches the cursor, dragging shows a guide line
// and releasing applies the new size as a single undoable command.

/** How close to a border the pointer has to be, in document pixels. */
const EDGE_TOLERANCE = 5;
/** Univer's default cell padding, used when a table doesn't set its own. */
const DEFAULT_CELL_MARGIN = { start: 10, end: 10, top: 5, bottom: 5 };

/**
 * What Univer's hover service reports for the cell under the pointer. Its
 * own interface isn't part of the package's public types, so the shape is
 * mirrored here.
 */
interface TableCellBound {
  rect: { left: number; right: number; top: number; bottom: number };
  pageIndex: number;
  rowIndex: number;
  colIndex: number;
  tableId: string;
}

/** Just enough of the skeleton to map a slice's row back to the table's. */
type SkeletonLike = {
  getSkeletonData: () => { pages?: { skeTables?: Map<string, { rows: { index: number }[] }> }[] } | undefined | null;
};

/**
 * A table that spills onto another page is laid out as several skeleton
 * "slices" — `${tableId}#-#0`, `#-#1`, ... — and each slice numbers its
 * rows from zero again. Hover events report those slice coordinates, while
 * every table command addresses the source table, so the two are reconciled
 * here: the slice's id gives the table, and the slice's own row skeleton
 * carries the row's real index within it.
 */
function toSourceCell(cell: TableCellBound, skeleton: SkeletonLike | undefined): TableCellBound {
  const { tableId } = getTableIdAndSliceIndex(cell.tableId);
  if (tableId === cell.tableId) return cell;
  const rows = skeleton?.getSkeletonData()?.pages?.[cell.pageIndex]?.skeTables?.get(cell.tableId)?.rows;
  return { ...cell, tableId, rowIndex: rows?.[cell.rowIndex]?.index ?? cell.rowIndex };
}

type ResizeTarget =
  | { kind: "column"; tableId: string; columnIndex: number; position: number }
  | { kind: "row"; tableId: string; rowIndex: number; position: number; top: number };

interface SceneLike {
  getAncestorScale: () => { scaleX: number; scaleY: number };
  getViewport: (key: string) => { viewportScrollX: number; viewportScrollY: number } | null | undefined;
  setCursor: (cursor: CURSOR_TYPE) => void;
  resetCursor: () => void;
  onPointerMove$: { subscribeEvent: (callback: (event: PointerEvent) => void) => IDisposable | undefined };
}

/** Univer's canvas coordinates: scaled by zoom and offset by the scroll. */
function toSceneCoords(scene: SceneLike, offsetX: number, offsetY: number) {
  const { scaleX, scaleY } = scene.getAncestorScale();
  const viewport = scene.getViewport("viewMain");
  if (!viewport) return { x: offsetX, y: offsetY, scaleX, scaleY };
  return {
    x: offsetX / scaleX + viewport.viewportScrollX,
    y: offsetY / scaleY + viewport.viewportScrollY,
    scaleX,
    scaleY,
  };
}

/**
 * Univer reports the hovered cell's *content* box, inset by the cell's
 * padding, while the border a user aims at sits out at the edge of the
 * padding. Reconstructing the border positions from the cell's own margins
 * is what makes the hit zone land where the line is drawn.
 */
/** The cell's border box: its content rect grown by the cell's padding. */
function cellBorderBox(cell: TableCellBound, margin: ITableCellMargin | undefined) {
  return {
    left: cell.rect.left - (margin?.start?.v ?? DEFAULT_CELL_MARGIN.start),
    right: cell.rect.right + (margin?.end?.v ?? DEFAULT_CELL_MARGIN.end),
    top: cell.rect.top - (margin?.top?.v ?? DEFAULT_CELL_MARGIN.top),
    bottom: cell.rect.bottom + (margin?.bottom?.v ?? DEFAULT_CELL_MARGIN.bottom),
  };
}

/**
 * Univer reports the hovered cell's *content* box, inset by the cell's
 * padding, while the border a user aims at sits out at the edge of that
 * padding. Reconstructing the border positions from the cell's own margins
 * is what puts the hit zone where the line is actually drawn.
 */
function resolveTarget(cell: TableCellBound, margin: ITableCellMargin | undefined, x: number, y: number): ResizeTarget | null {
  const { left, right, top, bottom } = cellBorderBox(cell, margin);
  if (Math.abs(x - left) <= EDGE_TOLERANCE && cell.colIndex > 0) {
    return { kind: "column", tableId: cell.tableId, columnIndex: cell.colIndex - 1, position: left };
  }
  if (Math.abs(x - right) <= EDGE_TOLERANCE) {
    return { kind: "column", tableId: cell.tableId, columnIndex: cell.colIndex, position: right };
  }
  if (Math.abs(y - bottom) <= EDGE_TOLERANCE) {
    return { kind: "row", tableId: cell.tableId, rowIndex: cell.rowIndex, position: bottom, top };
  }
  if (Math.abs(y - top) <= EDGE_TOLERANCE && cell.rowIndex > 0) {
    return { kind: "row", tableId: cell.tableId, rowIndex: cell.rowIndex - 1, position: top, top };
  }
  return null;
}

/** Whether the pointer is still on (or right beside) this cell. */
function isNearCell(cell: TableCellBound, margin: ITableCellMargin | undefined, x: number, y: number): boolean {
  const { left, right, top, bottom } = cellBorderBox(cell, margin);
  return (
    x >= left - EDGE_TOLERANCE && x <= right + EDGE_TOLERANCE && y >= top - EDGE_TOLERANCE && y <= bottom + EDGE_TOLERANCE
  );
}

function createGuide(container: HTMLElement, vertical: boolean): HTMLElement {
  const guide = document.createElement("div");
  guide.style.position = "absolute";
  guide.style.zIndex = "40";
  guide.style.pointerEvents = "none";
  // Word draws a thin dotted line while a border is being dragged.
  if (vertical) {
    guide.style.top = "0";
    guide.style.bottom = "0";
    guide.style.width = "0";
    guide.style.borderLeft = "1px dashed rgba(0, 0, 0, 0.65)";
  } else {
    guide.style.left = "0";
    guide.style.right = "0";
    guide.style.height = "0";
    guide.style.borderTop = "1px dashed rgba(0, 0, 0, 0.65)";
  }
  container.appendChild(guide);
  return guide;
}

/**
 * Wires mouse resizing of table borders into a document's canvas.
 *
 * @param injector The Univer instance's injector.
 * @param unitId The document being edited.
 * @param getContainer Returns the element the canvas is rendered into.
 */
export function createTableResizeInteraction(
  injector: Injector,
  unitId: string,
  getContainer: () => HTMLElement | null,
): IDisposable {
  let inner: IDisposable | null = null;
  let disposed = false;

  // The document's render modules (the hover service among them) only exist
  // once Univer has rendered for the first time; asking for them any earlier
  // throws a dependency-resolution error.
  void injector
    .get(LifecycleService)
    .onStage(LifecycleStages.Rendered)
    .then(() => {
      if (disposed) return;
      inner = attachTableResize(injector, unitId, getContainer);
    })
    .catch(() => {
      // A document torn down before it finished rendering simply never gets
      // the interaction; nothing to clean up.
    });

  return toDisposable(() => {
    disposed = true;
    inner?.dispose();
    inner = null;
  });
}

function attachTableResize(
  injector: Injector,
  unitId: string,
  getContainer: () => HTMLElement | null,
): IDisposable {
  const renderManagerService = injector.get(IRenderManagerService);
  const commandService = injector.get(ICommandService);
  const univerInstanceService = injector.get(IUniverInstanceService);
  const render = renderManagerService.getRenderUnitById(unitId);
  if (!render) return toDisposable(() => {});

  const cellMargin = (cell: TableCellBound): ITableCellMargin | undefined => {
    const table = univerInstanceService
      .getCurrentUnitOfType<DocumentDataModel>(UniverInstanceType.UNIVER_DOC)
      ?.getSnapshot().tableSource?.[cell.tableId];
    return table?.tableRows?.[cell.rowIndex]?.tableCells?.[cell.colIndex]?.margin ?? table?.cellMargin;
  };

  const scene = render.scene as unknown as SceneLike;
  const eventManager = render.with(DocEventManagerService);

  let hoveredCell: TableCellBound | null = null;
  let target: ResizeTarget | null = null;
  let cursorApplied = false;

  const clearCursor = () => {
    if (!cursorApplied) return;
    cursorApplied = false;
    scene.resetCursor();
  };

  // The service reports no cell while the pointer sits in the padding
  // between two cells - which is exactly where a border is - so the last
  // cell is kept until the pointer moves clear of it (see resolveTarget).
  const skeletonManagerService = render.with(DocSkeletonManagerService);
  const hoverSubscription = eventManager.hoverTableCellRealTime$.subscribe((cell) => {
    if (cell) hoveredCell = toSourceCell(cell, skeletonManagerService?.getSkeleton() as SkeletonLike | undefined);
  });

  const pointerMove = scene.onPointerMove$.subscribeEvent((event) => {
    if (!hoveredCell || event.buttons > 0) return;
    const { x, y } = toSceneCoords(scene, event.offsetX, event.offsetY);
    const margin = cellMargin(hoveredCell);
    if (!isNearCell(hoveredCell, margin, x, y)) {
      hoveredCell = null;
      target = null;
      clearCursor();
      return;
    }
    target = resolveTarget(hoveredCell, margin, x, y);
    if (target) {
      cursorApplied = true;
      scene.setCursor(target.kind === "column" ? CURSOR_TYPE.COLUMN_RESIZE : CURSOR_TYPE.ROW_RESIZE);
    } else {
      clearCursor();
    }
  });

  // Univer's own canvas handler owns pointerdown (the scene never sees it,
  // and letting it through would start a cell selection instead of a
  // resize), so the drag starts from a capture-phase DOM listener that
  // swallows the event when a border is under the pointer.
  const onPointerDown = (event: PointerEvent) => {
    const container = getContainer();
    if (!target || !container) return;
    event.preventDefault();
    event.stopPropagation();

    const dragged = target;
    const { scaleX, scaleY } = scene.getAncestorScale();
    const scale = dragged.kind === "column" ? scaleX : scaleY;
    const startClient = dragged.kind === "column" ? event.clientX : event.clientY;
    const containerRect = container.getBoundingClientRect();
    const guide = createGuide(container, dragged.kind === "column");
    const place = (client: number) => {
      if (dragged.kind === "column") guide.style.left = `${client - containerRect.left}px`;
      else guide.style.top = `${client - containerRect.top}px`;
    };
    place(startClient);

    // The drag is followed on the window so it keeps working when the
    // pointer leaves the canvas, exactly like dragging a border in Word.
    const onMove = (moveEvent: PointerEvent) => {
      place(dragged.kind === "column" ? moveEvent.clientX : moveEvent.clientY);
    };
    const onUp = (upEvent: PointerEvent) => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      guide.remove();
      const endClient = dragged.kind === "column" ? upEvent.clientX : upEvent.clientY;
      // Screen pixels are zoomed document pixels; the model wants the latter.
      const delta = (endClient - startClient) / (scale || 1);
      if (Math.abs(delta) < 1) return;
      if (dragged.kind === "column") {
        void commandService.executeCommand(ResizeTableColumnCommandId, {
          tableId: dragged.tableId,
          columnIndex: dragged.columnIndex,
          delta,
        });
      } else {
        void commandService.executeCommand(ResizeTableRowCommandId, {
          tableId: dragged.tableId,
          rowIndex: dragged.rowIndex,
          height: dragged.position - dragged.top + delta,
        });
      }
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
  };

  const container = getContainer();
  container?.addEventListener("pointerdown", onPointerDown, true);

  return toDisposable(() => {
    hoverSubscription.unsubscribe();
    pointerMove?.dispose();
    container?.removeEventListener("pointerdown", onPointerDown, true);
    clearCursor();
  });
}
