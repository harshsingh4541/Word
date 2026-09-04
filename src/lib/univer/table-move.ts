import {
  ICommandService,
  IUniverInstanceService,
  LifecycleService,
  LifecycleStages,
  UniverInstanceType,
  toDisposable,
} from "@univerjs/core";
import type { DocumentDataModel, IDisposable, Injector } from "@univerjs/core";
import { DocEventManagerService, MoveDocBlockCommand } from "@univerjs/docs-ui";
import { getTableIdAndSliceIndex, IRenderManagerService } from "@univerjs/engine-render";
import { DocSkeletonManagerService } from "@univerjs/docs";

// Word shows a four-directional move handle at the top-left corner of a
// hovered table. Clicking and dragging it repositions the entire table
// block within the document, just like the table move handle in Word.

const DEFAULT_MARGIN = { start: 10, end: 10, top: 5, bottom: 5 };

interface TableCellBound {
  rect: { left: number; right: number; top: number; bottom: number };
  pageIndex: number;
  rowIndex: number;
  colIndex: number;
  tableId: string;
}

type SkeletonLike = {
  getSkeletonData: () => {
    pages?: {
      skeTables?: Map<string, { rows: { index: number }[] }>;
      lines?: { top: number; st: number }[];
    }[];
  } | null | undefined;
};

interface SceneLike {
  getAncestorScale: () => { scaleX: number; scaleY: number };
  getViewport: (key: string) => { viewportScrollX: number; viewportScrollY: number } | null | undefined;
}

function toSourceCell(cell: TableCellBound, skeleton: SkeletonLike | undefined): TableCellBound {
  const { tableId } = getTableIdAndSliceIndex(cell.tableId);
  if (tableId === cell.tableId) return cell;
  const rows = skeleton?.getSkeletonData()?.pages?.[cell.pageIndex]?.skeTables?.get(cell.tableId)?.rows;
  return { ...cell, tableId, rowIndex: rows?.[cell.rowIndex]?.index ?? cell.rowIndex };
}

/** Convert document coordinates to container-relative screen coordinates. */
function docToContainer(
  scene: SceneLike,
  docX: number,
  docY: number,
  canvas: HTMLCanvasElement,
  container: HTMLElement,
): { x: number; y: number } {
  const { scaleX, scaleY } = scene.getAncestorScale();
  const viewport = scene.getViewport("viewMain");
  const scrollX = viewport?.viewportScrollX ?? 0;
  const scrollY = viewport?.viewportScrollY ?? 0;
  const canvasRect = canvas.getBoundingClientRect();
  const containerRect = container.getBoundingClientRect();
  return {
    x: (docX - scrollX) * scaleX + canvasRect.left - containerRect.left,
    y: (docY - scrollY) * scaleY + canvasRect.top - containerRect.top,
  };
}

/** Word-style four-arrow move handle icon. */
const MOVE_ICON_SVG = `<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 14 14" fill="currentColor">
  <path d="M7 0L5 3h1.5V7.5H2V6L0 8l2 2V8.5h3.5V13H4l2 2 2-2H6.5V8.5H10V10l2-2-2-2V7.5H6.5V3H8L7 0z"/>
</svg>`;

function createHandle(container: HTMLElement): HTMLElement {
  const handle = document.createElement("div");
  handle.title = "Move table";
  handle.style.cssText = [
    "position: absolute",
    "width: 20px",
    "height: 20px",
    "background: white",
    "border: 1px solid #aaa",
    "border-radius: 3px",
    "cursor: move",
    "display: flex",
    "align-items: center",
    "justify-content: center",
    "z-index: 50",
    "pointer-events: auto",
    "box-shadow: 0 1px 4px rgba(0,0,0,0.25)",
    "user-select: none",
    "color: #444",
  ].join("; ");
  handle.innerHTML = MOVE_ICON_SVG;
  handle.hidden = true;
  container.appendChild(handle);
  return handle;
}

function createDropLine(container: HTMLElement): HTMLElement {
  const line = document.createElement("div");
  line.style.cssText = [
    "position: absolute",
    "left: 0",
    "right: 0",
    "height: 0",
    "border-top: 2px solid #0078d4",
    "z-index: 45",
    "pointer-events: none",
  ].join("; ");
  line.hidden = true;
  container.appendChild(line);
  return line;
}

/**
 * Find the document character offset whose rendered Y position is nearest to
 * the given container-relative Y coordinate, so a table can be moved before
 * the block that lives at that position.
 */
function findDropOffset(
  skeleton: SkeletonLike | undefined,
  containerY: number,
  canvas: HTMLCanvasElement,
  container: HTMLElement,
  scene: SceneLike,
  docModel: DocumentDataModel,
): number {
  const pages = skeleton?.getSkeletonData()?.pages;
  const bodyLength = docModel.getBody()?.dataStream.length ?? 0;
  if (!pages?.length) return bodyLength;

  const { scaleY } = scene.getAncestorScale();
  const viewport = scene.getViewport("viewMain");
  const scrollY = viewport?.viewportScrollY ?? 0;
  const canvasTop = canvas.getBoundingClientRect().top - container.getBoundingClientRect().top;

  // Convert container Y → document Y
  const docY = (containerY - canvasTop) / scaleY + scrollY;

  let bestOffset = 0;
  let bestDist = Infinity;
  for (const page of pages) {
    const lines = page.lines ?? [];
    for (const line of lines) {
      const dist = Math.abs(line.top - docY);
      if (dist < bestDist) {
        bestDist = dist;
        bestOffset = line.st;
      }
    }
  }
  return bestOffset;
}

export function createTableMoveInteraction(
  injector: Injector,
  unitId: string,
  getContainer: () => HTMLElement | null,
): IDisposable {
  let inner: IDisposable | null = null;
  let disposed = false;

  void injector
    .get(LifecycleService)
    .onStage(LifecycleStages.Rendered)
    .then(() => {
      if (disposed) return;
      inner = attachTableMove(injector, unitId, getContainer);
    })
    .catch(() => {});

  return toDisposable(() => {
    disposed = true;
    inner?.dispose();
    inner = null;
  });
}

function attachTableMove(
  injector: Injector,
  unitId: string,
  getContainer: () => HTMLElement | null,
): IDisposable {
  const renderManagerService = injector.get(IRenderManagerService);
  const commandService = injector.get(ICommandService);
  const univerInstanceService = injector.get(IUniverInstanceService);
  const render = renderManagerService.getRenderUnitById(unitId);
  if (!render) return toDisposable(() => {});

  const scene = render.scene as unknown as SceneLike;
  const eventManager = render.with(DocEventManagerService);
  const skeletonManagerService = render.with(DocSkeletonManagerService);

  // Per-table top-left corner doc-coordinate, updated when (0,0) cell is seen.
  const tableTopLeft = new Map<string, { x: number; y: number }>();

  let hoveredTableId: string | null = null;
  let mouseOnHandle = false;

  let handle: HTMLElement | null = null;
  let dropLine: HTMLElement | null = null;
  let containerRef: HTMLElement | null = null;

  function getCanvas() {
    return containerRef?.querySelector<HTMLCanvasElement>("canvas") ?? null;
  }

  function ensureHandle(container: HTMLElement) {
    if (containerRef === container && handle) return;
    handle?.remove();
    dropLine?.remove();
    containerRef = container;
    handle = createHandle(container);
    dropLine = createDropLine(container);
    handle.addEventListener("pointerenter", () => { mouseOnHandle = true; });
    handle.addEventListener("pointerleave", () => { mouseOnHandle = false; });
    handle.addEventListener("pointerdown", onHandlePointerDown);
  }

  function positionHandle() {
    if (!handle || !containerRef || !hoveredTableId) return;
    const pos = tableTopLeft.get(hoveredTableId);
    if (!pos) return;
    const canvas = getCanvas();
    if (!canvas) return;
    const { x, y } = docToContainer(scene, pos.x, pos.y, canvas, containerRef);
    // Clamp to at least 4px from container edges so the handle stays reachable
    // even when the table is near the very top or left edge of the page.
    handle.style.left = `${Math.max(4, x - 20)}px`;
    handle.style.top = `${Math.max(4, y - 20)}px`;
    handle.hidden = false;
  }

  const hoverSub = eventManager.hoverTableCellRealTime$.subscribe((cell) => {
    if (!cell) {
      if (!mouseOnHandle) {
        hoveredTableId = null;
        if (handle) handle.hidden = true;
      }
      return;
    }

    const skeleton = skeletonManagerService?.getSkeleton() as SkeletonLike | undefined;
    const resolved = toSourceCell(cell, skeleton);

    const container = getContainer();
    if (container) ensureHandle(container);

    hoveredTableId = resolved.tableId;

    // Record the table's top-left corner when we see the (0, 0) cell.
    if (resolved.rowIndex === 0 && resolved.colIndex === 0) {
      tableTopLeft.set(resolved.tableId, {
        x: resolved.rect.left - DEFAULT_MARGIN.start,
        y: resolved.rect.top - DEFAULT_MARGIN.top,
      });
    }

    positionHandle();
  });

  function onHandlePointerDown(event: PointerEvent) {
    if (!hoveredTableId || !containerRef) return;
    event.preventDefault();
    event.stopPropagation();

    const capturedTableId = hoveredTableId;

    const onMove = (moveEvent: PointerEvent) => {
      if (!dropLine || !containerRef) return;
      const y = moveEvent.clientY - containerRef.getBoundingClientRect().top;
      dropLine.style.top = `${y}px`;
      dropLine.hidden = false;
    };

    const onUp = (upEvent: PointerEvent) => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      if (dropLine) dropLine.hidden = true;
      if (!containerRef) return;

      const docModel = univerInstanceService.getCurrentUnitOfType<DocumentDataModel>(UniverInstanceType.UNIVER_DOC);
      if (!docModel) return;

      // Find the table's source range from the body.
      const tableEntry = docModel.getBody()?.tables?.find((t) => t.tableId === capturedTableId);
      if (!tableEntry) return;

      const canvas = getCanvas();
      if (!canvas) return;

      const skeleton = skeletonManagerService?.getSkeleton() as SkeletonLike | undefined;
      const containerY = upEvent.clientY - containerRef.getBoundingClientRect().top;
      const targetOffset = findDropOffset(skeleton, containerY, canvas, containerRef, scene, docModel);

      // Do not move onto itself.
      if (targetOffset >= tableEntry.startIndex && targetOffset <= tableEntry.endIndex) return;

      void commandService.executeCommand(MoveDocBlockCommand.id, {
        unitId,
        sourceRange: { startOffset: tableEntry.startIndex, endOffset: tableEntry.endIndex },
        targetOffset,
      });
    };

    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
  }

  return toDisposable(() => {
    hoverSub.unsubscribe();
    handle?.remove();
    dropLine?.remove();
    handle = null;
    dropLine = null;
    hoveredTableId = null;
    containerRef = null;
    tableTopLeft.clear();
  });
}
