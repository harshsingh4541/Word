import type {
  DocumentDataModel,
  IAccessor,
  Nullable,
  IColorStyle,
  ICommand,
  IMutationInfo,
  JSONXActions,
} from "@univerjs/core";
import type { IRichTextEditingMutationParams } from "@univerjs/docs";
import {
  BooleanNumber,
  CommandType,
  DashStyleType,
  HorizontalAlign,
  ICommandService,
  IUniverInstanceService,
  JSONX,
  TableRowHeightRule,
  TableLayoutType,
  TableSizeType,
  TextX,
  TextXActionType,
  UniverInstanceType,
  VerticalAlignmentType,
} from "@univerjs/core";
import { DocSelectionManagerService, DocSkeletonManagerService, RichTextEditingMutation } from "@univerjs/docs";
import { getTableIdAndSliceIndex, IRenderManagerService } from "@univerjs/engine-render";
import { getCommandSkeleton } from "@univerjs/docs-ui";

// These commands fill a real gap in Univer's open-source Docs table plugin:
// the data model (ITableCell.backgroundColor/borderTop.../vAlign, ITableRow.trHeight,
// ITable.layout) already supports all of this and the renderer already paints it —
// there just aren't any built-in commands or toolbar buttons to set them yet.
// Every command below only ever writes plain properties via JSONX ops (no dataStream
// edits), which keeps them in the same low-risk category as Univer's own
// column-width-on-insert logic. Cell merge/split needs dataStream edits too
// (removing table cells changes the document's text stream) and is intentionally
// not included here — that's real-but-separate follow-up work.

// Univer's dataStream table tokens, by character code so no control
// characters end up pasted into the source.
const TABLE_ROW_START_TOKEN = String.fromCharCode(0x1b);
const TABLE_CELL_START_TOKEN = String.fromCharCode(0x1c);

export type SelectedTableRange = {
  tableId: string;
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
};

// A caret merely positioned inside one cell (no drag-selected block) never
// produces a rectRange — confirmed in 1.0.0-beta.2, where a plain click
// into a cell yields getRectRanges() === []. Which cell it is has to come
// from the document itself: the selection's skeleton path names a *slice*
// of the table, not the table (a table that spills onto a second page is
// laid out as several skeleton tables, `${tableId}#-#0`, `#-#1`, ... each
// numbering its rows from zero again), so reading row and column off the
// path silently addresses the wrong cell — or no cell at all, since the
// slice id is not a key in `tableSource`.
//
// Counting the model's own row/cell tokens up to the caret has neither
// problem: the tokens are the table, whole, however it happens to be
// paginated.
function findCellAtOffset(
  docDataModel: DocumentDataModel,
  offset: number,
): { tableId: string; row: number; column: number } | null {
  const body = docDataModel.getBody();
  const table = body?.tables?.find((t) => offset > t.startIndex && offset < t.endIndex);
  if (!body || !table) return null;

  const { dataStream } = body;
  let row = -1;
  let column = -1;
  for (let i = table.startIndex; i < table.endIndex && i < offset; i++) {
    const char = dataStream[i];
    if (char === TABLE_ROW_START_TOKEN) {
      row++;
      column = -1;
    } else if (char === TABLE_CELL_START_TOKEN) {
      column++;
    }
  }
  if (row < 0 || column < 0) return null;

  return { tableId: table.tableId, row, column };
}

// The single source of truth for "what table cell(s) is the user's live
// selection touching right now" — used both as the command-side fallback
// below and by DocsEditor's selection subscription that drives the Table
// Design ribbon's visibility.
export function resolveLiveTableRange(
  docSelectionManagerService: DocSelectionManagerService,
  docDataModel: Nullable<DocumentDataModel>,
): SelectedTableRange | null {
  const rectRanges = docSelectionManagerService.getRectRanges();
  const rectRange = rectRanges?.find((r) => r.tableId);
  if (rectRange) {
    // A rect range's rows and columns are already absolute, but its
    // tableId is the slice's.
    return remember({
      tableId: getTableIdAndSliceIndex(rectRange.tableId).tableId,
      startRow: Math.min(rectRange.startRow, rectRange.endRow),
      endRow: Math.max(rectRange.startRow, rectRange.endRow),
      startColumn: Math.min(rectRange.startColumn, rectRange.endColumn),
      endColumn: Math.max(rectRange.startColumn, rectRange.endColumn),
    });
  }

  const offset = docSelectionManagerService.getDocRanges()?.[0]?.startOffset;
  const cell = docDataModel && offset != null ? findCellAtOffset(docDataModel, offset) : null;
  if (!cell) return null;

  return remember({
    tableId: cell.tableId,
    startRow: cell.row,
    endRow: cell.row,
    startColumn: cell.column,
    endColumn: cell.column,
  });
}

// Opening a ribbon dropdown (a colour picker, a row-height menu) moves DOM
// focus off the canvas, and Univer clears its live selection when that
// happens — so by the time the click that applies the command lands, the
// live lookup above can already return null. Remembering the last real
// table selection keeps those commands working; the Table Design tab is
// only on screen while a table selection exists in the first place.
let rememberedRange: SelectedTableRange | null = null;

function remember(range: SelectedTableRange): SelectedTableRange {
  rememberedRange = range;
  return range;
}

/** Called when an editor is torn down so no range leaks into the next one. */
export function clearRememberedTableRange() {
  rememberedRange = null;
}

// Focusing any input outside the canvas (typing a border width, a row
// height...) clears Univer's live selection before the button click that
// applies it ever fires. Every command below therefore accepts an explicit
// `range` in its params — captured by the panel from a selection-change
// subscription while the selection was still live — and only falls back to
// querying live selection when one isn't provided.
function resolveTableRange(accessor: IAccessor, explicit?: SelectedTableRange | null): SelectedTableRange | null {
  if (explicit) return explicit;
  const docSelectionManagerService = accessor.get(DocSelectionManagerService);
  const docDataModel = accessor
    .get(IUniverInstanceService)
    .getCurrentUnitOfType<DocumentDataModel>(UniverInstanceType.UNIVER_DOC);
  return resolveLiveTableRange(docSelectionManagerService, docDataModel) ?? rememberedRange;
}

function getDocAndTable(accessor: IAccessor, tableId: string) {
  const univerInstanceService = accessor.get(IUniverInstanceService);
  const docDataModel = univerInstanceService.getCurrentUnitOfType<DocumentDataModel>(
    UniverInstanceType.UNIVER_DOC,
  );
  if (!docDataModel) return null;

  const table = docDataModel.getSnapshot().tableSource?.[tableId];
  if (!table) return null;

  return { docDataModel, table };
}

function setProperty(
  jsonX: ReturnType<typeof JSONX.getInstance>,
  rawActions: JSONXActions[],
  path: (string | number)[],
  oldVal: unknown,
  newVal: unknown,
) {
  if (oldVal === undefined && newVal === undefined) return;

  const op =
    oldVal === undefined
      ? jsonX.insertOp(path, newVal)
      : newVal === undefined
        ? jsonX.removeOp(path, oldVal)
        : jsonX.replaceOp(path, oldVal, newVal);

  if (op) rawActions.push(op as JSONXActions);
}

// The RETAIN edit above is enough to trigger a repaint (proven for border/
// background), but properties that affect box dimensions — row height, in
// particular — are cached in the skeleton at layout time and a repaint
// alone reads the stale cached height. Force a real layout recalculation.
function forceRelayout(accessor: IAccessor, unitId: string) {
  const renderManagerService = accessor.get(IRenderManagerService);
  const render = renderManagerService.getRenderUnitById(unitId);
  const skeletonManagerService = render?.with(DocSkeletonManagerService);
  const skeleton = skeletonManagerService?.getSkeleton();
  skeleton?.makeDirty(true);
  skeleton?.calculate();
  // Recalculating the skeleton alone updates cached layout data but doesn't
  // repaint — the canvas still shows the old frame until the scene is told
  // to redraw. This last line is what actually made row-height changes
  // (and any other layout-affecting property) visible on screen.
  render?.scene.makeDirty(true);
  render?.mainComponent?.makeDirty(true);
  void render?.scene.requestRender();
}

function runMutation(
  accessor: IAccessor,
  docDataModel: DocumentDataModel,
  rawActions: JSONXActions[],
): boolean {
  if (!rawActions.length) return false;

  const commandService = accessor.get(ICommandService);
  const docSelectionManagerService = accessor.get(DocSelectionManagerService);
  const activeTextRange = docSelectionManagerService.getActiveTextRange();

  // Property-only JSONX ops (no textX component) don't reliably trigger a
  // skeleton rebuild — Univer's own table commands always bundle a textX
  // edit alongside their JSONX ops for exactly this reason. A full-length
  // RETAIN is a no-op edit that forces the same re-layout/repaint pass.
  const body = docDataModel.getBody();
  const bodyLength = body?.dataStream.length ?? 0;
  const jsonX = JSONX.getInstance();
  const allActions = [...rawActions];

  if (bodyLength > 0) {
    const textX = new TextX();
    textX.push({ t: TextXActionType.RETAIN, len: bodyLength });
    // Equivalent to core's internal getRichTextEditPath(docDataModel, '') —
    // not part of core's public export surface, so inlined here.
    const editOp = jsonX.editOp(textX.serialize(), ["body"]);
    if (editOp) allActions.push(editOp as JSONXActions);
  }

  const actions = allActions.reduce(
    (acc, cur) => JSONX.compose(acc, cur),
    null as JSONXActions,
  );

  const doMutation: IMutationInfo<IRichTextEditingMutationParams> = {
    id: RichTextEditingMutation.id,
    params: {
      unitId: docDataModel.getUnitId(),
      actions,
      textRanges: activeTextRange ? [activeTextRange] : [],
    },
  };

  const result = commandService.syncExecuteCommand(doMutation.id, doMutation.params);
  if (result) forceRelayout(accessor, docDataModel.getUnitId());
  return Boolean(result);
}

// ---------------------------------------------------------------------------
// Cell background color
// ---------------------------------------------------------------------------

export interface ISetTableCellBackgroundParams {
  color: string | null;
  range?: SelectedTableRange | null;
}

export const SetTableCellBackgroundCommandId = "dockaro.command.table-cell-background";

export const SetTableCellBackgroundCommand: ICommand<ISetTableCellBackgroundParams> = {
  id: SetTableCellBackgroundCommandId,
  type: CommandType.COMMAND,
  handler: (accessor, params) => {
    if (!params) return false;
    const range = resolveTableRange(accessor, params.range);
    if (!range) return false;
    const found = getDocAndTable(accessor, range.tableId);
    if (!found) return false;
    const { docDataModel, table } = found;

    const jsonX = JSONX.getInstance();
    const rawActions: JSONXActions[] = [];

    for (let r = range.startRow; r <= range.endRow; r++) {
      const row = table.tableRows[r];
      if (!row) continue;
      for (let c = range.startColumn; c <= range.endColumn; c++) {
        const cell = row.tableCells[c];
        if (!cell) continue;
        const newVal: IColorStyle | undefined = params.color ? { rgb: params.color } : undefined;
        setProperty(
          jsonX,
          rawActions,
          ["tableSource", range.tableId, "tableRows", r, "tableCells", c, "backgroundColor"],
          cell.backgroundColor,
          newVal,
        );
      }
    }

    return runMutation(accessor, docDataModel, rawActions);
  },
};

// ---------------------------------------------------------------------------
// Cell border (per side, color + width, or clear)
// ---------------------------------------------------------------------------

export type BorderSide = "Top" | "Bottom" | "Left" | "Right";

export interface ISetTableCellBorderParams {
  sides: BorderSide[];
  color: string | null;
  width: number;
  /** Word's "Line Style": solid, dotted or dashed. Solid when omitted. */
  dashStyle?: DashStyleType;
  range?: SelectedTableRange | null;
}

export const SetTableCellBorderCommandId = "dockaro.command.table-cell-border";

export const SetTableCellBorderCommand: ICommand<ISetTableCellBorderParams> = {
  id: SetTableCellBorderCommandId,
  type: CommandType.COMMAND,
  handler: (accessor, params) => {
    if (!params || params.sides.length === 0) return false;
    const range = resolveTableRange(accessor, params.range);
    if (!range) return false;
    const found = getDocAndTable(accessor, range.tableId);
    if (!found) return false;
    const { docDataModel, table } = found;

    const jsonX = JSONX.getInstance();
    const rawActions: JSONXActions[] = [];

    for (let r = range.startRow; r <= range.endRow; r++) {
      const row = table.tableRows[r];
      if (!row) continue;
      for (let c = range.startColumn; c <= range.endColumn; c++) {
        const cell = row.tableCells[c];
        if (!cell) continue;

        for (const side of params.sides) {
          const key = `border${side}` as const;
          // Word's "No border" has to leave a border behind: Univer's
          // renderer falls back to a default grey line whenever a cell has
          // no border property at all, and only skips drawing when the
          // border it finds is zero-width or transparent. Deleting the
          // property would put the default line back.
          const newVal = params.color
            ? {
                color: { rgb: params.color },
                width: { v: params.width },
                dashStyle: params.dashStyle ?? DashStyleType.SOLID,
              }
            : { color: { rgb: "transparent" }, width: { v: 0 }, dashStyle: DashStyleType.SOLID };

          setProperty(
            jsonX,
            rawActions,
            ["tableSource", range.tableId, "tableRows", r, "tableCells", c, key],
            (cell as Record<string, unknown>)[key],
            newVal,
          );
        }
      }
    }

    return runMutation(accessor, docDataModel, rawActions);
  },
};

// ---------------------------------------------------------------------------
// View gridlines
// ---------------------------------------------------------------------------

// Word's View Gridlines draws faint dashed lines wherever a table has no
// border, so a borderless table is still editable instead of invisible.
// Univer has no non-printing overlay layer to hang real gridlines off — the
// renderer only paints borders it finds on cells — so these are ordinary
// borders in a style reserved for the purpose.
//
// Reserving the style is also what removes them: any side still carrying
// exactly this style is a gridline, so toggling off needs no record of what
// was painted, and undo/redo can't desync that record. The trade-off is
// that a border set to exactly this style by hand reads as a gridline.
// Word can't express that case at all, since its gridlines aren't borders.
const GRIDLINE_COLOR = "#D0D0D0";
const GRIDLINE_WIDTH = 1;

const GRIDLINE_BORDER = {
  color: { rgb: GRIDLINE_COLOR },
  width: { v: GRIDLINE_WIDTH },
  dashStyle: DashStyleType.DASH,
};

// Same shape SetTableCellBorderCommand clears to, and for the same reason:
// a cell with no border property at all gets the renderer's default grey
// line back, so "off" has to be an explicit invisible border.
const CLEARED_BORDER = {
  color: { rgb: "transparent" },
  width: { v: 0 },
  dashStyle: DashStyleType.SOLID,
};

const GRIDLINE_SIDES: BorderSide[] = ["Top", "Bottom", "Left", "Right"];

type StoredBorder = { color?: IColorStyle; width?: { v?: number }; dashStyle?: DashStyleType };

/** Nothing visible on this side: absent, transparent, or zero-width. */
function isBorderInvisible(border: unknown): boolean {
  if (!border) return true;
  const { color, width } = border as StoredBorder;
  const rgb = color?.rgb;
  if (!rgb || rgb.toLowerCase() === "transparent") return true;
  return !width?.v;
}

function isGridline(border: unknown): boolean {
  if (!border) return false;
  const { color, width, dashStyle } = border as StoredBorder;
  return (
    (color?.rgb ?? "").toLowerCase() === GRIDLINE_COLOR.toLowerCase() &&
    width?.v === GRIDLINE_WIDTH &&
    dashStyle === DashStyleType.DASH
  );
}

export interface IToggleTableGridlinesParams {
  visible: boolean;
  range?: SelectedTableRange | null;
}

export const ToggleTableGridlinesCommandId = "dockaro.command.table-gridlines";

export const ToggleTableGridlinesCommand: ICommand<IToggleTableGridlinesParams> = {
  id: ToggleTableGridlinesCommandId,
  type: CommandType.COMMAND,
  handler: (accessor, params) => {
    if (!params) return false;
    const range = resolveTableRange(accessor, params.range);
    if (!range) return false;
    const found = getDocAndTable(accessor, range.tableId);
    if (!found) return false;
    const { docDataModel, table } = found;

    const jsonX = JSONX.getInstance();
    const rawActions: JSONXActions[] = [];

    // Gridlines cover the whole table, not the selected cells: this is a
    // view toggle, not a formatting brush.
    table.tableRows.forEach((row, r) => {
      row.tableCells.forEach((cell, c) => {
        for (const side of GRIDLINE_SIDES) {
          const key = `border${side}` as const;
          const current = (cell as Record<string, unknown>)[key];

          // Showing never touches a real border, and hiding only touches
          // the reserved style — anything else is the user's formatting.
          if (params.visible ? !isBorderInvisible(current) : !isGridline(current)) continue;

          setProperty(
            jsonX,
            rawActions,
            ["tableSource", range.tableId, "tableRows", r, "tableCells", c, key],
            current,
            params.visible ? GRIDLINE_BORDER : CLEARED_BORDER,
          );
        }
      });
    });

    return runMutation(accessor, docDataModel, rawActions);
  },
};

// ---------------------------------------------------------------------------
// Cell vertical alignment
// ---------------------------------------------------------------------------

export interface ISetTableCellVAlignParams {
  vAlign: VerticalAlignmentType;
  range?: SelectedTableRange | null;
}

export const SetTableCellVAlignCommandId = "dockaro.command.table-cell-valign";

export const SetTableCellVAlignCommand: ICommand<ISetTableCellVAlignParams> = {
  id: SetTableCellVAlignCommandId,
  type: CommandType.COMMAND,
  handler: (accessor, params) => {
    if (!params) return false;
    const range = resolveTableRange(accessor, params.range);
    if (!range) return false;
    const found = getDocAndTable(accessor, range.tableId);
    if (!found) return false;
    const { docDataModel, table } = found;

    const jsonX = JSONX.getInstance();
    const rawActions: JSONXActions[] = [];

    for (let r = range.startRow; r <= range.endRow; r++) {
      const row = table.tableRows[r];
      if (!row) continue;
      for (let c = range.startColumn; c <= range.endColumn; c++) {
        const cell = row.tableCells[c];
        if (!cell) continue;
        setProperty(
          jsonX,
          rawActions,
          ["tableSource", range.tableId, "tableRows", r, "tableCells", c, "vAlign"],
          cell.vAlign,
          params.vAlign,
        );
      }
    }

    return runMutation(accessor, docDataModel, rawActions);
  },
};

// ---------------------------------------------------------------------------
// Cell margins (Word's Table Layout > Cell Margins)
// ---------------------------------------------------------------------------
//
// The padding inside every cell, set for the table as a whole - which is
// how Word's own Table Options dialog scopes it. Univer's model already
// carries it (ITable.cellMargin) and the renderer already reads it; there
// was just no way to change it.

export interface ISetTableCellMarginParams {
  margin: { start: number; end: number; top: number; bottom: number };
  range?: SelectedTableRange | null;
}

export const SetTableCellMarginCommandId = "dockaro.command.table-cell-margin";

export const SetTableCellMarginCommand: ICommand<ISetTableCellMarginParams> = {
  id: SetTableCellMarginCommandId,
  type: CommandType.COMMAND,
  handler: (accessor, params) => {
    if (!params) return false;
    const range = resolveTableRange(accessor, params.range);
    if (!range) return false;
    const found = getDocAndTable(accessor, range.tableId);
    if (!found) return false;
    const { docDataModel, table } = found;

    const { start, end, top, bottom } = params.margin;
    const jsonX = JSONX.getInstance();
    const rawActions: JSONXActions[] = [];
    setProperty(jsonX, rawActions, ["tableSource", range.tableId, "cellMargin"], table.cellMargin, {
      start: { v: start },
      end: { v: end },
      top: { v: top },
      bottom: { v: bottom },
    });

    // Per-cell overrides would win over the table default, so Univer's own
    // per-cell margins (written when a table is created) are cleared too.
    table.tableRows.forEach((row, r) => {
      row.tableCells.forEach((cell, c) => {
        if (!cell.margin) return;
        setProperty(
          jsonX,
          rawActions,
          ["tableSource", range.tableId, "tableRows", r, "tableCells", c, "margin"],
          cell.margin,
          undefined,
        );
      });
    });

    if (!runMutation(accessor, docDataModel, rawActions)) return false;
    forceRelayout(accessor, docDataModel.getUnitId());
    return true;
  },
};

// ---------------------------------------------------------------------------
// Row height
// ---------------------------------------------------------------------------

/** Univer's default cell padding, used when a table sets none of its own. */
const DEFAULT_ROW_CELL_MARGIN = { start: 10, end: 10, top: 5, bottom: 5 };

/**
 * A conservative height for one line of text at the editor's default font
 * size, used to decide how much vertical padding a requested row height can
 * still afford. Deliberately an estimate: it only decides how much padding to
 * trim, and the AT_LEAST rule below guarantees the row still grows to fit its
 * real text — so guessing low costs a little padding and guessing high never
 * clips anything.
 */
const APPROX_LINE_HEIGHT = 20;

export interface ISetTableRowHeightParams {
  /**
   * `"fixed"` is Word's "At least": the typed height is a floor, so the row
   * never crops its text. Reaching a *smaller* height is handled by trimming
   * the cells' vertical padding instead — see the command below.
   */
  mode: "auto" | "fixed";
  height?: number;
  range?: SelectedTableRange | null;
}

export const SetTableRowHeightCommandId = "dockaro.command.table-row-height";

export const SetTableRowHeightCommand: ICommand<ISetTableRowHeightParams> = {
  id: SetTableRowHeightCommandId,
  type: CommandType.COMMAND,
  handler: (accessor, params) => {
    if (!params) return false;
    const range = resolveTableRange(accessor, params.range);
    if (!range) return false;
    const found = getDocAndTable(accessor, range.tableId);
    if (!found) return false;
    const { docDataModel, table } = found;

    const jsonX = JSONX.getInstance();
    const rawActions: JSONXActions[] = [];

    // AT_LEAST, never EXACT. The renderer derives a row's content height as
    // `cell.height + marginTop + marginBottom` and then applies the rule:
    // AT_LEAST is `Math.max(contentHeight, val.v)`, so the row always fits its
    // text, while EXACT is a flat `rowHeight = val.v` that crops whatever
    // doesn't fit — EXACT is why text was being cut off.
    //
    // AT_LEAST alone can't shrink a row, though: the vertical cell padding is
    // part of contentHeight, so with the default 5px top + 5px bottom a
    // single-line row never drops below roughly 30px however small a height is
    // asked for — which is why decreasing appeared to do nothing.
    //
    // So the padding is trimmed to fit the requested height as well. That
    // lowers contentHeight, letting AT_LEAST actually reach a shorter row,
    // and because the rule is still AT_LEAST the text is never clipped.
    const isAuto = params.mode === "auto";
    const requestedHeight = Math.max(MIN_ROW_HEIGHT, params.height ?? 30);
    const newTrHeight = isAuto
      ? { val: { v: 30 }, hRule: TableRowHeightRule.AUTO }
      : { val: { v: requestedHeight }, hRule: TableRowHeightRule.AT_LEAST };

    // What the requested height can spare for padding once a line of text has
    // taken its share. Padding is only ever reduced, so growing a row leaves
    // the user's own cell margins alone.
    const paddingRoom = Math.max(0, Math.floor((requestedHeight - APPROX_LINE_HEIGHT) / 2));

    for (let r = range.startRow; r <= range.endRow; r++) {
      const row = table.tableRows[r];
      if (!row) continue;

      setProperty(
        jsonX,
        rawActions,
        ["tableSource", range.tableId, "tableRows", r, "trHeight"],
        row.trHeight,
        newTrHeight,
      );

      // AutoFit means "grow with the content", so it keeps the full padding.
      if (isAuto) continue;

      row.tableCells.forEach((cell, c) => {
        const current = cell.margin ?? table.cellMargin;
        const start = current?.start?.v ?? DEFAULT_ROW_CELL_MARGIN.start;
        const end = current?.end?.v ?? DEFAULT_ROW_CELL_MARGIN.end;
        const top = current?.top?.v ?? DEFAULT_ROW_CELL_MARGIN.top;
        const bottom = current?.bottom?.v ?? DEFAULT_ROW_CELL_MARGIN.bottom;
        const newTop = Math.min(top, paddingRoom);
        const newBottom = Math.min(bottom, paddingRoom);
        if (newTop === top && newBottom === bottom) return;

        setProperty(
          jsonX,
          rawActions,
          ["tableSource", range.tableId, "tableRows", r, "tableCells", c, "margin"],
          cell.margin,
          { start: { v: start }, end: { v: end }, top: { v: newTop }, bottom: { v: newBottom } },
        );
      });
    }

    return runMutation(accessor, docDataModel, rawActions);
  },
};

// ---------------------------------------------------------------------------
// Column width — Univer's Docs table has no interactive drag-to-resize
// (confirmed: no such command exists anywhere in docs-ui, only an internal
// auto-fit-to-page-width helper). This is the reliable substitute: type an
// exact width for the selected column(s) instead of dragging a border.
// ---------------------------------------------------------------------------

export interface ISetTableColumnWidthParams {
  width: number;
  range?: SelectedTableRange | null;
}

export const SetTableColumnWidthCommandId = "dockaro.command.table-column-width";

export const SetTableColumnWidthCommand: ICommand<ISetTableColumnWidthParams> = {
  id: SetTableColumnWidthCommandId,
  type: CommandType.COMMAND,
  handler: (accessor, params) => {
    if (!params || params.width <= 0) return false;
    const range = resolveTableRange(accessor, params.range);
    if (!range) return false;
    const found = getDocAndTable(accessor, range.tableId);
    if (!found) return false;
    const { docDataModel, table } = found;

    const jsonX = JSONX.getInstance();
    const rawActions: JSONXActions[] = [];

    for (let c = range.startColumn; c <= range.endColumn; c++) {
      const column = table.tableColumns[c];
      if (!column) continue;
      setProperty(
        jsonX,
        rawActions,
        ["tableSource", range.tableId, "tableColumns", c, "size", "width", "v"],
        column.size.width.v,
        params.width,
      );
    }

    return runMutation(accessor, docDataModel, rawActions);
  },
};

// ---------------------------------------------------------------------------
// Drag-resizing a column border, the way Word does it
// ---------------------------------------------------------------------------
//
// Dragging a border in Word moves that border: the column on its left takes
// the delta and the column on its right gives it back, so the table keeps
// its overall width. Dragging the table's own right edge widens the table
// instead. Both are one command so a drag is a single undo step.

/** Word refuses to shrink a column below roughly this width. */
const MIN_COLUMN_WIDTH = 24;

export interface IResizeTableColumnParams {
  tableId: string;
  /** Index of the column on the left of the dragged border. */
  columnIndex: number;
  /** Movement in document pixels; positive widens the left column. */
  delta: number;
}

export const ResizeTableColumnCommandId = "dockaro.command.table-resize-column";

export const ResizeTableColumnCommand: ICommand<IResizeTableColumnParams> = {
  id: ResizeTableColumnCommandId,
  type: CommandType.COMMAND,
  handler: (accessor, params) => {
    if (!params || !params.delta) return false;
    const found = getDocAndTable(accessor, params.tableId);
    if (!found) return false;
    const { docDataModel, table } = found;

    const left = table.tableColumns[params.columnIndex];
    if (!left) return false;
    const right = table.tableColumns[params.columnIndex + 1];
    const leftWidth = left.size?.width?.v ?? 0;

    const jsonX = JSONX.getInstance();
    const rawActions: JSONXActions[] = [];
    const setColumnWidth = (index: number, from: number, to: number) => {
      setProperty(
        jsonX,
        rawActions,
        ["tableSource", params.tableId, "tableColumns", index, "size", "width", "v"],
        from,
        to,
      );
    };

    if (right) {
      const rightWidth = right.size?.width?.v ?? 0;
      // Clamp the drag so neither side collapses; the border simply stops.
      const delta = Math.max(
        MIN_COLUMN_WIDTH - leftWidth,
        Math.min(rightWidth - MIN_COLUMN_WIDTH, params.delta),
      );
      if (!delta) return false;
      setColumnWidth(params.columnIndex, leftWidth, leftWidth + delta);
      setColumnWidth(params.columnIndex + 1, rightWidth, rightWidth - delta);
    } else {
      const delta = Math.max(MIN_COLUMN_WIDTH - leftWidth, params.delta);
      if (!delta) return false;
      setColumnWidth(params.columnIndex, leftWidth, leftWidth + delta);
      const tableWidth = table.size?.width?.v;
      if (typeof tableWidth === "number") {
        setProperty(
          jsonX,
          rawActions,
          ["tableSource", params.tableId, "size", "width", "v"],
          tableWidth,
          tableWidth + delta,
        );
      }
    }

    return runMutation(accessor, docDataModel, rawActions);
  },
};

// ---------------------------------------------------------------------------
// Drag-resizing a row border
// ---------------------------------------------------------------------------

/** Word's minimum row height, in document pixels. */
const MIN_ROW_HEIGHT = 16;

export interface IResizeTableRowParams {
  tableId: string;
  rowIndex: number;
  /** The row's new height in document pixels. */
  height: number;
}

export const ResizeTableRowCommandId = "dockaro.command.table-resize-row";

export const ResizeTableRowCommand: ICommand<IResizeTableRowParams> = {
  id: ResizeTableRowCommandId,
  type: CommandType.COMMAND,
  handler: (accessor, params) => {
    if (!params) return false;
    const found = getDocAndTable(accessor, params.tableId);
    if (!found) return false;
    const { docDataModel, table } = found;
    const row = table.tableRows[params.rowIndex];
    if (!row) return false;

    const jsonX = JSONX.getInstance();
    const rawActions: JSONXActions[] = [];
    setProperty(
      jsonX,
      rawActions,
      ["tableSource", params.tableId, "tableRows", params.rowIndex, "trHeight"],
      row.trHeight,
      // Dragging a row border in Word sets a minimum height, not a fixed
      // one: the row still grows if its content needs more space.
      { val: { v: Math.max(MIN_ROW_HEIGHT, params.height) }, hRule: TableRowHeightRule.AT_LEAST },
    );

    return runMutation(accessor, docDataModel, rawActions);
  },
};

// ---------------------------------------------------------------------------
// Table layout (auto-fit vs fixed column widths)
// ---------------------------------------------------------------------------

export interface ISetTableLayoutParams {
  layout: "auto" | "fixed";
  range?: SelectedTableRange | null;
}

export const SetTableLayoutCommandId = "dockaro.command.table-layout";

export const SetTableLayoutCommand: ICommand<ISetTableLayoutParams> = {
  id: SetTableLayoutCommandId,
  type: CommandType.COMMAND,
  handler: (accessor, params) => {
    if (!params) return false;
    const range = resolveTableRange(accessor, params.range);
    if (!range) return false;
    const found = getDocAndTable(accessor, range.tableId);
    if (!found) return false;
    const { docDataModel, table } = found;

    const jsonX = JSONX.getInstance();
    const rawActions: JSONXActions[] = [];

    const newVal = params.layout === "auto" ? TableLayoutType.AUTO_FIT : TableLayoutType.FIXED;

    setProperty(
      jsonX,
      rawActions,
      ["tableSource", range.tableId, "layout"],
      table.layout,
      newVal,
    );

    return runMutation(accessor, docDataModel, rawActions);
  },
};

// ---------------------------------------------------------------------------
// Banded (striped) rows — whole table, two alternating colors
// ---------------------------------------------------------------------------

export interface ISetTableBandedRowsParams {
  enabled: boolean;
  colorOdd?: string;
  colorEven?: string;
  range?: SelectedTableRange | null;
}

export const SetTableBandedRowsCommandId = "dockaro.command.table-banded-rows";

export const SetTableBandedRowsCommand: ICommand<ISetTableBandedRowsParams> = {
  id: SetTableBandedRowsCommandId,
  type: CommandType.COMMAND,
  handler: (accessor, params) => {
    if (!params) return false;
    const range = resolveTableRange(accessor, params.range);
    if (!range) return false;
    const found = getDocAndTable(accessor, range.tableId);
    if (!found) return false;
    const { docDataModel, table } = found;

    const jsonX = JSONX.getInstance();
    const rawActions: JSONXActions[] = [];

    table.tableRows.forEach((row, r) => {
      const isOdd = r % 2 === 0;
      const color = !params.enabled
        ? undefined
        : isOdd
          ? (params.colorOdd ?? "#F7F7F8")
          : (params.colorEven ?? "#FFFFFF");
      const newVal: IColorStyle | undefined = color ? { rgb: color } : undefined;

      row.tableCells.forEach((cell, c) => {
        setProperty(
          jsonX,
          rawActions,
          ["tableSource", range.tableId, "tableRows", r, "tableCells", c, "backgroundColor"],
          cell.backgroundColor,
          newVal,
        );
      });
    });

    return runMutation(accessor, docDataModel, rawActions);
  },
};

// ---------------------------------------------------------------------------
// Repeat header row
// ---------------------------------------------------------------------------

export interface ISetTableHeaderRowParams {
  enabled: boolean;
  range?: SelectedTableRange | null;
}

export const SetTableHeaderRowCommandId = "dockaro.command.table-header-row";

export const SetTableHeaderRowCommand: ICommand<ISetTableHeaderRowParams> = {
  id: SetTableHeaderRowCommandId,
  type: CommandType.COMMAND,
  handler: (accessor, params) => {
    if (!params) return false;
    const range = resolveTableRange(accessor, params.range);
    if (!range) return false;
    const found = getDocAndTable(accessor, range.tableId);
    if (!found) return false;
    const { docDataModel, table } = found;

    const firstRow = table.tableRows[0];
    if (!firstRow) return false;

    const jsonX = JSONX.getInstance();
    const rawActions: JSONXActions[] = [];

    setProperty(
      jsonX,
      rawActions,
      ["tableSource", range.tableId, "tableRows", 0, "repeatHeaderRow"],
      firstRow.repeatHeaderRow,
      params.enabled ? BooleanNumber.TRUE : undefined,
    );
    setProperty(
      jsonX,
      rawActions,
      ["tableSource", range.tableId, "tableRows", 0, "isFirstRow"],
      firstRow.isFirstRow,
      params.enabled ? BooleanNumber.TRUE : undefined,
    );

    return runMutation(accessor, docDataModel, rawActions);
  },
};

// ---------------------------------------------------------------------------
// Merge cells (horizontal, same row only)
// ---------------------------------------------------------------------------
//
// NOTE: this command is data-correct but deliberately not wired into
// TableRibbon's UI right now. Root-caused against Univer's own layout
// engine source (engine-render/src/components/docs/layout/block/table.ts):
// a merge is represented by leaving the absorbed cell IN PLACE with
// columnSpan set to 0 (isCoveredTableCell checks `columnSpan === 0`) while
// the anchor cell's columnSpan is set to the spanned count — confirmed
// correct, and confirmed to update the persisted snapshot correctly. BUT
// the only span-aware sizing function in that file, applyMergedCellSpanHeights,
// computes height for rowSpan and has no width-equivalent for columnSpan —
// so the anchor cell's box never actually widens on screen. The covered
// cell's content/background/border correctly stop painting
// (_drawTable/_drawTableCellBackgrounds in engine-render's document.ts both
// check isMergedCellCovered), it just leaves a visually blank gap instead
// of a wider cell. This is a genuine gap in this Univer version's table
// renderer, not fixable from application code — revisit if a newer Univer
// release adds column-span width handling. Kept here (unused) since the
// data model side is correct and this becomes a one-line UI change
// (re-add the Merge button in TableRibbon) once upstream catches up.
//
// An earlier version of this command instead deleted the absorbed cells'
// dataStream content and removed them from the tableCells array (mirroring
// Univer's DocTableDeleteColumnsCommand). That passed Univer's own
// structural-integrity checks but silently failed to render — array
// removal desyncs a row's cell count from its sibling rows and from
// tableColumns, which the layout engine doesn't expect. Property-only
// edits (this version) avoid that whole class of problem and match every
// other command in this file's low-risk category.
//
// Scoped to a single row deliberately — a real rowSpan merge needs the
// same columnSpan-style tombstoning applied per-row, which is realistic
// follow-up work, not included here. Splitting a merged cell back apart is
// the mirror image (columnSpan back to 1 on the tombstoned cells).

export interface IMergeTableCellsParams {
  range?: SelectedTableRange | null;
}

export const MergeTableCellsCommandId = "dockaro.command.table-merge-cells";

export const MergeTableCellsCommand: ICommand<IMergeTableCellsParams> = {
  id: MergeTableCellsCommandId,
  type: CommandType.COMMAND,
  handler: (accessor, params) => {
    const range = resolveTableRange(accessor, params?.range);
    if (!range) return false;
    if (range.startRow !== range.endRow) return false;
    if (range.startColumn === range.endColumn) return false;

    const found = getDocAndTable(accessor, range.tableId);
    if (!found) return false;
    const { docDataModel, table } = found;

    const row = table.tableRows[range.startRow];
    const anchorCell = row?.tableCells[range.startColumn];
    if (!row || !anchorCell) return false;

    const jsonX = JSONX.getInstance();
    const rawActions: JSONXActions[] = [];

    setProperty(
      jsonX,
      rawActions,
      ["tableSource", range.tableId, "tableRows", range.startRow, "tableCells", range.startColumn, "columnSpan"],
      anchorCell.columnSpan,
      range.endColumn - range.startColumn + 1,
    );

    for (let c = range.startColumn + 1; c <= range.endColumn; c++) {
      const cell = row.tableCells[c];
      if (!cell) continue;
      setProperty(
        jsonX,
        rawActions,
        ["tableSource", range.tableId, "tableRows", range.startRow, "tableCells", c, "columnSpan"],
        cell.columnSpan,
        0,
      );
    }

    if (!runMutation(accessor, docDataModel, rawActions)) return false;
    placeCaretInCell(accessor, docDataModel, range.tableId, range.startRow, range.startColumn);
    return true;
  },
};

// ---------------------------------------------------------------------------
// Splitting a merged cell back apart (Word's "Split Cells")
// ---------------------------------------------------------------------------
//
// The mirror image of the merge above: the anchor goes back to spanning one
// column and every cell it absorbed (marked with a zero span) becomes a
// normal cell again.

export interface ISplitTableCellsParams {
  range?: SelectedTableRange | null;
}

export const SplitTableCellsCommandId = "dockaro.command.table-split-cells";

export const SplitTableCellsCommand: ICommand<ISplitTableCellsParams> = {
  id: SplitTableCellsCommandId,
  type: CommandType.COMMAND,
  handler: (accessor, params) => {
    const range = resolveTableRange(accessor, params?.range);
    if (!range) return false;
    const found = getDocAndTable(accessor, range.tableId);
    if (!found) return false;
    const { docDataModel, table } = found;

    const jsonX = JSONX.getInstance();
    const rawActions: JSONXActions[] = [];
    const anchors: { row: number; column: number }[] = [];

    for (let r = range.startRow; r <= range.endRow; r++) {
      const row = table.tableRows[r];
      if (!row) continue;
      // Walk back to the anchor: the caret can sit anywhere in the merged run.
      let anchorIndex = range.startColumn;
      while (anchorIndex > 0 && row.tableCells[anchorIndex]?.columnSpan === 0) anchorIndex--;
      const anchor = row.tableCells[anchorIndex];
      const span = anchor?.columnSpan ?? 1;
      if (!anchor || span <= 1) continue;

      setProperty(
        jsonX,
        rawActions,
        ["tableSource", range.tableId, "tableRows", r, "tableCells", anchorIndex, "columnSpan"],
        anchor.columnSpan,
        1,
      );
      for (let c = anchorIndex + 1; c < anchorIndex + span; c++) {
        const cell = row.tableCells[c];
        if (!cell) continue;
        setProperty(
          jsonX,
          rawActions,
          ["tableSource", range.tableId, "tableRows", r, "tableCells", c, "columnSpan"],
          cell.columnSpan,
          1,
        );
      }
      anchors.push({ row: r, column: anchorIndex });
    }

    if (!runMutation(accessor, docDataModel, rawActions)) return false;
    const first = anchors[0];
    if (first) placeCaretInCell(accessor, docDataModel, range.tableId, first.row, first.column);
    return true;
  },
};

// ---------------------------------------------------------------------------
// Cell alignment (9-way grid: 3 horizontal x 3 vertical)
// ---------------------------------------------------------------------------
//
// Vertical alignment is a table-cell property (existing pattern above).
// Horizontal alignment of a cell's text is a PARAGRAPH property — Univer
// has no separate "cell horizontal align," Word doesn't either, it's just
// the paragraph(s) inside the cell. Finding those paragraphs needs the
// cell's startIndex/endIndex, which only exists in the layout viewModel
// (not in tableSource metadata) — same tree-walk verified working by the
// merge command above. Bundled into one atomic command instead of two
// separate ones (cell vAlign + native doc.command.align-*) because the
// native align commands run against Univer's LIVE selection, which the
// ribbon's own inputs routinely clear (see resolveTableRange's docs) —
// running both through our own resolved `range` avoids that entirely.
function findCellRange(
  accessor: IAccessor,
  docDataModel: DocumentDataModel,
  tableId: string,
  row: number,
  column: number,
): { startIndex: number; endIndex: number } | null {
  const docSkeletonManagerService = getCommandSkeleton(accessor, docDataModel.getUnitId());
  if (!docSkeletonManagerService) return null;
  const viewModel = docSkeletonManagerService.getViewModel();

  const body = docDataModel.getBody();
  const tableMeta = body?.tables?.find((t) => t.tableId === tableId);
  if (!body || !tableMeta) return null;

  type TableCellNode = { startIndex: number; endIndex: number };
  type TableRowNode = { children: TableCellNode[] };
  type TableNode = { startIndex: number; children: TableRowNode[] };

  let tableNode: TableNode | null = null;
  for (const section of viewModel.getChildren()) {
    for (const paragraph of section.children) {
      const node = paragraph.children[0];
      if (node && node.startIndex === tableMeta.startIndex) {
        tableNode = node as unknown as TableNode;
        break;
      }
    }
    if (tableNode) break;
  }
  if (!tableNode) return null;

  const cell = tableNode.children[row]?.children[column];
  return cell ? { startIndex: cell.startIndex, endIndex: cell.endIndex } : null;
}

/**
 * Where a table cell begins in the document's character stream, read from
 * the model's own table tokens rather than the skeleton - the skeleton's
 * table node is not addressable straight after a structural mutation.
 */
function findCellStartOffset(
  docDataModel: DocumentDataModel,
  tableId: string,
  row: number,
  column: number,
): number | null {
  const body = docDataModel.getBody();
  const table = body?.tables?.find((t) => t.tableId === tableId);
  if (!body || !table) return null;

  const { dataStream } = body;
  let rowIndex = -1;
  let columnIndex = -1;
  for (let i = table.startIndex; i < table.endIndex; i++) {
    const char = dataStream[i];
    if (char === TABLE_ROW_START_TOKEN) {
      rowIndex++;
      columnIndex = -1;
    } else if (char === TABLE_CELL_START_TOKEN) {
      columnIndex++;
      if (rowIndex === row && columnIndex === column) return i + 1;
    }
  }
  return null;
}

/**
 * Leaves the caret inside a cell after a structural change, the way Word
 * does: merging two cells drops the cursor into the merged one instead of
 * clearing the selection (which would also drop the Table Design tab).
 */
function placeCaretInCell(
  accessor: IAccessor,
  docDataModel: DocumentDataModel,
  tableId: string,
  row: number,
  column: number,
) {
  const offset = findCellStartOffset(docDataModel, tableId, row, column);
  if (offset == null) return;
  accessor
    .get(DocSelectionManagerService)
    .replaceDocRanges([{ startOffset: offset, endOffset: offset }]);
}

export interface ISetTableCellAlignParams {
  horizontal: HorizontalAlign;
  vertical: VerticalAlignmentType;
  range?: SelectedTableRange | null;
}

export const SetTableCellAlignCommandId = "dockaro.command.table-cell-align";

export const SetTableCellAlignCommand: ICommand<ISetTableCellAlignParams> = {
  id: SetTableCellAlignCommandId,
  type: CommandType.COMMAND,
  handler: (accessor, params) => {
    if (!params) return false;
    const range = resolveTableRange(accessor, params.range);
    if (!range) return false;
    const found = getDocAndTable(accessor, range.tableId);
    if (!found) return false;
    const { docDataModel, table } = found;
    const body = docDataModel.getBody();
    if (!body?.paragraphs) return false;

    const jsonX = JSONX.getInstance();
    const rawActions: JSONXActions[] = [];

    for (let r = range.startRow; r <= range.endRow; r++) {
      const row = table.tableRows[r];
      if (!row) continue;
      for (let c = range.startColumn; c <= range.endColumn; c++) {
        const cell = row.tableCells[c];
        if (!cell) continue;

        setProperty(
          jsonX,
          rawActions,
          ["tableSource", range.tableId, "tableRows", r, "tableCells", c, "vAlign"],
          cell.vAlign,
          params.vertical,
        );

        const cellRange = findCellRange(accessor, docDataModel, range.tableId, r, c);
        if (!cellRange) continue;
        body.paragraphs.forEach((paragraph, pIndex) => {
          if (paragraph.startIndex < cellRange.startIndex || paragraph.startIndex > cellRange.endIndex) return;
          setProperty(
            jsonX,
            rawActions,
            ["body", "paragraphs", pIndex, "paragraphStyle", "horizontalAlign"],
            paragraph.paragraphStyle?.horizontalAlign,
            params.horizontal,
          );
        });
      }
    }

    return runMutation(accessor, docDataModel, rawActions);
  },
};

// ---------------------------------------------------------------------------
// Fit table to window (stretch to the page's content width, evenly)
// ---------------------------------------------------------------------------

export interface ISetTableFitToWindowParams {
  range?: SelectedTableRange | null;
}

export const SetTableFitToWindowCommandId = "dockaro.command.table-fit-to-window";

export const SetTableFitToWindowCommand: ICommand<ISetTableFitToWindowParams> = {
  id: SetTableFitToWindowCommandId,
  type: CommandType.COMMAND,
  handler: (accessor, params) => {
    const range = resolveTableRange(accessor, params?.range);
    if (!range) return false;
    const found = getDocAndTable(accessor, range.tableId);
    if (!found) return false;
    const { docDataModel, table } = found;

    const docStyle = docDataModel.getSnapshot().documentStyle;
    const pageWidth = docStyle?.pageSize?.width ?? 794;
    const marginLeft = docStyle?.marginLeft ?? 72;
    const marginRight = docStyle?.marginRight ?? 72;
    const contentWidth = Math.max(100, pageWidth - marginLeft - marginRight);
    const columnCount = table.tableColumns.length || 1;
    const perColumnWidth = Math.floor(contentWidth / columnCount);

    const jsonX = JSONX.getInstance();
    const rawActions: JSONXActions[] = [];

    setProperty(jsonX, rawActions, ["tableSource", range.tableId, "layout"], table.layout, TableLayoutType.FIXED);
    setProperty(
      jsonX,
      rawActions,
      ["tableSource", range.tableId, "size"],
      table.size,
      { type: TableSizeType.SPECIFIED, width: { v: contentWidth } },
    );

    table.tableColumns.forEach((col, i) => {
      setProperty(
        jsonX,
        rawActions,
        ["tableSource", range.tableId, "tableColumns", i, "size"],
        col.size,
        { type: TableSizeType.SPECIFIED, width: { v: perColumnWidth } },
      );
    });

    return runMutation(accessor, docDataModel, rawActions);
  },
};

export const ALL_TABLE_STYLE_COMMANDS: ICommand[] = [
  SetTableCellMarginCommand,
  ResizeTableColumnCommand,
  ResizeTableRowCommand,
  SetTableCellBackgroundCommand,
  SetTableCellBorderCommand,
  ToggleTableGridlinesCommand,
  SetTableCellVAlignCommand,
  SetTableCellAlignCommand,
  SetTableRowHeightCommand,
  SetTableColumnWidthCommand,
  SetTableLayoutCommand,
  SetTableFitToWindowCommand,
  SetTableBandedRowsCommand,
  SetTableHeaderRowCommand,
  MergeTableCellsCommand,
  SplitTableCellsCommand,
];
