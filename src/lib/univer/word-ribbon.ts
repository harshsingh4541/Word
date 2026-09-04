import { DashStyleType, IContextService, UniverInstanceType } from "@univerjs/core";
import type { IAccessor, IDisposable, Injector } from "@univerjs/core";
import {
  COLOR_PICKER_COMPONENT,
  ComponentManager,
  IMenuManagerService,
  IRibbonService,
  IShortcutService,
  IconManager,
  KeyCode,
  MenuItemType,
  MetaKeys,
  RibbonInsertGroup,
  RibbonStartGroup,
  getMenuHiddenObservable,
} from "@univerjs/ui";
import type { IMenuButtonItem, IMenuSelectorItem, IValueOption, MenuSchemaType } from "@univerjs/ui";
import TableGridPicker, { parseTableSize } from "@/components/editors/TableGridPicker";
import TableSizeField from "@/components/editors/TableSizeField";
import { whenDocAndEditorFocused } from "@univerjs/docs-ui";
import {
  DocSelectAllCommand,
  DocTableTabCommand,
  AlignCenterCommand,
  AlignJustifyCommand,
  AlignLeftCommand,
  AlignRightCommand,
  CreateDocTableCommand,
  DocCreateTableOperation,
  DocTableDeleteColumnsCommand,
  DocTableDeleteRowsCommand,
  DocTableDeleteTableCommand,
  DocTableInsertColumnLeftCommand,
  DocTableInsertColumnRightCommand,
  DocTableInsertRowAboveCommand,
  DocTableInsertRowBellowCommand,
} from "@univerjs/docs-ui";
import {
  AdjustWidthDoubleIcon,
  CancelMergeIcon,
  HorizontalBorderDoubleIcon,
  LineIndentDecreaseIcon,
  LineIndentIncreaseIcon,
  MergeAllIcon,
  TableBorderStyleIcon,
  AlignBottomIcon,
  AlignTopIcon,
  AllBorderIcon,
  AutoHeightDoubleIcon,
  CommentIcon,
  DocsMultiIcon,
  DownBorderDoubleIcon,
  ExportIcon,
  FlipHorizontalIcon,
  LeftBorderDoubleIcon,
  NoBorderIcon,
  PrintIcon,
  RightBorderDoubleIcon,
  RowHeightTallIcon,
  ShapeFrameIcon,
  ShapeRectIcon,
  ShrinkToFitIcon,
  UpBorderDoubleIcon,
  ZoomInIcon,
} from "@univerjs/icons";
import {
  EditHeaderFooterCommandId,
  ExportDocumentCommandId,
  InsertBlankPageCommandId,
  InsertPageBreakCommandId,
  MARGIN_PRESETS,
  PAGE_SIZE_PRESETS,
  SetIndentCommandId,
  SetLineSpacingCommandId,
  SetPageMarginsCommandId,
  SetPageOrientationCommandId,
  SetPageSizeCommandId,
  SetParagraphSpaceCommandId,
  SetZoomCommandId,
  PX_PER_INCH,
} from "./word-commands";
import { BORDER_WEIGHTS, SetBorderPenCommandId, getBorderPen, pointsToPixels } from "./border-pen";
import { ToggleSpellCheckCommandId } from "./spell-check";
import { SetWatermarkCommandId } from "./watermark";
import { ResolveTrackedChangesCommandId, ToggleTrackChangesCommandId } from "./track-changes";
import {
  InsertCoverPageCommandId,
  InsertTableOfContentsCommandId,
  SetColumnsCommandId,
  SetHeaderFooterOptionsCommandId,
  SetParagraphBorderCommandId,
  SetParagraphShadingCommandId,
} from "./word-features";
import {
  MergeTableCellsCommandId,
  SetTableBandedRowsCommandId,
  SetTableCellAlignCommandId,
  SetTableCellBackgroundCommandId,
  SetTableCellMarginCommandId,
  SetTableCellBorderCommandId,
  SetTableColumnWidthCommandId,
  SetTableFitToWindowCommandId,
  SetTableHeaderRowCommandId,
  SetTableLayoutCommandId,
  SetTableRowHeightCommandId,
  SplitTableCellsCommandId,
  type BorderSide,
} from "./table-style-commands";

/** Context key set by DocsEditor when the cursor is inside a table. */
export const WORD_CURSOR_IN_TABLE_CTX = "dockaro.ctx.cursorInTable";

// Word's ribbon, built out of Univer's own ribbon machinery rather than a
// second toolbar bolted on next to it. Univer's `grid` ribbon already has
// the shape Word's has — a tab strip over grouped, two-row controls — so
// what's left is: name the tabs the way Word names them, move Univer's
// page/header items onto a Layout tab, add the Word features Univer has no
// toolbar entry for, and give tables a real contextual tab. Everything
// here is a menu item bound to a registered command, so the same actions
// stay reachable from context menus and shortcuts.

/** Univer's fixed ribbon positions, re-labelled as Word's tabs. */
export const WORD_TAB = {
  FILE: "ribbon.others",
  HOME: "ribbon.start",
  INSERT: "ribbon.insert",
  LAYOUT: "ribbon.formulas",
  REVIEW: "ribbon.data",
  VIEW: "ribbon.view",
  TABLE: "ribbon.tableDesign",
} as const;

const WORD_GROUP = {
  FILE: "ribbon.others.others",
  LAYOUT_PAGE: "ribbon.formulas.basic",
  LAYOUT_PARAGRAPH: "ribbon.formulas.others",
  REVIEW_COMMENTS: "ribbon.data.rules",
  VIEW_ZOOM: "ribbon.view.display",
  TABLE_STYLE: "ribbon.tableDesign.style",
  TABLE_BORDERS: "ribbon.tableDesign.borders",
  TABLE_LAYOUT: "ribbon.tableDesign.cells",
  TABLE_ROWS: "ribbon.tableDesign.rows",
} as const;

/** Registered component id for the Word-style table size grid. */
const TABLE_GRID_PICKER_COMPONENT = "dockaro.component.table-grid-picker";
/** Registered component id for Word's row height / column width box. */
const TABLE_SIZE_FIELD_COMPONENT = "dockaro.component.table-size-field";

const COMMENT_PANEL_COMMAND_ID = "docs.operation.toggle-comment-panel";
const ADD_COMMENT_COMMAND_ID = "docs.operation.start-add-comment";
/** Univer's own items that Word keeps on other tabs than Univer does. */
export const HEADER_FOOTER_PANEL_COMMAND_ID = "doc.command.open-header-footer-panel";
export const PAGE_SETTING_COMMAND_ID = "docs.operation.open-page-setting";
/** Univer's own Table button, replaced here by Word's size grid. */
export const UNIVER_TABLE_MENU_ID = "doc.menu.table";

/**
 * Tab names, and titles for every control this app adds. Univer's own
 * locale keys stay untouched; `ui.ribbon.*` is overridden so the tab strip
 * reads Home / Insert / Layout / Review / View like Word's does.
 */
export const WORD_UI_LOCALE = {
  ui: {
    ribbon: {
      start: "Home",
      startDesc: "Fonts, paragraphs and styles.",
      insert: "Insert",
      insertDesc: "Tables, pictures, links and breaks.",
      formulas: "Layout",
      formulasDesc: "Margins, orientation, size and paragraph spacing.",
      data: "Review",
      dataDesc: "Comments and revisions.",
      view: "View",
      viewDesc: "Zoom and display options.",
      others: "File",
      othersDesc: "Export, print and page setup.",
    },
  },
  dockaro: {
    editing: {
      selectAll: "Select all",
    },
    file: {
      export: "Export",
      exportWord: "Word document (.docx)",
      exportPdf: "PDF",
      exportHtml: "Web page (.html)",
      print: "Print",
      pageSetup: "Page setup",
    },
    layout: {
      columns: "Columns",
      columnsOne: "One",
      columnsTwo: "Two",
      columnsThree: "Three",
      columnsLeft: "Left",
      columnsRight: "Right",
      columnsLine: "Two with line between",
      headerFooterOptions: "Header options",
      watermark: "Watermark",
      watermarkConfidential: "CONFIDENTIAL",
      watermarkDraft: "DRAFT",
      watermarkSample: "SAMPLE",
      watermarkNone: "No watermark",
      differentFirstPage: "Different first page",
      sameFirstPage: "Same first page",
      differentOddEven: "Different odd & even pages",
      sameOddEven: "Same odd & even pages",
      margins: "Margins",
      marginsNormal: "Normal (1 inch)",
      marginsNarrow: "Narrow (0.5 inch)",
      marginsModerate: "Moderate",
      marginsWide: "Wide (2 inch)",
      orientation: "Orientation",
      portrait: "Portrait",
      landscape: "Landscape",
      size: "Size",
      pageBreak: "Page break",
      blankPage: "Blank page",
      headerFooter: "Header & footer",
      spaceBefore: "Space before",
      spaceAfter: "Space after",
      spaceNone: "None",
    },
    paragraph: {
      shading: "Shading",
      borderBottom: "Bottom border",
      borderThin: "Thin line",
      borderThick: "Thick line",
      borderDashed: "Dashed line",
      borderNone: "No line",
      lineSpacing: "Line spacing",
      indentIncrease: "Increase indent",
      indentDecrease: "Decrease indent",
    },
    review: {
      trackChanges: "Track changes",
      trackOn: "Track changes on",
      trackOff: "Track changes off",
      resolveChanges: "Accept / reject",
      acceptAll: "Accept all changes",
      rejectAll: "Reject all changes",
      spelling: "Spelling",
      spellingOn: "Check spelling",
      spellingOff: "Off",
      newComment: "New comment",
      comments: "Comments",
    },
    view: {
      zoom: "Zoom",
    },
    insert: {
      coverPage: "Cover page",
      coverPlain: "Plain",
      coverBanded: "Banded",
      tableOfContents: "Contents",
      tocOneLevel: "Headings 1",
      tocThreeLevels: "Headings 1-3",
      tocFiveLevels: "Headings 1-5",
      table: "Table",
      tableDialog: "Insert table...",
    },
    table: {
      title: "Table Design",
      titleDesc: "Shading, borders and cell layout for the selected table.",
      shading: "Shading",
      shadingClear: "No shading",
      borders: "Borders",
      borderColor: "Pen color",
      borderStyle: "Line style",
      borderWeight: "Line weight",
      borderSolid: "Solid",
      borderDashed: "Dashed",
      borderDotted: "Dotted",
      merge: "Merge cells",
      split: "Split cells",
      bordersAll: "All borders",
      bordersNone: "No border",
      borderTop: "Top border",
      borderBottom: "Bottom border",
      borderLeft: "Left border",
      borderRight: "Right border",
      headerRow: "Header row",
      bandedRows: "Banded rows",
      bandedGrey: "Banded - grey",
      bandedBlue: "Banded - blue",
      bandedPurple: "Banded - purple",
      cellMargins: "Cell margins",
      cellMarginsNormal: "Normal",
      cellMarginsNarrow: "Narrow",
      cellMarginsNone: "None",
      off: "Off",
      align: "Cell alignment",
      rowHeight: "Row height",
      columnWidth: "Column width",
      autoFit: "AutoFit",
      autoFitContents: "AutoFit contents",
      autoFitWindow: "AutoFit window",
      fixedWidth: "Fixed column width",
      insert: "Insert",
      insertRowAbove: "Insert row above",
      insertRowBelow: "Insert row below",
      insertColumnLeft: "Insert column left",
      insertColumnRight: "Insert column right",
      delete: "Delete",
      deleteRow: "Delete row",
      deleteColumn: "Delete column",
      deleteTable: "Delete table",
    },
  },
};

/** Icons this app's ribbon items use that Univer doesn't already register. */
const WORD_ICONS = {
  AdjustWidthDoubleIcon,
  CancelMergeIcon,
  HorizontalBorderDoubleIcon,
  MergeAllIcon,
  TableBorderStyleIcon,
  // docs-ui imports these two but never registers them, so its own indent
  // menu ids resolve to an empty icon; registering them here fixes both.
  LineIndentDecreaseIcon,
  LineIndentIncreaseIcon,
  AlignBottomIcon,
  AlignTopIcon,
  AllBorderIcon,
  AutoHeightDoubleIcon,
  CommentIcon,
  DocsMultiIcon,
  DownBorderDoubleIcon,
  ExportIcon,
  FlipHorizontalIcon,
  LeftBorderDoubleIcon,
  NoBorderIcon,
  PrintIcon,
  RightBorderDoubleIcon,
  RowHeightTallIcon,
  ShapeFrameIcon,
  ShapeRectIcon,
  ShrinkToFitIcon,
  UpBorderDoubleIcon,
  ZoomInIcon,
};

interface ButtonOptions {
  id: string;
  commandId?: string;
  icon: string;
  title: string;
  params?: Record<string, unknown>;
}

function button(accessor: IAccessor, options: ButtonOptions): IMenuButtonItem {
  return {
    id: options.id,
    commandId: options.commandId,
    type: MenuItemType.BUTTON,
    icon: options.icon,
    title: options.title,
    tooltip: options.title,
    params: options.params,
    hidden$: getMenuHiddenObservable(accessor, UniverInstanceType.UNIVER_DOC),
  };
}

interface SelectorOptions {
  id: string;
  commandId?: string;
  icon: string;
  title: string;
  selections: IValueOption[];
}

function selector(accessor: IAccessor, options: SelectorOptions): IMenuSelectorItem {
  return {
    id: options.id,
    commandId: options.commandId,
    type: MenuItemType.SUBITEMS,
    icon: options.icon,
    title: options.title,
    tooltip: options.title,
    selections: options.selections,
    hidden$: getMenuHiddenObservable(accessor, UniverInstanceType.UNIVER_DOC),
  };
}

/** A dropdown entry that always sends the same fixed params. */
function option(label: string, params: Record<string, unknown>, icon?: string): IValueOption {
  return { label, value: label, icon, params: () => params };
}

/**
 * A Borders entry. Like Word, it draws with whatever the pen is set to at
 * the moment it is clicked rather than a fixed colour and weight.
 */
function borderOption(label: string, sides: BorderSide[], icon: string): IValueOption {
  return {
    label,
    value: label,
    icon,
    params: () => {
      const pen = getBorderPen();
      return { sides, color: pen.color, width: pen.width, dashStyle: pen.dashStyle };
    },
  };
}

const LINE_SPACINGS = [1, 1.15, 1.5, 2, 2.5, 3];
const ZOOM_LEVELS = [50, 75, 100, 125, 150, 200];
const ROW_HEIGHTS = [24, 32, 40, 48, 64];
const COLUMN_WIDTHS = [80, 100, 120, 160, 200];
/** Word's own "space before/after paragraph" presets, in points. */
const PARAGRAPH_SPACES = [0, 6, 12, 18];
const ALL_BORDER_SIDES: BorderSide[] = ["Top", "Bottom", "Left", "Right"];

const CELL_ALIGNMENTS: { label: string; icon: string; horizontal: number; vertical: number }[] = [
  // HorizontalAlign LEFT/CENTER/RIGHT = 1/2/3, VerticalAlignmentType TOP/CENTER/BOTTOM = 2/3/4.
  { label: "Top left", icon: "AlignTopIcon", horizontal: 1, vertical: 2 },
  { label: "Top center", icon: "AlignTopIcon", horizontal: 2, vertical: 2 },
  { label: "Top right", icon: "AlignTopIcon", horizontal: 3, vertical: 2 },
  { label: "Middle left", icon: "HorizontallyIcon", horizontal: 1, vertical: 3 },
  { label: "Middle center", icon: "HorizontallyIcon", horizontal: 2, vertical: 3 },
  { label: "Middle right", icon: "HorizontallyIcon", horizontal: 3, vertical: 3 },
  { label: "Bottom left", icon: "AlignBottomIcon", horizontal: 1, vertical: 4 },
  { label: "Bottom center", icon: "AlignBottomIcon", horizontal: 2, vertical: 4 },
  { label: "Bottom right", icon: "AlignBottomIcon", horizontal: 3, vertical: 4 },
];

function buildWordMenuSchema(): MenuSchemaType {
  return {
    [WORD_GROUP.FILE]: {
      [ExportDocumentCommandId]: {
        order: 0,
        gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true },
        menuItemFactory: (accessor: IAccessor) =>
          selector(accessor, {
            id: ExportDocumentCommandId,
            icon: "ExportIcon",
            title: "dockaro.file.export",
            selections: [
              option("dockaro.file.exportWord", { format: "word" }),
              option("dockaro.file.exportPdf", { format: "pdf" }),
              option("dockaro.file.exportHtml", { format: "html" }),
            ],
          }),
      },
      "dockaro.menu.print": {
        order: 1,
        gridLayout: { row: 1, column: 2, rowSpan: 2, showLabel: true },
        menuItemFactory: (accessor: IAccessor) =>
          button(accessor, {
            id: "dockaro.menu.print",
            commandId: ExportDocumentCommandId,
            icon: "PrintIcon",
            title: "dockaro.file.print",
            params: { format: "pdf" },
          }),
      },
      "dockaro.menu.page-setup": {
        order: 2,
        gridLayout: { row: 1, column: 3, rowSpan: 2, showLabel: true },
        menuItemFactory: (accessor: IAccessor) =>
          button(accessor, {
            id: "dockaro.menu.page-setup",
            commandId: PAGE_SETTING_COMMAND_ID,
            icon: "DocSettingIcon",
            title: "dockaro.file.pageSetup",
          }),
      },
    },

    [RibbonStartGroup.LAYOUT]: {
      [SetIndentCommandId]: {
        order: 20,
        gridLayout: { row: 1, column: 3 },
        menuItemFactory: (accessor: IAccessor) =>
          button(accessor, {
            id: SetIndentCommandId,
            icon: "LineIndentIncreaseIcon",
            title: "dockaro.paragraph.indentIncrease",
            params: { direction: "increase" },
          }),
      },
      "dockaro.menu.indent-decrease": {
        order: 21,
        gridLayout: { row: 2, column: 3 },
        menuItemFactory: (accessor: IAccessor) =>
          button(accessor, {
            id: "dockaro.menu.indent-decrease",
            commandId: SetIndentCommandId,
            icon: "LineIndentDecreaseIcon",
            title: "dockaro.paragraph.indentDecrease",
            params: { direction: "decrease" },
          }),
      },
      [SetParagraphShadingCommandId]: {
        order: 30,
        gridLayout: { row: 1, column: 5 },
        menuItemFactory: (accessor: IAccessor): IMenuSelectorItem => ({
          id: SetParagraphShadingCommandId,
          type: MenuItemType.SUBITEMS,
          icon: "PaintBucketDoubleIcon",
          title: "dockaro.paragraph.shading",
          tooltip: "dockaro.paragraph.shading",
          selections: [
            {
              label: { name: COLOR_PICKER_COMPONENT, hoverable: false, selectable: false },
              params: (value?: string | number) => ({ color: value }),
            },
          ],
          hidden$: getMenuHiddenObservable(accessor, UniverInstanceType.UNIVER_DOC),
        }),
      },
      [SetParagraphBorderCommandId]: {
        order: 31,
        gridLayout: { row: 2, column: 5 },
        menuItemFactory: (accessor: IAccessor) =>
          selector(accessor, {
            id: SetParagraphBorderCommandId,
            icon: "DownBorderDoubleIcon",
            title: "dockaro.paragraph.borderBottom",
            selections: [
              option("dockaro.paragraph.borderThin", { enabled: true, width: 1 }),
              option("dockaro.paragraph.borderThick", { enabled: true, width: 2 }),
              option("dockaro.paragraph.borderDashed", {
                enabled: true,
                width: 1,
                dashStyle: DashStyleType.DASH,
              }),
              option("dockaro.paragraph.borderNone", { enabled: false }),
            ],
          }),
      },
      [SetLineSpacingCommandId]: {
        order: 22,
        gridLayout: { row: 1, column: 4, rowSpan: 2 },
        menuItemFactory: (accessor: IAccessor) =>
          selector(accessor, {
            id: SetLineSpacingCommandId,
            icon: "RowHeightTallIcon",
            title: "dockaro.paragraph.lineSpacing",
            selections: LINE_SPACINGS.map((value) => option(value.toFixed(2).replace(/0$/, ""), { value })),
          }),
      },
    },

    [RibbonInsertGroup.MEDIA]: {
      "dockaro.menu.table": {
        order: -1,
        gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true },
        menuItemFactory: (accessor: IAccessor): IMenuSelectorItem => ({
          id: "dockaro.menu.table",
          commandId: CreateDocTableCommand.id,
          type: MenuItemType.SUBITEMS,
          icon: "GridIcon",
          title: "dockaro.insert.table",
          tooltip: "dockaro.insert.table",
          selections: [
            {
              label: { name: TABLE_GRID_PICKER_COMPONENT, hoverable: false, selectable: false },
              params: (value?: string | number) => parseTableSize(value) ?? undefined,
            },
            {
              label: "dockaro.insert.tableDialog",
              value: "dialog",
              id: DocCreateTableOperation.id,
              icon: "GridIcon",
            },
          ],
          hidden$: getMenuHiddenObservable(accessor, UniverInstanceType.UNIVER_DOC),
        }),
      },
      [InsertPageBreakCommandId]: {
        order: 10,
        gridLayout: { row: 1, column: 5, rowSpan: 2, showLabel: true },
        menuItemFactory: (accessor: IAccessor) =>
          button(accessor, {
            id: InsertPageBreakCommandId,
            icon: "DocsMultiIcon",
            title: "dockaro.layout.pageBreak",
          }),
      },
      [InsertCoverPageCommandId]: {
        order: 8,
        gridLayout: { row: 1, column: 8, rowSpan: 2, showLabel: true },
        menuItemFactory: (accessor: IAccessor) =>
          selector(accessor, {
            id: InsertCoverPageCommandId,
            icon: "ShapeRectIcon",
            title: "dockaro.insert.coverPage",
            selections: [
              option("dockaro.insert.coverPlain", { design: "plain" }),
              option("dockaro.insert.coverBanded", { design: "banded" }),
            ],
          }),
      },
      [InsertTableOfContentsCommandId]: {
        order: 9,
        gridLayout: { row: 1, column: 9, rowSpan: 2, showLabel: true },
        menuItemFactory: (accessor: IAccessor) =>
          selector(accessor, {
            id: InsertTableOfContentsCommandId,
            icon: "OrderIcon",
            title: "dockaro.insert.tableOfContents",
            selections: [
              option("dockaro.insert.tocThreeLevels", { levels: 3 }),
              option("dockaro.insert.tocOneLevel", { levels: 1 }),
              option("dockaro.insert.tocFiveLevels", { levels: 5 }),
            ],
          }),
      },
      [InsertBlankPageCommandId]: {
        order: 11,
        gridLayout: { row: 1, column: 6, rowSpan: 2, showLabel: true },
        menuItemFactory: (accessor: IAccessor) =>
          button(accessor, {
            id: InsertBlankPageCommandId,
            icon: "DocsMultiIcon",
            title: "dockaro.layout.blankPage",
          }),
      },
      [EditHeaderFooterCommandId]: {
        order: 12,
        gridLayout: { row: 1, column: 7, rowSpan: 2, showLabel: true },
        menuItemFactory: (accessor: IAccessor) =>
          button(accessor, {
            id: EditHeaderFooterCommandId,
            icon: "HeaderFooterIcon",
            title: "dockaro.layout.headerFooter",
          }),
      },
    },

    [WORD_GROUP.LAYOUT_PAGE]: {
      [SetPageMarginsCommandId]: {
        order: 0,
        gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true },
        menuItemFactory: (accessor: IAccessor) =>
          selector(accessor, {
            id: SetPageMarginsCommandId,
            icon: "ShapeFrameIcon",
            title: "dockaro.layout.margins",
            selections: (Object.keys(MARGIN_PRESETS) as (keyof typeof MARGIN_PRESETS)[]).map((preset) =>
              option(`dockaro.layout.margins${preset[0].toUpperCase()}${preset.slice(1)}`, { preset }),
            ),
          }),
      },
      [SetPageOrientationCommandId]: {
        order: 1,
        gridLayout: { row: 1, column: 2, rowSpan: 2, showLabel: true },
        menuItemFactory: (accessor: IAccessor) =>
          selector(accessor, {
            id: SetPageOrientationCommandId,
            icon: "FlipHorizontalIcon",
            title: "dockaro.layout.orientation",
            selections: [
              option("dockaro.layout.portrait", { orientation: "portrait" }),
              option("dockaro.layout.landscape", { orientation: "landscape" }),
            ],
          }),
      },
      [SetPageSizeCommandId]: {
        order: 2,
        gridLayout: { row: 1, column: 3, rowSpan: 2, showLabel: true },
        menuItemFactory: (accessor: IAccessor) =>
          selector(accessor, {
            id: SetPageSizeCommandId,
            icon: "ShapeRectIcon",
            title: "dockaro.layout.size",
            selections: PAGE_SIZE_PRESETS.map((size) => option(size, { size })),
          }),
      },
      "dockaro.menu.layout-page-break": {
        order: 3,
        gridLayout: { row: 1, column: 4, rowSpan: 2, showLabel: true },
        menuItemFactory: (accessor: IAccessor) =>
          button(accessor, {
            id: "dockaro.menu.layout-page-break",
            commandId: InsertPageBreakCommandId,
            icon: "DocsMultiIcon",
            title: "dockaro.layout.pageBreak",
          }),
      },
      "dockaro.menu.layout-header-footer": {
        order: 4,
        gridLayout: { row: 1, column: 5, rowSpan: 2, showLabel: true },
        menuItemFactory: (accessor: IAccessor) =>
          button(accessor, {
            id: "dockaro.menu.layout-header-footer",
            commandId: EditHeaderFooterCommandId,
            icon: "HeaderFooterIcon",
            title: "dockaro.layout.headerFooter",
          }),
      },
      [SetColumnsCommandId]: {
        order: 5,
        gridLayout: { row: 1, column: 6, rowSpan: 2, showLabel: true },
        menuItemFactory: (accessor: IAccessor) =>
          selector(accessor, {
            id: SetColumnsCommandId,
            icon: "ShrinkToFitIcon",
            title: "dockaro.layout.columns",
            selections: [
              option("dockaro.layout.columnsOne", { count: 1 }),
              option("dockaro.layout.columnsTwo", { count: 2 }),
              option("dockaro.layout.columnsThree", { count: 3 }),
              option("dockaro.layout.columnsLeft", { count: 2, weights: [1, 2] }),
              option("dockaro.layout.columnsRight", { count: 2, weights: [2, 1] }),
              option("dockaro.layout.columnsLine", { count: 2, separator: true }),
            ],
          }),
      },
      [SetWatermarkCommandId]: {
        order: 7,
        gridLayout: { row: 1, column: 8, rowSpan: 2, showLabel: true },
        menuItemFactory: (accessor: IAccessor) =>
          selector(accessor, {
            id: SetWatermarkCommandId,
            icon: "ShapeRectIcon",
            title: "dockaro.layout.watermark",
            selections: [
              option("dockaro.layout.watermarkConfidential", { text: "CONFIDENTIAL" }),
              option("dockaro.layout.watermarkDraft", { text: "DRAFT" }),
              option("dockaro.layout.watermarkSample", { text: "SAMPLE" }),
              option("dockaro.layout.watermarkNone", { text: null }),
            ],
          }),
      },
      [SetHeaderFooterOptionsCommandId]: {
        order: 6,
        gridLayout: { row: 1, column: 7, rowSpan: 2, showLabel: true },
        menuItemFactory: (accessor: IAccessor) =>
          selector(accessor, {
            id: SetHeaderFooterOptionsCommandId,
            icon: "HeaderFooterIcon",
            title: "dockaro.layout.headerFooterOptions",
            selections: [
              option("dockaro.layout.differentFirstPage", { useFirstPage: true }),
              option("dockaro.layout.sameFirstPage", { useFirstPage: false }),
              option("dockaro.layout.differentOddEven", { oddEven: true }),
              option("dockaro.layout.sameOddEven", { oddEven: false }),
            ],
          }),
      },
    },

    [WORD_GROUP.LAYOUT_PARAGRAPH]: {
      [SetParagraphSpaceCommandId]: {
        order: 0,
        gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true },
        menuItemFactory: (accessor: IAccessor) =>
          selector(accessor, {
            id: SetParagraphSpaceCommandId,
            icon: "AutoHeightDoubleIcon",
            title: "dockaro.layout.spaceBefore",
            selections: PARAGRAPH_SPACES.map((points) =>
              option(points === 0 ? "dockaro.layout.spaceNone" : `${points} pt`, {
                above: (points * PX_PER_INCH) / 72,
              }),
            ),
          }),
      },
      "dockaro.menu.space-after": {
        order: 1,
        gridLayout: { row: 1, column: 2, rowSpan: 2, showLabel: true },
        menuItemFactory: (accessor: IAccessor) =>
          selector(accessor, {
            id: "dockaro.menu.space-after",
            commandId: SetParagraphSpaceCommandId,
            icon: "AutoHeightDoubleIcon",
            title: "dockaro.layout.spaceAfter",
            selections: PARAGRAPH_SPACES.map((points) =>
              option(points === 0 ? "dockaro.layout.spaceNone" : `${points} pt`, {
                below: (points * PX_PER_INCH) / 72,
              }),
            ),
          }),
      },
    },

    [WORD_GROUP.REVIEW_COMMENTS]: {
      [ToggleTrackChangesCommandId]: {
        order: -3,
        gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true },
        menuItemFactory: (accessor: IAccessor) =>
          selector(accessor, {
            id: ToggleTrackChangesCommandId,
            icon: "CommentIcon",
            title: "dockaro.review.trackChanges",
            selections: [
              option("dockaro.review.trackOn", { enabled: true }),
              option("dockaro.review.trackOff", { enabled: false }),
            ],
          }),
      },
      [ResolveTrackedChangesCommandId]: {
        order: -2,
        gridLayout: { row: 1, column: 2, rowSpan: 2, showLabel: true },
        menuItemFactory: (accessor: IAccessor) =>
          selector(accessor, {
            id: ResolveTrackedChangesCommandId,
            icon: "CheckMarkIcon",
            title: "dockaro.review.resolveChanges",
            selections: [
              option("dockaro.review.acceptAll", { action: "accept" }),
              option("dockaro.review.rejectAll", { action: "reject" }),
            ],
          }),
      },
      [ToggleSpellCheckCommandId]: {
        order: -1,
        gridLayout: { row: 1, column: 3, rowSpan: 2, showLabel: true },
        menuItemFactory: (accessor: IAccessor) =>
          selector(accessor, {
            id: ToggleSpellCheckCommandId,
            icon: "FontColorDoubleIcon",
            title: "dockaro.review.spelling",
            selections: [
              option("dockaro.review.spellingOn", { enabled: true }),
              option("dockaro.review.spellingOff", { enabled: false }),
            ],
          }),
      },
      "dockaro.menu.new-comment": {
        order: 0,
        gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true },
        menuItemFactory: (accessor: IAccessor) =>
          button(accessor, {
            id: "dockaro.menu.new-comment",
            commandId: ADD_COMMENT_COMMAND_ID,
            icon: "CommentIcon",
            title: "dockaro.review.newComment",
          }),
      },
      "dockaro.menu.comment-panel": {
        order: 1,
        gridLayout: { row: 1, column: 2, rowSpan: 2, showLabel: true },
        menuItemFactory: (accessor: IAccessor) =>
          button(accessor, {
            id: "dockaro.menu.comment-panel",
            commandId: COMMENT_PANEL_COMMAND_ID,
            icon: "CommentIcon",
            title: "dockaro.review.comments",
          }),
      },
    },

    [WORD_GROUP.VIEW_ZOOM]: {
      [SetZoomCommandId]: {
        order: 0,
        gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true },
        menuItemFactory: (accessor: IAccessor) =>
          selector(accessor, {
            id: SetZoomCommandId,
            icon: "ZoomInIcon",
            title: "dockaro.view.zoom",
            selections: ZOOM_LEVELS.map((value) => option(`${value}%`, { value })),
          }),
      },
    },
  };
}

/**
 * Word shows table tools only while the cursor is inside a table, on a tab
 * that appears next to the permanent ones. Univer's ribbon supports exactly
 * that through `contextual`, so the table controls live in a real
 * contextual tab instead of a second toolbar that is always on screen.
 */
function buildRootMenuOverrides(): MenuSchemaType {
  return {
    ribbon: {
      // Word opens with File first; Univer's spare "others" tab is where
      // this app's File items live, so it moves to the front of the strip.
      [WORD_TAB.FILE]: { order: -1 },
      [WORD_TAB.TABLE]: {
        order: 6,
        title: "dockaro.table.title",
        contextual: true,

        [WORD_GROUP.TABLE_STYLE]: {
          order: 0,
          [SetTableCellBackgroundCommandId]: {
            order: 0,
            gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true },
            menuItemFactory: (accessor: IAccessor): IMenuSelectorItem => ({
              id: SetTableCellBackgroundCommandId,
              type: MenuItemType.SUBITEMS,
              icon: "PaintBucketDoubleIcon",
              title: "dockaro.table.shading",
              tooltip: "dockaro.table.shading",
              selections: [
                {
                  label: { name: COLOR_PICKER_COMPONENT, hoverable: false, selectable: false },
                  params: (value?: string | number) => ({ color: value }),
                },
              ],
              hidden$: getMenuHiddenObservable(accessor, UniverInstanceType.UNIVER_DOC),
            }),
          },
          "dockaro.menu.table-shading-clear": {
            order: 1,
            gridLayout: { row: 1, column: 2, rowSpan: 2, showLabel: true },
            menuItemFactory: (accessor: IAccessor) =>
              button(accessor, {
                id: "dockaro.menu.table-shading-clear",
                commandId: SetTableCellBackgroundCommandId,
                icon: "NoColorDoubleIcon",
                title: "dockaro.table.shadingClear",
                params: { color: null },
              }),
          },
          [SetTableHeaderRowCommandId]: {
            order: 2,
            gridLayout: { row: 1, column: 3, rowSpan: 2, showLabel: true },
            menuItemFactory: (accessor: IAccessor) =>
              selector(accessor, {
                id: SetTableHeaderRowCommandId,
                icon: "GridIcon",
                title: "dockaro.table.headerRow",
                selections: [
                  option("dockaro.table.headerRow", { enabled: true }),
                  option("dockaro.table.off", { enabled: false }),
                ],
              }),
          },
          [SetTableBandedRowsCommandId]: {
            order: 3,
            gridLayout: { row: 1, column: 4, rowSpan: 2, showLabel: true },
            menuItemFactory: (accessor: IAccessor) =>
              selector(accessor, {
                id: SetTableBandedRowsCommandId,
                icon: "GridIcon",
                title: "dockaro.table.bandedRows",
                selections: [
                  option("dockaro.table.bandedGrey", {
                    enabled: true,
                    colorOdd: "#F2F2F2",
                    colorEven: "#FFFFFF",
                  }),
                  option("dockaro.table.bandedBlue", {
                    enabled: true,
                    colorOdd: "#DEEAF6",
                    colorEven: "#FFFFFF",
                  }),
                  option("dockaro.table.bandedPurple", {
                    enabled: true,
                    colorOdd: "#F2EEFC",
                    colorEven: "#FFFFFF",
                  }),
                  option("dockaro.table.off", { enabled: false }),
                ],
              }),
          },
        },

        [WORD_GROUP.TABLE_BORDERS]: {
          order: 1,
          [SetTableCellBorderCommandId]: {
            order: 0,
            gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true },
            menuItemFactory: (accessor: IAccessor) =>
              selector(accessor, {
                id: SetTableCellBorderCommandId,
                icon: "AllBorderIcon",
                title: "dockaro.table.borders",
                selections: [
                  borderOption("dockaro.table.bordersAll", ALL_BORDER_SIDES, "AllBorderIcon"),
                  borderOption("dockaro.table.borderTop", ["Top"], "UpBorderDoubleIcon"),
                  borderOption("dockaro.table.borderBottom", ["Bottom"], "DownBorderDoubleIcon"),
                  borderOption("dockaro.table.borderLeft", ["Left"], "LeftBorderDoubleIcon"),
                  borderOption("dockaro.table.borderRight", ["Right"], "RightBorderDoubleIcon"),
                  option(
                    "dockaro.table.bordersNone",
                    { sides: ALL_BORDER_SIDES, color: null, width: 1 },
                    "NoBorderIcon",
                  ),
                ],
              }),
          },
          "dockaro.menu.border-style": {
            order: 1,
            gridLayout: { row: 1, column: 2, rowSpan: 2, showLabel: true },
            menuItemFactory: (accessor: IAccessor) =>
              selector(accessor, {
                id: "dockaro.menu.border-style",
                commandId: SetBorderPenCommandId,
                icon: "TableBorderStyleIcon",
                title: "dockaro.table.borderStyle",
                selections: [
                  option("dockaro.table.borderSolid", { dashStyle: DashStyleType.SOLID }),
                  option("dockaro.table.borderDashed", { dashStyle: DashStyleType.DASH }),
                  option("dockaro.table.borderDotted", { dashStyle: DashStyleType.DOT }),
                ],
              }),
          },
          "dockaro.menu.border-weight": {
            order: 2,
            gridLayout: { row: 1, column: 3, rowSpan: 2, showLabel: true },
            menuItemFactory: (accessor: IAccessor) =>
              selector(accessor, {
                id: "dockaro.menu.border-weight",
                commandId: SetBorderPenCommandId,
                icon: "HorizontalBorderDoubleIcon",
                title: "dockaro.table.borderWeight",
                selections: BORDER_WEIGHTS.map((points) =>
                  option(`${points} pt`, { width: pointsToPixels(points) }),
                ),
              }),
          },
          "dockaro.menu.table-border-color": {
            order: 3,
            gridLayout: { row: 1, column: 4, rowSpan: 2, showLabel: true },
            menuItemFactory: (accessor: IAccessor): IMenuSelectorItem => ({
              id: "dockaro.menu.table-border-color",
              commandId: SetBorderPenCommandId,
              type: MenuItemType.SUBITEMS,
              icon: "FontColorDoubleIcon",
              title: "dockaro.table.borderColor",
              tooltip: "dockaro.table.borderColor",
              selections: [
                {
                  label: { name: COLOR_PICKER_COMPONENT, hoverable: false, selectable: false },
                  params: (value?: string | number) => ({ color: value }),
                },
              ],
              hidden$: getMenuHiddenObservable(accessor, UniverInstanceType.UNIVER_DOC),
            }),
          },
        },

        [WORD_GROUP.TABLE_LAYOUT]: {
          order: 2,
          [SetTableCellAlignCommandId]: {
            order: 0,
            gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true },
            menuItemFactory: (accessor: IAccessor) =>
              selector(accessor, {
                id: SetTableCellAlignCommandId,
                icon: "HorizontallyIcon",
                title: "dockaro.table.align",
                selections: CELL_ALIGNMENTS.map((alignment) =>
                  option(alignment.label, { horizontal: alignment.horizontal, vertical: alignment.vertical }, alignment.icon),
                ),
              }),
          },
          [SetTableRowHeightCommandId]: {
            order: 1,
            gridLayout: { row: 1, column: 2, rowSpan: 2, showLabel: true },
            menuItemFactory: (accessor: IAccessor) =>
              selector(accessor, {
                id: SetTableRowHeightCommandId,
                icon: "AutoHeightDoubleIcon",
                title: "dockaro.table.rowHeight",
                selections: [
                  {
                    label: {
                      name: TABLE_SIZE_FIELD_COMPONENT,
                      hoverable: false,
                      selectable: false,
                      props: { title: "Height" },
                    },
                    params: (value?: string | number) => ({ mode: "fixed", height: Number(value) }),
                  },
                  option("dockaro.table.autoFit", { mode: "auto" }),
                  ...ROW_HEIGHTS.map((height) => option(`${height} px`, { mode: "fixed", height })),
                ],
              }),
          },
          [SetTableColumnWidthCommandId]: {
            order: 2,
            gridLayout: { row: 1, column: 3, rowSpan: 2, showLabel: true },
            menuItemFactory: (accessor: IAccessor) =>
              selector(accessor, {
                id: SetTableColumnWidthCommandId,
                icon: "AdjustWidthDoubleIcon",
                title: "dockaro.table.columnWidth",
                selections: [
                  {
                    label: {
                      name: TABLE_SIZE_FIELD_COMPONENT,
                      hoverable: false,
                      selectable: false,
                      props: { title: "Width" },
                    },
                    params: (value?: string | number) => ({ width: Number(value) }),
                  },
                  ...COLUMN_WIDTHS.map((width) => option(`${width} px`, { width })),
                ],
              }),
          },
          [MergeTableCellsCommandId]: {
            order: 4,
            gridLayout: { row: 1, column: 5, rowSpan: 2, showLabel: true },
            menuItemFactory: (accessor: IAccessor) =>
              button(accessor, {
                id: MergeTableCellsCommandId,
                icon: "MergeAllIcon",
                title: "dockaro.table.merge",
              }),
          },
          [SplitTableCellsCommandId]: {
            order: 5,
            gridLayout: { row: 1, column: 6, rowSpan: 2, showLabel: true },
            menuItemFactory: (accessor: IAccessor) =>
              button(accessor, {
                id: SplitTableCellsCommandId,
                icon: "CancelMergeIcon",
                title: "dockaro.table.split",
              }),
          },
          [SetTableCellMarginCommandId]: {
            order: 6,
            gridLayout: { row: 1, column: 7, rowSpan: 2, showLabel: true },
            menuItemFactory: (accessor: IAccessor) =>
              selector(accessor, {
                id: SetTableCellMarginCommandId,
                icon: "ShapeFrameIcon",
                title: "dockaro.table.cellMargins",
                selections: [
                  option("dockaro.table.cellMarginsNormal", {
                    margin: { start: 10, end: 10, top: 5, bottom: 5 },
                  }),
                  option("dockaro.table.cellMarginsNarrow", {
                    margin: { start: 6, end: 6, top: 1, bottom: 1 },
                  }),
                  option("dockaro.table.cellMarginsNone", {
                    margin: { start: 0, end: 0, top: 0, bottom: 0 },
                  }),
                ],
              }),
          },
          [SetTableLayoutCommandId]: {
            order: 3,
            gridLayout: { row: 1, column: 4, rowSpan: 2, showLabel: true },
            menuItemFactory: (accessor: IAccessor) =>
              selector(accessor, {
                id: SetTableLayoutCommandId,
                icon: "ShrinkToFitIcon",
                title: "dockaro.table.autoFit",
                selections: [
                  option("dockaro.table.autoFitContents", { layout: "auto" }),
                  { label: "dockaro.table.autoFitWindow", value: "window", id: SetTableFitToWindowCommandId },
                  option("dockaro.table.fixedWidth", { layout: "fixed" }),
                ],
              }),
          },
        },

        [WORD_GROUP.TABLE_ROWS]: {
          order: 3,
          "dockaro.menu.table-insert": {
            order: 0,
            gridLayout: { row: 1, column: 1, rowSpan: 2, showLabel: true },
            menuItemFactory: (accessor: IAccessor) =>
              selector(accessor, {
                id: "dockaro.menu.table-insert",
                icon: "InsertRowAboveDoubleIcon",
                title: "dockaro.table.insert",
                selections: [
                  {
                    label: "dockaro.table.insertRowAbove",
                    value: "row-above",
                    id: DocTableInsertRowAboveCommand.id,
                    icon: "InsertRowAboveDoubleIcon",
                  },
                  {
                    label: "dockaro.table.insertRowBelow",
                    value: "row-below",
                    id: DocTableInsertRowBellowCommand.id,
                    icon: "InsertRowBelowDoubleIcon",
                  },
                  {
                    label: "dockaro.table.insertColumnLeft",
                    value: "column-left",
                    id: DocTableInsertColumnLeftCommand.id,
                    icon: "LeftInsertColumnDoubleIcon",
                  },
                  {
                    label: "dockaro.table.insertColumnRight",
                    value: "column-right",
                    id: DocTableInsertColumnRightCommand.id,
                    icon: "RightInsertColumnDoubleIcon",
                  },
                ],
              }),
          },
          "dockaro.menu.table-delete": {
            order: 1,
            gridLayout: { row: 1, column: 2, rowSpan: 2, showLabel: true },
            menuItemFactory: (accessor: IAccessor) =>
              selector(accessor, {
                id: "dockaro.menu.table-delete",
                icon: "DeleteRowDoubleIcon",
                title: "dockaro.table.delete",
                selections: [
                  {
                    label: "dockaro.table.deleteRow",
                    value: "row",
                    id: DocTableDeleteRowsCommand.id,
                    icon: "DeleteRowDoubleIcon",
                  },
                  {
                    label: "dockaro.table.deleteColumn",
                    value: "column",
                    id: DocTableDeleteColumnsCommand.id,
                    icon: "DeleteColumnDoubleIcon",
                  },
                  {
                    label: "dockaro.table.deleteTable",
                    value: "table",
                    id: DocTableDeleteTableCommand.id,
                    icon: "DeleteTableDoubleIcon",
                  },
                ],
              }),
          },
        },
      },
    },
  };
}

export interface WordRibbon extends IDisposable {
  /** Shows or hides the contextual Table Design tab, as Word does. */
  setTableContextActive: (active: boolean) => void;
}

export function installWordRibbon(injector: Injector): WordRibbon {
  const iconManager = injector.get(IconManager);
  const componentManager = injector.get(ComponentManager);
  const menuManagerService = injector.get(IMenuManagerService);
  const ribbonService = injector.get(IRibbonService);

  const iconDisposable = iconManager.register(WORD_ICONS);
  const shortcutService = injector.get(IShortcutService);
  // Univer's shortcut panel throws if a group has no title of its own.
  const wordShortcut = (
    id: string,
    binding: number,
    description: string,
    extra?: { staticParameters?: object; priority?: number },
  ) =>
    shortcutService.registerShortcut({
      id,
      binding,
      preconditions: whenDocAndEditorFocused,
      description,
      group: "10_global-shortcut",
      groupTitle: "ui.global-shortcut",
      ...extra,
    });
  // Precondition: cursor is inside a table (for Tab navigation shortcuts).
  // The context key is set by DocsEditor whenever the selection changes.
  const isInTable = (contextService: IContextService): boolean =>
    whenDocAndEditorFocused(contextService) &&
    (contextService.getContextValue(WORD_CURSOR_IN_TABLE_CTX) ?? false);

  const shortcutDisposables = [
    // Word's Ctrl+Enter inserts a page break; Univer binds nothing to it.
    wordShortcut(InsertPageBreakCommandId, KeyCode.ENTER | MetaKeys.CTRL_COMMAND, "dockaro.layout.pageBreak"),
    // Univer binds alignment to Google Docs' Ctrl+Shift+L/E/R/J. Word's are
    // Ctrl+L/E/R/J, so those are added alongside — Univer's keep working.
    wordShortcut(AlignLeftCommand.id, KeyCode.L | MetaKeys.CTRL_COMMAND, "toolbar.alignLeft"),
    wordShortcut(AlignCenterCommand.id, KeyCode.E | MetaKeys.CTRL_COMMAND, "toolbar.alignCenter"),
    wordShortcut(AlignRightCommand.id, KeyCode.R | MetaKeys.CTRL_COMMAND, "toolbar.alignRight"),
    wordShortcut(AlignJustifyCommand.id, KeyCode.J | MetaKeys.CTRL_COMMAND, "toolbar.alignJustify"),
    // Univer's Ctrl+A widens the selection a step at a time (Google Docs:
    // sentence, paragraph, then document). Word selects the whole document
    // on the first press, so this takes the binding over — higher priority
    // wins when two shortcuts share one.
    wordShortcut(DocSelectAllCommand.id, KeyCode.A | MetaKeys.CTRL_COMMAND, "dockaro.editing.selectAll", {
      staticParameters: { wholeDocument: true },
      priority: 100,
    }),
    // Tab / Shift+Tab move the cursor between table cells, exactly as Word
    // does. Priority 200 ensures these beat any other Tab binding while the
    // cursor is inside a table. Outside a table the precondition is false and
    // the key reaches Univer's default handler unchanged.
    shortcutService.registerShortcut({
      id: DocTableTabCommand.id,
      binding: KeyCode.TAB,
      preconditions: isInTable,
      staticParameters: { shift: false },
      description: "dockaro.table.tab",
      group: "10_global-shortcut",
      groupTitle: "ui.global-shortcut",
      priority: 200,
    }),
    shortcutService.registerShortcut({
      id: DocTableTabCommand.id,
      binding: KeyCode.TAB | MetaKeys.SHIFT,
      preconditions: isInTable,
      staticParameters: { shift: true },
      description: "dockaro.table.tabBack",
      group: "10_global-shortcut",
      groupTitle: "ui.global-shortcut",
      priority: 200,
    }),
  ];
  const componentDisposable = componentManager.register(TABLE_GRID_PICKER_COMPONENT, TableGridPicker);
  const sizeFieldDisposable = componentManager.register(TABLE_SIZE_FIELD_COMPONENT, TableSizeField);
  menuManagerService.appendRootMenu(buildRootMenuOverrides());
  menuManagerService.mergeMenu(buildWordMenuSchema());

  let tableTabVisible = false;

  return {
    setTableContextActive: (active: boolean) => {
      if (active === tableTabVisible) return;
      tableTabVisible = active;
      // Word reveals the tab but leaves the active one alone, so a click
      // into a table never yanks the ribbon out from under the user.
      if (active) ribbonService.showContextualTab(WORD_TAB.TABLE);
      else ribbonService.hideContextualTab(WORD_TAB.TABLE);
    },
    dispose: () => {
      ribbonService.hideAllContextualTabs();
      componentDisposable.dispose();
      sizeFieldDisposable.dispose();
      for (const disposable of shortcutDisposables) disposable.dispose();
      iconDisposable.dispose();
    },
  };
}

/** Command ids Univer puts on its own tabs that this ribbon relocates. */
export const RELOCATED_UNIVER_MENU_ITEMS = {
  [HEADER_FOOTER_PANEL_COMMAND_ID]: { hidden: true },
  [PAGE_SETTING_COMMAND_ID]: { hidden: true },
  [UNIVER_TABLE_MENU_ID]: { hidden: true },
};
