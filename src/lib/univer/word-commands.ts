import {
  CommandType,
  DocumentFlavor,
  ICommandService,
  IUniverInstanceService,
  PAGE_SIZE,
  PageOrientType,
  SectionType,
  SpacingRule,
  UniverInstanceType,
} from "@univerjs/core";
import type {
  DocumentDataModel,
  IAccessor,
  ICommand,
  IDocumentData,
  IDocumentStyle,
  IParagraphStyle,
} from "@univerjs/core";
import { DocSelectionManagerService, DocSkeletonManagerService } from "@univerjs/docs";
import { DocSelectionRenderService } from "@univerjs/docs-ui";
import { DocumentEditArea, IRenderManagerService } from "@univerjs/engine-render";
import { exportDocument, type ExportFormat } from "./doc-export";

// Every Word action this app adds on top of Univer's own is a real
// registered command, not a React click handler. That is what lets the
// same action appear in the ribbon, in a context menu and behind a
// keyboard shortcut and still be one implementation — Univer's own
// toolbar only ever calls `commandService.executeCommand(id, params)`.

export const SetLineSpacingCommandId = "dockaro.command.line-spacing";
export const SetParagraphSpaceCommandId = "dockaro.command.paragraph-space";
export const SetIndentCommandId = "dockaro.command.indent";
export const InsertPageBreakCommandId = "dockaro.command.page-break";
export const InsertBlankPageCommandId = "dockaro.command.blank-page";
export const EditHeaderFooterCommandId = "dockaro.command.header-footer";
export const ExportDocumentCommandId = "dockaro.command.export";
export const SetPageMarginsCommandId = "dockaro.command.page-margins";
export const SetPageOrientationCommandId = "dockaro.command.page-orientation";
export const SetPageSizeCommandId = "dockaro.command.page-size";
export const SetZoomCommandId = "dockaro.command.zoom";

const DOC_PARAGRAPH_SETTING_COMMAND_ID = "doc-paragraph-setting.command";
const DOC_PAGE_SETUP_COMMAND_ID = "docs.command.page-setup";
const DOC_SET_ZOOM_RATIO_COMMAND_ID = "doc.command.set-zoom-ratio";

// 96 CSS px = 1 inch, which is the unit every page/margin value in the
// document model uses.
export const PX_PER_INCH = 96;

/** Word's four Layout > Margins presets, in inches. */
export const MARGIN_PRESETS = {
  normal: { top: 1, bottom: 1, left: 1, right: 1 },
  narrow: { top: 0.5, bottom: 0.5, left: 0.5, right: 0.5 },
  moderate: { top: 1, bottom: 1, left: 0.75, right: 0.75 },
  wide: { top: 1, bottom: 1, left: 2, right: 2 },
} as const;

export type MarginPreset = keyof typeof MARGIN_PRESETS;

/** The paper sizes Word's Layout > Size menu offers, ordered as Word does. */
export const PAGE_SIZE_PRESETS = ["Letter", "Legal", "A4", "A5", "A3", "Executive", "Tabloid"] as const;

export type PageSizePreset = (typeof PAGE_SIZE_PRESETS)[number];

/** Word's Increase/Decrease Indent step is half an inch. */
const INDENT_STEP_PX = PX_PER_INCH / 2;

/**
 * The few facade methods these commands need from the live document. Kept
 * structural so this module doesn't depend on which package re-exports
 * `FDocument` in a given Univer release.
 */
export interface WordDocumentApi {
  getId: () => string;
  getName: () => string;
  setName: (name: string) => unknown;
  save: () => IDocumentData;
  insertColumnBreak: (offset: number) => boolean;
  setSelection: (startOffset: number, endOffset: number) => unknown;
  ensurePageHeader: (pageIndex?: number) => string;
  insertSectionBreak: (offset: number, options?: { nextSectionType?: SectionType }) => unknown;
}

export interface WordCommandContext {
  doc: WordDocumentApi;
  /** The element Univer's canvas + contenteditable live in. */
  getContainer: () => HTMLElement | null;
}

function getDocModel(accessor: IAccessor): DocumentDataModel | null {
  const univerInstanceService = accessor.get(IUniverInstanceService);
  return univerInstanceService.getCurrentUnitOfType<DocumentDataModel>(UniverInstanceType.UNIVER_DOC) ?? null;
}

function getCursorOffset(accessor: IAccessor): number | null {
  const service = accessor.get(DocSelectionManagerService);
  return service.getActiveTextRange()?.startOffset ?? null;
}

/**
 * insertColumnBreak/page-setup only update the data model. Without a forced
 * relayout the canvas keeps hit-testing clicks against the stale skeleton,
 * so a click right after using one lands at the wrong offset (confirmed:
 * typed text landed at document start instead of the click point).
 */
function forceRelayout(accessor: IAccessor, unitId: string) {
  const render = accessor.get(IRenderManagerService).getRenderUnitById(unitId);
  const skeleton = render?.with(DocSkeletonManagerService)?.getSkeleton();
  skeleton?.makeDirty(true);
  skeleton?.calculate();
  render?.scene.makeDirty(true);
  render?.mainComponent?.makeDirty(true);
  void render?.scene.requestRender();
}

function applyParagraphStyle(
  accessor: IAccessor,
  paragraph: Partial<IParagraphStyle>,
): Promise<boolean> {
  const commandService = accessor.get(ICommandService);
  return commandService.executeCommand(DOC_PARAGRAPH_SETTING_COMMAND_ID, { paragraph }) as Promise<boolean>;
}

/**
 * `docs.command.page-setup` replaces the whole page geometry in one go, so
 * every caller has to send the values it isn't changing too. This reads the
 * current geometry and applies a patch on top.
 */
function applyPageSetup(
  accessor: IAccessor,
  patch: Partial<Pick<IDocumentStyle, "marginTop" | "marginBottom" | "marginLeft" | "marginRight" | "pageOrient" | "pageSize">>,
): boolean {
  const docModel = getDocModel(accessor);
  if (!docModel) return false;
  const style = docModel.getDocumentStyle();
  const commandService = accessor.get(ICommandService);
  void commandService.executeCommand(DOC_PAGE_SETUP_COMMAND_ID, {
    documentFlavor: style.documentFlavor ?? DocumentFlavor.TRADITIONAL,
    marginTop: style.marginTop,
    marginBottom: style.marginBottom,
    marginLeft: style.marginLeft,
    marginRight: style.marginRight,
    pageOrient: style.pageOrient ?? PageOrientType.PORTRAIT,
    pageSize: style.pageSize,
    ...patch,
  });
  forceRelayout(accessor, docModel.getUnitId());
  return true;
}

export interface ISetLineSpacingParams {
  value: number;
}

export interface ISetParagraphSpaceParams {
  /** Points of space added above/below the paragraph, Word's own unit. */
  above?: number;
  below?: number;
}

export interface ISetIndentParams {
  /** Word's Increase/Decrease Indent buttons: a half-inch step. */
  direction?: "increase" | "decrease";
  /** Absolute indents in document pixels, as dragged on the ruler. */
  indentStart?: number;
  indentEnd?: number;
  indentFirstLine?: number;
}

export interface IExportDocumentParams {
  format: ExportFormat;
}

export interface ISetPageMarginsParams {
  /** One of Word's Margins presets. */
  preset?: MarginPreset;
  /** Explicit margins in document pixels, as dragged on the rulers. */
  marginTop?: number;
  marginBottom?: number;
  marginLeft?: number;
  marginRight?: number;
}

export interface ISetPageOrientationParams {
  orientation: "portrait" | "landscape";
}

export interface ISetPageSizeParams {
  size: PageSizePreset;
}

export interface ISetZoomParams {
  /** Zoom as a percentage, matching Word's status-bar zoom control. */
  value: number;
}

export function createWordCommands(context: WordCommandContext): ICommand[] {
  const { doc, getContainer } = context;

  const setLineSpacing: ICommand<ISetLineSpacingParams> = {
    id: SetLineSpacingCommandId,
    type: CommandType.COMMAND,
    handler: async (accessor, params) => {
      if (!params) return false;
      // spacingRule AUTO is what makes lineSpacing a multiplier. Without it
      // the renderer treats the value as an absolute size, clamped to the
      // normal line height — so it looks like nothing happened.
      return applyParagraphStyle(accessor, { lineSpacing: params.value, spacingRule: SpacingRule.AUTO });
    },
  };

  const setParagraphSpace: ICommand<ISetParagraphSpaceParams> = {
    id: SetParagraphSpaceCommandId,
    type: CommandType.COMMAND,
    handler: async (accessor, params) => {
      if (!params) return false;
      const style: Partial<IParagraphStyle> = {};
      if (params.above !== undefined) style.spaceAbove = { v: params.above };
      if (params.below !== undefined) style.spaceBelow = { v: params.below };
      if (Object.keys(style).length === 0) return false;
      return applyParagraphStyle(accessor, style);
    },
  };

  const setIndent: ICommand<ISetIndentParams> = {
    id: SetIndentCommandId,
    type: CommandType.COMMAND,
    handler: async (accessor, params) => {
      if (!params) return false;
      if (!params.direction) {
        const style: Partial<IParagraphStyle> = {};
        if (params.indentStart !== undefined) style.indentStart = { v: Math.max(0, params.indentStart) };
        if (params.indentEnd !== undefined) style.indentEnd = { v: Math.max(0, params.indentEnd) };
        if (params.indentFirstLine !== undefined) {
          style.indentFirstLine = { v: Math.max(0, params.indentFirstLine) };
        }
        if (Object.keys(style).length === 0) return false;
        return applyParagraphStyle(accessor, style);
      }
      const docModel = getDocModel(accessor);
      const offset = getCursorOffset(accessor);
      if (!docModel || offset == null) return false;
      const paragraphs = docModel.getBody()?.paragraphs ?? [];
      const current = paragraphs.find((p) => p.startIndex >= offset) ?? paragraphs[paragraphs.length - 1];
      const indent = current?.paragraphStyle?.indentStart?.v ?? 0;
      const next = params.direction === "increase" ? indent + INDENT_STEP_PX : Math.max(0, indent - INDENT_STEP_PX);
      return applyParagraphStyle(accessor, { indentStart: { v: next } });
    },
  };

  const insertPageBreak: ICommand = {
    id: InsertPageBreakCommandId,
    type: CommandType.COMMAND,
    handler: async (accessor) => {
      const offset = getCursorOffset(accessor);
      if (offset == null) return false;
      // Univer has no dedicated page-break command, but insertColumnBreak's
      // own docs say it plainly: "In a single-column section, the
      // traditional renderer advances to the next physical page" — which is
      // exactly Word's Ctrl+Enter, since every document here is
      // single-column TRADITIONAL flavor.
      doc.insertColumnBreak(offset);
      forceRelayout(accessor, doc.getId());

      // Without moving the cursor past the break nothing tells the viewport
      // to scroll, so the new page sits off-screen below the fold and the
      // button looks like it did nothing.
      doc.setSelection(offset + 1, offset + 1);

      // setSelection only moves Univer's model of the cursor; real DOM focus
      // stays on the ribbon button that was just clicked, so the next space
      // bar press re-activates that button instead of typing. Put focus back
      // on Univer's editable surface, like a click into the page would.
      getContainer()?.querySelector<HTMLElement>('[contenteditable="true"]')?.focus();
      return true;
    },
  };

  const insertBlankPage: ICommand = {
    id: InsertBlankPageCommandId,
    type: CommandType.COMMAND,
    handler: async (accessor) => {
      const offset = getCursorOffset(accessor);
      if (offset == null) return false;
      // Word's Insert > Blank Page is two page breaks with an empty page
      // between them: one ends the current page, the other ends the blank
      // one so the following content keeps its own page.
      doc.insertColumnBreak(offset);
      doc.insertColumnBreak(offset + 1);
      forceRelayout(accessor, doc.getId());
      // Leave the caret on the blank page, ready to type.
      doc.setSelection(offset + 1, offset + 1);
      getContainer()?.querySelector<HTMLElement>('[contenteditable="true"]')?.focus();
      return true;
    },
  };

  const editHeaderFooter: ICommand = {
    id: EditHeaderFooterCommandId,
    type: CommandType.COMMAND,
    handler: async (accessor) => {
      // Univer's own "Header & footer" panel only shows real options when
      // edit focus is ALREADY inside a header/footer — otherwise it renders
      // "Header & footer settings are disabled". Its own way in is
      // double-clicking the page's top margin; this does the same thing in
      // one click: ensure a header segment exists, then move the same
      // edit-area/segment state that double-click sets.
      const unitId = doc.getId();
      const headerSegmentId = doc.ensurePageHeader(0);
      const render = accessor.get(IRenderManagerService).getRenderUnitById(unitId);
      if (!render) return false;

      render.with(DocSkeletonManagerService).getViewModel().setEditArea(DocumentEditArea.HEADER);
      const selectionRenderService = render.with(DocSelectionRenderService);
      selectionRenderService.setSegment(headerSegmentId);
      selectionRenderService.setSegmentPage(0);
      forceRelayout(accessor, unitId);
      return true;
    },
  };

  const exportDoc: ICommand<IExportDocumentParams> = {
    id: ExportDocumentCommandId,
    type: CommandType.COMMAND,
    handler: async (_accessor, params) => {
      if (!params) return false;
      await exportDocument(doc.save(), params.format);
      return true;
    },
  };

  const setPageMargins: ICommand<ISetPageMarginsParams> = {
    id: SetPageMarginsCommandId,
    type: CommandType.COMMAND,
    handler: async (accessor, params) => {
      if (!params) return false;
      if (!params.preset) {
        const patch: Parameters<typeof applyPageSetup>[1] = {};
        if (params.marginTop !== undefined) patch.marginTop = Math.max(0, params.marginTop);
        if (params.marginBottom !== undefined) patch.marginBottom = Math.max(0, params.marginBottom);
        if (params.marginLeft !== undefined) patch.marginLeft = Math.max(0, params.marginLeft);
        if (params.marginRight !== undefined) patch.marginRight = Math.max(0, params.marginRight);
        if (Object.keys(patch).length === 0) return false;
        return applyPageSetup(accessor, patch);
      }
      const preset = MARGIN_PRESETS[params.preset];
      if (!preset) return false;
      return applyPageSetup(accessor, {
        marginTop: preset.top * PX_PER_INCH,
        marginBottom: preset.bottom * PX_PER_INCH,
        marginLeft: preset.left * PX_PER_INCH,
        marginRight: preset.right * PX_PER_INCH,
      });
    },
  };

  const setPageOrientation: ICommand<ISetPageOrientationParams> = {
    id: SetPageOrientationCommandId,
    type: CommandType.COMMAND,
    handler: async (accessor, params) => {
      if (!params) return false;
      const docModel = getDocModel(accessor);
      if (!docModel) return false;
      const size = docModel.getDocumentStyle().pageSize ?? PAGE_SIZE.A4;
      const width = size.width ?? PAGE_SIZE.A4.width;
      const height = size.height ?? PAGE_SIZE.A4.height;
      const landscape = params.orientation === "landscape";
      // Orientation in Word is the page turned on its side, so the two
      // dimensions swap; keep whichever ordering the target orientation
      // needs rather than blindly swapping an already-correct page.
      const long = Math.max(width, height);
      const short = Math.min(width, height);
      return applyPageSetup(accessor, {
        pageOrient: landscape ? PageOrientType.LANDSCAPE : PageOrientType.PORTRAIT,
        pageSize: landscape ? { width: long, height: short } : { width: short, height: long },
      });
    },
  };

  const setPageSize: ICommand<ISetPageSizeParams> = {
    id: SetPageSizeCommandId,
    type: CommandType.COMMAND,
    handler: async (accessor, params) => {
      if (!params) return false;
      const size = PAGE_SIZE[params.size];
      if (!size) return false;
      const docModel = getDocModel(accessor);
      const landscape = docModel?.getDocumentStyle().pageOrient === PageOrientType.LANDSCAPE;
      return applyPageSetup(accessor, {
        pageSize: landscape ? { width: size.height, height: size.width } : { width: size.width, height: size.height },
      });
    },
  };

  const setZoom: ICommand<ISetZoomParams> = {
    id: SetZoomCommandId,
    type: CommandType.COMMAND,
    handler: async (accessor, params) => {
      if (!params) return false;
      const commandService = accessor.get(ICommandService);
      return Boolean(await commandService.executeCommand(DOC_SET_ZOOM_RATIO_COMMAND_ID, {
        zoomRatio: params.value / 100,
      }));
    },
  };

  // Univer's own right-click "Section Settings" submenu (Continuous / Next
  // page / Next column / Even page / Odd page) executes these exact ids
  // directly - confirmed from the crash it throws otherwise: `[CommandService]:
  // command "doc.menu.section-break.continuous" is not registered.` The menu
  // is real UI shipped by docs-ui; only the handlers behind it are missing in
  // this version, so they are supplied here.
  const sectionBreaks: { id: string; nextSectionType: SectionType }[] = [
    { id: "doc.menu.section-break.continuous", nextSectionType: SectionType.CONTINUOUS },
    { id: "doc.menu.section-break.next-page", nextSectionType: SectionType.NEXT_PAGE },
    { id: "doc.menu.section-break.next-column", nextSectionType: SectionType.NEXT_COLUMN },
    { id: "doc.menu.section-break.even-page", nextSectionType: SectionType.EVEN_PAGE },
    { id: "doc.menu.section-break.odd-page", nextSectionType: SectionType.ODD_PAGE },
  ];
  const sectionBreakCommands: ICommand[] = sectionBreaks.map(({ id, nextSectionType }) => ({
    id,
    type: CommandType.COMMAND,
    handler: async (accessor) => {
      const offset = getCursorOffset(accessor);
      if (offset == null) return false;
      if (!doc.insertSectionBreak(offset, { nextSectionType })) return false;
      forceRelayout(accessor, doc.getId());
      return true;
    },
  }));

  return [
    ...sectionBreakCommands,
    setLineSpacing,
    setParagraphSpace,
    setIndent,
    insertPageBreak,
    insertBlankPage,
    editHeaderFooter,
    exportDoc,
    setPageMargins,
    setPageOrientation,
    setPageSize,
    setZoom,
  ];
}
