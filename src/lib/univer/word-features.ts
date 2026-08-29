import {
  BooleanNumber,
  ColumnSeparatorType,
  CommandType,
  DashStyleType,
  HorizontalAlign,
  IUniverInstanceService,
  NamedStyleType,
  TabStopAlignment,
  TabStopLeader,
  UniverInstanceType,
} from "@univerjs/core";
import type { DocumentDataModel, IAccessor, ICommand, IParagraphStyle } from "@univerjs/core";
import { DocSelectionManagerService, DocSkeletonManagerService } from "@univerjs/docs";
import { IRenderManagerService } from "@univerjs/engine-render";
import type { FDocument, FDocumentParagraph } from "@univerjs/docs/facade";

// Word features that Univer's document model and renderer already support
// but ships no UI for: newspaper columns, per-section header/footer
// variants, paragraph shading and rules, cover pages and a table of
// contents. Each one is a command so it can sit on the ribbon, in a menu
// and behind a shortcut like everything else.

export const SetColumnsCommandId = "dockaro.command.columns";
export const SetHeaderFooterOptionsCommandId = "dockaro.command.header-footer-options";
export const SetParagraphShadingCommandId = "dockaro.command.paragraph-shading";
export const SetParagraphBorderCommandId = "dockaro.command.paragraph-border";
export const InsertCoverPageCommandId = "dockaro.command.cover-page";
export const InsertTableOfContentsCommandId = "dockaro.command.table-of-contents";

const PX_PER_INCH = 96;
/** Word's default gap between columns is half an inch. */
const COLUMN_GAP_PX = PX_PER_INCH / 2;

export interface ISetColumnsParams {
  count: number;
  /** Word's "Left" and "Right" presets: one narrow column beside a wide one. */
  weights?: number[];
  separator?: boolean;
}

export interface ISetHeaderFooterOptionsParams {
  /** Word's "Different first page". */
  useFirstPage?: boolean;
  /** Word's "Different odd & even pages". */
  oddEven?: boolean;
}

export interface ISetParagraphShadingParams {
  /** null clears the shading, as Word's "No Color" does. */
  color: string | null;
}

export interface ISetParagraphBorderParams {
  enabled: boolean;
  color?: string;
  /** Border thickness in pixels. */
  width?: number;
  dashStyle?: DashStyleType;
}

export interface IInsertCoverPageParams {
  design: "plain" | "banded";
}

export interface IInsertTableOfContentsParams {
  /** How many heading levels to list, as Word's "Show levels" does. */
  levels?: number;
}

function getDocModel(accessor: IAccessor): DocumentDataModel | null {
  return (
    accessor.get(IUniverInstanceService).getCurrentUnitOfType<DocumentDataModel>(UniverInstanceType.UNIVER_DOC) ??
    null
  );
}

function getCursorOffset(accessor: IAccessor): number | null {
  return accessor.get(DocSelectionManagerService).getActiveTextRange()?.startOffset ?? null;
}

/** Repaints after a structural change; see word-commands.ts for why. */
function forceRelayout(accessor: IAccessor, unitId: string) {
  const render = accessor.get(IRenderManagerService).getRenderUnitById(unitId);
  const skeleton = render?.with(DocSkeletonManagerService).getSkeleton();
  skeleton?.makeDirty(true);
  skeleton?.calculate();
  render?.mainComponent?.makeDirty(true);
}

/** The width text actually occupies on the page, in pixels. */
function contentWidth(docModel: DocumentDataModel): number {
  const style = docModel.getDocumentStyle();
  const width = style.pageSize?.width ?? 794;
  return width - (style.marginLeft ?? PX_PER_INCH) - (style.marginRight ?? PX_PER_INCH);
}

// ---------------------------------------------------------------------------
// Columns
// ---------------------------------------------------------------------------

function columnsCommand(doc: FDocument): ICommand<ISetColumnsParams> {
  return {
    id: SetColumnsCommandId,
    type: CommandType.COMMAND,
    handler: async (accessor, params) => {
      if (!params || params.count < 1) return false;
      const docModel = getDocModel(accessor);
      if (!docModel) return false;
      const offset = getCursorOffset(accessor);
      const section = (offset == null ? null : doc.getSectionAt(offset)) ?? doc.getSection(0);
      if (!section) return false;

      const separator = params.separator
        ? ColumnSeparatorType.BETWEEN_EACH_COLUMN
        : ColumnSeparatorType.NONE;

      if (params.count === 1) {
        // Word's "One" is the absence of column properties, not one column
        // of full width — that is what lets the text keep resizing with the
        // margins.
        section.setColumnProperties([], separator);
      } else {
        const total = contentWidth(docModel) - COLUMN_GAP_PX * (params.count - 1);
        const weights = params.weights ?? Array.from({ length: params.count }, () => 1);
        const sum = weights.reduce((a, b) => a + b, 0);
        section.setColumns(params.count, {
          gap: COLUMN_GAP_PX,
          widths: weights.map((weight) => Math.round((total * weight) / sum)),
          separator,
        });
      }

      forceRelayout(accessor, docModel.getUnitId());
      return true;
    },
  };
}

// ---------------------------------------------------------------------------
// Header and footer options
// ---------------------------------------------------------------------------

function headerFooterOptionsCommand(doc: FDocument): ICommand<ISetHeaderFooterOptionsParams> {
  return {
    id: SetHeaderFooterOptionsCommandId,
    type: CommandType.COMMAND,
    handler: async (accessor, params) => {
      if (!params) return false;
      const current = doc.getHeaderFooterOptions();
      const flag = (value: boolean | undefined, fallback: BooleanNumber | undefined) =>
        value === undefined ? fallback : value ? BooleanNumber.TRUE : BooleanNumber.FALSE;

      const applied = doc.setHeaderFooterOptions({
        useFirstPageHeaderFooter: flag(params.useFirstPage, current.useFirstPageHeaderFooter),
        evenAndOddHeaders: flag(params.oddEven, current.evenAndOddHeaders),
      });
      if (!applied) return false;

      // A variant with no body of its own renders as an empty header rather
      // than falling back to the default one, which reads as broken; create
      // it up front so the option takes visible effect immediately.
      const section = doc.getSection(0);
      if (params.useFirstPage) {
        section?.ensureHeader("first");
        section?.ensureFooter("first");
      }
      if (params.oddEven) {
        section?.ensureHeader("even");
        section?.ensureFooter("even");
      }

      forceRelayout(accessor, doc.getId());
      return true;
    },
  };
}

// ---------------------------------------------------------------------------
// Paragraph shading and rules
// ---------------------------------------------------------------------------

/** The paragraphs the selection touches, or the one holding the caret. */
function selectedParagraphs(accessor: IAccessor, doc: FDocument): FDocumentParagraph[] {
  const range = accessor.get(DocSelectionManagerService).getActiveTextRange();
  if (!range) return [];
  const start = Math.min(range.startOffset, range.endOffset);
  const end = Math.max(range.startOffset, range.endOffset);

  return doc.getParagraphs().filter((paragraph) => {
    const { startOffset, endOffset } = paragraph.getRange();
    return startOffset <= end && endOffset >= start;
  });
}

function applyToSelectedParagraphs(accessor: IAccessor, doc: FDocument, style: IParagraphStyle): boolean {
  const paragraphs = selectedParagraphs(accessor, doc);
  if (paragraphs.length === 0) return false;
  for (const paragraph of paragraphs) paragraph.setStyle(style);
  forceRelayout(accessor, doc.getId());
  return true;
}

function paragraphShadingCommand(doc: FDocument): ICommand<ISetParagraphShadingParams> {
  return {
    id: SetParagraphShadingCommandId,
    type: CommandType.COMMAND,
    handler: async (accessor, params) => {
      if (!params) return false;
      return applyToSelectedParagraphs(accessor, doc, {
        shading: params.color ? { backgroundColor: { rgb: params.color } } : undefined,
      });
    },
  };
}

function paragraphBorderCommand(doc: FDocument): ICommand<ISetParagraphBorderParams> {
  return {
    id: SetParagraphBorderCommandId,
    type: CommandType.COMMAND,
    handler: async (accessor, params) => {
      if (!params) return false;
      // Univer's renderer paints a paragraph's bottom border only — Word's
      // Bottom Border, the one people actually use as a rule under a
      // heading. Box/top/left/right have model support but nothing draws
      // them, so they are deliberately not offered.
      return applyToSelectedParagraphs(accessor, doc, {
        borderBottom: params.enabled
          ? {
              color: { rgb: params.color ?? "#000000" },
              width: params.width ?? 1,
              dashStyle: params.dashStyle ?? DashStyleType.SOLID,
              padding: 4,
            }
          : undefined,
      });
    },
  };
}

// ---------------------------------------------------------------------------
// Cover page
// ---------------------------------------------------------------------------

function coverPageCommand(doc: FDocument): ICommand<IInsertCoverPageParams> {
  return {
    id: InsertCoverPageCommandId,
    type: CommandType.COMMAND,
    handler: async (accessor, params) => {
      const banded = (params?.design ?? "plain") === "banded";
      const year = new Date().getFullYear();

      const lines: { text: string; style: IParagraphStyle }[] = [
        { text: "", style: {} },
        { text: "", style: {} },
        {
          text: "[Document title]",
          style: {
            namedStyleType: NamedStyleType.TITLE,
            horizontalAlign: HorizontalAlign.CENTER,
            ...(banded
              ? {
                  shading: { backgroundColor: { rgb: "#1F3864" } },
                  textStyle: { cl: { rgb: "#FFFFFF" } },
                }
              : {}),
          },
        },
        {
          text: "[Document subtitle]",
          style: { namedStyleType: NamedStyleType.SUBTITLE, horizontalAlign: HorizontalAlign.CENTER },
        },
        { text: "", style: {} },
        { text: String(year), style: { horizontalAlign: HorizontalAlign.CENTER } },
        { text: "[Company name]", style: { horizontalAlign: HorizontalAlign.CENTER } },
      ];

      // Inserted back to front at offset 0 so no offset arithmetic is
      // needed: each insertion pushes the previous one down.
      const inserted: FDocumentParagraph[] = [];
      for (let i = lines.length - 1; i >= 0; i--) {
        const paragraph = doc.insertParagraph(0, lines[i].text);
        if (!paragraph) return false;
        inserted.unshift(paragraph);
      }
      inserted.forEach((paragraph, index) => paragraph.setStyle(lines[index].style));

      // The cover gets a page to itself, exactly as Word's does.
      const end = inserted[inserted.length - 1]?.getRange().endOffset;
      if (end != null) doc.insertColumnBreak(end);

      forceRelayout(accessor, doc.getId());
      return true;
    },
  };
}

// ---------------------------------------------------------------------------
// Table of contents
// ---------------------------------------------------------------------------

const HEADING_LEVELS: Partial<Record<NamedStyleType, number>> = {
  [NamedStyleType.HEADING_1]: 1,
  [NamedStyleType.HEADING_2]: 2,
  [NamedStyleType.HEADING_3]: 3,
  [NamedStyleType.HEADING_4]: 4,
  [NamedStyleType.HEADING_5]: 5,
};

/**
 * Which printed page a character sits on. Only the laid-out skeleton knows,
 * so this walks up from the glyph at that offset to its page and asks for
 * that page's position in the document.
 */
function pageNumberAt(accessor: IAccessor, unitId: string, offset: number): number | null {
  const skeleton = accessor
    .get(IRenderManagerService)
    .getRenderUnitById(unitId)
    ?.with(DocSkeletonManagerService)
    .getSkeleton();
  const pages = skeleton?.getSkeletonData()?.pages;
  if (!skeleton || !pages) return null;

  type Node = { parent?: Node; sections?: unknown };
  let node = skeleton.findNodeByCharIndex(offset) as unknown as Node | null | undefined;
  while (node && !node.sections) node = node.parent;
  if (!node) return null;

  const index = pages.indexOf(node as never);
  return index === -1 ? null : index + 1;
}

function tableOfContentsCommand(doc: FDocument): ICommand<IInsertTableOfContentsParams> {
  return {
    id: InsertTableOfContentsCommandId,
    type: CommandType.COMMAND,
    handler: async (accessor, params) => {
      const docModel = getDocModel(accessor);
      const offset = getCursorOffset(accessor);
      if (!docModel || offset == null) return false;
      const maxLevel = params?.levels ?? 3;
      const unitId = docModel.getUnitId();

      const entries = doc
        .getParagraphs()
        .map((paragraph) => {
          const named = paragraph.getInfo().paragraph.paragraphStyle?.namedStyleType;
          const level = named == null ? undefined : HEADING_LEVELS[named];
          if (!level || level > maxLevel) return null;
          const text = paragraph.getText().trim();
          if (!text) return null;
          return { text, level, page: pageNumberAt(accessor, unitId, paragraph.getRange().startOffset) };
        })
        .filter((entry) => entry != null);

      if (entries.length === 0) return false;

      // A right-aligned tab stop at the right margin with a dot leader is
      // exactly how Word lays a contents line out: title, dots, page number.
      const tabStop = {
        offset: contentWidth(docModel),
        alignment: TabStopAlignment.END,
        leader: TabStopLeader.DOT,
      };

      const lines: { text: string; style: IParagraphStyle }[] = [
        { text: "Contents", style: { namedStyleType: NamedStyleType.HEADING_1 } },
        ...entries.map((entry) => ({
          text: `${entry.text}\t${entry.page ?? ""}`,
          style: {
            tabStops: [tabStop],
            indentStart: { v: (entry.level - 1) * (PX_PER_INCH / 4) },
          },
        })),
        { text: "", style: {} },
      ];

      const inserted: FDocumentParagraph[] = [];
      for (let i = lines.length - 1; i >= 0; i--) {
        const paragraph = doc.insertParagraph(offset, lines[i].text);
        if (!paragraph) return false;
        inserted.unshift(paragraph);
      }
      inserted.forEach((paragraph, index) => paragraph.setStyle(lines[index].style));

      forceRelayout(accessor, unitId);
      return true;
    },
  };
}

export function createWordFeatureCommands(doc: FDocument): ICommand[] {
  return [
    columnsCommand(doc),
    headerFooterOptionsCommand(doc),
    paragraphShadingCommand(doc),
    paragraphBorderCommand(doc),
    coverPageCommand(doc),
    tableOfContentsCommand(doc),
  ];
}
