import {
  BaselineOffset,
  BooleanNumber,
  CustomRangeType,
  SpacingRule,
  HorizontalAlign,
  ImageSourceType,
  NamedStyleType,
  PageOrientType,
  TableRowHeightRule,
  VerticalAlignmentType,
} from "@univerjs/core";
import type {
  ICustomRange,
  Nullable,
  ICustomTable,
  IDocumentBody,
  IDocumentData,
  IParagraph,
  IReferenceSource,
  ITable,
  ITableCell,
  ITextRun,
  ITextStyle,
} from "@univerjs/core";
import {
  COLUMN_BREAK,
  CUSTOM_BLOCK,
  CUSTOM_RANGE_END,
  CUSTOM_RANGE_START,
  DOCS_END,
  PAGE_BREAK,
  PARAGRAPH,
  SECTION_BREAK,
  TAB,
  TABLE_CELL_END,
  TABLE_CELL_START,
  TABLE_END,
  TABLE_ROW_END,
  TABLE_ROW_START,
  TABLE_START,
} from "./tokens";

// Univer's data model -> WordprocessingML (the XML inside a .docx).
//
// This is a real OOXML writer, not the HTML-with-a-.doc-extension trick the
// Word export used before: Word opens HTML-in-.doc only after warning that
// "the file format and extension don't match", other readers mis-detect it
// entirely, and everything HTML can't carry (real named styles, list
// numbering, section geometry, headers and footers) was lost on the way
// out. A genuine .docx opens silently everywhere.

/** All document-model geometry is in 96-DPI pixels; OOXML wants twips. */
const TWIPS_PER_PX = 15; // 1440 twips per inch / 96 px per inch
/** OOXML expresses image sizes in EMUs. */
const EMU_PER_PX = 9525;
/** One point, in twips: below this an "at least" line height is a no-op. */
const MIN_MEANINGFUL_LINE_TWIPS = 20;

export function pxToTwips(px: number): number {
  return Math.round(px * TWIPS_PER_PX);
}

export function escapeXml(value: string): string {
  return value.replace(/[&<>"']/g, (c) => {
    switch (c) {
      case "&":
        return "&amp;";
      case "<":
        return "&lt;";
      case ">":
        return "&gt;";
      case '"':
        return "&quot;";
      default:
        return "&apos;";
    }
  });
}

/** OOXML colors are bare RRGGBB hex - no leading hash, no rgb() form. */
function toHexColor(color: Nullable<string>): string | null {
  if (!color) return null;
  const value = color.trim();
  const hexMatch = /^#?([0-9a-f]{6})$/i.exec(value);
  if (hexMatch) return hexMatch[1].toUpperCase();
  const shortHex = /^#?([0-9a-f]{3})$/i.exec(value);
  if (shortHex) {
    return shortHex[1]
      .split("")
      .map((c) => c + c)
      .join("")
      .toUpperCase();
  }
  const rgb = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i.exec(value);
  if (rgb) {
    return [rgb[1], rgb[2], rgb[3]]
      .map((n) => Math.max(0, Math.min(255, Number(n))).toString(16).padStart(2, "0"))
      .join("")
      .toUpperCase();
  }
  return null;
}

/** Word style ids for the named styles Univer applies through its Styles menu. */
const NAMED_STYLE_IDS: Record<number, string> = {
  [NamedStyleType.TITLE]: "Title",
  [NamedStyleType.SUBTITLE]: "Subtitle",
  [NamedStyleType.HEADING_1]: "Heading1",
  [NamedStyleType.HEADING_2]: "Heading2",
  [NamedStyleType.HEADING_3]: "Heading3",
  [NamedStyleType.HEADING_4]: "Heading4",
  [NamedStyleType.HEADING_5]: "Heading5",
};

const ALIGNMENT: Record<number, string> = {
  [HorizontalAlign.LEFT]: "left",
  [HorizontalAlign.CENTER]: "center",
  [HorizontalAlign.RIGHT]: "right",
  [HorizontalAlign.JUSTIFIED]: "both",
  [HorizontalAlign.BOTH]: "both",
  [HorizontalAlign.DISTRIBUTED]: "distribute",
};

const CELL_VERTICAL_ALIGN: Record<number, string> = {
  [VerticalAlignmentType.TOP]: "top",
  [VerticalAlignmentType.CENTER]: "center",
  [VerticalAlignmentType.BOTTOM]: "bottom",
};

export interface DocxRelationship {
  id: string;
  type: string;
  target: string;
  targetMode?: "External";
}

export interface DocxImagePart {
  /** Zip path, e.g. word/media/image1.png. */
  path: string;
  extension: string;
  data: Uint8Array;
}

/**
 * Collects the relationships and numbering definitions the body conversion
 * discovers as it walks the document, so the caller can emit matching
 * `.rels` and `numbering.xml` parts.
 */
export class DocxContext {
  readonly relationships: DocxRelationship[] = [];
  readonly imageParts: DocxImagePart[] = [];
  /** listId -> numbering id used in numbering.xml. */
  readonly listNumbering = new Map<string, { numId: number; listType: string }>();
  /** drawingId -> relationship id, filled in before the body is converted. */
  readonly imageRelations = new Map<string, string>();
  private relCounter = 0;
  private imageCounter = 0;
  private numCounter = 0;

  addRelationship(type: string, target: string, targetMode?: "External"): string {
    const id = `rId${++this.relCounter}`;
    this.relationships.push({ id, type, target, targetMode });
    return id;
  }

  addImage(data: Uint8Array, extension: string): string {
    const index = ++this.imageCounter;
    this.imageParts.push({ path: `word/media/image${index}.${extension}`, extension, data });
    return this.addRelationship(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
      `media/image${index}.${extension}`,
    );
  }

  numIdForList(listId: string, listType: string): number {
    const existing = this.listNumbering.get(listId);
    if (existing) return existing.numId;
    const numId = ++this.numCounter;
    this.listNumbering.set(listId, { numId, listType });
    return numId;
  }
}

export const HYPERLINK_STYLE_ID = "Hyperlink";

function runProperties(style: ITextStyle | undefined, characterStyleId?: string): string {
  if (!style && !characterStyleId) return "";
  const parts: string[] = [];
  // Word's own Hyperlink character style is what makes a link look like a
  // link (blue, underlined) and keeps it consistent with links the user
  // adds later in Word itself.
  if (characterStyleId) parts.push(`<w:rStyle w:val="${characterStyleId}"/>`);
  if (style?.ff) {
    const font = escapeXml(style.ff);
    parts.push(`<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:eastAsia="${font}" w:cs="${font}"/>`);
  }
  if (style?.bl === BooleanNumber.TRUE) parts.push("<w:b/><w:bCs/>");
  if (style?.it === BooleanNumber.TRUE) parts.push("<w:i/><w:iCs/>");
  if (style?.st?.s === BooleanNumber.TRUE) parts.push("<w:strike/>");
  if (style?.ul?.s === BooleanNumber.TRUE) parts.push('<w:u w:val="single"/>');
  const color = toHexColor(style?.cl?.rgb);
  if (color) parts.push(`<w:color w:val="${color}"/>`);
  const background = toHexColor(style?.bg?.rgb);
  if (background) parts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${background}"/>`);
  if (typeof style?.fs === "number") {
    // OOXML font sizes are in half-points.
    const halfPoints = Math.round(style.fs * 2);
    parts.push(`<w:sz w:val="${halfPoints}"/><w:szCs w:val="${halfPoints}"/>`);
  }
  if (style?.va === BaselineOffset.SUBSCRIPT) parts.push('<w:vertAlign w:val="subscript"/>');
  if (style?.va === BaselineOffset.SUPERSCRIPT) parts.push('<w:vertAlign w:val="superscript"/>');
  return parts.length ? `<w:rPr>${parts.join("")}</w:rPr>` : "";
}

function paragraphProperties(paragraph: IParagraph | undefined, context: DocxContext): string {
  if (!paragraph) return "";
  const style = paragraph.paragraphStyle;
  const parts: string[] = [];

  const namedStyle = style?.namedStyleType;
  if (namedStyle && NAMED_STYLE_IDS[namedStyle]) parts.push(`<w:pStyle w:val="${NAMED_STYLE_IDS[namedStyle]}"/>`);

  if (paragraph.bullet) {
    const numId = context.numIdForList(paragraph.bullet.listId, paragraph.bullet.listType);
    parts.push(`<w:numPr><w:ilvl w:val="${paragraph.bullet.nestingLevel ?? 0}"/><w:numId w:val="${numId}"/></w:numPr>`);
  }

  const shading = toHexColor(style?.shading?.backgroundColor?.rgb);
  if (shading) parts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${shading}"/>`);

  const spacing: string[] = [];
  if (style?.spaceAbove?.v) spacing.push(`w:before="${pxToTwips(style.spaceAbove.v)}"`);
  if (style?.spaceBelow?.v) spacing.push(`w:after="${pxToTwips(style.spaceBelow.v)}"`);
  if (style?.lineSpacing) {
    // `lineSpacing` means two different things in Univer depending on
    // `spacingRule`, exactly as it does in Word: with AUTO it is a
    // multiplier (Word counts those in 240ths of a line), and otherwise it
    // is an absolute px height. Univer stamps `lineSpacing: 2` with no rule
    // on every table-cell paragraph, which its renderer treats as "at least
    // 2px" - i.e. a no-op - so exporting that as a multiplier would
    // double-space every table in Word.
    if (style.spacingRule === SpacingRule.AUTO) {
      spacing.push(`w:line="${Math.round(style.lineSpacing * 240)}" w:lineRule="auto"`);
    } else if (style.spacingRule === SpacingRule.EXACT) {
      spacing.push(`w:line="${pxToTwips(style.lineSpacing)}" w:lineRule="exact"`);
    } else if (pxToTwips(style.lineSpacing) >= MIN_MEANINGFUL_LINE_TWIPS) {
      // Anything below a point can never raise a real line's height, so
      // Univer's placeholder values are dropped rather than written out.
      spacing.push(`w:line="${pxToTwips(style.lineSpacing)}" w:lineRule="atLeast"`);
    }
  }
  if (spacing.length) parts.push(`<w:spacing ${spacing.join(" ")}/>`);

  const indent: string[] = [];
  if (style?.indentStart?.v) indent.push(`w:left="${pxToTwips(style.indentStart.v)}"`);
  if (style?.indentEnd?.v) indent.push(`w:right="${pxToTwips(style.indentEnd.v)}"`);
  if (style?.indentFirstLine?.v) indent.push(`w:firstLine="${pxToTwips(style.indentFirstLine.v)}"`);
  if (style?.hanging?.v) indent.push(`w:hanging="${pxToTwips(style.hanging.v)}"`);
  if (indent.length) parts.push(`<w:ind ${indent.join(" ")}/>`);

  if (style?.horizontalAlign && ALIGNMENT[style.horizontalAlign]) {
    parts.push(`<w:jc w:val="${ALIGNMENT[style.horizontalAlign]}"/>`);
  }

  // Word-compatible pagination flags Univer stores with the same meaning.
  if (style?.pageBreakBefore === BooleanNumber.TRUE) parts.push("<w:pageBreakBefore/>");
  if (style?.keepLines === BooleanNumber.TRUE) parts.push("<w:keepLines/>");
  if (style?.keepNext === BooleanNumber.TRUE) parts.push("<w:keepNext/>");
  if (style?.widowControl === BooleanNumber.FALSE) parts.push('<w:widowControl w:val="0"/>');

  const paragraphMarkStyle = runProperties(style?.textStyle);
  if (paragraphMarkStyle) parts.push(paragraphMarkStyle);

  return parts.length ? `<w:pPr>${parts.join("")}</w:pPr>` : "";
}

interface BodyScope {
  body: IDocumentBody;
  source: IReferenceSource;
  context: DocxContext;
  paragraphsByMark: Map<number, IParagraph>;
  tablesByStart: Map<number, ICustomTable>;
  hyperlinksByStart: Map<number, ICustomRange>;
}

function createScope(body: IDocumentBody, source: IReferenceSource, context: DocxContext): BodyScope {
  const paragraphsByMark = new Map<number, IParagraph>();
  for (const paragraph of body.paragraphs ?? []) paragraphsByMark.set(paragraph.startIndex, paragraph);
  const tablesByStart = new Map<number, ICustomTable>();
  for (const table of body.tables ?? []) tablesByStart.set(table.startIndex, table);
  const hyperlinksByStart = new Map<number, ICustomRange>();
  for (const range of body.customRanges ?? []) {
    if (range.rangeType === CustomRangeType.HYPERLINK) hyperlinksByStart.set(range.startIndex, range);
  }
  return { body, source, context, paragraphsByMark, tablesByStart, hyperlinksByStart };
}

/** The style that applies at a character offset, from the sorted text runs. */
function styleAt(runs: ITextRun[], offset: number): ITextStyle | undefined {
  for (const run of runs) {
    if (offset >= run.st && offset < run.ed) return run.ts;
    if (run.st > offset) break;
  }
  return undefined;
}

function hashString(value: string): number {
  let hash = 0;
  for (let i = 0; i < value.length; i++) hash = (hash * 31 + value.charCodeAt(i)) | 0;
  return hash;
}

function renderDrawing(scope: BodyScope, blockId: string): string {
  const drawing = scope.source.drawings?.[blockId];
  const relId = drawing ? scope.context.imageRelations.get(drawing.drawingId) : undefined;
  if (!drawing || !relId) return "";
  const size = drawing.docTransform?.size;
  const width = Math.round((size?.width ?? 200) * EMU_PER_PX);
  const height = Math.round((size?.height ?? 200) * EMU_PER_PX);
  const name = escapeXml(drawing.title || drawing.name || "Picture");
  const docPrId = (Math.abs(hashString(drawing.drawingId)) % 100000) + 1;
  // Inline (in-line-with-text) placement keeps the export simple and always
  // valid; Word re-flows it exactly like an inline picture of its own.
  return (
    "<w:r><w:drawing>" +
    '<wp:inline distT="0" distB="0" distL="0" distR="0">' +
    `<wp:extent cx="${width}" cy="${height}"/>` +
    `<wp:docPr id="${docPrId}" name="${name}"/>` +
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<pic:nvPicPr><pic:cNvPr id="0" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${width}" cy="${height}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
    "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"
  );
}

/** Builds the runs for the character range [start, end). */
function renderRuns(scope: BodyScope, start: number, end: number): string {
  const { dataStream } = scope.body;
  const runs = (scope.body.textRuns ?? [])
    .filter((run) => run.ed > start && run.st < end)
    .sort((a, b) => a.st - b.st);
  const customBlocks = scope.body.customBlocks ?? [];

  let out = "";
  let buffer = "";
  let bufferStyle: ITextStyle | undefined;
  // Hyperlinks are plain index ranges over the text, not sentinel-delimited:
  // Univer records them in `customRanges` as an inclusive [startIndex,
  // endIndex] pair and leaves the dataStream itself untouched (verified on a
  // saved document). So the link has to be opened and closed by position.
  let hyperlinkEnd: number | null = null;

  const flush = () => {
    if (!buffer) return;
    // xml:space="preserve" is what stops Word from dropping the leading and
    // trailing spaces the user actually typed.
    const linkStyle = hyperlinkEnd !== null ? HYPERLINK_STYLE_ID : undefined;
    out += `<w:r>${runProperties(bufferStyle, linkStyle)}<w:t xml:space="preserve">${escapeXml(buffer)}</w:t></w:r>`;
    buffer = "";
  };

  const closeHyperlink = () => {
    if (hyperlinkEnd === null) return;
    flush();
    out += "</w:hyperlink>";
    hyperlinkEnd = null;
  };

  for (let i = start; i < end; i++) {
    const char = dataStream[i];

    const link = scope.hyperlinksByStart.get(i);
    const url = typeof link?.properties?.url === "string" ? link.properties.url : undefined;
    if (link && url && hyperlinkEnd === null) {
      flush();
      const relId = scope.context.addRelationship(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        url,
        "External",
      );
      out += `<w:hyperlink r:id="${relId}">`;
      hyperlinkEnd = link.endIndex;
    }

    // Sentinel-delimited ranges (mentions and other whole-entity ranges) do
    // put markers in the stream; they carry no text of their own.
    if (char === CUSTOM_RANGE_START) continue;
    if (char === CUSTOM_RANGE_END) {
      closeHyperlink();
      continue;
    }
    if (char === CUSTOM_BLOCK) {
      flush();
      const block = customBlocks.find((b) => b.startIndex === i);
      if (block) out += renderDrawing(scope, block.blockId);
      continue;
    }
    if (char === TAB) {
      flush();
      out += `<w:r>${runProperties(styleAt(runs, i))}<w:tab/></w:r>`;
      continue;
    }

    if (char === PAGE_BREAK || char === COLUMN_BREAK) {
      flush();
      out += '<w:r><w:br w:type="page"/></w:r>';
      continue;
    }
    if (char === DOCS_END || char === SECTION_BREAK) continue;

    const style = styleAt(runs, i);
    if (style !== bufferStyle) {
      flush();
      bufferStyle = style;
    }
    buffer += char;

    // `endIndex` is inclusive, so the link closes after this character.
    if (hyperlinkEnd !== null && i >= hyperlinkEnd) closeHyperlink();
  }

  flush();
  closeHyperlink();
  return out;
}

function renderParagraph(scope: BodyScope, start: number, end: number): string {
  const paragraph = scope.paragraphsByMark.get(end);
  return `<w:p>${paragraphProperties(paragraph, scope.context)}${renderRuns(scope, start, end)}</w:p>`;
}

function cellProperties(cell: ITableCell | undefined, columnWidthPx: number | undefined): string {
  const parts: string[] = [];
  const width = cell?.size?.width?.v ?? columnWidthPx;
  parts.push(width ? `<w:tcW w:w="${pxToTwips(width)}" w:type="dxa"/>` : '<w:tcW w:w="0" w:type="auto"/>');
  if (cell?.columnSpan && cell.columnSpan > 1) parts.push(`<w:gridSpan w:val="${cell.columnSpan}"/>`);
  if (cell?.rowSpan && cell.rowSpan > 1) parts.push('<w:vMerge w:val="restart"/>');

  const borders = (["Top", "Left", "Bottom", "Right"] as const)
    .map((side) => {
      const border = cell?.[`border${side}` as const];
      if (!border) return "";
      const color = toHexColor(border.color?.rgb) ?? "auto";
      // OOXML border width is in eighths of a point.
      const eighths = Math.max(2, Math.round((border.width?.v ?? 1) * 0.75 * 8));
      return `<w:${side.toLowerCase()} w:val="single" w:sz="${eighths}" w:space="0" w:color="${color}"/>`;
    })
    .join("");
  if (borders) parts.push(`<w:tcBorders>${borders}</w:tcBorders>`);

  const fill = toHexColor(cell?.backgroundColor?.rgb);
  if (fill) parts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>`);
  if (cell?.vAlign && CELL_VERTICAL_ALIGN[cell.vAlign]) {
    parts.push(`<w:vAlign w:val="${CELL_VERTICAL_ALIGN[cell.vAlign]}"/>`);
  }
  return `<w:tcPr>${parts.join("")}</w:tcPr>`;
}

function tableIndent(source: ITable | undefined): string {
  const indent = source?.indent?.v;
  return indent ? `<w:tblInd w:w="${pxToTwips(indent)}" w:type="dxa"/>` : "";
}

function renderTable(scope: BodyScope, table: ICustomTable): string {
  const source = scope.source.tableSource?.[table.tableId];
  const { dataStream } = scope.body;
  const columns = source?.tableColumns ?? [];
  const grid = columns.map((column) => `<w:gridCol w:w="${pxToTwips(column.size?.width?.v ?? 100)}"/>`).join("");
  const totalWidth = columns.reduce((sum, column) => sum + (column.size?.width?.v ?? 0), 0);

  const rowsXml: string[] = [];
  let rowIndex = 0;
  let i = table.startIndex + 1;

  while (i < table.endIndex && dataStream[i] === TABLE_ROW_START) {
    const row = source?.tableRows?.[rowIndex];
    const cellsXml: string[] = [];
    let cellIndex = 0;
    i += 1;

    while (i < table.endIndex && dataStream[i] === TABLE_CELL_START) {
      const cellStart = i + 1;
      // Nested tables carry their own cell-end tokens, so only a cell end at
      // depth zero closes this cell.
      let depth = 0;
      let cursor = cellStart;
      while (cursor < table.endIndex) {
        const char = dataStream[cursor];
        if (char === TABLE_START) depth++;
        else if (char === TABLE_END) depth--;
        else if (char === TABLE_CELL_END && depth === 0) break;
        cursor++;
      }
      const cell = row?.tableCells?.[cellIndex];
      const content = renderBlocks(scope, cellStart, cursor);
      cellsXml.push(`<w:tc>${cellProperties(cell, columns[cellIndex]?.size?.width?.v)}${content || "<w:p/>"}</w:tc>`);
      cellIndex++;
      i = cursor + 1;
    }

    if (i < table.endIndex && dataStream[i] === TABLE_ROW_END) i += 1;

    const trProps: string[] = [];
    if (row?.trHeight?.val?.v && row.trHeight.hRule !== TableRowHeightRule.AUTO) {
      const rule = row.trHeight.hRule === TableRowHeightRule.EXACT ? "exact" : "atLeast";
      trProps.push(`<w:trHeight w:val="${pxToTwips(row.trHeight.val.v)}" w:hRule="${rule}"/>`);
    }
    if (row?.repeatHeaderRow === BooleanNumber.TRUE) trProps.push("<w:tblHeader/>");
    if (row?.cantSplit === BooleanNumber.TRUE) trProps.push("<w:cantSplit/>");
    rowsXml.push(`<w:tr>${trProps.length ? `<w:trPr>${trProps.join("")}</w:trPr>` : ""}${cellsXml.join("")}</w:tr>`);
    rowIndex++;
  }

  const tableProps =
    "<w:tblPr>" +
    `<w:tblW w:w="${totalWidth ? pxToTwips(totalWidth) : 0}" w:type="${totalWidth ? "dxa" : "auto"}"/>` +
    tableIndent(source) +
    "<w:tblBorders>" +
    ["top", "left", "bottom", "right", "insideH", "insideV"]
      .map((side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="A6A6A6"/>`)
      .join("") +
    "</w:tblBorders>" +
    '<w:tblLayout w:type="fixed"/>' +
    "</w:tblPr>";

  return `<w:tbl>${tableProps}<w:tblGrid>${grid}</w:tblGrid>${rowsXml.join("")}</w:tbl>`;
}

/** Converts the character range [start, end) into block-level XML. */
function renderBlocks(scope: BodyScope, start: number, end: number): string {
  const { dataStream } = scope.body;
  let out = "";
  let segmentStart = start;
  let i = start;

  while (i < end) {
    const char = dataStream[i];
    const table = char === TABLE_START ? scope.tablesByStart.get(i) : undefined;
    if (table) {
      out += renderTable(scope, table);
      i = table.endIndex;
      segmentStart = i;
      continue;
    }
    if (char === PARAGRAPH) {
      out += renderParagraph(scope, segmentStart, i);
      i += 1;
      segmentStart = i;
      continue;
    }
    i += 1;
  }

  // Text after the last paragraph mark still has to reach Word as a
  // paragraph, or it would silently vanish from the export.
  if (segmentStart < end) {
    const trailing = Array.from(dataStream.slice(segmentStart, end));
    if (trailing.some((char) => char !== DOCS_END && char !== SECTION_BREAK)) {
      out += renderParagraph(scope, segmentStart, end);
    }
  }
  return out;
}

export function bodyToXml(body: IDocumentBody, source: IReferenceSource, context: DocxContext): string {
  return renderBlocks(createScope(body, source, context), 0, body.dataStream.length);
}

export function sectionProperties(snapshot: IDocumentData, headerRelId?: string, footerRelId?: string): string {
  const style = snapshot.documentStyle ?? {};
  const width = pxToTwips(style.pageSize?.width ?? 794);
  const height = pxToTwips(style.pageSize?.height ?? 1123);
  const orient = style.pageOrient === PageOrientType.LANDSCAPE ? ' w:orient="landscape"' : "";
  const parts = [
    headerRelId ? `<w:headerReference w:type="default" r:id="${headerRelId}"/>` : "",
    footerRelId ? `<w:footerReference w:type="default" r:id="${footerRelId}"/>` : "",
    `<w:pgSz w:w="${width}" w:h="${height}"${orient}/>`,
    `<w:pgMar w:top="${pxToTwips(style.marginTop ?? 72)}" w:right="${pxToTwips(style.marginRight ?? 72)}"` +
      ` w:bottom="${pxToTwips(style.marginBottom ?? 72)}" w:left="${pxToTwips(style.marginLeft ?? 72)}"` +
      ` w:header="${pxToTwips(style.marginHeader ?? 36)}" w:footer="${pxToTwips(style.marginFooter ?? 36)}" w:gutter="0"/>`,
    '<w:cols w:space="720"/>',
    '<w:docGrid w:linePitch="360"/>',
  ];
  return `<w:sectPr>${parts.join("")}</w:sectPr>`;
}

function mimeToExtension(mime: string): string {
  if (mime.includes("jpeg") || mime.includes("jpg")) return "jpeg";
  if (mime.includes("gif")) return "gif";
  if (mime.includes("svg")) return "svg";
  if (mime.includes("bmp")) return "bmp";
  return "png";
}

export async function loadImageData(
  source: string,
  sourceType: ImageSourceType | string,
): Promise<{ data: Uint8Array; extension: string } | null> {
  try {
    if (!source) return null;
    if (source.startsWith("data:") || sourceType === ImageSourceType.BASE64) {
      const match = /^data:(image\/[a-z+]+);base64,(.*)$/i.exec(source);
      if (!match) return null;
      const binary = atob(match[2]);
      const data = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) data[i] = binary.charCodeAt(i);
      return { data, extension: mimeToExtension(match[1]) };
    }
    const response = await fetch(source);
    if (!response.ok) return null;
    const buffer = await response.arrayBuffer();
    return {
      data: new Uint8Array(buffer),
      extension: mimeToExtension(response.headers.get("content-type") ?? "image/png"),
    };
  } catch {
    // An image we can't read (revoked blob URL, CORS, offline) must not take
    // the whole export down - the rest of the document still exports.
    return null;
  }
}
