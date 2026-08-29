import type { IDocumentBody, IDocumentData, IParagraph, IReferenceSource, ITextRun } from "@univerjs/core";
import { NAMED_STYLE_MAP } from "@univerjs/core";
import { convertBodyToHtml } from "@univerjs/docs-ui";
import { DOCX_MIME_TYPE, buildDocxBlob } from "./docx/package";

// Word export writes a real .docx (see docx/ooxml.ts). PDF and "Web page"
// go through convertBodyToHtml, the one export-adjacent helper that ships
// in the open-source packages: browsers print HTML to PDF natively, and a
// plain .html file is a genuinely useful third format. Neither needs
// @univerjs-pro/docs-exchange-client, which is Pro-only.

export type ExportFormat = "word" | "pdf" | "html";

const PX_PER_INCH = 96;

function pxToIn(px: number | undefined, fallback: number): string {
  return `${((px ?? fallback) / PX_PER_INCH).toFixed(2)}in`;
}

function getDocTitle(snapshot: IDocumentData): string {
  return snapshot.title?.trim() || "Untitled document";
}

// Applying a named style (Heading 1, Title, Subtitle, ...) only ever writes
// `paragraphStyle.namedStyleType` onto the paragraph — confirmed via
// debug_heading.js: applying Heading 2 through the real command left
// `textRuns` completely empty and added no `textStyle` to the paragraph.
// The bold/size/color every heading visibly gets in the live editor is a
// pure runtime default the canvas resolves from this same NAMED_STYLE_MAP
// on the fly; none of it is ever persisted into the document. convertBody-
// ToHtml has no such runtime fallback — it only emits `class="UniverHeading"
// aria-level="N"` with no visual styling at all — so every named-style
// paragraph exports as plain, unstyled text (confirmed: an exported resume
// with colored Heading 2 section titles opened in Word with those titles
// in plain black, non-bold text). Synthesizing the same default the canvas
// already applies, as a textRun, is what makes the export match what the
// user actually sees on screen.
function paragraphContentRanges(paragraphs: IParagraph[]): { start: number; end: number; paragraph: IParagraph }[] {
  const ranges: { start: number; end: number; paragraph: IParagraph }[] = [];
  let cursor = 0;
  for (const paragraph of paragraphs) {
    ranges.push({ start: cursor, end: paragraph.startIndex, paragraph });
    cursor = paragraph.startIndex + 1;
  }
  return ranges;
}

// Works around a bug in @univerjs/docs-ui's convertBodyToHtml: its inline
// slicer advances a shared cursor for every textRun in the WHOLE document,
// not just ones inside the current paragraph/cell. Any textRun located
// later in the dataStream — even one nowhere near the current slice — force-
// advances that cursor to the slice's end, silently dropping any plain
// (unstyled) text that comes before it. Confirmed by exporting a document
// with several styled headings and plain paragraphs between them: every
// plain paragraph ahead of a later styled run came out empty. Tiling the
// textRuns array so every character in the stream is already "covered" by
// some run sidesteps the bug entirely, since each paragraph's own slice
// then always finds a fully-intersecting run and never falls through to
// the buggy trailing-append path. Gaps inside a named-style paragraph get
// that style's NAMED_STYLE_MAP default instead of a no-op, for the reason
// above; every other gap gets a no-op {} run same as before.
function fillTextRunGaps(body: IDocumentBody): ITextRun[] {
  const { dataStream, textRuns = [], paragraphs = [] } = body;
  const paragraphRanges = paragraphContentRanges(paragraphs);
  const sorted = [...textRuns].sort((a, b) => a.st - b.st);
  const filled: ITextRun[] = [];

  const fillGap = (st: number, ed: number) => {
    for (const range of paragraphRanges) {
      const segStart = Math.max(st, range.start);
      const segEnd = Math.min(ed, range.end);
      if (segStart >= segEnd) continue;
      const namedStyleType = range.paragraph.paragraphStyle?.namedStyleType;
      const defaultStyle = namedStyleType != null ? NAMED_STYLE_MAP[namedStyleType] : null;
      filled.push({ st: segStart, ed: segEnd, ts: defaultStyle ?? {} });
    }
  };

  let cursor = 0;
  for (const run of sorted) {
    if (run.st > cursor) fillGap(cursor, run.st);
    filled.push(run);
    cursor = Math.max(cursor, run.ed);
  }
  if (cursor < dataStream.length) fillGap(cursor, dataStream.length);
  return filled;
}

// convertBodyToHtml never emits per-column widths on the <table>/<td>
// markup — only the table's own overall width. A real document's column
// widths (set via the table-column-width command, or dragged in the UI)
// are silently dropped on export: every consumer (a browser, Word) falls
// back to distributing width evenly across columns, which can badly
// mismatch what the column-width command actually did to the live
// document (confirmed: a 5-column table with deliberately uneven widths —
// narrow date column, wide description column — opened in Word with all
// columns roughly equal and headers wrapping oddly). A <colgroup> is the
// standard, Word-compatible way to carry column widths independent of
// individual cells, and matches what Word's own "Save as Web Page" emits
// for a table with custom column widths.
function injectColumnWidths(html: string, snapshot: IDocumentData): string {
  const tableRanges = snapshot.body?.tables ?? [];
  const tableSource = snapshot.tableSource ?? {};
  let tableIndex = 0;
  return html.replace(/<table class="MsoNormalTable UniverTable"([^>]*)><tbody>/g, (match, attrs: string) => {
    const tableRange = tableRanges[tableIndex++];
    const table = tableRange ? tableSource[tableRange.tableId] : undefined;
    const widths = table?.tableColumns
      ?.map((c) => c.size?.width?.v)
      .filter((w): w is number => typeof w === "number");
    if (!widths || widths.length === 0) return match;
    const colgroup = `<colgroup>${widths.map((w) => `<col style="width: ${w}px;">`).join("")}</colgroup>`;
    return `<table class="MsoNormalTable UniverTable"${attrs}>${colgroup}<tbody>`;
  });
}

function buildHtmlBody(snapshot: IDocumentData): string {
  const body = snapshot.body;
  if (!body) return convertBodyToHtml(snapshot);
  const patched: IDocumentData = {
    ...snapshot,
    body: { ...body, textRuns: fillTextRunGaps(body) },
  };
  return injectColumnWidths(convertBodyToHtml(patched), snapshot);
}

/** The first header/footer defined, which is the document's default one. */
function firstSegment<T>(map: Record<string, T> | undefined): T | undefined {
  const values = Object.values(map ?? {});
  return values.length ? values[0] : undefined;
}

/**
 * A header or footer converted on its own. It is a full document body in
 * its own right - it can hold tables and drawings - so it goes through the
 * same conversion as the main body, against its own reference source.
 */
function buildSegmentHtml(snapshot: IDocumentData, segment: IReferenceSource & { body: IDocumentBody }): string {
  return buildHtmlBody({ ...snapshot, ...segment } as IDocumentData);
}

/**
 * Wraps the document so its header and footer repeat on every printed page.
 * A table's thead and tfoot are the one construct browsers reliably repeat
 * across page breaks - Chrome ignores CSS `@page` margin boxes, which is
 * how Word itself would place them - so a header and footer set in the
 * editor now survive into PDF and print instead of being dropped.
 */
function withRunningHeaderFooter(snapshot: IDocumentData, content: string): string {
  const header = firstSegment(snapshot.headers);
  const footer = firstSegment(snapshot.footers);
  if (!header && !footer) return content;

  const headerRow = header
    ? `<thead><tr><td><div class="RunningHeader">${buildSegmentHtml(snapshot, header)}</div></td></tr></thead>`
    : "";
  const footerRow = footer
    ? `<tfoot><tr><td><div class="RunningFooter">${buildSegmentHtml(snapshot, footer)}</div></td></tr></tfoot>`
    : "";
  return `<table class="PageLayout">${headerRow}${footerRow}<tbody><tr><td>${content}</td></tr></tbody></table>`;
}

/** Styling for the running header and footer rows. */
const RUNNING_HEADER_CSS = `
  table.PageLayout { width: 100%; border-collapse: collapse; }
  table.PageLayout > thead > tr > td,
  table.PageLayout > tbody > tr > td,
  table.PageLayout > tfoot > tr > td { padding: 0; border: none; }
  .RunningHeader { padding-bottom: 8px; }
  .RunningFooter { padding-top: 8px; }`;

function buildPageCss(snapshot: IDocumentData) {
  const style = snapshot.documentStyle;
  const width = pxToIn(style?.pageSize?.width ?? undefined, 8.27 * PX_PER_INCH);
  const height = pxToIn(style?.pageSize?.height ?? undefined, 11.69 * PX_PER_INCH);
  const marginTop = pxToIn(style?.marginTop, 72);
  const marginBottom = pxToIn(style?.marginBottom, 72);
  const marginLeft = pxToIn(style?.marginLeft, 72);
  const marginRight = pxToIn(style?.marginRight, 72);
  return { width, height, marginTop, marginBottom, marginLeft, marginRight };
}

function downloadBlob(content: string | Blob, mime: string, filename: string) {
  const blob = typeof content === "string" ? new Blob([content], { type: mime }) : content;
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// A plain, portable HTML file — opens in any browser, editable in Word too
// (Word opens .html natively), no MS-specific markup.
export function exportAsHtml(snapshot: IDocumentData) {
  const title = getDocTitle(snapshot);
  const body = withRunningHeaderFooter(snapshot, buildHtmlBody(snapshot));
  const html = `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>${escapeHtml(title)}</title>
<style>
  body { font-family: Arial, Helvetica, sans-serif; color: #1b1c1f; max-width: 800px; margin: 40px auto; }
  p.UniverNormal, p.UniverHeading { margin: 0; }
  table.UniverTable { border-collapse: collapse; }
  table.UniverTable td.UniverTableCell { border: 1px solid #ccc; padding: 4px 8px; }${RUNNING_HEADER_CSS}
</style>
</head>
<body>
${body}
</body>
</html>`;
  downloadBlob(html, "text/html", `${title}.html`);
}

// A genuine .docx — the OOXML package Word, Google Docs, Pages and
// LibreOffice all open without a word of complaint. The previous export
// wrote HTML into a .doc file, which is why users hit Word's "The file
// format and extension of 'doc.doc' don't match. The file could be
// corrupted or unsafe." warning on every single open, and why the file
// silently lost styles, list numbering and page setup on the way in.
export async function exportAsWord(snapshot: IDocumentData) {
  const title = getDocTitle(snapshot);
  const blob = await buildDocxBlob(snapshot, title);
  downloadBlob(blob, DOCX_MIME_TYPE, `${title}.docx`);
}

// PDF via the browser's own print pipeline — a hidden iframe (not
// window.open) avoids popup-blocker issues entirely, since nothing new
// opens; the user picks "Save as PDF" in their browser's native print
// dialog, which every major browser supports without any library.
export function exportAsPdf(snapshot: IDocumentData) {
  const title = getDocTitle(snapshot);
  const body = withRunningHeaderFooter(snapshot, buildHtmlBody(snapshot));
  const { width, height, marginTop, marginBottom, marginLeft, marginRight } = buildPageCss(snapshot);
  const html = `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>${escapeHtml(title)}</title>
<style>
  @page {
    size: ${width} ${height};
    margin: ${marginTop} ${marginRight} ${marginBottom} ${marginLeft};
  }
  body { font-family: Arial, Helvetica, sans-serif; color: #1b1c1f; margin: 0; }
  p.UniverNormal, p.UniverHeading { margin: 0; }
  table.UniverTable { border-collapse: collapse; }
  table.UniverTable td.UniverTableCell { border: 1px solid #999; padding: 4px 8px; }${RUNNING_HEADER_CSS}
</style>
</head>
<body>
${body}
</body>
</html>`;

  const iframe = document.createElement("iframe");
  iframe.style.position = "fixed";
  iframe.style.right = "0";
  iframe.style.bottom = "0";
  iframe.style.width = "0";
  iframe.style.height = "0";
  iframe.style.border = "0";
  document.body.appendChild(iframe);

  const cleanup = () => {
    // Give the print dialog a moment to actually open before the iframe
    // (and the document it holds) gets torn down from under it.
    setTimeout(() => document.body.removeChild(iframe), 1000);
  };

  const frameWindow = iframe.contentWindow;
  if (!frameWindow) {
    document.body.removeChild(iframe);
    return;
  }
  frameWindow.document.open();
  frameWindow.document.write(html);
  frameWindow.document.close();
  frameWindow.onafterprint = cleanup;
  iframe.onload = () => {
    frameWindow.focus();
    frameWindow.print();
    // onafterprint isn't reliable across all browsers, so also clean up
    // on a fallback timer.
    setTimeout(cleanup, 60000);
  };
}

function escapeHtml(s: string): string {
  return s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]!);
}

/** Single entry point used by the ribbon's Export commands. */
export async function exportDocument(snapshot: IDocumentData, format: ExportFormat): Promise<void> {
  if (format === "word") return exportAsWord(snapshot);
  if (format === "pdf") return exportAsPdf(snapshot);
  return exportAsHtml(snapshot);
}
