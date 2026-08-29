import { NAMED_STYLE_MAP, NAMED_STYLE_SPACE_MAP, NamedStyleType } from "@univerjs/core";
import type { IDocumentData, ITextStyle, Nullable } from "@univerjs/core";
import { DocxContext, HYPERLINK_STYLE_ID, bodyToXml, escapeXml, loadImageData, pxToTwips, sectionProperties } from "./ooxml";
import { createZipBlob, textToBytes, type ZipEntry } from "./zip";

export const DOCX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

// The namespace set Word itself writes on w:document. `wp`, `a` and `pic`
// are only needed for images, but Word accepts (and re-writes) the full set
// on every document, and declaring them up front keeps the picture markup
// valid without conditional namespace juggling.
const DOCUMENT_NAMESPACES = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
].join(" ");

const REL_TYPE = {
  officeDocument: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
  coreProperties: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
  styles: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
  numbering: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
  header: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
  footer: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
};

/** Word style id + display name for each Univer named style. */
const STYLE_DEFINITIONS: { id: string; name: string; type: NamedStyleType; outlineLevel?: number }[] = [
  { id: "Title", name: "Title", type: NamedStyleType.TITLE },
  { id: "Subtitle", name: "Subtitle", type: NamedStyleType.SUBTITLE },
  { id: "Heading1", name: "heading 1", type: NamedStyleType.HEADING_1, outlineLevel: 0 },
  { id: "Heading2", name: "heading 2", type: NamedStyleType.HEADING_2, outlineLevel: 1 },
  { id: "Heading3", name: "heading 3", type: NamedStyleType.HEADING_3, outlineLevel: 2 },
  { id: "Heading4", name: "heading 4", type: NamedStyleType.HEADING_4, outlineLevel: 3 },
  { id: "Heading5", name: "heading 5", type: NamedStyleType.HEADING_5, outlineLevel: 4 },
];

function styleRunProperties(style: Nullable<ITextStyle>): string {
  if (!style) return "";
  const parts: string[] = [];
  if (style.bl) parts.push("<w:b/><w:bCs/>");
  if (style.it) parts.push("<w:i/><w:iCs/>");
  if (style.cl?.rgb) {
    const hex = String(style.cl.rgb).replace("#", "").toUpperCase();
    if (/^[0-9A-F]{6}$/.test(hex)) parts.push(`<w:color w:val="${hex}"/>`);
  }
  if (typeof style.fs === "number") {
    const halfPoints = Math.round(style.fs * 2);
    parts.push(`<w:sz w:val="${halfPoints}"/><w:szCs w:val="${halfPoints}"/>`);
  }
  return parts.join("");
}

/**
 * Heading/Title/Subtitle formatting lives only in Univer's runtime
 * NAMED_STYLE_MAP - applying "Heading 2" writes nothing but a
 * `namedStyleType` onto the paragraph, and the bold/size/color you see on
 * screen is resolved by the canvas on the fly. Emitting the same values as
 * real Word styles is what makes an exported document look like the one on
 * screen, and it also lights up Word's own Styles gallery and navigation
 * pane instead of exporting flat, unstyled text.
 */
function buildStylesXml(): string {
  const styles = STYLE_DEFINITIONS.map((definition) => {
    const textStyle = NAMED_STYLE_MAP[definition.type];
    const spacing = NAMED_STYLE_SPACE_MAP[definition.type];
    const spacingAttrs: string[] = [];
    if (spacing?.spaceAbove?.v) spacingAttrs.push(`w:before="${pxToTwips(spacing.spaceAbove.v)}"`);
    if (spacing?.spaceBelow?.v) spacingAttrs.push(`w:after="${pxToTwips(spacing.spaceBelow.v)}"`);
    const pPr =
      `<w:pPr>${spacingAttrs.length ? `<w:spacing ${spacingAttrs.join(" ")}/>` : ""}` +
      `${definition.outlineLevel !== undefined ? `<w:outlineLvl w:val="${definition.outlineLevel}"/>` : ""}` +
      "<w:keepNext/><w:keepLines/></w:pPr>";
    const rPr = styleRunProperties(textStyle);
    return (
      `<w:style w:type="paragraph" w:styleId="${definition.id}">` +
      `<w:name w:val="${definition.name}"/><w:basedOn w:val="Normal"/><w:qFormat/>` +
      pPr +
      (rPr ? `<w:rPr>${rPr}</w:rPr>` : "") +
      "</w:style>"
    );
  }).join("");

  return (
    `${XML_DECLARATION}<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    "<w:docDefaults><w:rPrDefault><w:rPr>" +
    '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/>' +
    '<w:sz w:val="22"/><w:szCs w:val="22"/>' +
    "</w:rPr></w:rPrDefault>" +
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>' +
    "</w:docDefaults>" +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/>' +
    '<w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:contextualSpacing/></w:pPr></w:style>' +
    `<w:style w:type="character" w:styleId="${HYPERLINK_STYLE_ID}"><w:name w:val="Hyperlink"/>` +
    '<w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>' +
    styles +
    "</w:styles>"
  );
}

// The exact glyphs Word writes for its three nested bullet levels: a Symbol
// font bullet, a Courier New "o", and a Wingdings square. They live in the
// fonts' private-use range, so they are written as code points.
const BULLET_GLYPHS = [String.fromCharCode(0xf0b7), "o", String.fromCharCode(0xf0a7)];
const BULLET_FONTS = ["Symbol", "Courier New", "Wingdings"];

function abstractNumbering(id: number, listType: string): string {
  const ordered = listType.startsWith("ORDER");
  const levels = Array.from({ length: 9 }, (_, level) => {
    const indent = pxToTwips(36 * (level + 1));
    const hanging = pxToTwips(18);
    if (ordered) {
      const formats = ["decimal", "lowerLetter", "lowerRoman"];
      const format = formats[level % formats.length];
      return (
        `<w:lvl w:ilvl="${level}"><w:start w:val="1"/><w:numFmt w:val="${format}"/>` +
        `<w:lvlText w:val="%${level + 1}."/><w:lvlJc w:val="left"/>` +
        `<w:pPr><w:ind w:left="${indent}" w:hanging="${hanging}"/></w:pPr></w:lvl>`
      );
    }
    const glyph = BULLET_GLYPHS[level % BULLET_GLYPHS.length];
    const font = BULLET_FONTS[level % BULLET_FONTS.length];
    return (
      `<w:lvl w:ilvl="${level}"><w:start w:val="1"/><w:numFmt w:val="bullet"/>` +
      `<w:lvlText w:val="${escapeXml(glyph)}"/><w:lvlJc w:val="left"/>` +
      `<w:pPr><w:ind w:left="${indent}" w:hanging="${hanging}"/></w:pPr>` +
      `<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:hint="default"/></w:rPr></w:lvl>`
    );
  }).join("");
  return `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="hybridMultilevel"/>${levels}</w:abstractNum>`;
}

function buildNumberingXml(context: DocxContext): string {
  const entries = [...context.listNumbering.values()];
  const abstracts = entries.map((entry) => abstractNumbering(entry.numId, entry.listType)).join("");
  const nums = entries
    .map((entry) => `<w:num w:numId="${entry.numId}"><w:abstractNumId w:val="${entry.numId}"/></w:num>`)
    .join("");
  return `${XML_DECLARATION}<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${abstracts}${nums}</w:numbering>`;
}

function buildRelsXml(context: DocxContext): string {
  const rels = context.relationships
    .map(
      (rel) =>
        `<Relationship Id="${rel.id}" Type="${rel.type}" Target="${escapeXml(rel.target)}"` +
        `${rel.targetMode ? ` TargetMode="${rel.targetMode}"` : ""}/>`,
    )
    .join("");
  return `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`;
}

function buildContentTypesXml(context: DocxContext, hasHeader: boolean, hasFooter: boolean): string {
  const extensions = new Set(["rels|application/vnd.openxmlformats-package.relationships+xml", "xml|application/xml"]);
  for (const image of context.imageParts) extensions.add(`${image.extension}|image/${image.extension}`);
  const defaults = [...extensions]
    .map((entry) => {
      const [extension, type] = entry.split("|");
      return `<Default Extension="${extension}" ContentType="${type}"/>`;
    })
    .join("");
  const overrides = [
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>',
    hasHeader
      ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
      : "",
    hasFooter
      ? '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
      : "",
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
  ].join("");
  return `${XML_DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${defaults}${overrides}</Types>`;
}

function buildCorePropsXml(title: string): string {
  const now = new Date().toISOString().replace(/\.\d{3}Z$/, "Z");
  return (
    `${XML_DECLARATION}<cp:coreProperties ` +
    'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
    'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
    'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    `<dc:title>${escapeXml(title)}</dc:title>` +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>` +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>` +
    "</cp:coreProperties>"
  );
}

/** The first header/footer defined in the document, used as the default one. */
function firstSegment<T>(map: Record<string, T> | undefined): T | undefined {
  if (!map) return undefined;
  const values = Object.values(map);
  return values.length ? values[0] : undefined;
}

export async function buildDocxBlob(snapshot: IDocumentData, title: string): Promise<Blob> {
  const context = new DocxContext();

  // Relationship order matters only for readability, but images have to be
  // resolved before the body is converted: the body needs their rel ids.
  for (const drawing of Object.values(snapshot.drawings ?? {})) {
    const source = (drawing as { source?: string }).source;
    const sourceType = (drawing as { imageSourceType?: string }).imageSourceType ?? "";
    if (!source) continue;
    const image = await loadImageData(source, sourceType);
    if (!image) continue;
    context.imageRelations.set(drawing.drawingId, context.addImage(image.data, image.extension));
  }

  const header = firstSegment(snapshot.headers);
  const footer = firstSegment(snapshot.footers);
  const headerRelId = header ? context.addRelationship(REL_TYPE.header, "header1.xml") : undefined;
  const footerRelId = footer ? context.addRelationship(REL_TYPE.footer, "footer1.xml") : undefined;
  context.addRelationship(REL_TYPE.styles, "styles.xml");
  context.addRelationship(REL_TYPE.numbering, "numbering.xml");

  const bodyXml = snapshot.body ? bodyToXml(snapshot.body, snapshot, context) : "";
  const headerXml = header ? bodyToXml(header.body, header, context) : "";
  const footerXml = footer ? bodyToXml(footer.body, footer, context) : "";

  const documentXml =
    `${XML_DECLARATION}<w:document ${DOCUMENT_NAMESPACES}><w:body>` +
    // Word discards a body with no paragraph at all, so an empty document
    // still gets one.
    (bodyXml || "<w:p/>") +
    sectionProperties(snapshot, headerRelId, footerRelId) +
    "</w:body></w:document>";

  const entries: ZipEntry[] = [
    { name: "[Content_Types].xml", data: textToBytes(buildContentTypesXml(context, Boolean(header), Boolean(footer))) },
    {
      name: "_rels/.rels",
      data: textToBytes(
        `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rId1" Type="${REL_TYPE.officeDocument}" Target="word/document.xml"/>` +
          `<Relationship Id="rId2" Type="${REL_TYPE.coreProperties}" Target="docProps/core.xml"/>` +
          "</Relationships>",
      ),
    },
    { name: "docProps/core.xml", data: textToBytes(buildCorePropsXml(title)) },
    { name: "word/document.xml", data: textToBytes(documentXml) },
    { name: "word/_rels/document.xml.rels", data: textToBytes(buildRelsXml(context)) },
    { name: "word/styles.xml", data: textToBytes(buildStylesXml()) },
    { name: "word/numbering.xml", data: textToBytes(buildNumberingXml(context)) },
  ];

  if (header) {
    entries.push({
      name: "word/header1.xml",
      data: textToBytes(`${XML_DECLARATION}<w:hdr ${DOCUMENT_NAMESPACES}>${headerXml || "<w:p/>"}</w:hdr>`),
    });
  }
  if (footer) {
    entries.push({
      name: "word/footer1.xml",
      data: textToBytes(`${XML_DECLARATION}<w:ftr ${DOCUMENT_NAMESPACES}>${footerXml || "<w:p/>"}</w:ftr>`),
    });
  }
  for (const image of context.imageParts) entries.push({ name: image.path, data: image.data });

  return createZipBlob(entries, DOCX_MIME_TYPE);
}
