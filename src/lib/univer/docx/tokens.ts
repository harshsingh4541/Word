// Univer's dataStream mixes text with single-character control tokens
// (DataStreamTreeTokenType). They are written here as character codes
// rather than literals so the source stays readable and greppable.
export const TABLE_START = String.fromCharCode(0x1a);
export const TABLE_ROW_START = String.fromCharCode(0x1b);
export const TABLE_CELL_START = String.fromCharCode(0x1c);
export const TABLE_CELL_END = String.fromCharCode(0x1d);
export const TABLE_ROW_END = String.fromCharCode(0x0e);
export const TABLE_END = String.fromCharCode(0x0f);
export const CUSTOM_RANGE_START = String.fromCharCode(0x1f);
export const CUSTOM_RANGE_END = String.fromCharCode(0x1e);
export const PARAGRAPH = String.fromCharCode(0x0d);
export const SECTION_BREAK = String.fromCharCode(0x0a);
export const COLUMN_BREAK = String.fromCharCode(0x0b);
export const PAGE_BREAK = String.fromCharCode(0x0c);
export const TAB = String.fromCharCode(0x09);
export const CUSTOM_BLOCK = String.fromCharCode(0x08);
export const DOCS_END = String.fromCharCode(0x00);
