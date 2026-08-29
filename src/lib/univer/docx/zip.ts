// A .docx file is an OOXML package: a plain ZIP archive of XML parts. This
// is the smallest ZIP writer that produces an archive Word accepts —
// stored (uncompressed) entries only, which the format allows and Word
// opens without complaint. Writing it here keeps the export dependency-free
// (no JSZip, no Univer Pro exchange service) and small.

export interface ZipEntry {
  name: string;
  data: Uint8Array;
}

const CRC32_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[i] = c >>> 0;
  }
  return table;
})();

function crc32(data: Uint8Array): number {
  let crc = 0xffffffff;
  for (let i = 0; i < data.length; i++) crc = CRC32_TABLE[(crc ^ data[i]) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

/** ZIP stores timestamps in the MS-DOS packed format, not epoch seconds. */
function dosDateTime(date: Date): { time: number; date: number } {
  const time = (date.getHours() << 11) | (date.getMinutes() << 5) | (Math.floor(date.getSeconds() / 2) & 0x1f);
  const dosDate = (((date.getFullYear() - 1980) & 0x7f) << 9) | ((date.getMonth() + 1) << 5) | date.getDate();
  return { time, date: dosDate };
}

class ByteWriter {
  private chunks: Uint8Array[] = [];
  private _length = 0;

  get length(): number {
    return this._length;
  }

  push(bytes: Uint8Array) {
    this.chunks.push(bytes);
    this._length += bytes.length;
  }

  uint16(value: number) {
    const b = new Uint8Array(2);
    new DataView(b.buffer).setUint16(0, value, true);
    this.push(b);
  }

  uint32(value: number) {
    const b = new Uint8Array(4);
    new DataView(b.buffer).setUint32(0, value >>> 0, true);
    this.push(b);
  }

  toBlob(type: string): Blob {
    return new Blob(this.chunks as BlobPart[], { type });
  }
}

export function textToBytes(text: string): Uint8Array {
  return new TextEncoder().encode(text);
}

export function createZipBlob(entries: ZipEntry[], mimeType: string): Blob {
  const writer = new ByteWriter();
  const { time, date } = dosDateTime(new Date());
  const central: { nameBytes: Uint8Array; crc: number; size: number; offset: number }[] = [];

  for (const entry of entries) {
    const nameBytes = textToBytes(entry.name);
    const crc = crc32(entry.data);
    central.push({ nameBytes, crc, size: entry.data.length, offset: writer.length });

    writer.uint32(0x04034b50); // local file header
    writer.uint16(20); // version needed
    writer.uint16(0x0800); // UTF-8 file names
    writer.uint16(0); // stored, no compression
    writer.uint16(time);
    writer.uint16(date);
    writer.uint32(crc);
    writer.uint32(entry.data.length);
    writer.uint32(entry.data.length);
    writer.uint16(nameBytes.length);
    writer.uint16(0); // extra field length
    writer.push(nameBytes);
    writer.push(entry.data);
  }

  const centralStart = writer.length;
  for (const item of central) {
    writer.uint32(0x02014b50); // central directory header
    writer.uint16(20); // version made by
    writer.uint16(20); // version needed
    writer.uint16(0x0800);
    writer.uint16(0);
    writer.uint16(time);
    writer.uint16(date);
    writer.uint32(item.crc);
    writer.uint32(item.size);
    writer.uint32(item.size);
    writer.uint16(item.nameBytes.length);
    writer.uint16(0); // extra
    writer.uint16(0); // comment
    writer.uint16(0); // disk number
    writer.uint16(0); // internal attributes
    writer.uint32(0); // external attributes
    writer.uint32(item.offset);
    writer.push(item.nameBytes);
  }

  const centralSize = writer.length - centralStart;
  writer.uint32(0x06054b50); // end of central directory
  writer.uint16(0);
  writer.uint16(0);
  writer.uint16(central.length);
  writer.uint16(central.length);
  writer.uint32(centralSize);
  writer.uint32(centralStart);
  writer.uint16(0); // archive comment
  return writer.toBlob(mimeType);
}
