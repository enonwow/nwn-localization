export type ZipInputFile = {
  name: string;
  bytes: Uint8Array;
};

type ZipEntry = {
  nameBytes: Uint8Array;
  bytes: Uint8Array;
  crc32: number;
  offset: number;
  dosTime: number;
  dosDate: number;
};

const CRC32_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i += 1) {
    let c = i;
    for (let j = 0; j < 8; j += 1) {
      c = (c & 1) !== 0 ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
    }
    table[i] = c >>> 0;
  }
  return table;
})();

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (let i = 0; i < bytes.length; i += 1) {
    crc = CRC32_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function toDosDateTime(date: Date): { dosDate: number; dosTime: number } {
  const year = Math.min(2107, Math.max(1980, date.getFullYear()));
  const month = date.getMonth() + 1;
  const day = date.getDate();
  const hours = date.getHours();
  const minutes = date.getMinutes();
  const seconds = Math.floor(date.getSeconds() / 2);
  const dosTime = ((hours & 0x1f) << 11) | ((minutes & 0x3f) << 5) | (seconds & 0x1f);
  const dosDate = (((year - 1980) & 0x7f) << 9) | ((month & 0x0f) << 5) | (day & 0x1f);
  return { dosDate, dosTime };
}

function writeUint16(target: Uint8Array, offset: number, value: number): void {
  target[offset] = value & 0xff;
  target[offset + 1] = (value >>> 8) & 0xff;
}

function writeUint32(target: Uint8Array, offset: number, value: number): void {
  target[offset] = value & 0xff;
  target[offset + 1] = (value >>> 8) & 0xff;
  target[offset + 2] = (value >>> 16) & 0xff;
  target[offset + 3] = (value >>> 24) & 0xff;
}

function concatChunks(chunks: Uint8Array[]): Uint8Array {
  const total = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const result = new Uint8Array(total);
  let offset = 0;
  for (let i = 0; i < chunks.length; i += 1) {
    const chunk = chunks[i];
    result.set(chunk, offset);
    offset += chunk.length;
  }
  return result;
}

export function buildZipArchive(files: ZipInputFile[]): Uint8Array {
  const encoder = new TextEncoder();
  const now = new Date();
  const { dosDate, dosTime } = toDosDateTime(now);
  const entries: ZipEntry[] = [];
  const localChunks: Uint8Array[] = [];
  let localOffset = 0;

  for (let i = 0; i < files.length; i += 1) {
    const file = files[i];
    const safeName = String(file.name || "").replace(/\\/g, "/");
    const nameBytes = encoder.encode(safeName);
    const data = file.bytes;
    const fileCrc = crc32(data);

    const localHeader = new Uint8Array(30 + nameBytes.length);
    writeUint32(localHeader, 0, 0x04034b50);
    writeUint16(localHeader, 4, 20);
    writeUint16(localHeader, 6, 0x0800); // UTF-8
    writeUint16(localHeader, 8, 0); // stored (no compression)
    writeUint16(localHeader, 10, dosTime);
    writeUint16(localHeader, 12, dosDate);
    writeUint32(localHeader, 14, fileCrc);
    writeUint32(localHeader, 18, data.length);
    writeUint32(localHeader, 22, data.length);
    writeUint16(localHeader, 26, nameBytes.length);
    writeUint16(localHeader, 28, 0);
    localHeader.set(nameBytes, 30);

    localChunks.push(localHeader, data);
    entries.push({
      nameBytes,
      bytes: data,
      crc32: fileCrc,
      offset: localOffset,
      dosDate,
      dosTime,
    });
    localOffset += localHeader.length + data.length;
  }

  const centralStart = localOffset;
  const centralChunks: Uint8Array[] = [];
  let centralSize = 0;
  for (let i = 0; i < entries.length; i += 1) {
    const entry = entries[i];
    const centralHeader = new Uint8Array(46 + entry.nameBytes.length);
    writeUint32(centralHeader, 0, 0x02014b50);
    writeUint16(centralHeader, 4, 20); // version made by
    writeUint16(centralHeader, 6, 20); // version needed
    writeUint16(centralHeader, 8, 0x0800); // UTF-8
    writeUint16(centralHeader, 10, 0); // stored
    writeUint16(centralHeader, 12, entry.dosTime);
    writeUint16(centralHeader, 14, entry.dosDate);
    writeUint32(centralHeader, 16, entry.crc32);
    writeUint32(centralHeader, 20, entry.bytes.length);
    writeUint32(centralHeader, 24, entry.bytes.length);
    writeUint16(centralHeader, 28, entry.nameBytes.length);
    writeUint16(centralHeader, 30, 0); // extra len
    writeUint16(centralHeader, 32, 0); // comment len
    writeUint16(centralHeader, 34, 0); // disk number start
    writeUint16(centralHeader, 36, 0); // internal attrs
    writeUint32(centralHeader, 38, 0); // external attrs
    writeUint32(centralHeader, 42, entry.offset);
    centralHeader.set(entry.nameBytes, 46);
    centralChunks.push(centralHeader);
    centralSize += centralHeader.length;
  }

  const eocd = new Uint8Array(22);
  writeUint32(eocd, 0, 0x06054b50);
  writeUint16(eocd, 4, 0); // disk number
  writeUint16(eocd, 6, 0); // central dir start disk
  writeUint16(eocd, 8, entries.length);
  writeUint16(eocd, 10, entries.length);
  writeUint32(eocd, 12, centralSize);
  writeUint32(eocd, 16, centralStart);
  writeUint16(eocd, 20, 0); // comment len

  return concatChunks([...localChunks, ...centralChunks, eocd]);
}

