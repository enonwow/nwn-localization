import { describe, expect, it } from "vitest";

import { buildZipArchive } from "../zip";

function readUint16(bytes: Uint8Array, offset: number): number {
  return bytes[offset] | (bytes[offset + 1] << 8);
}

function readUint32(bytes: Uint8Array, offset: number): number {
  return (
    bytes[offset]
    | (bytes[offset + 1] << 8)
    | (bytes[offset + 2] << 16)
    | (bytes[offset + 3] << 24)
  ) >>> 0;
}

function findEocdOffset(bytes: Uint8Array): number {
  for (let i = bytes.length - 22; i >= 0; i -= 1) {
    if (readUint32(bytes, i) === 0x06054b50) {
      return i;
    }
  }
  return -1;
}

type ParsedZipEntry = {
  name: string;
  data: Uint8Array;
};

function parseStoredZipEntries(zipBytes: Uint8Array): ParsedZipEntry[] {
  const eocdOffset = findEocdOffset(zipBytes);
  if (eocdOffset < 0) {
    throw new Error("EOCD signature not found.");
  }

  const totalEntries = readUint16(zipBytes, eocdOffset + 10);
  const centralDirOffset = readUint32(zipBytes, eocdOffset + 16);
  const decoder = new TextDecoder();
  const entries: ParsedZipEntry[] = [];
  let cursor = centralDirOffset;

  for (let i = 0; i < totalEntries; i += 1) {
    expect(readUint32(zipBytes, cursor)).toBe(0x02014b50);
    const nameLength = readUint16(zipBytes, cursor + 28);
    const extraLength = readUint16(zipBytes, cursor + 30);
    const commentLength = readUint16(zipBytes, cursor + 32);
    const localHeaderOffset = readUint32(zipBytes, cursor + 42);
    const nameStart = cursor + 46;
    const nameEnd = nameStart + nameLength;
    const name = decoder.decode(zipBytes.slice(nameStart, nameEnd));

    expect(readUint32(zipBytes, localHeaderOffset)).toBe(0x04034b50);
    const localNameLength = readUint16(zipBytes, localHeaderOffset + 26);
    const localExtraLength = readUint16(zipBytes, localHeaderOffset + 28);
    const compressedSize = readUint32(zipBytes, localHeaderOffset + 18);
    const dataStart = localHeaderOffset + 30 + localNameLength + localExtraLength;
    const dataEnd = dataStart + compressedSize;
    entries.push({
      name,
      data: zipBytes.slice(dataStart, dataEnd),
    });

    cursor = nameEnd + extraLength + commentLength;
  }

  return entries;
}

describe("ZIP archive builder", () => {
  it("builds valid empty archive", () => {
    const zip = buildZipArchive([]);
    expect(zip).toHaveLength(22);
    expect(readUint32(zip, 0)).toBe(0x06054b50);
    expect(readUint16(zip, 8)).toBe(0);
    expect(readUint16(zip, 10)).toBe(0);
    expect(parseStoredZipEntries(zip)).toEqual([]);
  });

  it("stores multiple files and normalizes path separators", () => {
    const files = [
      { name: "tlc\\en\\dialog.tlk", bytes: new Uint8Array([1, 2, 3, 4]) },
      { name: "tlc/pl/dialogf.tlk", bytes: new Uint8Array([]) },
    ];

    const zip = buildZipArchive(files);
    const entries = parseStoredZipEntries(zip);

    expect(entries).toHaveLength(2);
    expect(entries[0].name).toBe("tlc/en/dialog.tlk");
    expect(entries[1].name).toBe("tlc/pl/dialogf.tlk");
    expect(Array.from(entries[0].data)).toEqual([1, 2, 3, 4]);
    expect(entries[1].data.byteLength).toBe(0);
  });
});

