import { readFile } from "fs/promises";
import { ZipReader, BlobReader, TextWriter, BlobWriter } from "@zip.js/zip.js";
import * as crypto from "crypto";

export type CellValue = string | number | boolean | Date | null;

export interface CellCoordinates {
  row: number;
  col: number;
}

export interface CellRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

export function colToIndex(col: string): number {
  let index = 0;
  for (let i = 0; i < col.length; i++) {
    index = index * 26 + (col.charCodeAt(i) - 64);
  }
  return index - 1;
}

export function indexToCol(index: number): string {
  let col = "";
  let n = index + 1;
  while (n > 0) {
    const mod = (n - 1) % 26;
    col = String.fromCharCode(65 + mod) + col;
    n = Math.floor((n - 1) / 26);
  }
  return col;
}

export function cellRef(row: number, col: number): string {
  return `${indexToCol(col)}${row + 1}`;
}

export function parseCellRef(ref: string): CellCoordinates {
  const match = /^([A-Z]+)(\d+)$/.exec(ref);
  if (!match) throw new Error(`Invalid cell reference: ${ref}`);
  return {
    row: parseInt(match[2] ?? "0", 10) - 1,
    col: colToIndex(match[1] ?? "A"),
  };
}

export function parseRange(ref: string): CellRange {
  const [start, end] = ref.split(":");
  if (!end) {
    const coords = parseCellRef(start);
    return {
      startRow: coords.row,
      startCol: coords.col,
      endRow: coords.row,
      endCol: coords.col,
    };
  }
  const startCoords = parseCellRef(start);
  const endCoords = parseCellRef(end);
  return {
    startRow: startCoords.row,
    startCol: startCoords.col,
    endRow: endCoords.row,
    endCol: endCoords.col,
  };
}

export async function readZipEntry(
  zip: InstanceType<typeof ZipReader>,
  path: string,
): Promise<string | null> {
  const entries = await zip.getEntries();
  const entry = entries.find((e: { filename: string }) => e.filename === path);
  if (!entry) return null;
  const writer = new TextWriter();
  const data = await (entry as any).getData(writer);
  return data ?? null;
}

export async function openZip(filePath: string) {
  const buffer = await readFile(filePath);
  const blob = new Blob([buffer]);
  return new ZipReader(new BlobReader(blob));
}

export function escapeXml(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

export function unescapeXml(str: string): string {
  return str
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

export function generateUuid(): string {
  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
    const r = (Math.random() * 16) | 0;
    const v = c === "x" ? r : (r & 0x3) | 0x8;
    return v.toString(16);
  });
}

export async function createZipWriter() {
  const { ZipWriter } = await import("@zip.js/zip.js");
  return new ZipWriter(new BlobWriter());
}

export async function addZipEntry(zip: any, path: string, content: string) {
  const { TextReader } = await import("@zip.js/zip.js");
  await zip.add(path, new TextReader(content));
}

export async function finalizeZip(zip: any): Promise<Uint8Array> {
  const blob = await zip.close();
  return new Uint8Array(await blob.arrayBuffer());
}

export function deriveKey(password: string, salt: Uint8Array): Buffer {
  return crypto.pbkdf2Sync(password, salt, 100000, 32, "sha512");
}

export function encryptData(data: Uint8Array, key: Buffer): Uint8Array {
  const iv = crypto.randomBytes(16);
  const cipher = crypto.createCipheriv("aes-256-cbc", key, iv);
  const encrypted = Buffer.concat([cipher.update(data), cipher.final()]);
  return Buffer.concat([iv, encrypted]);
}

export function decryptData(data: Uint8Array, key: Buffer): Uint8Array {
  const iv = data.subarray(0, 16);
  const encrypted = data.subarray(16);
  const decipher = crypto.createDecipheriv("aes-256-cbc", key, iv);
  return Buffer.concat([decipher.update(encrypted), decipher.final()]);
}
