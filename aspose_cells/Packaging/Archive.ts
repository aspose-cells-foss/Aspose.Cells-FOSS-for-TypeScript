import AdmZip from "adm-zip";

/**
 * Output archive writer. Mirrors the ZipArchive write access used by the .NET serializer.
 * Entries are stored in insertion order and UTF-8 encoded without a BOM.
 */
export class ZipWriter {
  private readonly _zip: AdmZip;

  constructor() {
    this._zip = new AdmZip();
  }

  addText(path: string, content: string): void {
    this._zip.addFile(normalizeEntryPath(path), Buffer.from(content, "utf8"));
  }

  addBinary(path: string, content: Buffer | Uint8Array): void {
    this._zip.addFile(normalizeEntryPath(path), Buffer.from(content));
  }

  toBuffer(): Buffer {
    return Buffer.from(this._zip.toBuffer());
  }
}

/**
 * Input archive reader. Mirrors ZipArchive read access, including missing-entry behavior.
 */
export class ZipReader {
  private readonly _zip: AdmZip;

  constructor(data: Buffer | Uint8Array) {
    this._zip = new AdmZip(Buffer.from(data));
  }

  static fromFile(filePath: string): ZipReader {
    return new ZipReader(Buffer.from(new AdmZip(filePath).toBuffer()));
  }

  hasEntry(path: string): boolean {
    return this._zip.getEntry(normalizeEntryPath(path)) !== null;
  }

  readText(path: string): string {
    const entry = this._zip.getEntry(normalizeEntryPath(path));
    if (!entry) {
      throw new MissingPackageEntryException(path);
    }

    return entry.getData().toString("utf8");
  }

  tryReadText(path: string): string | null {
    const entry = this._zip.getEntry(normalizeEntryPath(path));
    return entry ? entry.getData().toString("utf8") : null;
  }

  readBinary(path: string): Buffer {
    const entry = this._zip.getEntry(normalizeEntryPath(path));
    if (!entry) {
      throw new MissingPackageEntryException(path);
    }

    return Buffer.from(entry.getData());
  }

  tryReadBinary(path: string): Buffer | null {
    const entry = this._zip.getEntry(normalizeEntryPath(path));
    return entry ? Buffer.from(entry.getData()) : null;
  }

  listEntries(): string[] {
    return this._zip.getEntries().map((entry) => entry.entryName);
  }
}

export class MissingPackageEntryException extends Error {
  constructor(entryPath: string) {
    super(`Missing zip entry '${entryPath}'.`);
    this.name = "MissingPackageEntryException";
  }
}

export function normalizeEntryPath(entryPath: string): string {
  return entryPath.replace(/^\/+/, "").replace(/\\/g, "/");
}
