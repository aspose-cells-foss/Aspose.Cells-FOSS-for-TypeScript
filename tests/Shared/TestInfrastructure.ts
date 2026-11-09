import AdmZip from "adm-zip";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import * as fs from "fs";
import * as path from "path";
import * as os from "os";

/**
 * A single named test case, mirroring Aspose.Cells_FOSS.Testing.TestCase.
 */
export class TestCase {
  constructor(
    public readonly name: string,
    public readonly body: () => void,
  ) {}
}

/**
 * Minimal console test runner mirroring Aspose.Cells_FOSS.Testing.TestRunner.
 *
 * Supports the same command line filter syntax as the .NET runner:
 *   --filter name
 *   --filter=name
 */
export class TestRunner {
  static run(suiteName: string, tests: TestCase[]): number {
    const filter = TestRunner.readNameFilterFromCommandLine();
    const selected = TestRunner.selectTests(tests, filter);

    if (filter && selected.length === 0) {
      console.log(`Running ${suiteName} (0 tests)`);
      console.log(`No tests matched filter '${filter}'.`);
      return 1;
    }

    console.log(`Running ${suiteName} (${selected.length} tests)`);
    let failed = 0;

    for (const test of selected) {
      try {
        test.body();
        console.log(`PASS ${test.name}`);
      } catch (exception) {
        failed++;
        console.log(`FAIL ${test.name}`);
        console.log(exception instanceof Error ? exception.stack : String(exception));
      }
    }

    console.log(`${suiteName}: ${selected.length - failed} passed, ${failed} failed`);
    return failed === 0 ? 0 : 1;
  }

  static runAndExit(suiteName: string, tests: TestCase[]): void {
    process.exitCode = TestRunner.run(suiteName, tests);
  }

  private static readNameFilterFromCommandLine(): string {
    const args = process.argv;
    for (let index = 0; index < args.length; index++) {
      const argument = args[index];
      if (argument.toLowerCase() === "--filter") {
        return index + 1 < args.length ? args[index + 1] : "";
      }

      const prefix = "--filter=";
      if (argument.toLowerCase().startsWith(prefix)) {
        return argument.substring(prefix.length);
      }
    }

    return "";
  }

  private static selectTests(tests: TestCase[], filter: string): TestCase[] {
    if (!filter) {
      return tests.slice();
    }

    return tests.filter((test) => test.name.toLowerCase().indexOf(filter.toLowerCase()) >= 0);
  }
}

/**
 * Assertion helpers mirroring Aspose.Cells_FOSS.Testing.AssertEx.
 */
export class AssertEx {
  static true(condition: boolean, message?: string): void {
    if (!condition) {
      throw new Error(message ?? "Expected condition to be true.");
    }
  }

  static false(condition: boolean, message?: string): void {
    if (condition) {
      throw new Error(message ?? "Expected condition to be false.");
    }
  }

  static equal<T>(expected: T, actual: T, message?: string): void {
    if (!AssertEx.valuesEqual(expected, actual)) {
      throw new Error(message ?? `Expected '${AssertEx.describe(expected)}', got '${AssertEx.describe(actual)}'.`);
    }
  }

  static notEqual<T>(expected: T, actual: T, message?: string): void {
    if (AssertEx.valuesEqual(expected, actual)) {
      throw new Error(message ?? `Expected value to differ from '${AssertEx.describe(expected)}'.`);
    }
  }

  static notNull<T>(value: T | null | undefined, message?: string): void {
    if (value === null || value === undefined) {
      throw new Error(message ?? "Expected value to be non-null.");
    }
  }

  static null(value: unknown, message?: string): void {
    if (value !== null && value !== undefined) {
      throw new Error(message ?? `Expected value to be null, got '${AssertEx.describe(value)}'.`);
    }
  }

  static contains(expectedSubstring: string, actual: string, message?: string): void {
    if (actual === null || actual === undefined || actual.indexOf(expectedSubstring) < 0) {
      throw new Error(message ?? `Expected '${actual}' to contain '${expectedSubstring}'.`);
    }
  }

  static doesNotContain(expectedSubstring: string, actual: string, message?: string): void {
    if (actual !== null && actual !== undefined && actual.indexOf(expectedSubstring) >= 0) {
      throw new Error(message ?? `Expected '${actual}' to not contain '${expectedSubstring}'.`);
    }
  }

  static throws(action: () => void, message?: string): Error {
    try {
      action();
    } catch (exception) {
      if (exception instanceof Error) {
        return exception;
      }

      return new Error(String(exception));
    }

    throw new Error(message ?? "Expected an exception to be thrown.");
  }

  private static valuesEqual(expected: unknown, actual: unknown): boolean {
    if (expected === actual) {
      return true;
    }

    if (expected instanceof Date && actual instanceof Date) {
      return expected.getTime() === actual.getTime();
    }

    if (expected === null || expected === undefined || actual === null || actual === undefined) {
      return false;
    }

    const expectedAny = expected as any;
    if (typeof expectedAny.equals === "function") {
      return Boolean(expectedAny.equals(actual));
    }

    if (typeof expected === "object" && typeof actual === "object") {
      return JSON.stringify(expected) === JSON.stringify(actual);
    }

    return false;
  }

  private static describe(value: unknown): string {
    if (value === null) {
      return "null";
    }

    if (value === undefined) {
      return "undefined";
    }

    return String(value);
  }
}

/**
 * In-memory stream replacement for System.IO.MemoryStream usage in the tests.
 * Workbook save overloads accept this object and write the produced bytes into it.
 */
export class MemoryStream {
  private _buffer: Buffer;

  constructor(buffer?: Buffer | Uint8Array) {
    this._buffer = buffer ? Buffer.from(buffer) : Buffer.alloc(0);
  }

  get length(): number {
    return this._buffer.length;
  }

  get position(): number {
    return 0;
  }

  set position(_value: number) {
    // Position is not modeled; the in-memory writer always exposes the full buffer.
  }

  toArray(): Buffer {
    return this._buffer;
  }

  toBuffer(): Buffer {
    return this._buffer;
  }

  write(buffer: Buffer | Uint8Array): void {
    this._buffer = Buffer.from(buffer);
  }

  readAll(): Buffer {
    return this._buffer;
  }
}

/**
 * Resolves the repository root by walking up from the current working directory
 * until a package.json is found. Mirrors the repository root discovery used by
 * the .NET TemporaryDirectory helper.
 */
export function resolveRepositoryRoot(startDirectory?: string): string {
  let directory = path.resolve(startDirectory ?? process.cwd());

  while (true) {
    if (fs.existsSync(path.join(directory, "package.json"))) {
      return directory;
    }

    const parent = path.dirname(directory);
    if (parent === directory) {
      return process.cwd();
    }

    directory = parent;
  }
}

/**
 * Creates an isolated output directory under <repo>/output/<suiteName>/<guid>.
 * Generated artifacts are preserved for manual inspection, matching the .NET helper.
 */
export class TemporaryDirectory {
  readonly rootPath: string;

  constructor(suiteName: string) {
    const repositoryRoot = resolveRepositoryRoot();
    const unique = `${Date.now().toString(16)}${Math.random().toString(16).slice(2)}`;
    this.rootPath = path.join(repositoryRoot, "output", suiteName, unique);
    fs.mkdirSync(this.rootPath, { recursive: true });
  }

  getPath(fileName: string): string {
    return path.join(this.rootPath, fileName);
  }

  dispose(): void {
    // Artifacts are intentionally preserved on disk for inspection.
  }
}

/**
 * Returns (and creates when needed) a shared output directory under <repo>/output/<name>.
 */
export function getOrCreateOutputDirectory(subDirectory: string): string {
  const outputDir = path.join(resolveRepositoryRoot(), "output", subDirectory);
  fs.mkdirSync(outputDir, { recursive: true });
  return outputDir;
}

/**
 * Reads a zip package entry as UTF-8 text. Mirrors the .NET test helper that
 * opened the saved XLSX with System.IO.Compression.ZipArchive.
 */
export class ZipPackageHelper {
  static readEntryText(packagePath: string, entryPath: string): string {
    const archive = new AdmZip(packagePath);
    const entry = archive.getEntry(entryPath.replace(/^\/+/, ""));
    if (!entry) {
      throw new Error("Missing zip entry: " + entryPath);
    }

    return entry.getData().toString("utf8");
  }

  static readEntryTextFromBuffer(buffer: Buffer | Uint8Array, entryPath: string): string {
    const archive = new AdmZip(Buffer.from(buffer));
    const entry = archive.getEntry(entryPath.replace(/^\/+/, ""));
    if (!entry) {
      throw new Error("Missing zip entry: " + entryPath);
    }

    return entry.getData().toString("utf8");
  }

  static entryExists(packagePath: string, entryPath: string): boolean {
    const archive = new AdmZip(packagePath);
    return archive.getEntry(entryPath.replace(/^\/+/, "")) !== null;
  }

  static listEntries(packagePath: string): string[] {
    const archive = new AdmZip(packagePath);
    return archive.getEntries().map((entry) => entry.entryName);
  }

  /** Replaces one entry's text content using the supplied rewrite delegate. */
  static rewriteEntryText(packagePath: string, entryPath: string, rewrite: (content: string) => string): void {
    const archive = new AdmZip(packagePath);
    const normalized = entryPath.replace(/^\/+/, "");
    const entry = archive.getEntry(normalized);
    if (!entry) {
      throw new Error("Missing zip entry: " + entryPath);
    }

    const replacement = rewrite(entry.getData().toString("utf8"));
    archive.updateFile(normalized, Buffer.from(replacement, "utf8"));
    archive.writeZip(packagePath);
  }

  /** Parses an XML entry, applies the DOM mutation, and writes the serialized document back. */
  static rewriteXmlEntry(packagePath: string, entryPath: string, mutate: (document: Document) => void): void {
    ZipPackageHelper.rewriteEntryText(packagePath, entryPath, (content) => {
      const declaration = content.trimStart().startsWith("<?xml")
        ? content.substring(0, content.indexOf("?>") + 2)
        : "";
      const document = new DOMParser().parseFromString(content, "text/xml");
      mutate(document);
      const serialized = new XMLSerializer().serializeToString(document);
      return declaration.length > 0 && !serialized.startsWith("<?xml") ? declaration + serialized : serialized;
    });
  }

  /** Moves an entry to a new path inside the same zip package. */
  static moveEntry(packagePath: string, sourceEntryPath: string, destinationEntryPath: string): void {
    const archive = new AdmZip(packagePath);
    const normalizedSource = sourceEntryPath.replace(/^\/+/, "");
    const normalizedDestination = destinationEntryPath.replace(/^\/+/, "");
    const sourceEntry = archive.getEntry(normalizedSource);
    if (!sourceEntry) {
      throw new Error("Missing zip entry: " + sourceEntryPath);
    }

    if (archive.getEntry(normalizedDestination)) {
      archive.deleteFile(normalizedDestination);
    }

    archive.addFile(normalizedDestination, sourceEntry.getData());
    archive.deleteFile(normalizedSource);
    archive.writeZip(packagePath);
  }
}
