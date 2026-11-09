import { existsSync, readdirSync, statSync } from "fs";
import { extname, isAbsolute, join, resolve } from "path";
import { PaperSizeType, PdfSaveOptions, Workbook } from "../../aspose_cells";

class PrintOverrides {
  private paper: PaperSizeType | null = null;
  private scale: number | null = null;
  private readonly margins: Array<number | null> = [null, null, null, null];

  public Read(args: string[], cursor: { index: number }): boolean {
    const name = args[cursor.index];
    const marginNames = ["--margin-left-cm", "--margin-right-cm", "--margin-top-cm", "--margin-bottom-cm"];
    const marginIndex = marginNames.indexOf(name);
    if (name !== "--paper-size" && name !== "--scale" && marginIndex < 0) {
      return false;
    }

    if (cursor.index + 1 >= args.length) {
      throw new Error(name + " requires a value.");
    }

    const value = args[++cursor.index];
    if (name === "--paper-size") {
      const key = "Paper" + value;
      const match = Object.keys(PaperSizeType).find((candidate) => candidate.toLowerCase() === key.toLowerCase());
      if (match === undefined || typeof (PaperSizeType as any)[match] !== "number") {
        throw new Error("Unsupported paper size: " + value);
      }

      this.paper = (PaperSizeType as any)[match] as PaperSizeType;
    } else if (name === "--scale") {
      const scale = Number(value);
      if (!/^[+-]?\d+$/.test(value.trim()) || !Number.isInteger(scale) || scale < 10 || scale > 400) {
        throw new Error("--scale must be between 10 and 400.");
      }

      this.scale = scale;
    } else {
      const margin = Number(value);
      if (!/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(value.trim()) || Number.isNaN(margin) || !Number.isFinite(margin) || margin < 0) {
        throw new Error(name + " requires a non-negative finite number.");
      }

      this.margins[marginIndex] = margin;
    }

    return true;
  }

  public Apply(workbook: Workbook): void {
    for (let index = 0; index < workbook.Worksheets.Count; index++) {
      const setup = workbook.Worksheets[index].PageSetup;
      if (this.paper !== null) {
        setup.PaperSize = this.paper;
      }

      if (this.scale !== null) {
        setup.Scale = this.scale;
      }

      if (this.margins[0] !== null) {
        setup.LeftMargin = this.margins[0]!;
      }

      if (this.margins[1] !== null) {
        setup.RightMargin = this.margins[1]!;
      }

      if (this.margins[2] !== null) {
        setup.TopMargin = this.margins[2]!;
      }

      if (this.margins[3] !== null) {
        setup.BottomMargin = this.margins[3]!;
      }

      console.log("Print settings: paper=" + PaperSizeType[setup.PaperSize] + ", scale=" + setup.Scale
        + ", margins (cm)=" + setup.LeftMargin.toString() + ","
        + setup.RightMargin.toString() + ","
        + setup.TopMargin.toString() + ","
        + setup.BottomMargin.toString());
    }
  }
}

function resolveRepositoryRoot(): string {
  let directory = resolve(process.cwd());
  while (true) {
    const inputDirectory = join(directory, "samples", "files");
    if (existsSync(inputDirectory)) {
      return directory;
    }

    const parent = resolve(directory, "..");
    if (parent === directory) {
      throw new Error("Could not locate the repository root containing the samples/files directory.");
    }

    directory = parent;
  }
}

function resolveInputPaths(repositoryRoot: string, args: string[]): string[] {
  if (args !== null && args.length > 0) {
    return resolveExplicitPaths(repositoryRoot, args);
  }

  return resolveAllInputPaths(repositoryRoot);
}

function resolveExplicitPaths(repositoryRoot: string, args: string[]): string[] {
  const results: string[] = [];
  for (const inputArgument of args) {
    if (inputArgument === null || inputArgument === undefined || inputArgument.trim().length === 0) {
      continue;
    }

    results.push(resolveExplicitPath(repositoryRoot, inputArgument));
  }

  if (results.length === 0) {
    throw new Error("At least one non-empty .xlsx path must be provided.");
  }

  return results;
}

function resolveAllInputPaths(repositoryRoot: string): string[] {
  const inputDirectory = join(repositoryRoot, "samples", "files");
  const files = collectXlsxFiles(inputDirectory, true);
  files.sort((left, right) => {
    const loweredLeft = left.toLowerCase();
    const loweredRight = right.toLowerCase();
    return loweredLeft < loweredRight ? -1 : loweredLeft > loweredRight ? 1 : 0;
  });

  if (files.length === 0) {
    throw new Error("Could not find any .xlsx files under the samples/files directory.");
  }

  return files;
}

function collectXlsxFiles(directory: string, recursive: boolean): string[] {
  const results: string[] = [];
  for (const entry of readdirSync(directory)) {
    const fullPath = join(directory, entry);
    if (statSync(fullPath).isDirectory()) {
      if (recursive) {
        results.push(...collectXlsxFiles(fullPath, true));
      }

      continue;
    }

    if (extname(fullPath).toLowerCase() === ".xlsx") {
      results.push(fullPath);
    }
  }

  return results;
}

function resolveExplicitPath(repositoryRoot: string, inputArgument: string): string {
  let candidatePath = inputArgument;
  if (!isAbsolute(candidatePath)) {
    candidatePath = join(repositoryRoot, candidatePath);
    if (!existsSync(candidatePath)) {
      candidatePath = join(repositoryRoot, "samples", "files", inputArgument);
    }
  }

  if (!existsSync(candidatePath) || !statSync(candidatePath).isFile()) {
    throw new Error("The specified XLSX file does not exist.");
  }

  if (extname(candidatePath).toLowerCase() !== ".xlsx") {
    throw new Error("The input file must have an .xlsx extension.");
  }

  return resolve(candidatePath);
}

function parseConditionalDate(value: string): Date {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (match === null) {
    throw new Error("--conditional-date requires a yyyy-MM-dd date.");
  }

  const date = new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
  if (
    date.getUTCFullYear() !== Number(match[1]) ||
    date.getUTCMonth() !== Number(match[2]) - 1 ||
    date.getUTCDate() !== Number(match[3])
  ) {
    throw new Error("--conditional-date requires a yyyy-MM-dd date.");
  }

  return date;
}

const repositoryRoot = resolveRepositoryRoot();
let evaluationDate: Date | null = null;
const printOverrides = new PrintOverrides();
const inputArguments: string[] = [];
const args = process.argv.slice(2);

for (const cursor = { index: 0 }; cursor.index < args.length; cursor.index++) {
  if (printOverrides.Read(args, cursor)) {
    continue;
  }

  if (args[cursor.index] === "--conditional-date") {
    if (cursor.index + 1 >= args.length) {
      throw new Error("--conditional-date requires a yyyy-MM-dd date.");
    }

    evaluationDate = parseConditionalDate(args[cursor.index + 1]);
    cursor.index++;
    continue;
  }

  inputArguments.push(args[cursor.index]);
}

const inputPaths = resolveInputPaths(repositoryRoot, inputArguments);

for (const inputPath of inputPaths) {
  const outputPath = inputPath.substring(0, inputPath.length - extname(inputPath).length) + ".pdf";

  const workbook = new Workbook(inputPath);
  printOverrides.Apply(workbook);
  const pdfOptions = new PdfSaveOptions();
  pdfOptions.ConditionalFormattingEvaluationDate = evaluationDate;
  workbook.Save(outputPath, pdfOptions);

  console.log("Input: " + inputPath);
  console.log("Output: " + outputPath);
}
