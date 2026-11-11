import * as fs from "fs";
import * as path from "path";
import { resolveRepositoryRoot, skipTest } from "../Shared/TestInfrastructure";

/**
 * Shared helpers for the PdfRendering test suite.
 *
 * The upstream .NET suite renders workbooks with SkiaSharp and inspects the
 * resulting pixels, PDF fonts and internal chart projections. This TypeScript
 * port resolves every fixture from the repository `Input/` folder. When a
 * fixture is absent the test is reported as SKIP instead of FAIL. Tests whose
 * bodies depend on SkiaSharp rasterization or .NET-only internals are also
 * reported as SKIP with an explicit reason.
 */

export const RASTER_REASON = "requires .NET-only SkiaSharp rasterization";

export const CHART_INTERNALS_REASON =
  "requires the .NET-only chart XML parser and chart renderer internals";

export const SHAPE_INTERNALS_REASON =
  "requires the .NET-only shape geometry and pattern renderer internals";

export const PDF_INSPECTION_REASON =
  "requires .NET-only PDF font and content stream inspection helpers";

export const TEXT_LAYOUT_REASON =
  "requires .NET-only distributed text layout internals";

export const ASYNC_PDF_REASON =
  "PDF export is asynchronous in the TypeScript port";

/** Absolute path of the repository Input folder. */
export function inputRootDirectory(): string {
  return path.join(resolveRepositoryRoot(), "Input");
}

/**
 * Resolves a fixture relative to `Input/`. Aborts the current test with a
 * `missing Input/<path>` reason when the file does not exist.
 */
export function requireInput(relativePath: string): string {
  const normalized = relativePath.replace(/\\/g, "/");
  const fullPath = path.join(inputRootDirectory(), normalized);
  if (!fs.existsSync(fullPath)) {
    skipTest(`missing Input/${normalized}`);
  }

  return fullPath;
}

/** Skips when an already-resolved absolute fixture path does not exist. */
export function requireInputFile(fullPath: string): void {
  if (fs.existsSync(fullPath)) {
    return;
  }

  const relativePath = path.relative(inputRootDirectory(), fullPath).replace(/\\/g, "/");
  skipTest(`missing Input/${relativePath}`);
}

/** Resolves every fixture and skips on the first missing file. */
export function requireInputs(...relativePaths: string[]): void {
  for (const relativePath of relativePaths) {
    requireInput(relativePath);
  }
}

/**
 * Guard used by tests whose fidelity checks need .NET-only facilities.
 * The fixture check runs first so a missing input always wins as the skip
 * reason, matching the upstream input-resolution behavior.
 */
export function guardedTest(requiredInputs: string[], reason: string): void {
  requireInputs(...requiredInputs);
  skipTest(reason);
}
