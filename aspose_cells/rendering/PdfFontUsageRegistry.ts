import type { PdfType3FontUsage } from "./PdfType3FontUsage";

/**
 * Registry of Type3 font usages, keyed by (font family, matrix orientation).
 * Collects glyph coverage across all text runs and delegates subset creation.
 */
export class PdfFontUsageRegistry {
  Usages: PdfType3FontUsage[] = [];

  private _createUsage: (typeface: unknown, usePositiveYMatrix: boolean) => PdfType3FontUsage;

  constructor(createUsage: (typeface: unknown, usePositiveYMatrix: boolean) => PdfType3FontUsage) {
    this._createUsage = createUsage;
  }

  /** Register text from every run so we know which glyphs to embed. */
  AddRuns(runs: { text: string; typeface: unknown; usePositiveYFontMatrix: boolean }[]): void {
    for (const run of runs) {
      const usage = this.UsageFor(run.typeface, run.usePositiveYFontMatrix);
      usage.AddText(run.text);
    }
  }

  /** Assign resource indices and finalise each subset (max 255 glyphs per subset). */
  FinalizeSubsets(): void {
    let resourceIndex = 1;
    for (const usage of this.Usages) {
      usage.FinalizeSubsets(resourceIndex);
      // resourceIndex is mutated inside FinalizeSubsets
      resourceIndex = (usage as unknown as { _nextResourceIndex?: number })._nextResourceIndex ?? resourceIndex;
    }
  }

  UsageFor(typeface: unknown, usePositiveYFontMatrix: boolean): PdfType3FontUsage {
    for (const u of this.Usages) {
      if (u.Typeface === typeface && u.UsePositiveYFontMatrix === usePositiveYFontMatrix) return u;
    }
    const usage = this._createUsage(typeface, usePositiveYFontMatrix);
    this.Usages.push(usage);
    return usage;
  }

  Dispose(): void {
    for (const u of this.Usages) u.Dispose();
    this.Usages = [];
  }
}
