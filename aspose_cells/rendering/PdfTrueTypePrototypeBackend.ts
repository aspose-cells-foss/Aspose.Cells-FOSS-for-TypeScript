import type { PdfTextFontBackend } from "./PdfTextFontBackend";
import { PdfType3TextFontBackend } from "./PdfType3TextFontBackend";
import { PdfTrueTypePrototypeUsage, type ITrueTypeFont } from "./PdfTrueTypePrototypeUsage";
import { PdfTrueTypePrototypeFontBuilder } from "./PdfTrueTypePrototypeFontBuilder";
import { PdfEncodedTextSegment } from "./PdfEncodedTextSegment";
import type { PdfType3FontUsage } from "./PdfType3FontUsage";

/**
 * Prototype TrueType backend – falls back to Type3 for any run whose text
 * contains non-ASCII codepoints or whose typeface doesn't expose TrueType
 * tables.  ASCII-heavy worksheets thus avoid per-glyph Type3 streams.
 */
export class PdfTrueTypePrototypeBackend implements PdfTextFontBackend {
  private _fallback: PdfType3TextFontBackend;
  private _trueTypeUsages: PdfTrueTypePrototypeUsage[] = [];
  private _builder = new PdfTrueTypePrototypeFontBuilder();
  private _fontProvider: (typeface: unknown) => ITrueTypeFont | null;
  private _createFallbackUsage: (typeface: unknown, usePositiveY: boolean) => PdfType3FontUsage;

  constructor(
    fontProvider: (typeface: unknown) => ITrueTypeFont | null,
    createFallbackUsage: (typeface: unknown, usePositiveY: boolean) => PdfType3FontUsage,
  ) {
    this._fontProvider = fontProvider;
    this._createFallbackUsage = createFallbackUsage;
    this._fallback = new PdfType3TextFontBackend(createFallbackUsage);
  }

  Initialize(runs: { text: string; typeface: unknown; usePositiveYFontMatrix: boolean }[]): void {
    const fallbackRuns: { text: string; typeface: unknown; usePositiveYFontMatrix: boolean }[] = [];
    for (const run of runs) {
      if (this.isEligible(run)) {
        const usage = this.usageFor(run.typeface);
        usage.AddText(run.text);
      } else {
        fallbackRuns.push(run);
      }
    }
    this._fallback.Initialize(fallbackRuns);
  }

  BuildObjects(nextObjectId: number): { objects: Map<number, Uint8Array>; nextObjectId: number } {
    const result = this._builder.BuildObjects(this._trueTypeUsages, nextObjectId);
    const fb = this._fallback.BuildObjects(result.nextObjectId);
    for (const [k, v] of fb.objects) result.objects.set(k, v);
    result.nextObjectId = fb.nextObjectId;
    return result;
  }

  BuildFontResourceText(): string {
    const parts: string[] = ["/Font <<"];
    for (const usage of this._trueTypeUsages) {
      parts.push(` /${usage.ResourceName} ${usage.FontObjectId} 0 R`);
    }
    const fbText = this._fallback.BuildFontResourceText();
    // Extract the inner " /X N 0 R /Y M 0 R" of "/Font << ... >>" and splice it in.
    const innerMatch = fbText.match(/<<([\s\S]*)>>/);
    if (innerMatch) parts.push(innerMatch[1]);
    parts.push(" >>");
    return parts.join("");
  }

  EncodeRun(run: { text: string; typeface: unknown; usePositiveYFontMatrix: boolean }): PdfEncodedTextSegment[] {
    if (!this.isEligible(run)) return this._fallback.EncodeRun(run);
    const usage = this.usageFor(run.typeface);
    const seg = new PdfEncodedTextSegment();
    seg.ResourceName = usage.ResourceName;
    seg.EncodedBytes = usage.EncodeText(run.text);
    seg.StartXPt = 0;
    seg.RequiresTopDownTextFlip = true;
    return [seg];
  }

  private isEligible(run: { text: string; typeface: unknown; usePositiveYFontMatrix: boolean }): boolean {
    if (!run?.text) return false;
    // ASCII only – TrueType prototype emits WinAnsiEncoding.
    for (let i = 0; i < run.text.length; i++) {
      const cp = run.text.charCodeAt(i);
      if (cp < 32 || cp > 126) return false;
    }
    return this._fontProvider(run.typeface) !== null;
  }

  private usageFor(typeface: unknown): PdfTrueTypePrototypeUsage {
    for (const u of this._trueTypeUsages) if (u.Typeface === typeface) return u;
    const font = this._fontProvider(typeface);
    if (!font) throw new Error("PdfTrueTypePrototypeBackend: missing font provider for typeface");
    const usage = new PdfTrueTypePrototypeUsage(typeface, font);
    this._trueTypeUsages.push(usage);
    return usage;
  }

  Dispose(): void {
    this._fallback.Dispose();
    for (const u of this._trueTypeUsages) u.Dispose();
    this._trueTypeUsages = [];
  }
}
