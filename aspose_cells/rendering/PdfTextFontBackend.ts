import type { PdfType3FontUsage } from "./PdfType3FontUsage";
import type { PdfType3FontSubset } from "./PdfType3FontSubset";
import type { PdfType3FontBuilder } from "./PdfType3FontBuilder";
import { PdfFontUsageRegistry } from "./PdfFontUsageRegistry";
import type { PdfEncodedTextSegment } from "./PdfEncodedTextSegment";

/**
 * Backend for producing font objects and encoded text segments for PDF.
 * Mirrors the .NET `PdfTextFontBackend` interface.
 */
export interface PdfTextFontBackend {
  Initialize(runs: { text: string; typeface: unknown; usePositiveYFontMatrix: boolean }[]): void;
  BuildObjects(nextObjectId: number): { objects: Map<number, Uint8Array>; nextObjectId: number };
  BuildFontResourceText(): string;
  EncodeRun(run: { text: string; typeface: unknown; usePositiveYFontMatrix: boolean }): PdfEncodedTextSegment[];
}
