/**
 * Optional PDF text post-processing session.
 * Mirrors the .NET PdfTextDocumentSession concept - stores text runs for Type3/TrueType
 * font subsetting after the initial PDF generation.
 */
export class PdfTextRunRecord {
  constructor(
    public text: string,
    public fontName: string,
    public fontSize: number,
    public xPt: number,
    public yPt: number,
    public color: string,
  ) {}
}

export class PdfTextDocumentSession {
  runs: PdfTextRunRecord[] = [];

  get hasRuns(): boolean {
    return this.runs.length > 0;
  }

  addRun(run: PdfTextRunRecord): void {
    this.runs.push(run);
  }
}
