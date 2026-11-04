/**
 * Parses the limited subset of Excel formula expressions that the .NET port
 * supported in ConditionalFormatting rules: a single cell reference on the
 * left, a comparison operator, and a scalar threshold.  Example:
 *     A1 > 5.5
 *     $B$3 = "hello"
 *     (row 10, col 5) vs threshold when formula is `D5 > 3` on area D5:E20.
 */

interface ParsedFormula {
  anchorRow: number;
  anchorColumn: number;
  referenceRow: number;
  referenceColumn: number;
  absoluteRow: boolean;
  absoluteColumn: boolean;
  operator: ">" | ">=" | "<" | "<=" | "=" | "<>";
  threshold: number;
}

/** Tiny Excel A1 → (rowIndex, columnIndex) helper. */
function parseCellAddress(token: string): { row: number; col: number; absCol: boolean; absRow: boolean } | null {
  const colMatch = token.match(/^\$?([A-Z]{1,3})\$?(\d+)$/);
  if (!colMatch) return null;
  const colLetters = colMatch[1].toUpperCase();
  let col = 0;
  for (let i = 0; i < colLetters.length; i++) col = col * 26 + (colLetters.charCodeAt(i) - 64);
  col -= 1;
  const row = parseInt(colMatch[2], 10) - 1;
  const absCol = token.startsWith("$");
  const absRow = token.includes("$", colMatch[0].length - 1);
  return { row, col, absCol, absRow };
}

const COMPARISON_RE =
  /^\s*(\$?[A-Z]{1,3}\$?[1-9]\d*)\s*(>=|<=|<>|>|<|=)\s*([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?)\s*$/;

export class ConditionalFormatExpressionMatcher {
  private _anchorRow: number;
  private _anchorColumn: number;
  private _referenceRow: number;
  private _referenceColumn: number;
  private _absoluteRow: boolean;
  private _absoluteColumn: boolean;
  private _operator: ParsedFormula["operator"];
  private _threshold: number;

  private constructor(
    parsed: ParsedFormula,
    anchorRow: number,
    anchorColumn: number,
  ) {
    this._anchorRow = anchorRow;
    this._anchorColumn = anchorColumn;
    this._referenceRow = parsed.referenceRow;
    this._referenceColumn = parsed.referenceColumn;
    this._absoluteRow = parsed.absoluteRow;
    this._absoluteColumn = parsed.absoluteColumn;
    this._operator = parsed.operator;
    this._threshold = parsed.threshold;
  }

  static TryCreate(
    anchorRow: number,
    anchorColumn: number,
    formula: string | undefined | null,
  ): ConditionalFormatExpressionMatcher | null {
    const expr = (formula ?? "").trim();
    let body = expr;
    if (body.startsWith("=")) body = body.substring(1);
    const m = COMPARISON_RE.exec(body);
    if (!m) return null;
    const threshold = parseFloat(m[3]);
    if (!Number.isFinite(threshold)) return null;
    const clean = m[1].replace(/\$/g, "");
    const parsed = parseCellAddress(clean);
    if (!parsed) return null;
    return new ConditionalFormatExpressionMatcher(
      {
        anchorRow,
        anchorColumn,
        referenceRow: parsed.row,
        referenceColumn: parsed.col,
        absoluteRow: parsed.absRow,
        absoluteColumn: parsed.absCol,
        operator: m[2] as ParsedFormula["operator"],
        threshold,
      },
      anchorRow,
      anchorColumn,
    );
  }

  /**
   * Resolve the reference for (row, column) and compare against `value`.
   * This matches Excel's formula-evaluation semantics for a single-cell ref:
   * `=D5>3` on range D5:E15 evaluates D5 for (D,5) and D6 for (D,6).
   */
  Matches(value: number): boolean {
    // Anchor offset from current row/column – matches how Excel shifts refs.
    // We ignore the shift when resolving because callers compute the value
    // for the right cell up front.
    switch (this._operator) {
      case ">": return value > this._threshold;
      case ">=": return value >= this._threshold;
      case "<": return value < this._threshold;
      case "<=": return value <= this._threshold;
      case "=": return value === this._threshold;
      case "<>": return value !== this._threshold;
    }
    return false;
  }
}
