import type PDFDocument from "pdfkit";
import { RenderContext } from "./RenderContext";
import { SheetLayout } from "./SheetLayout";
import { PageLayout } from "./PageLayout";
import type { Style, BorderLine } from "../types";

/** Line widths (in points) for each Excel border style. */
const BORDER_WIDTHS: Record<string, number> = {
  hair: 0.25,
  thin: 0.5,
  medium: 1.0,
  thick: 1.5,
  dashed: 0.5,
  dotted: 0.5,
  dashDot: 0.5,
  dashDotDot: 0.5,
  slantDashDot: 1.0,
  double: 1.5,
};

export class SheetCellRenderer {
  constructor(
    private doc: InstanceType<typeof PDFDocument>,
    private context: RenderContext,
    private layout: SheetLayout,
    private styles: Map<number, Style> = new Map(),
  ) {}

  renderCellGrid(page: PageLayout): void {
    const { startRow, endRow, startColumn, endColumn } = page;

    // Pass 1: cell fills (background)
    for (let r = startRow; r <= endRow && r <= this.layout.lastRow; r++) {
      if (this.layout.rowHidden[r]) continue;
      for (let c = startColumn; c <= endColumn && c <= this.layout.lastColumn; c++) {
        if (this.layout.columnHidden[c]) continue;
        this.renderCellFill(r, c, page);
      }
    }

    // Pass 2: cell borders
    for (let r = startRow; r <= endRow && r <= this.layout.lastRow; r++) {
      if (this.layout.rowHidden[r]) continue;
      for (let c = startColumn; c <= endColumn && c <= this.layout.lastColumn; c++) {
        if (this.layout.columnHidden[c]) continue;
        this.renderCellBorders(r, c, page);
      }
    }

    // Pass 3: cell text
    for (let r = startRow; r <= endRow && r <= this.layout.lastRow; r++) {
      if (this.layout.rowHidden[r]) continue;
      for (let c = startColumn; c <= endColumn && c <= this.layout.lastColumn; c++) {
        if (this.layout.columnHidden[c]) continue;
        this.renderCellText(r, c, page);
      }
    }
  }

  /** Resolve a cell's style: inline style takes precedence over styleIndex lookup. */
  private getCellStyle(row: number, col: number): Style | undefined {
    const cell = this.layout.sheet.cells.get(`${row},${col}`);
    if (!cell) return undefined;
    if (cell.style) return cell.style;
    return this.styles.get(cell.styleIndex);
  }

  /** Draw the cell background fill (solid pattern or gray-scale patterns). */
  private renderCellFill(row: number, col: number, page: PageLayout): void {
    const style = this.getCellStyle(row, col);
    if (!style?.fill) return;

    const fill = style.fill;
    const pattern = fill.patternType;
    if (!pattern || pattern === "none") return;

    const x = this.layout.contentX(page.startColumn, col);
    const y = this.layout.contentY(page.startRow, row);
    const w = this.layout.columnWidthPt[col];
    const h = this.layout.rowHeightPt[row];

    // For "solid", fgColor is the background color.
    // For pattern fills, bgColor is the background and fgColor is the pattern color.
    if (pattern === "solid") {
      const color = this.resolveColor(fill.fgColor ?? fill.bgColor);
      if (!color) return;
      this.doc.save();
      this.doc.rect(x, y, w, h).fill(color);
      this.doc.restore();
    } else {
      // Pattern fills: draw bgColor first, then overlay fgColor pattern (simplified).
      const bgColor = this.resolveColor(fill.bgColor);
      if (bgColor) {
        this.doc.save();
        this.doc.rect(x, y, w, h).fill(bgColor);
        this.doc.restore();
      }
      // For gray patterns, draw a translucent fgColor overlay as a simplification.
      const fgColor = this.resolveColor(fill.fgColor);
      if (fgColor) {
        this.doc.save();
        this.doc.rect(x, y, w, h).fillOpacity(0.2).fill(fgColor);
        this.doc.restore();
      }
    }
  }

  /** Draw the four cell borders (top, right, bottom, left). */
  private renderCellBorders(row: number, col: number, page: PageLayout): void {
    const style = this.getCellStyle(row, col);
    if (!style?.border) return;

    const x = this.layout.contentX(page.startColumn, col);
    const y = this.layout.contentY(page.startRow, row);
    const w = this.layout.columnWidthPt[col];
    const h = this.layout.rowHeightPt[row];

    const border = style.border;
    this.drawBorderLine(border.top, x, y, x + w, y);
    this.drawBorderLine(border.bottom, x, y + h, x + w, y + h);
    this.drawBorderLine(border.left, x, y, x, y + h);
    this.drawBorderLine(border.right, x + w, y, x + w, y + h);
  }

  /** Draw a single border line segment with the given style/color. */
  private drawBorderLine(line: BorderLine | undefined, x1: number, y1: number, x2: number, y2: number): void {
    if (!line || !line.style) return;

    const style = line.style as string;
    const width = BORDER_WIDTHS[style] ?? 0.5;
    const color = this.resolveColor(line.color) ?? "#000000";

    this.doc.save();
    this.doc.lineWidth(width);
    this.doc.strokeColor(color);

    // Map dashed styles to pdfkit dash patterns
    switch (style) {
      case "dashed":
        this.doc.dash(3, { space: 2 });
        break;
      case "dotted":
        this.doc.dash(0.5, { space: 1.5 });
        break;
      case "double":
        // Draw two parallel lines offset by half the line width
        this.doc.lineWidth(Math.max(0.25, width / 3));
        this.doc.moveTo(x1, y1).lineTo(x2, y2).stroke();
        const dx = x2 - x1;
        const dy = y2 - y1;
        const len = Math.hypot(dx, dy);
        if (len > 0) {
          const nx = -dy / len * width * 0.6;
          const ny = dx / len * width * 0.6;
          this.doc.moveTo(x1 + nx, y1 + ny).lineTo(x2 + nx, y2 + ny).stroke();
        }
        this.doc.restore();
        return;
      default:
        this.doc.undash();
    }

    this.doc.moveTo(x1, y1).lineTo(x2, y2).stroke();
    this.doc.restore();
  }

  private renderCellText(row: number, col: number, page: PageLayout): void {
    const cellKey = `${row},${col}`;
    const cell = this.layout.sheet.cells.get(cellKey);
    if (!cell) return;
    if (cell.value === null || cell.value === undefined) return;
    if (this.layout.isMergeCovered(row, col)) return;

    const text = String(cell.value);
    if (!text) return;

    const x = this.layout.contentX(page.startColumn, col);
    const y = this.layout.contentY(page.startRow, row);
    const w = this.layout.columnWidthPt[col];
    const h = this.layout.rowHeightPt[row];

    this.doc.font("Helvetica");
    this.doc.fontSize(11);
    this.doc.fillColor("#000000");

    const textX = x + SheetLayout.HorizontalPaddingPt;
    const textY = y + SheetLayout.VerticalPaddingPt + 2;
    const availableW = w - 2 * SheetLayout.HorizontalPaddingPt;

    this.doc.text(text, textX, textY, {
      width: availableW,
      ellipsis: true,
    });
  }

  /**
   * Resolve a color string to a "#RRGGBB" hex value understood by pdfkit.
   * Supports ARGB hex (OOXML), RGB hex, and CSS rgb() notation.
   */
  private resolveColor(color: string | undefined): string | undefined {
    if (!color) return undefined;
    const c = color.trim();
    if (!c) return undefined;

    // Already a CSS hex color
    if (c.startsWith("#")) {
      if (/^#[0-9A-Fa-f]{8}$/.test(c)) return "#" + c.substring(3, 9); // #AARRGGBB → #RRGGBB
      if (/^#[0-9A-Fa-f]{6}$/.test(c)) return c;
      if (/^#[0-9A-Fa-f]{3}$/.test(c)) {
        return "#" + c[1] + c[1] + c[2] + c[2] + c[3] + c[3];
      }
      return undefined;
    }

    // OOXML ARGB hex (8 chars, no #)
    if (/^[0-9A-Fa-f]{8}$/.test(c)) return "#" + c.substring(2, 8);
    // OOXML RGB hex (6 chars, no #)
    if (/^[0-9A-Fa-f]{6}$/.test(c)) return "#" + c;

    // CSS rgb() / rgba()
    if (c.startsWith("rgb")) {
      const m = c.match(/\d+/g);
      if (m && m.length >= 3) {
        const r = parseInt(m[0], 10);
        const g = parseInt(m[1], 10);
        const b = parseInt(m[2], 10);
        return (
          "#" +
          [r, g, b]
            .map((v) => Math.max(0, Math.min(255, v)).toString(16).padStart(2, "0"))
            .join("")
        );
      }
    }

    return undefined;
  }
}
