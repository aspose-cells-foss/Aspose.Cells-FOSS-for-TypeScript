import type PDFDocument from "pdfkit";
import { RenderContext } from "./RenderContext";
import { PageLayout } from "./PageLayout";
import { SheetLayout } from "./SheetLayout";
import { SheetCellRenderer } from "./SheetCellRenderer";
import { ShapeRenderer } from "./ShapeRenderer";
import type { ExtractedPicture } from "./WorksheetDataExtractor";

/**
 * Renders a single PDF page using pdfkit. Handles cell grid, pictures, and shapes.
 */
export class SheetPdfRenderer {
  /** EMU per point (914400 EMU/in / 72 pt/in). */
  private static readonly EmuPerPoint = 12700;

  constructor(
    private doc: InstanceType<typeof PDFDocument>,
    private context: RenderContext,
  ) {}

  renderPage(page: PageLayout): void {
    const layout = page.sheet as SheetLayout;

    // Set up page transform: content origin + scale
    this.doc.save();
    this.doc.translate(page.contentOriginXPt, page.contentOriginYPt);
    if (page.scaleFactor !== 1) {
      this.doc.scale(page.scaleFactor);
    }

    // Render cells
    const cellRenderer = new SheetCellRenderer(this.doc, this.context, layout);
    cellRenderer.renderCellGrid(page);

    // Render pictures anchored on this page
    SheetPdfRenderer.drawPictures(this.doc, page, layout);

    // Render shapes
    const shapeRenderer = new ShapeRenderer(this.doc, this.context, layout);
    shapeRenderer.renderAll(page);

    this.doc.restore();
  }

  /**
   * Draws each anchored picture that intersects this page at its EMU-precise rectangle.
   * pdfkit handles the PNG/JPEG decoding internally. The active page clip naturally
   * trims any pixels that fall outside the printable content bounds.
   */
  private static drawPictures(
    doc: InstanceType<typeof PDFDocument>,
    page: PageLayout,
    layout: SheetLayout,
  ): void {
    const pictures: ExtractedPicture[] = (layout.sheet as any)?.pictures ?? [];
    const shapes = (layout.sheet as any)?.shapes as any[] | undefined;
    if (pictures.length === 0) return;

    // Page coverage in sheet-local points (for intersection check)
    const pageLeftColEdge = layout.columnStartPt[page.startColumn];
    const pageRightColEdge = layout.columnStartPt[page.endColumn + 1] ?? pageLeftColEdge;
    const pageTopRowEdge = layout.rowStartPt[page.startRow];
    const pageBottomRowEdge = layout.rowStartPt[page.endRow + 1] ?? pageTopRowEdge;

    for (let i = 0; i < pictures.length; i++) {
      const p = pictures[i];
      if (!p.data || p.data.length === 0) continue;

      const left = layout.columnStartPt[p.fromCol] + p.fromColOff / SheetPdfRenderer.EmuPerPoint;
      const right = layout.columnStartPt[p.toCol] + p.toColOff / SheetPdfRenderer.EmuPerPoint;
      const top = layout.rowStartPt[p.fromRow] + p.fromRowOff / SheetPdfRenderer.EmuPerPoint;
      const bottom = layout.rowStartPt[p.toRow] + p.toRowOff / SheetPdfRenderer.EmuPerPoint;

      const x = Math.min(left, right);
      const y = Math.min(top, bottom);
      const width = Math.abs(right - left);
      const height = Math.abs(bottom - top);

      if (width <= 0.01 || height <= 0.01) continue;

      // Intersects this page at all?
      if (
        right < pageLeftColEdge || left > pageRightColEdge ||
        bottom < pageTopRowEdge || top > pageBottomRowEdge
      ) {
        continue;
      }

      const shape = shapes?.find((s) => (s as any)?.type === "picture"
        && (s as any)?.fromCol === p.fromCol
        && (s as any)?.fromRow === p.fromRow);
      const rotationDeg = (shape as any)?.rotation ?? 0;
      const flipH = !!(shape as any)?.flipH;
      const flipV = !!(shape as any)?.flipV;

      const hasTransform = Math.abs(rotationDeg) > 0.01 || flipH || flipV;

      try {
        if (!hasTransform) {
          doc.image(p.data, x, y, { width, height });
        } else {
          doc.save();
          const cx = x + width / 2;
          const cy = y + height / 2;
          // Move picture center to origin, rotate, flip, move back.
          doc.translate(cx, cy);
          if (Math.abs(rotationDeg) > 0.01) doc.rotate(rotationDeg);
          if (flipH || flipV) {
            doc.scale(flipH ? -1 : 1, flipV ? -1 : 1);
          }
          doc.image(p.data, -width / 2, -height / 2, { width, height });
          doc.restore();
        }
      } catch {
        // Some image formats (e.g. unsupported TIFF) cause pdfkit to throw.
        // Skip rather than abort the whole page.
      }
    }
  }
}
