import type PDFDocument from "pdfkit";
import { RenderContext } from "./RenderContext";
import { PageLayout } from "./PageLayout";
import { SheetLayout } from "./SheetLayout";
import { SheetCellRenderer } from "./SheetCellRenderer";
import { ShapeRenderer } from "./ShapeRenderer";

/**
 * Renders a single PDF page using pdfkit. Handles cell grid, shapes, and charts.
 */
export class SheetPdfRenderer {
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

    // Render shapes
    const shapeRenderer = new ShapeRenderer(this.doc, this.context, layout);
    shapeRenderer.renderAll(page);

    this.doc.restore();
  }
}
