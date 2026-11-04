import type { ParsedChart, ChartSeries } from "./ChartData.js";

/**
 * 给柱状图和条形图的柱/条绘制边框。
 * 使用 pdfkit 的 stroke 功能。
 */
export class ChartBarBorder {
    /**
     * 给指定矩形绘制柱/条边框。
     * @param doc PDFKit 文档实例
     * @param rect 矩形区域
     * @param series 图表序列
     */
    static Draw(
        doc: any,
        rect: { left: number; top: number; right: number; bottom: number },
        series: ChartSeries
    ): void {
        if (series == null || !series.HasVisibleLine) {
            return;
        }
        const width = Math.max(0.5, series.LineWidthPt);
        const color = series.LineColor.toPdfKitColor();
        doc.lineWidth(width);
        doc.strokeColor(color);
        doc.rect(rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top);
        doc.stroke();
    }
}
