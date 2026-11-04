import type { ParsedChart, ChartSeries } from "./ChartData.js";

/**
 * 处理图表区域的图案填充（如斜线、点状等）。
 * PDFKit 不直接支持图案，这里返回 false 表示需要回退到纯色填充。
 */
export class ChartPatternRenderer {
    /**
     * 绘制图表区域背景图案。
     * @returns true 表示成功绘制了图案，false 表示应回退到纯色
     */
    static DrawChartArea(
        doc: any,
        rect: { left: number; top: number; right: number; bottom: number },
        chart: ParsedChart
    ): boolean {
        if (chart == null || !chart.HasChartAreaPattern) {
            return false;
        }
        // PDFKit 不原生支持 hatch 图案，简化处理
        return false;
    }

    /**
     * 给单个序列矩形绘制图案（用于图例 swatch）。
     * @returns true 表示成功绘制，false 表示需要回退
     */
    static DrawRect(
        doc: any,
        rect: { left: number; top: number; right: number; bottom: number },
        series: ChartSeries,
        fallbackColor: string
    ): boolean {
        if (series == null || !series.HasPatternFill) {
            return false;
        }
        // 简化：暂时不支持图案，使用纯色
        doc.save();
        doc.fillColor(fallbackColor);
        doc.rect(rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top);
        doc.fill();
        doc.restore();
        return true;
    }
}
