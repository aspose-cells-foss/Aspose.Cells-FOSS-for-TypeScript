import type { ParsedChart } from "./ChartData.js";
import { ChartDateAxisMath } from "./ChartDateAxisMath.js";

/**
 * 折线图分类轴渲染器。
 * 处理折线图的分类轴（可能是日期轴）的刻度和网格线绘制。
 */
export class LineChartCategoryAxisRenderer {
    /**
     * 绘制折线图的分类轴。
     */
    static Draw(
        doc: any,
        plot: { left: number; top: number; right: number; bottom: number },
        chart: ParsedChart,
        dateSystem: number
    ): void {
        if (chart == null) return;

        doc.save();
        doc.lineWidth(1);
        doc.strokeColor("#D9D9D9");

        // 基础轴线
        doc.moveTo(plot.left, plot.bottom);
        doc.lineTo(plot.right, plot.bottom);
        doc.stroke();

        // 分类网格线
        if (chart.HasCategoryGridlines) {
            doc.strokeColor(chart.CategoryGridlineColor.toPdfKitColor());
            doc.lineWidth(Math.max(0.5, chart.CategoryGridlineWidthPt));

            if (chart.HasDateCategoryAxis) {
                LineChartCategoryAxisRenderer.DrawDateGridlines(
                    doc, plot, chart, dateSystem
                );
            } else {
                LineChartCategoryAxisRenderer.DrawCategoryGridlines(
                    doc, plot, chart
                );
            }
        }

        doc.restore();
    }

    private static DrawCategoryGridlines(
        doc: any,
        plot: { left: number; top: number; right: number; bottom: number },
        chart: ParsedChart
    ): void {
        const count = Math.max(chart.Categories.length, chart.CategoryValues.length);
        if (count <= 0) return;

        const slot = (plot.right - plot.left) / count;
        for (let i = 1; i < count; i++) {
            const x = plot.left + i * slot;
            doc.moveTo(x, plot.top);
            doc.lineTo(x, plot.bottom);
            doc.stroke();
        }
    }

    private static DrawDateGridlines(
        doc: any,
        plot: { left: number; top: number; right: number; bottom: number },
        chart: ParsedChart,
        dateSystem: number
    ): void {
        const minDate = ChartDateAxisMath.GetAxisMinDate(chart);
        const maxDate = ChartDateAxisMath.GetAxisMaxDate(chart);
        if (maxDate <= minDate) return;

        const months = (maxDate - minDate) / 30.4375;
        const majorUnit = chart.DateAxisMajorUnit ?? 1;
        const steps = Math.ceil(months / majorUnit);
        if (steps <= 0) return;

        for (let i = 1; i <= steps; i++) {
            const ratio = i / steps;
            const x = plot.left + ratio * (plot.right - plot.left);
            doc.moveTo(x, plot.top);
            doc.lineTo(x, plot.bottom);
            doc.stroke();
        }
    }
}
