import type { ParsedChart, Rect } from "./ChartData.js";

/**
 * 散点图的绘图区域布局计算。
 */
export class ScatterChartPlotLayout {
    /**
     * 计算散点图的绘图区域。
     */
    static Resolve(
        area: Rect,
        chart: ParsedChart
    ): Rect {
        if (chart == null) {
            return { left: area.left, top: area.top + area.height * 0.1, right: area.right, bottom: area.bottom - area.height * 0.1 };
        }

        if (chart.HasManualPlot) {
            const w = area.right - area.left;
            const h = area.bottom - area.top;
            return {
                left: area.left + chart.PlotX * w,
                top: area.top + chart.PlotY * h,
                right: area.left + (chart.PlotX + chart.PlotW) * w,
                bottom: area.top + (chart.PlotY + chart.PlotH) * h
            };
        }

        const leftPadding = Math.max(area.width * 0.1, 44);
        const topPadding = chart.Title ? area.height * 0.15 : area.height * 0.08;
        const rightPadding = Math.max(area.width * 0.05, 20);
        const bottomPadding = Math.max(area.height * 0.1, 24);

        return {
            left: area.left + leftPadding,
            top: area.top + topPadding,
            right: area.right - rightPadding,
            bottom: area.bottom - bottomPadding
        };
    }
}
