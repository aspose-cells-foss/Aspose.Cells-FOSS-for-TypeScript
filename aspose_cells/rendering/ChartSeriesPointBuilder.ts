import type { ParsedChart } from "./ChartData.js";
import { ChartDateAxisMath } from "./ChartDateAxisMath.js";

/**
 * 构建序列点的屏幕坐标。
 * 处理日期轴和常规分类轴两种情况。
 */
export class ChartSeriesPointBuilder {
    /**
     * 针对日期轴构建序列点坐标。返回 null 表示不是日期轴。
     */
    static BuildDateSeriesPoints(
        chart: ParsedChart,
        values: Array<{ hasValue: boolean; value: number }>,
        plot: { left: number; right: number; top: number; bottom: number },
        min: number,
        max: number,
        dateSystem: number
    ): Array<{ x: number; y: number }> | null {
        if (chart == null || !chart.HasDateCategoryAxis) {
            return null;
        }

        const plotMinDate = ChartDateAxisMath.NormalizeDateSerial(
            ChartDateAxisMath.GetAxisMinDate(chart),
            dateSystem
        );
        const plotMaxDate = ChartDateAxisMath.NormalizeDateSerial(
            ChartDateAxisMath.GetAxisMaxDate(chart),
            dateSystem
        );

        if (plotMaxDate <= plotMinDate) return null;

        const points: Array<{ x: number; y: number }> = [];
        for (let i = 0; i < values.length; i++) {
            if (!values[i].hasValue) continue;
            if (i >= chart.CategoryValues.length || !chart.CategoryValues[i].hasValue) continue;

            const serial = ChartDateAxisMath.NormalizeDateSerial(
                chart.CategoryValues[i].value,
                dateSystem
            );
            const x = ChartDateAxisMath.DateCategoryX(serial, plotMinDate, plotMaxDate, plot);
            const y = ChartSeriesPointBuilder.ValueToY(values[i].value, min, max, plot);
            points.push({ x, y });
        }

        return points.length > 0 ? points : null;
    }

    private static ValueToY(
        value: number,
        min: number,
        max: number,
        plot: { top: number; bottom: number }
    ): number {
        if (max <= min) return plot.bottom;
        const ratio = (value - min) / (max - min);
        return plot.bottom - ratio * (plot.bottom - plot.top);
    }
}
