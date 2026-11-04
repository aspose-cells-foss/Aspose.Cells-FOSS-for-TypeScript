import type { ParsedChart } from "./ChartData.js";
import { ChartDateAxisMath } from "./ChartDateAxisMath.js";

/**
 * Area 3D 图表日期轴数学工具。
 */
export class Area3DDateAxisMath {
    /**
     * 按月构建刻度标签。
     */
    static BuildMonthlyTicks(
        chart: ParsedChart,
        dateSystem: number,
        _culture: string | null
    ): Array<[number, string]> {
        const min = ChartDateAxisMath.GetAxisMinDate(chart);
        const max = ChartDateAxisMath.GetAxisMaxDate(chart);
        if (!chart.HasDateCategoryAxis || max <= min) return [];

        const result: Array<[number, string]> = [];
        const months = Math.ceil((max - min) / 30.4375);
        for (let i = 0; i <= Math.min(months, 12); i++) {
            const ratio = months > 0 ? i / months : 0;
            const serial = min + (max - min) * ratio;
            const pos = ChartDateAxisMath.DateCategoryX(serial, min, max, { left: 0, right: 1 });
            result.push([pos, ""]);
        }
        return result;
    }

    /**
     * 获取分类在 X 轴上的归一化位置。
     */
    static CategoryPosition(
        chart: ParsedChart,
        index: number,
        count: number,
        dateSystem: number
    ): number {
        if (chart.HasDateCategoryAxis && chart.CategoryValues.length > index
            && chart.CategoryValues[index].hasValue) {
            const min = ChartDateAxisMath.GetAxisMinDate(chart);
            const max = ChartDateAxisMath.GetAxisMaxDate(chart);
            if (max > min) {
                const serial = chart.CategoryValues[index].value;
                return (serial - min) / (max - min);
            }
        }
        return count > 0 ? index / count : 0;
    }
}
