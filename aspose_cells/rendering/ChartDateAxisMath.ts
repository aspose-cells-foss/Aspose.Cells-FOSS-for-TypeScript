import type { ParsedChart } from "./ChartData.js";

/**
 * 日期轴的数学计算工具。
 * 处理 Excel 日期序列号的归一化、范围计算和屏幕坐标映射。
 */
export class ChartDateAxisMath {
    /**
     * 将日期序列号归一化到统一的日期系统（Windows 1900 = 0）。
     * Mac 1904 系统的日期需要减去 1462 天。
     */
    static NormalizeDateSerial(serial: number, dateSystem: number): number {
        if (dateSystem === 1904) {
            return serial - 1462;
        }
        return serial;
    }

    /**
     * 反向归一化：将统一日期转回图表的原始日期系统。
     */
    static DenormalizeDateSerial(normalized: number, dateSystem: number): number {
        if (dateSystem === 1904) {
            return normalized + 1462;
        }
        return normalized;
    }

    /**
     * 获取日期轴的最小日期序列号。
     */
    static GetAxisMinDate(chart: ParsedChart): number {
        if (chart == null || !chart.HasDateCategoryAxis) return 0;
        if (chart.DateAxisMin != null && chart.DateAxisMin.hasValue) {
            return chart.DateAxisMin.value;
        }
        let min = Infinity;
        for (const cv of chart.CategoryValues) {
            if (cv.hasValue && cv.value < min) min = cv.value;
        }
        return isFinite(min) ? min : 0;
    }

    /**
     * 获取日期轴的最大日期序列号。
     */
    static GetAxisMaxDate(chart: ParsedChart): number {
        if (chart == null || !chart.HasDateCategoryAxis) return 1;
        if (chart.DateAxisMax != null && chart.DateAxisMax.hasValue) {
            return chart.DateAxisMax.value;
        }
        let max = -Infinity;
        for (const cv of chart.CategoryValues) {
            if (cv.hasValue && cv.value > max) max = cv.value;
        }
        return isFinite(max) ? max : 1;
    }

    /**
     * 将日期序列号映射到绘图区域的 X 坐标。
     */
    static DateCategoryX(
        serial: number,
        plotMinDate: number,
        plotMaxDate: number,
        plot: { left: number; right: number }
    ): number {
        if (plotMaxDate <= plotMinDate) return plot.left;
        const ratio = (serial - plotMinDate) / (plotMaxDate - plotMinDate);
        return plot.left + ratio * (plot.right - plot.left);
    }

    /**
     * 规范化点的日期序列号到绘图范围。
     */
    static NormalizePointSerial(
        chart: ParsedChart,
        serial: number,
        dateSystem: number
    ): number {
        return ChartDateAxisMath.NormalizeDateSerial(serial, dateSystem);
    }
}
