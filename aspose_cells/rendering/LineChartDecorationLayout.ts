import type { ParsedChart } from "./ChartData.js";

/**
 * 折线图装饰元素布局。
 * 计算标题中心 Y 坐标、图例位置、间距等。
 */
export class LineChartDecorationLayout {
    /**
     * 计算标题中心 Y 坐标。
     */
    static TitleCenterY(
        area: { top: number; height: number },
        chart: ParsedChart,
        defaultCenterY: number
    ): number {
        if (chart == null || !chart.Title) return defaultCenterY;

        if (chart.HasManualTitle) {
            return area.top + (chart.TitleY + (chart.TitleFont?.Size ?? 9) * 0.5) * area.height;
        }

        return defaultCenterY;
    }

    /**
     * 图例中心 Y 坐标。
     */
    static LegendCenterY(
        area: { top: number; bottom: number; height: number },
        chart: ParsedChart,
        defaultCenterY: number
    ): number {
        if (chart == null) return defaultCenterY;

        if (chart.HasManualLegend) {
            return area.top + (chart.LegendY + chart.LegendH * 0.5) * area.height;
        }

        return defaultCenterY;
    }

    /**
     * 图例文本与 swatch 之间的间距。
     */
    static TextGap(chart: ParsedChart, defaultValue: number): number {
        if (chart == null) return defaultValue;
        return 2;
    }

    /**
     * 图例条目之间的间距。
     */
    static ItemGap(chart: ParsedChart, defaultValue: number): number {
        if (chart == null) return defaultValue;
        return 8;
    }

    /**
     * 图例 swatch 宽度。
     */
    static SwatchWidth(chart: ParsedChart, defaultValue: number): number {
        if (chart == null) return defaultValue;
        return 9;
    }

    /**
     * 图例水平偏移。
     */
    static HorizontalOffset(chart: ParsedChart, defaultValue: number): number {
        if (chart == null) return defaultValue;
        return 0;
    }

    /**
     * 是否使用自动底部图例布局。
     */
    static UsesAutomaticBottomLegend(chart: ParsedChart): boolean {
        if (chart == null) return false;
        return !chart.HasManualLegend && chart.LegendPosition === "b";
    }
}
