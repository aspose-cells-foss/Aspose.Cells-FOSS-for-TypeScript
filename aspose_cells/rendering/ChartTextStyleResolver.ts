import type { ParsedChart } from "./ChartData.js";

/**
 * 解析图表文字样式（字体、大小、颜色、对齐等）。
 * 为标题、轴标签、图例文字等提供统一的样式查询。
 */
export class ChartTextStyleResolver {
    /**
     * 获取标题字体。
     */
    static ResolveTitleFont(chart: ParsedChart): { name: string; size: number } {
        if (chart == null) return { name: "Calibri", size: 9 };
        if (chart.TitleFont) {
            return { name: chart.TitleFont.Name, size: chart.TitleFont.Size };
        }
        return { name: "Calibri", size: (chart.TextSizePt ?? 9) * 1.25 };
    }

    /**
     * 获取轴标签字号。
     */
    static ResolveAxisLabelSize(chart: ParsedChart): number {
        return chart == null ? 9 : (chart.TextSizePt ?? 9);
    }

    /**
     * 获取图例字号。
     */
    static ResolveLegendSize(chart: ParsedChart): number {
        return chart == null ? 9 : (chart.TextSizePt ?? 9);
    }

    /**
     * 获取数据标签字号。
     */
    static ResolveDataLabelSize(chart: ParsedChart): number {
        return chart == null ? 9 : (chart.TextSizePt ?? 9);
    }
}
