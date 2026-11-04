import type { ParsedChart } from "./ChartData.js";

interface FontContext {
    fontSize: number;
}

/**
 * 处理图表分类标签的换行和尺寸计算。
 */
export class ChartCategoryLabels {
    /**
     * 将多行标签拆分为字符串数组。支持 \n 换行。
     */
    static Lines(label: string): string[] {
        if (label == null) {
            return [""];
        }
        return label.split(/[\r\n]/);
    }

    /**
     * 单行标签的高度（以 PDF points 为单位）。
     */
    static LineAdvance(fontContext: FontContext | null): number {
        if (fontContext == null) {
            return 13;
        }
        return Math.round(fontContext.fontSize * 1.3);
    }

    /**
     * 多行标签相对于单行标签的额外高度。
     */
    static ExtraHeight(
        chart: ParsedChart,
        fontContext: FontContext | null
    ): number {
        if (chart == null || chart.Categories == null) {
            return 0;
        }
        const lineAdvance = ChartCategoryLabels.LineAdvance(fontContext);
        let extra = 0;
        for (const category of chart.Categories) {
            const lines = ChartCategoryLabels.Lines(category ?? "");
            if (lines.length > 1) {
                extra = Math.max(extra, (lines.length - 1) * lineAdvance);
            }
        }
        return extra;
    }
}
