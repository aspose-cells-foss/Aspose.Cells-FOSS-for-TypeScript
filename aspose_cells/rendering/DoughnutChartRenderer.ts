import type { ParsedChart, ChartSeries } from "./ChartData.js";

/**
 * 环形图渲染器。
 * 与饼图类似，但中心有一个空心圆。
 */
export class DoughnutChartRenderer {
    /**
     * 绘制环形图，返回绘图区域矩形供后续元素使用。
     */
    static Draw(
        doc: any,
        area: { left: number; top: number; right: number; bottom: number },
        chart: ParsedChart
    ): { left: number; top: number; right: number; bottom: number } {
        const size = Math.min(area.right - area.left, area.bottom - area.top) * 0.55;
        const centerX = (area.left + area.right) / 2;
        const centerY = area.top + (area.bottom - area.top) * 0.45;
        const rect = {
            left: centerX - size / 2,
            top: centerY - size / 2,
            right: centerX + size / 2,
            bottom: centerY + size / 2
        };

        if (chart == null || chart.Series.length === 0) return rect;

        const series = chart.Series[0];
        const outerRadius = size / 2;
        const holeRatio = chart.DoughnutHoleSizeRatio ?? 0.5;
        const innerRadius = outerRadius * holeRatio;

        let sum = 0;
        for (const v of series.Values) {
            if (v.hasValue) sum += Math.abs(v.value);
        }
        if (sum <= 0) return rect;

        let startAngle = -Math.PI / 2; // -90 度，顶部开始

        for (let i = 0; i < series.Values.length; i++) {
            const v = series.Values[i];
            if (!v.hasValue) continue;

            const sweep = (Math.abs(v.value) / sum) * Math.PI * 2;
            const endAngle = startAngle + sweep;
            const color = DoughnutChartRenderer.GetSliceColor(chart, i);

            DoughnutChartRenderer.DrawSlice(doc, centerX, centerY, outerRadius, innerRadius, startAngle, endAngle, color);
            startAngle = endAngle;
        }

        return rect;
    }

    private static DrawSlice(
        doc: any,
        cx: number,
        cy: number,
        outerR: number,
        innerR: number,
        startAngle: number,
        endAngle: number,
        colorHex: string
    ): void {
        doc.save();
        doc.fillColor(colorHex);
        doc.moveTo(cx + innerR * Math.cos(startAngle), cy + innerR * Math.sin(startAngle));

        // 外弧
        const outerSteps = Math.max(8, Math.ceil(Math.abs(endAngle - startAngle) / 0.1));
        for (let i = 1; i <= outerSteps; i++) {
            const a = startAngle + (endAngle - startAngle) * (i / outerSteps);
            doc.lineTo(cx + outerR * Math.cos(a), cy + outerR * Math.sin(a));
        }

        // 内弧（反向）
        for (let i = outerSteps - 1; i >= 0; i--) {
            const a = startAngle + (endAngle - startAngle) * (i / outerSteps);
            doc.lineTo(cx + innerR * Math.cos(a), cy + innerR * Math.sin(a));
        }

        doc.closePath();
        doc.fill();
        doc.restore();
    }

    private static GetSliceColor(chart: ParsedChart, index: number): string {
        if (chart.Series[0].PointColors && index < chart.Series[0].PointColors.length) {
            const pc = chart.Series[0].PointColors[index];
            if (pc) {
                const h = (v: number) => Math.round(v).toString(16).padStart(2, "0");
                return `#${h(pc.red)}${h(pc.green)}${h(pc.blue)}`;
            }
        }
        const colors = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"];
        return colors[index % colors.length];
    }
}
