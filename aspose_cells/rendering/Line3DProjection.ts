import type { ParsedChart } from "./ChartData.js";

/**
 * 折线图 3D 投影变换。
 * 将 3D 坐标 (x, y, depth) 投影到 2D 屏幕坐标。
 */
export class Line3DProjection {
    private readonly area: { left: number; top: number; right: number; bottom: number };
    private readonly chart: ParsedChart;
    private readonly pointCount: number;
    private readonly depthPercent: number;
    private readonly depthScale: number;

    constructor(
        area: { left: number; top: number; right: number; bottom: number },
        chart: ParsedChart,
        pointCount: number
    ) {
        this.area = area;
        this.chart = chart;
        this.pointCount = Math.max(1, pointCount);
        this.depthPercent = chart.DepthPercent ?? 100;
        this.depthScale = Math.max(0.25, Math.min(2, this.depthPercent / 100));
    }

    /**
     * 将 3D 坐标投影为 2D 屏幕坐标。
     * x: 0-1 沿分类轴, y: 0-1 沿值轴, depth: 0-1 沿序列深度
     */
    Project(x: number, y: number, depth: number): { x: number; y: number } {
        const w = this.area.right - this.area.left;
        const h = this.area.bottom - this.area.top;
        const depthOffsetX = depth * 12 * this.depthScale;
        const depthOffsetY = -depth * 8 * this.depthScale;

        return {
            x: this.area.left + x * w + depthOffsetX,
            y: this.area.bottom - y * h + depthOffsetY
        };
    }

    /**
     * 返回序列在深度轴上的归一化位置。
     */
    SeriesDepth(seriesIndex: number): number {
        if (this.chart.Series.length <= 1) return 0;
        return seriesIndex / (this.chart.Series.length - 1);
    }

    /**
     * 计算一条 ribbon 面的四角坐标。
     */
    RibbonCorners(
        prev: { x: number; y: number },
        curr: { x: number; y: number },
        seriesIndex: number
    ): Array<{ x: number; y: number }> {
        const d = this.SeriesDepth(seriesIndex);
        const dNext = this.SeriesDepth(Math.min(seriesIndex + 1, this.chart.Series.length - 1));

        const frontPrev = this.Project(prev.x, prev.y, d);
        const frontCurr = this.Project(curr.x, curr.y, d);
        const backPrev = this.Project(prev.x, prev.y, dNext);
        const backCurr = this.Project(curr.x, curr.y, dNext);

        return [frontPrev, frontCurr, backCurr, backPrev];
    }

    /**
     * 计算 ribbon 面的阴影系数。
     */
    RibbonShade(
        prev: { x: number; y: number },
        curr: { x: number; y: number },
        corners: Array<{ x: number; y: number }>
    ): number {
        // 简化的光影计算
        return 0.75;
    }
}
