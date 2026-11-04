import type { ParsedChart } from "./ChartData.js";

/**
 * 柱状图/条形图 3D 投影变换。
 */
export class Column3DProjection {
    private readonly area: { left: number; top: number; right: number; bottom: number };
    private readonly chart: ParsedChart;
    private readonly categoryCount: number;
    private readonly labelReserve: number;
    private readonly depthPercent: number;
    private readonly depthX: number;
    private readonly depthY: number;

    constructor(
        area: { left: number; top: number; right: number; bottom: number },
        chart: ParsedChart,
        categoryCount: number,
        labelReserve: number = 0
    ) {
        this.area = area;
        this.chart = chart;
        this.categoryCount = Math.max(1, categoryCount);
        this.labelReserve = labelReserve;
        this.depthPercent = Math.max(25, Math.min(200, chart.DepthPercent ?? 100));

        const plotH = area.bottom - area.top;
        this.depthX = plotH * 0.045 * (this.depthPercent / 100);
        this.depthY = -plotH * 0.03 * (this.depthPercent / 100);
    }

    /**
     * 投影 3D 坐标到 2D 屏幕。
     */
    Project(x: number, y: number, z: number): { x: number; y: number } {
        const plotW = this.area.right - this.area.left;
        const plotH = this.area.bottom - this.area.top;

        return {
            x: this.area.left + x * plotW + z * this.depthX,
            y: this.area.bottom - y * plotH + z * this.depthY
        };
    }

    /**
     * 获取相机深度（用于排序隐藏面）。
     */
    CameraDepth(x: number, y: number, z: number): number {
        // 越大表示越远（越应该先绘制）
        return x * 100 + z * 50;
    }

    /**
     * 计算一个柱体的 8 个角点。
     */
    ColumnCorners(
        centerX: number,
        depthZ: number,
        width: number,
        depth: number,
        bottomY: number,
        topY: number
    ): Array<{ x: number; y: number }> {
        const hw = width / 2;
        const corners3D: Array<[number, number, number]> = [
            [centerX - hw, bottomY, 0],       // 0: front-bottom-left
            [centerX - hw, topY, 0],           // 1: front-top-left
            [centerX + hw, topY, 0],           // 2: front-top-right
            [centerX + hw, bottomY, 0],        // 3: front-bottom-right
            [centerX - hw, bottomY, depthZ],   // 4: back-bottom-left
            [centerX - hw, topY, depthZ],      // 5: back-top-left
            [centerX + hw, topY, depthZ],      // 6: back-top-right
            [centerX + hw, bottomY, depthZ]    // 7: back-bottom-right
        ];

        return corners3D.map(([x, y, z]) => this.Project(x, y, z));
    }

    /**
     * 获取柱体墙的尺寸。
     */
    WallSize(side: boolean): { width: number; height: number } {
        if (side) {
            return {
                width: Math.abs(this.depthX),
                height: this.area.bottom - this.area.top
            };
        }
        return {
            width: this.area.right - this.area.left,
            height: Math.abs(this.depthY)
        };
    }

    /**
     * 计算列宽比例。
     */
    static ColumnWidth(chart: ParsedChart, slotCount: number): number {
        const gapRatio = (chart.GapWidthPercent ?? 150) / 100;
        return 1 / (slotCount * (1 + gapRatio));
    }

    /**
     * 计算柱体深度比例。
     */
    static ColumnDepth(chart: ParsedChart): number {
        return Math.min(0.5, Math.max(0.15, (chart.DepthPercent ?? 100) / 400));
    }
}
