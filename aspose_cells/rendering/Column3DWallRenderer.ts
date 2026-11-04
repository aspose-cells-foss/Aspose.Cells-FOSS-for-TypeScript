import type { ParsedChart } from "./ChartData.js";
import { Column3DProjection } from "./Column3DProjection.js";

/**
 * 3D 柱体图墙壁渲染器。
 * 绘制侧墙和后墙，支持纹理和纯色填充。
 */
export class Column3DWallRenderer {
    /**
     * 绘制墙壁。
     */
    static Draw(
        doc: any,
        p: Column3DProjection,
        chart: ParsedChart
    ): void {
        // 后墙
        Column3DWallRenderer.DrawWall(doc, p, false, chart);
        // 侧墙
        Column3DWallRenderer.DrawWall(doc, p, true, chart);
    }

    private static DrawWall(
        doc: any,
        p: Column3DProjection,
        side: boolean,
        chart: ParsedChart
    ): void {
        const corners = side
            ? [p.Project(0, 0, 0), p.Project(0, 1, 0), p.Project(0, 1, 1), p.Project(0, 0, 1)]
            : [p.Project(0, 0, 1), p.Project(1, 0, 1), p.Project(1, 1, 1), p.Project(0, 1, 1)];

        const wallColor = side ? "#F0F0F0" : "#E8E8E8";

        doc.save();
        doc.fillColor(wallColor);
        doc.moveTo(corners[0].x, corners[0].y);
        for (let i = 1; i < corners.length; i++) {
            doc.lineTo(corners[i].x, corners[i].y);
        }
        doc.closePath();
        doc.fill();
        doc.restore();
    }
}
