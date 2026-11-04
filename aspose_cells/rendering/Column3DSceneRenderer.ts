import type { ParsedChart, ChartSeries } from "./ChartData.js";
import { Column3DProjection } from "./Column3DProjection.js";

/**
 * 3D 柱体场景渲染器。
 * 绘制墙壁、网格、柱体和标签。
 */
export class Column3DSceneRenderer {
    /**
     * 绘制完整的 3D 柱状图场景。
     */
    Draw(
        doc: any,
        area: { left: number; top: number; right: number; bottom: number },
        chart: ParsedChart,
        min: number,
        max: number,
        major: number
    ): void {
        // 计算分类数量
        let count = chart.Categories.length;
        for (const series of chart.Series) {
            count = Math.max(count, series.Values.length);
        }
        if (count === 0) return;

        const projection = new Column3DProjection(area, chart, count);

        // 绘制墙壁和网格
        Column3DSceneRenderer.DrawWalls(doc, projection, chart, min, max, major, count);

        // 绘制柱体
        Column3DSceneRenderer.DrawColumns(doc, projection, chart, min, max, count);

        // 绘制标签
        Column3DSceneRenderer.DrawLabels(doc, projection, chart, min, max, major, count);
    }

    private static DrawWalls(
        doc: any,
        p: Column3DProjection,
        chart: ParsedChart,
        min: number,
        max: number,
        major: number,
        count: number
    ): void {
        doc.save();
        doc.lineWidth(0.6);
        doc.strokeColor("#D9D9D9");

        // 网格线
        if (major > 0) {
            const steps = Math.min(100, Math.round((max - min) / major));
            for (let i = 0; i <= steps; i++) {
                const y = i * major / (max - min);
                const frontLeft = p.Project(0, y, 0);
                const backLeft = p.Project(0, y, 1);
                const backRight = p.Project(1, y, 1);
                const frontRight = p.Project(1, y, 0);

                doc.moveTo(frontLeft.x, frontLeft.y);
                doc.lineTo(backLeft.x, backLeft.y);
                doc.lineTo(backRight.x, backRight.y);
                doc.moveTo(frontLeft.x, frontLeft.y);
                doc.lineTo(frontRight.x, frontRight.y);
                doc.stroke();
            }
        }

        // 底部前沿线
        const bl = p.Project(0, 0, 0);
        const br = p.Project(1, 0, 0);
        doc.moveTo(bl.x, bl.y);
        doc.lineTo(br.x, br.y);
        doc.stroke();

        // 柱之间的分隔
        for (let i = 0; i <= count; i++) {
            const x = i / count;
            const p1 = p.Project(x, 0, 0);
            doc.moveTo(p1.x, p1.y);
            doc.lineTo(p1.x, p1.y + 2.8);
            doc.stroke();
        }

        doc.restore();
    }

    private static DrawColumns(
        doc: any,
        p: Column3DProjection,
        chart: ParsedChart,
        min: number,
        max: number,
        count: number
    ): void {
        if (max <= min) return;

        const width = Column3DProjection.ColumnWidth(chart, count);
        const depth = Column3DProjection.ColumnDepth(chart);
        const baseline = Math.max(0, Math.min(1, -min / (max - min)));

        // 收集所有柱体并按深度排序
        const order: number[] = [];
        for (let s = 0; s < chart.Series.length; s++) {
            for (let c = 0; c < count; c++) {
                if (c < chart.Series[s].Values.length && chart.Series[s].Values[c].hasValue) {
                    order.push(s * count + c);
                }
            }
        }

        order.sort((a, b) => {
            const sa = Math.floor(a / count);
            const sc = a % count;
            const sb = Math.floor(b / count);
            const bc = b % count;
            return sc - bc || sb - sa;
        });

        for (const idx of order) {
            const s = Math.floor(idx / count);
            const c = idx % count;
            const series = chart.Series[s];
            const v = series.Values[c].value;
            const bottomY = Math.min(baseline, Math.max(0, Math.min(1, (v - min) / (max - min))));
            const topY = Math.max(baseline, Math.min(1, Math.max(0, (v - min) / (max - min))));
            const centerX = (c + 0.5) / count;
            const depthZ = (s + 0.5) / chart.Series.length;

            const corners = p.ColumnCorners(centerX, depthZ, width, depth, bottomY, topY);
            Column3DSceneRenderer.DrawColumnFaces(doc, corners, series);
        }
    }

    private static DrawColumnFaces(
        doc: any,
        corners: Array<{ x: number; y: number }>,
        series: ChartSeries
    ): void {
        const color = series.Color;
        const h = (v: number) => Math.round(v).toString(16).padStart(2, "0");
        const baseColor = `#${h(color.red)}${h(color.green)}${h(color.blue)}`;

        // 5 个面：front(top), back, side-left, side-right, top
        const faces: Array<{ indices: number[]; shade: number }> = [
            { indices: [0, 3, 7, 4], shade: 0.72 }, // 左侧面（后）
            { indices: [3, 2, 6, 7], shade: 0.80 }, // 右侧面
            { indices: [1, 5, 6, 2], shade: 0.66 }, // 后侧面
            { indices: [4, 7, 6, 5], shade: 0.85 }, // 顶面（如果可见）
            { indices: [0, 4, 5, 1], shade: 1.00 }  // 前面
        ];

        for (const face of faces) {
            const pts = face.indices.map(i => corners[i]);
            // 背面剔除
            const cross = (pts[1].x - pts[0].x) * (pts[2].y - pts[0].y) - (pts[1].y - pts[0].y) * (pts[2].x - pts[0].x);
            if (cross <= 0) continue;

            doc.save();
            const lighten = face.shade;
            const r = Math.round(color.red * lighten);
            const g = Math.round(color.green * lighten);
            const b = Math.round(color.blue * lighten);
            doc.fillColor(`#${h(r)}${h(g)}${h(b)}`);
            doc.moveTo(pts[0].x, pts[0].y);
            for (let i = 1; i < pts.length; i++) {
                doc.lineTo(pts[i].x, pts[i].y);
            }
            doc.closePath();
            doc.fill();
            doc.restore();
        }
    }

    private static DrawLabels(
        _doc: any,
        _p: Column3DProjection,
        _chart: ParsedChart,
        _min: number,
        _max: number,
        _major: number,
        _count: number
    ): void {
        // 标签由外部 ChartRenderer 处理
    }
}
