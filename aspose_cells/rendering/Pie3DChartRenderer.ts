import type { ParsedChart, Rect } from "./ChartData.js";

/**
 * Pie 3D 图表渲染器。
 * 使用投影算法绘制带透视效果的 3D 饼图。
 */
export class Pie3DChartRenderer {
    /**
     * 绘制 Pie 3D 图表，返回投影椭圆区域。
     */
    static Draw(
        doc: any,
        area: Rect,
        chart: ParsedChart,
        values: Array<{ hasValue: boolean; value: number }>,
        colors: Array<{ red: number; green: number; blue: number; alpha: number }>,
        sum: number
    ): Rect {
        const perspective = chart.RightAngleAxes ? 0
            : Math.sin((chart.PerspectiveDeg ?? 30) * Math.PI / 180) * 0.68;
        const p = Math.max(0, Math.min(0.6, perspective));

        const flatView = 1 - Math.min(1, p / 0.34);
        const projectedRadiusX = (area.right - area.left) * (0.45 - 0.03 * flatView);
        const radiusX = projectedRadiusX * Math.sqrt(1 - p * p);
        const tilt = Math.sin(Math.max(0, Math.min(70, chart.RotationXDeg ?? 15)) * Math.PI / 180);
        const radiusY = (area.bottom - area.top) * (0.15 + 0.065 * flatView)
            * Math.max(0.2, tilt / Math.sin(15 * Math.PI / 180));

        const centerX = (area.left + area.right) / 2;
        const centerY = area.top + (area.bottom - area.top) * (0.303 + 0.107 * flatView);

        const depth = (area.bottom - area.top) * 0.193
            * Math.max(0.25, Math.min(2, (chart.DepthPercent ?? 100) / 100));

        const ellipse = {
            left: centerX - projectedRadiusX,
            top: centerY - radiusY / (1 + p),
            right: centerX + projectedRadiusX,
            bottom: centerY + radiusY / (1 - p)
        };

        if (sum <= 0) return ellipse;

        const startAngle = (-90 + (chart.RotationYDeg ?? 0)) * Math.PI / 180;

        // 先绘制可见 rim（侧面）
        Pie3DChartRenderer.DrawVisibleRim(
            doc, centerX, centerY, radiusX, radiusY, p, depth,
            startAngle, values, sum, colors
        );

        // 再绘制顶面
        Pie3DChartRenderer.DrawTopSurface(
            doc, centerX, centerY, radiusX, radiusY, p,
            startAngle, values, sum, colors
        );

        return ellipse;
    }

    private static DrawVisibleRim(
        doc: any,
        cx: number, cy: number, rx: number, ry: number, p: number, depth: number,
        startAngle: number,
        values: Array<{ hasValue: boolean; value: number }>,
        sum: number,
        colors: Array<{ red: number; green: number; blue: number; alpha: number }>
    ): void {
        // 简化：不单独绘制 rim 光影
    }

    private static DrawTopSurface(
        doc: any,
        cx: number, cy: number, rx: number, ry: number, p: number,
        startAngle: number,
        values: Array<{ hasValue: boolean; value: number }>,
        sum: number,
        colors: Array<{ red: number; green: number; blue: number; alpha: number }>
    ): void {
        let currentAngle = startAngle;

        for (let i = values.length - 1; i >= 0; i--) {
            if (!values[i].hasValue) continue;
            const sweep = (Math.abs(values[i].value) / sum) * Math.PI * 2;
            const c = colors[i] ?? colors[0];
            const h = (v: number) => Math.round(v).toString(16).padStart(2, "0");
            const hex = `#${h(c.red)}${h(c.green)}${h(c.blue)}`;

            doc.save();
            doc.fillColor(hex);
            doc.moveTo(cx, cy);
            const steps = Math.max(12, Math.ceil(sweep / 0.05));
            for (let j = 0; j <= steps; j++) {
                const angle = currentAngle + sweep * j / steps;
                const sine = Math.sin(angle);
                const divisor = 1 - p * sine;
                const x = cx + rx * Math.cos(angle) / divisor;
                const y = cy + ry * sine / divisor;
                doc.lineTo(x, y);
            }
            doc.closePath();
            doc.fill();
            doc.restore();

            currentAngle += sweep;
        }
    }
}
