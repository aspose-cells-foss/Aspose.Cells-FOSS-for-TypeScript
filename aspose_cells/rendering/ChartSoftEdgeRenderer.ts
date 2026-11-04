
/**
 * 绘制带有软边缘（柔和边缘）的矩形填充。
 * 在 PDFKit 中通过阴影效果或渐变模拟。
 */
export class ChartSoftEdgeRenderer {
    /**
     * 绘制带软边缘的纯色填充矩形。
     * @param doc PDFKit 文档
     * @param rect 矩形区域
     * @param color 填充颜色
     * @param radius 软边缘半径（PDF points）
     */
    static DrawSolidFill(
        doc: any,
        rect: { left: number; top: number; right: number; bottom: number },
        color: string,
        radius: number
    ): void {
        const r = Math.min(radius, Math.min(rect.right - rect.left, rect.bottom - rect.top) * 0.5);
        doc.save();
        doc.fillColor(color);
        if (r > 0) {
            doc.roundedRect(rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top, r);
        } else {
            doc.rect(rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top);
        }
        doc.fill();
        doc.restore();
    }
}
