
/**
 * 3D 气泡球体渲染器。
 * 使用径向渐变模拟球体表面光影。
 */
export class BubbleSphereRenderer {
    /**
     * 绘制一个 3D 球体效果。
     */
    static Draw(
        doc: any,
        center: { x: number; y: number },
        radius: number,
        colorHex: string
    ): void {
        // 解析颜色
        const r = parseInt(colorHex.slice(1, 3), 16);
        const g = parseInt(colorHex.slice(3, 5), 16);
        const b = parseInt(colorHex.slice(5, 7), 16);

        // 高光点偏移
        const highlightOffset = radius * 0.3;
        const highlightX = center.x - highlightOffset;
        const highlightY = center.y - highlightOffset;

        // 径向渐变：从高光到原色再到深色
        doc.save();
        doc.fillColor(colorHex);
        doc.circle(center.x, center.y, radius);
        doc.fill();

        // 叠加高光（更浅的颜色）
        const highlightHex = BubbleSphereRenderer.Lighten(r, g, b, 0.45);
        doc.globalAlpha(0.55);
        doc.fillColor(highlightHex);
        doc.circle(highlightX, highlightY, radius * 0.5);
        doc.fill();

        // 底部阴影
        const shadowHex = BubbleSphereRenderer.Darken(r, g, b, 0.4);
        doc.globalAlpha(0.45);
        doc.fillColor(shadowHex);
        doc.circle(center.x + highlightOffset * 0.5, center.y + highlightOffset * 0.5, radius * 0.6);
        doc.fill();

        doc.restore();
    }

    static DrawInLegend(
        doc: any,
        center: { x: number; y: number },
        radius: number,
        colorHex: string
    ): void {
        BubbleSphereRenderer.Draw(doc, center, radius, colorHex);
    }

    private static Lighten(r: number, g: number, b: number, factor: number): string {
        const nr = Math.round(r + (255 - r) * factor);
        const ng = Math.round(g + (255 - g) * factor);
        const nb = Math.round(b + (255 - b) * factor);
        return BubbleSphereRenderer.ToHex(nr, ng, nb);
    }

    private static Darken(r: number, g: number, b: number, factor: number): string {
        const nr = Math.round(r * (1 - factor));
        const ng = Math.round(g * (1 - factor));
        const nb = Math.round(b * (1 - factor));
        return BubbleSphereRenderer.ToHex(nr, ng, nb);
    }

    private static ToHex(r: number, g: number, b: number): string {
        const h = (v: number) => v.toString(16).padStart(2, "0");
        return `#${h(r)}${h(g)}${h(b)}`;
    }
}
