
/**
 * 图表 PDF 图像缓存。
 * 缓存图表渲染过程中嵌入的图像数据，避免重复解码。
 */
export class ChartPdfImageCache {
    private readonly _cache = new Map<string, { ref: string; data: Buffer }>();

    /**
     * 获取缓存的图像引用。如果不存在则添加。
     */
    GetOrAdd(key: string, imageData: Buffer): { ref: string; data?: Buffer } {
        const cached = this._cache.get(key);
        if (cached) {
            return { ref: cached.ref };
        }
        const ref = `chart_img_${this._cache.size}`;
        this._cache.set(key, { ref, data: imageData });
        return { ref, data: imageData };
    }

    /**
     * 将缓存的图像全部嵌入到 PDF 文档中。
     */
    EmbedAll(doc: any): void {
        for (const [, entry] of this._cache) {
            try {
                doc.registerImage(entry.ref, entry.data);
            } catch {
                // 忽略已经注册的图像
            }
        }
    }

    /**
     * 清空缓存。
     */
    Clear(): void {
        this._cache.clear();
    }

    /**
     * 缓存条目数。
     */
    get Count(): number {
        return this._cache.size;
    }
}
