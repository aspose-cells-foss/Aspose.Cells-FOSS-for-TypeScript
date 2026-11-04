import { ColorValue, RgbaColor } from "./ColorValue";

/**
 * Resolves a ColorValue (ARGB direct, theme reference, or indexed palette) into a concrete RgbaColor.
 */
export class RenderColor {
  /** Office default theme palette (clrScheme document order). */
  private static readonly DefaultScheme: RgbaColor[] = [
    new RgbaColor(0x00, 0x00, 0x00), // dk1 (text1)
    new RgbaColor(0xff, 0xff, 0xff), // lt1 (bg1)
    new RgbaColor(0x44, 0x54, 0x6a), // dk2 (text2)
    new RgbaColor(0xe7, 0xe6, 0xe6), // lt2 (bg2)
    new RgbaColor(0x44, 0x72, 0xc4), // accent1
    new RgbaColor(0xed, 0x7d, 0x31), // accent2
    new RgbaColor(0xa5, 0xa5, 0xa5), // accent3
    new RgbaColor(0xff, 0xc0, 0x00), // accent4
    new RgbaColor(0x5b, 0x9b, 0xd5), // accent5
    new RgbaColor(0x70, 0xad, 0x47), // accent6
    new RgbaColor(0x05, 0x63, 0xc1), // hlink
    new RgbaColor(0x95, 0x4f, 0x72), // folHlink
  ];

  private static readonly IndexedPalette = RenderColor.buildIndexedPalette();
  private readonly scheme: RgbaColor[];
  majorFontName = "Calibri Light";
  minorFontName = "Calibri";

  private constructor(scheme: RgbaColor[]) {
    this.scheme = scheme;
  }

  static fromWorkbook(_themeXml?: string): RenderColor {
    // For now, use default scheme. Theme XML parsing can be added later.
    return new RenderColor([...RenderColor.DefaultScheme]);
  }

  /** Resolves the color, returns fallback when value is empty/unset. */
  resolve(color: ColorValue | undefined, fallback: RgbaColor): RgbaColor {
    if (!color) return fallback;

    if (color.ThemeIndex !== undefined) {
      const baseColor = this.resolveTheme(color.ThemeIndex);
      return RenderColor.applyTint(baseColor, color.Tint ?? 0);
    }

    if (color.Indexed !== undefined) {
      return RenderColor.resolveIndexed(color.Indexed, fallback);
    }

    if (color.isUnset) {
      return fallback;
    }

    return new RgbaColor(color.R, color.G, color.B, color.A || 255);
  }

  /** Resolves a DrawingML scheme color name ("accent1", "tx1", "bg1", etc.). */
  resolveSchemeName(name: string, fallback: RgbaColor): RgbaColor {
    let index = -1;
    switch (name) {
      case "dk1": case "tx1": index = 0; break;
      case "lt1": case "bg1": index = 1; break;
      case "dk2": case "tx2": index = 2; break;
      case "lt2": case "bg2": index = 3; break;
      case "accent1": index = 4; break;
      case "accent2": index = 5; break;
      case "accent3": index = 6; break;
      case "accent4": index = 7; break;
      case "accent5": index = 8; break;
      case "accent6": index = 9; break;
      case "hlink": index = 10; break;
      case "folHlink": index = 11; break;
      default: return fallback;
    }
    return index >= 0 && index < this.scheme.length ? this.scheme[index] : fallback;
  }

  /** Applies DrawingML luminance modulation/offset (lumMod/lumOff) to a color. */
  static applyLuma(color: RgbaColor, lumMod: number, lumOff: number): RgbaColor {
    let h = 0, s = 0, l = 0;
    RenderColor.toHsl(color, (v) => { h = v[0]; s = v[1]; l = v[2]; });
    l = Math.max(0, Math.min(1, l * lumMod + lumOff));
    return RenderColor.fromHsl(h, s, l, color.a);
  }

  private static toHsl(color: RgbaColor, cb: (v: number[]) => void): void {
    const r = color.r / 255;
    const g = color.g / 255;
    const b = color.b / 255;
    const max = Math.max(r, g, b);
    const min = Math.min(r, g, b);
    let h = 0, s = 0;
    const l = (max + min) * 0.5;
    if (Math.abs(max - min) < 1e-7) {
      cb([0, 0, l]);
      return;
    }
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (Math.abs(max - r) < 1e-7) h = (g - b) / d + (g < b ? 6 : 0);
    else if (Math.abs(max - g) < 1e-7) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
    h /= 6;
    cb([h, s, l]);
  }

  private static fromHsl(hue: number, sat: number, lum: number, alpha: number): RgbaColor {
    if (sat <= 0) {
      const g = Math.round(Math.max(0, Math.min(1, lum)) * 255);
      return new RgbaColor(g, g, g, alpha);
    }
    const q = lum < 0.5 ? lum * (1 + sat) : lum + sat - lum * sat;
    const p = 2 * lum - q;
    const c = (t: number) => {
      if (t < 0) t += 1;
      if (t > 1) t -= 1;
      if (t < 1 / 6) return p + (q - p) * 6 * t;
      if (t < 0.5) return q;
      if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
      return p;
    };
    const r = Math.round(Math.max(0, Math.min(1, c(hue + 1 / 3))) * 255);
    const g = Math.round(Math.max(0, Math.min(1, c(hue))) * 255);
    const b = Math.round(Math.max(0, Math.min(1, c(hue - 1 / 3))) * 255);
    return new RgbaColor(r, g, b, alpha);
  }

  private resolveTheme(themeIndex: number): RgbaColor {
    // Style theme index swaps first two dark/light pairs vs clrScheme order.
    let schemeIndex: number;
    switch (themeIndex) {
      case 0: schemeIndex = 1; break; // lt1
      case 1: schemeIndex = 0; break; // dk1
      case 2: schemeIndex = 3; break; // lt2
      case 3: schemeIndex = 2; break; // dk2
      default: schemeIndex = themeIndex; break;
    }
    if (schemeIndex >= 0 && schemeIndex < this.scheme.length) {
      return this.scheme[schemeIndex];
    }
    return RgbaColor.black;
  }

  private static resolveIndexed(index: number, fallback: RgbaColor): RgbaColor {
    if (index === 64 || index === 65) return fallback;
    if (index >= 0 && index < RenderColor.IndexedPalette.length) {
      return RenderColor.IndexedPalette[index];
    }
    return fallback;
  }

  private static applyTint(color: RgbaColor, tint: number): RgbaColor {
    if (tint === 0) return color;
    const t = (ch: number) => {
      let v: number;
      if (tint < 0) v = ch * (1 + tint);
      else v = ch * (1 - tint) + 255 * tint;
      return Math.max(0, Math.min(255, Math.round(v)));
    };
    return new RgbaColor(t(color.r), t(color.g), t(color.b), color.a);
  }

  private static buildIndexedPalette(): RgbaColor[] {
    const rgb = [
      0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF,
      0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF,
      0x800000, 0x008000, 0x000080, 0x808000, 0x800080, 0x008080, 0xC0C0C0, 0x808080,
      0x9999FF, 0x993366, 0xFFFFCC, 0xCCFFFF, 0x660066, 0xFF8080, 0x0066CC, 0xCCCCFF,
      0x000080, 0xFF00FF, 0xFFFF00, 0x00FFFF, 0x800080, 0x800000, 0x008080, 0x0000FF,
      0x00CCFF, 0xCCFFFF, 0xCCFFCC, 0xFFFF99, 0x99CCFF, 0xFF99CC, 0xCC99FF, 0xFFCC99,
      0x3366FF, 0x33CCCC, 0x99CC00, 0xFFCC00, 0xFF9900, 0xFF6600, 0x666699, 0x969696,
      0x003366, 0x339966, 0x003300, 0x333300, 0x993300, 0x993366, 0x333399, 0x333333,
    ];
    return rgb.map((v) => new RgbaColor(
      (v >> 16) & 0xff,
      (v >> 8) & 0xff,
      v & 0xff,
    ));
  }
}
