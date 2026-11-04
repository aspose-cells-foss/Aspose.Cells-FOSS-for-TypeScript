/**
 * TypeScript equivalent of Aspose.Cells_FOSS.Core.ColorValue.
 * Represents a color which may be:
 * - Direct ARGB (A, R, G, B fields)
 * - Theme reference (ThemeIndex + optional Tint)
 * - Indexed palette (Indexed)
 */
export class ColorValue {
  R = 0;
  G = 0;
  B = 0;
  A = 0;
  ThemeIndex?: number;
  Tint?: number;
  Indexed?: number;

  static fromArgb(r: number, g: number, b: number, a = 255): ColorValue {
    const c = new ColorValue();
    c.R = r;
    c.G = g;
    c.B = b;
    c.A = a;
    return c;
  }

  static fromTheme(index: number, tint = 0): ColorValue {
    const c = new ColorValue();
    c.ThemeIndex = index;
    c.Tint = tint;
    return c;
  }

  static fromIndexed(index: number): ColorValue {
    const c = new ColorValue();
    c.Indexed = index;
    return c;
  }

  /** Returns true if this is the "unset" sentinel (all zero, no theme, no indexed). */
  get isUnset(): boolean {
    return (
      this.A === 0 &&
      this.R === 0 &&
      this.G === 0 &&
      this.B === 0 &&
      this.ThemeIndex === undefined &&
      this.Indexed === undefined
    );
  }
}

/** Simple RGBA color used by the rendering pipeline. */
export class RgbaColor {
  constructor(
    public r: number,
    public g: number,
    public b: number,
    public a = 255,
  ) {}

  static fromHex(hex: string): RgbaColor {
    // Accepts "RRGGBB", "AARRGGBB", "#RRGGBB"
    let h = hex.replace(/^#/, "");
    let a = 255;
    if (h.length === 8) {
      a = parseInt(h.substring(0, 2), 16);
      h = h.substring(2);
    }
    const r = parseInt(h.substring(0, 2), 16);
    const g = parseInt(h.substring(2, 4), 16);
    const b = parseInt(h.substring(4, 6), 16);
    return new RgbaColor(r, g, b, a);
  }

  static black = new RgbaColor(0, 0, 0, 255);
  static white = new RgbaColor(255, 255, 255, 255);
  static transparent = new RgbaColor(0, 0, 0, 0);

  /** Format for pdfkit: "#RRGGBB" or "#RRGGBBAA" */
  toPdfKitColor(): string {
    const r = this.r.toString(16).padStart(2, "0");
    const g = this.g.toString(16).padStart(2, "0");
    const b = this.b.toString(16).padStart(2, "0");
    if (this.a < 255) {
      const a = this.a.toString(16).padStart(2, "0");
      return `#${r}${g}${b}${a}`;
    }
    return `#${r}${g}${b}`;
  }
}
