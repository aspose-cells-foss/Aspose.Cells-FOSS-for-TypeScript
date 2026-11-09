export class ColorValue {
  public readonly A: number = 0;
  public readonly R: number = 0;
  public readonly G: number = 0;
  public readonly B: number = 0;
  public readonly ThemeIndex: number | null = null;
  public readonly Tint: number | null = null;
  public readonly Indexed: number | null = null;

  /** Initializes a default (zero) color value. */
  constructor();
  /** Initializes an RGB color value. */
  constructor(a: number, r: number, g: number, b: number);
  /** Initializes a theme-based color value. */
  constructor(themeIndex: number, tint: number | null);
  /** Initializes an indexed color value. */
  constructor(indexed: number);
  constructor(a?: number, b?: number | null, c?: number | null, d?: number) {
    if (a === undefined) {
      return;
    }

    if (d !== undefined) {
      this.A = a;
      this.R = b as number;
      this.G = c as number;
      this.B = d;
      return;
    }

    if (c === undefined) {
      this.ThemeIndex = a;
      this.Tint = b === undefined ? null : b;
      return;
    }

    this.Indexed = a;
  }

  public Equals(other: ColorValue | null): boolean {
    if (other == null) {
      return false;
    }

    return (
      this.A === other.A &&
      this.R === other.R &&
      this.G === other.G &&
      this.B === other.B &&
      this.ThemeIndex === other.ThemeIndex &&
      this.Tint === other.Tint &&
      this.Indexed === other.Indexed
    );
  }

  public hashKey(): number {
    let hash = this.A;
    hash = (Math.imul(hash, 397) ^ this.R) | 0;
    hash = (Math.imul(hash, 397) ^ this.G) | 0;
    hash = (Math.imul(hash, 397) ^ this.B) | 0;
    hash = (Math.imul(hash, 397) ^ (this.ThemeIndex ?? 0)) | 0;
    hash = (Math.imul(hash, 397) ^ (this.Tint == null ? 0 : this.Tint | 0)) | 0;
    hash = (Math.imul(hash, 397) ^ (this.Indexed ?? 0)) | 0;
    return hash;
  }

  public clone(): ColorValue {
    if (this.ThemeIndex !== null) {
      return new ColorValue(this.ThemeIndex, this.Tint);
    }

    if (this.Indexed !== null) {
      return new ColorValue(this.Indexed);
    }

    return new ColorValue(this.A, this.R, this.G, this.B);
  }
}
