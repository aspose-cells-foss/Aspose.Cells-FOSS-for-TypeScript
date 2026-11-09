/**
 * Represents an ARGB (alpha, red, green, blue) color.
 * Ported from Aspose.Cells_FOSS.Color.
 */

/**
 * Core color payload used by FromCore/ToCore.
 * Mirrors the internal ColorValue fields of the .NET implementation.
 */
export interface ColorValue {
  readonly A: number;
  readonly R: number;
  readonly G: number;
  readonly B: number;
  readonly ThemeIndex?: number | null;
  readonly Tint?: number | null;
  readonly Indexed?: number | null;
}

function toByte(value: number): number {
  return value & 0xff;
}

export class Color {
  private readonly alpha: number;
  private readonly red: number;
  private readonly green: number;
  private readonly blue: number;

  public constructor(a: number = 0, r: number = 0, g: number = 0, b: number = 0) {
    this.alpha = toByte(a);
    this.red = toByte(r);
    this.green = toByte(g);
    this.blue = toByte(b);
  }

  /** Gets the alpha component. */
  public get A(): number {
    return this.alpha;
  }

  /** Gets the red component. */
  public get R(): number {
    return this.red;
  }

  /** Gets the green component. */
  public get G(): number {
    return this.green;
  }

  /** Gets the blue component. */
  public get B(): number {
    return this.blue;
  }

  /** Gets the empty color (0, 0, 0, 0). */
  public static get Empty(): Color {
    return new Color(0, 0, 0, 0);
  }

  /** Creates a color from the four ARGB components. */
  public static FromArgb(a: number, r: number, g: number, b: number): Color;
  /** Creates a color from a 32-bit ARGB value. */
  public static FromArgb(argb: number): Color;
  public static FromArgb(a: number, r?: number, g?: number, b?: number): Color {
    if (r === undefined || g === undefined || b === undefined) {
      return new Color((a >> 24) & 0xff, (a >> 16) & 0xff, (a >> 8) & 0xff, a & 0xff);
    }
    return new Color(a, r, g, b);
  }

  /** Creates a color from a core color payload. */
  public static FromCore(value: ColorValue): Color {
    return new Color(value.A, value.R, value.G, value.B);
  }

  /** Converts this color to a core color payload. */
  public ToCore(): ColorValue {
    return {
      A: this.alpha,
      R: this.red,
      G: this.green,
      B: this.blue,
      ThemeIndex: null,
      Tint: null,
      Indexed: null,
    };
  }

  /** Returns true when all four components are zero. */
  public get IsEmpty(): boolean {
    return this.alpha === 0 && this.red === 0 && this.green === 0 && this.blue === 0;
  }

  /** Gets the 32-bit ARGB value of this color. */
  public ToArgb(): number {
    return (this.alpha << 24) | (this.red << 16) | (this.green << 8) | this.blue;
  }

  /** Determines whether the specified color equals this instance. */
  public Equals(other: Color | null | undefined): boolean {
    if (other === null || other === undefined) {
      return false;
    }
    return (
      this.alpha === other.alpha &&
      this.red === other.red &&
      this.green === other.green &&
      this.blue === other.blue
    );
  }

  /** Lowercase alias of Equals for generic assertion helpers. */
  public equals(other: Color | null | undefined): boolean {
    return this.Equals(other);
  }

  public ToString(): string {
    return `Color [A=${this.alpha}, R=${this.red}, G=${this.green}, B=${this.blue}]`;
  }

  public toString(): string {
    return this.ToString();
  }
}
