/**
 * Central place for the unit conversions used when leaving Excel's native units.
 * All layout coordinates downstream of this type are expressed in PDF points (1pt = 1/72 inch).
 */
export class RenderUnits {
  /** Points per inch (PDF user-space unit is the point). */
  static readonly PointsPerInch = 72;

  /** Excel measures pixels at 96 DPI; a pixel is therefore 0.75 points. */
  static readonly PixelsToPoints = 72 / 96;

  /** Inverse of PixelsToPoints: points expressed in 96-DPI pixels. */
  static readonly PointsToPixels = 96 / 72;

  static inchesToPoints(inches: number): number {
    return inches * RenderUnits.PointsPerInch;
  }

  static pixelsToPointsValue(pixels: number): number {
    return pixels * RenderUnits.PixelsToPoints;
  }

  /** Converts Excel column width (digit '0' characters) to pixels per ECMA-376. */
  static columnWidthCharsToPixels(widthChars: number, maxDigitWidth: number): number {
    if (maxDigitWidth <= 0) {
      maxDigitWidth = 7.5;
    }
    return (256 * widthChars + Math.trunc(128 / maxDigitWidth)) / 256 * maxDigitWidth;
  }

  static columnWidthCharsToPoints(widthChars: number, maxDigitWidth: number): number {
    return RenderUnits.pixelsToPointsValue(
      RenderUnits.columnWidthCharsToPixels(widthChars, maxDigitWidth)
    );
  }

  /** Converts Excel's default column width (8.43 chars) to points. */
  static defaultColumnWidthCharsToPoints(displayChars: number, maxDigitWidth: number): number {
    if (maxDigitWidth <= 0) {
      maxDigitWidth = 7.5;
    }
    const pixels = Math.round(displayChars * maxDigitWidth) + 4;
    return RenderUnits.pixelsToPointsValue(pixels);
  }
}
