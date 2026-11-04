import { RenderUnits } from "./RenderUnits";

/** Maps an OOXML paper-size code to portrait dimensions in points. */
export class PaperSizes {
  static getDimensionsPoints(
    paperSizeCode: number,
  ): { widthPt: number; heightPt: number } {
    let widthIn: number;
    let heightIn: number;
    switch (paperSizeCode) {
      case 1: widthIn = 8.5; heightIn = 11.0; break; // Letter
      case 3: widthIn = 11.0; heightIn = 17.0; break; // Tabloid
      case 5: widthIn = 8.5; heightIn = 14.0; break; // Legal
      case 8: widthIn = 11.69; heightIn = 16.54; break; // A3
      case 9: widthIn = 210 / 25.4; heightIn = 297 / 25.4; break; // A4
      case 11: widthIn = 5.83; heightIn = 8.27; break; // A5
      case 13: widthIn = 7.17; heightIn = 10.12; break; // B5 (JIS)
      default: widthIn = 210 / 25.4; heightIn = 297 / 25.4; break;
    }
    return {
      widthPt: RenderUnits.inchesToPoints(widthIn),
      heightPt: RenderUnits.inchesToPoints(heightIn),
    };
  }
}
