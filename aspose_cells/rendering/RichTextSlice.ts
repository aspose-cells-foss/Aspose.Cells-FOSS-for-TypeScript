import type { CellFontContext } from "./CellFontContext";
import { RgbaColor } from "./ColorValue";

/** One homogeneous-style slice of rich text. */
export class RichTextSlice {
  Text = "";
  FontName = "";
  FontSize = 11;
  Bold = false;
  Italic = false;
  FontContext?: CellFontContext;
  WidthPt = 0;
  AscentPt = 0;
  DescentPt = 0;
  Color: RgbaColor = RgbaColor.black;
}
