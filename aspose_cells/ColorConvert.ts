import { Color } from "./Color";
import { ColorValue as CoreColorValue } from "./Core/ColorValue";

/** Converts a public Color into the core color payload. */
export function toCoreColor(color: Color): CoreColorValue {
  return new CoreColorValue(color.A, color.R, color.G, color.B);
}
