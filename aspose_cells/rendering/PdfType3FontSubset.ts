/**
 * A single Type3 font subset. One PDF Type3 font object covers at most 255 glyphs.
 * Tokens are stored as full string rune (supplementary codepoints use a 2-element
 * UTF-16 surrogate pair as a single token).
 */
export class PdfType3FontSubset {
  /** PDF resource name e.g. "T3F1". */
  ResourceName = "";
  /** True for fonts whose vertical metrics need a flipped FontMatrix. */
  UsePositiveYFontMatrix = false;
  /** Full-string tokens (one per glyph). */
  Tokens: string[] = [];
  /** Maps each token to its 1-based code (0 is unused / invalid). */
  TokenCodes: Map<string, number> = new Map();
  /** Object ID of the /Font object in the target PDF. */
  FontObjectId = 0;
  /** Object ID of the /ToUnicode CMap in the target PDF. */
  ToUnicodeObjectId = 0;
}
