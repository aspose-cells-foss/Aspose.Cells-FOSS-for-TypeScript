import { RgbaColor } from "./ColorValue";

/**
 * Extracts the `[ColorName]` directive from a Excel number format code section.
 * Recognised names match the MS-DOS palette (black, blue, cyan, green,
 * magenta, red, white, yellow).  Case insensitive.
 */
export class NumberFormatTextColor {
  static Resolve(formatCode: string | undefined | null): RgbaColor | null {
    if (!formatCode) return null;
    const sections = splitFormatSections(formatCode);
    // Pick the first section that actually renders text (positive, negative,
    // zero or text) – colour applies to whichever one is active.  We scan
    // every section's raw text for a [ColorName] directive; the first match
    // wins, which matches the typical one-section case.
    for (const section of sections) {
      let inQuote = false;
      for (let i = 0; i < section.length; i++) {
        const ch = section[i];
        if (ch === '"') {
          inQuote = !inQuote;
          continue;
        }
        if (inQuote) continue;
        if (ch === "\\" || ch === "_" || ch === "*") {
          i++;
          continue;
        }
        if (ch !== "[") continue;
        const end = section.indexOf("]", i + 1);
        if (end < 0) break;
        const name = section.substring(i + 1, end);
        const color = lookupNamedColor(name);
        if (color) return color;
        i = end;
      }
    }
    return null;
  }
}

function splitFormatSections(formatCode: string): string[] {
  const sections: string[] = [];
  let current = "";
  let inQuote = false;
  for (let i = 0; i < formatCode.length; i++) {
    const ch = formatCode[i];
    if (ch === '"') {
      inQuote = !inQuote;
      current += ch;
      continue;
    }
    if (!inQuote && ch === ";") {
      sections.push(current);
      current = "";
      continue;
    }
    current += ch;
  }
  sections.push(current);
  return sections;
}

function lookupNamedColor(name: string): RgbaColor | null {
  switch (name.toLowerCase()) {
    case "black": return new RgbaColor(0, 0, 0);
    case "blue": return new RgbaColor(0, 0, 255);
    case "cyan": return new RgbaColor(0, 255, 255);
    case "green": return new RgbaColor(0, 255, 0);
    case "magenta": return new RgbaColor(255, 0, 255);
    case "red": return new RgbaColor(255, 0, 0);
    case "white": return new RgbaColor(255, 255, 255);
    case "yellow": return new RgbaColor(255, 255, 0);
    default: return null;
  }
}
