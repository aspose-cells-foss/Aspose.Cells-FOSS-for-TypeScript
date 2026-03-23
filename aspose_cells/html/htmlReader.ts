import { cellRef } from "../util";

export class HtmlReader {
  private html: string;
  private pos: number = 0;
  private len: number;
  tagName: string = "";
  isStartTag: boolean = false;
  isEndTag: boolean = false;
  isSelfClosing: boolean = false;
  text: string = "";
  private currentTagStart: number = -1;
  private currentTagEnd: number = -1;
  private tagBuffer: string = "";

  constructor(html: string) {
    this.html = html;
    this.len = html.length;
  }

  read(): boolean {
    this.tagName = "";
    this.isStartTag = false;
    this.isEndTag = false;
    this.isSelfClosing = false;
    this.text = "";

    while (this.pos < this.len) {
      const char = this.html[this.pos];

      if (char === "<") {
        const nextChars = this.html.slice(this.pos, this.pos + 4).toLowerCase();
        if (nextChars.startsWith("<!--")) {
          this.skipComment();
          continue;
        }

        this.currentTagStart = this.pos;
        this.pos++;

        if (this.pos < this.len && this.html[this.pos] === "/") {
          this.isEndTag = true;
          this.pos++;
        } else {
          this.isStartTag = true;
        }

        this.readTagName();
        this.readAttributes();

        this.currentTagEnd = this.pos;
        if (this.pos < this.len && this.html[this.pos] === "/") {
          this.isSelfClosing = true;
          this.pos++;
        }
        if (this.pos < this.len && this.html[this.pos] === ">") {
          this.pos++;
        }

        return true;
      } else if (char === ">") {
        this.pos++;
      } else {
        this.readText();
        if (this.text.length > 0) {
          return true;
        }
      }
    }

    return false;
  }

  private skipComment(): void {
    const end = this.html.indexOf("-->", this.pos);
    if (end !== -1) {
      this.pos = end + 3;
    } else {
      this.pos = this.len;
    }
  }

  private readTagName(): void {
    const start = this.pos;
    while (this.pos < this.len) {
      const char = this.html[this.pos];
      if (/[a-zA-Z0-9\-]/.test(char)) {
        this.pos++;
      } else {
        break;
      }
    }
    this.tagName = this.html.slice(start, this.pos).toLowerCase();
  }

  private readAttributes(): void {
    let attrBuffer = "";
    while (this.pos < this.len) {
      const char = this.html[this.pos];
      if (char === ">") {
        break;
      }
      attrBuffer += char;
      this.pos++;
    }
    this.tagBuffer = attrBuffer;
  }

  private readText(): void {
    const start = this.pos;
    while (this.pos < this.len) {
      const char = this.html[this.pos];
      if (char === "<") {
        break;
      }
      this.pos++;
    }
    this.text = this.html.slice(start, this.pos).replace(/&nbsp;/g, " ");
  }

  getAttribute(name: string): string | null {
    const regex = new RegExp(`${name}=["']([^"']*)["']`, "i");
    const match = this.tagBuffer.match(regex);
    return match ? match[1] : null;
  }

  getOuterHtml(): string {
    if (this.currentTagStart === -1 || this.currentTagEnd === -1) {
      return "";
    }
    return this.html.slice(this.currentTagStart, this.currentTagEnd);
  }

  getInnerHtml(): string {
    const outer = this.getOuterHtml();
    if (!outer) return "";

    const endTag = `</${this.tagName}>`;
    const startPos = outer.indexOf(">");
    const endPos = outer.lastIndexOf(endTag);

    if (startPos === -1 || endPos === -1) return "";

    return outer.slice(startPos + 1, endPos);
  }

  querySelectorAll(selector: string): HtmlReader[] {
    const results: HtmlReader[] = [];
    const originalPos = this.pos;

    this.pos = 0;
    while (this.read()) {
      if (this.tagName === selector && this.isStartTag) {
        results.push(new HtmlReader(this.getOuterHtml()));
      }
    }

    this.pos = originalPos;
    return results;
  }

  querySelector(selector: string): HtmlReader | null {
    const results = this.querySelectorAll(selector);
    return results.length > 0 ? results[0] : null;
  }

  getFullElement(): string {
    const tagName = this.tagName;
    if (!tagName) return "";

    const startPos = this.currentTagStart;
    if (startPos === -1) return "";

    let depth = 1;
    let pos = this.currentTagEnd;

    while (pos < this.len && depth > 0) {
      if (this.html.slice(pos, pos + 4).toLowerCase() === "<!--") {
        const endComment = this.html.indexOf("-->", pos + 4);
        if (endComment !== -1) {
          pos = endComment + 3;
          continue;
        }
      }

      if (this.html[pos] === "<") {
        if (this.html.slice(pos, pos + 2) === "</") {
          const closeTagMatch = this.html.slice(pos).match(/^<\/?([a-zA-Z0-9\-]+)[^>]*>/i);
          if (closeTagMatch && closeTagMatch[1].toLowerCase() === tagName) {
            depth--;
            if (depth === 0) {
              pos = pos + closeTagMatch[0].length;
              break;
            }
          }
        } else {
          const openTagMatch = this.html.slice(pos).match(/^<([a-zA-Z0-9\-]+)[^>]*>/i);
          if (openTagMatch && openTagMatch[1].toLowerCase() === tagName) {
            depth++;
          }
        }
      }
      pos++;
    }

    return this.html.slice(startPos, pos);
  }

  getElementsByTagName(tagName: string): HtmlReader[] {
    const results: HtmlReader[] = [];
    const originalPos = this.pos;

    this.pos = 0;
    while (this.read()) {
      if (this.tagName === tagName.toLowerCase()) {
        results.push(new HtmlReader(this.getOuterHtml()));
      }
    }

    this.pos = originalPos;
    return results;
  }

  get textContent(): string {
    let text = "";
    const originalPos = this.pos;

    this.pos = 0;
    while (this.read()) {
      if (this.isStartTag || this.isEndTag) {
        text += " ";
      }
      text += this.text;
    }

    this.pos = originalPos;
    return text.replace(/\s+/g, " ").trim();
  }
}
