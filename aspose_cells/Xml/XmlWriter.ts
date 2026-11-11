/**
 * Minimal deterministic XML tree builder used by the XLSX writer.
 *
 * The .NET implementation builds documents with System.Xml.Linq. This builder
 * provides the same deterministic element/attribute ordering while keeping the
 * produced text predictable and independent from any DOM serializer.
 */
export class XmlElement {
  readonly name: string;
  private readonly _attributes: Array<[string, string]> = [];
  private readonly _children: Array<XmlElement | string> = [];
  private readonly _attributeNames = new Set<string>();
  private readonly _spacedSelfClosing: boolean;

  private constructor(name: string, spacedSelfClosing = false) {
    this.name = name;
    this._spacedSelfClosing = spacedSelfClosing;
  }

  static create(name: string): XmlElement {
    return new XmlElement(name);
  }

  /**
   * Creates an element whose empty form is serialized as `<name ... />`.
   *
   * Only the autoFilter writer opts in: its emitted markup is asserted against
   * the System.Xml.Linq output in golden tests. Every other XML part keeps the
   * default `<name .../>` form, so this factory cannot change unrelated output.
   */
  static createWithSpacedSelfClosing(name: string): XmlElement {
    return new XmlElement(name, true);
  }

  /**
   * Creates an element that declares the given default namespace.
   * The declaration is emitted first so the produced document keeps the
   * "xmlns" attribute ordering used by the .NET writer.
   */
  static createWithNamespace(name: string, namespaceUri: string): XmlElement {
    const element = new XmlElement(name);
    element.attr("xmlns", namespaceUri);
    return element;
  }

  /** Sets an attribute. Null/undefined values are ignored, matching XAttribute omission. */
  attr(name: string, value: string | number | boolean | null | undefined): XmlElement {
    if (value === null || value === undefined) {
      return this;
    }

    const text = String(value);
    if (this._attributeNames.has(name)) {
      const index = this._attributes.findIndex((pair) => pair[0] === name);
      this._attributes[index] = [name, text];
      return this;
    }

    this._attributeNames.add(name);
    this._attributes.push([name, text]);
    return this;
  }

  /** Sets a boolean attribute using the OOXML "1" convention. */
  boolAttr(name: string, value: boolean | null | undefined): XmlElement {
    if (value === null || value === undefined) {
      return this;
    }

    return this.attr(name, value ? 1 : 0);
  }

  child(child: XmlElement | string | null | undefined): XmlElement {
    if (child === null || child === undefined) {
      return this;
    }

    this._children.push(child);
    return this;
  }

  children(children: Array<XmlElement | string | null | undefined>): XmlElement {
    for (const child of children) {
      this.child(child);
    }

    return this;
  }

  get childCount(): number {
    return this._children.length;
  }

  get childrenItems(): ReadonlyArray<XmlElement | string> {
    return this._children;
  }

  get attributes(): ReadonlyArray<[string, string]> {
    return this._attributes;
  }

  getAttribute(name: string): string | undefined {
    const pair = this._attributes.find((item) => item[0] === name);
    return pair ? pair[1] : undefined;
  }

  toString(): string {
    let attributes = "";
    for (const [name, value] of this._attributes) {
      attributes += ` ${name}="${escapeXmlAttribute(value)}"`;
    }

    if (this._children.length === 0) {
      return this._spacedSelfClosing ? `<${this.name}${attributes} />` : `<${this.name}${attributes}/>`;
    }

    let inner = "";
    for (const child of this._children) {
      inner += typeof child === "string" ? escapeXmlText(child) : child.toString();
    }

    return `<${this.name}${attributes}>${inner}</${this.name}>`;
  }
}

/** Builds a full XML declaration followed by the root element. */
export function xmlDocument(root: XmlElement, standalone = true): string {
  const declaration = standalone
    ? '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    : '<?xml version="1.0" encoding="UTF-8"?>';
  return declaration + root.toString();
}

export function escapeXmlText(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

export function escapeXmlAttribute(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** Formats a double using invariant culture, mirroring C# double.ToString(InvariantCulture). */
export function formatInvariant(value: number): string {
  if (!Number.isFinite(value)) {
    return "0";
  }

  if (Number.isInteger(value)) {
    return String(value);
  }

  return String(value);
}
