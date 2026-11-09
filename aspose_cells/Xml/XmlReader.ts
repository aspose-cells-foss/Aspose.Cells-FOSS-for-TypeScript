import { DOMParser } from "@xmldom/xmldom";

/**
 * Small helper around @xmldom/xmldom used by the XLSX reader.
 * Element matching is performed on the local name so that both default
 * namespaces and prefixed documents are handled uniformly.
 */
export interface XmlNode {
  nodeName: string;
  nodeType?: number;
  textContent: string | null;
  childNodes: ArrayLike<XmlNode> | null;
  getAttribute(name: string): string | null;
}

export function localName(nodeName: string): string {
  const separator = nodeName.indexOf(":");
  return separator >= 0 ? nodeName.substring(separator + 1) : nodeName;
}

export function parseXml(text: string): XmlNode {
  const parser = new DOMParser({
    errorHandler: {
      warning: () => undefined,
      error: () => undefined,
      fatalError: () => undefined,
    },
  });

  return parser.parseFromString(text, "text/xml") as unknown as XmlNode;
}

/** Returns direct child elements with the given local name. */
export function elementChildren(node: XmlNode | null | undefined, name: string): XmlNode[] {
  const result: XmlNode[] = [];
  if (!node || !node.childNodes) {
    return result;
  }

  for (let index = 0; index < node.childNodes.length; index++) {
    const child = node.childNodes[index];
    if (localName(child.nodeName) === name) {
      result.push(child);
    }
  }

  return result;
}

/** Returns the first direct child element with the given local name. */
export function firstElement(node: XmlNode | null | undefined, name: string): XmlNode | null {
  if (!node || !node.childNodes) {
    return null;
  }

  for (let index = 0; index < node.childNodes.length; index++) {
    const child = node.childNodes[index];
    if (localName(child.nodeName) === name) {
      return child;
    }
  }

  return null;
}

/** Returns all descendant elements with the given local name (any depth). */
export function descendants(node: XmlNode | null | undefined, name: string): XmlNode[] {
  const result: XmlNode[] = [];
  if (!node || !node.childNodes) {
    return result;
  }

  const visit = (current: XmlNode): void => {
    if (!current.childNodes) {
      return;
    }

    for (let index = 0; index < current.childNodes.length; index++) {
      const child = current.childNodes[index];
      if (localName(child.nodeName) === name) {
        result.push(child);
      }

      visit(child);
    }
  };

  visit(node);
  return result;
}

/** Returns the root element of a parsed document. */
export function documentRoot(document: XmlNode): XmlNode | null {
  if (!document.childNodes) {
    return null;
  }

  for (let index = 0; index < document.childNodes.length; index++) {
    const child = document.childNodes[index];
    if (typeof child.nodeType === "number") {
      if (child.nodeType === 1) {
        return child;
      }

      continue;
    }

    if (localName(child.nodeName) !== "#text" && localName(child.nodeName) !== "#comment" && localName(child.nodeName) !== "xml") {
      return child;
    }
  }

  return null;
}

/** Returns the concatenated text of all descendants. */
export function textOf(node: XmlNode | null | undefined): string {
  return node === null || node === undefined ? "" : (node.textContent ?? "");
}

export function attribute(node: XmlNode | null | undefined, name: string): string | null {
  if (!node) {
    return null;
  }

  // @xmldom returns an empty string for missing attributes, so treat "" as absent
  // and fall back to a local-name lookup that also resolves prefixed attributes.
  const value = node.getAttribute(name);
  if (value !== null && value !== undefined && value.length > 0) {
    return value;
  }

  const nodes = (node as unknown as { attributes?: ArrayLike<XmlNode> }).attributes;
  if (nodes) {
    for (let index = 0; index < nodes.length; index++) {
      if (localName(nodes[index].nodeName) === name) {
        return nodes[index].textContent;
      }
    }
  }

  return null;
}

export function requiredAttribute(node: XmlNode | null | undefined, name: string): string {
  const value = attribute(node, name);
  return value === null ? "" : value;
}

export function attributeBool(node: XmlNode | null | undefined, name: string): boolean {
  const value = attribute(node, name);
  if (value === null) {
    return false;
  }

  const normalized = value.trim().toLowerCase();
  return normalized === "1" || normalized === "true" || normalized === "on";
}

export function attributeOptionalBool(node: XmlNode | null | undefined, name: string): boolean | null {
  const value = attribute(node, name);
  if (value === null || value.trim().length === 0) {
    return null;
  }

  const normalized = value.trim().toLowerCase();
  return normalized === "1" || normalized === "true" || normalized === "on";
}

export function attributeNumber(node: XmlNode | null | undefined, name: string): number | null {
  const value = attribute(node, name);
  if (value === null || value.trim().length === 0) {
    return null;
  }

  const parsed = Number(value);
  return Number.isNaN(parsed) ? null : parsed;
}

export function attributeInt(node: XmlNode | null | undefined, name: string): number | null {
  const value = attributeNumber(node, name);
  return value === null ? null : Math.trunc(value);
}
