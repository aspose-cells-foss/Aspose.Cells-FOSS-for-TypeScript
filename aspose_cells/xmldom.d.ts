declare module "@xmldom/xmldom" {
  export class DOMParser {
    constructor();
    parseFromString(xml: string, type: string): Document;
  }
}
