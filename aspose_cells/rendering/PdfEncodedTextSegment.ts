/**
 * A segment of encoded text for insertion into a PDF content stream.
 * Each segment references a single font resource and carries a byte-encoded glyph sequence.
 */
export class PdfEncodedTextSegment {
  /** The PDF font resource name (e.g. "/T3F1"). */
  ResourceName = "";
  /** Raw bytes produced by the backend's encoding (WinAnsi or custom). */
  EncodedBytes: Uint8Array = new Uint8Array(0);
  /** x position in points for the Tm operator. */
  StartXPt = 0;
  /** TrueType prototype output requires an additional -Y flip. */
  RequiresTopDownTextFlip = false;
}
