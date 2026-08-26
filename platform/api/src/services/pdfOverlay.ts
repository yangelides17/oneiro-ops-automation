/**
 * PDF signature overlay service.
 *
 * Uses pdf-lib to overlay a principal's signature image + name + title + date
 * onto a PDF document during the approval step.
 *
 * Ported from webapp/server.js approve-signin and approve-cert-payroll handlers
 * which use pdf-lib to draw signature overlays.
 */
import { PDFDocument, rgb, StandardFonts } from 'pdf-lib';

export interface SignatureOverlayParams {
  /** The PDF bytes to overlay onto. */
  pdfBytes: Buffer;
  /** Base64-encoded PNG signature image. */
  signatureB64: string;
  /** Signer's printed name. */
  name: string;
  /** Signer's title (e.g., "Principal"). */
  title: string;
  /** Date string to print (e.g., "08/24/2026"). */
  dateStr: string;
  /**
   * Where to place the signature block on the last page.
   * x, y are from bottom-left in PDF points (72 pts/inch).
   * Defaults to bottom-right area if not specified.
   */
  position?: { x: number; y: number };
}

/**
 * Overlay a signature image + name + title + date onto a PDF.
 * Returns the modified PDF as a Buffer.
 */
export async function overlaySignature(params: SignatureOverlayParams): Promise<Buffer> {
  const {
    pdfBytes,
    signatureB64,
    name,
    title,
    dateStr,
    position,
  } = params;

  const pdfDoc = await PDFDocument.load(pdfBytes);
  const pages = pdfDoc.getPages();
  if (pages.length === 0) throw new Error('PDF has no pages');

  const page = pages[pages.length - 1]; // overlay on last page
  const { width, height } = page.getSize();
  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);

  // Default position: bottom-right signature block area
  const x = position?.x ?? width - 250;
  const y = position?.y ?? 80;

  // Draw signature image if provided
  if (signatureB64) {
    try {
      const sigData = Buffer.from(signatureB64, 'base64');
      const sigImage = await pdfDoc.embedPng(sigData);
      const sigDims = sigImage.scale(0.4); // scale to reasonable size
      const maxW = 150;
      const maxH = 50;
      const scale = Math.min(maxW / sigDims.width, maxH / sigDims.height, 1);
      page.drawImage(sigImage, {
        x,
        y: y + 20,
        width: sigDims.width * scale,
        height: sigDims.height * scale,
      });
    } catch {
      // If signature image fails to embed (invalid PNG, etc.), continue without it
    }
  }

  // Draw name, title, date as text below the signature
  const fontSize = 9;
  const lineHeight = fontSize + 3;

  if (name) {
    page.drawText(name, { x, y: y + 5, size: fontSize, font, color: rgb(0, 0, 0) });
  }
  if (title) {
    page.drawText(title, { x, y: y + 5 - lineHeight, size: fontSize, font, color: rgb(0, 0, 0) });
  }
  if (dateStr) {
    page.drawText(dateStr, { x, y: y + 5 - lineHeight * 2, size: fontSize, font, color: rgb(0, 0, 0) });
  }

  const result = await pdfDoc.save();
  return Buffer.from(result);
}
