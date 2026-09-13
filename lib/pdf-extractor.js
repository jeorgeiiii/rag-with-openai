/**
 * PDF Text Extraction Utility
 *
 * Extracts text from PDF files using pdf-parse (v2 class-based API)
 */

import { PDFParse } from 'pdf-parse';
import { CanvasFactory } from 'pdf-parse/worker';

/**
 * Extract text from PDF buffer
 *
 * @param {Buffer} pdfBuffer - PDF file as buffer
 * @returns {Promise<{text: string, pages: number, info: object}>}
 */
export async function extractTextFromPDF(pdfBuffer) {
  const parser = new PDFParse({ data: pdfBuffer, CanvasFactory });
  try {
    console.log('[PDF Extractor] Starting extraction, buffer size:', pdfBuffer.length);

    // pdf-parse's fake-worker message channel (used when no real Worker
    // thread is available, e.g. Node) can't handle two concurrent commands
    // on the same parser instance, so these must run sequentially.
    const textResult = await parser.getText();
    const infoResult = await parser.getInfo();
    console.log('[PDF Extractor] Extraction complete, pages:', infoResult.total);

    return {
      text: textResult.text,
      pages: infoResult.total,
      info: infoResult.info || {}
    };
  } catch (error) {
    console.error('[PDF Extractor] Error details:', {
      message: error.message,
      stack: error.stack,
      name: error.name
    });
    throw new Error(`Failed to extract text from PDF: ${error.message}`);
  } finally {
    await parser.destroy();
  }
}

/**
 * Extract text from PDF file (from Blob/File)
 *
 * @param {File|Blob} file - PDF file
 * @returns {Promise<{text: string, pages: number, info: object}>}
 */
export async function extractTextFromPDFFile(file) {
  try {
    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);
    return await extractTextFromPDF(buffer);
  } catch (error) {
    console.error('[PDF Extractor] Error reading PDF file:', error);
    throw new Error('Failed to read PDF file');
  }
}

/**
 * Validate PDF file
 *
 * @param {File|Blob} file - File to validate
 * @returns {{valid: boolean, error: string|null}}
 */
export function validatePDF(file) {
  // Check file type
  if (file.type !== 'application/pdf') {
    return { valid: false, error: 'File must be a PDF' };
  }

  // Check file size (max 25MB - balanced for performance and capacity)
  const maxSize = 25 * 1024 * 1024; // 25MB
  if (file.size > maxSize) {
    return { valid: false, error: 'PDF file too large (max 25MB)' };
  }

  if (file.size === 0) {
    return { valid: false, error: 'PDF file is empty' };
  }

  return { valid: true, error: null };
}
