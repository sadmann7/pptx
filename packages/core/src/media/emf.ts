export type EmfContent =
  | { type: "pdf"; data: Uint8Array }
  | { type: "bitmap"; imageData: ImageData }
  | { type: "empty" }
  | { type: "unsupported" };

const EMR_EOF = 14;
const EMR_COMMENT = 70;
const EMR_STRETCHDIBITS = 81;

const GDIC_COMMENT_ID = 0x43494447; // "GDIC".
const GDIC_BEGINGROUP = 0x00000002;
const GDIC_MULTIFORMATS = 0x40000004;

// EMF header signature at offset 40.
const EMF_SIGNATURE = 0x464d4520; // " EMF".

const PDF_HEADER = [0x25, 0x50, 0x44, 0x46]; // "%PDF".
const PDF_EOF = [0x25, 0x25, 0x45, 0x4f, 0x46]; // "%%EOF".

const BI_RGB = 0;
const MAX_EMF_BITMAP_PIXELS = 16_777_216;

// PPTX OLE previews are usually EMFs wrapping a PDF in a GDI comment or a
// bitmap in a STRETCHDIBITS record, so drawing records aren't interpreted.
export function parseEmfContent(data: Uint8Array): EmfContent {
  if (data.length < 44) return { type: "unsupported" };

  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);

  if (view.getUint32(40, true) !== EMF_SIGNATURE) {
    return { type: "unsupported" };
  }

  let offset = 0;
  let recordCount = 0;

  while (offset + 8 <= data.length) {
    const recordType = view.getUint32(offset, true);
    const recordSize = view.getUint32(offset + 4, true);

    if (recordSize < 8 || offset + recordSize > data.length) break;

    recordCount++;

    if (recordType === EMR_EOF) break;

    if (recordType === EMR_COMMENT && recordSize > 16) {
      const result = parseGdiComment(data, view, offset, recordSize);
      if (result) return result;
    }

    if (recordType === EMR_STRETCHDIBITS && recordSize > 80) {
      const result = parseStretchDibits(data, view, offset, recordSize);
      if (result) return result;
    }

    offset += recordSize;
  }

  // The header and EOF records alone make an empty EMF.
  if (recordCount <= 2) {
    return { type: "empty" };
  }

  return { type: "unsupported" };
}

function parseGdiComment(
  data: Uint8Array,
  view: DataView,
  offset: number,
  recordSize: number,
): EmfContent | null {
  // Record layout: type(4) + size(4) + cbData(4) + commentId(4) + ...
  if (offset + 16 > data.length) return null;

  const commentId = view.getUint32(offset + 12, true);

  if (commentId === GDIC_COMMENT_ID && offset + 20 <= data.length) {
    const publicType = view.getUint32(offset + 16, true);

    if (publicType === GDIC_BEGINGROUP) {
      const recordData = data.subarray(offset + 8, offset + recordSize);
      const pdf = extractPdfFromBuffer(recordData);
      if (pdf) return { type: "pdf", data: pdf };
    }

    if (publicType === GDIC_MULTIFORMATS && offset + 24 <= data.length) {
      const result = parseMultiformats(data, view, offset, recordSize);
      if (result) return result;
    }
  }

  if (recordSize > 100) {
    const recordData = data.subarray(offset + 8, offset + recordSize);
    const pdf = extractPdfFromBuffer(recordData);
    if (pdf) return { type: "pdf", data: pdf };
  }

  return null;
}

function parseMultiformats(
  data: Uint8Array,
  view: DataView,
  offset: number,
  _recordSize: number,
): EmfContent | null {
  // Offsets from the record start: commentIdentifier +12, publicCommentIdentifier
  // +16, outputRect +20 (a 16-byte RECTL), countFormats +36, then the format
  // descriptors from +40, each { signature, version, cbData, offData } of 4 bytes.
  if (offset + 40 > data.length) return null;

  const countFormats = view.getUint32(offset + 36, true);
  const descriptorStart = offset + 40;

  for (let i = 0; i < countFormats && i < 10; i++) {
    const descOff = descriptorStart + i * 16;
    if (descOff + 16 > data.length) break;

    const cbData = view.getUint32(descOff + 8, true);
    const offData = view.getUint32(descOff + 12, true);

    // offData is relative to the start of the record.
    const dataStart = offset + offData;
    if (dataStart + cbData > data.length || cbData === 0) continue;

    const formatData = data.subarray(dataStart, dataStart + cbData);
    const pdf = extractPdfFromBuffer(formatData);
    if (pdf) return { type: "pdf", data: pdf };
  }

  return null;
}

function extractPdfFromBuffer(buf: Uint8Array): Uint8Array | null {
  const pdfStart = findSequence(buf, PDF_HEADER);
  if (pdfStart === -1) return null;

  // Incrementally updated PDFs have several %%EOF markers; take the last.
  let pdfEnd = -1;
  for (let i = buf.length - PDF_EOF.length; i >= pdfStart; i--) {
    if (matchesAt(buf, i, PDF_EOF)) {
      pdfEnd = i + PDF_EOF.length;
      break;
    }
  }

  if (pdfEnd === -1) {
    pdfEnd = buf.length;
  }

  return buf.slice(pdfStart, pdfEnd);
}

function parseStretchDibits(
  data: Uint8Array,
  view: DataView,
  offset: number,
  _recordSize: number,
): EmfContent | null {
  // STRETCHDIBITS offsets from the record start: type 0, size 4, rclBounds 8 (16
  // bytes), then 4-byte xDest 24, yDest 28, xSrc 32, ySrc 36, cxSrc 40, cySrc 44,
  // offBmiSrc 48, cbBmiSrc 52, offBitsSrc 56, cbBitsSrc 60, iUsageSrc 64, dwRop 68,
  // cxDest 72 and cyDest 76.
  if (offset + 80 > data.length) return null;

  const offBmiSrc = view.getUint32(offset + 48, true);
  const cbBmiSrc = view.getUint32(offset + 52, true);
  const offBitsSrc = view.getUint32(offset + 56, true);
  const cbBitsSrc = view.getUint32(offset + 60, true);

  if (cbBmiSrc === 0 || cbBitsSrc === 0) return null;

  const bmiStart = offset + offBmiSrc;
  if (bmiStart + 40 > data.length) return null;

  const biWidth = view.getInt32(bmiStart + 4, true);
  const biHeight = view.getInt32(bmiStart + 8, true);
  const biBitCount = view.getUint16(bmiStart + 14, true);
  const biCompression = view.getUint32(bmiStart + 16, true);

  if (biCompression !== BI_RGB) return null;
  if (biBitCount !== 24 && biBitCount !== 32) return null;

  const width = Math.abs(biWidth);
  const height = Math.abs(biHeight);
  if (width === 0 || height === 0 || width > 8192 || height > 8192) return null;
  if (width * height > MAX_EMF_BITMAP_PIXELS) return null;

  const bitsStart = offset + offBitsSrc;
  if (bitsStart + cbBitsSrc > data.length) return null;

  const bytesPerPixel = biBitCount / 8;
  // DIB rows are padded to 4-byte boundaries.
  const rowStride = Math.ceil((width * bytesPerPixel) / 4) * 4;
  const requiredBits = rowStride * height;
  if (cbBitsSrc < requiredBits) return null;

  const bitsData = data.subarray(bitsStart, bitsStart + requiredBits);

  // Negative height means top-down rows; positive means bottom-up.
  const topDown = biHeight < 0;

  const imageData = new ImageData(width, height);

  for (let y = 0; y < height; y++) {
    const srcRow = topDown ? y : height - 1 - y;
    const srcOffset = srcRow * rowStride;
    const dstOffset = y * width * 4;

    for (let x = 0; x < width; x++) {
      const srcIdx = srcOffset + x * bytesPerPixel;
      if (srcIdx + bytesPerPixel > bitsData.length) break;

      // DIB stores BGR(A).
      imageData.data[dstOffset + x * 4 + 0] = bitsData[srcIdx + 2];
      imageData.data[dstOffset + x * 4 + 1] = bitsData[srcIdx + 1];
      imageData.data[dstOffset + x * 4 + 2] = bitsData[srcIdx + 0];
      imageData.data[dstOffset + x * 4 + 3] = biBitCount === 32 ? bitsData[srcIdx + 3] : 255;
    }
  }

  return { type: "bitmap", imageData };
}

function findSequence(buf: Uint8Array, seq: number[]): number {
  const end = buf.length - seq.length;
  for (let i = 0; i <= end; i++) {
    if (matchesAt(buf, i, seq)) return i;
  }
  return -1;
}

function matchesAt(buf: Uint8Array, offset: number, seq: number[]): boolean {
  for (let j = 0; j < seq.length; j++) {
    if (buf[offset + j] !== seq[j]) return false;
  }
  return true;
}
