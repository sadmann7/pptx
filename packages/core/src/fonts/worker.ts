/**
 * Web Worker entry for embedded-font decompression.
 *
 * Receives one job per message and posts the decoded TrueType bytes back
 * as a transferable ArrayBuffer (or `null` when decoding fails).
 */

import { copyToArrayBuffer, decodeEmbeddedFontPart } from "./decode";
import type { MtxErrorCode } from "./mtx";

export interface FontWorkerRequest {
  path: string;
  bytes: ArrayBuffer;
  fontKey?: string;
}

export interface FontWorkerResponse {
  path: string;
  buffer: ArrayBuffer | null;
  /** Why decoding failed, set whenever `buffer` is `null`. */
  message?: string;
  code?: MtxErrorCode;
}

const workerScope = self as unknown as {
  onmessage: ((event: MessageEvent<FontWorkerRequest>) => void) | null;
  postMessage(message: FontWorkerResponse, transfer?: Transferable[]): void;
};

workerScope.onmessage = (event) => {
  const { path, bytes, fontKey } = event.data;
  try {
    const result = decodeEmbeddedFontPart(new Uint8Array(bytes), fontKey);
    if (result.ok && result.bytes.length > 0) {
      const buffer = copyToArrayBuffer(result.bytes);
      workerScope.postMessage({ path, buffer }, [buffer]);
    } else {
      workerScope.postMessage({
        path,
        buffer: null,
        message: result.ok ? "Decoded font is empty" : result.message,
        code: result.ok ? undefined : result.code,
      });
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    workerScope.postMessage({ path, buffer: null, message });
  }
};
