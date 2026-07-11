/**
 * Web Worker entry for embedded-font decompression.
 *
 * Receives one job per message and posts the decoded TrueType bytes back
 * as a transferable ArrayBuffer (or `null` when decoding fails).
 */

import { decodeEmbeddedFont, toStandaloneArrayBuffer } from "./decode";

export interface FontWorkerRequest {
  path: string;
  bytes: ArrayBuffer;
  fontKey?: string;
}

export interface FontWorkerResponse {
  path: string;
  buffer: ArrayBuffer | null;
}

const workerScope = self as unknown as {
  onmessage: ((event: MessageEvent<FontWorkerRequest>) => void) | null;
  postMessage(message: FontWorkerResponse, transfer?: Transferable[]): void;
};

workerScope.onmessage = (event) => {
  const { path, bytes, fontKey } = event.data;
  try {
    const decoded = decodeEmbeddedFont(new Uint8Array(bytes), fontKey);
    if (decoded && decoded.length > 0) {
      const buffer = toStandaloneArrayBuffer(decoded);
      workerScope.postMessage({ path, buffer }, [buffer]);
    } else {
      workerScope.postMessage({ path, buffer: null });
    }
  } catch {
    workerScope.postMessage({ path, buffer: null });
  }
};
