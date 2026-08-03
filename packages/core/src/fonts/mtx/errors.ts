export type MtxErrorCode =
  | "BOUNDS"
  | "INVALID_EOT"
  | "INVALID_MTX"
  | "INVALID_CTF"
  | "INVALID_SFNT"
  | "UNSUPPORTED"
  | "LIMIT_EXCEEDED";

export class MtxError extends Error {
  readonly code: MtxErrorCode;
  readonly offset: number | undefined;

  constructor(code: MtxErrorCode, message: string, offset?: number) {
    super(offset === undefined ? message : `${message} (offset ${offset})`);
    this.name = "MtxError";
    this.code = code;
    this.offset = offset;
  }
}

export function fail(code: MtxErrorCode, message: string, offset?: number): never {
  throw new MtxError(code, message, offset);
}
