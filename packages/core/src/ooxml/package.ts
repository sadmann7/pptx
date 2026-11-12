import JSZip from "jszip";

import { SafeXmlNode } from "./xml";

const DEFAULT_XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

// Model paths like slide ids and rel targets may be percent-decoded while zip
// entry paths keep the original encoding.
export function decodeZipPath(path: string): string {
  return path
    .split("/")
    .map((segment) => {
      try {
        return decodeURIComponent(segment);
      } catch {
        return segment;
      }
    })
    .join("/");
}

// Stateless helpers shared across parts (a deck has hundreds of entries).
let sharedSerializer: XMLSerializer | undefined;
let sharedEncoder: TextEncoder | undefined;
let sharedDecoder: TextDecoder | undefined;

function encodeText(text: string): Uint8Array {
  sharedEncoder ??= new TextEncoder();
  const encoded = sharedEncoder.encode(text);
  // TextEncoder may come from another realm (e.g. jsdom), whose Uint8Array
  // fails JSZip's instanceof checks; rewrap in this realm's constructor.
  return encoded instanceof Uint8Array ? encoded : new Uint8Array(encoded);
}

// XMLSerializer never emits the `<?xml ...?>` declaration, so a re-serialized
// part keeps the source's, trailing newline included.
function extractXmlDeclaration(bytes: Uint8Array): string | undefined {
  sharedDecoder ??= new TextDecoder("utf-8");
  const head = sharedDecoder.decode(bytes.subarray(0, 256));
  const match = /^<\?xml.*?\?>\r?\n?/.exec(head);
  return match?.[0];
}

export interface PptxSaveOptions {
  /** Zip compression for the written archive. @default "DEFLATE". */
  compression?: "STORE" | "DEFLATE";
}

/**
 * The source package kept for round-trip saves, which copies untouched parts
 * byte for byte so animations, comments, and vendor extensions survive an edit.
 * Don't detach the input buffer, such as by transferring it to a worker, while
 * saving is still needed.
 */
export class PptxPackage {
  // Entries stay compressed until read.
  private readonly zip: JSZip;
  // Normalized or percent-decoded path to the actual zip entry key.
  private readonly pathAliases = new Map<string, string>();
  private readonly overrides = new Map<string, Uint8Array>();
  private readonly deleted = new Set<string>();
  private readonly xmlRoots = new Map<string, SafeXmlNode>();
  private readonly dirty = new Set<string>();

  constructor(zip: JSZip) {
    this.zip = zip;
    for (const [key, file] of Object.entries(zip.files)) {
      if (file.dir) continue;
      const normalized = key.replace(/\\/g, "/");
      this.pathAliases.set(normalized, key);
      const decoded = decodeZipPath(normalized);
      if (decoded !== normalized && !this.pathAliases.has(decoded)) {
        this.pathAliases.set(decoded, key);
      }
    }
  }

  /** Lists the canonical paths of every part in the package. */
  paths(): string[] {
    const result = new Set<string>();
    for (const [key, file] of Object.entries(this.zip.files)) {
      if (file.dir || this.deleted.has(key)) continue;
      result.add(key);
    }
    for (const key of this.overrides.keys()) {
      if (!this.deleted.has(key)) result.add(key);
    }
    return [...result];
  }

  has(path: string): boolean {
    const key = this.canonical(path);
    if (this.deleted.has(key)) return false;
    return this.overrides.has(key) || Boolean(this.zip.file(key));
  }

  /** The live parsed XML root for a part, when the model registered one. */
  getXmlRoot(path: string): SafeXmlNode | undefined {
    return this.xmlRoots.get(this.canonical(path));
  }

  /**
   * Links a part to its parsed XML root so DOM edits save. The first
   * registration wins, since a theme reused as a chart theme override is parsed
   * more than once and only the document the model references must be saved.
   *
   * @internal
   */
  registerXmlRoot(path: string, root: SafeXmlNode): void {
    if (!root.exists()) return;
    const key = this.canonical(path);
    if (!this.xmlRoots.has(key)) {
      this.xmlRoots.set(key, root);
    }
  }

  /** Marks a part as edited so save re-serializes it from the live DOM. */
  markDirty(path: string): void {
    const key = this.canonical(path);
    if (!this.xmlRoots.has(key)) {
      throw new Error(
        `PptxPackage.markDirty: no XML document registered for "${path}". ` +
          "Use setEntry() to replace raw part content directly.",
      );
    }
    this.dirty.add(key);
  }

  isDirty(path: string): boolean {
    return this.dirty.has(this.canonical(path));
  }

  /** Replaces or adds a part's raw content, discarding its registered XML root and dirty flag. */
  setEntry(path: string, content: Uint8Array | string): void {
    const key = this.canonical(path);
    this.overrides.set(key, typeof content === "string" ? encodeText(content) : content);
    this.deleted.delete(key);
    this.dirty.delete(key);
    this.xmlRoots.delete(key);
  }

  /** Removes a part, returning whether it existed. */
  deleteEntry(path: string): boolean {
    const key = this.canonical(path);
    const existed = this.has(key);
    this.deleted.add(key);
    this.overrides.delete(key);
    this.dirty.delete(key);
    this.xmlRoots.delete(key);
    return existed;
  }

  /** Reads a part's content as save would write it. */
  async readBytes(path: string): Promise<Uint8Array | undefined> {
    const key = this.canonical(path);
    if (this.deleted.has(key)) return undefined;

    if (this.dirty.has(key)) {
      return this.serializeDirtyPart(key);
    }

    const override = this.overrides.get(key);
    if (override) return override;

    const file = this.zip.file(key);
    if (!file || file.dir) return undefined;
    return file.async("uint8array");
  }

  /** UTF-8 decoded {@link readBytes}. */
  async readText(path: string): Promise<string | undefined> {
    const bytes = await this.readBytes(path);
    if (bytes === undefined) return undefined;
    sharedDecoder ??= new TextDecoder("utf-8");
    return sharedDecoder.decode(bytes);
  }

  /** Writes the package to a new .pptx archive. */
  async save(options: PptxSaveOptions = {}): Promise<Uint8Array> {
    const out = new JSZip();

    for (const key of this.paths()) {
      // JSZip accepts a promise as content; source entries are only
      // decompressed when the output archive is generated.
      out.file(
        key,
        this.readBytes(key).then((bytes) => {
          if (!bytes) throw new Error(`PptxPackage: part "${key}" disappeared during save.`);
          return bytes;
        }),
        { binary: true },
      );
    }

    return out.generateAsync({
      type: "uint8array",
      compression: options.compression ?? "DEFLATE",
    });
  }

  private async serializeDirtyPart(key: string): Promise<Uint8Array> {
    const element = this.xmlRoots.get(key)?.element;
    if (!element) {
      throw new Error(`PptxPackage: dirty part "${key}" has no XML document to serialize.`);
    }

    sharedSerializer ??= new XMLSerializer();
    const xml = sharedSerializer.serializeToString(element);
    const original = this.overrides.get(key) ?? (await this.zip.file(key)?.async("uint8array"));
    const declaration = (original && extractXmlDeclaration(original)) ?? DEFAULT_XML_DECLARATION;
    return encodeText(declaration + xml);
  }

  private canonical(path: string): string {
    return this.pathAliases.get(path) ?? this.pathAliases.get(decodeZipPath(path)) ?? path;
  }
}
