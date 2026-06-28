export interface ResolvedMedia {
  mediaPath: string;
  data: Uint8Array;
}

export interface MediaResolver {
  resolve(target: string): Promise<ResolvedMedia | undefined>;
}

export function getMimeType(path: string): string {
  const ext = path.split('.').pop()?.toLowerCase() || '';
  const mimeMap: Record<string, string> = {
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    gif: 'image/gif',
    svg: 'image/svg+xml',
    bmp: 'image/bmp',
    tiff: 'image/tiff',
    tif: 'image/tiff',
    emf: 'image/x-emf',
    wmf: 'image/x-wmf',
    webp: 'image/webp',
    mp4: 'video/mp4',
    m4v: 'video/mp4',
    webm: 'video/webm',
    mp3: 'audio/mpeg',
    wav: 'audio/wav',
    m4a: 'audio/mp4',
    ogg: 'audio/ogg',
  };
  return mimeMap[ext] || 'application/octet-stream';
}

function stripUriSuffix(target: string): string {
  const suffixIndex = target.search(/[?#]/);
  return suffixIndex >= 0 ? target.slice(0, suffixIndex) : target;
}

function normalizePathSegments(path: string): string[] {
  const normalized: string[] = [];
  for (const part of path.replace(/\\/g, '/').split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') { normalized.pop(); continue; }
    normalized.push(part);
  }
  return normalized;
}

function decodeUriPathSegment(segment: string): string {
  try { return decodeURIComponent(segment); } catch { return segment; }
}

function mediaRelativePath(target: string): string {
  const parts = normalizePathSegments(stripUriSuffix(target));
  const mediaIndex = parts.lastIndexOf('media');
  const mediaParts =
    mediaIndex >= 0 && mediaIndex < parts.length - 1
      ? parts.slice(mediaIndex + 1)
      : parts.slice(-1);
  return mediaParts.join('/');
}

export function resolveMediaPath(target: string): string {
  const decodedPath = mediaRelativePath(target).split('/').map(decodeUriPathSegment).join('/');
  return `ppt/media/${decodedPath}`;
}

export function resolveMediaPathCandidates(target: string): string[] {
  const rawRelativePath = mediaRelativePath(target);
  const decodedPath = resolveMediaPath(target);
  const rawPath = `ppt/media/${rawRelativePath}`;
  return decodedPath === rawPath ? [decodedPath] : [decodedPath, rawPath];
}

export function findMediaByTarget(
  target: string,
  media: Map<string, Uint8Array>,
): ResolvedMedia | undefined {
  for (const mediaPath of resolveMediaPathCandidates(target)) {
    const data = media.get(mediaPath);
    if (data) return { mediaPath, data };
  }
  return undefined;
}

export async function findMediaByTargetAsync(
  target: string,
  media: Map<string, Uint8Array>,
  resolver?: MediaResolver,
): Promise<ResolvedMedia | undefined> {
  const eager = findMediaByTarget(target, media);
  if (eager) return eager;
  return resolver?.resolve(target);
}

export function getOrCreateBlobUrl(
  mediaPath: string,
  data: Uint8Array | ArrayBuffer,
  cache: Map<string, string>,
): string {
  let url = cache.get(mediaPath);
  if (!url) {
    const mime = getMimeType(mediaPath);
    const blob = new Blob([data as BlobPart], { type: mime });
    url = URL.createObjectURL(blob);
    cache.set(mediaPath, url);
  }
  return url;
}
