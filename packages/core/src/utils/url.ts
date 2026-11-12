const ALLOWED_PROTOCOLS = new Set(["http:", "https:", "mailto:"]);
const ALLOWED_MEDIA_PROTOCOLS = new Set(["http:", "https:"]);

function getUrlProtocol(url: string): string | undefined {
  try {
    return new URL(url).protocol.toLowerCase();
  } catch {
    return undefined;
  }
}

export function isAllowedExternalUrl(url: string): boolean {
  const protocol = getUrlProtocol(url);
  return protocol !== undefined && ALLOWED_PROTOCOLS.has(protocol);
}

export function isAllowedExternalMediaUrl(url: string): boolean {
  const protocol = getUrlProtocol(url);
  return protocol !== undefined && ALLOWED_MEDIA_PROTOCOLS.has(protocol);
}
