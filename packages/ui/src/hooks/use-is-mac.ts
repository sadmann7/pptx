import * as React from "react";

function subscribe() {
  return () => {};
}

function getIsMac() {
  const platform =
    (navigator as Navigator & { userAgentData?: { platform?: string } }).userAgentData?.platform ??
    navigator.platform;
  return /mac|iphone|ipad|ipod/i.test(platform);
}

/** Whether the user is on an Apple platform, where shortcuts use ⌘ instead of Ctrl. */
function useIsMac() {
  return React.useSyncExternalStore(subscribe, getIsMac, () => false);
}

export { useIsMac };
