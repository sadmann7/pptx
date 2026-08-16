import * as React from "react";

import { useDelayRender } from "remotion";
import { createHighlighterCore, type HighlighterCore, type ThemedToken } from "shiki/core";
import { createJavaScriptRegexEngine } from "shiki/engine/javascript";
import tsx from "shiki/langs/tsx.mjs";
import githubDark from "shiki/themes/github-dark.mjs";

const LANG = "tsx";
const THEME = "github-dark";

let highlighter: Promise<HighlighterCore> | null = null;

/**
 * One grammar and one theme, imported directly rather than through Shiki's
 * bundled loader, so the video ships a highlighter instead of every language it
 * knows. The JavaScript engine avoids pulling in the Oniguruma wasm.
 */
function getHighlighter() {
  highlighter ??= createHighlighterCore({
    langs: [tsx],
    themes: [githubDark],
    engine: createJavaScriptRegexEngine(),
  });
  return highlighter;
}

/**
 * Tokenises a snippet into one array per line, holding Remotion's render until
 * the highlighter is ready so no frame is captured against unstyled code.
 */
export function useHighlightedLines(code: string): ThemedToken[][] | null {
  const { delayRender, continueRender, cancelRender } = useDelayRender();
  const [handle] = React.useState(() => delayRender("highlight snippet"));
  const [lines, setLines] = React.useState<ThemedToken[][] | null>(null);

  React.useEffect(() => {
    let cancelled = false;
    getHighlighter()
      .then((core) => {
        if (cancelled) return;
        setLines(core.codeToTokens(code, { lang: LANG, theme: THEME }).tokens);
        continueRender(handle);
      })
      .catch((error) => {
        if (!cancelled) cancelRender(error);
      });
    return () => {
      cancelled = true;
      continueRender(handle);
    };
  }, [code, handle, continueRender, cancelRender]);

  return lines;
}
