import { isMarkdownPreferred, rewritePath } from "fumadocs-core/negotiation";
import { NextRequest, NextResponse } from "next/server";

import { DOCS_CONTENT_ROUTE, DOCS_ROUTE } from "@/lib/constants";

const docsPathRewriter = rewritePath(
  `${DOCS_ROUTE}{/*path}`,
  `${DOCS_CONTENT_ROUTE}{/*path}/content.md`,
);
const suffixPathRewriter = rewritePath(
  `${DOCS_ROUTE}{/*path}.md`,
  `${DOCS_CONTENT_ROUTE}{/*path}/content.md`,
);

export default function proxy(request: NextRequest) {
  const result = suffixPathRewriter.rewrite(request.nextUrl.pathname);
  if (result) {
    return NextResponse.rewrite(new URL(result, request.nextUrl));
  }

  if (isMarkdownPreferred(request)) {
    const result = docsPathRewriter.rewrite(request.nextUrl.pathname);

    if (result) {
      return NextResponse.rewrite(new URL(result, request.nextUrl));
    }
  }

  return NextResponse.next();
}
