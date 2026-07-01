import type { BaseLayoutProps } from "fumadocs-ui/layouts/shared";

import { siteConfig } from "@/lib/site";

export function baseOptions(): BaseLayoutProps {
  return {
    nav: {
      // JSX supported
      title: siteConfig.name,
    },
    githubUrl: siteConfig.links.github,
  };
}
