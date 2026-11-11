import type { BaseLayoutProps } from "fumadocs-ui/layouts/shared";

import { siteConfig } from "@/lib/site";

export function baseOptions(): BaseLayoutProps {
  return {
    nav: {
      title: <SiteLogo />,
    },
    githubUrl: siteConfig.links.github,
  };
}

export function SiteLogo() {
  return (
    <>
      <span className="flex size-5.5 items-center justify-center rounded-[5px] bg-primary">
        <span className="h-2 w-3 rounded-[1.5px] border-[1.5px] border-primary-foreground" />
      </span>
      <span className="font-semibold tracking-tight">{siteConfig.name}</span>
    </>
  );
}
