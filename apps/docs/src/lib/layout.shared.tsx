import type { BaseLayoutProps } from "fumadocs-ui/layouts/shared";

import { LogoMark } from "@/components/icons";
import { GITHUB_URL, SITE_NAME } from "@/lib/site";

export function baseOptions(): BaseLayoutProps {
  return {
    nav: {
      title: <SiteLogo />,
    },
    githubUrl: GITHUB_URL,
  };
}

export function SiteLogo() {
  return (
    <>
      <span className="flex size-5.5 items-center justify-center rounded-[5px] bg-primary text-primary-foreground">
        <LogoMark className="size-4" />
      </span>
      <span className="font-semibold tracking-tight">{SITE_NAME}</span>
    </>
  );
}
