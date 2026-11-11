import { MAIN_CONTENT_ID, SkipNav } from "@pptx/ui/components/skip-nav";
import { HomeLayout } from "fumadocs-ui/layouts/home";

import { SiteHeader } from "@/components/site-header";
import { baseOptions } from "@/lib/layout.shared";

export default function Layout({ children }: LayoutProps<"/">) {
  return (
    <>
      <SkipNav />
      <HomeLayout {...baseOptions()} slots={{ header: SiteHeader }}>
        <div id={MAIN_CONTENT_ID} className="flex flex-1 scroll-mt-14 flex-col md:scroll-mt-16">
          {children}
        </div>
      </HomeLayout>
    </>
  );
}
