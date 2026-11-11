import { MAIN_CONTENT_ID, SkipNav } from "@pptx/ui/components/skip-nav";
import { HomeLayout } from "fumadocs-ui/layouts/home";

import { baseOptions } from "@/lib/layout.shared";

export default function Layout({ children }: LayoutProps<"/">) {
  return (
    <>
      <SkipNav />
      <HomeLayout
        {...baseOptions()}
        links={[
          { text: "Docs", url: "/docs" },
          { text: "Components", url: "/docs/components" },
          { text: "Editing", url: "/docs/editing" },
          { text: "Store", url: "/docs/store" },
        ]}
      >
        <div id={MAIN_CONTENT_ID} className="flex flex-1 scroll-mt-14 flex-col">
          {children}
        </div>
      </HomeLayout>
    </>
  );
}
