import * as React from "react";

import { PresentationPlayground } from "@/components/presentation-playground";

export default function PgPage({ searchParams }: PageProps<"/pg">) {
  return (
    <React.Suspense
      fallback={
        <div className="mx-auto flex w-full max-w-(--fd-layout-width) flex-col gap-2 p-4">
          <div className="h-9" />
          <div className="h-[calc(100dvh-(--spacing(32)))] rounded-md border" />
        </div>
      }
    >
      <PresentationPlayground searchParams={searchParams} />
    </React.Suspense>
  );
}
