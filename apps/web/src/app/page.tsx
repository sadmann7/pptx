"use client";

import { Presentation } from "@pptx/react";
import * as React from "react";

export default function IndexPage() {
  const [file, setFile] = React.useState<File | null>(null);

  return (
    <div className="container mx-auto max-w-3xl px-4 py-2">
      <input
        type="file"
        onChange={(e) => {
          const file = e.target.files?.[0];
          if (file) {
            setFile(file);
          }
        }}
      />
      <Presentation.Root file={file}>
        <div style={{ display: "flex", height: "100vh" }}>
          <Presentation.Thumbnails style={{ width: 160 }} />
          <div style={{ flex: 1, display: "flex", flexDirection: "column" }}>
            <Presentation.Viewport style={{ flex: 1 }} autoFit>
              <Presentation.Slide />
            </Presentation.Viewport>
            <Presentation.Notes style={{ height: 120 }} />
          </div>
        </div>
      </Presentation.Root>
    </div>
  );
}
