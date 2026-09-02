"use client";

import { TooltipProvider } from "@pptx/ui/components/tooltip";
import type { RootProviderProps } from "fumadocs-ui/provider/base";
import { RootProvider } from "fumadocs-ui/provider/next";
interface ProvidersProps extends RootProviderProps {
  children: React.ReactNode;
}

export function Providers({ children, ...props }: ProvidersProps) {
  return (
    <RootProvider {...props}>
      <TooltipProvider>{children}</TooltipProvider>
    </RootProvider>
  );
}
