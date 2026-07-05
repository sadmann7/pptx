import type * as React from "react";

import type { Metadata, Viewport } from "next";

import { Providers } from "@/components/providers";
import { TailwindIndicator } from "@/components/tailwind-indicator";
import { Toaster } from "@/components/ui/sonner";
import { fontMono, fontSans } from "@/lib/fonts";
import { siteConfig } from "@/lib/site";
import { cn } from "@/lib/utils";
import "@/styles/globals.css";

export const metadata: Metadata = {
  metadataBase: new URL(siteConfig.url),
  title: {
    default: siteConfig.name,
    template: `%s - ${siteConfig.name}`,
  },
  description: siteConfig.description,
  keywords: ["pptx", "pptx-viewer", "pptx-parser", "powerpoint", "presentation", "slides"],
  authors: [
    {
      name: "sadmann7",
      url: "https://www.sadmn.com",
    },
  ],
  creator: "sadmann7",
  openGraph: {
    type: "website",
    locale: "en_US",
    url: siteConfig.url,
    title: siteConfig.name,
    description: siteConfig.description,
    siteName: siteConfig.name,
    images: [
      {
        url: "/opengraph-image.png",
        alt: siteConfig.name,
      },
    ],
  },
  twitter: {
    card: "summary_large_image",
    title: siteConfig.name,
    description: siteConfig.description,
    images: [
      {
        url: "/opengraph-image.png",
        alt: siteConfig.name,
      },
    ],
    creator: "@sadmann17",
  },
  icons: {
    icon: "/icon.png",
  },
};

export const viewport: Viewport = {
  themeColor: [
    { media: "(prefers-color-scheme: light)", color: "white" },
    { media: "(prefers-color-scheme: dark)", color: "black" },
  ],
};

interface RootLayoutProps {
  children: React.ReactNode;
}

export default function RootLayout({ children }: RootLayoutProps) {
  return (
    <html lang="en" suppressHydrationWarning>
      <body
        className={cn(
          "isolate min-h-screen font-sans antialiased",
          fontSans.variable,
          fontMono.variable,
        )}
      >
        <TailwindIndicator />
        <Toaster />
        <Providers
          theme={{
            attribute: "class",
            defaultTheme: "system",
            enableSystem: true,
            disableTransitionOnChange: true,
          }}
        >
          {children}
        </Providers>
      </body>
    </html>
  );
}
